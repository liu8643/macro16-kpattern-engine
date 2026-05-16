# -*- coding: utf-8 -*-
"""
GTC Workflow Orchestrator V17
每日增量更新 / 快速重算排行 / AI選股TOP20 工作邊界控制模組

設計目的
========
本模組是依「每日增量更新_重建排行_AI選股TOP20_工作重組分析」重新定義三個按鈕/流程的責任邊界：

1. Daily_Update_Master：唯一重型資料入口
   - 允許外部 API、price_history 寫入、financial_feature_daily 建立、EPS Matrix BUILD/DECISION、
     ranking_result 更新、teacher snapshot 建立。
2. FAST_RANK_REBUILD：快速重算排行
   - 只允許讀 DB 快取、重算技術/策略分數、更新 ranking_result、必要欄位 teacher merge。
   - 禁止外部 API、禁止重建 EPS Matrix、禁止重建 financial_feature_daily。
3. AI_TOP20_VIEW：純查詢/顯示
   - 只允許讀取 ranking_result / trade_plan / teacher_strategy_snapshot 快照。
   - 禁止任何 update / rebuild / build / 外部 API / 全量 merge。

整合方式
========
在主程式中引入：

    from gtc_workflow_orchestrator_v17 import (
        GTCWorkflowOrchestrator,
        WorkflowMode,
        WorkflowViolation,
    )

使用原則：
- 每日增量更新按鈕：呼叫 orchestrator.daily_update_master(...)
- 重建排行按鈕：改名為「快速重算排行」，呼叫 orchestrator.fast_rank_rebuild(...)
- AI選股TOP20：呼叫 orchestrator.ai_top20_view(...)

重要：
本模組不直接改寫你的 DB schema，不直接抓資料，不直接重算股票。
它是「流程總閘 + 防誤用 guard + log驗收工具」，用來防止主程式再次把三個流程混在一起。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
import functools
import logging
import re
import threading
import time
import uuid


class WorkflowMode(str, Enum):
    """三個標準工作模式。"""

    INIT_MARKET = "INIT_MARKET"
    BUILD_FULL_HISTORY = "BUILD_FULL_HISTORY"
    DAILY_UPDATE_MASTER = "Daily_Update_Master"
    FAST_RANK_REBUILD = "FAST_RANK_REBUILD"
    AI_TOP20_VIEW = "AI_TOP20_VIEW"


class Operation(str, Enum):
    """可被 guard 管控的工作項目。"""

    EXTERNAL_API_FETCH = "external_api_fetch"
    PRICE_HISTORY_WRITE = "price_history_write"
    EXTERNAL_TABLE_WRITE = "external_table_write"
    FINANCIAL_FEATURE_WRITE = "financial_feature_daily_write"
    EPS_MATRIX_BUILD = "eps_matrix_build"
    EPS_MATRIX_DECISION_BULK = "eps_matrix_decision_bulk"
    TECHNICAL_SCORE_RECALC = "technical_score_recalc"
    RANKING_RESULT_WRITE = "ranking_result_write"
    TEACHER_FULL_MERGE = "teacher_full_merge"
    TEACHER_LIGHT_MERGE = "teacher_light_merge"
    TEACHER_SNAPSHOT_WRITE = "teacher_snapshot_write"
    SNAPSHOT_READ = "snapshot_read"
    UI_RENDER = "ui_render"
    EXPORT_EXCEL = "export_excel"
    HEAVY_MAINTHREAD_WORK = "heavy_mainthread_work"


class WorkflowViolation(RuntimeError):
    """流程違規：代表某個模式呼叫了不該做的重工作。"""


@dataclass(frozen=True)
class PermissionRule:
    operation: Operation
    daily_update_master: bool
    fast_rank_rebuild: bool
    ai_top20_view: bool
    reason: str


# 權限矩陣：對應分析報告「06_驗收條件」
PERMISSION_MATRIX: Dict[Operation, PermissionRule] = {
    Operation.EXTERNAL_API_FETCH: PermissionRule(
        Operation.EXTERNAL_API_FETCH, True, False, False,
        "外部 API 只允許每日增量更新；重算排行與 TOP20 必須只讀快取。",
    ),
    Operation.PRICE_HISTORY_WRITE: PermissionRule(
        Operation.PRICE_HISTORY_WRITE, True, False, False,
        "price_history 是原始行情資料，只能由每日增量更新寫入。",
    ),
    Operation.EXTERNAL_TABLE_WRITE: PermissionRule(
        Operation.EXTERNAL_TABLE_WRITE, True, False, False,
        "external_valuation/revenue/margin 等外部表只能由每日增量更新寫入。",
    ),
    Operation.FINANCIAL_FEATURE_WRITE: PermissionRule(
        Operation.FINANCIAL_FEATURE_WRITE, True, False, False,
        "financial_feature_daily / EPS Matrix 快取只能每日建立一次。",
    ),
    Operation.EPS_MATRIX_BUILD: PermissionRule(
        Operation.EPS_MATRIX_BUILD, True, False, False,
        "EPS MATRIX BUILD 禁止出現在快速重算排行與 AI TOP20。",
    ),
    Operation.EPS_MATRIX_DECISION_BULK: PermissionRule(
        Operation.EPS_MATRIX_DECISION_BULK, True, False, False,
        "2080/2267 筆 EPS MATRIX DECISION 屬於重型批次，只能每日更新跑一次。",
    ),
    Operation.TECHNICAL_SCORE_RECALC: PermissionRule(
        Operation.TECHNICAL_SCORE_RECALC, True, True, False,
        "技術分數可在每日更新後或快速重算排行時重算；TOP20不得重算。",
    ),
    Operation.RANKING_RESULT_WRITE: PermissionRule(
        Operation.RANKING_RESULT_WRITE, True, True, False,
        "ranking_result 可由每日更新或快速重排更新；TOP20只能讀。",
    ),
    Operation.TEACHER_FULL_MERGE: PermissionRule(
        Operation.TEACHER_FULL_MERGE, True, False, False,
        "TeacherStrategy 全量 merge 只能在每日更新或背景批次執行。",
    ),
    Operation.TEACHER_LIGHT_MERGE: PermissionRule(
        Operation.TEACHER_LIGHT_MERGE, True, True, False,
        "快速重算排行只允許必要欄位 teacher light merge；TOP20只讀快照。",
    ),
    Operation.TEACHER_SNAPSHOT_WRITE: PermissionRule(
        Operation.TEACHER_SNAPSHOT_WRITE, True, False, False,
        "teacher_strategy_snapshot 寫入屬於每日批次，不應由重排/TOP20觸發。",
    ),
    Operation.SNAPSHOT_READ: PermissionRule(
        Operation.SNAPSHOT_READ, True, True, True,
        "三種流程都可讀快照，但只有 AI_TOP20_VIEW 應以讀快照為主。",
    ),
    Operation.UI_RENDER: PermissionRule(
        Operation.UI_RENDER, True, True, True,
        "UI render 三者都允許，但 MainThread 僅能做顯示，不得做重工作。",
    ),
    Operation.EXPORT_EXCEL: PermissionRule(
        Operation.EXPORT_EXCEL, True, True, True,
        "匯出可讀目前快照；不可在匯出時觸發重建。",
    ),
    Operation.HEAVY_MAINTHREAD_WORK: PermissionRule(
        Operation.HEAVY_MAINTHREAD_WORK, False, False, False,
        "任何模式都不應在 MainThread 做重型 bulk work。",
    ),
}


@dataclass
class WorkflowEvent:
    run_id: str
    mode: WorkflowMode
    stage: str
    status: str
    message: str = ""
    rows: int = 0
    duration_sec: float = 0.0
    thread_name: str = field(default_factory=lambda: threading.current_thread().name)
    event_time: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


@dataclass
class WorkflowContext:
    mode: WorkflowMode
    run_id: str
    strict: bool = True
    allow_mainthread_render_only: bool = True
    metadata: Dict[str, Any] = field(default_factory=dict)


class WorkflowLogger:
    """統一流程 log：輸出 START/END/duration/rows/thread/mode。"""

    def __init__(self, logger: Optional[logging.Logger] = None, log_cb: Optional[Callable[[str], None]] = None):
        self.logger = logger or logging.getLogger("gtc.workflow")
        self.log_cb = log_cb
        self.events: List[WorkflowEvent] = []

    def emit(self, event: WorkflowEvent) -> None:
        self.events.append(event)
        text = (
            f"[WORKFLOW][{event.mode.value}][{event.status}] "
            f"run_id={event.run_id} stage={event.stage} rows={event.rows} "
            f"duration={event.duration_sec:.3f}s thread={event.thread_name} {event.message}"
        )
        try:
            self.logger.info(text)
        except Exception:
            pass
        if self.log_cb:
            try:
                self.log_cb(text)
            except Exception:
                pass

    def start(self, ctx: WorkflowContext, stage: str, message: str = "") -> float:
        self.emit(WorkflowEvent(ctx.run_id, ctx.mode, stage, "START", message=message))
        return time.perf_counter()

    def end(self, ctx: WorkflowContext, stage: str, start_time: float, rows: int = 0, message: str = "") -> None:
        self.emit(WorkflowEvent(
            ctx.run_id, ctx.mode, stage, "END",
            message=message,
            rows=int(rows or 0),
            duration_sec=max(time.perf_counter() - start_time, 0.0),
        ))

    def warning(self, ctx: WorkflowContext, stage: str, message: str = "") -> None:
        self.emit(WorkflowEvent(ctx.run_id, ctx.mode, stage, "WARNING", message=message))

    def violation(self, ctx: WorkflowContext, stage: str, message: str = "") -> None:
        self.emit(WorkflowEvent(ctx.run_id, ctx.mode, stage, "VIOLATION", message=message))


class WorkflowGuard:
    """流程權限總閘。"""

    def __init__(self, logger: Optional[WorkflowLogger] = None):
        self.workflow_logger = logger or WorkflowLogger()

    @staticmethod
    def _allowed(mode: WorkflowMode, operation: Operation) -> bool:
        rule = PERMISSION_MATRIX[operation]
        if mode == WorkflowMode.DAILY_UPDATE_MASTER:
            return rule.daily_update_master
        if mode == WorkflowMode.FAST_RANK_REBUILD:
            return rule.fast_rank_rebuild
        if mode == WorkflowMode.AI_TOP20_VIEW:
            return rule.ai_top20_view
        return False

    def assert_allowed(self, ctx: WorkflowContext, operation: Operation, stage: str = "") -> None:
        if operation not in PERMISSION_MATRIX:
            raise WorkflowViolation(f"未知工作項目：{operation}")
        if not self._allowed(ctx.mode, operation):
            rule = PERMISSION_MATRIX[operation]
            msg = f"禁止在 {ctx.mode.value} 執行 {operation.value}：{rule.reason}"
            self.workflow_logger.violation(ctx, stage or operation.value, msg)
            if ctx.strict:
                raise WorkflowViolation(msg)

    def assert_not_heavy_mainthread(self, ctx: WorkflowContext, stage: str) -> None:
        if threading.current_thread().name == "MainThread":
            msg = f"重型工作不得在 MainThread 執行：mode={ctx.mode.value}, stage={stage}"
            self.workflow_logger.violation(ctx, stage, msg)
            if ctx.strict:
                raise WorkflowViolation(msg)


def _safe_len(obj: Any) -> int:
    try:
        return int(len(obj))
    except Exception:
        return 0


def _call_optional(fn: Optional[Callable[..., Any]], *args: Any, **kwargs: Any) -> Any:
    if fn is None:
        return None
    return fn(*args, **kwargs)


class GTCWorkflowOrchestrator:
    """三大流程重新組合後的標準入口。

    注意：
    - 本類只負責流程協調與防錯，不直接實作你的資料抓取或股票評分。
    - 你要把主程式既有函式以 callable 傳入，例如 update_daily_fn、build_financial_feature_fn。
    """

    def __init__(
        self,
        logger: Optional[logging.Logger] = None,
        log_cb: Optional[Callable[[str], None]] = None,
        strict: bool = True,
    ):
        self.workflow_logger = WorkflowLogger(logger=logger, log_cb=log_cb)
        self.guard = WorkflowGuard(self.workflow_logger)
        self.strict = strict

    def _ctx(self, mode: WorkflowMode, **metadata: Any) -> WorkflowContext:
        return WorkflowContext(
            mode=mode,
            run_id=f"{mode.value}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}",
            strict=self.strict,
            metadata=dict(metadata or {}),
        )

    @contextmanager
    def stage(
        self,
        ctx: WorkflowContext,
        stage_name: str,
        operation: Optional[Operation] = None,
        heavy: bool = False,
        message: str = "",
    ):
        if operation is not None:
            self.guard.assert_allowed(ctx, operation, stage_name)
        if heavy:
            self.guard.assert_not_heavy_mainthread(ctx, stage_name)
        start_time = self.workflow_logger.start(ctx, stage_name, message=message)
        rows = 0
        try:
            result = yield
            rows = _safe_len(result)
        finally:
            self.workflow_logger.end(ctx, stage_name, start_time, rows=rows)

    def initialize_market(self, *, build_universe_fn: Callable[..., Any], refresh_status_fn: Optional[Callable[..., Any]] = None, **kwargs: Any) -> Dict[str, Any]:
        """初始化全市場：只允許主檔/分類初始化與狀態刷新，不得重排行。"""
        ctx = self._ctx(WorkflowMode.INIT_MARKET)
        result: Dict[str, Any] = {"run_id": ctx.run_id, "mode": ctx.mode.value}
        with self.stage(ctx, "01_build_market_universe", None, heavy=True):
            result["universe"] = build_universe_fn(**kwargs)
        if refresh_status_fn:
            with self.stage(ctx, "02_refresh_master_status", Operation.UI_RENDER, heavy=False):
                result["ui"] = refresh_status_fn(**kwargs)
        return result

    def build_full_history(self, *, build_history_fn: Callable[..., Any], refresh_status_fn: Optional[Callable[..., Any]] = None, **kwargs: Any) -> Dict[str, Any]:
        """建立完整歷史：只允許補建 price_history，不得重排行或重建 EPS。"""
        ctx = self._ctx(WorkflowMode.BUILD_FULL_HISTORY)
        result: Dict[str, Any] = {"run_id": ctx.run_id, "mode": ctx.mode.value}
        with self.stage(ctx, "01_build_full_price_history", None, heavy=True):
            result["price_history"] = build_history_fn(**kwargs)
        if refresh_status_fn:
            with self.stage(ctx, "02_refresh_build_status", Operation.UI_RENDER, heavy=False):
                result["ui"] = refresh_status_fn(**kwargs)
        return result

    def daily_update_master(
        self,
        *,
        update_price_history_fn: Callable[..., Any],
        sync_external_tables_fn: Optional[Callable[..., Any]] = None,
        build_financial_feature_daily_fn: Optional[Callable[..., Any]] = None,
        eps_matrix_decision_fn: Optional[Callable[..., Any]] = None,
        rebuild_ranking_from_cache_fn: Optional[Callable[..., Any]] = None,
        build_teacher_snapshot_fn: Optional[Callable[..., Any]] = None,
        refresh_ui_fn: Optional[Callable[..., Any]] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """每日增量更新：唯一重型資料入口。

        允許：
        - 外部資料抓取
        - price_history / external_* / financial_feature_daily 寫入
        - EPS Matrix BUILD/DECISION
        - ranking_result 產出
        - teacher snapshot 建立
        """
        ctx = self._ctx(WorkflowMode.DAILY_UPDATE_MASTER)
        result: Dict[str, Any] = {"run_id": ctx.run_id, "mode": ctx.mode.value}

        with self.stage(ctx, "01_update_price_history", Operation.PRICE_HISTORY_WRITE, heavy=True):
            result["price_history"] = update_price_history_fn(**kwargs)

        if sync_external_tables_fn:
            with self.stage(ctx, "02_sync_external_tables", Operation.EXTERNAL_TABLE_WRITE, heavy=True):
                result["external_tables"] = sync_external_tables_fn(**kwargs)

        if build_financial_feature_daily_fn:
            with self.stage(ctx, "03_build_financial_feature_daily", Operation.FINANCIAL_FEATURE_WRITE, heavy=True):
                result["financial_feature_daily"] = build_financial_feature_daily_fn(**kwargs)

        if eps_matrix_decision_fn:
            with self.stage(ctx, "04_eps_matrix_decision_bulk", Operation.EPS_MATRIX_DECISION_BULK, heavy=True):
                result["eps_matrix_decision"] = eps_matrix_decision_fn(**kwargs)

        if rebuild_ranking_from_cache_fn:
            # 每日更新完成後可自動產出 ranking_result，但仍應讀取已建好的快取。
            with self.stage(ctx, "05_rebuild_ranking_result_after_update", Operation.RANKING_RESULT_WRITE, heavy=True):
                result["ranking_result"] = rebuild_ranking_from_cache_fn(rank_mode="daily_after_update", **kwargs)

        if build_teacher_snapshot_fn:
            with self.stage(ctx, "06_build_teacher_strategy_snapshot", Operation.TEACHER_SNAPSHOT_WRITE, heavy=True):
                result["teacher_snapshot"] = build_teacher_snapshot_fn(**kwargs)

        if refresh_ui_fn:
            with self.stage(ctx, "07_refresh_ui", Operation.UI_RENDER, heavy=False):
                result["ui"] = refresh_ui_fn(**kwargs)

        return result

    def fast_rank_rebuild(
        self,
        *,
        read_cache_fn: Callable[..., Any],
        rebuild_ranking_from_cache_fn: Callable[..., Any],
        teacher_light_merge_fn: Optional[Callable[..., Any]] = None,
        write_ranking_result_fn: Optional[Callable[..., Any]] = None,
        refresh_ui_fn: Optional[Callable[..., Any]] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """快速重算排行：原「重建排行」改名後的正確定位。

        允許：
        - 讀取 price_history / financial_feature_daily / teacher snapshot
        - 重算技術分數與策略分數
        - 更新 ranking_result

        禁止：
        - 外部 API
        - 更新 price_history / external_* / financial_feature_daily
        - EPS MATRIX BUILD
        - EPS MATRIX DECISION 逐檔
        - TeacherStrategy 全量 merge
        """
        ctx = self._ctx(WorkflowMode.FAST_RANK_REBUILD)
        result: Dict[str, Any] = {"run_id": ctx.run_id, "mode": ctx.mode.value}

        with self.stage(ctx, "01_read_cache", Operation.SNAPSHOT_READ, heavy=False):
            result["cache"] = read_cache_fn(**kwargs)

        with self.stage(ctx, "02_recalculate_technical_and_rank", Operation.TECHNICAL_SCORE_RECALC, heavy=True):
            result["ranking_df"] = rebuild_ranking_from_cache_fn(
                rank_mode="fast",
                allow_external_api=False,
                allow_eps_build=False,
                allow_financial_feature_write=False,
                **kwargs,
            )

        if teacher_light_merge_fn:
            with self.stage(ctx, "03_teacher_light_merge", Operation.TEACHER_LIGHT_MERGE, heavy=True):
                result["ranking_df"] = teacher_light_merge_fn(
                    result.get("ranking_df"),
                    only_columns=[
                        "stock_id", "teacher_score", "teacher_final_decision",
                        "teacher_light", "teacher_ui_bucket", "teacher_priority",
                        "teacher_no_trade_reason",
                    ],
                    **kwargs,
                )

        if write_ranking_result_fn:
            with self.stage(ctx, "04_write_ranking_result", Operation.RANKING_RESULT_WRITE, heavy=True):
                result["write_result"] = write_ranking_result_fn(result.get("ranking_df"), **kwargs)

        if refresh_ui_fn:
            with self.stage(ctx, "05_refresh_ui", Operation.UI_RENDER, heavy=False):
                result["ui"] = refresh_ui_fn(result.get("ranking_df"), **kwargs)

        return result

    def ai_top20_view(
        self,
        *,
        read_top20_snapshot_fn: Callable[..., Any],
        render_ui_fn: Optional[Callable[..., Any]] = None,
        export_fn: Optional[Callable[..., Any]] = None,
        top_n: int = 20,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """AI選股TOP20：純讀取快照，不得觸發任何 rebuild/update/build。"""
        ctx = self._ctx(WorkflowMode.AI_TOP20_VIEW, top_n=top_n)
        result: Dict[str, Any] = {"run_id": ctx.run_id, "mode": ctx.mode.value, "top_n": top_n}

        with self.stage(ctx, "01_read_top20_snapshot", Operation.SNAPSHOT_READ, heavy=False):
            result["top20"] = read_top20_snapshot_fn(top_n=top_n, **kwargs)

        if render_ui_fn:
            with self.stage(ctx, "02_render_ui_top20", Operation.UI_RENDER, heavy=False):
                result["ui"] = render_ui_fn(result.get("top20"), **kwargs)

        if export_fn:
            with self.stage(ctx, "03_export_top20_snapshot", Operation.EXPORT_EXCEL, heavy=False):
                result["export"] = export_fn(result.get("top20"), **kwargs)

        return result


# -----------------------------
# Log 驗收工具
# -----------------------------


@dataclass
class LogViolation:
    line_no: int
    level: str
    workflow: str
    violation_type: str
    evidence: str
    recommendation: str


FAST_RANK_FORBIDDEN_PATTERNS: Sequence[Tuple[str, str]] = (
    (r"\[EPS MATRIX\]\[BUILD\]", "FAST_RANK_REBUILD 不得執行 EPS MATRIX BUILD"),
    (r"\[EPS MATRIX\]\[MERGE\].*R7", "FAST_RANK_REBUILD 不得觸發 EPS MATRIX R7 MERGE"),
    (r"\[EPS MATRIX\]\[DECISION\]", "FAST_RANK_REBUILD 不得逐檔 EPS MATRIX DECISION"),
    (r"financial_feature_daily.*rows|replace_financial_feature|financial_feature_batch", "FAST_RANK_REBUILD 不得寫 financial_feature_daily"),
    (r"official fetch|TWSE|TPEx|Yahoo|requests|read timeout", "FAST_RANK_REBUILD 不得抓外部資料"),
)

AI_TOP20_FORBIDDEN_PATTERNS: Sequence[Tuple[str, str]] = (
    (r"rebuild|重建排行|ranking_rebuild", "AI_TOP20_VIEW 不得觸發重建排行"),
    (r"\[EPS MATRIX\]\[BUILD\]|\[EPS MATRIX\]\[MERGE\]|\[EPS MATRIX\]\[DECISION\]", "AI_TOP20_VIEW 不得觸發 EPS Matrix"),
    (r"update_daily|每日資料更新|price_history|financial_feature_daily", "AI_TOP20_VIEW 不得觸發資料更新"),
    (r"TeacherStrategy.*merged context=(ranking_rebuild|show_top20)", "AI_TOP20_VIEW 不得全量 TeacherStrategy merge"),
)

MAINTHREAD_HEAVY_PATTERNS: Sequence[Tuple[str, str]] = (
    (r"MainThread.*\[EPS MATRIX\]\[BUILD\]", "EPS MATRIX BUILD 不得在 MainThread"),
    (r"MainThread.*重排行進度\s+\d+/\d+", "重排行 bulk loop 不得在 MainThread"),
    (r"MainThread.*TeacherStrategy.*merged.*rows=\d{3,}", "TeacherStrategy 大量 merge 不得在 MainThread"),
)


def scan_workflow_log_for_violations(log_text: str) -> List[LogViolation]:
    """掃描既有 gtc_ai_trading log，找出三流程混用證據。

    使用方式：
        text = Path("logs/gtc_ai_trading_yyyymmdd.log").read_text(encoding="utf-8", errors="ignore")
        violations = scan_workflow_log_for_violations(text)
    """
    violations: List[LogViolation] = []
    current_workflow = "UNKNOWN"

    for i, line in enumerate(str(log_text or "").splitlines(), start=1):
        # 粗略判斷目前流程區段
        if "每日增量更新" in line or "update_daily" in line:
            current_workflow = WorkflowMode.DAILY_UPDATE_MASTER.value
        elif "重建排行" in line or "重排行進度" in line or "ranking_rebuild" in line:
            current_workflow = WorkflowMode.FAST_RANK_REBUILD.value
        elif "show_top20" in line or "AI選股TOP20" in line or "TOP20" in line:
            current_workflow = WorkflowMode.AI_TOP20_VIEW.value

        if current_workflow == WorkflowMode.FAST_RANK_REBUILD.value:
            for pattern, reco in FAST_RANK_FORBIDDEN_PATTERNS:
                if re.search(pattern, line, flags=re.IGNORECASE):
                    violations.append(LogViolation(
                        i, "P0", current_workflow, "FAST_RANK_FORBIDDEN_WORK", line.strip(), reco
                    ))

        if current_workflow == WorkflowMode.AI_TOP20_VIEW.value:
            for pattern, reco in AI_TOP20_FORBIDDEN_PATTERNS:
                if re.search(pattern, line, flags=re.IGNORECASE):
                    violations.append(LogViolation(
                        i, "P0", current_workflow, "AI_TOP20_FORBIDDEN_WORK", line.strip(), reco
                    ))

        for pattern, reco in MAINTHREAD_HEAVY_PATTERNS:
            if re.search(pattern, line, flags=re.IGNORECASE):
                violations.append(LogViolation(
                    i, "P1", current_workflow, "MAINTHREAD_HEAVY_WORK", line.strip(), reco
                ))

    return violations


def violations_to_rows(violations: Iterable[LogViolation]) -> List[Dict[str, Any]]:
    return [
        {
            "line_no": v.line_no,
            "priority": v.level,
            "workflow": v.workflow,
            "violation_type": v.violation_type,
            "evidence": v.evidence,
            "recommendation": v.recommendation,
        }
        for v in violations
    ]


def build_permission_matrix_rows() -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for op, rule in PERMISSION_MATRIX.items():
        rows.append({
            "operation": op.value,
            "daily_update_master": "ALLOW" if rule.daily_update_master else "DENY",
            "fast_rank_rebuild": "ALLOW" if rule.fast_rank_rebuild else "DENY",
            "ai_top20_view": "ALLOW" if rule.ai_top20_view else "DENY",
            "reason": rule.reason,
        })
    return rows


def require_build_tag(expected_tag: str, actual_tag: str) -> None:
    """用於確認 EXE 是否真的打包到最新主程式。

    主程式啟動時建議：
        APP_BUILD_TAG = "WORKFLOW_SPLIT_V17_YYYYMMDD"
        require_build_tag("WORKFLOW_SPLIT_V17_YYYYMMDD", APP_BUILD_TAG)
        log_info(f"[BUILD_TAG] {APP_BUILD_TAG}")
    """
    if str(expected_tag).strip() != str(actual_tag).strip():
        raise WorkflowViolation(f"BUILD_TAG 不一致：expected={expected_tag}, actual={actual_tag}")


__all__ = [
    "WorkflowMode",
    "Operation",
    "WorkflowViolation",
    "WorkflowContext",
    "WorkflowEvent",
    "WorkflowLogger",
    "WorkflowGuard",
    "GTCWorkflowOrchestrator",
    "scan_workflow_log_for_violations",
    "violations_to_rows",
    "build_permission_matrix_rows",
    "require_build_tag",
]
