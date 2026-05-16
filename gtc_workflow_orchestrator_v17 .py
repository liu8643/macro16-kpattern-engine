# -*- coding: utf-8 -*-
"""
GTC Workflow Orchestrator V17.1
四流程總閘：初始化全市場 / 建立完整歷史 / 每日增量更新 / 重建排行

此模組只做流程管制與驗收，不實作股票下載或排名邏輯。
目的：防止不同按鈕之間重覆執行 full reload / full merge / full rebuild。
"""
from __future__ import annotations
from dataclasses import dataclass, field
from enum import Enum
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
import logging, re, threading, time, uuid

class WorkflowMode(str, Enum):
    INIT_MARKET = "INIT_MARKET"
    BUILD_FULL_HISTORY = "BUILD_FULL_HISTORY"
    DAILY_UPDATE_MASTER = "Daily_Update_Master"
    FAST_RANK_REBUILD = "FAST_RANK_REBUILD"
    AI_TOP20_VIEW = "AI_TOP20_VIEW"

class Operation(str, Enum):
    MASTER_UNIVERSE_FETCH = "master_universe_fetch"
    MASTER_UNIVERSE_WRITE = "master_universe_write"
    CLASSIFICATION_QA = "classification_qa"
    FULL_HISTORY_BUILD = "full_history_build"
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
    pass

@dataclass(frozen=True)
class PermissionRule:
    operation: Operation
    init_market: bool
    build_full_history: bool
    daily_update_master: bool
    fast_rank_rebuild: bool
    ai_top20_view: bool
    reason: str

PERMISSION_MATRIX: Dict[Operation, PermissionRule] = {
    Operation.MASTER_UNIVERSE_FETCH: PermissionRule(Operation.MASTER_UNIVERSE_FETCH, True, False, False, False, False, "只有初始化全市場可抓主檔。"),
    Operation.MASTER_UNIVERSE_WRITE: PermissionRule(Operation.MASTER_UNIVERSE_WRITE, True, False, False, False, False, "只有初始化全市場可覆寫 stocks_master。"),
    Operation.CLASSIFICATION_QA: PermissionRule(Operation.CLASSIFICATION_QA, True, False, True, True, True, "分類 QA 可讀取，但不可在重排行時重新抓主檔。"),
    Operation.FULL_HISTORY_BUILD: PermissionRule(Operation.FULL_HISTORY_BUILD, False, True, False, False, False, "完整歷史建庫只寫 price_history 歷史資料，不做排名或基本面。"),
    Operation.EXTERNAL_API_FETCH: PermissionRule(Operation.EXTERNAL_API_FETCH, True, True, True, False, False, "外部 API 僅允許初始化、建庫、每日增量；重排行/TOP20 禁止。"),
    Operation.PRICE_HISTORY_WRITE: PermissionRule(Operation.PRICE_HISTORY_WRITE, False, True, True, False, False, "price_history 只允許完整建庫或每日增量寫入。"),
    Operation.EXTERNAL_TABLE_WRITE: PermissionRule(Operation.EXTERNAL_TABLE_WRITE, False, False, True, False, False, "external_* 只能每日增量更新寫入。"),
    Operation.FINANCIAL_FEATURE_WRITE: PermissionRule(Operation.FINANCIAL_FEATURE_WRITE, False, False, True, False, False, "financial_feature_daily 只能每日增量建立。"),
    Operation.EPS_MATRIX_BUILD: PermissionRule(Operation.EPS_MATRIX_BUILD, False, False, True, False, False, "EPS Matrix build 只能每日增量。"),
    Operation.EPS_MATRIX_DECISION_BULK: PermissionRule(Operation.EPS_MATRIX_DECISION_BULK, False, False, True, False, False, "EPS Matrix bulk decision 只能每日增量。"),
    Operation.TECHNICAL_SCORE_RECALC: PermissionRule(Operation.TECHNICAL_SCORE_RECALC, False, False, True, True, False, "技術分數只允許每日增量後或快速重排行。"),
    Operation.RANKING_RESULT_WRITE: PermissionRule(Operation.RANKING_RESULT_WRITE, False, False, True, True, False, "ranking_result 只允許每日增量或快速重排行寫入。"),
    Operation.TEACHER_FULL_MERGE: PermissionRule(Operation.TEACHER_FULL_MERGE, False, False, True, False, False, "Teacher full merge 只能每日增量或指定批次。"),
    Operation.TEACHER_LIGHT_MERGE: PermissionRule(Operation.TEACHER_LIGHT_MERGE, False, False, True, True, False, "快速重排行只允許 teacher snapshot/light merge。"),
    Operation.TEACHER_SNAPSHOT_WRITE: PermissionRule(Operation.TEACHER_SNAPSHOT_WRITE, False, False, True, False, False, "Teacher snapshot 寫入只屬每日批次。"),
    Operation.SNAPSHOT_READ: PermissionRule(Operation.SNAPSHOT_READ, True, True, True, True, True, "五種流程都可讀快照。"),
    Operation.UI_RENDER: PermissionRule(Operation.UI_RENDER, True, True, True, True, True, "UI 僅允許顯示，不可觸發重工作。"),
    Operation.EXPORT_EXCEL: PermissionRule(Operation.EXPORT_EXCEL, False, False, True, True, True, "匯出只能讀既有結果。"),
    Operation.HEAVY_MAINTHREAD_WORK: PermissionRule(Operation.HEAVY_MAINTHREAD_WORK, False, False, False, False, False, "任何流程都不得在 MainThread 做重工作。"),
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
    metadata: Dict[str, Any] = field(default_factory=dict)

class WorkflowLogger:
    def __init__(self, logger: Optional[logging.Logger]=None, log_cb: Optional[Callable[[str], None]]=None):
        self.logger = logger or logging.getLogger("gtc.workflow")
        self.log_cb = log_cb
        self.events: List[WorkflowEvent] = []
    def emit(self, event: WorkflowEvent) -> None:
        self.events.append(event)
        text = f"[WORKFLOW][{event.mode.value}][{event.status}] run_id={event.run_id} stage={event.stage} rows={event.rows} duration={event.duration_sec:.3f}s thread={event.thread_name} {event.message}"
        try: self.logger.info(text)
        except Exception: pass
        if self.log_cb:
            try: self.log_cb(text)
            except Exception: pass
    def start(self, ctx: WorkflowContext, stage: str, message: str="") -> float:
        self.emit(WorkflowEvent(ctx.run_id, ctx.mode, stage, "START", message=message)); return time.perf_counter()
    def end(self, ctx: WorkflowContext, stage: str, start_time: float, rows: int=0, message: str="") -> None:
        self.emit(WorkflowEvent(ctx.run_id, ctx.mode, stage, "END", rows=int(rows or 0), duration_sec=max(time.perf_counter()-start_time,0), message=message))
    def violation(self, ctx: WorkflowContext, stage: str, message: str="") -> None:
        self.emit(WorkflowEvent(ctx.run_id, ctx.mode, stage, "VIOLATION", message=message))

class WorkflowGuard:
    def __init__(self, logger: Optional[WorkflowLogger]=None):
        self.workflow_logger = logger or WorkflowLogger()
    @staticmethod
    def _allowed(mode: WorkflowMode, operation: Operation) -> bool:
        rule = PERMISSION_MATRIX[operation]
        return {
            WorkflowMode.INIT_MARKET: rule.init_market,
            WorkflowMode.BUILD_FULL_HISTORY: rule.build_full_history,
            WorkflowMode.DAILY_UPDATE_MASTER: rule.daily_update_master,
            WorkflowMode.FAST_RANK_REBUILD: rule.fast_rank_rebuild,
            WorkflowMode.AI_TOP20_VIEW: rule.ai_top20_view,
        }.get(mode, False)
    def assert_allowed(self, ctx: WorkflowContext, operation: Operation, stage: str="") -> None:
        if operation not in PERMISSION_MATRIX:
            raise WorkflowViolation(f"未知工作項目：{operation}")
        if not self._allowed(ctx.mode, operation):
            rule = PERMISSION_MATRIX[operation]
            msg = f"禁止在 {ctx.mode.value} 執行 {operation.value}：{rule.reason}"
            self.workflow_logger.violation(ctx, stage or operation.value, msg)
            if ctx.strict: raise WorkflowViolation(msg)
    def assert_not_heavy_mainthread(self, ctx: WorkflowContext, stage: str) -> None:
        if threading.current_thread().name == "MainThread":
            msg = f"重型工作不得在 MainThread 執行：mode={ctx.mode.value}, stage={stage}"
            self.workflow_logger.violation(ctx, stage, msg)
            if ctx.strict: raise WorkflowViolation(msg)

def _safe_len(obj: Any) -> int:
    try: return int(len(obj))
    except Exception: return 0

class GTCWorkflowOrchestrator:
    def __init__(self, logger: Optional[logging.Logger]=None, log_cb: Optional[Callable[[str], None]]=None, strict: bool=True):
        self.workflow_logger = WorkflowLogger(logger=logger, log_cb=log_cb)
        self.guard = WorkflowGuard(self.workflow_logger)
        self.strict = strict
    def _ctx(self, mode: WorkflowMode, **metadata: Any) -> WorkflowContext:
        return WorkflowContext(mode=mode, run_id=f"{mode.value}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}", strict=self.strict, metadata=dict(metadata or {}))
    @contextmanager
    def stage(self, ctx: WorkflowContext, stage_name: str, operation: Optional[Operation]=None, heavy: bool=False, message: str=""):
        if operation is not None: self.guard.assert_allowed(ctx, operation, stage_name)
        if heavy: self.guard.assert_not_heavy_mainthread(ctx, stage_name)
        start_time = self.workflow_logger.start(ctx, stage_name, message=message)
        try:
            yield
        finally:
            self.workflow_logger.end(ctx, stage_name, start_time)
    def initialize_market(self, *, build_universe_fn: Callable[..., Any], import_master_fn: Callable[..., Any], classification_qa_fn: Optional[Callable[..., Any]]=None, refresh_ui_fn: Optional[Callable[..., Any]]=None, **kwargs: Any) -> Dict[str, Any]:
        ctx=self._ctx(WorkflowMode.INIT_MARKET); result={"run_id":ctx.run_id,"mode":ctx.mode.value}
        with self.stage(ctx,"01_build_market_universe",Operation.MASTER_UNIVERSE_FETCH,heavy=True): result["universe"]=build_universe_fn(**kwargs)
        with self.stage(ctx,"02_import_master",Operation.MASTER_UNIVERSE_WRITE,heavy=True): result["import_master"]=import_master_fn(result.get("universe"),**kwargs)
        if classification_qa_fn:
            with self.stage(ctx,"03_classification_qa",Operation.CLASSIFICATION_QA,heavy=False): result["classification_qa"]=classification_qa_fn(**kwargs)
        if refresh_ui_fn:
            with self.stage(ctx,"04_refresh_ui_master_only",Operation.UI_RENDER,heavy=False): result["ui"]=refresh_ui_fn(**kwargs)
        return result
    def build_full_history(self, *, build_history_fn: Callable[..., Any], refresh_ui_fn: Optional[Callable[..., Any]]=None, **kwargs: Any) -> Dict[str, Any]:
        ctx=self._ctx(WorkflowMode.BUILD_FULL_HISTORY); result={"run_id":ctx.run_id,"mode":ctx.mode.value}
        with self.stage(ctx,"01_build_full_price_history",Operation.FULL_HISTORY_BUILD,heavy=True): result["history"]=build_history_fn(**kwargs)
        if refresh_ui_fn:
            with self.stage(ctx,"02_refresh_ui_history_status_only",Operation.UI_RENDER,heavy=False): result["ui"]=refresh_ui_fn(**kwargs)
        return result
    def daily_update_master(self, *, update_price_history_fn: Callable[..., Any], sync_external_tables_fn: Optional[Callable[..., Any]]=None, build_financial_feature_daily_fn: Optional[Callable[..., Any]]=None, eps_matrix_decision_fn: Optional[Callable[..., Any]]=None, rebuild_ranking_from_cache_fn: Optional[Callable[..., Any]]=None, build_teacher_snapshot_fn: Optional[Callable[..., Any]]=None, refresh_ui_fn: Optional[Callable[..., Any]]=None, **kwargs: Any) -> Dict[str, Any]:
        ctx=self._ctx(WorkflowMode.DAILY_UPDATE_MASTER); result={"run_id":ctx.run_id,"mode":ctx.mode.value}
        with self.stage(ctx,"01_update_price_history",Operation.PRICE_HISTORY_WRITE,heavy=True): result["price_history"]=update_price_history_fn(**kwargs)
        if sync_external_tables_fn:
            with self.stage(ctx,"02_sync_external_tables",Operation.EXTERNAL_TABLE_WRITE,heavy=True): result["external_tables"]=sync_external_tables_fn(**kwargs)
        if build_financial_feature_daily_fn:
            with self.stage(ctx,"03_build_financial_feature_daily",Operation.FINANCIAL_FEATURE_WRITE,heavy=True): result["financial_feature_daily"]=build_financial_feature_daily_fn(**kwargs)
        if eps_matrix_decision_fn:
            with self.stage(ctx,"04_eps_matrix_decision_bulk",Operation.EPS_MATRIX_DECISION_BULK,heavy=True): result["eps_matrix_decision"]=eps_matrix_decision_fn(**kwargs)
        if rebuild_ranking_from_cache_fn:
            with self.stage(ctx,"05_rebuild_ranking_result_after_update",Operation.RANKING_RESULT_WRITE,heavy=True): result["ranking_result"]=rebuild_ranking_from_cache_fn(rank_mode="daily_after_update", **kwargs)
        if build_teacher_snapshot_fn:
            with self.stage(ctx,"06_build_teacher_strategy_snapshot",Operation.TEACHER_SNAPSHOT_WRITE,heavy=True): result["teacher_snapshot"]=build_teacher_snapshot_fn(**kwargs)
        if refresh_ui_fn:
            with self.stage(ctx,"07_refresh_ui",Operation.UI_RENDER,heavy=False): result["ui"]=refresh_ui_fn(**kwargs)
        return result
    def fast_rank_rebuild(self, *, read_cache_fn: Callable[..., Any], rebuild_ranking_from_cache_fn: Callable[..., Any], teacher_light_merge_fn: Optional[Callable[..., Any]]=None, write_ranking_result_fn: Optional[Callable[..., Any]]=None, refresh_ui_fn: Optional[Callable[..., Any]]=None, **kwargs: Any) -> Dict[str, Any]:
        ctx=self._ctx(WorkflowMode.FAST_RANK_REBUILD); result={"run_id":ctx.run_id,"mode":ctx.mode.value}
        with self.stage(ctx,"01_read_cache",Operation.SNAPSHOT_READ,heavy=False): result["cache"]=read_cache_fn(**kwargs)
        with self.stage(ctx,"02_recalculate_technical_and_rank",Operation.TECHNICAL_SCORE_RECALC,heavy=True): result["ranking_df"]=rebuild_ranking_from_cache_fn(rank_mode="fast", allow_external_api=False, allow_eps_build=False, allow_financial_feature_write=False, **kwargs)
        if teacher_light_merge_fn:
            with self.stage(ctx,"03_teacher_light_merge",Operation.TEACHER_LIGHT_MERGE,heavy=True): result["ranking_df"]=teacher_light_merge_fn(result.get("ranking_df"), **kwargs)
        if write_ranking_result_fn:
            with self.stage(ctx,"04_write_ranking_result",Operation.RANKING_RESULT_WRITE,heavy=True): result["write_result"]=write_ranking_result_fn(result.get("ranking_df"), **kwargs)
        if refresh_ui_fn:
            with self.stage(ctx,"05_refresh_ui_rank_only",Operation.UI_RENDER,heavy=False): result["ui"]=refresh_ui_fn(result.get("ranking_df"), **kwargs)
        return result
    def ai_top20_view(self, *, read_top20_snapshot_fn: Callable[..., Any], render_ui_fn: Optional[Callable[..., Any]]=None, export_fn: Optional[Callable[..., Any]]=None, top_n: int=20, **kwargs: Any) -> Dict[str, Any]:
        ctx=self._ctx(WorkflowMode.AI_TOP20_VIEW, top_n=top_n); result={"run_id":ctx.run_id,"mode":ctx.mode.value,"top_n":top_n}
        with self.stage(ctx,"01_read_top20_snapshot",Operation.SNAPSHOT_READ,heavy=False): result["top20"]=read_top20_snapshot_fn(top_n=top_n, **kwargs)
        if render_ui_fn:
            with self.stage(ctx,"02_render_ui_top20",Operation.UI_RENDER,heavy=False): result["ui"]=render_ui_fn(result.get("top20"), **kwargs)
        if export_fn:
            with self.stage(ctx,"03_export_top20_snapshot",Operation.EXPORT_EXCEL,heavy=False): result["export"]=export_fn(result.get("top20"), **kwargs)
        return result

@dataclass
class LogViolation:
    line_no: int; level: str; workflow: str; violation_type: str; evidence: str; recommendation: str

FAST_RANK_FORBIDDEN_PATTERNS: Sequence[Tuple[str,str]] = (
    (r"\[EPS MATRIX\]\[BUILD\]", "FAST_RANK_REBUILD 不得執行 EPS MATRIX BUILD"),
    (r"\[EPS MATRIX\]\[DECISION\]", "FAST_RANK_REBUILD 不得逐檔 EPS MATRIX DECISION"),
    (r"financial_feature_daily.*rows|replace_financial_feature|financial_feature_batch", "FAST_RANK_REBUILD 不得寫 financial_feature_daily"),
    (r"official fetch|TWSE|TPEx|Yahoo|requests|read timeout", "FAST_RANK_REBUILD 不得抓外部資料"),
)
AI_TOP20_FORBIDDEN_PATTERNS: Sequence[Tuple[str,str]] = (
    (r"rebuild|重建排行|ranking_rebuild", "AI_TOP20_VIEW 不得觸發重建排行"),
    (r"\[EPS MATRIX\]\[BUILD\]|\[EPS MATRIX\]\[MERGE\]|\[EPS MATRIX\]\[DECISION\]", "AI_TOP20_VIEW 不得觸發 EPS Matrix"),
    (r"update_daily|每日資料更新|price_history|financial_feature_daily", "AI_TOP20_VIEW 不得觸發資料更新"),
)
MAINTHREAD_HEAVY_PATTERNS: Sequence[Tuple[str,str]] = (
    (r"MainThread.*\[EPS MATRIX\]\[BUILD\]", "EPS MATRIX BUILD 不得在 MainThread"),
    (r"MainThread.*重排行進度\s+\d+/\d+", "重排行 bulk loop 不得在 MainThread"),
)
def scan_workflow_log_for_violations(log_text: str) -> List[LogViolation]:
    violations: List[LogViolation] = []; current="UNKNOWN"
    for i,line in enumerate(str(log_text or "").splitlines(),1):
        if "初始化全市場" in line or "init_market" in line: current=WorkflowMode.INIT_MARKET.value
        elif "建立完整歷史" in line or "build_history" in line: current=WorkflowMode.BUILD_FULL_HISTORY.value
        elif "每日增量更新" in line or "update_daily" in line or "Daily_Update_Master" in line: current=WorkflowMode.DAILY_UPDATE_MASTER.value
        elif "重建排行" in line or "快速重算排行" in line or "ranking_rebuild" in line or "FAST_RANK_REBUILD" in line: current=WorkflowMode.FAST_RANK_REBUILD.value
        elif "show_top20" in line or "AI選股TOP20" in line or "TOP20" in line: current=WorkflowMode.AI_TOP20_VIEW.value
        if current == WorkflowMode.FAST_RANK_REBUILD.value:
            for pattern,reco in FAST_RANK_FORBIDDEN_PATTERNS:
                if re.search(pattern,line,re.I): violations.append(LogViolation(i,"P0",current,"FAST_RANK_FORBIDDEN_WORK",line.strip(),reco))
        if current == WorkflowMode.AI_TOP20_VIEW.value:
            for pattern,reco in AI_TOP20_FORBIDDEN_PATTERNS:
                if re.search(pattern,line,re.I): violations.append(LogViolation(i,"P0",current,"AI_TOP20_FORBIDDEN_WORK",line.strip(),reco))
        for pattern,reco in MAINTHREAD_HEAVY_PATTERNS:
            if re.search(pattern,line,re.I): violations.append(LogViolation(i,"P1",current,"MAINTHREAD_HEAVY_WORK",line.strip(),reco))
    return violations

def violations_to_rows(violations: Iterable[LogViolation]) -> List[Dict[str,Any]]:
    return [v.__dict__.copy() for v in violations]

def build_permission_matrix_rows() -> List[Dict[str,Any]]:
    return [{"operation":op.value,"init_market":"ALLOW" if r.init_market else "DENY","build_full_history":"ALLOW" if r.build_full_history else "DENY","daily_update_master":"ALLOW" if r.daily_update_master else "DENY","fast_rank_rebuild":"ALLOW" if r.fast_rank_rebuild else "DENY","ai_top20_view":"ALLOW" if r.ai_top20_view else "DENY","reason":r.reason} for op,r in PERMISSION_MATRIX.items()]

def require_build_tag(expected_tag: str, actual_tag: str) -> None:
    if str(expected_tag).strip()!=str(actual_tag).strip(): raise WorkflowViolation(f"BUILD_TAG 不一致：expected={expected_tag}, actual={actual_tag}")

__all__ = ["WorkflowMode","Operation","WorkflowViolation","WorkflowContext","WorkflowEvent","WorkflowLogger","WorkflowGuard","GTCWorkflowOrchestrator","scan_workflow_log_for_violations","violations_to_rows","build_permission_matrix_rows","require_build_tag"]
