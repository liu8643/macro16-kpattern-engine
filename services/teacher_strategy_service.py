try:
    import pandas as pd
except Exception:
    pd = None

try:
    from teacher_strategy.teacher_strategy_engine import TeacherStrategyEngine
except Exception:
    TeacherStrategyEngine = None


class TeacherStrategyService:
    """TeacherStrategyService V4.1 FINAL：批次執行 TeacherStrategyEngine.analyze_row()。

    服務層只負責批次、欄位標準化、型別安全；不寫老師策略邏輯。
    """

    TEACHER_COLUMNS = [
        "stock_id", "stock_name",
        "teacher_strategy_class", "teacher_final_decision", "teacher_light", "teacher_gate",
        "teacher_trade_allowed", "teacher_ui_bucket", "teacher_priority",
        "teacher_no_trade_reason", "teacher_block_reason",
        "position_stage", "position_score",
        "two_high_fail", "weak_gate", "weak_score",
        "rotation", "rotation_level", "sector_strength_score",
        "flow_score", "rs_score",
        "low_base_type", "low_base_score", "low_base_reason",
        "teacher_strategy_score", "teacher_rank", "teacher_rank_seed",
        "teacher_buy_zone", "teacher_stop_loss", "teacher_target_price",
        "teacher_reason", "teacher_source",
    ]

    DEFAULT_ROW = {
        "stock_id": "", "stock_name": "",
        "teacher_strategy_class": "未評估", "teacher_final_decision": "WATCH",
        "teacher_light": "⚫", "teacher_gate": "NE",
        "teacher_trade_allowed": 0, "teacher_ui_bucket": "未評估", "teacher_priority": 9999,
        "teacher_no_trade_reason": "TeacherStrategyService 尚未完成評估或資料不足",
        "teacher_block_reason": "",
        "position_stage": "未知", "position_score": 0.0,
        "two_high_fail": False, "weak_gate": "NE", "weak_score": 0.0,
        "rotation": "未知", "rotation_level": "", "sector_strength_score": 0.0,
        "flow_score": 0.0, "rs_score": 0.0,
        "low_base_type": "非低位階", "low_base_score": 0.0, "low_base_reason": "",
        "teacher_strategy_score": 0.0, "teacher_rank": 9999, "teacher_rank_seed": 9999,
        "teacher_buy_zone": "", "teacher_stop_loss": "", "teacher_target_price": "",
        "teacher_reason": "TeacherStrategyService 未完成評估或資料不足",
        "teacher_source": "teacher_strategy_service_v4_1_final",
    }

    NUMERIC_COLUMNS = [
        "position_score", "weak_score", "sector_strength_score", "flow_score", "rs_score",
        "low_base_score", "teacher_strategy_score", "teacher_rank", "teacher_rank_seed",
        "teacher_priority", "teacher_trade_allowed",
    ]

    TEXT_COLUMNS = [c for c in TEACHER_COLUMNS if c not in NUMERIC_COLUMNS and c != "two_high_fail"]

    def __init__(self):
        self.engine = TeacherStrategyEngine() if TeacherStrategyEngine else None

    def _default_result(self, row=None):
        out = dict(self.DEFAULT_ROW)
        if row is not None:
            try:
                out["stock_id"] = str(row.get("stock_id", row.get("代號", "")))
            except Exception:
                pass
            try:
                out["stock_name"] = str(row.get("stock_name", row.get("名稱", "")))
            except Exception:
                pass
        return out

    def _normalize_result(self, result, row=None):
        out = self._default_result(row)
        if isinstance(result, dict):
            for k, v in result.items():
                if k in out:
                    out[k] = v
        if not out.get("stock_id"):
            try:
                out["stock_id"] = str(row.get("stock_id", row.get("代號", "")))
            except Exception:
                pass
        if not out.get("stock_name"):
            try:
                out["stock_name"] = str(row.get("stock_name", row.get("名稱", "")))
            except Exception:
                pass
        return out

    def run(self, ranking_df=None, price_history_df=None, market_df=None, institutional_df=None, **kwargs):
        if pd is None:
            return None
        if ranking_df is None:
            return pd.DataFrame(columns=self.TEACHER_COLUMNS)
        try:
            if ranking_df.empty:
                return pd.DataFrame(columns=self.TEACHER_COLUMNS)
        except Exception:
            return pd.DataFrame(columns=self.TEACHER_COLUMNS)

        results = []
        if self.engine is None:
            for _, row in ranking_df.iterrows():
                results.append(self._default_result(row))
        else:
            for _, row in ranking_df.iterrows():
                try:
                    results.append(self._normalize_result(self.engine.analyze_row(row), row))
                except Exception as exc:
                    fallback = self._default_result(row)
                    fallback["teacher_reason"] = f"TeacherStrategy analyze_row failed: {exc}"
                    fallback["teacher_source"] = "teacher_strategy_service_v4_1_fallback"
                    results.append(fallback)

        out = pd.DataFrame(results)
        for col in self.TEACHER_COLUMNS:
            if col not in out.columns:
                out[col] = self.DEFAULT_ROW.get(col, "")
        out = out[self.TEACHER_COLUMNS].copy()

        for col in self.NUMERIC_COLUMNS:
            try:
                default = self.DEFAULT_ROW.get(col, 0.0)
                out[col] = pd.to_numeric(out[col], errors="coerce").fillna(default)
            except Exception:
                out[col] = self.DEFAULT_ROW.get(col, 0.0)

        try:
            out["two_high_fail"] = out["two_high_fail"].fillna(False).astype(str).str.upper().isin(["TRUE", "1", "YES", "Y", "是"])
        except Exception:
            out["two_high_fail"] = False

        for col in self.TEXT_COLUMNS:
            try:
                out[col] = out[col].fillna(self.DEFAULT_ROW.get(col, "")).astype(str)
            except Exception:
                out[col] = str(self.DEFAULT_ROW.get(col, ""))

        return out
