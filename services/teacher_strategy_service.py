try:
    import pandas as pd
except Exception:
    pd = None

try:
    from teacher_strategy.teacher_strategy_engine import TeacherStrategyEngine
except Exception:
    TeacherStrategyEngine = None


class TeacherStrategyService:
    """
    TeacherStrategyService V2

    功能：
    1. 服務層，不寫老師策略邏輯
    2. 批次執行 TeacherStrategyEngine.analyze_row()
    3. 輸出 teacher_df 給主程式 merge 回 ranking_df / trade_plan
    4. 保證欄位固定，避免 UI / Excel / DB 缺欄位崩潰
    """

    TEACHER_COLUMNS = [
        "stock_id",
        "stock_name",
        "teacher_strategy_class",
        "teacher_final_decision",
        "teacher_light",
        "teacher_gate",
        "position_stage",
        "position_score",
        "two_high_fail",
        "weak_gate",
        "weak_score",
        "rotation",
        "sector_strength_score",
        "flow_score",
        "rs_score",
        "teacher_buy_zone",
        "teacher_stop_loss",
        "teacher_target_price",
        "teacher_reason",
        "teacher_source",
    ]

    DEFAULT_ROW = {
        "stock_id": "",
        "stock_name": "",
        "teacher_strategy_class": "未評估",
        "teacher_final_decision": "WATCH",
        "teacher_light": "⚫",
        "teacher_gate": "NE",
        "position_stage": "未知",
        "position_score": 0.0,
        "two_high_fail": False,
        "weak_gate": "NE",
        "weak_score": 0.0,
        "rotation": "未知",
        "sector_strength_score": 0.0,
        "flow_score": 0.0,
        "rs_score": 0.0,
        "teacher_buy_zone": "",
        "teacher_stop_loss": "",
        "teacher_target_price": "",
        "teacher_reason": "TeacherStrategyService 未完成評估或資料不足",
        "teacher_source": "teacher_strategy_service_v2",
    }

    def __init__(self):
        self.engine = None
        if TeacherStrategyEngine:
            self.engine = TeacherStrategyEngine()

    def _default_result(self, row=None):
        out = dict(self.DEFAULT_ROW)

        if row is not None:
            try:
                out["stock_id"] = str(row.get("stock_id", ""))
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
                out["stock_id"] = str(row.get("stock_id", ""))
            except Exception:
                pass

        if not out.get("stock_name"):
            try:
                out["stock_name"] = str(row.get("stock_name", row.get("名稱", "")))
            except Exception:
                pass

        return out

    def run(
        self,
        ranking_df=None,
        price_history_df=None,
        market_df=None,
        institutional_df=None,
        **kwargs
    ):
        """
        主程式呼叫入口

        Parameters:
        ranking_df:
            RankingEngine.rebuild() 或 trade_plan 前的股票清單

        price_history_df:
            保留給後續版本使用，目前不在 Service 層計算策略

        market_df / institutional_df:
            保留給市場級 AI / 法人資金流升級
        """

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
            return pd.DataFrame(results, columns=self.TEACHER_COLUMNS)

        for _, row in ranking_df.iterrows():
            try:
                result = self.engine.analyze_row(row)
                results.append(self._normalize_result(result, row))
            except Exception as exc:
                fallback = self._default_result(row)
                fallback["teacher_reason"] = f"TeacherStrategy analyze_row failed: {exc}"
                fallback["teacher_source"] = "teacher_strategy_service_v2_fallback"
                results.append(fallback)

        out = pd.DataFrame(results)

        for col in self.TEACHER_COLUMNS:
            if col not in out.columns:
                out[col] = self.DEFAULT_ROW.get(col, "")

        out = out[self.TEACHER_COLUMNS].copy()

        # 型別安全處理，避免主程式 merge / fillna / Excel 輸出錯誤
        for col in [
            "position_score",
            "weak_score",
            "sector_strength_score",
            "flow_score",
            "rs_score",
        ]:
            try:
                out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
            except Exception:
                out[col] = 0.0

        try:
            out["two_high_fail"] = out["two_high_fail"].fillna(False).astype(bool)
        except Exception:
            out["two_high_fail"] = False

        for col in [
            "stock_id",
            "stock_name",
            "teacher_strategy_class",
            "teacher_final_decision",
            "teacher_light",
            "teacher_gate",
            "position_stage",
            "weak_gate",
            "rotation",
            "teacher_buy_zone",
            "teacher_stop_loss",
            "teacher_target_price",
            "teacher_reason",
            "teacher_source",
        ]:
            try:
                out[col] = out[col].fillna("").astype(str)
            except Exception:
                out[col] = ""

        return out
