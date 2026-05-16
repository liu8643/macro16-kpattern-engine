# V9_WORKFLOW_COMPAT: service must not trigger workflow update/rebuild by itself.
try:
    import pandas as pd
except Exception:
    pd = None

try:
    from teacher_strategy.teacher_strategy_engine import TeacherStrategyEngine
except Exception:
    try:
        from teacher_strategy_engine import TeacherStrategyEngine
    except Exception:
        TeacherStrategyEngine = None


class TeacherStrategyService:
    """TeacherStrategyService V17-R3。

    重點：
    1. 批次執行 TeacherStrategyEngine.analyze_row()。
    2. 補 teacher_score / teacher_priority，讓老師策略能獨立排序，不再等同 TOP20。
    3. 保留固定欄位，避免 UI / Excel / DB 因缺欄位中斷。
    """

    TEACHER_COLUMNS = [
        "stock_id", "stock_name",
        "teacher_strategy_class", "teacher_final_decision", "teacher_light", "teacher_gate",
        "teacher_trade_allowed", "teacher_ui_bucket", "teacher_priority", "teacher_score",
        "teacher_no_trade_reason", "position_stage", "position_score",
        "low_base_type", "low_base_score", "two_high_fail", "weak_gate", "weak_score", "teacher_exclusion_level",
        "rotation", "sector_strength_score", "flow_score", "rs_score",
        "teacher_buy_zone", "teacher_stop_loss", "teacher_target_price", "teacher_reason", "teacher_source",
    ]

    DEFAULT_ROW = {
        "stock_id": "", "stock_name": "",
        "teacher_strategy_class": "未評估", "teacher_final_decision": "WATCH",
        "teacher_light": "⚫", "teacher_gate": "NE",
        "teacher_trade_allowed": 0, "teacher_ui_bucket": "未評估", "teacher_priority": 9999, "teacher_score": 0.0,
        "teacher_no_trade_reason": "TeacherStrategyService 未完成評估或資料不足",
        "position_stage": "未知", "position_score": 0.0,
        "low_base_type": "非低位階", "low_base_score": 0.0,
        "two_high_fail": False, "weak_gate": "NE", "weak_score": 0.0, "teacher_exclusion_level": "NE",
        "rotation": "未知", "sector_strength_score": 0.0,
        "flow_score": 0.0, "rs_score": 0.0,
        "teacher_buy_zone": "", "teacher_stop_loss": "", "teacher_target_price": "",
        "teacher_reason": "TeacherStrategyService 未完成評估或資料不足",
        "teacher_source": "teacher_strategy_service_v17_r3",
    }

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
                out[k] = v
        return out

    def _finalize(self, out):
        if pd is None:
            return out
        if out is None or out.empty:
            return pd.DataFrame(columns=self.TEACHER_COLUMNS)

        for col, default in self.DEFAULT_ROW.items():
            if col not in out.columns:
                out[col] = default
            else:
                out[col] = out[col].fillna(default)

        def _num(col, default=0.0):
            return pd.to_numeric(out[col], errors="coerce").fillna(default) if col in out.columns else pd.Series(default, index=out.index)

        def _txt(col, default=""):
            return out[col].fillna(default).astype(str) if col in out.columns else pd.Series(default, index=out.index, dtype="object")

        decision = _txt("teacher_final_decision", "WATCH").str.upper().str.replace(" ", "_", regex=False)
        weak_gate = _txt("weak_gate", "NE").str.upper()
        teacher_gate = _txt("teacher_gate", "NE").str.upper()
        weak_score = _num("weak_score", 0.0)
        position_score = _num("position_score", 0.0)
        teacher_score = _num("teacher_score", 0.0)
        rr = _num("rr_live", 0.0) if "rr_live" in out.columns else _num("rr", 0.0)
        rsi = _num("rsi14", 50.0) if "rsi14" in out.columns else _num("rsi", 50.0)
        price_dev = _num("price_deviation", 0.0) if "price_deviation" in out.columns else _num("price_dev", 0.0)
        two_high = _txt("two_high_fail", "False").str.upper().isin(["TRUE", "1", "YES", "Y", "是"])

        buy_decision = decision.isin(["BUY", "LOW_BUY"])
        wait_decision = decision.str.contains("WAIT|WATCH|PULLBACK", case=False, na=False)
        block_decision = decision.str.contains("AVOID|REDUCE|BLOCK|SELL|排除|減碼", case=False, na=False)

        hard_block = (
            weak_gate.eq("BLOCK") |
            teacher_gate.eq("BLOCK") |
            two_high |
            block_decision |
            (weak_score >= 70) |
            (rsi >= 78) |
            (price_dev >= 0.20)
        )
        soft_block = (
            weak_gate.eq("WARNING") |
            (position_score < 60) |
            ((rr > 0) & (rr < 1.0))
        )
        allowed = buy_decision & (~hard_block) & (~soft_block)
        out["teacher_trade_allowed"] = allowed.astype(int)

        bucket = pd.Series("觀察", index=out.index, dtype="object")
        bucket.loc[allowed & decision.eq("BUY")] = "今日可買"
        bucket.loc[allowed & decision.eq("LOW_BUY")] = "低位階翻多"
        bucket.loc[(~allowed) & wait_decision & (~hard_block)] = "等拉回"
        bucket.loc[hard_block] = "排除"
        out["teacher_ui_bucket"] = bucket

        priority_map = {"今日可買": 1, "低位階翻多": 2, "等拉回": 3, "觀察": 4, "排除": 9, "未評估": 99}
        out["teacher_priority"] = out["teacher_ui_bucket"].map(priority_map).fillna(99).astype(int)
        out["teacher_score"] = teacher_score.clip(lower=0, upper=100).round(2)

        reasons = []
        for i in out.index:
            rs = []
            if bool(hard_block.loc[i]):
                if str(weak_gate.loc[i]) == "BLOCK": rs.append("weak_gate=BLOCK")
                if bool(two_high.loc[i]): rs.append("兩高不過")
                if float(weak_score.loc[i]) >= 70: rs.append("weak_score>=70")
                if float(rsi.loc[i]) >= 78: rs.append("RSI過熱")
                if float(price_dev.loc[i]) >= 0.20: rs.append("乖離過大")
                if bool(block_decision.loc[i]): rs.append(f"teacher_decision={decision.loc[i]}")
            elif bool(soft_block.loc[i]):
                if str(weak_gate.loc[i]) == "WARNING": rs.append("weak_gate=WARNING")
                if float(position_score.loc[i]) < 60: rs.append("position_score<60")
                if float(rr.loc[i]) > 0 and float(rr.loc[i]) < 1.0: rs.append("RR<1")
            elif bool(allowed.loc[i]):
                rs.append("Teacher Gate PASS")
            else:
                rs.append("未達BUY/LOW_BUY條件")
            reasons.append("；".join(rs))
        out["teacher_no_trade_reason"] = reasons

        # 排序：老師策略使用 teacher_priority + teacher_score，不再維持 TOP20 原排序。
        sort_cols = [c for c in ["teacher_priority", "teacher_score", "rank_all"] if c in out.columns]
        if sort_cols:
            ascending = [True if c == "teacher_priority" else False if c == "teacher_score" else True for c in sort_cols]
            out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

        for col in self.TEACHER_COLUMNS:
            if col not in out.columns:
                out[col] = self.DEFAULT_ROW.get(col, "")
        return out.reset_index(drop=True)

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
                fallback = self._default_result(row)
                fallback["teacher_reason"] = "TeacherStrategyEngine 匯入失敗，請檢查 teacher_strategy package"
                fallback["teacher_source"] = "teacher_strategy_service_v17_r3_engine_unavailable"
                results.append(fallback)
        else:
            for _, row in ranking_df.iterrows():
                try:
                    results.append(self._normalize_result(self.engine.analyze_row(row), row))
                except Exception as exc:
                    fallback = self._default_result(row)
                    fallback["teacher_reason"] = f"TeacherStrategy analyze_row failed: {exc}"
                    fallback["teacher_source"] = "teacher_strategy_service_v17_r3_fallback"
                    results.append(fallback)

        out = pd.DataFrame(results)
        # 把原 ranking 欄位帶回，讓 UI/Excel 可顯示代號、名稱、rank_all、rr 等欄位。
        try:
            base = ranking_df.copy().reset_index(drop=True)
            if "stock_id" in base.columns and "stock_id" in out.columns:
                extra_cols = [c for c in base.columns if c not in out.columns or c in ["rr", "rr_live", "rank_all", "total_score"]]
                out = out.merge(base[extra_cols + ["stock_id"]].drop_duplicates("stock_id"), on="stock_id", how="left")
        except Exception:
            pass
        return self._finalize(out)
