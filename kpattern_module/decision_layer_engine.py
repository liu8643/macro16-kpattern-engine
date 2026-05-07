import pandas as pd


class KPatternDecisionLayerEngine:
    """KPattern 對主系統的穩定輸出層。不得直接覆蓋 trade_allowed。"""

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        score = pd.to_numeric(x.get("kpattern_score"), errors="coerce").fillna(0.0)
        vol = pd.to_numeric(x.get("kpattern_volume_score"), errors="coerce").fillna(0.0)
        risk_penalty = x.get("kpattern_risk_flag", "NO").astype(str).str.upper().map({"HIGH": 20.0, "MID": 8.0}).fillna(0.0)
        # final_trade_score 是型態子系統分數，不取代主程式 total_score / model_score。
        x["final_trade_score"] = (50.0 + score + vol - risk_penalty).clip(0, 100).round(2)
        # 主程式只讀微調值；控制在 -8 ~ +6，避免型態分數過度影響排名。
        x["kpattern_rank_adjustment"] = (score * 0.10 + vol * 0.25 - risk_penalty * 0.10).clip(-8, 6).round(2)
        x["kpattern_signal"] = "NEUTRAL"
        x.loc[x["kpattern_trade_bias"].isin(["BUY"]), "kpattern_signal"] = "BUY"
        x.loc[x["kpattern_trade_bias"].isin(["WATCH", "WAIT"]), "kpattern_signal"] = "WATCH"
        x.loc[x["kpattern_trade_bias"].eq("AVOID"), "kpattern_signal"] = "AVOID"
        return x
