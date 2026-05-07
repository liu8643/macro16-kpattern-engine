import pandas as pd


class PatternScoreEngine:
    """型態分數引擎。輸出 -100~100 的 kpattern_score 與交易偏向。"""

    BASE = {
        "GRAVESTONE": (-45, "AVOID", "高檔墓碑/長上影，追高風險"),
        "SHOOTING_STAR": (-35, "AVOID", "流星線，疑似上檔賣壓"),
        "THREE_BLACK_CROWS": (-50, "AVOID", "三烏鴉，短線轉弱"),
        "THREE_WHITE_SOLDIERS": (25, "BUY", "紅三兵，連續攻擊"),
        "HAMMER": (20, "WAIT", "錘子線，低接觀察"),
        "INVERTED_HAMMER": (10, "WATCH", "倒錘線，需後續確認"),
        "BULLISH_ENGULFING": (28, "BUY", "多方吞噬，轉強訊號"),
        "BEARISH_ENGULFING": (-35, "AVOID", "空方吞噬，轉弱訊號"),
        "NEUTRAL": (0, "NEUTRAL", "無明確型態"),
    }

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        x["kpattern_score"] = 0.0
        x["kpattern_trade_bias"] = "NEUTRAL"
        x["kpattern_risk_flag"] = "NO"
        x["kpattern_reason"] = "無明確型態"
        x["kpattern_trade_advice"] = "觀察"

        for code, (score, bias, reason) in self.BASE.items():
            mask = x["kpattern_code"].eq(code)
            x.loc[mask, "kpattern_score"] = float(score)
            x.loc[mask, "kpattern_trade_bias"] = bias
            x.loc[mask, "kpattern_reason"] = reason

        high = x["kpattern_position"].eq("高位階")
        low = x["kpattern_position"].eq("低位階")
        volume_high = x["kpattern_volume_confirm"].isin(["放量", "爆量"])

        risk_codes = x["kpattern_code"].isin(["GRAVESTONE", "SHOOTING_STAR", "THREE_BLACK_CROWS", "BEARISH_ENGULFING"])
        x.loc[risk_codes & high, "kpattern_score"] -= 12
        x.loc[risk_codes & high & volume_high, "kpattern_score"] -= 8
        x.loc[risk_codes & (high | volume_high), "kpattern_risk_flag"] = "HIGH"

        buy_codes = x["kpattern_code"].isin(["THREE_WHITE_SOLDIERS", "HAMMER", "BULLISH_ENGULFING"])
        x.loc[buy_codes & low, "kpattern_score"] += 8
        x.loc[buy_codes & volume_high, "kpattern_score"] += 4
        x.loc[buy_codes & high, "kpattern_score"] -= 10
        x.loc[buy_codes & high, "kpattern_reason"] = x.loc[buy_codes & high, "kpattern_reason"].astype(str) + "；高位階防追高"

        x["kpattern_score"] = x["kpattern_score"].clip(-100, 100).round(2)
        x.loc[x["kpattern_trade_bias"].eq("AVOID"), "kpattern_trade_advice"] = "強制排除/禁追高"
        x.loc[x["kpattern_trade_bias"].eq("BUY"), "kpattern_trade_advice"] = "可列入候選，仍需RR與價格區間確認"
        x.loc[x["kpattern_trade_bias"].isin(["WATCH", "WAIT"]), "kpattern_trade_advice"] = "觀察/等待低接"
        return x
