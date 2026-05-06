import pandas as pd


class PatternScoreEngine:
    """
    K線型態分數引擎

    功能：
    1. 將 pattern 轉成 pattern_score
    2. 產生 pattern_direction
    3. 產生 pattern_reason
    """

    SCORE_MAP = {
        "GRAVESTONE": (-30, "AVOID", "墓碑線：長上影反轉風險"),
        "SHOOTING_STAR": (-25, "AVOID", "流星線：高檔賣壓風險"),
        "THREE_BLACK_CROWS": (-40, "AVOID", "三烏鴉：短線轉弱風險"),
        "HAMMER": (25, "WATCH", "錘子線：低檔止跌觀察"),
        "INVERTED_HAMMER": (15, "WATCH", "倒錘線：低檔試攻觀察"),
        "BULLISH_ENGULFING": (30, "BUY", "多方吞噬：轉強訊號"),
        "BEARISH_ENGULFING": (-30, "AVOID", "空方吞噬：轉弱訊號"),
    }

    PRIORITY = {
        "AVOID": 5,
        "ATTACK": 4,
        "BUY": 3,
        "WAIT": 2,
        "WATCH": 1,
        "NEUTRAL": 0,
    }

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        df["pattern_score"] = 0
        df["pattern_direction"] = "NEUTRAL"
        df["pattern_reason"] = "無明確型態"

        for pattern, (score, direction, reason) in self.SCORE_MAP.items():
            cond = df["pattern"] == pattern

            df.loc[cond, "pattern_score"] = score
            df.loc[cond, "pattern_direction"] = direction
            df.loc[cond, "pattern_reason"] = reason

        df["pattern_priority"] = df["pattern_direction"].map(self.PRIORITY).fillna(0).astype(int)

        return df
