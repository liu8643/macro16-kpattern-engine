import pandas as pd


class SakataPatternEngine:
    """
    阪田型態判斷（Phase 1）
    目前包含：
    1. 墓碑線
    2. 流星線
    3. 三烏鴉
    4. 錘子線
    5. 倒錘線
    6. 多方吞噬
    7. 空方吞噬
    """

    def classify(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        # 預設
        df["pattern"] = None

        # === 墓碑線（Gravestone）===
        cond_gravestone = (
            (df["body_ratio"] <= 0.15) &
            (df["upper_shadow_ratio"] >= 0.65) &
            (df["close_position"] <= 0.25)
        )
        df.loc[cond_gravestone, "pattern"] = "GRAVESTONE"

        # === 流星線（Shooting Star）===
        cond_shooting = (
            (df["upper_shadow"] >= df["body"] * 2) &
            (df["lower_shadow_ratio"] <= 0.2) &
            (df["body_ratio"] <= 0.4)
        )
        df.loc[cond_shooting, "pattern"] = "SHOOTING_STAR"

        # === 三烏鴉（Three Black Crows）===
        cond_three_black = (
            (df["is_bear"]) &
            (df["close"] < df["close"].shift(1)) &
            (df["close"].shift(1) < df["close"].shift(2)) &
            (df["is_bear"].shift(1)) &
            (df["is_bear"].shift(2))
        )
        df.loc[cond_three_black, "pattern"] = "THREE_BLACK_CROWS"

        # === 錘子線（Hammer）===
        cond_hammer = (
            (df["lower_shadow"] >= df["body"] * 2) &
            (df["upper_shadow_ratio"] <= 0.2) &
            (df["body_ratio"] <= 0.4)
        )
        df.loc[cond_hammer, "pattern"] = "HAMMER"

        # === 倒錘線（Inverted Hammer）===
        cond_inverted_hammer = (
            (df["upper_shadow"] >= df["body"] * 2) &
            (df["lower_shadow_ratio"] <= 0.2) &
            (df["body_ratio"] <= 0.4)
        )
        df.loc[cond_inverted_hammer, "pattern"] = "INVERTED_HAMMER"

        # === 多方吞噬（Bullish Engulfing）===
        cond_bullish_engulfing = (
            (df["is_bull"]) &
            (df["is_bear"].shift(1)) &
            (df["close"] >= df["open"].shift(1)) &
            (df["open"] <= df["close"].shift(1))
        )
        df.loc[cond_bullish_engulfing, "pattern"] = "BULLISH_ENGULFING"

        # === 空方吞噬（Bearish Engulfing）===
        cond_bearish_engulfing = (
            (df["is_bear"]) &
            (df["is_bull"].shift(1)) &
            (df["close"] <= df["open"].shift(1)) &
            (df["open"] >= df["close"].shift(1))
        )
        df.loc[cond_bearish_engulfing, "pattern"] = "BEARISH_ENGULFING"

        return df
