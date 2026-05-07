import pandas as pd


class SakataPatternEngine:
    """Phase 1 核心8型態辨識。只辨識型態，不決定是否下單。"""

    def classify(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        x["kpattern_code"] = "NEUTRAL"
        x["kpattern_name"] = "無明確型態"

        gravestone = (
            (x["k_body_ratio"] <= 0.15) &
            (x["k_upper_shadow_ratio"] >= 0.65) &
            (x["k_close_position"] <= 0.25)
        )
        x.loc[gravestone, ["kpattern_code", "kpattern_name"]] = ["GRAVESTONE", "墓碑線"]

        shooting_star = (
            (x["k_upper_shadow"] >= x["k_body"] * 2.0) &
            (x["k_lower_shadow_ratio"] <= 0.20) &
            (x["k_body_ratio"] <= 0.40)
        )
        x.loc[shooting_star, ["kpattern_code", "kpattern_name"]] = ["SHOOTING_STAR", "流星線"]

        three_black = (
            x["k_is_bear"].fillna(False) &
            x["k_is_bear"].shift(1).fillna(False) &
            x["k_is_bear"].shift(2).fillna(False) &
            (x["close"] < x["close"].shift(1)) &
            (x["close"].shift(1) < x["close"].shift(2)) &
            (x["low"] < x["low"].shift(1))
        )
        x.loc[three_black, ["kpattern_code", "kpattern_name"]] = ["THREE_BLACK_CROWS", "三烏鴉"]

        red_three = (
            x["k_is_bull"].fillna(False) &
            x["k_is_bull"].shift(1).fillna(False) &
            x["k_is_bull"].shift(2).fillna(False) &
            (x["close"] > x["close"].shift(1)) &
            (x["close"].shift(1) > x["close"].shift(2)) &
            (x["k_body_ratio"] >= 0.40)
        )
        x.loc[red_three, ["kpattern_code", "kpattern_name"]] = ["THREE_WHITE_SOLDIERS", "紅三兵"]

        hammer = (
            (x["k_lower_shadow"] >= x["k_body"] * 2.0) &
            (x["k_upper_shadow_ratio"] <= 0.30) &
            (x["k_body_ratio"] <= 0.45) &
            (x["k_close_position"] >= 0.50)
        )
        x.loc[hammer, ["kpattern_code", "kpattern_name"]] = ["HAMMER", "錘子線"]

        inverted_hammer = (
            (x["k_upper_shadow"] >= x["k_body"] * 2.0) &
            (x["k_lower_shadow_ratio"] <= 0.20) &
            (x["k_body_ratio"] <= 0.45) &
            (x["k_close_position"] >= 0.35)
        )
        # 流星線與倒錘型態相近，真正有效性由位置模組決定；避免覆蓋墓碑線。
        x.loc[inverted_hammer & x["kpattern_code"].eq("NEUTRAL"), ["kpattern_code", "kpattern_name"]] = ["INVERTED_HAMMER", "倒錘線"]

        bullish_engulfing = (
            x["k_is_bull"].fillna(False) &
            (x["k_prev_close"] < x["k_prev_open"]) &
            (x["close"] >= x["k_prev_open"]) &
            (x["open"] <= x["k_prev_close"])
        )
        x.loc[bullish_engulfing, ["kpattern_code", "kpattern_name"]] = ["BULLISH_ENGULFING", "多方吞噬"]

        bearish_engulfing = (
            x["k_is_bear"].fillna(False) &
            (x["k_prev_close"] > x["k_prev_open"]) &
            (x["close"] <= x["k_prev_open"]) &
            (x["open"] >= x["k_prev_close"])
        )
        x.loc[bearish_engulfing, ["kpattern_code", "kpattern_name"]] = ["BEARISH_ENGULFING", "空方吞噬"]

        return x
