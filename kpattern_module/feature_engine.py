import pandas as pd


class KPatternFeatureEngine:
    """
    K線基礎特徵計算
    """

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        # 基本K線
        df["body"] = abs(df["close"] - df["open"])
        df["range"] = df["high"] - df["low"]

        # 上影線 / 下影線
        df["upper_shadow"] = df["high"] - df[["open", "close"]].max(axis=1)
        df["lower_shadow"] = df[["open", "close"]].min(axis=1) - df["low"]

        # 比例
        df["body_ratio"] = df["body"] / (df["range"] + 1e-9)
        df["upper_shadow_ratio"] = df["upper_shadow"] / (df["range"] + 1e-9)
        df["lower_shadow_ratio"] = df["lower_shadow"] / (df["range"] + 1e-9)

        # 收盤位置（0~1）
        df["close_position"] = (df["close"] - df["low"]) / (df["range"] + 1e-9)

        return df
