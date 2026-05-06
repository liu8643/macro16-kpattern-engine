import pandas as pd


class PatternSelectorEngine:
    """
    選股引擎（TOP20 / 可下單）
    """

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        # 排序（依分數）
        df = df.sort_values("pattern_score", ascending=False)

        # 排名
        df["rank"] = range(1, len(df) + 1)

        # TOP20
        df["is_top20"] = df["rank"] <= 20

        # TOP5
        df["is_top5"] = df["rank"] <= 5

        # 今日可下單
        df["trade_ready"] = (
            (df["decision"] == "BUY") &
            (df["pattern_score"] >= 20)
        )

        return df
