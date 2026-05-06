import pandas as pd


class PatternDecisionEngine:
    """
    Decision Layer（核心決策引擎）

    功能：
    1. 判斷是否可交易
    2. 產生最終決策（BUY / WAIT / AVOID）
    3. 建立 ranking score
    """

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        # === 基礎條件 ===
        df["trade_allowed"] = True

        # === 禁止條件（高風險）===
        df.loc[df["pattern_direction"] == "AVOID", "trade_allowed"] = False

        # === 初始決策 ===
        df["decision"] = "WAIT"

        # === BUY 條件 ===
        cond_buy = (
            (df["pattern_direction"].isin(["BUY", "ATTACK"])) &
            (df["pattern_score"] >= 20)
        )

        df.loc[cond_buy, "decision"] = "BUY"

        # === AVOID ===
        df.loc[df["trade_allowed"] == False, "decision"] = "AVOID"

        # === 排名分數（可擴充）===
        df["rank_score"] = df["pattern_score"]

        return df
