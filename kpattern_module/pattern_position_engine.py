import pandas as pd
import numpy as np


class PatternPositionEngine:
    """計算20/60/120日位階。只提供位置資訊，不覆蓋原系統分數。"""

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        close = pd.to_numeric(x.get("close"), errors="coerce")
        high = pd.to_numeric(x.get("high"), errors="coerce")
        low = pd.to_numeric(x.get("low"), errors="coerce")
        hi120 = high.rolling(120, min_periods=20).max()
        lo120 = low.rolling(120, min_periods=20).min()
        pos = (close - lo120) / (hi120 - lo120).replace(0, np.nan)
        x["kpattern_position_ratio"] = pos.clip(0, 1)
        x["kpattern_position"] = "資料不足"
        x.loc[x["kpattern_position_ratio"] <= 0.35, "kpattern_position"] = "低位階"
        x.loc[(x["kpattern_position_ratio"] > 0.35) & (x["kpattern_position_ratio"] < 0.70), "kpattern_position"] = "中位階"
        x.loc[x["kpattern_position_ratio"] >= 0.70, "kpattern_position"] = "高位階"
        return x
