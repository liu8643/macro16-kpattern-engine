import pandas as pd
import numpy as np


class KPatternFeatureEngine:
    """計算K棒基礎特徵。只產生特徵，不做交易決策。"""

    REQUIRED = ["open", "high", "low", "close", "volume"]

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        for c in self.REQUIRED:
            if c not in x.columns:
                x[c] = np.nan
            x[c] = pd.to_numeric(x[c], errors="coerce")

        x["k_body"] = (x["close"] - x["open"]).abs()
        x["k_range"] = (x["high"] - x["low"]).abs().clip(lower=1e-9)
        x["k_upper_shadow"] = x["high"] - x[["open", "close"]].max(axis=1)
        x["k_lower_shadow"] = x[["open", "close"]].min(axis=1) - x["low"]
        x["k_body_ratio"] = x["k_body"] / x["k_range"]
        x["k_upper_shadow_ratio"] = x["k_upper_shadow"] / x["k_range"]
        x["k_lower_shadow_ratio"] = x["k_lower_shadow"] / x["k_range"]
        x["k_close_position"] = (x["close"] - x["low"]) / x["k_range"]
        x["k_is_bull"] = x["close"] > x["open"]
        x["k_is_bear"] = x["close"] < x["open"]
        x["k_prev_open"] = x["open"].shift(1)
        x["k_prev_close"] = x["close"].shift(1)
        return x
