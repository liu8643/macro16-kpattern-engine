import pandas as pd
import numpy as np


class VolumeConfirmEngine:
    """判斷量能確認。只新增 pattern_volume_confirm 與 volume_confirm_score。"""

    def build(self, df: pd.DataFrame) -> pd.DataFrame:
        x = df.copy()
        volume = pd.to_numeric(x.get("volume"), errors="coerce")
        vol5 = volume.rolling(5, min_periods=3).mean()
        vol20 = volume.rolling(20, min_periods=5).mean()
        ratio = volume / vol20.replace(0, np.nan)
        x["kpattern_volume_ratio"] = ratio.fillna(1.0)
        x["kpattern_volume_confirm"] = "中性"
        x["kpattern_volume_score"] = 0.0
        x.loc[ratio >= 2.0, ["kpattern_volume_confirm", "kpattern_volume_score"]] = ["爆量", -5.0]
        x.loc[(ratio >= 1.3) & (ratio < 2.0), ["kpattern_volume_confirm", "kpattern_volume_score"]] = ["放量", 5.0]
        x.loc[(ratio >= 1.05) & (ratio < 1.3), ["kpattern_volume_confirm", "kpattern_volume_score"]] = ["溫和放量", 3.0]
        x.loc[ratio < 0.75, ["kpattern_volume_confirm", "kpattern_volume_score"]] = ["量縮", -2.0]
        return x
