import pandas as pd

from .kpattern_feature_engine import KPatternFeatureEngine
from .sakata_pattern_engine import SakataPatternEngine
from .pattern_position_engine import PatternPositionEngine
from .volume_confirm_engine import VolumeConfirmEngine
from .pattern_score_engine import PatternScoreEngine
from .decision_layer_engine import KPatternDecisionLayerEngine


class KPatternPipeline:
    """KPattern 固定管線。

    主程式只需要呼叫：
        hist = KPatternPipeline().run(hist, stock_id=stock_id)
        info = KPatternPipeline.latest_dict(hist)

    後續新增 24 / 48 型態時，只修改 engines 或 config，不再修改主程式。
    """

    DEFAULT = {
        "kpattern_code": "NEUTRAL",
        "kpattern_name": "無明確型態",
        "kpattern_signal": "NEUTRAL",
        "kpattern_position": "資料不足",
        "kpattern_volume_confirm": "中性",
        "kpattern_score": 0.0,
        "kpattern_rank_adjustment": 0.0,
        "kpattern_trade_bias": "NEUTRAL",
        "kpattern_risk_flag": "NO",
        "kpattern_reason": "無明確型態",
        "kpattern_trade_advice": "觀察",
        "final_trade_score": 0.0,
    }

    def __init__(self):
        self.feature = KPatternFeatureEngine()
        self.pattern = SakataPatternEngine()
        self.position = PatternPositionEngine()
        self.volume = VolumeConfirmEngine()
        self.score = PatternScoreEngine()
        self.decision = KPatternDecisionLayerEngine()

    def run(self, df: pd.DataFrame, price_history_df: pd.DataFrame | None = None, stock_id: str | None = None, **kwargs) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        x = df.copy()
        if "date" in x.columns:
            x = x.sort_values("date").reset_index(drop=True)
        if len(x) < 5:
            for k, v in self.DEFAULT.items():
                x[k] = v
            x["kpattern_code"] = "INSUFFICIENT_DATA"
            x["kpattern_name"] = "資料不足"
            x["kpattern_reason"] = "少於5根K棒，無法判斷型態"
            return x
        try:
            x = self.feature.build(x)
            x = self.pattern.classify(x)
            x = self.position.build(x)
            x = self.volume.build(x)
            x = self.score.build(x)
            x = self.decision.build(x)
        except Exception as exc:
            for k, v in self.DEFAULT.items():
                x[k] = v
            x["kpattern_code"] = "PLUGIN_ERROR"
            x["kpattern_name"] = "KPattern執行失敗"
            x["kpattern_reason"] = str(exc)
        return x

    @staticmethod
    def latest_dict(df: pd.DataFrame) -> dict:
        out = dict(KPatternPipeline.DEFAULT)
        if df is None or df.empty:
            return out
        try:
            row = df.iloc[-1]
            for k in out.keys():
                if k in df.columns:
                    val = row.get(k)
                    if pd.notna(val):
                        out[k] = val
        except Exception as exc:
            out["kpattern_code"] = "LATEST_ERROR"
            out["kpattern_name"] = "KPattern讀取失敗"
            out["kpattern_reason"] = str(exc)
        return out
