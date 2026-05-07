from kpattern_module.kpattern_feature_engine import KPatternFeatureEngine
from kpattern_module.sakata_pattern_engine import SakataPatternEngine
from kpattern_module.pattern_position_engine import PatternPositionEngine
from kpattern_module.volume_confirm_engine import VolumeConfirmEngine
from kpattern_module.pattern_score_engine import PatternScoreEngine
from kpattern_module.decision_layer_engine import KPatternDecisionLayerEngine


class KPatternPipeline:
    """
    KPattern 主流程
    只負責串接各子引擎，不直接改主程式交易邏輯。
    """

    def __init__(self):
        self.feature = KPatternFeatureEngine()
        self.sakata = SakataPatternEngine()
        self.position = PatternPositionEngine()
        self.volume = VolumeConfirmEngine()
        self.score = PatternScoreEngine()
        self.decision = KPatternDecisionLayerEngine()

    def run(self, df, price_history_df=None, **kwargs):
        if df is None:
            return df

        x = df.copy()

        x = self.feature.run(x, price_history_df=price_history_df, **kwargs)
        x = self.sakata.run(x, price_history_df=price_history_df, **kwargs)
        x = self.position.run(x, price_history_df=price_history_df, **kwargs)
        x = self.volume.run(x, price_history_df=price_history_df, **kwargs)
        x = self.score.run(x, price_history_df=price_history_df, **kwargs)
        x = self.decision.run(x, price_history_df=price_history_df, **kwargs)

        return x
