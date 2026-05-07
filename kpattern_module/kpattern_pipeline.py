from kpattern_module.kpattern_feature_engine import KPatternFeatureEngine
from kpattern_module.sakata_pattern_engine import SakataPatternEngine
from kpattern_module.pattern_position_engine import PatternPositionEngine
from kpattern_module.volume_confirm_engine import VolumeConfirmEngine
from kpattern_module.pattern_score_engine import PatternScoreEngine
from kpattern_module.decision_layer_engine import KPatternDecisionLayerEngine


class KPatternPipeline:
    """
    KPattern 主流程
    串接各子引擎，輸出主程式固定讀取的 kpattern_* 欄位。

    注意：
    - 各子引擎目前的方法名稱不是全部叫 run。
    - feature / position / volume / score / decision 使用 build。
    - sakata 使用 classify。
    - 因此 pipeline 不可呼叫 self.xxx.run()，否則會出現 AttributeError。
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

        # 1. K棒基礎特徵
        x = self.feature.build(x)

        # 2. 阪田 / K線型態辨識
        x = self.sakata.classify(x)

        # 3. 位階判斷
        x = self.position.build(x)

        # 4. 量能確認
        x = self.volume.build(x)

        # 5. 型態分數與交易偏向
        x = self.score.build(x)

        # 6. 穩定輸出層：final_trade_score / kpattern_rank_adjustment / kpattern_signal
        x = self.decision.build(x)

        return x
