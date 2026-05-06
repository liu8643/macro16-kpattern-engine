from .feature_engine import KPatternFeatureEngine
from .sakata_engine import SakataPatternEngine
from .position_engine import PatternPositionEngine
from .volume_engine import VolumeConfirmEngine
from .score_engine import PatternScoreEngine
from .decision_engine import PatternDecisionEngine


class KPatternPipeline:
    """
    KPattern 主流程
    負責串接所有子引擎
    """

    def __init__(self):
        self.feature = KPatternFeatureEngine()
        self.pattern = SakataPatternEngine()
        self.position = PatternPositionEngine()
        self.volume = VolumeConfirmEngine()
        self.score = PatternScoreEngine()
        self.decision = PatternDecisionEngine()

    def run(self, df):
        """
        主執行流程
        """

        # 1. K線特徵
        df = self.feature.build(df)

        # 2. 型態判斷
        df = self.pattern.classify(df)

        # 3. 位階判斷
        df = self.position.build(df)

        # 4. 量能判斷
        df = self.volume.build(df)

        # 5. 分數與方向
        df = self.score.build(df)

        # 6. 決策層
        df = self.decision.build(df)

        return df
