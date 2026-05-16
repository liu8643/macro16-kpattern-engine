KPATTERN_READ_ONLY_GUARD = True  # V17.1: KPattern service must be read-only; no external API / no DB rebuild.
from kpattern_module.kpattern_pipeline import KPatternPipeline


class KPatternService:
    """相容層：主程式可透過 Service 呼叫 KPatternPipeline。"""

    def __init__(self):
        self.pipeline = KPatternPipeline()

    def run(self, df, price_history=None, **kwargs):
        return self.pipeline.run(
            df,
            price_history_df=price_history,
            **kwargs
        )
