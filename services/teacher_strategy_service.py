try:
    from teacher_strategy.teacher_strategy_engine import TeacherStrategyEngine
except Exception:
    TeacherStrategyEngine = None


class TeacherStrategyService:

    def __init__(self):

        self.engine = None

        if TeacherStrategyEngine:
            self.engine = TeacherStrategyEngine()

    def run(
        self,
        ranking_df=None,
        price_history_df=None,
        market_df=None,
        institutional_df=None
    ):

        if self.engine is None:
            return None

        results = []

        if ranking_df is None:
            return None

        for _, row in ranking_df.iterrows():

            result = self.engine.analyze_row(row)

            result["stock_id"] = row.get("stock_id")

            results.append(result)

        import pandas as pd

        return pd.DataFrame(results)
