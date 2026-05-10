class LowBaseReversalEngine:
    """低位階翻多輔助分類引擎。"""

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            return float(v) if v is not None else default
        except Exception:
            return default

    def analyze(self, row):
        close = self._num(row, "close")
        ma20 = self._num(row, "ma20")
        ma60 = self._num(row, "ma60")
        revenue_yoy = self._num(row, "revenue_yoy", 0)
        dividend_yield = self._num(row, "dividend_yield", 0)
        if close > 0 and ma20 > 0 and ma60 > 0 and close > ma20 and close <= ma60 * 1.1:
            if revenue_yoy >= 10:
                return {"low_base_type": "營收成長低位階翻多", "low_base_score": 85}
            if dividend_yield >= 4:
                return {"low_base_type": "高殖利率防守低位階", "low_base_score": 75}
            return {"low_base_type": "技術低位階翻多", "low_base_score": 80}
        return {"low_base_type": "非低位階", "low_base_score": 0}
