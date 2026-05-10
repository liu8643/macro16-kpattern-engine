class LowBaseReversalEngine:
    """低位階翻多交易化輔助分類引擎。

    Phase4 FINAL：
    - 不只回傳分類，還回傳原因與可交易分數。
    - 判斷低位階、站回月線、接近季線、營收/殖利率/RS/資金輔助。
    - 給 TeacherStrategyEngine 作為 LOW_BUY / 低位階翻多清單來源。
    """

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            return float(v) if v is not None else default
        except Exception:
            return default

    def analyze(self, row):
        close = self._num(row, "close", self._num(row, "現價", 0.0))
        ma20 = self._num(row, "ma20", self._num(row, "MA20", 0.0))
        ma60 = self._num(row, "ma60", self._num(row, "MA60", 0.0))
        ma120 = self._num(row, "ma120", 0.0)
        rsi = self._num(row, "rsi14", self._num(row, "rsi", 50.0))
        revenue_yoy = self._num(row, "revenue_yoy", 0.0)
        dividend_yield = self._num(row, "dividend_yield", 0.0)
        flow_score = self._num(row, "flow_score", self._num(row, "institutional_score", 50.0))
        rs_score = self._num(row, "rs_score", self._num(row, "relative_strength_score", 50.0))
        price_dev = self._num(row, "price_deviation", self._num(row, "price_dev", 0.0))

        reasons = []
        score = 0.0
        low_base_type = "非低位階"

        if close <= 0:
            return {"low_base_type": "資料不足", "low_base_score": 0.0, "low_base_reason": "缺少現價，無法判斷低位階"}

        above_ma20 = ma20 > 0 and close >= ma20
        near_ma60 = ma60 > 0 and close <= ma60 * 1.12
        below_ma120_or_near_base = (ma120 > 0 and close <= ma120 * 1.05) or (ma60 > 0 and close <= ma60 * 1.12) or price_dev <= 0.10

        if above_ma20:
            score += 25
            reasons.append("站回月線")
        if near_ma60:
            score += 20
            reasons.append("仍接近季線，位階未過高")
        if below_ma120_or_near_base:
            score += 15
            reasons.append("中長期位階仍偏低")
        if 40 <= rsi <= 68:
            score += 10
            reasons.append("RSI未過熱")
        if revenue_yoy >= 10:
            score += 15
            reasons.append("營收年增支撐")
        if dividend_yield >= 4:
            score += 10
            reasons.append("殖利率具防守性")
        if flow_score >= 55:
            score += 5
            reasons.append("資金流向改善")
        if rs_score >= 55:
            score += 5
            reasons.append("相對強弱轉佳")

        score = max(0.0, min(100.0, round(score, 2)))
        if score >= 85 and revenue_yoy >= 10:
            low_base_type = "營收成長低位階翻多"
        elif score >= 80:
            low_base_type = "技術低位階翻多"
        elif score >= 75 and dividend_yield >= 4:
            low_base_type = "高殖利率防守低位階"
        elif score >= 65:
            low_base_type = "低位階觀察"
        else:
            low_base_type = "非低位階"

        return {
            "low_base_type": low_base_type,
            "low_base_score": score,
            "low_base_reason": "；".join(reasons) if reasons else "未達低位階翻多條件",
        }
