class TwoHighFailEngine:
    """
    顧奎國老師策略：兩高不過 / 弱勢排除引擎

    用途：
    1. 判斷兩高不過
    2. 判斷高檔爆量不漲
    3. 判斷高檔背離
    4. 判斷假突破
    5. 產出 weak_gate 給 teacher_strategy_engine 強制排除
    """

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            if v is None:
                return default
            return float(v)
        except Exception:
            return default

    def analyze(self, row):
        close = self._num(row, "close")
        high = self._num(row, "high", close)
        ma20 = self._num(row, "ma20")
        ma60 = self._num(row, "ma60")
        rsi = self._num(row, "rsi14", self._num(row, "rsi", 50))
        volume_ratio = self._num(row, "volume_ratio", 1.0)
        price_dev = self._num(row, "price_deviation", self._num(row, "price_dev", 0.0))
        macd_hist = self._num(row, "macd_hist", 0.0)

        recent_high_1 = self._num(row, "recent_high_1", 0)
        recent_high_2 = self._num(row, "recent_high_2", 0)
        prev_high = self._num(row, "prev_high", recent_high_1)

        reasons = []
        score = 0

        if close <= 0:
            return {
                "two_high_fail": False,
                "weak_gate": "NE",
                "weak_score": 0,
                "weak_reason": "資料不足，無法判斷兩高不過"
            }

        # 1. 兩高不過：收盤價連續無法站上兩個前高
        if recent_high_1 > 0 and recent_high_2 > 0:
            if close < recent_high_1 and close < recent_high_2:
                score += 35
                reasons.append("兩高不過：收盤價未能站上前兩個高點")

        # 2. 高檔爆量不漲：量大但收盤無法推升
        if volume_ratio >= 1.8 and price_dev >= 0.15:
            if high > close * 1.01:
                score += 25
                reasons.append("高檔爆量不漲：量能放大但收盤無法站穩高點")

        # 3. RSI高檔背離風險
        if rsi >= 72 and macd_hist < 0:
            score += 20
            reasons.append("高檔背離：RSI偏高但MACD柱狀體轉弱")

        # 4. 假突破：突破前高但收盤跌回
        if prev_high > 0 and high > prev_high and close < prev_high:
            score += 30
            reasons.append("假突破：盤中突破前高但收盤跌回壓力下方")

        # 5. 跌破月線：轉弱確認
        if ma20 > 0 and close < ma20:
            score += 15
            reasons.append("跌破月線：短線結構轉弱")

        # 6. 跌破季線：中期轉弱
        if ma60 > 0 and close < ma60:
            score += 20
            reasons.append("跌破季線：中期結構轉弱")

        two_high_fail = score >= 35

        if score >= 70:
            weak_gate = "BLOCK"
        elif score >= 35:
            weak_gate = "WARNING"
        else:
            weak_gate = "PASS"

        return {
            "two_high_fail": bool(two_high_fail),
            "weak_gate": weak_gate,
            "weak_score": round(score, 2),
            "weak_reason": "；".join(reasons) if reasons else "未觸發兩高不過或弱勢排除條件"
        }
