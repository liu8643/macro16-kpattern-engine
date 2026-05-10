class TwoHighFailEngine:
    """顧奎國老師策略：兩高不過 / 弱勢排除 Gate。

    Phase4 FINAL：
    - 支援 recent_high_1 / recent_high_2 / prev_high 缺欄位時的安全降級。
    - 以 row 內可用欄位判定兩高不過、假突破、高檔爆量不漲、均線轉弱。
    - 回傳 weak_gate / weak_score / weak_reason，供老師策略總閘與排除清單 UI 使用。
    """

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            if v is None:
                return default
            return float(v)
        except Exception:
            return default

    def _first_num(self, row, keys, default=0.0):
        for key in keys:
            v = self._num(row, key, None)
            if v is not None and v != 0:
                return v
        return default

    def analyze(self, row):
        close = self._num(row, "close", self._num(row, "現價", 0.0))
        high = self._num(row, "high", close)
        ma20 = self._num(row, "ma20", self._num(row, "MA20", 0.0))
        ma60 = self._num(row, "ma60", self._num(row, "MA60", 0.0))
        rsi = self._num(row, "rsi14", self._num(row, "rsi", self._num(row, "RSI", 50)))
        volume_ratio = self._num(row, "volume_ratio", self._num(row, "量比", 1.0))
        price_dev = self._num(row, "price_deviation", self._num(row, "price_dev", self._num(row, "乖離", 0.0)))
        macd_hist = self._num(row, "macd_hist", 0.0)

        # recent_high_1/2 若主程式沒有給，改抓常見壓力欄位，避免 Gate 永遠不會觸發。
        recent_high_1 = self._first_num(row, ["recent_high_1", "prev_high", "resistance", "壓力", "target_price"], 0.0)
        recent_high_2 = self._first_num(row, ["recent_high_2", "recent_high", "high_20", "high_60", "target_1382"], 0.0)
        prev_high = self._first_num(row, ["prev_high", "recent_high_1", "resistance", "壓力"], recent_high_1)

        reasons = []
        score = 0.0
        if close <= 0:
            return {
                "two_high_fail": False,
                "weak_gate": "NE",
                "weak_score": 0.0,
                "weak_reason": "資料不足，無法判斷兩高不過",
            }

        if recent_high_1 > 0 and recent_high_2 > 0 and close < recent_high_1 and close < recent_high_2:
            score += 35
            reasons.append("兩高不過：收盤價未能站上前兩個高點")
        if volume_ratio >= 1.8 and price_dev >= 0.15 and high > close * 1.01:
            score += 25
            reasons.append("高檔爆量不漲：量能放大但收盤無法站穩高點")
        if rsi >= 72 and macd_hist < 0:
            score += 20
            reasons.append("高檔背離：RSI偏高但MACD柱狀體轉弱")
        if prev_high > 0 and high > prev_high and close < prev_high:
            score += 30
            reasons.append("假突破：盤中突破前高但收盤跌回壓力下方")
        if ma20 > 0 and close < ma20:
            score += 15
            reasons.append("跌破月線：短線結構轉弱")
        if ma60 > 0 and close < ma60:
            score += 20
            reasons.append("跌破季線：中期結構轉弱")

        two_high_fail = score >= 35
        weak_gate = "BLOCK" if score >= 70 else ("WARNING" if score >= 35 else "PASS")
        return {
            "two_high_fail": bool(two_high_fail),
            "weak_gate": weak_gate,
            "weak_score": round(score, 2),
            "weak_reason": "；".join(reasons) if reasons else "未觸發兩高不過或弱勢排除條件",
        }
