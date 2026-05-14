class TwoHighFailEngine:
    """顧奎國老師策略：兩高不過 / 弱勢排除 Gate V17-R3。

    設計原則：
    1. 只負責「排除 / 警告」風險，不直接產生 BUY。
    2. BLOCK 只給明確弱勢或假突破；WARNING 給高檔過熱但尚未跌破的狀態。
    3. 輸出欄位固定，供 TeacherStrategyEngine / UI / Excel 共用。
    """

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            if v is None or v == "":
                return default
            return float(v)
        except Exception:
            return default

    def _first_num(self, row, keys, default=0.0):
        for k in keys:
            v = self._num(row, k, None)
            if v is not None:
                return v
        return default

    def analyze(self, row):
        close = self._first_num(row, ["close", "收盤", "最新收盤", "現價"], 0.0)
        high = self._first_num(row, ["high", "最高"], close)
        ma20 = self._first_num(row, ["ma20", "MA20"], 0.0)
        ma60 = self._first_num(row, ["ma60", "MA60"], 0.0)
        rsi = self._first_num(row, ["rsi14", "rsi", "RSI14", "RSI"], 50.0)
        volume_ratio = self._first_num(row, ["volume_ratio", "量比", "vol_ratio"], 1.0)
        price_dev = self._first_num(row, ["price_deviation", "price_dev", "乖離"], 0.0)
        macd_hist = self._first_num(row, ["macd_hist", "MACD_Hist", "macd柱"], 0.0)
        recent_high_1 = self._first_num(row, ["recent_high_1", "前高1", "high_1"], 0.0)
        recent_high_2 = self._first_num(row, ["recent_high_2", "前高2", "high_2"], 0.0)
        prev_high = self._first_num(row, ["prev_high", "前高"], recent_high_1)

        reasons = []
        score = 0.0
        hard_flags = []
        warning_flags = []

        if close <= 0:
            return {
                "two_high_fail": False,
                "weak_gate": "NE",
                "weak_score": 0.0,
                "weak_reason": "資料不足，無法判斷兩高不過",
                "teacher_exclusion_level": "NE",
            }

        # 兩高不過：必須同時有兩個有效前高，且收盤未站上兩高。
        if recent_high_1 > 0 and recent_high_2 > 0 and close < recent_high_1 and close < recent_high_2:
            score += 35
            hard_flags.append("TWO_HIGH_FAIL")
            reasons.append("兩高不過：收盤價未能站上前兩個高點")

        # 假突破：盤中過前高，但收盤跌回前高下。
        if prev_high > 0 and high > prev_high and close < prev_high:
            score += 30
            hard_flags.append("FALSE_BREAKOUT")
            reasons.append("假突破：盤中突破前高但收盤跌回壓力下方")

        # 高檔爆量不漲：視為警告或排除加分，不單獨強制 BLOCK。
        if volume_ratio >= 1.8 and price_dev >= 0.12 and high > close * 1.01:
            score += 20
            warning_flags.append("HIGH_VOLUME_NO_ADVANCE")
            reasons.append("高檔爆量不漲：量能放大但收盤無法站穩高點")

        # 高檔背離。
        if rsi >= 72 and macd_hist < 0:
            score += 20
            warning_flags.append("RSI_MACD_DIVERGENCE")
            reasons.append("高檔背離：RSI偏高但MACD柱狀體轉弱")

        # 均線破壞。
        if ma20 > 0 and close < ma20:
            score += 15
            warning_flags.append("BELOW_MA20")
            reasons.append("跌破月線：短線結構轉弱")
        if ma60 > 0 and close < ma60:
            score += 20
            hard_flags.append("BELOW_MA60")
            reasons.append("跌破季線：中期結構轉弱")

        two_high_fail = "TWO_HIGH_FAIL" in hard_flags
        # BLOCK 條件：明確兩高不過/假突破 + 總分，或跌破季線且總分偏高。
        if score >= 70 or ("FALSE_BREAKOUT" in hard_flags and score >= 50) or (two_high_fail and score >= 35):
            weak_gate = "BLOCK"
            exclusion = "HARD_BLOCK"
        elif score >= 35:
            weak_gate = "WARNING"
            exclusion = "WARNING"
        else:
            weak_gate = "PASS"
            exclusion = "PASS"

        return {
            "two_high_fail": bool(two_high_fail),
            "weak_gate": weak_gate,
            "weak_score": round(score, 2),
            "weak_reason": "；".join(reasons) if reasons else "未觸發兩高不過或弱勢排除條件",
            "teacher_exclusion_level": exclusion,
        }
