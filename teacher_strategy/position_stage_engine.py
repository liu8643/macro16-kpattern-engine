class PositionStageEngine:
    """
    奎國老師策略：位階判斷引擎
    輸出：低位階翻多、主升初段、主升3浪、高位階過熱、修正段、ABC修正待確認。
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
        ma5 = self._num(row, "ma5")
        ma10 = self._num(row, "ma10")
        ma20 = self._num(row, "ma20")
        ma60 = self._num(row, "ma60")
        rsi = self._num(row, "rsi14", self._num(row, "rsi", 50))
        volume_ratio = self._num(row, "volume_ratio", 1.0)
        price_dev = self._num(row, "price_deviation", self._num(row, "price_dev", 0.0))
        macd_hist = self._num(row, "macd_hist", 0.0)

        if close <= 0 or ma20 <= 0 or ma60 <= 0:
            return {"position_stage": "資料不足", "position_score": 0, "position_reason": "缺少 close / ma20 / ma60，無法判斷位階"}

        above_ma20 = close > ma20
        above_ma60 = close > ma60
        ma_bull = ma5 > ma10 > ma20 > ma60 if ma5 > 0 and ma10 > 0 else False
        ma_turn_up = close > ma20 and ma20 >= ma60 * 0.98
        low_base = close <= ma60 * 1.12
        high_dev = price_dev >= 0.18 or close >= ma60 * 1.35
        overheat = rsi >= 75 and volume_ratio >= 1.5

        if above_ma20 and above_ma60 and high_dev and overheat:
            return {"position_stage": "高位階過熱", "position_score": 30, "position_reason": "股價乖離過大且 RSI/量能偏熱，屬高位階過熱，不適合追價"}
        if ma_bull and 55 <= rsi <= 72 and volume_ratio >= 1.2 and macd_hist >= 0:
            return {"position_stage": "主升3浪", "position_score": 95, "position_reason": "均線多頭排列、量能放大、RSI健康、MACD維持正向，符合主升3浪特徵"}
        if ma_turn_up and low_base and 45 <= rsi <= 65 and macd_hist >= 0:
            return {"position_stage": "主升初段", "position_score": 85, "position_reason": "股價站回中期均線，乖離尚低，屬低位階轉強或主升初段"}
        if close > ma20 and close <= ma60 * 1.10 and rsi >= 42 and volume_ratio >= 1.0:
            return {"position_stage": "低位階翻多", "position_score": 80, "position_reason": "股價站上月線但尚未大幅乖離季線，屬低位階翻多觀察"}
        if close < ma20 and close >= ma60 * 0.95 and rsi >= 38:
            return {"position_stage": "ABC修正待確認", "position_score": 55, "position_reason": "股價跌破月線但仍接近季線，可能處於ABC修正或修正末端，需等待止跌K棒"}
        if close < ma20 and close < ma60:
            return {"position_stage": "修正段", "position_score": 35, "position_reason": "股價跌破月線與季線，結構偏弱，屬修正段"}
        if above_ma20 and above_ma60 and not high_dev:
            return {"position_stage": "中位階整理", "position_score": 65, "position_reason": "股價在均線上方但尚未形成明確主升或過熱，屬中位階整理"}
        return {"position_stage": "中性觀察", "position_score": 50, "position_reason": "條件未形成明確主升、低位階翻多或修正段，暫列中性觀察"}
