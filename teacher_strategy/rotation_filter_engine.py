class RotationFilterEngine:
    """
    市場級 AI：類股輪動引擎

    功能：
    1. 判斷市場主流族群
    2. 判斷 AI 主流 / 補漲 / 弱勢
    3. 提供 teacher_strategy_engine 使用
    4. 不在 UI 層重算
    """

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            if v is None:
                return default
            return float(v)
        except Exception:
            return default

    def _text(self, row, key, default=""):
        try:
            v = row.get(key, default)
            if v is None:
                return default
            return str(v)
        except Exception:
            return default

    def analyze(self, row):

        industry = self._text(row, "industry")
        sub_theme = self._text(row, "sub_theme")
        theme = self._text(row, "theme")

        sector_strength = self._num(row, "sector_strength", 50)
        flow_score = self._num(row, "flow_score", 50)
        rs_score = self._num(row, "rs_score", 50)
        volume_ratio = self._num(row, "volume_ratio", 1.0)
        price_dev = self._num(row, "price_deviation", 0)

        rotation = "中性"
        sector_strength_score = sector_strength
        rotation_reason = []

        text_all = f"{industry} {theme} {sub_theme}"

        # =========================
        # AI / 主流族群
        # =========================

        ai_keywords = [
            "AI",
            "伺服器",
            "CPO",
            "矽光子",
            "ASIC",
            "HBM",
            "BBU",
            "液冷",
            "機器人",
            "邊緣AI",
            "光通訊",
            "CoWoS",
            "AEC",
            "高速傳輸",
        ]

        # =========================
        # 非AI主流
        # =========================

        strong_keywords = [
            "重電",
            "軍工",
            "低軌衛星",
            "散熱",
            "網通",
            "工業電腦",
            "車用",
        ]

        # =========================
        # 弱勢族群
        # =========================

        weak_keywords = [
            "營建",
            "塑化",
            "觀光",
            "百貨",
            "傳產",
        ]

        # =========================
        # AI主流
        # =========================

        if any(k in text_all for k in ai_keywords):

            rotation = "AI主流"

            sector_strength_score += 20

            rotation_reason.append("屬AI主流族群")

            if flow_score >= 60:
                sector_strength_score += 10
                rotation_reason.append("資金流入AI主流")

            if rs_score >= 60:
                sector_strength_score += 10
                rotation_reason.append("相對強弱優於大盤")

        # =========================
        # 強勢主流
        # =========================

        elif any(k in text_all for k in strong_keywords):

            rotation = "強勢主流"

            sector_strength_score += 12

            rotation_reason.append("屬市場強勢輪動族群")

        # =========================
        # 弱勢族群
        # =========================

        elif any(k in text_all for k in weak_keywords):

            rotation = "弱勢族群"

            sector_strength_score -= 20

            rotation_reason.append("屬市場弱勢族群")

        # =========================
        # 補漲股
        # =========================

        if (
            sector_strength >= 60
            and rs_score >= 55
            and price_dev < 0.12
            and volume_ratio >= 1.2
        ):
            sector_strength_score += 8
            rotation_reason.append("具補漲條件")

        # =========================
        # 爆量轉強
        # =========================

        if volume_ratio >= 1.8 and rs_score >= 60:
            sector_strength_score += 5
            rotation_reason.append("量價同步轉強")

        # =========================
        # 高檔過熱
        # =========================

        if price_dev >= 0.20:
            sector_strength_score -= 10
            rotation_reason.append("乖離過大，需防追高")

        # =========================
        # 評分限制
        # =========================

        sector_strength_score = max(0, min(100, sector_strength_score))

        # =========================
        # 輪動等級
        # =========================

        if sector_strength_score >= 85:
            rotation_level = "S"
        elif sector_strength_score >= 70:
            rotation_level = "A"
        elif sector_strength_score >= 55:
            rotation_level = "B"
        elif sector_strength_score >= 40:
            rotation_level = "C"
        else:
            rotation_level = "D"

        return {
            "rotation": rotation,
            "rotation_level": rotation_level,
            "sector_strength_score": round(sector_strength_score, 2),
            "rotation_reason": "；".join(rotation_reason)
            if rotation_reason
            else "無明顯主流輪動特徵",
        }
