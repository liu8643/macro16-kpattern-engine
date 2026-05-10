class RotationFilterEngine:
    """
    類股輪動判定
    """

    def analyze(self, row):
        try:
            sector_strength = row.get("sector_strength", 0)

            if sector_strength >= 80:
                return "主流"

            elif sector_strength >= 60:
                return "輪動"

            return "弱勢"

        except Exception:
            return "未知"
