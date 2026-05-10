class PositionStageEngine:
    """
    判斷股票位階：
    低位階 / 中位階 / 高位階
    """

    def analyze(self, row):
        try:
            price = row.get("close", 0)
            ma20 = row.get("ma20", 0)
            ma60 = row.get("ma60", 0)

            if price > ma20 > ma60:
                return "低位階翻多"

            elif price > ma60:
                return "中位階"

            return "高位階"

        except Exception:
            return "未知"
