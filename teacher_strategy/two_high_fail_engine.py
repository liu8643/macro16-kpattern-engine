class TwoHighFailEngine:
    """
    兩高不過判定
    """

    def analyze(self, row):
        try:
            high1 = row.get("recent_high_1", 0)
            high2 = row.get("recent_high_2", 0)
            close = row.get("close", 0)

            if close < high1 and close < high2:
                return True

            return False

        except Exception:
            return False
