from .position_stage_engine import PositionStageEngine
from .two_high_fail_engine import TwoHighFailEngine
from .rotation_filter_engine import RotationFilterEngine


class TeacherStrategyEngine:
    """
    顧奎國老師策略主引擎

    原則：
    1. UI 不計算策略，只讀本引擎輸出
    2. 主程式只 Hook / run / merge / display / export
    3. 本引擎整合：位階、兩高不過、類股輪動、資金/RS/RR
    """

    def __init__(self):
        self.position_engine = PositionStageEngine()
        self.two_high_engine = TwoHighFailEngine()
        self.rotation_engine = RotationFilterEngine()

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

    def _price_zone(self, row):
        close = self._num(row, "close")
        ma20 = self._num(row, "ma20")
        ma60 = self._num(row, "ma60")
        atr_pct = self._num(row, "atr_pct", 0.03)

        if close <= 0:
            return "", "", ""

        if atr_pct <= 0:
            atr_pct = 0.03

        entry_low = close * (1 - atr_pct * 0.8)
        entry_high = close * (1 - atr_pct * 0.2)

        stop_base = ma20 if ma20 > 0 else close * 0.94
        stop_loss = min(stop_base * 0.98, close * 0.94)

        target_base = ma60 if ma60 > close else close
        target_price = max(close * 1.08, target_base * 1.05)

        return (
            f"{entry_low:.2f} ~ {entry_high:.2f}",
            f"{stop_loss:.2f}",
            f"{target_price:.2f}",
        )

    def _calc_gate(self, decision):
        if decision in ("BUY", "LOW BUY"):
            return "PASS"
        if decision in ("WAIT_PULLBACK", "WATCH"):
            return "WATCH"
        return "BLOCK"

    def analyze_row(self, row):
        result = {}

        stock_id = self._text(row, "stock_id")
        stock_name = self._text(row, "stock_name")

        position_result = self.position_engine.analyze(row)
        if isinstance(position_result, dict):
            position_stage = position_result.get("position_stage", "未知")
            position_score = self._num(position_result, "position_score", 0)
            position_reason = position_result.get("position_reason", "")
        else:
            position_stage = str(position_result)
            position_score = 50
            position_reason = ""

        two_high_fail = self.two_high_engine.analyze(row)
        two_high_fail = bool(two_high_fail)

        rotation_result = self.rotation_engine.analyze(row)
        if isinstance(rotation_result, dict):
            rotation = rotation_result.get("rotation", "未知")
            sector_strength_score = self._num(rotation_result, "sector_strength_score", 0)
            rotation_reason = rotation_result.get("rotation_reason", "")
        else:
            rotation = str(rotation_result)
            sector_strength_score = self._num(row, "sector_strength", 0)
            rotation_reason = ""

        flow_score = self._num(row, "flow_score", self._num(row, "institutional_score", 50))
        rs_score = self._num(row, "rs_score", self._num(row, "relative_strength_score", 50))
        rr = self._num(row, "rr_live", self._num(row, "rr", 0))
        rsi = self._num(row, "rsi14", self._num(row, "rsi", 50))
        price_dev = self._num(row, "price_deviation", self._num(row, "price_dev", 0))

        teacher_buy_zone, teacher_stop_loss, teacher_target_price = self._price_zone(row)

        teacher_strategy_class = "觀察"
        teacher_final_decision = "WATCH"
        teacher_light = "🔵"
        reason_parts = []

        # 1. 強制排除：兩高不過
        if two_high_fail:
            teacher_strategy_class = "排除"
            teacher_final_decision = "AVOID"
            teacher_light = "🔴"
            reason_parts.append("兩高不過成立，弱勢反彈或壓力未突破，不得進今日可買")

        # 2. 高位階過熱：不追
        elif position_stage == "高位階過熱" or price_dev >= 0.20 or rsi >= 78:
            teacher_strategy_class = "排除"
            teacher_final_decision = "REDUCE"
            teacher_light = "🔴"
            reason_parts.append("高位階過熱或乖離過大，禁止追高，偏減碼或避開")

        # 3. 主升3浪：主攻
        elif position_stage == "主升3浪" and sector_strength_score >= 70 and flow_score >= 50 and rs_score >= 55:
            teacher_strategy_class = "主攻"
            teacher_final_decision = "BUY"
            teacher_light = "🟢"
            reason_parts.append("主升3浪 + 類股強 + 資金/RS支撐，列入老師策略主攻")

        # 4. 主升初段 / 低位階翻多：低接
        elif position_stage in ("主升初段", "低位階翻多") and sector_strength_score >= 55:
            teacher_strategy_class = "低接"
            teacher_final_decision = "LOW BUY"
            teacher_light = "🟢"
            reason_parts.append("低位階翻多或主升初段，適合拉回低接與卡位")

        # 5. 中位階整理：等拉回
        elif position_stage == "中位階整理" and rs_score >= 50:
            teacher_strategy_class = "等拉回"
            teacher_final_decision = "WAIT_PULLBACK"
            teacher_light = "🟡"
            reason_parts.append("中位階整理且相對強弱尚可，但不追價，等待拉回")

        # 6. ABC修正：觀察
        elif position_stage == "ABC修正待確認":
            teacher_strategy_class = "觀察"
            teacher_final_decision = "WATCH"
            teacher_light = "🔵"
            reason_parts.append("可能為ABC修正或修正末端，需等待止跌K棒確認")

        # 7. 修正段：避開
        elif position_stage == "修正段":
            teacher_strategy_class = "排除"
            teacher_final_decision = "AVOID"
            teacher_light = "🔴"
            reason_parts.append("股價結構偏弱，屬修正段，不主動進場")

        else:
            teacher_strategy_class = "觀察"
            teacher_final_decision = "WATCH"
            teacher_light = "🔵"
            reason_parts.append("條件尚未形成主攻或低接，列觀察")

        if rr and rr < 1.0 and teacher_final_decision in ("BUY", "LOW BUY"):
            teacher_final_decision = "WATCH"
            teacher_light = "🔵"
            teacher_strategy_class = "觀察"
            reason_parts.append("RR低於1，主攻/低接降級為觀察")

        if position_reason:
            reason_parts.append(position_reason)

        if rotation_reason:
            reason_parts.append(rotation_reason)

        result["stock_id"] = stock_id
        result["stock_name"] = stock_name
        result["teacher_strategy_class"] = teacher_strategy_class
        result["teacher_final_decision"] = teacher_final_decision
        result["teacher_light"] = teacher_light
        result["position_stage"] = position_stage
        result["position_score"] = round(position_score, 2)
        result["two_high_fail"] = two_high_fail
        result["rotation"] = rotation
        result["sector_strength_score"] = round(sector_strength_score, 2)
        result["flow_score"] = round(flow_score, 2)
        result["rs_score"] = round(rs_score, 2)
        result["teacher_buy_zone"] = teacher_buy_zone
        result["teacher_stop_loss"] = teacher_stop_loss
        result["teacher_target_price"] = teacher_target_price
        result["teacher_reason"] = "；".join(reason_parts)
        result["teacher_gate"] = self._calc_gate(teacher_final_decision)
        result["teacher_source"] = "teacher_strategy_engine_v2"

        return result
