try:
    from .position_stage_engine import PositionStageEngine
    from .two_high_fail_engine import TwoHighFailEngine
    from .rotation_filter_engine import RotationFilterEngine
    from .low_base_reversal_engine import LowBaseReversalEngine
except Exception:
    from position_stage_engine import PositionStageEngine
    from two_high_fail_engine import TwoHighFailEngine
    from rotation_filter_engine import RotationFilterEngine
    from low_base_reversal_engine import LowBaseReversalEngine


class TeacherStrategyEngine:
    """顧奎國老師策略主引擎 V17-R3：獨立老師分數，不再只沿用 TOP20 排序。"""

    def __init__(self):
        self.position_engine = PositionStageEngine()
        self.two_high_engine = TwoHighFailEngine()
        self.rotation_engine = RotationFilterEngine()
        self.low_base_engine = LowBaseReversalEngine()

    def _num(self, row, key, default=0.0):
        try:
            v = row.get(key, default)
            if v is None or v == "":
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
        atr_pct = self._num(row, "atr_pct", self._num(row, "atr", 0.03))
        if close <= 0:
            return "", "", ""
        if atr_pct <= 0 or atr_pct > 0.20:
            atr_pct = 0.03
        entry_low = close * (1 - atr_pct * 0.8)
        entry_high = close * (1 - atr_pct * 0.2)
        stop_base = ma20 if ma20 > 0 else close * 0.94
        stop_loss = min(stop_base * 0.98, close * 0.94)
        target_base = ma60 if ma60 > close else close
        target_price = max(close * 1.08, target_base * 1.05)
        return f"{entry_low:.2f} ~ {entry_high:.2f}", f"{stop_loss:.2f}", f"{target_price:.2f}"

    def _calc_gate(self, decision):
        d = str(decision).upper().replace(" ", "_")
        if d in ("BUY", "LOW_BUY"):
            return "PASS"
        if d in ("WAIT_PULLBACK", "WATCH"):
            return "WATCH"
        if d in ("REDUCE", "AVOID", "BLOCK", "SELL"):
            return "BLOCK"
        return "WATCH"

    def _teacher_score(self, position_score, sector_strength_score, flow_score, rs_score, rr, weak_score, low_base_score, decision):
        rr_score = 0.0
        try:
            rr = float(rr)
            if rr > 0:
                rr_score = min(rr / 2.0 * 100.0, 100.0)
        except Exception:
            rr_score = 0.0
        base = (
            position_score * 0.30 +
            sector_strength_score * 0.20 +
            flow_score * 0.15 +
            rs_score * 0.15 +
            rr_score * 0.10 +
            low_base_score * 0.10
        )
        penalty = min(max(weak_score, 0.0), 100.0) * 0.45
        d = str(decision).upper().replace(" ", "_")
        bonus = 8 if d == "BUY" else 5 if d == "LOW_BUY" else 0
        return round(max(0.0, min(100.0, base - penalty + bonus)), 2)

    def analyze_row(self, row):
        stock_id = self._text(row, "stock_id", self._text(row, "代號"))
        stock_name = self._text(row, "stock_name", self._text(row, "名稱"))

        position_result = self.position_engine.analyze(row)
        position_stage = position_result.get("position_stage", "未知") if isinstance(position_result, dict) else str(position_result)
        position_score = self._num(position_result, "position_score", 50) if isinstance(position_result, dict) else 50
        position_reason = position_result.get("position_reason", "") if isinstance(position_result, dict) else ""

        low_base_result = self.low_base_engine.analyze(row)
        low_base_type = low_base_result.get("low_base_type", "非低位階") if isinstance(low_base_result, dict) else "非低位階"
        low_base_score = self._num(low_base_result, "low_base_score", 0) if isinstance(low_base_result, dict) else 0

        two_high_result = self.two_high_engine.analyze(row)
        if isinstance(two_high_result, dict):
            two_high_fail = bool(two_high_result.get("two_high_fail", False))
            weak_gate = two_high_result.get("weak_gate", "PASS")
            weak_score = self._num(two_high_result, "weak_score", 0)
            weak_reason = two_high_result.get("weak_reason", "")
            teacher_exclusion_level = two_high_result.get("teacher_exclusion_level", weak_gate)
        else:
            two_high_fail = bool(two_high_result)
            weak_gate = "BLOCK" if two_high_fail else "PASS"
            weak_score = 100 if two_high_fail else 0
            weak_reason = "兩高不過成立" if two_high_fail else ""
            teacher_exclusion_level = weak_gate

        rotation_result = self.rotation_engine.analyze(row)
        rotation = rotation_result.get("rotation", "未知") if isinstance(rotation_result, dict) else str(rotation_result)
        sector_strength_score = self._num(rotation_result, "sector_strength_score", 0) if isinstance(rotation_result, dict) else self._num(row, "sector_strength", 0)
        rotation_reason = rotation_result.get("rotation_reason", "") if isinstance(rotation_result, dict) else ""

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

        if weak_gate == "BLOCK" or two_high_fail:
            teacher_strategy_class = "排除"
            teacher_final_decision = "AVOID"
            teacher_light = "🔴"
            reason_parts.append("弱勢排除Gate成立，兩高不過/假突破/跌破關鍵均線，不得進今日可買")
        elif position_stage == "高位階過熱" or price_dev >= 0.20 or rsi >= 78:
            teacher_strategy_class = "排除"
            teacher_final_decision = "REDUCE"
            teacher_light = "🔴"
            reason_parts.append("高位階過熱或乖離過大，禁止追高，偏減碼或避開")
        elif position_stage == "主升3浪" and sector_strength_score >= 65 and flow_score >= 50 and rs_score >= 55:
            teacher_strategy_class = "主攻"
            teacher_final_decision = "BUY"
            teacher_light = "🟢"
            reason_parts.append("主升3浪 + 類股強 + 資金/RS支撐，列入老師策略主攻")
        elif (position_stage in ("主升初段", "低位階翻多") or low_base_score >= 75) and sector_strength_score >= 50:
            teacher_strategy_class = "低接"
            teacher_final_decision = "LOW BUY"
            teacher_light = "🟢"
            reason_parts.append("低位階翻多或主升初段，適合拉回低接與卡位")
        elif position_stage == "中位階整理" and rs_score >= 50:
            teacher_strategy_class = "等拉回"
            teacher_final_decision = "WAIT_PULLBACK"
            teacher_light = "🟡"
            reason_parts.append("中位階整理且相對強弱尚可，不追價，等待拉回")
        elif position_stage == "ABC修正待確認":
            teacher_strategy_class = "觀察"
            teacher_final_decision = "WATCH"
            teacher_light = "🔵"
            reason_parts.append("可能為ABC修正或修正末端，需等待止跌K棒確認")
        elif position_stage == "修正段":
            teacher_strategy_class = "排除"
            teacher_final_decision = "AVOID"
            teacher_light = "🔴"
            reason_parts.append("股價結構偏弱，屬修正段，不主動進場")
        elif weak_gate == "WARNING":
            teacher_strategy_class = "觀察"
            teacher_final_decision = "WATCH"
            teacher_light = "🔵"
            reason_parts.append("弱勢Gate警告，先列觀察，不列今日可買")
        else:
            reason_parts.append("條件尚未形成主攻或低接，列觀察")

        if rr and rr < 1.0 and teacher_final_decision in ("BUY", "LOW BUY"):
            teacher_final_decision = "WATCH"
            teacher_light = "🔵"
            teacher_strategy_class = "觀察"
            reason_parts.append("RR低於1，主攻/低接降級為觀察")

        teacher_score = self._teacher_score(position_score, sector_strength_score, flow_score, rs_score, rr, weak_score, low_base_score, teacher_final_decision)

        if low_base_type and low_base_type != "非低位階":
            reason_parts.append(low_base_type)
        if position_reason:
            reason_parts.append(position_reason)
        if weak_reason:
            reason_parts.append(weak_reason)
        if rotation_reason:
            reason_parts.append(rotation_reason)

        return {
            "stock_id": stock_id,
            "stock_name": stock_name,
            "teacher_strategy_class": teacher_strategy_class,
            "teacher_final_decision": teacher_final_decision,
            "teacher_light": teacher_light,
            "teacher_gate": self._calc_gate(teacher_final_decision),
            "teacher_score": teacher_score,
            "position_stage": position_stage,
            "position_score": round(position_score, 2),
            "low_base_type": low_base_type,
            "low_base_score": round(low_base_score, 2),
            "two_high_fail": two_high_fail,
            "weak_gate": weak_gate,
            "weak_score": round(weak_score, 2),
            "teacher_exclusion_level": teacher_exclusion_level,
            "rotation": rotation,
            "sector_strength_score": round(sector_strength_score, 2),
            "flow_score": round(flow_score, 2),
            "rs_score": round(rs_score, 2),
            "teacher_buy_zone": teacher_buy_zone,
            "teacher_stop_loss": teacher_stop_loss,
            "teacher_target_price": teacher_target_price,
            "teacher_reason": "；".join(reason_parts),
            "teacher_source": "teacher_strategy_engine_v17_r3_semantic_unified",
        }
