from .position_stage_engine import PositionStageEngine
from .two_high_fail_engine import TwoHighFailEngine
from .rotation_filter_engine import RotationFilterEngine


class TeacherStrategyEngine:

    def __init__(self):
        self.position_engine = PositionStageEngine()
        self.two_high_engine = TwoHighFailEngine()
        self.rotation_engine = RotationFilterEngine()

    def analyze_row(self, row):

        result = {}

        position_stage = self.position_engine.analyze(row)
        two_high_fail = self.two_high_engine.analyze(row)
        rotation = self.rotation_engine.analyze(row)

        result["position_stage"] = position_stage
        result["two_high_fail"] = two_high_fail
        result["rotation"] = rotation

        # 老師策略決策
        if two_high_fail:
            result["teacher_final_decision"] = "AVOID"
            result["teacher_light"] = "🔴"

        elif position_stage == "低位階翻多" and rotation == "主流":
            result["teacher_final_decision"] = "LOW BUY"
            result["teacher_light"] = "🟢"

        else:
            result["teacher_final_decision"] = "WATCH"
            result["teacher_light"] = "🔵"

        return result
