# -*- coding: utf-8 -*-
"""資料品質 Gate：每日增量更新後與重排行前的快取品質檢查。"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, Dict

@dataclass
class QualityResult:
    passed: bool
    level: str
    reason: str
    metrics: Dict[str, Any]

class DataQualityGate:
    def __init__(self, max_ne_ratio: float = 0.80, min_feature_rows: int = 1):
        self.max_ne_ratio = float(max_ne_ratio)
        self.min_feature_rows = int(min_feature_rows)
    def check_financial_feature_cache(self, feature_rows: int = 0, ne_ratio: float = 1.0) -> QualityResult:
        feature_rows = int(feature_rows or 0)
        ne_ratio = float(ne_ratio if ne_ratio is not None else 1.0)
        if feature_rows < self.min_feature_rows:
            return QualityResult(False, "P0", "financial_feature_daily 無資料，禁止視為可下單依據", {"feature_rows": feature_rows, "ne_ratio": ne_ratio})
        if ne_ratio >= self.max_ne_ratio:
            return QualityResult(False, "P0", f"financial_feature_daily NE_ratio={ne_ratio:.2%} 過高", {"feature_rows": feature_rows, "ne_ratio": ne_ratio})
        return QualityResult(True, "OK", "financial_feature_daily 品質通過", {"feature_rows": feature_rows, "ne_ratio": ne_ratio})
    def check_market_snapshot(self, rows: int = 0, source_level: str = "") -> QualityResult:
        rows = int(rows or 0)
        source_level = str(source_level or "")
        if rows <= 0:
            return QualityResult(False, "P0", "market_snapshot 無資料", {"rows": rows, "source_level": source_level})
        if "fallback" in source_level.lower():
            return QualityResult(False, "WARN", "market_snapshot 使用 fallback，僅允許分析，不可直接控制下單", {"rows": rows, "source_level": source_level})
        return QualityResult(True, "OK", "market_snapshot 品質通過", {"rows": rows, "source_level": source_level})
