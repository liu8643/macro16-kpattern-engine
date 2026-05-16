# -*- coding: utf-8 -*-
"""GTC data quality gate.
集中檢查每日增量/排行前資料品質，不執行外部 API。
"""
from __future__ import annotations
from dataclasses import dataclass

@dataclass
class DataQualityResult:
    ok: bool
    level: str
    reason: str
    metrics: dict

def check_official_hit_ratio(official_rows: int, master_rows: int, min_ratio: float = 0.50) -> DataQualityResult:
    ratio = float(official_rows or 0) / float(master_rows or 1)
    ok = ratio >= float(min_ratio)
    return DataQualityResult(ok=ok, level="PASS" if ok else "WARN", reason=f"official_hit_ratio={ratio:.2%}", metrics={"official_rows": official_rows, "master_rows": master_rows, "ratio": ratio})

def check_financial_ne_ratio(ne_ratio: float, max_ratio: float = 0.80) -> DataQualityResult:
    ok = float(ne_ratio or 0) < float(max_ratio)
    return DataQualityResult(ok=ok, level="PASS" if ok else "WARN", reason=f"financial_feature_daily NE_ratio={float(ne_ratio or 0):.2%}", metrics={"ne_ratio": float(ne_ratio or 0)})
