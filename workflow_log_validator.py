# -*- coding: utf-8 -*-
"""Workflow log validator wrapper.
用於驗收：檢查重建排行是否誤跑外部 API/EPS build，或 UI 是否重覆 refresh。
"""
from __future__ import annotations
from pathlib import Path

VIOLATION_KEYWORDS = [
    ("FAST_RANK_REBUILD", "external API"),
    ("FAST_RANK_REBUILD", "EPS BUILD"),
    ("AI_TOP20_VIEW", "rebuild"),
    ("AI_TOP20_VIEW", "update_price_history"),
]

def scan_log_text(text: str) -> list[dict]:
    rows = []
    lines = str(text or "").splitlines()
    for i, line in enumerate(lines, 1):
        for mode, keyword in VIOLATION_KEYWORDS:
            if mode in line and keyword.lower() in line.lower():
                rows.append({"line": i, "mode": mode, "keyword": keyword, "message": line})
    return rows

def scan_workflow_log_for_violations(path: str | Path) -> list[dict]:
    p = Path(path)
    if not p.exists():
        return [{"line": 0, "mode": "FILE", "keyword": "missing", "message": str(p)}]
    return scan_log_text(p.read_text(encoding="utf-8", errors="ignore"))
