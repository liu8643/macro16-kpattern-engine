# -*- coding: utf-8 -*-
"""單一策略設定載入器，避免門檻散落在 UI / Decision / Ranking。"""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict

DEFAULT_STRATEGY_CONFIG: Dict[str, Any] = {
    "rr_live_min": 1.5,
    "rsi_max": 72,
    "price_deviation_max": 0.03,
    "wave_trade_score_min": 82,
    "model_score_min": 82,
    "liquidity_status_required": "PASS",
}

class StrategyConfigLoader:
    _cache: Dict[str, Any] | None = None
    def __init__(self, path: str | Path = "strategy_config.json"):
        self.path = Path(path)
    def load(self, force: bool = False) -> Dict[str, Any]:
        if self.__class__._cache is not None and not force:
            return dict(self.__class__._cache)
        cfg = dict(DEFAULT_STRATEGY_CONFIG)
        if self.path.exists():
            try:
                user_cfg = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(user_cfg, dict): cfg.update(user_cfg)
            except Exception:
                pass
        self.__class__._cache = cfg
        return dict(cfg)
