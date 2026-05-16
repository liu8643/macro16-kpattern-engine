# -*- coding: utf-8 -*-
"""GTC strategy config singleton.
目前提供安全預設；後續可改讀 JSON，但主程式只從這裡取值。
"""
from __future__ import annotations
import json
from pathlib import Path

DEFAULT_STRATEGY_CONFIG = {
    "rr_live_min": 1.5,
    "rsi_max": 72,
    "price_deviation_max": 0.03,
    "wave_trade_score_min": 82,
    "model_score_min": 82,
    "daily_yahoo_fallback_limit": 120,
    "official_hit_ratio_min": 0.50,
}

class StrategyConfigManager:
    def __init__(self, path: str | None = None):
        self.path = Path(path) if path else Path(__file__).with_name("strategy_config.json")
        self.data = dict(DEFAULT_STRATEGY_CONFIG)
        self.load()
    def load(self):
        try:
            if self.path.exists():
                raw = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    self.data.update(raw)
        except Exception:
            pass
        return self.data
    def get(self, key, default=None):
        return self.data.get(key, default)

STRATEGY_CONFIG = StrategyConfigManager()
