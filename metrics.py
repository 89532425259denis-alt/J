# -*- coding: utf-8 -*-
"""metrics — аналитика генераций (LLM-вызовы, подгонка страниц, итоги работ).

Пишет JSONL и умеет считать сводку: успешность моделей, среднее число
итераций подгонки, точность попадания в страницы.
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass, field, asdict
from typing import Optional

_METRICS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "metrics.jsonl")
_MAX_BYTES = 10 * 1024 * 1024


@dataclass
class GenerationMetrics:
    """Одно событие аналитики. Все поля опциональны, кроме event."""
    event: str = "generic"                # llm_call | work | image_search | ...
    model_key: Optional[str] = None
    provider: Optional[str] = None
    success: Optional[bool] = None
    duration_seconds: Optional[float] = None
    note: str = ""
    doc_type: Optional[str] = None
    topic: Optional[str] = None
    pages: Optional[int] = None
    final_pages: Optional[int] = None
    iterations: Optional[int] = None
    user_id: Optional[int] = None
    ts: float = field(default_factory=time.time)


class Metrics:
    def __init__(self, path: str = _METRICS_FILE) -> None:
        self.path = path

    def log(self, m: GenerationMetrics) -> None:
        try:
            if os.path.exists(self.path) and os.path.getsize(self.path) > _MAX_BYTES:
                with open(self.path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                with open(self.path, "w", encoding="utf-8") as f:
                    f.writelines(lines[len(lines) // 2:])
            record = {k: v for k, v in asdict(m).items() if v is not None and v != ""}
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(json.dumps(record, ensure_ascii=False) + "\n")
        except Exception:
            pass  # аналитика никогда не ломает генерацию

    def summary(self, days: int = 7) -> dict:
        cutoff = time.time() - days * 86400
        total = ok = 0
        page_hits = page_events = 0
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                for line in f:
                    try:
                        r = json.loads(line)
                    except Exception:
                        continue
                    if float(r.get("ts", 0)) < cutoff:
                        continue
                    total += 1
                    if r.get("success"):
                        ok += 1
                    if r.get("event") == "work" and r.get("pages") and r.get("final_pages"):
                        page_events += 1
                        if abs(int(r["final_pages"]) - int(r["pages"])) <= 1:
                            page_hits += 1
        except FileNotFoundError:
            pass
        except Exception:
            pass
        return {
            "total": total,
            "success": ok,
            "page_accuracy": (page_hits / page_events) if page_events else None,
        }


_metrics_singleton: Optional[Metrics] = None


def get_metrics() -> Metrics:
    global _metrics_singleton
    if _metrics_singleton is None:
        _metrics_singleton = Metrics()
    return _metrics_singleton
