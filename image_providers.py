# -*- coding: utf-8 -*-
"""image_providers — модульный агрегатор поиска изображений.

CallableProvider оборачивает любую async-функцию поиска.
AggregateImageSearcher соблюдает общий бюджет времени, дедуплицирует
результаты, применяет фильтры и «прощает» падение любого провайдера
(graceful degradation) — один упавший источник не срывает генерацию.
"""

from __future__ import annotations

import asyncio
import inspect
import random
import time
from typing import Any, Awaitable, Callable

ImageDict = dict  # {"url": .., "bytes": .., "caption": .., "source": .., "_query": ..}


class CallableProvider:
    """Обёртка над функцией поиска. Функция: (query, count) -> list[dict]."""

    def __init__(self, name: str, fn: Callable[..., Awaitable[list]]) -> None:
        self.name = name
        self.fn = fn

    async def search(self, query: str, count: int, timeout: float = 20.0) -> list[ImageDict]:
        try:
            if inspect.iscoroutinefunction(self.fn):
                coro = self.fn(query, count)
            else:
                coro = asyncio.to_thread(self.fn, query, count)
            result = await asyncio.wait_for(coro, timeout=timeout)
            if not result:
                return []
            out = []
            for item in result:
                if isinstance(item, dict):
                    item.setdefault("_provider", self.name)
                    item.setdefault("_query", query)
                    out.append(item)
            return out
        except (asyncio.TimeoutError, Exception) as e:
            print(f"[IMG:{self.name}] пропущен: {type(e).__name__}: {str(e)[:120]}")
            return []


def _img_key(img: ImageDict) -> str:
    return str(img.get("url") or img.get("source") or id(img))


class AggregateImageSearcher:
    """Опрашивает провайдеры по очереди, пока не наберёт count изображений.

    - deadline_sec: общий бюджет времени на весь поиск.
    - result_filter(img) -> bool: пропускать ли изображение.
    - should_avoid(key) -> bool: анти-повтор (например, картинки прошлой работы).
    - avoid_key(img) -> str: ключ для анти-повтора.
    """

    def __init__(
        self,
        providers: list[CallableProvider],
        deadline_sec: float = 45.0,
        rng: random.Random | None = None,
    ) -> None:
        self.providers = list(providers or [])
        self.deadline_sec = max(5.0, float(deadline_sec))
        self.rng = rng or random.Random()

    async def _collect(
        self,
        queries: list[str],
        count: int,
        *,
        result_filter: Callable[[ImageDict], bool] | None = None,
        should_avoid: Callable[[str], bool] | None = None,
        avoid_key: Callable[[ImageDict], str] | None = None,
        already: list[ImageDict] | None = None,
    ) -> list[ImageDict]:
        start = time.monotonic()
        found: list[ImageDict] = list(already or [])
        seen: set[str] = {_img_key(i) for i in found}

        for query in queries:
            if len(found) >= count:
                break
            remaining = self.deadline_sec - (time.monotonic() - start)
            if remaining <= 1.0:
                break
            need = count - len(found)
            per_provider_timeout = min(20.0, remaining)

            # FIX: провайдеры опрашиваются ПАРАЛЛЕЛЬНО (было — по очереди),
            # чтобы один зависший/медленный источник не съедал весь
            # deadline_sec и не оставлял остальных без шанса ответить.
            tasks = [
                provider.search(query, need, timeout=per_provider_timeout)
                for provider in self.providers
            ]
            try:
                provider_results = await asyncio.wait_for(
                    asyncio.gather(*tasks, return_exceptions=True),
                    timeout=per_provider_timeout + 2.0,
                )
            except asyncio.TimeoutError:
                provider_results = []

            for results in provider_results:
                if isinstance(results, BaseException):
                    continue
                if not results:
                    continue
                for img in results:
                    if len(found) >= count:
                        break
                    key = _img_key(img)
                    if key in seen:
                        continue
                    if result_filter is not None:
                        try:
                            if not result_filter(img):
                                continue
                        except Exception:
                            continue
                    if should_avoid is not None and avoid_key is not None:
                        try:
                            if should_avoid(avoid_key(img)):
                                continue
                        except Exception:
                            pass
                    seen.add(key)
                    found.append(img)
        return found

    async def search_all(
        self,
        queries: list[str],
        count: int,
        *,
        result_filter: Callable[[ImageDict], bool] | None = None,
        should_avoid: Callable[[str], bool] | None = None,
        avoid_key: Callable[[ImageDict], str] | None = None,
    ) -> list[ImageDict]:
        return await self._collect(
            list(queries or []), max(0, int(count)),
            result_filter=result_filter,
            should_avoid=should_avoid,
            avoid_key=avoid_key,
        )

    async def refill(
        self,
        queries: list[str],
        images: list[ImageDict],
        count: int,
        *,
        result_filter: Callable[[ImageDict], bool] | None = None,
    ) -> list[ImageDict]:
        """Добор без анти-повтора: разрешаем ранее использованные картинки."""
        if len(images) >= count:
            return images
        return await self._collect(
            list(queries or []), max(0, int(count)),
            result_filter=result_filter,
            already=list(images),
        )
