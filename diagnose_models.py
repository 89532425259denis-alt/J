# -*- coding: utf-8 -*-
"""diagnose_models.py — быстрая проверка, реально ли работают API-ключи
моделей, БЕЗ запуска телеграм-бота. Запускать на том же сервере/окружении,
где крутится основной процесс (важно: та же .env / TOKENS / bot_config.json).

Использование:
    python3 diagnose_models.py

Выведет для каждой модели: есть ли ключ, и реальный ответ/ошибку API.
"""
import asyncio
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gost_assistant as ga  # noqa: E402


async def check_model(key: str, info: dict) -> None:
    has_key = bool(info.get("api_key"))
    print(f"\n=== {key} ({info.get('name')}) ===")
    print(f"  base_url: {info.get('base_url')}")
    print(f"  model:    {info.get('model')}")
    print(f"  api_key:  {'ЕСТЬ (' + info['api_key'][:6] + '...)' if has_key else '❌ ПУСТОЙ'}")
    if not has_key:
        print("  → пропускаю вызов: ключ не задан в окружении")
        return
    try:
        text = await ga.chat_with_model(
            info,
            [{"role": "user", "content": "Ответь одним словом: тест"}],
            max_tokens=20,
        )
        if text and text.strip():
            print(f"  ✅ ОТВЕТИЛА: {text.strip()[:80]!r}")
        else:
            print("  ⚠️ Пустой ответ (без исключения)")
    except Exception as e:
        print(f"  ❌ ОШИБКА: {type(e).__name__}: {e}")


async def main() -> None:
    print("Проверяю AI_MODELS...")
    for key, info in ga.AI_MODELS.items():
        await check_model(key, info)

    print("\n=== include_images / image providers ===")
    try:
        imgs = await ga.prepare_work_images(
            "тестовая тема для диагностики",
            "информатика",
            10,
            model_key=ga.FREE_MODEL_KEY,
            image_count=2,
        )
        print(f"  Найдено изображений: {len(imgs)}")
        for im in imgs[:3]:
            print(f"   - {im.get('_provider')}: {str(im.get('url'))[:70]}")
    except Exception as e:
        print(f"  ❌ ОШИБКА поиска изображений: {type(e).__name__}: {e}")


if __name__ == "__main__":
    asyncio.run(main())
