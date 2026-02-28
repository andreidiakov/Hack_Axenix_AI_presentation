"""
Точка входа: тема → content.json → result.pptx

Пайплайн:
  1. Параллельно:
       - Сервис приоритезации → ссылка на шаблон в Google Drive
       - Агентная система    → генерирует content.json
  2. Скачиваем шаблон с Google Drive
  3. Собираем PPTX

Запуск:
  python main.py
  python main.py "Ваша тема"
"""

import sys
import logging
import concurrent.futures

from agent_system import generate_content_json
from generation_pres import build_presentation
from google_drive import download_template

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

TOPIC = "Будущее искусственного интеллекта в образовании"

# Локальный шаблон — используется если Google Drive недоступен
FALLBACK_TEMPLATE = "test.pptx"


# ── Заглушка сервиса приоритезации ────────────────────────────────────────────

def get_template_link() -> str:
    """
    Заглушка сервиса приоритезации.

    TODO: заменить на HTTP-запрос к реальному сервису.
    Сервис определяет нужный шаблон и возвращает ссылку на Google Drive.
    """
    log.info("Prioritization stub: возвращаю ссылку на шаблон...")

    # Заглушка — фиксированная ссылка
    return "https://docs.google.com/presentation/d/1WwA5CecwJYxEenZIIrUXyytPBJBuv6_B/edit"


# ── Main ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    topic = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else TOPIC
    log.info(f"Тема: {topic}")

    # Шаг 1: параллельно запускаем приоритезацию и генерацию контента
    log.info("Запускаю параллельно: приоритезацию и агентную систему...")
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        future_link    = executor.submit(get_template_link)
        future_content = executor.submit(generate_content_json, topic, "content.json")

        gdrive_link = future_link.result()
        log.info(f"Приоритезация готова: {gdrive_link}")

        future_content.result()  # ждём завершения агентов
        log.info("Агентная система: content.json готов")

    # Шаг 2: скачиваем шаблон с Google Drive (fallback → локальный файл)
    try:
        template_path = download_template(gdrive_link, local_path="template.pptx")
    except Exception as e:
        log.warning(f"Google Drive недоступен ({e}), использую локальный шаблон: {FALLBACK_TEMPLATE}")
        template_path = FALLBACK_TEMPLATE

    # Шаг 3: собираем PPTX
    log.info("Собираю презентацию...")
    build_presentation(
        template_path  = template_path,
        structure_path = "structure.json",
        content_path   = "content.json",
        output_path    = "result.pptx",
    )

    log.info("Готово → result.pptx")
