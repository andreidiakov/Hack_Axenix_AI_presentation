"""
Точка входа: тема → content.json → result.pptx

Пайплайн:
  1. Сервис приоритезации (заглушка) → ссылка на шаблон Google Drive
  2. Скачиваем шаблон
  3. Анализируем шаблон → динамическая структура (template_parser)
  4. Агенты генерируют content.json под эту структуру
  5. Собираем PPTX

Запуск:
  python main.py
  python main.py "Ваша тема"
"""

import sys
import json
import logging
import concurrent.futures

from agent_system import generate_content_json, call_llm, load_prompt, parse_json_safe
from generation_pres import build_presentation
from google_drive import download_template
from template_parser import build_structure

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

TOPIC             = "Будущее искусственного интеллекта в образовании"
FALLBACK_TEMPLATE = "test.pptx"
STRUCTURE_PATH    = "structure.json"


# ── Заглушка сервиса приоритезации ────────────────────────────────────────────

def get_template_link() -> str:
    """
    Заглушка сервиса приоритезации.

    TODO: заменить на HTTP-запрос к реальному сервису.
    Сервис определяет нужный шаблон и возвращает ссылку на Google Drive.
    """
    log.info("Prioritization stub: возвращаю ссылку на шаблон...")
    return "https://docs.google.com/presentation/d/1ead93ZPCCb0IxmGTqBWOMa5MyozVsMzF/edit?usp=drive_link&ouid=102886056468162060057&rtpof=true&sd=true"


# ── Main ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    topic = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else TOPIC
    log.info(f"Тема: {topic}")

    # Шаг 1: получаем ссылку на шаблон (параллельно — когда сервис будет реальным
    # и будет делать тяжёлую работу, тут можно добавить что-то параллельное)
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future_link = executor.submit(get_template_link)
        gdrive_link = future_link.result()
    log.info(f"Шаблон: {gdrive_link}")

    # Шаг 2: скачиваем шаблон (fallback → локальный файл)
    try:
        template_path = download_template(gdrive_link, local_path="template.pptx")
    except Exception as e:
        log.warning(f"Google Drive недоступен ({e}), использую локальный шаблон: {FALLBACK_TEMPLATE}")
        template_path = FALLBACK_TEMPLATE

    # Шаг 3: анализируем шаблон → строим динамическую структуру
    log.info(f"Анализирую шаблон: {template_path}")
    structure = build_structure(template_path, call_llm, load_prompt, parse_json_safe)

    # Сохраняем для отладки
    with open(STRUCTURE_PATH, "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)
    log.info(f"Структура → {STRUCTURE_PATH}")

    # Шаг 4: агенты генерируют content.json под структуру конкретного шаблона
    generate_content_json(topic, structure, output_path="content.json")

    # Шаг 5: собираем PPTX
    log.info("Собираю презентацию...")
    build_presentation(
        template_path  = template_path,
        structure_path = STRUCTURE_PATH,
        content_path   = "content.json",
        output_path    = "result.pptx",
    )

    log.info("Готово → result.pptx")
