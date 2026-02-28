"""
Точка входа: тема → content.json → result.pptx

Пайплайн:
  1. template_selector  — выбирает оптимальный шаблон из Google Sheets через LLM
  2. Скачиваем шаблон с Google Drive
  3. Анализируем шаблон → динамическая структура (template_parser)
  4. Агенты генерируют content.json под эту структуру
  5. Собираем PPTX

Запуск:
  python main.py
  python main.py "Ваша тема"
  python main.py "Ваша тема" --style minimalism --theme dark
  python main.py "Ваша тема" --prompt "деловая аудитория, строгий стиль"
"""

import argparse
import json
import logging
import os

from dotenv import load_dotenv

load_dotenv()

from agent_system import generate_content_json, call_llm, load_prompt, parse_json_safe
from generation_pres import build_presentation
from google_drive import download_template
from template_parser import build_structure
from template_selector import select_template

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

DEFAULT_TOPIC          = "Будущее искусственного интеллекта в образовании"
FALLBACK_LOCAL_TEMPLATE = os.getenv("FALLBACK_LOCAL_TEMPLATE", "test.pptx")
STRUCTURE_PATH         = "structure.json"


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Генератор презентаций на основе LLM")
    ap.add_argument("topic",   nargs="*",  default=[],  help="Тема презентации")
    ap.add_argument("--style", default=os.getenv("PRESENTATION_STYLE", ""),
                    help="Желаемый стиль (minimalism, GSB, Axenix...)")
    ap.add_argument("--theme", default=os.getenv("PRESENTATION_THEME", ""),
                    help="Тема оформления (dark/light)")
    ap.add_argument("--prompt", default="",
                    help="Дополнительный контекст для выбора шаблона")
    args = ap.parse_args()

    topic = " ".join(args.topic).strip() if args.topic else DEFAULT_TOPIC
    log.info(f"Тема: {topic}")

    # Шаг 1: выбираем шаблон из Google Sheets через LLM
    log.info("Выбираю шаблон...")
    gdrive_link = select_template(
        topic  = topic,
        style  = args.style,
        theme  = args.theme,
        prompt = args.prompt,
    )
    log.info(f"Шаблон: {gdrive_link}")

    # Шаг 2: скачиваем шаблон (fallback → локальный файл)
    try:
        template_path = download_template(gdrive_link, local_path="template.pptx")
    except Exception as e:
        log.warning(f"Google Drive недоступен ({e}), использую локальный шаблон: {FALLBACK_LOCAL_TEMPLATE}")
        template_path = FALLBACK_LOCAL_TEMPLATE

    # Шаг 3: анализируем шаблон → строим динамическую структуру
    log.info(f"Анализирую шаблон: {template_path}")
    structure = build_structure(template_path, call_llm, load_prompt, parse_json_safe)

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