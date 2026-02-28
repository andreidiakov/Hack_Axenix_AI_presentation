"""
Точка входа: тема → content.json → result.pptx

Запуск:
  python main.py
  python main.py "Ваша тема"
"""

import sys
from agent_system import generate_content_json
from generation_pres import build_presentation

TOPIC = "Будущее искусственного интеллекта в образовании"

if __name__ == "__main__":
    topic = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else TOPIC

    # Шаг 1: агенты генерируют content.json
    generate_content_json(topic, output_path="content.json")

    # Шаг 2: generation_pres собирает PPTX из шаблона
    build_presentation(
        template_path  = "test.pptx",
        structure_path = "structure.json",
        content_path   = "content.json",
        output_path    = "result.pptx",
    )
