"""
Агентная система генерации content.json для презентации.

Пайплайн:
  1. PlannerAgent  — выбирает типы слайдов и их порядок под тему
  2. WriterAgent   — генерирует контент каждого слайда (вызывается N раз)
  3. Assembler     — собирает итоговый content.json

Вход  : тема (строка)
Выход : content.json → затем передаётся в generation_pres.py
"""

import json
import re
import logging
from pathlib import Path
import httpx
from openai import OpenAI

# ── Логирование ────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── LLM client ────────────────────────────────────────────────────────────────

client = OpenAI(
    base_url="http://172.28.4.29:8000/v1",
    api_key="dummy",
    timeout=httpx.Timeout(timeout=600.0, connect=60.0),
)
MODEL_NAME = "/model"

PROMPTS_DIR = Path(__file__).parent / "prompts"


def load_prompt(filename: str) -> str:
    path = PROMPTS_DIR / filename
    log.debug(f"Загружаю промпт: {path}")
    return path.read_text(encoding="utf-8").strip()


def call_llm(system_prompt: str, user_prompt: str, temperature: float = 0.7) -> str:
    log.debug(f"Запрос к LLM (temperature={temperature})")
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
    )
    return response.choices[0].message.content


def parse_json_safe(raw: str) -> dict | list:
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r'```(?:json)?\s*([\s\S]+?)```', raw)
        if match:
            return json.loads(match.group(1))
        raise


# ── Схема слайдов ─────────────────────────────────────────────────────────────

SLIDE_SCHEMA: dict[str, dict] = {
    "TITLE": {
        "description": "Титульный слайд. Название и подзаголовок.",
        "fields": {
            "{{TITLE_TITLE}}": "string — главный заголовок (≤8 слов)",
            "{{SUBTITLE}}":    "string — подзаголовок / автор / дата",
        },
        "list_fields": [],
    },
    "BULLETS_6": {
        "description": "Слайд с шестью тезисами/фактами.",
        "fields": {
            "{{TITLE_BULLETS_6}}": "string — заголовок слайда",
            "{{BULLET_1}}":        "string — тезис 1",
            "{{BULLET_2}}":        "string — тезис 2",
            "{{BULLET_3}}":        "string — тезис 3",
            "{{BULLET_4}}":        "string — тезис 4",
            "{{BULLET_5}}":        "string — тезис 5",
            "{{BULLET_6}}":        "string — тезис 6",
        },
        "list_fields": [],
    },
    "BULLETS_4": {
        "description": "Слайд с четырьмя ключевыми тезисами.",
        "fields": {
            "{{TITLE_BULLETS_4}}": "string — заголовок слайда",
            "{{BULLET_1}}":        "string — тезис 1",
            "{{BULLET_2}}":        "string — тезис 2",
            "{{BULLET_3}}":        "string — тезис 3",
            "{{BULLET_4}}":        "string — тезис 4",
        },
        "list_fields": [],
    },
    "COMPARE": {
        "description": "Сравнительный слайд: левая и правая колонки.",
        "fields": {
            "{{TITLE_COMPARE}}": "string — заголовок слайда",
            "{{LEFT_TITLE}}":    "string — заголовок левой колонки",
            "{{LEFT_ITEMS}}":    "list   — пункты левой колонки",
            "{{RIGHT_TITLE}}":   "string — заголовок правой колонки",
            "{{RIGHT_ITEMS}}":   "list   — пункты правой колонки",
        },
        "list_fields": ["{{LEFT_ITEMS}}", "{{RIGHT_ITEMS}}"],
    },
    "LEFT_TEXT": {
        "description": "Слайд с заголовком и развёрнутым текстом слева.",
        "fields": {
            "{{TITLE_LEFT_TEXT}}": "string — заголовок слайда",
            "{{ITEMS}}":           "list   — текстовые пункты",
        },
        "list_fields": ["{{ITEMS}}"],
    },
    "RIGHT_TEXT": {
        "description": "Слайд с заголовком и развёрнутым текстом справа.",
        "fields": {
            "{{TITLE_RIGHT_TEXT}}": "string — заголовок слайда",
            "{{ITEMS}}":            "list   — текстовые пункты",
        },
        "list_fields": ["{{ITEMS}}"],
    },
    "THREE_COLUMNS": {
        "description": "Три колонки — структура, сравнение или три аспекта.",
        "fields": {
            "{{TITLE_THREE_COLUMNS}}": "string — заголовок слайда",
            "{{SUBTITLE_COLUMN_1}}":   "string — заголовок колонки 1",
            "{{SUBTITLE_COLUMN_2}}":   "string — заголовок колонки 2",
            "{{SUBTITLE_COLUMN_3}}":   "string — заголовок колонки 3",
            "{{ITEM_COLUMN_1}}":       "string — текст колонки 1",
            "{{ITEM_COLUMN_2}}":       "string — текст колонки 2",
            "{{ITEM_COLUMN_3}}":       "string — текст колонки 3",
        },
        "list_fields": [],
    },
    "TIMELINE": {
        "description": "Временная шкала с пятью этапами/точками.",
        "fields": {
            "{{TITLE_TIMELINE}}": "string — заголовок слайда",
            "{{POINT_1}}":        "string — этап 1",
            "{{POINT_2}}":        "string — этап 2",
            "{{POINT_3}}":        "string — этап 3",
            "{{POINT_4}}":        "string — этап 4",
            "{{POINT_5}}":        "string — этап 5",
        },
        "list_fields": [],
    },
    "CLOSE": {
        "description": "Завершающий слайд: благодарность / призыв к действию.",
        "fields": {
            "{{TITLE_CLOSE}}": "string — финальный заголовок",
            "{{SUBTITLE}}":    "string — подзаголовок / контакты",
        },
        "list_fields": [],
    },
}


# ── Agent 1: PlannerAgent ─────────────────────────────────────────────────────

def run_planner(topic: str) -> list[dict]:
    log.info("PlannerAgent: запускаю планирование структуры...")

    system_prompt = load_prompt("planner_system.txt")

    types_desc = json.dumps(
        {k: v["description"] for k, v in SLIDE_SCHEMA.items()},
        ensure_ascii=False, indent=2,
    )
    user_prompt = f"""ТЕМА: {topic}

ДОСТУПНЫЕ ТИПЫ СЛАЙДОВ:
{types_desc}

Составь план презентации."""

    raw = call_llm(system_prompt, user_prompt, temperature=0.6)
    log.debug(f"PlannerAgent ответ:\n{raw}")

    data = parse_json_safe(raw)
    plan = data.get("plan") or data.get("slides") or []
    plan = [p for p in plan if p.get("slide_type") in SLIDE_SCHEMA]

    log.info(f"PlannerAgent: получил план из {len(plan)} слайдов")
    for i, p in enumerate(plan, 1):
        log.info(f"  {i}. {p['slide_type']:15} — {p['purpose'][:60]}")

    return plan


# ── Agent 2: WriterAgent ──────────────────────────────────────────────────────

def run_writer(topic: str, plan_context: str, slide_type: str, purpose: str) -> dict:
    log.info(f"WriterAgent: генерирую слайд {slide_type} — {purpose[:50]}...")

    system_prompt = load_prompt("writer_system.txt")
    schema = SLIDE_SCHEMA[slide_type]

    fields_lines = "\n".join(
        f'  "{k}": {v}' for k, v in schema["fields"].items()
    )
    list_hint = load_prompt("list_format_hint.txt") if schema["list_fields"] else ""

    user_prompt = f"""ТЕМА ПРЕЗЕНТАЦИИ: {topic}

ПЛАН ПРЕЗЕНТАЦИИ (контекст):
{plan_context}

ТВОЙ СЛАЙД:
  Тип:    {slide_type}
  Задача: {purpose}
  ({schema["description"]})

ПОЛЯ ДЛЯ ЗАПОЛНЕНИЯ:
{fields_lines}

List-поля этого слайда: {schema["list_fields"] if schema["list_fields"] else "нет"}
{list_hint}
Верни JSON с ключом "replacements"."""

    raw = call_llm(system_prompt, user_prompt, temperature=0.65)
    log.debug(f"WriterAgent ответ для {slide_type}:\n{raw}")

    data = parse_json_safe(raw)
    reps: dict = data.get("replacements", data)

    for key in schema["fields"]:
        if key not in reps:
            reps[key] = "" if key not in schema["list_fields"] else []

    for key in schema["list_fields"]:
        if isinstance(reps.get(key), str):
            text = reps[key].strip()
            reps[key] = [{"type": "bullet", "value": text}] if text else []

    log.info(f"WriterAgent: слайд {slide_type} готов ({len(reps)} полей)")
    return reps


# ── Main pipeline ─────────────────────────────────────────────────────────────

def generate_content_json(topic: str, output_path: str = "content.json") -> dict:
    log.info(f"{'='*60}")
    log.info(f"Тема: {topic}")
    log.info(f"{'='*60}")

    # Шаг 1: Планирование
    plan = run_planner(topic)
    if not plan:
        raise RuntimeError("PlannerAgent вернул пустой план. Проверь соединение с LLM.")

    plan_context = "\n".join(
        f"{i}. {p['slide_type']}: {p['purpose']}"
        for i, p in enumerate(plan, 1)
    )

    # Шаг 2: Генерация контента слайдов
    log.info(f"WriterAgent: начинаю генерацию {len(plan)} слайдов...")
    slides_out = []

    for i, item in enumerate(plan, 1):
        slide_type = item.get("slide_type", "")
        purpose    = item.get("purpose", "")
        log.info(f"[{i}/{len(plan)}] Обрабатываю слайд {slide_type}")
        replacements = run_writer(topic, plan_context, slide_type, purpose)
        slides_out.append({
            "slide_type":   slide_type,
            "replacements": replacements,
        })

    # Шаг 3: Сборка content.json
    log.info(f"Assembler: записываю {output_path}...")
    content = {"slides": slides_out}

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    log.info(f"Готово: {output_path} ({len(slides_out)} слайдов)")
    return content


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    TOPIC = "Будущее искусственного интеллекта в образовании"
    generate_content_json(TOPIC)
