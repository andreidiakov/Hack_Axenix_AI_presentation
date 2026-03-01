"""
Агентная система генерации content.json для презентации.

Пайплайн:
  1. PlannerAgent  — выбирает типы слайдов и их порядок под тему
  2. WriterAgent   — генерирует контент каждого слайда (вызывается N раз)
  3. Assembler     — собирает итоговый content.json

Вход  : тема (строка) + динамическая структура шаблона (из template_parser)
Выход : content.json → затем передаётся в generation_pres.py
"""

import concurrent.futures
import json
import logging
import os
import re
from pathlib import Path
from string import Template

import httpx
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

# ── Логирование ────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── LLM client ────────────────────────────────────────────────────────────────

_LLM_BASE_URL = os.getenv("LLM_BASE_URL", "http://172.28.4.29:8000/v1")
_LLM_API_KEY  = os.getenv("LLM_API_KEY", "dummy")
_LLM_TIMEOUT  = float(os.getenv("LLM_TIMEOUT", "600"))
MODEL_NAME    = os.getenv("LLM_MODEL", "/model")

client = OpenAI(
    base_url=_LLM_BASE_URL,
    api_key=_LLM_API_KEY,
    timeout=httpx.Timeout(timeout=_LLM_TIMEOUT, connect=min(60.0, _LLM_TIMEOUT)),
)

PROMPTS_DIR = Path(__file__).parent / "prompts"
DATA_DIR    = Path(__file__).parent / "data"


def load_prompt(filename: str, **kwargs) -> str:
    """
    Загружает промпт из файла.
    Если переданы kwargs — подставляет переменные через string.Template.
    Использует safe_substitute: незаполненные ${var} остаются как есть.
    """
    path = PROMPTS_DIR / filename
    log.debug(f"Загружаю промпт: {filename}")
    text = path.read_text(encoding="utf-8").strip()
    return Template(text).safe_substitute(**kwargs) if kwargs else text


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


# ── Конвертация динамической структуры в схему агентов ────────────────────────

def structure_to_schema(structure: dict) -> dict:
    """
    Преобразует structure (формат template_parser) в рабочую схему для агентов.

    Вход  (structure["slides"][i]):
      slide_index, slide_type, description, replacements, list_fields

    Выход (schema[slide_type]):
      description  : str
      slide_index  : int        — индекс в шаблоне (для generation_pres)
      fields       : list[str]  — ключи для замены
      list_fields  : list[str]  — ключи, ожидающие список буллетов
    """
    schema: dict = {}
    type_count: dict[str, int] = {}
    for s in structure.get("slides", []):
        base  = s["slide_type"]
        count = type_count.get(base, 0) + 1
        type_count[base] = count
        # Первое вхождение — без суффикса; последующие — _2, _3 ...
        key = base if count == 1 else f"{base}_{count}"
        schema[key] = {
            "description": s.get("description", ""),
            "slide_index": s["slide_index"],
            "fields":      list(s.get("replacements", {}).keys()),
            "list_fields": s.get("list_fields", []),
        }
    return schema


# ── Agent 1: PlannerAgent ─────────────────────────────────────────────────────

def run_planner(topic: str, schema: dict, n_slides: int = 7) -> list[dict]:
    log.info(f"PlannerAgent: планирую {n_slides} слайдов...")

    types_desc = json.dumps(
        {k: v["description"] for k, v in schema.items()},
        ensure_ascii=False, indent=2,
    )

    system_prompt = load_prompt("planner_system.txt", n_slides=n_slides)
    user_prompt   = load_prompt("planner_user.txt", topic=topic, types_desc=types_desc, n_slides=n_slides)

    raw = call_llm(system_prompt, user_prompt, temperature=0.6)
    log.debug(f"PlannerAgent ответ:\n{raw}")

    data = parse_json_safe(raw)
    plan = data.get("plan") or data.get("slides") or []
    plan = [p for p in plan if p.get("slide_type") in schema]
    if len(plan) > n_slides:
        log.warning(f"PlannerAgent вернул {len(plan)} слайдов, обрезаю до {n_slides}")
        plan = plan[:n_slides]
    elif len(plan) < n_slides:
        log.warning(f"PlannerAgent вернул только {len(plan)} слайдов из {n_slides}")

    log.info(f"PlannerAgent: план из {len(plan)} слайдов")
    for i, p in enumerate(plan, 1):
        log.info(f"  {i}. {p['slide_type']:20} — {p['purpose'][:60]}")

    return plan


# ── Agent 2: WriterAgent ─────────────────────────────────────────���────────────

def run_writer(topic: str,
               plan_context: str,
               slide_type: str,
               purpose: str,
               schema: dict) -> dict:
    log.info(f"WriterAgent: генерирую слайд {slide_type} — {purpose[:50]}...")

    slide_info  = schema[slide_type]
    fields      = slide_info["fields"]
    list_fields = list(slide_info["list_fields"])

    # Авто-определяем list-поля по имени, если schema не указала явно:
    # {{ITEMS}}, {{LEFT_ITEMS}}, {{RIGHT_ITEMS}} и т.п. — всегда списки
    _auto = [f for f in fields
             if f.strip('{}').strip() == 'ITEMS'
             or f.strip('{}').strip().endswith('_ITEMS')]
    for f in _auto:
        if f not in list_fields:
            list_fields.append(f)

    fields_lines    = "\n".join(f'  "{k}"' for k in fields)
    list_hint       = load_prompt("writer_list_format_hint.txt") if list_fields else ""
    list_fields_str = ", ".join(list_fields) if list_fields else "нет"

    system_prompt = load_prompt("writer_system.txt")
    user_prompt   = load_prompt(
        "writer_user.txt",
        topic           = topic,
        plan_context    = plan_context,
        slide_type      = slide_type,
        purpose         = purpose,
        description     = slide_info["description"],
        fields_lines    = fields_lines,
        list_fields_str = list_fields_str,
        list_hint       = list_hint,
    )

    raw = call_llm(system_prompt, user_prompt, temperature=0.65)
    log.debug(f"WriterAgent ответ для {slide_type}:\n{raw}")

    data = parse_json_safe(raw)
    reps: dict = data.get("replacements", data)

    for key in fields:
        if key not in reps:
            reps[key] = [] if key in list_fields else ""

    for key in list_fields:
        if isinstance(reps.get(key), str):
            text = reps[key].strip()
            reps[key] = [{"type": "bullet", "value": text}] if text else []

    log.info(f"WriterAgent: слайд {slide_type} готов ({len(reps)} полей)")
    return reps


# ── Main pipeline ─────────────────────────────────────────────────────────────

def generate_content_json(topic: str,
                          structure: dict,
                          output_path: str = None,
                          n_slides: int = 7) -> dict:
    """
    Полный пайплайн: тема + структура шаблона → content.json.

    structure  — результат template_parser.build_structure()
    output_path — путь к файлу; по умолчанию data/content.json
    """
    if output_path is None:
        output_path = str(DATA_DIR / "content.json")

    log.info(f"{'='*60}")
    log.info(f"Тема: {topic}")
    log.info(f"{'='*60}")

    schema = structure_to_schema(structure)
    log.info(f"Схема из шаблона: {list(schema.keys())}")

    # Шаг 1: Планирование
    plan = run_planner(topic, schema, n_slides=n_slides)
    if not plan:
        raise RuntimeError("PlannerAgent вернул пустой план. Проверь соединение с LLM.")

    plan_context = "\n".join(
        f"{i}. {p['slide_type']}: {p['purpose']}"
        for i, p in enumerate(plan, 1)
    )

    # Шаг 2: Параллельная генерация контента слайдов
    log.info(f"WriterAgent: параллельная генерация {len(plan)} слайдов...")

    def _write_one(args):
        i, item = args
        slide_type = item.get("slide_type", "")
        purpose    = item.get("purpose", "")
        log.info(f"[{i}/{len(plan)}] Слайд {slide_type}")
        replacements = run_writer(topic, plan_context, slide_type, purpose, schema)
        return {"slide_type": slide_type, "replacements": replacements}

    n_workers = min(len(plan), 5)
    with concurrent.futures.ThreadPoolExecutor(max_workers=n_workers) as pool:
        slides_out = list(pool.map(_write_one, enumerate(plan, 1)))

    # Шаг 3: Сборка content.json
    log.info(f"Assembler: записываю {output_path}...")
    content = {"slides": slides_out}

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    log.info(f"Готово: {output_path} ({len(slides_out)} слайдов)")
    return content


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    from template_parser import build_structure

    TOPIC         = "Будущее искусственного интеллекта в образовании"
    TEMPLATE_PATH = str(DATA_DIR / "template.pptx")

    log.info(f"Анализирую шаблон: {TEMPLATE_PATH}")
    structure = build_structure(TEMPLATE_PATH, call_llm, load_prompt, parse_json_safe)

    generate_content_json(TOPIC, structure)
