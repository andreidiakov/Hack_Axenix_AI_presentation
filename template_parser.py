"""
Динамическое извлечение структуры из PPTX-шаблона.

Заменяет фиксированный structure.json — структура строится каждый раз
заново из входящей презентации.

Поддерживает два типа шаблонов:
  1. С плейсхолдерами {{FIELD}} — извлекаются напрямую из XML
  2. Без маркеров        — LLM определяет, какие тексты нужно заменить

Результат build_structure() совместим с форматом structure.json
и передаётся в agent_system.py и generation_pres.py.
"""

import re
import json
import zipfile
import logging
from lxml import etree

log = logging.getLogger(__name__)

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

PLACEHOLDER_RE = re.compile(r'\{\{[^}]+\}\}')
SLIDE_FILE_RE  = re.compile(r'^ppt/slides/slide(\d+)\.xml$')


# ── Парсинг XML слайдов ────────────────────────────────────────────────────────

def _extract_texts(slide_xml: bytes) -> list[str]:
    """Извлекает все непустые текстовые строки из XML одного слайда."""
    tree = etree.fromstring(slide_xml)
    texts = []
    for para in tree.iter(f'{{{NS_A}}}p'):
        parts = []
        for run in para.findall(f'{{{NS_A}}}r'):
            t = run.find(f'{{{NS_A}}}t')
            if t is not None and t.text:
                parts.append(t.text)
        text = ''.join(parts).strip()
        if text:
            texts.append(text)
    return texts


def _parse_raw_slides(pptx_path: str) -> list[dict]:
    """
    Читает PPTX, возвращает сырые данные по каждому слайду:
      - slide_index  : 0-based
      - texts        : все текстовые строки слайда
      - placeholders : найденные {{...}} (дедуплицированы, порядок сохранён)
    """
    with zipfile.ZipFile(pptx_path, 'r') as z:
        files = {name: z.read(name) for name in z.namelist()}

    slide_nums = sorted(
        int(m.group(1))
        for name in files
        if (m := SLIDE_FILE_RE.match(name))
    )

    raw = []
    for num in slide_nums:
        xml   = files[f'ppt/slides/slide{num}.xml']
        texts = _extract_texts(xml)
        all_text = ' '.join(texts)
        placeholders = list(dict.fromkeys(PLACEHOLDER_RE.findall(all_text)))

        log.debug(
            f"Слайд {num}: {len(texts)} текстов, "
            f"{len(placeholders)} {{{{}}}} плейсхолдеров: {placeholders}"
        )
        raw.append({
            "slide_index":  num - 1,
            "texts":        texts,
            "placeholders": placeholders,
        })

    log.info(f"Прочитано {len(raw)} слайдов из {pptx_path}")
    return raw


# ── LLM-анализ шаблона ────────────────────────────────────────────────────────

def _build_llm_description(raw_slides: list[dict]) -> str:
    """Формирует текстовое описание слайдов для LLM."""
    parts = []
    for s in raw_slides:
        idx = s["slide_index"]
        if s["placeholders"]:
            ph = ", ".join(s["placeholders"])
            parts.append(
                f"Слайд {idx + 1} (индекс {idx}):\n"
                f"  Найдены плейсхолдеры {{{{...}}}}: {ph}"
            )
        else:
            texts_preview = "\n    ".join(f'"{t}"' for t in s["texts"][:12])
            parts.append(
                f"Слайд {idx + 1} (индекс {idx}):\n"
                f"  Текстовые блоки (без {{{{}}}} маркеров):\n    {texts_preview}"
            )
    return "\n\n".join(parts)


def _analyze_with_llm(raw_slides: list[dict], call_llm, load_prompt, parse_json_safe) -> list[dict]:
    """
    Вызывает LLM для анализа всех слайдов шаблона.
    Возвращает список analyzed-слайдов с полями:
      slide_index, slide_type, description, replacements, list_fields
    """
    system_prompt = load_prompt("template_analyzer_system.txt")
    user_prompt   = load_prompt(
        "template_analyzer_user.txt",
        slide_count = len(raw_slides),
        slides_desc = _build_llm_description(raw_slides),
    )

    log.info("TemplateAnalyzer: отправляю запрос к LLM...")
    raw = call_llm(system_prompt, user_prompt, temperature=0.2)
    log.debug(f"TemplateAnalyzer ответ:\n{raw}")

    data = parse_json_safe(raw)
    analyzed = data.get("slides", [])
    log.info(f"TemplateAnalyzer: получил описание {len(analyzed)} слайдов")
    return analyzed


# ── Основная функция ──────────────────────────────────────────────────────────

def build_structure(pptx_path: str, call_llm, load_prompt, parse_json_safe) -> dict:
    """
    Строит динамическую структуру шаблона из PPTX.

    Алгоритм:
      1. Парсим PPTX — получаем тексты и {{...}} плейсхолдеры каждого слайда
      2. Передаём всё в LLM — он именует типы слайдов, описывает их назначение,
         а для слайдов без {{}} — определяет какие тексты являются плейсхолдерами
      3. Возвращаем structure-dict, совместимый с generation_pres.build_presentation()

    Формат результата:
    {
        "slides": [
            {
                "slide_index":  0,
                "slide_type":   "TITLE",
                "description":  "Титульный слайд",
                "replacements": {"{{TITLE_TITLE}}": "", "{{SUBTITLE}}": ""},
                "list_fields":  []
            },
            ...
        ]
    }
    """
    raw_slides = _parse_raw_slides(pptx_path)
    analyzed   = _analyze_with_llm(raw_slides, call_llm, load_prompt, parse_json_safe)

    structure_slides = []
    for item in analyzed:
        idx        = item.get("slide_index", 0)
        slide_type = item.get("slide_type", f"SLIDE_{idx}")
        desc       = item.get("description", "")
        list_fields = item.get("list_fields", [])

        # replacements: принимаем и dict {key: ""} и list [key, ...]
        raw_reps = item.get("replacements", {})
        if isinstance(raw_reps, list):
            replacements = {k: "" for k in raw_reps}
        else:
            replacements = {k: "" for k in raw_reps}

        structure_slides.append({
            "slide_index":  idx,
            "slide_type":   slide_type,
            "description":  desc,
            "replacements": replacements,
            "list_fields":  list_fields,
        })

    structure = {"slides": structure_slides}
    log.info(f"build_structure: построена структура из {len(structure_slides)} слайдов")
    return structure
