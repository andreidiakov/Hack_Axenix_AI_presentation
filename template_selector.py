"""
Выбор лучшего шаблона презентации (ссылка на Google Drive) на основе:
  - content.json (сгенерированного вашим agent_system.py)
  - Excel-файла со списком доступных шаблонов (link/style/color/description)

Результат: печатает В STDOUT ТОЛЬКО выбранную ссылку.
Если ничего не подходит — печатает fallback-ссылку (задаётся параметром).

Пример:
  python template_selector.py \
    --content content.json \
    --excel templates.xlsx \
    --fallback "https://drive.google.com/...." \
    --topic "Будущее искусственного интеллекта в образовании"
"""

import argparse
import json
import logging
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import httpx
import pandas as pd
from openai import OpenAI

# ── Логирование ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── LLM client (как у вас) ─────────────────────────────────────────────────────
def build_client(base_url: str, api_key: str, timeout_s: float) -> OpenAI:
    return OpenAI(
        base_url=base_url,
        api_key=api_key,
        timeout=httpx.Timeout(timeout=timeout_s, connect=min(60.0, timeout_s)),
    )

# ── Utils ──────────────────────────────────────────────────────────────────────
def parse_json_safe(raw: str) -> Any:
    """
    Пытается извлечь JSON:
      - либо raw целиком JSON
      - либо JSON внутри ```json ... ```
    """
    raw = raw.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"```(?:json)?\s*([\s\S]+?)```", raw)
        if match:
            return json.loads(match.group(1).strip())
        # иногда модель возвращает просто ссылку — обработаем отдельно выше
        raise

def normalize_str(x: Any) -> str:
    return str(x).strip() if x is not None else ""

def first_nonempty(*vals: Any) -> str:
    for v in vals:
        s = normalize_str(v)
        if s:
            return s
    return ""

def looks_like_url(s: str) -> bool:
    s = s.strip()
    return s.startswith("http://") or s.startswith("https://")

def extract_first_url(text: str) -> Optional[str]:
    m = re.search(r"(https?://\S+)", text)
    return m.group(1).rstrip(").,;\"'") if m else None

# ── Excel reading ──────────────────────────────────────────────────────────────
COL_ALIASES = {
    "link",
    "style",
    "theme" ,
    "description"
    }

def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    cols_lower = {c.lower().strip(): c for c in df.columns}
    detected: Dict[str, str] = {}

    for key, aliases in COL_ALIASES.items():
        for a in aliases:
            if a.lower().strip() in cols_lower:
                detected[key] = cols_lower[a.lower().strip()]
                break

    # link обязателен
    if "link" not in detected:
        raise ValueError(
            "Не найдена колонка со ссылкой. Ожидаю одну из: "
            + ", ".join(COL_ALIASES["link"])
        )

    # необязательные
    for opt in ["style", "color", "description"]:
        if opt not in detected:
            detected[opt] = None

    return detected

def read_templates_excel(path: str) -> List[Dict[str, str]]:
    df = pd.read_excel(path)
    cols = detect_columns(df)

    templates: List[Dict[str, str]] = []
    for _, row in df.iterrows():
        link = normalize_str(row[cols["link"]])
        if not looks_like_url(link):
            continue

        style = normalize_str(row[cols["style"]]) if cols["style"] else ""
        color = normalize_str(row[cols["color"]]) if cols["color"] else ""
        desc  = normalize_str(row[cols["description"]]) if cols["description"] else ""

        templates.append({
            "link": link,
            "style": style,
            "color": color,
            "description": desc,
        })

    if not templates:
        raise ValueError("В Excel не найдено ни одной валидной ссылки (http/https).")

    return templates

# ── Content summarization ──────────────────────────────────────────────────────
def summarize_content_json(content: Dict[str, Any], max_chars: int = 8000) -> str:
    """
    Делаем компактный контекст для LLM:
    - типы слайдов
    - заголовки/ключевые буллеты (если есть)
    """
    slides = content.get("slides", [])
    parts: List[str] = []
    parts.append(f"Всего слайдов: {len(slides)}")

    for i, s in enumerate(slides, 1):
        st = normalize_str(s.get("slide_type"))
        reps = s.get("replacements", {}) or {}
        # вытащим любые поля, которые выглядят как заголовки/буллеты
        titles = []
        bullets = []
        for k, v in reps.items():
            kk = k.lower()
            if "title" in kk:
                if isinstance(v, str) and v.strip():
                    titles.append(v.strip())
            if "bullet" in kk:
                if isinstance(v, str) and v.strip():
                    bullets.append(v.strip())

        # list-поля (часто ITEMS/LEFT_ITEMS/RIGHT_ITEMS)
        for k, v in reps.items():
            if isinstance(v, list) and v:
                # поддержим варианты: ["a","b"] или [{"type":"bullet","value":"..."}]
                extracted = []
                for item in v[:6]:
                    if isinstance(item, str):
                        extracted.append(item.strip())
                    elif isinstance(item, dict):
                        val = item.get("value")
                        if isinstance(val, str) and val.strip():
                            extracted.append(val.strip())
                if extracted:
                    bullets.extend(extracted)

        line = f"{i}. {st}"
        if titles:
            line += f" | title: {titles[0]}"
        if bullets:
            line += f" | ключевые пункты: " + "; ".join(bullets[:4])
        parts.append(line)

    out = "\n".join(parts)
    return out[:max_chars]

# ── LLM selection ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """Ты — ассистент, который выбирает лучший шаблон презентации.
Дано:
1) Тема/контекст презентации (из content.json)
2) Список доступных шаблонов (ссылка, стиль, цвет, описание)

Задача:
- Выбрать ОДНУ лучшую ссылку на шаблон.
- Если ни один не подходит, вернуть первую в списке.

Правила:
- Учитывай соответствие теме, тону (minimalism / professional / fun / GSB / Axenix), и цветовой схеме (dark/light).
- Если в списке есть несколько подходящих — выбирай самый уместный и универсальный для аудитории.
- Ответ возвращай СТРОГО в JSON:
  {"selected": "<URL>"}
- Никакого дополнительного текста.
"""

def call_llm_select(
    client: OpenAI,
    model: str,
    topic: str,
    content_summary: str,
    templates: List[Dict[str, str]],
    temperature: float = 0.2,
) -> Dict[str, str]:
    # Ограничим число вариантов, чтобы не раздувать промпт (если Excel большой)
    # Сначала оставим все, но если очень много — возьмём первые 60.
    max_templates = 60
    trimmed = templates[:max_templates]

    templates_block = "\n".join(
        f"{idx+1}) link: {t['link']}\n   style: {t['style'] or '-'}\n   color: {t['color'] or '-'}\n   description: {t['description'] or '-'}"
        for idx, t in enumerate(trimmed)
    )

    user_prompt = f"""ТЕМА: {topic}

КОНТЕКСТ ИЗ content.json (кратко):
{content_summary}

ДОСТУПНЫЕ ШАБЛОНЫ:
{templates_block}
"""

    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
    )
    raw = resp.choices[0].message.content.strip()

    # иногда модель может вернуть просто ссылку — подстрахуемся
    if looks_like_url(raw) and raw.count("{") == 0:
        return {"selected": raw, "reason": "model_returned_url"}

    data = parse_json_safe(raw)
    selected = normalize_str(data.get("selected"))
    reason = normalize_str(data.get("reason"))

    return {"selected": selected, "reason": reason}

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", required=True, help="Путь к content.json")
    ap.add_argument("--excel", required=True, help="Путь к Excel с шаблонами")
    ap.add_argument("--fallback", required=True, help="Ссылка, которую печатать если NO_MATCH")
    ap.add_argument("--topic", default="", help="Тема (если не хотите извлекать из контекста)")
    ap.add_argument("--base-url", default=os.getenv("OPENAI_BASE_URL", "http://172.28.4.29:8000/v1"))
    ap.add_argument("--api-key", default=os.getenv("OPENAI_API_KEY", "dummy"))
    ap.add_argument("--model", default=os.getenv("OPENAI_MODEL", "/model"))
    ap.add_argument("--timeout", type=float, default=600.0)
    args = ap.parse_args()

    # 1) читаем content.json
    with open(args.content, "r", encoding="utf-8") as f:
        content = json.load(f)

    # 2) читаем excel
    templates = read_templates_excel(args.excel)

    # 3) готовим контекст
    content_summary = summarize_content_json(content)
    topic = args.topic.strip() or "Презентация (тема не указана явно)"

    # 4) LLM selection
    client = build_client(args.base_url, args.api_key, args.timeout)
    try:
        out = call_llm_select(
            client=client,
            model=args.model,
            topic=topic,
            content_summary=content_summary,
            templates=templates,
        )
        selected = out.get("selected", "").strip()

        if selected == "NO_MATCH" or not looks_like_url(selected):
            # fallback
            print(args.fallback.strip())
        else:
            print(selected)

    except Exception as e:
        # на любой ошибке — безопасный fallback (и в stdout только ссылка)
        log.error(f"Ошибка выбора шаблона: {e}")
        print(args.fallback.strip())

if __name__ == "__main__":
    main()
