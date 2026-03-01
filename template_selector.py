"""
Выбор лучшего шаблона презентации из Google Sheets через LLM.

Входные параметры:
  - topic:        тема презентации
  - content_text: первичный текст/контент (опционально)
  - prompt:       дополнительный промт/контекст (опционально)
  - style:        желаемый стиль (minimalism, GSB, Axenix...) (опционально)
  - theme:        тема оформления (dark/light) (опционально)

Список шаблонов берётся из Google Sheets (TEMPLATES_SHEET_URL в .env).
Формат таблицы: link | style | theme | num_slides | description

Экспортирует:
  select_template(...) -> str   # URL выбранного шаблона

CLI:
  python template_selector.py --topic "..." [--style minimalism] [--theme dark]
"""

import io
import json
import logging
import os
import re
from pathlib import Path
from string import Template
from typing import Any, Dict, List, Optional

import httpx
import pandas as pd
import requests
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

log = logging.getLogger(__name__)

PROMPTS_DIR = Path(__file__).parent / "prompts"

# ── Конфиг из .env ─────────────────────────────────────────────────────────────
LLM_BASE_URL        = os.getenv("LLM_BASE_URL", "http://172.28.4.29:8000/v1")
LLM_API_KEY         = os.getenv("LLM_API_KEY", "dummy")
LLM_MODEL           = os.getenv("LLM_MODEL", "/model")
LLM_TIMEOUT         = float(os.getenv("LLM_TIMEOUT", "600"))
TEMPLATES_SHEET_URL = os.getenv("TEMPLATES_SHEET_URL", "")
FALLBACK_URL        = os.getenv("FALLBACK_TEMPLATE_URL", "")


# ── LLM client ─────────────────────────────────────────────────────────────────
def _build_client() -> OpenAI:
    return OpenAI(
        base_url=LLM_BASE_URL,
        api_key=LLM_API_KEY,
        timeout=httpx.Timeout(timeout=LLM_TIMEOUT, connect=min(60.0, LLM_TIMEOUT)),
    )


# ── Промты ─────────────────────────────────────────────────────────────────────
def _load_prompt(filename: str, **kwargs) -> str:
    path = PROMPTS_DIR / filename
    text = path.read_text(encoding="utf-8").strip()
    return Template(text).safe_substitute(**kwargs) if kwargs else text


# ── Utils ───────────────────────────────────────────────────────────────────────
def _normalize(x: Any) -> str:
    return str(x).strip() if x is not None else ""


def _is_url(s: str) -> bool:
    s = s.strip()
    return s.startswith("http://") or s.startswith("https://")


def _parse_json(raw: str) -> Any:
    raw = raw.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"```(?:json)?\s*([\s\S]+?)```", raw)
        if match:
            return json.loads(match.group(1).strip())
        raise


# ── Загрузка таблицы из Google Sheets ──────────────────────────────────────────
def _extract_sheet_id(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        raise ValueError(f"Не удалось извлечь ID таблицы из ссылки: {url}")
    return match.group(1)


def _download_templates_df(sheets_url: str) -> pd.DataFrame:
    sheet_id = _extract_sheet_id(sheets_url)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=0"
    log.info("Скачиваю список шаблонов из Google Sheets...")
    resp = requests.get(export_url, timeout=30)
    resp.raise_for_status()
    return pd.read_csv(io.StringIO(resp.text))


def _filter_df(
    df: pd.DataFrame,
    style: Optional[str] = None,
    theme: Optional[str] = None,
) -> pd.DataFrame:
    """
    Фильтрует DataFrame по style и theme через pandas.
    Если после фильтра строк нет — возвращает оригинал.
    """
    cols = {c.lower().strip(): c for c in df.columns}
    style_col = cols.get("style")
    theme_col = cols.get("theme")
    result = df.copy()

    if theme and theme_col:
        filtered = result[result[theme_col].str.strip().str.lower() == theme.lower()]
        if not filtered.empty:
            log.info(f"Pandas-фильтр theme={theme!r}: {len(filtered)}/{len(result)} строк")
            result = filtered
        else:
            log.warning(f"Pandas-фильтр theme={theme!r}: 0 совпадений — фильтр не применяется")

    if style and style_col:
        filtered = result[
            result[style_col].str.strip().str.lower().str.contains(style.lower(), na=False)
        ]
        if not filtered.empty:
            log.info(f"Pandas-фильтр style={style!r}: {len(filtered)}/{len(result)} строк")
            result = filtered
        else:
            log.warning(f"Pandas-фильтр style={style!r}: 0 совпадений — фильтр не применяется")

    return result if not result.empty else df


def _parse_templates(df: pd.DataFrame) -> List[Dict[str, str]]:
    cols = {c.lower().strip(): c for c in df.columns}

    link_col  = cols.get("link")
    style_col = cols.get("style")
    theme_col = cols.get("theme")
    desc_col  = cols.get("description")

    if not link_col:
        raise ValueError("В таблице не найдена колонка 'link'")

    templates: List[Dict[str, str]] = []
    for _, row in df.iterrows():
        link = _normalize(row[link_col])
        if not _is_url(link):
            continue
        templates.append({
            "link":        link,
            "style":       _normalize(row[style_col]) if style_col else "",
            "theme":       _normalize(row[theme_col]) if theme_col else "",
            "description": _normalize(row[desc_col])  if desc_col  else "",
        })

    if not templates:
        raise ValueError("В таблице не найдено ни одной валидной ссылки (http/https).")

    return templates


# ── LLM выбор шаблона ──────────────────────────────────────────────────────────
def _call_llm_select(
    client: OpenAI,
    topic: str,
    content_summary: str,
    templates: List[Dict[str, str]],
    style: Optional[str] = None,
    theme: Optional[str] = None,
    extra_prompt: Optional[str] = None,
) -> str:
    templates_block = "\n".join(
        f"{i+1}) link: {t['link']}\n   style: {t['style'] or '-'}"
        f"\n   theme: {t['theme'] or '-'}\n   description: {t['description'] or '-'}"
        for i, t in enumerate(templates[:60])
    )

    system_prompt = _load_prompt("template_selector_system.txt")
    user_prompt = _load_prompt(
        "template_selector_user.txt",
        topic           = topic,
        style_hint      = f"Желаемый стиль: {style}\n" if style else "",
        theme_hint      = f"Желаемая тема оформления: {theme}\n" if theme else "",
        extra_prompt    = f"Дополнительный контекст: {extra_prompt}\n" if extra_prompt else "",
        content_summary = content_summary,
        templates_block = templates_block,
    )

    resp = client.chat.completions.create(
        model=LLM_MODEL,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_prompt},
        ],
        temperature=0.2,
    )
    raw = resp.choices[0].message.content.strip()

    # Иногда модель возвращает просто ссылку
    if _is_url(raw) and "{" not in raw:
        return raw

    data = _parse_json(raw)
    return _normalize(data.get("selected", ""))


# ── Публичный API ───────────────────────────────────────────────────────────────
def select_template(
    topic: str,
    content_text: str = "",
    prompt: str = "",
    style: str = "",
    theme: str = "",
    sheets_url: str = "",
    fallback_url: str = "",
) -> str:
    """
    Выбирает лучший шаблон из Google Sheets через LLM.

    Args:
        topic:        тема презентации
        content_text: первичный текст/контент презентации (опционально)
        prompt:       дополнительный промт/контекст для выбора (опционально)
        style:        желаемый стиль (minimalism, GSB, Axenix...) (опционально)
        theme:        тема оформления (dark/light) (опционально)
        sheets_url:   URL Google Sheets с шаблонами (по умолчанию из .env)
        fallback_url: fallback URL при ошибке (по умолчанию из .env)

    Returns:
        URL выбранного шаблона Google Drive
    """
    _sheets_url = sheets_url  or TEMPLATES_SHEET_URL
    _fallback   = fallback_url or FALLBACK_URL

    if not _sheets_url:
        log.warning("TEMPLATES_SHEET_URL не задан, возвращаю fallback")
        return _fallback

    try:
        df = _download_templates_df(_sheets_url)
        log.info(f"Загружено шаблонов из таблицы: {len(df)}")

        # Фильтрация в pandas — строго по style и theme
        df_filtered = _filter_df(df, style=style or None, theme=theme or None)
        templates   = _parse_templates(df_filtered)
        valid_links = {t["link"] for t in templates}

        log.info(f"После фильтрации: {len(templates)} шаблонов → отправляю в LLM")
        for t in templates:
            log.info(f"  [{t['style']:12} / {t['theme']:5}] {t['link'][:70]}")

        content_summary = content_text[:3000] if content_text else ""

        client   = _build_client()
        selected = _call_llm_select(
            client          = client,
            topic           = topic,
            content_summary = content_summary,
            templates       = templates,
            style           = style   or None,
            theme           = theme   or None,
            extra_prompt    = prompt  or None,
        )
        log.info(f"LLM вернул: {selected!r}")

        # Жёсткая проверка: URL должен быть из нашего отфильтрованного списка
        if selected and selected in valid_links:
            log.info(f"Выбран шаблон (из списка): {selected}")
            return selected

        if selected and _is_url(selected):
            log.warning(f"LLM вернул URL не из нашего списка: {selected!r} — беру первый из отфильтрованных")
        else:
            log.warning(f"LLM вернул невалидный URL: {selected!r} — беру первый из отфильтрованных")

        return templates[0]["link"]

    except Exception as e:
        log.error(f"Ошибка выбора шаблона: {e}")
        return _fallback


# ── CLI ─────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    ap = argparse.ArgumentParser(description="Выбор шаблона презентации")
    ap.add_argument("--topic",      default="", help="Тема презентации")
    ap.add_argument("--content",    default="", help="Путь к файлу с текстом контента (опционально)")
    ap.add_argument("--prompt",     default="", help="Дополнительный промт (опционально)")
    ap.add_argument("--style",      default="", help="Желаемый стиль (minimalism, GSB...)")
    ap.add_argument("--theme",      default="", help="Тема оформления (dark/light)")
    ap.add_argument("--sheets-url", default="", help="URL Google Sheets с шаблонами")
    ap.add_argument("--fallback",   default="", help="Fallback URL при ошибке")
    args = ap.parse_args()

    content_text = ""
    if args.content and os.path.isfile(args.content):
        with open(args.content, encoding="utf-8") as f:
            content_text = f.read()

    url = select_template(
        topic        = args.topic,
        content_text = content_text,
        prompt       = args.prompt,
        style        = args.style,
        theme        = args.theme,
        sheets_url   = args.sheets_url,
        fallback_url = args.fallback,
    )
    print(url)