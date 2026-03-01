"""
FastAPI-сервер для PrezAI.

Эндпоинты:
  GET  /          → index.html
  GET  /health    → {\"status\": \"ok\"}
  POST /api/chat  → multipart/form-data { text, style?, theme?, slides?, template? } → PPTX-файл (blob)

Запуск:
  python api.py
  uvicorn api:app --host 0.0.0.0 --port 8080 --reload
"""

import json
import logging
import os
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional

import uvicorn
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response

load_dotenv()

from agent_system import call_llm, generate_content_json, load_prompt, parse_json_safe
from generation_pres import build_presentation
from google_drive import download_template
from template_parser import build_structure
from template_selector import select_template

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("api")

BASE_DIR       = Path(__file__).parent
_fallback_name = os.getenv("FALLBACK_LOCAL_TEMPLATE", "test.pptx")
FALLBACK_LOCAL = str(BASE_DIR / _fallback_name)
WORKERS        = int(os.getenv("API_WORKERS", "4"))

app      = FastAPI(title="PrezAI", version="1.0")
executor = ThreadPoolExecutor(max_workers=WORKERS)


# ── Пайплайн (запускается в потоке) ────────────────────────────────────────────
def _run_pipeline(
    text: str,
    style: str,
    theme: str,
    slides: int,
    workdir: Path,
    custom_template_bytes: bytes | None = None,
    team_members: list | None = None,
) -> bytes:
    tag = workdir.name

    if custom_template_bytes:
        # Пользователь загрузил свой шаблон — пропускаем выбор и скачивание
        template_path = str(workdir / "template.pptx")
        with open(template_path, "wb") as f:
            f.write(custom_template_bytes)
        log.info(f"[{tag}] Используется загруженный шаблон ({len(custom_template_bytes)//1024} KB)")
    else:
        # Шаг 1: выбираем шаблон из Google Sheets
        gdrive_link = select_template(topic=text, style=style, theme=theme)
        log.info(f"[{tag}] Шаблон выбран: {gdrive_link}")

        # Шаг 2: скачиваем шаблон (fallback → локальный файл)
        template_local = str(workdir / "template.pptx")
        try:
            template_path = download_template(gdrive_link, local_path=template_local)
        except Exception as e:
            log.warning(f"[{tag}] Drive FAIL для {gdrive_link!r} — {e}, пробую локальный fallback")
            if not Path(FALLBACK_LOCAL).exists():
                raise RuntimeError(
                    f"Шаблон с Google Drive недоступен ({e}), "
                    f"и локальный fallback '{FALLBACK_LOCAL}' не найден. "
                    "Убедитесь, что файл шаблона открыт по ссылке (Anyone with the link → Viewer)."
                )
            template_path = FALLBACK_LOCAL
            log.info(f"[{tag}] Использую локальный шаблон: {FALLBACK_LOCAL}")

    # Шаг 3: анализируем шаблон → структура
    structure = build_structure(template_path, call_llm, load_prompt, parse_json_safe)
    structure_path = str(workdir / "structure.json")
    with open(structure_path, "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    # Шаг 4: генерируем контент
    content_path = str(workdir / "content.json")
    generate_content_json(text, structure, output_path=content_path, n_slides=slides)

    # Шаг 5: собираем PPTX
    output_path = workdir / "result.pptx"
    build_presentation(
        template_path  = template_path,
        structure_path = structure_path,
        content_path   = content_path,
        output_path    = str(output_path),
        team_members   = team_members,
    )

    pptx_bytes = output_path.read_bytes()

    # Сохраняем копии в рабочую директорию проекта для отладки
    import shutil
    shutil.copy(content_path,   str(BASE_DIR / "content.json"))
    shutil.copy(structure_path, str(BASE_DIR / "structure.json"))
    (BASE_DIR / "result.pptx").write_bytes(pptx_bytes)
    log.info(f"[{tag}] Файлы сохранены локально: result.pptx, content.json, structure.json")

    return pptx_bytes


# ── Эндпоинты ───────────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/chat")
async def api_chat(
    text:         str            = Form(...),
    style:        str            = Form(""),
    theme:        str            = Form(""),
    slides:       int            = Form(10),
    template:     Optional[UploadFile] = File(None),
    team_members: str            = Form(""),
):
    if not text.strip():
        raise HTTPException(status_code=400, detail="text обязателен")

    custom_bytes = await template.read() if template and template.filename else None
    if custom_bytes:
        log.info(f"Получен кастомный шаблон: {template.filename!r} ({len(custom_bytes)//1024} KB)")

    parsed_team: list | None = None
    if team_members.strip():
        try:
            parsed_team = json.loads(team_members)
            log.info(f"Команда: {len(parsed_team)} участников")
        except Exception as e:
            log.warning(f"Не удалось распарсить team_members: {e}")

    import asyncio
    loop = asyncio.get_running_loop()

    with tempfile.TemporaryDirectory(prefix="prezai_") as tmpdir:
        workdir = Path(tmpdir)
        log.info(f"[{workdir.name}] Запрос: text={text[:60]!r} style={style!r} theme={theme!r} slides={slides}")
        try:
            pptx_bytes = await loop.run_in_executor(
                executor,
                _run_pipeline,
                text, style, theme, slides, workdir, custom_bytes, parsed_team,
            )
        except Exception as e:
            log.exception(f"[{workdir.name}] Ошибка пайплайна")
            raise HTTPException(status_code=500, detail=str(e))

    return Response(
        content    = pptx_bytes,
        media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers    = {"Content-Disposition": 'attachment; filename="result.pptx"'},
    )


@app.get("/")
def root():
    return FileResponse(str(BASE_DIR / "index.html"))


@app.get("/download")
async def download_file():
    file_path = BASE_DIR / "result.pptx"
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path=file_path,
        filename="result.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


@app.get("/{filename}")
def static_file(filename: str):
    path = BASE_DIR / filename
    if path.exists() and path.is_file() and not path.name.startswith("."):
        return FileResponse(str(path))
    raise HTTPException(status_code=404, detail="Not found")


# ── Entry point ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.getenv("API_PORT", "8080"))
    log.info(f"Запускаю сервер на http://0.0.0.0:{port}")
    uvicorn.run("api:app", host="0.0.0.0", port=port, reload=False)
