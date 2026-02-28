"""
FastAPI-сервер для PrezAI.

Эндпоинты:
  GET  /          → index.html
  GET  /health    → {"status": "ok"}
  POST /api/chat  → { text, style?, theme?, slides? } → PPTX-файл (blob)

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

import uvicorn
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel

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

BASE_DIR             = Path(__file__).parent
FALLBACK_LOCAL       = os.getenv("FALLBACK_LOCAL_TEMPLATE", "test.pptx")
WORKERS              = int(os.getenv("API_WORKERS", "4"))

app      = FastAPI(title="PrezAI", version="1.0")
executor = ThreadPoolExecutor(max_workers=WORKERS)


# ── Схема запроса ───────────────────────────────────────────────────────────────
class GenerateRequest(BaseModel):
    text:   str
    style:  str = ""
    theme:  str = ""
    slides: int = 10


# ── Пайплайн (запускается в потоке) ────────────────────────────────────────────
def _run_pipeline(req: GenerateRequest, workdir: Path) -> bytes:
    tag = workdir.name

    # Шаг 1: выбираем шаблон
    gdrive_link = select_template(
        topic  = req.text,
        style  = req.style,
        theme  = req.theme,
    )
    log.info(f"[{tag}] Шаблон: {gdrive_link}")

    # Шаг 2: скачиваем шаблон (fallback → локальный файл)
    template_local = str(workdir / "template.pptx")
    try:
        template_path = download_template(gdrive_link, local_path=template_local)
    except Exception as e:
        log.warning(f"[{tag}] Drive недоступен ({e}), использую локальный шаблон")
        template_path = FALLBACK_LOCAL

    # Шаг 3: анализируем шаблон → структура
    structure = build_structure(template_path, call_llm, load_prompt, parse_json_safe)
    structure_path = str(workdir / "structure.json")
    with open(structure_path, "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    # Шаг 4: генерируем контент
    content_path = str(workdir / "content.json")
    generate_content_json(req.text, structure, output_path=content_path)

    # Шаг 5: собираем PPTX
    output_path = workdir / "result.pptx"
    build_presentation(
        template_path  = template_path,
        structure_path = structure_path,
        content_path   = content_path,
        output_path    = str(output_path),
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
async def api_chat(req: GenerateRequest):
    if not req.text.strip():
        raise HTTPException(status_code=400, detail="text обязателен")

    import asyncio
    loop = asyncio.get_running_loop()

    with tempfile.TemporaryDirectory(prefix="prezai_") as tmpdir:
        workdir = Path(tmpdir)
        log.info(f"[{workdir.name}] Запрос: text={req.text[:60]!r} style={req.style!r} theme={req.theme!r}")
        try:
            pptx_bytes = await loop.run_in_executor(
                executor, _run_pipeline, req, workdir
            )
        except Exception as e:
            log.exception(f"[{workdir.name}] Ошибка пайплайна")
            raise HTTPException(status_code=500, detail=str(e))

    return Response(
        content     = pptx_bytes,
        media_type  = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers     = {"Content-Disposition": 'attachment; filename="result.pptx"'},
    )


@app.get("/")
def root():
    return FileResponse(str(BASE_DIR / "index.html"))


@app.get("/{filename}")
def static_file(filename: str):
    path = BASE_DIR / filename
    if path.exists() and path.is_file() and not path.name.startswith("."):
        return FileResponse(str(path))
    raise HTTPException(status_code=404, detail="Not found")

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

# ── Entry point ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.getenv("API_PORT", "8080"))
    log.info(f"Запускаю сервер на http://0.0.0.0:{port}")
    uvicorn.run("api:app", host="0.0.0.0", port=port, reload=False)
