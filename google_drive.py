"""
Утилиты для работы с Google Drive.

Скачивает PPTX-шаблон по публичной ссылке Google Drive.
Загрузка обратно — TODO (требует OAuth2/Service Account).
"""

import re
import logging
import requests

log = logging.getLogger(__name__)


def extract_file_id(gdrive_url: str) -> str:
    """Извлекает ID файла из любого формата ссылки Google Drive."""
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', gdrive_url)
    if not match:
        raise ValueError(f"Не удалось извлечь ID файла из ссылки: {gdrive_url}")
    return match.group(1)


def download_template(gdrive_url: str, local_path: str = "template.pptx") -> str:
    """
    Скачивает PPTX с Google Drive по экспортному URL.
    Файл должен быть открыт для просмотра (Anyone with the link).

    Возвращает путь к сохранённому файлу.
    """
    file_id = extract_file_id(gdrive_url)
    export_url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"

    log.info(f"Google Drive: скачиваю шаблон (ID={file_id})...")
    log.debug(f"URL: {export_url}")

    session = requests.Session()
    response = session.get(export_url, stream=True, timeout=60)
    response.raise_for_status()

    # Google иногда возвращает страницу подтверждения антивируса для больших файлов
    content_type = response.headers.get("Content-Type", "")
    if "text/html" in content_type:
        log.warning("Google Drive вернул HTML (возможно, нужна авторизация или подтверждение)")
        raise RuntimeError(
            "Не удалось скачать файл: Google Drive требует авторизации. "
            "Убедитесь, что файл открыт по ссылке (Anyone with the link → Viewer)."
        )

    with open(local_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)

    size_kb = sum(1 for _ in open(local_path, "rb")) // 1024
    log.info(f"Google Drive: шаблон сохранён → {local_path}")
    return local_path
