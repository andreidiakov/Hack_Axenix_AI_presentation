"""
Утилиты для работы с Google Drive.

Скачивает PPTX-шаблон по публичной ссылке Google Drive / Google Slides.
Загрузка обратно — TODO (требует OAuth2/Service Account).
"""

import re
import logging
import requests

log = logging.getLogger(__name__)


def extract_file_id(gdrive_url: str) -> str:
    """Извлекает ID файла из любого формата ссылки Google Drive/Slides."""
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', gdrive_url)
    if not match:
        raise ValueError(f"Не удалось извлечь ID файла из ссылки: {gdrive_url}")
    return match.group(1)


def download_template(gdrive_url: str, local_path: str = "template.pptx") -> str:
    """
    Скачивает PPTX с Google Drive по публичной ссылке.
    Пробует прямой экспорт, потом fallback через uc?export=download.

    Файл должен быть открыт для просмотра (Anyone with the link).
    Возвращает путь к сохранённому файлу.
    """
    file_id = extract_file_id(gdrive_url)

    urls_to_try = [
        f"https://docs.google.com/presentation/d/{file_id}/export/pptx",
        f"https://drive.google.com/uc?export=download&id={file_id}&confirm=t",
    ]

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    last_error = None
    for url in urls_to_try:
        log.info(f"Google Drive: пробую URL → {url}")
        try:
            resp = session.get(url, stream=True, timeout=60, allow_redirects=True)
            resp.raise_for_status()
        except Exception as e:
            log.warning(f"Google Drive: FAIL [{url}] — {e}")
            last_error = e
            continue

        content_type = resp.headers.get("Content-Type", "")
        log.info(f"Google Drive: Content-Type={content_type!r}")

        # Google иногда возвращает HTML со страницей подтверждения (антивирус)
        if "text/html" in content_type:
            html = resp.text
            # Ищем confirm-токен в теле страницы
            m = re.search(r'confirm=([^&"\']+)', html)
            if m:
                confirm_token = m.group(1)
                confirm_url = (
                    f"https://drive.google.com/uc?export=download"
                    f"&id={file_id}&confirm={confirm_token}"
                )
                log.info(f"Google Drive: страница подтверждения, retry → {confirm_url}")
                try:
                    resp = session.get(confirm_url, stream=True, timeout=60, allow_redirects=True)
                    resp.raise_for_status()
                    if "text/html" not in resp.headers.get("Content-Type", ""):
                        break  # успех
                except Exception as e:
                    last_error = e
            log.warning(f"Google Drive: FAIL (HTML-ответ) для URL={url}")
            last_error = RuntimeError(
                f"Google Drive вернул HTML для {url}. "
                "Убедитесь, что файл открыт по ссылке (Anyone with the link → Viewer)."
            )
            continue

        # Успешный ответ с бинарными данными
        break
    else:
        raise last_error or RuntimeError("Не удалось скачать шаблон с Google Drive")

    with open(local_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    size_kb = resp.headers.get("Content-Length", "?")
    log.info(f"Google Drive: шаблон сохранён → {local_path} (Content-Length: {size_kb})")
    return local_path
