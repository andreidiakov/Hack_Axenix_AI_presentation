# PrezAI — AI-генератор презентаций

**PrezAI** — сервис для автоматической генерации PPTX-презентаций по текстовой теме. Пользователь вводит тему (и опционально: стиль, оформление, количество слайдов, список команды), система выдаёт готовый `.pptx`-файл через браузер или API.

---

## Как это реально работает

### Общий пайплайн

```
Тема пользователя
  → [1] TemplateSelector   — выбирает шаблон через LLM + Google Sheets
  → [2] Google Drive       — скачивает .pptx шаблон (fallback → локальный файл)
  → [3] TemplateParser     — анализирует XML шаблона, извлекает {{плейсхолдеры}}
  → [4] PlannerAgent       — планирует N слайдов (какие типы, в каком порядке)
  → [5] WriterAgent ×N     — параллельно генерирует текст для каждого слайда
  → [6] PPTX Builder       — собирает итоговый .pptx на уровне ZIP/XML
  → result.pptx
```

### Детально по шагам

**1. TemplateSelector** (`template_selector.py`)

Загружает таблицу шаблонов из Google Sheets (название, стиль, ссылка на Drive). Отправляет в LLM тему + метаданные шаблонов. LLM выбирает наиболее подходящий вариант и возвращает Google Drive ссылку.

**2. Скачивание шаблона** (`google_drive.py`)

Скачивает `.pptx` по ссылке с Google Drive. Если Drive недоступен — использует локальный fallback-файл (`test.pptx`). Пользователь может загрузить свой шаблон через веб-интерфейс, тогда шаги 1–2 пропускаются.

**3. TemplateParser** (`template_parser.py`)

Парсит ZIP-структуру `.pptx` на уровне XML. Извлекает плейсхолдеры вида `{{TITLE}}`, `{{ITEMS}}`, `{{BULLET_1}}` из каждого слайда. Для шаблонов без явных маркеров вызывает LLM: он сам определяет, какие текстовые блоки являются заглушками. Результат — `structure.json` с описанием каждого типа слайда и его полей.

**4. PlannerAgent** (`agent_system.py`)

Получает тему и список доступных типов слайдов из `structure.json`. LLM планирует структуру: какие типы слайдов использовать, в каком порядке, и зачем нужен каждый слайд. Первый слайд всегда `TITLE`, последний — `CLOSE`.

**5. WriterAgent** (`agent_system.py`)

Для каждого слайда из плана LLM заполняет поля `replacements`: заголовки, текст, буллеты. Все слайды генерируются параллельно (до 5 потоков). Списковые поля (например, `{{ITEMS}}`) получают массив объектов `{"type": "bullet"|"text"|"numbered", "value": "..."}`.

**6. PPTX Builder** (`generation_pres.py`)

Работает напрямую с ZIP-архивом `.pptx` через `lxml`:
- Клонирует XML-слайды из шаблона в нужном порядке (поддерживает дубли типов)
- Вставляет текст через прямую замену в XML, сохраняя шрифты и форматирование шаблона
- Пересобирает служебные файлы: `presentation.xml`, `ppt/_rels/presentation.xml.rels`, `[Content_Types].xml`
- Если передан список участников команды — добавляет слайд `TITLE_TEAM` автоматически

---

## Структура проекта

```
├── api.py                   # FastAPI-сервер: HTTP-эндпоинты, оркестрация пайплайна
├── agent_system.py          # PlannerAgent + WriterAgent + LLM-клиент
├── template_selector.py     # Выбор шаблона: LLM + Google Sheets
├── template_parser.py       # Анализ структуры PPTX-шаблона (XML → JSON)
├── generation_pres.py       # Сборка итогового PPTX (ZIP/XML манипуляции)
├── google_drive.py          # Скачивание шаблонов с Google Drive
├── main.py                  # CLI-точка входа
├── index.html               # Веб-интерфейс (чат с AI)
├── prompts/
│   ├── planner_system.txt        # Системный промпт: планировщик структуры
│   ├── planner_user.txt          # Пользовательский промпт: планировщик
│   ├── writer_system.txt         # Системный промпт: автор слайдов
│   ├── writer_user.txt           # Пользовательский промпт: автор слайдов
│   ├── writer_list_format_hint.txt  # Формат списков
│   ├── template_analyzer_system.txt # LLM-анализатор шаблонов
│   ├── template_analyzer_user.txt
│   ├── template_selector_system.txt # LLM-выбор шаблона
│   └── template_selector_user.txt
├── .env                     # Конфигурация (LLM, Google Sheets, fallback)
├── test.pptx                # Локальный fallback-шаблон
├── structure.json           # Последняя разобранная структура шаблона (debug)
├── content.json             # Последний сгенерированный контент (debug)
└── result.pptx              # Последний результат (debug)
```

---

## Установка

```bash
pip install openai httpx python-dotenv lxml python-pptx pandas requests fastapi uvicorn
```

### Конфигурация `.env`

```env
# OpenAI-совместимый LLM
LLM_BASE_URL=http://your-llm-server:8000/v1
LLM_API_KEY=your-api-key
LLM_MODEL=your-model-name
LLM_TIMEOUT=600.0

# Google Sheets с каталогом шаблонов
TEMPLATES_SHEET_URL=https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit

# Локальный fallback если Google Drive недоступен
FALLBACK_LOCAL_TEMPLATE=test.pptx
```

---

## Запуск

### Веб-сервер (основной режим)

```bash
python api.py
# или
uvicorn api:app --host 0.0.0.0 --port 8080 --reload
```

Открыть в браузере: `http://localhost:8080`

### CLI

```bash
python main.py "Тренды ИИ в образовании"
python main.py "Квартальный отчёт" --style minimalism --theme dark
```

---

## API

| Метод | Путь | Описание |
|-------|------|----------|
| `GET` | `/` | Веб-интерфейс |
| `GET` | `/health` | Статус сервера |
| `POST` | `/api/chat` | Генерация PPTX |
| `GET` | `/download` | Скачать последний result.pptx |

### POST `/api/chat` — параметры формы

| Поле | Тип | Описание |
|------|-----|----------|
| `text` | string (обязательно) | Тема презентации |
| `style` | string | Стиль (minimalism, corporate и т.д.) |
| `theme` | string | Оформление (light, dark и т.д.) |
| `slides` | int | Количество слайдов (по умолчанию 10) |
| `template` | file | Свой .pptx шаблон (опционально) |
| `team_members` | JSON string | Список команды `[{"name":"...","role":"..."}]` |

Ответ — бинарный `.pptx` файл.

---

## Формат шаблонов

Шаблон — обычный `.pptx` файл с плейсхолдерами вида `{{ИМЯ_ПОЛЯ}}` в текстовых блоках. TemplateParser извлекает их автоматически. Для шаблонов без маркеров — LLM сам определяет заглушки.

Пример `content.json`:

```json
{
  "slides": [
    {
      "slide_type": "TITLE",
      "replacements": {
        "{{TITLE_TITLE}}": "ИИ в образовании: тренды 2025",
        "{{SUBTITLE}}": "Персонализация, автоматизация, новые возможности"
      }
    },
    {
      "slide_type": "BULLETS_4",
      "replacements": {
        "{{TITLE_BULLETS_4}}": "Ключевые направления",
        "{{ITEMS}}": [
          {"type": "bullet", "value": "Персонализированные траектории обучения"},
          {"type": "bullet", "value": "Автоматическая оценка знаний"},
          {"type": "bullet", "value": "Генерация учебных материалов"}
        ]
      }
    }
  ]
}
```

---

## Технологии

- **Python 3.11+**
- **FastAPI + Uvicorn** — HTTP API и веб-сервер
- **OpenAI SDK** — OpenAI-совместимый LLM (любой: vLLM, LM Studio, etc.)
- **lxml** — прямая работа с XML внутри PPTX
- **python-pptx** — вспомогательный анализ презентаций
- **pandas** — таблица шаблонов из Google Sheets
- **httpx** — HTTP-клиент с таймаутами

---

*Разработано на хакатоне Axenix AI.*
