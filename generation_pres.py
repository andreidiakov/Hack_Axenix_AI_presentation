"""
Сборщик презентаций из шаблона.

Вход:
  - test.pptx         — шаблон с типизированными слайдами
  - structure.json    — схема: slide_type → slide_index + доступные replacements
  - content.json      — структура итоговой презентации

Выход:
  - result.pptx       — готовая презентация

Поддерживает: дубли слайдов, изменение порядка, удаление, mixed-content (text/bullet/numbered).
"""

import json
import copy
import zipfile
import re
from lxml import etree

# ── Namespace constants ───────────────────────────────────────────────────────
NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'

# Типы связей (relationship types)
SLIDE_REL_TYPE  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
NOTES_REL_TYPE  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
SLIDE_CT        = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

# Регулярка для фильтрации слайд-файлов из шаблона при копировании
SLIDE_FILE_RE = re.compile(r'^ppt/slides/(slide\d+\.xml|_rels/slide\d+\.xml\.rels)$')


# ── Text replacement helpers ──────────────────────────────────────────────────

def _build_paragraphs_from_list(template_para, items):
    """
    Раскрывает один placeholder-параграф в несколько параграфов.

    Принимает:
      template_para — <a:p> элемент с placeholder (нужен для наследования стиля)
      items         — список dict'ов: {"type": "text"|"bullet"|"numbered", "value": "..."}

    Возвращает список новых <a:p> с нужным форматированием буллетов.
    Стиль шрифта (размер, цвет, жирность) берётся из первого run шаблона.
    """
    # Копируем rPr (run properties) из шаблонного параграфа для сохранения стиля
    runs = template_para.findall(f'{{{NS_A}}}r')
    template_rpr = None
    if runs:
        rpr = runs[0].find(f'{{{NS_A}}}rPr')
        if rpr is not None:
            template_rpr = copy.deepcopy(rpr)

    # Берём выравнивание из pPr шаблона
    template_ppr = template_para.find(f'{{{NS_A}}}pPr')
    algn = template_ppr.get('algn', 'l') if template_ppr is not None else 'l'

    new_paragraphs = []
    for item in items:
        item_type  = item.get('type', 'bullet')
        item_value = item.get('value', '')

        p   = etree.Element(f'{{{NS_A}}}p')
        ppr = etree.SubElement(p, f'{{{NS_A}}}pPr')
        ppr.set('algn', algn)

        if item_type == 'text':
            # Обычный текст — без символа буллета
            etree.SubElement(ppr, f'{{{NS_A}}}buNone')
        elif item_type == 'bullet':
            # Маркированный список — символ •
            ppr.set('marL', '342900')
            ppr.set('indent', '-342900')
            etree.SubElement(ppr, f'{{{NS_A}}}buChar').set('char', '•')
        elif item_type == 'numbered':
            # Нумерованный список — 1. 2. 3.
            ppr.set('marL', '342900')
            ppr.set('indent', '-342900')
            etree.SubElement(ppr, f'{{{NS_A}}}buAutoNum').set('type', 'arabicPeriod')

        # Run наследует стиль (шрифт, цвет) из шаблонного параграфа
        r = etree.SubElement(p, f'{{{NS_A}}}r')
        if template_rpr is not None:
            r.append(copy.deepcopy(template_rpr))
        etree.SubElement(r, f'{{{NS_A}}}t').text = item_value

        new_paragraphs.append(p)

    return new_paragraphs


def replace_in_slide(slide_xml: bytes, replacements: dict) -> bytes:
    """
    Заменяет placeholders в XML одного слайда.

    replacements — словарь: ключ → значение
      - строка   → простая замена текста
      - список   → mixed-content (несколько параграфов с разными типами буллетов)

    Возвращает модифицированный XML как bytes.
    """
    tree = etree.fromstring(slide_xml)

    # Разбиваем замены на два типа
    simple_reps = {k: v for k, v in replacements.items() if isinstance(v, str)}
    list_reps   = {k: v for k, v in replacements.items() if isinstance(v, list)}

    # ── 1. Простые замены строк ───────────────────────────────────────────────
    # Проходим по всем параграфам, собираем текст из runs, заменяем placeholder
    for para in tree.iter(f'{{{NS_A}}}p'):
        runs = para.findall(f'{{{NS_A}}}r')
        if not runs:
            continue

        # Склеиваем текст всех runs параграфа (placeholder может быть разбит)
        full_text = ''.join(
            (r.find(f'{{{NS_A}}}t').text or '')
            for r in runs
            if r.find(f'{{{NS_A}}}t') is not None
        )

        new_text = full_text
        changed  = False
        for ph, val in simple_reps.items():
            if ph in new_text:
                new_text = new_text.replace(ph, val)
                changed = True

        if not changed:
            continue

        # Пишем результат в первый run, очищаем остальные
        first_t = runs[0].find(f'{{{NS_A}}}t')
        if first_t is not None:
            first_t.text = new_text
        for run in runs[1:]:
            t_el = run.find(f'{{{NS_A}}}t')
            if t_el is not None:
                t_el.text = ''

    # ── 2. List-замены: один placeholder → несколько параграфов ──────────────
    # txBody — это p:txBody (NS_P), параграфы внутри — a:p (NS_A)
    for txBody in tree.iter(f'{{{NS_P}}}txBody'):
        for para in list(txBody):
            if para.tag != f'{{{NS_A}}}p':
                continue
            runs = para.findall(f'{{{NS_A}}}r')
            if not runs:
                continue

            full_text = ''.join(
                (r.find(f'{{{NS_A}}}t').text or '')
                for r in runs
                if r.find(f'{{{NS_A}}}t') is not None
            )

            for ph, items in list_reps.items():
                if ph in full_text:
                    if not items:
                        # Пустой список — оставляем параграф нетронутым,
                        # иначе XML останется без параграфа и PPTX сломается
                        break
                    # Удаляем шаблонный параграф и вставляем на его место новые
                    idx = list(txBody).index(para)
                    txBody.remove(para)
                    for i, new_para in enumerate(_build_paragraphs_from_list(para, items)):
                        txBody.insert(idx + i, new_para)
                    break  # один placeholder за раз, выходим из inner loop

    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


# ── Slide rels helper ─────────────────────────────────────────────────────────

def _clean_slide_rels(rels_xml: bytes) -> bytes:
    """
    Очищает _rels файл слайда от ссылок на notesSlide.

    Проблема: при дублировании слайда оба output-слайда указывали бы
    на одну и ту же notesSlide из шаблона — это ломает валидацию PowerPoint.
    Решение: убираем notesSlide ref; слайды без заметок работают корректно.
    """
    tree = etree.fromstring(rels_xml)
    for rel in tree.findall(f'{{{NS_REL}}}Relationship'):
        if rel.get('Type') == NOTES_REL_TYPE:
            tree.remove(rel)
    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


# ── PPTX structure rebuilders ─────────────────────────────────────────────────

def _rebuild_pres_rels(template_rels_xml: bytes, num_slides: int):
    """
    Пересобирает ppt/_rels/presentation.xml.rels.

    Удаляет все старые slide-ссылки.
    Добавляет новые, используя rId ПОСЛЕ существующих non-slide rId'ов
    → избегаем конфликт с rId'ами theme/viewProps/presProps/slideMaster/notesMaster.

    Возвращает: (bytes XML, list[str] новых rId для слайдов)
    """
    tree = etree.fromstring(template_rels_xml)

    # Удаляем старые slide rels
    for rel in tree.findall(f'{{{NS_REL}}}Relationship'):
        if rel.get('Type') == SLIDE_REL_TYPE:
            tree.remove(rel)

    # Находим максимальный rId среди оставшихся (non-slide) связей
    # чтобы новые rId'ы не конфликтовали с theme, viewProps, presProps и т.д.
    existing_ids = []
    for rel in tree.findall(f'{{{NS_REL}}}Relationship'):
        m = re.match(r'rId(\d+)', rel.get('Id', ''))
        if m:
            existing_ids.append(int(m.group(1)))
    start = max(existing_ids, default=0) + 1

    # Добавляем новые slide rels начиная с rId(start)
    slide_rids = []
    for i in range(num_slides):
        rid = f'rId{start + i}'
        rel = etree.SubElement(tree, f'{{{NS_REL}}}Relationship')
        rel.set('Id',     rid)
        rel.set('Type',   SLIDE_REL_TYPE)
        rel.set('Target', f'slides/slide{i + 1}.xml')
        slide_rids.append(rid)

    xml_bytes = etree.tostring(tree, xml_declaration=True,
                               encoding='UTF-8', standalone=True)
    return xml_bytes, slide_rids


def _rebuild_pres_xml(template_pres_xml: bytes, slide_rids: list) -> bytes:
    """
    Пересобирает ppt/presentation.xml: обновляет sldIdLst.

    Использует rId'ы, переданные из _rebuild_pres_rels,
    чтобы sldIdLst ссылался на те же rId'ы, что и .rels файл.
    """
    tree = etree.fromstring(template_pres_xml)

    sldIdLst = tree.find(f'{{{NS_P}}}sldIdLst')
    if sldIdLst is None:
        sldIdLst = etree.SubElement(tree, f'{{{NS_P}}}sldIdLst')

    # Очищаем старый список слайдов
    for child in list(sldIdLst):
        sldIdLst.remove(child)

    # Заполняем новый список; id начинается с 256 (стандарт OOXML)
    for i, rid in enumerate(slide_rids):
        sld = etree.SubElement(sldIdLst, f'{{{NS_P}}}sldId')
        sld.set('id',             str(256 + i))
        sld.set(f'{{{NS_R}}}id', rid)

    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


def _rebuild_content_types(template_ct_xml: bytes, num_slides: int) -> bytes:
    """
    Пересобирает [Content_Types].xml: заменяет записи о слайдах.

    Убирает Override'ы для старых slide{N}.xml,
    добавляет Override'ы для новых slide1.xml … slideN.xml.
    """
    tree = etree.fromstring(template_ct_xml)

    # Удаляем существующие override'ы для слайдов
    for override in tree.findall(f'{{{NS_CT}}}Override'):
        if re.match(r'/ppt/slides/slide\d+\.xml$', override.get('PartName', '')):
            tree.remove(override)

    # Добавляем override'ы для output-слайдов
    for i in range(num_slides):
        override = etree.SubElement(tree, f'{{{NS_CT}}}Override')
        override.set('PartName',    f'/ppt/slides/slide{i + 1}.xml')
        override.set('ContentType', SLIDE_CT)

    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


# ── Main builder ──────────────────────────────────────────────────────────────

def build_presentation(template_path: str,
                       structure_path: str,
                       content_path:   str,
                       output_path:    str):
    """
    Собирает итоговую презентацию из шаблона по content.json.

    Алгоритм:
      1. Читает structure.json → маппинг slide_type → номер слайда в шаблоне
      2. Читает content.json  → список слайдов с типами и заменами
      3. Для каждого слайда: клонирует XML из шаблона, применяет замены
      4. Пересобирает служебные файлы PPTX (rels, sldIdLst, Content_Types)
      5. Записывает result.pptx

    Поддерживает дубли slide_type, произвольный порядок, пропуски.
    """
    # Загружаем схему шаблона и входной контент
    with open(structure_path, encoding='utf-8') as f:
        structure = json.load(f)
    with open(content_path, encoding='utf-8') as f:
        content = json.load(f)

    # Маппинг: slide_type → 1-based номер файла слайда в шаблоне
    # Дублирующиеся типы получают суффикс _2, _3 ... (как в structure_to_schema)
    type_to_tmpl_num: dict[str, int] = {}
    _type_count: dict[str, int] = {}
    for s in structure['slides']:
        base  = s['slide_type']
        cnt   = _type_count.get(base, 0) + 1
        _type_count[base] = cnt
        key   = base if cnt == 1 else f"{base}_{cnt}"
        type_to_tmpl_num[key] = s['slide_index'] + 1

    # Загружаем все файлы шаблона в память
    with zipfile.ZipFile(template_path, 'r') as z:
        tmpl_files = {name: z.read(name) for name in z.namelist()}

    # Стартуем с копии всех файлов шаблона, кроме слайд-файлов
    # (слайды будем добавлять заново в нужном порядке)
    out_files = {k: v for k, v in tmpl_files.items()
                 if not SLIDE_FILE_RE.match(k)}

    # ── Обрабатываем слайды из content.json ──────────────────────────────────
    out_count = 0
    for conf in content.get('slides', []):
        slide_type   = conf.get('slide_type', '')
        replacements = conf.get('replacements', {})

        tmpl_num = type_to_tmpl_num.get(slide_type)
        if tmpl_num is None:
            print(f"[WARN] Неизвестный slide_type='{slide_type}', пропускаем")
            continue

        # Клонируем XML шаблонного слайда и применяем замены текста
        tmpl_xml = tmpl_files[f'ppt/slides/slide{tmpl_num}.xml']
        out_xml  = replace_in_slide(tmpl_xml, replacements)

        out_count += 1
        out_files[f'ppt/slides/slide{out_count}.xml'] = out_xml

        # Копируем _rels шаблонного слайда, очищая notesSlide-ссылки.
        # Без очистки: дублированные слайды ссылались бы на одну notesSlide
        # → PowerPoint ругался бы на невалидные cross-references.
        tmpl_rels_path = f'ppt/slides/_rels/slide{tmpl_num}.xml.rels'
        if tmpl_rels_path in tmpl_files:
            cleaned_rels = _clean_slide_rels(tmpl_files[tmpl_rels_path])
            out_files[f'ppt/slides/_rels/slide{out_count}.xml.rels'] = cleaned_rels

        print(f'[OK] Слайд {out_count}: {slide_type}  ({len(replacements)} замен)')

    # ── Пересобираем служебные файлы PPTX ────────────────────────────────────

    # Rels: удаляем старые slide-ссылки, добавляем новые с безопасными rId'ами
    # (rId'ы начинаются ПОСЛЕ максимального существующего non-slide rId)
    new_rels_xml, slide_rids = _rebuild_pres_rels(
        tmpl_files['ppt/_rels/presentation.xml.rels'], out_count)
    out_files['ppt/_rels/presentation.xml.rels'] = new_rels_xml

    # presentation.xml: обновляем sldIdLst с теми же rId'ами
    out_files['ppt/presentation.xml'] = _rebuild_pres_xml(
        tmpl_files['ppt/presentation.xml'], slide_rids)

    # Content_Types: регистрируем новые slide{N}.xml
    out_files['[Content_Types].xml'] = _rebuild_content_types(
        tmpl_files['[Content_Types].xml'], out_count)

    # ── Пишем итоговый файл ───────────────────────────────────────────────────
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in out_files.items():
            z.writestr(name, data)

    print(f'\nГотово → {output_path}  ({out_count} слайдов)')


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == '__main__':
    build_presentation(
        template_path  = 'test.pptx',
        structure_path = 'structure.json',
        content_path   = 'content.json',
        output_path    = 'result.pptx',
    )
