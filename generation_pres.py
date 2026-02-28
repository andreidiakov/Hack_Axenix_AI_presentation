"""
Сборщик презентаций из шаблона.

Вход:
  - test.pptx         — шаблон с 9 типами слайдов
  - structure.json    — схема: slide_type → slide_index + доступные replacements
  - content.json      — структура итоговой презентации: список слайдов с типами и заменами

Выход:
  - result.pptx       — готовая презентация
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

SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
SLIDE_CT       = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

SLIDE_RE = re.compile(r'^ppt/slides/(slide\d+\.xml|_rels/slide\d+\.xml\.rels)$')


# ── Text replacement helpers ──────────────────────────────────────────────────

def _build_paragraphs_from_list(template_para, items):
    """
    Принимает параграф-шаблон с placeholder и список item'ов:
        {"type": "text"|"bullet"|"numbered", "value": "..."}
    Возвращает список новых <a:p> с нужным форматированием.
    """
    runs = template_para.findall(f'{{{NS_A}}}r')
    template_rpr = None
    if runs:
        rpr = runs[0].find(f'{{{NS_A}}}rPr')
        if rpr is not None:
            template_rpr = copy.deepcopy(rpr)

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
            etree.SubElement(ppr, f'{{{NS_A}}}buNone')
        elif item_type == 'bullet':
            ppr.set('marL', '342900')
            ppr.set('indent', '-342900')
            etree.SubElement(ppr, f'{{{NS_A}}}buChar').set('char', '•')
        elif item_type == 'numbered':
            ppr.set('marL', '342900')
            ppr.set('indent', '-342900')
            etree.SubElement(ppr, f'{{{NS_A}}}buAutoNum').set('type', 'arabicPeriod')

        r = etree.SubElement(p, f'{{{NS_A}}}r')
        if template_rpr is not None:
            r.append(copy.deepcopy(template_rpr))
        etree.SubElement(r, f'{{{NS_A}}}t').text = item_value

        new_paragraphs.append(p)

    return new_paragraphs


def replace_in_slide(slide_xml: bytes, replacements: dict) -> bytes:
    """
    Заменяет placeholders в XML одного слайда.
    Значение — строка (простая замена) или список dict'ов (mixed-content).
    """
    tree = etree.fromstring(slide_xml)

    simple_reps = {k: v for k, v in replacements.items() if isinstance(v, str)}
    list_reps   = {k: v for k, v in replacements.items() if isinstance(v, list)}

    # Простые замены строк
    for para in tree.iter(f'{{{NS_A}}}p'):
        runs = para.findall(f'{{{NS_A}}}r')
        if not runs:
            continue
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
        first_t = runs[0].find(f'{{{NS_A}}}t')
        if first_t is not None:
            first_t.text = new_text
        for run in runs[1:]:
            t_el = run.find(f'{{{NS_A}}}t')
            if t_el is not None:
                t_el.text = ''

    # List-замены: placeholder → несколько параграфов
    # txBody — p:txBody (NS_P), параграфы внутри — a:p (NS_A)
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
                    idx = list(txBody).index(para)
                    txBody.remove(para)
                    for i, new_para in enumerate(_build_paragraphs_from_list(para, items)):
                        txBody.insert(idx + i, new_para)
                    break

    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


# ── PPTX structure rebuilders ─────────────────────────────────────────────────

def _rebuild_pres_rels(template_rels_xml: bytes, num_slides: int) -> bytes:
    """Пересобирает presentation.xml.rels: убирает старые слайды, добавляет новые."""
    tree = etree.fromstring(template_rels_xml)
    for rel in tree.findall(f'{{{NS_REL}}}Relationship'):
        if rel.get('Type') == SLIDE_REL_TYPE:
            tree.remove(rel)
    for i in range(num_slides):
        rel = etree.SubElement(tree, f'{{{NS_REL}}}Relationship')
        rel.set('Id',     f'rId{i + 2}')
        rel.set('Type',   SLIDE_REL_TYPE)
        rel.set('Target', f'slides/slide{i + 1}.xml')
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def _rebuild_pres_xml(template_pres_xml: bytes, num_slides: int) -> bytes:
    """Пересобирает presentation.xml: обновляет sldIdLst."""
    tree = etree.fromstring(template_pres_xml)
    sldIdLst = tree.find(f'{{{NS_P}}}sldIdLst')
    if sldIdLst is None:
        sldIdLst = etree.SubElement(tree, f'{{{NS_P}}}sldIdLst')
    for child in list(sldIdLst):
        sldIdLst.remove(child)
    for i in range(num_slides):
        sld = etree.SubElement(sldIdLst, f'{{{NS_P}}}sldId')
        sld.set('id',             str(256 + i))
        sld.set(f'{{{NS_R}}}id', f'rId{i + 2}')
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def _rebuild_content_types(template_ct_xml: bytes, num_slides: int) -> bytes:
    """Пересобирает [Content_Types].xml: заменяет записи о слайдах."""
    tree = etree.fromstring(template_ct_xml)
    for override in tree.findall(f'{{{NS_CT}}}Override'):
        if re.match(r'/ppt/slides/slide\d+\.xml$', override.get('PartName', '')):
            tree.remove(override)
    for i in range(num_slides):
        override = etree.SubElement(tree, f'{{{NS_CT}}}Override')
        override.set('PartName',    f'/ppt/slides/slide{i + 1}.xml')
        override.set('ContentType', SLIDE_CT)
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


# ── Main builder ──────────────────────────────────────────────────────────────

def build_presentation(template_path: str,
                       structure_path: str,
                       content_path:   str,
                       output_path:    str):
    """
    Собирает итоговую презентацию из шаблона по content.json.

    content.json["slides"] — список:
        {
          "slide_type":   "COMPARE",          # тип из structure.json
          "replacements": { "{{KEY}}": "..." } # замены (str или list)
        }
    Слайды могут повторяться, отсутствовать, идти в любом порядке.
    """
    with open(structure_path, encoding='utf-8') as f:
        structure = json.load(f)
    with open(content_path, encoding='utf-8') as f:
        content = json.load(f)

    # slide_type → индекс шаблонного слайда (1-based для имён файлов)
    type_to_tmpl = {s['slide_type']: s['slide_index'] + 1
                    for s in structure['slides']}

    # Загружаем все файлы шаблона
    with zipfile.ZipFile(template_path, 'r') as z:
        tmpl_files = {name: z.read(name) for name in z.namelist()}

    # Базовый набор файлов: всё кроме слайдов
    out_files = {k: v for k, v in tmpl_files.items() if not SLIDE_RE.match(k)}

    # Обрабатываем слайды из content.json
    out_count = 0
    for conf in content.get('slides', []):
        slide_type   = conf.get('slide_type', '')
        replacements = conf.get('replacements', {})

        tmpl_num = type_to_tmpl.get(slide_type)
        if tmpl_num is None:
            print(f"[WARN] Неизвестный slide_type='{slide_type}', пропускаем")
            continue

        # Клонируем шаблонный слайд и применяем замены
        tmpl_xml = tmpl_files[f'ppt/slides/slide{tmpl_num}.xml']
        out_xml  = replace_in_slide(tmpl_xml, replacements)

        out_count += 1
        out_files[f'ppt/slides/slide{out_count}.xml'] = out_xml

        # Копируем rels шаблонного слайда (ссылка на layout)
        tmpl_rels = f'ppt/slides/_rels/slide{tmpl_num}.xml.rels'
        if tmpl_rels in tmpl_files:
            out_files[f'ppt/slides/_rels/slide{out_count}.xml.rels'] = tmpl_files[tmpl_rels]

        print(f'[OK] Слайд {out_count}: {slide_type}  ({len(replacements)} замен)')

    # Пересобираем служебные файлы PPTX
    out_files['ppt/_rels/presentation.xml.rels'] = _rebuild_pres_rels(
        tmpl_files['ppt/_rels/presentation.xml.rels'], out_count)
    out_files['ppt/presentation.xml'] = _rebuild_pres_xml(
        tmpl_files['ppt/presentation.xml'], out_count)
    out_files['[Content_Types].xml'] = _rebuild_content_types(
        tmpl_files['[Content_Types].xml'], out_count)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in out_files.items():
            z.writestr(name, data)

    print(f'\nГотово → {output_path}  ({out_count} слайдов)')


# ── Entry point ───────────────────────────────────────────────────────────────
build_presentation(
    template_path  = 'test.pptx',
    structure_path = 'structure.json',
    content_path   = 'content.json',
    output_path    = 'result.pptx',
)
