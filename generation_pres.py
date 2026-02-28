# template_filler.py
import json
import copy
import zipfile
import shutil
from lxml import etree

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

def _get_slide_path(index: int) -> str:
    return f"ppt/slides/slide{index + 1}.xml"


def _build_paragraphs_from_list(template_para, items):
    """
    Принимает параграф-шаблон с placeholder и список item'ов вида:
        {"type": "text"|"bullet"|"numbered", "value": "..."}
    Возвращает список новых <a:p> элементов с нужным форматированием.
    """
    # Копируем стиль run (шрифт, цвет, размер) из шаблонного параграфа
    runs = template_para.findall(f'{{{NS_A}}}r')
    template_rpr = None
    if runs:
        rpr = runs[0].find(f'{{{NS_A}}}rPr')
        if rpr is not None:
            template_rpr = copy.deepcopy(rpr)

    # Выравнивание из шаблонного параграфа
    template_ppr = template_para.find(f'{{{NS_A}}}pPr')
    algn = template_ppr.get('algn', 'l') if template_ppr is not None else 'l'

    new_paragraphs = []
    for item in items:
        item_type = item.get('type', 'bullet')
        item_value = item.get('value', '')

        p = etree.Element(f'{{{NS_A}}}p')

        # Параграф-свойства: тип буллета
        ppr = etree.SubElement(p, f'{{{NS_A}}}pPr')
        ppr.set('algn', algn)

        if item_type == 'text':
            etree.SubElement(ppr, f'{{{NS_A}}}buNone')
        elif item_type == 'bullet':
            ppr.set('marL', '342900')
            ppr.set('indent', '-342900')
            bu = etree.SubElement(ppr, f'{{{NS_A}}}buChar')
            bu.set('char', '•')
        elif item_type == 'numbered':
            ppr.set('marL', '342900')
            ppr.set('indent', '-342900')
            bu = etree.SubElement(ppr, f'{{{NS_A}}}buAutoNum')
            bu.set('type', 'arabicPeriod')

        # Run с форматированием шаблона
        r = etree.SubElement(p, f'{{{NS_A}}}r')
        if template_rpr is not None:
            r.append(copy.deepcopy(template_rpr))
        t = etree.SubElement(r, f'{{{NS_A}}}t')
        t.text = item_value

        new_paragraphs.append(p)

    return new_paragraphs


def replace_in_slide(slide_xml: bytes, replacements: dict) -> bytes:
    """
    Заменяет placeholders в XML одного слайда.
    Значение может быть строкой (простая замена) или списком dict'ов
    {"type": "text"|"bullet"|"numbered", "value": "..."} (mixed-content).
    """
    tree = etree.fromstring(slide_xml)

    simple_reps = {k: v for k, v in replacements.items() if isinstance(v, str)}
    list_reps   = {k: v for k, v in replacements.items() if isinstance(v, list)}

    # ── Простые замены строк ──────────────────────────────────────────────────
    for para in tree.iter(f'{{{NS_A}}}p'):
        runs = para.findall(f'{{{NS_A}}}r')
        if not runs:
            continue

        full_text = "".join(
            (r.find(f'{{{NS_A}}}t').text or "")
            for r in runs
            if r.find(f'{{{NS_A}}}t') is not None
        )

        new_text = full_text
        changed = False
        for placeholder, value in simple_reps.items():
            if placeholder in new_text:
                new_text = new_text.replace(placeholder, value)
                changed = True

        if not changed:
            continue

        first_t = runs[0].find(f'{{{NS_A}}}t')
        if first_t is not None:
            first_t.text = new_text
        for run in runs[1:]:
            t_el = run.find(f'{{{NS_A}}}t')
            if t_el is not None:
                t_el.text = ""

    # ── List-замены: один placeholder → несколько параграфов ─────────────────
    # txBody — это p:txBody (NS_P), параграфы внутри — a:p (NS_A)
    for txBody in tree.iter(f'{{{NS_P}}}txBody'):
        children = list(txBody)
        for para in children:
            if para.tag != f'{{{NS_A}}}p':
                continue
            runs = para.findall(f'{{{NS_A}}}r')
            if not runs:
                continue

            full_text = "".join(
                (r.find(f'{{{NS_A}}}t').text or "")
                for r in runs
                if r.find(f'{{{NS_A}}}t') is not None
            )

            for placeholder, items in list_reps.items():
                if placeholder in full_text:
                    idx = list(txBody).index(para)
                    txBody.remove(para)
                    for i, new_para in enumerate(_build_paragraphs_from_list(para, items)):
                        txBody.insert(idx + i, new_para)
                    break

    return etree.tostring(tree, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


def fill_template(template_path: str, content_json: dict, output_path: str):
    shutil.copy2(template_path, output_path)
    slides_data = content_json.get("slides", [])

    with zipfile.ZipFile(output_path, 'r') as zin:
        all_files = {name: zin.read(name) for name in zin.namelist()}

    for slide_conf in slides_data:
        idx = slide_conf["slide_index"]
        replacements = slide_conf["replacements"]
        slide_path = _get_slide_path(idx)

        if slide_path not in all_files:
            print(f"[WARN] Слайд {idx} не найден в шаблоне, пропускаем")
            continue

        all_files[slide_path] = replace_in_slide(all_files[slide_path], replacements)
        print(f"[OK] Слайд {idx + 1}: заменено {len(replacements)} placeholder(s)")

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)

    print(f"\nГотово: {output_path}")


# ── Запуск ────────────────────────────────────────────────────────────────────
with open("content.json", encoding="utf-8") as f:
    content = json.load(f)

fill_template("test.pptx", content, "result.pptx")
