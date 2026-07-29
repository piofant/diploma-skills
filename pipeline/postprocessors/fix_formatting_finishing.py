#!/usr/bin/env python3
"""
Post-process diploma_body.docx ПОСЛЕ Pandoc + tables + headings:

1. APPENDICES: «### Приложение А. Название» → page break + center «Приложение А»
   + center bold «Название» на следующей строке. Применяется для всех
   приложений А, Б, В, Г, Д, Е, Ж, И, К.

2. STRUCTURAL ELEMENTS (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ,
   СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ, ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ, СОДЕРЖАНИЕ):
   align=center, first_line_indent=0 + page break before (т.к. структурные
   элементы по регламенту начинаются с новой страницы).

3. CHAPTERS («1 АНАЛИЗ...», «2 ПРОЕКТИРОВАНИЕ...», «3 РЕАЛИЗАЦИЯ»): pageBreakBefore.

4. TABLE CAPTIONS (Normal-параграфы, начинающиеся с «Таблица N.M – ...»):
   first_line_indent=0, align=left.
"""
import sys, os, shutil, zipfile, re
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

STRUCTURAL_TEXTS = {
    'ВВЕДЕНИЕ',
    'ЗАКЛЮЧЕНИЕ',
    'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
    'СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ',
    'ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ',
    'СОДЕРЖАНИЕ',
}

# Главы (нумерованные H2 верхнего уровня)
CHAPTER_RX = re.compile(r'^\s*[1-9]\s+[А-ЯЁ]')

# «Приложение А. Название» / «Приложение А — Название» / «Приложение А»
APPENDIX_RX = re.compile(r'^\s*Приложение\s+([А-ЯЁ])\.?\s*(.*)\s*$')

# Подпись таблицы: «Таблица 1.1 – Сводный...», «Таблица А.1 – Структура...»
TABLE_CAPTION_RX = re.compile(r'^\s*Таблица\s+([\dА-ЯA-Z]+\.[\dА-ЯA-Z]+|\d+)\s*[–—-]\s*')


def _para_text(p):
    out = []
    for t in p.iter(f'{W}t'):
        if t.text:
            out.append(t.text)
    return ''.join(out).strip()


def _style_id(p):
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        return None
    pStyle = pPr.find(f'{W}pStyle')
    if pStyle is None:
        return None
    return pStyle.get(f'{W}val')


def _get_or_create_pPr(p):
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        pPr = ET.Element(f'{W}pPr')
        p.insert(0, pPr)
    return pPr


def _set_align(p, val):
    """Set <w:jc w:val='center|left|...'/>"""
    pPr = _get_or_create_pPr(p)
    jc = pPr.find(f'{W}jc')
    if jc is None:
        jc = ET.SubElement(pPr, f'{W}jc')
    jc.set(f'{W}val', val)


def _set_indent_zero(p):
    """Remove first-line indent (set firstLine=0)."""
    pPr = _get_or_create_pPr(p)
    ind = pPr.find(f'{W}ind')
    if ind is None:
        ind = ET.SubElement(pPr, f'{W}ind')
    ind.set(f'{W}firstLine', '0')
    # also wipe left, in case
    if ind.attrib.get(f'{W}left') and ind.attrib[f'{W}left'] != '0':
        pass  # keep left as-is for now


def _set_page_break_before(p):
    pPr = _get_or_create_pPr(p)
    pbb = pPr.find(f'{W}pageBreakBefore')
    if pbb is None:
        ET.SubElement(pPr, f'{W}pageBreakBefore')


def _strip_runs_keep_text(p, new_text=None, bold=None):
    """Remove all runs from p; insert a single new run with new_text (or keep original)."""
    if new_text is None:
        new_text = _para_text(p)
    # remove existing runs
    for r in list(p.findall(f'{W}r')):
        p.remove(r)
    # add new run
    r = ET.SubElement(p, f'{W}r')
    rPr = ET.SubElement(r, f'{W}rPr')
    rfonts = ET.SubElement(rPr, f'{W}rFonts')
    rfonts.set(f'{W}ascii', 'Times New Roman')
    rfonts.set(f'{W}hAnsi', 'Times New Roman')
    rfonts.set(f'{W}cs', 'Times New Roman')
    rfonts.set(f'{W}eastAsia', 'Times New Roman')
    sz = ET.SubElement(rPr, f'{W}sz')
    sz.set(f'{W}val', '28')
    szcs = ET.SubElement(rPr, f'{W}szCs')
    szcs.set(f'{W}val', '28')
    color = ET.SubElement(rPr, f'{W}color')
    color.set(f'{W}val', '000000')
    if bold:
        ET.SubElement(rPr, f'{W}b')
        ET.SubElement(rPr, f'{W}bCs')
    t = ET.SubElement(r, f'{W}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = new_text


def fix(xml_bytes):
    root = ET.fromstring(xml_bytes)
    body = root.find(f'{W}body')
    if body is None:
        body = root
    stats = {
        'structural': 0,
        'chapters': 0,
        'appendices': 0,
        'table_captions': 0,
    }

    # Pass 1: walk all paragraphs, mark structural / chapter / table-caption
    paragraphs = list(body.iter(f'{W}p'))
    for p in paragraphs:
        sid = _style_id(p) or ''
        text = _para_text(p)

        # 1) Structural elements (in Heading 2 mostly)
        if sid.startswith('Heading') and text in STRUCTURAL_TEXTS:
            _set_align(p, 'center')
            _set_indent_zero(p)
            _set_page_break_before(p)
            stats['structural'] += 1
            continue

        # 2) Chapter headings («1 АНАЛИЗ...» / «2 ПРОЕКТИРОВАНИЕ...»)
        if sid.startswith('Heading') and CHAPTER_RX.match(text) and len(text) < 200:
            _set_page_break_before(p)
            stats['chapters'] += 1
            continue

        # 4) Table captions (BodyText/Normal/Compact/FirstParagraph, starts with «Таблица N.M – ...»)
        if (not sid or sid in ('Normal', 'Compact', 'BodyText', 'FirstParagraph')) and TABLE_CAPTION_RX.match(text):
            _set_align(p, 'left')
            _set_indent_zero(p)
            stats['table_captions'] += 1
            continue

    # Pass 2: APPENDIX reformatting — НА МЕСТЕ, без удаления/вставки в body
    # (удалять параграф из body требует поиска родителя; XML дерево разное у элементов).
    # Простой подход: найти параграфы с текстом «Приложение Х. Название», прямо в них
    # переписать на «Приложение Х» с pageBreakBefore center + ВСТАВИТЬ следующий параграф
    # «Название» bold center сразу за ним.
    # Идём через body.findall, чтобы получить прямой parent.
    direct_children = list(body)
    for idx, elem in enumerate(direct_children):
        if elem.tag != f'{W}p':
            continue
        sid = _style_id(elem) or ''
        text = _para_text(elem)
        if not sid.startswith('Heading'):
            continue
        m = APPENDIX_RX.match(text)
        if not m:
            continue
        letter, title = m.group(1), m.group(2).strip()
        # «Приложение Х» в текущем параграфе
        _set_align(elem, 'center')
        _set_indent_zero(elem)
        _set_page_break_before(elem)
        # стиль оставляем тот же (Heading X) чтобы попасть в outline для TOC
        _strip_runs_keep_text(elem, new_text=f'Приложение {letter}', bold=False)
        # Если title непустой — вставить ещё параграф «Название» bold center.
        # ВАЖНО: стиль = 'BodyText' (а не Heading), чтобы НЕ попало в outline/TOC —
        # иначе каждое приложение даст 2 entries: «Приложение А» и название отдельно.
        # Outline покажет только короткое «Приложение А», а полное имя останется
        # в текущем Heading-параграфе — для этого расширим текст первого параграфа.
        if title:
            title = title.rstrip('. ')
            # Расширим первый параграф: «Приложение А. Структура базы данных...»
            # чтобы в outline / TOC увидели и букву, и название.
            # Но визуально на странице первый параграф = «Приложение А», под ним bold.
            # Решение: текст в outline через скрытый <w:t> не работает — outline
            # читается из видимого. Компромисс: оставить первый параграф = только
            # «Приложение А» (короткая запись TOC), второй = название без стиля
            # Heading (в TOC не попадёт). В постпроцессоре gen_static_toc можно
            # склеить, но проще оставить TOC коротким — это допустимо по ГОСТ.
            new_p = ET.Element(f'{W}p')
            new_pPr = ET.SubElement(new_p, f'{W}pPr')
            # BodyText вместо Heading — не попадает в outline
            new_pStyle = ET.SubElement(new_pPr, f'{W}pStyle')
            new_pStyle.set(f'{W}val', 'BodyText')
            new_jc = ET.SubElement(new_pPr, f'{W}jc')
            new_jc.set(f'{W}val', 'center')
            new_ind = ET.SubElement(new_pPr, f'{W}ind')
            new_ind.set(f'{W}firstLine', '0')
            _strip_runs_keep_text(new_p, new_text=title, bold=True)
            # insert after current elem
            current_idx = list(body).index(elem)
            body.insert(current_idx + 1, new_p)
        stats['appendices'] += 1

    return ET.tostring(root, encoding='UTF-8', xml_declaration=True), stats


def main(in_path, out_path):
    tmp = out_path + '.tmp'
    stats = None
    with zipfile.ZipFile(in_path, 'r') as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == 'word/document.xml':
                data, stats = fix(data)
            zout.writestr(name, data)
    shutil.move(tmp, out_path)
    print(f"  fix_formatting_finishing: {stats}")


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else None
    dst = sys.argv[2] if len(sys.argv) > 2 else src
    if not src or not os.path.exists(src):
        print(f"Usage: {sys.argv[0]} input.docx [output.docx]"); sys.exit(1)
    main(src, dst)
