#!/usr/bin/env python3
"""
Replace the empty Word {TOC} field with a STATIC table of contents.

Google Docs does NOT compute Word field codes on .docx import, so the
`{ TOC \\o "1-2" }` field renders as an empty "СОДЕРЖАНИЕ" page. We extract
the real heading->page map from the LibreOffice-rendered PDF outline
(authoritative ВКР pagination) and write literal TOC paragraphs with a
right dot-leader tab + page number. Not bold, black (ГЭК secretary req).

Usage: gen_static_toc.py <docx> <rendered_pdf> [out_docx]
"""
import sys, os, shutil, zipfile
from xml.etree import ElementTree as ET
from pypdf import PdfReader

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % WNS
ET.register_namespace('w', WNS)

RIGHT_TAB = 9300          # dxa: page text width (page 11906 - L1701 - R850 ≈ 9355)
INDENT = {1: 0, 2: 283, 3: 567}   # dxa left indent by level (~0 / 0.5cm / 1cm)


def outline(pdf_path, max_level=3, page_offset=0):
    """page_offset compensates for <w:pgNumType w:start="N"/> in section properties.
    pypdf returns physical page index (1-based), but the footer shows N + physical - 1.
    If pgNumType start=2 → page_offset=1 to match footer."""
    r = PdfReader(pdf_path)
    res = []
    def walk(items, lvl=1):
        for it in items:
            if isinstance(it, list):
                walk(it, lvl + 1)
            else:
                if lvl <= max_level:
                    try:
                        pg = r.get_destination_page_number(it) + 1 + page_offset
                    except Exception:
                        pg = None
                    if pg:
                        res.append((lvl, it.title.strip(), pg))
    walk(r.outline)
    return res


def make_toc_paragraph(level, title, page):
    p = ET.Element(f'{W}p')
    pPr = ET.SubElement(p, f'{W}pPr')
    tabs = ET.SubElement(pPr, f'{W}tabs')
    tab = ET.SubElement(tabs, f'{W}tab')
    tab.set(f'{W}val', 'right'); tab.set(f'{W}leader', 'dot'); tab.set(f'{W}pos', str(RIGHT_TAB))
    ind = ET.SubElement(pPr, f'{W}ind')
    ind.set(f'{W}left', str(INDENT.get(level, 0))); ind.set(f'{W}firstLine', '0')
    sp = ET.SubElement(pPr, f'{W}spacing')
    sp.set(f'{W}after', '0'); sp.set(f'{W}line', '276'); sp.set(f'{W}lineRule', 'auto')
    def run(text):
        r = ET.SubElement(p, f'{W}r')
        rPr = ET.SubElement(r, f'{W}rPr')
        rf = ET.SubElement(rPr, f'{W}rFonts')
        rf.set(f'{W}ascii', 'Times New Roman'); rf.set(f'{W}hAnsi', 'Times New Roman'); rf.set(f'{W}cs', 'Times New Roman')
        c = ET.SubElement(rPr, f'{W}color'); c.set(f'{W}val', '000000')
        sz = ET.SubElement(rPr, f'{W}sz'); sz.set(f'{W}val', '28')
        szcs = ET.SubElement(rPr, f'{W}szCs'); szcs.set(f'{W}val', '28')
        t = ET.SubElement(r, f'{W}t'); t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        return r
    run(title)
    rt = ET.SubElement(p, f'{W}r')
    ET.SubElement(ET.SubElement(rt, f'{W}rPr'), f'{W}color').set(f'{W}val', '000000')
    ET.SubElement(rt, f'{W}tab')
    run(str(page))
    return p


def _is_static_toc_entry(el):
    """Detect static TOC entry: paragraph with tab+leader='dot' tabs."""
    if el.tag != f'{W}p':
        return False
    pPr = el.find(f'{W}pPr')
    if pPr is None:
        return False
    tabs = pPr.find(f'{W}tabs')
    if tabs is None:
        return False
    for tab in tabs.findall(f'{W}tab'):
        if tab.get(f'{W}leader') == 'dot':
            return True
    return False


def main(docx_in, pdf_in, docx_out, page_offset=1):
    # page_offset=1 matches add_page_numbers_no_title.py with <w:pgNumType w:start="2"/>
    entries = outline(pdf_in, page_offset=page_offset)
    if not entries:
        print("WARN: no outline entries — TOC left as-is");
        if docx_in != docx_out: shutil.copy(docx_in, docx_out)
        return
    tmp = docx_out + '.tmp'
    with zipfile.ZipFile(docx_in, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == 'word/document.xml':
                root = ET.fromstring(data)
                body = root.find(f'{W}body')
                kids = list(body)

                # Mode 1: find TOC field (first pass)
                fld_idx = None
                for i, el in enumerate(kids):
                    if el.tag != f'{W}p':
                        continue
                    for it in el.iter(f'{W}instrText'):
                        if it.text and ' TOC ' in it.text:
                            fld_idx = i; break
                    if fld_idx is not None:
                        break

                if fld_idx is not None:
                    body.remove(kids[fld_idx])
                    for off, (lvl, title, pg) in enumerate(entries):
                        body.insert(fld_idx + off, make_toc_paragraph(lvl, title, pg))
                    print(f"Static TOC injected: {len(entries)} entries (replaced field at idx {fld_idx})")
                else:
                    # Mode 2: update existing static TOC (subsequent passes)
                    # Find first static-TOC-entry paragraph, delete all consecutive ones,
                    # then insert fresh entries based on new outline.
                    start_idx = None
                    end_idx = None
                    for i, el in enumerate(kids):
                        if _is_static_toc_entry(el):
                            if start_idx is None:
                                start_idx = i
                            end_idx = i
                        elif start_idx is not None:
                            break
                    if start_idx is None:
                        print("WARN: neither TOC field nor static TOC entries found — left as-is")
                    else:
                        # remove existing entries
                        for el in kids[start_idx:end_idx + 1]:
                            body.remove(el)
                        # insert fresh entries
                        for off, (lvl, title, pg) in enumerate(entries):
                            body.insert(start_idx + off, make_toc_paragraph(lvl, title, pg))
                        print(f"Static TOC UPDATED: {len(entries)} entries (replaced {end_idx - start_idx + 1} old entries at idx {start_idx})")
                data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
            zout.writestr(name, data)
    shutil.move(tmp, docx_out)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f"Usage: {sys.argv[0]} <docx> <rendered_pdf> [out_docx]"); sys.exit(1)
    src, pdf = sys.argv[1], sys.argv[2]
    dst = sys.argv[3] if len(sys.argv) > 3 else src
    main(src, pdf, dst)
