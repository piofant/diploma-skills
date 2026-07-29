#!/usr/bin/env python3
"""
Post-processing fix: set all tables to fixed full-width layout with sane
column widths AND 12pt font in cells (allowed by university requirements for large tables).
Also: disable auto-shrink, prevent narrow Russian-text columns.
"""
import sys
import zipfile
import shutil
import os
from xml.etree import ElementTree as ET


NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# Page width 11906 - left 1701 - right 850 = 9355 DXA
FULL_WIDTH = 9300


def fix_tables_in_xml(xml_content):
    """Update <w:tbl> elements:
    - fixed layout, full width
    - column widths proportional or fallback wide-first
    - 12pt font in cells (suitable for wider content)
    - no abzac indent in cells
    """
    ET.register_namespace('w', NS['w'])
    root = ET.fromstring(xml_content)

    for tbl in root.iter(f'{W}tbl'):
        tblPr = tbl.find(f'{W}tblPr')
        if tblPr is None:
            continue

        # Force table width = full page width in DXA
        tblW = tblPr.find(f'{W}tblW')
        if tblW is None:
            tblW = ET.SubElement(tblPr, f'{W}tblW')
        tblW.set(f'{W}type', 'dxa')
        tblW.set(f'{W}w', str(FULL_WIDTH))

        # Force fixed layout
        tblLayout = tblPr.find(f'{W}tblLayout')
        if tblLayout is None:
            tblLayout = ET.SubElement(tblPr, f'{W}tblLayout')
        tblLayout.set(f'{W}type', 'fixed')

        # Force visible borders on all sides + inside (GOST вуза style)
        tblBorders = tblPr.find(f'{W}tblBorders')
        if tblBorders is None:
            tblBorders = ET.SubElement(tblPr, f'{W}tblBorders')
        else:
            # clear children
            for child in list(tblBorders):
                tblBorders.remove(child)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = ET.SubElement(tblBorders, f'{W}{side}')
            b.set(f'{W}val', 'single')
            b.set(f'{W}sz', '4')   # 0.5pt
            b.set(f'{W}space', '0')
            b.set(f'{W}color', '000000')

        # Disable autoshrink — many Office variants honor this
        tblLook = tblPr.find(f'{W}tblLook')
        if tblLook is None:
            tblLook = ET.SubElement(tblPr, f'{W}tblLook')

        # Recalculate tblGrid column widths
        tblGrid = tbl.find(f'{W}tblGrid')
        if tblGrid is not None:
            cols = tblGrid.findall(f'{W}gridCol')
            n = len(cols)
            if n > 0:
                # Measure max text length per column (across all rows)
                col_lens = [80] * n  # min reasonable
                rows = tbl.findall(f'{W}tr')
                for tr in rows:
                    cells = tr.findall(f'{W}tc')
                    if len(cells) != n:
                        continue
                    for i, tc in enumerate(cells):
                        cell_text = ''.join(t.text or '' for t in tc.iter(f'{W}t'))
                        col_lens[i] = max(col_lens[i], len(cell_text))

                # Read pandoc-suggested widths
                widths = []
                for c in cols:
                    try:
                        widths.append(int(c.get(f'{W}w', '1000')))
                    except (TypeError, ValueError):
                        widths.append(1000)
                total = sum(widths) or n

                # Decide: if columns are clearly uneven by content length, use
                # content-proportional widths; otherwise pandoc-scaled.
                max_len = max(col_lens)
                min_len = min(col_lens)
                content_uneven = max_len > 3 * min_len  # at least 3× spread

                if content_uneven:
                    # Allocate proportional to max content length
                    L_total = sum(col_lens)
                    widths = [max(int(FULL_WIDTH * cl / L_total), 700) for cl in col_lens]
                    widths[-1] = FULL_WIDTH - sum(widths[:-1])
                elif total < 6000:
                    # Pandoc collapsed columns — sensible fallback
                    if n == 4:
                        widths = [4500, 1500, 1500, 1800]
                    elif n == 5:
                        widths = [3700, 1400, 1400, 1400, 1400]
                    elif n == 6:
                        widths = [1800, 2200, 800, 900, 1100, 2500]
                    elif n == 3:
                        widths = [3500, 2900, 2900]
                    else:
                        widths = [FULL_WIDTH // n] * n
                else:
                    scale = FULL_WIDTH / total
                    widths = [max(int(w * scale), 600) for w in widths]
                    widths[-1] = FULL_WIDTH - sum(widths[:-1])
                for c, w in zip(cols, widths):
                    c.set(f'{W}w', str(w))

                # Mark first row as repeating header (so it shows up
                # at the top of every page when the table spans pages)
                all_trs = tbl.findall(f'{W}tr')
                if all_trs:
                    first_tr = all_trs[0]
                    trPr = first_tr.find(f'{W}trPr')
                    if trPr is None:
                        trPr = ET.SubElement(first_tr, f'{W}trPr')
                        first_tr.insert(0, trPr)
                    if trPr.find(f'{W}tblHeader') is None:
                        ET.SubElement(trPr, f'{W}tblHeader')
                    # Also prevent rows from splitting across pages
                    for tr in all_trs:
                        trPr2 = tr.find(f'{W}trPr')
                        if trPr2 is None:
                            trPr2 = ET.SubElement(tr, f'{W}trPr')
                            tr.insert(0, trPr2)
                        if trPr2.find(f'{W}cantSplit') is None:
                            ET.SubElement(trPr2, f'{W}cantSplit')

                # Update each row's cells
                for tr in tbl.findall(f'{W}tr'):
                    cells = tr.findall(f'{W}tc')
                    if len(cells) != n:
                        continue
                    for tc, w in zip(cells, widths):
                        tcPr = tc.find(f'{W}tcPr')
                        if tcPr is None:
                            tcPr = ET.SubElement(tc, f'{W}tcPr')
                            tc.insert(0, tcPr)
                        tcW = tcPr.find(f'{W}tcW')
                        if tcW is None:
                            tcW = ET.SubElement(tcPr, f'{W}tcW')
                        tcW.set(f'{W}type', 'dxa')
                        tcW.set(f'{W}w', str(w))

                        # Set font 12pt + remove abzac indent in cells
                        for p in tc.findall(f'{W}p'):
                            # Remove abzac indent in paragraph properties
                            pPr = p.find(f'{W}pPr')
                            if pPr is None:
                                pPr = ET.SubElement(p, f'{W}pPr')
                                p.insert(0, pPr)
                            # Remove existing ind element to kill abzac indent
                            ind = pPr.find(f'{W}ind')
                            if ind is not None:
                                ind.set(f'{W}firstLine', '0')
                                ind.set(f'{W}left', '0')
                            else:
                                ind = ET.SubElement(pPr, f'{W}ind')
                                ind.set(f'{W}firstLine', '0')
                                ind.set(f'{W}left', '0')
                            # Force 12pt for paragraph runs in cells
                            for r in p.findall(f'{W}r'):
                                rPr = r.find(f'{W}rPr')
                                if rPr is None:
                                    rPr = ET.SubElement(r, f'{W}rPr')
                                    r.insert(0, rPr)
                                sz = rPr.find(f'{W}sz')
                                if sz is None:
                                    sz = ET.SubElement(rPr, f'{W}sz')
                                sz.set(f'{W}val', '24')  # 24 half-pt = 12pt
                                szCs = rPr.find(f'{W}szCs')
                                if szCs is None:
                                    szCs = ET.SubElement(rPr, f'{W}szCs')
                                szCs.set(f'{W}val', '24')

    # ----- Center images and figure captions -----
    DRAWING = f'{W}drawing'
    body = root.find(f'{W}body')
    if body is not None:
        paragraphs = list(body.iter(f'{W}p'))
        for i, p in enumerate(paragraphs):
            # Detect: paragraph contains a <w:drawing> (inline image)
            has_drawing = any(p.iter(DRAWING))
            if has_drawing:
                _set_jc(p, 'center')
                _kill_indent(p)
                # Next paragraph is the caption "Рисунок N – ..." — center it too
                if i + 1 < len(paragraphs):
                    nxt = paragraphs[i + 1]
                    txt = ''.join((t.text or '') for t in nxt.iter(f'{W}t'))
                    if txt.lstrip().startswith('Рисунок'):
                        _set_jc(nxt, 'center')
                        _kill_indent(nxt)

    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _set_jc(p, val):
    """Set <w:jc w:val=val/> in paragraph properties."""
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        pPr = ET.SubElement(p, f'{W}pPr')
        p.insert(0, pPr)
    jc = pPr.find(f'{W}jc')
    if jc is None:
        jc = ET.SubElement(pPr, f'{W}jc')
    jc.set(f'{W}val', val)


def _kill_indent(p):
    """Remove first-line/left indent so centering is not offset."""
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        return
    ind = pPr.find(f'{W}ind')
    if ind is None:
        ind = ET.SubElement(pPr, f'{W}ind')
    ind.set(f'{W}firstLine', '0')
    ind.set(f'{W}left', '0')


def main(in_path, out_path):
    tmp_path = out_path + '.tmp'
    with zipfile.ZipFile(in_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name == 'word/document.xml':
                    data = fix_tables_in_xml(data)
                zout.writestr(name, data)
    shutil.move(tmp_path, out_path)
    print(f"Tables fixed: {out_path}")


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else None
    dst = sys.argv[2] if len(sys.argv) > 2 else src
    if not src or not os.path.exists(src):
        print(f"Usage: {sys.argv[0]} input.docx [output.docx]")
        sys.exit(1)
    main(src, dst)
