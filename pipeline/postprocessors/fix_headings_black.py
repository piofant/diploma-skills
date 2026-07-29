#!/usr/bin/env python3
"""
Post-process: force DIRECT black run colour on every heading paragraph.

Why: the reference template already defines Heading 1..9 styles as black,
but Google Docs' .docx importer remaps named heading styles to ITS OWN
theme (blue) and ignores the style-level colour. Direct (run-level)
character formatting survives the import, so we inject <w:color w:val="000000"/>
into every run of every Heading* / Title / TOC-heading paragraph and strip
theme-colour attributes. Idempotent.
"""
import sys, os, shutil, zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

HEADING_STYLES = {f'Heading{i}' for i in range(1, 10)} | {
    'Heading1','Heading2','Heading3','Heading4','Heading5',
    'Title','Subtitle','TOCHeading',
}


def _style_id(p):
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        return None
    pStyle = pPr.find(f'{W}pStyle')
    if pStyle is None:
        return None
    return pStyle.get(f'{W}val')


def _force_black_run(r):
    rPr = r.find(f'{W}rPr')
    if rPr is None:
        rPr = ET.Element(f'{W}rPr')
        r.insert(0, rPr)
    color = rPr.find(f'{W}color')
    if color is None:
        color = ET.SubElement(rPr, f'{W}color')
    color.set(f'{W}val', '000000')
    # kill any theme colour so Google can't re-tint
    for a in (f'{W}themeColor', f'{W}themeTint', f'{W}themeShade'):
        if a in color.attrib:
            del color.attrib[a]


def _strip_bold_italic_run(r):
    """Hard-set b=0 and i=0 on a run, overriding any direct char formatting.

    Pandoc generates direct <w:b/> / <w:i/> on heading runs for some H-levels,
    which overrides the style-level bold=False. We hard-set b w:val=0 to nuke it.
    Requirement of the ГЭК secretary: H4+ (третий и далее уровень нумерованного
    заголовка) не выделяется полужирным.
    """
    rPr = r.find(f'{W}rPr')
    if rPr is None:
        rPr = ET.Element(f'{W}rPr')
        r.insert(0, rPr)
    # Bold off
    b = rPr.find(f'{W}b')
    if b is None:
        b = ET.SubElement(rPr, f'{W}b')
    b.set(f'{W}val', '0')
    # bCs (complex script bold) off too
    bcs = rPr.find(f'{W}bCs')
    if bcs is None:
        bcs = ET.SubElement(rPr, f'{W}bCs')
    bcs.set(f'{W}val', '0')
    # Italic off
    i = rPr.find(f'{W}i')
    if i is None:
        i = ET.SubElement(rPr, f'{W}i')
    i.set(f'{W}val', '0')
    iCs = rPr.find(f'{W}iCs')
    if iCs is None:
        iCs = ET.SubElement(rPr, f'{W}iCs')
    iCs.set(f'{W}val', '0')


def _heading_level(sid):
    """Return 1-9 for Heading{N}/Heading {N}, else None."""
    if not sid:
        return None
    import re
    m = re.match(r'^Heading\s*([1-9])$', sid)
    if m:
        return int(m.group(1))
    return None


def fix(xml_bytes):
    root = ET.fromstring(xml_bytes)
    n_par = n_run = n_h4_stripped = 0
    for p in root.iter(f'{W}p'):
        sid = _style_id(p)
        if sid and (sid in HEADING_STYLES or sid.startswith('Heading')):
            n_par += 1
            lvl = _heading_level(sid)
            for r in p.iter(f'{W}r'):
                _force_black_run(r)
                n_run += 1
                # H4+ (`#### 1.1.1` и глубже) — НЕ bold по требованию ГЭК вуза
                if lvl is not None and lvl >= 4:
                    _strip_bold_italic_run(r)
                    n_h4_stripped += 1
    print(f"  fix_headings: H4+ runs un-bolded/un-italicised: {n_h4_stripped}")
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True), n_par, n_run


def main(in_path, out_path):
    tmp = out_path + '.tmp'
    npar = nrun = 0
    with zipfile.ZipFile(in_path, 'r') as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == 'word/document.xml':
                data, npar, nrun = fix(data)
            zout.writestr(name, data)
    shutil.move(tmp, out_path)
    print(f"Headings forced black: {npar} heading paragraphs, {nrun} runs -> {out_path}")


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else None
    dst = sys.argv[2] if len(sys.argv) > 2 else src
    if not src or not os.path.exists(src):
        print(f"Usage: {sys.argv[0]} input.docx [output.docx]"); sys.exit(1)
    main(src, dst)
