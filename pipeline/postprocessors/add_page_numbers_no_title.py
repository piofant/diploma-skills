#!/usr/bin/env python3
"""
Add page numbers WITHOUT title page suppression — for ИСУ upload.
Сквозная нумерация с 1 на каждой странице (СОДЕРЖАНИЕ = 1).
ИСУ автоматически приклеит свой титул/задание/аннотацию перед файлом.

Differs from add_page_numbers.py: NO <w:titlePg/>, NO first-page empty footer.
"""
import sys, os, shutil, zipfile, re

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

FOOTER_PG = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="%s" xmlns:r="%s"><w:p><w:pPr><w:jc w:val="center"/>
<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>1</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>''' % (WNS, RNS)


def main(src, dst):
    tmp = dst + '.tmp'
    with zipfile.ZipFile(src, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    if 'word/footer_pgnum.xml' in files:
        if src != dst:
            shutil.copy(src, dst)
        print("page numbers already present — skip")
        return

    files['word/footer_pgnum.xml'] = FOOTER_PG.encode()

    ct = files['[Content_Types].xml'].decode()
    add = '<Override PartName="/word/footer_pgnum.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    ct = ct.replace('</Types>', add + '</Types>')
    files['[Content_Types].xml'] = ct.encode()

    rp = 'word/_rels/document.xml.rels'
    rels = files[rp].decode()
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    n1 = max(used) + 1
    rid_pg = f'rId{n1}'
    reladd = f'<Relationship Id="{rid_pg}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer_pgnum.xml"/>'
    rels = rels.replace('</Relationships>', reladd + '</Relationships>')
    files[rp] = rels.encode()

    doc = files['word/document.xml'].decode()
    m = re.search(r'(<w:sectPr\b[^>]*>)(.*?)(</w:sectPr>)', doc, re.S)
    if not m:
        raise SystemExit("sectPr not found")
    open_tag, inner, close_tag = m.group(1), m.group(2), m.group(3)
    # remove any existing footer refs / titlePg / pgNumType
    inner = re.sub(r'<w:footerReference\b[^/]*/>', '', inner)
    inner = inner.replace('<w:titlePg/>', '').replace('<w:titlePg />', '')
    inner = re.sub(r'<w:pgNumType\b[^/]*/>', '', inner)
    # add single default footer (no titlePg = number prints on first page too)
    # pgNumType start="2": ИСУ автоматически приклеивает свой титул как стр.1,
    # поэтому моя нумерация стартует с 2 (Содержание = стр.2).
    refs = (f'<w:footerReference xmlns:r="{RNS}" w:type="default" r:id="{rid_pg}"/>'
            f'<w:pgNumType w:start="2"/>')
    new_sect = open_tag + refs + inner + close_tag
    doc = doc[:m.start()] + new_sect + doc[m.end():]
    files['word/document.xml'] = doc.encode()

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.writestr(n, files[n])
        zo.writestr('word/footer_pgnum.xml', files['word/footer_pgnum.xml'])
    shutil.move(tmp, dst)
    print(f"Page numbers added (no title-page suppression) -> {dst}")


if __name__ == '__main__':
    if len(sys.argv) < 2 or not os.path.exists(sys.argv[1]):
        print(f"Usage: {sys.argv[0]} <docx> [out_docx]"); sys.exit(1)
    s = sys.argv[1]
    d = sys.argv[2] if len(sys.argv) > 2 else s
    main(s, d)
