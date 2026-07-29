#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Post-process DOCX: remove first-line indent in СПИСОК СОКРАЩЕНИЙ and ТЕРМИНЫ sections.
Required by ГЭК-секретарь (секретаря ГЭК, 2026-05-14):
«В Терминах... и Сокращениях ... нет красной строки нигде»
"""
import sys
from docx import Document
from docx.shared import Cm

ABBREV_TITLE = 'СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ'
TERMS_TITLE = 'ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ'
INTRO_TITLE = 'ВВЕДЕНИЕ'

def main(path):
    doc = Document(path)
    in_target_section = False
    fixed_count = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        style_name = p.style.name if p.style else ''
        is_heading = 'Heading' in style_name or 'Title' in style_name

        if is_heading:
            # Enter / exit target sections
            if text in (ABBREV_TITLE, TERMS_TITLE):
                in_target_section = True
                continue
            elif in_target_section:
                # Hit a different heading → exit
                in_target_section = False

        if in_target_section and text:
            # Remove first-line indent
            p.paragraph_format.first_line_indent = Cm(0)
            fixed_count += 1

    doc.save(path)
    print(f'  Removed first-line indent on {fixed_count} paragraphs (Сокращения + Термины)')

if __name__ == '__main__':
    main(sys.argv[1])
