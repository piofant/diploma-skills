#!/bin/bash
# Сборка ВКР БЕЗ титульного листа для заливки в ИСУ.
# секретарь ГЭК: «он у вас есть автоматически прикрепляется системой» — ИСУ
# сам приклеивает свой титул/задание/аннотацию. Мой файл должен начинаться
# с СОДЕРЖАНИЯ.
#
# Output: VKR_Lutsenko_XXXXX_no_title.{docx,pdf}

set -e

DIPLOMA_DIR="/path/to/your/diploma"
SRC="$DIPLOMA_DIR/diploma_body.md"
BUILD_DIR="$DIPLOMA_DIR/prs_defence_21-04/docx_build"
OUT_DOCX="$DIPLOMA_DIR/VKR_Lutsenko_XXXXX_no_title.docx"
OUT_PDF="$DIPLOMA_DIR/VKR_Lutsenko_XXXXX_no_title.pdf"
TEMPLATE="$BUILD_DIR/template.docx"
STRIPPED_MD="$BUILD_DIR/_body_stripped.md"

mkdir -p "$BUILD_DIR"

echo ">> [1/8] Генерация template.docx (без titlepage)"
python3 "$BUILD_DIR/make_reference.py"

echo ">> [2/8] Чистка markdown (убрать первый H1)"
awk '
  BEGIN { skipped_h1 = 0 }
  /^# Разработка/ && !skipped_h1 { skipped_h1 = 1; next }
  { print }
' "$SRC" > "$STRIPPED_MD"

echo ">> [3/8] pandoc: md -> body.docx"
pandoc "$STRIPPED_MD" \
  -o "$BUILD_DIR/body.docx" \
  --from markdown \
  --to docx \
  --reference-doc="$TEMPLATE" \
  --resource-path="$DIPLOMA_DIR" \
  --columns=200 \
  --default-image-extension=png \
  --wrap=none

echo ">> [4/8] Сборка БЕЗ титула + TOC-поле + СОДЕРЖАНИЕ"
python3 "$BUILD_DIR/merge_no_title.py"

echo ">> [5/8] Пост-обработка: tables + headings + formatting + page numbers"
python3 "$BUILD_DIR/fix_tables.py" "$BUILD_DIR/diploma_body.docx"
python3 "$BUILD_DIR/fix_headings_black.py" "$BUILD_DIR/diploma_body.docx"
python3 "$BUILD_DIR/fix_indent_abbrev.py" "$BUILD_DIR/diploma_body.docx"
python3 "$BUILD_DIR/fix_formatting_finishing.py" "$BUILD_DIR/diploma_body.docx"
python3 "$BUILD_DIR/add_page_numbers_no_title.py" "$BUILD_DIR/diploma_body.docx"

cp "$BUILD_DIR/diploma_body.docx" "$OUT_DOCX"

echo ">> [6/8] LibreOffice: docx -> pdf (+ обновление TOC)"
python3 "$BUILD_DIR/update_and_convert.py" "$OUT_DOCX" "$OUT_PDF"

echo ">> [7/8] Static TOC из outline PDF (Google Docs не вычисляет TOC-поле)"
python3 "$BUILD_DIR/gen_static_toc.py" "$OUT_DOCX" "$OUT_PDF"

echo ">> [8/10] Финальный PDF из docx со static TOC"
python3 "$BUILD_DIR/update_and_convert.py" "$OUT_DOCX" "$OUT_PDF"

# Сходимость: static TOC может занять больше страниц чем LibreOffice TOC-поле,
# из-за чего номера в TOC съезжают на 1-2 страницы. Повторный gen_static_toc
# с new outline + rerender выравнивает.
echo ">> [9/10] Convergence pass: re-generate static TOC from final outline"
python3 "$BUILD_DIR/gen_static_toc.py" "$OUT_DOCX" "$OUT_PDF"
echo ">> [10/10] LibreOffice: re-render PDF после convergence"
python3 "$BUILD_DIR/update_and_convert.py" "$OUT_DOCX" "$OUT_PDF"

echo ""
echo ">> Готово (БЕЗ титула, для ИСУ):"
echo "   DOCX: $OUT_DOCX"
echo "   PDF:  $OUT_PDF"
ls -la "$OUT_DOCX" "$OUT_PDF" 2>/dev/null
