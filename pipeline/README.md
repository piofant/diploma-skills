# Build-pipeline: `markdown → docx → pdf`

Рабочий пайплайн сборки ВКР из одного markdown-файла. Тот самый, о котором говорят скиллы `diploma-format` и `diploma-antiplag` — здесь он целиком, а не описанием.

Собирает `.docx` и `.pdf`, которые:
- не разваливаются при импорте в Google Docs (это отдельный ад, см. [TROUBLESHOOTING.md](TROUBLESHOOTING.md))
- имеют работающее оглавление с номерами страниц
- проходят формальные требования: TNR 14, интервал 1,5, отступ 1,25 см, поля 30/15/20/20 мм

---

## Зависимости

```bash
brew install pandoc          # или apt install pandoc
brew install --cask libreoffice
pip install python-docx docxcompose pypdf
```

`soffice` должен быть в `PATH`.

---

## Запуск

```bash
# поправьте пути в начале build.sh под свой проект
bash build.sh
```

На вход — `diploma_body.md`, на выход — `.docx` + `.pdf`.

---

## Как это работает

```
diploma_body.md
      │
      ├─ [1] make_reference.py ──────→ template.docx (стили: TNR 14, 1.5, отступы, поля)
      │
      ├─ [2] awk ────────────────────→ убрать первый H1 (тема — на титуле)
      │
      ├─ [3] pandoc --reference-doc ─→ body.docx
      │
      ├─ [4] merge_no_title.py ──────→ + «СОДЕРЖАНИЕ» + TOC-поле + разрыв страницы
      │
      ├─ [5] пост-обработка:
      │       fix_tables.py               границы и фикс-ширина таблиц
      │       fix_headings_black.py       чёрный цвет на runs + снять bold с H4+
      │       fix_indent_abbrev.py        убрать красную строку в «Сокращениях»
      │       fix_formatting_finishing.py структурные элементы, главы, приложения, подписи таблиц
      │       add_page_numbers_no_title.py footer с PAGE-полем
      │
      ├─ [6] update_and_convert.py ──→ PDF (LibreOffice раскрывает TOC, строит outline)
      │
      ├─ [7] gen_static_toc.py ──────→ статический TOC из outline PDF
      │
      ├─ [8] update_and_convert.py ──→ PDF заново
      │
      ├─ [9] gen_static_toc.py ──────→ convergence-pass (номера съезжают после [7])
      │
      └─ [10] update_and_convert.py ─→ финальный PDF
```

Шаги 7–10 выглядят избыточно, но нужны: статический TOC занимает больше страниц, чем поле-заглушка, из-за чего вся пагинация сдвигается. Второй проход выравнивает.

---

## Что делает каждый скрипт

| Скрипт | Задача |
|---|---|
| `make_reference.py` | Генерит `template.docx` — reference-doc для pandoc. Все стили: Normal, Heading 1–9, TOC-стили, Compact. H4+ принудительно не-bold. |
| `make_titlepage.py` | Титульный лист отдельным `.docx`. **Нужен не всем** — некоторые вузовские системы приклеивают титул сами. |
| `merge_no_title.py` | Склейка + вставка «СОДЕРЖАНИЕ» и Word-поля `{TOC}`. Версия без титула. |
| `fix_tables.py` | Фиксированная ширина на всю страницу + видимые границы. Без этого Google Docs схлопывает borderless-таблицы pandoc в 1–2 символа. |
| `fix_headings_black.py` | Прямой `<w:color w:val="000000"/>` на каждом run заголовка + `<w:b w:val="0"/>` для H4+. Style-level цвет Google Docs игнорирует. |
| `fix_indent_abbrev.py` | Снимает `first-line indent` в разделах «Сокращения» и «Термины». |
| `fix_formatting_finishing.py` | Структурные элементы (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ…) — по центру, без отступа, с новой страницы. Главы — с новой страницы. Приложения — «Приложение А» + название bold ниже. Подписи таблиц — слева без отступа. |
| `add_page_numbers_no_title.py` | Footer с полем `PAGE`, `pgNumType start="2"` (если титул приклеивает система). |
| `gen_static_toc.py` | Читает outline из отрендеренного PDF через `pypdf`, генерит статические абзацы TOC с dot-leader табами. Умеет обновлять уже вставленный TOC (второй проход). |
| `update_and_convert.py` | `docx → pdf` через LibreOffice headless. |

---

## Адаптация под свой вуз

1. **Стили** — `make_reference.py`, функция `configure_style()`. Размеры, интервалы, отступы, bold по уровням.
2. **Поля страницы** — там же, секция `section.*_margin`.
3. **Титул** — `make_titlepage.py` целиком под свой шаблон, либо выкинуть если система генерит сама.
4. **Нумерация** — `add_page_numbers_no_title.py`, параметр `pgNumType w:start`.
5. **Структурные элементы** — `fix_formatting_finishing.py`, константа `STRUCTURAL_TEXTS`.

---

## Известные грабли

Отдельным файлом: **[TROUBLESHOOTING.md](TROUBLESHOOTING.md)** — 7 проблем, каждая стоила вечера.
