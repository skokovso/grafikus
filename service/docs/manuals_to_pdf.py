#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""manuals_to_pdf.py — рендерит руководства из Markdown в PDF:
нумерованные разделы (1., 1.1., 1.1.1), таблицы, списки, жирные акценты.
Источники: project_info/manuals/*.md; если папка пуста — USER_MANUAL*.md
в корне проекта (поддержка старой раскладки). PDF: project_info/manuals/pdf/.
Символы, которых нет в Arial/DejaVu (эмодзи, ✓, ⬇, ), заменяются
печатными эквивалентами — в печати не будет пустых квадратов.
Требует: pip install fpdf2
Запуск: python service/docs/manuals_to_pdf.py
"""
import sys
from pathlib import Path

try:
    from fpdf import FPDF
except ImportError:
    sys.exit("Нужна библиотека fpdf2:  pip install fpdf2")


def _find_root(start: Path) -> Path:
    """Корень проекта: ближайший предок с .git; иначе папка скрипта."""
    for p in [start, *start.parents]:
        if (p / ".git").exists():
            return p
    return start.parent


BASE_DIR = _find_root(Path(__file__).resolve())
MANUALS_DIR = BASE_DIR / "project_info" / "manuals"
PDF_DIR = MANUALS_DIR / "pdf"

FONT_CANDIDATES = [
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf",
     "C:/Windows/Fonts/cour.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"),
]

# Замена символов, которых нет в Arial/DejaVu, на печатные эквиваленты.
# MD-источники сохраняют эмодзи для веба; в PDF они читаются словами.
GLYPH_REPL = [
    ("🎯 ", ""), ("🔓 ", ""), ("📋 ", ""), ("📄 ", ""), ("✎ ", ""), ("⬇ ", ""),
    ("👁 ", ""), ("👁", ""), ("⬇", ""), ("🎯", ""), ("🔓", ""), ("📋", ""),
    ("📄", ""), ("✎", ""),
    ("🗑", "[удалить]"), ("✓", "[отметка]"), ("＋", "+"),
]


def sanitize(text: str) -> str:
    for old, new in GLYPH_REPL:
        text = text.replace(old, new)
    return (text.replace("  ", " ").replace(" ;", ";")
                .replace(" ,", ",").replace("( ", "(").replace(" )", ")"))


class ManualPDF(FPDF):
    def __init__(self, doc_title):
        super().__init__()
        self.doc_title = doc_title
        self.set_margins(15, 15, 15)
        self.set_auto_page_break(auto=True, margin=18)

    def footer(self):
        self.set_y(-12)
        self.set_font("reg", "", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 8, f"— {self.page_no()} —", align="C")

    def title_block(self, text):
        self.set_font("reg", "B", 16)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 8, text, markdown=False)
        self.ln(2)
        self.set_draw_color(180, 180, 180)
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(4)

    def heading(self, text, size):
        self.ln(2)
        self.set_font("reg", "B", size)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, size * 0.55, text, markdown=False)
        self.ln(1)

    def paragraph(self, text):
        self.set_font("reg", "", 10)
        self.set_text_color(30, 30, 30)
        self.multi_cell(0, 5.2, text.replace("`", ""), markdown=True)
        self.ln(1.2)

    def bullet(self, text):
        self.set_font("reg", "", 10)
        self.set_text_color(30, 30, 30)
        x = self.get_x()
        self.set_x(x + 4)
        self.multi_cell(0, 5.2, "•  " + text.replace("`", ""), markdown=True)
        self.ln(0.6)

    def code_block(self, lines):
        self.set_font("mono", "", 9)
        self.set_fill_color(243, 243, 243)
        self.set_text_color(40, 40, 40)
        for ln in lines:
            self.cell(0, 4.6, "  " + ln, fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def table_block(self, rows):
        self.set_font("reg", "", 9)
        self.set_text_color(30, 30, 30)
        with self.table(borders_layout="ALL", first_row_as_headings=True,
                        line_height=4.8, padding=1.5) as tbl:
            for r in rows:
                tbl.row([c.replace("`", "") for c in r])
        self.ln(3)


def is_sep(row):
    return all(set(c.strip()) <= set("-: ") for c in row if c.strip())


def parse_md(text):
    """MD-подмножество -> список блоков (type, payload)."""
    blocks, buf, fence, table = [], [], False, []
    n2 = n3 = n4 = 0
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.startswith("```"):
            if fence:
                blocks.append(("code", buf))
                buf = []
            fence = not fence
            continue
        if fence:
            buf.append(line)
            continue
        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if not is_sep(cells):
                table.append(cells)
            continue
        if table:
            blocks.append(("table", table))
            table = []
        if not line.strip():
            continue
        if line.startswith("# "):
            blocks.append(("title", line[2:].strip()))
        elif line.startswith("## "):
            n2 += 1
            n3 = n4 = 0
            blocks.append(("h2", f"{n2}. {line[3:].strip()}"))
        elif line.startswith("### "):
            n3 += 1
            n4 = 0
            blocks.append(("h3", f"{n2}.{n3}. {line[4:].strip()}"))
        elif line.startswith("#### "):
            n4 += 1
            blocks.append(("h4", f"{n2}.{n3}.{n4}. {line[5:].strip()}"))
        elif line.startswith("- "):
            blocks.append(("bullet", line[2:].strip()))
        else:
            blocks.append(("p", line.strip()))
    if table:
        blocks.append(("table", table))
    return blocks


def render(md_path: Path, pdf_path: Path):
    blocks = parse_md(sanitize(md_path.read_text(encoding="utf-8")))
    doc = ManualPDF(md_path.stem)
    reg, bold, mono = next((f for f in FONT_CANDIDATES if Path(f[0]).exists()),
                           FONT_CANDIDATES[0])
    doc.add_font("reg", "", reg)
    doc.add_font("reg", "B", bold if Path(bold).exists() else reg)
    doc.add_font("mono", "", mono if Path(mono).exists() else reg)
    doc.add_page()
    for kind, payload in blocks:
        if kind == "title":
            doc.title_block(payload)
        elif kind == "h2":
            doc.heading(payload, 13)
        elif kind == "h3":
            doc.heading(payload, 11.5)
        elif kind == "h4":
            doc.heading(payload, 10.5)
        elif kind == "p":
            doc.paragraph(payload)
        elif kind == "bullet":
            doc.bullet(payload)
        elif kind == "code":
            doc.code_block(payload)
        elif kind == "table":
            doc.table_block(payload)
    doc.output(pdf_path)
    print(f"✅ PDF: {pdf_path} ({doc.page_no()} стр.)")


def main():
    sources = sorted(MANUALS_DIR.glob("*.md")) if MANUALS_DIR.exists() else []
    if not sources:
        sources = sorted(BASE_DIR.glob("USER_MANUAL*.md"))
    if not sources:
        sys.exit("Нет руководств: создайте project_info/manuals/*.md "
                 "(или USER_MANUAL*.md в корне) — например, make_docs.py")
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    for md in sources:
        render(md, PDF_DIR / (md.stem + ".pdf"))


if __name__ == "__main__":
    main()