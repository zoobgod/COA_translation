"""
Word document generation module.

Creates a formatted Russian-language COA Word document from translated text,
using a pre-prepared docx template via docxtpl, or building one from scratch
with python-docx if no template is available.
"""

import io
import logging
import os
import re
from datetime import datetime

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docxtpl import DocxTemplate

logger = logging.getLogger(__name__)

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")
TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "coa_template.docx")


def generate_doc_from_template(
    translated_text: str,
    original_filename: str,
    extraction_method: str,
    model_used: str,
) -> bytes:
    """
    Generate a Word document using the docxtpl template.

    Args:
        translated_text: The translated Russian text
        original_filename: Name of the original PDF file
        extraction_method: Which PDF extraction method was used
        model_used: Which OpenAI model was used for translation

    Returns:
        bytes of the generated .docx file
    """
    if os.path.exists(TEMPLATE_PATH):
        return _generate_from_template(
            translated_text, original_filename, extraction_method, model_used
        )
    else:
        return _generate_from_scratch(
            translated_text, original_filename, extraction_method, model_used
        )


def _generate_from_template(
    translated_text: str,
    original_filename: str,
    extraction_method: str,
    model_used: str,
) -> bytes:
    """Generate document using the docxtpl template."""
    doc = DocxTemplate(TEMPLATE_PATH)

    # Parse the translated text into sections
    sections = _parse_translated_sections(translated_text)

    context = {
        "title": "СЕРТИФИКАТ АНАЛИЗА",
        "subtitle": "Перевод на русский язык",
        "original_filename": original_filename,
        "translation_date": datetime.now().strftime("%d.%m.%Y"),
        "model_used": model_used,
        "extraction_method": extraction_method,
        "translated_content": translated_text,
        "sections": sections,
    }

    doc.render(context)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _generate_from_scratch(
    translated_text: str,
    original_filename: str,
    extraction_method: str,
    model_used: str,
) -> bytes:
    """Generate a formatted Word document from scratch using python-docx."""
    doc = Document()

    # Configure default font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(11)

    # Configure page margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(0.8)

    # --- Title ---
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run("СЕРТИФИКАТ АНАЛИЗА")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_run.font.name = "Times New Roman"

    # --- Subtitle ---
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = subtitle_para.add_run("(Перевод на русский язык)")
    sub_run.font.size = Pt(12)
    sub_run.font.name = "Times New Roman"
    sub_run.font.color.rgb = RGBColor(100, 100, 100)

    # --- Metadata ---
    doc.add_paragraph()  # spacer
    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    meta_items = [
        ("Исходный файл:", original_filename),
        ("Дата перевода:", datetime.now().strftime("%d.%m.%Y")),
        ("Модель перевода:", model_used),
        ("Метод извлечения:", extraction_method),
    ]

    for label, value in meta_items:
        p = doc.add_paragraph()
        label_run = p.add_run(label + " ")
        label_run.bold = True
        label_run.font.size = Pt(9)
        label_run.font.name = "Times New Roman"
        label_run.font.color.rgb = RGBColor(80, 80, 80)
        value_run = p.add_run(value)
        value_run.font.size = Pt(9)
        value_run.font.name = "Times New Roman"
        value_run.font.color.rgb = RGBColor(80, 80, 80)

    # --- Divider ---
    divider = doc.add_paragraph()
    divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
    div_run = divider.add_run("─" * 60)
    div_run.font.size = Pt(8)
    div_run.font.color.rgb = RGBColor(180, 180, 180)

    # --- Translated Content ---
    _add_translated_content(doc, translated_text)

    # --- Footer divider ---
    doc.add_paragraph()
    divider2 = doc.add_paragraph()
    divider2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    div_run2 = divider2.add_run("─" * 60)
    div_run2.font.size = Pt(8)
    div_run2.font.color.rgb = RGBColor(180, 180, 180)

    # --- Disclaimer ---
    disclaimer = doc.add_paragraph()
    disclaimer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    disc_run = disclaimer.add_run(
        "Данный документ является переводом оригинального Сертификата анализа.\n"
        "Перевод выполнен с использованием искусственного интеллекта с применением\n"
        "фармацевтического глоссария. Рекомендуется верификация специалистом."
    )
    disc_run.font.size = Pt(8)
    disc_run.font.name = "Times New Roman"
    disc_run.font.color.rgb = RGBColor(140, 140, 140)
    disc_run.italic = True

    # Save to bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _add_translated_content(doc: Document, translated_text: str):
    """
    Parse the translated text and add it to the document with
    appropriate formatting for headers, tables, and body text.
    """
    lines = translated_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if not line:
            i += 1
            continue

        # Detect page separators
        if re.match(r"^---\s*(Страница|Page)\s*\d+", line, re.IGNORECASE):
            page_para = doc.add_paragraph()
            page_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            page_run = page_para.add_run(line.strip("- "))
            page_run.font.size = Pt(9)
            page_run.font.color.rgb = RGBColor(120, 120, 120)
            page_run.italic = True
            i += 1
            continue

        # Detect table rows (lines with | separator)
        if "|" in line and _looks_like_table_row(line):
            table_rows = []
            while i < len(lines) and "|" in lines[i] and _looks_like_table_row(lines[i]):
                cells = [c.strip() for c in lines[i].split("|") if c.strip()]
                table_rows.append(cells)
                i += 1

            if table_rows:
                _add_table(doc, table_rows)
            continue

        # Detect section headers (all caps or short bold-looking lines)
        if _is_section_header(line):
            heading = doc.add_heading(level=2)
            heading_run = heading.add_run(line)
            heading_run.font.name = "Times New Roman"
            heading_run.font.size = Pt(12)
            i += 1
            continue

        # Regular text paragraph
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        i += 1


def _looks_like_table_row(line: str) -> bool:
    """Check if a line looks like a table row."""
    parts = line.split("|")
    return len(parts) >= 2


def _is_section_header(line: str) -> bool:
    """Heuristic to detect section headers."""
    stripped = line.strip(":-= ")
    if not stripped:
        return False
    # All uppercase and relatively short
    if stripped.isupper() and len(stripped) < 80:
        return True
    # Common header patterns
    if re.match(r"^\d+\.\s+[А-ЯA-Z]", stripped):
        return True
    return False


def _add_table(doc: Document, rows: list):
    """Add a formatted table to the document."""
    if not rows:
        return

    max_cols = max(len(row) for row in rows)

    # Normalize rows to have the same number of columns
    for row in rows:
        while len(row) < max_cols:
            row.append("")

    table = doc.add_table(rows=len(rows), cols=max_cols)
    table.style = "Table Grid"

    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i, j)
            cell.text = cell_text

            # Format cell text
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(10)

            # Bold the header row
            if i == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

    doc.add_paragraph()  # spacer after table


def _parse_translated_sections(text: str) -> list[dict]:
    """Parse translated text into sections for template rendering."""
    sections = []
    current_section = {"title": "", "content": ""}

    for line in text.split("\n"):
        stripped = line.strip()
        if _is_section_header(stripped) and stripped:
            if current_section["content"].strip():
                sections.append(current_section)
            current_section = {"title": stripped, "content": ""}
        else:
            current_section["content"] += line + "\n"

    if current_section["content"].strip() or current_section["title"]:
        sections.append(current_section)

    return sections
