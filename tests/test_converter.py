"""Unit tests for core.converter including UTF-8 Arabic text preservation."""

from __future__ import annotations

import json
from pathlib import Path

import fitz
import pandas as pd
import pytest
from docx import Document

from core.converter import (
    compress_image,
    convert_file,
    convert_image,
    csv_to_json,
    csv_to_xlsx,
    docx_to_md,
    json_to_csv,
    json_to_xlsx,
    md_to_docx,
    md_to_pdf,
    pdf_to_docx,
    pdf_to_md,
    pdf_to_txt,
    xlsx_to_csv,
    xlsx_to_json,
)
from core.utils import ConversionError

ARABIC_SAMPLE = "مرحبا بالعالم - اختبار التحويل"


def test_docx_to_md_and_back_with_arabic(tmp_path: Path) -> None:
    docx_path = tmp_path / "sample.docx"
    doc = Document()
    doc.add_heading("عنوان رئيسي", level=1)
    doc.add_paragraph(f"فقرة تجريبية: {ARABIC_SAMPLE}")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "الاسم"
    table.cell(0, 1).text = "القيمة"
    table.cell(1, 0).text = "اختبار"
    table.cell(1, 1).text = "123"
    doc.save(str(docx_path))

    md_path = docx_to_md(docx_path)
    assert md_path.exists()
    assert md_path.parent == tmp_path
    md_content = md_path.read_text(encoding="utf-8")
    assert "عنوان رئيسي" in md_content
    assert ARABIC_SAMPLE in md_content
    assert "الاسم" in md_content

    # Convert back to docx
    out_docx = md_to_docx(md_path)
    assert out_docx.exists()
    assert out_docx != docx_path  # Never overwrite original
    reloaded = Document(str(out_docx))
    all_text = "\n".join(p.text for p in reloaded.paragraphs)
    assert ARABIC_SAMPLE in all_text


def test_md_to_docx_formatting(tmp_path: Path) -> None:
    md_path = tmp_path / "notes.md"
    md_path.write_text(
        "# Heading One\n\n"
        "Paragraph with **bold** and *italic* and `code`.\n\n"
        "- Item A\n- Item B\n\n"
        "| Col1 | Col2 |\n| --- | --- |\n| A | B |\n",
        encoding="utf-8",
    )

    out_docx = md_to_docx(md_path)
    assert out_docx.exists()
    doc = Document(str(out_docx))
    assert any("Heading One" in p.text for p in doc.paragraphs)
    assert len(doc.tables) == 1


def test_pdf_to_txt_and_md(tmp_path: Path) -> None:
    pdf_path = tmp_path / "doc.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Main Title", fontsize=22)
    page.insert_text((72, 120), "Body paragraph line.", fontsize=11)
    doc.save(str(pdf_path))
    doc.close()

    txt_path = pdf_to_txt(pdf_path)
    assert txt_path.exists()
    txt = txt_path.read_text(encoding="utf-8")
    assert "Main Title" in txt
    assert "Body paragraph line." in txt

    md_path = pdf_to_md(pdf_path)
    assert md_path.exists()
    md = md_path.read_text(encoding="utf-8")
    assert "# Main Title" in md
    assert "Body paragraph line." in md


def test_tabular_roundtrips_with_arabic_utf8_sig(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "data.xlsx"
    df = pd.DataFrame(
        [
            {"الاسم": "أحمد", "المدينة": "الرياض", "الدرجة": 95},
            {"الاسم": "سارة", "المدينة": "جدة", "الدرجة": 98},
        ]
    )
    df.to_excel(xlsx_path, index=False, engine="openpyxl")

    # XLSX -> CSV (verify UTF-8-SIG BOM for Excel Arabic compatibility)
    csv_path = xlsx_to_csv(xlsx_path)
    assert csv_path.exists()
    raw_bytes = csv_path.read_bytes()
    assert raw_bytes.startswith(b"\xef\xbb\xbf")
    csv_text = csv_path.read_text(encoding="utf-8-sig")
    assert "الرياض" in csv_text

    # CSV -> JSON
    json_path = csv_to_json(csv_path)
    assert json_path.exists()
    records = json.loads(json_path.read_text(encoding="utf-8"))
    assert records[0]["الاسم"] == "أحمد"
    assert records[1]["المدينة"] == "جدة"

    # JSON -> XLSX
    xlsx_out = json_to_xlsx(json_path)
    assert xlsx_out.exists()
    df_reloaded = pd.read_excel(xlsx_out, engine="openpyxl")
    assert df_reloaded.iloc[0]["المدينة"] == "الرياض"

    # XLSX -> JSON -> CSV -> XLSX
    json2 = xlsx_to_json(xlsx_path)
    csv2 = json_to_csv(json2)
    xlsx2 = csv_to_xlsx(csv2)
    assert json2.exists() and csv2.exists() and xlsx2.exists()


def test_convert_file_dispatcher_and_errors(tmp_path: Path) -> None:
    missing = tmp_path / "missing.docx"
    with pytest.raises(ConversionError):
        convert_file(missing, "md")

    txt = tmp_path / "plain.txt"
    txt.write_text("hello", encoding="utf-8")
    with pytest.raises(ConversionError):
        convert_file(txt, "xlsx")


def test_md_to_pdf_and_pdf_to_docx(tmp_path: Path) -> None:
    md_path = tmp_path / "report.md"
    md_path.write_text(
        "# Quarterly Summary\n\n"
        "This is a **local** markdown to PDF test.\n\n"
        "- Fast\n- Offline\n",
        encoding="utf-8",
    )

    pdf_path = md_to_pdf(md_path)
    assert pdf_path.exists()
    assert pdf_path.suffix == ".pdf"
    with fitz.open(str(pdf_path)) as doc:
        assert len(doc) >= 1
        assert "Quarterly Summary" in doc[0].get_text("text")

    docx_path = pdf_to_docx(pdf_path)
    assert docx_path.exists()
    assert docx_path.suffix == ".docx"
    reloaded = Document(str(docx_path))
    full_text = "\n".join(p.text for p in reloaded.paragraphs)
    assert "Quarterly Summary" in full_text


def test_image_conversions_and_compression(tmp_path: Path) -> None:
    from PIL import Image, ImageDraw

    png_path = tmp_path / "sample.png"
    img = Image.new("RGBA", (1000, 800), (20, 100, 220, 200))
    draw = ImageDraw.Draw(img)
    for x in range(0, 1000, 25):
        draw.line([(x, 0), (x, 800)], fill=(255, 255, 255, 255), width=3)
    img.save(png_path, format="PNG")

    jpg_path = convert_image(png_path, "jpg")
    assert jpg_path.exists()
    assert jpg_path.suffix == ".jpg"

    webp_path = convert_file(jpg_path, "webp")
    assert webp_path.exists()
    assert webp_path.suffix == ".webp"

    png_back = convert_file(webp_path, "png")
    assert png_back.exists()
    assert png_back.suffix == ".png"

    compressed_jpg = compress_image(jpg_path, quality="high")
    assert compressed_jpg.exists()
    assert compressed_jpg.stat().st_size < jpg_path.stat().st_size
