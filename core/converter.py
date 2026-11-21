"""Pure, modular document and tabular conversion functions."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from core.utils import ConversionError, build_output_path, ensure_input_file


# ---------------------------------------------------------------------------
# Document conversions
# ---------------------------------------------------------------------------


def docx_to_md(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert a .docx file to clean Markdown preserving headings, tables, lists, and code."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".md", output_path=output_path)

    try:
        import mammoth
        from markdownify import markdownify as html_to_md

        with src.open("rb") as f:
            result = mammoth.convert_to_html(f)
        html = result.value
        md_text = html_to_md(
            html,
            heading_style="ATX",
            bullets="-",
            code_language="",
        )
        # Clean excessive blank lines
        lines = [line.rstrip() for line in md_text.splitlines()]
        cleaned = "\n".join(lines).strip() + "\n"
        out.write_text(cleaned, encoding="utf-8")
        return out
    except Exception as exc:
        # Optional pypandoc fallback if installed
        try:
            import pypandoc  # type: ignore

            pypandoc.convert_file(str(src), "gfm", outputfile=str(out))
            return out
        except Exception:
            raise ConversionError(f"Failed to convert DOCX to Markdown: {exc}") from exc


def _add_formatted_runs(paragraph: Any, element: Any) -> None:
    """Recursively add inline formatted runs (bold, italic, code) to a python-docx paragraph."""
    from bs4 import NavigableString, Tag

    for child in element.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                paragraph.add_run(text)
        elif isinstance(child, Tag):
            name = child.name.lower()
            if name in ("strong", "b"):
                run = paragraph.add_run(child.get_text())
                run.bold = True
            elif name in ("em", "i"):
                run = paragraph.add_run(child.get_text())
                run.italic = True
            elif name == "code":
                run = paragraph.add_run(child.get_text())
                run.font.name = "Consolas"
            elif name == "br":
                paragraph.add_run("\n")
            else:
                _add_formatted_runs(paragraph, child)


def md_to_docx(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert a Markdown (.md) file to a styled .docx document."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".docx", output_path=output_path)

    try:
        import markdown
        from bs4 import BeautifulSoup, Tag
        from docx import Document
        from docx.shared import Pt

        md_text = src.read_text(encoding="utf-8-sig")
        html = markdown.markdown(
            md_text,
            extensions=["tables", "fenced_code", "nl2br", "sane_lists"],
        )
        soup = BeautifulSoup(html, "html.parser")
        doc = Document()

        # Set a clean default font that supports Latin & Arabic well
        normal_style = doc.styles["Normal"]
        normal_style.font.name = "Calibri"
        normal_style.font.size = Pt(11)

        for node in soup.children:
            if not isinstance(node, Tag):
                text = str(node).strip()
                if text:
                    doc.add_paragraph(text)
                continue

            tag = node.name.lower()
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level = min(int(tag[1]), 4)
                p = doc.add_heading(level=level)
                _add_formatted_runs(p, node)
            elif tag == "p":
                p = doc.add_paragraph()
                _add_formatted_runs(p, node)
            elif tag in ("ul", "ol"):
                style = "List Bullet" if tag == "ul" else "List Number"
                for li in node.find_all("li", recursive=False):
                    p = doc.add_paragraph(style=style)
                    _add_formatted_runs(p, li)
            elif tag == "pre":
                p = doc.add_paragraph()
                run = p.add_run(node.get_text().rstrip("\n"))
                run.font.name = "Consolas"
                run.font.size = Pt(9.5)
            elif tag == "blockquote":
                p = doc.add_paragraph(style="Quote" if "Quote" in doc.styles else None)
                _add_formatted_runs(p, node)
            elif tag == "table":
                rows = node.find_all("tr")
                if rows:
                    first_cells = rows[0].find_all(["th", "td"])
                    col_count = len(first_cells)
                    if col_count > 0:
                        table = doc.add_table(rows=len(rows), cols=col_count)
                        table.style = "Table Grid"
                        for r_idx, row in enumerate(rows):
                            cells = row.find_all(["th", "td"])
                            for c_idx, cell in enumerate(cells[:col_count]):
                                p = table.cell(r_idx, c_idx).paragraphs[0]
                                _add_formatted_runs(p, cell)
                                if r_idx == 0 and cell.name == "th":
                                    for run in p.runs:
                                        run.bold = True
            else:
                text = node.get_text().strip()
                if text:
                    doc.add_paragraph(text)

        doc.save(str(out))
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert Markdown to DOCX: {exc}") from exc


def docx_to_pdf(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """
    Convert a .docx file to .pdf using local MS Word COM (via docx2pdf),
    falling back to LibreOffice (`soffice`) if Word is unavailable.
    """
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".pdf", output_path=output_path)

    errors: list[str] = []

    # Primary: docx2pdf (MS Word COM on Windows)
    if sys.platform == "win32":
        com_initialized = False
        try:
            import pythoncom  # type: ignore

            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            pass

        try:
            from docx2pdf import convert as d2p_convert

            d2p_convert(str(src), str(out))
            if out.exists() and out.stat().st_size > 0:
                return out
        except Exception as exc:
            errors.append(f"docx2pdf/Word COM: {exc}")
        finally:
            if com_initialized:
                try:
                    import pythoncom  # type: ignore

                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # Fallback: LibreOffice headless
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out.parent), str(src)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=creation_flags,
            )
            generated = out.parent / f"{src.stem}.pdf"
            if generated.exists():
                if generated != out:
                    generated.replace(out)
                return out
        except Exception as exc:
            errors.append(f"LibreOffice: {exc}")

    detail = "; ".join(errors) if errors else "Neither MS Word nor LibreOffice was found."
    raise ConversionError(f"Failed to convert DOCX to PDF: {detail}")


def pdf_to_txt(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Extract clean UTF-8 plain text from a PDF using PyMuPDF (fitz)."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".txt", output_path=output_path)

    try:
        import fitz  # PyMuPDF

        doc = fitz.open(str(src))
        try:
            pages_text: list[str] = []
            for page in doc:
                pages_text.append(page.get_text("text").strip())
            content = "\n\n".join(t for t in pages_text if t)
            out.write_text(content + ("\n" if content else ""), encoding="utf-8")
            return out
        finally:
            doc.close()
    except Exception as exc:
        raise ConversionError(f"Failed to convert PDF to TXT: {exc}") from exc


def pdf_to_md(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Extract structured Markdown from a PDF using PyMuPDF (or pymupdf4llm if installed)."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".md", output_path=output_path)

    # Try pymupdf4llm if installed
    try:
        import pymupdf4llm  # type: ignore

        md_text = pymupdf4llm.to_markdown(str(src))
        out.write_text(md_text, encoding="utf-8")
        return out
    except ImportError:
        pass
    except Exception:
        pass

    # Built-in PyMuPDF structured extraction
    try:
        import fitz

        doc = fitz.open(str(src))
        try:
            sections: list[str] = []
            for page_idx, page in enumerate(doc, start=1):
                data = page.get_text("dict")
                font_sizes: list[float] = []
                for block in data.get("blocks", []):
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("text", "").strip():
                                font_sizes.append(float(span.get("size", 11.0)))

                body_size = sorted(font_sizes)[len(font_sizes) // 2] if font_sizes else 11.0
                page_lines: list[str] = [f"<!-- Page {page_idx} -->"]

                for block in data.get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    for line in block.get("lines", []):
                        spans = line.get("spans", [])
                        text = "".join(s.get("text", "") for s in spans).strip()
                        if not text:
                            continue
                        max_size = max((float(s.get("size", body_size)) for s in spans), default=body_size)
                        if max_size >= body_size * 1.45:
                            page_lines.append(f"# {text}")
                        elif max_size >= body_size * 1.2:
                            page_lines.append(f"## {text}")
                        else:
                            page_lines.append(text)
                    page_lines.append("")

                sections.append("\n".join(page_lines).strip())

            out.write_text("\n\n".join(sections).strip() + "\n", encoding="utf-8")
            return out
        finally:
            doc.close()
    except Exception as exc:
        raise ConversionError(f"Failed to convert PDF to Markdown: {exc}") from exc


def pdf_to_docx(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """
    Convert a .pdf file to .docx using pdf2docx, falling back to PyMuPDF + python-docx
    text extraction if pdf2docx fails or is unavailable.
    """
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".docx", output_path=output_path)

    try:
        from pdf2docx import Converter  # type: ignore

        cv = Converter(str(src))
        try:
            cv.convert(str(out))
        finally:
            cv.close()
        if out.exists() and out.stat().st_size > 0:
            return out
    except Exception:
        pass

    # Fallback: extract text per page via PyMuPDF and build a clean DOCX
    try:
        import pymupdf
        from docx import Document

        pdf_doc = pymupdf.open(str(src))
        try:
            docx_doc = Document()
            for idx, page in enumerate(pdf_doc):
                text = page.get_text("text").strip()
                if text:
                    for para in text.split("\n\n"):
                        cleaned = para.strip()
                        if cleaned:
                            docx_doc.add_paragraph(cleaned)
                if idx < len(pdf_doc) - 1:
                    docx_doc.add_page_break()
            docx_doc.save(str(out))
            return out
        finally:
            pdf_doc.close()
    except Exception as exc:
        raise ConversionError(f"Failed to convert PDF to DOCX: {exc}") from exc


def md_to_pdf(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """
    Convert a Markdown (.md) file directly to .pdf using markdown -> HTML -> PyMuPDF Story,
    requiring no external Word or LibreOffice installation.
    """
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".pdf", output_path=output_path)

    try:
        import markdown
        import pymupdf

        md_text = src.read_text(encoding="utf-8-sig")
        body_html = markdown.markdown(
            md_text,
            extensions=["tables", "fenced_code", "nl2br", "sane_lists"],
        )
        full_html = (
            "<html><head><style>"
            "body { font-family: sans-serif; font-size: 11pt; line-height: 1.45; }"
            "h1, h2, h3 { color: #1f2937; margin-top: 14pt; }"
            "pre, code { font-family: monospace; background: #f3f4f6; padding: 2pt 4pt; }"
            "table { border-collapse: collapse; width: 100%; margin: 8pt 0; }"
            "th, td { border: 1px solid #d1d5db; padding: 4pt 6pt; text-align: left; }"
            "th { background: #f3f4f6; }"
            "</style></head><body>"
            f"{body_html}"
            "</body></html>"
        )

        story = pymupdf.Story(html=full_html)
        mediabox = pymupdf.paper_rect("a4")
        where = mediabox + (50, 50, -50, -50)

        writer = pymupdf.DocumentWriter(str(out))
        try:
            more = True
            while more:
                dev = writer.begin_page(mediabox)
                more, _ = story.place(where)
                story.draw(dev, None)
                writer.end_page()
        finally:
            writer.close()
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert Markdown to PDF: {exc}") from exc


# ---------------------------------------------------------------------------
# Image conversions & compression
# ---------------------------------------------------------------------------

_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
_IMAGE_QUALITY_PRESETS: dict[str, tuple[int | None, int]] = {
    "low": (None, 88),      # Light compression
    "medium": (2048, 72),   # Balanced
    "high": (1280, 50),     # Aggressive compression
}


def convert_image(
    input_path: str | Path,
    target_format: str,
    output_path: str | Path | None = None,
) -> Path:
    """Convert an image between PNG, JPG/JPEG, WebP, and BMP."""
    src = ensure_input_file(input_path)
    fmt = target_format.lower().lstrip(".")
    if f".{fmt}" not in _IMAGE_EXTS:
        raise ConversionError(f"Unsupported target image format: '.{fmt}'")

    out = build_output_path(src, f".{fmt}", output_path=output_path)
    pil_fmt = "JPEG" if fmt in ("jpg", "jpeg") else fmt.upper()

    try:
        from PIL import Image

        with Image.open(src) as im:
            if pil_fmt == "JPEG" and im.mode in ("RGBA", "LA", "P"):
                bg = Image.new("RGB", im.size, (255, 255, 255))
                rgba = im.convert("RGBA")
                bg.paste(rgba, mask=rgba.split()[3])
                im_out = bg
            elif pil_fmt == "JPEG" and im.mode != "RGB":
                im_out = im.convert("RGB")
            else:
                im_out = im

            save_kwargs: dict[str, Any] = {}
            if pil_fmt in ("JPEG", "WEBP"):
                save_kwargs["quality"] = 90
            elif pil_fmt == "PNG":
                save_kwargs["optimize"] = True

            im_out.save(out, format=pil_fmt, **save_kwargs)
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert image to {fmt}: {exc}") from exc


def compress_image(
    input_path: str | Path,
    quality: str = "medium",
    output_path: str | Path | None = None,
) -> Path:
    """Compress an image (PNG, JPG, WebP) with 'low', 'medium', or 'high' compression preset."""
    src = ensure_input_file(input_path)
    preset = quality.lower()
    if preset not in _IMAGE_QUALITY_PRESETS:
        raise ConversionError(f"Invalid quality preset '{quality}'. Choose: low, medium, high.")

    max_dim, q_val = _IMAGE_QUALITY_PRESETS[preset]
    ext = src.suffix.lower()
    out = build_output_path(src, ext, suffix=f"_compressed_{preset}", output_path=output_path)

    try:
        from PIL import Image

        with Image.open(src) as im:
            if max_dim is not None and max(im.size) > max_dim:
                scale = max_dim / float(max(im.size))
                new_size = (max(1, int(im.width * scale)), max(1, int(im.height * scale)))
                im = im.resize(new_size, Image.Resampling.LANCZOS)

            if ext in (".jpg", ".jpeg"):
                if im.mode != "RGB":
                    im = im.convert("RGB")
                im.save(out, format="JPEG", quality=q_val, optimize=True)
            elif ext == ".webp":
                im.save(out, format="WEBP", quality=q_val, method=6)
            elif ext == ".png":
                if preset == "high" and im.mode not in ("P",):
                    im = im.convert("P", palette=Image.Palette.ADAPTIVE, colors=256)
                im.save(out, format="PNG", optimize=True, compress_level=9)
            else:
                im.save(out)
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to compress image: {exc}") from exc


# ---------------------------------------------------------------------------
# Tabular / Data conversions (UTF-8-SIG for Arabic compatibility in CSV)
# ---------------------------------------------------------------------------


def _read_csv_flexible(src: Path) -> Any:
    import pandas as pd

    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin1"):
        try:
            return pd.read_csv(src, encoding=enc)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(src)


def xlsx_to_csv(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert .xlsx to .csv with UTF-8-SIG encoding for Arabic compatibility."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".csv", output_path=output_path)
    try:
        import pandas as pd

        df = pd.read_excel(src, engine="openpyxl")
        df.to_csv(out, index=False, encoding="utf-8-sig")
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert XLSX to CSV: {exc}") from exc


def csv_to_xlsx(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert .csv to .xlsx."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".xlsx", output_path=output_path)
    try:
        df = _read_csv_flexible(src)
        df.to_excel(out, index=False, engine="openpyxl")
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert CSV to XLSX: {exc}") from exc


def xlsx_to_json(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert .xlsx to UTF-8 .json (preserving Arabic characters unescaped)."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".json", output_path=output_path)
    try:
        import pandas as pd

        df = pd.read_excel(src, engine="openpyxl")
        json_str = df.to_json(orient="records", force_ascii=False, indent=2)
        out.write_text(json_str, encoding="utf-8")
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert XLSX to JSON: {exc}") from exc


def json_to_xlsx(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert .json to .xlsx."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".xlsx", output_path=output_path)
    try:
        import pandas as pd

        df = pd.read_json(src, encoding="utf-8")
        df.to_excel(out, index=False, engine="openpyxl")
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert JSON to XLSX: {exc}") from exc


def csv_to_json(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert .csv to UTF-8 .json."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".json", output_path=output_path)
    try:
        df = _read_csv_flexible(src)
        json_str = df.to_json(orient="records", force_ascii=False, indent=2)
        out.write_text(json_str, encoding="utf-8")
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert CSV to JSON: {exc}") from exc


def json_to_csv(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Convert .json to .csv with UTF-8-SIG encoding for Arabic compatibility."""
    src = ensure_input_file(input_path)
    out = build_output_path(src, ".csv", output_path=output_path)
    try:
        import pandas as pd

        df = pd.read_json(src, encoding="utf-8")
        df.to_csv(out, index=False, encoding="utf-8-sig")
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert JSON to CSV: {exc}") from exc


# ---------------------------------------------------------------------------
# Unified dispatcher
# ---------------------------------------------------------------------------

_CONVERSION_MAP = {
    (".docx", "md"): docx_to_md,
    (".md", "docx"): md_to_docx,
    (".md", "pdf"): md_to_pdf,
    (".docx", "pdf"): docx_to_pdf,
    (".pdf", "docx"): pdf_to_docx,
    (".pdf", "md"): pdf_to_md,
    (".pdf", "txt"): pdf_to_txt,
    (".xlsx", "csv"): xlsx_to_csv,
    (".csv", "xlsx"): csv_to_xlsx,
    (".xlsx", "json"): xlsx_to_json,
    (".json", "xlsx"): json_to_xlsx,
    (".csv", "json"): csv_to_json,
    (".json", "csv"): json_to_csv,
}


def convert_file(
    input_path: str | Path,
    target_format: str,
    output_path: str | Path | None = None,
) -> Path:
    """Dispatch conversion based on input extension and target format."""
    src = ensure_input_file(input_path)
    ext = src.suffix.lower()
    target = target_format.lower().lstrip(".")

    if ext in _IMAGE_EXTS and f".{target}" in _IMAGE_EXTS:
        return convert_image(src, target, output_path=output_path)

    handler = _CONVERSION_MAP.get((ext, target))
    if handler is None:
        raise ConversionError(f"Unsupported conversion: '{ext}' -> '.{target}'")
    return handler(src, output_path=output_path)
