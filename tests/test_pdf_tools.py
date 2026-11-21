"""Unit tests for core.pdf_tools and CLI batching."""

from __future__ import annotations

import io
from pathlib import Path

import fitz
from PIL import Image, ImageDraw

from cli import main as cli_main
from core.pdf_tools import (
    compress_pdf,
    extract_pdf_images,
    extract_pdf_text,
    images_to_pdf,
    merge_pdfs,
    split_pdf,
)


def _create_sample_image(path: Path, color: tuple[int, int, int], size: tuple[int, int] = (800, 600)) -> Path:
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    for i in range(0, size[0], 40):
        draw.line([(i, 0), (i, size[1])], fill=(255, 255, 255), width=2)
    img.save(path, format="PNG")
    return path


def _create_sample_pdf(path: Path, num_pages: int = 3, with_image: bool = False) -> Path:
    doc = fitz.open()
    for i in range(1, num_pages + 1):
        page = doc.new_page()
        page.insert_text((72, 72), f"Sample Page {i}", fontsize=16)
        if with_image and i == 1:
            img = Image.new("RGB", (1200, 900), color=(30, 120, 200))
            draw = ImageDraw.Draw(img)
            for x in range(0, 1200, 20):
                draw.line([(x, 0), (x, 900)], fill=(250, 220, 50), width=3)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            page.insert_image(fitz.Rect(72, 120, 472, 420), stream=buf.getvalue())
    doc.save(str(path))
    doc.close()
    return path


def test_merge_and_split_pdf(tmp_path: Path) -> None:
    p1 = _create_sample_pdf(tmp_path / "part1.pdf", num_pages=2)
    p2 = _create_sample_pdf(tmp_path / "part2.pdf", num_pages=2)

    merged = merge_pdfs([p2, p1])
    assert merged.exists()
    with fitz.open(str(merged)) as mdoc:
        assert len(mdoc) == 4

    # Split range
    range_outs = split_pdf(merged, pages="1-2, 4")
    assert len(range_outs) == 1
    with fitz.open(str(range_outs[0])) as rdoc:
        assert len(rdoc) == 3

    # Burst all pages
    burst_outs = split_pdf(merged, burst=True)
    assert len(burst_outs) == 4
    assert all(p.exists() for p in burst_outs)


def test_compress_pdf_and_extract_text(tmp_path: Path) -> None:
    pdf = _create_sample_pdf(tmp_path / "big.pdf", num_pages=1, with_image=True)
    compressed = compress_pdf(pdf, quality="high")
    assert compressed.exists()
    assert compressed.stat().st_size < pdf.stat().st_size

    txt = extract_pdf_text(pdf)
    assert txt.exists()
    assert "Sample Page 1" in txt.read_text(encoding="utf-8")


def test_images_to_pdf_and_extract_images(tmp_path: Path) -> None:
    img1 = _create_sample_image(tmp_path / "01.png", (200, 50, 50))
    img2 = _create_sample_image(tmp_path / "02.png", (50, 180, 90))

    pdf_out = images_to_pdf([img1, img2])
    assert pdf_out.exists()
    with fitz.open(str(pdf_out)) as doc:
        assert len(doc) == 2

    extracted = extract_pdf_images(pdf_out)
    assert len(extracted) == 2
    assert all(p.exists() for p in extracted)


def test_cli_non_interactive_commands(tmp_path: Path) -> None:
    p1 = _create_sample_pdf(tmp_path / "a.pdf", num_pages=1)
    p2 = _create_sample_pdf(tmp_path / "b.pdf", num_pages=1)

    rc = cli_main(["pdf-merge", str(p1), str(p2)])
    assert rc == 0
    assert (tmp_path / "a_merged.pdf").exists()
