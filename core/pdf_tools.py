"""Local PDF manipulation utilities powered by PyMuPDF (fitz), Pillow, and img2pdf."""

from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Iterable, Literal

from core.utils import ConversionError, build_output_path, ensure_input_file

QualityPreset = Literal["low", "medium", "high"]

_QUALITY_SETTINGS: dict[str, tuple[int, int]] = {
    # (max_image_dimension_px, jpeg_quality)
    "low": (1800, 82),      # Light compression, high visual fidelity
    "medium": (1280, 65),   # Balanced size & quality
    "high": (900, 45),      # Aggressive compression, smallest size
}


def _natural_sort_key(path: Path) -> list[int | str]:
    parts = re.split(r"(\d+)", path.name.lower())
    return [int(p) if p.isdigit() else p for p in parts]


def compress_pdf(
    input_path: str | Path,
    quality: QualityPreset = "medium",
    output_path: str | Path | None = None,
) -> Path:
    """
    Compress a PDF by recompressing embedded raster images and optimizing PDF streams.
    Quality presets: 'low' (light compression), 'medium' (balanced), 'high' (maximum compression).
    """
    src = ensure_input_file(input_path)
    preset = quality.lower()
    if preset not in _QUALITY_SETTINGS:
        raise ConversionError(f"Invalid quality preset '{quality}'. Choose: low, medium, high.")

    max_dim, jpeg_quality = _QUALITY_SETTINGS[preset]
    out = build_output_path(src, ".pdf", suffix=f"_compressed_{preset}", output_path=output_path)

    try:
        import fitz
        from PIL import Image

        doc = fitz.open(str(src))
        try:
            seen_xrefs: set[int] = set()
            for page in doc:
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    if xref in seen_xrefs:
                        continue
                    seen_xrefs.add(xref)

                    try:
                        base_img = doc.extract_image(xref)
                        img_bytes = base_img.get("image")
                        if not img_bytes:
                            continue

                        with Image.open(io.BytesIO(img_bytes)) as pil_img:
                            if pil_img.mode in ("RGBA", "LA", "P"):
                                pil_img = pil_img.convert("RGB")
                            elif pil_img.mode != "RGB":
                                pil_img = pil_img.convert("RGB")

                            w, h = pil_img.size
                            if max(w, h) > max_dim:
                                scale = max_dim / float(max(w, h))
                                new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
                                pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)

                            buf = io.BytesIO()
                            pil_img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
                            new_bytes = buf.getvalue()

                            if len(new_bytes) < len(img_bytes):
                                page.replace_image(xref, stream=new_bytes)
                    except Exception:
                        # Skip individual unreplaceable images (e.g., masks/special colorspaces)
                        continue

            doc.save(
                str(out),
                garbage=4,
                deflate=True,
                clean=True,
            )
            return out
        finally:
            doc.close()
    except Exception as exc:
        raise ConversionError(f"Failed to compress PDF: {exc}") from exc


def merge_pdfs(
    input_paths: Iterable[str | Path],
    output_path: str | Path | None = None,
    sort_by_name: bool = True,
) -> Path:
    """Merge multiple PDF files into a single PDF."""
    paths = [ensure_input_file(p) for p in input_paths]
    if len(paths) < 2:
        raise ConversionError("At least 2 PDF files are required to merge.")

    if sort_by_name:
        paths = sorted(paths, key=_natural_sort_key)

    out = build_output_path(paths[0], ".pdf", suffix="_merged", output_path=output_path)

    try:
        import fitz

        merged = fitz.open()
        try:
            for pdf_path in paths:
                with fitz.open(str(pdf_path)) as src_doc:
                    merged.insert_pdf(src_doc)
            merged.save(str(out), garbage=3, deflate=True)
            return out
        finally:
            merged.close()
    except Exception as exc:
        raise ConversionError(f"Failed to merge PDFs: {exc}") from exc


def _parse_page_ranges(spec: str, total_pages: int) -> list[int]:
    """Parse 1-based page range spec like '1-3, 5' into 0-based page indices."""
    indices: list[int] = []
    for part in spec.split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            start_s, end_s = token.split("-", 1)
            start = int(start_s.strip())
            end = int(end_s.strip())
            if start < 1 or end > total_pages or start > end:
                raise ConversionError(f"Invalid page range '{token}' for {total_pages}-page PDF.")
            indices.extend(range(start - 1, end))
        else:
            page_num = int(token)
            if page_num < 1 or page_num > total_pages:
                raise ConversionError(f"Page {page_num} out of bounds (1-{total_pages}).")
            indices.append(page_num - 1)

    if not indices:
        raise ConversionError("No valid pages specified.")
    return indices


def split_pdf(
    input_path: str | Path,
    pages: str | None = None,
    burst: bool = False,
    output_dir: str | Path | None = None,
) -> list[Path]:
    """
    Split a PDF either by extracting a page range spec (e.g. '1-3, 5')
    or bursting every page into individual PDF files.
    """
    src = ensure_input_file(input_path)
    target_dir = Path(output_dir).expanduser().resolve() if output_dir else src.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    try:
        import fitz

        doc = fitz.open(str(src))
        try:
            total = len(doc)
            if total == 0:
                raise ConversionError("PDF has 0 pages.")

            outputs: list[Path] = []
            if burst or not pages:
                for idx in range(total):
                    single = fitz.open()
                    try:
                        single.insert_pdf(doc, from_page=idx, to_page=idx)
                        out_file = build_output_path(
                            target_dir / src.name,
                            ".pdf",
                            suffix=f"_page_{idx + 1:03d}",
                        )
                        single.save(str(out_file))
                        outputs.append(out_file)
                    finally:
                        single.close()
                return outputs

            selected = _parse_page_ranges(pages, total)
            subset = fitz.open()
            try:
                for idx in selected:
                    subset.insert_pdf(doc, from_page=idx, to_page=idx)
                out_file = build_output_path(
                    target_dir / src.name,
                    ".pdf",
                    suffix="_split",
                )
                subset.save(str(out_file))
                return [out_file]
            finally:
                subset.close()
        finally:
            doc.close()
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Failed to split PDF: {exc}") from exc


def images_to_pdf(
    image_paths: Iterable[str | Path],
    output_path: str | Path | None = None,
    sort_by_name: bool = True,
) -> Path:
    """Convert one or more JPG/PNG/WebP images into a single PDF."""
    paths = [ensure_input_file(p) for p in image_paths]
    if not paths:
        raise ConversionError("No image files provided.")

    if sort_by_name:
        paths = sorted(paths, key=_natural_sort_key)

    out = build_output_path(paths[0], ".pdf", suffix="_images" if len(paths) > 1 else "", output_path=output_path)

    try:
        import img2pdf
        from PIL import Image

        prepared_buffers: list[bytes] = []
        for img_path in paths:
            with Image.open(img_path) as im:
                # img2pdf rejects alpha channels and some WebP modes; normalize to RGB JPEG/PNG bytes
                if im.mode in ("RGBA", "LA", "P"):
                    bg = Image.new("RGB", im.size, (255, 255, 255))
                    rgba = im.convert("RGBA")
                    bg.paste(rgba, mask=rgba.split()[3])
                    im_rgb = bg
                elif im.mode != "RGB":
                    im_rgb = im.convert("RGB")
                else:
                    im_rgb = im

                buf = io.BytesIO()
                im_rgb.save(buf, format="PNG")
                prepared_buffers.append(buf.getvalue())

        pdf_bytes = img2pdf.convert(prepared_buffers)
        out.write_bytes(pdf_bytes)
        return out
    except Exception as exc:
        raise ConversionError(f"Failed to convert images to PDF: {exc}") from exc


def extract_pdf_images(
    input_path: str | Path,
    output_dir: str | Path | None = None,
) -> list[Path]:
    """Extract all embedded images from a PDF into a folder alongside the PDF."""
    src = ensure_input_file(input_path)
    out_folder = Path(output_dir).expanduser().resolve() if output_dir else (src.parent / f"{src.stem}_images")
    out_folder.mkdir(parents=True, exist_ok=True)

    try:
        import fitz

        doc = fitz.open(str(src))
        try:
            extracted: list[Path] = []
            seen_xrefs: set[int] = set()
            counter = 1
            for page_idx, page in enumerate(doc, start=1):
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    if xref in seen_xrefs:
                        continue
                    seen_xrefs.add(xref)
                    base_img = doc.extract_image(xref)
                    ext = base_img.get("ext", "png")
                    img_bytes = base_img.get("image")
                    if not img_bytes:
                        continue
                    out_file = out_folder / f"page_{page_idx:03d}_img_{counter:03d}.{ext}"
                    out_file.write_bytes(img_bytes)
                    extracted.append(out_file)
                    counter += 1
            return extracted
        finally:
            doc.close()
    except Exception as exc:
        raise ConversionError(f"Failed to extract images from PDF: {exc}") from exc


def extract_pdf_text(input_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Extract UTF-8 text from a PDF file."""
    from core.converter import pdf_to_txt

    return pdf_to_txt(input_path, output_path=output_path)
