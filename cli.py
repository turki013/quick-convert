"""Fast, silent CLI entrypoint for Quick Convert (terminal & Windows Context Menu)."""

from __future__ import annotations

import argparse
import hashlib
import os
import sys
import tempfile
import time
from pathlib import Path
from typing import Callable

from core.converter import compress_image, convert_file
from core.pdf_tools import (
    compress_pdf,
    extract_pdf_images,
    extract_pdf_text,
    images_to_pdf,
    merge_pdfs,
    split_pdf,
)
from core.utils import ConversionError, notify_windows

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}


def _acquire_batch_paths(action: str, file_path: Path, wait_seconds: float = 0.75) -> list[Path] | None:
    """
    Coordinate multiple parallel Windows Explorer context-menu invocations.

    When a user selects N files and clicks a context menu item, Explorer launches N
    separate processes almost simultaneously. Each process appends its file path to
    a shared batch file. Only the leader process (which created the batch) waits
    briefly for siblings, reads the full list, cleans up, and returns the paths.
    Non-leader processes return None and exit silently.
    """
    batch_root = Path(tempfile.gettempdir()) / "quick_convert_batch"
    batch_root.mkdir(parents=True, exist_ok=True)

    parent_hash = hashlib.sha1(str(file_path.parent).lower().encode("utf-8")).hexdigest()[:12]
    batch_file = batch_root / f"{action}_{parent_hash}.txt"
    lock_file = batch_root / f"{action}_{parent_hash}.lock"

    # Acquire simple spin-lock using atomic O_CREAT | O_EXCL
    deadline = time.monotonic() + 5.0
    fd = None
    while time.monotonic() < deadline:
        try:
            fd = os.open(str(lock_file), os.O_CREAT | os.O_EXCL | os.O_RDWR)
            break
        except FileExistsError:
            # Break stale lock older than 10 seconds
            try:
                if time.time() - lock_file.stat().st_mtime > 10.0:
                    lock_file.unlink(missing_ok=True)
            except OSError:
                pass
            time.sleep(0.02)

    is_leader = False
    try:
        is_leader = not batch_file.exists() or (time.time() - batch_file.stat().st_mtime > 5.0)
        mode = "w" if is_leader else "a"
        with batch_file.open(mode, encoding="utf-8") as f:
            f.write(str(file_path.resolve()) + "\n")
    finally:
        if fd is not None:
            os.close(fd)
        lock_file.unlink(missing_ok=True)

    if not is_leader:
        return None

    # Leader waits for sibling Explorer processes to append their paths
    last_mtime = batch_file.stat().st_mtime
    start = time.monotonic()
    while time.monotonic() - start < wait_seconds:
        time.sleep(0.15)
        try:
            cur_mtime = batch_file.stat().st_mtime
            if cur_mtime > last_mtime:
                last_mtime = cur_mtime
                # Extend slightly if files are still arriving, capped at 2.0s total
                if time.monotonic() - start > wait_seconds - 0.2 and (time.monotonic() - start) < 2.0:
                    wait_seconds += 0.25
        except OSError:
            break

    # Collect & deduplicate preserving order
    raw_lines = batch_file.read_text(encoding="utf-8").splitlines()
    batch_file.unlink(missing_ok=True)

    unique_paths: list[Path] = []
    seen: set[str] = set()
    for line in raw_lines:
        clean = line.strip()
        if clean and clean not in seen:
            seen.add(clean)
            unique_paths.append(Path(clean))
    return unique_paths


def _run_with_notify(notify: bool, title: str, fn: Callable[[], str]) -> int:
    try:
        msg = fn()
        if notify:
            notify_windows(title, msg, is_error=False)
        else:
            print(msg)
        return 0
    except ConversionError as exc:
        if notify:
            notify_windows(f"{title} - Error", str(exc), is_error=True)
        else:
            print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        if notify:
            notify_windows(f"{title} - Unexpected Error", str(exc), is_error=True)
        else:
            print(f"Unexpected error: {exc}", file=sys.stderr)
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="quick-convert",
        description="Fast, 100% local file converter and PDF utility.",
    )
    parser.add_argument(
        "--notify",
        action="store_true",
        help="Show Windows toast notification on completion or error.",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Aggregate multi-select Windows Explorer invocations into a single operation.",
    )

    sub = parser.add_subparsers(dest="command", required=True)

    # convert
    p_conv = sub.add_parser("convert", help="Convert a file to a target format.")
    p_conv.add_argument("input", help="Source file path")
    p_conv.add_argument(
        "--to",
        required=True,
        help="Target format (md, docx, pdf, txt, csv, xlsx, json, png, jpg, webp)",
    )
    p_conv.add_argument("-o", "--output", default=None, help="Optional explicit output path")

    # compress-image
    p_cimg = sub.add_parser("compress-image", help="Compress an image file.")
    p_cimg.add_argument("input", help="Source image path")
    p_cimg.add_argument("--quality", choices=["low", "medium", "high"], default="medium")
    p_cimg.add_argument("-o", "--output", default=None)

    # pdf-compress
    p_comp = sub.add_parser("pdf-compress", help="Compress a PDF file.")
    p_comp.add_argument("input", help="Source PDF path")
    p_comp.add_argument("--quality", choices=["low", "medium", "high"], default="medium")
    p_comp.add_argument("-o", "--output", default=None)

    # pdf-merge
    p_merge = sub.add_parser("pdf-merge", help="Merge multiple PDF files into one.")
    p_merge.add_argument("inputs", nargs="+", help="PDF files to merge")
    p_merge.add_argument("-o", "--output", default=None)

    # pdf-split
    p_split = sub.add_parser("pdf-split", help="Split a PDF by page ranges or burst all pages.")
    p_split.add_argument("input", help="Source PDF path")
    p_split.add_argument("--pages", default=None, help="Page range e.g. '1-3, 5'")
    p_split.add_argument("--burst", action="store_true", help="Burst into single-page PDFs")
    p_split.add_argument("-o", "--output-dir", default=None)

    # pdf-extract-images
    p_ext_img = sub.add_parser("pdf-extract-images", help="Extract embedded images from a PDF.")
    p_ext_img.add_argument("input", help="Source PDF path")
    p_ext_img.add_argument("-o", "--output-dir", default=None)

    # pdf-extract-text
    p_ext_txt = sub.add_parser("pdf-extract-text", help="Extract plain text from a PDF.")
    p_ext_txt.add_argument("input", help="Source PDF path")
    p_ext_txt.add_argument("-o", "--output", default=None)

    # images-to-pdf
    p_i2p = sub.add_parser("images-to-pdf", help="Convert image(s) to a single PDF.")
    p_i2p.add_argument("inputs", nargs="+", help="Image files")
    p_i2p.add_argument("-o", "--output", default=None)

    # folder-merge-pdfs
    p_f_merge = sub.add_parser("folder-merge-pdfs", help="Merge all PDFs inside a folder.")
    p_f_merge.add_argument("folder", help="Folder path")

    # folder-images-to-pdf
    p_f_img = sub.add_parser("folder-images-to-pdf", help="Convert all images in a folder to PDF.")
    p_f_img.add_argument("folder", help="Folder path")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    notify: bool = args.notify

    if args.command == "convert":
        return _run_with_notify(
            notify,
            "Quick Convert",
            lambda: f"Saved: {convert_file(args.input, args.to, output_path=args.output).name}",
        )

    if args.command == "compress-image":
        return _run_with_notify(
            notify,
            "Compress Image",
            lambda: f"Saved: {compress_image(args.input, quality=args.quality, output_path=args.output).name}",
        )

    if args.command == "pdf-compress":
        return _run_with_notify(
            notify,
            "PDF Compress",
            lambda: f"Saved: {compress_pdf(args.input, quality=args.quality, output_path=args.output).name}",
        )

    if args.command == "pdf-merge":
        if args.batch and len(args.inputs) == 1:
            collected = _acquire_batch_paths("pdf_merge", Path(args.inputs[0]))
            if collected is None:
                return 0
            inputs = collected
        else:
            inputs = [Path(p) for p in args.inputs]

        return _run_with_notify(
            notify,
            "Merge PDFs",
            lambda: f"Merged {len(inputs)} PDFs -> {merge_pdfs(inputs, output_path=args.output).name}",
        )

    if args.command == "pdf-split":
        def _do_split() -> str:
            outs = split_pdf(args.input, pages=args.pages, burst=args.burst or not args.pages, output_dir=args.output_dir)
            return f"Created {len(outs)} PDF file(s) in {outs[0].parent.name}"

        return _run_with_notify(notify, "Split PDF", _do_split)

    if args.command == "pdf-extract-images":
        def _do_ext_img() -> str:
            outs = extract_pdf_images(args.input, output_dir=args.output_dir)
            if not outs:
                return "No embedded images found in PDF."
            return f"Extracted {len(outs)} image(s) to {outs[0].parent.name}"

        return _run_with_notify(notify, "Extract PDF Images", _do_ext_img)

    if args.command == "pdf-extract-text":
        return _run_with_notify(
            notify,
            "Extract PDF Text",
            lambda: f"Saved: {extract_pdf_text(args.input, output_path=args.output).name}",
        )

    if args.command == "images-to-pdf":
        if args.batch and len(args.inputs) == 1:
            collected = _acquire_batch_paths("images_to_pdf", Path(args.inputs[0]))
            if collected is None:
                return 0
            inputs = collected
        else:
            inputs = [Path(p) for p in args.inputs]

        return _run_with_notify(
            notify,
            "Images to PDF",
            lambda: f"Converted {len(inputs)} image(s) -> {images_to_pdf(inputs, output_path=args.output).name}",
        )

    if args.command == "folder-merge-pdfs":
        folder = Path(args.folder).expanduser().resolve()
        pdfs = sorted(
            [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf" and not p.stem.endswith("_merged")],
            key=lambda p: p.name.lower(),
        )
        return _run_with_notify(
            notify,
            "Merge Folder PDFs",
            lambda: f"Merged {len(pdfs)} PDFs -> {merge_pdfs(pdfs).name}",
        )

    if args.command == "folder-images-to-pdf":
        folder = Path(args.folder).expanduser().resolve()
        imgs = sorted(
            [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS],
            key=lambda p: p.name.lower(),
        )
        return _run_with_notify(
            notify,
            "Folder Images to PDF",
            lambda: f"Converted {len(imgs)} image(s) -> {images_to_pdf(imgs).name}",
        )

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
