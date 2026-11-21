"""
Cloud-ready Telegram Bot (aiogram v3) powered by Quick Convert's core engine.

Supports:
- All core conversions (.docx <-> .md, .docx/.md/images -> .pdf, .pdf -> .docx/.md/.txt,
  PDF compress/split/extract-images, image format conversion/compression, .xlsx/.csv/.json)
- Multi-file batch sessions (/merge and /img2pdf)
- Optional ALLOWED_USER_IDS access control
- Hybrid Webhook (if WEBHOOK_URL is set) or Long-Polling + aiohttp /health server
- Automatic temporary file cleanup after each operation
"""

from __future__ import annotations

import asyncio
import logging
import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Any, Awaitable, Callable

from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
    TelegramObject,
)
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
from aiohttp import web

from core.converter import compress_image, convert_file
from core.pdf_tools import (
    compress_pdf,
    extract_pdf_images,
    images_to_pdf,
    merge_pdfs,
    split_pdf,
)
from core.utils import ConversionError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("quick_convert.bot")

router = Router()

WORK_DIR = Path(tempfile.gettempdir()) / "quick_convert_bot"
WORK_DIR.mkdir(parents=True, exist_ok=True)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}

# Per-user single-file store & batch session store
USER_FILES: dict[int, Path] = {}
BATCH_SESSIONS: dict[int, dict[str, Any]] = {}


def _parse_allowed_users() -> set[int]:
    raw = os.getenv("ALLOWED_USER_IDS", "").strip()
    if not raw:
        return set()
    ids: set[int] = set()
    for part in raw.split(","):
        token = part.strip()
        if token.lstrip("-").isdigit():
            ids.add(int(token))
    return ids


ALLOWED_USERS = _parse_allowed_users()


class AccessControlMiddleware(BaseMiddleware):
    """Restrict bot usage to ALLOWED_USER_IDS when configured."""

    async def __call__(
        self,
        handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict[str, Any],
    ) -> Any:
        if not ALLOWED_USERS:
            return await handler(event, data)

        user = getattr(event, "from_user", None)
        if user is not None and user.id not in ALLOWED_USERS:
            if isinstance(event, Message):
                await event.answer("⛔ Private bot: access restricted.")
            elif isinstance(event, CallbackQuery):
                await event.answer("⛔ Access restricted.", show_alert=True)
            return None
        return await handler(event, data)


def _sanitize_filename(name: str) -> str:
    """Keep Arabic/Unicode characters intact while stripping path traversal/unsafe chars."""
    base = Path(name).name
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base).strip()
    return cleaned or f"file_{uuid.uuid4().hex[:8]}"


def _cleanup_paths(*paths: Path | None) -> None:
    for p in paths:
        if p is None:
            continue
        try:
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            elif p.exists():
                p.unlink(missing_ok=True)
        except OSError:
            pass


def _actions_for_extension(ext: str) -> list[tuple[str, str]]:
    ext = ext.lower()
    if ext == ".docx":
        return [("To Markdown (.md)", "conv:md"), ("To PDF (.pdf)", "conv:pdf")]
    if ext == ".md":
        return [("To Word (.docx)", "conv:docx"), ("To PDF (.pdf)", "conv:pdf")]
    if ext == ".pdf":
        return [
            ("To Word (.docx)", "conv:docx"),
            ("To Markdown (.md)", "conv:md"),
            ("To Text (.txt)", "conv:txt"),
            ("Compress (Medium)", "pdf_comp:medium"),
            ("Compress (High)", "pdf_comp:high"),
            ("Burst Pages", "pdf_burst"),
            ("Extract Images", "pdf_ext_img"),
        ]
    if ext == ".xlsx":
        return [("To CSV (.csv)", "conv:csv"), ("To JSON (.json)", "conv:json")]
    if ext == ".csv":
        return [("To Excel (.xlsx)", "conv:xlsx"), ("To JSON (.json)", "conv:json")]
    if ext == ".json":
        return [("To Excel (.xlsx)", "conv:xlsx"), ("To CSV (.csv)", "conv:csv")]
    if ext in IMAGE_EXTS:
        return [
            ("Convert to PDF", "img2pdf"),
            ("Compress Image", "img_comp:medium"),
            ("To PNG", "conv:png"),
            ("To JPG", "conv:jpg"),
            ("To WebP", "conv:webp"),
        ]
    return []


def _build_keyboard(actions: list[tuple[str, str]]) -> InlineKeyboardMarkup:
    rows = [[InlineKeyboardButton(text=label, callback_data=cb)] for label, cb in actions]
    return InlineKeyboardMarkup(inline_keyboard=rows)


def _batch_keyboard(mode: str) -> InlineKeyboardMarkup:
    done_label = "✅ Done / Merge PDFs" if mode == "merge" else "✅ Done / Create PDF"
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=done_label, callback_data=f"batch:{mode}")],
            [InlineKeyboardButton(text="❌ Cancel", callback_data="batch:cancel")],
        ]
    )


@router.message(CommandStart())
async def cmd_start(message: Message) -> None:
    await message.answer(
        "👋 **Quick Convert Bot** (100% Local Engine)\n\n"
        "Send any file (`.docx`, `.md`, `.pdf`, `.xlsx`, `.csv`, `.json`, or image) for instant conversion.\n\n"
        "**Batch Commands:**\n"
        "• `/merge` — Send multiple PDFs and merge them into one\n"
        "• `/img2pdf` — Send multiple images and combine into one PDF\n"
        "• `/cancel` — Cancel active batch session",
        parse_mode="Markdown",
    )


@router.message(Command("merge"))
async def cmd_merge(message: Message) -> None:
    if message.from_user is None:
        return
    uid = message.from_user.id
    if uid in BATCH_SESSIONS:
        _cleanup_paths(BATCH_SESSIONS[uid].get("dir"))
    session_dir = WORK_DIR / f"batch_{uid}_{uuid.uuid4().hex[:8]}"
    session_dir.mkdir(parents=True, exist_ok=True)
    BATCH_SESSIONS[uid] = {"mode": "merge", "files": [], "dir": session_dir}
    await message.answer(
        "📄 **PDF Merge Mode started.**\nSend 2 or more PDF files, then press **Done**.",
        parse_mode="Markdown",
    )


@router.message(Command("img2pdf"))
async def cmd_img2pdf(message: Message) -> None:
    if message.from_user is None:
        return
    uid = message.from_user.id
    if uid in BATCH_SESSIONS:
        _cleanup_paths(BATCH_SESSIONS[uid].get("dir"))
    session_dir = WORK_DIR / f"batch_{uid}_{uuid.uuid4().hex[:8]}"
    session_dir.mkdir(parents=True, exist_ok=True)
    BATCH_SESSIONS[uid] = {"mode": "img2pdf", "files": [], "dir": session_dir}
    await message.answer(
        "🖼️ **Images-to-PDF Mode started.**\nSend 1 or more images/photos, then press **Done**.",
        parse_mode="Markdown",
    )


@router.message(Command("cancel"))
async def cmd_cancel(message: Message) -> None:
    if message.from_user is None:
        return
    session = BATCH_SESSIONS.pop(message.from_user.id, None)
    if session:
        _cleanup_paths(session.get("dir"))
        await message.answer("❌ Batch session cancelled.")
    else:
        await message.answer("No active batch session.")


@router.message(F.document)
async def handle_document(message: Message, bot: Bot) -> None:
    doc = message.document
    if doc is None or message.from_user is None:
        return

    uid = message.from_user.id
    file_name = _sanitize_filename(doc.file_name or "uploaded_file")
    ext = Path(file_name).suffix.lower()

    # Check if user is in a batch session
    session = BATCH_SESSIONS.get(uid)
    if session is not None:
        mode = session["mode"]
        if mode == "merge" and ext != ".pdf":
            await message.answer("⚠️ Merge mode expects `.pdf` files. Send a PDF or `/cancel`.")
            return
        if mode == "img2pdf" and ext not in IMAGE_EXTS:
            await message.answer("⚠️ Images-to-PDF mode expects image files (`.jpg`, `.png`, `.webp`).")
            return

        idx = len(session["files"]) + 1
        local_path = session["dir"] / f"{idx:03d}_{file_name}"
        tg_file = await bot.get_file(doc.file_id)
        if tg_file.file_path:
            await bot.download_file(tg_file.file_path, destination=local_path)
        session["files"].append(local_path)
        await message.answer(
            f"➕ Added `{file_name}` ({len(session['files'])} file(s) queued).",
            reply_markup=_batch_keyboard(mode),
            parse_mode="Markdown",
        )
        return

    actions = _actions_for_extension(ext)
    if not actions:
        await message.answer(f"Unsupported file type: `{ext}`", parse_mode="Markdown")
        return

    # Clean previous single file for this user
    _cleanup_paths(USER_FILES.pop(uid, None))

    user_dir = WORK_DIR / f"single_{uid}"
    user_dir.mkdir(parents=True, exist_ok=True)
    local_path = user_dir / file_name

    tg_file = await bot.get_file(doc.file_id)
    if tg_file.file_path:
        await bot.download_file(tg_file.file_path, destination=local_path)

    USER_FILES[uid] = local_path
    await message.answer(
        f"Received `{file_name}`. Choose an action:",
        reply_markup=_build_keyboard(actions),
        parse_mode="Markdown",
    )


@router.message(F.photo)
async def handle_photo(message: Message, bot: Bot) -> None:
    if not message.photo or message.from_user is None:
        return

    uid = message.from_user.id
    photo = message.photo[-1]
    session = BATCH_SESSIONS.get(uid)

    if session is not None and session["mode"] == "img2pdf":
        idx = len(session["files"]) + 1
        local_path = session["dir"] / f"{idx:03d}_photo_{photo.file_unique_id}.jpg"
        tg_file = await bot.get_file(photo.file_id)
        if tg_file.file_path:
            await bot.download_file(tg_file.file_path, destination=local_path)
        session["files"].append(local_path)
        await message.answer(
            f"➕ Added photo ({len(session['files'])} queued).",
            reply_markup=_batch_keyboard("img2pdf"),
        )
        return

    _cleanup_paths(USER_FILES.pop(uid, None))
    user_dir = WORK_DIR / f"single_{uid}"
    user_dir.mkdir(parents=True, exist_ok=True)
    local_path = user_dir / f"photo_{photo.file_unique_id}.jpg"

    tg_file = await bot.get_file(photo.file_id)
    if tg_file.file_path:
        await bot.download_file(tg_file.file_path, destination=local_path)

    USER_FILES[uid] = local_path
    await message.answer(
        "Received image. Choose an action:",
        reply_markup=_build_keyboard(_actions_for_extension(".jpg")),
    )


@router.callback_query(F.data.startswith("batch:"))
async def handle_batch_callback(callback: CallbackQuery) -> None:
    if callback.from_user is None or callback.data is None or callback.message is None:
        return

    uid = callback.from_user.id
    action = callback.data.split(":", 1)[1]
    session = BATCH_SESSIONS.pop(uid, None)

    if action == "cancel" or session is None:
        if session:
            _cleanup_paths(session.get("dir"))
        await callback.answer("Cancelled.")
        await callback.message.answer("❌ Batch session cancelled.")
        return

    files: list[Path] = session["files"]
    session_dir: Path = session["dir"]

    try:
        if action == "merge":
            if len(files) < 2:
                BATCH_SESSIONS[uid] = session
                await callback.answer("Send at least 2 PDFs first.", show_alert=True)
                return
            await callback.answer("Merging PDFs...")
            out_path = await asyncio.to_thread(merge_pdfs, files, session_dir / "merged.pdf", False)
            await callback.message.answer_document(FSInputFile(out_path))
        elif action == "img2pdf":
            if not files:
                BATCH_SESSIONS[uid] = session
                await callback.answer("Send at least 1 image first.", show_alert=True)
                return
            await callback.answer("Creating PDF...")
            out_path = await asyncio.to_thread(images_to_pdf, files, session_dir / "images.pdf", False)
            await callback.message.answer_document(FSInputFile(out_path))
    except ConversionError as exc:
        await callback.message.answer(f"Operation failed: {exc}")
    finally:
        if uid not in BATCH_SESSIONS:
            _cleanup_paths(session_dir)


@router.callback_query()
async def handle_action(callback: CallbackQuery) -> None:
    if callback.from_user is None or callback.data is None or callback.message is None:
        return

    uid = callback.from_user.id
    src = USER_FILES.get(uid)
    if src is None or not src.exists():
        await callback.answer("Please send a file first.", show_alert=True)
        return

    await callback.answer("Processing...")
    action = callback.data
    generated_outputs: list[Path] = []
    extra_dir: Path | None = None

    try:
        if action.startswith("conv:"):
            target_fmt = action.split(":", 1)[1]
            out_path = await asyncio.to_thread(convert_file, src, target_fmt)
            generated_outputs.append(out_path)
            await callback.message.answer_document(FSInputFile(out_path))
        elif action.startswith("pdf_comp:"):
            preset = action.split(":", 1)[1]
            out_path = await asyncio.to_thread(compress_pdf, src, preset)  # type: ignore[arg-type]
            generated_outputs.append(out_path)
            await callback.message.answer_document(FSInputFile(out_path))
        elif action == "pdf_burst":
            outs = await asyncio.to_thread(split_pdf, src, None, True)
            generated_outputs.extend(outs)
            for out_file in outs[:10]:
                await callback.message.answer_document(FSInputFile(out_file))
            if len(outs) > 10:
                await callback.message.answer(f"Sent first 10 of {len(outs)} pages.")
        elif action == "pdf_ext_img":
            extra_dir = src.parent / f"{src.stem}_images_{uuid.uuid4().hex[:6]}"
            outs = await asyncio.to_thread(extract_pdf_images, src, extra_dir)
            if not outs:
                await callback.message.answer("No embedded images found in this PDF.")
            else:
                for img_file in outs[:10]:
                    await callback.message.answer_document(FSInputFile(img_file))
                if len(outs) > 10:
                    await callback.message.answer(f"Sent first 10 of {len(outs)} images.")
        elif action == "img2pdf":
            out_path = await asyncio.to_thread(images_to_pdf, [src])
            generated_outputs.append(out_path)
            await callback.message.answer_document(FSInputFile(out_path))
        elif action.startswith("img_comp:"):
            preset = action.split(":", 1)[1]
            out_path = await asyncio.to_thread(compress_image, src, preset)
            generated_outputs.append(out_path)
            await callback.message.answer_document(FSInputFile(out_path))
    except ConversionError as exc:
        await callback.message.answer(f"Conversion failed: {exc}")
    except Exception as exc:
        logger.exception("Unexpected error in bot callback")
        await callback.message.answer(f"Unexpected error: {exc}")
    finally:
        USER_FILES.pop(uid, None)
        _cleanup_paths(src, extra_dir, *generated_outputs)


async def _health_handler(_request: web.Request) -> web.Response:
    return web.json_response({"status": "ok", "service": "quick-convert-bot"})


async def main() -> None:
    token = os.getenv("QUICK_CONVERT_BOT_TOKEN") or os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError(
            "Missing Telegram Bot Token. Set BOT_TOKEN or QUICK_CONVERT_BOT_TOKEN environment variable."
        )

    bot = Bot(token=token)
    dp = Dispatcher()
    dp.message.middleware(AccessControlMiddleware())
    dp.callback_query.middleware(AccessControlMiddleware())
    dp.include_router(router)

    port = int(os.getenv("PORT", "8000"))
    webhook_url = os.getenv("WEBHOOK_URL", "").strip()

    app = web.Application()
    app.router.add_get("/", _health_handler)
    app.router.add_get("/health", _health_handler)

    if webhook_url:
        webhook_path = "/webhook"
        full_webhook_url = f"{webhook_url.rstrip('/')}{webhook_path}"
        SimpleRequestHandler(dispatcher=dp, bot=bot).register(app, path=webhook_path)
        setup_application(app, dp, bot=bot)
        await bot.set_webhook(full_webhook_url)
        logger.info("Starting Quick Convert Bot in WEBHOOK mode on port %s (%s)", port, full_webhook_url)

        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, host="0.0.0.0", port=port)
        await site.start()
        try:
            await asyncio.Event().wait()
        finally:
            await runner.cleanup()
            await bot.session.close()
    else:
        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, host="0.0.0.0", port=port)
        await site.start()
        logger.info("Health-check server listening on 0.0.0.0:%s (/health)", port)
        try:
            await bot.delete_webhook(drop_pending_updates=False)
            logger.info("Starting Quick Convert Bot in POLLING mode.")
            await dp.start_polling(bot)
        finally:
            await runner.cleanup()
            await bot.session.close()


if __name__ == "__main__":
    asyncio.run(main())
