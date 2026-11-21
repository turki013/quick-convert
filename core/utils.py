"""Shared utilities for Quick Convert."""

from __future__ import annotations

import logging
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger("quick_convert")


class ConversionError(Exception):
    """Raised when a file conversion or PDF operation fails."""


def ensure_input_file(path: str | Path) -> Path:
    """Validate that an input path exists and is a file."""
    resolved = Path(path).expanduser().resolve()
    if not resolved.exists():
        raise ConversionError(f"Input file not found: {resolved}")
    if not resolved.is_file():
        raise ConversionError(f"Input path is not a file: {resolved}")
    return resolved


def build_output_path(
    input_path: str | Path,
    extension: str,
    suffix: str = "",
    output_path: str | Path | None = None,
) -> Path:
    """
    Build a non-overwriting output path in the same directory as input_path
    unless output_path is explicitly provided.
    """
    if output_path is not None:
        out = Path(output_path).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        return out

    src = Path(input_path).expanduser().resolve()
    ext = extension if extension.startswith(".") else f".{extension}"
    stem = f"{src.stem}{suffix}"
    candidate = src.with_name(f"{stem}{ext}")

    # Avoid overwriting the source file or an existing file
    if candidate == src or candidate.exists():
        counter = 1
        while True:
            numbered = src.with_name(f"{stem}_{counter}{ext}")
            if numbered != src and not numbered.exists():
                return numbered
            counter += 1
    return candidate


def notify_windows(title: str, message: str, is_error: bool = False) -> None:
    """
    Show a non-blocking Windows toast notification.
    Uses winotify when available and falls back to PowerShell WinRT toast.
    """
    if sys.platform != "win32":
        return

    try:
        from winotify import Notification, audio  # type: ignore

        toast = Notification(
            app_id="Quick Convert",
            title=title[:120],
            msg=message[:240],
            duration="short",
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
        return
    except Exception as exc:
        logger.debug("winotify failed, falling back to PowerShell toast: %s", exc)

    # PowerShell non-blocking fallback
    safe_title = title.replace("'", "''")[:120]
    safe_msg = message.replace("'", "''")[:240]
    ps_script = (
        "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;"
        "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null;"
        f"$xml = New-Object Windows.Data.Xml.Dom.XmlDocument;"
        f"$xml.LoadXml('<toast><visual><binding template=\"ToastGeneric\"><text>{safe_title}</text><text>{safe_msg}</text></binding></visual></toast>');"
        "$toast = [Windows.UI.Notifications.ToastNotification]::new($xml);"
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Quick Convert').Show($toast);"
    )
    try:
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
        )
    except Exception as exc:
        logger.debug("PowerShell toast fallback failed: %s", exc)
