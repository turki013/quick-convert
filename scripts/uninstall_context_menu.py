"""
Unregister 'Quick Convert' from both the Windows 11 Modern Context Menu (Sparse Package)
and the classic HKCU Registry fallback.
"""

from __future__ import annotations

import subprocess
import sys
import winreg

KEYS_TO_REMOVE = [
    r"Software\Classes\*\shell\QuickConvert",
    r"Software\Classes\Directory\Background\shell\QuickConvert",
    r"Software\Classes\CLSID\{8F3C7E21-4B9A-4D8E-9C12-6A7B8C9D0E1F}",
    r"Software\Classes\CLSID\{9E4D8F32-5C0B-4E9F-AD23-7B8C9D0E1F2A}",
]


def _delete_tree(root: int, subkey: str) -> bool:
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                    _delete_tree(root, rf"{subkey}\{child}")
                except OSError:
                    break
        winreg.DeleteKey(root, subkey)
        return True
    except FileNotFoundError:
        return False


def uninstall() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Windows Registry context menu removal is only supported on Windows.")

    # Remove Windows 11 Sparse Package if registered
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "Get-AppxPackage -Name 'QuickConvert.ShellExtension' | Remove-AppxPackage",
        ],
        capture_output=True,
    )

    removed = 0
    for key_path in KEYS_TO_REMOVE:
        if _delete_tree(winreg.HKEY_CURRENT_USER, key_path):
            removed += 1
            print(f"Removed: HKCU\\{key_path}")
        else:
            print(f"Not found (already clean): HKCU\\{key_path}")

    print(f"Done. Removed {removed} registry key(s) and checked Sparse Package removal.")


if __name__ == "__main__":
    uninstall()
