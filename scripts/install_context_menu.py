"""
Register 'Quick Convert' in Windows 11/10 Explorer:
1. Modern Windows 11 Context Menu (primary right-click menu, no Shift+F10)
   via a Sparse Package (AppxManifest.xml) + compiled IExplorerCommand COM DLL.
2. Classic HKCU Registry cascading submenu as a fallback.

Writes only to HKEY_CURRENT_USER (no administrator privileges required).
"""

from __future__ import annotations

import subprocess
import sys
import winreg
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CLI_PATH = PROJECT_ROOT / "cli.py"
SHELL_DIR = PROJECT_ROOT / "shell_extension"
SHELL_CS = SHELL_DIR / "QuickConvertShell.cs"
SHELL_DLL = SHELL_DIR / "QuickConvertShell.dll"
MANIFEST_PATH = SHELL_DIR / "AppxManifest.xml"

FILE_MENU_KEY = r"Software\Classes\*\shell\QuickConvert"
BG_MENU_KEY = r"Software\Classes\Directory\Background\shell\QuickConvert"

COM_CLASSES = [
    ("8F3C7E21-4B9A-4D8E-9C12-6A7B8C9D0E1F", "QuickConvertShell.QuickConvertRootCommand"),
    ("9E4D8F32-5C0B-4E9F-AD23-7B8C9D0E1F2A", "QuickConvertShell.QuickConvertBgCommand"),
]


def _resolve_pythonw() -> Path:
    venv_pythonw = PROJECT_ROOT / ".venv" / "Scripts" / "pythonw.exe"
    if venv_pythonw.exists():
        return venv_pythonw
    exe_dir = Path(sys.executable).resolve().parent
    candidate = exe_dir / "pythonw.exe"
    if candidate.exists():
        return candidate
    return Path(sys.executable).resolve()


def _set_reg_value(key_path: str, name: str, value: str) -> None:
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE) as key:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)


def _add_submenu_item(
    parent_shell_path: str,
    item_id: str,
    label: str,
    command_args: str,
    pythonw: Path,
    applies_to: str | None = None,
) -> None:
    item_key = rf"{parent_shell_path}\shell\{item_id}"
    _set_reg_value(item_key, "MUIVerb", label)
    if applies_to:
        _set_reg_value(item_key, "AppliesTo", applies_to)

    cmd = f'"{pythonw}" "{CLI_PATH}" --notify {command_args}'
    _set_reg_value(rf"{item_key}\command", "", cmd)


def _build_shell_dll() -> Path:
    """Compile QuickConvertShell.cs into QuickConvertShell.dll using built-in .NET csc.exe."""
    if SHELL_DLL.exists() and SHELL_DLL.stat().st_mtime >= SHELL_CS.stat().st_mtime:
        return SHELL_DLL

    csc_candidates = sorted(
        Path(r"C:\Windows\Microsoft.NET\Framework64").glob("v4*/csc.exe"),
        reverse=True,
    )
    if not csc_candidates:
        raise RuntimeError("Could not find .NET Framework x64 csc.exe to compile QuickConvertShell.dll")

    csc = csc_candidates[0]
    subprocess.run(
        [
            str(csc),
            "/nologo",
            "/target:library",
            "/platform:x64",
            "/optimize+",
            f"/out:{SHELL_DLL}",
            str(SHELL_CS),
        ],
        check=True,
    )
    return SHELL_DLL


def _register_hkcu_com(dll_path: Path) -> None:
    """Register .NET COM-visible IExplorerCommand classes under HKCU (no admin required)."""
    codebase = dll_path.resolve().as_uri()
    assembly_name = "QuickConvertShell, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"

    for clsid, class_name in COM_CLASSES:
        clsid_key = rf"Software\Classes\CLSID\{{{clsid}}}"
        inproc_key = rf"{clsid_key}\InprocServer32"
        _set_reg_value(clsid_key, "", class_name)
        _set_reg_value(inproc_key, "", r"C:\Windows\System32\mscoree.dll")
        _set_reg_value(inproc_key, "ThreadingModel", "Both")
        _set_reg_value(inproc_key, "Assembly", assembly_name)
        _set_reg_value(inproc_key, "Class", class_name)
        _set_reg_value(inproc_key, "RuntimeVersion", "v4.0.30319")
        _set_reg_value(inproc_key, "CodeBase", codebase)


def _is_developer_mode_enabled() -> bool:
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock",
        ) as key:
            val, _ = winreg.QueryValueEx(key, "AllowDevelopmentWithoutDevLicense")
            return int(val) == 1
    except OSError:
        return False


def _register_sparse_package() -> bool:
    """Register the Windows 11 Sparse Package via Add-AppxPackage -ExternalLocation."""
    ps_cmd = (
        f'Add-AppxPackage -Register "{MANIFEST_PATH}" -ExternalLocation "{PROJECT_ROOT}"'
    )
    if not _is_developer_mode_enabled():
        print("\n[!] Windows Developer Mode is currently OFF.")
        print("    To show 'Quick Convert' in the primary Windows 11 right-click menu (without Shift+F10):")
        print("    1. Enable Developer Mode: Settings -> System -> For developers -> Developer Mode (ON)")
        print("    2. Run this one-line PowerShell command:")
        print(f"       {ps_cmd}")
        return False

    proc = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_cmd],
        capture_output=True,
        text=True,
    )
    if proc.returncode == 0:
        print("Windows 11 Modern Context Menu (Sparse Package) registered successfully!")
        return True

    print(f"\n[!] Sparse Package registration failed: {proc.stderr.strip()}")
    print(f"    Manual command: {ps_cmd}")
    return False


def _install_classic_fallback(pythonw: Path) -> None:
    """Register classic HKCU cascading menu as fallback (Shift+F10 / Windows 10)."""
    _set_reg_value(FILE_MENU_KEY, "MUIVerb", "Quick Convert")
    _set_reg_value(FILE_MENU_KEY, "SubCommands", "")
    _set_reg_value(FILE_MENU_KEY, "Icon", "imageres.dll,-109")

    file_items = [
        ("01_to_pdf", "Convert to PDF (.pdf)", 'convert "%1" --to pdf', ".docx OR .md"),
        ("02_to_docx", "Convert to Word (.docx)", 'convert "%1" --to docx', ".pdf OR .md"),
        ("03_to_md", "Convert to Markdown (.md)", 'convert "%1" --to md', ".docx OR .pdf"),
        ("04_to_txt", "Extract Text (.txt)", 'convert "%1" --to txt', ".pdf"),
        ("05_pdf_merge", "Merge Selected PDFs", '--batch pdf-merge "%1"', ".pdf"),
        ("06_pdf_comp_med", "Compress PDF (Medium)", 'pdf-compress "%1" --quality medium', ".pdf"),
        ("07_pdf_comp_high", "Compress PDF (Max Compression)", 'pdf-compress "%1" --quality high', ".pdf"),
        ("08_pdf_burst", "Split PDF (Burst Pages)", 'pdf-split "%1" --burst', ".pdf"),
        ("09_pdf_ext_img", "Extract PDF Images", 'pdf-extract-images "%1"', ".pdf"),
        ("10_img_to_pdf", "Selected Images to PDF", '--batch images-to-pdf "%1"', ".jpg OR .jpeg OR .png OR .webp OR .bmp"),
        ("11_img_comp", "Compress Image", 'compress-image "%1" --quality medium', ".jpg OR .jpeg OR .png OR .webp"),
        ("12_to_png", "Convert to PNG (.png)", 'convert "%1" --to png', ".jpg OR .jpeg OR .webp OR .bmp"),
        ("13_to_jpg", "Convert to JPG (.jpg)", 'convert "%1" --to jpg', ".png OR .webp OR .bmp"),
        ("14_to_webp", "Convert to WebP (.webp)", 'convert "%1" --to webp', ".png OR .jpg OR .jpeg OR .bmp"),
        ("15_to_csv", "Convert to CSV (.csv)", 'convert "%1" --to csv', ".xlsx OR .json"),
        ("16_to_xlsx", "Convert to Excel (.xlsx)", 'convert "%1" --to xlsx', ".csv OR .json"),
        ("17_to_json", "Convert to JSON (.json)", 'convert "%1" --to json', ".xlsx OR .csv"),
    ]

    for item_id, label, args, applies_to in file_items:
        _add_submenu_item(FILE_MENU_KEY, item_id, label, args, pythonw, applies_to=applies_to)

    _set_reg_value(BG_MENU_KEY, "MUIVerb", "Quick Convert")
    _set_reg_value(BG_MENU_KEY, "SubCommands", "")
    _set_reg_value(BG_MENU_KEY, "Icon", "imageres.dll,-109")

    _add_submenu_item(
        BG_MENU_KEY,
        "01_folder_merge_pdfs",
        "Merge All PDFs in Folder",
        'folder-merge-pdfs "%V"',
        pythonw,
    )
    _add_submenu_item(
        BG_MENU_KEY,
        "02_folder_images_to_pdf",
        "All Images in Folder to PDF",
        'folder-images-to-pdf "%V"',
        pythonw,
    )


def install() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Windows Registry context menu installation is only supported on Windows.")

    pythonw = _resolve_pythonw()
    dll_path = _build_shell_dll()
    _register_hkcu_com(dll_path)
    _install_classic_fallback(pythonw)
    print(f"Classic HKCU fallback menu installed (interpreter: {pythonw}).")
    _register_sparse_package()


if __name__ == "__main__":
    install()
