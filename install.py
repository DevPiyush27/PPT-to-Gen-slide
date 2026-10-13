"""install.py - Automated Dependency Setup and System Environment Verifier.

Checks and installs required Python packages from requirements.txt,
and audits external system binaries (Poppler, Microsoft PowerPoint, LibreOffice, FFmpeg).
"""

import importlib
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

# Mapping package names in requirements.txt to their importable module names
REQUIRED_MODULES = {
    "python-pptx": "pptx",
    "pdf2image": "pdf2image",
    "comtypes": "comtypes",
    "python-dotenv": "dotenv",
    "requests": "requests",
    "edge-tts": "edge_tts",
    "moviepy": "moviepy",
}


def refresh_windows_path() -> None:
    """Reloads PATH from Windows User and System registry into current process environment."""
    if platform.system() != "Windows":
        return

    try:
        import winreg  # type: ignore

        paths = []
        for root, key_path in [
            (winreg.HKEY_CURRENT_USER, r"Environment"),
            (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
        ]:
            try:
                with winreg.OpenKey(root, key_path) as k:
                    val, _ = winreg.QueryValueEx(k, "Path")
                    paths.extend(val.split(";"))
            except OSError:
                continue

        # Also inspect winget local packages folder for poppler if present
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        if local_app_data:
            winget_pkgs = Path(local_app_data) / "Microsoft" / "WinGet" / "Packages"
            if winget_pkgs.exists():
                for poppler_bin in winget_pkgs.glob("*Poppler*/poppler-*/Library/bin"):
                    if poppler_bin.is_dir() and str(poppler_bin) not in paths:
                        paths.append(str(poppler_bin))

        current_paths = os.environ.get("PATH", "").split(";")
        combined = list(dict.fromkeys([p for p in paths + current_paths if p.strip()]))
        os.environ["PATH"] = ";".join(combined)
    except Exception:
        pass


def check_python_dependencies() -> bool:
    """Verifies all required Python modules can be imported.

    Returns True if all are present, False otherwise.
    """
    missing = []
    for pkg_name, module_name in REQUIRED_MODULES.items():
        try:
            mod = importlib.import_module(module_name)
            if pkg_name == "moviepy":
                version_str = getattr(mod, "__version__", "0.0.0")
                try:
                    major = int(version_str.split(".")[0])
                    if major < 2:
                        print(f"[-] moviepy is version {version_str}, but >= 2.0.0 is required.")
                        missing.append(pkg_name)
                except (ValueError, IndexError):
                    pass
        except ImportError:
            missing.append(pkg_name)

    return len(missing) == 0


def install_python_dependencies(req_file: Path) -> bool:
    """Installs packages from requirements.txt using pip."""
    if not req_file.exists():
        print(f"[!] Error: {req_file} not found.")
        return False

    print(f"[*] Installing dependencies from {req_file.name} via pip...")
    cmd = [sys.executable, "-m", "pip", "install", "-r", str(req_file)]
    try:
        result = subprocess.run(cmd, check=True)
        return result.returncode == 0
    except subprocess.CalledProcessError as e:
        print(f"[!] Failed to install dependencies: {e}")
        return False


def check_powerpoint_or_libreoffice() -> tuple[bool, str]:
    """Checks for Microsoft PowerPoint COM automation or LibreOffice headless in PATH.

    Returns (is_available, description_message).
    """
    # 1. Check Microsoft PowerPoint via COM (Windows only)
    if platform.system() == "Windows":
        try:
            import win32com.client  # type: ignore # noqa: F401
            # Or test comtypes
            import comtypes.client
            app = comtypes.client.CreateObject("PowerPoint.Application")
            app.Quit()
            return True, "Microsoft PowerPoint COM automation is available and verified."
        except Exception:
            # Fall back to checking powershell COM check
            try:
                res = subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     "try { $app = New-Object -ComObject PowerPoint.Application; $app.Quit(); exit 0 } catch { exit 1 }"],
                    capture_output=True,
                    timeout=8
                )
                if res.returncode == 0:
                    return True, "Microsoft PowerPoint COM automation is available."
            except Exception:
                pass

    # 2. Check LibreOffice
    soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice_path:
        return True, f"LibreOffice headless is available at: {soffice_path}"

    # Also check typical Windows install locations for LibreOffice
    default_lo_paths = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for p in default_lo_paths:
        if p.exists():
            return True, f"LibreOffice headless found at: {p}"

    return False, (
        "Neither Microsoft PowerPoint nor LibreOffice was detected!\n"
        "  -> Action required: Install Microsoft PowerPoint (standard on Windows Office/Microsoft 365)\n"
        "     OR install LibreOffice: winget install TheDocumentFoundation.LibreOffice"
    )


def check_poppler() -> tuple[bool, str]:
    """Checks whether pdftoppm (Poppler utility) is available in PATH.

    Returns (is_available, description_message).
    """
    refresh_windows_path()
    pdftoppm_path = shutil.which("pdftoppm")
    if pdftoppm_path:
        return True, f"Poppler (pdftoppm) is available at: {pdftoppm_path}"

    return False, (
        "Poppler binary 'pdftoppm' not found in PATH!\n"
        "  pdf2image requires Poppler to rasterize PDF slides to PNG images.\n"
        "  -> Action required on Windows:\n"
        "     1. Run: winget install oschwartz10612.Poppler\n"
        "     2. Or download from: https://github.com/oschwartz10612/poppler-windows/releases\n"
        "        and add the extracted 'Library/bin' folder to your System PATH."
    )


def check_ffmpeg() -> tuple[bool, str]:
    """Checks whether ffmpeg is available in PATH for video encoding."""
    refresh_windows_path()
    ffmpeg_path = shutil.which("ffmpeg")
    if ffmpeg_path:
        return True, f"FFmpeg is available at: {ffmpeg_path}"

    return False, (
        "FFmpeg not found in PATH!\n"
        "  MoviePy requires FFmpeg to encode and render MP4 presentations.\n"
        "  -> Action required: winget install Gyan.FFmpeg"
    )


def run_system_audit() -> dict[str, bool]:
    """Runs a complete audit of all external binaries and prints actionable diagnostics."""
    print("\n" + "=" * 60)
    print(" SYSTEM DEPENDENCIES AUDIT")
    print("=" * 60)

    results = {}

    # PowerPoint / LibreOffice
    ppt_ok, ppt_msg = check_powerpoint_or_libreoffice()
    results["office"] = ppt_ok
    if ppt_ok:
        print(f"[+] Presentation Engine: {ppt_msg}")
    else:
        print(f"[!] Presentation Engine Missing:\n    {ppt_msg}")

    # Poppler
    poppler_ok, poppler_msg = check_poppler()
    results["poppler"] = poppler_ok
    if poppler_ok:
        print(f"[+] PDF Rasterizer: {poppler_msg}")
    else:
        print(f"[!] PDF Rasterizer Missing:\n    {poppler_msg}")

    # FFmpeg
    ffmpeg_ok, ffmpeg_msg = check_ffmpeg()
    results["ffmpeg"] = ffmpeg_ok
    if ffmpeg_ok:
        print(f"[+] Video Codec (FFmpeg): {ffmpeg_msg}")
    else:
        print(f"[!] Video Codec Missing:\n    {ffmpeg_msg}")

    print("=" * 60 + "\n")
    return results


def setup_all() -> bool:
    """Orchestrates complete installation and verification."""
    print("=== PPT-to-Animation: Setup & Dependency Verifier ===")
    req_file = Path(__file__).parent / "requirements.txt"

    if not check_python_dependencies():
        print("[*] Missing or outdated Python packages detected. Installing...")
        success = install_python_dependencies(req_file)
        if not success:
            print("[!] Python package installation failed.")
            return False
    else:
        print("[+] All Python package dependencies are satisfied.")

    # System audit
    audit = run_system_audit()
    all_ok = all(audit.values())
    if all_ok:
        print("[+] System environment is fully prepared!")
    else:
        print("[*] Note: Pipeline will still attempt fallback converters where possible.")

    return True


if __name__ == "__main__":
    success = setup_all()
    sys.exit(0 if success else 1)
