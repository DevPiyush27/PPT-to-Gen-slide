"""config.py - Global Configuration and Environment Settings.

Handles:
1. Automatic creation of /assets/images/, /assets/audio/, and /output/ directories.
2. Loading .env variables via python-dotenv.
3. Deferred GEMINI_API_KEY fallback check with user-facing notification.
4. Auto-detection of Poppler path for reliable pdf2image execution on Windows.
"""

import os
import platform
import shutil
import sys
from pathlib import Path
from dotenv import load_dotenv

# Base project root
BASE_DIR = Path(__file__).resolve().parent

# Load .env file (with fallback to .env.example if user edited that directly)
ENV_PATH = BASE_DIR / ".env"
ENV_EXAMPLE_PATH = BASE_DIR / ".env.example"

if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
elif ENV_EXAMPLE_PATH.exists():
    load_dotenv(dotenv_path=ENV_EXAMPLE_PATH)

# Directory Paths
ASSETS_DIR = BASE_DIR / "assets"
IMAGES_DIR = ASSETS_DIR / "images"
AUDIO_DIR = ASSETS_DIR / "audio"
OUTPUT_DIR = BASE_DIR / "output"

# Auto-create all required directories on module initialization
for directory in [ASSETS_DIR, IMAGES_DIR, AUDIO_DIR, OUTPUT_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

# API Key handling with non-crashing deferred fallback check
RAW_API_KEY = (os.getenv("OPENROUTER_API_KEY") or os.getenv("GEMINI_API_KEY", "")).strip()
HAS_API_KEY = bool(
    RAW_API_KEY and RAW_API_KEY not in ["your_actual_api_key_here", "your_gemini_api_key_here"]
)

if not HAS_API_KEY:
    print("[!] Warning: OPENROUTER_API_KEY is not set. Add it to your .env file to enable AI narration generation.")
    OPENROUTER_API_KEY = ""
else:
    print("[+] Loaded OPENROUTER_API_KEY successfully from environment.")
    OPENROUTER_API_KEY = RAW_API_KEY

# Backward-compatibility alias
GEMINI_API_KEY = OPENROUTER_API_KEY

# Model configuration for OpenRouter
DEFAULT_MODEL = os.getenv("OPENROUTER_MODEL", "nvidia/nemotron-3.5-lightning:free")
FALLBACK_MODELS = [
    DEFAULT_MODEL,
    "nvidia/nemotron-3.5-lightning",
    "meta-llama/llama-3.3-70b-instruct:free",
]


# Default TTS Voice: Grounded, formal, academic delivery
DEFAULT_TTS_VOICE = os.getenv("TTS_VOICE", "en-US-ChristopherNeural")


def locate_poppler_bin() -> str | None:
    """Finds the directory containing pdftoppm for pdf2image.

    Checks PATH, then scans common Windows locations (winget, chocolatey, etc.).
    """
    # 1. Check if pdftoppm is already directly in PATH
    existing = shutil.which("pdftoppm")
    if existing:
        return str(Path(existing).parent)

    # 2. Check Windows registry / Winget package cache
    if platform.system() == "Windows":
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        if local_app_data:
            winget_pkgs = Path(local_app_data) / "Microsoft" / "WinGet" / "Packages"
            if winget_pkgs.exists():
                for candidate in winget_pkgs.glob("*Poppler*/poppler-*/Library/bin"):
                    if (candidate / "pdftoppm.exe").exists():
                        # Prepend to PATH for current process
                        os.environ["PATH"] = f"{candidate};" + os.environ.get("PATH", "")
                        return str(candidate)

        common_locations = [
            Path(r"C:\Program Files\Poppler\Library\bin"),
            Path(r"C:\Program Files (x86)\Poppler\Library\bin"),
            Path(r"C:\ProgramData\chocolatey\lib\poppler\tools"),
        ]
        for loc in common_locations:
            if (loc / "pdftoppm.exe").exists():
                os.environ["PATH"] = f"{loc};" + os.environ.get("PATH", "")
                return str(loc)

    return None


POPPLER_BIN_DIR = locate_poppler_bin()
