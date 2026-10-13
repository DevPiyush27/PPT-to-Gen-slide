"""tts_engine.py - Asynchronous Edge-TTS Audio Narration Generator.

Generates grounded, formal, and authoritative spoken narration audio MP3s
for each presentation slide using Microsoft Edge TTS (en-US-ChristopherNeural).
"""

from pathlib import Path
from typing import Any
import edge_tts

from config import DEFAULT_TTS_VOICE


async def generate_narration_audio(
    script_data: list[dict[str, Any]],
    output_dir: str | Path,
    voice: str = DEFAULT_TTS_VOICE,
) -> list[str]:
    """Asynchronous function to synthesize formal audio narration for each slide script.

    Args:
        script_data: List of dicts with 'slide_number' and 'narration_script'.
        output_dir: Target directory where MP3 files will be stored.
        voice: Edge-TTS voice identifier (defaults to 'en-US-ChristopherNeural' via config).

    Returns:
        list[str]: Sequential paths to generated slide_{n}.mp3 files.
    """
    out_dir = Path(output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    # Use grounded corporate voice if default or empty
    selected_voice = voice if voice and voice.strip() else DEFAULT_TTS_VOICE

    audio_paths: list[str] = []

    for item in script_data:
        slide_num = item.get("slide_number", 1)
        narration_text = item.get("narration_script", "")

        # Validation: Verify text is not empty or pure whitespace to prevent silent audio
        if not narration_text or not narration_text.strip():
            raise ValueError(
                f"Narration text for slide {slide_num} is empty or blank! "
                "The TTS engine requires non-empty text to generate audible speech."
            )

        clean_text = narration_text.strip()
        audio_path = out_dir / f"slide_{slide_num}.mp3"
        print(f"[*] Synthesizing formal narration for Slide {slide_num} ({selected_voice})...")

        # Communicate object initialization with formal voice profile
        communicate = edge_tts.Communicate(text=clean_text, voice=selected_voice)

        # Mandate: Strictly await the asynchronous save operation
        await communicate.save(str(audio_path))

        # Verification: Confirm the audio file was written and is not empty
        if not audio_path.exists() or audio_path.stat().st_size == 0:
            raise RuntimeError(
                f"Audio generation failed for slide {slide_num}: "
                f"Output file {audio_path.name} is missing or 0 bytes."
            )

        print(f"[+] Saved audio ({audio_path.stat().st_size} bytes): {audio_path.name}")
        audio_paths.append(str(audio_path))

    return audio_paths
