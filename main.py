"""main.py - End-to-End PPTX-to-Video Pipeline CLI Entrypoint.

Orchestrates the entire modular workflow:
1. Environment & system dependency verification
2. PPTX slide text, bullet points, and speaker notes extraction
3. Slide image rasterization (PowerPoint COM -> PDF -> PNG via pdf2image)
4. AI narration scripting (Gemini 2.5 Flash structured JSON / energetic mock)
5. Edge-TTS asynchronous narration audio synthesis
6. MoviePy 2.0 sequential video assembly and MP4 export
"""

import argparse
import asyncio
import os
import sys
from pathlib import Path

from config import (
    AUDIO_DIR,
    OPENROUTER_API_KEY,
    HAS_API_KEY,
    IMAGES_DIR,
    OUTPUT_DIR,
    DEFAULT_TTS_VOICE,
)
from extractor import convert_pptx_to_images, extract_slide_text
from ai_director import generate_script
from tts_engine import generate_narration_audio
from video_renderer import build_presentation_video
from install import check_python_dependencies, run_system_audit


def parse_arguments() -> argparse.Namespace:
    """Parses command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Convert a .pptx presentation into a fully narrated video presentation.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--input", "-i",
        type=str,
        required=False,
        default="",
        help="Path to the input PowerPoint (.pptx) file.",
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        required=False,
        default="",
        help="Target path for the output .mp4 video file (defaults to /output/<presentation_name>.mp4).",
    )
    parser.add_argument(
        "--mock", "-m",
        action="store_true",
        help="Force mock AI scripting for offline or dry-run testing without Gemini API key.",
    )
    parser.add_argument(
        "--voice", "-v",
        type=str,
        default=DEFAULT_TTS_VOICE,
        help="Edge-TTS voice name (e.g. en-US-GuyNeural, en-US-AriaNeural).",
    )
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Run system and package dependency checks only and exit.",
    )
    return parser.parse_args()


async def run_pipeline(
    pptx_path: Path,
    output_video_path: Path,
    use_mock: bool,
    voice: str,
) -> Path:
    """Runs the complete presentation-to-video pipeline."""
    print("\n" + "=" * 70)
    print(f"[*] Starting Pipeline for: {pptx_path.name}")
    print("=" * 70)

    # 1. Slide Text Extraction
    print(f"\n[Step 1/5] Extracting slide text, bullets, and speaker notes...")
    slides_data = extract_slide_text(pptx_path)
    total_slides = len(slides_data)
    print(f"[+] Extracted textual data for {total_slides} slides.")

    for s in slides_data:
        preview_bullets = len(s['bullet_points'])
        has_notes = bool(s['speaker_notes'])
        print(f"    - Slide {s['slide_number']}: {preview_bullets} text blocks, speaker notes: {'yes' if has_notes else 'none'}")

    # 2. Slide Image Rasterization
    print(f"\n[Step 2/5] Converting slides to PNG images (comtypes -> PDF -> pdf2image)...")
    images = convert_pptx_to_images(pptx_path, IMAGES_DIR)
    print(f"[+] Successfully generated {len(images)} slide PNGs in {IMAGES_DIR}")

    # 3. AI Narration Scripting
    print(f"\n[Step 3/5] Generating narration script...")
    api_key_to_use = "" if use_mock else OPENROUTER_API_KEY
    if not use_mock and not HAS_API_KEY:
        print("[!] No OPENROUTER_API_KEY detected in .env. Automatically enabling mock narration script.")
        use_mock = True

    script_data = await generate_script(slides_data, api_key=api_key_to_use)
    print(f"[+] Narration script prepared for {len(script_data)} slides.")
    for sc in script_data:
        preview = sc['narration_script'][:65] + "..." if len(sc['narration_script']) > 65 else sc['narration_script']
        print(f"    - Slide {sc['slide_number']}: \"{preview}\"")

    # 4. Asynchronous Edge-TTS Synthesis
    print(f"\n[Step 4/5] Synthesizing narration audio with Edge-TTS (voice: {voice})...")
    audio_files = await generate_narration_audio(script_data, AUDIO_DIR, voice=voice)
    print(f"[+] Generated {len(audio_files)} narration audio files in {AUDIO_DIR}")

    # 5. Video Assembly via MoviePy 2.0
    print(f"\n[Step 5/5] Assembling presentation video using MoviePy 2.0...")
    final_output = build_presentation_video(
        slide_image_paths=images,
        audio_paths=audio_files,
        output_path=output_video_path,
    )

    print("\n" + "=" * 70)
    print(f"[SUCCESS] Video Presentation Generated Successfully!")
    print(f"Output File: {final_output}")
    print(f"File Size  : {os.path.getsize(final_output) / (1024 * 1024):.2f} MB")
    print("=" * 70 + "\n")
    return Path(final_output)


def main() -> None:
    args = parse_arguments()

    # Pre-flight check: Verify python & system dependencies
    if not check_python_dependencies():
        print("[!] Missing Python packages detected. Run 'python install.py' first.")
        sys.exit(1)

    if args.check_only:
        audit = run_system_audit()
        sys.exit(0 if all(audit.values()) else 1)

    # Determine input file
    input_file = args.input.strip()
    if not input_file:
        print("\n[!] Error: No input file specified.")
        print("Usage: python main.py --input <path_to_presentation.pptx> [--mock]")
        print("Run 'python main.py --help' for all options.")
        sys.exit(1)

    pptx_path = Path(input_file).resolve()
    if not pptx_path.exists():
        print(f"[!] Error: File '{pptx_path}' does not exist.")
        sys.exit(1)

    # Determine output file path
    if args.output.strip():
        output_path = Path(args.output).resolve()
    else:
        output_path = OUTPUT_DIR / f"{pptx_path.stem}_presentation.mp4"

    asyncio.run(
        run_pipeline(
            pptx_path=pptx_path,
            output_video_path=output_path,
            use_mock=args.mock,
            voice=args.voice,
        )
    )


if __name__ == "__main__":
    main()
