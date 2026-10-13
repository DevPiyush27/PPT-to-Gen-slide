"""video_renderer.py - MoviePy 2.0 Sequential Video Assembler.

Assembles sequential slide video presentation using modern MoviePy 2.0 APIs:
- Explicit imports (no wildcards)
- Immutable chaining: .with_duration() and .with_audio()
- Smooth transitions with .with_effects([CrossFadeIn(duration)])
- Sequential assembly with concatenate_videoclips(clips, method="compose")
- Clean rendering to MP4 via libx264 and aac
"""

from pathlib import Path
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
from moviepy.video.fx import CrossFadeIn


def build_presentation_video(
    slide_image_paths: list[str],
    audio_paths: list[str],
    output_path: str | Path,
    fps: int = 24,
    transition_duration: float = 0.5,
) -> str:
    """Builds an animated, narrated presentation video from slide images and audio files.

    Args:
        slide_image_paths: List of file paths to slide PNG images.
        audio_paths: List of file paths to corresponding slide narration MP3s.
        output_path: Target output MP4 file path.
        fps: Frames per second (default 24).
        transition_duration: Duration in seconds for slide transition effect.

    Returns:
        str: Absolute path to the rendered output MP4 video.
    """
    out_file = Path(output_path).resolve()
    # Container format enforcement: strictly require .mp4 to prevent audio mute in player containers
    if out_file.suffix.lower() != ".mp4":
        out_file = out_file.with_suffix(".mp4")
    out_file.parent.mkdir(parents=True, exist_ok=True)

    if len(slide_image_paths) != len(audio_paths):
        raise ValueError(
            f"Mismatched counts: {len(slide_image_paths)} images vs {len(audio_paths)} audio files."
        )

    if not slide_image_paths:
        raise ValueError("Cannot build video: No slide images provided.")

    clips = []
    audio_clips = []

    try:
        for idx, (img_path, aud_path) in enumerate(zip(slide_image_paths, audio_paths), start=1):
            print(f"[*] Processing slide {idx}: {Path(img_path).name} with {Path(aud_path).name}...")

            # 1. Load audio clip
            audio = AudioFileClip(aud_path)
            audio_clips.append(audio)

            # 2. MoviePy 2.0 Immutable Clip Mandate:
            # .with_duration() and .with_audio() return new clips and do NOT mutate in place.
            # We MUST reassign the variable: clip = clip.with_duration(...).with_audio(...)
            clip = ImageClip(img_path)
            clip = clip.with_duration(audio.duration).with_audio(audio)

            # 3. Apply smooth entrance animation using .with_effects()
            if idx > 1 and audio.duration > transition_duration:
                # Apply crossfade transition for slides following the title slide
                clip = clip.with_effects([CrossFadeIn(transition_duration)])

            clips.append(clip)

        print(f"[*] Concatenating {len(clips)} slide clips sequentially...")
        # Sequential assembly with compose method
        final_clip = concatenate_videoclips(clips, method="compose")

        print(f"[*] Rendering final video to {out_file.name} (codec: libx264, audio: aac)...")
        # Mandate: export code must use .mp4 extension, codec="libx264", audio_codec="aac"
        final_clip.write_videofile(
            str(out_file),
            fps=fps,
            codec="libx264",
            audio_codec="aac",
            preset="medium",
            logger="bar",
        )

        print(f"[+] Presentation video created successfully: {out_file}")
        return str(out_file)

    finally:
        # Explicit cleanup to release file locks on Windows
        for c in clips:
            try:
                c.close()
            except Exception:
                pass
        for a in audio_clips:
            try:
                a.close()
            except Exception:
                pass
