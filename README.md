# PPT-to-Animation: Automated Presentation Video Narrator

A modular Python pipeline that converts any Microsoft PowerPoint (`.pptx`) presentation into an energetic, narrated video presentation (`.mp4`) with synchronized slides, high-quality speech synthesis, and smooth cinematic transitions.

---

## Architecture Overview

```
.pptx Presentation
       │
       ├──► 1. Text & Notes Extractor (python-pptx)
       │         └──► Bullet Points & Speaker Notes
       │
       ├──► 2. Slide Rasterizer (PowerPoint COM via comtypes -> PDF -> pdf2image)
       │         └──► High-Resolution Slide PNGs (assets/images/slide_{n}.png)
       │
       ├──► 3. AI Script Director (google-genai)
       │         ├──► Gemini 2.5 Flash / 2.0 Flash (Strict JSON Schema)
       │         └──► Energetic Cartoon Narrator Dry-Run Fallback
       │
       ├──► 4. Voice Synthesizer (edge-tts via asyncio)
       │         └──► Natural Spoken Audio MP3s (assets/audio/slide_{n}.mp3)
       │
       └──► 5. Sequential Video Assembly (MoviePy >= 2.0.0)
                 └──► Final Production Video (output/<name>_presentation.mp4)
```

---

## Directory Structure

```
├── requirements.txt         # Pinned pipeline dependencies
├── install.py               # Automated dependency installer & system environment auditor
├── .env.example             # Configuration template for GEMINI_API_KEY
├── config.py                # Environment loader, asset directories & non-crashing key handler
├── extractor.py             # Slide text/notes parsing & PowerPoint COM to PNG conversion
├── ai_director.py           # Gemini structured JSON script generator & mock narrator
├── tts_engine.py            # Asynchronous Edge-TTS generator (en-US-GuyNeural)
├── video_renderer.py        # MoviePy 2.0 sequential video assembler
├── main.py                  # End-to-end CLI runner
└── create_demo_pptx.py      # Demo presentation generator for testing
```

---

## Setup & Quickstart

### 1. Automated Dependency Setup
Run the automated installer to check and install all Python packages and audit system binaries:
```bash
python install.py
```

### 2. Configure Environment (Optional for Testing)
Copy `.env.example` to `.env`:
```bash
cp .env.example .env
```
Add your Google Gemini API key:
```env
GEMINI_API_KEY=your_actual_gemini_api_key_here
```
> **Deferred API Key Handling**: If `GEMINI_API_KEY` is omitted or empty, the pipeline will notify you and seamlessly fall back to an energetic mock cartoon narrator, allowing full end-to-end dry-run testing.

### 3. Generate a Test Presentation
```bash
python create_demo_pptx.py
```

### 4. Run the Pipeline
Convert your presentation into an animated video:
```bash
# Offline/Dry-run mode (no API key required)
python main.py --input demo_presentation.pptx --mock

# Production mode with Gemini 2.5 Flash
python main.py --input demo_presentation.pptx
```

---

## CLI Options

| Flag | Description | Default |
|------|-------------|---------|
| `--input`, `-i` | Path to the PowerPoint (`.pptx`) file | *(Required)* |
| `--output`, `-o` | Custom path for output `.mp4` video | `output/<name>_presentation.mp4` |
| `--mock`, `-m` | Force mock narrator script without calling Gemini API | `False` |
| `--voice`, `-v` | Edge-TTS voice identifier | `en-US-GuyNeural` |
| `--check-only` | Audit system and Python dependencies then exit | `False` |

---

## MoviePy 2.0 Compliance
This codebase strictly conforms to modern MoviePy 2.0 conventions:
- Explicit module imports: `from moviepy import ImageClip, AudioFileClip, concatenate_videoclips`.
- Immutable method chaining: `.with_duration()` and `.with_audio()`.
- Smooth entrance animation: `.with_effects([CrossFadeIn(duration=0.5)])`.
- Sequential assembly: `concatenate_videoclips(clips, method="compose")`.
