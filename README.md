# 🎬 PPTX to AI Video Generator

An automated Python pipeline that transforms standard PowerPoint (`.pptx`) presentations into fully narrated, academic video presentations (`.mp4`). 

Instead of relying on expensive cloud subscriptions, this tool is designed to be 100% free and runs locally. It parses slide text and speaker notes, generates a formal, professional narration script using **NVIDIA Nemotron 3.5 Lightning** (via OpenRouter), synthesizes human-sounding speech using **Edge-TTS**, and stitches the slides and audio together into a polished video using **MoviePy**.

## ✨ Key Features
* **Automated Slide Extraction:** Converts `.pptx` slides into high-quality images and extracts all bullet points, text, and hidden speaker notes.
* **Smart AI Director:** Uses OpenRouter's API to generate concise, academic, and formal narration scripts tailored to each slide's content.
* **High-Quality TTS:** Synthesizes natural-sounding voiceovers using Microsoft Edge-TTS (defaulting to a calm, professional documentary tone).
* **Local Video Rendering:** Combines slide images and audio sequentially into an MP4 file with precise timing using MoviePy 2.0.
* **Cost-Free Architecture:** Utilizes free-tier LLM APIs and local rendering libraries to avoid continuous billing.
* **Fallback Mechanisms:** Includes a mock-generation offline fallback if API keys are missing or the network drops.

## 🛠️ Tech Stack
* **Language:** Python 3.12
* **Slide Processing:** `python-pptx`, `comtypes` (LibreOffice/Word engine), `pdf2image` (Poppler)
* **AI / LLM:** `requests` (OpenRouter REST API)
* **Audio / Video:** `edge-tts`, `moviepy>=2.0.0`

## ⚙️ Prerequisites
1. **Python 3.11 or 3.12:** Recommended to avoid C-extension build issues in newer Python 3.13 versions.
2. **Poppler:** Required by `pdf2image` to convert slide PDFs into image frames. Ensure Poppler is installed and added to your system's PATH.


