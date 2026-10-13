"""ai_director.py - Formal & Academic AI Narration Script Generator using OpenRouter.

Transforms slide bullets and speaker notes into a concise, formal, and academic
narration script using NVIDIA Nemotron via OpenRouter and strict structured JSON output.
Features a grounded offline mock generator adhering to corporate/academic standards.
"""

import asyncio
import json
import re
from typing import Any
import requests

from config import DEFAULT_MODEL, FALLBACK_MODELS


def generate_mock_script(slides_data: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Generates a formal, direct, and academic narration script locally without an API key.

    Adheres strictly to professional standards: no greetings, banter, or filler phrases.
    """
    mock_results: list[dict[str, Any]] = []

    for item in slides_data:
        slide_num = item.get("slide_number", 1)
        bullets = item.get("bullet_points", [])
        notes = item.get("speaker_notes", "")

        statements: list[str] = []

        if bullets:
            primary = bullets[0].rstrip(".")
            statements.append(f"This section focuses on {primary}.")

            if len(bullets) > 1:
                secondary_points = "; ".join(b.rstrip(".") for b in bullets[1:4])
                statements.append(f"Key technical aspects include: {secondary_points}.")
        else:
            statements.append("This slide presents the primary conceptual overview.")

        if notes:
            clean_notes = notes.strip().rstrip(".")
            statements.append(f"Specifically, {clean_notes}.")

        narration = " ".join(statements)

        mock_results.append({
            "slide_number": slide_num,
            "narration_script": narration,
        })

    return mock_results


def clean_json_response(raw_text: str) -> str:
    """Strips markdown formatting and code blocks from LLM response to safely parse JSON."""
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\n?", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\n?```$", "", text)
    return text.strip()


async def generate_script(
    slides_data: list[dict[str, Any]],
    api_key: str = "",
) -> list[dict[str, Any]]:
    """Asynchronous function to generate a formal and academic narration script for all slides.

    Uses OpenRouter API (NVIDIA Nemotron 3.5 Lightning) with standard requests.

    Args:
        slides_data: List of dicts with slide_number, bullet_points, speaker_notes.
        api_key: OpenRouter API key.

    Returns:
        List of dicts: [{'slide_number': int, 'narration_script': str}, ...]
    """
    cleaned_key = api_key.strip() if api_key else ""
    if not cleaned_key or cleaned_key in ["your_actual_api_key_here", "your_gemini_api_key_here"]:
        print("[*] Dry-run/Mock Mode: No OPENROUTER_API_KEY supplied. Generating formal academic narration.")
        return generate_mock_script(slides_data)

    system_prompt = (
        "You are an executive corporate and academic technical presenter. "
        "Your objective is to provide a concise, direct, and rigorous spoken narration "
        "for each slide based strictly on the provided bullet points and speaker notes.\n\n"
        "Strict Formatting & Tone Guidelines:\n"
        "1. Tone: Formal, objective, authoritative, and direct.\n"
        "2. Strictly Prohibit Conversational Greetings: Do NOT include 'Welcome', 'Hello', "
        "'Hey guys', 'Good morning', or any introductory salutations.\n"
        "3. Strictly Prohibit Filler & Rhetorical Banter: Do NOT use rhetorical questions, "
        "emotional adjectives, jokes, conversational transitions, or fluff.\n"
        "4. Strictly Prohibit Conversational Sign-offs: Do NOT include 'Thank you for watching', "
        "'That wraps it up', 'Hope you enjoyed', or any closing remarks.\n"
        "5. Direct Factual Explanation: Immediately explain the technical facts, architecture, "
        "methodologies, and data on each slide in clear, professional sentences.\n"
        "6. Technical Precision: Faithfully integrate all facts and speaker notes without omission.\n\n"
        "CRITICAL OUTPUT FORMAT: You must output ONLY a valid JSON array. Do not include any conversational "
        "text before or after the JSON. The JSON array must contain objects with exactly two keys: "
        "'slide_number' (integer) and 'narration_script' (string)."
    )

    user_prompt = f"Slides Content:\n{json.dumps(slides_data, indent=2)}"

    headers = {
        "Authorization": f"Bearer {cleaned_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/ppt-to-animation",
        "X-Title": "PPT to Animation",
    }

    target_models = list(dict.fromkeys(FALLBACK_MODELS))
    if "nvidia/nemotron-3.5-lightning:free" not in target_models:
        target_models.insert(0, "nvidia/nemotron-3.5-lightning:free")

    for model_name in target_models:
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.2,
        }

        print(f"[*] Querying AI Director using OpenRouter ({model_name})...")

        try:
            response = await asyncio.to_thread(
                requests.post,
                url="https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                data=json.dumps(payload),
                timeout=60,
            )
            response.raise_for_status()
            data = response.json()

            raw_content = data["choices"][0]["message"]["content"]
            print(f"[+] Successfully received response from OpenRouter ({model_name}).")

            cleaned_text = clean_json_response(raw_content)
            raw_json = json.loads(cleaned_text)

            validated: list[dict[str, Any]] = []
            for item in raw_json:
                validated.append({
                    "slide_number": int(item["slide_number"]),
                    "narration_script": str(item["narration_script"]),
                })

            print(f"[+] Successfully generated formal narration script for {len(validated)} slides via {model_name}.")
            return validated

        except requests.exceptions.RequestException as e:
            print(f"[-] OpenRouter API connection failed for {model_name}: {e}")
        except (json.JSONDecodeError, KeyError) as e:
            print(f"[-] Failed to parse LLM structured JSON from {model_name} ({e}).")
        except Exception as e:
            print(f"[-] An unexpected error occurred with {model_name}: {e}")

    print("[!] Warning: Remote generation failed. Falling back to local formal generator.")
    return generate_mock_script(slides_data)
