"""
Provision PPTX Dubbing - Generate Narration
=============================================
Takes translated text segments and generates TTS narration audio
using the ElevenLabs API.

Usage:
    python generate_narration.py --config path/to/config.json
"""

import argparse
import json
import os
import time
from pathlib import Path

import requests


def load_config(config_path: str) -> dict:
    """Load project configuration from JSON file."""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_path(project_dir: str, relative_path: str) -> str:
    """Resolve a path relative to the project directory."""
    return os.path.join(project_dir, relative_path)


def get_audio_duration(file_path: str, ffprobe_path: str = "ffprobe") -> float:
    """Get the duration of an audio file in seconds using ffprobe."""
    import subprocess

    cmd = [
        ffprobe_path, "-v", "quiet",
        "-print_format", "json",
        "-show_format",
        file_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print(f"WARNING: ffprobe failed for {file_path}: {result.stderr}")
        return 0.0

    data = json.loads(result.stdout)
    return float(data.get("format", {}).get("duration", 0.0))


def generate_speech_elevenlabs(
    text: str,
    api_key: str,
    voice_id: str,
    model_id: str = "eleven_multilingual_v2",
    voice_settings: dict = None,
    output_path: str = "output.mp3",
) -> bool:
    """
    Generate speech audio from text using ElevenLabs API.

    Returns True on success, False on failure.
    """
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"

    headers = {
        "Accept": "audio/mpeg",
        "Content-Type": "application/json",
        "xi-api-key": api_key,
    }

    # Default voice settings if not provided
    if voice_settings is None:
        voice_settings = {
            "stability": 0.5,
            "similarity_boost": 0.75,
            "style": 0.0,
            "use_speaker_boost": True,
        }

    payload = {
        "text": text,
        "model_id": model_id,
        "voice_settings": voice_settings,
    }

    try:
        response = requests.post(url, json=payload, headers=headers, timeout=120)

        if response.status_code == 200:
            with open(output_path, "wb") as f:
                f.write(response.content)
            return True
        else:
            print(f"ERROR: ElevenLabs API returned {response.status_code}: {response.text}")
            return False

    except requests.exceptions.RequestException as e:
        print(f"ERROR: Request failed: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Generate TTS narration for translated slides")
    parser.add_argument("--config", required=True, help="Path to config.json")
    args = parser.parse_args()

    config = load_config(args.config)
    project_dir = config["project_dir"]

    # Resolve paths
    narration_dir = resolve_path(project_dir, config["narration_dir"])
    narration_texts_path = os.path.join(narration_dir, "narration_texts.json")
    segments_info_path = os.path.join(narration_dir, "segments_info.json")

    os.makedirs(narration_dir, exist_ok=True)

    # Load narration texts (produced by translate_pptx.py)
    if not os.path.exists(narration_texts_path):
        print(f"ERROR: Narration texts not found at {narration_texts_path}")
        print("Run translate_pptx.py first to generate narration texts.")
        return

    with open(narration_texts_path, "r", encoding="utf-8") as f:
        narration_texts = json.load(f)

    print(f"Generating narration for {len(narration_texts)} slides...")

    # ElevenLabs config
    api_key = config["elevenlabs_api_key"]
    voice_id = config["elevenlabs_voice_id"]
    model_id = config.get("elevenlabs_model_id", "eleven_multilingual_v2")
    voice_settings = config.get("elevenlabs_voice_settings", None)
    ffprobe_path = config.get("ffprobe_path", "ffprobe")

    segments_info = []

    for item in narration_texts:
        slide_idx = item["slide_index"]
        text = item["narration_text"]

        if not text.strip():
            print(f"  Slide {slide_idx + 1}: No text, skipping narration")
            segments_info.append({
                "slide_index": slide_idx,
                "audio_file": None,
                "duration_seconds": 0.0,
                "text": "",
            })
            continue

        output_filename = f"slide_{slide_idx + 1:03d}.mp3"
        output_path = os.path.join(narration_dir, output_filename)

        print(f"  Slide {slide_idx + 1}: Generating narration ({len(text)} chars)...")

        success = generate_speech_elevenlabs(
            text=text,
            api_key=api_key,
            voice_id=voice_id,
            model_id=model_id,
            voice_settings=voice_settings,
            output_path=output_path,
        )

        if success:
            duration = get_audio_duration(output_path, ffprobe_path)
            print(f"    Saved: {output_filename} ({duration:.2f}s)")

            segments_info.append({
                "slide_index": slide_idx,
                "audio_file": output_filename,
                "duration_seconds": duration,
                "text": text,
            })
        else:
            print(f"    FAILED: Could not generate narration for slide {slide_idx + 1}")
            segments_info.append({
                "slide_index": slide_idx,
                "audio_file": None,
                "duration_seconds": 0.0,
                "text": text,
                "error": True,
            })

        # Rate limit: ElevenLabs has rate limits, add a small delay
        time.sleep(0.5)

    # Save segments info
    with open(segments_info_path, "w", encoding="utf-8") as f:
        json.dump(segments_info, f, ensure_ascii=False, indent=2)

    print(f"\nNarration generation complete.")
    print(f"Segments info saved to: {segments_info_path}")

    # Summary
    total_duration = sum(s["duration_seconds"] for s in segments_info)
    generated = sum(1 for s in segments_info if s["audio_file"] is not None)
    failed = sum(1 for s in segments_info if s.get("error"))
    print(f"  Generated: {generated}/{len(segments_info)} slides")
    print(f"  Failed: {failed}")
    print(f"  Total narration duration: {total_duration:.2f}s")


if __name__ == "__main__":
    main()
