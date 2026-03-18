"""
Provision PPTX Dubbing - Detect Slide Transitions
====================================================
Extracts frames from the exported video at regular intervals, sends them
to Gemini 3.1 Pro for frame-by-frame analysis to detect slide transitions,
and outputs a timing map.

Uses Gemini 3.1 Pro (NOT OpenCV) for visual analysis of slide boundaries.

Usage:
    python detect_transitions.py --config path/to/config.json
"""

import argparse
import base64
import json
import os
import subprocess
import time
from pathlib import Path

import google.generativeai as genai


def load_config(config_path: str) -> dict:
    """Load project configuration from JSON file."""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_path(project_dir: str, relative_path: str) -> str:
    """Resolve a path relative to the project directory."""
    return os.path.join(project_dir, relative_path)


def get_video_duration(video_path: str, ffprobe_path: str = "ffprobe") -> float:
    """Get the duration of a video file in seconds."""
    cmd = [
        ffprobe_path, "-v", "quiet",
        "-print_format", "json",
        "-show_format",
        video_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    data = json.loads(result.stdout)
    return float(data["format"]["duration"])


def extract_frames(
    video_path: str,
    frames_dir: str,
    interval: float = 0.5,
    ffmpeg_path: str = "ffmpeg",
) -> list[dict]:
    """
    Extract frames from the video at the specified interval using FFmpeg.

    Returns a list of dicts with frame info:
    [{"frame_index": 0, "timestamp": 0.0, "path": "/path/to/frame_0000.jpg"}, ...]
    """
    os.makedirs(frames_dir, exist_ok=True)

    # Clean existing frames
    for f in Path(frames_dir).glob("frame_*.jpg"):
        f.unlink()

    # Extract frames using FFmpeg
    output_pattern = os.path.join(frames_dir, "frame_%04d.jpg")
    cmd = [
        ffmpeg_path,
        "-i", video_path,
        "-vf", f"fps=1/{interval}",
        "-q:v", "2",  # High quality JPEG
        output_pattern,
        "-y",  # Overwrite
    ]

    print(f"Extracting frames every {interval}s...")
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print(f"FFmpeg error: {result.stderr}")
        raise RuntimeError("Frame extraction failed")

    # List extracted frames
    frames = []
    for frame_path in sorted(Path(frames_dir).glob("frame_*.jpg")):
        # FFmpeg numbers from 1, extract the number
        frame_num = int(frame_path.stem.split("_")[1])
        timestamp = (frame_num - 1) * interval

        frames.append({
            "frame_index": frame_num - 1,
            "timestamp": round(timestamp, 2),
            "path": str(frame_path),
        })

    print(f"  Extracted {len(frames)} frames")
    return frames


def encode_frame_to_base64(frame_path: str) -> str:
    """Read a frame image and encode it as base64."""
    with open(frame_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def detect_transitions_with_gemini(
    frames: list[dict],
    api_key: str,
    model_name: str = "gemini-3.1-pro",
    batch_size: int = 20,
) -> list[dict]:
    """
    Send batches of frames to Gemini 3.1 Pro and ask it to identify
    which frames represent slide transitions.

    Returns a list of transition events:
    [{"timestamp": 5.0, "from_slide": 1, "to_slide": 2}, ...]
    """
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    all_transitions = []
    current_slide = 1

    for batch_start in range(0, len(frames), batch_size):
        batch = frames[batch_start : batch_start + batch_size]

        print(f"  Analyzing frames {batch_start} to {batch_start + len(batch) - 1}...")

        # Build multimodal content: images + text prompt
        content_parts = []

        # Add context about the timestamps
        timestamps_info = ", ".join(
            f"Frame {f['frame_index']}: {f['timestamp']}s" for f in batch
        )

        content_parts.append(
            f"I am analyzing a PowerPoint presentation exported as video. "
            f"Below are {len(batch)} frames extracted at regular intervals. "
            f"The frame timestamps are: {timestamps_info}\n\n"
            f"The current slide number before this batch is {current_slide}.\n\n"
            f"Analyze these frames and identify any slide transitions. "
            f"A slide transition occurs when the visual content changes significantly, "
            f"indicating a new slide. Ignore animation effects within the same slide "
            f"(elements appearing/moving). Focus on complete content changes.\n\n"
            f"Return a JSON array of transition events. Each event should have:\n"
            f'- "frame_index": the index of the frame where the NEW slide first appears\n'
            f'- "timestamp": the timestamp in seconds\n'
            f'- "from_slide": the slide number before the transition\n'
            f'- "to_slide": the slide number after the transition\n\n'
            f"If no transitions occur in this batch, return an empty array: []\n"
            f"Return ONLY the JSON array, nothing else."
        )

        # Add frame images
        for frame in batch:
            image_data = encode_frame_to_base64(frame["path"])
            content_parts.append({
                "mime_type": "image/jpeg",
                "data": image_data,
            })

        try:
            response = model.generate_content(content_parts)
            response_text = response.text.strip()

            # Strip markdown code fences if present
            if response_text.startswith("```"):
                lines = response_text.split("\n")
                response_text = "\n".join(lines[1:-1])

            batch_transitions = json.loads(response_text)

            if batch_transitions:
                all_transitions.extend(batch_transitions)
                # Update current slide number for next batch
                current_slide = batch_transitions[-1]["to_slide"]
                print(f"    Found {len(batch_transitions)} transition(s)")
            else:
                print(f"    No transitions in this batch")

        except json.JSONDecodeError as e:
            print(f"    WARNING: Could not parse Gemini response: {e}")
            print(f"    Raw response: {response_text[:200]}")
        except Exception as e:
            print(f"    ERROR: Gemini API call failed: {e}")

        # Rate limit: be respectful of API limits
        time.sleep(1.0)

    return all_transitions


def build_slide_timing_map(transitions: list[dict], video_duration: float) -> list[dict]:
    """
    Convert transition events into a slide timing map with start/end times.

    Returns:
    [
        {"slide_number": 1, "start_time": 0.0, "end_time": 5.0, "duration": 5.0},
        {"slide_number": 2, "start_time": 5.0, "end_time": 12.5, "duration": 7.5},
        ...
    ]
    """
    timing_map = []

    if not transitions:
        # Single slide or no transitions detected
        timing_map.append({
            "slide_number": 1,
            "start_time": 0.0,
            "end_time": video_duration,
            "duration": video_duration,
        })
        return timing_map

    # First slide: from 0 to the first transition
    first_transition = transitions[0]
    timing_map.append({
        "slide_number": first_transition["from_slide"],
        "start_time": 0.0,
        "end_time": first_transition["timestamp"],
        "duration": first_transition["timestamp"],
    })

    # Middle slides: between consecutive transitions
    for i in range(len(transitions)):
        start = transitions[i]["timestamp"]
        end = transitions[i + 1]["timestamp"] if i + 1 < len(transitions) else video_duration
        slide_num = transitions[i]["to_slide"]

        timing_map.append({
            "slide_number": slide_num,
            "start_time": round(start, 2),
            "end_time": round(end, 2),
            "duration": round(end - start, 2),
        })

    return timing_map


def main():
    parser = argparse.ArgumentParser(description="Detect slide transitions in exported video")
    parser.add_argument("--config", required=True, help="Path to config.json")
    args = parser.parse_args()

    config = load_config(args.config)
    project_dir = config["project_dir"]

    # Resolve paths
    video_path = resolve_path(project_dir, config["exported_video"])
    frames_dir = resolve_path(project_dir, config["frames_dir"])
    output_dir = resolve_path(project_dir, config["output_dir"])
    timing_map_path = os.path.join(output_dir, "slide_timing_map.json")

    os.makedirs(output_dir, exist_ok=True)

    if not os.path.exists(video_path):
        print(f"ERROR: Exported video not found: {video_path}")
        print("Run export_pptx_video.ps1 first.")
        return

    # Get video duration
    ffprobe_path = config.get("ffprobe_path", "ffprobe")
    video_duration = get_video_duration(video_path, ffprobe_path)
    print(f"Video duration: {video_duration:.2f}s")

    # Extract frames
    interval = config.get("frame_extraction_interval", 0.5)
    ffmpeg_path = config.get("ffmpeg_path", "ffmpeg")
    frames = extract_frames(video_path, frames_dir, interval, ffmpeg_path)

    # Detect transitions with Gemini 3.1 Pro
    print("\nDetecting slide transitions with Gemini 3.1 Pro...")
    transitions = detect_transitions_with_gemini(
        frames,
        api_key=config["gemini_api_key"],
        model_name=config.get("gemini_model_vision", "gemini-3.1-pro"),
        batch_size=config.get("frame_batch_size", 20),
    )

    print(f"\nDetected {len(transitions)} transitions")

    # Build timing map
    timing_map = build_slide_timing_map(transitions, video_duration)

    # Save timing map
    with open(timing_map_path, "w", encoding="utf-8") as f:
        json.dump(timing_map, f, indent=2)

    print(f"\nSlide timing map saved to: {timing_map_path}")
    print("\nSlide Timing Summary:")
    print(f"{'Slide':>6} {'Start':>8} {'End':>8} {'Duration':>10}")
    print("-" * 36)
    for entry in timing_map:
        print(
            f"{entry['slide_number']:>6} "
            f"{entry['start_time']:>7.2f}s "
            f"{entry['end_time']:>7.2f}s "
            f"{entry['duration']:>9.2f}s"
        )


if __name__ == "__main__":
    main()
