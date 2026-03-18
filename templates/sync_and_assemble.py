"""
Provision PPTX Dubbing - Sync and Assemble
=============================================
Reads the slide timing map and narration segments, positions each narration
clip at the correct timestamp, applies smart-trim for duration matching,
and produces the final dubbed video.

Usage:
    python sync_and_assemble.py --config path/to/config.json
"""

import argparse
import json
import os
import subprocess
from pathlib import Path


def load_config(config_path: str) -> dict:
    """Load project configuration from JSON file."""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_path(project_dir: str, relative_path: str) -> str:
    """Resolve a path relative to the project directory."""
    return os.path.join(project_dir, relative_path)


def get_media_duration(file_path: str, ffprobe_path: str = "ffprobe") -> float:
    """Get duration of a media file in seconds."""
    cmd = [
        ffprobe_path, "-v", "quiet",
        "-print_format", "json",
        "-show_format",
        file_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        return 0.0
    data = json.loads(result.stdout)
    return float(data.get("format", {}).get("duration", 0.0))


def run_ffmpeg(cmd: list[str], description: str = "") -> bool:
    """Run an FFmpeg command and handle errors."""
    if description:
        print(f"  {description}")

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ERROR: FFmpeg failed: {result.stderr[:500]}")
        return False
    return True


def create_silent_audio(duration: float, output_path: str, ffmpeg_path: str = "ffmpeg") -> bool:
    """Create a silent audio track of the specified duration."""
    cmd = [
        ffmpeg_path,
        "-f", "lavfi",
        "-i", f"anullsrc=r=44100:cl=stereo",
        "-t", str(duration),
        "-c:a", "aac",
        "-b:a", "128k",
        output_path,
        "-y",
    ]
    return run_ffmpeg(cmd, f"Creating silent audio track ({duration:.2f}s)")


def compute_smart_trim(
    timing_map: list[dict],
    segments_info: list[dict],
    target_duration: float | None,
    max_ratio: float = 2.0,
    min_padding: float = 0.5,
) -> list[dict]:
    """
    Compute smart-trim adjustments for slides that are too long.

    The smart-trim algorithm:
    1. Calculate total video duration vs target duration.
    2. Identify trimmable slides (where slide duration >> narration duration).
    3. Proportionally trim excess time from trimmable slides.
    4. Never trim below narration duration + min_padding.

    Returns the timing_map with adjusted durations and an 'adjusted_end_time' field.
    """
    if target_duration is None:
        # No target duration specified; use the video as-is
        for entry in timing_map:
            entry["adjusted_start_time"] = entry["start_time"]
            entry["adjusted_end_time"] = entry["end_time"]
            entry["adjusted_duration"] = entry["duration"]
            entry["trimmed"] = False
        return timing_map

    total_video_duration = sum(entry["duration"] for entry in timing_map)
    excess = total_video_duration - target_duration

    if excess <= 0:
        # Video is already shorter than target, no trimming needed
        for entry in timing_map:
            entry["adjusted_start_time"] = entry["start_time"]
            entry["adjusted_end_time"] = entry["end_time"]
            entry["adjusted_duration"] = entry["duration"]
            entry["trimmed"] = False
        return timing_map

    print(f"\nSmart-trim: Video is {excess:.2f}s longer than target ({target_duration:.2f}s)")

    # Build narration duration lookup by slide number
    narration_lookup = {}
    for seg in segments_info:
        slide_num = seg["slide_index"] + 1  # Convert 0-indexed to 1-indexed
        narration_lookup[slide_num] = seg["duration_seconds"]

    # Identify trimmable slides and their excess
    trimmable = []
    for entry in timing_map:
        slide_num = entry["slide_number"]
        narration_dur = narration_lookup.get(slide_num, 0.0)
        min_duration = narration_dur + min_padding
        slide_excess = entry["duration"] - min_duration

        if slide_excess > 0 and entry["duration"] / max(narration_dur, 0.1) > 1.2:
            trimmable.append({
                "slide_number": slide_num,
                "current_duration": entry["duration"],
                "min_duration": min_duration,
                "max_trimmable": slide_excess,
            })

    total_trimmable = sum(t["max_trimmable"] for t in trimmable)

    if total_trimmable <= 0:
        print("  No slides can be trimmed. Output will exceed target duration.")
        for entry in timing_map:
            entry["adjusted_start_time"] = entry["start_time"]
            entry["adjusted_end_time"] = entry["end_time"]
            entry["adjusted_duration"] = entry["duration"]
            entry["trimmed"] = False
        return timing_map

    # Calculate trim ratio (how much of each slide's excess to remove)
    trim_ratio = min(excess / total_trimmable, 1.0)

    print(f"  Trimmable slides: {len(trimmable)}")
    print(f"  Total trimmable time: {total_trimmable:.2f}s")
    print(f"  Trim ratio: {trim_ratio:.2%}")

    # Build trim map
    trim_amounts = {}
    for t in trimmable:
        trim_amount = t["max_trimmable"] * trim_ratio
        trim_amounts[t["slide_number"]] = trim_amount

    # Apply trims and recalculate timestamps
    cumulative_offset = 0.0
    for entry in timing_map:
        slide_num = entry["slide_number"]
        trim = trim_amounts.get(slide_num, 0.0)

        entry["adjusted_start_time"] = round(entry["start_time"] - cumulative_offset, 3)
        entry["adjusted_duration"] = round(entry["duration"] - trim, 3)
        entry["adjusted_end_time"] = round(entry["adjusted_start_time"] + entry["adjusted_duration"], 3)
        entry["trimmed"] = trim > 0.01

        if entry["trimmed"]:
            print(f"  Slide {slide_num}: trimmed {trim:.2f}s "
                  f"({entry['duration']:.2f}s -> {entry['adjusted_duration']:.2f}s)")

        cumulative_offset += trim

    return timing_map


def overlay_narration_segments(
    video_path: str,
    timing_map: list[dict],
    segments_info: list[dict],
    narration_dir: str,
    output_path: str,
    ffmpeg_path: str = "ffmpeg",
    max_speedup: float = 1.1,
) -> bool:
    """
    Overlay narration audio segments onto the video at the correct timestamps.

    Steps:
    1. Create a silent audio track matching video duration.
    2. For each slide with narration, position the audio at the slide's start time.
    3. If narration is longer than slide duration, speed it up (up to max_speedup).
    4. Merge the narration audio track with the video.
    """
    # Get video duration
    video_duration = get_media_duration(video_path, ffmpeg_path.replace("ffmpeg", "ffprobe"))

    # Build filter complex for overlaying all narration segments
    inputs = ["-i", video_path]
    filter_parts = []
    audio_inputs = []
    input_idx = 1  # 0 is the video

    for seg in segments_info:
        if seg["audio_file"] is None:
            continue

        slide_num = seg["slide_index"] + 1
        audio_path = os.path.join(narration_dir, seg["audio_file"])

        if not os.path.exists(audio_path):
            print(f"  WARNING: Audio file not found: {audio_path}")
            continue

        # Find the matching timing entry
        timing_entry = None
        for t in timing_map:
            if t["slide_number"] == slide_num:
                timing_entry = t
                break

        if timing_entry is None:
            print(f"  WARNING: No timing entry for slide {slide_num}")
            continue

        # Determine start time and available duration
        start_time = timing_entry.get("adjusted_start_time", timing_entry["start_time"])
        available_duration = timing_entry.get("adjusted_duration", timing_entry["duration"])
        narration_duration = seg["duration_seconds"]

        inputs.extend(["-i", audio_path])

        # Apply tempo adjustment if narration is too long
        if narration_duration > available_duration and max_speedup > 1.0:
            speedup = min(narration_duration / available_duration, max_speedup)
            filter_parts.append(
                f"[{input_idx}:a]atempo={speedup:.4f},adelay={int(start_time * 1000)}|{int(start_time * 1000)}[a{input_idx}]"
            )
        else:
            filter_parts.append(
                f"[{input_idx}:a]adelay={int(start_time * 1000)}|{int(start_time * 1000)}[a{input_idx}]"
            )

        audio_inputs.append(f"[a{input_idx}]")
        input_idx += 1

    if not audio_inputs:
        print("  No narration segments to overlay. Copying video as-is.")
        cmd = [ffmpeg_path, "-i", video_path, "-c", "copy", output_path, "-y"]
        return run_ffmpeg(cmd, "Copying video without narration")

    # Merge all audio streams
    mix_inputs = "".join(audio_inputs)
    filter_parts.append(
        f"{mix_inputs}amix=inputs={len(audio_inputs)}:duration=longest:dropout_transition=0[mixed]"
    )

    filter_complex = ";".join(filter_parts)

    # Build final command
    cmd = [
        ffmpeg_path,
        *inputs,
        "-filter_complex", filter_complex,
        "-map", "0:v",      # Video from original
        "-map", "[mixed]",   # Mixed narration audio
        "-c:v", "copy",      # Copy video stream (no re-encode)
        "-c:a", "aac",
        "-b:a", "192k",
        "-shortest",
        output_path,
        "-y",
    ]

    return run_ffmpeg(cmd, "Merging narration with video")


def apply_smart_trim_to_video(
    input_video: str,
    timing_map: list[dict],
    output_path: str,
    ffmpeg_path: str = "ffmpeg",
) -> bool:
    """
    Apply smart-trim by re-cutting the video based on adjusted timestamps.
    Uses FFmpeg's trim filter to cut each slide segment and concatenate them.
    """
    has_trims = any(entry.get("trimmed", False) for entry in timing_map)
    if not has_trims:
        # No trimming needed, just rename/copy
        if input_video != output_path:
            cmd = [ffmpeg_path, "-i", input_video, "-c", "copy", output_path, "-y"]
            return run_ffmpeg(cmd, "No trimming needed, copying output")
        return True

    print("\nApplying smart-trim to video...")

    # Build FFmpeg filter for trimming and concatenating
    filter_parts = []
    concat_parts_v = []
    concat_parts_a = []

    for i, entry in enumerate(timing_map):
        start = entry["start_time"]
        # Trim from the original start, but only keep adjusted_duration worth
        duration = entry.get("adjusted_duration", entry["duration"])

        filter_parts.append(
            f"[0:v]trim=start={start}:duration={duration},setpts=PTS-STARTPTS[v{i}]"
        )
        filter_parts.append(
            f"[0:a]atrim=start={start}:duration={duration},asetpts=PTS-STARTPTS[a{i}]"
        )
        concat_parts_v.append(f"[v{i}]")
        concat_parts_a.append(f"[a{i}]")

    n = len(timing_map)
    concat_inputs = "".join(
        f"{concat_parts_v[i]}{concat_parts_a[i]}" for i in range(n)
    )
    filter_parts.append(
        f"{concat_inputs}concat=n={n}:v=1:a=1[outv][outa]"
    )

    filter_complex = ";".join(filter_parts)

    cmd = [
        ffmpeg_path,
        "-i", input_video,
        "-filter_complex", filter_complex,
        "-map", "[outv]",
        "-map", "[outa]",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "18",
        "-c:a", "aac",
        "-b:a", "192k",
        output_path,
        "-y",
    ]

    return run_ffmpeg(cmd, f"Smart-trimming {sum(1 for e in timing_map if e.get('trimmed'))} slides")


def main():
    parser = argparse.ArgumentParser(description="Sync narration and assemble final dubbed video")
    parser.add_argument("--config", required=True, help="Path to config.json")
    args = parser.parse_args()

    config = load_config(args.config)
    project_dir = config["project_dir"]

    # Resolve paths
    video_path = resolve_path(project_dir, config["exported_video"])
    narration_dir = resolve_path(project_dir, config["narration_dir"])
    output_dir = resolve_path(project_dir, config["output_dir"])
    timing_map_path = os.path.join(output_dir, "slide_timing_map.json")
    segments_info_path = os.path.join(narration_dir, "segments_info.json")
    final_output = os.path.join(output_dir, "final_dubbed.mp4")
    temp_narrated = os.path.join(output_dir, "temp_narrated.mp4")

    ffmpeg_path = config.get("ffmpeg_path", "ffmpeg")
    ffprobe_path = config.get("ffprobe_path", "ffprobe")

    # Load timing map
    if not os.path.exists(timing_map_path):
        print(f"ERROR: Slide timing map not found: {timing_map_path}")
        print("Run detect_transitions.py first.")
        return

    with open(timing_map_path, "r", encoding="utf-8") as f:
        timing_map = json.load(f)

    # Load segments info
    if not os.path.exists(segments_info_path):
        print(f"ERROR: Segments info not found: {segments_info_path}")
        print("Run generate_narration.py first.")
        return

    with open(segments_info_path, "r", encoding="utf-8") as f:
        segments_info = json.load(f)

    print(f"Loaded timing map: {len(timing_map)} slides")
    print(f"Loaded narration segments: {len(segments_info)} segments")

    # --- Smart-trim computation ---
    target_duration = config.get("target_duration_seconds", None)
    max_ratio = config.get("max_slide_duration_ratio", 2.0)
    min_padding = config.get("min_slide_padding_seconds", 0.5)
    max_speedup = config.get("max_narration_speedup", 1.1)

    timing_map = compute_smart_trim(
        timing_map, segments_info, target_duration, max_ratio, min_padding
    )

    # --- Overlay narration ---
    print("\nOverlaying narration segments...")
    success = overlay_narration_segments(
        video_path=video_path,
        timing_map=timing_map,
        segments_info=segments_info,
        narration_dir=narration_dir,
        output_path=temp_narrated,
        ffmpeg_path=ffmpeg_path,
        max_speedup=max_speedup,
    )

    if not success:
        print("ERROR: Failed to overlay narration. Aborting.")
        return

    # --- Apply smart-trim ---
    success = apply_smart_trim_to_video(
        input_video=temp_narrated,
        timing_map=timing_map,
        output_path=final_output,
        ffmpeg_path=ffmpeg_path,
    )

    if not success:
        print("ERROR: Failed to apply smart-trim. Aborting.")
        return

    # Clean up temp file
    if os.path.exists(temp_narrated) and temp_narrated != final_output:
        os.unlink(temp_narrated)

    # --- Summary ---
    final_duration = get_media_duration(final_output, ffprobe_path)
    print(f"\n{'=' * 50}")
    print(f"Final dubbed video: {final_output}")
    print(f"Duration: {final_duration:.2f}s")
    if target_duration:
        diff = final_duration - target_duration
        print(f"Target duration: {target_duration:.2f}s (diff: {diff:+.2f}s)")
    print(f"{'=' * 50}")

    # Save the updated timing map with trim info
    adjusted_map_path = os.path.join(output_dir, "slide_timing_map_adjusted.json")
    with open(adjusted_map_path, "w", encoding="utf-8") as f:
        json.dump(timing_map, f, indent=2)
    print(f"Adjusted timing map saved to: {adjusted_map_path}")


if __name__ == "__main__":
    main()
