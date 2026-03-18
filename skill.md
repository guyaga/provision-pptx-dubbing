---
name: provision-pptx-dubbing
description: "Translate PowerPoint presentations and create dubbed video versions with synced narration for Provision ISR. Handles PPTX text translation, TTS generation via ElevenLabs, PowerPoint video export, slide transition detection with Gemini 3.1 Pro, and audio synchronization. Use for: presentation dubbing, PPTX translation, multilingual training videos. Triggers: provision pptx, translate presentation, dub presentation, pptx dubbing, provision slides"
allowed-tools: Read, Write, Edit, Bash, Glob, Grep
---

# Provision PPTX Dubbing Pipeline

Translate PowerPoint presentations into other languages and produce dubbed video versions with synchronized narration. Built for Provision ISR to localize training and product presentations.

## Pipeline Overview

The dubbing pipeline has 7 stages:

1. **Extract & Translate** - Open the PPTX (a zip of XML files), extract all slide text, translate via Gemini AI, replace text in the XML, save translated PPTX.
2. **Generate Narration** - Send translated text segments to ElevenLabs TTS API, save individual MP3 files per slide.
3. **Export to Video** - Use PowerPoint COM automation (via PowerShell) to export the translated PPTX as an MP4 video.
4. **Detect Slide Transitions** - Extract frames from the exported video at regular intervals, send them to Gemini 3.1 Pro for frame-by-frame analysis to detect exact slide transition timestamps.
5. **Sync & Assemble** - Map narration audio segments to the detected slide timestamps, overlay audio onto the video.
6. **Smart-Trim** - Trim slides that are too long (due to embedded videos, 3D animations, long transitions) to match original duration.
7. **Final Output** - Produce the final dubbed video with synced narration.

## Prerequisites

### Software
- Python 3.10+
- PowerShell 5.1+ (Windows)
- FFmpeg (must be on PATH)
- Microsoft PowerPoint (for COM automation video export)

### Python Packages
```
pip install python-pptx google-generativeai elevenlabs requests
```

### API Keys
- **Google Gemini API Key** - For translation and slide transition detection
- **ElevenLabs API Key** - For TTS narration generation

### FFmpeg
Download from https://ffmpeg.org/download.html and ensure `ffmpeg` is on your PATH.

## Project Structure

When starting a new dubbing project, create this structure:

```
project_dir/
    config.json              # Project configuration
    input/
        original.pptx        # Source presentation
    translated/
        translated.pptx      # Output: translated presentation
    narration/
        slide_001.mp3        # Output: per-slide narration audio
        slide_002.mp3
        ...
        segments_info.json   # Narration timing metadata
    video/
        exported.mp4         # Output: PowerPoint video export
        frames/              # Extracted frames for transition detection
    output/
        slide_timing_map.json # Detected slide transitions
        final_dubbed.mp4     # Final output video
```

Create this structure with:
```bash
mkdir -p project_dir/{input,translated,narration,video/frames,output}
```

## Step-by-Step Workflow

### Step 1: Configure the Project

Copy `templates/config.json` to your project directory and fill in:
- Path to the source PPTX
- Target language
- Gemini API key
- ElevenLabs API key and voice ID
- Output preferences

```bash
cp templates/config.json project_dir/config.json
# Edit config.json with your settings
```

### Step 2: Translate the PPTX

```bash
python templates/translate_pptx.py --config project_dir/config.json
```

This will:
- Open the source PPTX
- Extract all text from every slide
- Send text to Gemini for translation (preserving formatting placeholders)
- Replace text in the PPTX XML
- Save the translated PPTX to `translated/translated.pptx`

### Step 3: Generate Narration Audio

```bash
python templates/generate_narration.py --config project_dir/config.json
```

This will:
- Read the translated text for each slide
- Call ElevenLabs API to generate speech for each slide
- Save individual MP3 files in `narration/`
- Write `narration/segments_info.json` with duration metadata

### Step 4: Export Translated PPTX to Video

```powershell
powershell -ExecutionPolicy Bypass -File templates/export_pptx_video.ps1 -ConfigPath project_dir/config.json
```

This will:
- Open the translated PPTX in PowerPoint via COM
- Optionally set fast animation delays to reduce dead air
- Export as MP4 video to `video/exported.mp4`

### Step 5: Detect Slide Transitions

```bash
python templates/detect_transitions.py --config project_dir/config.json
```

This will:
- Extract frames from the exported video at regular intervals using FFmpeg
- Send batches of frames to Gemini 3.1 Pro for analysis
- Gemini identifies which frames correspond to slide transitions
- Output `output/slide_timing_map.json` with start/end timestamps per slide

### Step 6: Sync and Assemble Final Video

```bash
python templates/sync_and_assemble.py --config project_dir/config.json
```

This will:
- Read the slide timing map and narration segment info
- Position each narration clip at the correct timestamp
- Smart-trim slides that exceed target duration
- Use FFmpeg to merge everything into the final dubbed video

## Configuration Reference

See `templates/config.json` for all available options. Key settings:

| Setting | Description |
|---------|-------------|
| `source_pptx` | Path to original PPTX file |
| `target_language` | Target language (e.g., "Spanish", "Hebrew", "French") |
| `gemini_api_key` | Google Gemini API key |
| `gemini_model_translation` | Model for translation (default: gemini-2.0-flash) |
| `gemini_model_vision` | Model for frame analysis (default: gemini-3.1-pro) |
| `elevenlabs_api_key` | ElevenLabs API key |
| `elevenlabs_voice_id` | Voice ID for TTS |
| `frame_extraction_interval` | Seconds between extracted frames (default: 0.5) |
| `max_slide_duration_ratio` | Max ratio of slide duration to narration (for smart-trim) |

## ElevenLabs Voice Setup

### Finding Voice IDs
1. Go to https://elevenlabs.io/voice-library
2. Browse or search for voices in your target language
3. Add a voice to your library
4. Go to https://elevenlabs.io/app/voice-lab to find the voice ID

### Recommended Voices by Language
- **Spanish**: Search for native Spanish speakers with clear enunciation
- **Hebrew**: Use voices trained on Hebrew; ensure RTL text handling in PPTX
- **French**: Professional French voices work well for corporate presentations
- **Arabic**: Ensure the voice supports MSA (Modern Standard Arabic)

### Voice Settings in config.json
```json
{
    "elevenlabs_voice_id": "YOUR_VOICE_ID",
    "elevenlabs_model_id": "eleven_multilingual_v2",
    "elevenlabs_voice_settings": {
        "stability": 0.5,
        "similarity_boost": 0.75,
        "style": 0.0,
        "use_speaker_boost": true
    }
}
```

## Handling Common Issues

### Animation Delays / Dead Air
PowerPoint presentations often have complex animations that create long pauses in the exported video. The pipeline handles this by:
- Setting animation delays to their minimum via COM automation before export
- Detecting actual slide boundaries (not animation steps) via Gemini vision
- Smart-trimming excess duration

### Embedded Videos in Slides
Slides with embedded videos will have longer durations in the export. The pipeline:
- Detects these slides via Gemini frame analysis (it sees video content playing)
- Preserves the embedded video duration but trims surrounding dead time
- Adjusts narration placement to avoid overlapping with embedded audio

### Timing Mismatches
If narration is longer than the slide duration:
- The smart-trim algorithm extends the slide slightly or speeds up narration by up to 10%
- If the mismatch is severe (>2x), it flags the slide for manual review

If narration is shorter than the slide duration:
- The narration is centered within the slide duration
- Remaining time is left as original audio (or silence)

### RTL Languages (Hebrew, Arabic)
- python-pptx handles RTL text correctly in most cases
- Verify the translated PPTX renders correctly in PowerPoint before exporting
- Some bullet formatting may need manual adjustment

## Smart-Trim Approach

The smart-trim algorithm ensures the final dubbed video matches the original presentation duration:

1. **Calculate target duration** from the original video export (or a user-specified duration).
2. **Sum all slide durations** from the transition detection.
3. **Identify trimmable slides** - slides where detected duration significantly exceeds narration duration.
4. **Proportionally trim** trimmable slides, prioritizing:
   - Slides with the most excess (dead air after animations complete)
   - Slides without embedded video content
5. **Apply minimum duration** - never trim a slide below its narration duration + 0.5s padding.
6. **Re-encode** the video with adjusted timestamps using FFmpeg's `trim` and `concat` filters.

## Template Files

- `templates/translate_pptx.py` - PPTX text extraction and translation
- `templates/generate_narration.py` - ElevenLabs TTS generation
- `templates/export_pptx_video.ps1` - PowerPoint COM video export
- `templates/detect_transitions.py` - Gemini 3.1 Pro slide transition detection
- `templates/sync_and_assemble.py` - Audio sync and final video assembly
- `templates/config.json` - Configuration template
