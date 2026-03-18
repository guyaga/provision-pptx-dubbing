<#
.SYNOPSIS
    Export a PowerPoint presentation to video using COM automation.

.DESCRIPTION
    Opens the translated PPTX in PowerPoint, optionally sets fast animation
    delays to minimize dead air, and exports the presentation as an MP4 video.

.PARAMETER ConfigPath
    Path to the project config.json file.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File export_pptx_video.ps1 -ConfigPath .\config.json
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

# --- Load Configuration ---
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    exit 1
}

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$projectDir = $config.project_dir

# Resolve paths relative to project directory
$translatedPptx = Join-Path $projectDir $config.translated_pptx
$exportedVideo = Join-Path $projectDir $config.exported_video

# Ensure the video output directory exists
$videoDir = Split-Path $exportedVideo -Parent
if (-not (Test-Path $videoDir)) {
    New-Item -ItemType Directory -Path $videoDir -Force | Out-Null
}

# Convert to absolute paths
$translatedPptx = (Resolve-Path $translatedPptx).Path
$exportedVideo = Join-Path (Resolve-Path $videoDir).Path (Split-Path $exportedVideo -Leaf)

# --- Configuration values ---
$resolution = $config.video_export_resolution
if (-not $resolution) { $resolution = "1920x1080" }
$resParts = $resolution -split "x"
$width = [int]$resParts[0]
$height = [int]$resParts[1]

$fps = $config.video_export_fps
if (-not $fps) { $fps = 30 }

$animDelayOverride = $config.animation_delay_override_seconds
if (-not $animDelayOverride) { $animDelayOverride = 0.1 }

Write-Host "=== Provision PPTX Video Export ==="
Write-Host "  Input:  $translatedPptx"
Write-Host "  Output: $exportedVideo"
Write-Host "  Resolution: ${width}x${height} @ ${fps}fps"
Write-Host ""

# --- Start PowerPoint COM ---
$pptApp = $null
$presentation = $null

try {
    Write-Host "Starting PowerPoint..."
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    Write-Host "Opening presentation: $translatedPptx"
    $presentation = $pptApp.Presentations.Open($translatedPptx, $false, $false, $true)

    # --- Optionally speed up animations ---
    # This reduces animation delays to minimize dead air in the exported video.
    # Set to 0 or remove this block to preserve original animation timing.
    if ($animDelayOverride -gt 0) {
        Write-Host "Setting animation delays to ${animDelayOverride}s..."
        foreach ($slide in $presentation.Slides) {
            $timeline = $slide.TimeLine
            if ($timeline.MainSequence.Count -gt 0) {
                foreach ($effect in $timeline.MainSequence) {
                    # Set delay after previous animation
                    try {
                        $effect.Timing.TriggerDelayTime = $animDelayOverride
                    }
                    catch {
                        # Some effects may not support delay modification
                    }
                }
            }
        }
    }

    # --- Export to Video ---
    Write-Host "Exporting to video (this may take several minutes)..."

    # CreateVideo parameters:
    #   FileName: output path
    #   UseTimingsAndNarrations: use slide timings
    #   DefaultSlideDuration: seconds per slide if no timing set
    #   VertResolution: vertical resolution
    #   FramesPerSecond: FPS
    #   Quality: 0-100

    $defaultSlideDuration = 5  # seconds per slide if no timing is set
    $quality = 85

    $presentation.CreateVideo(
        $exportedVideo,
        $true,            # UseTimingsAndNarrations
        $defaultSlideDuration,
        $height,          # VertResolution
        $fps,             # FramesPerSecond
        $quality          # Quality
    )

    # Wait for the export to complete
    # PowerPoint exports asynchronously; poll until done
    Write-Host "Waiting for export to complete..."
    $timeout = 600  # 10 minute timeout
    $elapsed = 0

    while ($presentation.CreateVideoStatus -eq 1) {
        # 1 = ppMediaTaskStatusInProgress
        Start-Sleep -Seconds 2
        $elapsed += 2
        if ($elapsed % 30 -eq 0) {
            Write-Host "  Still exporting... (${elapsed}s elapsed)"
        }
        if ($elapsed -ge $timeout) {
            Write-Error "Video export timed out after ${timeout}s"
            break
        }
    }

    $status = $presentation.CreateVideoStatus
    # Status codes:
    # 0 = ppMediaTaskStatusNone
    # 1 = ppMediaTaskStatusInProgress
    # 2 = ppMediaTaskStatusQueued
    # 3 = ppMediaTaskStatusDone
    # 4 = ppMediaTaskStatusFailed

    if ($status -eq 3) {
        Write-Host ""
        Write-Host "Video export complete: $exportedVideo"

        # Verify the output file exists
        if (Test-Path $exportedVideo) {
            $fileInfo = Get-Item $exportedVideo
            $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
            Write-Host "  File size: ${sizeMB} MB"
        }
    }
    else {
        Write-Error "Video export failed with status: $status"
    }
}
catch {
    Write-Error "Error during export: $_"
    Write-Error $_.Exception.Message
    Write-Error $_.ScriptStackTrace
}
finally {
    # --- Cleanup ---
    if ($presentation) {
        try {
            $presentation.Close()
        }
        catch {
            Write-Warning "Could not close presentation: $_"
        }
    }
    if ($pptApp) {
        try {
            $pptApp.Quit()
        }
        catch {
            Write-Warning "Could not quit PowerPoint: $_"
        }
    }

    # Release COM objects
    if ($presentation) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null }
    if ($pptApp) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "PowerPoint COM cleanup complete."
}
