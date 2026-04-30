# Theia 明察秋毫

💚 🍵 [Buy Ming a tea to show your support](https://buymeacoffee.com/ming_qiu)

A suite of VFX editorial tools for DaVinci Resolve Studio.

Theia connects to the DaVinci Resolve scripting API to export clip inventories, generate frame counter videos, and manage shot metadata — all through standalone GUIs launched directly from Resolve's script menu.

> **Requires DaVinci Resolve Studio.** The free version does not include the scripting API.

Named after the Greek Titan goddess of sight and heavenly light. 明察秋毫 — to see the finest details with keen observation.

## Tools

### Clip Inventory

Exports all visible clips on selected video tracks to an Excel spreadsheet with:

- Thumbnail images captured from the timeline
- Reel names and track numbers
- Record In/Out and Source In timecodes
- Cut order

Handles multi-track occlusion (only exports what's actually visible) and transition detection.

### Frame Counter

Generates MP4 videos of VFX frame numbers.

- Customizable dimensions (up to 3840x2160)
- Preset and custom frame rates
- Timecode metadata embedded via ffmpeg
- Configurable font

### Add Metadata

Reads a clip inventory spreadsheet and works in three optional modes — all can run together in a single pass:

- **FCPXML Titles** — exports selected metadata columns as FCPXML basic title files, one file per column, for import into Resolve as title tracks
- **SRT Subtitles** — exports the same columns as SRT subtitle files for import as subtitle tracks
- **Frame Counter** — adds a new video track to the timeline and places frame counter clips at each shot position (only for shots that have any metadata in the spreadsheet). The starting frame number is configurable. Optionally selects a VFX shot code column to rename each placed clip with its shot code — these named clips then serve as the frame counter track consumed by the Shot List tool.

Metadata columns are read from column G onwards in the spreadsheet. Record In/Out timecodes are read from columns D and E. Supports 23.976, 24, 25, 29.97, 30, and 60 fps timelines, with auto-detection from the open Resolve timeline.

### Shot List

Exports a structured VFX shot list from the current timeline to Excel, with two sheets:

- **Shots** — one row per shot with sequence, cut order, editorial name, shot code, cut in/out frames, work in/out, cut duration, and BG/FG retime flags
- **Elements** — one row per element per shot with element name, clip in/out frames, head in/tail out (with scan handles), reel name, scale/reposition summary, and retime details

Shot boundaries, shot codes, and frame numbers all come from a single **frame counter track** — clips on that track define each shot, their names are the VFX shot codes, and their source timecode carries the frame numbering.

Clip In/Out for BG elements is calculated from the frame counter source TC. Non-BG elements are calculated relative to their overlapping BG clip. Retimes are detected and summarised automatically (percentage for simple retimes, frame-mapped table for non-linear retimes across merged clips).

Optionally compare against a previous shot list Excel to flag cuts that have changed.

## Installation

### Prerequisites

- **macOS** (installer is macOS-only)
- **Python 3.9+** (Homebrew Python recommended on Apple Silicon)
- **DaVinci Resolve Studio** running with scripting API enabled
- **ffmpeg** (for frame counter timecode metadata)

### Setup

```bash
chmod +x install.command
./install.command
```

The installer will:

1. Create a virtual environment at `/Library/Application Support/Theia/venv`
2. Install Python dependencies (PySide6, openpyxl, Pillow, timecode, moviepy)
3. Copy GUI scripts to `/Library/Application Support/Theia/`
4. Install bridge scripts to Resolve's user script directory

Each user needs to run the installer once to set up their Resolve bridge scripts.

### Prep Your Resolve

Make sure the DaVinci Resolve scripting API is set up. In Resolve, go to `Help → Documentation → Developer`. In the directory that opens, find the Scripting folder and follow the instructions in README.txt.

## Usage

In DaVinci Resolve: **Workspace → Scripts → Edit → [Tool Name]**

Tools can also be run directly:

```bash
"/Library/Application Support/Theia/venv/bin/python3" \
  "/Library/Application Support/Theia/clip_inventory_gui.py"
```

## Workflows

### Export a clip inventory

1. Open a timeline in Resolve
2. Launch **Clip Inventory** from Scripts menu
3. Select video tracks to export
4. Export to Excel — you get thumbnails, timecodes, and reel names
5. Add metadata columns (VFX shot codes, work assignments, vendor) in the spreadsheet

### Generate frame counters

1. Launch **Frame Counter** from Scripts menu
2. Set dimensions, frame range, FPS, and font
3. Choose output directory
4. Generate — produces an MP4 with burned-in frame numbers

### Add metadata to timeline

1. Export a clip inventory with the **Clip Inventory** tool and fill in metadata columns (VFX shot codes, assignments, etc.) from column G onwards in the spreadsheet
2. Generate a frame counter video with the **Frame Counter** tool
3. Launch **Add Metadata** from Scripts menu
4. Select the spreadsheet and choose which metadata columns to export
5. Enable **FCPXML Titles** and/or **SRT Subtitles** and set output directories
6. Enable **Frame Counter**, select the generated video, set the starting frame number, and choose the VFX shot code column to name each clip
7. Click Go — SRT/FCPXML files are written to disk and frame counter clips are placed on a new timeline track, each named with its VFX shot code
8. In Resolve, import the FCPXML or SRT files as needed

### Export a VFX shot list

1. Ensure your timeline has a frame counter track with clips named by VFX shot code
2. Launch **Shot List** from Scripts menu
3. Select the frame counter track, element track range, and handle sizes
4. Optionally load a previous shot list Excel to diff against
5. Export — produces an Excel with Shots and Elements sheets

## Project Structure

```
theia/
├── install.command          # macOS installer
├── uninstall.command        # macOS uninstaller
├── bridges/                 # DaVinci Resolve bridge scripts
│   ├── 01 Clip Inventory.py
│   ├── 02 Frame Counter.py
│   ├── 03 Add Metadata.py
│   └── 04 Shot List.py
├── scripts/                 # GUI applications
│   ├── add_metadata_gui.py
│   ├── clip_inventory_gui.py
│   ├── frame_counter_gui.py
│   └── shot_list_gui.py
└── resources/
    ├── fonts/
    │   └── SF-Pro-Text-Regular.otf
    └── graphics/
        ├── add_metadata_icon.png
        ├── clip_inventory_icon.png
        └── frame_counter_icon.png
```

## Troubleshooting

**"Could not import DaVinci Resolve API"** — Make sure Resolve Studio is running (not the free version) and the scripting API is set up.

**"No project/timeline is currently open"** — Open a project and timeline in Resolve before launching scripts.

**Architecture issues on Apple Silicon** — The installer detects ARM64 vs x86_64 and handles Rosetta compatibility automatically. If you have issues, try reinstalling with Homebrew Python (`brew install python`).
