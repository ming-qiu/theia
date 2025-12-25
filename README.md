# DaVinci Resolve VFX Editor Scripts

A collection of Python scripts for exporting shot information and VFX data from DaVinci Resolve timelines.

**Note: These scripts require DaVinci Resolve Studio. The free version does not include the scripting API.**

## Prep Your Resolve

Before anything please make sure the Scripting API of your DaVinci Resolve Studio works.

In Resolve, go to `Help -> Documentation -> Developer`. In the directory that pops up, go to Scripting. You will find a README.txt, which will guide you through the API setup.

_Throw the txt into any LLM to format it better._

## Quick Start (5 Minutes)

### 1. Install Dependencies

**macOS/Linux:**
```bash
chmod +x install_dependencies.sh
./install_dependencies.sh
```

**Windows:**
```cmd
install_dependencies.bat
```

The installer will use your system Python and install all required packages.

### 2. Create a Clip Inventory

In Resolve, open any timeline with clips on video track 1, then:

```bash
python3 clip-inventory.py
```

That's it! You'll have an inventory of clips with thumbnails.


## Scripts Overview

### `clip-inventory.py` - Export Clip Lists with Thumbnails

Exports the Record In, Record Out, and Source In timcodes of all clips on one track with thumbnails to an Excel spreadsheet.

**Options:**
- `--bg-track` Track number of the BG track (default: 1)
- `--file-name` Output file name (default: clip_inventory.xlsx)

**Quick Usage:**
```bash
python3 clip-inventory.py --file-name my_inventory.xlsx --bg-track 2
```

**What You Get:**
- Thumbnail image
- Reel names
- Record TC In/Out (timeline positions)
- Source TC (original media timecode)
- Empty columns for shot metadata such as VFX Shot Code, VFX Work, and Vendor (fill in manually)

**When to Use:**
- Need visual reference of all shots
- Creating a clip inventory for review
- Preparing to assign VFX shot codes

### `frame-counter.py` - Generate Frame Counter Videos

Generate an mp4 video that counts frame numbers.

**Options:**
- `--w` Width of the frame counter video (default: 160)
- `--h` Height of the frame counter video (default: 80)
- `--begin` Beginning frame number (default: 1009)
- `--end` Ending frame number (default: 2000)
- `--fps` Frame rate (default: 24)
- `--dest` Output directory of the frame counter (default: ./frame-counters)
- `--font` Path to font file (default: ./SF-Pro-Text-Regular.otf)

**Quick Usage:**
```bash
python3 frame-counter.py --begin 1001 --end 2300 --fps 25
```

**When to Use:**
- Need a frame counter for shots

### `shot-metadata.py` - Generate Shot Metadata Subtitles

Takes a clip inventory sheet, and exports all the metadata columns as subtitles.
Option to add frame counters to all the shots specified in the sheet.

**Options:**
- `--sheet` Excel sheet that contains metadata (default: clip_inventory.xlsx)
- `--metadata-from` Export metadata from which column (default: G)
- `--metadata-to` Export metadata to which column (default: G)
- `--frame-counter` Path to frame counter video file
- `--first-frame` Starting frame number for frame counters
- `--fps` Timeline FPS (default: 24)


**Quick Usage:**
```bash
python3 shot-metadata.py --excel ./shot_list.xlsx --metadata-from G --metadata-to I --frame-counter ./frame-counters/frame_counter_24fps.mp4 --first-frame 1009 --fps 24
```

**Workflow:**
1. (Export clip info using `clip-inventory.py`)
2. (Fill shot metadata into Excel file)
3. (Generate frame counters using `frame-counter.py`)
4. Run this script to add frame counters and generate SRT files
5. Import the generated SRT files into Resolve
6. Adjust the size and position of the frame counter

### `shot-list.py` - _Comprehensive VFX Shot List THIS IS UNDER CONSTRUCTION DO NOT USE_

Creates detailed shot and element breakdown with optional ShotGrid sync.

**Quick Usage:**
```bash
python3 shot-list.py --output shot_list.xlsx
```

**Advanced Usage:**
```bash
# Specify timeline and handles
python3 shot-list.py --timeline "master" --work-handle 12 --scan-handle 48 --output vfx_shots.xlsx

# With ShotGrid sync and editorial change report
python3 shot-list.py --shotgun --output vfx_shots.xlsx
```

**Half-frame Problem and the Workaround:**

Sometimes the Resolve API can return clip start and end times that are a half frame more than the actual value, causing precision problems when calculating retimes etc. For example, on a 25 fps timeline, when the clip starts at 21:35:14:04, the API could read a clip start time of 77516.81999999999 (seconds), while the correct number should be 77516.8. The API's reading is off by 0.5 * 1 / 25 second = 0.02 frame. This only happens to some projects, and it's unclear what leads to this problem.

The current workaround is the `--half-frame` flag. The script will print all clips' start times to your terminal. If those numbers are a half frame `(0.5 * 1 / fps)` off, rerun the script with the `--half-frame` flag. To check if the numbers are off, you can pick a clip in Resolve that starts at HH:MM:SS:00, and see if the corresponding clip start time in your terminal is an integer.

**Options:**
- `--timeline NAME` - Specify timeline name (default: current active timeline)
- `--sequence NAME` - Sequence name (default: extracted from timeline name before first `_`)
- `--output PATH` - Output Excel file path (default: `shot_list.xlsx`)
- `--init-cut-in N` - Initial cut-in frame number (default: 1009)
- `--work-handle N` - Work handle frames (default: 8)
- `--scan-handle N` - Scan handle frames (default: 24)
- `--half-frame` - Apply half-frame offset correction
- `--shotgun` - Query ShotGrid for previous edit info and calculate changes

**What You Get:**

**"Shots" Sheet:**
- Sequence name
- Cut order
- Editorial name (of the ScanBg element)
- Shot code
- Change to cut (if using ShotGrid)
- Work In/Out, Cut In/Out
- Cut duration
- Bg/Fg retime flags
- Cut In TC

**"Elements" Sheet:**
- All shot info
- Element name (ScanBg, ScanFg01, ScanFg02, etc.)
- Cut In/Out for the shot
- Clip In/Out TC and frame numbers
- Retime summary (speed percentages or frame mappings)
- Scale & Reposition info

**Features:**
- Automatic retime detection (constant speed and non-linear)
- Scale and reposition detection
- ShotGrid integration for tracking editorial changes

**Requirements:**
- Any reference track should be named "ref" in Resolve and will be ignored by the script
- Subtitle track must contain shot codes (one subtitle item per shot defines the shot span)

## Common Workflows

### Workflow 1: Add VFX shot codes and VFX work as subtitles

**Goal:** Document which shots are VFX shots, give them names, and optionally document the work required.

1. **In Resolve:**
   - Consolidate all shots to video track 1

2. **Export Clips:**
   ```bash
   python3 clip-inventory.py all_shots.xlsx
   ```

3. **Review & Assign:**
   - Open `all_shots.xlsx`
   - Review thumbnails
   - Fill Column G with shot codes (e.g., "SH010", "SH020")
   - Fill Column H with VFX work required (optional)

4. **Import as Subtitles:**
   ```bash
   python3 shot-metadata.py all_shots.xlsx
   ```

5. **In Resolve:**
   - Import the generated .srt

### Workflow 2: Create VFX Shot List

**Goal:** Export comprehensive shot breakdown for VFX team

1. **In Resolve:**
   - Make sure there is a subtitle track that contains VFX shot codes
   - If the subtitles also contain VFX work or any other text, make sure the shot code is the first part of the text

2. **Export:**
   ```bash
   python3 shot-list.py --output project_vfx_shots.xlsx
   ```

3. **Review:**
   - Open Excel file
   - Check "Shots" sheet for overall shot info
   - Check "Elements" sheet for per-clip details

### Workflow 3: ShotGrid-Tracked Editorial Changes

**Goal:** Compare current edit with previous ShotGrid cut

1. **Configure ShotGrid** (one-time setup):
   - Copy `.env.example` to `.env`
   - Add your `SCRIPT_KEY`
   - Copy `api.json.example` to `api.json`
   - Add your `SERVER_PATH` and `SCRIPT_USER`
   - SCRIPT_KEY can also be stored as an environment variable of your computer

2. **In Resolve:**
   - Project name should be the same as the project's code in ShotGrid

3. **Export with ShotGrid Sync:**
   ```bash
   python3 shot-list.py --shotgun --output updated_cut.xlsx
   ```

4. **Review Changes:**
   - "Change to Cut" column shows In/Out frame differences
   - Positive = moved to the right on timeline, Negative = moved to the left on timeline

## Installation Details

### Prerequisites
- **DaVinci Resolve Studio** (scripting API not available in free version)
- **Python 3.6+** installed on your system

### What Gets Installed
The installer adds these Python packages to your system Python:
- `openpyxl` - Excel file creation
- `Pillow` - Image processing for thumbnails
- `timecode` - Timecode calculations
- `shotgun-api3` - ShotGrid integration (optional)
- `python-dotenv` - Environment configuration

### Manual Installation

If the automatic installer doesn't work:

```bash
# macOS/Linux
python3 -m pip install -r requirements.txt

# Windows
python -m pip install -r requirements.txt
```

### Don't Have Python Installed?

Download and install Python from [python.org](https://www.python.org/downloads/)

**Windows users:** Make sure to check "Add Python to PATH" during installation!

## Troubleshooting

### "Could not import DaVinci Resolve API"
**Solution:** 
- Ensure DaVinci Resolve Studio is running (not the free version)
- The Resolve API only works when Resolve is open and running
- Verify the API is properly set up (see "Prep Your Resolve" section)

### "No project is currently open"
**Solution:** Open a project in DaVinci Resolve before running scripts

### "No timeline is currently open"
**Solution:** 
- Open a timeline, OR
- Use `--timeline "Timeline_Name"` to specify one

### "No subtitle items found"
**Solution:** 
- For `shot-list.py`: Add shot codes to a subtitle track
- Each subtitle item defines one shot's time span

### ShotGrid Connection Issues
**Solution:**
- Verify `.env` file exists with correct `SCRIPT_KEY`
- Check `api.json` exists and contains valid credentials
- Test ShotGrid connection separately
- Ensure project code in Resolve matches ShotGrid project

### Installation Failed
**Solution:**
- Check internet connection (pip needs to download packages)
- Try manual installation command
- Ensure you have write permissions for Python package installation
- On macOS/Linux, you may need to use `pip3` instead of `pip`


## Technical Details

### Shot List Frame Calculations

- **INIT_CUT_IN** (default 1009): The frame number assigned to the start of the first shot
- **Cut In/Out**: Shot boundaries in VFX frame numbering system
- **Work In/Out**: Cut In/Out +/- WORK_HANDLE (for artist working space)
- **Clip In/Out**: Element frame numbers in same VFX frame system
- **Scan Handles**: Additional frames for scanning/pre-comp (not shown in sheets)

### Retime Detection

The scripts detect:
- **Constant speed retimes**: Shows as percentage (e.g., "50%", "200%")
- **Non-linear retimes**: Shows frame mappings (e.g., "1009 -> 1009, 1017 -> 1025")
- **Sequential clip merging**: Groups clips from same reel with back-to-back source frames

### Scale & Repo Presentation

- **Scale**: The script detects the Zoom X value at the in point of the clip. Keyframed scale not documented precisely at this time
- **Repo**: Not available at this time. Stay tuned for updates

### Track Naming Convention

- Track 1: `ScanBg` (background element)
- Track 2: `ScanFg01` (foreground element 1)
- Track 3: `ScanFg02` (foreground element 2)
- Tracks with "ref" in name are skipped

### ShotGrid Integration

When using `--shotgun`:
- Matches shots by code and project
- Retrieves previous cut timecodes
- Calculates frame shift between old and new cut
- Applies shift to all shot/element frame numbers
- Records differences in "Change to Cut" columns


## Tips & Best Practices

- **Save First**: Always save your Resolve project before running scripts
- **Backup**: Keep backups of important Excel files before re-exporting
- **Consistent Naming**: Use consistent shot code formats (e.g., SH010, SH020)
- **Reference Tracks**: Name reference tracks with "ref" so they're automatically skipped
- **Timeline Organization**: Keep elements on consecutive video tracks without gaps
- **Subtitle Precision**: Ensure subtitle items exactly span each shot's duration
- **Test First**: Run on a small test timeline before processing large projects


## File Structure

```
resolve-vfx-scripts/
├── README.md                  # This file
├── requirements.txt           # Python dependencies
├── install_dependencies.sh    # macOS/Linux installer
├── install_dependencies.bat   # Windows installer
├── .env.example               # ShotGrid config template
├── api.json.example           # ShotGrid config template
├── frame-counter.py           # Generate frame counter videos
├── clip-inventory.py          # Export clip lists with thumbnails
├── shot-metadata.py           # Generate shot metadata subtitles
```


## Support

For issues or questions:
1. Check this README's troubleshooting section
2. Verify all prerequisites (Resolve Studio running, project/timeline open)
3. Confirm dependencies are installed (re-run installer if needed)


## License & Credits

These scripts use the DaVinci Resolve API and require DaVinci Resolve Studio to be installed and running.