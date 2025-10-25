# DaVinci Resolve VFX Editor Scripts

A collection of Python scripts for exporting shot information and VFX data from DaVinci Resolve timelines.

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

The installer will automatically find your DaVinci Resolve Python and install everything needed.

### 2. Try a Clip Inventory

In Resolve, open any timeline with only 1 video track, go to the Color page in Resolve, then:

```bash
python clip-inventory.py my_inventory.xlsx
```

That's it! You'll have an inventory of clips with thumbnails.

---

## Scripts Overview

### `clip-inventory.py` - Clip Export with Thumbnails

Works with timelines with only 1 video track. Exports all clips from video track 1 with thumbnails to Excel.

**Quick Usage:**
```bash
python clip-inventory.py my_inventory.xlsx
```

**What You Get:**
- Thumbnail image for each clip
- Reel names
- Record TC In/Out (timeline positions)
- Source TC (original media timecode)
- Empty columns for VFX Shot Code and VFX Work (fill in manually)

**When to Use:**
- Need visual reference of all clips
- Creating a clip inventory for review
- Preparing to assign VFX shot codes

---

### `shot-code-vfx-work.py` - Import Shot Codes and VFX Work as Subtitles

Reads VFX shot codes from Excel and creates an SRT subtitle file for import.

**Quick Usage:**
```bash
python shot-code-vfx-work.py my_inventory.xlsx
```

**Workflow:**
1. (Export clips using `clip-inventory.py`)
2. Fill in Column G (VFX Shot Code) in the Excel file
3. Optionally fill Column H (VFX Work) with descriptions
4. Run this script to generate an SRT file
5. Import the generated SRT file into Resolve

**What It Does:**
- Reads columns D (Record TC In), E (Record TC Out), G (VFX Shot Code), H (VFX Work)
- Creates properly formatted SRT subtitle file
- Multi-line subtitles if VFX Work is provided

---

### `shot-list.py` - Comprehensive VFX Shot List

Creates detailed shot and element breakdown with optional ShotGrid sync.

**Quick Usage:**
```bash
python shot-list.py --output shot_list.xlsx
```

**Advanced Usage:**
```bash
# Specify timeline and handles
python shot-list.py --timeline "master" --work-handle 12 --scan-handle 48 --output vfx_shots.xlsx

# With ShotGrid sync and editorial change report
python shot-list.py --shotgun --output vfx_shots.xlsx
```

🫨 **Half-frame Problem and the Workaround:**
Sometimes the Resolve API can return clip start and end times that are a half frame more than the actual value, causing precision problems when calculating retimes etc.. For example, on a 25 fps timeline, when the clip starts at 21:35:14:04, the API could read a clip start time of 77516.81999999999 (seconds), while the correct number should be 77516.8. The API's reading is off by 0.5 * 1 / 25 second = 0.02 frame. This only happens to some projects, and it's unclear what leads to this problem.

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

---

## Common Workflows

### Workflow 1: Add VFX shot codes and VFX work as subtitles

**Goal:** Document which shot are VFX shots, give them names, and optionally document the work required.

1. **In Resolve:**
   - Consolidate all shots to video track 1
   - Go to the Color page

2. **Export Clips:**
   ```bash
   python clip-inventory.py all_shots.xlsx
   ```

3. **Review & Assign:**
   - Open `all_shots.xlsx`
   - Review thumbnails
   - Fill Column G with shot codes (e.g., "SH010", "SH020")
   - Fill Column H with VFX work required (optional)

4. **Import as Subtitles:**
   ```bash
   python shot-code-vfx-work.py all_shots.xlsx
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
   python shot-list.py --output project_vfx_shots.xlsx
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
   - Project name should be the same as the project's code in Shotgrid

2. **Export with ShotGrid Sync:**
   ```bash
   python shot-list.py --shotgun --output updated_cut.xlsx
   ```

3. **Review Changes:**
   - "Change to Cut" column shows In/Out frame differences
   - Positive = moved to the right on timeline, Negative = moved to the left on timeline

---

## Installation Details

### Prerequisites
- **DaVinci Resolve** (Studio or Free)
- **Python 3.6+** (comes with Resolve)

### What Gets Installed
The installer adds these Python packages to Resolve's Python:
- `openpyxl` - Excel file creation
- `Pillow` - Image processing for thumbnails
- `timecode` - Timecode calculations
- `shotgun-api3` - ShotGrid integration (optional)

### Manual Installation

If the automatic installer doesn't work:

**Find Resolve's Python:**
- macOS: `/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Frameworks/Python.framework/Versions/3.6/bin/python3`
- Windows: `C:\Program Files\Blackmagic Design\DaVinci Resolve\python.exe`
- Linux: `/opt/resolve/bin/python`

**Install dependencies:**
```bash
/path/to/resolve/python -m pip install -r requirements.txt
```

---

## Troubleshooting

### "Could not import DaVinci Resolve API"
**Solution:** 
- Make sure Resolve is running
- Use Resolve's Python, not system Python
- Check that scripts are being run from within Resolve's Python environment

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

### Missing Thumbnails (clip-inventory.py)
**Solution:** Make sure you are on the Color page in Resolve when running the script

### ShotGrid Connection Issues
**Solution:**
- Verify `.env` file exists with correct `SCRIPT_KEY`
- Check `~/config/sg_plugins/api.json` exists and contains valid credentials
- Test ShotGrid connection separately
- Ensure project code in Resolve matches ShotGrid project

### Installation Failed
**Solution:**
- Check internet connection (pip needs to download packages)
- Try manual installation command
- Ensure you have write permissions to Resolve's Python directory
- On macOS/Linux, you may need `sudo` for some Python installations

---

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
- **Repo**: Not available at this time :( Stay tuned for updates

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

---

## Tips & Best Practices

- **Save First**: Always save your Resolve project before running scripts
- **Backup**: Keep backups of important Excel files before re-exporting
- **Consistent Naming**: Use consistent shot code formats (e.g., SH010, SH020)
- **Reference Tracks**: Name reference tracks with "ref" so they're automatically skipped
- **Timeline Organization**: Keep elements on consecutive video tracks without gaps
- **Subtitle Precision**: Ensure subtitle items exactly span each shot's duration
- **Test First**: Run on a small test timeline before processing large projects

---

## File Structure

```
resolve-vfx-scripts/
├── README.md                    # This file
├── requirements.txt            # Python dependencies
├── install_dependencies.sh     # macOS/Linux installer
├── install_dependencies.bat    # Windows installer
├── .env.example               # ShotGrid config template
├── api.json.example           # ShotGrid config template
├── clip-inventory.py          # Clip export with thumbnails
├── shot-code-vfx-work.py      # Generate subtitles of VFX shot codes
└── shot-list.py               # VFX shot list export
```

---

## Support

For issues or questions:
1. Check this README's troubleshooting section
2. Verify all prerequisites (Resolve running, project/timeline open)
3. Confirm dependencies are installed (re-run installer if needed)

---

## License & Credits

These scripts use the DaVinci Resolve API and require DaVinci Resolve to be installed and running.
