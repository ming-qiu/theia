# Theia 明察秋毫: DaVinci Resolve VFX Editor Scripts

A collection of Python scripts for exporting shot information and VFX data from DaVinci Resolve timelines.

**Note: These scripts require DaVinci Resolve Studio. The free version does not include the scripting API.**

In Greek mythology, Theia was a Titan goddess of sight and shining heavenly light, daughter of Uranus and Gaea, who embodied the bright blue sky and bestowed radiance on gold, silver, and gems.

“明察秋毫”是一个汉语成语，意思是目光敏锐，能看清秋天鸟兽新长出来的细毛，比喻能洞察一切，看出极其细微的地方，出自《孟子·梁惠王上》，多含褒义，形容人精明、观察力强。

## Prep Your Resolve

Before anything please make sure the Scripting API of your DaVinci Resolve Studio works.

In Resolve, go to `Help -> Documentation -> Developer`. In the directory that pops up, go to Scripting. You will find a README.txt, which will guide you through the API setup.

_Throw the txt into any LLM to format it better._

## Quick Start (5 Minutes)

### 1. Install Dependencies

**macOS/Linux:**
```bash
chmod +x config/install_dependencies.sh
./config/install_dependencies.sh
```

**Windows:**
```cmd
config\install_dependencies.bat
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

### `shot-list.py` - Comprehensive VFX Shot List

Creates detailed shot and element breakdown with the option to compare the current edit to the last edit.

**Quick Usage:**
```bash
python3 shot-list.py --output shot_list.xlsx
```

**Advanced Usage:**
```bash
# Specify timeline and handles
python3 shot-list.py --timeline "master_0306" --work-handle 12 --scan-handle 48 --output vfx_shots.xlsx --top 7 --counter-track 9 --old-timeline "master_0205"
```

**Options:**
- `--timeline` Timeline name to process (default: current active timeline)
- `--output` Output Excel file path (default: `shot_list.xlsx`)
- `--bottom` Track number of the bottom video layer to process (default: 1)
- `--top` Track number of the top video layer to process (default: 4)
- `--counter-track` Track number containing frame counter clips (default: 5)
- `--work-handle` Work handle frames added to Cut In/Out (default: 8)
- `--scan-handle` Scan handle frames for scanning/pre-comp (default: 24)
- `--old-timeline` Timeline name of the previous edit for comparison (default: None)

**How It Works:**

1. **Shot Detection**: Reads shot codes from the first non-empty subtitle track. Each subtitle item defines one shot's time span on the timeline.

2. **Frame Counter Track**: Uses clips on the `--counter-track` to determine the Cut In and Cut Out frame numbers for each shot. The script reads the source timecode from frame counter clips that fall within each shot's subtitle span.

3. **Element Extraction**: Processes all video tracks between `--bottom` and `--top` (excluding the counter track) to extract element information:
   - Track 1: `ScanBg` (background element)
   - Track 2+: `ScanFg01`, `ScanFg02`, etc. (foreground elements)

4. **Retime Detection**: Automatically detects and documents:
   - **Constant speed retimes**: Displayed as percentage (e.g., "50%", "200%")
   - **Non-linear retimes**: Displayed as frame mappings (e.g., "1009 -> 1009, 1017 -> 1025")
   - Groups consecutive clips from the same reel with back-to-back source frames into merged elements

5. **Comparison Mode**: When `--old-timeline` is specified, compares current edit with the previous edit and calculates frame differences in the "Change to Cut" column.

**What You Get:**

**"Shots" Sheet:**
- `Sequence`: Sequence name extracted from shot code
- `Cut Order`: Sequential order of shots (1, 2, 3...)
- `Editorial Name`: Reel name from the ScanBg element (earliest background clip)
- `Shot Code`: Shot code from subtitle track (first token extracted)
- `Change to Cut`: Frame differences if comparing to old timeline (e.g., "In: +5, Out: -3")
- `Work In/Out`: Cut In/Out ± work handle frames
- `Cut In/Out`: Frame numbers in VFX frame system (from counter track)
- `Cut Duration`: Duration of the shot in frames
- `Bg Retime`: "x" if background has retime, empty otherwise
- `Fg Retime`: "x" if any foreground element has retime, empty otherwise
- `Cut In TC`: Timecode of Cut In from ScanBg element

**"Elements" Sheet:**
- `Sequence`: Sequence name
- `Cut Order`: Shot order
- `Editorial Name`: Reel name of the element
- `Shot Code`: Shot code
- `Element`: Element name (ScanBg, ScanFg01, ScanFg02, etc.)
- `Cut In/Out`: Shot boundaries in VFX frame system
- `Clip In TC`: Source timecode in point
- `Clip In Frames`: Source frame number in point
- `Clip In`: Element in point in VFX frame system
- `Clip Out`: Element out point in VFX frame system
- `Clip Out Frames`: Source frame number out point
- `Clip Out TC`: Source timecode out point
- `Clip Duration`: Duration of the clip in frames
- `Retime Summary`: Retime information (percentage or frame mappings)
- `Scale & Repo`: Scale percentage and reposition coordinates (if applicable)

**Features:**
- Automatic retime detection (constant speed and non-linear)
- Scale and reposition detection from clip properties
- Frame-accurate cut point calculation using counter track
- Element grouping for non-linear retimes
- Comparison with previous edit versions

**Requirements:**
- **Subtitle track** must contain shot codes (one subtitle item per shot defines the shot span)
- **Frame counter track** (`--counter-track`) must contain frame counter clips within each shot's span
- Video tracks between `--bottom` and `--top` should contain the elements to extract
- Shot codes in subtitles should be the first token (text before any whitespace)

## Common Workflows

### Workflow 1: Add VFX shot codes and VFX work as subtitles

Document which shots are VFX shots, give them names, and optionally document the work required.

1. **In Resolve:**
   - Consolidate all shots to video track 1

2. **Export Clips:**
   ```bash
   python3 clip-inventory.py all_shots.xlsx
   ```

3. **Review & Assign:**
   - Open `all_shots.xlsx`
   - Fill Column G and on with metadata of the shots, such as VFX shot code

4. **Generate Frame Counter:**
   ```bash
   python3 frame-counter.py --fps 25 --begin 977 --end 2000
   ```

5. **Add Frame Counters to Timeline:**
   ```bash
   python3 shot-metadata.py --sheet all_shots.xlsx --metadata-from G --metadata-to J --frame-counter ./frame-counters/frame_counter_25fps.mp4 --first-frame 1009 --fps 25
   ```

6. **In Resolve:**
   - Import the generated .srt files

### Workflow 2: Create VFX Shot List

**Goal:** Export comprehensive shot breakdown for VFX team

1. **In Resolve:**
   - Make sure there is a subtitle track that contains VFX shot codes (one subtitle item per shot)
   - If the subtitles also contain VFX work or any other text, make sure the shot code is the first part of the text
   - Ensure a frame counter track is set up with frame counter clips within each shot's span (default track 5)
   - Organize elements on video tracks: Track 1 = ScanBg, Track 2+ = ScanFg elements

2. **Export:**
   ```bash
   python3 shot-list.py --output cut_v3_shot_list.xlsx
   ```
   
   Or with custom track configuration and an older edit to compare to:
   ```bash
   python3 shot-list.py --bottom 1 --top 6 --counter-track 7 --output cut_v3_shot_list.xlsx --old-timeline "cut_v2"
   ```

3. **Review:**
   - Open Excel file
   - Check "Shots" sheet for overall shot info (cut points, handles, retime flags)
   - Check "Elements" sheet for per-clip details (timecodes, retime info, scale/repo)

## Installation Details

### Prerequisites
- **DaVinci Resolve Studio** (scripting API not available in free version)
- **Python 3.6+** installed on your system

### What Gets Installed
The installer adds these Python packages to your system Python:
- `openpyxl` - Excel file creation
- `Pillow` - Image processing for thumbnails
- `timecode` - Timecode calculations

### Manual Installation

If the automatic installer doesn't work:

```bash
# macOS/Linux
python3 -m pip install -r config/requirements.txt

# Windows
python -m pip install -r config\requirements.txt
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
- Ensure at least one subtitle track has items

### "Counter track not found or no counter clips in shot"
**Solution:**
- Ensure frame counter clips are placed on the track specified by `--counter-track` (default: 5)
- Each shot's subtitle span must contain at least one frame counter clip
- Frame counter clips should have valid source timecode
- Check that counter clips fall within the shot's subtitle item boundaries

### Installation Failed
**Solution:**
- Check internet connection (pip needs to download packages)
- Try manual installation command
- Ensure you have write permissions for Python package installation
- On macOS/Linux, you may need to use `pip3` instead of `pip`


## Technical Details

### Shot List Frame Calculations

- **Cut In/Out**: Shot boundaries in VFX frame numbering system, determined by reading source timecode from frame counter clips on the counter track
- **Work In/Out**: Cut In/Out ± work handle frames (default: 8 frames) for artist working space
- **Clip In/Out**: Element frame numbers in same VFX frame system, calculated relative to shot Cut In
- **Scan Handles**: Additional frames for scanning/pre-comp (default: 24 frames, used in HeadIn/TailOut calculations but not shown in main sheets)
- **HeadIn/TailOut**: Clip boundaries extended by scan handles for scanning purposes

### Retime Detection

The scripts detect:
- **Constant speed retimes**: Shows as percentage (e.g., "50%", "200%")
- **Non-linear retimes**: Shows frame mappings (e.g., "1009 -> 1009, 1017 -> 1025")
- **Sequential clip merging**: Groups clips from same reel with back-to-back source frames

### Scale & Repo Presentation

- **Scale**: The script detects the Zoom X value at the in point of the clip. Displayed as percentage (e.g., "Scale: 150%")
- **Repo**: Not available at this time. Stay tuned for updates
- Note: Keyframed scale/repo changes within a clip are not documented precisely at this time

### Track Naming Convention

- Track 1: `ScanBg` (background element)
- Track 2: `ScanFg01` (foreground element 1)
- Track 3: `ScanFg02` (foreground element 2)
- Tracks with "ref" in name are skipped

### Timeline Comparison

When using `--old-timeline`:
- Matches shots by Shot Code between current and old timeline
- Compares Cut In/Out frame numbers for each matched shot
- Calculates frame differences (In: ±N, Out: ±M)
- Detects retime changes between edits
- Records all differences in the "Change to Cut" column of the Shots sheet


## Tips & Best Practices

- **Save First**: Always save your Resolve project before running scripts
- **Backup**: Keep backups of important Excel files before re-exporting
- **Consistent Naming**: Use consistent shot code formats (e.g., SH010, SH020)
- **Frame Counter Track**: Ensure frame counter clips are placed on the counter track and fall within each shot's subtitle span
- **Reference Tracks**: Name reference tracks with "ref" so they're automatically skipped
- **Timeline Organization**: Keep elements on consecutive video tracks without gaps
- **Subtitle Precision**: Ensure subtitle items exactly span each shot's duration
- **Shot Code Format**: Place shot code as the first token in subtitle text (before any whitespace or additional text)
- **Track Configuration**: Use `--bottom` and `--top` to specify the range of tracks containing elements (counter track is automatically excluded)
- **Test First**: Run on a small test timeline before processing large projects


## File Structure

```
theia/
├── README.md                  # This file
├── config/                    # Configuration and installation files
│   ├── requirements.txt       # Python dependencies
│   ├── install_dependencies.sh # macOS/Linux installer
│   └── install_dependencies.bat # Windows installer
├── frame-counter.py           # Generate frame counter videos
├── clip-inventory.py          # Export clip lists with thumbnails
├── shot-metadata.py           # Generate shot metadata subtitles
├── shot-list.py               # Export comprehensive VFX shot list with elements
├── frame-counters/            # Output directory for frame counter videos
├── SF-Pro-Text-Regular.otf    # Font file for frame counter generation
```


## Support

For issues or questions:
1. Check this README's troubleshooting section
2. Verify all prerequisites (Resolve Studio running, project/timeline open)
3. Confirm dependencies are installed (re-run installer if needed)


## License & Credits

These scripts use the DaVinci Resolve API and require DaVinci Resolve Studio to be installed and running.