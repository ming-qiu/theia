"""
DaVinci Resolve - Import VFX Shot Codes from Excel to Timeline Subtitles
Reads Excel file and creates SRT subtitle file with VFX shot codes.
"""

import sys
from timecode import Timecode

try:
    from openpyxl import load_workbook
except ImportError:
    print("Error: openpyxl package required. Install with: pip install openpyxl")
    sys.exit(1)

try:
    import DaVinciResolveScript as dvr
except ImportError:
    print("Error: Could not import DaVinci Resolve API")
    sys.exit(1)


def tc_to_srt_format(tc_obj):
    """Convert Timecode object to SRT format: HH:MM:SS,mmm"""
    frames = tc_obj.frames
    fps = float(tc_obj.framerate)
    
    total_seconds = frames / fps
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    milliseconds = int((total_seconds % 1) * 1000)
    
    return f"{hours:02d}:{minutes:02d}:{seconds:02d},{milliseconds:03d}"


def import_subtitles_from_excel(excel_path):
    """Read Excel and create SRT subtitle file."""
    
    # Load Excel file
    print(f"Loading Excel file: {excel_path}")
    wb = load_workbook(excel_path)
    ws = wb.active
    
    # Connect to Resolve
    print("Connecting to DaVinci Resolve...")
    resolve = dvr.scriptapp("Resolve")
    if not resolve:
        print("Error: Could not connect to DaVinci Resolve")
        sys.exit(1)
    
    project = resolve.GetProjectManager().GetCurrentProject()
    timeline = project.GetCurrentTimeline()
    timeline_fps = float(timeline.GetSetting("timelineFrameRate"))
    
    print(f"Timeline: {timeline.GetName()}")
    print(f"FPS: {timeline_fps}")
    
    # Ensure subtitle track exists
    if timeline.GetTrackCount("subtitle") == 0:
        print("Creating subtitle track...")
        timeline.AddTrack("subtitle", None)
    
    # Read VFX shot codes from Excel
    shot_codes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 7:
            continue
        
        record_tc_in = row[3]   # Column D
        record_tc_out = row[4]  # Column E
        vfx_shot_code = row[6]  # Column G
        vfx_work = row[7] if len(row) > 7 else None  # Column H
        
        if vfx_shot_code and str(vfx_shot_code).strip():
            # Build subtitle text with shot code and optional VFX work
            subtitle_text = str(vfx_shot_code).strip()
            if vfx_work and str(vfx_work).strip():
                subtitle_text += f"\n{str(vfx_work).strip()}"
            
            shot_codes.append({
                'tc_in': str(record_tc_in),
                'tc_out': str(record_tc_out),
                'text': subtitle_text
            })
    
    print(f"Found {len(shot_codes)} VFX shot codes")
    
    if not shot_codes:
        print("No VFX shot codes found. Nothing to import.")
        return
    
    # Create SRT content
    srt_lines = []
    for idx, shot in enumerate(shot_codes, start=1):
        tc_in = Timecode(timeline_fps, shot['tc_in'])
        tc_out = Timecode(timeline_fps, shot['tc_out'])
        
        srt_lines.append(str(idx))
        srt_lines.append(f"{tc_to_srt_format(tc_in)} --> {tc_to_srt_format(tc_out)}")
        srt_lines.append(shot['text'])
        srt_lines.append("")
    
    # Write SRT file
    srt_filename = excel_path.replace('.xlsx', '_subtitles.srt')
    print(f"\nWriting SRT file: {srt_filename}")
    
    with open(srt_filename, 'w', encoding='utf-8') as f:
        f.write('\n'.join(srt_lines))
    
    print("SRT file created successfully!")
    print("\nTo import subtitles:")
    print("1. Go to Edit page in Resolve")
    print("2. Right-click on subtitle track")
    print("3. Select 'Import Subtitle...'")
    print(f"4. Choose: {srt_filename}")


if __name__ == "__main__":
    excel_file = sys.argv[1] if len(sys.argv) > 1 else "timeline_cut_points.xlsx"
    import_subtitles_from_excel(excel_file)