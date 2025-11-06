"""
DaVinci Resolve - Import VFX Shot Codes from Excel to Timeline Subtitles
Reads Excel file and creates SRT subtitle files based on selected columns.
"""

import sys
import argparse
from timecode import Timecode
from openpyxl import load_workbook


def create_srt_file(ws, output_path, column_index, fps):
    """Create SRT file from specific Excel column."""
    
    # Read data from Excel
    subtitles = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= column_index:
            continue
        
        content = row[column_index]
        if content and str(content).strip():
            subtitles.append({
                'tc_in': Timecode(fps, str(row[3])),    # Column D
                'tc_out': Timecode(fps, str(row[4])),   # Column E
                'text': str(content).strip()
            })
    
    if not subtitles:
        return None
    
    # Create SRT content
    srt_lines = []
    for idx, sub in enumerate(subtitles, start=1):
        # Convert to SRT format: HH:MM:SS,mmm
        def to_srt(tc):
            total_sec = tc.frames / float(tc.framerate)
            h = int(total_sec // 3600)
            m = int((total_sec % 3600) // 60)
            s = int(total_sec % 60)
            ms = int((total_sec % 1) * 1000)
            return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"
        
        srt_lines.append(str(idx))
        srt_lines.append(f"{to_srt(sub['tc_in'])} --> {to_srt(sub['tc_out'])}")
        srt_lines.append(sub['text'])
        srt_lines.append("")
    
    # Write SRT file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(srt_lines))
    
    return output_path


def main():
    parser = argparse.ArgumentParser(description='Create SRT subtitle files from Excel columns')
    parser.add_argument('--excel', default="clip_inventory.xlsx", help='Path to Excel file')
    parser.add_argument('--shot-code', action='store_true', help='Create SRT from column G')
    parser.add_argument('--vfx-work', action='store_true', help='Create SRT from column H')
    parser.add_argument('--vendor', action='store_true', help='Create SRT from column I')
    parser.add_argument('--fps', type=float, default=24.0, help='Timeline FPS (default: 24)')
    
    args = parser.parse_args()
    
    # Check at least one flag is set
    if not (args.shot_code or args.vfx_work or args.vendor):
        print("ERROR: At least one flag must be specified\n")
        print("Usage examples:")
        print(f"  python {sys.argv[0]} --shot-code")
        print(f"  python {sys.argv[0]} --shot-code --vfx-work --fps 23.976\n")
        print("Flags:")
        print("  --excel      Path to Excel file (default: clip_inventory.xlsx)")
        print("  --shot-code  Create SRT from column G (VFX Shot Code)")
        print("  --vfx-work   Create SRT from column H (VFX Work)")
        print("  --vendor     Create SRT from column I (Vendor)")
        print("  --fps        Timeline FPS (default: 24)")
        sys.exit(1)
    
    print(f"Loading: {args.excel}")
    print(f"FPS: {args.fps}\n")
    
    # Load Excel once
    wb = load_workbook(args.excel)
    ws = wb.active
    
    # Create SRT files
    base_path = args.excel.replace('.xlsx', '')
    
    if args.shot_code:
        output = f"shot_code.srt"
        if create_srt_file(ws, output, 6, args.fps):
            print(f"Created: {output}")
    
    if args.vfx_work:
        output = f"vfx_work.srt"
        if create_srt_file(ws, output, 7, args.fps):
            print(f"Created: {output}")
    
    if args.vendor:
        output = f"vendor.srt"
        if create_srt_file(ws, output, 8, args.fps):
            print(f"Created: {output}")
    
    print("\nTo import into Resolve:")
    print("1. Go to Edit page")
    print("2. Right-click on subtitle track")
    print("3. Select 'Import Subtitle...'")


if __name__ == "__main__":
    main()