"""
Theia - Shot List GUI
Export VFX shot list with elements from a DaVinci Resolve timeline to Excel.
"""

import os
import sys
import math
import traceback
from pathlib import Path
from collections import defaultdict

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QFileDialog,
    QMessageBox, QProgressBar, QTextEdit, QGroupBox, QSpinBox
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from timecode import Timecode

# Import DaVinci Resolve API
try:
    import DaVinciResolveScript as dvr
except ImportError:
    resolve_script_api = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules"
    if resolve_script_api not in sys.path:
        sys.path.append(resolve_script_api)
    try:
        import DaVinciResolveScript as dvr
    except ImportError:
        dvr = None

# -------- Utils --------

def get_timeline_fps(timeline):
    fps_str = timeline.GetSetting("timelineFrameRate") or ""
    try:
        fps = float(fps_str)
        return fps
    except Exception:
        return 24.0

def fps_to_str(fps):
    if float(fps).is_integer():
        return str(int(round(fps)))
    return f"{fps:.3f}".rstrip('0').rstrip('.')

def safe_get(d, k, default=None):
    try:
        return d.get(k, default)
    except Exception:
        return default

def title_clip_text(item):
    """Extract text from a Basic Title / Fusion Title clip on a video track."""
    props = item.GetProperty() or {}
    text = safe_get(props, "Text", "")
    if not text:
        text = item.GetName() or ""
    return text.strip()

def element_name_for_track(idx):
    if idx == 1:
        return "ScanBg"
    return f"ScanFg{idx-1:02d}"

def get_reel_name(clip_props):
    for k in ("Clip Name", "File Name", "Reel Name"):
        v = clip_props.get(k)
        if v:
            return v
    return "placeholder"

def get_sequence_name(shot_code):
    if shot_code.find('_') != -1:
        return shot_code.split('_')[0]
    elif shot_code.find('-') != -1:
        return shot_code.split('-')[0]
    else:
        return 'sequence_name'

def shot_editorial_name_from_bg(bg_elements):
    """Pick the ScanBg reel name for the shot (earliest bg element)."""
    if not bg_elements:
        return ""
    e = sorted(bg_elements, key=lambda x: (x["TimelineStart"], x["TimelineEnd"]))[0]
    return e.get("ReelName", "") or ""

def best_bg_cut_in_tc(bg_elements, fps):
    """Choose Cut In TC from ScanBg (track 1). Prefer the element that starts closest to the shot start."""
    if not bg_elements:
        return ""
    e = sorted(bg_elements, key=lambda x: x["ClipInFrames"])[0]
    return e["ClipInTC"]

def get_clip_tc(timeline_item, fps):
    """
    Works for online/offline items. Assumes clip fps == timeline fps.
    Returns: dict with ClipInTC, ClipInFrames, ClipOutTC, ClipOutFrames
    """
    tin = timeline_item.GetSourceStartTime()
    tout = timeline_item.GetSourceEndTime()

    if tin is not None and tout is not None:
        source_dur_seconds = tout - tin
        source_dur_frames = source_dur_seconds * fps

        timeline_dur = timeline_item.GetDuration(False) or 0

        speed = (source_dur_frames - 1) / (timeline_dur - 1) if timeline_dur > 0 else 1.0

        if abs(speed - 1) < 1e-3:
            speed = 1

        actual_source_frames = timeline_dur * speed

        src_in_frames = int(round(tin * fps))
        src_out_frames = int(round(src_in_frames + actual_source_frames))

    elif timeline_item.GetDuration(False):
        dur = timeline_item.GetDuration(False)
        src_in_frames = 0
        src_out_frames = int(dur)
    else:
        src_in_frames = 0
        src_out_frames = 1

    # Timecode lib requires frames >= 1; clamp for placeholders / clips with no real source TC
    tc_in_frames = max(1, src_in_frames)
    tc_out_frames = max(1, src_out_frames)

    return {
        "ClipInTC":      repr(Timecode(fps_to_str(fps), frames=tc_in_frames)),
        "ClipInFrames":  src_in_frames,
        "ClipOutTC":     repr(Timecode(fps_to_str(fps), frames=tc_out_frames)),
        "ClipOutFrames": src_out_frames,
    }

def _is_back_to_back(prev_src_out, curr_src_in, tol=1):
    return curr_src_in == prev_src_out or curr_src_in == prev_src_out + 1 or abs(curr_src_in - prev_src_out) <= tol

def _fmt_percent(val):
    f = float(str(val).strip())
    p = f * 100.0
    return f"{int(round(p))}%" if abs(p - round(p)) < 1e-6 else f"{p:.2f}%"

def retime_summary(elements_by_track, fps):
    """
    For each track's elements (within a shot), detect non-linear retimes:
    - Group consecutive clips by same ReelName when their source frames are back-to-back
    - Merge grouped clips into a single element representing the non-linear retime
    - For merged elements, calculate frame mappings showing timeline_frame -> source_frame
    - Write RetimeSummary as: "x1 -> y1, x2 -> y2, ..."
    - Sets HasRetime (True if any segment speed != 1 or non-linear retime detected)
    - MODIFIES elements_by_track in place, replacing groups with merged elements
    """
    for track_num, track in elements_by_track.items():
        track.sort(key=lambda e: (e["TimelineStart"], e["TimelineEnd"]))

        for clip in track:
            ti = clip["TimelineItem"]
            src_frames = max(0, int(clip["ClipOutFrames"] - clip["ClipInFrames"]))
            tl_frames = int(ti.GetDuration(False) or (clip["TimelineEnd"] - clip["TimelineStart"]))

            speed = (src_frames / tl_frames) if tl_frames > 0 else None

            clip["SourceDur"] = src_frames
            clip["TimelineDur"] = tl_frames
            clip["Speed"] = speed
            clip["RetimeFPS"] = (fps * speed) if speed is not None else None
            clip["RetimeSummary"] = ""
            clip["HasRetime"] = (speed is not None) and (abs(speed - 1.0) > 1e-3)

        merged_track = []
        i = 0
        n = len(track)

        while i < n:
            group = [track[i]]
            reel = track[i].get("ReelName", "")
            j = i + 1

            while j < n:
                same_reel = (track[j].get("ReelName", "") == reel)
                if not same_reel:
                    break
                if not _is_back_to_back(group[-1]["ClipOutFrames"], track[j]["ClipInFrames"]):
                    break
                group.append(track[j])
                j += 1

            any_retime = any(g["HasRetime"] for g in group)

            if any_retime and len(group) > 1:
                first = group[0]
                last = group[-1]

                mappings = []
                current_timeline_frame = first["ClipIn"]
                current_source_frame = first["ClipIn"]

                mappings.append(f"{current_timeline_frame} -> {current_source_frame}")
                current_source_frame -= 1
                current_timeline_frame -= 1

                for g in group:
                    timeline_end = current_timeline_frame + g["TimelineDur"]
                    source_end = current_source_frame + g["SourceDur"]
                    mappings.append(f"{timeline_end} -> {source_end}")
                    current_timeline_frame = timeline_end
                    current_source_frame = source_end

                summary = ", ".join(mappings)

                merged_element = {
                    "TrackIndex": first["TrackIndex"],
                    "ShotCode": first["ShotCode"],
                    "ElementName": first["ElementName"],
                    "TimelineItem": first["TimelineItem"],
                    "TimelineStart": first["TimelineStart"],
                    "TimelineEnd": last["TimelineEnd"],
                    "ClipIn": first["ClipIn"],
                    "ClipOut": last["ClipOut"],
                    "ClipInTC": first["ClipInTC"],
                    "ClipOutTC": last["ClipOutTC"],
                    "ClipInFrames": first["ClipInFrames"],
                    "ClipOutFrames": last["ClipOutFrames"],
                    "RetimeSummary": summary,
                    "ScaleRepo": first["ScaleRepo"],
                    "HasRetime": True,
                    "ReelName": reel,
                    "Props": first["Props"],
                    "ClipProps": first["ClipProps"],
                    "HeadIn": first["HeadIn"],
                    "TailOut": last["TailOut"],
                    "SourceDur": sum(g["SourceDur"] for g in group) - 1,
                    "TimelineDur": sum(g["TimelineDur"] for g in group) - 1,
                    "Speed": None,
                    "RetimeFPS": None,
                }

                merged_track.append(merged_element)

            elif any_retime:
                for g in group:
                    if g["HasRetime"]:
                        g["RetimeSummary"] = _fmt_percent(g['Speed'])
                    merged_track.append(g)
            else:
                merged_track.extend(group)

            i = j

        elements_by_track[track_num] = merged_track

def summarize_scale_repo(props):
    zx = safe_get(props, "ZoomX")
    px = safe_get(props, "PositionX")
    py = safe_get(props, "PositionY")
    parts = []
    if zx != 1:
        parts.append(f"Scale: {_fmt_percent(zx)}")
    if px is not None or py is not None:
        parts.append(f"Repo: {px},{py}")
    return " ".join(parts)

# -------- Old Excel loading --------

def load_old_shot_list_excel(excel_path, fps):
    """
    Read the Shots sheet from an old shot list Excel.
    Returns dict mapping ShotCode -> {CutIn, CutOut, CutInTC, CutInTCFrames, ...}

    Columns: A=Sequence, B=Cut Order, C=Editorial Name, D=Shot Code,
    E=Change to Cut, F=Work In, G=Cut In, H=Cut Out, I=Work Out,
    J=Cut Duration, K=Bg Retime, L=Fg Retime, M=Cut In TC
    """
    wb = load_workbook(excel_path, read_only=True)
    try:
        ws = wb["Shots"]
    except KeyError:
        ws = wb.active

    old_shots = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 13:
            continue
        shot_code = row[3]  # Column D
        if not shot_code or not str(shot_code).strip():
            continue

        cut_in = int(row[6]) if row[6] is not None else None
        cut_out = int(row[7]) if row[7] is not None else None
        cut_in_tc = str(row[12]).strip() if row[12] else ""

        cut_in_tc_frames = None
        if cut_in_tc:
            try:
                cut_in_tc_frames = Timecode(fps_to_str(fps), cut_in_tc).frames
            except Exception:
                pass

        old_shots[str(shot_code).strip()] = {
            'CutIn': cut_in,
            'CutOut': cut_out,
            'CutInTC': cut_in_tc,
            'CutInTCFrames': cut_in_tc_frames,
        }

    wb.close()
    return old_shots

def compare_with_old_excel(current_shots, old_shots_dict):
    """Compare current shots against old Excel data. Modifies current_shots in place."""
    for shot in current_shots:
        code = shot['ShotCode']
        old = old_shots_dict.get(code)
        if not old or old['CutIn'] is None or old['CutOut'] is None:
            continue

        d_in = shot['CutIn'] - old['CutIn']
        d_out = shot['CutOut'] - old['CutOut']
        change_to_cut = ""
        if d_in != 0 or d_out != 0:
            change_to_cut = (
                f"In: {d_in}, Out: {d_out}" if (d_in != 0 and d_out != 0)
                else f"In: {d_in}" if d_in != 0
                else f"Out: {d_out}"
            )
        shot['ChangeToCut'] = change_to_cut

# -------- Worker --------

class ShotListWorker(QThread):
    """Threaded worker for shot list extraction and Excel export."""
    progress = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, shot_code_track, frame_counter_track, bg_track,
                 bottom_track, top_track, cut_in_frame, old_excel_path,
                 work_handle, scan_handle, output_path, input_sequence, fps):
        super().__init__()
        self.shot_code_track = shot_code_track
        self.frame_counter_track = frame_counter_track
        self.bg_track = bg_track
        self.bottom_track = bottom_track
        self.top_track = top_track
        self.cut_in_frame = cut_in_frame
        self.old_excel_path = old_excel_path
        self.work_handle = work_handle
        self.scan_handle = scan_handle
        self.output_path = output_path
        self.input_sequence = input_sequence
        self.fps = fps

    def log(self, msg):
        self.progress.emit(msg)

    def run(self):
        try:
            self.log("Connecting to DaVinci Resolve...")
            if dvr is None:
                self.finished.emit(False, "DaVinci Resolve API not available")
                return

            resolve = dvr.scriptapp("Resolve")
            if not resolve:
                self.finished.emit(False, "Could not connect to DaVinci Resolve")
                return

            project = resolve.GetProjectManager().GetCurrentProject()
            if not project:
                self.finished.emit(False, "No project open in Resolve")
                return

            timeline = project.GetCurrentTimeline()
            if not timeline:
                self.finished.emit(False, "No timeline open in Resolve")
                return

            fps = self.fps
            self.log(f"Timeline: {timeline.GetName()} | FPS: {fps}")

            # Load old shot list Excel if provided
            old_shots_dict = None
            if self.old_excel_path:
                self.log(f"Loading old shot list: {os.path.basename(self.old_excel_path)}")
                old_shots_dict = load_old_shot_list_excel(self.old_excel_path, fps)
                self.log(f"  Found {len(old_shots_dict)} shots in old Excel")

            # Get shot code items from title clips on video track
            shot_items = timeline.GetItemListInTrack("video", self.shot_code_track) or []
            if not shot_items:
                self.finished.emit(False,
                    f"No clips found on Shot Code Track {self.shot_code_track}")
                return
            shot_items_sorted = sorted(shot_items, key=lambda c: c.GetStart(False))
            self.log(f"Found {len(shot_items_sorted)} shot code clips on track {self.shot_code_track}")

            # Get frame counter clips if Mode A
            counters = None
            if self.frame_counter_track is not None:
                counters = timeline.GetItemListInTrack("video", self.frame_counter_track) or []
                self.log(f"Frame counter track {self.frame_counter_track}: {len(counters)} clips")

            # Pre-pull video items for element tracks
            v_tracks = timeline.GetTrackCount("video")
            element_labels = {i: element_name_for_track(i) for i in range(1, v_tracks + 1)}

            skip_tracks = {self.shot_code_track}
            if self.frame_counter_track is not None:
                skip_tracks.add(self.frame_counter_track)

            track_items = {}
            for i in range(self.bottom_track, self.top_track + 1):
                if i not in skip_tracks:
                    track_items[i] = timeline.GetItemListInTrack("video", i) or []

            # Also pre-pull BG track items for Mode B if needed
            bg_items_all = None
            if self.frame_counter_track is None and self.bg_track is not None:
                bg_items_all = timeline.GetItemListInTrack("video", self.bg_track) or []

            # Process each shot
            shots_rows = []
            elements_rows = []
            cut_order = 0

            for shot_item in shot_items_sorted:
                shot_code = title_clip_text(shot_item)
                if not shot_code:
                    continue

                cut_order += 1
                shot_start = shot_item.GetStart(False)
                shot_end = shot_item.GetEnd(False)
                shot_dur = shot_end - shot_start

                self.log(f"==== Cut {cut_order}: {shot_code} [{shot_start}-{shot_end}] ====")

                # Determine cut_in and cut_out
                cut_in = None
                cut_out = None

                if counters is not None:
                    # MODE A: Frame counter track
                    for counter in counters:
                        cs = counter.GetStart(False)
                        ce = counter.GetEnd(False)
                        if cs >= shot_start and ce <= shot_end:
                            tc_info = get_clip_tc(counter, fps)
                            cut_in = tc_info['ClipInFrames']
                            cut_out = tc_info['ClipOutFrames'] - 1
                            self.log(f"  Counter: Cut In={cut_in}, Cut Out={cut_out}")
                            break
                    if cut_in is None:
                        self.log(f"  WARNING: No counter clip found, using default {self.cut_in_frame}")
                        cut_in = self.cut_in_frame
                        cut_out = cut_in + shot_dur - 1
                else:
                    # MODE B: BG track fallback
                    bg_clips_in_shot = []
                    for clip in (bg_items_all or []):
                        cs = clip.GetStart(False)
                        ce = clip.GetEnd(False)
                        if cs >= shot_start and ce <= shot_end:
                            bg_clips_in_shot.append(clip)

                    if bg_clips_in_shot:
                        earliest_bg = min(bg_clips_in_shot, key=lambda c: c.GetStart(False))
                        bg_tc_info = get_clip_tc(earliest_bg, fps)
                        current_bg_source_frames = bg_tc_info['ClipInFrames']

                        if (old_shots_dict
                                and shot_code in old_shots_dict
                                and old_shots_dict[shot_code]['CutIn'] is not None
                                and old_shots_dict[shot_code]['CutInTCFrames'] is not None):
                            old = old_shots_dict[shot_code]
                            cut_in = old['CutIn'] + (current_bg_source_frames - old['CutInTCFrames'])
                            cut_out = cut_in + shot_dur - 1
                            self.log(f"  BG mapped from old: Cut In={cut_in}, Cut Out={cut_out}")
                        else:
                            cut_in = self.cut_in_frame
                            cut_out = cut_in + shot_dur - 1
                            self.log(f"  BG default: Cut In={cut_in}, Cut Out={cut_out}")
                    else:
                        cut_in = self.cut_in_frame
                        cut_out = cut_in + shot_dur - 1
                        self.log(f"  WARNING: No BG clips found, using default {self.cut_in_frame}")

                # Collect elements on [bottom..top] tracks
                elements_by_track = defaultdict(list)

                for track in range(self.bottom_track, self.top_track + 1):
                    if track in skip_tracks:
                        continue
                    for clip in (track_items.get(track) or []):
                        clip_start = clip.GetStart(False)
                        clip_end = clip.GetEnd(False)
                        if clip_start >= shot_start and clip_end <= shot_end:
                            mpi = clip.GetMediaPoolItem()
                            clip_props = (mpi.GetClipProperty() if mpi else {}) or {}

                            reel = clip.GetName()

                            tc_info = get_clip_tc(clip, fps)

                            clip_in_tc = tc_info["ClipInTC"]
                            clip_out_tc = tc_info["ClipOutTC"]
                            clip_in_tc_frames = tc_info["ClipInFrames"]
                            clip_out_tc_frames = tc_info["ClipOutFrames"]

                            clip_in = int(cut_in + (clip_start - shot_start))
                            clip_out = int(clip_in + (clip_out_tc_frames - clip_in_tc_frames)) - 1

                            props = clip.GetProperty() or {}
                            scalerpo_sum = summarize_scale_repo(props)

                            elements_by_track[track].append({
                                "TrackIndex": track,
                                "ShotCode": shot_code,
                                "ElementName": element_labels[track],
                                "TimelineItem": clip,
                                "TimelineStart": clip_start,
                                "TimelineEnd": clip_end,
                                "ClipIn": clip_in,
                                "ClipOut": clip_out,
                                "ClipInTC": clip_in_tc,
                                "ClipOutTC": clip_out_tc,
                                "ClipInFrames": clip_in_tc_frames,
                                "ClipOutFrames": clip_out_tc_frames,
                                "ClipDuration": clip_out - clip_in + 1,
                                "HasRetime": False,
                                "RetimeSummary": "",
                                "ScaleRepo": scalerpo_sum,
                                "ReelName": reel,
                                "Props": props,
                                "ClipProps": clip_props,
                                "HeadIn": int(clip_in - self.scan_handle),
                                "TailOut": int((clip_out_tc_frames - clip_in_tc_frames) + clip_in + self.scan_handle),
                            })

                # Shot metadata from BG elements
                bg_elems = elements_by_track.get(self.bottom_track, [])
                cut_in_tc = best_bg_cut_in_tc(bg_elems, fps) if bg_elems else ""
                shot_editorial_name = shot_editorial_name_from_bg(bg_elems)

                work_in = int(cut_in - self.work_handle)
                work_out = int(cut_out + self.work_handle)

                retime_summary(elements_by_track, fps)

                bg_retime = "x" if any(e["HasRetime"] for e in elements_by_track.get(self.bottom_track, [])) else ""
                fg_retime = "x" if any(
                    e["HasRetime"]
                    for track, lst in elements_by_track.items()
                    if track > self.bottom_track for e in lst
                ) else ""

                # Determine sequence name by input and shot code
                if self.input_sequence == None:
                    sequence = get_sequence_name(shot_code)
                else:
                    sequence = self.input_sequence

                shots_rows.append({
                    "Sequence": sequence,
                    "CutOrder": cut_order,
                    "EditorialName": shot_editorial_name,
                    "ShotCode": shot_code,
                    "ChangeToCut": None,
                    "WorkIn": work_in,
                    "CutIn": int(cut_in),
                    "CutOut": int(cut_out),
                    "WorkOut": work_out,
                    "CutDuration": shot_dur,
                    "BgRetime": bg_retime,
                    "FgRetime": fg_retime,
                    "CutInTC": cut_in_tc,
                })

                for track in range(1, v_tracks + 1):
                    for e in sorted(elements_by_track.get(track, []), key=lambda x: (x["TimelineStart"], x["TimelineEnd"])):
                        elements_rows.append({
                            "Sequence": sequence,
                            "CutOrder": cut_order,
                            "EditorialName": e["ReelName"],
                            "ShotCode": shot_code,
                            "Element": e["ElementName"],
                            "ShotCutIn": int(cut_in),
                            "ShotCutOut": int(cut_out),
                            "ClipInTC": e["ClipInTC"],
                            "ClipInFrames": int(e["ClipInFrames"]),
                            "ClipIn": int(e["ClipIn"]),
                            "ClipOut": int(e["ClipOut"]),
                            "ClipOutFrames": int(e["ClipOutFrames"]),
                            "ClipOutTC": e["ClipOutTC"],
                            "ClipDuration": e['ClipDuration'],
                            "Retime": e["RetimeSummary"],
                            "ScaleRepo": e["ScaleRepo"],
                        })

            # Sort
            shots_rows.sort(key=lambda r: r["CutOrder"])
            elements_rows.sort(key=lambda r: r["CutOrder"])

            # Compare with old Excel
            if old_shots_dict:
                compare_with_old_excel(shots_rows, old_shots_dict)

            # -------- Excel output --------
            self.log("")
            self.log("Writing Excel...")

            wb = Workbook()
            ws_shots = wb.active
            ws_shots.title = "Shots"
            shots_cols = [
                "Sequence", "Cut Order", "Editorial Name", "Shot Code", "Change to Cut",
                "Work In", "Cut In", "Cut Out", "Work Out",
                "Cut Duration", "Bg Retime", "Fg Retime", "Cut In TC"
            ]
            ws_shots.append(shots_cols)
            for r in shots_rows:
                ws_shots.append([
                    r["Sequence"], r["CutOrder"], r["EditorialName"], r["ShotCode"], r["ChangeToCut"],
                    r["WorkIn"], r["CutIn"], r["CutOut"], r["WorkOut"],
                    r["CutDuration"], r["BgRetime"], r["FgRetime"], r["CutInTC"]
                ])

            ws_elems = wb.create_sheet(title="Elements")
            elems_cols = [
                "Sequence", "Cut Order", "Editorial Name", "Shot Code", "Element",
                "Cut In", "Cut Out", "Clip Duration",
                "Clip In TC", "Clip In Frames", "Clip In", "Clip Out", "Clip Out Frames", "Clip Out TC",
                "Retime Summary", "Scale & Repo"
            ]
            ws_elems.append(elems_cols)
            for r in elements_rows:
                ws_elems.append([
                    r["Sequence"], r["CutOrder"], r["EditorialName"], r["ShotCode"], r["Element"],
                    r["ShotCutIn"], r["ShotCutOut"], r['ClipDuration'],
                    r["ClipInTC"], r["ClipInFrames"], r["ClipIn"], r["ClipOut"], r["ClipOutFrames"], r["ClipOutTC"],
                    r["Retime"], r["ScaleRepo"]
                ])

            # Auto-width
            for ws in (ws_shots, ws_elems):
                for col_idx, _ in enumerate(ws[1], start=1):
                    max_len = 0
                    for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                        val = row[0].value
                        if val is None:
                            continue
                        max_len = max(max_len, len(str(val)))
                    ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 50)

            out_path = os.path.abspath(self.output_path)
            wb.save(out_path)
            self.log(f"✓ Wrote Excel: {out_path}")
            self.log(f"  {len(shots_rows)} shots, {len(elements_rows)} elements")
            self.finished.emit(True, f"Exported {len(shots_rows)} shots to {os.path.basename(out_path)}")

        except Exception as e:
            self.log(f"ERROR: {e}")
            self.log(traceback.format_exc())
            self.finished.emit(False, str(e))

# -------- GUI --------

class ShotListGUI(QMainWindow):
    """Main GUI window for shot list export."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.setup_ui()
        self.populate_tracks()

    def setup_ui(self):
        self.setWindowTitle("Theia - Shot List")
        self.setMinimumWidth(750)
        self.setMinimumHeight(700)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Title
        title = QLabel("Shot List")
        font = QFont()
        font.setPointSize(16)
        font.setBold(True)
        title.setFont(font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        layout.addSpacing(5)

        # ---- Track Configuration ----
        track_group = QGroupBox("Track Configuration")
        track_layout = QVBoxLayout()

        # Shot Code Track
        sc_row = QHBoxLayout()
        sc_row.addWidget(QLabel("Shot Code Track:"))
        self.shot_code_track_combo = QComboBox()
        self.shot_code_track_combo.setMinimumWidth(200)
        sc_row.addWidget(self.shot_code_track_combo)
        sc_row.addStretch()
        track_layout.addLayout(sc_row)

        # Frame Counter Track
        fc_row = QHBoxLayout()
        fc_row.addWidget(QLabel("Frame Counter Track:"))
        self.counter_track_combo = QComboBox()
        self.counter_track_combo.setMinimumWidth(200)
        self.counter_track_combo.currentTextChanged.connect(self.on_counter_track_changed)
        fc_row.addWidget(self.counter_track_combo)
        fc_row.addStretch()
        track_layout.addLayout(fc_row)

        # BG Track (visible only when counter = None)
        bg_row = QHBoxLayout()
        self.bg_track_label = QLabel("BG Track:")
        bg_row.addWidget(self.bg_track_label)
        self.bg_track_combo = QComboBox()
        self.bg_track_combo.setMinimumWidth(200)
        bg_row.addWidget(self.bg_track_combo)
        bg_row.addStretch()
        track_layout.addLayout(bg_row)
        self.bg_track_label.hide()
        self.bg_track_combo.hide()

        # Bottom / Top tracks
        bt_row = QHBoxLayout()
        bt_row.addWidget(QLabel("Bottom Track:"))
        self.bottom_spin = QSpinBox()
        self.bottom_spin.setMinimum(1)
        self.bottom_spin.setMaximum(99)
        self.bottom_spin.setValue(1)
        bt_row.addWidget(self.bottom_spin)
        bt_row.addSpacing(20)
        bt_row.addWidget(QLabel("Top Track:"))
        self.top_spin = QSpinBox()
        self.top_spin.setMinimum(1)
        self.top_spin.setMaximum(99)
        self.top_spin.setValue(4)
        bt_row.addWidget(self.top_spin)
        bt_row.addStretch()

        refresh_btn = QPushButton("↻")
        refresh_btn.setMaximumWidth(40)
        refresh_btn.setToolTip("Refresh track list from Resolve")
        refresh_btn.clicked.connect(self.populate_tracks)
        bt_row.addWidget(refresh_btn)

        track_layout.addLayout(bt_row)
        track_group.setLayout(track_layout)
        layout.addWidget(track_group)

        # ---- Sequence Name Setting ----
        seq_group = QGroupBox("Sequence Name")
        seq_layout = QVBoxLayout()

        seq_row = QHBoxLayout()
        seq_row.addWidget(QLabel("Sequence Name:"))
        self.seq_name = QLineEdit("")
        self.seq_name.setPlaceholderText("Optional")
        seq_row.addWidget(self.seq_name)
        seq_row.addStretch()
        seq_layout.addLayout(seq_row)
        seq_group.setLayout(seq_layout)

        layout.addWidget(seq_group)

        # ---- Frame Number Settings ----
        frame_group = QGroupBox("Frame Number Settings")
        frame_layout = QVBoxLayout()

        ci_row = QHBoxLayout()
        ci_row.addWidget(QLabel("Cut In Frame Number:"))
        self.cut_in_spin = QSpinBox()
        self.cut_in_spin.setMinimum(0)
        self.cut_in_spin.setMaximum(9999)
        self.cut_in_spin.setValue(1009)
        ci_row.addWidget(self.cut_in_spin)
        ci_row.addStretch()
        frame_layout.addLayout(ci_row)

        old_row = QHBoxLayout()
        old_row.addWidget(QLabel("Old Shot List:"))
        self.old_excel_input = QLineEdit("")
        self.old_excel_input.setPlaceholderText("Optional - previous shot list .xlsx for comparison & frame mapping")
        old_row.addWidget(self.old_excel_input)
        browse_old_btn = QPushButton("Browse...")
        browse_old_btn.setMaximumWidth(100)
        browse_old_btn.clicked.connect(self.browse_old_excel)
        old_row.addWidget(browse_old_btn)
        frame_layout.addLayout(old_row)

        frame_group.setLayout(frame_layout)
        layout.addWidget(frame_group)

        # ---- Handles ----
        handle_group = QGroupBox("Handles")
        handle_layout = QHBoxLayout()

        handle_layout.addWidget(QLabel("Work Handle:"))
        self.work_handle_spin = QSpinBox()
        self.work_handle_spin.setMinimum(0)
        self.work_handle_spin.setMaximum(999)
        self.work_handle_spin.setValue(8)
        handle_layout.addWidget(self.work_handle_spin)

        handle_layout.addSpacing(20)

        handle_layout.addWidget(QLabel("Scan Handle:"))
        self.scan_handle_spin = QSpinBox()
        self.scan_handle_spin.setMinimum(0)
        self.scan_handle_spin.setMaximum(999)
        self.scan_handle_spin.setValue(24)
        handle_layout.addWidget(self.scan_handle_spin)

        handle_layout.addStretch()
        handle_group.setLayout(handle_layout)
        layout.addWidget(handle_group)

        # ---- Output ----
        output_group = QGroupBox("Output")
        output_layout = QHBoxLayout()

        output_layout.addWidget(QLabel("Output File:"))
        self.output_input = QLineEdit(str(Path.home() / "Downloads" / "shot_list.xlsx"))
        output_layout.addWidget(self.output_input)
        browse_output_btn = QPushButton("Browse...")
        browse_output_btn.setMaximumWidth(100)
        browse_output_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(browse_output_btn)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # ---- Frame Rate ----
        fps_group = QGroupBox("Frame Rate")
        fps_layout = QHBoxLayout()

        fps_layout.addWidget(QLabel("Timeline FPS:"))
        self.fps_combo = QComboBox()
        self.fps_combo.setMaximumWidth(120)
        self.fps_combo.addItems(["23.976", "24", "25", "30", "60", "Custom..."])
        self.fps_combo.setCurrentText("24")
        self.fps_combo.currentTextChanged.connect(self.on_fps_changed)
        fps_layout.addWidget(self.fps_combo)

        self.custom_fps_input = QLineEdit()
        self.custom_fps_input.setMaximumWidth(100)
        self.custom_fps_input.setPlaceholderText("Enter FPS")
        self.custom_fps_input.hide()
        self.custom_fps_input.textChanged.connect(self.validate_custom_fps)
        fps_layout.addWidget(self.custom_fps_input)

        fps_layout.addStretch()
        fps_group.setLayout(fps_layout)
        layout.addWidget(fps_group)

        # ---- Go button ----
        self.go_btn = QPushButton("Go")
        self.go_btn.setMinimumHeight(40)
        self.go_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.go_btn)

        # Progress bar
        self.progress = QProgressBar()
        self.progress.hide()
        layout.addWidget(self.progress)

        # Log
        layout.addWidget(QLabel("Log:"))
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

    def populate_tracks(self):
        """Populate track dropdowns from the current Resolve timeline."""
        self.shot_code_track_combo.clear()
        self.counter_track_combo.clear()
        self.bg_track_combo.clear()

        self.counter_track_combo.addItem("None", None)

        if dvr is None:
            self.log.append("⚠️  DaVinci Resolve API not available")
            return

        try:
            resolve = dvr.scriptapp("Resolve")
            if not resolve:
                self.log.append("⚠️  Could not connect to Resolve")
                return

            project = resolve.GetProjectManager().GetCurrentProject()
            if not project:
                self.log.append("⚠️  No project open")
                return

            timeline = project.GetCurrentTimeline()
            if not timeline:
                self.log.append("⚠️  No timeline open")
                return

            track_count = timeline.GetTrackCount("video")
            self.log.append(f"✓ Timeline: {timeline.GetName()} ({track_count} video tracks)")

            for i in range(1, track_count + 1):
                name = timeline.GetTrackName("video", i) or ""
                label = f"Track {i}" + (f" ({name})" if name else "")
                self.shot_code_track_combo.addItem(label, i)
                self.counter_track_combo.addItem(label, i)
                self.bg_track_combo.addItem(label, i)

            # Try to auto-detect FPS from timeline
            try:
                fps = get_timeline_fps(timeline)
                fps_str = fps_to_str(fps)
                idx = self.fps_combo.findText(fps_str)
                if idx >= 0:
                    self.fps_combo.setCurrentIndex(idx)
                else:
                    self.fps_combo.setCurrentText("Custom...")
                    self.custom_fps_input.setText(str(fps))
                    self.custom_fps_input.show()
            except Exception:
                pass

        except Exception as e:
            self.log.append(f"⚠️  Resolve connection error: {e}")

    def on_counter_track_changed(self, text):
        """Show/hide BG track based on counter track selection."""
        is_none = (text == "None")
        self.bg_track_label.setVisible(is_none)
        self.bg_track_combo.setVisible(is_none)

    def on_fps_changed(self, text):
        if text == "Custom...":
            self.custom_fps_input.show()
            self.custom_fps_input.setFocus()
        else:
            self.custom_fps_input.hide()

    def validate_custom_fps(self, text):
        if not text:
            return
        try:
            fps = float(text)
            if fps <= 0:
                self.custom_fps_input.setStyleSheet("background-color: #ffcccc;")
            else:
                self.custom_fps_input.setStyleSheet("")
        except ValueError:
            self.custom_fps_input.setStyleSheet("background-color: #ffcccc;")

    def get_fps(self):
        if self.fps_combo.currentText() == "Custom...":
            try:
                fps = float(self.custom_fps_input.text())
                if fps <= 0:
                    return None
                return fps
            except ValueError:
                return None
        else:
            return float(self.fps_combo.currentText())

    def browse_old_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Old Shot List",
            self.old_excel_input.text() or str(Path.home() / "Downloads"),
            "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.old_excel_input.setText(path)

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Shot List As",
            self.output_input.text(),
            "Excel Files (*.xlsx)"
        )
        if path:
            self.output_input.setText(path)

    def start_processing(self):
        fps = self.get_fps()
        if fps is None:
            QMessageBox.warning(self, "Error", "Please enter a valid FPS value")
            return

        shot_code_track = self.shot_code_track_combo.currentData()
        if shot_code_track is None:
            QMessageBox.warning(self, "Error", "Please select a Shot Code Track")
            return

        counter_text = self.counter_track_combo.currentText()
        if counter_text == "None":
            frame_counter_track = None
            bg_track = self.bg_track_combo.currentData()
            if bg_track is None:
                QMessageBox.warning(self, "Error",
                    "When Frame Counter Track is None, please select a BG Track")
                return
        else:
            frame_counter_track = self.counter_track_combo.currentData()
            bg_track = None

        output = self.output_input.text()
        if not output:
            QMessageBox.warning(self, "Error", "Please specify an output file")
            return

        old_excel = self.old_excel_input.text().strip()
        if old_excel and not os.path.exists(old_excel):
            QMessageBox.warning(self, "Error", "Old Shot List file not found")
            return
        old_excel_path = old_excel if old_excel else None

        self.log.clear()
        self.go_btn.setEnabled(False)
        self.progress.setRange(0, 0)
        self.progress.show()

        self.worker = ShotListWorker(
            shot_code_track=shot_code_track,
            frame_counter_track=frame_counter_track,
            bg_track=bg_track,
            bottom_track=self.bottom_spin.value(),
            top_track=self.top_spin.value(),
            cut_in_frame=self.cut_in_spin.value(),
            old_excel_path=old_excel_path,
            work_handle=self.work_handle_spin.value(),
            scan_handle=self.scan_handle_spin.value(),
            output_path=output,
            input_sequence=self.seq_name.text() or None,
            fps=fps,
        )
        self.worker.progress.connect(self.update_log)
        self.worker.finished.connect(self.processing_done)
        self.worker.start()

    def update_log(self, msg):
        self.log.append(msg)
        self.log.verticalScrollBar().setValue(
            self.log.verticalScrollBar().maximum()
        )

    def processing_done(self, success, msg):
        self.go_btn.setEnabled(True)
        self.progress.hide()

        if success:
            QMessageBox.information(self, "Success", msg)
        else:
            QMessageBox.critical(self, "Error", f"Export failed: {msg}")


def main():
    app = QApplication(sys.argv)

    theia_dir = Path("/Library/Application Support/Theia")
    icon_path = theia_dir / "resources" / "graphics" / "shot_list_icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = ShotListGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
