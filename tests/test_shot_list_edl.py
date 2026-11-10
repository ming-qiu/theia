"""EDL-authoritative export regressions from DLA_1105 (24 fps)."""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import shot_list_gui as shot_list


class EDLMappingTests(unittest.TestCase):
    def setUp(self):
        self.events = shot_list.parse_edl(Path(__file__).parent / "fixtures/dlag_dissolves.edl")

    def test_incoming_dissolves_select_correct_event_without_neighbor_handles(self):
        for start, end, source_in, half, speed in (
            (90691, 90771, "16:32:04:22", 12, 19.2),
            (90857, 91006, "16:32:09:15", 24, 12.0),
        ):
            with self.subTest(start=start):
                selected = [(event, shot_list.edl_event_for_shot(event, 24, start, end))
                            for event in self.events]
                selected = [(event, info) for event, info in selected if info is not None]
                self.assertEqual(len(selected), 1)
                event, info = selected[0]
                self.assertEqual(event["src_in"], source_in)
                self.assertEqual(event["dissolve_in"], half)
                self.assertEqual(event["retime_fps"], speed)
                self.assertEqual(info["TimelineStart"], start - half)
                self.assertEqual(info["TimelineEnd"], end)

    def test_zero_length_marker_does_not_modify_unrelated_previous_event(self):
        self.assertNotIn("dissolve_out", self.events[0])
        self.assertIsNone(self.events[0]["retime_fps"])

    def test_matching_outgoing_marker_preserves_dissolve_and_retime(self):
        outgoing = self.events[3]
        self.assertEqual(outgoing["clip_name"], "17/18_10_3_C")
        self.assertEqual(outgoing["dissolve_out"], 24)
        self.assertEqual(outgoing["retime_fps"], 12)

    def test_event_spanning_shots_is_trimmed_in_source_time(self):
        event = dict(self.events[0], rec_in="01:00:00:00", rec_out="01:00:10:00",
                     src_in="02:00:00:00", src_out="02:00:05:00", retime_fps=12)
        info = shot_list.edl_event_for_shot(event, 24, 86448, 86496)
        self.assertEqual(info["ClipInTC"], "02:00:01:00")
        self.assertEqual(info["ClipDuration"], 48)
        self.assertEqual(info["ClipOutTC"], "02:00:01:23")

    def test_export_uses_one_foreground_edl_event_without_reading_segments(self):
        # Resolve V2 has two segments (105541-105690 and 105690-105887).
        # Export must use event 014 once, even if timeline segmentation differs.
        fixture = Path(__file__).parent / "fixtures/dlag_0400_fg.edl"
        counter = Mock()
        counter.GetName.return_value = "DLAG_0400"
        counter.GetStart.return_value = 105541
        counter.GetEnd.return_value = 105887
        counter.GetDuration.return_value = 346
        counter.GetSourceStartTime.return_value = 1001 / 24
        timeline = Mock()
        timeline.GetName.return_value = "DLA_1105"

        def items(kind, track):
            self.assertEqual((kind, track), ("video", 4))
            return [counter]

        timeline.GetItemListInTrack.side_effect = items
        resolve = Mock()
        resolve.GetProjectManager.return_value.GetCurrentProject.return_value.GetCurrentTimeline.return_value = timeline
        with tempfile.TemporaryDirectory() as folder, patch.object(shot_list, "dvr") as api:
            api.scriptapp.return_value = resolve
            output = str(Path(folder) / "shots.xlsx")
            worker = shot_list.ShotListWorker(4, {2: str(fixture)}, None, 8, 8, output, None, 24)
            result = []
            worker.finished.connect(lambda ok, message: result.append((ok, message)))
            worker.run()
            self.assertTrue(result[0][0], result)
            workbook = load_workbook(output, data_only=True)
            try:
                rows = list(workbook["Elements"].values)
                self.assertEqual(len(rows), 2)
                element = dict(zip(rows[0], rows[1]))
                self.assertEqual(element["Element"], "ScanFg01")
                self.assertEqual(element["Clip Duration (with dissolve before retime)"], 346)
                self.assertEqual(element["Clip In TC"], "14:21:26:12")
                self.assertEqual(element["Clip Out TC"], "14:21:40:21")
                self.assertIsNone(element["Scale & Repo"])
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
