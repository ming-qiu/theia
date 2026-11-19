"""Regression checks for EDL dissolve ranges and retimed source frames."""

import importlib.util
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("shot_list_gui", ROOT / "scripts/shot_list_gui.py")
shot_list = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(shot_list)


class DissolveTimingTests(unittest.TestCase):
    def event(self, speed=None, outgoing=5, incoming=0):
        return {
            "src_in": "00:29:19:11",
            "rec_in": "01:00:01:00",
            "rec_out": "01:00:01:05",
            "retime_fps": speed,
            "dissolve_in": incoming,
            "dissolve_out": outgoing,
        }

    def test_outgoing_dissolve_includes_both_halves(self):
        info = shot_list.edl_event_for_shot(self.event(), 25, 90025, 90035)
        self.assertEqual(info["TimelineStart"], 90025)
        self.assertEqual(info["TimelineEnd"], 90040)
        self.assertEqual(info["ClipDuration"], 15)
        self.assertEqual(info["ClipOutFrames"] - info["ClipInFrames"] + 1, 15)

    def test_dissolve_tail_does_not_create_neighbor_element(self):
        self.assertIsNone(shot_list.edl_event_for_shot(self.event(), 25, 90035, 90045))

    def test_incoming_dissolve_does_not_create_previous_element(self):
        event = self.event(outgoing=0, incoming=5)
        event.update(rec_in="01:00:01:05", rec_out="01:00:01:20")
        self.assertIsNone(shot_list.edl_event_for_shot(event, 25, 90025, 90035))
        info = shot_list.edl_event_for_shot(event, 25, 90035, 90045)
        self.assertEqual(info["ClipDuration"], 15)

    def test_retime_uses_full_dissolve_range(self):
        for speed, source_length in [(12.5, 8), (50, 30)]:
            with self.subTest(speed=speed):
                info = shot_list.edl_event_for_shot(self.event(speed=speed), 25, 90025, 90035)
                self.assertEqual(info["ClipDuration"], 15)
                self.assertEqual(info["ClipOutFrames"] - info["ClipInFrames"] + 1, source_length)
                element = dict(info, ClipIn=1001, ClipOut=1015, HeadIn=976,
                               TailOut=1040, HasRetime=True, RetimeFPS=speed,
                               Speed=speed / 25, RetimeSummary="", ScaleRepo="")
                tracks = {1: [element]}
                shot_list.retime_summary(tracks, 25, 25)
                self.assertEqual(element["ClipOut"], 1015)
                self.assertEqual(element["ClipDuration"], source_length)
                self.assertEqual(element["ClipOutFrames"] - element["ClipInFrames"] + 1, source_length)
                self.assertEqual(element["TailOut"], 1001 + source_length + 25 - 1)

    def test_no_dissolve_still_uses_record_duration(self):
        info = shot_list.edl_event_for_shot(self.event(outgoing=0), 25, 90025, 90035)
        self.assertEqual(info["ClipDuration"], 5)
        self.assertEqual(info["TimelineEnd"], 90030)


if __name__ == "__main__":
    unittest.main()
