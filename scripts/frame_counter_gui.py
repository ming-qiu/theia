"""
Theia - Frame Counter Generator GUI
Generate frame counter videos with burned-in frame numbers
"""

import sys
import os
import shutil
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QFileDialog,
    QMessageBox, QProgressBar, QTextEdit, QGroupBox, QSpinBox
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon

from PIL import Image, ImageDraw, ImageFont
from moviepy.video.io.ImageSequenceClip import ImageSequenceClip
import subprocess
import tempfile


class FrameCounterWorker(QThread):
    """Threaded worker for frame counter generation."""
    progress = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, width, height, begin, end, fps, output_dir, font_path):
        super().__init__()
        self.width = width
        self.height = height
        self.begin = begin
        self.end = end
        self.fps = fps
        self.output_dir = output_dir
        self.font_path = font_path

    def log(self, msg):
        self.progress.emit(msg)

    def run(self):
        try:
            fps = self.fps
            w = self.width
            h = self.height
            begin = self.begin
            end = self.end

            # FPS label for filenames: 23.976 -> 23_976, 24.0 -> 24
            fps_label = f"{fps:g}".replace('.', '_')

            # Temp directory for frame images
            temp_frames_dir = tempfile.mkdtemp(prefix="fc_frames_")

            self.log(f"Settings: {w}x{h}, frames {begin}–{end}, {fps} fps")
            self.log(f"Font: {self.font_path}")
            self.log(f"Output: {self.output_dir}")
            self.log("=" * 50)

            # Load font
            try:
                font = ImageFont.truetype(self.font_path, int(0.75 * h))
            except Exception as e:
                self.log(f"WARNING: Could not load font '{self.font_path}': {e}")
                self.log("Falling back to default font")
                font = ImageFont.load_default()

            # Generate frame images
            total_frames = end - begin + 1
            self.log(f"Generating {total_frames} frame images...")
            for i, f in enumerate(range(begin, end + 1)):
                im = Image.new(mode="RGB", size=(w, h))
                draw = ImageDraw.Draw(im)
                draw.text((int(0.1 * h), int(0.1 * h)), str(f), font=font, fill=(255, 255, 255))
                file_name = str(f).zfill(4) + '.png'
                im.save(os.path.join(temp_frames_dir, file_name))

                # Log progress every 100 frames
                if (i + 1) % 100 == 0 or i == total_frames - 1:
                    self.log(f"  Frames: {i + 1}/{total_frames}")

            # Generate video from image sequence
            self.log("Creating video from frames...")
            clip = ImageSequenceClip(temp_frames_dir, fps=fps)

            temp_video_path = os.path.join(self.output_dir, f"temp_{fps_label}fps.mp4")
            clip.write_videofile(temp_video_path, fps=float(fps), logger=None)

            # Calculate starting timecode
            import timecode as tc_mod
            tc = tc_mod.Timecode(fps, frames=begin) + 1
            start_timecode = str(tc)
            self.log(f"Setting start timecode to: {start_timecode}")

            # Burn timecode metadata via ffmpeg
            video_path = os.path.join(self.output_dir, f"frame_counter_{fps_label}fps.mp4")

            ffmpeg_cmd = [
                'ffmpeg', '-i', temp_video_path,
                '-c', 'copy',
                '-timecode', start_timecode,
                '-y',
                video_path
            ]

            try:
                subprocess.run(ffmpeg_cmd, check=True, capture_output=True)
                self.log(f"✓ Timecode metadata applied")
                os.remove(temp_video_path)
            except subprocess.CalledProcessError as e:
                self.log(f"WARNING: ffmpeg timecode failed: {e.stderr.decode()}")
                self.log("  Saving video without timecode metadata")
                os.rename(temp_video_path, video_path)

            # Cleanup temp frames
            shutil.rmtree(temp_frames_dir, ignore_errors=True)

            self.log("=" * 50)
            self.log(f"✓ Done: {video_path}")
            self.finished.emit(True, f"Generated: {video_path}")

        except Exception as e:
            # Best-effort cleanup
            if 'temp_frames_dir' in locals():
                shutil.rmtree(temp_frames_dir, ignore_errors=True)
            if 'temp_video_path' in locals() and os.path.exists(temp_video_path):
                os.remove(temp_video_path)

            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.finished.emit(False, str(e))


class FrameCounterGUI(QMainWindow):
    """Main GUI window for Frame Counter Generator."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Theia - Frame Counter Generator")
        self.setMinimumSize(600, 650)

        # Main layout
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Title
        title = QLabel("Frame Counter Generator")
        font = QFont()
        font.setPointSize(16)
        font.setBold(True)
        title.setFont(font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        layout.addSpacing(5)

        # ── Video Size ──────────────────────────────────
        size_group = QGroupBox("Video Size")
        size_layout = QHBoxLayout()

        size_layout.addWidget(QLabel("Width:"))
        self.width_spin = QSpinBox()
        self.width_spin.setMinimum(100)
        self.width_spin.setMaximum(3840)
        self.width_spin.setValue(200)
        self.width_spin.setSuffix(" px")
        size_layout.addWidget(self.width_spin)

        size_layout.addSpacing(20)
        size_layout.addWidget(QLabel("Height:"))
        self.height_spin = QSpinBox()
        self.height_spin.setMinimum(50)
        self.height_spin.setMaximum(2160)
        self.height_spin.setValue(100)
        self.height_spin.setSuffix(" px")
        size_layout.addWidget(self.height_spin)

        size_layout.addStretch()
        size_group.setLayout(size_layout)
        layout.addWidget(size_group)

        # ── Frame Range ─────────────────────────────────
        range_group = QGroupBox("Frame Range")
        range_layout = QHBoxLayout()

        range_layout.addWidget(QLabel("Start:"))
        self.begin_spin = QSpinBox()
        self.begin_spin.setMinimum(0)
        self.begin_spin.setMaximum(999999)
        self.begin_spin.setValue(1001)
        range_layout.addWidget(self.begin_spin)

        range_layout.addSpacing(20)
        range_layout.addWidget(QLabel("End:"))
        self.end_spin = QSpinBox()
        self.end_spin.setMinimum(0)
        self.end_spin.setMaximum(999999)
        self.end_spin.setValue(2000)
        range_layout.addWidget(self.end_spin)

        range_layout.addSpacing(20)
        self.frame_count_label = QLabel("")
        self.frame_count_label.setStyleSheet("color: #666;")
        range_layout.addWidget(self.frame_count_label)

        range_layout.addStretch()
        range_group.setLayout(range_layout)
        layout.addWidget(range_group)

        # Wire up live frame-count label
        self.begin_spin.valueChanged.connect(self.update_frame_count)
        self.end_spin.valueChanged.connect(self.update_frame_count)
        self.update_frame_count()

        # ── Frame Rate ──────────────────────────────────
        fps_group = QGroupBox("Frame Rate")
        fps_layout = QHBoxLayout()

        fps_layout.addWidget(QLabel("FPS:"))
        self.fps_combo = QComboBox()
        self.fps_combo.setMaximumWidth(120)
        self.fps_combo.addItems(["23.976", "24", "25", "29.97", "30", "60", "Custom..."])
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

        # ── Font ────────────────────────────────────────
        font_group = QGroupBox("Font")
        font_layout = QHBoxLayout()

        font_layout.addWidget(QLabel("Font File:"))
        self.font_input = QLineEdit(str(Path("/Library/Application Support/Theia") / "resources" / "fonts" / "SF-Pro-Text-Regular.otf"))
        self.font_input.setPlaceholderText("Path to .otf / .ttf font file")
        font_layout.addWidget(self.font_input)

        browse_font_btn = QPushButton("Browse...")
        browse_font_btn.setMaximumWidth(100)
        browse_font_btn.clicked.connect(self.browse_font)
        font_layout.addWidget(browse_font_btn)

        font_group.setLayout(font_layout)
        layout.addWidget(font_group)

        # ── Output ──────────────────────────────────────
        output_group = QGroupBox("Output")
        output_layout = QHBoxLayout()

        output_layout.addWidget(QLabel("Output Directory:"))
        self.output_dir_input = QLineEdit(str(Path.home() / "Downloads/frame-counters"))
        self.output_dir_input.setPlaceholderText("Directory for output video")
        output_layout.addWidget(self.output_dir_input)

        browse_output_btn = QPushButton("Browse...")
        browse_output_btn.setMaximumWidth(100)
        browse_output_btn.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(browse_output_btn)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        layout.addSpacing(5)

        # ── Go button ───────────────────────────────────
        self.go_btn = QPushButton("Go")
        self.go_btn.setMinimumHeight(40)
        self.go_btn.clicked.connect(self.start_generation)
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

    # ── Live helpers ────────────────────────────────────
    def update_frame_count(self):
        count = self.end_spin.value() - self.begin_spin.value() + 1
        self.frame_count_label.setText(f"{count} frames" if count > 0 else "⚠️  invalid range")

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
            self.custom_fps_input.setStyleSheet("" if fps > 0 else "background-color: #ffcccc;")
        except ValueError:
            self.custom_fps_input.setStyleSheet("background-color: #ffcccc;")

    def get_fps(self):
        if self.fps_combo.currentText() == "Custom...":
            try:
                fps = float(self.custom_fps_input.text())
                return fps if fps > 0 else None
            except ValueError:
                return None
        return float(self.fps_combo.currentText())

    # ── Browse dialogs ──────────────────────────────────
    def browse_font(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Font File",
            self.font_input.text(),
            "Font Files (*.otf *.ttf *.TTF *.OTF);;All Files (*.*)"
        )
        if path:
            self.font_input.setText(path)

    def browse_output_dir(self):
        directory = QFileDialog.getExistingDirectory(
            self, "Select Output Directory",
            self.output_dir_input.text()
        )
        if directory:
            self.output_dir_input.setText(directory)

    # ── Kick off ────────────────────────────────────────
    def start_generation(self):
        # Validate frame range
        begin = self.begin_spin.value()
        end = self.end_spin.value()
        if end <= begin:
            QMessageBox.warning(self, "Error", "End frame must be greater than start frame")
            return

        # Validate FPS
        fps = self.get_fps()
        if fps is None:
            QMessageBox.warning(self, "Error", "Please enter a valid FPS value")
            return

        # Validate font (warn but don't block — PIL has a fallback)
        font_path = self.font_input.text()
        if not font_path or not os.path.exists(font_path):
            reply = QMessageBox.question(
                self, "Font Not Found",
                f"Font file not found:\n{font_path}\n\nContinue with default font?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        # Create output directory if needed
        output_dir = self.output_dir_input.text()
        if not output_dir:
            QMessageBox.warning(self, "Error", "Please specify an output directory")
            return
        os.makedirs(output_dir, exist_ok=True)

        # Launch worker
        self.log.clear()
        self.go_btn.setEnabled(False)
        self.progress.setRange(0, 0)
        self.progress.show()

        self.worker = FrameCounterWorker(
            width=self.width_spin.value(),
            height=self.height_spin.value(),
            begin=begin,
            end=end,
            fps=fps,
            output_dir=output_dir,
            font_path=font_path
        )
        self.worker.progress.connect(self.update_log)
        self.worker.finished.connect(self.generation_done)
        self.worker.start()

    # ── Callbacks ───────────────────────────────────────
    def update_log(self, msg):
        self.log.append(msg)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def generation_done(self, success, msg):
        self.go_btn.setEnabled(True)
        self.progress.hide()
        if success:
            QMessageBox.information(self, "Success", msg)
        else:
            QMessageBox.critical(self, "Error", f"Generation failed: {msg}")


def main():
    app = QApplication(sys.argv)

    theia_dir = Path("/Library/Application Support/Theia")
    icon_path = theia_dir / "resources" / "graphics" / "frame_counter_icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = FrameCounterGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()