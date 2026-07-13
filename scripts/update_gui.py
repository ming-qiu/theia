"""Check for, download, and launch Theia updates."""
import json
import os
import re
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QLabel, QMainWindow, QMessageBox, QProgressBar,
    QPushButton, QVBoxLayout, QWidget,
)


THEIA_DIR = Path("/Library/Application Support/Theia")
GITHUB_REPOSITORY = "ming-qiu/theia"
BRANCH_PATTERN = re.compile(r"^release/v(\d+(?:\.\d+)*)$")


def read_installed_version():
    """Read the version copied into the installed Theia directory."""
    try:
        version = (THEIA_DIR / "VERSION").read_text(encoding="utf-8").strip()
        if re.fullmatch(r"\d+(?:\.\d+)*", version):
            return version
    except OSError:
        pass
    return None


def version_tuple(version):
    return tuple(int(part) for part in version.split("."))


class CheckWorker(QThread):
    finished = Signal(bool, str, str)

    def run(self):
        try:
            result = subprocess.run(
                [
                    "/usr/bin/curl", "-fsSL", "--retry", "3",
                    "--connect-timeout", "15", "--max-time", "45",
                    "--proto", "=https",
                    "-H", "Accept: application/vnd.github+json",
                    "-H", "User-Agent: Theia-Updater",
                    f"https://api.github.com/repos/{GITHUB_REPOSITORY}/branches?per_page=100",
                ],
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                message = result.stderr.strip() or f"curl exited with status {result.returncode}"
                raise RuntimeError(message)
            branches = json.loads(result.stdout)

            releases = []
            for branch in branches:
                name = branch.get("name", "")
                match = BRANCH_PATTERN.fullmatch(name)
                if match:
                    releases.append((version_tuple(match.group(1)), match.group(1), name))
            if not releases:
                raise RuntimeError("GitHub has no versioned release branches.")

            _, version, branch = max(releases)
            self.finished.emit(True, version, branch)
        except Exception as error:
            self.finished.emit(False, "", str(error))


class DownloadWorker(QThread):
    finished = Signal(bool, str)

    def __init__(self, branch):
        super().__init__()
        self.branch = branch

    def run(self):
        try:
            work_dir = Path(tempfile.mkdtemp(prefix="theia-update-"))
            archive = work_dir / "theia.tar.gz"
            url = (
                f"https://github.com/{GITHUB_REPOSITORY}/archive/refs/heads/"
                f"{self.branch}.tar.gz"
            )
            with open(archive, "wb") as output:
                result = subprocess.run(
                    [
                        "/usr/bin/curl", "-fL", "--retry", "3",
                        "--connect-timeout", "15", "--max-time", "300",
                        "--proto", "=https", "-H", "User-Agent: Theia-Updater",
                        url,
                    ],
                    stdout=output,
                    stderr=subprocess.PIPE,
                )
            if result.returncode != 0:
                message = result.stderr.decode("utf-8", errors="replace").strip()
                raise RuntimeError(message or f"curl exited with status {result.returncode}")

            extract_dir = work_dir / "release"
            extract_dir.mkdir()
            with tarfile.open(archive, "r:gz") as bundle:
                destination = str(extract_dir.resolve())
                for member in bundle.getmembers():
                    target = str((extract_dir / member.name).resolve())
                    if os.path.commonpath([destination, target]) != destination:
                        raise RuntimeError("The downloaded archive contains an unsafe path.")
                bundle.extractall(extract_dir)

            installers = list(extract_dir.glob("*/install.command"))
            if len(installers) != 1:
                raise RuntimeError("The downloaded release does not contain install.command.")

            installer = installers[0]
            version_file = installer.parent / "VERSION"
            expected_version = BRANCH_PATTERN.fullmatch(self.branch).group(1)
            if not version_file.exists() or version_file.read_text(encoding="utf-8").strip() != expected_version:
                raise RuntimeError("The downloaded release has inconsistent version information.")

            installer.chmod(installer.stat().st_mode | 0o100)
            self.finished.emit(True, str(installer))
        except Exception as error:
            self.finished.emit(False, str(error))


class UpdateGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.installed_version = read_installed_version()
        self.latest_version = None
        self.latest_branch = None
        self.worker = None
        self.setup_ui()
        self.check_for_updates()

    def setup_ui(self):
        self.setWindowTitle("Theia - Update")
        self.setFixedSize(440, 250)
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        title = QLabel("Update Theia")
        title.setStyleSheet("font-size: 22px; font-weight: 600;")
        layout.addWidget(title)

        installed = self.installed_version or "Unknown (installed before version tracking)"
        self.installed_label = QLabel(f"Installed version: {installed}")
        self.latest_label = QLabel("Latest version: Checking…")
        self.status_label = QLabel("Contacting GitHub…")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.installed_label)
        layout.addWidget(self.latest_label)
        layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        layout.addWidget(self.progress)

        self.action_button = QPushButton("Checking for Updates…")
        self.action_button.setMinimumHeight(38)
        self.action_button.setEnabled(False)
        self.action_button.clicked.connect(self.on_action)
        layout.addWidget(self.action_button)

    def check_for_updates(self):
        self.progress.show()
        self.action_button.setEnabled(False)
        self.action_button.setText("Checking for Updates…")
        self.status_label.setText("Contacting GitHub…")
        self.worker = CheckWorker()
        self.worker.finished.connect(self.check_finished)
        self.worker.start()

    def check_finished(self, success, latest_version, branch_or_error):
        self.progress.hide()
        if not success:
            self.latest_label.setText("Latest version: Unable to check")
            self.status_label.setText(branch_or_error)
            self.action_button.setText("Try Again")
            self.action_button.setEnabled(True)
            self.latest_branch = None
            return

        self.latest_version = latest_version
        self.latest_branch = branch_or_error
        self.latest_label.setText(f"Latest version: {latest_version}")
        current = version_tuple(self.installed_version) if self.installed_version else None
        if current is not None and current >= version_tuple(latest_version):
            self.status_label.setText("Theia is up to date.")
            self.action_button.setText("Check Again")
        else:
            self.status_label.setText("A newer version of Theia is available.")
            self.action_button.setText(f"Download and Install {latest_version}")
        self.action_button.setEnabled(True)

    def on_action(self):
        if not self.latest_branch:
            self.check_for_updates()
            return

        current = version_tuple(self.installed_version) if self.installed_version else None
        if current is not None and current >= version_tuple(self.latest_version):
            self.check_for_updates()
            return

        answer = QMessageBox.question(
            self, "Install Theia Update",
            f"Download Theia {self.latest_version} from GitHub and open its installer?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        self.action_button.setEnabled(False)
        self.action_button.setText("Downloading…")
        self.status_label.setText("Downloading and preparing the installer…")
        self.progress.show()
        self.worker = DownloadWorker(self.latest_branch)
        self.worker.finished.connect(self.download_finished)
        self.worker.start()

    def download_finished(self, success, result):
        self.progress.hide()
        if not success:
            self.status_label.setText(f"Download failed: {result}")
            self.action_button.setText("Try Again")
            self.action_button.setEnabled(True)
            return
        try:
            subprocess.Popen(["/usr/bin/open", result])
        except Exception as error:
            self.status_label.setText(f"Could not open installer: {error}")
            self.action_button.setText("Try Again")
            self.action_button.setEnabled(True)
            return

        QMessageBox.information(
            self, "Installer Opened",
            "The latest installer has opened in Terminal. Follow its prompts to finish the update.",
        )
        self.close()


def main():
    app = QApplication(sys.argv)

    theia_dir = Path("/Library/Application Support/Theia")
    icon_path = theia_dir / "resources" / "graphics" / "clip_inventory_icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = UpdateGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
