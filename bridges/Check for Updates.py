#!/usr/bin/env python
"""Launch the Theia updater from DaVinci Resolve's Scripts menu."""
import os
import subprocess
from pathlib import Path


def main():
    theia_dir = Path("/Library/Application Support/Theia")
    python_exe = theia_dir / "venv/bin/python3"
    gui_script = theia_dir / "update_gui.py"
    log_dir = theia_dir / "log"
    log_file = log_dir / "theia_update_bridge_log.txt"

    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        with open(log_file, "w") as log:
            log.write("=== Theia Update Bridge ===\n")
            log.write(f"Python executable: {python_exe}\n")
            log.write(f"GUI script: {gui_script}\n")

            if not python_exe.exists():
                log.write("ERROR: Theia Python environment was not found.\n")
                return
            if not gui_script.exists():
                log.write("ERROR: update_gui.py was not found.\n")
                return

            env = os.environ.copy()
            env["PATH"] = f"{theia_dir}/venv/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"
            gui_log = log_dir / "theia_update_gui_log.txt"
            with open(gui_log, "w") as gui_output:
                process = subprocess.Popen(
                    [str(python_exe), str(gui_script)],
                    env=env,
                    stdout=gui_output,
                    stderr=subprocess.STDOUT,
                    cwd=str(theia_dir),
                )
            log.write(f"Update GUI launched with PID {process.pid}.\n")
    except Exception:
        with open(log_file, "a") as log:
            log.write("\nEXCEPTION:\n")
            import traceback
            log.write(traceback.format_exc())


if __name__ == "__main__":
    main()
