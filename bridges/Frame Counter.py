#!/usr/bin/env python
"""
Theia - Clip Inventory Bridge
Place in: ~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Edit/
"""
import subprocess
import os
import sys
import glob
from pathlib import Path

def main():
    theia_dir = Path("/Library/Application Support/Theia")
    python_exe = theia_dir / "venv/bin/python3"
    gui_script = theia_dir / "frame_counter_gui.py"
    
    # Write diagnostic info immediately
    log_file = Path.home() / "Desktop" / "theia_bridge_log.txt"
    
    try:
        with open(log_file, "w") as f:
            f.write("=== Theia Bridge Diagnostic ===\n")
            f.write(f"Python version: {sys.version}\n")
            f.write(f"Python executable: {sys.executable}\n")
            f.write(f"Current working directory: {os.getcwd()}\n")
            f.write(f"PATH: {os.environ.get('PATH', 'NOT SET')}\n\n")
            
            f.write(f"Theia dir: {theia_dir}\n")
            f.write(f"Theia exists: {theia_dir.exists()}\n")
            f.write(f"Python exe: {python_exe}\n")
            f.write(f"Python exists: {python_exe.exists()}\n")
            f.write(f"GUI script: {gui_script}\n")
            f.write(f"GUI exists: {gui_script.exists()}\n\n")
            
            if not python_exe.exists():
                f.write("ERROR: Python executable not found!\n")
                return
            
            if not gui_script.exists():
                f.write("ERROR: GUI script not found!\n")
                return
            
            # Set up clean environment
            env = os.environ.copy()
            env['PATH'] = f"{theia_dir}/venv/bin:/usr/local/bin:/usr/bin:/bin"
            
            # Check what architecture PIL was built for
            pil_libs = glob.glob(str(theia_dir / "venv/lib/python*/site-packages/PIL/_imaging*.so"))
            pil_arch = "unknown"
            
            if pil_libs:
                result = subprocess.run(["file", pil_libs[0]], capture_output=True, text=True)
                if "x86_64" in result.stdout and "arm64" not in result.stdout:
                    pil_arch = "x86_64"
                elif "arm64" in result.stdout:
                    pil_arch = "arm64"
            
            f.write(f"PIL architecture: {pil_arch}\n")
            f.write("Launching GUI...\n")
            
            # Launch GUI - force architecture to match PIL if needed
            gui_log = Path.home() / "Desktop" / "theia_gui_log.txt"
            
            with open(gui_log, "w") as gui_out:
                if pil_arch == "x86_64":
                    f.write("Forcing x86_64 mode to match PIL libraries\n")
                    process = subprocess.Popen(
                        ["arch", "-x86_64", str(python_exe), str(gui_script)],
                        env=env,
                        stdout=gui_out,
                        stderr=subprocess.STDOUT,
                        cwd=str(theia_dir)
                    )
                else:
                    f.write("Using native architecture\n")
                    process = subprocess.Popen(
                        [str(python_exe), str(gui_script)],
                        env=env,
                        stdout=gui_out,
                        stderr=subprocess.STDOUT,
                        cwd=str(theia_dir)
                    )
            
            f.write(f"Process launched with PID: {process.pid}\n")
            f.write(f"GUI output will be in: {gui_log}\n")
            f.write("SUCCESS!\n")
            
    except Exception as e:
        with open(log_file, "a") as f:
            f.write(f"\nEXCEPTION: {e}\n")
            import traceback
            f.write(traceback.format_exc())

if __name__ == "__main__":
    main()