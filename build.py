"""Build script for creating Windows executable using PyInstaller."""

import subprocess
import sys
import os


def build_exe():
    """Build the executable using PyInstaller."""
    # Platform-specific path separator for --add-data
    path_sep = ';' if os.name == 'nt' else ':'

    # Build command for PyInstaller
    cmd = [
        "pyinstaller",
        "--onefile",          # Create a single file
        "--windowed",         # No console window
        "--name", "ExcelCompare",  # Name of the executable
        "--add-data", f"src/excel_compare{path_sep}excel_compare",  # Include the package
        "--hidden-import", "tkinterdnd2",  # Ensure tkinterdnd2 is included
        "src/excel_compare/main.py"
    ]

    print("Building Windows executable...")
    print(f"Command: {' '.join(cmd)}")
    print("-" * 60)

    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(result.stdout)
        print("\n" + "=" * 60)
        print("Build completed successfully!")
        print(f"Executable location: dist/ExcelCompare.exe")
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with error:\n{e.stderr}")
        sys.exit(1)


if __name__ == "__main__":
    build_exe()
