"""Build script for cross-compiling Windows executable using Nuitka."""

import subprocess
import sys
import os


def build_windows_exe():
    """
    Build Windows executable using Nuitka with cross-compilation support.

    This script attempts to compile for Windows from Linux/WSL.
    For best results, run this on a Windows machine.
    """
    # Build command for Nuitka
    cmd = [
        "nuitka",
        "--standalone",                  # Create standalone executable
        "--onefile",                     # Single file output
        "--enable-plugin=tk-inter",      # Enable tkinter plugin
        "--follow-imports=tkinterdnd2",  # Include tkinterdnd2
        "--assume-yes-for-downloads",    # Auto-confirm downloads
        "--output-dir=dist",
        "--output-filename=ExcelCompare.exe",
        "src/excel_compare/main.py"
    ]

    # Try with Mingw-w64 for Windows cross-compilation
    mingw_cmd = [
        "nuitka",
        "--standalone",
        "--onefile",
        "--enable-plugin=tk-inter",
        "--include-package=tkinterdnd2",
        "--assume-yes-for-downloads",
        "--output-dir=dist",
        "--output-filename=ExcelCompare.exe",
        "--windows-console-mode=disable",  # No console window
        "--mingw64",                      # Use Mingw-w64 for Windows
        "src/excel_compare/main.py"
    ]

    # Remove empty icon argument if not exists
    mingw_cmd = [x for x in mingw_cmd if x]

    print("Attempting Windows cross-compilation with Nuitka...")
    print(f"Command: {' '.join(mingw_cmd)}")
    print("-" * 60)

    try:
        # First, try with Mingw
        result = subprocess.run(mingw_cmd, check=True, capture_output=True, text=True)
        print(result.stdout)

        print("\n" + "=" * 60)
        print("Build completed!")
        print(f"Executable location: dist/ExcelCompare.exe")

        if not os.path.exists("dist/ExcelCompare.exe"):
            print("\nNote: If ExcelCompare.exe was not created, cross-compilation failed.")
            print("Please run this script on a Windows machine for reliable results.")

    except subprocess.CalledProcessError as e:
        print(f"\nCross-compilation failed with error:\n{e.stderr}")
        print("\n" + "=" * 60)
        print("Fallback: Trying standard PyInstaller (Linux executable)...")
        print("For Windows exe, please run on a Windows machine.")

        # Fallback to standard build
        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            "--name", "ExcelCompare",
            "--add-data", "src/excel_compare:excel_compare",
            "--hidden-import", "tkinterdnd2",
            "src/excel_compare/main.py"
        ]

        try:
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            print(result.stdout)
            print("\nLinux executable created at: dist/ExcelCompare")
        except subprocess.CalledProcessError as e2:
            print(f"Build also failed: {e2.stderr}")
            sys.exit(1)


def create_windows_build_script():
    """Create a .bat script for Windows users."""
    bat_content = """@echo off
echo Building Windows executable for Excel Compare Tool...
echo.

uv sync
uv run python build.py

echo.
if exist "dist\\ExcelCompare.exe" (
    echo Build successful! Executable is at: dist\\ExcelCompare.exe
) else (
    echo Build failed. Check the output above for errors.
)

pause
"""

    with open("build_windows.bat", "w") as f:
        f.write(bat_content)

    print("Created build_windows.bat for Windows users")


if __name__ == "__main__":
    # Check if we're on Linux/WSL
    if os.name != 'nt':
        print("Running on Linux/WSL - attempting cross-compilation...")
        print("Note: Cross-compilation may fail. For best results, run on Windows.")
        print()
        create_windows_build_script()
        build_windows_exe()
    else:
        print("Running on Windows - using PyInstaller for native build...")
        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            "--name", "ExcelCompare",
            "--add-data", "src/excel_compare;excel_compare",
            "--hidden-import", "tkinterdnd2",
            "src/excel_compare/main.py"
        ]

        try:
            result = subprocess.run(cmd, check=True)
            print("\nBuild completed! Executable is at: dist\\ExcelCompare.exe")
        except subprocess.CalledProcessError as e:
            print(f"\nBuild failed: {e}")
            sys.exit(1)
