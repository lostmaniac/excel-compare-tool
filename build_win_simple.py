#!/usr/bin/env python3
"""
Simplified Windows cross-compilation script

This script attempts to create a Windows executable using Nuitka
without complex mingw configuration.
"""

import subprocess
import sys
import os

def compile_windows_exe():
    """Compile Windows exe using Nuitka."""

    print("=" * 60)
    print("Compiling Windows executable (Simplified)")
    print("=" * 60)
    print()

    # Clean previous builds
    if os.path.exists('build'):
        import shutil
        shutil.rmtree('build')
    if os.path.exists('dist'):
        import shutil
        shutil.rmtree('dist')

    # Nuitka command - simplified version
    cmd = [
        'uv', 'run', 'nuitka',
        '--standalone',
        '--onefile',
        '--include-package=tkinterdnd2',
        '--assume-yes-for-downloads',
        '--output-dir=dist',
        '--output-filename=ExcelCompare.exe',
        '--remove-output',
        'src/excel_compare/main.py'
    ]

    print("Running Nuitka command:")
    print(' '.join(cmd))
    print()
    print("=" * 60)
    print()

    try:
        # Run the compilation
        result = subprocess.run(cmd, check=True, capture_output=False)

        print()
        print("=" * 60)
        print("Build completed!")
        print("=" * 60)

        # Check if exe was created
        exe_path = 'dist/ExcelCompare.exe'
        if os.path.exists(exe_path):
            size = os.path.getsize(exe_path)
            size_mb = size / (1024 * 1024)
            print(f"✓ Executable created: {exe_path}")
            print(f"✓ File size: {size_mb:.2f} MB")

            # Check if it's actually a Windows PE file
            result = subprocess.run(['file', exe_path], capture_output=True, text=True)
            print(f"✓ File type: {result.stdout.strip()}")
        else:
            print("⚠ ExcelCompare.exe not found in dist/")

    except subprocess.CalledProcessError as e:
        print()
        print("=" * 60)
        print("Build failed!")
        print("=" * 60)
        print(f"Error code: {e.returncode}")

if __name__ == "__main__":
    compile_windows_exe()
