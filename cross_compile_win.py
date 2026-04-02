#!/usr/bin/env python3
"""
Cross-compile Windows executable from Linux using Nuitka

This script uses Nuitka with Mingw-w64 to create a Windows executable.
"""

import subprocess
import sys
import os

def compile_windows_exe():
    """Compile Windows exe using Nuitka with Mingw-w64."""

    # Set Mingw environment variables
    env = os.environ.copy()
    env['CC'] = 'x86_64-w64-mingw32-gcc'
    env['CXX'] = 'x86_64-w64-mingw32-g++'
    env['LD'] = 'x86_64-w64-mingw32-gcc'
    env['AR'] = 'x86_64-w64-mingw32-ar'
    env['RANLIB'] = 'x86_64-w64-mingw32-ranlib'

    print("=" * 60)
    print("Cross-compiling Windows executable from Linux")
    print("=" * 60)
    print()

    # Clean previous builds
    if os.path.exists('build'):
        import shutil
        shutil.rmtree('build')
    if os.path.exists('dist'):
        import shutil
        shutil.rmtree('dist')

    # Nuitka command for Windows cross-compilation
    cmd = [
        'uv', 'run', 'nuitka',
        '--standalone',
        '--onefile',
        '--enable-plugin=tk-inter',
        '--include-package=tkinterdnd2',
        '--follow-imports',  # Include all imported modules
        '--assume-yes-for-downloads',
        '--output-dir=dist',
        '--output-filename=ExcelCompare.exe',
        '--mingw64',  # Use Mingw-w64
        '--windows-console-mode=disable',  # No console
        '--remove-output',  # Clean up build files
        'src/excel_compare/main.py'
    ]

    print("Running Nuitka command:")
    print(' '.join(cmd))
    print()
    print("=" * 60)
    print()

    try:
        # Run the compilation
        result = subprocess.run(cmd, env=env, check=True, capture_output=False)

        print()
        print("=" * 60)
        print("Build completed!")
        print("=" * 60)

        # Check if exe was created
        exe_path = 'dist/ExcelCompare.exe'
        if os.path.exists(exe_path):
            size = os.path.getsize(exe_path)
            size_mb = size / (1024 * 1024)
            print(f"✓ Windows executable created: {exe_path}")
            print(f"✓ File size: {size_mb:.2f} MB")
        else:
            print("⚠ ExcelCompare.exe not found in dist/")
            print("Build may have partially succeeded")

    except subprocess.CalledProcessError as e:
        print()
        print("=" * 60)
        print("Build failed!")
        print("=" * 60)
        print(f"Error code: {e.returncode}")
        print()
        print("Troubleshooting:")
        print("1. Ensure mingw-w64 is installed:")
        print("   apt-get install mingw-w64 g++-mingw-w64-x86-64")
        print("2. Check Nuitka version (4.0+ recommended)")
        print("3. Try building without --onefile for debugging")

if __name__ == "__main__":
    compile_windows_exe()
