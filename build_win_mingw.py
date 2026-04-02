#!/usr/bin/env python3
"""
Windows cross-compilation using Mingw-w64

This script attempts true cross-compilation to Windows EXE.
"""

import subprocess
import sys
import os
import shutil

def compile_windows_exe():
    """Compile Windows exe using Nuitka with Mingw-w64."""

    # Set environment for cross-compilation
    env = os.environ.copy()

    # Force Mingw-w64 compilers
    env['CC'] = 'x86_64-w64-mingw32-gcc'
    env['CXX'] = 'x86_64-w64-mingw32-g++'
    env['LD'] = 'x86_64-w64-mingw32-gcc'
    env['AR'] = 'x86_64-w64-mingw32-ar'
    env['RANLIB'] = 'x86_64-w64-mingw32-ranlib'

    # Disable LTO for compatibility
    env['NUITKA_LTO_FLAGS'] = ''

    print("=" * 60)
    print("Cross-compiling Windows EXE using Mingw-w64")
    print("=" * 60)
    print()

    # Clean previous builds
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')

    # Nuitka command for Mingw-w64 cross-compilation
    cmd = [
        'uv', 'run', 'nuitka',
        '--standalone',
        '--onefile',
        '--enable-plugin=tk-inter',
        '--include-package=tkinterdnd2',
        '--assume-yes-for-downloads',
        '--output-dir=dist',
        '--output-filename=ExcelCompare.exe',
        '--mingw64',
        '--lto=no',  # Disable link-time optimization for compatibility
        '--jobs=4',  # Use 4 parallel jobs
        '--remove-output',
        'src/excel_compare/main.py'
    ]

    print("Environment settings:")
    print(f"  CC: {env.get('CC')}")
    print(f"  CXX: {env.get('CXX')}")
    print()
    print("Nuitka command:")
    print(' '.join(cmd))
    print()
    print("=" * 60)
    print()

    try:
        # Run the compilation with custom environment
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
            print(f"✓ Executable created: {exe_path}")
            print(f"✓ File size: {size_mb:.2f} MB")

            # Check file type
            result = subprocess.run(['file', exe_path], capture_output=True, text=True)
            file_info = result.stdout.strip()
            print(f"✓ File type: {file_info}")

            if 'PE32' in file_info or 'MS-DOS' in file_info:
                print("✓ SUCCESS: This is a Windows executable!")
            else:
                print("⚠ WARNING: This may not be a Windows executable")
                print("   File appears to be: " + file_info.split(':')[1].strip())
        else:
            print("⚠ ExcelCompare.exe not found in dist/")
            # List what's in dist
            if os.path.exists('dist'):
                print("Contents of dist/:")
                for item in os.listdir('dist'):
                    item_path = os.path.join('dist', item)
                    if os.path.isfile(item_path):
                        result = subprocess.run(['file', item_path], capture_output=True, text=True)
                        print(f"  {item}: {result.stdout.strip()}")

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
        print("2. Check that x86_64-w64-mingw32-gcc works:")
        print("   x86_64-w64-mingw32-gcc --version")
        print("3. For full debugging, remove --onefile flag")

if __name__ == "__main__":
    compile_windows_exe()
