#!/usr/bin/env python3
"""
Final Windows cross-compilation script using fixed GCC wrapper
"""

import subprocess
import sys
import os
import shutil

def compile_windows_exe():
    """Compile Windows exe using Nuitka with fixed Mingw-w64."""

    # Set environment for cross-compilation
    env = os.environ.copy()

    # Use the fixed GCC wrapper
    env['CC'] = '/usr/local/bin/x86_64-w64-mingw32-gcc-fixed'
    env['CXX'] = '/usr/local/bin/x86_64-w64-mingw32-gcc-fixed'
    env['LD'] = '/usr/local/bin/x86_64-w64-mingw32-gcc-fixed'
    env['AR'] = 'x86_64-w64-mingw32-ar'
    env['RANLIB'] = 'x86_64-w64-mingw32-ranlib'

    # Disable LTO for compatibility
    env['NUITKA_LTO_FLAGS'] = ''

    print("=" * 60)
    print("Cross-compiling Windows EXE (Final Attempt)")
    print("=" * 60)
    print()

    # Clean previous builds
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')

    # Nuitka command
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
        '--lto=no',
        '--jobs=4',
        '--remove-output',
        'src/excel_compare/main.py'
    ]

    print("Environment settings:")
    print(f"  CC: {env.get('CC')}")
    print(f"  CXX: {env.get('CXX')}")
    print()

    # Verify compiler works
    print("Verifying Mingw compiler...")
    try:
        result = subprocess.run([env['CC'], '--version'], capture_output=True, text=True)
        print(f"✓ Compiler: {result.stdout.strip().splitlines()[0]}")
    except Exception as e:
        print(f"✗ Compiler test failed: {e}")
        return

    print()
    print("Starting Nuitka compilation...")
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
            print(f"✓ Executable created: {exe_path}")
            print(f"✓ File size: {size_mb:.2f} MB")

            # Check file type
            result = subprocess.run(['file', exe_path], capture_output=True, text=True)
            file_info = result.stdout.strip()
            print(f"✓ File type: {file_info}")

            if 'PE32' in file_info or 'MS-DOS' in file_info:
                print("✓✓✓ SUCCESS! This is a Windows PE executable!")
                print()
                print("The executable is ready to run on Windows systems.")
                print("No Python installation required on target machine.")
            else:
                print("⚠ WARNING: This may not be a Windows executable")
                print(f"   File appears to be: {file_info.split(':', 1)[1].strip()}")
        else:
            print("⚠ ExcelCompare.exe not found in dist/")
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

if __name__ == "__main__":
    compile_windows_exe()
