"""
Cross-platform build script for Excel Compare Tool

This script attempts to build a Windows executable from Linux/WSL,
but recommends running on Windows for best results.
"""

import subprocess
import sys
import os
import platform


def check_dependencies():
    """Check if required build tools are available."""
    issues = []

    # Check Python version
    if sys.version_info < (3, 12):
        issues.append(f"Python {sys.version} is too old. Required: 3.12+")

    # Check PyInstaller
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__} found")
    except ImportError:
        issues.append("PyInstaller not found")

    # Check Nuitka
    try:
        import nuitka
        version = getattr(nuitka, '__version__', 'unknown')
        print(f"✓ Nuitka {version} found")
    except ImportError:
        print("✗ Nuitka not found (optional, for better performance)")

    # Check for Mingw on Linux/WSL
    if platform.system() != 'Windows':
        try:
            result = subprocess.run(['x86_64-w64-mingw32-gcc', '--version'],
                                   capture_output=True)
            if result.returncode != 0:
                print("✗ Mingw-w64 not found (needed for Windows cross-compilation)")
                issues.append("Mingw-w64 not installed (needed for Windows cross-compilation)")
            else:
                print("✓ Mingw-w64 found")
        except FileNotFoundError:
            print("✗ Mingw-w64 not found (needed for Windows cross-compilation)")
            issues.append("Mingw-w64 not installed (needed for Windows cross-compilation)")

    return issues


def build_with_pyinstaller(target_windows=False):
    """Build with PyInstaller."""
    print("\n" + "="*60)
    print("Building with PyInstaller...")
    print("="*60)

    if target_windows and platform.system() != 'Windows':
        print("WARNING: Building for Windows from Linux may not work properly")
        print("Recommendation: Run this script on Windows for best results\n")

    cmd = [
        "uv", "run", "pyinstaller",
        "--onefile",
        "--windowed" if target_windows else "",
        "--name", "ExcelCompare",
    ]

    if platform.system() == 'Windows':
        cmd.extend(["--add-data", "src\\excel_compare;excel_compare"])
    else:
        cmd.extend(["--add-data", "src/excel_compare:excel_compare"])

    cmd.extend([
        "--hidden-import", "tkinterdnd2",
        "--clean",
        "--noconfirm",
        "src/excel_compare/main.py"
    ])

    # Remove empty strings
    cmd = [x for x in cmd if x]

    print(f"Command: {' '.join(cmd)}\n")

    try:
        result = subprocess.run(cmd, check=True, capture_output=False)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        return False


def build_with_nuitka(target_windows=False):
    """Build with Nuitka (experimental for cross-compilation)."""
    print("\n" + "="*60)
    print("Building with Nuitka (experimental)...")
    print("="*60)

    if target_windows and platform.system() != 'Windows':
        print("WARNING: Nuitka cross-compilation is experimental")
        print("Requires Mingw-w64 and proper configuration\n")

    cmd = [
        "uv", "run", "nuitka",
        "--standalone",
        "--onefile",
        "--enable-plugin=tk-inter",
        "--include-package=tkinterdnd2",
        "--assume-yes-for-downloads",
        "--output-dir=dist",
    ]

    if target_windows:
        cmd.append("--mingw64")
        cmd.append("--windows-console-mode=disable")
        cmd.append("--output-filename=ExcelCompare.exe")
    else:
        cmd.append("--output-filename=ExcelCompare")

    cmd.append("src/excel_compare/main.py")

    print(f"Command: {' '.join(cmd)}\n")

    try:
        result = subprocess.run(cmd, check=True, capture_output=False)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        return False


def main():
    """Main build function."""
    print("="*60)
    print("  Excel Compare Tool - Build Script")
    print("="*60)
    print(f"Platform: {platform.system()}")
    print(f"Python: {sys.version}")
    print()

    # Check dependencies
    issues = check_dependencies()
    if issues:
        print("\n⚠ Issues found:")
        for issue in issues:
            print(f"  - {issue}")
        print("\nAttempting to build anyway...\n")

    # Determine target
    target_windows = platform.system() != 'Windows'

    if target_windows:
        print("="*60)
        print("⚠ CROSS-COMPILATION MODE")
        print("="*60)
        print("You're building for Windows from Linux/WSL.")
        print("This may not work perfectly.")
        print("\nFor best results, run this on Windows with:")
        print("  python compile_all.py")
        print()
        response = input("Continue anyway? (y/N): ")
        if response.lower() != 'y':
            print("\n✗ Build cancelled")
            return

    # Try PyInstaller first
    success = build_with_pyinstaller(target_windows)

    # If failed, try Nuitka
    if not success:
        print("\nPyInstaller failed, trying Nuitka...\n")
        success = build_with_nuitka(target_windows)

    # Check results
    print("\n" + "="*60)
    print("BUILD SUMMARY")
    print("="*60)

    exe_name = "ExcelCompare.exe" if target_windows else "ExcelCompare"
    exe_path = os.path.join("dist", exe_name)

    if os.path.exists(exe_path):
        size = os.path.getsize(exe_path)
        size_mb = size / (1024 * 1024)
        print(f"✓ Build successful!")
        print(f"  Location: {exe_path}")
        print(f"  Size: {size_mb:.2f} MB")

        if target_windows:
            print("\n⚠ Note: This was cross-compiled from Linux.")
            print("  Test thoroughly before distribution!")
    else:
        print("✗ Build failed")
        print("\nRecommendations:")
        if target_windows:
            print("  - Run this script on Windows for native compilation")
            print("  - Or copy the project to Windows and run build_windows.bat")
        else:
            print("  - Check error messages above")
            print("  - Ensure all dependencies are installed")

    print("="*60)


if __name__ == "__main__":
    main()
