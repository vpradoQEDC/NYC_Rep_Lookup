import PyInstaller
import os
import shutil
import sys

# Configuration
APP_NAME = "NYC_Representatives_Lookup"
ICON_FILE = "app_icon.ico"  # Optional - will use default if not present
OUTPUT_DIR = "dist"

# Build command
build_command = [
    'pyinstaller',
    '--onefile',
    '--windowed',
    f'--name={APP_NAME}',
    '--add-data=dashboard.html:.',
    f'--distpath={OUTPUT_DIR}',
    '--specpath=build',
    '--workpath=build',
    '--hidden-import=PyQt5',
    '--hidden-import=PyQt5.QtWebEngineWidgets',
    '--hidden-import=pandas',
    '--hidden-import=requests',
    '--hidden-import=bs4',
    'main.py'
]

# Add icon if it exists
if os.path.exists(ICON_FILE):
    build_command.insert(3, f'--icon={ICON_FILE}')

# Run PyInstaller
print("Building Windows .EXE...")
print(f"Command: {' '.join(build_command)}")
os.system(' '.join(build_command))

print("\n" + "="*60)
print(f"✅ Build complete!")
print(f"📦 Output: {OUTPUT_DIR}/{APP_NAME}.exe")
print(f"📂 Size: ~180-200 MB (includes Python runtime)")
print("="*60)
