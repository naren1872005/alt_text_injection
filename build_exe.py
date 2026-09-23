"""
Build script to package PDF Alt-Text Extractor into a standalone Windows Executable.
"""
import os
import subprocess
import sys

def build():
    print("=" * 60)
    print("Building PDF Alt-Text Extractor Standalone Executable...")
    print("=" * 60)

    # Command arguments for PyInstaller
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--name", "PDF_AltText_Tool",
        "--add-data", f"static{os.pathsep}static",
        "--hidden-import", "uvicorn.logging",
        "--hidden-import", "uvicorn.loops",
        "--hidden-import", "uvicorn.loops.auto",
        "--hidden-import", "uvicorn.protocols",
        "--hidden-import", "uvicorn.protocols.http",
        "--hidden-import", "uvicorn.protocols.http.auto",
        "--hidden-import", "uvicorn.protocols.websockets",
        "--hidden-import", "uvicorn.protocols.websockets.auto",
        "--hidden-import", "uvicorn.lifespan",
        "--hidden-import", "uvicorn.lifespan.on",
        "--hidden-import", "pymupdf",
        "--hidden-import", "fitz",
        "--hidden-import", "cv2",
        "--hidden-import", "PIL",
        "--hidden-import", "openpyxl",
        "--hidden-import", "xlrd",
        "--hidden-import", "imagehash",
        "app.py"
    ]

    print("Running command:\n", " ".join(cmd))
    result = subprocess.run(cmd)

    if result.returncode == 0:
        print("\n" + "=" * 60)
        print("SUCCESS! Build complete.")
        print(f"Output folder: {os.path.abspath('dist/PDF_AltText_Tool')}")
        print("You can zip the 'dist/PDF_AltText_Tool' folder and share it with your office mates!")
        print("=" * 60)
    else:
        print("\nBuild failed. Please check the error log above.")

if __name__ == "__main__":
    build()
