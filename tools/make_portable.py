"""Put a portable Python with the app's libraries in python/, so the folder runs on any 64-bit Windows PC.

    python tools/make_portable.py

Run once, on a PC with internet. It downloads the official "embeddable" Python (the same version as the
Python running this script) and installs requirements.txt into it, leaving out the test-only libraries.
Start.bat then uses python/python.exe (with -s, so the PC's own per-user libraries are never mixed in): copy the whole folder to a USB stick and double-click Start.bat on
any PC. Nothing is installed there and no internet is needed. Running it again rebuilds python/.
"""
import io
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "python"
TEST_ONLY = {"pymupdf"}                  # only the tests read PDFs back


def main() -> None:
    if sys.platform != "win32":
        sys.exit("Run this on Windows: the portable Python is a Windows build.")
    v = sys.version_info
    version = f"{v.major}.{v.minor}.{v.micro}"
    url = f"https://www.python.org/ftp/python/{version}/python-{version}-embed-amd64.zip"

    if TARGET.exists():
        shutil.rmtree(TARGET)
    print(f"Downloading Python {version} ...")
    with urllib.request.urlopen(url) as r:
        zipfile.ZipFile(io.BytesIO(r.read())).extractall(TARGET)

    # The ._pth file is the embeddable Python's whole search path: add the libraries, the app folder
    # (one up, so "python -m app.web" works from anywhere) and turn on site.
    pth = TARGET / f"python{v.major}{v.minor}._pth"
    pth.write_text(f"python{v.major}{v.minor}.zip\n.\nLib\\site-packages\n..\nimport site\n", encoding="ascii")

    requirements = [line.strip() for line in (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()
                    if line.strip() and not line.startswith("#")
                    and line.split("=")[0].strip().lower() not in TEST_ONLY]
    print("Installing", ", ".join(requirements), "...")
    subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                    "--target", str(TARGET / "Lib" / "site-packages"),
                    "--only-binary=:all:", "--platform", "win_amd64",
                    "--python-version", f"{v.major}.{v.minor}", "--implementation", "cp",
                    *requirements], check=True)

    check = subprocess.run([str(TARGET / "python.exe"), "-s", "-c", "import flask, openpyxl, xlrd, docx, reportlab, PIL, app.web"],
                           cwd=ROOT.parent, capture_output=True, text=True)
    if check.returncode:
        sys.exit("The portable Python can't load the app:\n" + check.stderr)
    size = sum(f.stat().st_size for f in TARGET.rglob("*") if f.is_file()) / 1e6
    print(f"Done: {TARGET} ({size:.0f} MB). Copy the whole folder to a USB stick and double-click Start.bat.")


if __name__ == "__main__":
    main()
