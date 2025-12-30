"""Create a minimal .xlsm test file with an embedded image and set it read-only.

Usage:
    python tools/create_test_xlsm.py

The script will create: input/test_sample.xlsm (overwrites if present)
It embeds a small PNG image into column D and marks the file read-only.
"""
import os
import base64
from pathlib import Path
import tempfile
import time

try:
    import win32com.client
except Exception as e:
    print("win32com is required to run this script. Install pywin32 and try again.")
    raise

DATA_DIR = Path(__file__).resolve().parents[1] / 'input'
DATA_DIR.mkdir(parents=True, exist_ok=True)
OUT_FILE = DATA_DIR / 'test_sample.xlsm'

# Tiny 1x1 red PNG (base64)
PNG_B64 = (
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII='
)

print(f"Creating test file: {OUT_FILE}")

# Write image to temp file
tmp_img = tempfile.gettempdir() + os.sep + f"test_img_{int(time.time())}.png"
with open(tmp_img, 'wb') as f:
    f.write(base64.b64decode(PNG_B64))

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Add()
try:
    sht = wb.Sheets(1)
    sht.Name = 'SpecTest'

    # Headers
    sht.Cells(1, 1).Value = 'ID'
    sht.Cells(1, 4).Value = 'Image'

    # Sample rows
    for i in range(2, 7):
        sht.Cells(i, 1).Value = f'EL-{i-1:02d}'

    # Insert image into column D anchored at row 2
    left = sht.Cells(2, 4).Left
    top = sht.Rows(2).Top
    # Width and height set so it's visible in print
    sht.Shapes.AddPicture(tmp_img, False, True, left, top, 300, 200)

    # Save as xlsm (FileFormat=52)
    out_path = str(OUT_FILE)
    if OUT_FILE.exists():
        OUT_FILE.unlink()
    wb.SaveAs(out_path, FileFormat=52)
    wb.Close(SaveChanges=False)
    excel.Quit()

    # Mark file read-only
    try:
        os.chmod(out_path, 0o444)
    except Exception:
        # Windows alternative: set readonly attribute
        try:
            import ctypes
            FILE_ATTRIBUTE_READONLY = 0x01
            attrs = ctypes.windll.kernel32.GetFileAttributesW(out_path)
            if attrs != -1:
                ctypes.windll.kernel32.SetFileAttributesW(out_path, attrs | FILE_ATTRIBUTE_READONLY)
        except Exception:
            pass

    print(f"Created read-only .xlsm test file: {out_path}")
finally:
    # Cleanup temp image
    try:
        os.remove(tmp_img)
    except Exception:
        pass
