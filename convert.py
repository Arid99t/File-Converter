"""
FileConverter - right-click "Convert to > PDF / JPEG / Word" for Windows.

Usage:
    convert.py pdf  <file> [<file> ...]
    convert.py jpeg <file> [<file> ...]
    convert.py docx <file> [<file> ...]
    convert.py --install      (register the right-click menu for the current user)
    convert.py --uninstall    (remove it)
    convert.py --check        (report which conversion engines are available)

Output is written next to the source file. Existing files are never overwritten:
"report.pdf" becomes "report (1).pdf" if needed.
"""
import os
import sys
import shutil
import subprocess
import tempfile
import time
import traceback

APP_NAME = "FileConverter"
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), APP_NAME)
LOG_FILE = os.path.join(DATA_DIR, "log.txt")

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".heic", ".heif"}
PDF_EXT = {".pdf"}
WORD_EXT = {".docx", ".doc", ".rtf", ".odt"}
SLIDES_EXT = {".pptx", ".ppt", ".odp"}
SHEET_EXT = {".xlsx", ".xls", ".ods", ".csv"}
TEXT_EXT = {".txt", ".md"}

# Which targets appear in the menu for each source type (never "convert to itself").
TARGETS = {}
for e in IMAGE_EXT:
    TARGETS[e] = ["pdf", "jpeg", "docx"] if e not in {".jpg", ".jpeg"} else ["pdf", "docx"]
for e in PDF_EXT:
    TARGETS[e] = ["jpeg", "docx"]
for e in WORD_EXT:
    TARGETS[e] = ["pdf", "jpeg"] + ([] if e == ".docx" else ["docx"])
for e in SLIDES_EXT:
    TARGETS[e] = ["pdf", "jpeg"]
for e in SHEET_EXT:
    TARGETS[e] = ["pdf", "jpeg"]
for e in TEXT_EXT:
    TARGETS[e] = ["pdf", "jpeg", "docx"]

LABELS = {"pdf": "PDF", "jpeg": "JPEG image", "docx": "Word document"}
OUT_EXT = {"pdf": ".pdf", "jpeg": ".jpg", "docx": ".docx"}


class ConversionError(Exception):
    pass


# --------------------------------------------------------------------------- helpers

def log(msg):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + msg + "\n")
    except Exception:
        pass


def unique_path(path):
    """Return path, or 'name (1).ext', 'name (2).ext'... if it exists."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def out_path_for(src, target):
    base = os.path.splitext(src)[0]
    return unique_path(base + OUT_EXT[target])


def show_error(title, text):
    if os.name == "nt":
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, text, title, 0x10 | 0x40000)  # MB_ICONERROR | MB_TOPMOST
    else:
        print(f"[{title}] {text}", file=sys.stderr)


class FileLock:
    """Serialises Office automation when many files are converted at once
    (Windows launches one process per selected file)."""

    def __init__(self, name):
        os.makedirs(DATA_DIR, exist_ok=True)
        self.path = os.path.join(DATA_DIR, name + ".lock")
        self.fh = None

    def __enter__(self):
        self.fh = open(self.path, "a+")
        if os.name == "nt":
            import msvcrt
            deadline = time.time() + 600
            while True:
                try:
                    self.fh.seek(0)
                    msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
                    return self
                except OSError:
                    if time.time() > deadline:
                        raise ConversionError("Timed out waiting for another conversion to finish.")
                    time.sleep(0.3)
        else:
            import fcntl
            fcntl.flock(self.fh, fcntl.LOCK_EX)
        return self

    def __exit__(self, *a):
        try:
            if os.name == "nt":
                import msvcrt
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.fh, fcntl.LOCK_UN)
        finally:
            self.fh.close()


# --------------------------------------------------------------------------- engines

def find_soffice():
    cands = [shutil.which("soffice"), shutil.which("soffice.exe")]
    for root in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if root:
            cands.append(os.path.join(root, "LibreOffice", "program", "soffice.exe"))
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def office_available(progid):
    if os.name != "nt":
        return False
    import winreg
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID"))
        return True
    except OSError:
        return False


def soffice_convert(src, fmt, out):
    soffice = find_soffice()
    if not soffice:
        return False
    with tempfile.TemporaryDirectory() as td:
        profile = "file:///" + os.path.join(td, "profile").replace("\\", "/")
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless",
               "--convert-to", fmt, "--outdir", td, src]
        flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
        subprocess.run(cmd, check=False, timeout=300, creationflags=flags,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        produced = os.path.join(td, os.path.splitext(os.path.basename(src))[0] + "." + fmt.split(":")[0])
        if not os.path.isfile(produced):
            raise ConversionError("LibreOffice could not convert this file.")
        shutil.move(produced, out)
    return True


def com_app(progid):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    return win32com.client.DispatchEx(progid)


def word_convert(src, out, fmt_code):
    """fmt_code: 17 = PDF, 16 = DOCX. Returns True if Word did it."""
    if not office_available("Word.Application"):
        return False
    with FileLock("word"):
        word = com_app("Word.Application")
        try:
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(src), ConfirmConversions=False,
                                      ReadOnly=True, AddToRecentFiles=False)
            try:
                if fmt_code == 17:
                    doc.ExportAsFixedFormat(os.path.abspath(out), 17)
                else:
                    doc.SaveAs2(os.path.abspath(out), FileFormat=fmt_code)
            finally:
                doc.Close(0)
        finally:
            word.Quit()
    return True


def excel_to_pdf(src, out):
    if not office_available("Excel.Application"):
        return False
    with FileLock("excel"):
        xl = com_app("Excel.Application")
        try:
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(os.path.abspath(src), ReadOnly=True, AddToMru=False)
            try:
                wb.ExportAsFixedFormat(0, os.path.abspath(out))
            finally:
                wb.Close(False)
        finally:
            xl.Quit()
    return True


def powerpoint_to_pdf(src, out):
    if not office_available("PowerPoint.Application"):
        return False
    with FileLock("powerpoint"):
        pp = com_app("PowerPoint.Application")
        try:
            pres = pp.Presentations.Open(os.path.abspath(src), ReadOnly=True, Untitled=False, WithWindow=False)
            try:
                pres.SaveAs(os.path.abspath(out), 32)  # ppSaveAsPDF
            finally:
                pres.Close()
        finally:
            pp.Quit()
    return True


NO_OFFICE = ("This needs Microsoft Office or LibreOffice to convert, and neither was found.\n\n"
             "Install LibreOffice (free) from libreoffice.org, then try again.")


def office_to_pdf(src, out):
    ext = os.path.splitext(src)[1].lower()
    if ext in WORD_EXT and word_convert(src, out, 17):
        return
    if ext in SHEET_EXT and excel_to_pdf(src, out):
        return
    if ext in SLIDES_EXT and powerpoint_to_pdf(src, out):
        return
    if soffice_convert(src, "pdf", out):
        return
    raise ConversionError(NO_OFFICE)


# --------------------------------------------------------------------------- image helpers

def open_image(src):
    from PIL import Image
    if os.path.splitext(src)[1].lower() in {".heic", ".heif"}:
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except ImportError:
            raise ConversionError("HEIC support is not installed (pillow-heif).")
    img = Image.open(src)
    if getattr(img, "n_frames", 1) > 1:
        return img  # multi-page (e.g. TIFF): keep all frames
    try:
        from PIL import ImageOps
        img = ImageOps.exif_transpose(img)  # respect phone rotation
    except Exception:
        pass
    return img


def flatten_rgb(img):
    from PIL import Image
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        return bg
    return img.convert("RGB")


def image_frames(src):
    from PIL import ImageSequence
    img = open_image(src)
    if os.path.splitext(src)[1].lower() in {".tif", ".tiff"} and getattr(img, "n_frames", 1) > 1:
        return [flatten_rgb(f.copy()) for f in ImageSequence.Iterator(img)]
    return [flatten_rgb(img)]


def pdf_to_jpegs(pdf, src_for_naming, dpi=200):
    """Render each page. One page -> name.jpg; several -> folder 'name (JPEG)'."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    doc = fitz.open(pdf)
    try:
        base = os.path.splitext(src_for_naming)[0]
        if doc.page_count == 1:
            out = unique_path(base + ".jpg")
            doc[0].get_pixmap(dpi=dpi).save(out, jpg_quality=92)
            return out
        folder = unique_path(base + " (JPEG)")
        os.makedirs(folder)
        name = os.path.basename(base)
        width = max(2, len(str(doc.page_count)))
        for i, page in enumerate(doc, 1):
            page.get_pixmap(dpi=dpi).save(os.path.join(folder, f"{name}_p{i:0{width}d}.jpg"), jpg_quality=92)
        return folder
    finally:
        doc.close()


def text_to_pdf(src, out):
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    import html
    with open(src, "r", encoding="utf-8-sig", errors="replace") as f:
        text = f.read()
    body = "<pre style='font-family: Consolas, monospace; font-size: 10pt; white-space: pre-wrap;'>" \
           + html.escape(text) + "</pre>"
    story = fitz.Story(html=body)
    writer = fitz.DocumentWriter(out)
    mediabox = fitz.paper_rect("a4")
    where = mediabox + (54, 54, -54, -54)
    more = 1
    while more:
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()


def text_to_docx(src, out):
    from docx import Document
    from docx.shared import Pt
    with open(src, "r", encoding="utf-8-sig", errors="replace") as f:
        lines = f.read().splitlines()
    d = Document()
    style = d.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    for line in lines:
        d.add_paragraph(line)
    d.save(out)


def images_to_docx(frames, out):
    from docx import Document
    from docx.shared import Inches
    d = Document()
    sec = d.sections[0]
    max_w = (sec.page_width - sec.left_margin - sec.right_margin) / 914400  # inches
    max_h = (sec.page_height - sec.top_margin - sec.bottom_margin) / 914400 - 0.3
    with tempfile.TemporaryDirectory() as td:
        for i, fr in enumerate(frames):
            p = os.path.join(td, f"f{i}.jpg")
            fr.save(p, "JPEG", quality=95)
            w, h = fr.size
            scale = min(max_w / w, max_h / h)
            if i:
                d.add_page_break()
            d.add_picture(p, width=Inches(w * scale))
        d.save(out)


# --------------------------------------------------------------------------- dispatcher

def convert(src, target):
    src = os.path.abspath(src)
    ext = os.path.splitext(src)[1].lower()
    if target not in TARGETS.get(ext, []):
        raise ConversionError(f"Can't convert {ext} files to {LABELS.get(target, target)}.")

    # ---- images
    if ext in IMAGE_EXT:
        frames = image_frames(src)
        out = out_path_for(src, target)
        if target == "pdf":
            frames[0].save(out, "PDF", resolution=150.0, save_all=True, append_images=frames[1:])
        elif target == "jpeg":
            if len(frames) == 1:
                frames[0].save(out, "JPEG", quality=92)
            else:
                with tempfile.TemporaryDirectory() as td:
                    tmp = os.path.join(td, "x.pdf")
                    frames[0].save(tmp, "PDF", resolution=150.0, save_all=True, append_images=frames[1:])
                    out = pdf_to_jpegs(tmp, src)
        else:
            images_to_docx(frames, out)
        return out

    # ---- PDF
    if ext in PDF_EXT:
        if target == "jpeg":
            return pdf_to_jpegs(src, src)
        out = out_path_for(src, "docx")
        from pdf2docx import Converter
        cv = Converter(src)
        try:
            cv.convert(out)
        finally:
            cv.close()
        return out

    # ---- text
    if ext in TEXT_EXT:
        if target == "docx":
            out = out_path_for(src, "docx")
            text_to_docx(src, out)
            return out
        if target == "pdf":
            out = out_path_for(src, "pdf")
            text_to_pdf(src, out)
            return out
        with tempfile.TemporaryDirectory() as td:
            tmp = os.path.join(td, "x.pdf")
            text_to_pdf(src, tmp)
            return pdf_to_jpegs(tmp, src)

    # ---- Office documents
    if target == "docx":  # .doc / .rtf / .odt -> .docx
        out = out_path_for(src, "docx")
        if word_convert(src, out, 16) or soffice_convert(src, "docx", out):
            return out
        raise ConversionError(NO_OFFICE)
    if target == "pdf":
        out = out_path_for(src, "pdf")
        office_to_pdf(src, out)
        return out
    # jpeg: go via a temporary PDF
    with tempfile.TemporaryDirectory() as td:
        tmp = os.path.join(td, "x.pdf")
        office_to_pdf(src, tmp)
        return pdf_to_jpegs(tmp, src)


# --------------------------------------------------------------------------- install / uninstall

MENU_KEY = "FileConverter"
CLASSES = r"Software\Classes\SystemFileAssociations"


def _pythonw():
    exe = sys.executable
    w = os.path.join(os.path.dirname(exe), "pythonw.exe")
    return w if os.path.isfile(w) else exe


def install():
    import winreg
    script = os.path.abspath(__file__)
    pyw = _pythonw()
    icon = os.path.join(os.path.dirname(script), "icon.ico")
    for ext, targets in sorted(TARGETS.items()):
        root = rf"{CLASSES}\{ext}\shell\{MENU_KEY}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, root) as k:
            winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, "Convert to")
            winreg.SetValueEx(k, "SubCommands", 0, winreg.REG_SZ, "")
            if os.path.isfile(icon):
                winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon)
        for i, t in enumerate(targets, 1):
            sub = rf"{root}\shell\{i:02d}_{t}"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub) as k:
                winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, LABELS[t])
                winreg.SetValueEx(k, "MultiSelectModel", 0, winreg.REG_SZ, "Player")
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub + r"\command") as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, f'"{pyw}" "{script}" {t} "%1"')
    print(f"Right-click menu registered for {len(TARGETS)} file types.")


def _delete_tree(root, path):
    import winreg
    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS) as k:
            while True:
                try:
                    child = winreg.EnumKey(k, 0)
                except OSError:
                    break
                _delete_tree(root, path + "\\" + child)
        winreg.DeleteKey(root, path)
    except FileNotFoundError:
        pass


def uninstall():
    import winreg
    for ext in TARGETS:
        _delete_tree(winreg.HKEY_CURRENT_USER, rf"{CLASSES}\{ext}\shell\{MENU_KEY}")
    print("Right-click menu removed.")


def check():
    mods = {}
    for m in ("PIL", "pillow_heif", "pymupdf", "pdf2docx", "docx", "win32com"):
        try:
            __import__(m)
            mods[m] = "ok"
        except Exception as e:
            mods[m] = f"MISSING ({e.__class__.__name__})"
    for m, s in mods.items():
        print(f"  python module {m:12s} {s}")
    print(f"  Microsoft Word       {'found' if office_available('Word.Application') else 'not found'}")
    print(f"  Microsoft Excel      {'found' if office_available('Excel.Application') else 'not found'}")
    print(f"  Microsoft PowerPoint {'found' if office_available('PowerPoint.Application') else 'not found'}")
    print(f"  LibreOffice          {find_soffice() or 'not found'}")


# --------------------------------------------------------------------------- main

def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 1
    cmd = argv[1]
    if cmd == "--install":
        install(); return 0
    if cmd == "--uninstall":
        uninstall(); return 0
    if cmd == "--check":
        check(); return 0
    target = cmd.lower()
    if target == "jpg":
        target = "jpeg"
    failures = []
    for src in argv[2:]:
        try:
            out = convert(src, target)
            log(f"OK   {src} -> {out}")
        except ConversionError as e:
            failures.append((src, str(e)))
            log(f"FAIL {src}: {e}")
        except Exception as e:
            failures.append((src, f"{e.__class__.__name__}: {e}"))
            log(f"FAIL {src}:\n{traceback.format_exc()}")
    if failures:
        msg = "\n\n".join(f"{os.path.basename(s)}\n{m}" for s, m in failures)
        show_error("FileConverter - conversion failed", msg + f"\n\nDetails: {LOG_FILE}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
