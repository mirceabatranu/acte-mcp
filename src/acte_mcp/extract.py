"""Citirea continutului unui act, ca Claude sa-l poata prelucra (ex. pentru import in SAGA).

PDF cu text -> text; PDF scanat -> imagini ale paginilor; poze -> imagine;
Excel/CSV -> tabel text; XML (e-Factura), TXT, MT940 -> text; DOCX -> text; ZIP -> lista.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from pathlib import Path

from mcp.server.fastmcp import Image

MAX_CHARS = 60_000
MAX_IMAGE_BYTES = 1_500_000
TEXT_EXT = {".txt", ".xml", ".json", ".html", ".htm", ".sta", ".mt940", ".940", ".log", ".md"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff", ".heic"}


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1250", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def _pages(spec: str | None, total: int, default_max: int) -> list[int]:
    if not spec:
        return list(range(min(total, default_max)))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out += range(int(a) - 1, min(int(b), total))
        elif part:
            out.append(int(part) - 1)
    return [p for p in out if 0 <= p < total]


def _image(pil) -> Image:
    from PIL import Image as PILImage

    img = pil.convert("RGB") if pil.mode not in ("RGB", "L") else pil
    max_side = 2000
    if max(img.size) > max_side:
        img.thumbnail((max_side, max_side), PILImage.LANCZOS)
    quality = 85
    while True:
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality)
        if buf.tell() <= MAX_IMAGE_BYTES or quality <= 40:
            return Image(data=buf.getvalue(), format="jpeg")
        quality -= 15


def read_pdf(path: Path, pagini: str | None, ca_imagine: bool) -> list:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    total = len(reader.pages)
    idx = _pages(pagini, total, 20)
    texts = []
    for i in idx:
        try:
            texts.append((i, reader.pages[i].extract_text() or ""))
        except Exception:  # noqa: BLE001
            texts.append((i, ""))
    chars = sum(len(t.strip()) for _, t in texts)
    scanned = chars < 20 * max(1, len(texts))
    out: list = [f"{path.name}: PDF, {total} pagini" + (" (scanat: trimit imaginile paginilor)" if scanned else "")]
    if not scanned and not ca_imagine:
        body = "\n\n".join(f"--- pagina {i + 1} ---\n{t}" for i, t in texts)
        out.append(body[:MAX_CHARS] + ("\n... (trunchiat; cere alte pagini)" if len(body) > MAX_CHARS else ""))
        return out
    import pypdfium2 as pdfium

    doc = pdfium.PdfDocument(str(path))
    for i in _pages(pagini, total, 5):
        out.append(f"--- pagina {i + 1} ---")
        out.append(_image(doc[i].render(scale=2.0).to_pil()))
    doc.close()
    if total > 5 and not pagini:
        out.append(f"(am trimis primele 5 pagini din {total}; cere restul cu pagini='6-10')")
    return out


def read_excel(path: Path) -> list:
    rows_out: list[str] = []
    if path.suffix.lower() == ".xls":
        import xlrd

        book = xlrd.open_workbook(str(path))
        for sh in book.sheets():
            rows_out.append(f"=== foaia {sh.name} ({sh.nrows} randuri) ===")
            for r in range(min(sh.nrows, 1000)):
                rows_out.append("\t".join(str(v) for v in sh.row_values(r)))
    else:
        import openpyxl

        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows_out.append(f"=== foaia {ws.title} ===")
            for n, row in enumerate(ws.iter_rows(values_only=True)):
                if n >= 1000:
                    rows_out.append("... (trunchiat la 1000 de randuri)")
                    break
                if any(v is not None for v in row):
                    rows_out.append("\t".join("" if v is None else str(v) for v in row))
        wb.close()
    text = "\n".join(rows_out)
    return [f"{path.name}: tabel", text[:MAX_CHARS]]


def read_csv(path: Path) -> list:
    text = _decode(path.read_bytes())
    try:
        dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t|")
        rows = list(csv.reader(io.StringIO(text), dialect))
        text = "\n".join("\t".join(r) for r in rows[:2000])
    except csv.Error:
        pass
    return [f"{path.name}: CSV", text[:MAX_CHARS]]


def read_docx(path: Path) -> list:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    xml = re.sub(r"</w:p>", "\n", xml)
    text = re.sub(r"<[^>]+>", "", xml)
    return [f"{path.name}: document Word", text[:MAX_CHARS]]


def read(path: Path, pagini: str | None = None, ca_imagine: bool = False) -> list:
    ext = path.suffix.lower()
    if ext == ".pdf":
        try:
            return read_pdf(path, pagini, ca_imagine)
        except Exception as exc:  # noqa: BLE001 - PDF stricat: incercam macar imaginile
            try:
                import pypdfium2 as pdfium

                doc = pdfium.PdfDocument(str(path))
                out: list = [f"{path.name}: PDF greu de citit ca text ({exc}); trimit imaginile paginilor"]
                for i in _pages(pagini, len(doc), 5):
                    out += [f"--- pagina {i + 1} ---", _image(doc[i].render(scale=2.0).to_pil())]
                doc.close()
                return out
            except Exception as exc2:  # noqa: BLE001
                return [f"{path.name}: nu pot deschide PDF-ul ({exc2})"]
    if ext in IMAGE_EXT:
        from PIL import Image as PILImage

        try:
            return [f"{path.name}: imagine", _image(PILImage.open(path))]
        except Exception as exc:  # noqa: BLE001 - ex. HEIC fara suport
            return [f"{path.name}: nu pot deschide imaginea ({exc})"]
    if ext in (".xlsx", ".xlsm", ".xls"):
        return read_excel(path)
    if ext == ".csv":
        return read_csv(path)
    if ext == ".docx":
        return read_docx(path)
    if ext == ".zip":
        with zipfile.ZipFile(path) as z:
            names = [f"{i.filename} ({i.file_size} octeti)" for i in z.infolist()]
        return [f"{path.name}: arhiva ZIP cu {len(names)} fisiere:\n" + "\n".join(names[:300])]
    if ext in TEXT_EXT or path.stat().st_size < 200_000:
        data = path.read_bytes()
        if b"\x00" not in data[:4000]:
            text = _decode(data)
            return [f"{path.name}: text", text[:MAX_CHARS]]
    return [f"{path.name}: format {ext or 'necunoscut'}, {path.stat().st_size} octeti — nu il pot citi direct"]
