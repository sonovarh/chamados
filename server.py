# -*- coding: utf-8 -*-
"""
Servidor local do Leitor de Chamados RH Sonova.
Lê automaticamente o arquivo .xlsx da mesma pasta e entrega JSON para o dashboard.
Não depende de pandas, openpyxl ou internet.
"""
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
import datetime as _dt
import json
import re
import traceback
import zipfile
import xml.etree.ElementTree as ET

PORT = 8000
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_NAME = "Chamados RH Sonova (respostas) 05-06-2026.xlsx"


def find_latest_workbook():
    candidates = sorted(
        BASE_DIR.glob("Chamados RH Sonova (respostas)*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if candidates:
        return candidates[0]
    return BASE_DIR / DEFAULT_WORKBOOK_NAME


WORKBOOK_PATH = find_latest_workbook()
WORKBOOK_NAME = WORKBOOK_PATH.name

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "officeRel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

DATE_NUM_FORMAT_IDS = {14, 15, 16, 17, 22, 27, 30, 36, 50, 57}
DATE_TOKENS = ("yy", "yyyy", "dd", "mm", "m/d", "d/m", "h:mm", "hh:mm", "[$-")


def _col_index(cell_ref: str) -> int:
    letters = re.sub(r"[^A-Z]", "", cell_ref.upper())
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _text_from_rich(node):
    parts = []
    for t in node.iterfind(".//main:t", NS):
        parts.append(t.text or "")
    return "".join(parts)


def _read_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return [_text_from_rich(si) for si in root.findall("main:si", NS)]


def _read_date_style_indexes(z):
    if "xl/styles.xml" not in z.namelist():
        return set()
    root = ET.fromstring(z.read("xl/styles.xml"))
    custom_formats = {}
    numfmts = root.find("main:numFmts", NS)
    if numfmts is not None:
        for fmt in numfmts.findall("main:numFmt", NS):
            try:
                fmt_id = int(fmt.attrib.get("numFmtId", "0"))
            except ValueError:
                continue
            code = (fmt.attrib.get("formatCode", "") or "").lower()
            custom_formats[fmt_id] = code
    date_styles = set()
    cellxfs = root.find("main:cellXfs", NS)
    if cellxfs is not None:
        for idx, xf in enumerate(cellxfs.findall("main:xf", NS)):
            try:
                num_id = int(xf.attrib.get("numFmtId", "0"))
            except ValueError:
                num_id = 0
            code = custom_formats.get(num_id, "")
            is_date = num_id in DATE_NUM_FORMAT_IDS or any(tok in code for tok in DATE_TOKENS)
            if is_date:
                date_styles.add(idx)
    return date_styles


def _excel_serial_to_iso(value):
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return value
    # Excel usa 1899-12-30 para compatibilidade com o bug de 1900.
    base = _dt.datetime(1899, 12, 30)
    dt = base + _dt.timedelta(days=serial)
    if abs(serial - int(serial)) < 1e-9:
        return dt.date().isoformat()
    return dt.isoformat(timespec="seconds")


def _read_workbook_sheet_map(z):
    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels = {}
    for rel in rel_root.findall("rel:Relationship", NS):
        target = rel.attrib.get("Target", "")
        if target.startswith("/"):
            # Caminho absoluto no pacote (ex.: "/xl/worksheets/sheet1.xml", gerado por openpyxl e outras ferramentas)
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target
        rels[rel.attrib.get("Id")] = target
    sheets = []
    for sheet in wb_root.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib.get("name", "Planilha")
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rels.get(rid)
        if target:
            sheets.append((name, target))
    return sheets


def _cell_value(cell, shared_strings, date_styles):
    cell_type = cell.attrib.get("t")
    style_id = int(cell.attrib.get("s", "-1")) if cell.attrib.get("s", "-1").isdigit() else -1

    if cell_type == "inlineStr":
        return _text_from_rich(cell)

    v = cell.find("main:v", NS)
    raw = v.text if v is not None else ""
    if raw is None:
        return ""

    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except Exception:
            return raw
    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"
    if cell_type == "str":
        return raw

    if style_id in date_styles and raw != "":
        return _excel_serial_to_iso(raw)

    # Mantém números como texto para não perder CPF e códigos longos.
    return raw


def _read_sheet(z, sheet_path, shared_strings, date_styles):
    root = ET.fromstring(z.read(sheet_path))
    rows = []
    for row in root.findall(".//main:sheetData/main:row", NS):
        values = []
        for cell in row.findall("main:c", NS):
            idx = _col_index(cell.attrib.get("r", "A1"))
            while len(values) <= idx:
                values.append("")
            values[idx] = _cell_value(cell, shared_strings, date_styles)
        rows.append(values)
    if not rows:
        return []
    headers = [str(x).strip() if str(x).strip() else f"COL_{i+1}" for i, x in enumerate(rows[0])]
    data = []
    for values in rows[1:]:
        if not any(str(v).strip() for v in values):
            continue
        item = {}
        for i, h in enumerate(headers):
            item[h] = values[i] if i < len(values) else ""
        data.append(item)
    return data


def read_xlsx_to_json(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path.name}")
    with zipfile.ZipFile(path) as z:
        shared_strings = _read_shared_strings(z)
        date_styles = _read_date_style_indexes(z)
        sheets = {}
        for name, target in _read_workbook_sheet_map(z):
            if target in z.namelist():
                sheets[name] = _read_sheet(z, target, shared_strings, date_styles)
    return sheets


class Handler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/workbook":
            try:
                workbook_path = find_latest_workbook()
                sheets = read_xlsx_to_json(workbook_path)
                total = sum(len(v) for v in sheets.values())
                payload = {
                    "ok": True,
                    "file": workbook_path.name,
                    "totalRows": total,
                    "sheets": sheets,
                    "readAt": _dt.datetime.now().isoformat(timespec="seconds"),
                }
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as exc:
                traceback.print_exc()
                body = json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode("utf-8")
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            return
        return super().do_GET()


def main():
    import webbrowser
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"Servidor iniciado em http://localhost:{PORT}/index.html")
    print(f"Lendo automaticamente o Excel mais recente da pasta: {find_latest_workbook()}")
    webbrowser.open(f"http://localhost:{PORT}/index.html")
    server.serve_forever()


if __name__ == "__main__":
    main()
