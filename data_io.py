from __future__ import annotations

import posixpath
import re
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKBOOK_NORMALIZED_NAME = "lista final 298 lojas 1.xlsx"

def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return " ".join(text.lower().split())

def normalize_filename(value: str) -> str:
    return " ".join(value.replace("_", " ").lower().split())

def is_backup(relative_path: str) -> bool:
    return any(part.strip().lower().startswith(("backup", "bakup")) for part in relative_path.replace("\\", "/").split("/"))

def find_store_workbook() -> Path | None:
    for path in ROOT.iterdir():
        if path.is_file() and path.suffix.lower() == ".xlsx" and normalize_filename(path.name) == WORKBOOK_NORMALIZED_NAME:
            return path
    return None

def list_report_files() -> list[dict[str, str]]:
    files: list[dict[str, str]] = []
    for folder_name in ("1.09", "9.09"):
        folder = ROOT / folder_name
        if not folder.is_dir():
            continue
        for path in folder.rglob("*"):
            if not path.is_file() or path.suffix.lower() != ".csv":
                continue
            relative = path.relative_to(ROOT).as_posix()
            if is_backup(relative):
                continue
            files.append({"name": path.name, "relativePath": relative})

    workbook = find_store_workbook()
    if workbook:
        files.append({"name": workbook.name, "relativePath": workbook.relative_to(ROOT).as_posix()})
    return sorted(files, key=lambda item: item["relativePath"].lower())

def column_index(cell_reference: str) -> int:
    letters = re.match(r"[A-Z]+", cell_reference.upper())
    if not letters:
        return 0
    value = 0
    for char in letters.group(0):
        value = value * 26 + ord(char) - ord("A") + 1
    return value - 1

def load_shared_strings(book: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(book.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.itertext()) for node in root.findall(f"{{{MAIN_NS}}}si")]

def workbook_sheet_paths(book: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(book.read("xl/workbook.xml"))
    relations = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    targets = {item.attrib["Id"]: item.attrib["Target"] for item in relations.findall(f"{{{PKG_REL_NS}}}Relationship")}
    result: dict[str, str] = {}
    sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
    if sheets is None:
        return result
    for sheet in sheets:
        rel_id = sheet.attrib.get(f"{{{REL_NS}}}id")
        target = targets.get(rel_id or "")
        if not target:
            continue
        if target.startswith("/"):
            target = target.lstrip("/")
        else:
            target = posixpath.normpath(posixpath.join("xl", target))
        result[sheet.attrib.get("name", "")] = target
    return result

def sheet_rows(book: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> list[dict[int, str]]:
    root = ET.fromstring(book.read(sheet_path))
    sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return []
    rows: list[dict[int, str]] = []
    for row_node in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        row: dict[int, str] = {}
        for cell in row_node.findall(f"{{{MAIN_NS}}}c"):
            ref = cell.attrib.get("r", "")
            index = column_index(ref)
            cell_type = cell.attrib.get("t", "")
            if cell_type == "inlineStr":
                value = "".join(cell.find(f"{{{MAIN_NS}}}is").itertext()) if cell.find(f"{{{MAIN_NS}}}is") is not None else ""
            else:
                value_node = cell.find(f"{{{MAIN_NS}}}v")
                value = value_node.text if value_node is not None and value_node.text is not None else ""
                if cell_type == "s" and value:
                    value = shared_strings[int(value)]
            row[index] = value
        rows.append(row)
    return rows

def read_store_sheet(book: zipfile.ZipFile, sheet_path: str, shared_strings: list[str], sheet_name: str) -> dict[str, dict[str, object]]:
    rows = sheet_rows(book, sheet_path, shared_strings)
    header_row_index = -1
    columns: dict[str, int] = {}
    for index, row in enumerate(rows[:80]):
        normalized = {normalize_text(value): col for col, value in row.items()}
        if "idfilial" in normalized and "shop ou rua" in normalized and "unid. log." in normalized:
            header_row_index = index
            columns = {
                "id": normalized["idfilial"],
                "shop": normalized["shop ou rua"],
                "cd": normalized["unid. log."],
            }
            break
    if header_row_index < 0:
        raise ValueError(f"Cabecalho de lojas nao encontrado na aba {sheet_name}.")

    result: dict[str, dict[str, object]] = {}
    for row in rows[header_row_index + 1 :]:
        id_match = re.search(r"\d{1,6}", row.get(columns["id"], ""))
        cd_match = re.match(r"\s*(\d{3,5})", row.get(columns["cd"], ""))
        if not id_match or not cd_match:
            continue
        store_id = str(int(id_match.group(0)))
        cd_code = str(int(cd_match.group(1))).zfill(4)
        result[store_id] = {
            "cd": cd_code,
            "isShop": normalize_text(row.get(columns["shop"], "")) == "shopping",
        }
    return result

def read_store_workbook(path: Path | None = None) -> dict[str, object]:
    workbook_path = path or find_store_workbook()
    if workbook_path is None or not workbook_path.is_file():
        raise FileNotFoundError("Lista final_298 lojas 1.xlsx nao encontrada na pasta principal.")
    with zipfile.ZipFile(workbook_path) as book:
        sheet_paths = workbook_sheet_paths(book)
        shared_strings = load_shared_strings(book)
        transp_path = next((path for name, path in sheet_paths.items() if normalize_text(name) == "transp"), None)
        custo_path = next((path for name, path in sheet_paths.items() if normalize_text(name) == "custo frete"), None)
        if not transp_path or not custo_path:
            raise ValueError("A planilha precisa conter as abas Transp e Custo Frete.")
        transp = read_store_sheet(book, transp_path, shared_strings, "Transp")
        custo = read_store_sheet(book, custo_path, shared_strings, "Custo Frete")

    if len(transp) != 298 or len(custo) != 298:
        raise ValueError(f"A planilha precisa conter 298 lojas por aba (Transp: {len(transp)}; Custo Frete: {len(custo)}).")
    ids = set(transp) | set(custo)
    if len(ids) != 298 or set(transp) != set(custo):
        raise ValueError("As abas Transp e Custo Frete precisam conter os mesmos 298 IDs de loja.")

    shops = {
        store_id: {
            "transp": transp[store_id]["cd"],
            "custo": custo[store_id]["cd"],
            "isShop": bool(transp[store_id]["isShop"] or custo[store_id]["isShop"]),
        }
        for store_id in sorted(ids, key=int)
    }
    return {"transpCount": len(transp), "custoCount": len(custo), "idCount": len(ids), "shops": shops}
