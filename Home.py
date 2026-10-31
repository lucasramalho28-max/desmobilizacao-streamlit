from __future__ import annotations

import csv
import hashlib
import hmac
import io
import json
import os
import shutil
import tempfile
import zipfile
import re
import sqlite3
from contextlib import contextmanager
import unicodedata
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from typing import Any

import streamlit as st

import data_io
from data_io import list_report_files, read_store_workbook

ROOT = data_io.ROOT


DB_PATH = Path(__file__).resolve().parent / "desmobilizacao.db"
REPORT_TEMPLATE = Path(__file__).resolve().parent / "report_template.html"
DEFAULT_SETTINGS = {
    "store_count_source": "transp",
    "codes_boa_to_qeb": "80",
    "codes_qeb_to_boa": "82",
    "codes_qeb_to_sld": "T1",
    "codes_av_return": "AV",
}
OFFICIAL_CDS = {
    "1088": "CD 1088 - RIBEIRAO PRETO SP",
    "1200": "CD 1200 - JUNDIAI SP",
    "1400": "CD 1400 - CAMACARI BA",
    "1401": "CD 1401 - DUQUE DE CAXIAS RJ",
    "1445": "CD 1445 - ESTEIO RS",
    "1475": "CD 1475 - PINHAIS PR",
    "1500": "CD 1500 - CONTAGEM MG",
    "1600": "CD 1600 - CAMPO GRANDE MS",
    "1624": "CD 1624 - ITAJAI SC",
    "1635": "CD 1635 - VARZEA GRANDE MT",
    "1640": "CD 1640 - SERRA ES",
    "1673": "CD 1673 - ITABAIANA SE",
    "1736": "CD 1736 - PARNAMIRIM RN",
    "1760": "CD 1760 - BETIM MG",
    "1792": "CD 1792 - CD",
    "1875": "CD 1875 - SAO LUIS MA",
    "1877": "CD 1877 - MARACANAU CE",
    "1887": "CD 1887 - CONDE PB",
    "1895": "CD 1895 - HIDROLANDIA GO",
    "2241": "CD 2241 - ANANINDEUA PA",
    "2600": "CD 2600 - MANAUS AM",
}



def set_source_root(path: Path) -> None:
    global ROOT
    ROOT = path
    data_io.ROOT = path


def extract_source_bundle(uploaded_file) -> Path:
    content = uploaded_file.getbuffer()
    bundle_hash = hashlib.sha256(content).hexdigest()
    extract_root = Path(tempfile.gettempdir()) / f"desmobilizacao_sources_{bundle_hash[:16]}"
    required = {"1.09": False, "9.09": False, "workbook": False}
    if not extract_root.exists():
        extract_root.mkdir(parents=True, exist_ok=True)
        total_uncompressed = 0
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            entries = [entry for entry in archive.infolist() if not entry.is_dir()]
            if len(entries) > 2000:
                raise ValueError("O ZIP contem mais de 2.000 arquivos.")
            total_uncompressed = sum(entry.file_size for entry in entries)
            if total_uncompressed > 1_000_000_000:
                raise ValueError("O tamanho descompactado do ZIP excede 1 GB.")
            for entry in entries:
                parts = PurePosixPath(entry.filename.replace("\\", "/")).parts
                if not parts or any(part in {"..", "."} for part in parts):
                    continue
                normalized_name = " ".join(Path(parts[-1]).stem.replace("_", " ").lower().split())
                workbook = normalized_name == "lista final 298 lojas 1" and parts[-1].lower().endswith(".xlsx")
                version_index = next((i for i, part in enumerate(parts) if part.lower() in {"1.09", "9.09"}), None)
                if workbook:
                    destination = extract_root / "Lista final_298 lojas 1.xlsx"
                    required["workbook"] = True
                elif version_index is not None and parts[-1].lower().endswith(".csv"):
                    version = parts[version_index].lower()
                    relative_parts = (version, *parts[version_index + 1 :])
                    destination = extract_root.joinpath(*relative_parts)
                    required[version] = True
                else:
                    continue
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(entry) as source, destination.open("wb") as output:
                    shutil.copyfileobj(source, output, length=1024 * 1024)
    else:
        required["1.09"] = any((extract_root / "1.09").rglob("*.csv"))
        required["9.09"] = any((extract_root / "9.09").rglob("*.csv"))
        required["workbook"] = (extract_root / "Lista final_298 lojas 1.xlsx").is_file()
    if not all(required.values()):
        raise ValueError("O ZIP precisa conter 1.09, 9.09 e Lista final_298 lojas 1.xlsx.")
    return extract_root

def normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value or "").strip().lower())
    return "".join(char for char in text if unicodedata.category(char) != "Mn")


def parse_number(value: object, brazilian: bool = False) -> float:
    text = str(value or "").strip()
    if not text:
        return 0.0
    if brazilian:
        text = text.replace(".", "").replace(",", ".")
    else:
        try:
            direct = float(text)
            if direct:
                return direct
        except ValueError:
            pass
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        match = re.match(r"^\s*[-+]?\d+(?:\.\d+)?", text)
        return float(match.group(0)) if match else 0.0


def csv_rows(path: Path):
    with path.open("rb") as raw:
        prefix = raw.read(4)
        raw.seek(0)
        if prefix.startswith((b"\xff\xfe", b"\xfe\xff")):
            encoding = "utf-16"
        elif b"\x00" in prefix:
            encoding = "utf-16-le"
        else:
            encoding = "utf-8-sig"
        text_file = io.TextIOWrapper(raw, encoding=encoding, errors="replace", newline="")
        yield from csv.DictReader(text_file, delimiter="\t")


def raw_code(filename: str, first_row: dict[str, Any] | None) -> str | None:
    if first_row:
        key = next((k for k in first_row if normalize_header(k) == "filial"), None)
        value = str(first_row.get(key, "")).strip() if key else ""
        match = re.search(r"\d{1,6}", value)
        if match:
            return str(int(match.group(0)))
    match = re.search(r"(?<!\d)(\d{3,5})(?!\d)", Path(filename).stem)
    return str(int(match.group(1))) if match else None


def empty_109() -> dict[str, float]:
    return {"initBoa": 0, "initQeb": 0, "initSld": 0, "volumeQeb": 0, "volumeTotal": 0}


def empty_909() -> dict[str, dict[str, float]]:
    return {"sumPecasByRC": {}, "negativeAdjustByRC": {}}


def add_109(summary: dict[str, float], row: dict[str, Any]) -> None:
    qty = 0.0
    reason_code = ""
    asn_type = ""
    volume = 0.0
    for key, value in row.items():
        header = normalize_header(key)
        text = str(value or "").strip()
        if header in {"qtde. pecas", "qtde pecas", "quantidade", "qtd", "estoque", "on hand"}:
            qty = parse_number(text, brazilian=True)
        elif header == "reason code":
            reason_code = text.upper()
        elif header == "tipo asn":
            asn_type = text.upper()
        elif "cubagem" in header or "volume" in header or header in {"m3", "mts"}:
            volume = parse_number(text, brazilian=True)

    item_qty = qty if qty > 0 else (0 if reason_code in {"VF", "S1", "69"} else 1)
    if "QEB" in asn_type or "QUEB" in asn_type:
        summary["initQeb"] += item_qty
        summary["volumeQeb"] += volume
    elif "SLD" in asn_type or "SALDO" in asn_type:
        summary["initSld"] += item_qty
    elif item_qty > 0:
        summary["initBoa"] += item_qty
    summary["volumeTotal"] += volume


def add_909(summary: dict[str, dict[str, float]], row: dict[str, Any]) -> None:
    reason_code = ""
    reason_desc = ""
    pieces = 0.0
    adjustment = 0.0
    for key, value in row.items():
        header = normalize_header(key)
        text = str(value or "").strip()
        if header == "reason code":
            reason_code = text.upper()
        elif header == "reason code description":
            reason_desc = text.upper()
        elif header in {"qtde. pecas", "qtde pecas"}:
            pieces = parse_number(value)
        elif header == "quantidade ajustada":
            adjustment = parse_number(value)

    if not reason_code and reason_desc:
        match = re.match(r"\s*([A-Z0-9]{1,4})\s*-", reason_desc)
        if match:
            reason_code = match.group(1)
    if "R07" in reason_desc or reason_code in {"07", "7"}:
        reason_code = "R07"
    if "R08" in reason_desc or reason_code in {"08", "8"}:
        reason_code = "R08"
    if not reason_code:
        return
    sums = summary["sumPecasByRC"]
    sums[reason_code] = sums.get(reason_code, 0) + pieces
    if adjustment < 0:
        negative = summary["negativeAdjustByRC"]
        negative[reason_code] = negative.get(reason_code, 0) + abs(adjustment)


def merge_summary(target: dict, addition: dict, kind: str) -> dict:
    if kind == "1.09":
        for key in target:
            target[key] += addition[key]
        return target
    for group in ("sumPecasByRC", "negativeAdjustByRC"):
        for code, value in addition[group].items():
            target[group][code] = target[group].get(code, 0) + value
    return target


def current_manifest() -> list[dict[str, Any]]:
    manifest = list_report_files()
    for item in manifest:
        path = data_io.ROOT / item["relativePath"]
        stat = path.stat()
        item["size"] = stat.st_size
        item["mtime_ns"] = stat.st_mtime_ns
    return manifest


def manifest_signature(manifest: list[dict[str, Any]]) -> str:
    stable = [{"path": x["relativePath"], "size": x["size"], "mtime_ns": x["mtime_ns"]} for x in manifest]
    return hashlib.sha256(json.dumps(stable, sort_keys=True).encode("utf-8")).hexdigest()


def discover_raw_bases(manifest: list[dict[str, Any]], store_data: dict[str, Any]) -> dict[str, Any]:
    files_109 = [x for x in manifest if x["relativePath"].replace("\\", "/").startswith("1.09/") and x["name"].lower().endswith(".csv")]
    files_909 = [x for x in manifest if x["relativePath"].replace("\\", "/").startswith("9.09/") and x["name"].lower().endswith(".csv")]
    if not files_109 or not files_909:
        raise ValueError("As pastas 1.09 e 9.09 precisam conter arquivos CSV.")
    if not any(x["name"].lower() == "lista final_298 lojas 1.xlsx" for x in manifest):
        raise ValueError("Lista final_298 lojas 1.xlsx nao encontrada na pasta principal.")

    base109: dict[str, dict[str, float]] = {}
    base909: dict[str, dict[str, dict[str, float]]] = {}
    errors: list[str] = []
    skipped = 0
    csv_files = [x for x in manifest if x["name"].lower().endswith(".csv")]

    for item in csv_files:
        relative = item["relativePath"].replace("\\", "/")
        parts = PurePosixPath(relative).parts
        if any(part.lower().startswith(("backup", "bakup")) for part in parts):
            skipped += 1
            continue
        kind = "1.09" if parts[0].startswith("1.09") else "9.09" if parts[0].startswith("9.09") else None
        if not kind:
            skipped += 1
            continue
        path = data_io.ROOT / relative
        try:
            iterator = iter(csv_rows(path))
            first = next(iterator, None)
            code = raw_code(item["name"], first)
            if not code:
                skipped += 1
                continue
            if kind == "1.09":
                if first is None:
                    skipped += 1
                    continue
                summary = empty_109()
                add_109(summary, first)
                for row in iterator:
                    add_109(summary, row)
                base109[code] = merge_summary(base109.get(code, empty_109()), summary, kind)
            else:
                summary = empty_909()
                if first is None:
                    if code in base109:
                        base909[code] = merge_summary(base909.get(code, empty_909()), summary, kind)
                        continue
                    skipped += 1
                    continue
                add_909(summary, first)
                for row in iterator:
                    add_909(summary, row)
                base909[code] = merge_summary(base909.get(code, empty_909()), summary, kind)
        except Exception as error:
            errors.append(f"{relative}: {error}")

    return {
        "manifest": manifest,
        "rawBase109": base109,
        "rawBase909": base909,
        "stores": store_data["shops"],
        "skipped": skipped,
        "errors": errors,
        "loaded_at": datetime.now(timezone.utc).isoformat(),
    }


def apply_store_source(payload: dict[str, Any], store_source: str) -> dict[str, Any]:
    stores = payload["stores"]

    def resolve(raw: str) -> str:
        if raw in OFFICIAL_CDS:
            return raw
        if raw in stores:
            return str(stores[raw].get(store_source) or raw).zfill(4)
        return str(int(raw)).zfill(4) if raw.isdigit() else raw

    mapped109: dict[str, dict[str, float]] = {}
    mapped909: dict[str, dict[str, dict[str, float]]] = {}
    for source, destination, kind in ((payload["rawBase109"], mapped109, "1.09"), (payload["rawBase909"], mapped909, "9.09")):
        for raw, summary in source.items():
            cd = resolve(raw)
            destination[cd] = merge_summary(destination.get(cd, empty_109() if kind == "1.09" else empty_909()), summary, kind)

    expected = set(OFFICIAL_CDS) | set(mapped109) | set(mapped909)
    missing109 = sorted(expected - set(mapped109))
    missing909 = sorted(expected - set(mapped909))
    issues = []
    if len(stores) != 298:
        issues.append(f"A lista de lojas possui {len(stores)} IDs, esperado 298")
    if missing109:
        issues.append("Sem dados 1.09 para: " + ", ".join(missing109))
    if missing909:
        issues.append("Sem dados 9.09 para: " + ", ".join(missing909))
    if payload["errors"]:
        issues.append(f"{len(payload['errors'])} arquivo(s) falharam na leitura")
    return {**payload, "base109": mapped109, "base909": mapped909, "validation_issues": issues}


def get_connection() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS source_cache (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            signature TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            loaded_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS sync_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            destination TEXT NOT NULL,
            signature TEXT NOT NULL,
            sent_at TEXT NOT NULL,
            status TEXT NOT NULL
        );
        """
    )
    return connection


@contextmanager
def db_connection():
    connection = get_connection()
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def load_settings() -> dict[str, str]:
    settings = dict(DEFAULT_SETTINGS)
    with db_connection() as connection:
        rows = connection.execute("SELECT key, value FROM settings").fetchall()
        for row in rows:
            if row["key"] in settings:
                settings[row["key"]] = row["value"]
    remote_url = online_database_url()
    if remote_url:
        try:
            import psycopg

            with psycopg.connect(remote_url) as connection:
                remote_rows = connection.execute(
                    "SELECT key, value FROM desmobilizacao_settings"
                ).fetchall()
            for key, value in remote_rows:
                if key in settings:
                    settings[key] = value
            if remote_rows:
                save_settings(settings)
        except Exception:
            pass
    return settings


def save_settings(settings: dict[str, str]) -> None:
    with db_connection() as connection:
        connection.executemany(
            "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            [(key, str(settings[key])) for key in DEFAULT_SETTINGS],
        )


def load_cache(signature: str) -> dict[str, Any] | None:
    with db_connection() as connection:
        row = connection.execute("SELECT payload_json FROM source_cache WHERE id=1 AND signature=?", (signature,)).fetchone()
    return json.loads(row["payload_json"]) if row else None


def save_cache(signature: str, payload: dict[str, Any]) -> None:
    json_text = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    with db_connection() as connection:
        connection.execute(
            "INSERT INTO source_cache(id,signature,payload_json,loaded_at) VALUES(1,?,?,?) "
            "ON CONFLICT(id) DO UPDATE SET signature=excluded.signature,payload_json=excluded.payload_json,loaded_at=excluded.loaded_at",
            (signature, json_text, payload["loaded_at"]),
        )


def log_sync(signature: str, status: str, destination: str = "postgresql") -> None:
    with db_connection() as connection:
        connection.execute(
            "INSERT INTO sync_log(destination,signature,sent_at,status) VALUES(?,?,?,?)",
            (destination, signature, datetime.now(timezone.utc).isoformat(), status),
        )


def requested_codes(value: str) -> list[str]:
    return [part.strip().upper() for part in str(value).split(",") if part.strip()]


def calculate_report(payload: dict[str, Any], settings: dict[str, str]) -> tuple[list[dict[str, Any]], dict[str, float]]:
    stores_by_cd: dict[str, dict[str, int]] = {cd: {"lojas": 0, "rua": 0, "shopping": 0} for cd in OFFICIAL_CDS}
    for row in payload["stores"].values():
        cd = str(row.get(settings["store_count_source"], ""))
        if not cd:
            continue
        counts = stores_by_cd.setdefault(cd, {"lojas": 0, "rua": 0, "shopping": 0})
        counts["lojas"] += 1
        counts["shopping" if row.get("isShop") else "rua"] += 1

    codes_80 = requested_codes(settings["codes_boa_to_qeb"])
    codes_82 = requested_codes(settings["codes_qeb_to_boa"])
    codes_t1 = requested_codes(settings["codes_qeb_to_sld"])
    codes_av = requested_codes(settings["codes_av_return"])
    cds = sorted(set(OFFICIAL_CDS) | set(payload["base109"]) | set(payload["base909"]), key=lambda value: (value != "1600", value))
    report_rows: list[dict[str, Any]] = []
    totals = {key: 0.0 for key in ("initBoa", "initQeb", "initSld", "boaToQeb", "qebToBoa", "qebToSld", "avReturn", "finalBoa", "finalQeb", "finalSld", "finalTotal", "numLojas", "volumeRetorno", "volumeMobiliario", "volumeQeb")}

    for cd in cds:
        store_info = stores_by_cd.get(cd, {"lojas": 0, "rua": 0, "shopping": 0})
        source109 = payload["base109"].get(cd, empty_109())
        source909 = payload["base909"].get(cd, empty_909())
        initial_boa = source109["initBoa"]
        initial_qeb = source109["initQeb"]
        initial_sld = source109["initSld"]
        sums = source909["sumPecasByRC"]
        negative = source909["negativeAdjustByRC"]
        boa_to_qeb = sum(sums.get(code, 0) for code in codes_80)
        qeb_to_boa = sum(sums.get(code, 0) for code in codes_82)
        qeb_to_sld = sum(sums.get(code, 0) for code in codes_t1)
        if qeb_to_sld > 0 and any(negative.get(code, 0) > 0 for code in codes_t1):
            qeb_to_sld -= 1
        av_return = sum(sums.get(code, 0) for code in codes_av)
        final_boa = max(0, initial_boa - boa_to_qeb + qeb_to_boa)
        final_qeb = max(0, initial_qeb + boa_to_qeb - qeb_to_boa - qeb_to_sld)
        final_sld = initial_sld + qeb_to_sld
        final_total = final_boa + final_qeb + final_sld + av_return
        row = {
            "CD": cd,
            "Unidade": OFFICIAL_CDS.get(cd, f"CD {cd}"),
            "Lojas": store_info["lojas"],
            "BOA inicial": initial_boa,
            "QEB inicial": initial_qeb,
            "SLD inicial": initial_sld,
            "Total inicial": initial_boa + initial_qeb + initial_sld,
            "BOA → QEB (80)": boa_to_qeb,
            "QEB → BOA (82)": qeb_to_boa,
            "QEB → SLD (T1)": qeb_to_sld,
            "Retorno AV": av_return,
            "BOA final": final_boa,
            "QEB final": final_qeb,
            "SLD final": final_sld,
            "Total final": final_total,
            "Retorno (mts)": 0,
            "Mobiliario/TI (mts)": 0,
            "Cubagem QEB (mts)": source109["volumeQeb"],
        }
        report_rows.append(row)
        totals["initBoa"] += initial_boa
        totals["initQeb"] += initial_qeb
        totals["initSld"] += initial_sld
        totals["boaToQeb"] += boa_to_qeb
        totals["qebToBoa"] += qeb_to_boa
        totals["qebToSld"] += qeb_to_sld
        totals["avReturn"] += av_return
        totals["finalBoa"] += final_boa
        totals["finalQeb"] += final_qeb
        totals["finalSld"] += final_sld
        totals["finalTotal"] += final_total
        totals["numLojas"] += store_info["lojas"]
        totals["volumeRetorno"] += row["Retorno (mts)"]
        totals["volumeMobiliario"] += row["Mobiliario/TI (mts)"]
        totals["volumeQeb"] += row["Cubagem QEB (mts)"]
    return report_rows, totals


def mapped_report_payload(payload: dict[str, Any], settings: dict[str, str]) -> dict[str, Any]:
    report_rows, _ = calculate_report(payload, settings)
    cd_set = {row["CD"] for row in report_rows}
    stores = {
        store_id: {"transp": row["transp"], "custo": row["custo"], "isShop": row["isShop"]}
        for store_id, row in payload["stores"].items()
    }
    raw109 = {cd: payload["base109"].get(cd, empty_109()) for cd in cd_set}
    raw909 = {cd: payload["base909"].get(cd, empty_909()) for cd in cd_set}
    return {"settings": settings, "stores": stores, "base109": raw109, "base909": raw909}


def build_report_html(payload: dict[str, Any], settings: dict[str, str]) -> str:
    html = REPORT_TEMPLATE.read_text(encoding="utf-8")
    report_data = json.dumps(mapped_report_payload(payload, settings), ensure_ascii=False).replace("</", "<\\/")
    marker = "  <script>\n    // ========================================================================="
    if marker not in html:
        raise RuntimeError("Nao encontrei o script da apresentacao HTML.")
    html = html.replace(marker, f"  <script>\n    window.STREAMLIT_DATA = {report_data};\n    // =========================================================================", 1)

    onload_pattern = re.compile(r"    window\.onload = function\(\) \{[\s\S]*?\n    \};")
    onload = """    window.onload = function() {
      const data = window.STREAMLIT_DATA;
      const settings = data.settings;
      document.getElementById('storeCountSource').value = settings.store_count_source;
      document.getElementById('codesBoaToQeb').value = settings.codes_boa_to_qeb;
      document.getElementById('codesQebToBoa').value = settings.codes_qeb_to_boa;
      document.getElementById('codesQebToSld').value = settings.codes_qeb_to_sld;
      document.getElementById('codesAvReturn').value = settings.codes_av_return;
      Object.keys(LOJAS_298_DB).forEach(id => delete LOJAS_298_DB[id]);
      Object.entries(data.stores).forEach(([id, row]) => { LOJAS_298_DB[id] = [row.transp, row.custo, row.isShop ? 1 : 0]; });
      Object.assign(rawFiles109, data.base109);
      Object.assign(rawFiles909, data.base909);
      storeWorkbookLoaded = true;
      dataLoadComplete = true;
      dataLoadValidationError = '';
      document.getElementById('statusRoot').textContent = 'Dados lidos e validados no Streamlit e no SQLite.';
      document.getElementById('sourceStatus').textContent = '21 CDs oficiais validados; 298 lojas carregadas.';
      loadSavedSettings = function() {};
      recalculateAllCDs();
    };"""
    html, count = onload_pattern.subn(onload, html, count=1)
    if count != 1:
        raise RuntimeError("Nao foi possivel preparar a apresentacao para o Streamlit.")
    css = """
    #controlPanel { display: block !important; margin: 0 0 8px; padding: 4px 8px; box-shadow: none; }
    #controlPanel .panel-header { justify-content: flex-end; padding: 0; border: 0; }
    #controlPanel .panel-header > div:first-child, #refreshDataButton, #statusRoot, #sourceStatus,
    #controlPanel .rules-bar, #controlPanel .branch-tags-box { display: none !important; }
    #controlPanel .btn-group { margin-left: auto; }
    """
    html = html.replace("  </style>", css + "\n  </style>", 1)
    return html


def online_database_url() -> str:
    env_url = os.getenv("DATABASE_URL", "").strip()
    if env_url:
        return env_url
    try:
        return str(st.secrets.get("database_url", "")).strip()
    except Exception:
        return ""


def sync_to_postgres(database_url: str, payload: dict[str, Any], signature: str) -> None:
    import psycopg
    from psycopg.types.json import Jsonb

    snapshot = {
        "manifest": payload["manifest"],
        "base109": payload["base109"],
        "base909": payload["base909"],
        "rawBase109": payload["rawBase109"],
        "rawBase909": payload["rawBase909"],
        "stores": payload["stores"],
        "errors": payload["errors"],
        "skipped": payload["skipped"],
        "loaded_at": payload["loaded_at"],
    }
    with psycopg.connect(database_url) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS desmobilizacao_snapshots (
              signature TEXT PRIMARY KEY,
              loaded_at TIMESTAMPTZ NOT NULL,
              snapshot JSONB NOT NULL
            )
            """
        )
        connection.execute(
            """
            INSERT INTO desmobilizacao_snapshots(signature, loaded_at, snapshot)
            VALUES (%s, %s, %s)
            ON CONFLICT(signature) DO UPDATE SET loaded_at=EXCLUDED.loaded_at, snapshot=EXCLUDED.snapshot
            """,
            (signature, payload["loaded_at"], Jsonb(snapshot)),
        )


def load_latest_postgres_snapshot(database_url: str) -> tuple[str, dict[str, Any]] | None:
    import psycopg

    with psycopg.connect(database_url) as connection:
        row = connection.execute(
            "SELECT signature, snapshot FROM desmobilizacao_snapshots ORDER BY loaded_at DESC LIMIT 1"
        ).fetchone()
    if not row:
        return None
    signature, snapshot = row
    payload = {
        "manifest": snapshot.get("manifest", []),
        "rawBase109": snapshot.get("rawBase109", snapshot.get("base109", {})),
        "rawBase909": snapshot.get("rawBase909", snapshot.get("base909", {})),
        "stores": snapshot.get("stores", {}),
        "errors": snapshot.get("errors", []),
        "skipped": snapshot.get("skipped", 0),
        "loaded_at": snapshot["loaded_at"],
    }
    return str(signature), payload


def sync_settings_to_postgres(database_url: str, settings: dict[str, str]) -> None:
    import psycopg

    with psycopg.connect(database_url) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS desmobilizacao_settings (
              key TEXT PRIMARY KEY,
              value TEXT NOT NULL,
              updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """
        )
        connection.executemany(
            """
            INSERT INTO desmobilizacao_settings(key, value, updated_at)
            VALUES (%s, %s, NOW())
            ON CONFLICT(key) DO UPDATE
            SET value=EXCLUDED.value, updated_at=NOW()
            """,
            [(key, str(settings[key])) for key in DEFAULT_SETTINGS],
        )


def require_report_password() -> None:
    if st.session_state.get("report_page_authorized"):
        return
    st.title("Relatorio de Desmobilizacao")
    st.caption("Acesso protegido")
    with st.form("report_login"):
        password = st.text_input("Senha de acesso ao relatorio", type="password")
        submitted = st.form_submit_button("Entrar", type="primary", width="stretch")
    if submitted:
        expected = os.getenv("APP_PASSWORD", "").strip()
        if not expected:
            try:
                expected = str(st.secrets.get("app_password", "")).strip()
            except Exception:
                expected = ""
        if not expected:
            st.error("Defina app_password nos secrets do Streamlit ou APP_PASSWORD no ambiente.")
        elif hmac.compare_digest(password, expected):
            st.session_state["report_page_authorized"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()


def main() -> None:
    st.set_page_config(page_title="Desmobilizacao | Relatorio", layout="wide")
    st.markdown(
        """
        <style>
        .stApp { background: #f4f6fa; }
        .block-container { padding-top: 1.5rem; max-width: 1600px; }
        div[data-testid="stMetric"] { background: white; border: 1px solid #e1e7f0; padding: 14px 16px; border-radius: 12px; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    require_report_password()
    settings = load_settings()
    manifest = current_manifest()
    remote_url = online_database_url()
    remote_snapshot = None
    if not manifest and remote_url:
        try:
            remote_snapshot = load_latest_postgres_snapshot(remote_url)
        except Exception as error:
            st.warning(f"Nao foi possivel ler o PostgreSQL online: {error}")

    if not manifest:
        st.info("Localmente, os diretorios raiz sao reconhecidos automaticamente. No deploy, envie um ZIP com 1.09, 9.09 e Lista final_298 lojas 1.xlsx, ou use o ultimo snapshot salvo no PostgreSQL.")
        uploaded_bundle = st.file_uploader("Bases para deploy (ZIP)", type=["zip"], accept_multiple_files=False)
        if uploaded_bundle is not None:
            try:
                source_root = extract_source_bundle(uploaded_bundle)
                set_source_root(source_root)
                manifest = current_manifest()
                st.success("Bases extraidas. Os dados serao atualizados automaticamente.")
            except (ValueError, zipfile.BadZipFile, OSError) as error:
                st.error(f"Nao foi possivel ler o ZIP: {error}")
                return

    local_signature = manifest_signature(manifest) if manifest else ""
    force_refresh = st.button("Atualizar dados", type="primary")
    payload = None
    if local_signature:
        if not force_refresh:
            payload = load_cache(local_signature)
        if payload is None:
            try:
                store_data = read_store_workbook()
                with st.spinner("Lendo 1.09 e 9.09 e salvando os dados no SQLite..."):
                    payload = discover_raw_bases(manifest, store_data)
                    save_cache(local_signature, payload)
            except Exception as error:
                st.error(f"Nao foi possivel carregar as bases: {error}")
                st.info("Deixe as pastas 1.09 e 9.09 e Lista final_298 lojas 1.xlsx ao lado deste aplicativo.")
                return
    elif remote_snapshot:
        signature, payload = remote_snapshot
        if force_refresh:
            try:
                remote_snapshot = load_latest_postgres_snapshot(remote_url)
                if remote_snapshot:
                    signature, payload = remote_snapshot
            except Exception as error:
                st.error(f"Nao foi possivel atualizar os dados do PostgreSQL: {error}")
                return
        st.caption("Dados recuperados do PostgreSQL online.")
    else:
        st.error("Nao encontrei as bases locais nem um snapshot PostgreSQL. Envie o ZIP com as bases para continuar.")
        return

    if payload is None:
        st.error("Nao consegui recuperar os dados do banco local ou online.")
        return
    payload = apply_store_source(payload, settings["store_count_source"])

    if payload["validation_issues"]:
        st.error("Bases carregadas com pendencias: " + "; ".join(payload["validation_issues"]))
        st.info("A apresentacao e o envio online ficam bloqueados ate completar todas as fontes.")
        return

    report_rows, totals = calculate_report(payload, settings)
    st.title("Desmobilizacao | Relatorio nacional")
    st.caption("Inicio")
    st.caption(f"Dados validados: {len(report_rows)} CDs · {int(totals['numLojas'])} lojas · salvo no SQLite em {payload['loaded_at'][:19].replace('T', ' ')} UTC")
    metrics = st.columns(5)
    metrics[0].metric("CDs", len(report_rows))
    metrics[1].metric("Lojas", f"{int(totals['numLojas']):,}".replace(",", "."))
    metrics[2].metric("Estoque inicial", f"{totals['initBoa'] + totals['initQeb'] + totals['initSld']:,.0f}".replace(",", "."))
    metrics[3].metric("Total final", f"{totals['finalTotal']:,.0f}".replace(",", "."))
    metrics[4].metric("Retorno AV", f"{totals['avReturn']:,.0f}".replace(",", "."))

    tab_summary, tab_cd, tab_method, tab_presentation = st.tabs(["Resumo geral", "CD por CD", "Metodologia", "Apresentacao"])
    with tab_summary:
        st.subheader("Consolidado geral")
        st.dataframe(report_rows, width="stretch", hide_index=True)
        st.markdown("#### Indicadores totais")
        all_totals = [
            {"Indicador": "BOA inicial", "Total": totals["initBoa"]},
            {"Indicador": "QEB inicial", "Total": totals["initQeb"]},
            {"Indicador": "SLD inicial", "Total": totals["initSld"]},
            {"Indicador": "BOA para QEB", "Total": totals["boaToQeb"]},
            {"Indicador": "QEB para BOA", "Total": totals["qebToBoa"]},
            {"Indicador": "QEB para SLD", "Total": totals["qebToSld"]},
            {"Indicador": "Retorno AV sem debito", "Total": totals["avReturn"]},
            {"Indicador": "BOA final", "Total": totals["finalBoa"]},
            {"Indicador": "QEB final", "Total": totals["finalQeb"]},
            {"Indicador": "SLD final", "Total": totals["finalSld"]},
            {"Indicador": "Total final calculado", "Total": totals["finalTotal"]},
            {"Indicador": "Retorno (mts)", "Total": totals["volumeRetorno"]},
            {"Indicador": "Mobiliario/TI (mts)", "Total": totals["volumeMobiliario"]},
            {"Indicador": "Cubagem QEB (mts)", "Total": totals["volumeQeb"]},
        ]
        st.dataframe(all_totals, width="stretch", hide_index=True)
    with tab_cd:
        selected_cd = st.selectbox("Filial", [row["CD"] for row in report_rows], index=0)
        selected = next(row for row in report_rows if row["CD"] == selected_cd)
        st.subheader(selected["Unidade"])
        st.dataframe([selected], width="stretch", hide_index=True)
        st.caption("As saidas que excedem o saldo disponivel deixam o saldo BOA ou QEB em zero.")
    with tab_method:
        st.subheader("Regras de calculo")
        st.markdown(
            f"""
            - **1.09:** classifica o estoque pelo campo Tipo ASN. QEB/QUEB vai para QEB; SLD/Saldo vai para SLD; os demais itens vao para BOA.
            - **CD 1600:** estoque inicial BOA, QEB e SLD lido diretamente da base 1.09, com a mesma regra aplicada a todas as filiais.
            - **9.09:** codigo(s) {settings['codes_boa_to_qeb']} transfere(m) BOA para QEB; {settings['codes_qeb_to_boa']} transfere(m) QEB para BOA; {settings['codes_qeb_to_sld']} transfere(m) QEB para SLD; {settings['codes_av_return']} representa retorno sem debito.
            - **Saldos:** BOA final = max(0, BOA inicial - BOA→QEB + QEB→BOA). QEB final = max(0, QEB inicial + BOA→QEB - QEB→BOA - QEB→SLD).
            - **Lojas:** {int(totals['numLojas'])} lojas lidas da aba {settings['store_count_source']} da planilha Lista final_298 lojas 1.xlsx.
            - **Persistencia:** agregados de CD, manifestos dos arquivos e configuracoes ficam no banco local SQLite. A pagina de configuracoes tem acesso protegido.
            """
        )
    with tab_presentation:
        report_html = build_report_html(payload, settings)
        st.download_button("Baixar apresentacao HTML", report_html, file_name="Relatorio_Desmobilizacao.html", mime="text/html")
        st.html(report_html, width="stretch", unsafe_allow_javascript=True)


if __name__ == "__main__":
    main()
