"""Proses finalisasi workbook delivery untuk fitur Tanpa Rumus.

Setiap output diproses dari salinan workbook sumbernya sendiri. Jika formula bulanan
menggunakan pola SUMPRODUCT Delivery yang dikenali, hasilnya dihitung lewat agregasi
Python sehingga Excel tidak perlu mengulang pencarian puluhan ribu kali ke 90.000
baris. Formula lokal yang tersisa tetap dihitung Excel sebelum dibekukan. Workbook
dengan pola formula berbeda memakai jalur kompatibilitas Excel.
"""
from __future__ import annotations

import io
import json
import logging
import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath
from typing import Any, Callable
from xml.etree import ElementTree as ET

from app.services.tanpa_rumus import (
    XL_AUTOMATION_SECURITY_FORCE_DISABLE,
    XL_CALC_MANUAL,
    CALC_MODE_BY_XML,
    XL_FILE_FORMAT,
    XL_SHEET_VISIBLE,
    ConversionStats,
    ExcelUnavailable,
    TanpaRumusError,
    TanpaRumusInputError,
    TanpaRumusTimeout,
    TanpaRumusValidationError,
    _EXCEL_LOCK,
    _Watchdog,
    _convert_worksheet,
    _import_xlwings,
    _shutdown_excel,
    _wait_for_calculation,
    inspect_output,
    inspect_workbook,
)

logger = logging.getLogger("fss.tanpa_rumus.delivery")
MAX_INPUT_BYTES = 50 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 2 * 1024**3
DEFAULT_TIMEOUT_SECONDS = 900.0
DATE_SHEET_NAMES = {"date", "data"}
MONTH_SHEET_RE = re.compile(r"^20\d{4}")
COUNTRY_CODE_RE = re.compile(r"^7R\d+$", re.IGNORECASE)
# Data FSS dari schedule export bersifat rapat/berurutan. Setelah 10.000 baris
# berturut-turut tanpa nilai pada kolom FSS, sisa sheet dianggap ekor helper.
# Batas konservatif ini menghindari parsing hampir satu juta formula helper.
EMPTY_FSS_TAIL_ROWS_TO_STOP = 10_000

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Header yang dikenali pada area sumber FSS. Kolom lain pada Date tidak akan disentuh.
HEADER_ALIASES = {
    "date": "Date",
    "vendor code": "Vendor code",
    "vendor name": "Vendor name",
    "receiving area": "Receiving Area",
    "recieving area": "Receiving Area",
    "delivery date": "Delivery date",
    "delivery period": "delivery period",
    "firm": "Firm",
    "classification": "Classification",
    "po number": "PO number",
    "item": "item",
    "parts number": "parts number",
    "part number": "parts number",
    "part name": "part name",
    "parts name": "part name",
    "order quantity": "order Quantity",
}
REQUIRED_HEADER_ANCHORS = {"Vendor code", "Receiving Area", "Delivery date", "parts number"}


@dataclass
class DeliveryWorkbookInput:
    filename: str
    content: bytes


@dataclass
class DateSheetPreview:
    sheet_name: str
    header_row: int
    data_start_row: int
    last_data_row: int
    receiving_column: int
    fss_columns: dict[str, int]
    receiving_counts: dict[str, int]
    blank_receiving_rows: int
    month_sheets: list[str]
    # Agregat internal untuk jalur cepat. Data ini tidak dikirim ke frontend.
    formula_key_totals_by_receiving: dict[str, dict[str, float]] = field(default_factory=dict)
    formula_key_totals_all: dict[str, float] = field(default_factory=dict)
    division_suffix_codes: set[str] = field(default_factory=set)
    fast_formula_supported: bool = False
    fast_formula_reason: str | None = None

    def to_dict(self, groups: list[dict[str, Any]]) -> dict[str, Any]:
        configured = {code for group in groups for code in group["codes"]}
        country_codes = {
            code: count for code, count in self.receiving_counts.items()
            if COUNTRY_CODE_RE.fullmatch(code)
        }
        unmapped = sorted(set(country_codes) - configured)
        return {
            "date_sheet": self.sheet_name,
            "header_row": self.header_row,
            "data_start_row": self.data_start_row,
            "last_data_row": self.last_data_row,
            "receiving_column": self.receiving_column,
            "fss_columns": self.fss_columns,
            "month_sheets": self.month_sheets,
            "receiving_areas": [
                {"code": code, "count": count, "is_country_code": bool(COUNTRY_CODE_RE.fullmatch(code))}
                for code, count in sorted(self.receiving_counts.items())
            ],
            "country_receiving_codes": [
                {"code": code, "count": count, "mapped": code in configured}
                for code, count in sorted(country_codes.items())
            ],
            "unmapped_codes": unmapped,
            "blank_receiving_rows": self.blank_receiving_rows,
            "row_count": sum(self.receiving_counts.values()) + self.blank_receiving_rows,
        }


def _local(tag: str) -> str:
    return tag.rpartition("}")[2]


def _column_number(reference: str) -> int:
    letters = "".join(ch for ch in reference if ch.isalpha()).upper()
    result = 0
    for char in letters:
        result = result * 26 + ord(char) - 64
    return result


def _canonical_header(value: Any) -> str | None:
    if value is None:
        return None
    # Delivery workbooks can contain a lowercase `date` helper column in
    # addition to the actual `Date` FSS field. Keep that helper outside the
    # FSS payload, matching the rule already used by sch_validator.
    if isinstance(value, str) and value.strip() == "date":
        return None
    normalized = " ".join(str(value).strip().split()).casefold()
    return HEADER_ALIASES.get(normalized)


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")) for item in root]


def _sheet_parts(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {item.attrib["Id"]: item.attrib["Target"] for item in rels_root}
    result: list[tuple[str, str]] = []
    for sheet in workbook_root.iter(f"{{{MAIN_NS}}}sheet"):
        relation_id = sheet.attrib.get(f"{{{REL_NS}}}id")
        target = rel_targets.get(relation_id or "")
        if not target:
            continue
        if target.startswith("/"):
            part = target.lstrip("/")
        else:
            part = os.path.normpath(os.path.join("xl", target)).replace("\\", "/")
        result.append((sheet.attrib.get("name", ""), part))
    return result


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t"))
    value = cell.find(f"{{{MAIN_NS}}}v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text.strip()


def _to_quantity(value: str) -> float | None:
    """Ubah nilai Order Quantity OOXML ke angka, tanpa menerima teks nonangka."""
    if not value:
        return 0.0
    try:
        return float(value.replace(",", ""))
    except (TypeError, ValueError):
        return None


def _delivery_formula_key(
    delivery_date: str,
    delivery_period: str,
    parts_number: str,
    receiving_area: str,
    division_suffix_codes: set[str] | None = None,
) -> str | None:
    """Bentuk key yang dipakai pola SUMPRODUCT Delivery: tanggal+cycle+part+divisi.

    Tanggal dibuat dengan helper DATE(P,Q,R). Akhiran divisi tidak ditebak dari
    pola kode; ia mengikuti daftar kode yang benar-benar ada pada formula Date!T.
    """
    from datetime import date, timedelta

    raw_date = (delivery_date or "").strip()
    if not raw_date:
        return None
    try:
        # OOXML bisa menyimpan tanggal sebagai YYYYMMDD, YYYY-MM-DD, ISO datetime,
        # atau serial Excel. Prioritaskan format tanggal eksplisit sebelum serial.
        compact = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", raw_date)
        separated = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ].*)?$", raw_date)
        if compact:
            year, month, day = map(int, compact.groups())
            serial = (date(year, month, day) - date(1899, 12, 30)).days
        elif separated:
            year, month, day = map(int, separated.groups())
            serial = (date(year, month, day) - date(1899, 12, 30)).days
        else:
            serial_value = float(raw_date)
            if not serial_value.is_integer() or serial_value < 1 or serial_value > 2958465:
                return None
            serial_date = date(1899, 12, 30) + timedelta(days=int(serial_value))
            serial = (serial_date - date(1899, 12, 30)).days
    except (ValueError, OverflowError):
        return None

    period = (delivery_period or "").strip()
    # Excel concatenation renders an integer numeric cycle as "1", not "1.0".
    try:
        numeric_period = float(period)
        if numeric_period.is_integer():
            period = str(int(numeric_period))
    except ValueError:
        pass

    code = (receiving_area or "").strip().upper()
    suffix_codes = {item.upper() for item in (division_suffix_codes or set())}
    suffix = "-KD" if code in suffix_codes else ""
    return f"{serial}{period}{parts_number.strip()}{suffix}"


_FSS_CELL_REF_RE = re.compile(rb'<c\b[^>]*\br="(?:C|D|E|F|G|H|I|J|K|L|M|N|O)(\d+)"')


def _last_fss_candidate_row(archive: zipfile.ZipFile, sheet_part: str) -> int:
    """Cari batas baris kandidat FSS cepat tanpa membangun satu juta objek XML.

    Pada workbook Delivery, formula helper di kolom A:B dan P:W dapat memanjang
    sampai baris 1.048.576 walaupun data FSS sebenarnya hanya puluhan ribu baris.
    Scan regex terkompilasi pada blok bytes jauh lebih ringan daripada membuat
    objek ElementTree untuk setiap baris helper yang tidak relevan.
    """
    last_row = 0
    carry = b""
    with archive.open(sheet_part, "r") as stream:
        while True:
            chunk = stream.read(1024 * 1024)
            if not chunk:
                break
            block = carry + chunk
            for match in _FSS_CELL_REF_RE.finditer(block):
                row_number = int(match.group(1))
                if row_number > last_row:
                    last_row = row_number
            carry = block[-128:]
    return last_row


def inspect_date_sheet(content: bytes, filename: str) -> DateSheetPreview:
    """Baca kode Receiving Area langsung dari OOXML tanpa membuka Excel.

    Pembacaan XML streaming menghindari pemindaian satu juta baris formula helper.
    """
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise TanpaRumusInputError("Workbook harus berformat .xlsx atau .xlsm.")
    if not content:
        raise TanpaRumusInputError("Workbook kosong.")
    if len(content) > MAX_INPUT_BYTES:
        raise TanpaRumusInputError("Ukuran workbook melebihi batas 50 MB.")

    try:
        archive = zipfile.ZipFile(io.BytesIO(content))
        total_uncompressed = sum(member.file_size for member in archive.infolist())
        if total_uncompressed > MAX_UNCOMPRESSED_BYTES:
            archive.close()
            raise TanpaRumusInputError(
                "Ukuran isi workbook setelah diekstrak melebihi batas aman 2 GB."
            )
        # Jangan menjalankan ZipFile.testzip() di seluruh workbook. Pada template
        # Delivery, Date XML bisa hampir 200 MB karena helper sampai baris Excel
        # maksimum. Member yang dipakai tetap diverifikasi CRC saat dibaca penuh.
        parts = _sheet_parts(archive)
        month_sheets = [name for name, _ in parts if MONTH_SHEET_RE.match(name)]
        date_matches = [(name, part) for name, part in parts if name.strip().casefold() in DATE_SHEET_NAMES]
        if not date_matches:
            raise TanpaRumusInputError("Worksheet 'Date' atau 'DATE' tidak ditemukan.")
        # Prioritaskan Date/DATE. DATA hanya dipilih jika memang tidak ada Date.
        date_matches.sort(key=lambda item: (item[0].strip().casefold() != "date", item[0]))
        sheet_name, sheet_part = date_matches[0]
        if len(date_matches) > 1 and date_matches[0][0].strip().casefold() == date_matches[1][0].strip().casefold():
            raise TanpaRumusInputError("Terdapat lebih dari satu worksheet kandidat Date yang ambigu.")
        if sheet_part not in archive.namelist():
            raise TanpaRumusInputError("XML worksheet Date tidak ditemukan dalam workbook.")
        shared = _shared_strings(archive)

        header_row = 0
        header_map: dict[str, int] = {}
        receiving_counts: Counter[str] = Counter()
        key_totals_by_receiving: dict[str, Counter[str]] = defaultdict(Counter)
        key_totals_all: Counter[str] = Counter()
        blank_receiving_rows = 0
        last_data_row = 0
        data_start = 0
        empty_fss_tail_rows = 0
        helper_key_formula = ""
        division_formula = ""
        division_suffix_codes: set[str] = set()
        fast_formula_reason: str | None = None

        with archive.open(sheet_part) as xml_stream:
            for _, row in ET.iterparse(xml_stream, events=("end",)):
                element = row
                if _local(element.tag) != "row":
                    continue
                row_number = int(element.attrib.get("r", "0") or 0)
                cells = element.findall(f"{{{MAIN_NS}}}c")

                if not header_row and 1 <= row_number <= 25:
                    found: dict[str, int] = {}
                    for cell in cells:
                        ref = cell.attrib.get("r", "")
                        col = _column_number(ref)
                        canonical = _canonical_header(_cell_text(cell, shared))
                        if canonical:
                            if canonical in found:
                                raise TanpaRumusInputError(
                                    f"Header '{canonical}' muncul lebih dari sekali pada kandidat header baris {row_number}."
                                )
                            found[canonical] = col
                    if REQUIRED_HEADER_ANCHORS.issubset(found):
                        header_row = row_number
                        header_map = found
                        data_start = header_row + 1
                        element.clear()
                        continue

                if header_row and row_number >= data_start:
                    values: dict[int, str] = {}
                    wanted_cols = set(header_map.values())
                    for cell in cells:
                        ref = cell.attrib.get("r", "")
                        col = _column_number(ref)
                        if col == 2 and not helper_key_formula:
                            helper = cell.find(f"{{{MAIN_NS}}}f")
                            helper_key_formula = (helper.text or "") if helper is not None else ""
                        if col == 20 and not division_formula:
                            helper = cell.find(f"{{{MAIN_NS}}}f")
                            division_formula = (helper.text or "") if helper is not None else ""
                        if col in wanted_cols:
                            values[col] = _cell_text(cell, shared)
                    if division_formula and not division_suffix_codes:
                        division_suffix_codes = {
                            match.upper()
                            for match in re.findall(
                                r'F\d+\s*=\s*"([^"]+)"\s*,\s*"-KD"',
                                division_formula,
                                re.IGNORECASE,
                            )
                        }
                    has_fss_data = any(value.strip() for value in values.values())
                    receiving_col = header_map["Receiving Area"]
                    receiving = values.get(receiving_col, "").strip().upper()
                    if has_fss_data:
                        last_data_row = row_number
                        empty_fss_tail_rows = 0
                        if receiving:
                            receiving_counts[receiving] += 1
                        else:
                            blank_receiving_rows += 1

                        # Siapkan agregat rumus SUMPRODUCT satu kali dari data mentah.
                        # Key meniru helper Date!B = S&H&M&T. Daftar akhiran -KD
                        # diambil dari formula Division pada template yang sedang dibaca.
                        required_key_fields = {
                            "Delivery date", "delivery period", "parts number", "order Quantity"
                        }
                        if required_key_fields.issubset(header_map):
                            date_value = values.get(header_map["Delivery date"], "")
                            period_value = values.get(header_map["delivery period"], "")
                            part_value = values.get(header_map["parts number"], "")
                            quantity = _to_quantity(values.get(header_map["order Quantity"], ""))
                            key = _delivery_formula_key(
                                date_value, period_value, part_value, receiving, division_suffix_codes
                            )
                            if quantity is None or key is None:
                                fast_formula_reason = "Nilai Delivery date atau Order Quantity tidak dapat dihitung dengan aman."
                            else:
                                key_totals_all[key] += quantity
                                key_totals_by_receiving[receiving][key] += quantity
                        else:
                            fast_formula_reason = "Header Delivery period, Parts number, atau Order Quantity tidak lengkap."
                    else:
                        empty_fss_tail_rows += 1
                element.clear()
                # Sheet delivery biasanya memiliki formula helper hingga baris
                # 1.048.576. Akhiri pembacaan setelah 10.000 baris berurutan
                # tidak lagi berisi nilai pada kolom FSS; jeda yang lebih pendek
                # di tengah data tetap ditoleransi dan penghitung akan di-reset.
                if header_row and row_number >= data_start and empty_fss_tail_rows >= EMPTY_FSS_TAIL_ROWS_TO_STOP:
                    break

        archive.close()
        if not header_row:
            raise TanpaRumusInputError(
                "Header worksheet Date tidak dikenali. Pastikan kolom Vendor code, Receiving Area, Delivery date, dan Parts Number tersedia."
            )
        if not month_sheets:
            raise TanpaRumusInputError(
                "Tidak ditemukan worksheet bulanan dengan nama diawali format YYYYMM, misalnya 202610."
            )
        if last_data_row < data_start:
            raise TanpaRumusInputError("Tidak ditemukan baris data FSS pada worksheet Date.")
        formula_shape_supported = bool(
            re.search(r"IFERROR\(S\d+&H\d+&M\d+&T\d+", helper_key_formula, re.IGNORECASE)
            and division_suffix_codes
        )
        if not formula_shape_supported and fast_formula_reason is None:
            fast_formula_reason = "Formula helper Date!B atau Division tidak sesuai pola Delivery yang dikenali."
        fast_formula_supported = formula_shape_supported and fast_formula_reason is None

        return DateSheetPreview(
            sheet_name=sheet_name,
            header_row=header_row,
            data_start_row=data_start,
            last_data_row=last_data_row,
            receiving_column=header_map["Receiving Area"],
            fss_columns=header_map,
            receiving_counts=dict(receiving_counts),
            blank_receiving_rows=blank_receiving_rows,
            month_sheets=month_sheets,
            formula_key_totals_by_receiving={code: dict(totals) for code, totals in key_totals_by_receiving.items()},
            formula_key_totals_all=dict(key_totals_all),
            division_suffix_codes=set(division_suffix_codes),
            fast_formula_supported=fast_formula_supported,
            fast_formula_reason=fast_formula_reason,
        )
    except TanpaRumusError:
        raise
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        raise TanpaRumusInputError("Workbook tidak dapat dibaca. Pastikan file Excel tidak rusak atau terkunci dengan password.") from exc


def inspect_delivery_source(content: bytes, filename: str) -> dict[str, Any]:
    """Pemeriksaan struktur minimum untuk sumber MMKI tanpa memindai seluruh sheet.

    MMKI tidak memerlukan agregasi Receiving Area, tetapi tetap harus mempunyai
    sheet sumber Date dan setidaknya satu sheet bulanan seperti pada proses akhir.
    Pembacaan hanya menyentuh workbook.xml dan relasi worksheet.
    """
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise TanpaRumusInputError("Workbook harus berformat .xlsx atau .xlsm.")
    if not content:
        raise TanpaRumusInputError("Workbook kosong.")
    if len(content) > MAX_INPUT_BYTES:
        raise TanpaRumusInputError("Ukuran workbook melebihi batas 50 MB.")

    try:
        with zipfile.ZipFile(io.BytesIO(content), "r") as archive:
            if sum(item.file_size for item in archive.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise TanpaRumusInputError("Ukuran isi workbook setelah diekstrak melebihi batas aman 2 GB.")
            sheets = _sheet_parts(archive)
    except TanpaRumusInputError:
        raise
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise TanpaRumusInputError("File sumber rusak atau struktur workbook tidak dapat dibaca.") from exc

    # Utamakan Date/DATE jika workbook memiliki keduanya dengan DATA.
    # Ini menyamakan prioritas dengan parser utama yang dipakai saat proses Excel.
    date_candidates = [name for name, _ in sheets if name.strip().casefold() == "date"]
    if not date_candidates:
        date_candidates = [name for name, _ in sheets if name.strip().casefold() == "data"]
    if not date_candidates:
        raise TanpaRumusInputError("Worksheet 'Date', 'DATE', atau 'DATA' tidak ditemukan.")
    if len(date_candidates) > 1:
        raise TanpaRumusInputError(
            "Ditemukan lebih dari satu worksheet bernama Date/DATE. Sisakan satu sheet sumber saja."
        )
    month_sheets = [name for name, _ in sheets if MONTH_SHEET_RE.match(name)]
    if not month_sheets:
        raise TanpaRumusInputError("Tidak ditemukan worksheet bulanan dengan nama diawali YYYYMM.")
    return {
        "filename": filename,
        "extension": extension,
        "date_sheet": date_candidates[0],
        "month_sheets": month_sheets,
        "sheet_names": [name for name, _ in sheets],
        "processing_seconds": 0.0,
    }
def _safe_stem(filename: str) -> tuple[str, str]:
    base = PureWindowsPath(filename).name
    extension = Path(base).suffix.lower()
    stem = Path(base).stem
    stem = re.sub(r"[\x00-\x1f<>:\"/\\|?*]", "_", stem).strip(" .")[:140]
    if extension not in {".xlsx", ".xlsm"}:
        raise TanpaRumusInputError("Hanya workbook .xlsx dan .xlsm yang didukung.")
    return stem or "workbook", extension


def _safe_label(label: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*]', "-", label).strip(" .")
    if not clean:
        raise TanpaRumusInputError("Label output tidak boleh kosong.")
    return clean[:140]


def build_output_tasks(
    mmki: DeliveryWorkbookInput | None,
    ckd: DeliveryWorkbookInput | None,
    make_mmki: bool,
    make_ckd_all: bool,
    make_ckd_groups: bool,
    selected_groups: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    tasks: list[dict[str, Any]] = []
    used_names: set[str] = set()

    def add(source: DeliveryWorkbookInput, kind: str, label: str, description: str, group=None):
        stem, extension = _safe_stem(source.filename)
        safe_label = _safe_label(label)
        candidate = f"{stem} ({safe_label}){extension}"
        base = Path(candidate).stem
        suffix = Path(candidate).suffix
        counter = 2
        while candidate.casefold() in used_names:
            candidate = f"{base}_{counter}{suffix}"
            counter += 1
        used_names.add(candidate.casefold())
        tasks.append({
            "id": f"output-{len(tasks) + 1}",
            "kind": kind,
            "source_name": source.filename,
            "source_content": source.content,
            "output_name": candidate,
            "description": description,
            "group": group,
        })

    if make_mmki:
        if mmki is None:
            raise TanpaRumusInputError("Pilih workbook MMKI terlebih dahulu.")
        add(mmki, "mmki_all", "TANPA RUMUS", "MMKI Delivery keseluruhan")
    if make_ckd_all:
        if ckd is None:
            raise TanpaRumusInputError("Pilih workbook CKD terlebih dahulu.")
        add(ckd, "ckd_all", "TANPA RUMUS", "CKD Delivery keseluruhan")
    if make_ckd_groups:
        if ckd is None:
            raise TanpaRumusInputError("Pilih workbook CKD terlebih dahulu.")
        if not selected_groups:
            raise TanpaRumusInputError("Pilih minimal satu kelompok receiving.")
        for group in selected_groups:
            add(
                ckd,
                "ckd_group",
                group["outputLabel"],
                f"CKD Delivery · {group['name']}",
                {"id": group["id"], "name": group["name"], "codes": group["codes"], "outputLabel": group["outputLabel"]},
            )
    if not tasks:
        raise TanpaRumusInputError("Pilih minimal satu jenis output.")
    if len(tasks) > 30:
        raise TanpaRumusInputError("Maksimal 30 file output dalam satu proses.")
    return tasks


def _excel_header_layout(date_sheet: Any) -> tuple[int, dict[str, int]]:
    preview = date_sheet.range("A1:AZ25").value
    if not isinstance(preview, list):
        raise TanpaRumusInputError("Header worksheet Date tidak dapat dibaca.")
    for row_index, row in enumerate(preview, start=1):
        mapping: dict[str, int] = {}
        for col_index, value in enumerate(row, start=1):
            canonical = _canonical_header(value)
            if canonical:
                if canonical in mapping:
                    raise TanpaRumusInputError(
                        f"Header '{canonical}' ditemukan lebih dari sekali pada baris {row_index}."
                    )
                mapping[canonical] = col_index
        if REQUIRED_HEADER_ANCHORS.issubset(mapping):
            return row_index, mapping
    raise TanpaRumusInputError(
        "Header worksheet Date tidak sesuai. Kolom Vendor code, Receiving Area, Delivery date, dan Parts Number harus tersedia."
    )


def _contiguous_runs(columns: list[int]) -> list[tuple[int, int]]:
    if not columns:
        return []
    ordered = sorted(set(columns))
    runs: list[tuple[int, int]] = []
    start = previous = ordered[0]
    for column in ordered[1:]:
        if column == previous + 1:
            previous = column
        else:
            runs.append((start, previous))
            start = previous = column
    runs.append((start, previous))
    return runs


def _clear_non_group_data(date_sheet: Any, codes: set[str], preview: DateSheetPreview) -> dict[str, int]:
    header_row, mapping = _excel_header_layout(date_sheet)
    if header_row != preview.header_row or mapping.get("Receiving Area") != preview.receiving_column:
        raise TanpaRumusValidationError(
            "Struktur header worksheet Date berubah antara preview dan proses. Pilih ulang file dan validasi kembali."
        )
    fss_columns = sorted(set(mapping.values()))
    if not fss_columns:
        raise TanpaRumusInputError("Kolom data FSS pada worksheet Date tidak ditemukan.")

    start_row = header_row + 1
    receiving_col = mapping["Receiving Area"]
    # Cari batas data dari kolom FSS, bukan UsedRange yang dapat memanjang karena formula helper.
    xl_up = -4162
    last_rows = []
    for col in fss_columns:
        last_rows.append(int(date_sheet.api.Cells(date_sheet.api.Rows.Count, col).End(xl_up).Row))
    last_row = max(last_rows)
    if last_row < start_row:
        raise TanpaRumusInputError("Tidak ada baris FSS yang dapat difilter pada worksheet Date.")

    # Pastikan tidak ada formula helper di kolom FSS yang akan dikosongkan.
    # Formula helper di kolom lain tidak disentuh.
    xl_cell_type_formulas = -4123
    contiguous = _contiguous_runs(fss_columns)
    for first_col, last_col in contiguous:
        block = date_sheet.range((start_row, first_col), (last_row, last_col))
        try:
            formulas = block.api.SpecialCells(xl_cell_type_formulas)
            if formulas is not None and int(formulas.CountLarge) > 0:
                raise TanpaRumusValidationError(
                    f"Rentang data FSS {block.address} pada sheet Date mengandung formula. "
                    "Proses dihentikan agar formula tidak ikut terhapus."
                )
        except TanpaRumusValidationError:
            raise
        except Exception:
            # Excel melempar error 1004 jika tidak ada formula pada rentang.
            pass

    receiving_values = date_sheet.range((start_row, receiving_col), (last_row, receiving_col)).value
    if not isinstance(receiving_values, list):
        receiving_values = [receiving_values]

    # Baca satu blok data FSS per kelompok kolom kontigu, lalu bersihkan baris yang bukan target.
    # Hanya kolom yang ditemukan dari header FSS yang disentuh; helper di luar kolom tersebut aman.
    cleared_rows: list[int] = []
    kept_rows = 0
    blank_rows = 0
    for offset, value in enumerate(receiving_values):
        if isinstance(value, list):
            value = value[0] if value else None
        code = "" if value is None else str(value).strip().upper()
        if not code:
            # Baris kosong tanpa receiving tidak dianggap order. Baris dengan data FSS lain
            # ditangani sebagai tidak cocok dan dikosongkan agar tidak bocor ke output negara.
            cleared_rows.append(start_row + offset)
            blank_rows += 1
        elif code in codes:
            kept_rows += 1
        else:
            cleared_rows.append(start_row + offset)

    # Gabungkan baris yang berurutan agar jumlah panggilan Excel minimal.
    row_runs: list[tuple[int, int]] = []
    if cleared_rows:
        run_start = previous = cleared_rows[0]
        for row_num in cleared_rows[1:]:
            if row_num == previous + 1:
                previous = row_num
            else:
                row_runs.append((run_start, previous))
                run_start = previous = row_num
        row_runs.append((run_start, previous))

    for first_row, last_clear_row in row_runs:
        for first_col, last_col in contiguous:
            date_sheet.range((first_row, first_col), (last_clear_row, last_col)).clear_contents()

    return {
        "source_data_rows": len(receiving_values),
        "kept_rows": kept_rows,
        "cleared_rows": len(cleared_rows),
        "blank_receiving_rows_cleared": blank_rows,
        "last_data_row": last_row,
    }


class _FastFormulaUnsupported(Exception):
    """Pola formula workbook tidak cocok dengan evaluator SUMPRODUCT yang aman."""


def _excel_concat(value: Any) -> str:
    """Konversi nilai sel seperti operator concatenation Excel untuk key formula."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        number = float(value)
        return str(int(number)) if number.is_integer() else format(number, ".15g")
    return str(value)


def _as_matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    """Normalisasi hasil COM Range.Formula/Value2 menjadi list dua dimensi."""
    if rows == 1 and columns == 1:
        return [[value]]
    if rows == 1:
        return [list(value)]
    if columns == 1:
        return [[item] for item in value]
    return [list(row) for row in value]


def _parse_delivery_sumproduct_formula(formula: str, date_sheet_name: str) -> dict[str, int] | None:
    """Kenali SUMPRODUCT key/quantity yang digunakan template Delivery CKD."""
    if not isinstance(formula, str) or not formula.lstrip().startswith("="):
        return None
    if "SUMPRODUCT" not in formula.upper():
        return None
    sheet = re.escape(date_sheet_name)
    prefix = rf"(?:'{sheet}'|{sheet})!"
    key_range = re.search(prefix + r"\$?B\$(\d+):\$?B\$(\d+)", formula, re.IGNORECASE)
    qty_range = re.search(prefix + r"\$?O\$(\d+):\$?O\$(\d+)", formula, re.IGNORECASE)
    header_ref = re.search(r"!\$?([A-Z]{1,3})\$?(\d+)\s*&", formula, re.IGNORECASE)
    cycle_ref = re.search(r"!\$?B\$?(\d+)\s*&", formula, re.IGNORECASE)
    part_ref = re.search(r"!\$?A\$?(\d+)", formula, re.IGNORECASE)
    if not all((key_range, qty_range, header_ref, cycle_ref, part_ref)):
        return None
    if key_range.group(1) != qty_range.group(1) or key_range.group(2) != qty_range.group(2):
        return None
    return {
        "data_start_row": int(key_range.group(1)),
        "data_end_row": int(key_range.group(2)),
        "header_column": _column_number(header_ref.group(1)),
        "header_row": int(header_ref.group(2)),
        "cycle_row": int(cycle_ref.group(1)),
        "parts_row": int(part_ref.group(1)),
    }


def _aggregate_delivery_keys(
    preview: DateSheetPreview,
    group: dict[str, Any] | None,
) -> dict[str, float]:
    if group is None:
        return preview.formula_key_totals_all
    totals: Counter[str] = Counter()
    for code in group["codes"]:
        totals.update(preview.formula_key_totals_by_receiving.get(str(code).upper(), {}))
    return dict(totals)


def _apply_fast_delivery_formulas(
    book: Any,
    date_sheet_name: str,
    month_sheets: list[str],
    preview: DateSheetPreview,
    group: dict[str, Any] | None,
    stats: ConversionStats,
) -> dict[str, Any]:
    """Ganti puluhan ribu SUMPRODUCT dengan lookup agregat O(n) di Python.

    Perubahan dilakukan pada sheet bulanan saja. Date akan dihapus dari file hasil,
    sehingga tidak perlu mengosongkan 80 ribu baris atau menghitung ulang sejuta
    formula helper pada Date. Formula lokal yang tersisa dihitung Excel setelah
    SUMPRODUCT berat diganti menjadi nilai.
    """
    if not preview.fast_formula_supported:
        raise _FastFormulaUnsupported(preview.fast_formula_reason or "Agregasi formula tidak didukung.")

    totals = _aggregate_delivery_keys(preview, group)
    aggregate_started = time.perf_counter()
    total_replaced = 0
    recalculated: list[str] = []
    for sheet_name in month_sheets:
        sheet = book.sheets[sheet_name]
        sheet_api = sheet.api
        used = sheet_api.UsedRange
        base_row = int(used.Row)
        base_col = int(used.Column)
        row_count = int(used.Rows.Count)
        col_count = int(used.Columns.Count)
        formulas = _as_matrix(used.Formula, row_count, col_count)
        values = _as_matrix(used.Value2, row_count, col_count)
        replaced_positions: list[tuple[int, int]] = []

        for row_index, row_values in enumerate(formulas):
            for col_index, formula in enumerate(row_values):
                if not isinstance(formula, str) or not formula.startswith("="):
                    continue
                date_ref_pattern = r"(?:'" + re.escape(date_sheet_name) + r"'|" + re.escape(date_sheet_name) + r")!"
                if not re.search(date_ref_pattern, formula, re.IGNORECASE):
                    continue
                parsed = _parse_delivery_sumproduct_formula(formula, date_sheet_name)
                if parsed is None:
                    raise _FastFormulaUnsupported(
                        f"Formula pada sheet '{sheet_name}' tidak cocok dengan pola SUMPRODUCT Delivery yang dikenali."
                    )
                if (
                    parsed["data_start_row"] != preview.data_start_row
                    or parsed["data_end_row"] < preview.last_data_row
                ):
                    raise _FastFormulaUnsupported(
                        f"Rentang Date pada formula sheet '{sheet_name}' tidak mencakup seluruh data FSS."
                    )

                header_row = parsed["header_row"] - base_row
                cycle_row = parsed["cycle_row"] - base_row
                parts_row = parsed["parts_row"] - base_row
                header_col = parsed["header_column"] - base_col
                cycle_col = 2 - base_col
                parts_col = 1 - base_col
                if not (
                    0 <= header_row < row_count
                    and 0 <= cycle_row < row_count
                    and 0 <= parts_row < row_count
                    and 0 <= header_col < col_count
                    and 0 <= cycle_col < col_count
                    and 0 <= parts_col < col_count
                ):
                    raise _FastFormulaUnsupported(
                        f"Referensi key pada formula sheet '{sheet_name}' berada di luar area worksheet yang terbaca."
                    )

                key = (
                    _excel_concat(values[header_row][header_col])
                    + _excel_concat(values[cycle_row][cycle_col])
                    + _excel_concat(values[parts_row][parts_col])
                )
                formulas[row_index][col_index] = totals.get(key, 0.0)
                replaced_positions.append((row_index, col_index))

        if replaced_positions:
            # Tulis hanya bounding block formula SUMPRODUCT. Menulis seluruh UsedRange
            # dapat menyentuh sel merged (mis. A8:B8 atau judul), lalu Excel menolak
            # assignment matriks COM. Blok data formula tidak mencakup area header merged.
            min_row = min(position[0] for position in replaced_positions)
            max_row = max(position[0] for position in replaced_positions)
            min_col = min(position[1] for position in replaced_positions)
            max_col = max(position[1] for position in replaced_positions)
            block = sheet_api.Range(
                sheet_api.Cells(base_row + min_row, base_col + min_col),
                sheet_api.Cells(base_row + max_row, base_col + max_col),
            )
            try:
                merged = block.MergeCells
            except Exception as exc:  # noqa: BLE001
                raise _FastFormulaUnsupported(
                    f"Area formula pada sheet '{sheet_name}' tidak dapat diperiksa untuk sel merged."
                ) from exc
            if merged is True:
                raise _FastFormulaUnsupported(
                    f"Area formula pada sheet '{sheet_name}' bertumpang tindih dengan sel merged."
                )
            block_matrix = [
                row[min_col:max_col + 1]
                for row in formulas[min_row:max_row + 1]
            ]
            # Satu assignment matriks menggantikan ribuan formula tanpa 1 COM call
            # per sel. Formula lokal lainnya di dalam blok tetap formula.
            block.Formula = tuple(tuple(row) for row in block_matrix)
            sheet_api.Calculate()
            total_replaced += len(replaced_positions)
            recalculated.append(sheet_name)

    if total_replaced == 0:
        raise _FastFormulaUnsupported("Tidak ditemukan formula SUMPRODUCT Date yang dapat dioptimalkan.")

    # Refresh sheet non-bulanan yang mungkin berisi ringkasan turunan dari sheet bulan.
    for sheet in book.sheets:
        name = str(sheet.name)
        if name.casefold() == date_sheet_name.casefold() or name in month_sheets:
            continue
        sheet.api.Calculate()

    elapsed = round(time.perf_counter() - aggregate_started, 2)
    stats.calculation_seconds = elapsed
    stats.calculation_mode = "python_sumproduct_aggregation"
    stats.warnings.append(
        f"Jalur cepat aktif: {total_replaced:,} formula SUMPRODUCT diganti dengan agregasi Python."
    )
    return {
        "strategy": "python_sumproduct_aggregation",
        "formula_cells_aggregated": total_replaced,
        "month_sheets_recalculated": recalculated,
        "processing_seconds": elapsed,
    }


def _formula_references_date(path: Path, excluded: set[str]) -> list[str]:
    """Cari formula pada sheet non-target yang masih merujuk ke Date sebelum sheet dihapus."""
    references: list[str] = []
    try:
        with zipfile.ZipFile(path) as archive:
            parts = dict(_sheet_parts(archive))
            pattern = re.compile(r"(?:'DATE'|'Date'|'DATA'|DATE|Date|DATA)!", re.IGNORECASE)
            for name, part in parts.items():
                if name.casefold() in excluded or name in excluded:
                    continue
                if part not in archive.namelist():
                    continue
                with archive.open(part) as stream:
                    for _, element in ET.iterparse(stream, events=("end",)):
                        if _local(element.tag) != "c":
                            continue
                        formula = element.find(f"{{{MAIN_NS}}}f")
                        if formula is not None and formula.text and pattern.search(formula.text):
                            references.append(f"{name}!{element.attrib.get('r', '?')}")
                            if len(references) >= 30:
                                return references
                        element.clear()
    except (KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise TanpaRumusValidationError("Formula workbook tidak dapat diperiksa sebelum sheet Date dihapus.") from exc
    return references


def _patch_calc_mode(path: Path, mode: str) -> None:
    """Pulihkan metadata calcMode tanpa mengubah namespace/prefix XML Excel.

    Jangan menggunakan regex ``(<calcPr ...)(/?>)`` yang greedy pada bagian
    atribut. Pada XML Excel yang memakai ``<calcPr .../>``, slash penutup dapat
    ikut tertangkap sebagai atribut lalu ditambahkan sekali lagi (``//>``),
    sehingga workbook.xml menjadi tidak valid. Patch ini bekerja pada bytes,
    membedakan atribut dari penutup tag, mempertahankan semua prefix, lalu
    mem-parse XML hasil sebelum mengganti file.
    """
    import tempfile as _tempfile

    if mode not in {"auto", "manual", "autoNoTable"}:
        mode = "auto"

    def patch_workbook_xml(data: bytes) -> bytes:
        # Lazy attribute capture membuat ``/`` pada self-closing tag tetap
        # berada di bagian penutup, bukan ikut tertelan ke atribut.
        calc_pattern = re.compile(
            rb"(?P<open><(?P<prefix>[A-Za-z_][\w.-]*:)?calcPr\b)"
            rb"(?P<attrs>[^>]*?)(?P<close>\s*/?>)"
        )
        match = calc_pattern.search(data)
        mode_bytes = mode.encode("ascii")
        if match:
            attrs = match.group("attrs")
            attr_pattern = re.compile(rb"(?<![\w:.-])calcMode\s*=\s*([\"'])[^\"']*\1")
            replacement_attr = b'calcMode="' + mode_bytes + b'"'
            if attr_pattern.search(attrs):
                attrs = attr_pattern.sub(replacement_attr, attrs, count=1)
            else:
                attrs = attrs.rstrip() + b" " + replacement_attr + b" "
            patched = data[:match.start()] + match.group("open") + attrs + match.group("close") + data[match.end():]
        else:
            # calcPr opsional. Jika tidak ada, sisipkan sebelum extLst (urutan
            # schema workbook) atau tepat sebelum elemen root penutup.
            root_match = re.search(rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?workbook\b", data)
            if not root_match:
                raise TanpaRumusValidationError("workbook.xml hasil tidak memiliki elemen root workbook.")
            prefix = root_match.group("prefix") or b""
            element = b"<" + prefix + b'calcPr calcMode="' + mode_bytes + b'"/>'
            ext_match = re.search(rb"<" + re.escape(prefix) + rb"extLst\b", data)
            if ext_match:
                insert_at = ext_match.start()
                patched = data[:insert_at] + element + data[insert_at:]
            else:
                close_match = re.search(rb"</" + re.escape(prefix) + rb"workbook\s*>", data)
                if not close_match:
                    raise TanpaRumusValidationError("Penutup elemen workbook tidak ditemukan di workbook.xml hasil.")
                patched = data[:close_match.start()] + element + data[close_match.start():]

        # Pastikan tidak ada kerusakan XML sebelum paket ditulis ke disk.
        try:
            root = ET.fromstring(patched)
        except ET.ParseError as exc:
            raise TanpaRumusValidationError(
                "Metadata kalkulasi tidak dapat dipulihkan karena workbook.xml hasil tidak valid."
            ) from exc
        calc_nodes = [node for node in root if _local(node.tag) == "calcPr"]
        if len(calc_nodes) != 1 or calc_nodes[0].get("calcMode") != mode:
            raise TanpaRumusValidationError("Metadata kalkulasi workbook.xml hasil gagal diverifikasi.")
        return patched

    fd, temp_name = _tempfile.mkstemp(prefix="calc_mode_", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    data = patch_workbook_xml(data)
                target.writestr(item, data)
        os.replace(temp_path, path)
        # Verifikasi bagian workbook.xml yang diubah. Jangan menjalankan
        # ZipFile.testzip() untuk seluruh workbook di sini: itu akan
        # mendekompresi ulang semua sheet besar dan menambah waktu proses.
        with zipfile.ZipFile(path, "r") as check:
            ET.fromstring(check.read("xl/workbook.xml"))
    finally:
        temp_path.unlink(missing_ok=True)


def _output_inspection(path: Path, extension: str, original_sheets: list[tuple[str, str]],
                        date_sheet_name: str, month_sheets: list[str]) -> dict[str, Any]:
    inspection = inspect_output(path, extension)
    output_names = [name for name, _ in inspection.sheets]
    expected_sheets = [sheet for sheet in original_sheets if sheet[0].casefold() != date_sheet_name.casefold()]
    problems: list[str] = []
    if inspection.sheets != expected_sheets:
        problems.append("nama, urutan, atau status visible/hidden worksheet berubah selain penghapusan sheet Date")
    if date_sheet_name.casefold() in {name.casefold() for name in output_names}:
        problems.append("worksheet Date masih ada pada file hasil")
    formula_counts = inspection.formulas_by_sheet
    for sheet_name in month_sheets:
        remaining = formula_counts.get(sheet_name, 0)
        if remaining:
            problems.append(f"masih ada {remaining} formula pada worksheet bulanan '{sheet_name}'")
    if problems:
        raise TanpaRumusValidationError("Validasi output gagal: " + "; ".join(problems) + ".")
    return {
        "sheets_before": [name for name, _ in original_sheets],
        "sheets_after": output_names,
        "date_sheet_removed": True,
        "month_sheets_frozen": month_sheets,
        "formulas_remaining_by_sheet": formula_counts,
        "formulas_remaining_total": sum(formula_counts.values()),
        "vba_preserved": inspection.has_vba_project,
        "external_links_remaining": inspection.external_link_count,
    }


def _calculate_workbook_smart(api: Any, stats: ConversionStats, timeout_seconds: float) -> None:
    """Hitung ulang hanya formula yang ditandai Excel sebagai perlu dihitung.

    CalculateFull menghitung ulang semua formula, termasuk lebih dari satu juta
    formula helper di sheet Date yang nantinya dihapus. Calculate memakai smart
    recalculation Excel: formula yang tidak berubah menggunakan cached result,
    sedangkan formula yang terdampak perubahan data akan dihitung ulang.
    """
    calculation_started = time.perf_counter()
    api.Calculate()
    _wait_for_calculation(api, time.monotonic() + timeout_seconds, timeout_seconds)
    stats.calculation_seconds = round(time.perf_counter() - calculation_started, 2)
    stats.calculation_mode = "calculate_smart"


def _run_one_output(
    source_path: Path,
    output_path: Path,
    source_name: str,
    mode: str,
    group: dict[str, Any] | None,
    timeout_seconds: float,
    workbook_info: Any | None = None,
    date_preview: DateSheetPreview | None = None,
) -> dict[str, Any]:
    extension = source_path.suffix.lower()
    info = workbook_info or inspect_workbook(source_path, extension)
    stats = ConversionStats()
    started = time.perf_counter()
    app = None
    pid: int | None = None
    watchdog: _Watchdog | None = None
    books: list[Any] = []
    input_copy = source_path.with_name(f"working_{uuid.uuid4().hex[:8]}{extension}")
    shutil.copy2(source_path, input_copy)
    output_path.unlink(missing_ok=True)
    filter_report: dict[str, Any] | None = None
    date_sheet_name = ""
    month_sheets: list[str] = []
    restore_warning: str | None = None
    try:
        xw = _import_xlwings()
        app = xw.App(visible=False, add_book=False)
        try:
            pid = int(app.pid)
        except Exception:
            pid = None
        watchdog = _Watchdog(app, timeout_seconds)
        watchdog.start()
        api = app.api
        # Workbook dibuka tanpa update external links/macro. Excel ini terisolasi dari sesi pengguna.
        api.DisplayAlerts = False
        api.ScreenUpdating = False
        api.EnableEvents = False
        api.AskToUpdateLinks = False
        api.AutomationSecurity = XL_AUTOMATION_SECURITY_FORCE_DISABLE
        books.append(app.books.add())
        api.Calculation = XL_CALC_MANUAL
        book = app.books.open(str(input_copy), update_links=False, read_only=False)
        books.append(book)
        workbook = book.api

        sheets = [workbook.Worksheets.Item(index) for index in range(1, int(workbook.Worksheets.Count) + 1)]
        date_candidates = [sheet for sheet in sheets if str(sheet.Name).strip().casefold() in DATE_SHEET_NAMES]
        if not date_candidates:
            raise TanpaRumusInputError("Worksheet 'Date' atau 'DATE' tidak ditemukan.")
        date_candidates.sort(key=lambda sheet: (str(sheet.Name).strip().casefold() != "date", str(sheet.Name)))
        date_sheet_name = str(date_candidates[0].Name)
        # `book.sheets[...]` menghasilkan wrapper xlwings. Wrapper ini diperlukan
        # oleh helper `.range()` / `.api`; Worksheet dari `workbook.Worksheets.Item`
        # adalah objek COM mentah dan tidak memiliki properti `.api`.
        date_sheet = book.sheets[date_sheet_name]
        if len(date_candidates) > 1 and str(date_candidates[0].Name).strip().casefold() == str(date_candidates[1].Name).strip().casefold():
            raise TanpaRumusInputError("Worksheet Date ambigu karena terdapat lebih dari satu kandidat.")

        month_sheets = [
            str(sheet.Name) for sheet in sheets
            if MONTH_SHEET_RE.match(str(sheet.Name))
        ]
        if not month_sheets:
            raise TanpaRumusInputError("Tidak ditemukan worksheet bulanan dengan nama diawali YYYYMM.")

        # Untuk template SUMPRODUCT yang dikenal, hitung hasil tiap kelompok langsung
        # dari data FSS dan hindari kalkulasi global satu juta formula helper di Date.
        # Date akan dihapus dari output sehingga tidak perlu benar-benar mengosongkan
        # kolom-kolom FSS di sheet tersebut.
        fast_calculation: dict[str, Any] | None = None
        if group is not None:
            preview = date_preview or inspect_date_sheet(source_path.read_bytes(), source_name)
            try:
                fast_calculation = _apply_fast_delivery_formulas(
                    book, date_sheet_name, month_sheets, preview, group, stats
                )
                all_data_rows = sum(preview.receiving_counts.values()) + preview.blank_receiving_rows
                kept_rows = sum(preview.receiving_counts.get(code, 0) for code in group["codes"])
                filter_report = {
                    "source_data_rows": all_data_rows,
                    "kept_rows": kept_rows,
                    "cleared_rows": max(0, all_data_rows - kept_rows),
                    "blank_receiving_rows_cleared": 0,
                    "last_data_row": preview.last_data_row,
                    "strategy": "virtual_filter_with_python_aggregation",
                    "formula_cells_aggregated": fast_calculation["formula_cells_aggregated"],
                }
            except _FastFormulaUnsupported as fast_error:
                # Jalur kompatibilitas: pertahankan perilaku Excel lama untuk template
                # yang rumusnya berbeda. Ini lebih lambat, tetapi tidak mengubah hasil.
                if bool(date_sheet.api.ProtectContents):
                    raise TanpaRumusInputError(
                        f"Worksheet '{date_sheet_name}' diproteksi. Buka proteksi sheet sebelum proses Tanpa Rumus."
                    ) from fast_error
                filter_report = _clear_non_group_data(date_sheet, set(group["codes"]), preview)
                stats.warnings.append(
                    f"Jalur cepat tidak digunakan ({fast_error}); kalkulasi Excel standar dijalankan."
                )
                _calculate_workbook_smart(api, stats, timeout_seconds)
        elif mode == "ckd_all" and date_preview is not None:
            # Untuk CKD keseluruhan, formula memakai semua baris. Gunakan agregasi
            # yang sama agar SUMPRODUCT tidak perlu menghitung ulang jutaan perbandingan.
            try:
                fast_calculation = _apply_fast_delivery_formulas(
                    book, date_sheet_name, month_sheets, date_preview, None, stats
                )
            except _FastFormulaUnsupported as fast_error:
                # Tidak ada data yang diubah pada output keseluruhan. Cached formula
                # dari workbook Update Schedule dipakai tanpa kalkulasi ulang global.
                stats.warnings.append(
                    f"Pola formula tidak cocok dengan jalur agregasi cepat ({fast_error}); cached result sumber dipakai."
                )
        # Output keseluruhan lain (mis. MMKI) tidak mengubah sumber data. Gunakan
        # cached result workbook dan jangan memicu Application.Calculate global.

        # Bekukan formula hanya pada worksheet bulanan (nama diawali YYYYMM).
        month_lookup = {name.casefold() for name in month_sheets}
        for sheet in sheets:
            name = str(sheet.Name)
            if name.casefold() in month_lookup:
                _convert_worksheet(api, sheet, stats, time.monotonic() + timeout_seconds, timeout_seconds)

        # Jangan hapus Date jika masih ada formula di worksheet lain yang merujuk ke sheet tersebut.
        remaining_references = _formula_references_date(
            input_copy,
            {date_sheet_name.casefold(), *(name.casefold() for name in month_sheets)},
        )
        if remaining_references:
            raise TanpaRumusValidationError(
                "Masih ada formula di worksheet non-bulanan yang merujuk ke Date: "
                + ", ".join(remaining_references[:10])
                + ". Sheet Date tidak dihapus agar formula tidak rusak."
            )

        if bool(workbook.ProtectStructure):
            raise TanpaRumusInputError(
                "Struktur workbook diproteksi sehingga sheet Date tidak dapat dihapus. Lepas proteksi struktur workbook terlebih dahulu."
            )

        # Date dihapus setelah nilai hasil kalkulasi pada sheet bulanan dibekukan.
        date_sheet.api.Delete()
        if int(workbook.Worksheets.Count) < 1:
            raise TanpaRumusValidationError("Workbook tidak memiliki worksheet tersisa setelah sheet Date dihapus.")

        # Pulihkan Calculation Mode asli sebelum SaveAs supaya Excel sendiri
        # menulis calcMode yang benar. Ini menghindari rewrite/recompress seluruh
        # paket OOXML setelah disimpan, yang mahal untuk workbook berukuran besar.
        api.Calculation = CALC_MODE_BY_XML.get(info.calc_mode, CALC_MODE_BY_XML["auto"])
        workbook.SaveAs(str(output_path.resolve()), XL_FILE_FORMAT[extension])
        book.close()
        # Jangan gunakan list.remove(Book): xlwings.__eq__ membandingkan COM App
        # yang sudah tertutup dan dapat melempar pywintypes.com_error.
        if books and books[-1] is book:
            books.pop()
        else:
            books[:] = [candidate for candidate in books if candidate is not book]

        # Excel normalnya menyimpan calcMode sesuai Calculation di atas. Untuk
        # workbook khusus yang tidak menyimpannya, gunakan patch XML tervalidasi
        # sebagai fallback saja (bukan pada setiap output).
        saved_info = inspect_workbook(output_path, extension)
        if saved_info.calc_mode != info.calc_mode:
            _patch_calc_mode(output_path, info.calc_mode)
        validation = _output_inspection(output_path, extension, info.sheets, date_sheet_name, month_sheets)
        if info.has_vba_project and not validation["vba_preserved"]:
            raise TanpaRumusValidationError("Makro VBA tidak terbawa ke file hasil .xlsm.")
        if not output_path.exists() or output_path.stat().st_size == 0:
            raise TanpaRumusValidationError("File output tidak terbentuk atau kosong.")

        stats.formulas_converted = sum(item["formulas_converted"] for item in stats.sheets)
        stats.formula_errors = stats.formula_errors
        return {
            "status": "succeeded",
            "source": source_name,
            "output": output_path.name,
            "mode": mode,
            "group": group,
            "date_filter": filter_report,
            "month_sheets": month_sheets,
            "formulas_converted": stats.formulas_converted,
            "formula_errors": stats.formula_errors,
            "formula_error_samples": stats.formula_error_samples,
            "calculation_mode": stats.calculation_mode,
            "calculation_seconds": stats.calculation_seconds,
            "fast_calculation": fast_calculation,
            "processing_seconds": round(time.perf_counter() - started, 2),
            "validation": validation,
            "warnings": stats.warnings,
        }
    except TanpaRumusError:
        raise
    except Exception as exc:  # noqa: BLE001
        if watchdog is not None and watchdog.fired.is_set():
            raise TanpaRumusTimeout(
                f"Excel melewati batas waktu {timeout_seconds:.0f} detik untuk output {output_path.name}."
            ) from exc
        logger.exception("Gagal membuat output Tanpa Rumus %s", output_path.name)
        raise TanpaRumusError(
            f"Gagal memproses {source_name} ({type(exc).__name__}). File sumber tetap tidak berubah."
        ) from exc
    finally:
        if watchdog is not None:
            watchdog.cancel()
        _shutdown_excel(app, pid, books)
        input_copy.unlink(missing_ok=True)
        # Excel dibuat oleh proses ini; pengaturan kalkulasinya tidak boleh memengaruhi Excel pengguna.
        # Jika instance gagal ditutup, _shutdown_excel akan mencoba mematikan PID yang dibuat aplikasi.


def create_delivery_zip(
    mmki: DeliveryWorkbookInput | None,
    ckd: DeliveryWorkbookInput | None,
    make_mmki: bool,
    make_ckd_all: bool,
    make_ckd_groups: bool,
    selected_groups: list[dict[str, Any]],
    all_groups_snapshot: list[dict[str, Any]],
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
    progress: Callable[[str, int, int], None] | None = None,
) -> tuple[Path, Path, dict[str, Any]]:
    process_started = time.perf_counter()
    tasks = build_output_tasks(mmki, ckd, make_mmki, make_ckd_all, make_ckd_groups, selected_groups)
    date_preview: DateSheetPreview | None = None
    preview_dict: dict[str, Any] | None = None
    preview_seconds = 0.0
    if ckd is not None and (make_ckd_all or make_ckd_groups):
        # Preview sekaligus membangun agregat key/quantity satu kali. Agregat ini
        # dipakai ulang untuk CKD keseluruhan dan seluruh output receiving group.
        preview_started = time.perf_counter()
        date_preview = inspect_date_sheet(ckd.content, ckd.filename)
        preview_seconds = round(time.perf_counter() - preview_started, 2)
        preview_dict = date_preview.to_dict(all_groups_snapshot)
        if make_ckd_groups:
            if preview_dict["unmapped_codes"]:
                raise TanpaRumusInputError(
                    "Ada kode receiving negara yang belum dipetakan: "
                    + ", ".join(preview_dict["unmapped_codes"])
                    + ". Tambahkan kode tersebut ke kelompok receiving sebelum membuat output CKD per kelompok."
                )
            if date_preview.blank_receiving_rows:
                raise TanpaRumusInputError(
                    f"Ditemukan {date_preview.blank_receiving_rows} baris data dengan Receiving Area kosong. "
                    "Periksa sheet Date sebelum membuat output per kelompok agar order tidak salah dikelompokkan."
                )

    work_dir = Path(tempfile.mkdtemp(prefix="fss_tanpa_rumus_delivery_"))
    outputs_dir = work_dir / "outputs"
    outputs_dir.mkdir(parents=True, exist_ok=True)
    reports: list[dict[str, Any]] = []
    used_output_names: set[str] = set()
    source_cache: dict[str, tuple[Path, Any]] = {}
    total = len(tasks)
    try:
        for index, task in enumerate(tasks, start=1):
            if progress:
                progress(f"Memproses {task['description']} ({index}/{total})", index - 1, total)
            source = mmki if task["kind"] == "mmki_all" else ckd
            if source is None:
                reports.append({"status": "failed", "output": task["output_name"], "error": "Workbook sumber tidak tersedia."})
                continue
            source_key = "mmki" if task["kind"] == "mmki_all" else "ckd"
            if source_key not in source_cache:
                source_path_cached = work_dir / f"source_{source_key}{Path(source.filename).suffix.lower()}"
                source_path_cached.write_bytes(source.content)
                source_info_cached = inspect_workbook(source_path_cached, source_path_cached.suffix.lower())
                source_cache[source_key] = (source_path_cached, source_info_cached)
            source_path, source_info = source_cache[source_key]

            output_name = task["output_name"]
            if output_name.casefold() in used_output_names:
                path = Path(output_name)
                counter = 2
                candidate = f"{path.stem}_{counter}{path.suffix}"
                while candidate.casefold() in used_output_names:
                    counter += 1
                    candidate = f"{path.stem}_{counter}{path.suffix}"
                output_name = candidate
            used_output_names.add(output_name.casefold())
            output_path = outputs_dir / output_name
            group = task["group"]
            mode = task["kind"]
            output_started = time.perf_counter()
            try:
                with _EXCEL_LOCK:
                    report = _run_one_output(
                        source_path,
                        output_path,
                        source.filename,
                        mode,
                        group,
                        timeout_seconds,
                        workbook_info=source_info,
                        date_preview=date_preview if source_key == "ckd" else None,
                    )
                report["output"] = output_name
                report["description"] = task["description"]
                reports.append(report)
            except Exception as exc:  # satu output gagal dicatat; output lain tetap independen
                logger.exception("Output Tanpa Rumus gagal: %s", output_name)
                reports.append({
                    "status": "failed",
                    "source": source.filename,
                    "output": output_name,
                    "description": task["description"],
                    "group": group,
                    "error": str(exc),
                    "processing_seconds": round(time.perf_counter() - output_started, 2),
                })
            if progress:
                progress(f"Selesai memproses {task['description']} ({index}/{total})", index, total)

        succeeded = [report for report in reports if report["status"] == "succeeded"]
        failed = len(reports) - len(succeeded)
        manifest = {
            "feature": "tanpa_rumus_delivery",
            "status": "succeeded" if failed == 0 else "partial" if succeeded else "failed",
            "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "total_outputs": total,
            "succeeded": len(succeeded),
            "failed": failed,
            "preview_seconds": preview_seconds,
            "processing_seconds": round(time.perf_counter() - process_started, 2),
            "receiving_preview": preview_dict,
            "groups_snapshot": all_groups_snapshot,
            "selected_groups": selected_groups,
            "files": reports,
            "note": (
                "Setiap output dibuat dari salinan workbook sumber yang sama. File sumber tidak diubah. "
                "Untuk template CKD yang dikenali, filter kelompok dihitung melalui agregasi per Receiving Area "
                "tanpa menulis ulang sheet Date yang akan dihapus. Jika pola formula tidak didukung, jalur fallback "
                "mengosongkan sel data FSS non-target dan menghitung ulang dengan Excel. Sheet bulanan dibekukan "
                "menjadi nilai, lalu Date dihapus."
            ),
        }
        if not succeeded:
            raise TanpaRumusError(
                "Tidak ada output Tanpa Rumus yang berhasil dibuat. Periksa laporan per output: "
                + "; ".join(report.get("error", "Gagal") for report in reports)
            )

        zip_path = work_dir / "FSS_Control_Center_Tanpa_Rumus.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for report in succeeded:
                output_path = outputs_dir / report["output"]
                archive.write(output_path, arcname=report["output"], compress_type=zipfile.ZIP_STORED)
            archive.writestr("manifest.json", json.dumps(manifest, indent=2, ensure_ascii=False))
            archive.writestr("laporan_tanpa_rumus.csv", _build_csv_report(reports).encode("utf-8-sig"))
        with zipfile.ZipFile(zip_path) as archive:
            if archive.testzip() is not None:
                raise TanpaRumusValidationError("Paket ZIP hasil tidak dapat dibuka kembali.")
        if progress:
            progress("Paket output siap diunduh", total, total)
        return zip_path, work_dir, manifest
    except Exception:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise


def _build_csv_report(reports: list[dict[str, Any]]) -> str:
    import csv
    import io as _io

    output = _io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Output", "Jenis", "Status", "Mode kalkulasi", "Formula menjadi nilai", "Waktu total output (detik)", "Waktu kalkulasi (detik)", "Receiving dipertahankan", "Receiving dikosongkan", "Pesan"])
    for report in reports:
        date_filter = report.get("date_filter") or {}
        writer.writerow([
            report.get("output", ""),
            report.get("description", report.get("mode", "")),
            "Berhasil" if report.get("status") == "succeeded" else "Gagal",
            report.get("calculation_mode", ""),
            report.get("formulas_converted", ""),
            report.get("processing_seconds", ""),
            report.get("calculation_seconds", ""),
            date_filter.get("kept_rows", ""),
            date_filter.get("cleared_rows", ""),
            report.get("error", ""),
        ])
    return output.getvalue()
