"""Tanpa Rumus: ubah semua formula Excel menjadi nilai hasil hitung terakhir.

Alur per workbook (Windows + Microsoft Excel desktop + xlwings):

1. Validasi file (ekstensi, ukuran, integritas paket OOXML, kecocokan format).
2. Salin ke folder sementara, buka di instance Excel terpisah (tanpa macro, tanpa update link).
3. Hitung ulang dan tunggu kalkulasi selesai (dengan timeout + watchdog yang mematikan Excel).
4. Untuk setiap worksheet (termasuk hidden / very hidden): cari sel berformula dan
   ganti dengan nilainya memakai Copy + PasteSpecial Values pada sel berformula saja.
   Setiap blok diverifikasi: nilai sebelum dan sesudah harus sama.
5. Simpan sebagai file baru, buka ulang di Excel, lalu validasi ulang secara independen
   dari XML paket (jumlah sheet, status sheet, nol formula, makro VBA tetap ada).
6. Kemas hanya file yang berhasil + laporan (manifest.json, laporan_tanpa_rumus.csv) ke satu ZIP.

File sumber yang diunggah tidak pernah dibuka untuk ditulis; semua proses memakai salinan
di folder sementara. openpyxl tidak dipakai untuk menyimpan ulang workbook.
"""
from __future__ import annotations

import csv
import io
import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path, PureWindowsPath
from typing import Any, Sequence
from xml.etree import ElementTree as ET

logger = logging.getLogger("fss.tanpa_rumus")

# ----------------------------------------------------------------------------
# Konfigurasi
# ----------------------------------------------------------------------------
OUTPUT_SUFFIX = "_TANPA_RUMUS"
SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm")
MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024
MAX_FILES = 20
MAX_UNCOMPRESSED_BYTES = 2 * 1024 ** 3
DEFAULT_TIMEOUT_SECONDS = 900.0  # batas total sesi Excel per workbook (buka -> simpan -> verifikasi)
BLOCK_CELLS = 200_000  # ukuran blok pemindaian formula agar memori & jumlah area tetap terkendali
REPORT_ERROR_SAMPLES = 10
QUIT_WAIT_POLLS = 20  # beri Excel sampai ~5 detik untuk keluar sendiri sebelum dimatikan paksa
QUIT_POLL_SECONDS = 0.25

# Konstanta COM Excel.
XL_CALC_AUTOMATIC = -4105
XL_CALC_MANUAL = -4135
XL_CALC_SEMIAUTOMATIC = 2
XL_CALC_DONE = 0
XL_CELL_TYPE_FORMULAS = -4123
XL_ERRORS = 16
XL_PASTE_VALUES = -4163
XL_SHEET_VISIBLE = -1
XL_SHEET_STATES = {-1: "visible", 0: "hidden", 2: "veryHidden"}
XL_FILE_FORMAT = {".xlsx": 51, ".xlsm": 52}
XL_AUTOMATION_SECURITY_FORCE_DISABLE = 3
CALC_MODE_BY_XML = {
    "auto": XL_CALC_AUTOMATIC,
    "manual": XL_CALC_MANUAL,
    "autoNoTable": XL_CALC_SEMIAUTOMATIC,
}

# Konstanta OOXML.
_MAIN_NAMESPACES = (
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
)
_F_TAGS = frozenset(f"{{{ns}}}f" for ns in _MAIN_NAMESPACES)
_ROW_TAGS = frozenset(f"{{{ns}}}row" for ns in _MAIN_NAMESPACES)
_CT_BY_EXTENSION = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
}
_CT_TEMPLATES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_WORKBOOK_PART = "xl/workbook.xml"
_VBA_PART = "xl/vbaProject.bin"

# Clipboard Windows bersifat global: hanya satu sesi Tanpa Rumus boleh berjalan pada satu waktu.
_EXCEL_LOCK = threading.Lock()


# ----------------------------------------------------------------------------
# Exception
# ----------------------------------------------------------------------------
class TanpaRumusError(RuntimeError):
    """Pemrosesan satu workbook gagal; pesan aman ditampilkan ke pengguna."""


class TanpaRumusInputError(TanpaRumusError):
    """File masukan tidak valid (ekstensi, ukuran, rusak, terproteksi, dsb.)."""


class TanpaRumusValidationError(TanpaRumusError):
    """Workbook hasil gagal divalidasi."""


class TanpaRumusTimeout(TanpaRumusError):
    """Excel melewati batas waktu pemrosesan."""


class ExcelUnavailable(TanpaRumusError):
    """Microsoft Excel desktop atau xlwings tidak tersedia."""


# ----------------------------------------------------------------------------
# Struktur data
# ----------------------------------------------------------------------------
@dataclass
class WorkbookInput:
    filename: str
    content: bytes = b""
    error: str | None = None  # kesalahan yang sudah ditemukan saat membaca upload


@dataclass
class WorkbookInfo:
    extension: str
    macro_enabled_type: bool
    has_vba_project: bool
    sheets: list[tuple[str, str]]  # (nama, state)
    external_link_count: int
    connection_count: int
    defined_name_count: int
    named_formula_count: int
    calc_mode: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class OutputInspection:
    extension_matches: bool
    sheets: list[tuple[str, str]]
    formulas_by_sheet: dict[str, int]
    has_vba_project: bool
    external_link_count: int
    table_calculated_columns: int

    @property
    def formulas_remaining(self) -> int:
        return sum(self.formulas_by_sheet.values())


@dataclass
class ConversionStats:
    calculation_mode: str = ""
    calculation_seconds: float = 0.0
    formulas_converted: int = 0
    formula_errors: int = 0
    formula_error_samples: list[str] = field(default_factory=list)
    sheets: list[dict[str, Any]] = field(default_factory=list)
    remaining_links: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class FileReport:
    index: int
    source: str
    output: str | None = None
    status: str = "failed"  # "succeeded" | "failed"
    sheet_count: int | None = None
    formulas_converted: int | None = None
    formulas_remaining: int | None = None
    processing_seconds: float = 0.0
    error: str | None = None
    warnings: list[str] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)
    output_path: Path | None = field(default=None, repr=False)

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "source": self.source,
            "output": self.output,
            "status": self.status,
            "sheet_count": self.sheet_count,
            "formulas_converted": self.formulas_converted,
            "formulas_remaining": self.formulas_remaining,
            "processing_seconds": self.processing_seconds,
            "error": self.error,
            "warnings": list(self.warnings),
            "details": self.details,
        }


# ----------------------------------------------------------------------------
# Nama file
# ----------------------------------------------------------------------------
def _safe_filename(name: str | None, index: int) -> str:
    base = PureWindowsPath(name or "").name  # menangani pemisah / maupun \
    base = re.sub(r'[\x00-\x1f<>:"/\\|?*]', "_", base).strip(" .")
    return base or f"file_{index:02d}"


def _output_name(source_name: str, used: set[str]) -> str:
    path = Path(source_name)
    stem, extension = path.stem[:150], path.suffix.lower()
    if not stem.upper().endswith(OUTPUT_SUFFIX):
        stem = f"{stem}{OUTPUT_SUFFIX}"
    candidate, counter = f"{stem}{extension}", 2
    while candidate.casefold() in used:
        candidate = f"{stem}_{counter}{extension}"
        counter += 1
    used.add(candidate.casefold())
    return candidate


# ----------------------------------------------------------------------------
# Inspeksi paket OOXML (tanpa Excel)
# ----------------------------------------------------------------------------
def _local(tag: str) -> str:
    return tag.rpartition("}")[2]


def _open_package(path: Path) -> zipfile.ZipFile:
    """Buka paket OOXML dan pastikan seluruh isi dapat dibaca (CRC valid)."""
    with path.open("rb") as handle:
        head = handle.read(8)
    if head == _OLE_MAGIC:
        raise TanpaRumusInputError(
            "File dienkripsi dengan kata sandi atau berformat Excel lama (.xls) yang diganti ekstensinya. "
            "Buka kunci/konversi file ke .xlsx atau .xlsm terlebih dahulu."
        )
    if not zipfile.is_zipfile(path):
        raise TanpaRumusInputError("File rusak atau bukan workbook Excel (.xlsx/.xlsm) yang valid.")
    try:
        archive = zipfile.ZipFile(path)
        try:
            if sum(item.file_size for item in archive.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise TanpaRumusInputError("Isi workbook terlalu besar setelah diekstrak (melebihi 2 GB).")
            bad_member = archive.testzip()
        except Exception:
            archive.close()
            raise
    except TanpaRumusInputError:
        raise
    except Exception as exc:
        raise TanpaRumusInputError("File rusak: paket workbook tidak dapat dibaca sepenuhnya.") from exc
    if bad_member is not None:
        archive.close()
        raise TanpaRumusInputError(f"File rusak: bagian '{bad_member}' pada paket workbook tidak valid.")
    return archive


def _workbook_content_type(archive: zipfile.ZipFile) -> str | None:
    try:
        root = ET.fromstring(archive.read("[Content_Types].xml"))
    except (KeyError, ET.ParseError) as exc:
        raise TanpaRumusInputError("File rusak: daftar tipe konten workbook tidak ditemukan atau tidak valid.") from exc
    for element in root:
        if _local(element.tag) == "Override" and element.get("PartName", "").lower() == "/xl/workbook.xml":
            return element.get("ContentType")
    return None


def _parse_workbook_xml(archive: zipfile.ZipFile) -> ET.Element:
    try:
        return ET.fromstring(archive.read(_WORKBOOK_PART))
    except KeyError as exc:
        raise TanpaRumusInputError("File bukan workbook Excel yang valid: 'xl/workbook.xml' tidak ditemukan.") from exc
    except ET.ParseError as exc:
        raise TanpaRumusInputError("File rusak: struktur workbook tidak dapat dibaca.") from exc


def _sheet_list(root: ET.Element) -> list[tuple[str, str]]:
    sheets: list[tuple[str, str]] = []
    for container in root:
        if _local(container.tag) == "sheets":
            for sheet in container:
                if _local(sheet.tag) == "sheet":
                    sheets.append((sheet.get("name", ""), sheet.get("state", "visible")))
    return sheets


def inspect_workbook(path: Path, extension: str) -> WorkbookInfo:
    """Validasi integritas workbook masukan dan kumpulkan informasi struktur."""
    archive = _open_package(path)
    with archive:
        content_type = _workbook_content_type(archive)
        if content_type in _CT_TEMPLATES:
            raise TanpaRumusInputError("File template (.xltx/.xltm) tidak didukung. Simpan sebagai .xlsx atau .xlsm.")
        expected = _CT_BY_EXTENSION[extension]
        if content_type != expected:
            if content_type == _CT_BY_EXTENSION[".xlsm"]:
                raise TanpaRumusInputError(
                    f"Ekstensi {extension} tidak sesuai isi file: workbook berformat makro (.xlsm). Ubah ekstensi menjadi .xlsm."
                )
            if content_type == _CT_BY_EXTENSION[".xlsx"]:
                raise TanpaRumusInputError(
                    f"Ekstensi {extension} tidak sesuai isi file: workbook berformat .xlsx. Ubah ekstensi menjadi .xlsx."
                )
            raise TanpaRumusInputError("Format workbook tidak dikenali sebagai .xlsx atau .xlsm.")

        root = _parse_workbook_xml(archive)
        sheets = _sheet_list(root)
        if not sheets:
            raise TanpaRumusInputError("Workbook tidak memiliki worksheet.")

        names = archive.namelist()
        defined_names: list[str] = []
        calc_mode = "auto"
        for element in root:
            tag = _local(element.tag)
            if tag == "fileSharing" and (element.get("reservationPassword") or element.get("hashValue")):
                raise TanpaRumusInputError(
                    "Workbook dilindungi kata sandi untuk menyimpan perubahan. Hapus kata sandi tersebut terlebih dahulu."
                )
            if tag == "calcPr":
                calc_mode = element.get("calcMode", "auto")
            if tag == "definedNames":
                defined_names = [(child.text or "") for child in element if _local(child.tag) == "definedName"]

        external_links = sum(
            1 for name in names if name.startswith("xl/externalLinks/") and name.endswith(".xml")
        )
        connection_count = 1 if "xl/connections.xml" in names else 0
        named_formulas = sum(1 for text in defined_names if "(" in text)
        has_vba = _VBA_PART in names

    warnings: list[str] = []
    if has_vba:
        warnings.append(
            "Makro VBA dipertahankan pada file .xlsm hasil, tetapi makro dinonaktifkan dan tidak dijalankan selama pemrosesan. "
            "Hitung ulang memakai Calculate agar fungsi buatan sendiri (UDF) tidak berubah menjadi #NAME?."
        )
    if external_links:
        warnings.append(
            f"Workbook memiliki {external_links} external link. Link tidak di-update; hitung ulang memakai Calculate "
            "agar nilai tersimpan dari link eksternal tidak berubah menjadi error."
        )
    if connection_count:
        warnings.append("Workbook memiliki koneksi data. Koneksi tidak dihapus dan tidak di-refresh.")
    if named_formulas:
        warnings.append(
            f"{named_formulas} defined name berisi formula dipertahankan apa adanya karena dapat dipakai "
            "validasi data, conditional formatting, atau grafik."
        )
    return WorkbookInfo(
        extension=extension,
        macro_enabled_type=content_type == _CT_BY_EXTENSION[".xlsm"],
        has_vba_project=has_vba,
        sheets=sheets,
        external_link_count=external_links,
        connection_count=connection_count,
        defined_name_count=len(defined_names),
        named_formula_count=named_formulas,
        calc_mode=calc_mode,
        warnings=warnings,
    )


def _count_sheet_formulas(archive: zipfile.ZipFile, part: str) -> int:
    """Hitung elemen <f> spreadsheetml pada satu worksheet (streaming, hemat memori).

    Hanya namespace utama spreadsheetml yang dihitung; <xm:f> pada extLst (x14) bukan formula sel.
    """
    count = 0
    with archive.open(part) as handle:
        for _event, element in ET.iterparse(handle, events=("end",)):
            if element.tag in _F_TAGS:
                count += 1
            elif element.tag in _ROW_TAGS:
                element.clear()
    return count


def inspect_output(path: Path, extension: str) -> OutputInspection:
    """Periksa workbook hasil langsung dari paket XML (independen dari Excel)."""
    try:
        archive = _open_package(path)
    except TanpaRumusInputError as exc:
        raise TanpaRumusValidationError(f"Validasi hasil gagal: {exc}") from exc
    with archive:
        try:
            content_type = _workbook_content_type(archive)
            root = _parse_workbook_xml(archive)
            names = archive.namelist()
            worksheet_parts = sorted(
                name for name in names
                if name.startswith("xl/worksheets/") and name.endswith(".xml") and name.count("/") == 2
            )
            formulas = {part: _count_sheet_formulas(archive, part) for part in worksheet_parts}
            table_calc = 0
            for name in names:
                if name.startswith("xl/tables/") and name.endswith(".xml"):
                    if b"calculatedColumnFormula" in archive.read(name):
                        table_calc += 1
        except TanpaRumusInputError as exc:
            raise TanpaRumusValidationError(f"Validasi hasil gagal: {exc}") from exc
        except ET.ParseError as exc:
            raise TanpaRumusValidationError("Validasi hasil gagal: XML worksheet pada hasil tidak valid.") from exc
        return OutputInspection(
            extension_matches=content_type == _CT_BY_EXTENSION[extension],
            sheets=_sheet_list(root),
            formulas_by_sheet=formulas,
            has_vba_project=_VBA_PART in names,
            external_link_count=sum(
                1 for name in names if name.startswith("xl/externalLinks/") and name.endswith(".xml")
            ),
            table_calculated_columns=table_calc,
        )


def _output_problems(info: WorkbookInfo, output: OutputInspection) -> list[str]:
    problems: list[str] = []
    if not output.extension_matches:
        problems.append(f"format file hasil tidak sesuai ekstensi {info.extension}")
    if len(output.sheets) != len(info.sheets):
        problems.append(f"jumlah worksheet berubah dari {len(info.sheets)} menjadi {len(output.sheets)}")
    elif output.sheets != info.sheets:
        problems.append("nama, urutan, atau status visible/hidden worksheet berubah")
    if output.formulas_remaining:
        problems.append(f"masih ada {output.formulas_remaining} formula pada worksheet hasil")
    if info.has_vba_project and not output.has_vba_project:
        problems.append("makro VBA tidak terbawa ke file hasil")
    return problems


# ----------------------------------------------------------------------------
# Excel (xlwings / COM)
# ----------------------------------------------------------------------------
def _import_xlwings():
    try:
        import xlwings as xw
    except ImportError as exc:
        raise ExcelUnavailable(
            "Paket xlwings belum terpasang. Jalankan 'python -m pip install xlwings' di environment backend."
        ) from exc
    return xw


def _process_alive(pid: int | None) -> bool:
    """True bila proses Windows dengan PID tersebut masih berjalan."""
    if not pid or os.name != "nt":
        return False
    import ctypes

    kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
    handle = kernel32.OpenProcess(0x00100000, False, int(pid))  # SYNCHRONIZE
    if not handle:
        return False
    try:
        return kernel32.WaitForSingleObject(handle, 0) == 0x102  # WAIT_TIMEOUT -> masih hidup
    finally:
        kernel32.CloseHandle(handle)


class _Watchdog:
    """Mematikan proses Excel bila sesi melewati batas waktu (panggilan COM sinkron tidak bisa di-interrupt)."""

    def __init__(self, app: Any, seconds: float) -> None:
        self.fired = threading.Event()
        self._app = app
        self._timer = threading.Timer(seconds, self._fire)
        self._timer.daemon = True

    def _fire(self) -> None:
        self.fired.set()
        try:
            self._app.kill()
        except Exception:  # noqa: BLE001 - proses mungkin sudah selesai
            pass

    def start(self) -> None:
        self._timer.start()

    def cancel(self) -> None:
        self._timer.cancel()


def _timeout_error(seconds: float) -> TanpaRumusTimeout:
    return TanpaRumusTimeout(
        f"Excel melewati batas waktu pemrosesan ({seconds:.0f} detik). Proses dibatalkan dan tidak ada output yang dibuat."
    )


def _check_deadline(deadline: float, seconds: float) -> None:
    if time.monotonic() > deadline:
        raise _timeout_error(seconds)


def _safe_set(api: Any, name: str, value: Any) -> None:
    try:
        setattr(api, name, value)
    except Exception:  # noqa: BLE001 - properti opsional
        logger.debug("Properti Excel %s tidak dapat diatur", name)


def _configure_application(api: Any) -> None:
    for name, value in (
        ("ScreenUpdating", False),
        ("DisplayAlerts", False),
        ("EnableEvents", False),
        ("AskToUpdateLinks", False),
        ("CalculateBeforeSave", False),
    ):
        _safe_set(api, name, value)
    api.AutomationSecurity = XL_AUTOMATION_SECURITY_FORCE_DISABLE  # jangan jalankan makro dari file unggahan
    api.Calculation = XL_CALC_MANUAL  # kalkulasi hanya saat diminta


def _wait_for_calculation(api: Any, deadline: float, seconds: float) -> None:
    while api.CalculationState != XL_CALC_DONE:
        _check_deadline(deadline, seconds)
        time.sleep(0.25)


def _special_cells(rng: Any, *args: Any) -> Any | None:
    """SpecialCells, atau None bila tidak ada sel yang cocok (Excel melempar error 1004)."""
    try:
        return rng.SpecialCells(*args)
    except Exception:  # noqa: BLE001
        return None


def _normalize(value: Any) -> Any:
    if isinstance(value, (tuple, list)):
        return tuple(_normalize(item) for item in value)
    return "" if value is None else value


def _describe(sheet_name: str, rng: Any) -> str:
    try:
        return f"{sheet_name}!{str(rng.Address).replace('$', '')}"
    except Exception:  # noqa: BLE001
        return sheet_name


def _paste_values(target: Any, sheet_name: str) -> None:
    """Ganti formula pada `target` dengan nilainya (Copy + PasteSpecial Values) dan verifikasi hasilnya."""
    before = target.Value2
    try:
        target.Copy()
        target.PasteSpecial(XL_PASTE_VALUES)
    except Exception:  # noqa: BLE001
        # Potongan array formula / dynamic array tidak boleh diubah sebagian: ubah per unit array.
        _paste_values_cellwise(target, sheet_name)
        return
    if _normalize(before) != _normalize(target.Value2):
        raise TanpaRumusValidationError(
            f"Nilai hasil konversi berbeda dari hasil hitung pada {_describe(sheet_name, target)}. Proses dibatalkan."
        )


def _paste_values_cellwise(target: Any, sheet_name: str) -> None:
    total = int(target.CountLarge)
    for position in range(1, total + 1):
        cell = target.Cells.Item(position)
        if not cell.HasFormula:
            continue  # sudah berubah menjadi nilai saat array yang memuatnya dikonversi
        unit = cell.CurrentArray if cell.HasArray else cell
        before = unit.Value2
        try:
            unit.Copy()
            unit.PasteSpecial(XL_PASTE_VALUES)
        except Exception as exc:  # noqa: BLE001
            raise TanpaRumusError(
                f"Excel tidak dapat mengganti formula pada {_describe(sheet_name, unit)} menjadi nilai "
                "(mungkin bagian dari data table atau struktur yang tidak dapat diubah sebagian)."
            ) from exc
        if _normalize(before) != _normalize(unit.Value2):
            raise TanpaRumusValidationError(
                f"Nilai hasil konversi berbeda dari hasil hitung pada {_describe(sheet_name, unit)}. Proses dibatalkan."
            )


def _convert_area(area: Any, sheet_name: str) -> None:
    hidden = area.EntireRow.Hidden
    if hidden is False:
        _paste_values(area, sheet_name)
        return
    # Ada baris tersembunyi/tersaring: Copy-Paste massal bisa menyalin hanya baris terlihat dan
    # menggeser nilai. Proses baris demi baris agar setiap baris berpasangan dengan dirinya sendiri.
    for position in range(1, int(area.Rows.Count) + 1):
        _paste_values(area.Rows.Item(position), sheet_name)


def _iter_formula_blocks(ws: Any, deadline: float, seconds: float):
    """Hasilkan Range sel berformula per blok baris pada UsedRange worksheet."""
    used = ws.UsedRange
    first_row, first_col = int(used.Row), int(used.Column)
    rows, cols = int(used.Rows.Count), int(used.Columns.Count)
    last_row, last_col = first_row + rows - 1, first_col + cols - 1
    rows_per_block = max(1, BLOCK_CELLS // max(1, cols))
    start = first_row
    while start <= last_row:
        end = min(start + rows_per_block - 1, last_row)
        _check_deadline(deadline, seconds)
        if start == first_row and end == last_row:
            block = used
        else:
            block = ws.Range(ws.Cells.Item(start, first_col), ws.Cells.Item(end, last_col))
        formulas = _special_cells(block, XL_CELL_TYPE_FORMULAS)
        if formulas is not None:
            yield formulas
        start = end + 1


_PROTECTION_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables",
)


def _unprotect_sheet(ws: Any, name: str) -> dict[str, Any] | None:
    if not bool(ws.ProtectContents):
        return None
    protection = ws.Protection
    snapshot = {
        "drawing": bool(ws.ProtectDrawingObjects),
        "scenarios": bool(ws.ProtectScenarios),
        "flags": [bool(getattr(protection, flag)) for flag in _PROTECTION_FLAGS],
    }
    message = (
        f"Worksheet '{name}' diproteksi dengan kata sandi, sehingga formulanya tidak dapat diubah. "
        "Buka proteksi sheet terlebih dahulu."
    )
    try:
        ws.Unprotect()
    except Exception as exc:  # noqa: BLE001
        raise TanpaRumusError(message) from exc
    if bool(ws.ProtectContents):
        raise TanpaRumusError(message)
    return snapshot


def _reprotect_sheet(ws: Any, snapshot: dict[str, Any], name: str) -> None:
    try:
        ws.Protect(None, snapshot["drawing"], True, snapshot["scenarios"], None, *snapshot["flags"])
    except Exception as exc:  # noqa: BLE001
        raise TanpaRumusError(
            f"Proteksi worksheet '{name}' tidak dapat dipulihkan; file dibatalkan agar proteksi tidak hilang diam-diam."
        ) from exc


def _count_formulas(ws: Any, deadline: float, seconds: float) -> int:
    return sum(int(block.CountLarge) for block in _iter_formula_blocks(ws, deadline, seconds))


def _convert_worksheet(api: Any, ws: Any, stats: ConversionStats, deadline: float, seconds: float) -> None:
    name = str(ws.Name)
    original_visible = ws.Visible
    snapshot: dict[str, Any] | None = None
    converted = 0
    try:
        if original_visible != XL_SHEET_VISIBLE:
            try:  # sheet tersembunyi ditampilkan sementara agar Paste selalu dapat dijalankan
                ws.Visible = XL_SHEET_VISIBLE
            except Exception:  # noqa: BLE001
                logger.debug("Sheet '%s' tidak dapat ditampilkan sementara", name)
        snapshot = _unprotect_sheet(ws, name)
        for formulas in _iter_formula_blocks(ws, deadline, seconds):
            converted += int(formulas.CountLarge)
            errors = _special_cells(formulas, XL_CELL_TYPE_FORMULAS, XL_ERRORS)
            if errors is not None:
                stats.formula_errors += int(errors.CountLarge)
                for position in range(1, int(errors.Areas.Count) + 1):
                    if len(stats.formula_error_samples) >= REPORT_ERROR_SAMPLES:
                        break
                    stats.formula_error_samples.append(_describe(name, errors.Areas.Item(position)))
            for position in range(1, int(formulas.Areas.Count) + 1):
                _convert_area(formulas.Areas.Item(position), name)
        remaining = _count_formulas(ws, deadline, seconds)
        if remaining:
            raise TanpaRumusValidationError(
                f"Masih ada {remaining} formula pada worksheet '{name}' setelah konversi "
                "(kemungkinan pada baris yang tersaring/tersembunyi). Hapus filter lalu coba lagi."
            )
        if snapshot is not None:
            _reprotect_sheet(ws, snapshot, name)
    finally:
        _safe_set(api, "CutCopyMode", False)
        if original_visible != XL_SHEET_VISIBLE:
            try:
                ws.Visible = original_visible
            except Exception:  # noqa: BLE001
                logger.debug("Status sheet '%s' tidak dapat dipulihkan", name)
    stats.formulas_converted += converted
    stats.sheets.append({
        "name": name,
        "state": XL_SHEET_STATES.get(original_visible, str(original_visible)),
        "formulas_converted": converted,
    })


def _shutdown_excel(app: Any, pid: int | None, books: list[Any]) -> None:
    for book in books:
        try:
            book.close()
        except Exception:  # noqa: BLE001
            pass
    if app is None:
        return
    try:
        app.quit()
    except Exception:  # noqa: BLE001
        pass
    for _ in range(QUIT_WAIT_POLLS):
        if not _process_alive(pid):
            return
        time.sleep(QUIT_POLL_SECONDS)
    try:
        app.kill()
    except Exception:  # noqa: BLE001
        logger.warning("Proses Excel (PID %s) tidak dapat dimatikan", pid)


def _convert_with_excel(
    input_path: Path,
    output_path: Path,
    info: WorkbookInfo,
    stats: ConversionStats,
    timeout_seconds: float,
) -> None:
    """Hitung ulang, ganti formula dengan nilai, simpan sebagai file baru, lalu buka ulang untuk verifikasi."""
    xw = _import_xlwings()
    with _EXCEL_LOCK:
        app = None
        pid: int | None = None
        watchdog: _Watchdog | None = None
        books: list[Any] = []
        deadline = time.monotonic() + timeout_seconds
        try:
            try:
                app = xw.App(visible=False, add_book=False)
            except Exception as exc:  # noqa: BLE001
                raise ExcelUnavailable(
                    "Microsoft Excel desktop tidak dapat dijalankan. Pastikan Excel terpasang dan dapat dibuka."
                ) from exc
            try:
                pid = int(app.pid)
            except Exception:  # noqa: BLE001
                pid = None
            watchdog = _Watchdog(app, timeout_seconds)
            watchdog.start()
            api = app.api
            books.append(app.books.add())  # workbook kosong: Application.Calculation baru bisa diatur bila ada workbook
            _configure_application(api)

            try:
                book = app.books.open(str(input_path), update_links=False, read_only=False)
            except Exception as exc:  # noqa: BLE001
                if watchdog.fired.is_set():
                    raise
                raise TanpaRumusInputError(
                    "Excel tidak dapat membuka workbook. File kemungkinan rusak, dilindungi kata sandi, "
                    "atau formatnya tidak valid."
                ) from exc
            books.append(book)
            workbook = book.api

            # 1-3. Hitung ulang dan tunggu sampai selesai.
            full = not (info.has_vba_project or info.external_link_count)
            stats.calculation_mode = "full" if full else "smart"
            started = time.perf_counter()
            if full:
                api.CalculateFull()
            else:
                api.Calculate()
            _wait_for_calculation(api, deadline, timeout_seconds)
            stats.calculation_seconds = round(time.perf_counter() - started, 2)

            # 4-7. Semua worksheet (termasuk hidden / very hidden).
            for position in range(1, int(workbook.Worksheets.Count) + 1):
                _convert_worksheet(api, workbook.Worksheets.Item(position), stats, deadline, timeout_seconds)

            try:
                links = workbook.LinkSources(1)
            except Exception:  # noqa: BLE001
                links = None
            if links:
                stats.remaining_links = [str(link) for link in links]
                stats.warnings.append(
                    f"{len(stats.remaining_links)} external link masih tercatat pada workbook (mis. dari defined name atau grafik): "
                    + "; ".join(stats.remaining_links[:3])
                )

            # Kembalikan mode kalkulasi asli agar membuka hasil tidak mengubah mode Excel pengguna.
            api.Calculation = CALC_MODE_BY_XML.get(info.calc_mode, XL_CALC_AUTOMATIC)

            # 8. Simpan sebagai file baru (file sumber tidak ditimpa).
            if output_path.exists():
                output_path.unlink()
            workbook.SaveAs(str(output_path.resolve()), XL_FILE_FORMAT[info.extension])
            book.close()
            books.remove(book)

            # Validasi: buka ulang hasil di Excel.
            reopened = app.books.open(str(output_path), update_links=False, read_only=True)
            books.append(reopened)
            reopened_api = reopened.api
            if int(reopened_api.Sheets.Count) != len(info.sheets):
                raise TanpaRumusValidationError("Validasi hasil gagal: jumlah worksheet berbeda saat file hasil dibuka ulang.")
            for position in range(1, int(reopened_api.Worksheets.Count) + 1):
                sheet = reopened_api.Worksheets.Item(position)
                left = _count_formulas(sheet, deadline, timeout_seconds)
                if left:
                    raise TanpaRumusValidationError(
                        f"Validasi hasil gagal: masih ada {left} formula pada worksheet '{sheet.Name}' saat file hasil dibuka ulang."
                    )
            reopened.close()
            books.remove(reopened)
        except TanpaRumusError:
            raise
        except Exception as exc:  # noqa: BLE001
            if watchdog is not None and watchdog.fired.is_set():
                raise _timeout_error(timeout_seconds) from exc
            logger.exception("Gagal memproses %s melalui Excel", input_path.name)
            raise TanpaRumusError(
                f"Gagal memproses workbook melalui Microsoft Excel ({type(exc).__name__}). File sumber tidak diubah."
            ) from exc
        finally:
            if watchdog is not None:
                watchdog.cancel()
            _shutdown_excel(app, pid, books)


# ----------------------------------------------------------------------------
# Pemrosesan banyak file
# ----------------------------------------------------------------------------
def _process_one(
    index: int,
    item: WorkbookInput,
    work_dir: Path,
    used_names: set[str],
    timeout_seconds: float,
) -> FileReport:
    started = time.perf_counter()
    safe_name = _safe_filename(item.filename, index)
    report = FileReport(index=index, source=safe_name)
    input_path: Path | None = None
    output_path: Path | None = None
    try:
        if item.error:
            raise TanpaRumusInputError(item.error)
        extension = Path(safe_name).suffix.lower()
        if extension not in SUPPORTED_EXTENSIONS:
            raise TanpaRumusInputError(f"{safe_name}: hanya file .xlsx dan .xlsm yang didukung.")
        if not item.content:
            raise TanpaRumusInputError(f"{safe_name}: file kosong.")
        if len(item.content) > MAX_FILE_SIZE_BYTES:
            raise TanpaRumusInputError(f"{safe_name}: ukuran file melebihi batas 50 MB.")

        input_path = work_dir / f"input_{index:02d}{extension}"
        input_path.write_bytes(item.content)
        item.content = b""  # lepaskan memori; salinan kerja ada di disk
        info = inspect_workbook(input_path, extension)
        report.sheet_count = len(info.sheets)

        stats = ConversionStats()
        output_path = work_dir / f"output_{index:02d}{extension}"
        try:
            _convert_with_excel(input_path, output_path, info, stats, timeout_seconds)
            if not output_path.exists() or output_path.stat().st_size == 0:
                raise TanpaRumusValidationError("Excel tidak menghasilkan file output.")
            output = inspect_output(output_path, extension)
            report.formulas_remaining = output.formulas_remaining
            problems = _output_problems(info, output)
            if problems:
                raise TanpaRumusValidationError("Validasi hasil gagal: " + "; ".join(problems) + ".")
        finally:
            report.warnings = [*info.warnings, *stats.warnings]
            report.details = {
                "calculation_mode": stats.calculation_mode,
                "calculation_seconds": stats.calculation_seconds,
                "sheets": stats.sheets,
                "external_links_in_source": info.external_link_count,
                "named_formulas_kept": info.named_formula_count,
            }

        if stats.formula_errors:
            samples = ", ".join(stats.formula_error_samples)
            report.warnings.append(
                f"{stats.formula_errors} formula menghasilkan error (#DIV/0!, #N/A, dsb.) dan disimpan sebagai nilai error "
                f"apa adanya. Contoh sel: {samples}."
            )
            report.details["formula_errors"] = {
                "count": stats.formula_errors,
                "samples": stats.formula_error_samples,
            }
        if output.table_calculated_columns:
            report.warnings.append(
                f"{output.table_calculated_columns} tabel Excel masih menyimpan definisi kolom terhitung; "
                "sel-selnya sudah berupa nilai, tetapi baris baru pada tabel itu dapat terisi formula otomatis."
            )
        if info.has_vba_project:
            report.details["macros_preserved"] = True  # sudah diverifikasi: vbaProject.bin ada pada hasil
        report.formulas_converted = stats.formulas_converted
        report.output = _output_name(safe_name, used_names)
        report.output_path = output_path
        report.status = "succeeded"
        logger.info("Tanpa Rumus berhasil: %s (%s formula)", safe_name, report.formulas_converted)
    except ExcelUnavailable:
        raise
    except TanpaRumusError as exc:
        report.status, report.error = "failed", str(exc)
        if report.sheet_count is not None:
            report.formulas_converted = 0  # tidak ada hasil yang diserahkan
        logger.warning("Tanpa Rumus gagal untuk %s: %s", safe_name, exc)
    except Exception as exc:  # noqa: BLE001 - satu file tidak boleh menggagalkan file lain
        report.status = "failed"
        report.error = f"Terjadi kesalahan tak terduga ({type(exc).__name__}). Periksa log backend."
        logger.exception("Kesalahan tak terduga saat memproses %s", safe_name)
    finally:
        report.processing_seconds = round(time.perf_counter() - started, 2)
        if input_path is not None:
            input_path.unlink(missing_ok=True)
        if report.status != "succeeded" and output_path is not None:
            output_path.unlink(missing_ok=True)
    return report


def _csv_safe(value: Any) -> str:
    text = "" if value is None else str(value)
    return "'" + text if text[:1] in ("=", "+", "-", "@", "\t", "\r") else text


def build_report_csv(reports: Sequence[FileReport]) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow([
        "No", "File Asli", "File Hasil", "Status", "Jumlah Worksheet", "Formula Diubah Menjadi Nilai",
        "Formula Tersisa", "Waktu Proses (detik)", "Peringatan", "Pesan Kesalahan",
    ])
    for report in reports:
        writer.writerow([
            report.index,
            _csv_safe(report.source),
            _csv_safe(report.output),
            "Berhasil" if report.status == "succeeded" else "Gagal",
            "" if report.sheet_count is None else report.sheet_count,
            "" if report.formulas_converted is None else report.formulas_converted,
            "" if report.formulas_remaining is None else report.formulas_remaining,
            report.processing_seconds,
            _csv_safe(" | ".join(report.warnings)),
            _csv_safe(report.error),
        ])
    return buffer.getvalue()


def build_manifest(reports: Sequence[FileReport]) -> dict[str, Any]:
    succeeded = [report for report in reports if report.status == "succeeded"]
    failed = len(reports) - len(succeeded)
    if not succeeded:
        status = "failed"
    elif failed:
        status = "partial"
    else:
        status = "succeeded"
    return {
        "feature": "tanpa_rumus",
        "status": status,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "total_files": len(reports),
        "succeeded": len(succeeded),
        "failed": failed,
        "total_formulas_converted": sum(report.formulas_converted or 0 for report in succeeded),
        "warning_count": sum(len(report.warnings) for report in reports),
        "note": (
            "File sumber tidak diubah. ZIP hanya berisi workbook yang lolos validasi (tanpa formula) dan laporan ini. "
            "File yang gagal tidak disertakan."
        ),
        "files": [report.to_dict() for report in reports],
    }


def _write_zip(zip_path: Path, reports: Sequence[FileReport], manifest: dict[str, Any]) -> None:
    expected = []
    with zipfile.ZipFile(zip_path, "w") as archive:
        for report in reports:
            if report.status == "succeeded" and report.output_path is not None and report.output:
                # Workbook sudah terkompresi; simpan tanpa kompresi ulang.
                archive.write(report.output_path, arcname=report.output, compress_type=zipfile.ZIP_STORED)
                expected.append(report.output)
        archive.writestr(
            "manifest.json",
            json.dumps(manifest, indent=2, ensure_ascii=False),
            compress_type=zipfile.ZIP_DEFLATED,
        )
        archive.writestr(
            "laporan_tanpa_rumus.csv",
            build_report_csv(reports).encode("utf-8-sig"),
            compress_type=zipfile.ZIP_DEFLATED,
        )
        expected.extend(["manifest.json", "laporan_tanpa_rumus.csv"])
    with zipfile.ZipFile(zip_path) as archive:  # pastikan ZIP dapat dibuka kembali
        if archive.testzip() is not None or sorted(archive.namelist()) != sorted(expected):
            raise TanpaRumusError("ZIP hasil tidak valid setelah dibuat ulang. Proses dibatalkan.")


def create_tanpa_rumus_zip(
    inputs: Sequence[WorkbookInput],
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> tuple[Path, Path, dict[str, Any]]:
    """Proses 1..MAX_FILES workbook dan kembalikan (path ZIP, folder kerja, manifest).

    Pemanggil bertanggung jawab menghapus folder kerja setelah ZIP selesai dikirim.
    Kegagalan satu file dicatat pada laporan dan tidak mengganggu file lain.
    """
    if not inputs:
        raise TanpaRumusError("Pilih minimal satu workbook untuk diproses.")
    if len(inputs) > MAX_FILES:
        raise TanpaRumusError(f"Maksimal {MAX_FILES} workbook dapat diproses sekaligus.")
    work_dir = Path(tempfile.mkdtemp(prefix="fss_tanpa_rumus_"))
    try:
        used_names: set[str] = set()
        reports = [
            _process_one(index, item, work_dir, used_names, timeout_seconds)
            for index, item in enumerate(inputs, start=1)
        ]
        manifest = build_manifest(reports)
        zip_path = work_dir / "FSS_Control_Center_Tanpa_Rumus.zip"
        _write_zip(zip_path, reports, manifest)
        return zip_path, work_dir, manifest
    except Exception:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise
