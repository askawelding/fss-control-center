import json
from urllib.parse import quote

from fastapi import Form
from fastapi.responses import StreamingResponse

from app.services.fss_builder import build_fss, FSSBuildError
import csv
from collections import Counter
from io import BytesIO, StringIO
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.openapi.utils import get_openapi
from openpyxl import load_workbook
from datetime import date, datetime, time
from app.routers.schedule import router as schedule_router
from app.routers.tanpa_rumus import router as tanpa_rumus_router


# ============================================================
# KONFIGURASI
# ============================================================

MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024
MAX_FILES = 20

SUPPORTED_EXTENSIONS = {".tsv", ".xlsx", ".xlsm"}

OUTPUT_COLUMNS = [
    "Date",
    "Vendor code",
    "Vendor name",
    "Receiving Area",
    "Delivery date",
    "delivery period",
    "Firm",
    "Classification",
    "PO number",
    "item",
    "parts number",
    "part name",
    "order Quantity",
]

# Header Firm sama dengan header output, kecuali kolom Firm.
FIRM_COLUMNS = [
    column
    for column in OUTPUT_COLUMNS
    if column != "Firm"
]

# Normalisasi hanya untuk pencocokan header.
# Nama header output tetap mengikuti OUTPUT_COLUMNS.
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
    "part name": "part name",
    "parts name": "part name",
    "order quantity": "order Quantity",
}


app = FastAPI(
    title="FSS Control Center API",
    description="Backend lokal untuk pengolahan FSS dan Excel.",
    version="0.1.0",
)
app.include_router(schedule_router)
app.include_router(tanpa_rumus_router)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-FSS-Summary", "X-Tanpa-Rumus-Summary"],
)


# ============================================================
# STATUS BACKEND
# ============================================================

@app.get("/api/health")
def health_check():
    return {
        "status": "ok",
        "app": "FSS Control Center",
        "message": "Backend lokal aktif.",
        "version": "0.1.0",
    }


# ============================================================
# PEMBACAAN HEADER
# ============================================================

def canonicalize_header(value):
    """Mencocokkan header tanpa membedakan kapitalisasi/spasi."""

    if value is None:
        return None

    normalized = " ".join(str(value).strip().split()).casefold()

    if not normalized:
        return None

    return HEADER_ALIASES.get(normalized)


def inspect_fss_header(header_values):
    """
    Memeriksa struktur header dan mengenali Firm/Forecast.
    Belum memeriksa isi setiap baris data.
    """

    headers = [
        str(value).strip() if value is not None else ""
        for value in header_values
    ]

    # Abaikan sel kosong di ujung kanan.
    while headers and not headers[-1]:
        headers.pop()

    canonical_headers = []
    unknown_headers = []

    for header in headers:
        if not header:
            continue

        canonical = canonicalize_header(header)

        if canonical is None:
            unknown_headers.append(header)
        else:
            canonical_headers.append(canonical)

    counts = Counter(canonical_headers)

    duplicate_headers = sorted(
        name
        for name, count in counts.items()
        if count > 1
    )

    # Berdasarkan spesifikasi:
    # Forecast memiliki kolom Firm, sedangkan Firm tidak.
    if "Firm" in canonical_headers:
        file_type = "Forecast"
        expected_columns = OUTPUT_COLUMNS
    else:
        file_type = "Firm"
        expected_columns = FIRM_COLUMNS

    missing_columns = [
        column
        for column in expected_columns
        if column not in canonical_headers
    ]

    correct_column_count = len(headers) == len(expected_columns)

    valid = (
        correct_column_count
        and not unknown_headers
        and not duplicate_headers
        and not missing_columns
    )

    if valid:
        message = (
            f"Header sesuai format {file_type}. "
            "Isi data belum diperiksa."
        )
    else:
        problems = []

        if not correct_column_count:
            problems.append(
                f"Jumlah kolom {len(headers)}, "
                f"seharusnya {len(expected_columns)} "
                f"untuk format {file_type}."
            )

        if missing_columns:
            problems.append(
                "Header wajib tidak ditemukan: "
                + ", ".join(missing_columns)
                + "."
            )

        if unknown_headers:
            problems.append(
                "Header tidak dikenali: "
                + ", ".join(unknown_headers)
                + "."
            )

        if duplicate_headers:
            problems.append(
                "Header duplikat: "
                + ", ".join(duplicate_headers)
                + "."
            )

        message = " ".join(problems)

    return {
        "valid": valid,
        "detected_type": file_type,
        "column_count": len(headers),
        "missing_columns": missing_columns,
        "unknown_headers": unknown_headers,
        "duplicate_headers": duplicate_headers,
        "row_preview": headers,
        "message": message,
    }


# ============================================================
# PEMBACA TSV
# ============================================================

def read_tsv_header(content):
    """Mencari baris pertama yang tidak kosong pada TSV."""

    if content.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = content.decode("utf-16")
    else:
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = content.decode("cp1252")
            except UnicodeDecodeError:
                text = content.decode("latin-1")

    reader = csv.reader(
        StringIO(text),
        delimiter="\t",
    )

    for row_number, row in enumerate(reader, start=1):
        if any(value.strip() for value in row):
            while row and not row[-1].strip():
                row.pop()

            return row_number, row

    return None, None


# ============================================================
# VALIDASI FILE FSS
# ============================================================

@app.post("/api/fss/validate")
async def validate_fss_files(
    files: list[UploadFile] = File(...),
):
    if len(files) > MAX_FILES:
        for file in files:
            await file.close()

        raise HTTPException(
            status_code=400,
            detail=f"Maksimal {MAX_FILES} file per pemeriksaan.",
        )

    results = []

    for file in files:
        filename = file.filename or "file-tanpa-nama"
        extension = Path(filename).suffix.lower()

        try:
            if extension not in SUPPORTED_EXTENSIONS:
                results.append({
                    "filename": filename,
                    "valid": False,
                    "message": (
                        "Format tidak didukung. "
                        "Gunakan .tsv, .xlsx, atau .xlsm."
                    ),
                })
                continue

            content = await file.read(
                MAX_FILE_SIZE_BYTES + 1
            )

            if len(content) > MAX_FILE_SIZE_BYTES:
                results.append({
                    "filename": filename,
                    "valid": False,
                    "message": "Ukuran file melebihi 50 MB.",
                })
                continue

            if not content:
                results.append({
                    "filename": filename,
                    "valid": False,
                    "message": "File kosong.",
                })
                continue

            # ------------------------------------------------
            # TSV
            # ------------------------------------------------

            if extension == ".tsv":
                header_row, header_values = read_tsv_header(
                    content
                )

                if header_values is None:
                    results.append({
                        "filename": filename,
                        "valid": False,
                        "message": (
                            "Tidak ditemukan baris header."
                        ),
                    })
                    continue

                header_check = inspect_fss_header(
                    header_values
                )

                results.append({
                    "filename": filename,
                    "format": "TSV",
                    "size_bytes": len(content),
                    "header_row": header_row,
                    **header_check,
                })

                continue

            # ------------------------------------------------
            # EXCEL
            # ------------------------------------------------

            workbook = None

            try:
                workbook = load_workbook(
                    BytesIO(content),
                    read_only=True,
                    data_only=False,
                    keep_vba=(extension == ".xlsm"),
                )

                sheet_names = [
                    sheet.title
                    for sheet in workbook.worksheets
                ]

                header_values = None
                header_row = None
                header_sheet = None

                # Cari baris pertama yang berisi nilai dalam
                # 10 baris pertama worksheet.
                for sheet in workbook.worksheets:
                    for row_number, row in enumerate(
                        sheet.iter_rows(
                            min_row=1,
                            max_row=10,
                            min_col=1,
                            max_col=30,
                            values_only=True,
                        ),
                        start=1,
                    ):
                        values = list(row)

                        if any(
                            value is not None
                            and str(value).strip()
                            for value in values
                        ):
                            header_values = values

                            while (
                                header_values
                                and header_values[-1] is None
                            ):
                                header_values.pop()

                            header_row = row_number
                            header_sheet = sheet.title
                            break

                    if header_values is not None:
                        break

                if header_values is None:
                    results.append({
                        "filename": filename,
                        "valid": False,
                        "message": (
                            "Header tidak ditemukan pada "
                            "10 baris pertama worksheet."
                        ),
                    })
                    continue

                header_check = inspect_fss_header(
                    header_values
                )

                results.append({
                    "filename": filename,
                    "format": extension.lstrip(".").upper(),
                    "size_bytes": len(content),
                    "sheets": sheet_names,
                    "header_sheet": header_sheet,
                    "header_row": header_row,
                    **header_check,
                })

            finally:
                if workbook is not None:
                    workbook.close()

        except Exception:
            results.append({
                "filename": filename,
                "valid": False,
                "message": (
                    "File gagal dibaca. "
                    "Periksa format atau kondisi file."
                ),
            })

        finally:
            await file.close()

    valid_count = sum(
        1
        for result in results
        if result.get("valid") is True
    )

    return {
        "status": "ok",
        "total_files": len(results),
        "valid_files": valid_count,
        "invalid_files": len(results) - valid_count,
        "files": results,
    }


# ============================================================
# PERBAIKAN SWAGGER UI UNTUK UNGGAH FILE
# ============================================================

def custom_openapi():
    if app.openapi_schema:
        return app.openapi_schema

    openapi_schema = get_openapi(
        title=app.title,
        version=app.version,
        description=app.description,
        routes=app.routes,
    )

    openapi_schema["openapi"] = "3.0.3"

    def fix_file_upload_schema(node):
        if isinstance(node, dict):
            if node.get("contentMediaType") == (
                "application/octet-stream"
            ):
                node.pop("contentMediaType", None)
                node["type"] = "string"
                node["format"] = "binary"

            for value in list(node.values()):
                fix_file_upload_schema(value)

        elif isinstance(node, list):
            for item in node:
                fix_file_upload_schema(item)

    fix_file_upload_schema(openapi_schema)

    app.openapi_schema = openapi_schema
    return app.openapi_schema


app.openapi = custom_openapi

def serialize_preview_value(value):
    """Mengubah nilai Excel menjadi tipe data yang aman untuk JSON."""

    if isinstance(value, (datetime, date, time)):
        return value.isoformat()

    if value is None or isinstance(value, (str, int, float, bool)):
        return value

    return str(value)

# ============================================================
# PREVIEW DATA FSS
# ============================================================

@app.post("/api/fss/preview")
async def preview_fss_endpoint(
    files: list[UploadFile] = File(...),
    release_date: str = Form(...),
):
    """Memproses file FSS dan mengembalikan contoh baris Firm/Forecast."""

    if not files:
        raise HTTPException(
            status_code=400,
            detail="Pilih minimal satu file Firm dan satu file Forecast.",
        )

    if len(files) > MAX_FILES:
        raise HTTPException(
            status_code=400,
            detail=f"Maksimal {MAX_FILES} file dapat diproses sekaligus.",
        )

    uploaded_files = []

    try:
        for file in files:
            filename = file.filename or "file"

            content = await file.read(MAX_FILE_SIZE_BYTES + 1)

            if len(content) > MAX_FILE_SIZE_BYTES:
                raise FSSBuildError(
                    f"{filename}: ukuran file melebihi 50 MB."
                )

            if not content:
                raise FSSBuildError(f"{filename}: file kosong.")

            uploaded_files.append((filename, content))

        # Gunakan proses utama agar hasil preview sama dengan hasil build.
        workbook_bytes, summary = build_fss(
            files=uploaded_files,
            release_date=release_date,
        )

        workbook = load_workbook(
            BytesIO(workbook_bytes),
            read_only=True,
            data_only=True,
        )

        try:
            worksheet = workbook["Data FSS"]

            headers = next(
                worksheet.iter_rows(
                    min_row=1,
                    max_row=1,
                    values_only=True,
                )
            )

            firm_count = int(summary["firm_rows"])
            forecast_count = int(summary["forecast_rows_after_filter"])

            preview_rows = []
            sample_per_type = 5

            # Ambil maksimal 5 baris Firm.
            firm_sample_count = min(firm_count, sample_per_type)

            if firm_sample_count > 0:
                for row in worksheet.iter_rows(
                    min_row=2,
                    max_row=1 + firm_sample_count,
                    values_only=True,
                ):
                    preview_rows.append({
                        "type": "Firm",
                        "values": [
                            serialize_preview_value(value)
                            for value in row
                        ],
                    })

            # Ambil maksimal 5 baris Forecast.
            forecast_sample_count = min(
                forecast_count,
                sample_per_type,
            )

            if forecast_sample_count > 0:
                forecast_start_row = 2 + firm_count

                for row in worksheet.iter_rows(
                    min_row=forecast_start_row,
                    max_row=forecast_start_row + forecast_sample_count - 1,
                    values_only=True,
                ):
                    preview_rows.append({
                        "type": "Forecast",
                        "values": [
                            serialize_preview_value(value)
                            for value in row
                        ],
                    })

            return {
                "status": "ok",
                "filename": summary["filename"],
                "columns": list(headers),
                "rows": preview_rows,
                "preview_count": len(preview_rows),
                "firm_rows": firm_count,
                "forecast_rows": forecast_count,
                "total_rows": summary["total_rows"],
            }

        finally:
            workbook.close()

    except FSSBuildError as exc:
        raise HTTPException(
            status_code=400,
            detail=str(exc),
        ) from exc

    except HTTPException:
        raise

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=(
                "Terjadi kesalahan saat mengambil preview Data FSS. "
                "Periksa log backend untuk detailnya."
            ),
        ) from exc

    finally:
        for file in files:
            await file.close()

@app.post("/api/fss/build")
async def build_fss_endpoint(
    files: list[UploadFile] = File(...),
    release_date: str = Form(...),
):
    """Gabungkan file Firm dan Forecast menjadi workbook Excel."""
    from io import BytesIO

    if not files:
        raise HTTPException(
            status_code=400,
            detail="Pilih minimal satu file Firm dan satu file Forecast.",
        )

    if len(files) > 20:
        raise HTTPException(
            status_code=400,
            detail="Maksimal 20 file dapat diproses sekaligus.",
        )

    uploaded_files = []

    try:
        for file in files:
            content = await file.read()

            if not content:
                raise FSSBuildError(
                    f"{file.filename}: file kosong."
                )

            uploaded_files.append(
                (file.filename or "file", content)
            )

        workbook_bytes, summary = build_fss(
            files=uploaded_files,
            release_date=release_date,
        )

        # Metadata ringkas dikirim melalui response header.
        public_summary = {
            "filename": summary["filename"],
            "firm_files": summary["firm_files"],
            "forecast_file": summary["forecast_file"],
            "firm_rows": summary["firm_rows"],
            "forecast_rows_before_filter": (
                summary["forecast_rows_before_filter"]
            ),
            "forecast_rows_removed": (
                summary["forecast_rows_removed"]
            ),
            "forecast_rows_after_filter": (
                summary["forecast_rows_after_filter"]
            ),
            "total_rows": summary["total_rows"],
            "duplicate_rows_warned": (
                summary["duplicate_rows_warned"]
            ),
            "warning_count": len(summary["warnings"])
                + summary["warnings_truncated"],
        }

        summary_header = quote(
            json.dumps(public_summary, ensure_ascii=False),
            safe="",
        )

        return StreamingResponse(
            BytesIO(workbook_bytes),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition": (
                    f'attachment; filename="{summary["filename"]}"'
                ),
                "X-FSS-Summary": summary_header,
            },
        )

    except FSSBuildError as exc:
        raise HTTPException(
            status_code=400,
            detail=str(exc),
        ) from exc

    except HTTPException:
        raise

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=(
                "Terjadi kesalahan saat membangun Data FSS. "
                "Periksa log backend untuk detailnya."
            ),
        ) from exc

    finally:
        for file in files:
            await file.close()
