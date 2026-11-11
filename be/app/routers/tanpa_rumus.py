"""API Tanpa Rumus: finalisasi workbook MMKI/CKD dan konfigurasi receiving group."""
from __future__ import annotations

import json
import logging
import shutil
import time
from pathlib import Path
from urllib.parse import quote

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

from app.services.tanpa_rumus import (
    MAX_FILE_SIZE_BYTES, MAX_FILES, ExcelUnavailable, TanpaRumusError,
    WorkbookInput, create_tanpa_rumus_zip,
)
from app.services.tanpa_rumus_config import (
    delete_group as delete_group_config,
    list_groups,
    list_runs,
    record_run,
    save_group,
)
from app.services.tanpa_rumus_delivery import (
    DeliveryWorkbookInput,
    create_delivery_zip,
    inspect_delivery_source,
    inspect_date_sheet,
)

logger = logging.getLogger("fss.tanpa_rumus")
router = APIRouter(prefix="/api/tanpa-rumus", tags=["Tanpa Rumus"])


class ReceivingGroupPayload(BaseModel):
    name: str = Field(min_length=1, max_length=80)
    codes: list[str] = Field(min_length=1)
    outputLabel: str = Field(min_length=1, max_length=140)


@router.get("/groups")
def get_receiving_groups():
    return {"groups": list_groups()}


@router.post("/groups", status_code=201)
def create_receiving_group(payload: ReceivingGroupPayload):
    try:
        group = save_group(payload.model_dump())
        return {"group": group}
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.put("/groups/{group_id}")
def update_receiving_group(group_id: str, payload: ReceivingGroupPayload):
    if not any(group["id"] == group_id for group in list_groups()):
        raise HTTPException(status_code=404, detail="Kelompok receiving tidak ditemukan.")
    try:
        group = save_group(payload.model_dump(), group_id=group_id)
        return {"group": group}
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@router.delete("/groups/{group_id}")
def remove_receiving_group(group_id: str):
    if not delete_group_config(group_id):
        raise HTTPException(status_code=404, detail="Kelompok receiving tidak ditemukan.")
    return {"deleted": True, "id": group_id}


@router.get("/history")
def get_tanpa_rumus_history(limit: int = 50):
    return {"items": list_runs(limit)}


@router.post("/preview-source")
async def preview_tanpa_rumus_source(workbook_file: UploadFile = File(...)):
    """Periksa struktur dasar workbook delivery untuk MMKI maupun CKD."""
    try:
        source = await _read_upload(workbook_file, "delivery", required=True)
        assert source is not None
        started = time.perf_counter()
        result = await run_in_threadpool(inspect_delivery_source, source.content, source.filename)
        result["processing_seconds"] = round(time.perf_counter() - started, 2)
        return result
    except TanpaRumusError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        logger.exception("Pemeriksaan struktur workbook Tanpa Rumus gagal")
        raise HTTPException(status_code=500, detail="Gagal membaca struktur workbook. Periksa log backend.") from exc
    finally:
        await workbook_file.close()


async def _read_upload(upload: UploadFile | None, label: str, required: bool) -> DeliveryWorkbookInput | None:
    if upload is None:
        if required:
            raise HTTPException(status_code=422, detail=f"Pilih workbook {label} terlebih dahulu.")
        return None
    filename = Path(upload.filename or f"{label}.xlsx").name
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise HTTPException(status_code=422, detail=f"File {label} harus berformat .xlsx atau .xlsm.")
    content = await upload.read(MAX_FILE_SIZE_BYTES + 1)
    if not content:
        raise HTTPException(status_code=422, detail=f"File {label} kosong.")
    if len(content) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(status_code=413, detail=f"File {label} melebihi batas 50 MB.")
    return DeliveryWorkbookInput(filename=filename, content=content)


@router.post("/preview-delivery")
async def preview_tanpa_rumus_delivery(ckd_file: UploadFile = File(...)):
    """Baca sheet Date CKD dan tampilkan kode receiving serta kode negara yang belum dipetakan."""
    try:
        ckd = await _read_upload(ckd_file, "CKD", required=True)
        assert ckd is not None
        started = time.perf_counter()
        preview = await run_in_threadpool(inspect_date_sheet, ckd.content, ckd.filename)
        groups = list_groups()
        result = preview.to_dict(groups)
        result["processing_seconds"] = round(time.perf_counter() - started, 2)
        return result
    except TanpaRumusError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        logger.exception("Preview Tanpa Rumus gagal")
        raise HTTPException(status_code=500, detail="Gagal membaca worksheet Date. Periksa log backend.") from exc
    finally:
        await ckd_file.close()


@router.post("/process-delivery")
async def process_tanpa_rumus_delivery(
    mmki_file: UploadFile | None = File(None),
    ckd_file: UploadFile | None = File(None),
    make_mmki: bool = Form(False),
    make_ckd_all: bool = Form(False),
    make_ckd_groups: bool = Form(False),
    selected_group_ids: str = Form("[]"),
):
    """Buat output MMKI keseluruhan, CKD keseluruhan, dan/atau CKD per kelompok receiving."""
    work_dir: Path | None = None
    all_groups: list[dict] = []
    selected_groups: list[dict] = []
    mmki: DeliveryWorkbookInput | None = None
    ckd: DeliveryWorkbookInput | None = None
    try:
        if not (make_mmki or make_ckd_all or make_ckd_groups):
            raise HTTPException(status_code=422, detail="Pilih minimal satu jenis output.")
        mmki = await _read_upload(mmki_file, "MMKI", required=make_mmki)
        ckd_required = make_ckd_all or make_ckd_groups
        ckd = await _read_upload(ckd_file, "CKD", required=ckd_required)

        try:
            selected_ids = json.loads(selected_group_ids or "[]")
            if not isinstance(selected_ids, list) or not all(isinstance(value, str) for value in selected_ids):
                raise ValueError("Format pilihan kelompok tidak valid.")
        except (json.JSONDecodeError, ValueError) as exc:
            raise HTTPException(status_code=422, detail="Pilihan kelompok receiving tidak valid.") from exc

        all_groups = list_groups()
        groups_by_id = {group["id"]: group for group in all_groups}
        if make_ckd_groups:
            if not selected_ids:
                raise HTTPException(status_code=422, detail="Pilih minimal satu kelompok receiving.")
            missing = [group_id for group_id in selected_ids if group_id not in groups_by_id]
            if missing:
                raise HTTPException(status_code=422, detail="Ada kelompok yang tidak ditemukan atau sudah dinonaktifkan. Muat ulang konfigurasi.")
        selected_groups = [groups_by_id[group_id] for group_id in selected_ids] if make_ckd_groups else []

        try:
            zip_path, work_dir, manifest = await run_in_threadpool(
                create_delivery_zip,
                mmki,
                ckd,
                make_mmki,
                make_ckd_all,
                make_ckd_groups,
                selected_groups,
                all_groups,
            )
        except ExcelUnavailable as exc:
            record_run(
                source_files=[item.filename for item in (mmki, ckd) if item is not None],
                selected_outputs=(
                    ([{"kind": "mmki_all"}] if make_mmki else [])
                    + ([{"kind": "ckd_all"}] if make_ckd_all else [])
                    + [
                        {"kind": "ckd_group", "group_id": group["id"], "name": group["name"], "codes": group["codes"], "outputLabel": group["outputLabel"]}
                        for group in selected_groups
                    ]
                ),
                group_snapshot=all_groups,
                status="failed",
                report={"status": "failed", "error": str(exc)},
            )
            raise HTTPException(status_code=503, detail=str(exc)) from exc
        except TanpaRumusError as exc:
            record_run(
                source_files=[item.filename for item in (mmki, ckd) if item is not None],
                selected_outputs=(
                    ([{"kind": "mmki_all"}] if make_mmki else [])
                    + ([{"kind": "ckd_all"}] if make_ckd_all else [])
                    + [
                        {"kind": "ckd_group", "group_id": group["id"], "name": group["name"], "codes": group["codes"], "outputLabel": group["outputLabel"]}
                        for group in selected_groups
                    ]
                ),
                group_snapshot=all_groups,
                status="failed",
                report={"status": "failed", "error": str(exc)},
            )
            raise HTTPException(status_code=422, detail=str(exc)) from exc

        summary = {
            "total_outputs": manifest["total_outputs"],
            "succeeded": manifest["succeeded"],
            "failed": manifest["failed"],
            "status": manifest["status"],
            "processing_seconds": manifest.get("processing_seconds"),
            "preview_seconds": manifest.get("preview_seconds"),
            "outputs": [
                {
                    "name": item.get("output"),
                    "status": item.get("status"),
                    "description": item.get("description"),
                    "processing_seconds": item.get("processing_seconds"),
                    "calculation_seconds": item.get("calculation_seconds"),
                    "calculation_mode": item.get("calculation_mode"),
                    "error": item.get("error"),
                }
                for item in manifest["files"]
            ],
        }
        record_run(
            source_files=[item.filename for item in (mmki, ckd) if item is not None],
            selected_outputs=[
                *([{"kind": "mmki_all"}] if make_mmki else []),
                *([{"kind": "ckd_all"}] if make_ckd_all else []),
                *[
                    {"kind": "ckd_group", "group_id": group["id"], "name": group["name"], "codes": group["codes"], "outputLabel": group["outputLabel"]}
                    for group in selected_groups
                ],
            ],
            group_snapshot=all_groups,
            status=manifest["status"],
            report=manifest,
        )
        summary_header = quote(json.dumps(summary, ensure_ascii=False), safe="")
        return FileResponse(
            path=zip_path,
            media_type="application/zip",
            filename="FSS_Control_Center_Tanpa_Rumus.zip",
            headers={"X-Tanpa-Rumus-Summary": summary_header},
            background=BackgroundTask(shutil.rmtree, str(work_dir), ignore_errors=True),
        )
    except HTTPException:
        if work_dir is not None:
            shutil.rmtree(work_dir, ignore_errors=True)
        raise
    except Exception as exc:  # noqa: BLE001
        logger.exception("Tanpa Rumus delivery gagal")
        if work_dir is not None:
            shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail="Tanpa Rumus gagal. Periksa log backend untuk detailnya.") from exc
    finally:
        if mmki_file is not None:
            await mmki_file.close()
        if ckd_file is not None:
            await ckd_file.close()


# Endpoint lama dipertahankan agar kompatibilitas versi frontend sebelumnya tidak rusak.
async def _read_legacy_inputs(files: list[UploadFile]) -> list[WorkbookInput]:
    inputs: list[WorkbookInput] = []
    for index, upload in enumerate(files, start=1):
        filename = Path(upload.filename or f"workbook_{index}.xlsx").name
        content = await upload.read(MAX_FILE_SIZE_BYTES + 1)
        error = None
        if len(content) > MAX_FILE_SIZE_BYTES:
            error = f"{filename}: ukuran file melebihi batas 50 MB."
            content = b""
        elif not content:
            error = f"{filename}: file kosong."
        inputs.append(WorkbookInput(filename, content, error))
    return inputs


@router.post("/process")
async def process_legacy_tanpa_rumus(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=422, detail="Pilih minimal satu workbook untuk diproses.")
    if len(files) > MAX_FILES:
        raise HTTPException(status_code=413, detail=f"Maksimal {MAX_FILES} workbook dapat diproses sekaligus.")
    work_dir = None
    try:
        inputs = await _read_legacy_inputs(files)
        try:
            zip_path, work_dir, manifest = await run_in_threadpool(create_tanpa_rumus_zip, inputs)
        except ExcelUnavailable as exc:
            raise HTTPException(status_code=503, detail=str(exc)) from exc
        except TanpaRumusError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

        if manifest["succeeded"] == 0:
            shutil.rmtree(work_dir, ignore_errors=True)
            reasons = "; ".join(item["error"] for item in manifest["files"] if item.get("error"))
            return JSONResponse(
                status_code=422,
                content={
                    "detail": "Tidak ada workbook yang berhasil diproses. " + reasons,
                    "report": manifest,
                },
            )

        summary = {
            "total_files": manifest["total_files"],
            "succeeded": manifest["succeeded"],
            "failed": manifest["failed"],
            "formulas_converted": manifest["total_formulas_converted"],
            "warning_count": manifest["warning_count"],
        }
        return FileResponse(
            path=zip_path,
            media_type="application/zip",
            filename="FSS_Control_Center_Tanpa_Rumus.zip",
            headers={"X-Tanpa-Rumus-Summary": quote(json.dumps(summary, ensure_ascii=False), safe="")},
            background=BackgroundTask(shutil.rmtree, str(work_dir), ignore_errors=True),
        )
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        logger.exception("Tanpa Rumus legacy gagal")
        if work_dir is not None:
            shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail="Terjadi kesalahan saat memproses Tanpa Rumus. Periksa log backend.") from exc
    finally:
        for upload in files:
            await upload.close()
