"""Endpoints for validating, previewing, and updating a flexible list of SCH workbooks."""
from __future__ import annotations

import shutil
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

from app.services.sch_reconciliation import (
    SCHReconciliationError,
    preview_schedule_reconciliation,
)
from app.services.sch_validator import SCHValidationError, inspect_sch_workbook
from app.services.sch_writer import ExcelUnavailable, SCHUpdateError, create_update_zip

router = APIRouter(prefix="/api/schedule", tags=["Update SCH Data"])
MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024
MAX_SCHEDULE_FILES = 20


async def _read_upload(upload: UploadFile, label: str) -> tuple[str, bytes]:
    filename = upload.filename or f"{label}.xlsx"
    content = await upload.read(MAX_FILE_SIZE_BYTES + 1)
    if len(content) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"{filename}: ukuran file melebihi batas 50 MB.",
        )
    if not content:
        raise HTTPException(status_code=422, detail=f"{filename}: file kosong.")
    return filename, content


def _check_schedule_count(schedule_files: list[UploadFile]) -> None:
    if not schedule_files:
        raise HTTPException(
            status_code=422,
            detail="Pilih minimal satu workbook schedule untuk diproses.",
        )
    if len(schedule_files) > MAX_SCHEDULE_FILES:
        raise HTTPException(
            status_code=413,
            detail=f"Maksimal {MAX_SCHEDULE_FILES} workbook schedule dapat diproses sekaligus.",
        )


async def _read_schedule_inputs(
    schedule_files: list[UploadFile],
) -> dict[str, tuple[str, bytes]]:
    """Read repeated schedule_files multipart fields into stable generic workbook IDs."""
    _check_schedule_count(schedule_files)
    inputs: dict[str, tuple[str, bytes]] = {}
    for index, upload in enumerate(schedule_files, start=1):
        role = f"workbook_{index}"
        inputs[role] = await _read_upload(upload, f"Workbook {index}")
    return inputs


@router.post("/validate")
async def validate_sch_workbooks(
    schedule_files: list[UploadFile] = File(...),
):
    """Validate one or more selected DATE workbooks. This endpoint is read-only."""
    _check_schedule_count(schedule_files)
    results = []
    try:
        for index, upload in enumerate(schedule_files, start=1):
            role = f"workbook_{index}"
            filename, content = await _read_upload(upload, f"Workbook {index}")
            try:
                item = inspect_sch_workbook(content, filename, role)
                item["label"] = f"Workbook {index}"
                results.append(item)
            except SCHValidationError as exc:
                results.append({
                    "role": role,
                    "label": f"Workbook {index}",
                    "filename": filename,
                    "valid": False,
                    "message": str(exc),
                    "warnings": [],
                })
    finally:
        for upload in schedule_files:
            await upload.close()

    total_files = len(results)
    valid_count = sum(1 for result in results if result.get("valid"))
    is_valid = valid_count == total_files
    return {
        "status": "ok" if is_valid else "needs_review",
        "valid": is_valid,
        "total_files": total_files,
        "valid_files": valid_count,
        "invalid_files": total_files - valid_count,
        "files": results,
        "note": "Pemeriksaan ini hanya membaca file. Belum ada perubahan atau output Excel yang dibuat.",
    }


@router.post("/preview")
async def preview_sch_reconciliation(
    master_file: UploadFile = File(...),
    schedule_files: list[UploadFile] = File(...),
):
    """Preview the selected schedule workbooks against one Data FSS master."""
    _check_schedule_count(schedule_files)
    try:
        master_name, master_content = await _read_upload(master_file, "Data FSS master")
        schedule_inputs = await _read_schedule_inputs(schedule_files)
        try:
            return preview_schedule_reconciliation(master_content, master_name, schedule_inputs)
        except SCHReconciliationError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
    finally:
        await master_file.close()
        for upload in schedule_files:
            await upload.close()


@router.post("/update")
async def update_sch_workbooks(
    master_file: UploadFile = File(...),
    schedule_files: list[UploadFile] = File(...),
    confirm_update: bool = Form(False),
):
    """Update only the selected one-to-twenty workbooks and return them as one ZIP.

    Original uploaded workbooks are never opened for writing. A selected subset is
    a complete job; files that are not selected are simply not touched.
    """
    _check_schedule_count(schedule_files)
    if not confirm_update:
        raise HTTPException(
            status_code=400,
            detail="Konfirmasi proses update belum diberikan. Jalankan dari tombol Update setelah preview diperiksa.",
        )

    work_dir: Path | None = None
    try:
        master_name, master_content = await _read_upload(master_file, "Data FSS master")
        schedule_inputs = await _read_schedule_inputs(schedule_files)

        try:
            zip_path, work_dir, _manifest = await run_in_threadpool(
                create_update_zip,
                master_content,
                master_name,
                schedule_inputs,
            )
        except ExcelUnavailable as exc:
            raise HTTPException(status_code=503, detail=str(exc)) from exc
        except SCHUpdateError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
        except SCHReconciliationError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

        return FileResponse(
            path=zip_path,
            media_type="application/zip",
            filename="FSS_Control_Center_Update_SCH.zip",
            background=BackgroundTask(shutil.rmtree, str(work_dir), ignore_errors=True),
        )
    finally:
        await master_file.close()
        for upload in schedule_files:
            await upload.close()
