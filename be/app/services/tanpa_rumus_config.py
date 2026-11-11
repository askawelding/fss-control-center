"""Penyimpanan konfigurasi kelompok receiving untuk fitur Tanpa Rumus.

Konfigurasi disimpan di SQLite lokal. Grup yang dihapus dinonaktifkan (soft delete)
agar riwayat konfigurasi proses lama tetap dapat ditelusuri.
"""
from __future__ import annotations

import json
import os
import re
import sqlite3
import uuid
from pathlib import Path
from typing import Any

DEFAULT_GROUPS: list[dict[str, Any]] = [
    {
        "id": "vietnam",
        "name": "Vietnam",
        "codes": ["7R01"],
        "outputLabel": "7R01 VIETNAM",
    },
    {
        "id": "malaysia-l300",
        "name": "Malaysia & L300",
        "codes": ["7R02", "7R03"],
        "outputLabel": "7R02 MALAYSIA & 7R03 L300",
    },
    {
        "id": "thailand",
        "name": "Thailand",
        "codes": ["7R06"],
        "outputLabel": "7R06 THAILAND",
    },
    {
        "id": "bangladesh-rv",
        "name": "Bangladesh & RV Malaysia",
        "codes": ["7R07", "7R08", "7R10"],
        "outputLabel": "7R07 BANGLADESH & 7R08 7R10 RV MALAYSIA",
    },
    {
        "id": "cmc-taiwan",
        "name": "CMC Taiwan",
        "codes": ["7R09", "7R11"],
        "outputLabel": "7R09 7R11 CMC TAIWAN",
    },
]


def database_path() -> Path:
    configured = os.environ.get("FSS_DATA_DIR")
    base = Path(configured).expanduser() if configured else Path(__file__).resolve().parents[2] / "data"
    base.mkdir(parents=True, exist_ok=True)
    return base / "fss_control_center.db"


def _connect() -> sqlite3.Connection:
    connection = sqlite3.connect(database_path(), timeout=20)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def normalize_codes(codes: list[str] | str) -> list[str]:
    raw = re.split(r"[,;\n]+", codes) if isinstance(codes, str) else codes
    result: list[str] = []
    seen: set[str] = set()
    for item in raw:
        code = str(item).strip().upper()
        if not code:
            continue
        if not re.fullmatch(r"[A-Z0-9]+", code):
            raise ValueError(f"Format kode receiving tidak valid: {code}")
        if code not in seen:
            result.append(code)
            seen.add(code)
    if not result:
        raise ValueError("Masukkan minimal satu kode receiving.")
    return result


def _row_to_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "name": row["name"],
        "codes": json.loads(row["codes_json"]),
        "outputLabel": row["output_label"],
        "active": bool(row["active"]),
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
    }


def initialize_database() -> None:
    with _connect() as db:
        db.execute(
            """CREATE TABLE IF NOT EXISTS receiving_groups (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                codes_json TEXT NOT NULL,
                output_label TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        db.execute(
            """CREATE TABLE IF NOT EXISTS tanpa_rumus_runs (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                source_files_json TEXT NOT NULL,
                selected_outputs_json TEXT NOT NULL,
                group_snapshot_json TEXT NOT NULL,
                status TEXT NOT NULL,
                report_json TEXT NOT NULL
            )"""
        )
        count = db.execute("SELECT COUNT(*) FROM receiving_groups WHERE active = 1").fetchone()[0]
        if count == 0:
            for index, group in enumerate(DEFAULT_GROUPS):
                db.execute(
                    """INSERT OR IGNORE INTO receiving_groups
                    (id, name, codes_json, output_label, active, sort_order)
                    VALUES (?, ?, ?, ?, 1, ?)""",
                    (
                        group["id"],
                        group["name"],
                        json.dumps(group["codes"]),
                        group["outputLabel"],
                        index,
                    ),
                )


def list_groups(active_only: bool = True) -> list[dict[str, Any]]:
    initialize_database()
    with _connect() as db:
        where = "WHERE active = 1" if active_only else ""
        rows = db.execute(
            f"SELECT * FROM receiving_groups {where} ORDER BY sort_order, name COLLATE NOCASE"
        ).fetchall()
    return [_row_to_dict(row) for row in rows]


def _validate_group_payload(payload: dict[str, Any]) -> tuple[str, list[str], str]:
    name = str(payload.get("name", "")).strip()
    output_label = str(payload.get("outputLabel", payload.get("output_label", ""))).strip()
    codes = normalize_codes(payload.get("codes", []))
    if not name:
        raise ValueError("Nama kelompok wajib diisi.")
    if not output_label:
        raise ValueError("Label nama file output wajib diisi.")
    if len(name) > 80 or len(output_label) > 140:
        raise ValueError("Nama kelompok atau label output terlalu panjang.")
    if any(char in output_label for char in '<>:"/\\|?*'):
        raise ValueError("Label output mengandung karakter yang tidak diperbolehkan pada nama file.")
    return name, codes, output_label


def save_group(payload: dict[str, Any], group_id: str | None = None) -> dict[str, Any]:
    initialize_database()
    name, codes, output_label = _validate_group_payload(payload)
    group_id = group_id or f"group-{uuid.uuid4().hex[:12]}"

    with _connect() as db:
        rows = db.execute(
            "SELECT id, codes_json FROM receiving_groups WHERE active = 1 AND id != ?",
            (group_id,),
        ).fetchall()
        conflicts: list[str] = []
        wanted = set(codes)
        for row in rows:
            existing = set(json.loads(row["codes_json"]))
            conflicts.extend(sorted(wanted.intersection(existing)))
        if conflicts:
            raise ValueError(
                "Kode receiving sudah digunakan kelompok lain: " + ", ".join(sorted(set(conflicts)))
            )

        existing = db.execute("SELECT id, sort_order FROM receiving_groups WHERE id = ?", (group_id,)).fetchone()
        if existing:
            db.execute(
                """UPDATE receiving_groups
                   SET name=?, codes_json=?, output_label=?, active=1, updated_at=CURRENT_TIMESTAMP
                   WHERE id=?""",
                (name, json.dumps(codes), output_label, group_id),
            )
        else:
            next_order = db.execute("SELECT COALESCE(MAX(sort_order), -1) + 1 FROM receiving_groups").fetchone()[0]
            db.execute(
                """INSERT INTO receiving_groups
                   (id, name, codes_json, output_label, active, sort_order)
                   VALUES (?, ?, ?, ?, 1, ?)""",
                (group_id, name, json.dumps(codes), output_label, next_order),
            )
        row = db.execute("SELECT * FROM receiving_groups WHERE id=?", (group_id,)).fetchone()
    return _row_to_dict(row)


def delete_group(group_id: str) -> bool:
    initialize_database()
    with _connect() as db:
        cursor = db.execute(
            "UPDATE receiving_groups SET active=0, updated_at=CURRENT_TIMESTAMP WHERE id=? AND active=1",
            (group_id,),
        )
        return cursor.rowcount > 0


def record_run(
    source_files: list[str],
    selected_outputs: list[dict[str, Any]],
    group_snapshot: list[dict[str, Any]],
    status: str,
    report: dict[str, Any],
) -> None:
    initialize_database()
    with _connect() as db:
        db.execute(
            """INSERT INTO tanpa_rumus_runs
               (id, source_files_json, selected_outputs_json, group_snapshot_json, status, report_json)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                uuid.uuid4().hex,
                json.dumps(source_files, ensure_ascii=False),
                json.dumps(selected_outputs, ensure_ascii=False),
                json.dumps(group_snapshot, ensure_ascii=False),
                status,
                json.dumps(report, ensure_ascii=False),
            ),
        )


def list_runs(limit: int = 50) -> list[dict[str, Any]]:
    initialize_database()
    limit = max(1, min(int(limit), 200))
    with _connect() as db:
        rows = db.execute(
            "SELECT * FROM tanpa_rumus_runs ORDER BY created_at DESC LIMIT ?", (limit,)
        ).fetchall()
    return [
        {
            "id": row["id"],
            "createdAt": row["created_at"],
            "sourceFiles": json.loads(row["source_files_json"]),
            "selectedOutputs": json.loads(row["selected_outputs_json"]),
            "groups": json.loads(row["group_snapshot_json"]),
            "status": row["status"],
            "report": json.loads(row["report_json"]),
        }
        for row in rows
    ]
