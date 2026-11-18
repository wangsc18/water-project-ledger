# -*- coding: utf-8 -*-
"""
SQLite 本地电子台账元数据底座
支持文件移动/重命名指纹追踪与 Excel 手工修订字段锁定保护
"""
from __future__ import annotations

import hashlib
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional


@dataclass
class LedgerRecord:
    id: Optional[int]
    rel_path: str
    project_name: str
    stage: str
    is_pending: int
    pending_reason: str
    archived_filename: str
    original_filename: str
    file_ext: str
    doc_date: str
    date_source: str
    doc_type: str
    milestone_key: str
    doc_no: str
    doc_no_manual: int
    amount_wan: str
    amount_manual: int
    remark: str
    extract_method: str
    sha256: str
    file_size: int
    mtime_ns: int
    archived_at: str


def compute_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


class LedgerDB:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_schema()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_schema(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rel_path TEXT UNIQUE NOT NULL,
                    project_name TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    is_pending INTEGER NOT NULL DEFAULT 0,
                    pending_reason TEXT DEFAULT '',
                    archived_filename TEXT NOT NULL,
                    original_filename TEXT NOT NULL,
                    file_ext TEXT NOT NULL,
                    doc_date TEXT NOT NULL,
                    date_source TEXT DEFAULT '',
                    doc_type TEXT NOT NULL,
                    milestone_key TEXT DEFAULT '',
                    doc_no TEXT DEFAULT '',
                    doc_no_manual INTEGER NOT NULL DEFAULT 0,
                    amount_wan TEXT DEFAULT '',
                    amount_manual INTEGER NOT NULL DEFAULT 0,
                    remark TEXT DEFAULT '',
                    extract_method TEXT DEFAULT '',
                    sha256 TEXT NOT NULL,
                    file_size INTEGER NOT NULL,
                    mtime_ns INTEGER NOT NULL,
                    archived_at TEXT NOT NULL
                )
                """
            )
            conn.execute("CREATE INDEX IF NOT EXISTS idx_docs_proj_stage ON documents(project_name, stage)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_docs_sha256 ON documents(sha256, file_size)")
            conn.commit()

    def _row_to_record(self, row: sqlite3.Row) -> LedgerRecord:
        return LedgerRecord(**dict(row))

    def list_all(self) -> List[LedgerRecord]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM documents ORDER BY is_pending ASC, project_name ASC, stage ASC, doc_date ASC, id ASC"
            ).fetchall()
            return [self._row_to_record(r) for r in rows]

    def get_by_id(self, rec_id: int) -> Optional[LedgerRecord]:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM documents WHERE id = ?", (rec_id,)).fetchone()
            return self._row_to_record(row) if row else None

    def get_by_rel_path(self, rel_path: str) -> Optional[LedgerRecord]:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM documents WHERE rel_path = ?", (rel_path,)).fetchone()
            return self._row_to_record(row) if row else None

    def insert(self, rec: LedgerRecord) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO documents (
                    rel_path, project_name, stage, is_pending, pending_reason,
                    archived_filename, original_filename, file_ext,
                    doc_date, date_source, doc_type, milestone_key,
                    doc_no, doc_no_manual, amount_wan, amount_manual, remark,
                    extract_method, sha256, file_size, mtime_ns, archived_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    rec.rel_path, rec.project_name, rec.stage, rec.is_pending, rec.pending_reason,
                    rec.archived_filename, rec.original_filename, rec.file_ext,
                    rec.doc_date, rec.date_source, rec.doc_type, rec.milestone_key,
                    rec.doc_no, rec.doc_no_manual, rec.amount_wan, rec.amount_manual, rec.remark,
                    rec.extract_method, rec.sha256, rec.file_size, rec.mtime_ns, rec.archived_at,
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    def update_moved_or_renamed(
        self,
        rec_id: int,
        rel_path: str,
        project_name: str,
        stage: str,
        is_pending: int,
        pending_reason: str,
        archived_filename: str,
        doc_date: str,
        doc_type: str,
        milestone_key: str,
        mtime_ns: int,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE documents
                SET rel_path = ?, project_name = ?, stage = ?, is_pending = ?, pending_reason = ?,
                    archived_filename = ?, doc_date = ?, doc_type = ?, milestone_key = ?, mtime_ns = ?
                WHERE id = ?
                """,
                (
                    rel_path, project_name, stage, is_pending, pending_reason,
                    archived_filename, doc_date, doc_type, milestone_key, mtime_ns, rec_id,
                ),
            )
            conn.commit()

    def update_manual_fields(
        self,
        rec_id: int,
        doc_no: Optional[str] = None,
        amount_wan: Optional[str] = None,
        remark: Optional[str] = None,
    ) -> None:
        rec = self.get_by_id(rec_id)
        if not rec:
            return
        new_doc_no = rec.doc_no
        new_doc_no_manual = rec.doc_no_manual
        if doc_no is not None and doc_no != rec.doc_no:
            new_doc_no = doc_no
            new_doc_no_manual = 1

        new_amount = rec.amount_wan
        new_amount_manual = rec.amount_manual
        if amount_wan is not None and amount_wan != rec.amount_wan:
            new_amount = amount_wan
            new_amount_manual = 1

        new_remark = rec.remark if remark is None else remark

        with self._connect() as conn:
            conn.execute(
                """
                UPDATE documents
                SET doc_no = ?, doc_no_manual = ?, amount_wan = ?, amount_manual = ?, remark = ?
                WHERE id = ?
                """,
                (new_doc_no, new_doc_no_manual, new_amount, new_amount_manual, new_remark, rec_id),
            )
            conn.commit()

    def delete_ids(self, ids: List[int]) -> None:
        if not ids:
            return
        with self._connect() as conn:
            placeholders = ",".join("?" for _ in ids)
            conn.execute(f"DELETE FROM documents WHERE id IN ({placeholders})", ids)
            conn.commit()

    def clear_all(self) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM documents")
            conn.execute("DELETE FROM sqlite_sequence WHERE name='documents'")
            conn.commit()
