from __future__ import annotations

import csv
import os
import re
import shutil
import sqlite3
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, X, Y, filedialog, messagebox, simpledialog, StringVar, Tk, Toplevel
import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook


APP_TITLE = "CIVIL HUB"
APP_SUBTITLE = "施工体制台帳作成・添付書類チェックシステム"
DB_PATH = Path(__file__).with_name("civilhub.db")
ATTACHMENTS_ROOT = Path(__file__).with_name("storage") / "attachments"
DOCUMENT_TEMPLATES = [
    "再下請通知書",
    "建設業許可証",
    "主任技術者資格証",
    "主任技術者の雇用確認資料",
    "社会保険加入確認資料",
    "労働保険関係資料",
    "安全衛生責任者選任書",
    "作業員名簿",
    "外国人就労関係書類",
    "契約書または注文請書",
]
DOCUMENT_STATUSES = ["未確認", "不足", "添付済み", "期限切れ", "不要"]
LEVEL_LABELS = {0: "元請", 1: "一次下請", 2: "二次下請", 3: "三次下請以降"}
STATUS_COLORS = {
    "添付済み": "#1f7a38",
    "不足": "#9f2d2d",
    "期限切れ": "#9f2d2d",
    "未確認": "#5f6368",
    "不要": "#607d8b",
}


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sanitize_filename(value: str) -> str:
    forbidden = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in forbidden else ch for ch in value.strip())
    return cleaned or "document"


def provisional_construction_no(project_name: str) -> str:
    safe_name = sanitize_filename(project_name).replace(" ", "_")
    return f"CFG_{safe_name}"[:80]


def normalize_label(value: str) -> str:
    return value.replace("\n", "").replace("\r", "").replace("\u3000", "").replace(" ", "").strip()


def parse_wareki_free_date(value: str) -> str:
    match = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", value)
    if not match:
        return value.strip()
    year, month, day = match.groups()
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"


def format_ledger_date(prefix: str, value: str) -> str:
    if not value:
        return ""
    match = re.match(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$", value)
    if match:
        year, month, day = match.groups()
        return f" {prefix} {int(year)}年{int(month)}月{int(day)}日"
    return value


def build_reflected_copy_path(source_path: Path) -> Path:
    base = source_path.with_name(f"{source_path.stem}_civilhub{source_path.suffix}")
    if not base.exists():
        return base
    version = 2
    while True:
        candidate = source_path.with_name(f"{source_path.stem}_civilhub_v{version}{source_path.suffix}")
        if not candidate.exists():
            return candidate
        version += 1


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


class Database:
    def __init__(self, db_path: Path) -> None:
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.initialize()

    def initialize(self) -> None:
        self.conn.executescript(
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                construction_no TEXT NOT NULL,
                client_name TEXT NOT NULL,
                start_date TEXT,
                end_date TEXT,
                site_agent TEXT,
                managing_engineer TEXT,
                chief_engineer TEXT,
                base_folder TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                parent_company_id INTEGER,
                level INTEGER NOT NULL,
                name TEXT NOT NULL,
                kana TEXT,
                representative TEXT,
                address TEXT,
                phone TEXT,
                license_no TEXT,
                license_expiry TEXT,
                work_type TEXT,
                contract_date TEXT,
                planned_start_date TEXT,
                planned_end_date TEXT,
                chief_engineer_name TEXT,
                chief_engineer_license TEXT,
                safety_manager TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
                FOREIGN KEY(parent_company_id) REFERENCES companies(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS required_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_id INTEGER NOT NULL,
                document_name TEXT NOT NULL,
                required INTEGER NOT NULL DEFAULT 1,
                status TEXT NOT NULL DEFAULT '未確認',
                expiry_date TEXT,
                note TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                required_document_id INTEGER NOT NULL,
                original_path TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                stored_filename TEXT NOT NULL,
                file_type TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(required_document_id) REFERENCES required_documents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS project_imports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                import_kind TEXT NOT NULL,
                source_path TEXT NOT NULL,
                source_sheet TEXT,
                imported_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            );
            """
        )
        self.conn.commit()

    def fetchall(self, query: str, params: tuple = ()) -> list[sqlite3.Row]:
        return list(self.conn.execute(query, params))

    def fetchone(self, query: str, params: tuple = ()) -> sqlite3.Row | None:
        return self.conn.execute(query, params).fetchone()

    def execute(self, query: str, params: tuple = ()) -> sqlite3.Cursor:
        cur = self.conn.execute(query, params)
        self.conn.commit()
        return cur

    def create_project(self, data: dict[str, str]) -> int:
        ts = now_text()
        cur = self.execute(
            """
            INSERT INTO projects (
                name, construction_no, client_name, start_date, end_date,
                site_agent, managing_engineer, chief_engineer, base_folder,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                data["name"],
                data["construction_no"],
                data["client_name"],
                data["start_date"],
                data["end_date"],
                data["site_agent"],
                data["managing_engineer"],
                data["chief_engineer"],
                data["base_folder"],
                ts,
                ts,
            ),
        )
        return int(cur.lastrowid)

    def list_projects(self) -> list[sqlite3.Row]:
        return self.fetchall(
            """
            SELECT *
            FROM projects
            ORDER BY updated_at DESC, id DESC
            """
        )

    def find_project_by_construction_no(self, construction_no: str) -> sqlite3.Row | None:
        return self.fetchone(
            """
            SELECT *
            FROM projects
            WHERE construction_no = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (construction_no,),
        )

    def update_project(self, project_id: int, data: dict[str, str]) -> None:
        self.execute(
            """
            UPDATE projects
            SET name = ?,
                construction_no = ?,
                client_name = ?,
                start_date = ?,
                end_date = ?,
                site_agent = ?,
                managing_engineer = ?,
                chief_engineer = ?,
                base_folder = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                data["name"],
                data["construction_no"],
                data["client_name"],
                data["start_date"],
                data["end_date"],
                data["site_agent"],
                data["managing_engineer"],
                data["chief_engineer"],
                data["base_folder"],
                now_text(),
                project_id,
            ),
        )

    def delete_project(self, project_id: int) -> None:
        self.execute("DELETE FROM projects WHERE id = ?", (project_id,))

    def reset_all_data(self) -> None:
        self.conn.execute("PRAGMA foreign_keys = OFF")
        try:
            self.conn.execute("DELETE FROM attachments")
            self.conn.execute("DELETE FROM required_documents")
            self.conn.execute("DELETE FROM companies")
            self.conn.execute("DELETE FROM project_imports")
            self.conn.execute("DELETE FROM projects")
            self.conn.commit()
        finally:
            self.conn.execute("PRAGMA foreign_keys = ON")

    def find_project_by_name(self, name: str) -> sqlite3.Row | None:
        return self.fetchone(
            """
            SELECT *
            FROM projects
            WHERE name = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (name,),
        )

    def upsert_project_import(self, project_id: int, import_kind: str, source_path: str, source_sheet: str) -> None:
        row = self.fetchone(
            """
            SELECT id
            FROM project_imports
            WHERE project_id = ? AND import_kind = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (project_id, import_kind),
        )
        ts = now_text()
        if row is None:
            self.execute(
                """
                INSERT INTO project_imports (
                    project_id, import_kind, source_path, source_sheet, imported_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (project_id, import_kind, source_path, source_sheet, ts, ts),
            )
            return
        self.execute(
            """
            UPDATE project_imports
            SET source_path = ?, source_sheet = ?, updated_at = ?
            WHERE id = ?
            """,
            (source_path, source_sheet, ts, row["id"]),
        )

    def get_latest_project_import(self, project_id: int, import_kind: str) -> sqlite3.Row | None:
        return self.fetchone(
            """
            SELECT *
            FROM project_imports
            WHERE project_id = ? AND import_kind = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (project_id, import_kind),
        )

    def create_company(self, data: dict[str, str | int | None]) -> int:
        ts = now_text()
        cur = self.execute(
            """
            INSERT INTO companies (
                project_id, parent_company_id, level, name, kana, representative,
                address, phone, license_no, license_expiry, work_type, contract_date,
                planned_start_date, planned_end_date, chief_engineer_name,
                chief_engineer_license, safety_manager, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                data["project_id"],
                data["parent_company_id"],
                data["level"],
                data["name"],
                data["kana"],
                data["representative"],
                data["address"],
                data["phone"],
                data["license_no"],
                data["license_expiry"],
                data["work_type"],
                data["contract_date"],
                data["planned_start_date"],
                data["planned_end_date"],
                data["chief_engineer_name"],
                data["chief_engineer_license"],
                data["safety_manager"],
                ts,
                ts,
            ),
        )
        company_id = int(cur.lastrowid)
        self.generate_required_documents(company_id)
        return company_id

    def generate_required_documents(self, company_id: int) -> None:
        ts = now_text()
        existing = {
            row["document_name"]
            for row in self.fetchall(
                "SELECT document_name FROM required_documents WHERE company_id = ?",
                (company_id,),
            )
        }
        for name in DOCUMENT_TEMPLATES:
            if name in existing:
                continue
            self.execute(
                """
                INSERT INTO required_documents (
                    company_id, document_name, required, status, expiry_date, note, created_at, updated_at
                ) VALUES (?, ?, 1, '未確認', '', '', ?, ?)
                """,
                (company_id, name, ts, ts),
            )

    def list_companies(self, project_id: int) -> list[sqlite3.Row]:
        return self.fetchall(
            """
            SELECT *
            FROM companies
            WHERE project_id = ?
            ORDER BY level ASC, id ASC
            """,
            (project_id,),
        )

    def get_company(self, company_id: int) -> sqlite3.Row | None:
        return self.fetchone("SELECT * FROM companies WHERE id = ?", (company_id,))

    def list_required_documents(self, company_id: int) -> list[sqlite3.Row]:
        return self.fetchall(
            """
            SELECT rd.*,
                   COUNT(a.id) AS attachment_count
            FROM required_documents rd
            LEFT JOIN attachments a ON a.required_document_id = rd.id
            WHERE rd.company_id = ?
            GROUP BY rd.id
            ORDER BY rd.id ASC
            """,
            (company_id,),
        )

    def update_required_document(self, document_id: int, status: str, expiry_date: str, note: str) -> None:
        self.execute(
            """
            UPDATE required_documents
            SET status = ?, expiry_date = ?, note = ?, updated_at = ?
            WHERE id = ?
            """,
            (status, expiry_date, note, now_text(), document_id),
        )

    def list_shortages(self, project_id: int) -> list[sqlite3.Row]:
        return self.fetchall(
            """
            SELECT c.name AS company_name, rd.document_name, rd.status
            FROM required_documents rd
            INNER JOIN companies c ON c.id = rd.company_id
            WHERE c.project_id = ?
              AND rd.status IN ('不足', '未確認', '期限切れ')
            ORDER BY c.level ASC, c.name ASC, rd.document_name ASC
            """,
            (project_id,),
        )

    def get_project(self, project_id: int) -> sqlite3.Row | None:
        return self.fetchone("SELECT * FROM projects WHERE id = ?", (project_id,))

    def add_attachment(self, document_id: int, original_path: Path, stored_path: Path) -> int:
        cur = self.execute(
            """
            INSERT INTO attachments (
                required_document_id, original_path, stored_path, original_filename,
                stored_filename, file_type, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                document_id,
                str(original_path),
                str(stored_path),
                original_path.name,
                stored_path.name,
                original_path.suffix.lower().lstrip("."),
                now_text(),
            ),
        )
        return int(cur.lastrowid)

    def list_attachments(self, document_id: int) -> list[sqlite3.Row]:
        return self.fetchall(
            """
            SELECT *
            FROM attachments
            WHERE required_document_id = ?
            ORDER BY id DESC
            """,
            (document_id,),
        )

    def get_attachment(self, attachment_id: int) -> sqlite3.Row | None:
        return self.fetchone("SELECT * FROM attachments WHERE id = ?", (attachment_id,))

    def next_versioned_filename(self, base_dir: Path, company_name: str, document_name: str, suffix: str) -> str:
        safe_company = sanitize_filename(company_name)
        safe_document = sanitize_filename(document_name)
        version = 1
        while True:
            candidate = f"{safe_company}_{safe_document}_v{version}{suffix}"
            if not (base_dir / candidate).exists():
                return candidate
            version += 1


@dataclass
class ProjectSelection:
    project_id: int | None = None
    company_id: int | None = None
    document_id: int | None = None
    attachment_id: int | None = None


class ProjectDialog:
    def __init__(self, master: tk.Misc, title: str = "工事登録", initial_data: dict[str, str] | None = None) -> None:
        self.top = Toplevel(master)
        self.top.title(title)
        self.top.transient(master)
        self.top.grab_set()
        self.result: dict[str, str] | None = None

        fields = [
            ("工事名", "name"),
            ("工事番号", "construction_no"),
            ("発注者", "client_name"),
            ("工期開始日", "start_date"),
            ("工期終了日", "end_date"),
            ("現場代理人", "site_agent"),
            ("監理技術者", "managing_engineer"),
            ("主任技術者", "chief_engineer"),
            ("保存フォルダ", "base_folder"),
        ]
        self.vars: dict[str, StringVar] = {}
        for index, (label, key) in enumerate(fields):
            ttk.Label(self.top, text=label).grid(row=index, column=0, padx=8, pady=4, sticky="w")
            var = StringVar(value=(initial_data or {}).get(key, ""))
            self.vars[key] = var
            entry = ttk.Entry(self.top, textvariable=var, width=42)
            entry.grid(row=index, column=1, padx=8, pady=4, sticky="ew")
            if key == "base_folder":
                ttk.Button(self.top, text="参照", command=self.pick_folder).grid(row=index, column=2, padx=8, pady=4)

        button_row = len(fields)
        ttk.Button(self.top, text="保存", command=self.submit).grid(row=button_row, column=1, padx=8, pady=10, sticky="e")
        ttk.Button(self.top, text="キャンセル", command=self.top.destroy).grid(row=button_row, column=2, padx=8, pady=10, sticky="w")
        self.top.columnconfigure(1, weight=1)

    def pick_folder(self) -> None:
        folder = filedialog.askdirectory(title="保存フォルダを選択")
        if folder:
            self.vars["base_folder"].set(folder)

    def submit(self) -> None:
        required = ["name", "construction_no", "client_name"]
        for key in required:
            if not self.vars[key].get().strip():
                messagebox.showerror("入力エラー", "工事名・工事番号・発注者は必須です。", parent=self.top)
                return
        self.result = {key: var.get().strip() for key, var in self.vars.items()}
        self.top.destroy()


class CompanyDialog:
    def __init__(self, master: tk.Misc, level: int, parent_name: str | None) -> None:
        self.top = Toplevel(master)
        self.top.title("業者登録")
        self.top.transient(master)
        self.top.grab_set()
        self.result: dict[str, str] | None = None

        title = LEVEL_LABELS.get(level, "下請")
        header = title if parent_name is None else f"{title} / 親: {parent_name}"
        ttk.Label(self.top, text=header).grid(row=0, column=0, columnspan=2, padx=8, pady=(8, 12), sticky="w")

        fields = [
            ("会社名", "name"),
            ("会社名カナ", "kana"),
            ("代表者名", "representative"),
            ("所在地", "address"),
            ("電話番号", "phone"),
            ("建設業許可番号", "license_no"),
            ("許可有効期限", "license_expiry"),
            ("担当工種", "work_type"),
            ("契約日", "contract_date"),
            ("施工開始予定日", "planned_start_date"),
            ("施工終了予定日", "planned_end_date"),
            ("主任技術者名", "chief_engineer_name"),
            ("主任技術者資格", "chief_engineer_license"),
            ("安全衛生責任者", "safety_manager"),
        ]
        self.vars: dict[str, StringVar] = {}
        for index, (label, key) in enumerate(fields, start=1):
            ttk.Label(self.top, text=label).grid(row=index, column=0, padx=8, pady=4, sticky="w")
            var = StringVar()
            self.vars[key] = var
            ttk.Entry(self.top, textvariable=var, width=40).grid(row=index, column=1, padx=8, pady=4, sticky="ew")

        button_row = len(fields) + 1
        ttk.Button(self.top, text="登録", command=self.submit).grid(row=button_row, column=1, padx=8, pady=10, sticky="e")
        ttk.Button(self.top, text="キャンセル", command=self.top.destroy).grid(row=button_row, column=0, padx=8, pady=10, sticky="w")
        self.top.columnconfigure(1, weight=1)

    def submit(self) -> None:
        if not self.vars["name"].get().strip():
            messagebox.showerror("入力エラー", "会社名は必須です。", parent=self.top)
            return
        self.result = {key: var.get().strip() for key, var in self.vars.items()}
        self.top.destroy()


class DocumentStatusDialog:
    def __init__(self, master: tk.Misc, document: sqlite3.Row) -> None:
        self.top = Toplevel(master)
        self.top.title("書類状態更新")
        self.top.transient(master)
        self.top.grab_set()
        self.result: dict[str, str] | None = None

        ttk.Label(self.top, text=document["document_name"]).grid(row=0, column=0, columnspan=2, padx=8, pady=(8, 12), sticky="w")

        self.status_var = StringVar(value=document["status"])
        self.expiry_var = StringVar(value=document["expiry_date"] or "")
        self.note_var = StringVar(value=document["note"] or "")

        ttk.Label(self.top, text="状態").grid(row=1, column=0, padx=8, pady=4, sticky="w")
        ttk.Combobox(self.top, textvariable=self.status_var, values=DOCUMENT_STATUSES, state="readonly").grid(row=1, column=1, padx=8, pady=4, sticky="ew")

        ttk.Label(self.top, text="有効期限").grid(row=2, column=0, padx=8, pady=4, sticky="w")
        ttk.Entry(self.top, textvariable=self.expiry_var).grid(row=2, column=1, padx=8, pady=4, sticky="ew")

        ttk.Label(self.top, text="備考").grid(row=3, column=0, padx=8, pady=4, sticky="nw")
        self.note_text = tk.Text(self.top, width=36, height=6)
        self.note_text.grid(row=3, column=1, padx=8, pady=4, sticky="ew")
        self.note_text.insert("1.0", document["note"] or "")

        ttk.Button(self.top, text="保存", command=self.submit).grid(row=4, column=1, padx=8, pady=10, sticky="e")
        ttk.Button(self.top, text="キャンセル", command=self.top.destroy).grid(row=4, column=0, padx=8, pady=10, sticky="w")
        self.top.columnconfigure(1, weight=1)

    def submit(self) -> None:
        self.result = {
            "status": self.status_var.get(),
            "expiry_date": self.expiry_var.get().strip(),
            "note": self.note_text.get("1.0", END).strip(),
        }
        self.top.destroy()


class CsvImportConfirmDialog:
    def __init__(self, master: tk.Misc, rows: list[dict[str, str]]) -> None:
        self.top = Toplevel(master)
        self.top.title("CSV取込確認")
        self.top.transient(master)
        self.top.grab_set()
        self.approved = False

        ttk.Label(
            self.top,
            text="工事番号で照合した取込候補です。内容を確認して取り込みを実行してください。",
        ).pack(anchor="w", padx=10, pady=(10, 8))

        columns = ("action", "construction_no", "name", "client_name", "start_date", "end_date")
        tree = ttk.Treeview(self.top, columns=columns, show="headings", height=12)
        tree.heading("action", text="取込区分")
        tree.heading("construction_no", text="工事番号")
        tree.heading("name", text="工事名")
        tree.heading("client_name", text="発注者")
        tree.heading("start_date", text="開始日")
        tree.heading("end_date", text="終了日")
        tree.column("action", width=80, anchor="center")
        tree.column("construction_no", width=120, anchor="center")
        tree.column("name", width=240)
        tree.column("client_name", width=180)
        tree.column("start_date", width=100, anchor="center")
        tree.column("end_date", width=100, anchor="center")
        tree.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

        for row in rows:
            tree.insert(
                "",
                END,
                values=(
                    row["action"],
                    row["construction_no"],
                    row["name"],
                    row["client_name"],
                    row["start_date"],
                    row["end_date"],
                ),
            )

        buttons = ttk.Frame(self.top)
        buttons.pack(fill=X, padx=10, pady=(0, 10))
        ttk.Button(buttons, text="取り込む", command=self.submit).pack(side=RIGHT)
        ttk.Button(buttons, text="キャンセル", command=self.top.destroy).pack(side=RIGHT, padx=(0, 8))

    def submit(self) -> None:
        self.approved = True
        self.top.destroy()


class CivilHubApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.db = Database(DB_PATH)
        self.selection = ProjectSelection()
        self.project_map: dict[str, int] = {}
        self.company_tree_map: dict[str, int] = {}
        self.document_tree_map: dict[str, int] = {}
        self.attachment_tree_map: dict[str, int] = {}

        self.root.title(f"{APP_TITLE} - {APP_SUBTITLE}")
        self.root.geometry("1500x900")

        self.style = ttk.Style(self.root)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.build_ui()
        self.refresh_projects()

    def build_ui(self) -> None:
        header = ttk.Frame(self.root, padding=10)
        header.pack(fill=X)
        ttk.Label(header, text=APP_TITLE, font=("Yu Gothic UI", 20, "bold")).pack(side=LEFT)
        ttk.Label(header, text=APP_SUBTITLE, font=("Yu Gothic UI", 12)).pack(side=LEFT, padx=(12, 0))
        ttk.Button(header, text="不足一覧", command=self.show_shortages).pack(side=RIGHT)

        body = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        body.pack(fill=BOTH, expand=True)

        left = ttk.Frame(body, padding=10)
        center = ttk.Frame(body, padding=10)
        right = ttk.Frame(body, padding=10)
        body.add(left, weight=2)
        body.add(center, weight=3)
        body.add(right, weight=3)

        self.build_project_panel(left)
        self.build_company_panel(center)
        self.build_detail_panel(right)

        bottom = ttk.Notebook(self.root)
        bottom.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
        docs_tab = ttk.Frame(bottom, padding=10)
        attach_tab = ttk.Frame(bottom, padding=10)
        bottom.add(docs_tab, text="添付書類チェック")
        bottom.add(attach_tab, text="ファイル添付")
        self.build_document_panel(docs_tab)
        self.build_attachment_panel(attach_tab)

        footer = ttk.Frame(self.root, padding=(10, 4))
        footer.pack(fill=X)
        self.status_var = StringVar(value=f"DB: {DB_PATH.name}")
        ttk.Label(footer, textvariable=self.status_var).pack(side=LEFT)
        ttk.Button(footer, text="開発用初期化", command=self.reset_debug_data).pack(side=RIGHT)

    def build_project_panel(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill=X)
        ttk.Label(top, text="工事一覧", font=("Yu Gothic UI", 12, "bold")).pack(side=LEFT)

        ttk.Label(parent, text="工事を選択して、施工体制と添付書類を管理します。").pack(anchor="w", pady=(8, 0))

        primary_buttons = ttk.Frame(parent)
        primary_buttons.pack(fill=X, pady=(8, 0))
        ttk.Button(primary_buttons, text="+ 工事登録", command=self.add_project).pack(side=LEFT)
        ttk.Button(primary_buttons, text="工事編集", command=self.edit_selected_project).pack(side=LEFT, padx=(6, 0))

        ledger_buttons = ttk.Frame(parent)
        ledger_buttons.pack(fill=X, pady=(6, 0))
        ttk.Button(ledger_buttons, text="台帳Excel取込", command=self.import_ledger_excel).pack(side=LEFT)
        ttk.Button(ledger_buttons, text="台帳反映コピー作成", command=self.reflect_project_to_ledger).pack(side=LEFT, padx=(6, 0))

        support_buttons = ttk.Frame(parent)
        support_buttons.pack(fill=X, pady=(6, 0))
        ttk.Button(support_buttons, text="工事再読込", command=self.refresh_projects).pack(side=LEFT)
        ttk.Button(support_buttons, text="CSV取込", command=self.import_projects_csv).pack(side=LEFT, padx=(6, 0))
        ttk.Button(support_buttons, text="工事削除", command=self.delete_selected_project).pack(side=LEFT, padx=(6, 0))

        self.project_list = tk.Listbox(parent, height=18)
        self.project_list.pack(fill=BOTH, expand=True, pady=(8, 8))
        self.project_list.bind("<<ListboxSelect>>", self.on_project_selected)

    def build_company_panel(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill=X)
        ttk.Label(top, text="施工体制ツリー", font=("Yu Gothic UI", 12, "bold")).pack(side=LEFT)
        ttk.Button(top, text="+ 元請", command=self.add_root_company).pack(side=RIGHT)
        ttk.Button(top, text="+ 子業者", command=self.add_child_company).pack(side=RIGHT, padx=(0, 6))
        ttk.Label(parent, text="工事を選択し、元請から順に施工体制を登録します。").pack(anchor="w", pady=(8, 0))

        columns = ("level", "work_type")
        self.company_tree = ttk.Treeview(parent, columns=columns, show="tree headings", height=22)
        self.company_tree.heading("#0", text="会社名")
        self.company_tree.heading("level", text="階層")
        self.company_tree.heading("work_type", text="担当工種")
        self.company_tree.column("#0", width=260)
        self.company_tree.column("level", width=100, anchor="center")
        self.company_tree.column("work_type", width=140)
        self.company_tree.pack(fill=BOTH, expand=True, pady=(6, 0))
        self.company_tree.bind("<<TreeviewSelect>>", self.on_company_selected)

    def build_detail_panel(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="選択業者の詳細", font=("Yu Gothic UI", 12, "bold")).pack(anchor="w")
        ttk.Label(parent, text="施工体制ツリーで業者を選択すると詳細を確認できます。").pack(anchor="w", pady=(8, 0))
        self.detail_text = tk.Text(parent, width=42, height=26)
        self.detail_text.pack(fill=BOTH, expand=True, pady=(6, 8))
        self.detail_text.configure(state="disabled")

        button_row = ttk.Frame(parent)
        button_row.pack(fill=X)
        ttk.Button(button_row, text="書類状態更新", command=self.edit_selected_document).pack(side=LEFT)
        ttk.Button(button_row, text="リネーム候補", command=self.show_rename_candidate).pack(side=LEFT, padx=(8, 0))

    def build_document_panel(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill=X)
        ttk.Label(top, text="必要添付書類", font=("Yu Gothic UI", 11, "bold")).pack(side=LEFT)
        ttk.Button(top, text="状態更新", command=self.edit_selected_document).pack(side=RIGHT)
        ttk.Label(parent, text="不足・未確認・期限切れを優先して確認してください。行のダブルクリックでも状態更新できます。").pack(anchor="w", pady=(8, 0))

        columns = ("document_name", "required", "status", "expiry_date", "attachment_count", "note")
        self.document_tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        self.document_tree.heading("document_name", text="書類名")
        self.document_tree.heading("required", text="必須")
        self.document_tree.heading("status", text="状態")
        self.document_tree.heading("expiry_date", text="有効期限")
        self.document_tree.heading("attachment_count", text="添付数")
        self.document_tree.heading("note", text="備考")
        self.document_tree.column("document_name", width=240)
        self.document_tree.column("required", width=60, anchor="center")
        self.document_tree.column("status", width=100, anchor="center")
        self.document_tree.column("expiry_date", width=120, anchor="center")
        self.document_tree.column("attachment_count", width=100, anchor="center")
        self.document_tree.column("note", width=380)
        self.document_tree.pack(fill=BOTH, expand=True, pady=(6, 0))
        self.document_tree.bind("<<TreeviewSelect>>", self.on_document_selected)
        self.document_tree.bind("<Double-1>", lambda _event: self.edit_selected_document())

        for status, color in STATUS_COLORS.items():
            self.document_tree.tag_configure(status, background=color, foreground="white")

    def build_attachment_panel(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill=X)
        ttk.Label(top, text="添付ファイル", font=("Yu Gothic UI", 11, "bold")).pack(side=LEFT)
        ttk.Button(top, text="+ 添付", command=self.attach_file).pack(side=RIGHT)
        ttk.Button(top, text="保存場所を開く", command=self.open_attachment_folder).pack(side=RIGHT, padx=(0, 6))
        ttk.Button(top, text="ファイルを開く", command=self.open_attachment).pack(side=RIGHT, padx=(0, 6))

        columns = ("original_filename", "stored_filename", "file_type", "created_at")
        self.attachment_tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        self.attachment_tree.heading("original_filename", text="元ファイル名")
        self.attachment_tree.heading("stored_filename", text="保存ファイル名")
        self.attachment_tree.heading("file_type", text="種類")
        self.attachment_tree.heading("created_at", text="登録日時")
        self.attachment_tree.column("original_filename", width=280)
        self.attachment_tree.column("stored_filename", width=280)
        self.attachment_tree.column("file_type", width=80, anchor="center")
        self.attachment_tree.column("created_at", width=150, anchor="center")
        self.attachment_tree.pack(fill=BOTH, expand=True, pady=(8, 0))
        self.attachment_tree.bind("<<TreeviewSelect>>", self.on_attachment_selected)

    def refresh_projects(self) -> None:
        self.project_list.delete(0, END)
        self.project_map.clear()
        selected_index: int | None = None
        for row in self.db.list_projects():
            term = f"{row['start_date'] or '-'} - {row['end_date'] or '-'}"
            line = f"{row['name']} / {row['construction_no']} / {row['client_name']} / {term}"
            self.project_map[line] = row["id"]
            current_index = self.project_list.size()
            self.project_list.insert(END, line)
            if row["id"] == self.selection.project_id:
                selected_index = current_index
        if selected_index is not None:
            self.project_list.selection_set(selected_index)
            self.project_list.activate(selected_index)
            self.on_project_selected()
        elif self.project_list.size() > 0 and self.selection.project_id is None:
            self.project_list.selection_set(0)
            self.on_project_selected()
        elif self.project_list.size() == 0:
            self.project_list.insert(END, "工事が登録されていません。「+ 工事登録」または「台帳Excel取込」から開始してください。")
            self.clear_project_selection()

    def on_project_selected(self, _event: object | None = None) -> None:
        selection = self.project_list.curselection()
        if not selection:
            self.clear_project_selection()
            return
        label = self.project_list.get(selection[0])
        if label not in self.project_map:
            self.clear_project_selection()
            return
        self.selection.project_id = self.project_map[label]
        self.selection.company_id = None
        self.selection.document_id = None
        self.selection.attachment_id = None
        self.refresh_company_tree()
        self.refresh_document_tree()
        self.refresh_attachment_tree()
        self.update_detail(None)

    def clear_project_selection(self) -> None:
        self.selection.project_id = None
        self.selection.company_id = None
        self.selection.document_id = None
        self.selection.attachment_id = None
        self.refresh_company_tree()
        self.refresh_document_tree()
        self.refresh_attachment_tree()
        self.update_detail(None)

    def refresh_company_tree(self) -> None:
        for item in self.company_tree.get_children():
            self.company_tree.delete(item)
        self.company_tree_map.clear()
        selected_tree_id: str | None = None

        project_id = self.selection.project_id
        if project_id is None:
            self.company_tree.insert("", END, text="左の工事を選択してください", values=("", ""))
            return

        companies = self.db.list_companies(project_id)
        if not companies:
            self.company_tree.insert("", END, text="+ 元請 から施工体制を登録してください", values=("", ""))
            return

        by_parent: dict[int | None, list[sqlite3.Row]] = {}
        for row in companies:
            by_parent.setdefault(row["parent_company_id"], []).append(row)

        def insert_nodes(parent_tree_id: str, parent_company_id: int | None) -> None:
            for row in by_parent.get(parent_company_id, []):
                tree_id = self.company_tree.insert(
                    parent_tree_id,
                    END,
                    text=row["name"],
                    values=(LEVEL_LABELS.get(row["level"], f"Lv{row['level']}"), row["work_type"] or ""),
                    open=True,
                )
                self.company_tree_map[tree_id] = row["id"]
                nonlocal selected_tree_id
                if row["id"] == self.selection.company_id:
                    selected_tree_id = tree_id
                insert_nodes(tree_id, row["id"])

        insert_nodes("", None)
        if selected_tree_id is not None:
            self.company_tree.selection_set(selected_tree_id)
            self.company_tree.focus(selected_tree_id)
            self.on_company_selected()

    def on_company_selected(self, _event: object | None = None) -> None:
        selected = self.company_tree.selection()
        if not selected:
            return
        tree_id = selected[0]
        if tree_id not in self.company_tree_map:
            return
        company_id = self.company_tree_map[tree_id]
        self.selection.company_id = company_id
        self.selection.document_id = None
        self.selection.attachment_id = None
        company = self.db.get_company(company_id)
        self.update_detail(company)
        self.refresh_document_tree()
        self.refresh_attachment_tree()

    def refresh_document_tree(self) -> None:
        for item in self.document_tree.get_children():
            self.document_tree.delete(item)
        self.document_tree_map.clear()

        company_id = self.selection.company_id
        if company_id is None:
            self.document_tree.insert("", END, values=("業者を選択すると必要添付書類を表示します", "", "", "", "", ""))
            return

        for row in self.db.list_required_documents(company_id):
            attachment_count = int(row["attachment_count"])
            item_id = self.document_tree.insert(
                "",
                END,
                values=(
                    row["document_name"],
                    "○" if row["required"] else "-",
                    row["status"],
                    row["expiry_date"] or "",
                    "0 (未添付)" if attachment_count == 0 else str(attachment_count),
                    row["note"] or "",
                ),
                tags=(row["status"],),
            )
            self.document_tree_map[item_id] = row["id"]

    def on_document_selected(self, _event: object | None = None) -> None:
        selected = self.document_tree.selection()
        if not selected:
            return
        if selected[0] not in self.document_tree_map:
            return
        self.selection.document_id = self.document_tree_map[selected[0]]
        self.selection.attachment_id = None
        self.refresh_attachment_tree()

    def refresh_attachment_tree(self) -> None:
        for item in self.attachment_tree.get_children():
            self.attachment_tree.delete(item)
        self.attachment_tree_map.clear()

        document_id = self.selection.document_id
        if document_id is None:
            self.attachment_tree.insert("", END, values=("書類を選択すると添付ファイルを表示します", "", "", ""))
            return
        for row in self.db.list_attachments(document_id):
            item_id = self.attachment_tree.insert(
                "",
                END,
                values=(row["original_filename"], row["stored_filename"], row["file_type"], row["created_at"]),
            )
            self.attachment_tree_map[item_id] = row["id"]

    def on_attachment_selected(self, _event: object | None = None) -> None:
        selected = self.attachment_tree.selection()
        if not selected:
            return
        if selected[0] not in self.attachment_tree_map:
            return
        self.selection.attachment_id = self.attachment_tree_map[selected[0]]

    def update_detail(self, company: sqlite3.Row | None) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", END)
        if company is None:
            self.detail_text.insert("1.0", "業者を選択すると詳細を表示します。\n\n工事を選択後、施工体制ツリーから対象業者を選択してください。")
        else:
            lines = [
                "基本情報",
                "------------------------------",
                f"会社名: {company['name']}",
                f"階層: {LEVEL_LABELS.get(company['level'], company['level'])}",
                f"会社名カナ: {company['kana'] or ''}",
                f"代表者名: {company['representative'] or ''}",
                f"所在地: {company['address'] or ''}",
                f"電話番号: {company['phone'] or ''}",
                "",
                "許可・担当",
                "------------------------------",
                f"建設業許可番号: {company['license_no'] or ''}",
                f"許可有効期限: {company['license_expiry'] or ''}",
                f"担当工種: {company['work_type'] or ''}",
                "",
                "契約・施工予定",
                "------------------------------",
                f"契約日: {company['contract_date'] or ''}",
                f"施工開始予定日: {company['planned_start_date'] or ''}",
                f"施工終了予定日: {company['planned_end_date'] or ''}",
                "",
                "技術者・安全衛生",
                "------------------------------",
                f"主任技術者名: {company['chief_engineer_name'] or ''}",
                f"主任技術者資格: {company['chief_engineer_license'] or ''}",
                f"安全衛生責任者: {company['safety_manager'] or ''}",
            ]
            self.detail_text.insert("1.0", "\n".join(lines))
        self.detail_text.configure(state="disabled")

    def add_project(self) -> None:
        dialog = ProjectDialog(self.root)
        self.root.wait_window(dialog.top)
        if dialog.result is None:
            return
        project_id = self.db.create_project(dialog.result)
        self.selection.project_id = project_id
        self.refresh_projects()
        self.status_var.set(f"工事を登録しました: {dialog.result['name']}")

    def edit_selected_project(self) -> None:
        project_id = self.selection.project_id
        if project_id is None:
            messagebox.showinfo("未選択", "編集する工事を選択してください。")
            return
        project = self.db.get_project(project_id)
        if project is None:
            return
        initial_data = {
            "name": project["name"] or "",
            "construction_no": project["construction_no"] or "",
            "client_name": project["client_name"] or "",
            "start_date": project["start_date"] or "",
            "end_date": project["end_date"] or "",
            "site_agent": project["site_agent"] or "",
            "managing_engineer": project["managing_engineer"] or "",
            "chief_engineer": project["chief_engineer"] or "",
            "base_folder": project["base_folder"] or "",
        }
        dialog = ProjectDialog(self.root, title="工事編集", initial_data=initial_data)
        self.root.wait_window(dialog.top)
        if dialog.result is None:
            return
        self.db.update_project(project_id, dialog.result)
        self.refresh_projects()
        self.status_var.set(f"工事を更新しました: {dialog.result['name']}")

    def delete_selected_project(self) -> None:
        project_id = self.selection.project_id
        if project_id is None:
            messagebox.showinfo("未選択", "削除する工事を選択してください。")
            return
        project = self.db.get_project(project_id)
        if project is None:
            return
        confirmed = messagebox.askyesno(
            "工事削除",
            f"工事「{project['name']}」を削除します。\n関連する業者、書類、添付情報も削除されます。よろしいですか。",
        )
        if not confirmed:
            return
        self.db.delete_project(project_id)
        self.selection.project_id = None
        self.refresh_projects()
        self.status_var.set(f"工事を削除しました: {project['name']}")

    def reset_debug_data(self) -> None:
        confirmed = messagebox.askyesno(
            "開発用初期化",
            "開発用初期化を実行します。\n登録済みの工事、業者、書類、添付情報、取込履歴をすべて削除します。\n通常運用では使用しない操作です。続行しますか。",
        )
        if not confirmed:
            return
        final_confirmed = messagebox.askyesno(
            "最終確認",
            "本当に全データを初期化しますか。\nこの操作は取り消せません。",
        )
        if not final_confirmed:
            return
        self.db.reset_all_data()
        if ATTACHMENTS_ROOT.exists():
            shutil.rmtree(ATTACHMENTS_ROOT, ignore_errors=True)
        ATTACHMENTS_ROOT.mkdir(parents=True, exist_ok=True)
        self.clear_project_selection()
        self.refresh_projects()
        self.status_var.set("開発用初期化を実行しました。")

    def import_projects_csv(self) -> None:
        csv_path = filedialog.askopenfilename(
            title="工事情報CSVを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if not csv_path:
            return

        records = self.load_project_csv(Path(csv_path))
        if not records:
            messagebox.showinfo("CSV取込", "取り込める工事データがありませんでした。")
            return

        preview_rows: list[dict[str, str]] = []
        for record in records:
            existing = self.db.find_project_by_construction_no(record["construction_no"])
            preview_rows.append(
                {
                    **record,
                    "action": "更新" if existing else "新規",
                    "project_id": str(existing["id"]) if existing else "",
                }
            )

        dialog = CsvImportConfirmDialog(self.root, preview_rows)
        self.root.wait_window(dialog.top)
        if not dialog.approved:
            return

        imported = 0
        updated = 0
        last_project_id: int | None = None
        for row in preview_rows:
            payload = {
                "name": row["name"],
                "construction_no": row["construction_no"],
                "client_name": row["client_name"],
                "start_date": row["start_date"],
                "end_date": row["end_date"],
                "site_agent": row["site_agent"],
                "managing_engineer": row["managing_engineer"],
                "chief_engineer": row["chief_engineer"],
                "base_folder": row["base_folder"],
            }
            if row["action"] == "更新" and row["project_id"]:
                project_id = int(row["project_id"])
                self.db.update_project(project_id, payload)
                updated += 1
                last_project_id = project_id
            else:
                project_id = self.db.create_project(payload)
                imported += 1
                last_project_id = project_id

        self.selection.project_id = last_project_id
        self.refresh_projects()
        self.status_var.set(f"CSV取込完了: 新規 {imported} 件 / 更新 {updated} 件")

    def import_ledger_excel(self) -> None:
        xlsx_path = filedialog.askopenfilename(
            title="施工体制台帳Excelを選択",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
        )
        if not xlsx_path:
            return

        record = self.load_ledger_project_info(Path(xlsx_path))
        if record is None:
            messagebox.showerror(
                "台帳Excel取込",
                "施工体制台帳Excelから工事情報を読み取れませんでした。\n先頭シートに工事名称、発注者、工期などの項目があるか確認してください。",
            )
            return

        existing = self.db.find_project_by_name(record["name"])
        if existing:
            payload = {
                "name": record["name"],
                "construction_no": existing["construction_no"] or provisional_construction_no(record["name"]),
                "client_name": record["client_name"],
                "start_date": record["start_date"],
                "end_date": record["end_date"],
                "site_agent": record["site_agent"],
                "managing_engineer": record["managing_engineer"],
                "chief_engineer": existing["chief_engineer"] or "",
                "base_folder": existing["base_folder"] or "",
            }
            self.db.update_project(existing["id"], payload)
            project_id = existing["id"]
            action = "更新"
        else:
            payload = {
                "name": record["name"],
                "construction_no": provisional_construction_no(record["name"]),
                "client_name": record["client_name"],
                "start_date": record["start_date"],
                "end_date": record["end_date"],
                "site_agent": record["site_agent"],
                "managing_engineer": record["managing_engineer"],
                "chief_engineer": "",
                "base_folder": "",
            }
            project_id = self.db.create_project(payload)
            action = "新規"

        self.db.upsert_project_import(project_id, "施工体制台帳", str(record["source_path"]), record["source_sheet"])
        self.selection.project_id = project_id
        self.refresh_projects()
        self.status_var.set(f"台帳Excel取込完了: {action} / {record['name']}")

    def reflect_project_to_ledger(self) -> None:
        project_id = self.selection.project_id
        if project_id is None:
            messagebox.showinfo("未選択", "反映する工事を選択してください。")
            return
        project = self.db.get_project(project_id)
        if project is None:
            return
        import_row = self.db.get_latest_project_import(project_id, "施工体制台帳")
        if import_row is None:
            messagebox.showinfo(
                "台帳未取込",
                "この工事には施工体制台帳Excelの取込履歴がありません。\n先に「台帳Excel取込」を実行してください。",
            )
            return

        source_path = Path(import_row["source_path"])
        if not source_path.exists():
            messagebox.showerror(
                "取込元Excelなし",
                f"台帳反映に使用する取込元Excelが見つかりません。\n移動または削除されていないか確認してください。\n\n{source_path}",
            )
            return

        output_path = build_reflected_copy_path(source_path)
        confirmed = messagebox.askyesno(
            "台帳反映コピー作成",
            f"取込元Excelは変更せず、反映済みコピーを作成します。\n\n出力予定ファイル:\n{output_path}\n\n実行しますか。",
        )
        if not confirmed:
            return
        workbook = load_workbook(source_path)
        sheet_name = import_row["source_sheet"] or workbook.sheetnames[0]
        sheet = workbook[sheet_name]

        self.write_project_to_ledger_sheet(sheet, project)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_path)
        self.status_var.set(f"台帳へ反映しました: {output_path.name}")
        messagebox.showinfo(
            "台帳反映コピー作成",
            f"原本は変更せず、反映済みコピーを保存しました。\n\n出力ファイル:\n{output_path.name}\n\n保存先:\n{output_path.parent}",
        )

    def load_project_csv(self, csv_path: Path) -> list[dict[str, str]]:
        standard_field_map = {
            "工事名": "name",
            "工事番号": "construction_no",
            "発注者": "client_name",
            "工期開始日": "start_date",
            "工期終了日": "end_date",
            "現場代理人": "site_agent",
            "監理技術者": "managing_engineer",
            "主任技術者": "chief_engineer",
            "保存フォルダ": "base_folder",
        }

        rows: list[dict[str, str]] = []
        loaded: list[dict[str, str]] | None = None
        for encoding in ("utf-8-sig", "cp932", "utf-8"):
            try:
                with csv_path.open("r", encoding=encoding, newline="") as handle:
                    reader = csv.DictReader(handle)
                    loaded = list(reader)
                break
            except UnicodeDecodeError:
                continue

        if loaded is None:
            messagebox.showerror("CSV取込", "CSVの文字コードを読み取れませんでした。")
            return rows

        if not loaded:
            return rows

        header_keys = set(loaded[0].keys())
        is_project_config_csv = {"project_name", "excel_output_dir"}.issubset(header_keys)

        if is_project_config_csv:
            for source in loaded:
                project_name = (source.get("project_name", "") or "").strip()
                if not project_name:
                    continue
                output_dir = (source.get("excel_output_dir", "") or "").strip()
                if output_dir:
                    base_folder = str((csv_path.parent / output_dir).resolve()) if not Path(output_dir).is_absolute() else output_dir
                else:
                    base_folder = ""
                rows.append(
                    {
                        "name": project_name,
                        "construction_no": provisional_construction_no(project_name),
                        "client_name": "",
                        "start_date": "",
                        "end_date": "",
                        "site_agent": "",
                        "managing_engineer": "",
                        "chief_engineer": "",
                        "base_folder": base_folder,
                    }
                )
            return rows

        for source in loaded:
            record = {dest: (source.get(src, "") or "").strip() for src, dest in standard_field_map.items()}
            if not record["construction_no"] or not record["name"]:
                continue
            rows.append(record)
        return rows

    def load_ledger_project_info(self, xlsx_path: Path) -> dict[str, str] | None:
        workbook = load_workbook(xlsx_path, data_only=False)
        sheet = workbook[workbook.sheetnames[0]]

        project_name = self.find_sheet_value(sheet, ["工事名称及び工事内容", "事業所名・現場ID"])
        client_name = self.find_sheet_value(sheet, ["発注者名及び住所"])
        start_date = self.find_sheet_value(sheet, ["工期"], offset_rows=0, occurrence=0)
        end_date = self.find_sheet_value(sheet, ["工期"], offset_rows=1, occurrence=0)
        site_agent = self.find_sheet_value(sheet, ["現場代理人名"])
        managing_engineer = self.find_sheet_value(sheet, ["監理技術者名主任技術者名"], value_index=1)

        if not project_name:
            return None

        return {
            "name": project_name,
            "client_name": client_name,
            "start_date": parse_wareki_free_date(start_date.replace("自", "").strip()) if start_date else "",
            "end_date": parse_wareki_free_date(end_date.replace("至", "").strip()) if end_date else "",
            "site_agent": site_agent,
            "managing_engineer": managing_engineer.replace("専任", "").strip() if managing_engineer else "",
            "source_path": str(xlsx_path),
            "source_sheet": sheet.title,
        }

    def write_project_to_ledger_sheet(self, sheet: object, project: sqlite3.Row) -> None:
        name = project["name"] or ""
        client_name = project["client_name"] or ""
        start_date = project["start_date"] or ""
        end_date = project["end_date"] or ""
        site_agent = project["site_agent"] or ""
        managing_engineer = project["managing_engineer"] or ""

        # 左側ブロックの主要な工事情報に限定して反映する
        if name:
            sheet["I6"] = name
            sheet["G21"] = name
        if client_name:
            sheet["G25"] = client_name
        if start_date:
            sheet["H29"] = format_ledger_date("自", start_date)
        if end_date:
            sheet["H30"] = format_ledger_date("至", end_date)
        if site_agent:
            sheet["G55"] = site_agent
        if managing_engineer:
            sheet["J57"] = managing_engineer

    def find_sheet_value(
        self,
        sheet: object,
        candidate_labels: list[str],
        offset_rows: int = 0,
        occurrence: int = 0,
        value_index: int = 0,
    ) -> str:
        normalized_labels = [normalize_label(label) for label in candidate_labels]
        matches: list[tuple[int, int]] = []
        for row in sheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                normalized = normalize_label(cell.value)
                if any(label in normalized for label in normalized_labels):
                    matches.append((cell.row, cell.column))

        if not matches or occurrence >= len(matches):
            return ""

        row_index, col_index = matches[occurrence]
        target_row = row_index + offset_rows
        found_values = 0
        for next_col in range(col_index + 1, sheet.max_column + 1):
            value = sheet.cell(target_row, next_col).value
            if value not in (None, ""):
                if found_values == value_index:
                    return str(value).strip()
                found_values += 1
        return ""

    def add_root_company(self) -> None:
        project_id = self.selection.project_id
        if project_id is None:
            messagebox.showinfo("未選択", "先に工事を選択してください。")
            return
        dialog = CompanyDialog(self.root, level=0, parent_name=None)
        self.root.wait_window(dialog.top)
        if dialog.result is None:
            return
        payload: dict[str, str | int | None] = dict(dialog.result)
        payload.update({"project_id": project_id, "parent_company_id": None, "level": 0})
        company_id = self.db.create_company(payload)
        self.selection.company_id = company_id
        self.refresh_company_tree()
        self.status_var.set(f"元請を登録しました: {dialog.result['name']}")

    def add_child_company(self) -> None:
        project_id = self.selection.project_id
        parent_company_id = self.selection.company_id
        if project_id is None or parent_company_id is None:
            messagebox.showinfo("未選択", "先に工事と親業者を選択してください。")
            return
        parent = self.db.get_company(parent_company_id)
        if parent is None:
            return
        level = int(parent["level"]) + 1
        dialog = CompanyDialog(self.root, level=level, parent_name=parent["name"])
        self.root.wait_window(dialog.top)
        if dialog.result is None:
            return
        payload: dict[str, str | int | None] = dict(dialog.result)
        payload.update({"project_id": project_id, "parent_company_id": parent_company_id, "level": level})
        company_id = self.db.create_company(payload)
        self.selection.company_id = company_id
        self.refresh_company_tree()
        self.status_var.set(f"子業者を登録しました: {dialog.result['name']}")

    def edit_selected_document(self) -> None:
        document_id = self.selection.document_id
        if document_id is None:
            messagebox.showinfo("未選択", "更新する書類を選択してください。")
            return
        document = self.db.fetchone("SELECT * FROM required_documents WHERE id = ?", (document_id,))
        if document is None:
            return
        dialog = DocumentStatusDialog(self.root, document)
        self.root.wait_window(dialog.top)
        if dialog.result is None:
            return
        self.db.update_required_document(
            document_id,
            dialog.result["status"],
            dialog.result["expiry_date"],
            dialog.result["note"],
        )
        self.refresh_document_tree()
        self.status_var.set(f"書類状態を更新しました: {document['document_name']}")

    def show_shortages(self) -> None:
        project_id = self.selection.project_id
        if project_id is None:
            messagebox.showinfo("未選択", "工事を選択してください。")
            return
        rows = self.db.list_shortages(project_id)
        if not rows:
            messagebox.showinfo("不足一覧", "不足・未確認・期限切れの書類はありません。")
            return
        lines = [f"{index}. [{row['status']}] {row['company_name']} / {row['document_name']}" for index, row in enumerate(rows, start=1)]
        messagebox.showinfo("不足一覧", f"不足・未確認・期限切れ: {len(rows)}件\n\n" + "\n".join(lines))

    def attach_file(self) -> None:
        document_id = self.selection.document_id
        company_id = self.selection.company_id
        project_id = self.selection.project_id
        missing = []
        if project_id is None:
            missing.append("工事")
        if company_id is None:
            missing.append("業者")
        if document_id is None:
            missing.append("必要添付書類")
        if missing:
            messagebox.showinfo("未選択", "添付前に次を選択してください。\n" + "、".join(missing))
            return

        path_str = filedialog.askopenfilename(
            title="添付ファイルを選択",
            filetypes=[
                ("許可ファイル", "*.pdf *.xlsx *.xls *.docx *.doc *.png *.jpg *.jpeg *.bmp"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if not path_str:
            return

        source = Path(path_str)
        project = self.db.get_project(project_id)
        company = self.db.get_company(company_id)
        document = self.db.fetchone("SELECT * FROM required_documents WHERE id = ?", (document_id,))
        if project is None or company is None or document is None:
            return

        project_folder = Path(project["base_folder"]).expanduser() if project["base_folder"] else ATTACHMENTS_ROOT / f"project_{project_id}"
        target_dir = project_folder / "CIVIL_HUB_添付" / sanitize_filename(company["name"]) / sanitize_filename(document["document_name"])
        target_dir.mkdir(parents=True, exist_ok=True)

        target_name = self.db.next_versioned_filename(target_dir, company["name"], document["document_name"], source.suffix.lower())
        target = target_dir / target_name
        shutil.copy2(source, target)
        self.db.add_attachment(document_id, source, target)
        self.db.update_required_document(document_id, "添付済み", document["expiry_date"] or "", document["note"] or "")
        self.refresh_document_tree()
        self.refresh_attachment_tree()
        self.status_var.set(f"ファイルを添付しました: {target.name}")

    def open_attachment(self) -> None:
        attachment_id = self.selection.attachment_id
        if attachment_id is None:
            messagebox.showinfo("未選択", "開く添付ファイルを選択してください。")
            return
        attachment = self.db.get_attachment(attachment_id)
        if attachment is None:
            return
        target = Path(attachment["stored_path"])
        if not target.exists():
            messagebox.showerror("ファイルなし", f"保存ファイルが見つかりません。\n{target}")
            return
        open_path(target)

    def open_attachment_folder(self) -> None:
        attachment_id = self.selection.attachment_id
        if attachment_id is not None:
            attachment = self.db.get_attachment(attachment_id)
            if attachment is not None:
                folder = Path(attachment["stored_path"]).parent
                folder.mkdir(parents=True, exist_ok=True)
                open_path(folder)
                return

        document_id = self.selection.document_id
        company_id = self.selection.company_id
        project_id = self.selection.project_id
        if document_id is None or company_id is None or project_id is None:
            messagebox.showinfo("未選択", "保存場所を開くには、先に工事・業者・必要添付書類を選択してください。")
            return

        project = self.db.get_project(project_id)
        company = self.db.get_company(company_id)
        document = self.db.fetchone("SELECT * FROM required_documents WHERE id = ?", (document_id,))
        if project is None or company is None or document is None:
            return
        project_folder = Path(project["base_folder"]).expanduser() if project["base_folder"] else ATTACHMENTS_ROOT / f"project_{project_id}"
        folder = project_folder / "CIVIL_HUB_添付" / sanitize_filename(company["name"]) / sanitize_filename(document["document_name"])
        folder.mkdir(parents=True, exist_ok=True)
        open_path(folder)

    def show_rename_candidate(self) -> None:
        document_id = self.selection.document_id
        company_id = self.selection.company_id
        project_id = self.selection.project_id
        if document_id is None or company_id is None or project_id is None:
            messagebox.showinfo("未選択", "先に工事・業者・書類を選択してください。")
            return
        project = self.db.get_project(project_id)
        company = self.db.get_company(company_id)
        document = self.db.fetchone("SELECT * FROM required_documents WHERE id = ?", (document_id,))
        if project is None or company is None or document is None:
            return

        suffix = simpledialog.askstring("拡張子", "拡張子を入力してください。例: .pdf", initialvalue=".pdf")
        if not suffix:
            return
        if not suffix.startswith("."):
            suffix = f".{suffix}"

        project_folder = Path(project["base_folder"]).expanduser() if project["base_folder"] else ATTACHMENTS_ROOT / f"project_{project_id}"
        target_dir = project_folder / "CIVIL_HUB_添付" / sanitize_filename(company["name"]) / sanitize_filename(document["document_name"])
        target_dir.mkdir(parents=True, exist_ok=True)
        candidate = self.db.next_versioned_filename(target_dir, company["name"], document["document_name"], suffix.lower())
        messagebox.showinfo("リネーム候補", candidate)


def main() -> None:
    ATTACHMENTS_ROOT.mkdir(parents=True, exist_ok=True)
    root = Tk()
    app = CivilHubApp(root)
    root.minsize(1200, 760)
    root.mainloop()


if __name__ == "__main__":
    main()
