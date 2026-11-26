import os
import re
import sqlite3
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import openpyxl
from openpyxl.styles import Font, PatternFill


def get_app_data_dir() -> Path:
    """Return a guaranteed writable local directory for data and exports."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        test_file = exe_dir / ".write_test"
        try:
            test_file.touch()
            test_file.unlink()
            return exe_dir
        except OSError:
            local_app = os.environ.get("LOCALAPPDATA") or str(Path.home())
            fallback = Path(local_app) / "BAF_Simple_Student_Register"
            fallback.mkdir(parents=True, exist_ok=True)
            return fallback
    return Path(__file__).resolve().parent


def reg_key(value: str | None) -> str:
    """Normalize registration numbers for duplicate checking."""
    return "".join((value or "").split()).upper()


def sanitize_filename(name: str) -> str:
    """Replace invalid Windows filename characters with underscores."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name.strip())
    cleaned = cleaned.rstrip(". ")
    return cleaned or "Class_Export"


class DatabaseManager:
    """Offline SQLite manager with foreign keys and integrity handling."""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON;")
        self._init_schema()

    def _init_schema(self) -> None:
        with self.conn:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS classes (
                  id INTEGER PRIMARY KEY,
                  name TEXT NOT NULL UNIQUE COLLATE NOCASE
                );
                """
            )
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS students (
                  id INTEGER PRIMARY KEY,
                  class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
                  name TEXT,
                  reg TEXT,
                  reg_key TEXT,
                  UNIQUE(class_id, reg_key)
                );
                """
            )

    def get_classes(self) -> list[sqlite3.Row]:
        cur = self.conn.execute("SELECT id, name FROM classes ORDER BY name COLLATE NOCASE ASC;")
        return cur.fetchall()

    def create_class(self, name: str) -> int:
        with self.conn:
            cur = self.conn.execute("INSERT INTO classes (name) VALUES (?);", (name.strip(),))
            return int(cur.lastrowid)

    def rename_class(self, class_id: int, new_name: str) -> None:
        with self.conn:
            self.conn.execute(
                "UPDATE classes SET name = ? WHERE id = ?;",
                (new_name.strip(), class_id),
            )

    def delete_class(self, class_id: int) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM classes WHERE id = ?;", (class_id,))

    def get_students(self, class_id: int) -> list[sqlite3.Row]:
        cur = self.conn.execute(
            "SELECT id, name, reg, reg_key FROM students WHERE class_id = ? ORDER BY id ASC;",
            (class_id,),
        )
        return cur.fetchall()

    def reg_exists(
        self, class_id: int, normalized_key: str, exclude_student_id: int | None = None
    ) -> bool:
        if not normalized_key:
            return False
        if exclude_student_id is not None:
            cur = self.conn.execute(
                "SELECT 1 FROM students WHERE class_id = ? AND reg_key = ? AND id != ? LIMIT 1;",
                (class_id, normalized_key, exclude_student_id),
            )
        else:
            cur = self.conn.execute(
                "SELECT 1 FROM students WHERE class_id = ? AND reg_key = ? LIMIT 1;",
                (class_id, normalized_key),
            )
        return cur.fetchone() is not None

    def insert_student(self, class_id: int, name: str, reg: str) -> int:
        norm = reg_key(reg) or None
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO students (class_id, name, reg, reg_key) VALUES (?, ?, ?, ?);",
                (class_id, name, reg, norm),
            )
            return int(cur.lastrowid)

    def update_student(self, student_id: int, name: str, reg: str) -> None:
        norm = reg_key(reg) or None
        with self.conn:
            self.conn.execute(
                "UPDATE students SET name = ?, reg = ?, reg_key = ? WHERE id = ?;",
                (name, reg, norm, student_id),
            )

    def delete_student(self, student_id: int) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM students WHERE id = ?;", (student_id,))

    def close(self) -> None:
        if self.conn:
            self.conn.close()


class StudentRegisterApp(tk.Tk):
    INITIAL_MIN_ROWS = 20
    DEFAULT_STATUS = (
        "Type a name, press Enter, type the registration number, "
        "then press Enter for the next student."
    )

    BG_MAIN = "#0b1220"
    BG_PANEL = "#101e30"
    BG_BTN = "#163449"
    BG_BTN_HOVER = "#20516a"
    BG_HEADER = "#193449"
    BORDER = "#28415b"
    BORDER_BTN = "#247b91"
    TEXT_MAIN = "#dfebff"
    TEXT_MUTED = "#9ec0e6"
    ACCENT = "#63e6df"

    def __init__(self, db: DatabaseManager, base_dir: Path):
        super().__init__()
        self.db = db
        self.base_dir = base_dir

        self.classes_data: list[tuple[int, str]] = []
        self.current_class_id: int | None = None
        self.current_class_name: str = ""

        self.rows_data: list[dict] = []
        self.iid_to_row_idx: dict[str, int] = {}

        self.active_editor: tk.Entry | None = None
        self.editing_row_idx: int | None = None
        self.editing_col_idx: int | None = None

        self.title("BAF | Simple Student Register")
        self.configure(bg=self.BG_MAIN)
        self._center_window(900, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._configure_ttk_styles()
        self._build_ui()
        self._load_classes()

    def _center_window(self, width: int, height: int) -> None:
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(760, 540)

    def _configure_ttk_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "Register.Treeview",
            background=self.BG_PANEL,
            foreground=self.TEXT_MAIN,
            fieldbackground=self.BG_PANEL,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            rowheight=30,
            font=("Segoe UI", 11),
        )
        style.map(
            "Register.Treeview",
            background=[("selected", self.BG_BTN_HOVER)],
            foreground=[("selected", "#ffffff")],
        )
        style.configure(
            "Register.Treeview.Heading",
            background=self.BG_HEADER,
            foreground=self.ACCENT,
            bordercolor=self.BORDER,
            relief="flat",
            font=("Segoe UI", 11, "bold"),
            padding=(8, 8),
        )
        style.map(
            "Register.Treeview.Heading",
            background=[("active", self.BG_BTN_HOVER)],
        )

    def _make_sidebar_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=self.BG_BTN,
            fg=self.TEXT_MAIN,
            activebackground=self.BG_BTN_HOVER,
            activeforeground=self.ACCENT,
            disabledforeground="#5b7694",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
            padx=12,
            pady=8,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=self.BORDER_BTN,
            highlightcolor=self.ACCENT,
            cursor="hand2",
        )
        btn.bind("<Enter>", lambda _e: btn.config(bg=self.BG_BTN_HOVER) if btn["state"] == "normal" else None)
        btn.bind("<Leave>", lambda _e: btn.config(bg=self.BG_BTN) if btn["state"] == "normal" else None)
        return btn

    def _build_ui(self) -> None:
        root_frame = tk.Frame(self, bg=self.BG_MAIN)
        root_frame.pack(fill="both", expand=True, padx=16, pady=16)

        sidebar = tk.Frame(root_frame, bg=self.BG_MAIN, width=230)
        sidebar.pack(side="left", fill="y", padx=(0, 16))
        sidebar.pack_propagate(False)

        brand_label = tk.Label(
            sidebar,
            text="BAF\nREGISTER",
            bg=self.BG_MAIN,
            fg=self.ACCENT,
            font=("Segoe UI", 22, "bold"),
            justify="left",
            anchor="w",
        )
        brand_label.pack(fill="x", pady=(0, 12))

        list_frame = tk.Frame(
            sidebar,
            bg=self.BG_PANEL,
            highlightthickness=1,
            highlightbackground=self.BORDER,
        )
        list_frame.pack(fill="x", pady=(0, 12))

        self.class_listbox = tk.Listbox(
            list_frame,
            bg=self.BG_PANEL,
            fg=self.TEXT_MAIN,
            selectbackground=self.BG_BTN_HOVER,
            selectforeground=self.ACCENT,
            font=("Segoe UI", 11),
            height=10,
            bd=0,
            highlightthickness=0,
            activestyle="none",
            exportselection=False,
        )
        self.class_listbox.pack(fill="both", expand=True, padx=6, pady=6)
        self.class_listbox.bind("<<ListboxSelect>>", self._on_class_selected)

