import os
import re
import sqlite3
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from PySide6.QtCore import QEvent, QObject, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStyleFactory,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)


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


def create_app_icon() -> QIcon:
    """Generate a crisp teal/navy BAF icon programmatically."""
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    painter.setBrush(QColor("#101e30"))
    painter.setPen(QPen(QColor("#63e6df"), 3))
    painter.drawRoundedRect(4, 4, 56, 56, 12, 12)

    painter.setPen(QColor("#63e6df"))
    font = QFont("Segoe UI", 18, QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "BAF")
    painter.end()
    return QIcon(pixmap)


class EnterKeyFilter(QObject):
    """Intercept Enter/Return keys inside table editors for rapid data entry."""

    def __init__(self, table: QTableWidget, on_enter_callback):
        super().__init__(table)
        self.table = table
        self.on_enter_callback = on_enter_callback

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.Type.KeyPress and event.key() in (
            Qt.Key.Key_Return,
            Qt.Key.Key_Enter,
        ):
            row = self.table.currentRow()
            col = self.table.currentColumn()
            if row >= 0 and col >= 0:
                self.table.setFocus()
                QTimer.singleShot(0, lambda r=row, c=col: self.on_enter_callback(r, c))
                return True
        return super().eventFilter(watched, event)


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

    def reg_exists(self, class_id: int, normalized_key: str, exclude_student_id: int | None = None) -> bool:
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


class StudentRegisterWindow(QMainWindow):
    INITIAL_MIN_ROWS = 20
    DEFAULT_STATUS = (
        "Type a name, press Enter, type the registration number, "
        "then press Enter for the next student."
    )

    def __init__(self, db: DatabaseManager, base_dir: Path):
        super().__init__()
        self.db = db
        self.base_dir = base_dir
        self.current_class_id: int | None = None
        self.current_class_name: str = ""
        self._loading = False

        self.setWindowTitle("BAF | Simple Student Register")
        self.resize(900, 700)
        self.setWindowIcon(create_app_icon())

        self._build_ui()
        self._apply_styles()
        self._load_classes()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(16)

        # Left Sidebar
        sidebar = QWidget()
        sidebar.setFixedWidth(230)
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.setSpacing(10)

        self.brand_label = QLabel("BAF\nREGISTER")
        self.brand_label.setObjectName("brandHeading")
        side_layout.addWidget(self.brand_label)

        self.class_list = QListWidget()
        self.class_list.setObjectName("classList")
        self.class_list.currentItemChanged.connect(self._on_class_selected)
        side_layout.addWidget(self.class_list, stretch=1)

        self.btn_new = QPushButton("+ New class")
        self.btn_rename = QPushButton("Rename")
        self.btn_delete = QPushButton("Delete class")
        self.btn_export = QPushButton("Export this class")

        self.btn_new.clicked.connect(self._handle_new_class)
        self.btn_rename.clicked.connect(self._handle_rename_class)
        self.btn_delete.clicked.connect(self._handle_delete_class)
        self.btn_export.clicked.connect(self._handle_export_class)

        for btn in (self.btn_new, self.btn_rename, self.btn_delete, self.btn_export):
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            side_layout.addWidget(btn)

        side_layout.addStretch(1)
        root_layout.addWidget(sidebar)

        # Main Content Area
        main_area = QWidget()
        main_layout = QVBoxLayout(main_area)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(12)

        self.class_heading = QLabel("Select or create a class")
        self.class_heading.setObjectName("classHeading")
        main_layout.addWidget(self.class_heading)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search name or registration number...")
        self.search_input.textChanged.connect(self._apply_search_filter)
        main_layout.addWidget(self.search_input)

        self.table = QTableWidget(self.INITIAL_MIN_ROWS, 2)
        self.table.setHorizontalHeaderLabels(["STUDENT NAME", "REGISTRATION NUMBER"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
