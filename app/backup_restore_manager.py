import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QMessageBox,
    QFileDialog,
)
from PySide6.QtCore import Qt

from database.schema import init_database


DB_NAME = "pharmaguard_operational_v12.db"
BACKUP_DIR_NAME = "backups"


def project_root():
    return Path.cwd()


def db_path():
    return project_root() / DB_NAME


def backup_dir():
    path = project_root() / BACKUP_DIR_NAME
    path.mkdir(exist_ok=True)
    return path


def make_backup():
    init_database()

    source = db_path()

    if not source.exists():
        raise FileNotFoundError(f"Database file not found: {source}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = backup_dir() / f"pharmaguard_backup_{timestamp}.db"

    shutil.copy2(source, target)

    return target


def list_backups():
    folder = backup_dir()
    return sorted(folder.glob("pharmaguard_backup_*.db"), reverse=True)


def restore_backup(backup_file):
    source = Path(backup_file)

    if not source.exists():
        raise FileNotFoundError(f"Backup file not found: {source}")

    target = db_path()

    if target.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safety_copy = backup_dir() / f"before_restore_safety_copy_{timestamp}.db"
        shutil.copy2(target, safety_copy)

    shutil.copy2(source, target)

    return target


class BackupRestoreManager(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Backup / Restore - PharmaGuard")
        self.resize(900, 600)

        layout = QVBoxLayout(self)

        title = QLabel("Backup / Restore / النسخ الاحتياطي والاسترجاع")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel(
            "This screen creates backups of the operational SQLite database. "
            "Restore should be used carefully. Close other PharmaGuard windows before restore."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self.db_label = QLabel(f"Database: {db_path()}")
        self.db_label.setWordWrap(True)
        layout.addWidget(self.db_label)

        buttons = QHBoxLayout()

        self.backup_btn = QPushButton("Create backup / إنشاء نسخة احتياطية")
        self.backup_btn.clicked.connect(self.create_backup_clicked)

        self.restore_btn = QPushButton("Restore selected backup / استرجاع النسخة المحددة")
        self.restore_btn.clicked.connect(self.restore_selected_clicked)

        self.refresh_btn = QPushButton("Refresh list / تحديث القائمة")
        self.refresh_btn.clicked.connect(self.load_backups)

        self.import_backup_btn = QPushButton("Restore from file / استرجاع من ملف")
        self.import_backup_btn.clicked.connect(self.restore_from_file_clicked)

        buttons.addWidget(self.backup_btn)
        buttons.addWidget(self.restore_btn)
        buttons.addWidget(self.import_backup_btn)
        buttons.addWidget(self.refresh_btn)
        buttons.addStretch()

        layout.addLayout(buttons)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.status_label)

        self.backup_list = QListWidget()
        layout.addWidget(self.backup_list)

        self.load_backups()

    def load_backups(self):
        self.backup_list.clear()

        backups = list_backups()

        for path in backups:
            self.backup_list.addItem(str(path))

        self.status_label.setText(f"Found {len(backups)} backup files.")

    def create_backup_clicked(self):
        try:
            path = make_backup()

            QMessageBox.information(
                self,
                "Backup created",
                f"Backup created successfully:\n{path}"
            )

            self.load_backups()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def selected_backup_path(self):
        item = self.backup_list.currentItem()

        if item is None:
            return None

        return item.text()

    def restore_selected_clicked(self):
        selected = self.selected_backup_path()

        if not selected:
            QMessageBox.warning(self, "No selection", "Select a backup first.")
            return

        reply = QMessageBox.question(
            self,
            "Confirm restore",
            "Restoring will replace the current database.\n\n"
            "A safety copy of the current database will be created first.\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No,
        )

        if reply != QMessageBox.Yes:
            return

        try:
            restored = restore_backup(selected)

            QMessageBox.information(
                self,
                "Restore done",
                f"Database restored successfully:\n{restored}\n\n"
                "Close and reopen other PharmaGuard windows."
            )

            self.load_backups()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    def restore_from_file_clicked(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose backup database file",
            str(backup_dir()),
            "Database Files (*.db);;All Files (*.*)"
        )

        if not path:
            return

        reply = QMessageBox.question(
            self,
            "Confirm restore",
            "Restoring will replace the current database.\n\n"
            "A safety copy of the current database will be created first.\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No,
        )

        if reply != QMessageBox.Yes:
            return

        try:
            restored = restore_backup(path)

            QMessageBox.information(
                self,
                "Restore done",
                f"Database restored successfully:\n{restored}\n\n"
                "Close and reopen other PharmaGuard windows."
            )

            self.load_backups()

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = BackupRestoreManager()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
