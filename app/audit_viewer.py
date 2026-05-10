import sys
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
)
from PySide6.QtCore import Qt

from database.schema import connect


def load_audit_logs(table_filter="All", action_filter="All", search_text="", limit=500):
    conn = connect()
    cur = conn.cursor()

    query = """
    SELECT
        a.audit_id,
        a.created_at,
        COALESCE(u.username, '') AS username,
        a.action,
        a.table_name,
        a.record_id,
        COALESCE(a.old_value, '') AS old_value,
        COALESCE(a.new_value, '') AS new_value
    FROM audit_logs a
    LEFT JOIN users u ON a.user_id = u.user_id
    WHERE 1 = 1
    """

    params = []

    if table_filter != "All":
        query += " AND a.table_name = ?"
        params.append(table_filter)

    if action_filter != "All":
        query += " AND a.action = ?"
        params.append(action_filter)

    if search_text.strip():
        query += """
        AND (
            a.table_name LIKE ?
            OR a.record_id LIKE ?
            OR a.old_value LIKE ?
            OR a.new_value LIKE ?
            OR u.username LIKE ?
        )
        """
        pattern = f"%{search_text.strip()}%"
        params.extend([pattern, pattern, pattern, pattern, pattern])

    query += " ORDER BY a.audit_id DESC LIMIT ?"
    params.append(limit)

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


class AuditViewer(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Audit Trail Viewer - PharmaGuard")
        self.resize(1300, 700)

        layout = QVBoxLayout(self)

        title = QLabel("Audit Trail Viewer / سجل المراجعة")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        note = QLabel("يعرض كل العمليات التي سجلتها قاعدة البيانات تلقائيًا.")
        note.setWordWrap(True)
        layout.addWidget(note)

        filters = QHBoxLayout()

        self.table_filter = QComboBox()
        self.table_filter.addItems([
            "All",
            "stock_movements",
            "items",
            "batches",
            "suppliers",
            "purchase_orders",
        ])

        self.action_filter = QComboBox()
        self.action_filter.addItems([
            "All",
            "INSERT",
            "UPDATE",
            "DELETE",
        ])

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search audit logs...")

        self.refresh_btn = QPushButton("Refresh / تحديث")
        self.refresh_btn.clicked.connect(self.load_table)

        filters.addWidget(QLabel("Table:"))
        filters.addWidget(self.table_filter)
        filters.addWidget(QLabel("Action:"))
        filters.addWidget(self.action_filter)
        filters.addWidget(QLabel("Search:"))
        filters.addWidget(self.search_box)
        filters.addWidget(self.refresh_btn)

        layout.addLayout(filters)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        layout.addWidget(self.table)

        self.load_table()

    def load_table(self):
        try:
            rows = load_audit_logs(
                table_filter=self.table_filter.currentText(),
                action_filter=self.action_filter.currentText(),
                search_text=self.search_box.text(),
                limit=500,
            )

            columns = [
                "audit_id",
                "created_at",
                "username",
                "action",
                "table_name",
                "record_id",
                "old_value",
                "new_value",
            ]

            self.table.setSortingEnabled(False)
            self.table.clear()
            self.table.setColumnCount(len(columns))
            self.table.setRowCount(len(rows))
            self.table.setHorizontalHeaderLabels(columns)

            for row_index, row in enumerate(rows):
                for col_index, value in enumerate(row):
                    item = QTableWidgetItem(str(value or ""))
                    self.table.setItem(row_index, col_index, item)

            self.table.resizeColumnsToContents()
            self.table.setSortingEnabled(True)

            self.status_label.setText(f"Loaded {len(rows)} audit rows.")

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = AuditViewer()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
