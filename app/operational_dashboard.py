import sys
from datetime import date, datetime, timedelta
from app.ui_helpers import create_scroll_layout

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
)
from PySide6.QtCore import Qt

from database.schema import connect, init_database
from database.repositories import list_recent_movements, list_low_stock_items
from database.stock_views import get_current_stock_rows
from app.session import get_current_user_label


def get_count(query):
    conn = connect()
    cur = conn.cursor()
    cur.execute(query)
    row = cur.fetchone()
    conn.close()
    return row[0] if row else 0


def get_today_movements():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            sm.movement_id,
            sm.movement_datetime,
            sm.movement_type,
            i.generic_name,
            sm.quantity,
            COALESCE(lf.location_name, '') AS from_location,
            COALESCE(lt.location_name, '') AS to_location,
            COALESCE(sm.reference_no, '') AS reference_no
        FROM stock_movements sm
        JOIN items i ON sm.item_id = i.item_id
        LEFT JOIN locations lf ON sm.from_location_id = lf.location_id
        LEFT JOIN locations lt ON sm.to_location_id = lt.location_id
        WHERE DATE(sm.movement_datetime) = DATE('now')
        ORDER BY sm.movement_id DESC
        LIMIT 100;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def get_po_status_counts():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT status, COUNT(*)
        FROM purchase_orders
        GROUP BY status
        ORDER BY status;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def parse_date(value):
    if not value:
        return None

    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def get_expiry_rows():
    stock_rows = get_current_stock_rows()
    today = date.today()
    near_limit = today + timedelta(days=90)

    result = []

    for row in stock_rows:
        expiry = parse_date(row.get("expiry_date"))
        stock = float(row.get("current_stock") or 0)

        if expiry is None or stock <= 0:
            continue

        if expiry < today:
            risk = "Expired"
        elif expiry <= near_limit:
            risk = "Near expiry"
        else:
            continue

        result.append([
            risk,
            row.get("generic_name", ""),
            row.get("location_name", ""),
            row.get("batch_number", ""),
            row.get("expiry_date", ""),
            stock,
        ])

    result.sort(key=lambda x: x[4])
    return result[:100]


def total_stock_quantity():
    rows = get_current_stock_rows()
    total = 0

    for row in rows:
        try:
            total += float(row.get("current_stock") or 0)
        except Exception:
            pass

    return round(total, 2)


class OperationalDashboard(QWidget):
    def __init__(self):
        super().__init__()

        init_database()

        self.setWindowTitle("Operational Dashboard - PharmaGuard")
        self.resize(1350, 800)

        layout = create_scroll_layout(self)

        title = QLabel("Operational Dashboard / لوحة التشغيل")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(title)

        self.user_label = QLabel(f"Current user: {get_current_user_label()}")
        self.user_label.setAlignment(Qt.AlignCenter)
        self.user_label.setStyleSheet("font-weight: bold; color: #2563eb;")
        layout.addWidget(self.user_label)

        note = QLabel(
            "هذه الشاشة تقرأ من قاعدة البيانات التشغيلية الجديدة: "
            "الأصناف، الحركات، الباتشات، طلبات الشراء، والصلاحيات."
        )
        note.setWordWrap(True)
        note.setAlignment(Qt.AlignCenter)
        layout.addWidget(note)

        self.refresh_btn = QPushButton("Refresh Dashboard / تحديث")
        self.refresh_btn.clicked.connect(self.load_dashboard)
        layout.addWidget(self.refresh_btn)

        self.cards_grid = QGridLayout()
        layout.addLayout(self.cards_grid)

        self.today_table = QTableWidget()
        self.expiry_table = QTableWidget()
        self.low_stock_table = QTableWidget()
        self.po_table = QTableWidget()
        self.recent_table = QTableWidget()

        layout.addWidget(QLabel("Today Movements / حركات اليوم"))
        layout.addWidget(self.today_table)

        layout.addWidget(QLabel("Near Expiry / Expired Batches / صلاحيات مهمة"))
        layout.addWidget(self.expiry_table)

        layout.addWidget(QLabel("Low Stock / أقل من الحد الأدنى"))
        layout.addWidget(self.low_stock_table)

        layout.addWidget(QLabel("Purchase Order Status / حالات طلبات الشراء"))
        layout.addWidget(self.po_table)

        layout.addWidget(QLabel("Recent Movements / آخر الحركات"))
        layout.addWidget(self.recent_table)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        self.load_dashboard()

    def clear_cards(self):
        while self.cards_grid.count():
            item = self.cards_grid.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def make_card(self, title, value):
        label = QLabel()
        label.setText(
            f"<div style='font-size:13px;color:#475569;'>{title}</div>"
            f"<div style='font-size:30px;font-weight:bold;color:#0f172a;'>{value}</div>"
        )
        label.setAlignment(Qt.AlignCenter)
        label.setMinimumHeight(90)
        label.setStyleSheet("""
            QLabel {
                background-color: #f8fafc;
                border: 1px solid #cbd5e1;
                border-radius: 12px;
                padding: 10px;
            }
        """)
        return label

    def set_table(self, table, columns, rows):
        table.setSortingEnabled(False)
        table.clear()
        table.setColumnCount(len(columns))
        table.setRowCount(len(rows))
        table.setHorizontalHeaderLabels(columns)

        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                table.setItem(r, c, QTableWidgetItem(str(value or "")))

        table.resizeColumnsToContents()
        table.setSortingEnabled(True)

    def load_dashboard(self):
        try:
            self.clear_cards()

            item_count = get_count("SELECT COUNT(*) FROM items WHERE is_active = 1;")
            batch_count = get_count("SELECT COUNT(*) FROM batches;")
            movement_count = get_count("SELECT COUNT(*) FROM stock_movements;")
            today_count = get_count("SELECT COUNT(*) FROM stock_movements WHERE DATE(movement_datetime) = DATE('now');")
            open_po_count = get_count("SELECT COUNT(*) FROM purchase_orders WHERE status NOT IN ('Closed', 'Cancelled');")
            user_count = get_count("SELECT COUNT(*) FROM users WHERE is_active = 1;")
            total_stock = total_stock_quantity()
            expiry_rows = get_expiry_rows()
            low_stock_rows = list_low_stock_items()

            cards = [
                ("Items / الأصناف", item_count),
                ("Batches / الباتشات", batch_count),
                ("Total Stock / إجمالي الرصيد", total_stock),
                ("All Movements / كل الحركات", movement_count),
                ("Today Movements / حركات اليوم", today_count),
                ("Open POs / طلبات مفتوحة", open_po_count),
                ("Users / المستخدمين", user_count),
                ("Expiry Alerts / تنبيهات الصلاحية", len(expiry_rows)),
                ("Low Stock / أقل من الحد الأدنى", len(low_stock_rows)),

            ]

            for index, card in enumerate(cards):
                row = index // 4
                col = index % 4
                self.cards_grid.addWidget(self.make_card(card[0], card[1]), row, col)

            self.set_table(
                self.today_table,
                [
                    "movement_id",
                    "datetime",
                    "type",
                    "generic_name",
                    "quantity",
                    "from_location",
                    "to_location",
                    "reference_no",
                ],
                get_today_movements(),
            )

            self.set_table(
                self.expiry_table,
                [
                    "risk",
                    "generic_name",
                    "location",
                    "batch_number",
                    "expiry_date",
                    "stock",
                ],
                expiry_rows,
            )

            self.set_table(
                self.low_stock_table,
                [
                    "item_id",
                    "item_code",
                    "generic_name",
                    "brand_name",
                    "current_stock",
                    "minimum_stock",
                ],
                low_stock_rows,
            )

            self.set_table(
                self.po_table,
                ["status", "count"],
                get_po_status_counts(),
            )

            self.set_table(
                self.recent_table,
                [
                    "movement_id",
                    "datetime",
                    "type",
                    "generic_name",
                    "quantity",
                    "from_location",
                    "to_location",
                    "reason",
                ],
                list_recent_movements(30),
            )

            self.status_label.setText("Dashboard refreshed successfully.")

        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    window = OperationalDashboard()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
