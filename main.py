
import sys
import os
import re
import csv
import sqlite3
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook, Workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QLineEdit, QDoubleSpinBox, QSpinBox, QMessageBox, QProgressBar,
    QHeaderView, QGroupBox, QFormLayout
)

APP_DIR = Path.home() / "ShopeeLiveMonitor"
APP_DIR.mkdir(exist_ok=True)
DB_PATH = APP_DIR / "products.db"

COLUMNS = [
    "Tên sản phẩm", "Giá", "Hoa hồng (%)", "Đơn hàng",
    "Lượt click", "Video đánh giá", "Shop Live",
    "Link sản phẩm", "Điểm", "Đánh giá"
]

ALIASES = {
    "Tên sản phẩm": [
        "tên sản phẩm", "sản phẩm", "product name", "name", "title"
    ],
    "Giá": ["giá", "giá bán", "price", "giá hiện tại"],
    "Hoa hồng (%)": [
        "hoa hồng", "hoa hồng (%)", "commission", "commission rate",
        "tỷ lệ hoa hồng", "phần trăm hoa hồng"
    ],
    "Đơn hàng": [
        "đơn hàng", "đã bán", "lượt bán", "sales", "orders",
        "sold", "số lượng bán"
    ],
    "Lượt click": ["lượt click", "click", "clicks"],
    "Video đánh giá": [
        "video đánh giá", "số video", "review video",
        "video review", "video"
    ],
    "Shop Live": [
        "shop live", "đang live", "số shop live", "live shops"
    ],
    "Link sản phẩm": [
        "link sản phẩm", "link", "url", "product link",
        "affiliate link"
    ],
}


def clean_number(value):
    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip().replace("đ", "").replace("₫", "")
    s = s.replace("%", "").replace(" ", "")

    if not s:
        return 0.0

    # Hỗ trợ số dạng 1.200.000 và 1,200,000
    if re.fullmatch(r"-?\d{1,3}([.,]\d{3})+", s):
        s = s.replace(".", "").replace(",", "")
    else:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return 0.0


def normalize_header(value):
    return str(value or "").strip().lower().replace("_", " ")


def init_database():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                price REAL,
                commission REAL,
                orders REAL,
                clicks REAL,
                videos REAL,
                live_shops REAL,
                link TEXT UNIQUE,
                score INTEGER,
                rating TEXT,
                updated_at TEXT
            )
        """)
        conn.commit()


def calculate_score(product):
    score = 0

    commission = product["commission"]
    orders = product["orders"]
    clicks = product["clicks"]
    videos = product["videos"]
    live = product["live_shops"]
    price = product["price"]

    if commission >= 25:
        score += 25
    elif commission >= 20:
        score += 22
    elif commission >= 15:
        score += 18
    elif commission > 0:
        score += 8

    if orders >= 1000:
        score += 25
    elif orders >= 500:
        score += 22
    elif orders >= 100:
        score += 17
    elif orders > 0:
        score += 8

    if clicks >= 1000:
        score += 15
    elif clicks >= 300:
        score += 12
    elif clicks > 0:
        score += 6

    if videos >= 10:
        score += 10
    elif videos > 0:
        score += 5

    if live >= 5:
        score += 15
    elif live >= 2:
        score += 12
    elif live > 0:
        score += 6

    if 20000 <= price <= 150000:
        score += 10
    elif price > 0:
        score += 4

    return min(score, 100)


def score_label(score):
    if score >= 80:
        return "Tiềm năng cao"
    if score >= 60:
        return "Nên thử nghiệm"
    return "Chưa ưu tiên"


class ShopeeMonitor(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Shopee Live Monitor")
        self.resize(1450, 850)

        init_database()

        self.products = []
        self.setup_ui()
        self.load_database()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        main = QVBoxLayout(central)
        main.setSpacing(10)

        title = QLabel("SHOPEE LIVE MONITOR")
        title.setStyleSheet(
            "font-size:26px;font-weight:bold;color:#16A34A;padding:8px;"
        )
        main.addWidget(title)

        subtitle = QLabel(
            "Công cụ lọc, phân tích và quản lý sản phẩm Shopee Affiliate"
        )
        main.addWidget(subtitle)

        stats = QHBoxLayout()
        self.total_label = QLabel("Tổng sản phẩm: 0")
        self.good_label = QLabel("Tiềm năng cao: 0")
        self.commission_label = QLabel("Hoa hồng đạt yêu cầu: 0")
        for label in (self.total_label, self.good_label,
                      self.commission_label):
            label.setStyleSheet(
                "background:#E8F5E9;padding:12px;border-radius:6px;"
                "font-weight:bold;"
            )
            stats.addWidget(label)
        main.addLayout(stats)

        actions = QHBoxLayout()

        self.import_button = QPushButton("📂 Nhập Excel/CSV")
        self.import_button.clicked.connect(self.import_files)

        self.export_button = QPushButton("📊 Xuất Excel")
        self.export_button.clicked.connect(self.export_excel)

        self.refresh_button = QPushButton("🔄 Tải lại dữ liệu")
        self.refresh_button.clicked.connect(self.load_database)

        self.delete_button = QPushButton("🗑 Xóa dữ liệu đang chọn")
        self.delete_button.clicked.connect(self.delete_selected)

        for button in (
            self.import_button, self.export_button,
            self.refresh_button, self.delete_button
        ):
            button.setMinimumHeight(38)
            actions.addWidget(button)

        main.addLayout(actions)

        filter_group = QGroupBox("BỘ LỌC SẢN PHẨM")
        filter_layout = QFormLayout(filter_group)

        self.keyword = QLineEdit()
        self.keyword.setPlaceholderText("Nhập tên sản phẩm hoặc từ khóa...")
        self.keyword.textChanged.connect(self.apply_filters)

        self.min_price = QSpinBox()
        self.min_price.setRange(0, 1000000000)
        self.min_price.setValue(20000)
        self.min_price.valueChanged.connect(self.apply_filters)

        self.max_price = QSpinBox()
        self.max_price.setRange(0, 1000000000)
        self.max_price.setValue(150000)
        self.max_price.valueChanged.connect(self.apply_filters)

        self.min_commission = QDoubleSpinBox()
        self.min_commission.setRange(0, 100)
        self.min_commission.setValue(15)
        self.min_commission.setSuffix("%")
        self.min_commission.valueChanged.connect(self.apply_filters)

        self.min_orders = QSpinBox()
        self.min_orders.setRange(0, 100000000)
        self.min_orders.setValue(100)
        self.min_orders.valueChanged.connect(self.apply_filters)

        filter_layout.addRow("Từ khóa:", self.keyword)
        filter_layout.addRow("Giá thấp nhất:", self.min_price)
        filter_layout.addRow("Giá cao nhất:", self.max_price)
        filter_layout.addRow("Hoa hồng tối thiểu:", self.min_commission)
        filter_layout.addRow("Đơn hàng tối thiểu:", self.min_orders)

        main.addWidget(filter_group)

        self.table = QTableWidget()
        self.table.setColumnCount(len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table.horizontalHeader().setStretchLastSection(True)

        main.addWidget(self.table)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        main.addWidget(self.progress)

        self.status = QLabel("Sẵn sàng. Hãy nhập file dữ liệu sản phẩm.")
        main.addWidget(self.status)

        self.setStyleSheet("""
            QMainWindow { background: #F7F8FA; }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #CCCCCC;
                border-radius: 7px;
                margin-top: 10px;
                padding: 10px;
            }
            QPushButton {
                background: #16A34A;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 14px;
                font-weight: bold;
            }
            QPushButton:hover { background: #15803D; }
            QLineEdit, QSpinBox, QDoubleSpinBox {
                padding: 7px;
                background: white;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
            QTableWidget {
                background: white;
                gridline-color: #DDDDDD;
            }
        """)

    def find_column(self, headers, target):
        aliases = ALIASES[target]
        normalized = [normalize_header(h) for h in headers]

        for alias in aliases:
            for index, header in enumerate(normalized):
                if header == alias:
                    return index

        return None

    def import_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn file dữ liệu Shopee",
            "",
            "Excel/CSV (*.xlsx *.xlsm *.csv)"
        )

        if not files:
            return

        imported = 0
        skipped = 0
        errors = []

        self.progress.setValue(0)

        for file_index, file_path in enumerate(files):
            try:
                if file_path.lower().endswith(".csv"):
                    rows = self.read_csv(file_path)
                else:
                    rows = self.read_excel(file_path)

                if not rows:
                    skipped += 1
                    continue

                headers = rows[0]
                indexes = {
                    key: self.find_column(headers, key)
                    for key in COLUMNS[:8]
                }

                if indexes["Tên sản phẩm"] is None:
                    raise ValueError("Không tìm thấy cột tên sản phẩm")

                if indexes["Link sản phẩm"] is None:
                    raise ValueError("Không tìm thấy cột link sản phẩm")

                for row in rows[1:]:
                    product = self.make_product(row, indexes)

                    if not product["name"] or not product["link"]:
                        continue

                    product["score"] = calculate_score(product)
                    product["rating"] = score_label(product["score"])

                    if self.save_product(product):
                        imported += 1

            except Exception as error:
                errors.append(
                    f"{Path(file_path).name}: {str(error)}"
                )

            self.progress.setValue(
                int((file_index + 1) / len(files) * 100)
            )

        self.load_database()

        message = (
            f"Đã thêm {imported} sản phẩm mới.\n"
            f"Đã bỏ qua {skipped} file rỗng."
        )

        if errors:
            message += "\n\nLỗi:\n" + "\n".join(errors[:10])

        QMessageBox.information(self, "Kết quả nhập dữ liệu", message)

    def read_excel(self, path):
        workbook = load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        workbook.close()
        return rows

    def read_csv(self, path):
        for encoding in ("utf-8-sig", "utf-8", "cp1258"):
            try:
                with open(path, "r", encoding=encoding, newline="") as f:
                    return list(csv.reader(f))
            except UnicodeDecodeError:
                continue
        raise ValueError("Không đọc được mã hóa CSV")

    def make_product(self, row, indexes):
        def get_value(field):
            index = indexes.get(field)
            if index is None or index >= len(row):
                return ""
            return row[index]

        return {
            "name": str(get_value("Tên sản phẩm") or "").strip(),
            "price": clean_number(get_value("Giá")),
            "commission": clean_number(get_value("Hoa hồng (%)")),
            "orders": clean_number(get_value("Đơn hàng")),
            "clicks": clean_number(get_value("Lượt click")),
            "videos": clean_number(get_value("Video đánh giá")),
            "live_shops": clean_number(get_value("Shop Live")),
            "link": str(get_value("Link sản phẩm") or "").strip(),
        }

    def save_product(self, product):
        try:
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("""
                    INSERT INTO products (
                        name, price, commission, orders, clicks,
                        videos, live_shops, link, score, rating, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(link) DO UPDATE SET
                        name=excluded.name,
                        price=excluded.price,
                        commission=excluded.commission,
                        orders=excluded.orders,
                        clicks=excluded.clicks,
                        videos=excluded.videos,
                        live_shops=excluded.live_shops,
                        score=excluded.score,
                        rating=excluded.rating,
                        updated_at=excluded.updated_at
                """, (
                    product["name"], product["price"],
                    product["commission"], product["orders"],
                    product["clicks"], product["videos"],
                    product["live_shops"], product["link"],
                    product["score"], product["rating"],
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ))
            return True
        except sqlite3.Error:
            return False

    def load_database(self):
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute("""
                SELECT name, price, commission, orders, clicks,
                       videos, live_shops, link, score, rating
                FROM products ORDER BY score DESC
            """).fetchall()

        self.products = []
        for row in rows:
            self.products.append({
                "name": row[0],
                "price": row[1],
                "commission": row[2],
                "orders": row[3],
                "clicks": row[4],
                "videos": row[5],
                "live_shops": row[6],
                "link": row[7],
                "score": row[8],
                "rating": row[9],
            })

        self.apply_filters()

    def apply_filters(self):
        keyword = self.keyword.text().strip().lower()
        min_price = self.min_price.value()
        max_price = self.max_price.value()
        min_commission = self.min_commission.value()
        min_orders = self.min_orders.value()

        filtered = []

        for product in self.products:
            if keyword and keyword not in product["name"].lower():
                continue

            if product["price"] < min_price:
                continue

            if max_price > 0 and product["price"] > max_price:
                continue

            if product["commission"] < min_commission:
                continue

            if product["orders"] < min_orders:
                continue

            filtered.append(product)

        self.show_products(filtered)

    def show_products(self, products):
        self.table.setRowCount(len(products))

        for row_index, product in enumerate(products):
            values = [
                product["name"],
                f'{product["price"]:,.0f}',
                f'{product["commission"]:g}%',
                f'{product["orders"]:,.0f}',
                f'{product["clicks"]:,.0f}',
                f'{product["videos"]:,.0f}',
                f'{product["live_shops"]:,.0f}',
                product["link"],
                str(product["score"]) + "/100",
                product["rating"]
            ]

            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if column in (1, 2, 3, 4, 5, 6, 8):
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignCenter
                    )
                self.table.setItem(row_index, column, item)

        total = len(self.products)
        good = sum(1 for p in self.products if p["score"] >= 80)
        commission_count = sum(
            1 for p in self.products
            if p["commission"] >= self.min_commission.value()
        )

        self.total_label.setText(f"Tổng sản phẩm: {total}")
        self.good_label.setText(f"Tiềm năng cao: {good}")
        self.commission_label.setText(
            f"Hoa hồng đạt yêu cầu: {commission_count}"
        )
        self.status.setText(f"Đang hiển thị {len(products)} sản phẩm.")

    def export_excel(self):
        if not self.products:
            QMessageBox.warning(
                self, "Thông báo", "Chưa có dữ liệu để xuất."
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu file Excel",
            "ShopeeLiveMonitor.xlsx",
            "Excel (*.xlsx)"
        )

        if not path:
            return

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Shopee Products"
        sheet.append(COLUMNS)

        for product in self.products:
            sheet.append([
                product["name"], product["price"],
                product["commission"], product["orders"],
                product["clicks"], product["videos"],
                product["live_shops"], product["link"],
                product["score"], product["rating"]
            ])

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        for column in sheet.columns:
            letter = column[0].column_letter
            max_length = max(
                len(str(cell.value or "")) for cell in column
            )
            sheet.column_dimensions[letter].width = min(
                max(max_length + 2, 12), 45
            )

        workbook.save(path)
        QMessageBox.information(
            self, "Thành công", f"Đã xuất Excel:\n{path}"
        )

    def delete_selected(self):
        selected = self.table.selectionModel().selectedRows()

        if not selected:
            QMessageBox.warning(
                self, "Thông báo", "Hãy chọn sản phẩm muốn xóa."
            )
            return

        reply = QMessageBox.question(
            self,
            "Xác nhận",
            f"Bạn có chắc muốn xóa {len(selected)} sản phẩm đã chọn?"
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        links = []
        for index in selected:
            link_item = self.table.item(index.row(), 7)
            if link_item:
                links.append(link_item.text())

        with sqlite3.connect(DB_PATH) as conn:
            conn.executemany(
                "DELETE FROM products WHERE link=?",
                [(link,) for link in links]
            )

        self.load_database()
        QMessageBox.information(
            self, "Thành công", "Đã xóa sản phẩm được chọn."
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ShopeeMonitor()
    window.show()
    sys.exit(app.exec())
