import sys
import os
import re
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap, QIcon, QColor, QFont


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def load_zip_lookup():
    zip_map = {}
    try:
        path = resource_path("Zipcodes-with-Reps-Complete.xlsx")
        if not os.path.exists(path):
            print(f"WARNING: Lookup file not found at {path}")
            return zip_map
        df = pd.read_excel(path, sheet_name="Full Scrape", dtype=str)
        df.columns = [c.strip() for c in df.columns]
        for _, row in df.iterrows():
            raw_zip = str(row.get("Zipcode", "")).strip()
            zipcode = raw_zip.split("-")[0].strip().zfill(5)
            if len(zipcode) != 5 or not zipcode.isdigit():
                continue
            council  = str(row.get("City Council District", "")).strip()
            assembly = str(row.get("State Assembly District", "")).strip()
            senate   = str(row.get("State Senate District", "")).strip()
            congress = str(row.get("US House District", "")).strip()
            if zipcode not in zip_map and council and council.lower() != "nan":
                zip_map[zipcode] = {
                    "City Council District":   council,
                    "State Assembly District": assembly,
                    "State Senate District":   senate,
                    "US House District":       congress,
                }
        print(f"ZIP lookup loaded: {len(zip_map)} unique zip codes.")
    except Exception as e:
        print(f"Error loading ZIP lookup: {e}")
    return zip_map


def scrape_mygovnyc(address, city, state="NY"):
    result = {
        "City Council District": "", "State Assembly District": "",
        "State Senate District": "", "US House District": ""
    }
    try:
        full_addr = f"{address}, {city}, {state}"
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        resp = requests.get(
            "https://www.mygovnyc.org/",
            params={"address": full_addr},
            headers=headers,
            timeout=12
        )
        if resp.status_code != 200:
            return result
        soup = BeautifulSoup(resp.text, "lxml")

        for el in soup.find_all(string=re.compile(r"city council", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"district\s+(\d+)", text, re.I)
            if m:
                result["City Council District"] = f"District {m.group(1)}"
                break

        for el in soup.find_all(string=re.compile(r"assembly", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"district\s+(\d+)", text, re.I)
            if m:
                result["State Assembly District"] = f"Assembly District {m.group(1)}"
                break

        for el in soup.find_all(string=re.compile(r"state senate", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"district\s+(\d+)", text, re.I)
            if m:
                result["State Senate District"] = f"Senate District {m.group(1)}"
                break

        for el in soup.find_all(string=re.compile(r"congressional|us house|u\.s\. house", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"(NY\s*\d+|\d+(?:st|nd|rd|th)\s+district)", text, re.I)
            if m:
                result["US House District"] = m.group(1).strip()
                break

    except Exception as e:
        print(f"Scrape error for '{address}': {e}")
    return result


def detect_col(columns, *keywords):
    for col in columns:
        lower = col.lower()
        if any(k in lower for k in keywords):
            return col
    return None


class LookupWorker(QThread):
    progress       = pyqtSignal(int, str)
    result_ready   = pyqtSignal(object)
    error_occurred = pyqtSignal(str)

    def __init__(self, df, zip_lookup):
        super().__init__()
        self.df         = df.copy()
        self.zip_lookup = zip_lookup
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        try:
            cols  = list(self.df.columns)
            total = len(self.df)

            zip_col   = detect_col(cols, "zip")
            addr_col  = detect_col(cols, "address")
            city_col  = detect_col(cols, "city")
            state_col = detect_col(cols, "state")

            council_list  = []
            assembly_list = []
            senate_list   = []
            congress_list = []

            for i, (_, row) in enumerate(self.df.iterrows()):
                if self._cancelled:
                    break

                raw_zip = str(row[zip_col]).strip()  if zip_col  else ""
                address = str(row[addr_col]).strip() if addr_col else ""
                city    = str(row[city_col]).strip() if city_col else ""
                state   = str(row[state_col]).strip().upper() if state_col else "NY"

                zipcode = raw_zip.split("-")[0].strip()
                if zipcode and len(zipcode) < 5:
                    zipcode = zipcode.zfill(5)

                council = assembly = senate = congress = ""

                if zipcode in self.zip_lookup:
                    d        = self.zip_lookup[zipcode]
                    council  = d.get("City Council District",   "")
                    assembly = d.get("State Assembly District", "")
                    senate   = d.get("State Senate District",   "")
                    congress = d.get("US House District",       "")
                elif state in ("NY", "NEW YORK") and address not in ("", "-", "'-", "nan"):
                    self.progress.emit(
                        int(i / total * 100),
                        f"[Web] Scraping {i+1}/{total}: {address[:40]}..."
                    )
                    scraped  = scrape_mygovnyc(address, city, state)
                    council  = scraped["City Council District"]
                    assembly = scraped["State Assembly District"]
                    senate   = scraped["State Senate District"]
                    congress = scraped["US House District"]
                    time.sleep(0.4)

                council_list.append(council)
                assembly_list.append(assembly)
                senate_list.append(senate)
                congress_list.append(congress)

                self.progress.emit(
                    int((i + 1) / total * 100),
                    f"Processing {i+1:,} of {total:,}   (ZIP: {zipcode})"
                )

            result = self.df.copy()
            result["City Council District"]   = council_list
            result["State Assembly District"] = assembly_list
            result["State Senate District"]   = senate_list
            result["US House District"]       = congress_list

            self.result_ready.emit(result)

        except Exception as e:
            import traceback
            self.error_occurred.emit(traceback.format_exc())


# ─────────────────────────────────────────────────────────────
#  Reusable card / stat widget
# ─────────────────────────────────────────────────────────────
class StatCard(QFrame):
    def __init__(self, label: str, value: str = "—", accent: str = "#1565C0"):
        super().__init__()
        self.setObjectName("statCard")
        self.accent = accent
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(4)

        self.val_lbl = QLabel(value)
        self.val_lbl.setAlignment(Qt.AlignCenter)
        font = QFont("Segoe UI", 22, QFont.Bold)
        self.val_lbl.setFont(font)
        self.val_lbl.setStyleSheet(f"color:{accent};")

        self.key_lbl = QLabel(label)
        self.key_lbl.setAlignment(Qt.AlignCenter)
        self.key_lbl.setStyleSheet("color:#546E7A;font-size:11px;font-weight:600;letter-spacing:0.5px;")

        layout.addWidget(self.val_lbl)
        layout.addWidget(self.key_lbl)

    def set_value(self, v: str):
        self.val_lbl.setText(v)


# ─────────────────────────────────────────────────────────────
#  Main Window
# ─────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df         = None
        self.result_df  = None
        self.worker     = None
        self.zip_lookup = load_zip_lookup()
        self._setup_ui()
        self._apply_theme()

    # ── UI Construction ──────────────────────────────────────
    def _setup_ui(self):
        self.setWindowTitle("NYC Representatives Lookup  —  QEDC")
        self.setMinimumSize(1180, 780)

        ico = resource_path("icon_256x256.ico")
        if os.path.exists(ico):
            self.setWindowIcon(QIcon(ico))

        root = QWidget()
        self.setCentralWidget(root)
        master = QVBoxLayout(root)
        master.setContentsMargins(0, 0, 0, 0)
        master.setSpacing(0)

        # ── Top navigation bar ───────────────────────────────
        nav = QFrame(); nav.setObjectName("navBar")
        nav.setFixedHeight(64)
        nav_lay = QHBoxLayout(nav)
        nav_lay.setContentsMargins(24, 0, 24, 0)

        logo_lbl = QLabel()
        logo_path = resource_path("QEDC-Full-Logo-Primary-Color.jpg")
        if os.path.exists(logo_path):
            px = QPixmap(logo_path).scaled(160, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_lbl.setPixmap(px)
        else:
            logo_lbl.setText("QEDC")
            logo_lbl.setStyleSheet("font-size:20px;font-weight:bold;color:#1565C0;")

        app_title = QLabel("NYC Representatives Lookup")
        app_title.setObjectName("navTitle")

        nav_lay.addWidget(logo_lbl)
        nav_lay.addStretch()
        nav_lay.addWidget(app_title)

        # ── Content area (padded) ────────────────────────────
        content_wrap = QWidget(); content_wrap.setObjectName("contentArea")
        content = QVBoxLayout(content_wrap)
        content.setContentsMargins(28, 20, 28, 20)
        content.setSpacing(16)

        # ── Toolbar card ─────────────────────────────────────
        toolbar = QFrame(); toolbar.setObjectName("toolCard")
        tb_lay  = QHBoxLayout(toolbar)
        tb_lay.setContentsMargins(20, 14, 20, 14)
        tb_lay.setSpacing(12)

        self.load_btn = QPushButton("  📂  Load File")
        self.load_btn.setObjectName("btnPrimary")
        self.load_btn.setFixedHeight(40)
        self.load_btn.clicked.connect(self.load_file)

        self.file_lbl = QLabel("No file loaded — supported formats: CSV, XLSX")
        self.file_lbl.setObjectName("fileLabel")
        self.file_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self.run_btn = QPushButton("  ▶  Run Lookup")
        self.run_btn.setObjectName("btnSuccess")
        self.run_btn.setFixedHeight(40)
        self.run_btn.setEnabled(False)
        self.run_btn.clicked.connect(self.run_lookup)

        self.export_btn = QPushButton("  💾  Export Excel")
        self.export_btn.setObjectName("btnWarning")
        self.export_btn.setFixedHeight(40)
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)

        for w in [self.load_btn, self.file_lbl, self.run_btn, self.export_btn]:
            tb_lay.addWidget(w)

        # ── Progress bar ─────────────────────────────────────
        prog_card = QFrame(); prog_card.setObjectName("toolCard")
        prog_lay  = QVBoxLayout(prog_card)
        prog_lay.setContentsMargins(20, 10, 20, 10)
        prog_lay.setSpacing(4)

        self.prog_bar = QProgressBar()
        self.prog_bar.setFixedHeight(8)
        self.prog_bar.setTextVisible(False)
        self.prog_bar.setVisible(False)

        self.status_lbl = QLabel("Load a CSV or Excel file to begin.")
        self.status_lbl.setObjectName("statusLabel")

        prog_lay.addWidget(self.status_lbl)
        prog_lay.addWidget(self.prog_bar)

        # ── Stat cards row ───────────────────────────────────
        stats_row = QHBoxLayout()
        stats_row.setSpacing(12)

        self.card_total     = StatCard("TOTAL RECORDS",      "—",      "#1565C0")
        self.card_matched   = StatCard("MATCHED",            "—",      "#2E7D32")
        self.card_unmatched = StatCard("UNMATCHED",          "—",      "#C62828")
        self.card_cache     = StatCard("ZIP CACHE ENTRIES",
                                       f"{len(self.zip_lookup):,}",    "#6A1B9A")

        for c in [self.card_total, self.card_matched, self.card_unmatched, self.card_cache]:
            stats_row.addWidget(c)

        # ── Section header ───────────────────────────────────
        tbl_hdr = QLabel("Results Preview")
        tbl_hdr.setObjectName("sectionHeader")

        # ── Table ─────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setObjectName("dataTable")
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.verticalHeader().setVisible(False)

        # ── Footer ────────────────────────────────────────────
        footer = QLabel("Queens Economic Development Corporation  •  NYC Representatives Lookup Tool")
        footer.setObjectName("footerLabel")
        footer.setAlignment(Qt.AlignCenter)

        # Assembly
        content.addWidget(toolbar)
        content.addWidget(prog_card)
        content.addLayout(stats_row)
        content.addWidget(tbl_hdr)
        content.addWidget(self.table)
        content.addWidget(footer)
        content.setStretch(4, 1)

        master.addWidget(nav)
        master.addWidget(content_wrap)

    # ── Theme ────────────────────────────────────────────────
    def _apply_theme(self):
        self.setStyleSheet("""
        /* ── Global ── */
        QMainWindow, QWidget {
            background: #F4F6F9;
            color: #263238;
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
            font-size: 13px;
        }

        /* ── Nav bar ── */
        #navBar {
            background: #FFFFFF;
            border-bottom: 2px solid #E3E8EF;
        }
        #navTitle {
            font-size: 17px;
            font-weight: 700;
            color: #1565C0;
            letter-spacing: 0.3px;
        }

        /* ── Content area ── */
        #contentArea { background: #F4F6F9; }

        /* ── Cards ── */
        #toolCard {
            background: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #DDE3ED;
        }
        #statCard {
            background: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #DDE3ED;
            min-width: 160px;
        }

        /* ── Buttons ── */
        QPushButton {
            padding: 0 20px;
            border-radius: 7px;
            font-weight: 600;
            font-size: 13px;
            border: none;
            min-width: 130px;
        }
        #btnPrimary {
            background: #1565C0;
            color: #FFFFFF;
        }
        #btnPrimary:hover  { background: #1976D2; }
        #btnPrimary:pressed{ background: #0D47A1; }

        #btnSuccess {
            background: #2E7D32;
            color: #FFFFFF;
        }
        #btnSuccess:hover  { background: #388E3C; }
        #btnSuccess:pressed{ background: #1B5E20; }
        #btnSuccess:disabled{ background: #BDBDBD; color: #757575; }

        #btnWarning {
            background: #E65100;
            color: #FFFFFF;
        }
        #btnWarning:hover  { background: #F4511E; }
        #btnWarning:pressed{ background: #BF360C; }
        #btnWarning:disabled{ background: #BDBDBD; color: #757575; }

        /* ── File label ── */
        #fileLabel {
            color: #455A64;
            font-size: 12px;
            padding: 0 8px;
        }

        /* ── Progress ── */
        QProgressBar {
            background: #E3E8EF;
            border-radius: 4px;
            border: none;
        }
        QProgressBar::chunk {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 #1565C0, stop:1 #42A5F5
            );
            border-radius: 4px;
        }

        /* ── Status ── */
        #statusLabel {
            color: #546E7A;
            font-size: 12px;
        }

        /* ── Section header ── */
        #sectionHeader {
            font-size: 14px;
            font-weight: 700;
            color: #37474F;
            padding: 4px 2px;
        }

        /* ── Table ── */
        #dataTable {
            background: #FFFFFF;
            alternate-background-color: #F8FAFC;
            border: 1px solid #DDE3ED;
            border-radius: 10px;
            gridline-color: transparent;
            outline: none;
        }
        QHeaderView::section {
            background: #1565C0;
            color: #FFFFFF;
            padding: 10px 12px;
            border: none;
            font-weight: 700;
            font-size: 12px;
            letter-spacing: 0.4px;
        }
        QHeaderView::section:first { border-top-left-radius: 8px; }
        QHeaderView::section:last  { border-top-right-radius: 8px; }

        QTableWidget::item {
            padding: 6px 10px;
            border-bottom: 1px solid #EEF1F5;
        }
        QTableWidget::item:selected {
            background: #BBDEFB;
            color: #0D47A1;
        }

        QScrollBar:vertical {
            background: #F4F6F9;
            width: 8px;
            margin: 0;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical {
            background: #B0BEC5;
            border-radius: 4px;
            min-height: 30px;
        }
        QScrollBar::handle:vertical:hover { background: #78909C; }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical { height: 0; }

        QScrollBar:horizontal {
            background: #F4F6F9;
            height: 8px;
            border-radius: 4px;
        }
        QScrollBar::handle:horizontal {
            background: #B0BEC5;
            border-radius: 4px;
        }
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal { width: 0; }

        /* ── Footer ── */
        #footerLabel {
            color: #90A4AE;
            font-size: 11px;
            padding: 8px 0 2px 0;
        }

        /* ── Message boxes ── */
        QMessageBox {
            background: #FFFFFF;
        }
        QMessageBox QPushButton {
            min-width: 80px;
            padding: 6px 16px;
            background: #1565C0;
            color: white;
            border-radius: 6px;
        }
        """)

    # ── Slots ────────────────────────────────────────────────
    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Data File", "",
            "Data Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            if path.lower().endswith((".xlsx", ".xls")):
                self.df = pd.read_excel(path, dtype=str)
            else:
                self.df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
            self.df.fillna("", inplace=True)
            self._show_table(self.df.head(100))
            self.file_lbl.setText(f"📄  {os.path.basename(path)}   ({len(self.df):,} rows)")
            self.run_btn.setEnabled(True)
            self.export_btn.setEnabled(False)
            self.result_df = None
            self.card_total.set_value(f"{len(self.df):,}")
            self.card_matched.set_value("—")
            self.card_unmatched.set_value("—")
            self.status_lbl.setText(f"✅  File loaded — {len(self.df):,} rows. Click  ▶ Run Lookup  to begin.")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))

    def _show_table(self, df):
        self.table.clear()
        if df is None or df.empty:
            return
        HIGHLIGHT = {
            "City Council District", "State Assembly District",
            "State Senate District", "US House District"
        }
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(list(df.columns))
        for r in range(len(df)):
            for c, col_name in enumerate(df.columns):
                val = str(df.iloc[r, c])
                if val == "nan":
                    val = ""
                item = QTableWidgetItem(val)
                if col_name in HIGHLIGHT:
                    if val:
                        item.setForeground(QColor("#1B5E20"))
                        item.setBackground(QColor("#F1F8E9"))
                    else:
                        item.setForeground(QColor("#B71C1C"))
                        item.setText("—")
                        item.setBackground(QColor("#FFF3E0"))
                self.table.setItem(r, c, item)

    def run_lookup(self):
        if self.df is None:
            return
        self.run_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.prog_bar.setValue(0)
        self.status_lbl.setText("Starting lookup…")
        self.worker = LookupWorker(self.df, self.zip_lookup)
        self.worker.progress.connect(self._on_progress)
        self.worker.result_ready.connect(self._on_result)
        self.worker.error_occurred.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, pct, msg):
        self.prog_bar.setValue(pct)
        self.status_lbl.setText(msg)

    def _on_result(self, df):
        self.result_df = df
        d_col    = "City Council District"
        matched  = (df[d_col].str.strip() != "").sum() if d_col in df.columns else 0
        unmatched = len(df) - matched
        self._show_table(df.head(200))
        self.card_total.set_value(f"{len(df):,}")
        self.card_matched.set_value(f"{matched:,}")
        self.card_unmatched.set_value(f"{unmatched:,}")
        self.prog_bar.setValue(100)
        self.status_lbl.setText(
            f"✅  Lookup complete  —  {matched:,} matched  |  {unmatched:,} unmatched  "
            f"(showing first 200 rows in preview)"
        )
        self.run_btn.setEnabled(True)
        self.export_btn.setEnabled(True)

    def _on_error(self, msg):
        QMessageBox.critical(self, "Lookup Error", msg)
        self.run_btn.setEnabled(True)
        self.prog_bar.setVisible(False)

    def export_results(self):
        if self.result_df is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Results", "results.xlsx", "Excel Workbook (*.xlsx)"
        )
        if not path:
            return
        try:
            self.result_df.to_excel(path, index=False, sheet_name="Results")
            QMessageBox.information(self, "Export Successful",
                                    f"File saved successfully:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
