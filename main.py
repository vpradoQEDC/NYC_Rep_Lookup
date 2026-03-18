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
    QTableWidgetItem, QHeaderView, QMessageBox, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap, QIcon, QColor


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def load_zip_lookup():
    """Build zipcode → district map from bundled Excel file."""
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
            # Only store if we have at least council district and it's not blank/nan
            if zipcode not in zip_map and council and council.lower() != "nan":
                zip_map[zipcode] = {
                    "City Council District":    council,
                    "State Assembly District":  assembly,
                    "State Senate District":    senate,
                    "US House District":        congress,
                }
        print(f"ZIP lookup loaded: {len(zip_map)} unique zip codes.")
    except Exception as e:
        print(f"Error loading ZIP lookup: {e}")
    return zip_map


def scrape_mygovnyc(address, city, state="NY"):
    """
    Fallback: scrape district info from mygovnyc.org for addresses
    not in the local cache. Returns dict with district keys.
    """
    result = {"City Council District": "", "State Assembly District": "",
              "State Senate District": "", "US House District": ""}
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

        # Council District
        for el in soup.find_all(string=re.compile(r"city council", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"district\s+(\d+)", text, re.I)
            if m:
                result["City Council District"] = f"District {m.group(1)}"
                break

        # Assembly District
        for el in soup.find_all(string=re.compile(r"assembly", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"district\s+(\d+)", text, re.I)
            if m:
                result["State Assembly District"] = f"Assembly District {m.group(1)}"
                break

        # Senate District
        for el in soup.find_all(string=re.compile(r"state senate", re.I)):
            text = el.parent.get_text(" ", strip=True)
            m = re.search(r"district\s+(\d+)", text, re.I)
            if m:
                result["State Senate District"] = f"Senate District {m.group(1)}"
                break

        # Congress
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
    """Find the first column whose name contains any of the keywords (case-insensitive)."""
    for col in columns:
        lower = col.lower()
        if any(k in lower for k in keywords):
            return col
    return None


class LookupWorker(QThread):
    progress      = pyqtSignal(int, str)
    result_ready  = pyqtSignal(object)
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

            zip_col     = detect_col(cols, "zip")
            addr_col    = detect_col(cols, "address")
            city_col    = detect_col(cols, "city")
            state_col   = detect_col(cols, "state")

            council_list  = []
            assembly_list = []
            senate_list   = []
            congress_list = []

            for i, (_, row) in enumerate(self.df.iterrows()):
                if self._cancelled:
                    break

                raw_zip  = str(row[zip_col]).strip()  if zip_col  else ""
                address  = str(row[addr_col]).strip() if addr_col else ""
                city     = str(row[city_col]).strip() if city_col else ""
                state    = str(row[state_col]).strip().upper() if state_col else "NY"

                # Clean zipcode
                zipcode = raw_zip.split("-")[0].strip()
                if zipcode and len(zipcode) < 5:
                    zipcode = zipcode.zfill(5)

                council = assembly = senate = congress = ""

                if zipcode in self.zip_lookup:
                    d = self.zip_lookup[zipcode]
                    council  = d.get("City Council District", "")
                    assembly = d.get("State Assembly District", "")
                    senate   = d.get("State Senate District", "")
                    congress = d.get("US House District", "")
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
                    time.sleep(0.4)   # polite rate limit

                council_list.append(council)
                assembly_list.append(assembly)
                senate_list.append(senate)
                congress_list.append(congress)

                self.progress.emit(
                    int((i + 1) / total * 100),
                    f"Processing {i+1:,} / {total:,}  (ZIP: {zipcode})"
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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df         = None
        self.result_df  = None
        self.worker     = None
        self.zip_lookup = load_zip_lookup()
        self._setup_ui()
        self._apply_theme()

    def _setup_ui(self):
        self.setWindowTitle("NYC Representatives Lookup  —  QEDC")
        self.setMinimumSize(1100, 720)

        ico = resource_path("icon_256x256.ico")
        if os.path.exists(ico):
            self.setWindowIcon(QIcon(ico))

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # ── Header ──────────────────────────────────────────────
        hdr = QFrame(); hdr.setObjectName("hdrFrame")
        hdr_lay = QHBoxLayout(hdr)
        logo_path = resource_path("QEDC-Full-Logo-Primary-Color.jpg")
        logo_lbl  = QLabel()
        if os.path.exists(logo_path):
            px = QPixmap(logo_path).scaled(220, 65, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_lbl.setPixmap(px)
        else:
            logo_lbl.setText("QEDC")
            logo_lbl.setStyleSheet("font-size:22px;font-weight:bold;color:#00D4FF;")
        title = QLabel("NYC Representatives Lookup")
        title.setObjectName("appTitle")
        hdr_lay.addWidget(logo_lbl)
        hdr_lay.addStretch()
        hdr_lay.addWidget(title)

        # ── Controls ─────────────────────────────────────────────
        ctrl = QFrame()
        ctrl_lay = QHBoxLayout(ctrl)
        self.load_btn   = QPushButton("📂  Load CSV / Excel")
        self.load_btn.setObjectName("btnBlue")
        self.load_btn.clicked.connect(self.load_file)

        self.file_lbl   = QLabel("No file loaded")
        self.file_lbl.setObjectName("fileLbl")

        self.run_btn    = QPushButton("▶  Look Up Representatives")
        self.run_btn.setObjectName("btnGreen")
        self.run_btn.setEnabled(False)
        self.run_btn.clicked.connect(self.run_lookup)

        self.export_btn = QPushButton("💾  Export to Excel")
        self.export_btn.setObjectName("btnOrange")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)

        for w in [self.load_btn, self.file_lbl, self.run_btn, self.export_btn]:
            ctrl_lay.addWidget(w)
        ctrl_lay.setStretch(1, 1)

        # ── Progress ─────────────────────────────────────────────
        self.prog_bar   = QProgressBar(); self.prog_bar.setVisible(False)
        self.status_lbl = QLabel(""); self.status_lbl.setAlignment(Qt.AlignCenter)
        self.status_lbl.setObjectName("statusLbl")

        # ── Stats ─────────────────────────────────────────────────
        stats = QFrame(); stats.setObjectName("statsFrame")
        s_lay = QHBoxLayout(stats)
        self.lbl_total     = QLabel("Total: —")
        self.lbl_matched   = QLabel("Matched: —")
        self.lbl_unmatched = QLabel("Unmatched: —")
        self.lbl_cached    = QLabel(f"ZIP Cache: {len(self.zip_lookup):,} entries")
        for l in [self.lbl_total, self.lbl_matched, self.lbl_unmatched, self.lbl_cached]:
            l.setObjectName("statLbl"); l.setAlignment(Qt.AlignCenter)
            s_lay.addWidget(l)

        # ── Table ─────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setObjectName("mainTable")
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        for w in [hdr, ctrl, self.prog_bar, self.status_lbl, stats, self.table]:
            layout.addWidget(w)
        layout.setStretch(5, 1)

    def _apply_theme(self):
        self.setStyleSheet("""
        QMainWindow,QWidget{background:#1A1A2E;color:#E0E0E0;
          font-family:'Segoe UI',Arial,sans-serif;font-size:13px;}
        #hdrFrame{background:#16213E;border-radius:8px;
          border:1px solid #0F3460;padding:8px;}
        #appTitle{font-size:20px;font-weight:bold;color:#00D4FF;}
        QPushButton{padding:9px 18px;border-radius:6px;
          font-weight:bold;font-size:13px;border:none;min-width:155px;}
        #btnBlue{background:#0F3460;color:#fff;}
        #btnBlue:hover{background:#1a5276;}
        #btnGreen{background:#1ABC9C;color:#000;}
        #btnGreen:hover{background:#17a589;}
        #btnGreen:disabled{background:#444;color:#777;}
        #btnOrange{background:#E67E22;color:#000;}
        #btnOrange:hover{background:#ca6f1e;}
        #btnOrange:disabled{background:#444;color:#777;}
        #fileLbl{color:#00D4FF;font-size:12px;padding:0 8px;}
        QProgressBar{border:1px solid #0F3460;border-radius:5px;
          background:#16213E;height:20px;text-align:center;color:#fff;}
        QProgressBar::chunk{background:#1ABC9C;border-radius:4px;}
        #statusLbl{color:#BDC3C7;font-size:12px;}
        #statsFrame{background:#16213E;border-radius:6px;
          border:1px solid #0F3460;padding:6px;}
        #statLbl{font-size:13px;font-weight:bold;color:#00D4FF;padding:4px 15px;}
        QTableWidget{background:#16213E;alternate-background-color:#1A1A2E;
          gridline-color:#0F3460;border:1px solid #0F3460;
          border-radius:6px;color:#E0E0E0;}
        QHeaderView::section{background:#0F3460;color:#00D4FF;
          padding:7px;border:none;font-weight:bold;}
        QTableWidget::item:selected{background:#1ABC9C;color:#000;}
        QScrollBar:vertical{background:#16213E;width:10px;}
        QScrollBar::handle:vertical{background:#0F3460;border-radius:5px;}
        """)

    # ── Slots ────────────────────────────────────────────────────
    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open File", "",
            "Data Files (*.csv *.xlsx *.xls);;CSV (*.csv);;Excel (*.xlsx *.xls)"
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
            self.file_lbl.setText(f"📄 {os.path.basename(path)}  ({len(self.df):,} rows)")
            self.run_btn.setEnabled(True)
            self.export_btn.setEnabled(False)
            self.result_df = None
            self.lbl_total.setText(f"Total: {len(self.df):,}")
            self.lbl_matched.setText("Matched: —")
            self.lbl_unmatched.setText("Unmatched: —")
            self.status_lbl.setText("File loaded. Click ▶ Look Up Representatives to begin.")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))

    def _show_table(self, df):
        self.table.clear()
        if df is None or df.empty:
            return
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(list(df.columns))
        district_cols = {"City Council District","State Assembly District",
                         "State Senate District","US House District"}
        for r in range(len(df)):
            for c, col_name in enumerate(df.columns):
                val = str(df.iloc[r, c])
                if val == "nan": val = ""
                item = QTableWidgetItem(val)
                if col_name in district_cols:
                    if val:
                        item.setForeground(QColor("#1ABC9C"))
                    else:
                        item.setForeground(QColor("#E74C3C"))
                        item.setText("—")
                self.table.setItem(r, c, item)

    def run_lookup(self):
        if self.df is None:
            return
        self.run_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.prog_bar.setValue(0)
        self.status_lbl.setText("Starting…")
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
        d_col = "City Council District"
        matched   = (df[d_col].str.strip() != "").sum() if d_col in df.columns else 0
        unmatched = len(df) - matched
        self._show_table(df.head(100))
        self.lbl_total.setText(f"Total: {len(df):,}")
        self.lbl_matched.setText(f"✅ Matched: {matched:,}")
        self.lbl_unmatched.setText(f"❌ Unmatched: {unmatched:,}")
        self.prog_bar.setValue(100)
        self.status_lbl.setText(f"✅ Done! {matched:,}/{len(df):,} matched with representatives.")
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
            self, "Save Results", "results.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            self.result_df.to_excel(path, index=False, sheet_name="Results")
            QMessageBox.information(self, "Saved", f"Results saved:\n{path}")
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
