import sys
import csv
import os
import sqlite3
import requests
import pandas as pd
import webbrowser
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QTabWidget, QProgressBar, QComboBox, QMessageBox, QScrollArea
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl, QSize
from PyQt5.QtGui import QPalette, QColor, QPixmap, QIcon
from bs4 import BeautifulSoup

# ============================================================================
# HELPER FUNCTIONS FOR FLEXIBLE COLUMN DETECTION
# ============================================================================

def detect_column_mappings(df):
    """
    Intelligently detect Address, City, State, Zipcode columns
    by checking column names, even if they don't match exactly.
    """
    columns = df.columns.str.lower().tolist()
    
    # Define keywords to look for
    address_keywords = ['address', 'street', 'addr', 'location', 'road', 'avenue', 'boulevard']
    city_keywords = ['city', 'municipality', 'town', 'metro_area', 'place']
    state_keywords = ['state', 'state_code', 'st', 'province', 'state_abbrev']
    zipcode_keywords = ['zip', 'zipcode', 'postal', 'zip_code', 'postcode', 'postal_code']
    
    mapping = {}
    original_columns = list(df.columns)
    
    # Find address column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in address_keywords):
            mapping['Address'] = original_columns[i]
            break
    
    # Find city column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in city_keywords):
            mapping['City'] = original_columns[i]
            break
    
    # Find state column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in state_keywords):
            mapping['State'] = original_columns[i]
            break
    
    # Find zipcode column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in zipcode_keywords):
            mapping['Zipcode'] = original_columns[i]
            break
    
    # Validate that all required columns are found
    required_keys = {'Address', 'City', 'State', 'Zipcode'}
    if not required_keys.issubset(mapping.keys()):
        missing = required_keys - mapping.keys()
        raise ValueError(f"Could not auto-detect columns: {missing}. Available columns: {original_columns}")
    
    return mapping

def load_file(file_path):
    """
    Load CSV or XLSX file and auto-detect columns.
    Returns DataFrame with standardized column names.
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.csv':
        df = pd.read_csv(file_path)
    elif file_ext in ['.xlsx', '.xls']:
        df = pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_ext}")
    
    # Auto-detect column mappings
    mapping = detect_column_mappings(df)
    
    # Rename columns to standard names
    df = df.rename(columns=mapping)
    
    # Keep only the columns we need
    df = df[['Address', 'City', 'State', 'Zipcode']].copy()
    
    # Clean data
    df = df.dropna(subset=['Address', 'Zipcode'])
    df['Address'] = df['Address'].astype(str).str.strip()
    df['City'] = df['City'].astype(str).str.strip()
    df['State'] = df['State'].astype(str).str.strip()
    df['Zipcode'] = df['Zipcode'].astype(str).str.strip()
    
    return df

# ============================================================================
# DATABASE SETUP
# ============================================================================

def setup_database():
    """Initialize SQLite database for caching representative data."""
    db_path = os.path.join(os.path.expanduser('~'), 'nyc_reps_cache.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS representatives (
            zipcode TEXT PRIMARY KEY,
            city_council TEXT,
            city_council_district TEXT,
            assembly_member TEXT,
            assembly_district TEXT,
            state_senator TEXT,
            state_senate_district TEXT,
            representative TEXT,
            congressional_district TEXT,
            borough_president TEXT,
            community_board TEXT,
            nyc_area BOOLEAN,
            timestamp DATETIME
        )
    ''')
    
    conn.commit()
    conn.close()
    return db_path

# ============================================================================
# NYC VALIDATION & DATA FETCHING
# ============================================================================

NYC_ZIPCODES = {
    # Manhattan
    '10001', '10002', '10003', '10004', '10005', '10006', '10007', '10008', '10009',
    '10010', '10011', '10012', '10013', '10014', '10016', '10017', '10018', '10019',
    '10020', '10021', '10022', '10023', '10024', '10025', '10026', '10027', '10028',
    '10029', '10030', '10031', '10032', '10033', '10034', '10035', '10036', '10037',
    '10038', '10039', '10040', '10041', '10043', '10044', '10045', '10046', '10047',
    '10048', '10049', '10050', '10051', '10055', '10060', '10069', '10075', '10081',
    '10082', '10087', '10090', '10095', '10096', '10097', '10098', '10099', '10101',
    '10102', '10103', '10104', '10105', '10106', '10107', '10108', '10109', '10110',
    '10111', '10112', '10113', '10114', '10115', '10116', '10117', '10118', '10119',
    '10120', '10121', '10122', '10123', '10124', '10125', '10126', '10128', '10151',
    '10152', '10153', '10154', '10155', '10156', '10157', '10158', '10159', '10160',
    '10161', '10162', '10163', '10164', '10165', '10166', '10167', '10168', '10169',
    '10170', '10171', '10172', '10173', '10174', '10175', '10176', '10177', '10178',
    '10179', '10185', '10199', '10203', '10210', '10211', '10212', '10213', '10214',
    '10215', '10216', '10217', '10218', '10219', '10220', '10221', '10222', '10223',
    '10224', '10225', '10226', '10227', '10228', '10280', '10281', '10282',
    # Bronx
    '10451', '10452', '10453', '10454', '10455', '10456', '10457', '10458', '10459',
    '10460', '10461', '10462', '10463', '10464', '10465', '10466', '10467', '10468',
    '10469', '10470', '10471', '10472', '10473', '10474', '10475',
    # Brooklyn
    '11201', '11202', '11203', '11204', '11205', '11206', '11207', '11208', '11209',
    '11210', '11211', '11212', '11213', '11214', '11215', '11216', '11217', '11218',
    '11219', '11220', '11221', '11222', '11223', '11224', '11225', '11226', '11227',
    '11228', '11229', '11230', '11231', '11232', '11233', '11234', '11235', '11236',
    '11237', '11238', '11239', '11241', '11242', '11243', '11244', '11245', '11246',
    '11247', '11248', '11249', '11251', '11252', '11256',
    # Queens
    '11354', '11355', '11356', '11357', '11358', '11359', '11360', '11361', '11362',
    '11363', '11364', '11365', '11366', '11367', '11368', '11369', '11370', '11371',
    '11372', '11373', '11374', '11375', '11376', '11377', '11378', '11379', '11380',
    '11381', '11382', '11383', '11384', '11385', '11386', '11387', '11388', '11389',
    '11390', '11391', '11392', '11393', '11394', '11395', '11396', '11397', '11398',
    '11399', '11401', '11402', '11403', '11404', '11405', '11406', '11407', '11408',
    '11409', '11410', '11411', '11412', '11413', '11414', '11415', '11416', '11417',
    '11418', '11419', '11420', '11421', '11422', '11423', '11424', '11425', '11426',
    '11427', '11428', '11429', '11430', '11431', '11432', '11433', '11434', '11435',
    '11436', '11691', '11692', '11693', '11694', '11695', '11696',
    # Staten Island
    '10301', '10302', '10303', '10304', '10305', '10306', '10307', '10308', '10309',
    '10310', '10311', '10312', '10313', '10314'
}

def validate_nyc_area(zipcode):
    """Check if zipcode is in NYC."""
    return str(zipcode).strip() in NYC_ZIPCODES

def fetch_representatives(address, city, state, zipcode):
    """Fetch representative data from MyGovNYC.org."""
    try:
        zipcode_str = str(zipcode).strip()
        
        if not validate_nyc_area(zipcode_str):
            return None
        
        url = f"https://www.mygov.nyc.gov/api/v1/representatives?zipcode={zipcode_str}"
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        data = response.json()
        representatives = {}
        
        for key in ['city_council_member', 'assembly_member', 'state_senator',
                    'representative', 'borough_president', 'community_board']:
            if key in data and data[key]:
                representatives[key] = data[key]
        
        return representatives if representatives else None
    except Exception as e:
        return None

def get_cached_or_fetch(zipcode, db_path):
    """Get representative data from cache or fetch from API."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM representatives WHERE zipcode = ?', (str(zipcode).strip(),))
    result = cursor.fetchone()
    
    if result:
        conn.close()
        return {
            'city_council': result[1],
            'city_council_district': result[2],
            'assembly_member': result[3],
            'assembly_district': result[4],
            'state_senator': result[5],
            'state_senate_district': result[6],
            'representative': result[7],
            'congressional_district': result[8],
            'borough_president': result[9],
            'community_board': result[10]
        }
    
    conn.close()
    return None

# ============================================================================
# WORKER THREAD FOR FILE PROCESSING
# ============================================================================

class FileProcessorThread(QThread):
    """Background thread for processing files with progress updates."""
    progress = pyqtSignal(int)  # Emit progress percentage
    status = pyqtSignal(str)  # Emit status message
    error = pyqtSignal(str)  # Emit error message
    finished_processing = pyqtSignal(pd.DataFrame)  # Emit processed data
    
    def __init__(self, file_path, db_path):
        super().__init__()
        self.file_path = file_path
        self.db_path = db_path
        self.results = []
    
    def run(self):
        try:
            # Step 1: Load file
            self.status.emit("📂 Loading file...")
            self.progress.emit(10)
            
            df = load_file(self.file_path)
            total_rows = len(df)
            
            if total_rows == 0:
                self.error.emit("File has no data rows!")
                return
            
            self.status.emit(f"✓ Loaded {total_rows} addresses. Processing...")
            self.progress.emit(20)
            
            # Step 2: Process each address
            results = []
            for idx, row in df.iterrows():
                progress_pct = 20 + int((idx / total_rows) * 70)  # 20-90%
                self.progress.emit(progress_pct)
                self.status.emit(f"🔍 Processing {idx + 1}/{total_rows}...")
                
                address = row['Address']
                city = row['City']
                state = row['State']
                zipcode = row['Zipcode']
                
                is_nyc = validate_nyc_area(zipcode)
                reps = get_cached_or_fetch(zipcode, self.db_path) if is_nyc else None
                
                result = {
                    'Address': address,
                    'City': city,
                    'State': state,
                    'Zipcode': zipcode,
                    'NYC Area': 'Yes' if is_nyc else 'No',
                    'City Council Member': reps.get('city_council', '') if reps else '',
                    'City Council District': reps.get('city_council_district', '') if reps else '',
                    'Assembly Member': reps.get('assembly_member', '') if reps else '',
                    'Assembly District': reps.get('assembly_district', '') if reps else '',
                    'State Senator': reps.get('state_senator', '') if reps else '',
                    'Senate District': reps.get('state_senate_district', '') if reps else '',
                    'Congressman': reps.get('representative', '') if reps else '',
                    'Congressional District': reps.get('congressional_district', '') if reps else '',
                    'Borough President': reps.get('borough_president', '') if reps else '',
                    'Community Board': reps.get('community_board', '') if reps else ''
                }
                results.append(result)
            
            self.progress.emit(95)
            self.status.emit("✓ Processing complete! Finalizing...")
            
            results_df = pd.DataFrame(results)
            self.progress.emit(100)
            self.status.emit("✓ Ready!")
            self.finished_processing.emit(results_df)
            
        except Exception as e:
            self.error.emit(f"Error: {str(e)}")

# ============================================================================
# MAIN APPLICATION
# ============================================================================

class NYCRepresentativesApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NYC Political Representatives Lookup")
        self.setGeometry(100, 100, 1400, 800)
        
        self.db_path = setup_database()
        self.results_df = None
        self.processor_thread = None
        
        # Get logo path
        self.logo_path = self.get_logo_path()
        
        self.init_ui()
    
    def get_logo_path(self):
        """Get the path to the logo file."""
        # Try multiple possible locations
        possible_paths = [
            os.path.join(os.path.dirname(__file__), 'QEDC-Full-Logo-Primary-Color.jpg'),
            os.path.join(os.path.expanduser('~'), 'QEDC-Full-Logo-Primary-Color.jpg'),
            'QEDC-Full-Logo-Primary-Color.jpg'
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        return None
    
    def init_ui(self):
        """Initialize the user interface."""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        
        # Top section: Logo + Header
        top_layout = QHBoxLayout()
        
        # Logo section (left)
        if self.logo_path:
            logo_widget = QLabel()
            pixmap = QPixmap(self.logo_path)
            pixmap = pixmap.scaledToHeight(100, Qt.SmoothTransformation)
            logo_widget.setPixmap(pixmap)
            logo_widget.setAlignment(Qt.AlignCenter)
            top_layout.addWidget(logo_widget)
        
        # Header section (right)
        header_layout = QVBoxLayout()
        header = QLabel("NYC Political Representatives Lookup")
        header.setStyleSheet("font-size: 20px; font-weight: bold; color: #0066cc;")
        header_layout.addWidget(header)
        
        subtitle = QLabel("Upload your address CSV or XLSX and retrieve political representatives")
        subtitle.setStyleSheet("color: #666666; font-size: 12px;")
        header_layout.addWidget(subtitle)
        
        top_layout.addLayout(header_layout, 1)
        main_layout.addLayout(top_layout)
        
        # Separator
        separator = QLabel("─" * 100)
        separator.setStyleSheet("color: #cccccc;")
        main_layout.addWidget(separator)
        
        # File upload section
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file chosen")
        self.file_label.setStyleSheet("color: #999999;")
        file_layout.addWidget(self.file_label)
        
        self.choose_file_btn = QPushButton("Choose CSV/XLSX File")
        self.choose_file_btn.clicked.connect(self.choose_file)
        self.choose_file_btn.setStyleSheet("""
            QPushButton {
                background-color: #0066cc;
                color: white;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0052a3;
            }
        """)
        file_layout.addWidget(self.choose_file_btn)
        
        main_layout.addLayout(file_layout)
        
        # Process button
        self.process_btn = QPushButton("Process Addresses")
        self.process_btn.clicked.connect(self.process_addresses)
        self.process_btn.setEnabled(False)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #008000;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #006600;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        main_layout.addWidget(self.process_btn)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #0066cc;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #0066cc;
            }
        """)
        main_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #0066cc; font-weight: bold;")
        main_layout.addWidget(self.status_label)
        
        # Tabs for results and analytics
        self.tabs = QTabWidget()
        
        # Results tab
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(14)
        self.results_table.setHorizontalHeaderLabels([
            'Address', 'City', 'State', 'Zipcode', 'NYC Area',
            'City Council', 'CC District', 'Assembly Member', 'Assembly District',
            'State Senator', 'Senate District', 'Congressman', 'Congress District',
            'Borough President'
        ])
        self.tabs.addTab(self.results_table, "Results")
        
        # Analytics tab
        self.analytics_label = QLabel("Analytics will appear here")
        self.tabs.addTab(self.analytics_label, "Analytics")
        
        main_layout.addWidget(self.tabs)
        
        # Export button
        self.export_btn = QPushButton("Export to Excel CSV")
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff6600;
                color: white;
                padding: 8px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e55c00;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        main_layout.addWidget(self.export_btn)
        
        # Support section - FOOTER
        footer_layout = QHBoxLayout()
        footer_separator = QLabel("─" * 100)
        footer_separator.setStyleSheet("color: #cccccc;")
        
        support_layout = QVBoxLayout()
        support_label = QLabel("❓ Questions or Support?")
        support_label.setStyleSheet("font-weight: bold; color: #333333;")
        support_layout.addWidget(support_label)
        
        self.support_btn = QPushButton("📧 Contact: Victor Prado (Vprado@Queensny.org)")
        self.support_btn.clicked.connect(self.open_support)
        self.support_btn.setCursor(Qt.PointingHandCursor)
        self.support_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffa500;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
                text-decoration: underline;
            }
            QPushButton:hover {
                background-color: #ff8c00;
            }
        """)
        support_layout.addWidget(self.support_btn)
        
        footer_layout.addLayout(support_layout)
        main_layout.addLayout(footer_layout)
        
        main_widget.setLayout(main_layout)
    
    def open_support(self):
        """Open support contact options."""
        msg = QMessageBox(self)
        msg.setWindowTitle("Support & Contact")
        msg.setIcon(QMessageBox.Information)
        msg.setText("For Questions or Support\n\n")
        msg.setInformativeText(
            "Name: Victor Prado\n"
            "Email: Vprado@Queensny.org\n\n"
            "Please include your issue description in the email."
        )
        
        # Add copy email button
        copy_btn = msg.addButton("Copy Email Address", QMessageBox.ActionRole)
        close_btn = msg.addButton("Close", QMessageBox.RejectRole)
        
        msg.exec_()
        
        if msg.clickedButton() == copy_btn:
            import pyperclip
            try:
                # Try to copy to clipboard
                os.system('echo Vprado@Queensny.org | clip')
                QMessageBox.information(self, "Copied", "Email address copied to clipboard!")
            except:
                QMessageBox.information(self, "Email", "Vprado@Queensny.org")
    
    def choose_file(self):
        """Open file dialog to choose CSV or XLSX."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Choose Address File", 
            "", 
            "CSV and Excel Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(f"✓ Selected: {os.path.basename(file_path)}")
            self.file_label.setStyleSheet("color: #008000;")
            self.process_btn.setEnabled(True)
    
    def process_addresses(self):
        """Process the selected file."""
        if not hasattr(self, 'file_path'):
            QMessageBox.warning(self, "Error", "Please select a file first")
            return
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Starting...")
        self.process_btn.setEnabled(False)
        self.choose_file_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        
        # Start worker thread
        self.processor_thread = FileProcessorThread(self.file_path, self.db_path)
        self.processor_thread.progress.connect(self.update_progress)
        self.processor_thread.status.connect(self.update_status)
        self.processor_thread.error.connect(self.show_error)
        self.processor_thread.finished_processing.connect(self.display_results)
        self.processor_thread.start()
    
    def update_progress(self, value):
        """Update progress bar."""
        self.progress_bar.setValue(value)
    
    def update_status(self, message):
        """Update status label."""
        self.status_label.setText(message)
    
    def show_error(self, error_msg):
        """Show error message."""
        QMessageBox.critical(self, "Error", error_msg)
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.choose_file_btn.setEnabled(True)
    
    def display_results(self, results_df):
        """Display results in table."""
        self.results_df = results_df
        self.results_table.setRowCount(len(results_df))
        
        for row_idx, row in results_df.iterrows():
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.results_table.setItem(row_idx, col_idx, item)
        
        self.results_table.resizeColumnsToContents()
        self.export_btn.setEnabled(True)
        
        # Update analytics
        nyc_count = len(results_df[results_df['NYC Area'] == 'Yes'])
        non_nyc_count = len(results_df[results_df['NYC Area'] == 'No'])
        analytics_text = (
            f"<b>Results Summary</b><br>"
            f"Total Addresses: {len(results_df)}<br>"
            f"NYC Addresses: {nyc_count}<br>"
            f"Non-NYC Addresses: {non_nyc_count}"
        )
        self.analytics_label.setText(analytics_text)
        self.analytics_label.setStyleSheet("font-size: 12px; padding: 20px;")
        
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.choose_file_btn.setEnabled(True)
        self.status_label.setText("✓ Processing complete!")
    
    def export_results(self):
        """Export results to CSV."""
        if self.results_df is None:
            QMessageBox.warning(self, "Error", "No results to export")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Save Results", 
            "", 
            "CSV Files (*.csv)"
        )
        
        if file_path:
            self.results_df.to_csv(file_path, index=False)
            QMessageBox.information(self, "Success", f"Results saved to {file_path}")

# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = NYCRepresentativesApp()
    window.show()
    sys.exit(app.exec_())
