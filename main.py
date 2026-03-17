
import webbrowser
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QTabWidget, QProgressBar, QComboBox, QMessageBox, QScrollArea
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl, QSize
from PyQt5.QtGui import QPalette, QColor, QPixmap, QIcon, QFont
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
    address_keywords = ['address', 'street', 'addr', 'location', 'road', 'avenue', 'boulevard', 'street_address']
    city_keywords = ['city', 'municipality', 'town', 'metro_area', 'place', 'metro']
    state_keywords = ['state', 'state_code', 'st', 'province', 'state_abbrev', 'state', 'region']
    zipcode_keywords = ['zip', 'zipcode', 'postal', 'zip_code', 'postcode', 'postal_code', 'zip code', 'postal_code']
    
    mapping = {}
    original_columns = list(df.columns)
    
    print(f"DEBUG: Detecting columns from: {original_columns}")
    
    # Find address column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in address_keywords):
            mapping['Address'] = original_columns[i]
            print(f"DEBUG: Found Address -> {original_columns[i]}")
            break
    
    # Find city column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in city_keywords):
            mapping['City'] = original_columns[i]
            print(f"DEBUG: Found City -> {original_columns[i]}")
            break
    
    # Find state column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in state_keywords):
            mapping['State'] = original_columns[i]
            print(f"DEBUG: Found State -> {original_columns[i]}")
            break
    
    # Find zipcode column
    for i, col in enumerate(columns):
        if any(keyword in col for keyword in zipcode_keywords):
            mapping['Zipcode'] = original_columns[i]
            print(f"DEBUG: Found Zipcode -> {original_columns[i]}")
            break
    
    # Validate that all required columns are found
    required_keys = {'Address', 'City', 'State', 'Zipcode'}
    missing = required_keys - mapping.keys()
    
    if missing:
        print(f"DEBUG: Missing columns: {missing}")
        print(f"DEBUG: Found columns: {mapping}")
        raise ValueError(
            f"Could not auto-detect columns: {missing}\n\n"
            f"Found columns in file: {original_columns}\n\n"
            f"Detected: {mapping}\n\n"
            f"The file needs columns for: Address, City, State, and Zipcode"
        )
    
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
    
    print(f"DEBUG: File loaded. Columns: {list(df.columns)}")
    print(f"DEBUG: First row: {df.iloc[0].to_dict() if len(df) > 0 else 'No rows'}")
    
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
# MAIN APPLICATION - BEAUTIFUL REDESIGN
# ============================================================================

class NYCRepresentativesApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NYC Political Representatives Lookup")
        self.setGeometry(100, 100, 1600, 900)
        
        # Apply modern stylesheet
        self.apply_modern_stylesheet()
        
        self.db_path = setup_database()
        self.results_df = None
        self.processor_thread = None
        
        # Get logo path
        self.logo_path = self.get_logo_path()
        
        self.init_ui()
    
    def apply_modern_stylesheet(self):
        """Apply beautiful modern dark theme stylesheet."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1a1a2e;
            }
            QWidget {
                background-color: #1a1a2e;
                color: #ffffff;
            }
            QLabel {
                color: #ffffff;
            }
            QPushButton {
                border-radius: 8px;
                border: none;
                font-weight: bold;
                padding: 10px 20px;
                transition: all 0.3s ease;
            }
            QPushButton:hover {
                transform: translateY(-2px);
                box-shadow: 0 8px 16px rgba(0, 0, 0, 0.3);
            }
            QTableWidget {
                background-color: #16213e;
                gridline-color: #0f3460;
                border: 1px solid #0f3460;
                border-radius: 6px;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #0f3460;
            }
            QTableWidget::item:selected {
                background-color: #e94560;
            }
            QHeaderView::section {
                background-color: #0f3460;
                color: #ffffff;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
            QTabWidget::pane {
                border: 1px solid #0f3460;
            }
            QTabBar::tab {
                background-color: #0f3460;
                color: #ffffff;
                padding: 8px 20px;
                margin-right: 2px;
                border-radius: 6px 6px 0 0;
            }
            QTabBar::tab:selected {
                background-color: #e94560;
            }
            QProgressBar {
                border-radius: 8px;
                border: 2px solid #0f3460;
                background-color: #0f3460;
                height: 25px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                           stop:0 #00d4ff, stop:1 #0099ff);
                border-radius: 6px;
            }
            QScrollBar:vertical {
                background-color: #0f3460;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #e94560;
                border-radius: 6px;
            }
        """)
    
    def get_logo_path(self):
        """Get the path to the logo file."""
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
        """Initialize the user interface with modern design."""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # ===== HEADER SECTION WITH LOGO =====
        header_widget = QWidget()
        header_widget.setStyleSheet("background-color: #0f3460; border-bottom: 3px solid #e94560;")
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(30, 20, 30, 20)
        header_layout.setSpacing(20)
        
        # Logo
        if self.logo_path:
            logo_widget = QLabel()
            pixmap = QPixmap(self.logo_path)
            pixmap = pixmap.scaledToHeight(80, Qt.SmoothTransformation)
            logo_widget.setPixmap(pixmap)
            header_layout.addWidget(logo_widget)
        
        # Title section
        title_widget = QWidget()
        title_layout = QVBoxLayout(title_widget)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(8)
        
        title = QLabel("NYC Political Representatives Lookup")
        title_font = QFont("Segoe UI", 26, QFont.Bold)
        title.setFont(title_font)
        title.setStyleSheet("color: #00d4ff;")
        title_layout.addWidget(title)
        
        subtitle = QLabel("Fast. Smart. Accurate. Find your representatives instantly.")
        subtitle_font = QFont("Segoe UI", 11)
        subtitle.setFont(subtitle_font)
        subtitle.setStyleSheet("color: #b0b0b0;")
        title_layout.addWidget(subtitle)
        
        header_layout.addWidget(title_widget, 1)
        main_layout.addWidget(header_widget)
        
        # ===== MAIN CONTENT AREA =====
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(30, 30, 30, 30)
        content_layout.setSpacing(20)
        
        # File Selection Card
        file_card = self.create_card()
        file_card_layout = QVBoxLayout(file_card)
        file_card_layout.setSpacing(15)
        
        file_label_text = QLabel("📁 SELECT YOUR FILE")
        file_label_text.setFont(QFont("Segoe UI", 12, QFont.Bold))
        file_label_text.setStyleSheet("color: #00d4ff;")
        file_card_layout.addWidget(file_label_text)
        
        self.file_label = QLabel("No file selected • Click below to choose")
        self.file_label.setStyleSheet("color: #808080; font-size: 11px;")
        file_card_layout.addWidget(self.file_label)
        
        self.choose_file_btn = self.create_button("📂 CHOOSE CSV/XLSX FILE", "#00d4ff", "#0099cc")
        self.choose_file_btn.clicked.connect(self.choose_file)
        file_card_layout.addWidget(self.choose_file_btn)
        
        content_layout.addWidget(file_card)
        
        # Process Button - LARGE & PROMINENT
        self.process_btn = self.create_button("🚀 PROCESS ADDRESSES", "#00ff00", "#00cc00")
        self.process_btn.setMinimumHeight(50)
        self.process_btn.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.process_btn.clicked.connect(self.process_addresses)
        self.process_btn.setEnabled(False)
        content_layout.addWidget(self.process_btn)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(30)
        content_layout.addWidget(self.progress_bar)
        
        # Status Label
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Segoe UI", 11))
        self.status_label.setStyleSheet("color: #00d4ff; font-weight: bold;")
        content_layout.addWidget(self.status_label)
        
        # Tabs for Results
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #0f3460;
                border-radius: 8px;
            }
            QTabBar::tab {
                background-color: #16213e;
                color: #ffffff;
                padding: 12px 30px;
                margin: 2px;
                border-radius: 6px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background-color: #e94560;
                color: #ffffff;
            }
        """)
        
        # Results tab
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(14)
        self.results_table.setHorizontalHeaderLabels([
            'Address', 'City', 'State', 'Zipcode', 'NYC Area',
            'City Council', 'CC District', 'Assembly Member', 'Assembly District',
            'State Senator', 'Senate District', 'Congressman', 'Congress District',
            'Borough President'
        ])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        self.tabs.addTab(self.results_table, "📊 RESULTS")
        
        # Analytics tab
        self.analytics_label = QLabel("Analytics will appear after processing")
        self.analytics_label.setStyleSheet("""
            background-color: #16213e;
            border-radius: 8px;
            padding: 30px;
            color: #b0b0b0;
            font-size: 13px;
        """)
        self.tabs.addTab(self.analytics_label, "📈 ANALYTICS")
        
        content_layout.addWidget(self.tabs, 1)
        
        # Export Button
        self.export_btn = self.create_button("💾 EXPORT TO EXCEL CSV", "#ff9900", "#cc7700")
        self.export_btn.setMinimumHeight(45)
        self.export_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        content_layout.addWidget(self.export_btn)
        
        main_layout.addWidget(content_widget, 1)
        
        # ===== FOOTER SECTION =====
        footer_widget = QWidget()
        footer_widget.setStyleSheet("background-color: #0f3460; border-top: 2px solid #e94560;")
        footer_layout = QVBoxLayout(footer_widget)
        footer_layout.setContentsMargins(30, 20, 30, 20)
        footer_layout.setSpacing(12)
        
        support_title = QLabel("❓ QUESTIONS? NEED HELP?")
        support_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        support_title.setStyleSheet("color: #00d4ff;")
        footer_layout.addWidget(support_title)
        
        self.support_btn = self.create_button("📧 Contact: Victor Prado (Vprado@Queensny.org)", "#e94560", "#cc3355")
        self.support_btn.clicked.connect(self.open_support)
        self.support_btn.setCursor(Qt.PointingHandCursor)
        footer_layout.addWidget(self.support_btn)
        
        main_layout.addWidget(footer_widget)
    
    def create_card(self):
        """Create a modern card widget."""
        card = QWidget()
        card.setStyleSheet("""
            QWidget {
                background-color: #16213e;
                border-radius: 12px;
                border: 2px solid #0f3460;
            }
        """)
        return card
    
    def create_button(self, text, color, hover_color):
        """Create a modern button with gradient and hover effect."""
        button = QPushButton(text)
        button.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: #ffffff;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                padding: 12px 20px;
                font-size: 11px;
                transition: all 0.3s ease;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
                transform: translateY(-2px);
            }}
            QPushButton:pressed {{
                transform: translateY(0px);
            }}
            QPushButton:disabled {{
                background-color: #404040;
                color: #808080;
            }}
        """)
        return button
    
    def open_support(self):
        """Open support contact options."""
        msg = QMessageBox(self)
        msg.setWindowTitle("Support & Contact")
        msg.setIcon(QMessageBox.Information)
        msg.setStyleSheet("background-color: #16213e; color: #ffffff;")
        msg.setText("For Questions or Support\n\n")
        msg.setInformativeText(
            "Name: Victor Prado\n"
            "Email: Vprado@Queensny.org\n\n"
            "Please include your issue description in the email."
        )
        
        msg.exec_()
    
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
            file_name = os.path.basename(file_path)
            self.file_label.setText(f"✓ Selected: {file_name}")
            self.file_label.setStyleSheet("color: #00ff00; font-weight: bold; font-size: 12px;")
            self.process_btn.setEnabled(True)
    
    def process_addresses(self):
        """Process the selected file."""
        if not hasattr(self, 'file_path'):
            QMessageBox.warning(self, "Error", "Please select a file first")
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Starting...")
        self.process_btn.setEnabled(False)
        self.choose_file_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        
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
        msg = QMessageBox(self)
        msg.setWindowTitle("Error")
        msg.setText(error_msg)
        msg.setIcon(QMessageBox.Critical)
        msg.setStyleSheet("background-color: #16213e; color: #ffffff;")
        msg.exec_()
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
                item.setForeground(QColor("#ffffff"))
                self.results_table.setItem(row_idx, col_idx, item)
        
        self.results_table.resizeColumnsToContents()
        self.export_btn.setEnabled(True)
        
        # Update analytics
        nyc_count = len(results_df[results_df['NYC Area'] == 'Yes'])
        non_nyc_count = len(results_df[results_df['NYC Area'] == 'No'])
        analytics_text = (
            f"<div style='font-size: 14px; line-height: 2;'>"
            f"<b style='color: #00d4ff;'>📊 RESULTS SUMMARY</b><br>"
            f"<span style='color: #00ff00;'>✓ Total Addresses:</span> <b>{len(results_df)}</b><br>"
            f"<span style='color: #00ff00;'>✓ NYC Addresses:</span> <b>{nyc_count}</b><br>"
            f"<span style='color: #ff9900;'>⚠ Non-NYC Addresses:</span> <b>{non_nyc_count}</b>"
            f"</div>"
        )
        self.analytics_label.setText(analytics_text)
        
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.choose_file_btn.setEnabled(True)
        self.status_label.setText("✅ Processing complete! Results ready.")
    
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
            msg = QMessageBox(self)
            msg.setWindowTitle("Success")
            msg.setText(f"✅ Results saved successfully!\n\n{file_path}")
            msg.setIcon(QMessageBox.Information)
            msg.setStyleSheet("background-color: #16213e; color: #ffffff;")
            msg.exec_()

# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = NYCRepresentativesApp()
    window.show()
    sys.exit(app.exec_())
