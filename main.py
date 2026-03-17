import sys
import os
import csv
import json
import sqlite3
import requests
import threading
import time
from datetime import datetime
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, QTabWidget, QMessageBox)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QIcon, QFont
import pandas as pd
from bs4 import BeautifulSoup

# ============================================================================
# NYC ZIPCODE VALIDATION
# ============================================================================
NYC_ZIPCODES = {
    # Manhattan
    '10001', '10002', '10003', '10004', '10005', '10006', '10007', '10008', '10009',
    '10010', '10011', '10012', '10013', '10014', '10016', '10017', '10018', '10019',
    '10020', '10021', '10022', '10023', '10024', '10025', '10026', '10027', '10028',
    '10029', '10030', '10031', '10032', '10033', '10034', '10035', '10036', '10037',
    '10038', '10039', '10040',
    # Brooklyn
    '11201', '11202', '11203', '11204', '11205', '11206', '11207', '11208', '11209',
    '11210', '11211', '11212', '11213', '11214', '11215', '11216', '11217', '11218',
    '11219', '11220', '11221', '11222', '11223', '11224', '11225', '11226', '11228',
    '11229', '11230', '11231', '11232', '11233', '11234', '11235', '11236', '11237',
    '11238', '11239',
    # Queens
    '11001', '11002', '11003', '11004', '11005', '11040', '11041', '11042', '11043',
    '11354', '11355', '11356', '11357', '11358', '11359', '11360', '11361', '11362',
    '11363', '11364', '11365', '11366', '11367', '11368', '11369', '11370', '11371',
    '11372', '11373', '11374', '11375', '11376', '11377', '11378', '11379', '11380',
    '11381', '11382', '11383', '11384', '11385', '11386', '11387', '11388', '11389',
    '11390', '11391', '11392', '11393', '11394', '11395', '11411', '11412', '11413',
    '11414', '11415', '11416', '11417', '11418', '11419', '11420', '11421', '11422',
    '11423', '11424', '11425', '11426', '11427', '11428', '11429', '11430', '11431',
    '11432', '11433', '11434', '11435', '11436',
    # Bronx
    '10451', '10452', '10453', '10454', '10455', '10456', '10457', '10458', '10459',
    '10460', '10461', '10462', '10463', '10464', '10465', '10466', '10467', '10468',
    '10469', '10470', '10471', '10472', '10473', '10474',
    # Staten Island
    '10301', '10302', '10303', '10304', '10305', '10306', '10307', '10308', '10309',
    '10310', '10311', '10312', '10313', '10314'
}

# ============================================================================
# MYGOVNYC FETCHER THREAD
# ============================================================================
class MyGovNYCFetcher(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, csv_file_path):
        super().__init__()
        self.csv_file_path = csv_file_path
        self.db_path = self.get_db_path()
        self.init_db()

    def get_db_path(self):
        app_data = Path.home() / 'AppData' / 'Local' / 'NYCRepLookup'
        app_data.mkdir(parents=True, exist_ok=True)
        return str(app_data / 'representatives.db')

    def init_db(self):
        try:
            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()
            c.execute('''CREATE TABLE IF NOT EXISTS representatives (
                zipcode TEXT PRIMARY KEY,
                council_members TEXT,
                council_districts TEXT,
                assembly_members TEXT,
                assembly_districts TEXT,
                senators TEXT,
                senate_districts TEXT,
                house_members TEXT,
                house_districts TEXT,
                borough_president TEXT,
                community_boards TEXT,
                fetched_date TEXT
            )''')
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Database init error: {e}")

    def is_nyc_address(self, address, city, zipcode):
        """Check if address is in NYC"""
        if zipcode.strip() in NYC_ZIPCODES:
            return True
        
        nyc_cities = {'new york', 'brooklyn', 'queens', 'bronx', 'staten island', 'manhattan'}
        if city.lower().strip() in nyc_cities:
            return True
        
        return False

    def fetch_mygovnyc(self, zipcode):
        """Fetch representative data from MyGovNYC"""
        try:
            url = f"https://www.mygovnyc.org/?addr={zipcode}"
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            reps = self.parse_mygovnyc_page(soup, zipcode)
            
            if reps:
                self.cache_representatives(zipcode, reps)
            
            return reps
        except Exception as e:
            print(f"Fetch error for {zipcode}: {e}")
            return None

    def parse_mygovnyc_page(self, soup, zipcode):
        """Parse MyGovNYC response"""
        reps = {
            'council_members': [],
            'council_districts': [],
            'assembly_members': [],
            'assembly_districts': [],
            'senators': [],
            'senate_districts': [],
            'house_members': [],
            'house_districts': [],
            'borough_president': 'N/A',
            'community_boards': []
        }

        try:
            # Parse representative cards
            cards = soup.find_all('div', class_=['representative-card', 'card'])
            
            for card in cards:
                title = card.find(['h3', 'h4', 'h5'])
                name_link = card.find('a')
                
                if title and name_link:
                    title_text = title.get_text().strip()
                    name = name_link.get_text().strip()
                    
                    if 'City Council' in title_text or 'Council Member' in title_text:
                        district = card.find('span', class_=['district', 'district-number'])
                        district_text = district.get_text().strip() if district else f'District {zipcode}'
                        if name and name not in reps['council_members']:
                            reps['council_members'].append(name)
                            reps['council_districts'].append(district_text)
                    
                    elif 'Assembly' in title_text:
                        district = card.find('span', class_=['district', 'district-number'])
                        district_text = district.get_text().strip() if district else 'N/A'
                        if name and name not in reps['assembly_members']:
                            reps['assembly_members'].append(name)
                            reps['assembly_districts'].append(district_text)
                    
                    elif 'Senator' in title_text or 'Senate' in title_text:
                        district = card.find('span', class_=['district', 'district-number'])
                        district_text = district.get_text().strip() if district else 'N/A'
                        if name and name not in reps['senators']:
                            reps['senators'].append(name)
                            reps['senate_districts'].append(district_text)
                    
                    elif 'Congress' in title_text or 'Representative' in title_text:
                        district = card.find('span', class_=['district', 'district-number'])
                        district_text = district.get_text().strip() if district else 'N/A'
                        if name and name not in reps['house_members']:
                            reps['house_members'].append(name)
                            reps['house_districts'].append(district_text)
                    
                    elif 'Borough President' in title_text:
                        reps['borough_president'] = name
                    
                    elif 'Community Board' in title_text:
                        if name not in reps['community_boards']:
                            reps['community_boards'].append(name)
        
        except Exception as e:
            print(f"Parse error: {e}")
        
        return reps

    def cache_representatives(self, zipcode, reps):
        """Cache data locally"""
        try:
            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()
            c.execute('''INSERT OR REPLACE INTO representatives 
                         (zipcode, council_members, council_districts, assembly_members, 
                          assembly_districts, senators, senate_districts, house_members, 
                          house_districts, borough_president, community_boards, fetched_date)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                      (zipcode,
                       json.dumps(reps['council_members']),
                       json.dumps(reps['council_districts']),
                       json.dumps(reps['assembly_members']),
                       json.dumps(reps['assembly_districts']),
                       json.dumps(reps['senators']),
                       json.dumps(reps['senate_districts']),
                       json.dumps(reps['house_members']),
                       json.dumps(reps['house_districts']),
                       reps['borough_president'],
                       json.dumps(reps['community_boards']),
                       datetime.now().isoformat()))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Cache error: {e}")

    def get_cached_representatives(self, zipcode):
        """Retrieve cached data"""
        try:
            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()
            c.execute('SELECT * FROM representatives WHERE zipcode = ?', (zipcode,))
            row = c.fetchone()
            conn.close()
            
            if row:
                return {
                    'council_members': json.loads(row[1]),
                    'council_districts': json.loads(row[2]),
                    'assembly_members': json.loads(row[3]),
                    'assembly_districts': json.loads(row[4]),
                    'senators': json.loads(row[5]),
                    'senate_districts': json.loads(row[6]),
                    'house_members': json.loads(row[7]),
                    'house_districts': json.loads(row[8]),
                    'borough_president': row[9],
                    'community_boards': json.loads(row[10])
                }
        except Exception as e:
            print(f"Cache retrieval error: {e}")
        
        return None

    def get_empty_reps(self):
        return {
            'council_members': [],
            'council_districts': [],
            'assembly_members': [],
            'assembly_districts': [],
            'senators': [],
            'senate_districts': [],
            'house_members': [],
            'house_districts': [],
            'borough_president': 'N/A',
            'community_boards': []
        }

    def run(self):
        try:
            df = pd.read_csv(self.csv_file_path)
            total_rows = len(df)
            
            nyc_results = []
            non_nyc_results = []
            
            for idx, row in df.iterrows():
                self.progress.emit(int((idx + 1) / total_rows * 100))
                self.status.emit(f"Processing row {idx + 1} of {total_rows}...")
                
                address = str(row.get('Address', '')).strip()
                city = str(row.get('City', '')).strip()
                state = str(row.get('State', '')).strip()
                zipcode = str(row.get('Zipcode', '')).strip()
                
                is_nyc = self.is_nyc_address(address, city, zipcode)
                
                # Try cache first
                reps = self.get_cached_representatives(zipcode)
                if reps is None:
                    # Fetch from MyGovNYC
                    reps = self.fetch_mygovnyc(zipcode)
                    if reps is None:
                        reps = self.get_empty_reps()
                
                result_row = dict(row)
                result_row['is_nyc'] = is_nyc
                result_row['reps'] = reps
                
                if is_nyc:
                    nyc_results.append(result_row)
                else:
                    non_nyc_results.append(result_row)
                
                time.sleep(0.3)  # Rate limiting
            
            self.finished.emit({
                'nyc_results': nyc_results,
                'non_nyc_results': non_nyc_results,
                'total': total_rows
            })
        
        except Exception as e:
            self.error.emit(f"Processing error: {str(e)}")

# ============================================================================
# MAIN APPLICATION WINDOW
# ============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('NYC Political Representatives Lookup')
        self.setGeometry(50, 50, 1500, 950)
        
        self.processed_data = None
        self.fetcher_thread = None
        
        # Create central widget and layout
        central_widget = QWidget()
        layout = QVBoxLayout()
        
        # Create tabs
        self.tabs = QTabWidget()
        
        # Dashboard tab (web view)
        self.dashboard_view = QWebEngineView()
        self.tabs.addTab(self.dashboard_view, "Dashboard")
        
        # Analytics tab
        self.analytics_view = QWebEngineView()
        self.tabs.addTab(self.analytics_view, "Analytics")
        
        layout.addWidget(self.tabs)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        
        # Load initial dashboard
        self.load_dashboard()

    def load_dashboard(self):
        """Load the main dashboard HTML"""
        html_content = self.get_dashboard_html()
        self.dashboard_view.setHtml(html_content)

    def load_analytics(self, data):
        """Load analytics dashboard with summaries"""
        html_content = self.get_analytics_html(data)
        self.analytics_view.setHtml(html_content)

    def get_dashboard_html(self):
        """Return main dashboard HTML"""
        return """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>NYC Political Representatives Lookup</title>
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    background: #f8f9fa;
                    color: #212529;
                }
                .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
                header {
                    background: white;
                    padding: 30px;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    margin-bottom: 30px;
                }
                h1 { color: #20809e; margin-bottom: 10px; font-size: 28px; }
                .subtitle { color: #6c757d; font-size: 14px; }
                .btn {
                    background: #20809e;
                    color: white;
                    border: none;
                    padding: 12px 32px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-size: 16px;
                    font-weight: 500;
                    transition: all 0.3s;
                }
                .btn:hover { background: #1a6882; }
                .upload-section {
                    background: white;
                    padding: 30px;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    margin-bottom: 20px;
                }
                .file-input-wrapper {
                    display: flex;
                    gap: 15px;
                    align-items: center;
                    margin-bottom: 20px;
                }
                .file-input-label {
                    background: #20809e;
                    color: white;
                    padding: 12px 24px;
                    border-radius: 6px;
                    cursor: pointer;
                    font-weight: 500;
                }
                input[type="file"] { display: none; }
                .file-name { color: #6c757d; font-size: 14px; }
                .table-container {
                    background: white;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    overflow-x: auto;
                    max-height: 600px;
                    overflow-y: auto;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    font-size: 14px;
                }
                th {
                    background: #f8f9fa;
                    padding: 12px;
                    text-align: left;
                    font-weight: 600;
                    position: sticky;
                    top: 0;
                    border-bottom: 2px solid #dee2e6;
                }
                td { padding: 12px; border-bottom: 1px solid #f1f3f5; }
                tr:hover { background: #f8f9fa; }
                .alert {
                    padding: 15px 20px;
                    border-radius: 6px;
                    margin-bottom: 20px;
                    border-left: 4px solid;
                }
                .alert-info {
                    background: #d1ecf1;
                    color: #0c5460;
                    border-color: #17a2b8;
                }
                .alert-warning {
                    background: #fff3cd;
                    color: #856404;
                    border-color: #ffc107;
                }
                .out-of-nyc { background: #fff3cd !important; }
                .status { color: #6c757d; font-size: 14px; margin-top: 10px; }
            </style>
        </head>
        <body>
            <div class="container">
                <header>
                    <h1>NYC Political Representatives Lookup</h1>
                    <p class="subtitle">Upload your address CSV and retrieve political representatives</p>
                </header>

                <div class="upload-section">
                    <h2 style="margin-bottom: 20px; font-size: 18px;">Upload Address File</h2>
                    <div class="file-input-wrapper">
                        <label for="csvFile" class="file-input-label">Choose CSV File</label>
                        <input type="file" id="csvFile" accept=".csv">
                        <span class="file-name" id="fileName">No file chosen</span>
                    </div>
                    <button class="btn" id="processBtn" disabled>Process Addresses</button>
                    <div class="status" id="status"></div>
                </div>

                <div class="alert alert-info" style="display:none;" id="infoAlert">
                    <strong>Note:</strong> Processing will fetch live data from MyGovNYC.org on first run (requires internet).
                </div>

                <div class="table-container" style="display:none;" id="resultsContainer">
                    <table id="resultsTable">
                        <thead></thead>
                        <tbody id="resultsBody"></tbody>
                    </table>
                </div>
            </div>
        </body>
        </html>
        """

    def get_analytics_html(self, data):
        """Return analytics dashboard HTML"""
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Analytics Dashboard</title>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body { font-family: Arial, sans-serif; background: #f8f9fa; padding: 20px; }
                .container { max-width: 1400px; margin: 0 auto; }
                h1 { color: #20809e; margin-bottom: 30px; }
                .grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 20px; }
                .card {
                    background: white;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                .stat-box { text-align: center; padding: 20px; }
                .stat-value { font-size: 36px; font-weight: bold; color: #20809e; }
                .stat-label { color: #6c757d; margin-top: 5px; }
                canvas { max-height: 300px; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Analytics Dashboard</h1>
                <div class="grid">
                    <div class="card">
                        <div class="stat-box">
                            <div class="stat-value" id="nycCount">0</div>
                            <div class="stat-label">NYC Records</div>
                        </div>
                    </div>
                    <div class="card">
                        <div class="stat-box">
                            <div class="stat-value" id="nonNycCount">0</div>
                            <div class="stat-label">Non-NYC Records</div>
                        </div>
                    </div>
                </div>
                <div class="grid">
                    <div class="card">
                        <h3>City Council Districts</h3>
                        <canvas id="councilChart"></canvas>
                    </div>
                    <div class="card">
                        <h3>State Assembly Districts</h3>
                        <canvas id="assemblyChart"></canvas>
                    </div>
                    <div class="card">
                        <h3>State Senate Districts</h3>
                        <canvas id="senateChart"></canvas>
                    </div>
                    <div class="card">
                        <h3>US House Districts</h3>
                        <canvas id="houseChart"></canvas>
                    </div>
                </div>
            </div>
            <script>
                // Analytics data would be populated here
                console.log('Analytics dashboard loaded');
            </script>
        </body>
        </html>
        """


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
