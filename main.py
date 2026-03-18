import sys
import os
from pathlib import Path
import pandas as pd
import requests
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QLabel,
                              QMessageBox, QHeaderView, QProgressBar)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QIcon, QPixmap, QColor
from PyQt5.QtCore import QThread, pyqtSignal
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json

# ============================================
# BEAUTIFUL DARK THEME UI WITH COLORED BUTTONS
# ============================================

class WorkerThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, file_path, app_window):
        super().__init__()
        self.file_path = file_path
        self.app = app_window
        
    def run(self):
        try:
            self.app.process_file_internal(self.file_path)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class NYC_RepresentativesLookup(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.data = None
        self.worker_thread = None
        
    def init_ui(self):
        # Window settings
        self.setWindowTitle('NYC Representatives Lookup')
        self.setGeometry(100, 100, 1200, 700)
        
        # Try to load icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'icon_256x256.ico')
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except:
            pass
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # ============ HEADER SECTION ============
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(20, 20, 20, 20)
        header_layout.setSpacing(15)
        
        # Logo
        logo_path = os.path.join(os.path.dirname(__file__), 'QEDC-Full-Logo-Primary-Color.jpg')
        if os.path.exists(logo_path):
            logo_label = QLabel()
            pixmap = QPixmap(logo_path)
            pixmap = pixmap.scaledToHeight(60, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            header_layout.addWidget(logo_label)
        
        # Title
        title_label = QLabel('NYC Representatives Lookup')
        title_font = QFont('Arial', 24, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet('color: #00d4ff;')
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        
        # Header container
        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        header_widget.setStyleSheet('background-color: #1a1a2e; border-bottom: 2px solid #00d4ff;')
        main_layout.addWidget(header_widget)
        
        # ============ CONTENT SECTION ============
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)
        
        # Button row 1
        button_row1 = QHBoxLayout()
        button_row1.setSpacing(10)
        
        self.choose_file_btn = self.create_button('📂 Choose File', '#00d4ff', 40, self.choose_file)
        self.process_btn = self.create_button('🚀 Process', '#00ff00', 50, self.process_file)
        
        button_row1.addWidget(self.choose_file_btn)
        button_row1.addWidget(self.process_btn)
        button_row1.addStretch()
        content_layout.addLayout(button_row1)
        
        # Button row 2
        button_row2 = QHBoxLayout()
        button_row2.setSpacing(10)
        
        self.export_btn = self.create_button('💾 Export', '#ff9900', 45, self.export_data)
        self.support_btn = self.create_button('📧 Support', '#e94560', 45, self.show_support)
        
        button_row2.addWidget(self.export_btn)
        button_row2.addWidget(self.support_btn)
        button_row2.addStretch()
        content_layout.addLayout(button_row2)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet('''
            QProgressBar {
                border: 2px solid #00d4ff;
                border-radius: 5px;
                background-color: #16213e;
                text-align: center;
                color: #f1f5f9;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #00d4ff,
                                           stop:1 #00ff00);
                border-radius: 3px;
            }
        ''')
        self.progress_bar.setValue(0)
        content_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel('Ready to process file...')
        self.status_label.setStyleSheet('color: #cbd5e1; font-size: 12px;')
        content_layout.addWidget(self.status_label)
        
        # Table
        self.table = QTableWidget()
        self.table.setStyleSheet('''
            QTableWidget {
                background-color: #16213e;
                gridline-color: #334155;
                color: #f1f5f9;
                border: 1px solid #334155;
                border-radius: 5px;
            }
            QTableWidget::item {
                padding: 5px;
                border: 1px solid #334155;
            }
            QTableWidget::item:selected {
                background-color: #00d4ff;
                color: #1a1a2e;
            }
            QHeaderView::section {
                background-color: #0f172a;
                color: #00d4ff;
                padding: 5px;
                border: 1px solid #00d4ff;
                font-weight: bold;
            }
            QScrollBar:vertical {
                background-color: #16213e;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background-color: #00d4ff;
                border-radius: 6px;
            }
            QScrollBar:horizontal {
                background-color: #16213e;
                height: 12px;
            }
            QScrollBar::handle:horizontal {
                background-color: #00d4ff;
                border-radius: 6px;
            }
        ''')
        content_layout.addWidget(self.table)
        
        content_widget = QWidget()
        content_widget.setLayout(content_layout)
        content_widget.setStyleSheet('background-color: #1a1a2e;')
        main_layout.addWidget(content_widget)
        
        # Apply dark theme to main window
        self.setStyleSheet('background-color: #1a1a2e;')
        
    def create_button(self, text, color, height, callback):
        """Create a styled button with hover effects"""
        btn = QPushButton(text)
        btn.setMinimumHeight(height)
        btn.setCursor(Qt.PointingHandCursor)
        
        btn.setStyleSheet(f'''
            QPushButton {{
                background-color: {color};
                color: #1a1a2e;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
                padding: 10px 20px;
                transition: all 0.2s;
            }}
            QPushButton:hover {{
                background-color: {self.lighten_color(color, 20)};
                transform: translateY(-2px);
                box-shadow: 0 8px 16px rgba(0, 0, 0, 0.3);
            }}
            QPushButton:pressed {{
                background-color: {self.darken_color(color, 20)};
                transform: translateY(0px);
            }}
        ''')
        
        btn.clicked.connect(callback)
        return btn
    
    @staticmethod
    def lighten_color(hex_color, percent):
        """Lighten a hex color"""
        hex_color = hex_color.lstrip('#')
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r = min(255, int(r + (255 - r) * percent / 100))
        g = min(255, int(g + (255 - g) * percent / 100))
        b = min(255, int(b + (255 - b) * percent / 100))
        return f'#{r:02x}{g:02x}{b:02x}'
    
    @staticmethod
    def darken_color(hex_color, percent):
        """Darken a hex color"""
        hex_color = hex_color.lstrip('#')
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r = max(0, int(r * (1 - percent / 100)))
        g = max(0, int(g * (1 - percent / 100)))
        b = max(0, int(b * (1 - percent / 100)))
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def choose_file(self):
        """File chooser dialog"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            'Select CSV or XLSX file',
            '',
            'CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*.*)'
        )
        
        if file_path:
            self.status_label.setText(f'Selected: {Path(file_path).name}')
            self.selected_file = file_path
    
    def process_file(self):
        """Process the selected file"""
        if not hasattr(self, 'selected_file'):
            QMessageBox.warning(self, 'Error', 'Please select a file first!')
            return
        
        self.process_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText('Processing... this may take a moment')
        
        # Create and start worker thread
        self.worker_thread = WorkerThread(self.selected_file, self)
        self.worker_thread.finished.connect(self.on_processing_finished)
        self.worker_thread.error.connect(self.on_processing_error)
        self.worker_thread.start()
        
        # Simulate progress
        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.update_progress)
        self.progress_value = 0
        self.progress_timer.start(500)
    
    def update_progress(self):
        """Update progress bar"""
        if self.progress_value < 90:
            self.progress_value += 10
            self.progress_bar.setValue(self.progress_value)
    
    def on_processing_finished(self):
        """Handle processing completion"""
        self.progress_timer.stop()
        self.progress_bar.setValue(100)
        self.process_btn.setEnabled(True)
        self.status_label.setText('✅ Processing complete! Ready to export.')
        QMessageBox.information(self, 'Success', 'File processed successfully!')
    
    def on_processing_error(self, error_msg):
        """Handle processing error"""
        self.progress_timer.stop()
        self.progress_bar.setValue(0)
        self.process_btn.setEnabled(True)
        self.status_label.setText('❌ Error during processing')
        QMessageBox.critical(self, 'Error', f'Error: {error_msg}')
    
    def process_file_internal(self, file_path):
        """Internal file processing"""
        # Read file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        self.data = df
        self.display_table(df)
    
    def display_table(self, df):
        """Display data in table"""
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str))
        
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setForeground(QColor('#f1f5f9'))
                self.table.setItem(i, j, item)
        
        # Resize columns
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
    
    def export_data(self):
        """Export data to Excel"""
        if self.data is None:
            QMessageBox.warning(self, 'Error', 'No data to export!')
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            'Save Excel file',
            'results.xlsx',
            'Excel Files (*.xlsx)'
        )
        
        if file_path:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Results"
                
                # Write headers
                for col_idx, col_name in enumerate(self.data.columns, 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.value = col_name
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="00d4ff", end_color="00d4ff", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # Write data
                for row_idx, row in enumerate(self.data.values, 2):
                    for col_idx, value in enumerate(row, 1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.value = value
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        cell.border = Border(
                            left=Side(style='thin', color='334155'),
                            right=Side(style='thin', color='334155'),
                            top=Side(style='thin', color='334155'),
                            bottom=Side(style='thin', color='334155')
                        )
                
                # Adjust column widths
                for col_idx in range(1, len(self.data.columns) + 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 20
                
                wb.save(file_path)
                QMessageBox.information(self, 'Success', f'Data exported to {file_path}')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Export failed: {str(e)}')
    
    def show_support(self):
        """Show support information"""
        support_info = "NYC Representatives Lookup\n\n" \
                      "📧 Email: support@example.com\n" \
                      "🌐 Website: https://example.com\n" \
                      "📱 Phone: (555) 123-4567\n\n" \
                      "Made with ❤️ by QEDC"
        QMessageBox.information(self, 'Support', support_info)


def main():
    app = QApplication(sys.argv)
    window = NYC_RepresentativesLookup()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
