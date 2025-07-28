import sys, os, re
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox,
    QTextEdit, QTableWidget, QTableWidgetItem, QComboBox, QLineEdit, QCheckBox, QSpinBox, QGroupBox,
    QProgressBar, QTabWidget, QSplitter, QHeaderView, QStyledItemDelegate, QListWidget, QDialog,
    QGridLayout, QButtonGroup, QRadioButton, QFormLayout, QMenuBar, QAction
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QSettings
from PyQt5.QtGui import QColor
import logging

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        # Running from a PyInstaller bundle
        return os.path.join(sys._MEIPASS, relative_path)
    # Running from normal script
    return os.path.join(os.path.abspath("."), relative_path)

# Path to your default template (bundled into the EXE)
#default_excel_template = resource_path("C:\\Users\\alyousefh\\Desktop\\365DataSience\\Python\\VSC\\Excel to SQL\\Default_Excel_template_File.xlsx")
default_excel_template = resource_path("Default_Excel_template_File.xlsx")

# --- DataHandler: All Pandas/Excel/JSON logic (no UI code) ---
class DataHandler:
    @staticmethod
    def load_excel_sheets(file_path):
        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None)
            return {name: df for name, df in all_sheets.items() if not df.empty and len(df.columns) > 0}
        except Exception as e:
            raise e

    @staticmethod
    def get_preview(df, n_rows):
        return df.head(n_rows)

    @staticmethod
    def validate_row(row, column_mappings, skip_arabic, validate_quality, arabic_pattern, sp_params):
        formatted_params = {}
        skip_row = False
        stats = {'skipped_arabic': 0, 'skipped_invalid_value': 0, 'skipped_empty': 0}
        
        for sp_param, excel_col in column_mappings.items():
            value = getattr(row, excel_col)
            
            # Check for empty/null values - ONLY if validation is enabled
            if validate_quality and (pd.isna(value) or (isinstance(value, str) and value.strip().lower() in ['nan', 'none', ''])):
                stats['skipped_empty'] += 1
                skip_row = True
                break
                
            # Check for Arabic text - controlled by skip_arabic flag
            if skip_arabic and isinstance(value, str) and arabic_pattern.search(str(value).strip()):
                stats['skipped_arabic'] += 1
                skip_row = True
                break
                
            # Handle string parameters
            if sp_param in ['item', 'Status']:
                str_value = str(value).strip().replace("'", "''")
                formatted_params[sp_param] = str_value
                
            # Handle numeric parameters - validation controlled by validate_quality
            elif sp_param in ['qty', 'Slp_Discount', 'Spv_Discount', 'Mgr_Discount', 'New_Current_Cost', 'new_Showroom']:
                try:
                    if isinstance(value, str):
                        clean_value = value.replace(",", "").replace("$", "").replace("%", "").strip()
                        
                        # Only validate if quality validation is enabled
                        if validate_quality and (not clean_value or clean_value.lower() in ['n/a', 'na', 'null', 'none']):
                            stats['skipped_invalid_value'] += 1
                            skip_row = True
                            break
                            
                        numeric_value = float(clean_value)
                    else:
                        numeric_value = float(value)
                        
                    # Validate range - ONLY if validation is enabled
                    if validate_quality and numeric_value < 0 and sp_param in ['qty', 'New_Current_Cost', 'new_Showroom']:
                        logging.debug(f"Negative value detected for {sp_param}: {numeric_value}")
                        stats['skipped_invalid_value'] += 1
                        skip_row = True
                        break
                        
                    formatted_params[sp_param] = numeric_value
                    
                except (ValueError, AttributeError, TypeError) as e:
                    # Only skip on error if validation is enabled
                    if validate_quality:
                        logging.debug(f"Invalid numeric value for {sp_param}: '{value}' - {str(e)}")
                        stats['skipped_invalid_value'] += 1
                        skip_row = True
                        break
                    else:
                        # If validation disabled, use string representation or default value
                        formatted_params[sp_param] = str(value) if value is not None else '0'
                        
            else:
                str_value = str(value).strip().replace("'", "''")
                formatted_params[sp_param] = str_value
                
        return skip_row, formatted_params, stats
    
    @staticmethod
    def validate_row_by_index(row, column_indices, skip_arabic, validate_quality, arabic_pattern, sp_params):
        """Validate row using column indices instead of attribute names"""
        formatted_params = {}
        skip_row = False
        stats = {'skipped_arabic': 0, 'skipped_invalid_value': 0, 'skipped_empty': 0}
        
        for sp_param, col_index in column_indices.items():
            # Access value by index position - this always works!
            value = row[col_index]
            
            # Rest of validation logic remains the same...
            if validate_quality and (pd.isna(value) or (isinstance(value, str) and value.strip().lower() in ['nan', 'none', ''])):
                stats['skipped_empty'] += 1
                skip_row = True
                break
                
            if skip_arabic and isinstance(value, str) and arabic_pattern.search(str(value).strip()):
                stats['skipped_arabic'] += 1
                skip_row = True
                break
                
            if sp_param in ['item', 'Status']:
                str_value = str(value).strip().replace("'", "''")
                formatted_params[sp_param] = str_value
                
            elif sp_param in ['qty', 'Slp_Discount', 'Spv_Discount', 'Mgr_Discount', 'New_Current_Cost', 'new_Showroom']:
                try:
                    if isinstance(value, str):
                        clean_value = value.replace(",", "").replace("$", "").replace("%", "").strip()
                        if validate_quality and (not clean_value or clean_value.lower() in ['n/a', 'na', 'null', 'none']):
                            stats['skipped_invalid_value'] += 1
                            skip_row = True
                            break
                        numeric_value = float(clean_value)
                    else:
                        numeric_value = float(value)
                        
                    if validate_quality and numeric_value < 0 and sp_param in ['qty', 'New_Current_Cost', 'new_Showroom']:
                        logging.debug(f"Negative value detected for {sp_param}: {numeric_value}")
                        stats['skipped_invalid_value'] += 1
                        skip_row = True
                        break
                        
                    formatted_params[sp_param] = numeric_value
                    
                except (ValueError, AttributeError, TypeError) as e:
                    if validate_quality:
                        logging.debug(f"Invalid numeric value for {sp_param}: '{value}' - {str(e)}")
                        stats['skipped_invalid_value'] += 1
                        skip_row = True
                        break
                    else:
                        formatted_params[sp_param] = str(value) if value is not None else '0'
                        
            else:
                str_value = str(value).strip().replace("'", "''")
                formatted_params[sp_param] = str_value
                
        return skip_row, formatted_params, stats
    
    

# --- Worker: QThread for heavy tasks (Excel loading, SQL generation) ---
class ExcelLoaderWorker(QThread):
    finished = pyqtSignal(dict, list)
    error = pyqtSignal(str)
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
    def run(self):
        try:
            sheets = DataHandler.load_excel_sheets(self.file_path)
            self.finished.emit(sheets, list(sheets.keys()))
        except Exception as e:
            self.error.emit(str(e))

class SQLGeneratorWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, list, dict)
    error = pyqtSignal(str)
    status_update = pyqtSignal(str)
    def __init__(self, df, sheet_name, sp_details, column_mappings, output_path, skip_arabic=True, validate_quality=True):
        super().__init__()
        self.df = df
        self.sheet_name = sheet_name
        self.sp_details = sp_details
        self.column_mappings = column_mappings
        self.output_path = output_path
        self.skip_arabic = skip_arabic
        self.validate_quality = validate_quality
    def run(self):
        try:
            start_time = datetime.now()
            arabic_pattern = re.compile(r'[\u0600-\u06FF]')
            sql_lines = []
            total_rows = len(self.df)
            stats = {
                'total_rows': total_rows,
                'processed_rows': 0,
                'skipped_arabic': 0,
                'skipped_invalid_value': 0,
                'skipped_empty': 0,
                'processing_time': 0
            }
            
            # Convert column names to indices for reliable access
            column_indices = {}
            df_columns = list(self.df.columns)
            for sp_param, excel_col in self.column_mappings.items():
                if excel_col in df_columns:
                    column_indices[sp_param] = df_columns.index(excel_col)
                else:
                    raise ValueError(f"Column '{excel_col}' not found in DataFrame columns: {df_columns}")
            
            # Use itertuples for performance
            for idx, row in enumerate(self.df.itertuples(index=False), 1):
                progress = int((idx / total_rows) * 100)
                self.progress.emit(progress)
                if idx % 100 == 0:
                    self.status_update.emit(f"Processing row {idx} of {total_rows}")
                
                skip_row, formatted_params, row_stats = DataHandler.validate_row_by_index(
                    row, column_indices, self.skip_arabic, self.validate_quality, arabic_pattern, self.sp_details['parameters']
                )
                
                stats['skipped_arabic'] += row_stats['skipped_arabic']
                stats['skipped_invalid_value'] += row_stats['skipped_invalid_value']
                stats['skipped_empty'] += row_stats['skipped_empty']
                if skip_row:
                    continue
                try:
                    sql = self.sp_details['sql_template'].format(**formatted_params)
                    sql_lines.append(sql)
                    stats['processed_rows'] += 1
                except KeyError as e:
                    self.error.emit(f"Missing parameter for SQL formatting: {e}. Check column mappings.")
                    return
                except Exception as e:
                    self.error.emit(f"Error formatting SQL for row {idx}: {e}")
                    return
            try:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(f"-- Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"-- Source Excel Sheet: {self.sheet_name}\n")
                    f.write(f"-- Stored Procedure/SQL Type: {self.sp_details.get('friendly_name', 'Unknown')}\n")
                    f.write(f"-- Total statements: {len(sql_lines)}\n\n")
                    for line in sql_lines:
                        f.write(line + '\nGO\n')
            except Exception as e:
                self.error.emit(f"File write error: {e}")
                return
            stats['processing_time'] = (datetime.now() - start_time).total_seconds()
            self.finished.emit(self.output_path, sql_lines, stats)
        except Exception as e:
            self.error.emit(str(e))

# --- ColorDelegate: For preview table coloring ---
class ColorDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.error_rows = set()
        self.warning_rows = set()
        self.arabic_rows = set()
    def paint(self, painter, option, index):
        if index.row() in self.error_rows:
            option.backgroundBrush = QColor(255, 200, 200)
        elif index.row() in self.warning_rows:
            option.backgroundBrush = QColor(255, 255, 200)
        elif index.row() in self.arabic_rows:
            option.backgroundBrush = QColor(200, 200, 255)
        super().paint(painter, option, index)

# --- SettingsDialog: For preferences ---
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setFixedSize(400, 300)
        layout = QVBoxLayout()
        theme_group = QGroupBox("Theme")
        theme_layout = QVBoxLayout()
        self.theme_group = QButtonGroup()
        self.light_theme = QRadioButton("Light Theme")
        self.dark_theme = QRadioButton("Dark Theme")
        self.theme_group.addButton(self.light_theme, 0)
        self.theme_group.addButton(self.dark_theme, 1)
        theme_layout.addWidget(self.light_theme)
        theme_layout.addWidget(self.dark_theme)
        theme_group.setLayout(theme_layout)
        defaults_group = QGroupBox("Default Settings")
        defaults_layout = QGridLayout()
        defaults_layout.addWidget(QLabel("Default Preview Rows:"), 0, 0)
        self.default_preview_rows = QSpinBox()
        self.default_preview_rows.setRange(1, 100)
        self.default_preview_rows.setValue(10)
        defaults_layout.addWidget(self.default_preview_rows, 0, 1)
        defaults_layout.addWidget(QLabel("Default Stored Procedure:"), 1, 0)
        self.default_sp_friendly_name = QLineEdit()
        defaults_layout.addWidget(self.default_sp_friendly_name, 1, 1)
        self.auto_save_settings = QCheckBox("Auto-save settings")
        defaults_layout.addWidget(self.auto_save_settings, 2, 0, 1, 2)
        defaults_group.setLayout(defaults_layout)
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Cancel")
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addWidget(theme_group)
        layout.addWidget(defaults_group)
        layout.addLayout(button_layout)
        self.setLayout(layout)
        self.load_settings()
    def load_settings(self):
        settings = QSettings('ExcelToSQL', 'Settings')
        theme = settings.value('theme', 'light')
        if theme == 'dark':
            self.dark_theme.setChecked(True)
        else:
            self.light_theme.setChecked(True)
        self.default_preview_rows.setValue(int(settings.value('default_preview_rows', 10)))
        self.default_sp_friendly_name.setText(settings.value('default_sp_friendly_name', 'Update Items Dropship Quantities'))
        self.auto_save_settings.setChecked(settings.value('auto_save_settings', True, type=bool))
    def save_settings(self):
        settings = QSettings('ExcelToSQL', 'Settings')
        theme = 'dark' if self.dark_theme.isChecked() else 'light'
        settings.setValue('theme', theme)
        settings.setValue('default_preview_rows', self.default_preview_rows.value())
        settings.setValue('default_sp_friendly_name', self.default_sp_friendly_name.text())
        settings.setValue('auto_save_settings', self.auto_save_settings.isChecked())

# --- MainWindow: All UI widgets/layouts ---
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # Configure logging 
        # This will log to both a file and the console
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('excel_to_sql.log'),
                logging.StreamHandler()  # Also print to console
            ]
        )
        
        self.setWindowTitle("Excel to SQL Script Generator v3.0")
        self.setGeometry(200, 100, 1200, 800)
        self.menubar = None
        self.df_all_sheets = {}
        self.selected_sheet_name = None
        self.current_df = None
        self.current_df_columns = []
        self.file_path = ""
        self.recent_files = []
        self.processing_history = []
        self.stored_procedures = {
            "Update Items Dropship Quantities": {
                "sql_template": "EXEC [dbo].[Hyou_UPDATE_EVS_ItemAddational_DROPSHIP_QTY_Excel] @ITEMNMBR = '{item}', @QTY = {qty:.3f}, @F1 = NULL, @F2 = NULL",
                "parameters": ["item", "qty"],
                "friendly_name": "Update Items Dropship Quantities"
            },
            "Update Markdown Discounts": {
                "sql_template": "EXEC [dbo].[HYOU_SP_UPDATE_Makdown_Discount_All_Levels] @ITEMNMBR = '{item}', @Slp_Markdown = {Slp_Discount:.3f}, @Spv_Markdown = {Spv_Discount:.3f}, @Mgr_Markdown = {Mgr_Discount:.3f}",
                "parameters": ["item", "Slp_Discount", "Spv_Discount", "Mgr_Discount"],
                "friendly_name": "Update Markdown Discounts"
            },
            "Update Items Current Cost": {
                "sql_template": "UPDATE IV00101 SET CURRCOST = {New_Current_Cost:.3f} WHERE ITEMNMBR = '{item}'",
                "parameters": ["item", "New_Current_Cost"],
                "friendly_name": "Update Items Current Cost"
            },
            "Update Items Status": {
                "sql_template": "UPDATE IV00101 SET USCATVLS_6 = '{Status}', INACTIVE = 1, ITEMTYPE = 2 WHERE ITEMNMBR = '{item}'",
                "parameters": ["item", "Status"],
                "friendly_name": "Update Items Status"
            },
            "Update Items Prices": {
                "sql_template": "EXEC [dbo].[HYOU_SP_UPDATE_Item_Price_IV00108&More] @ITEMNMBR = '{item}', @PRCLEVEL = 'SHOWROOM', @PRICE = {new_Showroom:.3f}",
                "parameters": ["item", "new_Showroom"],
                "friendly_name": "Update Items Prices"
            }
        }
        self.color_delegate = ColorDelegate()
        self.param_column_combos = {}
        self.mapping_widgets_layout = QFormLayout()
        self.load_settings()
        self.init_ui()
        self.setup_menu()
        self.apply_theme()
        self.setAcceptDrops(True)
        self.auto_save_timer = QTimer()
        self.auto_save_timer.timeout.connect(self.auto_save_settings)
        self.auto_save_timer.start(30000)
        # Heartbeat label for debugging UI freezes
        self.heartbeat_label = QLabel("UI Alive")
        self.heartbeat_timer = QTimer()
        self.heartbeat_timer.timeout.connect(lambda: self.heartbeat_label.setText(f"UI Alive: {datetime.now()}"))
        self.heartbeat_timer.start(1000)
    # --- Settings ---
    def load_settings(self):
        settings = QSettings('ExcelToSQL', 'Settings')
        self.theme = settings.value('theme', 'light')
        self.default_preview_rows = int(settings.value('default_preview_rows', 10))
        self.default_sp_friendly_name = settings.value('default_sp_friendly_name', 'Update Items Dropship Quantities')
        self.auto_save_enabled = settings.value('auto_save_settings', True, type=bool)
        self.recent_files = settings.value('recent_files', [], type=list)
        if len(self.recent_files) > 10:
            self.recent_files = self.recent_files[-10:]
    def save_settings(self):
        settings = QSettings('ExcelToSQL', 'Settings')
        settings.setValue('theme', self.theme)
        settings.setValue('default_preview_rows', self.default_preview_rows)
        if hasattr(self, 'sp_selector') and self.sp_selector is not None:
            settings.setValue('default_sp_friendly_name', self.sp_selector.currentText())
        else:
            settings.setValue('default_sp_friendly_name', self.default_sp_friendly_name)
        self.default_sp_friendly_name = settings.value('default_sp_friendly_name')
        settings.setValue('auto_save_settings', self.auto_save_enabled)
        settings.setValue('recent_files', self.recent_files)
    def auto_save_settings(self):
        if self.auto_save_enabled:
            self.save_settings()
    # --- Menu ---
    def setup_menu(self):
        self.menubar = QMenuBar(self)
        file_menu = self.menubar.addMenu('File')
        open_action = QAction('Open Excel File', self)
        open_action.setShortcut('Ctrl+O')
        open_action.triggered.connect(self.open_excel_dialog)
        file_menu.addAction(open_action)
        file_menu.addSeparator()
        self.recent_menu = file_menu.addMenu('Recent Files')
        self.update_recent_menu()
        file_menu.addSeparator()
        exit_action = QAction('Exit', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        settings_menu = self.menubar.addMenu('Settings')
        preferences_action = QAction('Preferences', self)
        preferences_action.triggered.connect(self.show_settings)
        settings_menu.addAction(preferences_action)
        help_menu = self.menubar.addMenu('Help')
        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    def update_recent_menu(self):
        self.recent_menu.clear()
        for file_path in self.recent_files:
            if os.path.exists(file_path):
                action = QAction(os.path.basename(file_path), self)
                action.setData(file_path)
                action.triggered.connect(lambda checked, path=file_path: self.load_recent_file(path))
                self.recent_menu.addAction(action)
        if not self.recent_files:
            no_recent_action = QAction('No recent files', self)
            no_recent_action.setEnabled(False)
            self.recent_menu.addAction(no_recent_action)
    def load_recent_file(self, file_path):
        self.file_path = file_path
        self.controller.load_excel_file_threaded(self.file_path)
    def show_settings(self):
        dialog = SettingsDialog(self)
        dialog.default_sp_friendly_name.setText(self.default_sp_friendly_name)
        if dialog.exec_() == QDialog.Accepted:
            dialog.save_settings()
            self.load_settings()
            self.apply_theme()
            index = self.sp_selector.findText(self.default_sp_friendly_name)
            if index != -1:
                self.sp_selector.setCurrentIndex(index)
            self.preview_rows_spin.setValue(self.default_preview_rows)
    def show_about(self):
        QMessageBox.about(self, "About",
            "Excel to SQL Script Generator v3.0\n\n"
            "A powerful tool for converting Excel data to SQL scripts\n"
            "with advanced features and customization options.")
    def apply_theme(self):
        if self.theme == 'dark':
            self.setStyleSheet("""
                QWidget { background-color: #2b2b2b; color: #ffffff; }
                QGroupBox { font-weight: bold; border: 2px solid #555555; border-radius: 8px; margin-top: 10px; padding-top: 10px; }
                QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px 0 5px; }
                QPushButton { background-color: #404040; border: 1px solid #555555; border-radius: 4px; padding: 8px; min-width: 80px; }
                QPushButton:hover { background-color: #505050; }
                QPushButton:pressed { background-color: #606060; }
                QLineEdit, QComboBox, QSpinBox { background-color: #404040; border: 1px solid #555555; border-radius: 4px; padding: 4px; }
                QTableWidget { background-color: #353535; alternate-background-color: #404040; gridline-color: #555555; }
                QTextEdit { background-color: #353535; border: 1px solid #555555; }
                QProgressBar { background-color: #404040; border: 1px solid #555555; border-radius: 4px; text-align: center; }
                QProgressBar::chunk { background-color: #4a9eff; border-radius: 3px; }
                QTabWidget::pane { border: 2px solid #555555; border-radius: 8px; background: #353535; }
                QTabBar::tab { background: #404040; color: #ffffff; border: 1px solid #555555; border-radius: 4px; padding: 8px; min-width: 100px; }
                QTabBar::tab:selected { background: #4a9eff; color: #ffffff; }
            """)
        else:
            self.setStyleSheet("""
                QGroupBox { font-weight: bold; border: 2px solid #cccccc; border-radius: 8px; margin-top: 10px; padding-top: 10px; }
                QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px 0 5px; }
                QPushButton { background-color: #f0f0f0; border: 1px solid #cccccc; border-radius: 4px; padding: 8px; min-width: 80px; }
                QPushButton:hover { background-color: #e0e0e0; }
                QPushButton:pressed { background-color: #d0d0d0; }
                QProgressBar { background-color: #f0f0f0; border: 1px solid #cccccc; border-radius: 4px; text-align: center; }
                QProgressBar::chunk { background-color: #4a9eff; border-radius: 3px; }
                QTabWidget::pane { border: 2px solid #cccccc; border-radius: 8px; background: #f0f0f0; }
                QTabBar::tab { background: #f0f0f0; color: #000000; border: 1px solid #cccccc; border-radius: 4px; padding: 8px; min-width: 100px; }
                QTabBar::tab:selected { background: #4a9eff; color: #ffffff; }
            """)
    # --- Drag & Drop ---
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                self.file_path = file_path
                self.controller.load_excel_file_threaded(self.file_path)
                break
    # --- UI Layout ---
    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setMenuBar(self.menubar)
        main_splitter = QSplitter(Qt.Horizontal)
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        file_group = QGroupBox("📁 File Selection")
        file_layout = QVBoxLayout()
        self.drop_zone = QLabel("📂 Drop Excel file here or click to browse")
        self.drop_zone.setAlignment(Qt.AlignCenter)
        self.drop_zone.setStyleSheet("QLabel { border: 2px dashed #cccccc; border-radius: 8px; padding: 20px; font-size: 14px; color: #666666; }")
        self.drop_zone.setMinimumHeight(80)
        self.drop_zone.mousePressEvent = lambda e: self.open_excel_dialog()
        self.file_label = QLabel("No file selected")
        self.file_button = QPushButton("Browse Files")
        self.file_button.clicked.connect(self.open_excel_dialog)
        sheet_layout = QHBoxLayout()
        self.sheet_label = QLabel("Sheet:")
        self.sheet_label.hide()
        self.sheet_selector = QComboBox()
        self.sheet_selector.hide()
        self.sheet_selector.currentIndexChanged.connect(self.on_sheet_changed)
        sheet_layout.addWidget(self.sheet_label)
        sheet_layout.addWidget(self.sheet_selector)
        file_layout.addWidget(self.drop_zone)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.file_button)
        file_layout.addLayout(sheet_layout)
        file_group.setLayout(file_layout)
        config_group = QGroupBox("⚙️ Configuration")
        config_layout = QVBoxLayout()
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Output file:"))
        self.output_path_input = QLineEdit("output_script.sql")
        self.browse_output_button = QPushButton("Browse")
        self.browse_output_button.clicked.connect(self.browse_output_file)
        output_layout.addWidget(self.output_path_input)
        output_layout.addWidget(self.browse_output_button)
        sp_layout = QHBoxLayout()
        sp_layout.addWidget(QLabel("Stored Procedure:"))
        self.sp_selector = QComboBox()
        self.sp_selector.addItems(list(self.stored_procedures.keys()))
        default_sp_index = self.sp_selector.findText(self.default_sp_friendly_name)
        if default_sp_index != -1:
            self.sp_selector.setCurrentIndex(default_sp_index)
        else:
            self.sp_selector.setCurrentIndex(0)
            self.default_sp_friendly_name = self.sp_selector.currentText()
        self.sp_selector.currentIndexChanged.connect(self.on_sp_changed)
        sp_layout.addWidget(self.sp_selector)
        self.skip_arabic_check = QCheckBox("Skip rows with Arabic text")
        self.skip_arabic_check.setChecked(True)
        self.validate_data_check = QCheckBox("Validate data quality")
        self.validate_data_check.setChecked(True)
        config_layout.addLayout(output_layout)
        config_layout.addLayout(sp_layout)
        config_layout.addWidget(self.skip_arabic_check)
        config_layout.addWidget(self.validate_data_check)
        config_group.setLayout(config_layout)
        mapping_group = QGroupBox("🔗 Column Mapping")
        mapping_group.setLayout(self.mapping_widgets_layout)
        process_group = QGroupBox("🚀 Processing")
        process_layout = QVBoxLayout()
        self.generate_button = QPushButton("Generate SQL Script")
        self.generate_button.clicked.connect(self.controller_generate_sql)
        self.generate_button.setEnabled(False)
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        self.status_label = QLabel("")
        self.status_label.hide()
        process_layout.addWidget(self.generate_button)
        process_layout.addWidget(self.progress_bar)
        process_layout.addWidget(self.status_label)
        process_group.setLayout(process_layout)
        left_layout.addWidget(file_group)
        left_layout.addWidget(config_group)
        left_layout.addWidget(mapping_group)
        left_layout.addWidget(process_group)
        left_layout.addStretch()
        left_panel.setLayout(left_layout)
        right_panel = QTabWidget()
        preview_tab = QWidget()
        preview_layout = QVBoxLayout()
        preview_controls = QHBoxLayout()
        preview_controls.addWidget(QLabel("Preview rows:"))
        self.preview_rows_spin = QSpinBox()
        self.preview_rows_spin.setRange(1, 100)
        self.preview_rows_spin.setValue(self.default_preview_rows)
        self.preview_rows_spin.valueChanged.connect(self.update_preview)
        preview_controls.addWidget(self.preview_rows_spin)
        preview_controls.addStretch()
        self.table_output = QTableWidget()
        self.table_output.setItemDelegate(self.color_delegate)
        self.table_output.setAlternatingRowColors(True)
        preview_layout.addLayout(preview_controls)
        preview_layout.addWidget(self.table_output)
        preview_tab.setLayout(preview_layout)
        stats_tab = QWidget()
        stats_layout = QVBoxLayout()
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)
        stats_layout.addWidget(self.stats_text)
        stats_tab.setLayout(stats_layout)
        log_tab = QWidget()
        log_layout = QVBoxLayout()
        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        log_layout.addWidget(self.text_output)
        log_tab.setLayout(log_layout)
        history_tab = QWidget()
        history_layout = QVBoxLayout()
        self.history_list = QListWidget()
        history_layout.addWidget(self.history_list)
        history_tab.setLayout(history_layout)
        right_panel.addTab(preview_tab, "📊 Preview")
        right_panel.addTab(stats_tab, "📈 Statistics")
        right_panel.addTab(log_tab, "📝 Log")
        right_panel.addTab(history_tab, "🕐 History")
        main_splitter.addWidget(left_panel)
        main_splitter.addWidget(right_panel)
        main_splitter.setSizes([400, 800])
        main_layout.addWidget(main_splitter)
        self.setLayout(main_layout)
        self.on_sp_changed()
    # --- UI Event Handlers ---
    def open_excel_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        self.file_path = file_path
        self.controller.load_excel_file_threaded(self.file_path)
    def browse_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save SQL Script", "", "SQL Files (*.sql);;All Files (*)")
        if file_path:
            self.output_path_input.setText(file_path)
    def on_sheet_changed(self):
        self.selected_sheet_name = self.sheet_selector.currentText()
        self.reload_sheet_data()
    def on_sp_changed(self):
        while self.mapping_widgets_layout.count():
            item = self.mapping_widgets_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.param_column_combos.clear()
        selected_sp_friendly_name = self.sp_selector.currentText()
        sp_details = self.stored_procedures.get(selected_sp_friendly_name)
        if sp_details:
            parameters = sp_details['parameters']
            for param in parameters:
                label = QLabel(f"{param.replace('_', ' ').title()} Column:")
                combo = QComboBox()
                combo.addItems(["-- Select Column --"] + [str(col) for col in self.current_df_columns])
                self.mapping_widgets_layout.addRow(label, combo)
                self.param_column_combos[param] = combo
                if self.current_df_columns:
                    for i, col_name in enumerate(self.current_df_columns):
                        if col_name.lower() == param.lower():
                            combo.setCurrentIndex(i + 1)
                            break
        else:
            self.text_output.append(f"Warning: Stored procedure '{selected_sp_friendly_name}' not found in definitions.")
        self.generate_button.setEnabled(bool(self.current_df_columns) and bool(sp_details))
    def update_preview(self):
        if self.selected_sheet_name and self.selected_sheet_name in self.df_all_sheets:
            self.reload_sheet_data()
        else:
            self.file_label.setText("No sheet selected or sheet not found.")
            self.table_output.clear()
            self.stats_text.clear()
    def reload_sheet_data(self):
        if not self.selected_sheet_name or self.selected_sheet_name not in self.df_all_sheets:
            self.text_output.append("No sheet selected or sheet data not available for preview.")
            self.table_output.clear()
            self.stats_text.clear()
            self.current_df = None
            self.current_df_columns = []
            self.on_sp_changed()
            self.generate_button.setEnabled(False)
            return
        self.current_df = self.df_all_sheets[self.selected_sheet_name]
        self.current_df_columns = list(self.current_df.columns)
        self.on_sp_changed()
        self.color_delegate.error_rows.clear()
        self.color_delegate.warning_rows.clear()
        self.color_delegate.arabic_rows.clear()
        preview_rows_count = self.preview_rows_spin.value()
        df_preview = self.current_df.head(preview_rows_count)
        self.table_output.setRowCount(df_preview.shape[0])
        self.table_output.setColumnCount(df_preview.shape[1])
        self.table_output.setHorizontalHeaderLabels(df_preview.columns)
        self.table_output.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        arabic_pattern = re.compile(r'[\u0600-\u06FF]')
        for r_idx, row in enumerate(df_preview.itertuples(index=False)):
            for c_idx, col_name in enumerate(df_preview.columns):
                value = row[c_idx]  # <-- Use index-based access, always works
                item = QTableWidgetItem(str(value))
                self.table_output.setItem(r_idx, c_idx, item)
                if self.validate_data_check.isChecked():
                    if pd.isna(value) or str(value).strip().lower() in ['nan', 'none', '']:
                        self.color_delegate.error_rows.add(r_idx)
                    elif isinstance(value, str) and arabic_pattern.search(value):
                        self.color_delegate.arabic_rows.add(r_idx)
        self.table_output.viewport().update()
        total_rows = len(self.current_df)
        self.stats_text.setText(
            f"--- Sheet Statistics: {self.selected_sheet_name} ---\n"
            f"Total Rows: {total_rows}\n"
            f"Total Columns: {len(self.current_df_columns)}\n"
            f"Columns: {', '.join(self.current_df_columns)}\n"
            f"\n(Note: Detailed processing statistics will be available after generating SQL script.)"
        )
        self.generate_button.setEnabled(True)
    def controller_generate_sql(self):
        logging.debug("Generate SQL button clicked - testing logging")
        self.controller.generate_sql()
    def add_to_recent_files(self, file_path):
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        self.recent_files.append(file_path)
        if len(self.recent_files) > 10:
            self.recent_files.pop(0)
        self.update_recent_menu()
    def update_history_list(self):
        self.history_list.clear()
        for entry in reversed(self.processing_history):
            item_text = (
                f"[{entry['timestamp']}] {entry['status']} - "
                f"File: {entry['file']}, Sheet: {entry['sheet']}, SP: {entry['sp_name']}, "
                f"Processed: {entry['processed_rows']} rows"
            )
            self.history_list.addItem(item_text)
    # Add this method to MainWindow class around line 575

def validate_column_mappings(self, column_mappings):
    """Validate that column mappings make sense based on data types"""
    
    if not self.current_df or self.current_df.empty:
        return True, ""
    
    validation_errors = []
    sample_size = min(100, len(self.current_df))  # Check first 100 rows for performance
    df_sample = self.current_df.head(sample_size)
    
    for param, excel_col in column_mappings.items():
        if excel_col == "-- Select Column --":
            continue
            
        try:
            # Get the actual column data
            if excel_col not in df_sample.columns:
                validation_errors.append(f"Column '{excel_col}' not found in Excel sheet")
                continue
                
            column_data = df_sample[excel_col].dropna()  # Remove NaN for analysis
            
            if column_data.empty:
                validation_errors.append(f"Column '{excel_col}' mapped to '{param}' contains only empty values")
                continue
            
            # Validate based on parameter type
            if param in ['qty', 'Slp_Discount', 'Spv_Discount', 'Mgr_Discount', 'New_Current_Cost', 'new_Showroom']:
                # Should be numeric
                numeric_count = 0
                non_numeric_examples = []
                
                for idx, value in column_data.head(20).items():  # Check first 20 non-null values
                    try:
                        if isinstance(value, str):
                            # Try to clean and convert
                            clean_value = value.replace(",", "").replace("$", "").replace("%", "").strip()
                            float(clean_value)
                        else:
                            float(value)
                        numeric_count += 1
                    except (ValueError, TypeError):
                        if len(non_numeric_examples) < 3:  # Collect up to 3 examples
                            non_numeric_examples.append(str(value)[:20])  # Truncate long values
                
                # If less than 70% are numeric, warn user
                if numeric_count < len(column_data.head(20)) * 0.7:
                    examples_str = ", ".join(f"'{ex}'" for ex in non_numeric_examples)
                    validation_errors.append(
                        f"Parameter '{param}' expects numeric values, but column '{excel_col}' contains "
                        f"non-numeric data. Examples: {examples_str}"
                    )
                    
            elif param in ['item', 'Status']:
                # Should be text-like
                text_count = 0
                for value in column_data.head(10):
                    if isinstance(value, (str, int, float)) and str(value).strip():
                        text_count += 1
                
                if text_count == 0:
                    validation_errors.append(
                        f"Parameter '{param}' expects text values, but column '{excel_col}' appears to be empty or invalid"
                    )
                    
        except Exception as e:
            validation_errors.append(f"Error analyzing column '{excel_col}' for parameter '{param}': {str(e)}")
    
    if validation_errors:
        error_message = "Column Mapping Issues Found:\n\n" + "\n".join(f"• {error}" for error in validation_errors)
        error_message += "\n\nPlease review your column mappings and try again."
        return False, error_message
    
    return True, ""

# Update the generate_sql method in AppController around line 850

def generate_sql(self):
    if self.sql_generator_thread and self.sql_generator_thread.isRunning():
        QMessageBox.warning(self.window, "Processing in Progress", "A script generation is already in progress. Please wait.")
        return
    if self.window.current_df is None or self.window.current_df.empty:
        QMessageBox.warning(self.window, "No Data", "Please load an Excel file and select a sheet with data first.")
        return
    output_path = self.window.output_path_input.text()
    if not output_path:
        QMessageBox.warning(self.window, "Output File Missing", "Please specify an output SQL file path.")
        return
    selected_sp_friendly_name = self.window.sp_selector.currentText()
    sp_details = self.window.stored_procedures.get(selected_sp_friendly_name)
    if not sp_details:
        QMessageBox.critical(self.window, "Invalid Stored Procedure", "Selected Stored Procedure definition not found.")
        return
    column_mappings = {}
    all_mappings_selected = True
    for param, combo in self.window.param_column_combos.items():
        selected_excel_col = combo.currentText()
        if selected_excel_col == "-- Select Column --":
            QMessageBox.warning(self.window, "Column Mapping Missing", f"Please map a column for '{param.replace('_', ' ').title()}'.")
            all_mappings_selected = False
            break
        column_mappings[param] = selected_excel_col
    if not all_mappings_selected:
        return
    required_params = set(sp_details['parameters'])
    mapped_params = set(column_mappings.keys())
    if not required_params.issubset(mapped_params):
        missing_params = required_params - mapped_params
        QMessageBox.critical(self.window, "Missing Mappings",
                             f"Not all required parameters for '{selected_sp_friendly_name}' are mapped. Missing: {', '.join(missing_params)}")
        return
    
    # NEW: Validate column mappings before processing
    is_valid, error_message = self.window.validate_column_mappings(column_mappings)
    if not is_valid:
        msg_box = QMessageBox(self.window)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("Column Mapping Validation")
        msg_box.setText("There are issues with your column mappings:")
        msg_box.setDetailedText(error_message)
        msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Ignore)
        msg_box.setDefaultButton(QMessageBox.Ok)
        
        result = msg_box.exec_()
        if result == QMessageBox.Ok:
            return  # User wants to fix mappings
        # If user clicks Ignore, continue processing
    
    self.window.text_output.append(f"Starting SQL generation for '{selected_sp_friendly_name}'...")
    self.window.progress_bar.setValue(0)
    self.window.progress_bar.show()
    self.window.status_label.setText("Initializing processing...")
    self.window.status_label.show()
    self.window.generate_button.setEnabled(False)
    self.sql_generator_thread = SQLGeneratorWorker(
        df=self.window.current_df,
        sheet_name=self.window.selected_sheet_name,
        sp_details=sp_details,
        column_mappings=column_mappings,
        output_path=output_path,
        skip_arabic=self.window.skip_arabic_check.isChecked(),
        validate_quality=self.window.validate_data_check.isChecked()
    )
    self.sql_generator_thread.progress.connect(self.window.progress_bar.setValue)
    self.sql_generator_thread.status_update.connect(self.window.status_label.setText)
    self.sql_generator_thread.finished.connect(self.on_processing_finished)
    self.sql_generator_thread.error.connect(self.on_processing_error)
    self.sql_generator_thread.start()

# Also add the safe column name handling to validate_row method (update around line 47)

@staticmethod
def validate_row(row, column_mappings, skip_arabic, validate_quality, arabic_pattern, sp_params):
    formatted_params = {}
    skip_row = False
    stats = {'skipped_arabic': 0, 'skipped_invalid_value': 0, 'skipped_empty': 0}
    
    for sp_param, excel_col in column_mappings.items():
        # Try to get the value with safe column name handling
        value = None
        possible_names = [
            excel_col,  # Original name
            excel_col.replace(' ', '_'),  # Space to underscore
            excel_col.replace('.', '_'),  # Dot to underscore  
            excel_col.replace(' ', '_').replace('.', '_'),  # Both
            re.sub(r'[^\w]', '_', excel_col),  # All special chars to underscore
        ]
        
        for name in possible_names:
            try:
                value = getattr(row, name)
                break
            except AttributeError:
                continue
        
        if value is None:
            # Last resort: show available attributes for debugging
            available_attrs = [attr for attr in dir(row) if not attr.startswith('_') and not callable(getattr(row, attr))]
            raise ValueError(f"Column '{excel_col}' not accessible. Available columns: {available_attrs[:10]}...")
        
        # Rest of validation logic remains the same...
        if validate_quality and (pd.isna(value) or (isinstance(value, str) and value.strip().lower() in ['nan', 'none', ''])):
            stats['skipped_empty'] += 1
            skip_row = True
            break
            
        if skip_arabic and isinstance(value, str) and arabic_pattern.search(str(value).strip()):
            stats['skipped_arabic'] += 1
            skip_row = True
            break
            
        if sp_param in ['item', 'Status']:
            str_value = str(value).strip().replace("'", "''")
            formatted_params[sp_param] = str_value
            
        elif sp_param in ['qty', 'Slp_Discount', 'Spv_Discount', 'Mgr_Discount', 'New_Current_Cost', 'new_Showroom']:
            try:
                if isinstance(value, str):
                    clean_value = value.replace(",", "").replace("$", "").replace("%", "").strip()
                    if validate_quality and (not clean_value or clean_value.lower() in ['n/a', 'na', 'null', 'none']):
                        stats['skipped_invalid_value'] += 1
                        skip_row = True
                        break
                    numeric_value = float(clean_value)
                else:
                    numeric_value = float(value)
                    
                if validate_quality and numeric_value < 0 and sp_param in ['qty', 'New_Current_Cost', 'new_Showroom']:
                    logging.debug(f"Negative value detected for {sp_param}: {numeric_value}")
                    stats['skipped_invalid_value'] += 1
                    skip_row = True
                    break
                    
                formatted_params[sp_param] = numeric_value
                
            except (ValueError, AttributeError, TypeError) as e:
                if validate_quality:
                    logging.debug(f"Invalid numeric value for {sp_param}: '{value}' - {str(e)}")
                    stats['skipped_invalid_value'] += 1
                    skip_row = True
                    break
                else:
                    formatted_params[sp_param] = str(value) if value is not None else '0'
                    
        else:
            str_value = str(value).strip().replace("'", "''")
            formatted_params[sp_param] = str_value
            
    return skip_row, formatted_params, stats        
            

# --- AppController: Connects everything, handles events, signals, slots ---
class AppController:
    def __init__(self, window: MainWindow):
        self.window = window
        self.window.controller = self
        self.excel_loader_thread = None
        self.sql_generator_thread = None
    def load_excel_file_threaded(self, file_path):
        self.window.df_all_sheets.clear()
        self.window.current_df = None
        self.window.file_label.setText("Loading Excel file...")
        self.window.drop_zone.setText("Loading...")
        self.window.text_output.append(f"Loading: {os.path.basename(file_path)}")
        self.window.generate_button.setEnabled(False)
        self.excel_loader_thread = ExcelLoaderWorker(file_path)
        self.excel_loader_thread.finished.connect(self.on_excel_loaded)
        self.excel_loader_thread.error.connect(self.on_excel_load_error)
        self.excel_loader_thread.start()
    def on_excel_loaded(self, sheets, sheet_names):
        self.window.df_all_sheets = sheets
        self.window.add_to_recent_files(self.window.file_path)
        if not sheet_names:
            self.window.file_label.setText("No non-empty sheets found.")
            self.window.drop_zone.setText("📂 Drop Excel file here or click to browse")
            self.window.text_output.append("No non-empty sheets found in the file.")
            self.window.generate_button.setEnabled(False)
            return
        self.window.file_label.setText(f"✔ File loaded: {os.path.basename(self.window.file_path)}")
        self.window.drop_zone.setText(f"✔ {os.path.basename(self.window.file_path)} loaded")
        self.window.text_output.clear()
        debug_info = f"Found {len(sheet_names)} non-empty sheet(s): {', '.join(sheet_names)}"
        self.window.text_output.append(debug_info)
        self.window.sheet_selector.clear()
        self.window.sheet_selector.addItems(sheet_names)
        if self.window.selected_sheet_name and self.window.selected_sheet_name in sheet_names:
            self.window.sheet_selector.setCurrentText(self.window.selected_sheet_name)
        else:
            self.window.sheet_selector.setCurrentIndex(0)
        self.window.sheet_selector.show()
        self.window.sheet_label.show()
        if len(sheet_names) > 1:
            self.window.text_output.append("Multiple sheets found - please select one from the dropdown.")
        else:
            self.window.text_output.append("Single sheet found - automatically selected.")
        self.window.selected_sheet_name = self.window.sheet_selector.currentText()
        self.window.reload_sheet_data()
        self.window.generate_button.setEnabled(True)
        self.window.file_path = ""
    def on_excel_load_error(self, message):
        self.window.file_label.setText("Error loading Excel file.")
        self.window.drop_zone.setText("📂 Drop Excel file here or click to browse")
        self.window.text_output.append(f"Error: {message}")
        QMessageBox.critical(self.window, "Error", message)
        self.window.generate_button.setEnabled(False)
    def generate_sql(self):
        if self.sql_generator_thread and self.sql_generator_thread.isRunning():
            QMessageBox.warning(self.window, "Processing in Progress", "A script generation is already in progress. Please wait.")
            return
        if self.window.current_df is None or self.window.current_df.empty:
            QMessageBox.warning(self.window, "No Data", "Please load an Excel file and select a sheet with data first.")
            return
        output_path = self.window.output_path_input.text()
        if not output_path:
            QMessageBox.warning(self.window, "Output File Missing", "Please specify an output SQL file path.")
            return
        selected_sp_friendly_name = self.window.sp_selector.currentText()
        sp_details = self.window.stored_procedures.get(selected_sp_friendly_name)
        if not sp_details:
            QMessageBox.critical(self.window, "Invalid Stored Procedure", "Selected Stored Procedure definition not found.")
            return
        column_mappings = {}
        all_mappings_selected = True
        for param, combo in self.window.param_column_combos.items():
            selected_excel_col = combo.currentText()
            if selected_excel_col == "-- Select Column --":
                QMessageBox.warning(self.window, "Column Mapping Missing", f"Please map a column for '{param.replace('_', ' ').title()}'.")
                all_mappings_selected = False
                break
            column_mappings[param] = selected_excel_col
        if not all_mappings_selected:
            return
        required_params = set(sp_details['parameters'])
        mapped_params = set(column_mappings.keys())
        if not required_params.issubset(mapped_params):
            missing_params = required_params - mapped_params
            QMessageBox.critical(self.window, "Missing Mappings",
                                 f"Not all required parameters for '{selected_sp_friendly_name}' are mapped. Missing: {', '.join(missing_params)}")
            return
        self.window.text_output.append(f"Starting SQL generation for '{selected_sp_friendly_name}'...")
        self.window.progress_bar.setValue(0)
        self.window.progress_bar.show()
        self.window.status_label.setText("Initializing processing...")
        self.window.status_label.show()
        self.window.generate_button.setEnabled(False)
        self.sql_generator_thread = SQLGeneratorWorker(
            df=self.window.current_df,
            sheet_name=self.window.selected_sheet_name,
            sp_details=sp_details,
            column_mappings=column_mappings,
            output_path=output_path,
            skip_arabic=self.window.skip_arabic_check.isChecked(),
            validate_quality=self.window.validate_data_check.isChecked()
        )
        self.sql_generator_thread.progress.connect(self.window.progress_bar.setValue)
        self.sql_generator_thread.status_update.connect(self.window.status_label.setText)
        self.sql_generator_thread.finished.connect(self.on_processing_finished)
        self.sql_generator_thread.error.connect(self.on_processing_error)
        self.sql_generator_thread.start()
    def on_processing_finished(self, output_path, sql_lines, stats):
        self.window.progress_bar.hide()
        self.window.status_label.hide()
        self.window.generate_button.setEnabled(True)
        self.window.text_output.append(f"SQL script generated successfully to: {output_path}")
        self.window.text_output.append(f"Total SQL statements generated: {len(sql_lines)}")
        stats_text = (
            f"--- Processing Statistics ---\n"
            f"Source Sheet: {self.window.selected_sheet_name}\n"
            f"Stored Procedure: {self.window.sp_selector.currentText()}\n"
            f"Total Rows in Excel: {stats['total_rows']}\n"
            f"Processed Rows (SQL statements generated): {stats['processed_rows']}\n"
            f"Skipped Rows (Empty/Invalid): {stats['skipped_empty'] + stats['skipped_invalid_value']}\n"
            f"  - Empty/NaN: {stats['skipped_empty']}\n"
            f"  - Invalid Quantity/Value: {stats['skipped_invalid_value']}\n"
            f"Skipped Rows (Arabic Text): {stats['skipped_arabic']}\n"
            f"Processing Time: {stats['processing_time']:.2f} seconds\n"
        )
        self.window.stats_text.setText(stats_text)
        self.window.text_output.append("\n" + stats_text)
        QMessageBox.information(self.window, "Success", f"SQL script generated successfully to:\n{output_path}")
        history_entry = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'file': os.path.basename(self.window.file_path),
            'sheet': self.window.selected_sheet_name,
            'sp_name': self.window.sp_selector.currentText(),
            'processed_rows': stats['processed_rows'],
            'output_path': output_path,
            'status': 'Success'
        }
        self.window.processing_history.append(history_entry)
        self.window.update_history_list()
        self.sql_generator_thread = None
    def on_processing_error(self, message):
        self.window.progress_bar.hide()
        self.window.status_label.hide()
        self.window.generate_button.setEnabled(True)
        self.window.text_output.append(f"Error during processing: {message}")
        QMessageBox.critical(self.window, "Error", f"An error occurred during SQL generation:\n{message}")
        history_entry = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'file': os.path.basename(self.window.file_path) if self.window.file_path else 'N/A',
            'sheet': self.window.selected_sheet_name if self.window.selected_sheet_name else 'N/A',
            'sp_name': self.window.sp_selector.currentText() if self.window.sp_selector.currentText() else 'N/A',
            'processed_rows': 0,
            'output_path': 'N/A',
            'status': f'Failed: {message}'
        }
        self.window.processing_history.append(history_entry)
        self.window.update_history_list()
        self.sql_generator_thread = None

# --- Main Entry Point ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    controller = AppController(main_window)
    main_window.show()
    sys.exit(app.exec_())
