"""
DocuShuttle - Email Forwarding Automation Tool (PyQt5 Version)

This application automates the process of forwarding emails from Outlook's Sent Items
folder based on configurable filters such as date range, subject keywords, and file numbers.

Features:
- Search and filter emails by subject, date range, and file number prefixes
- Forward emails automatically with configurable delays
- Track forwarded emails to prevent duplicates
- Multi-threaded operation for responsive GUI
- SQLite database for configuration and tracking

Author: Royal Payne
License: MIT
"""

import sys
import os
import datetime
import sqlite3
import threading
import re
import time
import json
import subprocess
import shutil
from queue import Queue, Empty
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# PyQt5 imports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTextEdit, QDateEdit,
    QCheckBox, QGroupBox, QTabWidget, QFrame, QMessageBox, QDialog,
    QFormLayout, QSpacerItem, QSizePolicy, QMenu, QAction, QToolButton,
    QTableWidget, QTableWidgetItem
)
from PyQt5.QtCore import Qt, QDate, QTimer, pyqtSignal, QObject, QThread, QPropertyAnimation, QPointF, QRectF, QEasingCurve
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor, QPixmap, QPainter, QPen, QBrush, QPainterPath, QRadialGradient, QLinearGradient
from PyQt5.QtWidgets import QSplashScreen, QProgressBar
import math
import random

# Windows COM integration
import win32com.client
import pythoncom
import pytz

# ============================================================================
# GLOBAL CONFIGURATION
# ============================================================================

# Determine the base path for resources (handles PyInstaller bundled exe)
if getattr(sys, 'frozen', False):
    BASE_PATH = sys._MEIPASS
    # For portable mode, get the directory where the exe is located
    EXE_DIR = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = BASE_PATH

ICON_PATH = os.path.join(BASE_PATH, 'myicon.ico')
ICON_PNG_PATH = os.path.join(BASE_PATH, 'myicon.png')

# Portable mode detection - check for portable.txt in exe directory
PORTABLE_MODE = os.path.exists(os.path.join(EXE_DIR, 'portable.txt'))


def get_app_data_dir():
    """Get the application data directory based on mode (portable or installed)."""
    if PORTABLE_MODE:
        # Portable mode: store data in 'data' subfolder next to exe
        data_dir = os.path.join(EXE_DIR, 'data')
    else:
        # Installed mode: use %LOCALAPPDATA%\DocuShuttle
        localappdata = os.environ.get('LOCALAPPDATA')
        if not localappdata:
            localappdata = os.path.expanduser('~')
        data_dir = os.path.join(localappdata, 'DocuShuttle')

    os.makedirs(data_dir, exist_ok=True)
    return data_dir

# Version and Update Configuration
APP_VERSION = "1.6.4"
GITHUB_REPO = "ProcessLogicLabs/DocuShuttle"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_CHECK_INTERVAL = 86400  # Check once per day (seconds)

# Constants
LOG_BUFFER_SIZE = 10
MAX_LOG_LINES = 1000
DEFAULT_TIMEZONE = 'US/Eastern'

# Thread lock for database access
db_lock = threading.Lock()

# Database path in app data folder (portable or installed)
def get_db_path():
    """Get the path to the database file."""
    try:
        db_dir = get_app_data_dir()
        db_path = os.path.join(db_dir, 'docushuttle.db')
        return db_path
    except Exception as e:
        # Log error and fallback to current directory
        try:
            error_log = os.path.join(get_app_data_dir(), 'error.log')
            with open(error_log, 'a') as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] get_db_path error: {e}\n")
        except:
            pass
        return 'docushuttle.db'

# ============================================================================
# STYLE CONSTANTS - OCRMill Light Theme
# ============================================================================
COLORS = {
    'primary': '#5D9A96',           # Muted teal accent
    'primary_hover': '#4A7B78',     # Darker muted teal for hover
    'primary_light': '#7FB3AF',     # Lighter muted teal for highlights
    'header_bg': '#FFFFFF',         # White header (OCRMill style)
    'header_text': '#5D9A96',       # Muted teal text on header
    'bg': '#F0F0F0',                # Light gray background
    'frame_bg': '#FFFFFF',          # White frame background
    'border': '#CCCCCC',            # Light gray border
    'text': '#333333',              # Dark gray text
    'text_secondary': '#666666',    # Medium gray text
    'input_bg': '#FFFFFF',          # White input background
    'input_border': '#CCCCCC',      # Gray input border
    'success': '#5DAE8B',           # Muted green for success
    'warning': '#D4A056',           # Muted orange for warning
    'tab_inactive': '#F5F5F5',      # Very light gray for inactive tabs
    'status_bar_bg': '#F0F0F0',     # Status bar background
}

STYLESHEET = f"""
QMainWindow {{
    background-color: {COLORS['bg']};
}}

QWidget {{
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 10pt;
}}

/* Header styling */
#headerFrame {{
    background-color: {COLORS['header_bg']};
    border-bottom: 1px solid {COLORS['border']};
    min-height: 60px;
    max-height: 60px;
}}

#brandLabel {{
    color: {COLORS['header_text']};
    font-size: 18pt;
    font-weight: bold;
}}

#brandAccent {{
    color: #9370A2;
    font-size: 18pt;
    font-weight: bold;
}}

/* Tab styling */
QTabWidget::pane {{
    border: 1px solid {COLORS['border']};
    background-color: {COLORS['frame_bg']};
    border-radius: 0px;
    border-top: none;
}}

QTabBar::tab {{
    background-color: {COLORS['tab_inactive']};
    color: {COLORS['text']};
    padding: 8px 20px;
    margin-right: 1px;
    border: 1px solid {COLORS['border']};
    border-bottom: none;
    border-top-left-radius: 3px;
    border-top-right-radius: 3px;
}}

QTabBar::tab:selected {{
    background-color: {COLORS['frame_bg']};
    color: {COLORS['primary']};
    font-weight: bold;
    border-bottom: 2px solid {COLORS['primary']};
}}

QTabBar::tab:hover:!selected {{
    background-color: #E8E8E8;
    color: {COLORS['text']};
}}

/* GroupBox styling */
QGroupBox {{
    background-color: {COLORS['frame_bg']};
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 10px;
    font-weight: bold;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 8px;
    color: {COLORS['primary']};
    background-color: {COLORS['frame_bg']};
}}

/* Input styling */
QLineEdit, QComboBox, QDateEdit {{
    padding: 8px 12px;
    border: 1px solid {COLORS['input_border']};
    border-radius: 4px;
    background-color: {COLORS['input_bg']};
    color: {COLORS['text']};
    min-height: 20px;
}}

QLineEdit:focus, QComboBox:focus, QDateEdit:focus {{
    border: 2px solid {COLORS['primary']};
}}

QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 25px;
    border-left: 1px solid {COLORS['input_border']};
    border-top-right-radius: 4px;
    border-bottom-right-radius: 4px;
    background-color: {COLORS['tab_inactive']};
}}

QComboBox::down-arrow {{
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid {COLORS['text']};
    margin-right: 5px;
}}

QComboBox::down-arrow:hover {{
    border-top-color: {COLORS['primary']};
}}

/* Button styling - OCRMill style */
QPushButton {{
    padding: 6px 16px;
    border-radius: 3px;
    font-weight: normal;
    min-width: 90px;
    border: 1px solid {COLORS['border']};
    background-color: {COLORS['frame_bg']};
    color: {COLORS['text']};
}}

QPushButton:hover {{
    background-color: #E8E8E8;
}}

QPushButton#primaryButton {{
    background-color: {COLORS['frame_bg']};
    color: {COLORS['text']};
    border: 1px solid {COLORS['border']};
}}

QPushButton#primaryButton:hover {{
    background-color: #E8E8E8;
}}

QPushButton#primaryButton:pressed {{
    background-color: #D8D8D8;
}}

QPushButton#primaryButton:disabled {{
    background-color: {COLORS['tab_inactive']};
    color: {COLORS['text_secondary']};
}}

QPushButton#secondaryButton {{
    background-color: {COLORS['frame_bg']};
    color: {COLORS['text']};
    border: 1px solid {COLORS['border']};
}}

QPushButton#secondaryButton:hover {{
    background-color: #E8E8E8;
}}

/* TextEdit styling */
QTextEdit {{
    border: 1px solid {COLORS['input_border']};
    border-radius: 4px;
    background-color: {COLORS['input_bg']};
    color: {COLORS['text']};
    padding: 8px;
}}

/* Menu button */
QToolButton#menuButton {{
    background-color: transparent;
    border: 1px solid {COLORS['border']};
    border-radius: 3px;
    color: {COLORS['text']};
    font-size: 14pt;
    padding: 6px;
}}

QToolButton#menuButton:hover {{
    background-color: #E8E8E8;
}}

/* Label styling */
QLabel {{
    color: {COLORS['text']};
}}

QLabel#sectionLabel {{
    color: {COLORS['text_secondary']};
    font-size: 9pt;
}}

/* Checkbox styling */
QCheckBox {{
    color: {COLORS['text']};
    spacing: 8px;
}}

QCheckBox::indicator {{
    width: 18px;
    height: 18px;
    border-radius: 3px;
    border: 1px solid {COLORS['input_border']};
}}

QCheckBox::indicator:checked {{
    background-color: {COLORS['primary']};
    border-color: {COLORS['primary']};
}}
"""


# ============================================================================
# AUTO-UPDATE SYSTEM
# ============================================================================
class UpdateSignals(QObject):
    """Signals for update checker thread."""
    update_available = pyqtSignal(str, str)  # version, download_url
    update_downloaded = pyqtSignal(str)  # path to downloaded file
    download_progress = pyqtSignal(int, int)  # bytes_downloaded, total_bytes
    update_error = pyqtSignal(str)
    no_update = pyqtSignal()


class UpdateChecker(QThread):
    """Background thread to check for and download updates."""

    def __init__(self, check_only=False):
        super().__init__()
        self.signals = UpdateSignals()
        self.check_only = check_only
        self.download_url = None
        self.new_version = None

    def run(self):
        """Check GitHub for updates and optionally download."""
        try:
            # Check for updates
            request = Request(GITHUB_API_URL)
            request.add_header('User-Agent', f'DocuShuttle/{APP_VERSION}')

            with urlopen(request, timeout=10) as response:
                data = json.loads(response.read().decode('utf-8'))

            latest_version = data.get('tag_name', '').lstrip('v')

            if not latest_version:
                self.signals.no_update.emit()
                return

            # Compare versions
            if self._version_compare(latest_version, APP_VERSION) > 0:
                # Find the exe asset
                assets = data.get('assets', [])
                download_url = None

                for asset in assets:
                    name = asset.get('name', '').lower()
                    if name.endswith('.exe') and 'setup' in name:
                        download_url = asset.get('browser_download_url')
                        break

                if not download_url:
                    # Try to find any exe
                    for asset in assets:
                        if asset.get('name', '').lower().endswith('.exe'):
                            download_url = asset.get('browser_download_url')
                            break

                if download_url:
                    self.new_version = latest_version
                    self.download_url = download_url
                    self.signals.update_available.emit(latest_version, download_url)

                    if not self.check_only:
                        self._download_update(download_url, latest_version)
                else:
                    self.signals.no_update.emit()
            else:
                self.signals.no_update.emit()

        except (URLError, HTTPError) as e:
            self.signals.update_error.emit(f"Network error: {str(e)}")
        except json.JSONDecodeError:
            self.signals.update_error.emit("Invalid response from update server")
        except Exception as e:
            self.signals.update_error.emit(f"Update check failed: {str(e)}")

    def _version_compare(self, v1, v2):
        """Compare two version strings. Returns >0 if v1>v2, <0 if v1<v2, 0 if equal."""
        def normalize(v):
            return [int(x) for x in re.sub(r'[^0-9.]', '', v).split('.')]

        v1_parts = normalize(v1)
        v2_parts = normalize(v2)

        # Pad shorter version with zeros
        max_len = max(len(v1_parts), len(v2_parts))
        v1_parts.extend([0] * (max_len - len(v1_parts)))
        v2_parts.extend([0] * (max_len - len(v2_parts)))

        for i in range(max_len):
            if v1_parts[i] > v2_parts[i]:
                return 1
            elif v1_parts[i] < v2_parts[i]:
                return -1
        return 0

    def _download_update(self, url, version):
        """Download the update installer with progress reporting."""
        filepath = None
        try:
            # Create updates directory in app data
            update_dir = os.path.join(get_app_data_dir(), 'updates')
            os.makedirs(update_dir, exist_ok=True)

            # Download file
            filename = f"DocuShuttle_Setup_v{version}.exe"
            filepath = os.path.join(update_dir, filename)

            # Remove old file if exists
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass

            request = Request(url)
            request.add_header('User-Agent', f'DocuShuttle/{APP_VERSION}')

            with urlopen(request, timeout=60) as response:
                total_size = int(response.headers.get('Content-Length', 0))
                downloaded = 0
                chunk_size = 8192

                with open(filepath, 'wb') as f:
                    while True:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        self.signals.download_progress.emit(downloaded, total_size)

            # Verify download completed successfully
            if filepath and os.path.exists(filepath):
                final_size = os.path.getsize(filepath)
                if total_size > 0 and final_size != total_size:
                    raise Exception(f"Download incomplete: expected {total_size} bytes, got {final_size}")
                self.signals.update_downloaded.emit(filepath)
            else:
                raise Exception("Downloaded file not found after download completed")

        except Exception as e:
            # Clean up partial download
            if filepath and os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass
            self.signals.update_error.emit(f"Download failed: {str(e)}")


def get_last_update_check():
    """Get timestamp of last update check from settings file."""
    settings_path = os.path.join(get_app_data_dir(), 'settings.json')
    try:
        if os.path.exists(settings_path):
            with open(settings_path, 'r') as f:
                settings = json.load(f)
                return settings.get('last_update_check', 0)
    except:
        pass
    return 0


def save_last_update_check():
    """Save timestamp of update check to settings file."""
    settings_path = os.path.join(get_app_data_dir(), 'settings.json')

    try:
        settings = {}

        if os.path.exists(settings_path):
            with open(settings_path, 'r') as f:
                settings = json.load(f)

        settings['last_update_check'] = time.time()

        with open(settings_path, 'w') as f:
            json.dump(settings, f)
    except:
        pass


def get_pending_update():
    """Check if there's a downloaded update waiting to be installed."""
    update_dir = os.path.join(get_app_data_dir(), 'updates')
    if os.path.exists(update_dir):
        for filename in os.listdir(update_dir):
            if filename.endswith('.exe') and 'Setup' in filename:
                return os.path.join(update_dir, filename)
    return None


def clear_pending_updates():
    """Remove any pending update files."""
    update_dir = os.path.join(get_app_data_dir(), 'updates')
    if os.path.exists(update_dir):
        try:
            shutil.rmtree(update_dir)
        except:
            pass


# ============================================================================
# WORKER SIGNALS
# ============================================================================
class WorkerSignals(QObject):
    """Signals for worker threads to communicate with GUI."""
    log_message = pyqtSignal(str)
    display_subject = pyqtSignal(str, str, str)  # subject, recipient, attachments
    operation_complete = pyqtSignal(int, int)
    search_complete = pyqtSignal(int, list)
    error = pyqtSignal(str)
    clear_subjects = pyqtSignal()


# ============================================================================
# DATABASE FUNCTIONS
# ============================================================================
def init_db():
    """Initialize SQLite database and create required tables."""
    db_path = get_db_path()
    new_db = not os.path.exists(db_path)

    # Log database path for debugging
    try:
        error_log_dir = get_app_data_dir()
        error_log = os.path.join(error_log_dir, 'error.log')
        with open(error_log, 'a') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] init_db: Initializing database at {db_path}\n")
    except:
        pass

    try:
        with db_lock:
            with sqlite3.connect(db_path, timeout=10) as conn:
                c = conn.cursor()
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Clients'")
                if not c.fetchone():
                    c.execute('''CREATE TABLE Clients
                                 (recipient TEXT PRIMARY KEY,
                                  start_date TEXT,
                                  end_date TEXT,
                                  file_number_prefix TEXT,
                                  subject_keyword TEXT,
                                  require_attachments TEXT,
                                  skip_forwarded TEXT,
                                  delay_seconds TEXT,
                                  created_at TIMESTAMP,
                                  customer_settings TEXT,
                                  selected_mid_customer TEXT)''')
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ForwardedEmails'")
                if not c.fetchone():
                    c.execute('''CREATE TABLE ForwardedEmails
                                 (file_number TEXT,
                                  recipient TEXT,
                                  forwarded_at TIMESTAMP,
                                  PRIMARY KEY (file_number, recipient))''')
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Settings'")
                if not c.fetchone():
                    c.execute('''CREATE TABLE Settings
                                 (key TEXT PRIMARY KEY,
                                  value TEXT)''')
                conn.commit()

        # Log success
        try:
            with open(error_log, 'a') as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] init_db: Database initialized successfully (new_db={new_db})\n")
        except:
            pass

        return new_db
    except Exception as e:
        # Log error
        try:
            with open(error_log, 'a') as f:
                import traceback
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] init_db ERROR: {str(e)}\n")
                f.write(traceback.format_exc())
                f.write("\n")
        except:
            pass
        raise Exception(f"Error initializing database: {str(e)}")


def load_email_addresses():
    """Load all distinct recipient email addresses from the database."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute("SELECT DISTINCT recipient FROM Clients WHERE recipient IS NOT NULL")
                return [row[0] for row in c.fetchall()]
    except Exception:
        return []


def save_setting(key, value):
    """Save a setting to the Settings table."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute("INSERT OR REPLACE INTO Settings (key, value) VALUES (?, ?)", (key, value))
                conn.commit()
    except Exception:
        pass


def load_setting(key):
    """Load a setting from the Settings table."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute("SELECT value FROM Settings WHERE key = ?", (key,))
                result = c.fetchone()
                return result[0] if result else None
    except Exception:
        return None


def load_config_for_email(recipient):
    """Load configuration for a specific email address."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute('''SELECT start_date, end_date, file_number_prefix, subject_keyword,
                             require_attachments, skip_forwarded, delay_seconds
                             FROM Clients WHERE recipient = ?''', (recipient,))
                return c.fetchone()
    except Exception:
        return None


def save_config(recipient, start_date, end_date, file_number_prefix, subject_keyword,
                require_attachments, skip_forwarded, delay_seconds):
    """Save configuration for a recipient."""
    created_at = datetime.datetime.now(pytz.timezone(DEFAULT_TIMEZONE)).strftime("%Y-%m-%d %H:%M:%S")
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute('''INSERT OR REPLACE INTO Clients
                             (recipient, start_date, end_date, file_number_prefix, subject_keyword,
                              require_attachments, skip_forwarded, delay_seconds, created_at, customer_settings,
                              selected_mid_customer)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                          (recipient, start_date, end_date, file_number_prefix, subject_keyword,
                           "1" if require_attachments else "0", "1" if skip_forwarded else "0",
                           str(delay_seconds), created_at, "", ""))
                conn.commit()
        return True
    except Exception:
        return False


def delete_config(recipient):
    """Delete configuration for a recipient."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute("DELETE FROM Clients WHERE recipient = ?", (recipient,))
                conn.commit()
                return c.rowcount > 0
    except Exception:
        return False


def check_if_forwarded_db(file_number, recipient):
    """Check if file number was previously forwarded."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                c.execute('''SELECT COUNT(*) FROM ForwardedEmails WHERE file_number = ? AND recipient = ?''',
                          (file_number, recipient.lower()))
                return c.fetchone()[0] > 0
    except Exception:
        return False


def log_forwarded_email(file_number, recipient):
    """Log forwarded email to database."""
    try:
        with db_lock:
            with sqlite3.connect(get_db_path(), timeout=10) as conn:
                c = conn.cursor()
                forwarded_at = datetime.datetime.now(pytz.timezone(DEFAULT_TIMEZONE)).strftime("%Y-%m-%d %H:%M:%S")
                c.execute('''INSERT OR REPLACE INTO ForwardedEmails (file_number, recipient, forwarded_at)
                             VALUES (?, ?, ?)''', (file_number, recipient.lower(), forwarded_at))
                conn.commit()
    except Exception:
        pass


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================
def validate_email(email):
    """Validate email address format."""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def sanitize_filter_value(value):
    """Sanitize input for MAPI Restrict filter."""
    if not value:
        return ""
    return value.replace("'", "''").replace("%", "%%")


def convert_date_format(date_str):
    """Convert date between formats."""
    if not date_str or not date_str.strip():
        return None
    try:
        parsed_date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return parsed_date.strftime("%m/%d/%Y")
    except ValueError:
        try:
            datetime.datetime.strptime(date_str, "%m/%d/%Y")
            return date_str
        except ValueError:
            return None


def extract_file_number(item, file_number_prefixes):
    """Extract file number from email."""
    try:
        if item.Attachments.Count > 0:
            attachment = item.Attachments.Item(1)
            filename = os.path.splitext(attachment.FileName)[0]
            for prefix in file_number_prefixes:
                match = re.search(rf'{prefix}\d{{{7-len(prefix)}}}', filename)
                if match:
                    return match.group(0)
        subject = item.Subject if item.Subject else ""
        for prefix in file_number_prefixes:
            match = re.search(rf'{prefix}\d{{{7-len(prefix)}}}', subject)
            if match:
                return match.group(0)
        return None
    except Exception:
        return None


# ============================================================================
# WORKER THREAD
# ============================================================================
class OutlookWorker(QThread):
    """Worker thread for Outlook operations."""

    def __init__(self, config, operation='forward'):
        super().__init__()
        self.config = config
        self.operation = operation
        self.signals = WorkerSignals()
        self.cancel_flag = False

    def cancel(self):
        """Set cancel flag to stop operation."""
        self.cancel_flag = True

    def run(self):
        """Execute the Outlook operation."""
        pythoncom.CoInitialize()
        try:
            if self.operation == 'forward':
                self._forward_emails()
            elif self.operation == 'search':
                self._search_emails()
        finally:
            pythoncom.CoUninitialize()

    def _log(self, message):
        """Emit log message signal."""
        self.signals.log_message.emit(message)

    def _get_outlook_folder(self, mapi):
        """Get Outlook Sent Items folder."""
        try:
            folder = mapi.GetDefaultFolder(5)
            self._log(f"Sent Items folder contains {folder.Items.Count} emails.")
            return folder
        except Exception as e:
            raise Exception(f"Error accessing Sent Items folder: {str(e)}")

    def _search_emails(self):
        """Search for matching emails."""
        try:
            config = self.config
            subject_keyword = config['subject_keyword']
            start_date_str = config['start_date']
            end_date_str = config['end_date']
            skip_forwarded = config['skip_forwarded']
            recipient = config['recipient']
            file_number_prefix = config.get('file_number_prefix', '')
            file_number_prefixes = [p.strip() for p in file_number_prefix.split(',') if p.strip()] if file_number_prefix else []

            local_tz = pytz.timezone(DEFAULT_TIMEZONE)
            start_date = local_tz.localize(datetime.datetime.strptime(start_date_str, "%m/%d/%Y"))
            end_date = local_tz.localize(datetime.datetime.strptime(end_date_str, "%m/%d/%Y") +
                                          datetime.timedelta(days=1) - datetime.timedelta(seconds=1))

            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as e:
                error_code = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                if error_code == -2147221005:
                    raise Exception(
                        "Cannot connect to Outlook. Please ensure:\n\n"
                        "1. Microsoft Outlook is installed\n"
                        "2. Outlook is open and running\n"
                        "3. Python and Outlook are both 32-bit or both 64-bit\n"
                        "4. Try running as Administrator"
                    )
                raise Exception(f"Failed to connect to Outlook: {str(e)}")
            mapi = outlook.GetNamespace("MAPI")
            folder = self._get_outlook_folder(mapi)
            folder.Items.Sort("[SentOn]", True)

            sanitized_subject = sanitize_filter_value(subject_keyword)
            restrict_filter = f"@SQL=\"urn:schemas:httpmail:subject\" ci_phrasematch '{sanitized_subject}'"

            try:
                filtered_items = folder.Items.Restrict(restrict_filter)
                total_emails = filtered_items.Count
            except Exception:
                filtered_items = folder.Items
                total_emails = filtered_items.Count

            self._log(f"Scanning {total_emails} emails...")
            matching_emails = []
            emails_scanned = 0

            for i, item in enumerate(filtered_items, 1):
                if self.cancel_flag:
                    break
                emails_scanned += 1

                if item.Class == 43:
                    try:
                        subject = item.Subject if item.Subject else "(No Subject)"
                        if not subject or subject_keyword.upper() not in subject.upper():
                            continue

                        file_number = None
                        if file_number_prefixes:
                            file_number = extract_file_number(item, file_number_prefixes)
                            if not file_number:
                                continue

                        sent_on = item.SentOn
                        if sent_on < start_date or sent_on > end_date:
                            continue

                        # Use file_number if available, otherwise use EntryID as unique identifier
                        tracking_id = file_number if file_number else item.EntryID

                        if skip_forwarded and check_if_forwarded_db(tracking_id, recipient):
                            continue

                        info = f"[{sent_on.strftime('%Y-%m-%d %H:%M:%S')}] {subject}"
                        if file_number:
                            info += f" (File Number: {file_number})"
                        matching_emails.append(info)
                    except Exception:
                        continue

                if i % 100 == 0:
                    self._log(f"Scanned {i}/{total_emails} emails...")

            self.signals.search_complete.emit(emails_scanned, matching_emails)

        except Exception as e:
            self.signals.error.emit(str(e))

    def _forward_emails(self):
        """Forward matching emails."""
        try:
            config = self.config
            recipient = config['recipient']
            subject_keyword = config['subject_keyword']
            start_date_str = config['start_date']
            end_date_str = config['end_date']
            file_number_prefix = config.get('file_number_prefix', '')
            file_number_prefixes = [p.strip() for p in file_number_prefix.split(',') if p.strip()] if file_number_prefix else []
            require_attachments = config['require_attachments']
            skip_forwarded = config['skip_forwarded']
            delay_seconds = float(config.get('delay_seconds', 0))

            local_tz = pytz.timezone(DEFAULT_TIMEZONE)
            start_date = local_tz.localize(datetime.datetime.strptime(start_date_str, "%m/%d/%Y"))
            end_date = local_tz.localize(datetime.datetime.strptime(end_date_str, "%m/%d/%Y") +
                                          datetime.timedelta(days=1) - datetime.timedelta(seconds=1))

            # Check if date range > 8 days
            date_range_days = (end_date.date() - start_date.date()).days
            if date_range_days > 8:
                delay_seconds = max(delay_seconds, 3.0)
                self._log(f"Date range of {date_range_days} days. Using 3-second delay.")

            self.signals.clear_subjects.emit()

            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as e:
                error_code = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                if error_code == -2147221005:
                    raise Exception(
                        "Cannot connect to Outlook. Please ensure:\n\n"
                        "1. Microsoft Outlook is installed\n"
                        "2. Outlook is open and running\n"
                        "3. Python and Outlook are both 32-bit or both 64-bit\n"
                        "4. Try running as Administrator"
                    )
                raise Exception(f"Failed to connect to Outlook: {str(e)}")
            mapi = outlook.GetNamespace("MAPI")
            self._log(f"Accessing Outlook account: {mapi.CurrentUser.Name}")

            folder = self._get_outlook_folder(mapi)
            folder.Items.Sort("[SentOn]", True)

            sanitized_subject = sanitize_filter_value(subject_keyword)
            restrict_filter = f"@SQL=\"urn:schemas:httpmail:subject\" ci_phrasematch '{sanitized_subject}'"

            try:
                filtered_items = folder.Items.Restrict(restrict_filter)
                total_emails = filtered_items.Count
            except Exception:
                filtered_items = folder.Items
                total_emails = filtered_items.Count

            self._log(f"Scanning {total_emails} emails...")
            emails_processed = 0
            emails_scanned = 0

            for i, item in enumerate(filtered_items, 1):
                if self.cancel_flag:
                    self._log(f"Operation cancelled. Scanned {emails_scanned}, forwarded {emails_processed}.")
                    break

                emails_scanned += 1

                if item.Class == 43:
                    try:
                        subject = item.Subject if item.Subject else "(No Subject)"
                        if not subject or subject_keyword.upper() not in subject.upper():
                            continue

                        file_number = None
                        if file_number_prefixes:
                            file_number = extract_file_number(item, file_number_prefixes)
                            if not file_number:
                                continue

                        sent_on = item.SentOn
                        if sent_on < start_date or sent_on > end_date:
                            continue

                        if require_attachments and item.Attachments.Count == 0:
                            continue

                        # Use file_number if available, otherwise use EntryID as unique identifier
                        tracking_id = file_number if file_number else item.EntryID

                        if skip_forwarded and check_if_forwarded_db(tracking_id, recipient):
                            continue

                        new_subject = file_number if file_number else subject

                        # Collect attachment names
                        attachment_names = []
                        if item.Attachments.Count > 0:
                            for att in item.Attachments:
                                attachment_names.append(att.FileName)
                        attachments_str = ", ".join(attachment_names) if attachment_names else "No attachments"

                        forward_email = item.Forward()
                        forward_email.To = recipient
                        forward_email.Subject = new_subject
                        forward_email.Send()

                        emails_processed += 1
                        self._log(f"Forwarded: {new_subject}")
                        # Show the sent subject (new_subject) in preview
                        self.signals.display_subject.emit(new_subject, recipient, attachments_str)

                        log_forwarded_email(tracking_id, recipient)

                        if delay_seconds > 0:
                            time.sleep(delay_seconds)
                    except Exception as e:
                        self._log(f"Error processing email: {str(e)}")
                        continue

                if i % 100 == 0:
                    self._log(f"Scanned {i}/{total_emails}, forwarded {emails_processed}...")

            self.signals.operation_complete.emit(emails_scanned, emails_processed)

        except Exception as e:
            self.signals.error.emit(str(e))


# ============================================================================
# CONFIGURATION DIALOG
# ============================================================================
class ConfigDialog(QDialog):
    """Configuration dialog with settings and instructions."""

    def __init__(self, parent=None, prefix="", delay="0", require_attach=True, skip_fwd=True, auto_update=False):
        super().__init__(parent)

        try:
            self.setWindowTitle("Configuration & Help")
            self.setFixedSize(520, 480)

            # Try to apply stylesheet
            try:
                self.setStyleSheet(STYLESHEET)
            except Exception as style_error:
                # Log to file if stylesheet fails
                try:
                    error_log = os.path.join(get_app_data_dir(), 'error.log')
                    with open(error_log, 'a') as f:
                        f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Stylesheet error: {style_error}\n")
                except:
                    pass

            layout = QVBoxLayout(self)
            layout.setContentsMargins(15, 15, 15, 15)
            layout.setSpacing(10)

            # Tab widget
            self.tabs = QTabWidget()
            layout.addWidget(self.tabs)

            # === Settings Tab ===
            settings_widget = QWidget()
            settings_layout = QVBoxLayout(settings_widget)
            settings_layout.setContentsMargins(15, 15, 15, 15)
            settings_layout.setSpacing(12)

            # Form layout
            form = QFormLayout()
            form.setSpacing(12)

            self.prefix_edit = QLineEdit(prefix)
            self.prefix_edit.setPlaceholderText("e.g., 759,123")
            self.prefix_edit.setToolTip(
                "Comma-separated list of file number prefixes to filter emails.\n"
                "Only emails with attachments or subjects containing these prefixes will be processed.\n"
                "Leave empty to process all matching emails."
            )
            form.addRow("File Number Prefixes:", self.prefix_edit)

            self.delay_edit = QLineEdit(delay)
            self.delay_edit.setPlaceholderText("Seconds between emails")
            self.delay_edit.setToolTip(
                "Time delay in seconds between forwarding each email.\n"
                "Use this to avoid overwhelming the mail server.\n"
                "Set to 0 for no delay."
            )
            form.addRow("Delay (Sec.):", self.delay_edit)

            self.require_attach_check = QCheckBox()
            self.require_attach_check.setChecked(require_attach)
            self.require_attach_check.setToolTip(
                "When checked, only emails with attachments will be forwarded.\n"
                "Uncheck to forward emails regardless of attachments."
            )
            form.addRow("Require Attachments:", self.require_attach_check)

            self.skip_fwd_check = QCheckBox()
            self.skip_fwd_check.setChecked(skip_fwd)
            self.skip_fwd_check.setToolTip(
                "When checked, emails that have already been forwarded will be skipped.\n"
                "This prevents duplicate forwards using the tracking database.\n"
                "Uncheck to re-forward previously forwarded emails."
            )
            form.addRow("Skip Previously Forwarded:", self.skip_fwd_check)

            self.auto_update_check = QCheckBox()
            self.auto_update_check.setChecked(auto_update)
            self.auto_update_check.setToolTip(
                "When checked, updates will be downloaded and installed automatically.\n"
                "The app will close and restart with the new version.\n"
                "Uncheck to be prompted before installing updates."
            )
            form.addRow("Auto-Install Updates:", self.auto_update_check)

            settings_layout.addLayout(form)
            settings_layout.addStretch()
            self.tabs.addTab(settings_widget, "Settings")

            # === Setup Instructions Tab ===
            setup_widget = QWidget()
            setup_layout = QVBoxLayout(setup_widget)
            setup_layout.setContentsMargins(15, 15, 15, 15)

            setup_text = QTextEdit()
            setup_text.setReadOnly(True)
            setup_text.setHtml("""
            <h3 style="color: #5D9A96; margin-bottom: 10px;">Initial Setup</h3>

            <p><b>1. Outlook Requirements:</b></p>
            <ul>
                <li>Microsoft Outlook must be installed and configured</li>
                <li>You must have at least one email account set up in Outlook</li>
                <li>Outlook should be running or able to start automatically</li>
            </ul>

            <p><b>2. Add Recipient Emails:</b></p>
            <ul>
                <li>Click the <b>Manage Emails</b> button in the main window</li>
                <li>Enter recipient email addresses (one per line or comma-separated)</li>
                <li>These are the addresses emails will be forwarded TO</li>
            </ul>

            <p><b>3. Configure Settings (Optional):</b></p>
            <ul>
                <li><b>File Number Prefixes:</b> Filter emails by file number (e.g., "759,123")</li>
                <li><b>Delay:</b> Set seconds between forwards to avoid rate limits</li>
                <li><b>Require Attachments:</b> Only forward emails with attachments</li>
                <li><b>Skip Previously Forwarded:</b> Prevent duplicate forwards</li>
            </ul>

            <p><b>4. First Run:</b></p>
            <ul>
                <li>Outlook may prompt you to allow access - click <b>Allow</b></li>
                <li>If using Exchange, ensure you have proper permissions</li>
            </ul>
            """)
            setup_layout.addWidget(setup_text)
            self.tabs.addTab(setup_widget, "Setup Instructions")

            # === Usage Instructions Tab ===
            usage_widget = QWidget()
            usage_layout = QVBoxLayout(usage_widget)
            usage_layout.setContentsMargins(15, 15, 15, 15)

            usage_text = QTextEdit()
            usage_text.setReadOnly(True)
            usage_text.setHtml("""
            <h3 style="color: #5D9A96; margin-bottom: 10px;">How to Use DocuShuttle</h3>

            <p><b>Basic Workflow:</b></p>
            <ol>
                <li>Select a <b>Recipient Email</b> from the dropdown (destination)</li>
                <li>Enter a <b>Subject Filter</b> to match emails (e.g., "Invoice")</li>
                <li>Set the <b>Date Range</b> for emails to search</li>
                <li>Click <b>Search & Forward</b> to process matching emails</li>
            </ol>

            <p><b>Understanding the Interface:</b></p>
            <ul>
                <li><b>Recipient Email:</b> Where forwarded emails will be sent</li>
                <li><b>Subject Filter:</b> Text to match in email subjects</li>
                <li><b>Start/End Date:</b> Date range to search in Sent Items</li>
                <li><b>Log Window:</b> Shows progress and results of operations</li>
            </ul>

            <p><b>Tips:</b></p>
            <ul>
                <li>Use specific subject filters to avoid forwarding unwanted emails</li>
                <li>Enable "Skip Previously Forwarded" to prevent duplicates</li>
                <li>Check the log window for detailed operation status</li>
                <li>Use "Require Attachments" if you only want document emails</li>
            </ul>

            <p><b>Troubleshooting:</b></p>
            <ul>
                <li><b>No emails found:</b> Check date range and subject filter</li>
                <li><b>Outlook errors:</b> Ensure Outlook is running and accessible</li>
                <li><b>Permission denied:</b> Allow DocuShuttle access in Outlook prompts</li>
            </ul>
            """)
            usage_layout.addWidget(usage_text)
            self.tabs.addTab(usage_widget, "Usage Instructions")

            # Buttons
            btn_layout = QHBoxLayout()
            btn_layout.addStretch()

            save_btn = QPushButton("Save")
            save_btn.setObjectName("primaryButton")
            save_btn.clicked.connect(self.accept)
            btn_layout.addWidget(save_btn)

            cancel_btn = QPushButton("Cancel")
            cancel_btn.setObjectName("secondaryButton")
            cancel_btn.clicked.connect(self.reject)
            btn_layout.addWidget(cancel_btn)

            layout.addLayout(btn_layout)

        except Exception as e:
            # Log critical error to file
            try:
                error_log = os.path.join(get_app_data_dir(), 'error.log')
                with open(error_log, 'a') as f:
                    import traceback
                    f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ConfigDialog error:\n")
                    f.write(traceback.format_exc())
                    f.write("\n")
            except:
                pass
            raise

    def get_values(self):
        """Return dialog values."""
        return {
            'prefix': self.prefix_edit.text(),
            'delay': self.delay_edit.text(),
            'require_attachments': self.require_attach_check.isChecked(),
            'skip_forwarded': self.skip_fwd_check.isChecked(),
            'auto_update': self.auto_update_check.isChecked()
        }


# ============================================================================
# UPDATE PROGRESS DIALOG
# ============================================================================
class UpdateProgressDialog(QDialog):
    """Progress dialog for update downloads."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Downloading Update")
        self.setFixedSize(400, 150)
        self.setStyleSheet(STYLESHEET)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.CustomizeWindowHint)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Status label
        self.status_label = QLabel("Downloading update...")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # Details label
        self.details_label = QLabel("")
        self.details_label.setAlignment(Qt.AlignCenter)
        self.details_label.setStyleSheet("color: #7D8A96; font-size: 10px;")
        layout.addWidget(self.details_label)

        layout.addStretch()

    def update_progress(self, downloaded, total):
        """Update progress bar with download progress."""
        if total > 0:
            percentage = int((downloaded / total) * 100)
            self.progress_bar.setValue(percentage)

            # Format sizes
            downloaded_mb = downloaded / (1024 * 1024)
            total_mb = total / (1024 * 1024)
            self.details_label.setText(f"{downloaded_mb:.1f} MB / {total_mb:.1f} MB")
        else:
            self.details_label.setText(f"{downloaded / (1024 * 1024):.1f} MB downloaded")

    def set_installing(self):
        """Change dialog to show installing status."""
        self.status_label.setText("Installing update...")
        self.progress_bar.setMaximum(0)  # Indeterminate progress
        self.details_label.setText("Application will restart automatically")


# ============================================================================
# MAIN WINDOW
# ============================================================================
class DocuShuttleWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.config_prefix = "76"
        self.config_delay = "0"
        self.config_require_attachments = True
        self.config_skip_forwarded = True
        self.config_auto_update = True  # Default to auto-update enabled
        self.update_checker = None
        self.pending_update_path = None
        self.progress_dialog = None

        self.init_ui()
        self.load_saved_state()
        # Delay update check until after window is shown
        QTimer.singleShot(1000, self.check_for_updates_on_startup)

    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("DocuShuttle")
        self.setMinimumSize(650, 600)
        self.resize(700, 650)

        # Set window icon
        if os.path.exists(ICON_PNG_PATH):
            self.setWindowIcon(QIcon(ICON_PNG_PATH))
        elif os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.setStyleSheet(STYLESHEET)

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Header
        header = QFrame()
        header.setObjectName("headerFrame")
        header.setFixedHeight(60)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 0, 20, 0)

        # Brand with logo icon
        brand_layout = QHBoxLayout()
        brand_layout.setSpacing(8)

        # Add logo icon (load from myicon.png)
        logo_label = QLabel()
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'myicon.png')
        if os.path.exists(icon_path):
            logo_pixmap = QPixmap(icon_path).scaled(36, 36, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(logo_pixmap)
        logo_label.setFixedSize(40, 40)
        brand_layout.addWidget(logo_label)

        # Brand text
        brand_text_layout = QHBoxLayout()
        brand_text_layout.setSpacing(0)
        brand_label = QLabel("Docu")
        brand_label.setObjectName("brandLabel")
        brand_text_layout.addWidget(brand_label)
        accent_label = QLabel("Shuttle")
        accent_label.setObjectName("brandAccent")
        brand_text_layout.addWidget(accent_label)
        brand_layout.addLayout(brand_text_layout)

        header_layout.addLayout(brand_layout)

        header_layout.addStretch()

        # Hamburger menu button in header for Configuration
        self.config_menu_btn = QToolButton()
        self.config_menu_btn.setText("☰")
        self.config_menu_btn.setFixedSize(36, 36)
        self.config_menu_btn.setStyleSheet(f"""
            QToolButton {{
                background-color: transparent;
                border: 1px solid {COLORS['border']};
                border-radius: 4px;
                font-size: 16pt;
                color: {COLORS['text']};
            }}
            QToolButton:hover {{
                background-color: #E8E8E8;
                border: 1px solid {COLORS['border']};
            }}
        """)
        self.config_menu_btn.setPopupMode(QToolButton.InstantPopup)

        # Create menu for config button
        config_menu = QMenu(self.config_menu_btn)
        config_menu.setStyleSheet(f"""
            QMenu {{
                background-color: {COLORS['frame_bg']};
                border: 1px solid {COLORS['border']};
                padding: 5px;
            }}
            QMenu::item {{
                padding: 8px 20px;
            }}
            QMenu::item:selected {{
                background-color: {COLORS['primary']};
                color: white;
            }}
        """)

        config_action = config_menu.addAction("Configuration...")
        config_action.triggered.connect(self.show_config_dialog)

        config_menu.addSeparator()

        check_update_action = config_menu.addAction("Check for Updates...")
        check_update_action.triggered.connect(self.manual_check_for_updates)

        about_action = config_menu.addAction(f"About DocuShuttle v{APP_VERSION}")
        about_action.triggered.connect(self.show_about_dialog)

        self.config_menu_btn.setMenu(config_menu)
        header_layout.addWidget(self.config_menu_btn)

        main_layout.addWidget(header)

        # Content area
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(15, 15, 15, 15)

        # Tab widget
        tabs = QTabWidget()
        content_layout.addWidget(tabs)

        # Search tab
        search_tab = QWidget()
        search_layout = QVBoxLayout(search_tab)
        search_layout.setContentsMargins(10, 15, 10, 10)
        search_layout.setSpacing(12)

        # Email Settings group
        email_group = QGroupBox("Email Settings")
        email_layout = QFormLayout(email_group)
        email_layout.setContentsMargins(15, 20, 15, 15)
        email_layout.setSpacing(12)

        # Forward To combobox with right-click context menu
        self.recipient_combo = QComboBox()
        self.recipient_combo.setEditable(True)
        self.recipient_combo.setMinimumWidth(320)
        self.recipient_combo.currentTextChanged.connect(self.on_recipient_changed)
        self.recipient_combo.setContextMenuPolicy(Qt.CustomContextMenu)
        self.recipient_combo.customContextMenuRequested.connect(self.show_email_context_menu)

        email_layout.addRow("Forward To:", self.recipient_combo)

        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("e.g., BILLING INVOICE")
        self.subject_edit.setText("BILLING INVOICE")
        email_layout.addRow("Subject Keyword:", self.subject_edit)

        search_layout.addWidget(email_group)

        # Date Range group
        date_group = QGroupBox("Date Range")
        date_layout = QHBoxLayout(date_group)
        date_layout.setContentsMargins(15, 20, 15, 15)
        date_layout.setSpacing(20)

        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Start Date:"))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setDisplayFormat("MM/dd/yyyy")
        start_layout.addWidget(self.start_date)
        date_layout.addLayout(start_layout)

        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("End Date:"))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setDisplayFormat("MM/dd/yyyy")
        end_layout.addWidget(self.end_date)
        date_layout.addLayout(end_layout)

        date_layout.addStretch()
        search_layout.addWidget(date_group)

        # Files Sent group
        files_group = QGroupBox("Files Sent")
        files_layout = QVBoxLayout(files_group)
        files_layout.setContentsMargins(15, 20, 15, 15)

        # Create table for Files Sent
        self.files_table = QTableWidget()
        self.files_table.setColumnCount(4)
        self.files_table.setHorizontalHeaderLabels(["Date/Time", "Sent Subject", "To", "Attachments"])

        # Set column widths
        self.files_table.setColumnWidth(0, 180)  # Date/Time
        self.files_table.setColumnWidth(1, 250)  # Sent Subject
        self.files_table.setColumnWidth(2, 220)  # To
        self.files_table.setColumnWidth(3, 200)  # Attachments

        # Table properties
        self.files_table.setAlternatingRowColors(True)
        self.files_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.files_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.files_table.setMinimumHeight(120)
        self.files_table.setSortingEnabled(True)
        self.files_table.verticalHeader().setVisible(False)

        files_layout.addWidget(self.files_table)

        search_layout.addWidget(files_group)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.preview_btn = QPushButton("Preview")
        self.preview_btn.setObjectName("secondaryButton")
        self.preview_btn.clicked.connect(self.preview_emails)
        btn_layout.addWidget(self.preview_btn)

        self.forward_btn = QPushButton("Scan and Forward")
        self.forward_btn.setObjectName("primaryButton")
        self.forward_btn.clicked.connect(self.scan_and_forward)
        btn_layout.addWidget(self.forward_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setObjectName("secondaryButton")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_operation)
        btn_layout.addWidget(self.cancel_btn)

        btn_layout.addStretch()
        search_layout.addLayout(btn_layout)

        tabs.addTab(search_tab, "  Search  ")

        # Log tab
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        log_layout.setContentsMargins(10, 15, 10, 10)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)

        tabs.addTab(log_tab, "  Log  ")

        main_layout.addWidget(content)

        # Initialize database and load emails
        init_db()
        self.refresh_email_list()

    def refresh_email_list(self):
        """Refresh the email combobox."""
        current = self.recipient_combo.currentText()
        self.recipient_combo.clear()
        emails = load_email_addresses()
        self.recipient_combo.addItems(emails)
        if current and current in emails:
            self.recipient_combo.setCurrentText(current)

    def load_saved_state(self):
        """Load saved application state."""
        last_email = load_setting('last_used_email')
        if last_email:
            idx = self.recipient_combo.findText(last_email)
            if idx >= 0:
                self.recipient_combo.setCurrentIndex(idx)

        last_start = load_setting('last_start_date')
        last_end = load_setting('last_end_date')
        if last_start:
            try:
                date = QDate.fromString(last_start, "MM/dd/yyyy")
                if date.isValid():
                    self.start_date.setDate(date)
            except Exception:
                pass
        if last_end:
            try:
                date = QDate.fromString(last_end, "MM/dd/yyyy")
                if date.isValid():
                    self.end_date.setDate(date)
            except Exception:
                pass

        # Load auto-update setting
        auto_update = load_setting('auto_update')
        if auto_update is not None:
            # Convert string to boolean (SQLite stores as string)
            if isinstance(auto_update, str):
                self.config_auto_update = auto_update.lower() in ('true', '1', 'yes')
            else:
                self.config_auto_update = bool(auto_update)

    def on_recipient_changed(self, text):
        """Handle recipient selection change."""
        if not text:
            return

        save_setting('last_used_email', text)

        config = load_config_for_email(text)
        if config:
            start_date, end_date, prefix, keyword, req_attach, skip_fwd, delay = config

            if start_date:
                converted = convert_date_format(start_date)
                if converted:
                    date = QDate.fromString(converted, "MM/dd/yyyy")
                    if date.isValid():
                        self.start_date.setDate(date)

            if end_date:
                converted = convert_date_format(end_date)
                if converted:
                    date = QDate.fromString(converted, "MM/dd/yyyy")
                    if date.isValid():
                        self.end_date.setDate(date)

            self.config_prefix = prefix or ""
            self.subject_edit.setText(keyword or "BILLING INVOICE")
            self.config_require_attachments = req_attach == "1"
            self.config_skip_forwarded = skip_fwd == "1"
            self.config_delay = delay or "0"

            self.log(f"Loaded configuration for '{text}'")

    def log(self, message):
        """Add message to log."""
        timestamp = datetime.datetime.now(pytz.timezone(DEFAULT_TIMEZONE)).strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

    def show_config_dialog(self):
        """Show configuration dialog."""
        try:
            dialog = ConfigDialog(
                self,
                self.config_prefix,
                self.config_delay,
                self.config_require_attachments,
                self.config_skip_forwarded,
                self.config_auto_update
            )

            if dialog.exec_() == QDialog.Accepted:
                values = dialog.get_values()
                self.config_prefix = values['prefix']
                self.config_delay = values['delay']
                self.config_require_attachments = values['require_attachments']
                self.config_skip_forwarded = values['skip_forwarded']
                self.config_auto_update = values['auto_update']
                # Save as string 'True' or 'False' for SQLite
                save_setting('auto_update', str(self.config_auto_update))
                self.log("Configuration updated")
        except Exception as e:
            # Log error and show user-friendly message
            error_msg = f"Failed to open configuration dialog: {str(e)}"
            self.log(error_msg)
            try:
                error_log = os.path.join(get_app_data_dir(), 'error.log')
                with open(error_log, 'a') as f:
                    import traceback
                    f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] show_config_dialog error:\n")
                    f.write(traceback.format_exc())
                    f.write("\n")
            except:
                pass
            QMessageBox.critical(
                self, "Configuration Error",
                f"Failed to open configuration dialog.\n\nError: {str(e)}\n\n"
                f"Check error.log in {get_app_data_dir()}"
            )

    def show_email_context_menu(self, position):
        """Show right-click context menu for email combobox."""
        context_menu = QMenu(self)
        context_menu.setStyleSheet(f"""
            QMenu {{
                background-color: {COLORS['frame_bg']};
                border: 1px solid {COLORS['border']};
                padding: 5px;
            }}
            QMenu::item {{
                padding: 8px 20px;
            }}
            QMenu::item:selected {{
                background-color: {COLORS['primary']};
                color: white;
            }}
        """)

        delete_action = context_menu.addAction("Delete Email")
        delete_action.triggered.connect(self.delete_current_config)

        context_menu.exec_(self.recipient_combo.mapToGlobal(position))

    def delete_current_config(self):
        """Delete current email configuration."""
        recipient = self.recipient_combo.currentText().strip()
        if not recipient:
            QMessageBox.warning(self, "Warning", "No email selected to delete.")
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            f"Delete configuration for '{recipient}'?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            if delete_config(recipient):
                self.log(f"Deleted configuration for '{recipient}'")
                self.refresh_email_list()
                self.recipient_combo.setCurrentText("")
            else:
                QMessageBox.warning(self, "Error", "Failed to delete configuration.")

    def validate_inputs(self):
        """Validate form inputs."""
        recipient = self.recipient_combo.currentText().strip()
        if not recipient or not validate_email(recipient):
            QMessageBox.warning(self, "Error", "Please enter a valid email address.")
            return False

        if not self.subject_edit.text().strip():
            QMessageBox.warning(self, "Error", "Subject keyword is required.")
            return False

        return True

    def get_config(self):
        """Get current configuration."""
        return {
            'recipient': self.recipient_combo.currentText().strip(),
            'subject_keyword': self.subject_edit.text().strip(),
            'start_date': self.start_date.date().toString("MM/dd/yyyy"),
            'end_date': self.end_date.date().toString("MM/dd/yyyy"),
            'file_number_prefix': self.config_prefix,
            'require_attachments': self.config_require_attachments,
            'skip_forwarded': self.config_skip_forwarded,
            'delay_seconds': self.config_delay
        }

    def set_buttons_enabled(self, enabled):
        """Enable/disable action buttons."""
        self.preview_btn.setEnabled(enabled)
        self.forward_btn.setEnabled(enabled)
        self.cancel_btn.setEnabled(not enabled)
        self.recipient_combo.setEnabled(enabled)

    def preview_emails(self):
        """Preview matching emails."""
        if not self.validate_inputs():
            return

        config = self.get_config()
        self.set_buttons_enabled(False)
        self.log("Starting email preview...")

        self.worker = OutlookWorker(config, 'search')
        self.worker.signals.log_message.connect(self.log)
        self.worker.signals.search_complete.connect(self.on_search_complete)
        self.worker.signals.error.connect(self.on_error)
        self.worker.finished.connect(lambda: self.set_buttons_enabled(True))
        self.worker.start()

    def on_search_complete(self, scanned, results):
        """Handle search completion."""
        msg = f"Found {len(results)} matching emails (scanned {scanned})"
        self.log(msg)

        if results:
            text = "\n".join(results)
            QMessageBox.information(self, "Preview Results", f"{msg}\n\n{text[:2000]}...")
        else:
            QMessageBox.information(self, "Preview Results", "No matching emails found.")

    def scan_and_forward(self):
        """Scan and forward matching emails."""
        if not self.validate_inputs():
            return

        # Prompt user to configure prefix if not set
        if not self.config_prefix.strip():
            reply = QMessageBox.question(
                self, "Configure File Number Prefix?",
                "No file number prefix is configured.\n\n"
                "Without a prefix, emails will be tracked by their unique ID, "
                "and the original subject line will be preserved.\n\n"
                "Would you like to configure a file number prefix now?",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            if reply == QMessageBox.Yes:
                self.show_config_dialog()
                return
            elif reply == QMessageBox.Cancel:
                return
            # If No, continue without prefix

        config = self.get_config()

        # Save configuration
        save_config(
            config['recipient'],
            config['start_date'],
            config['end_date'],
            config['file_number_prefix'],
            config['subject_keyword'],
            config['require_attachments'],
            config['skip_forwarded'],
            float(config['delay_seconds']) if config['delay_seconds'] else 0
        )

        save_setting('last_start_date', config['start_date'])
        save_setting('last_end_date', config['end_date'])

        self.refresh_email_list()
        self.set_buttons_enabled(False)
        self.log("Starting forward operation...")

        self.worker = OutlookWorker(config, 'forward')
        self.worker.signals.log_message.connect(self.log)
        self.worker.signals.display_subject.connect(self.display_subject)
        self.worker.signals.clear_subjects.connect(lambda: self.files_table.setRowCount(0))
        self.worker.signals.operation_complete.connect(self.on_forward_complete)
        self.worker.signals.error.connect(self.on_error)
        self.worker.finished.connect(lambda: self.set_buttons_enabled(True))
        self.worker.start()

    def display_subject(self, subject, recipient, attachments):
        """Display forwarded email details in table."""
        # Get current timestamp
        timestamp = datetime.datetime.now(pytz.timezone(DEFAULT_TIMEZONE)).strftime("%Y-%m-%d %H:%M:%S")

        # Disable sorting while adding row
        self.files_table.setSortingEnabled(False)

        # Add new row
        row_position = self.files_table.rowCount()
        self.files_table.insertRow(row_position)

        # Add data to columns
        self.files_table.setItem(row_position, 0, QTableWidgetItem(timestamp))
        self.files_table.setItem(row_position, 1, QTableWidgetItem(subject))
        self.files_table.setItem(row_position, 2, QTableWidgetItem(recipient))
        self.files_table.setItem(row_position, 3, QTableWidgetItem(attachments))

        # Re-enable sorting
        self.files_table.setSortingEnabled(True)

        # Scroll to the new row
        self.files_table.scrollToItem(self.files_table.item(row_position, 0))

    def on_forward_complete(self, scanned, forwarded):
        """Handle forward completion."""
        msg = f"Scanned {scanned} emails, forwarded {forwarded} emails."
        self.log(msg)
        QMessageBox.information(self, "Complete", msg)

    def on_error(self, error_msg):
        """Handle worker error."""
        self.log(f"Error: {error_msg}")
        QMessageBox.critical(self, "Error", error_msg)

    def cancel_operation(self):
        """Cancel current operation."""
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.log("Cancellation requested...")

    # ========================================================================
    # AUTO-UPDATE METHODS
    # ========================================================================
    def check_for_updates_on_startup(self):
        """Check for updates silently on startup."""
        # Check if enough time has passed since last check
        last_check = get_last_update_check()
        current_time = time.time()

        if current_time - last_check < UPDATE_CHECK_INTERVAL:
            # Check if there's a pending update
            pending = get_pending_update()
            if pending and os.path.exists(pending):
                self.prompt_install_update(pending)
            return

        # Start background update check
        self.start_update_check(silent=True)

    def manual_check_for_updates(self):
        """Manually trigger update check from menu."""
        self.log("Checking for updates...")
        self.start_update_check(silent=False)

    def start_update_check(self, silent=True):
        """Start the update checker thread."""
        if self.update_checker and self.update_checker.isRunning():
            return

        self.update_checker = UpdateChecker(check_only=True)
        self.update_checker.signals.update_available.connect(
            lambda ver, url: self.on_update_available(ver, url, silent))
        self.update_checker.signals.update_downloaded.connect(self.on_update_downloaded)
        self.update_checker.signals.update_error.connect(
            lambda err: self.on_update_error(err, silent))
        self.update_checker.signals.no_update.connect(
            lambda: self.on_no_update(silent))
        self.update_checker.start()

    def on_update_available(self, version, download_url, silent):
        """Handle update available signal."""
        save_last_update_check()

        if silent:
            # Silently download the update
            self.log(f"New version {version} available, downloading...")
            self.download_update(download_url, version)
        else:
            # Ask user if they want to download
            reply = QMessageBox.question(
                self, "Update Available",
                f"A new version ({version}) is available!\n\n"
                f"Would you like to download and install it?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.log(f"Downloading update {version}...")
                self.download_update(download_url, version)

    def download_update(self, url, version):
        """Download update in background with progress dialog."""
        # Create and show progress dialog
        self.progress_dialog = UpdateProgressDialog(self)
        self.progress_dialog.show()

        self.update_checker = UpdateChecker(check_only=False)
        self.update_checker.download_url = url
        self.update_checker.new_version = version
        self.update_checker.signals.download_progress.connect(self.on_download_progress)
        self.update_checker.signals.update_downloaded.connect(self.on_update_downloaded)
        self.update_checker.signals.update_error.connect(
            lambda err: self.on_update_error(err, False))
        self.update_checker.start()

    def on_download_progress(self, downloaded, total):
        """Update progress dialog with download progress."""
        if self.progress_dialog:
            self.progress_dialog.update_progress(downloaded, total)

    def on_update_downloaded(self, file_path):
        """Handle update downloaded signal."""
        self.pending_update_path = file_path
        self.log(f"Update downloaded: {file_path}")

        # Verify file exists and is valid
        if not os.path.exists(file_path):
            self.log(f"Error: Downloaded file not found at {file_path}")
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None
            QMessageBox.critical(
                self, "Update Error",
                f"Downloaded file not found: {file_path}"
            )
            return

        file_size = os.path.getsize(file_path)
        self.log(f"Update file size: {file_size / (1024*1024):.2f} MB")

        if self.config_auto_update:
            # Auto-install without prompting
            self.log("Auto-installing update...")
            if self.progress_dialog:
                self.progress_dialog.set_installing()
            self.install_update(file_path)
        else:
            # Close progress dialog
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None
            # Prompt user
            self.prompt_install_update(file_path)

    def prompt_install_update(self, file_path):
        """Prompt user to install the downloaded update."""
        reply = QMessageBox.question(
            self, "Update Ready",
            "A new update has been downloaded and is ready to install.\n\n"
            "The application will close and the installer will run.\n"
            "Would you like to install it now?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.install_update(file_path)

    def install_update(self, file_path):
        """Launch the installer and close the app."""
        try:
            # Close progress dialog first
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None

            # Verify file exists before proceeding
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Installer not found: {file_path}")

            # Log the action
            self.log(f"Launching installer: {file_path}")

            # Show message to user
            msg = QMessageBox(self)
            msg.setWindowTitle("Installing Update")
            msg.setText("The installer will launch now.\n\nPlease wait for the installation to complete.")
            msg.setIcon(QMessageBox.Information)
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

            # Launch the installer directly
            subprocess.Popen([file_path])

            self.log("Installer launched, closing application...")

            # Close the application
            QTimer.singleShot(500, QApplication.quit)

        except Exception as e:
            self.log(f"Installer launch error: {str(e)}")
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None
            QMessageBox.critical(
                self, "Update Error",
                f"Failed to launch installer:\n{str(e)}"
            )

    def on_update_error(self, error, silent):
        """Handle update error signal."""
        save_last_update_check()

        # Close progress dialog if open
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None

        if not silent:
            QMessageBox.warning(
                self, "Update Check Failed",
                f"Could not check for updates:\n{error}"
            )
        else:
            self.log(f"Update check failed: {error}")

    def on_no_update(self, silent):
        """Handle no update available signal."""
        save_last_update_check()
        if not silent:
            QMessageBox.information(
                self, "No Updates",
                f"You are running the latest version (v{APP_VERSION})."
            )
        else:
            self.log("No updates available")

    def show_about_dialog(self):
        """Show about dialog."""
        QMessageBox.about(
            self, "About DocuShuttle",
            f"<h2>DocuShuttle</h2>"
            f"<p>Version {APP_VERSION}</p>"
            f"<p>Email forwarding automation for Microsoft Outlook.</p>"
            f"<p>&copy; 2024 Process Logic Labs</p>"
            f"<p><a href='https://github.com/ProcessLogicLabs/DocuShuttle'>GitHub Repository</a></p>"
        )


# ============================================================================
# ANIMATED SPLASH SCREEN (Premium Design)
# ============================================================================
class AnimatedSplashScreen(QWidget):
    """Premium animated splash screen for DocuShuttle."""

    def __init__(self):
        super().__init__(None)

        # Timing
        self.start_time = time.time()
        self.fade_opacity = 1.0
        self.is_fading = False

        # Progress
        self.progress = 0
        self._target_progress = 0
        self._message = "Initializing..."

        # Animation states
        self.intro_progress = 0.0
        self.ring_rotation = 0.0
        self.pulse_phase = 0.0
        self.wave_offset = 0.0

        # Window setup
        self.splash_width = 540
        self.splash_height = 340
        self.setFixedSize(self.splash_width, self.splash_height)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_OpaquePaintEvent, True)
        self.setAutoFillBackground(False)

        # Animation timer
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._animate)
        self.timer.start(16)  # ~60 FPS

        # Progress timer
        self.progress_timer = QTimer(self)
        self.progress_timer.timeout.connect(self._update_progress)
        self.progress_timer.start(50)

        # Center on screen
        screen = QApplication.primaryScreen().geometry()
        self.move(
            (screen.width() - self.splash_width) // 2,
            (screen.height() - self.splash_height) // 2
        )

    def _animate(self):
        """Update animations."""
        elapsed = time.time() - self.start_time

        # Fade out
        if self.is_fading:
            self.fade_opacity = max(0, self.fade_opacity - 0.04)
            if self.fade_opacity <= 0:
                self.timer.stop()
                self.close()
                return

        # Intro animation (0 to 1.0s)
        if elapsed < 1.0:
            t = elapsed / 1.0
            self.intro_progress = 1 if t == 1 else 1 - pow(2, -10 * t)
        else:
            self.intro_progress = 1.0

        # Continuous animations
        self.ring_rotation = elapsed * 30
        self.pulse_phase = elapsed * 2.5
        self.wave_offset = elapsed * 80

        # Smooth progress (snap to target when close)
        diff = self._target_progress - self.progress
        if abs(diff) < 0.5:
            self.progress = self._target_progress
        else:
            self.progress += diff * 0.15

        self.update()

    def _update_progress(self):
        """Update progress bar."""
        if self._target_progress < 100:
            self._target_progress += 2
            # Update messages based on progress
            if self._target_progress < 20:
                self._message = "Initializing..."
            elif self._target_progress < 40:
                self._message = "Loading configuration..."
            elif self._target_progress < 60:
                self._message = "Connecting to Outlook..."
            elif self._target_progress < 80:
                self._message = "Loading email data..."
            else:
                self._message = "Almost ready..."
        else:
            self._message = "Ready!"
            self.progress_timer.stop()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)

        # Apply fade opacity
        if self.is_fading:
            painter.setOpacity(self.fade_opacity)

        # Draw solid background
        painter.fillRect(self.rect(), QColor(15, 23, 42))

        self._draw_background(painter)
        self._draw_orbital_rings(painter)
        self._draw_center_emblem(painter)
        self._draw_title(painter)
        self._draw_tagline(painter)
        self._draw_progress_area(painter)
        self._draw_corner_accents(painter)

        painter.end()

    def _draw_background(self, painter):
        """Draw premium gradient background."""
        painter.save()

        # Rich gradient background
        bg = QLinearGradient(0, 0, self.width(), self.height())
        bg.setColorAt(0, QColor(15, 23, 42))
        bg.setColorAt(0.5, QColor(30, 41, 59))
        bg.setColorAt(1, QColor(15, 23, 42))

        painter.setBrush(bg)
        painter.setPen(Qt.NoPen)
        painter.drawRect(self.rect())

        # Subtle top glow (teal for DocuShuttle)
        glow_rect = QRectF(0, 0, self.width(), 140)
        glow = QLinearGradient(0, 0, 0, 140)
        glow.setColorAt(0, QColor(93, 154, 150, 30))  # Muted teal
        glow.setColorAt(1, QColor(93, 154, 150, 0))
        painter.setBrush(glow)
        painter.drawRect(glow_rect)

        # Border
        painter.setBrush(Qt.NoBrush)
        border_pen = QPen(QColor(93, 154, 150, 120))  # Muted teal
        border_pen.setWidth(2)
        painter.setPen(border_pen)
        painter.drawRect(self.rect().adjusted(1, 1, -1, -1))

        painter.restore()

    def _draw_orbital_rings(self, painter):
        """Draw rotating orbital rings around center."""
        painter.save()

        cx, cy = self.width() / 2, 100
        opacity = self.intro_progress * 0.7

        # Outer ring
        painter.translate(cx, cy)
        painter.rotate(self.ring_rotation)

        # Draw ring as arc segments
        pen = QPen(QColor(93, 154, 150, int(200 * opacity)))  # Teal
        pen.setWidth(2)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)

        # Draw partial arcs
        painter.drawArc(QRectF(-50, -50, 100, 100), 0, 120 * 16)

        pen.setColor(QColor(147, 112, 162, int(200 * opacity)))  # Purple
        painter.setPen(pen)
        painter.drawArc(QRectF(-50, -50, 100, 100), 180 * 16, 120 * 16)

        # Inner ring (counter-rotate)
        painter.rotate(-self.ring_rotation * 2)

        pen.setColor(QColor(127, 179, 175, int(150 * opacity)))  # Light teal
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawArc(QRectF(-36, -36, 72, 72), 60 * 16, 120 * 16)

        pen.setColor(QColor(93, 154, 150, int(150 * opacity)))  # Teal
        painter.setPen(pen)
        painter.drawArc(QRectF(-36, -36, 72, 72), 240 * 16, 120 * 16)

        painter.restore()

    def _draw_center_emblem(self, painter):
        """Draw the central emblem with envelope icon."""
        painter.save()

        cx, cy = self.width() / 2, 100
        scale = self.intro_progress

        painter.translate(cx, cy)
        painter.scale(scale, scale)

        # Outer glow circle
        glow = QRadialGradient(0, 0, 45)
        glow.setColorAt(0, QColor(93, 154, 150, 60))  # Teal
        glow.setColorAt(0.6, QColor(147, 112, 162, 30))  # Purple
        glow.setColorAt(1, QColor(0, 0, 0, 0))
        painter.setBrush(glow)
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(QPointF(0, 0), 45, 45)

        # Main emblem circle
        emblem_bg = QRadialGradient(0, -8, 32)
        emblem_bg.setColorAt(0, QColor(51, 65, 85))
        emblem_bg.setColorAt(1, QColor(30, 41, 59))

        painter.setBrush(emblem_bg)
        pen = QPen(QColor(71, 85, 105))
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawEllipse(QPointF(0, 0), 28, 28)

        # Pulsing inner ring
        pulse = 0.85 + 0.15 * math.sin(self.pulse_phase)
        pen = QPen(QColor(93, 154, 150, 180))  # Teal
        pen.setWidth(2)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(QPointF(0, 0), 20 * pulse, 20 * pulse)

        # Draw envelope icon
        painter.setPen(Qt.NoPen)
        envelope_color = QColor(226, 232, 240)
        painter.setBrush(envelope_color)

        # Envelope body
        env_w, env_h = 16, 11
        painter.drawRect(int(-env_w/2), int(-env_h/2 + 1), env_w, env_h)

        # Envelope flap (triangle)
        flap_path = QPainterPath()
        flap_path.moveTo(-env_w/2, -env_h/2 + 1)
        flap_path.lineTo(0, 3)
        flap_path.lineTo(env_w/2, -env_h/2 + 1)
        flap_path.closeSubpath()

        painter.setBrush(QColor(200, 210, 220))
        painter.drawPath(flap_path)

        painter.restore()

    def _draw_title(self, painter):
        """Draw application title."""
        painter.save()

        opacity = max(0, (self.intro_progress - 0.2) / 0.8) if self.intro_progress > 0.2 else 0

        # Title font
        font = QFont("Segoe UI", 36, QFont.Light)
        font.setLetterSpacing(QFont.AbsoluteSpacing, 2)
        painter.setFont(font)

        title_rect = QRectF(0, 155, self.width(), 50)

        # Shadow
        painter.setPen(QColor(0, 0, 0, int(100 * opacity)))
        painter.drawText(title_rect.adjusted(2, 2, 2, 2), Qt.AlignCenter, "DocuShuttle")

        # Draw "Docu" in teal
        metrics = painter.fontMetrics()
        full_width = metrics.horizontalAdvance("DocuShuttle")
        start_x = (self.width() - full_width) / 2

        painter.setPen(QColor(93, 154, 150, int(255 * opacity)))  # Teal
        painter.drawText(int(start_x), 195, "Docu")

        # Draw "Shuttle" in purple
        docu_width = metrics.horizontalAdvance("Docu")
        painter.setPen(QColor(147, 112, 162, int(255 * opacity)))  # Purple
        painter.drawText(int(start_x + docu_width), 195, "Shuttle")

        painter.restore()

    def _draw_tagline(self, painter):
        """Draw tagline."""
        painter.save()

        opacity = max(0, (self.intro_progress - 0.4) / 0.6) if self.intro_progress > 0.4 else 0

        font = QFont("Segoe UI", 10)
        font.setLetterSpacing(QFont.AbsoluteSpacing, 2)
        painter.setFont(font)
        painter.setPen(QColor(148, 163, 184, int(255 * opacity)))

        painter.drawText(QRectF(0, 205, self.width(), 25), Qt.AlignCenter,
                        "EMAIL FORWARDING AUTOMATION")

        painter.restore()

    def _draw_progress_area(self, painter):
        """Draw progress bar and status."""
        painter.save()

        opacity = max(0, (self.intro_progress - 0.5) / 0.5) if self.intro_progress > 0.5 else 0

        # Status message
        font = QFont("Segoe UI", 10)
        painter.setFont(font)
        painter.setPen(QColor(148, 163, 184, int(255 * opacity)))
        painter.drawText(QRectF(0, 248, self.width(), 20), Qt.AlignCenter, self._message)

        # Progress bar dimensions
        bar_width = 320
        bar_height = 5
        bar_x = (self.width() - bar_width) / 2
        bar_y = 278

        # Track background
        painter.setBrush(QColor(51, 65, 85, int(255 * opacity)))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(QRectF(bar_x, bar_y, bar_width, bar_height), 2, 2)

        # Progress fill
        if self.progress > 0.5:
            fill_width = (self.progress / 100) * bar_width

            # Animated gradient (teal to purple)
            offset = self.wave_offset % (bar_width * 2)
            fill_grad = QLinearGradient(bar_x - offset, 0, bar_x + bar_width * 2 - offset, 0)
            fill_grad.setColorAt(0, QColor(93, 154, 150))   # Teal
            fill_grad.setColorAt(0.33, QColor(127, 179, 175))  # Light teal
            fill_grad.setColorAt(0.66, QColor(147, 112, 162))  # Purple
            fill_grad.setColorAt(1, QColor(93, 154, 150))   # Teal

            # Clip and draw
            painter.setClipRect(QRectF(bar_x, bar_y, fill_width, bar_height))
            painter.setBrush(fill_grad)
            painter.setOpacity(opacity)
            painter.drawRoundedRect(QRectF(bar_x, bar_y, bar_width, bar_height), 2, 2)
            painter.setClipping(False)

            # Top shine
            shine = QLinearGradient(0, bar_y, 0, bar_y + bar_height)
            shine.setColorAt(0, QColor(255, 255, 255, 70))
            shine.setColorAt(0.5, QColor(255, 255, 255, 0))
            painter.setClipRect(QRectF(bar_x, bar_y, fill_width, bar_height / 2))
            painter.setBrush(shine)
            painter.drawRoundedRect(QRectF(bar_x, bar_y, bar_width, bar_height), 2, 2)
            painter.setClipping(False)

        # Percentage text
        painter.setOpacity(opacity)
        pct_font = QFont("Segoe UI", 9)
        painter.setFont(pct_font)
        painter.setPen(QColor(100, 116, 139))
        painter.drawText(QRectF(bar_x + bar_width + 12, bar_y - 3, 50, 14),
                        Qt.AlignLeft | Qt.AlignVCenter, f"{int(self.progress)}%")

        painter.restore()

    def _draw_corner_accents(self, painter):
        """Draw corner accent decorations."""
        painter.save()

        opacity = self.intro_progress * 0.25

        # Top left - teal
        pen = QPen(QColor(93, 154, 150, int(255 * opacity)))
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawLine(15, 12, 40, 12)
        painter.drawLine(12, 15, 12, 40)

        # Top right - teal
        painter.drawLine(self.width() - 40, 12, self.width() - 15, 12)
        painter.drawLine(self.width() - 12, 15, self.width() - 12, 40)

        # Bottom left - purple
        pen.setColor(QColor(147, 112, 162, int(255 * opacity)))
        painter.setPen(pen)
        painter.drawLine(15, self.height() - 12, 40, self.height() - 12)
        painter.drawLine(12, self.height() - 40, 12, self.height() - 15)

        # Bottom right - purple
        painter.drawLine(self.width() - 40, self.height() - 12, self.width() - 15, self.height() - 12)
        painter.drawLine(self.width() - 12, self.height() - 40, self.width() - 12, self.height() - 15)

        painter.restore()

        # Version
        painter.save()
        opacity = max(0, (self.intro_progress - 0.6) / 0.4) if self.intro_progress > 0.6 else 0

        font = QFont("Segoe UI", 8)
        painter.setFont(font)
        painter.setPen(QColor(100, 116, 139, int(180 * opacity)))
        painter.drawText(QRectF(0, 308, self.width(), 20), Qt.AlignCenter, f"v{APP_VERSION}")
        painter.restore()

    def finish_splash(self, window):
        """Finish the splash and show main window."""
        self.is_fading = True
        self.timer.stop()
        self.progress_timer.stop()
        window.show()
        self.close()


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    """Application entry point."""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # Show animated splash screen
    splash = AnimatedSplashScreen()
    splash.show()
    app.processEvents()

    # Create the main window while splash is showing
    window = DocuShuttleWindow()

    # Wait for splash animation to complete (progress reaches 100%)
    def check_splash_done():
        if splash.progress >= 100:
            splash.finish_splash(window)
        else:
            QTimer.singleShot(100, check_splash_done)

    QTimer.singleShot(100, check_splash_done)

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
