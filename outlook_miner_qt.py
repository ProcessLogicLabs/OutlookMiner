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
from queue import Queue, Empty

# PyQt5 imports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTextEdit, QDateEdit,
    QCheckBox, QGroupBox, QTabWidget, QFrame, QMessageBox, QDialog,
    QFormLayout, QSpacerItem, QSizePolicy, QMenu, QAction, QToolButton
)
from PyQt5.QtCore import Qt, QDate, QTimer, pyqtSignal, QObject, QThread, QPropertyAnimation, QPointF, QRectF, QEasingCurve
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor, QPixmap, QPainter, QPen, QBrush, QPainterPath, QRadialGradient
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
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

ICON_PATH = os.path.join(BASE_PATH, 'myicon.ico')
ICON_PNG_PATH = os.path.join(BASE_PATH, 'myicon.png')

# Constants
LOG_BUFFER_SIZE = 10
MAX_LOG_LINES = 1000
DEFAULT_TIMEZONE = 'US/Eastern'

# Thread lock for database access
db_lock = threading.Lock()

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
# WORKER SIGNALS
# ============================================================================
class WorkerSignals(QObject):
    """Signals for worker threads to communicate with GUI."""
    log_message = pyqtSignal(str)
    display_subject = pyqtSignal(str)
    operation_complete = pyqtSignal(int, int)
    search_complete = pyqtSignal(int, list)
    error = pyqtSignal(str)
    clear_subjects = pyqtSignal()


# ============================================================================
# DATABASE FUNCTIONS
# ============================================================================
def init_db():
    """Initialize SQLite database and create required tables."""
    new_db = not os.path.exists('minerdb.db')
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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
        return new_db
    except Exception as e:
        raise Exception(f"Error initializing database: {str(e)}")


def load_email_addresses():
    """Load all distinct recipient email addresses from the database."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute("SELECT DISTINCT recipient FROM Clients WHERE recipient IS NOT NULL")
                return [row[0] for row in c.fetchall()]
    except Exception:
        return []


def save_setting(key, value):
    """Save a setting to the Settings table."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute("INSERT OR REPLACE INTO Settings (key, value) VALUES (?, ?)", (key, value))
                conn.commit()
    except Exception:
        pass


def load_setting(key):
    """Load a setting from the Settings table."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
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

                        if skip_forwarded and file_number and check_if_forwarded_db(file_number, recipient):
                            continue

                        sent_on = item.SentOn
                        if sent_on < start_date or sent_on > end_date:
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

                        if skip_forwarded and file_number and check_if_forwarded_db(file_number, recipient):
                            continue

                        sent_on = item.SentOn
                        if sent_on < start_date or sent_on > end_date:
                            continue

                        if require_attachments and item.Attachments.Count == 0:
                            continue

                        new_subject = file_number if file_number else subject

                        forward_email = item.Forward()
                        forward_email.To = recipient
                        forward_email.Subject = new_subject
                        forward_email.Send()

                        emails_processed += 1
                        self._log(f"Forwarded: {new_subject}")
                        self.signals.display_subject.emit(new_subject)

                        if file_number:
                            log_forwarded_email(file_number, recipient)

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
    """Configuration dialog for advanced settings."""

    def __init__(self, parent=None, prefix="", delay="0", require_attach=True, skip_fwd=True):
        super().__init__(parent)
        self.setWindowTitle("Configuration")
        self.setFixedSize(400, 280)
        self.setStyleSheet(STYLESHEET)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

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

        layout.addLayout(form)
        layout.addStretch()

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

    def get_values(self):
        """Return dialog values."""
        return {
            'prefix': self.prefix_edit.text(),
            'delay': self.delay_edit.text(),
            'require_attachments': self.require_attach_check.isChecked(),
            'skip_forwarded': self.skip_fwd_check.isChecked()
        }


# ============================================================================
# MAIN WINDOW
# ============================================================================
class DocuShuttleWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.config_prefix = ""
        self.config_delay = "0"
        self.config_require_attachments = True
        self.config_skip_forwarded = True

        self.init_ui()
        self.load_saved_state()

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

        self.files_text = QTextEdit()
        self.files_text.setReadOnly(False)
        self.files_text.setMinimumHeight(120)
        files_layout.addWidget(self.files_text)

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
        dialog = ConfigDialog(
            self,
            self.config_prefix,
            self.config_delay,
            self.config_require_attachments,
            self.config_skip_forwarded
        )

        if dialog.exec_() == QDialog.Accepted:
            values = dialog.get_values()
            self.config_prefix = values['prefix']
            self.config_delay = values['delay']
            self.config_require_attachments = values['require_attachments']
            self.config_skip_forwarded = values['skip_forwarded']
            self.log("Configuration updated")

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
        self.worker.signals.clear_subjects.connect(self.files_text.clear)
        self.worker.signals.operation_complete.connect(self.on_forward_complete)
        self.worker.signals.error.connect(self.on_error)
        self.worker.finished.connect(lambda: self.set_buttons_enabled(True))
        self.worker.start()

    def display_subject(self, subject):
        """Display forwarded subject."""
        self.files_text.append(subject)

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


# ============================================================================
# ANIMATED SPLASH SCREEN
# ============================================================================
class Envelope:
    """Represents an envelope that spirals into the vortex."""
    def __init__(self, x, y, size, angle, distance):
        self.x = x
        self.y = y
        self.size = size
        self.angle = angle  # Angle around the vortex center
        self.distance = distance  # Distance from center
        self.rotation = random.uniform(0, 360)  # Envelope rotation
        self.speed = random.uniform(0.8, 1.2)  # Speed multiplier
        self.opacity = 1.0


class AnimatedSplashScreen(QSplashScreen):
    """Animated splash screen with envelopes spiraling into a vortex."""

    def __init__(self):
        # Create a pixmap for the splash
        self.splash_width = 500
        self.splash_height = 350
        pixmap = QPixmap(self.splash_width, self.splash_height)
        pixmap.fill(Qt.transparent)
        super().__init__(pixmap)

        # Set window flags for transparency
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Vortex center
        self.center_x = self.splash_width // 2
        self.center_y = self.splash_height // 2 - 20

        # Create envelopes at various positions around the vortex
        self.envelopes = []
        self.create_envelopes()

        # Animation properties
        self.vortex_rotation = 0
        self.progress = 0

        # Timer for animation
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.animate)
        self.timer.start(30)  # ~33 FPS

        # Progress timer
        self.progress_timer = QTimer(self)
        self.progress_timer.timeout.connect(self.update_progress)
        self.progress_timer.start(50)

    def create_envelopes(self):
        """Create envelopes at random positions around the vortex."""
        self.envelopes = []
        for _ in range(12):
            angle = random.uniform(0, 360)
            distance = random.uniform(50, 120)  # Closer to center, away from edges
            size = random.uniform(18, 30)
            x = self.center_x + distance * math.cos(math.radians(angle))
            y = self.center_y + distance * math.sin(math.radians(angle))
            self.envelopes.append(Envelope(x, y, size, angle, distance))

    def animate(self):
        """Update animation state."""
        self.vortex_rotation += 3

        # Update each envelope - spiral toward center
        for env in self.envelopes:
            # Decrease distance (spiral inward)
            env.distance -= 1.5 * env.speed

            # Increase angle (rotate around center)
            env.angle += 4 * env.speed

            # Rotate the envelope itself
            env.rotation += 5 * env.speed

            # Update position based on polar coordinates
            env.x = self.center_x + env.distance * math.cos(math.radians(env.angle))
            env.y = self.center_y + env.distance * math.sin(math.radians(env.angle))

            # Shrink as it approaches center
            if env.distance < 60:
                env.size *= 0.96
                env.opacity = max(0, env.distance / 60)

            # Reset envelope if it reaches center or becomes too small
            if env.distance < 10 or env.size < 5:
                env.angle = random.uniform(0, 360)
                env.distance = random.uniform(100, 130)  # Respawn closer to center
                env.size = random.uniform(18, 30)
                env.rotation = random.uniform(0, 360)
                env.speed = random.uniform(0.8, 1.2)
                env.opacity = 1.0
                env.x = self.center_x + env.distance * math.cos(math.radians(env.angle))
                env.y = self.center_y + env.distance * math.sin(math.radians(env.angle))

        self.update()

    def update_progress(self):
        """Update progress bar."""
        self.progress += 2
        if self.progress >= 100:
            self.progress_timer.stop()
        self.update()

    def draw_envelope(self, painter, x, y, size, rotation, opacity):
        """Draw a simple envelope shape."""
        painter.save()
        painter.translate(x, y)
        painter.rotate(rotation)

        # Set opacity
        painter.setOpacity(opacity)

        # Envelope body (rectangle)
        envelope_color = QColor(COLORS['primary'])
        envelope_color.setAlpha(int(220 * opacity))
        painter.setBrush(QBrush(envelope_color))
        painter.setPen(QPen(QColor(255, 255, 255, int(180 * opacity)), 1))

        half_w = size / 2
        half_h = size / 3
        painter.drawRect(int(-half_w), int(-half_h), int(size), int(size * 0.66))

        # Envelope flap (triangle)
        flap_path = QPainterPath()
        flap_path.moveTo(-half_w, -half_h)
        flap_path.lineTo(0, half_h * 0.3)
        flap_path.lineTo(half_w, -half_h)
        flap_path.closeSubpath()

        flap_color = QColor(COLORS['primary_hover'])
        flap_color.setAlpha(int(200 * opacity))
        painter.setBrush(QBrush(flap_color))
        painter.drawPath(flap_path)

        painter.restore()

    def draw_vortex(self, painter):
        """Draw the central vortex effect."""
        # Draw spiral lines
        painter.setRenderHint(QPainter.Antialiasing)

        for i in range(4):
            offset_angle = self.vortex_rotation + (i * 90)

            # Create gradient spiral
            path = QPainterPath()
            start_dist = 15
            end_dist = 70

            for j in range(50):
                t = j / 49
                angle = offset_angle + t * 360
                dist = start_dist + t * (end_dist - start_dist)
                x = self.center_x + dist * math.cos(math.radians(angle))
                y = self.center_y + dist * math.sin(math.radians(angle))

                if j == 0:
                    path.moveTo(x, y)
                else:
                    path.lineTo(x, y)

            # Draw with gradient alpha
            pen_color = QColor(COLORS['primary'])
            pen_color.setAlpha(150)
            painter.setPen(QPen(pen_color, 3))
            painter.drawPath(path)

        # Draw center circle with gradient
        gradient = QRadialGradient(self.center_x, self.center_y, 25)
        gradient.setColorAt(0, QColor(COLORS['primary']))
        gradient.setColorAt(0.7, QColor(COLORS['primary_light']))
        gradient.setColorAt(1, QColor(255, 255, 255, 0))

        painter.setBrush(QBrush(gradient))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(self.center_x - 25, self.center_y - 25, 50, 50)

    def paintEvent(self, event):
        """Custom paint for the splash screen."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Transparent background - no fill

        # Draw brand text at top with shadow for visibility
        text_y = 50
        font = QFont('Segoe UI', 24, QFont.Bold)
        painter.setFont(font)

        # Draw text outline/glow for visibility on any background
        outline_color = QColor(255, 255, 255, 200)
        for dx in [-2, -1, 0, 1, 2]:
            for dy in [-2, -1, 0, 1, 2]:
                if dx != 0 or dy != 0:
                    painter.setPen(outline_color)
                    painter.drawText(170 + dx, text_y + dy, "Docu")
                    painter.drawText(245 + dx, text_y + dy, "Shuttle")

        # Draw "Docu" in teal
        painter.setPen(QColor(COLORS['primary']))
        painter.drawText(170, text_y, "Docu")

        # Draw "Shuttle" in muted purple
        painter.setPen(QColor(147, 112, 162))  # Muted purple
        painter.drawText(245, text_y, "Shuttle")

        # Draw subtitle with outline for visibility
        font = QFont('Segoe UI', 10)
        painter.setFont(font)
        subtitle_outline = QColor(255, 255, 255, 180)
        for dx in [-1, 0, 1]:
            for dy in [-1, 0, 1]:
                if dx != 0 or dy != 0:
                    painter.setPen(subtitle_outline)
                    painter.drawText(155 + dx, text_y + 25 + dy, "Email Forwarding Automation")
        painter.setPen(QColor(80, 80, 80))
        painter.drawText(155, text_y + 25, "Email Forwarding Automation")

        # Draw the vortex
        self.draw_vortex(painter)

        # Draw envelopes
        for env in self.envelopes:
            self.draw_envelope(painter, env.x, env.y, env.size, env.rotation, env.opacity)

        # Draw progress bar at bottom
        progress_y = self.splash_height - 50
        progress_width = self.splash_width - 100
        progress_x = 50
        progress_height = 6

        # Background track
        painter.setBrush(QBrush(QColor(COLORS['border'])))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(progress_x, progress_y, progress_width, progress_height, 3, 3)

        # Progress fill
        fill_width = int((self.progress / 100) * progress_width)
        if fill_width > 0:
            painter.setBrush(QBrush(QColor(COLORS['primary'])))
            painter.drawRoundedRect(progress_x, progress_y, fill_width, progress_height, 3, 3)

        # Loading text with shadow
        font = QFont('Segoe UI', 9)
        painter.setFont(font)
        painter.setPen(QColor(0, 0, 0, 80))
        loading_text = "Loading..." if self.progress < 100 else "Ready!"
        painter.drawText(progress_x + 1, progress_y + 23, loading_text)
        painter.drawText(progress_x + progress_width - 49, progress_y + 23, "v1.3.0")

        painter.setPen(QColor(COLORS['text_secondary']))
        painter.drawText(progress_x, progress_y + 22, loading_text)
        painter.drawText(progress_x + progress_width - 50, progress_y + 22, "v1.3.0")

        painter.end()

    def finish_splash(self, window):
        """Finish the splash and show main window."""
        self.timer.stop()
        self.progress_timer.stop()
        self.finish(window)


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
            window.show()
        else:
            QTimer.singleShot(100, check_splash_done)

    QTimer.singleShot(100, check_splash_done)

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
