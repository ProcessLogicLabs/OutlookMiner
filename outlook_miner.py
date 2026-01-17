"""
Outlook Miner - Email Forwarding Automation Tool

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

import tkinter as tk
from tkinter import messagebox, ttk, Toplevel
import win32com.client
import pythoncom
import os
import datetime
import pytz
from tkinter.scrolledtext import ScrolledText
import re
from tkcalendar import DateEntry
import sqlite3
import threading
from queue import Queue, Empty
import uuid
import time
from functools import wraps

# ============================================================================
# GLOBAL CONFIGURATION
# ============================================================================

# Determine the base path for resources (handles PyInstaller bundled exe)
import sys
if getattr(sys, 'frozen', False):
    # Running as compiled exe - use _MEIPASS for bundled data files
    BASE_PATH = sys._MEIPASS
else:
    # Running as script
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

ICON_PATH = os.path.join(BASE_PATH, 'myicon.ico')
ICON_PNG_PATH = os.path.join(BASE_PATH, 'myicon.png')

# Global flag to control operation cancellation
cancel_scan = False

# Queue for thread-safe GUI updates from worker threads
gui_queue = Queue()

# Thread lock for synchronized database access
db_lock = threading.Lock()

# Buffer for batching log messages to improve performance
log_buffer = []

# Flag to track if log frame is initialized (lazy loading)
log_frame_initialized = False

# Constants for application behavior
LOG_BUFFER_SIZE = 10  # Number of messages to batch before writing
MAX_LOG_LINES = 1000  # Maximum lines to keep in log display
PROGRESS_UPDATE_INTERVAL = 100  # Update progress every N emails
DEFAULT_TIMEZONE = 'US/Eastern'  # Default timezone for timestamps

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def debounce(wait):
    """
    Decorator to debounce function calls, preventing rapid successive executions.

    Args:
        wait (float): Minimum time in seconds between function executions

    Returns:
        function: Decorated function that implements debouncing

    Note:
        This is useful for preventing UI events from triggering too frequently
    """
    def decorator(func):
        last_call = 0
        last_args = None
        last_kwargs = None
        lock = threading.Lock()

        @wraps(func)
        def debounced(*args, **kwargs):
            nonlocal last_call, last_args, last_kwargs
            current_time = time.time()

            with lock:
                if current_time - last_call >= wait:
                    last_call = current_time
                    last_args = args
                    last_kwargs = kwargs
                    root.after(0, lambda: func(*last_args, **last_kwargs))
                else:
                    root.after(int(wait * 1000), lambda: func(*last_args, **last_kwargs) if time.time() - last_call >= wait else None)
                    last_args = args
                    last_kwargs = kwargs
        return debounced
    return decorator

def process_gui_queue():
    """
    Process queued GUI updates from worker threads in the main thread.

    This function runs periodically (every 200ms) to process any GUI updates
    that have been queued by worker threads. This ensures thread-safe GUI operations.

    Performance is monitored and logged if processing takes longer than 0.1 seconds.
    """
    start_time = time.time()
    try:
        # Process all queued updates
        while True:
            func, args = gui_queue.get_nowait()
            func(*args)
    except Empty:
        pass  # Queue is empty, which is normal

    # Monitor performance and warn if slow
    elapsed = time.time() - start_time
    if elapsed > 0.1:
        gui_safe_log_message(f"GUI queue processing took {elapsed:.2f} seconds")

    # Schedule next queue processing
    root.after(200, process_gui_queue)

def gui_safe_log_message(message):
    """
    Thread-safe logging function that queues messages for GUI display.

    Args:
        message (str): The log message to display

    Note:
        - Messages are batched for performance (batch size: LOG_BUFFER_SIZE)
        - If GUI is not initialized, messages are written to startup log file
        - Uses US/Eastern timezone for consistent timestamps
    """
    if log_frame_initialized:
        log_buffer.append(message)
        # Batch messages for better performance
        if len(log_buffer) >= LOG_BUFFER_SIZE:
            gui_queue.put((log_message_main_thread, (log_buffer[:],)))
            log_buffer.clear()
    else:
        # Log to file during startup before GUI is ready
        with open("outlook_miner_startup.log", 'a', encoding='utf-8') as f:
            timestamp = datetime.datetime.now(pytz.timezone(DEFAULT_TIMEZONE)).strftime("%Y-%m-%d %H:%M:%S %Z")
            f.write(f"[{timestamp}] {message}\n")

def log_message_main_thread(messages):
    """
    Display log messages in the GUI log window with timestamps.

    Args:
        messages (list or str): Single message or list of messages to display

    Note:
        - Automatically scrolls to show latest messages
        - Maintains a maximum of MAX_LOG_LINES to prevent memory issues
        - Must be called from main thread (GUI thread)
    """
    local_tz = pytz.timezone(DEFAULT_TIMEZONE)
    timestamp = datetime.datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S %Z")

    # Handle both single messages and batched messages
    for message in messages if isinstance(messages, list) else [messages]:
        log_text.insert(tk.END, f"[{timestamp}] {message}\n")

    # Auto-scroll to show latest message
    log_text.see(tk.END)

    # Trim old log lines to prevent excessive memory usage
    line_count = int(log_text.index('end-1c').split('.')[0])
    if line_count > MAX_LOG_LINES:
        log_text.delete('1.0', f'{line_count-MAX_LOG_LINES}.0')

def gui_safe_display_subject(subject):
    """
    Thread-safe function to queue subject display updates.

    Args:
        subject (str): Email subject to display in the "Files Sent" window
    """
    gui_queue.put((display_subject_main_thread, (subject,)))

def display_subject_main_thread(subject):
    """
    Display forwarded email subject in the GUI.

    Args:
        subject (str): Email subject to append to the display

    Note:
        Must be called from main thread
    """
    subject_text.insert(tk.END, f"{subject}\n")
    subject_text.see(tk.END)

# ============================================================================
# DATABASE FUNCTIONS
# ============================================================================

def init_db():
    """
    Initialize SQLite database and create required tables if they don't exist.

    Creates two tables:
    - Clients: Stores configuration for each recipient email
    - ForwardedEmails: Tracks which file numbers have been forwarded to prevent duplicates

    Raises:
        Exception: If database creation or initialization fails
    """
    new_db = not os.path.exists('minerdb.db')
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                # Check if Clients table exists
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
                # Check if ForwardedEmails table exists
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ForwardedEmails'")
                if not c.fetchone():
                    c.execute('''CREATE TABLE ForwardedEmails
                                 (file_number TEXT,
                                  recipient TEXT,
                                  forwarded_at TIMESTAMP,
                                  PRIMARY KEY (file_number, recipient))''')
                # Check if Settings table exists (for app preferences like last used email)
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Settings'")
                if not c.fetchone():
                    c.execute('''CREATE TABLE Settings
                                 (key TEXT PRIMARY KEY,
                                  value TEXT)''')
                conn.commit()
        if new_db:
            gui_safe_log_message("SQLite database 'minerdb.db' created and initialized with Clients and ForwardedEmails tables.")
        else:
            gui_safe_log_message("SQLite database 'minerdb.db' verified, Clients and ForwardedEmails tables ensured.")
    except Exception as e:
        error_message = f"Error initializing database 'minerdb.db' for Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
        raise

def load_email_addresses():
    """
    Load all distinct recipient email addresses from the database.

    Returns:
        list: List of email addresses, or empty list if error occurs
    """
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute("SELECT DISTINCT recipient FROM Clients WHERE recipient IS NOT NULL")
                emails = [row[0] for row in c.fetchall()]
        return emails
    except Exception as e:
        error_message = f"Error loading email addresses from database in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
        return []

def save_setting(key, value):
    """Save a setting to the Settings table."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute("INSERT OR REPLACE INTO Settings (key, value) VALUES (?, ?)",
                         (key, value))
                conn.commit()
    except Exception as e:
        gui_safe_log_message(f"Error saving setting '{key}': {str(e)}")

def load_setting(key):
    """Load a setting from the Settings table."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute("SELECT value FROM Settings WHERE key = ?", (key,))
                result = c.fetchone()
                return result[0] if result else None
    except Exception as e:
        gui_safe_log_message(f"Error loading setting '{key}': {str(e)}")
        return None

def save_last_used_email(email):
    """Save the last used email address to Settings table."""
    save_setting('last_used_email', email)

def load_last_used_email():
    """Load the last used email address from Settings table."""
    return load_setting('last_used_email')

def load_config_for_email(event):
    """Load configuration settings for the selected email from the Clients table."""
    recipient = recipient_combobox.get().strip()
    if not recipient:
        gui_safe_log_message("No email selected in Outlook Miner. GUI fields unchanged.")
        return
    # Save this as the last used email
    save_last_used_email(recipient)
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute('''SELECT start_date, end_date, file_number_prefix, subject_keyword,
                             require_attachments, skip_forwarded, delay_seconds
                             FROM Clients WHERE recipient = ?''', (recipient,))
                row = c.fetchone()
                if row:
                    config = {
                        'start_date': convert_date_format(row[0]) if row[0] and row[0].strip() else None,
                        'end_date': convert_date_format(row[1]) if row[1] and row[1].strip() else None,
                        'file_number_prefix': row[2] if row[2] is not None else "",
                        'subject_keyword': row[3] if row[3] is not None else "BILLING INVOICE",
                        'require_attachments': row[4] == "1" if row[4] is not None else True,
                        'skip_forwarded': row[5] == "1" if row[5] is not None else True,
                        'delay_seconds': row[6] if row[6] is not None else "0"
                    }
                    # Handle start_date
                    if config.get("start_date"):
                        try:
                            start_date_entry.set_date(config.get("start_date"))
                        except ValueError:
                            gui_safe_log_message(f"Warning: Invalid start_date '{config.get('start_date')}' for {recipient}. Clearing date in Outlook Miner.")
                            start_date_entry._set_date(None)
                    else:
                        start_date_entry._set_date(None)
                    # Handle end_date
                    if config.get("end_date"):
                        try:
                            end_date_entry.set_date(config.get("end_date"))
                        except ValueError:
                            gui_safe_log_message(f"Warning: Invalid end_date '{config.get('end_date')}' for {recipient}. Clearing date in Outlook Miner.")
                            end_date_entry._set_date(None)
                    else:
                        end_date_entry._set_date(None)
                    file_number_prefix_entry.delete(0, tk.END)
                    file_number_prefix_entry.insert(0, config.get("file_number_prefix", ""))
                    subject_keyword_entry.delete(0, tk.END)
                    subject_keyword_entry.insert(0, config.get("subject_keyword", "BILLING INVOICE"))
                    require_attachments_var.set(config.get("require_attachments", True))
                    skip_forwarded_var.set(config.get("skip_forwarded", True))
                    delay_seconds_entry.delete(0, tk.END)
                    delay_seconds_entry.insert(0, config.get("delay_seconds", "0"))
                    gui_safe_log_message(f"Loaded configuration for email '{recipient}' in Outlook Miner: Require Attachments={config.get('require_attachments')}, Skip Forwarded={config.get('skip_forwarded')}.")
                else:
                    gui_safe_log_message(f"No configuration found for email '{recipient}' in Outlook Miner. Using default values.")
                    start_date_entry._set_date(None)
                    end_date_entry._set_date(None)
                    file_number_prefix_entry.delete(0, tk.END)
                    subject_keyword_entry.delete(0, tk.END)
                    subject_keyword_entry.insert(0, "BILLING INVOICE")
                    require_attachments_var.set(True)
                    skip_forwarded_var.set(True)
                    delay_seconds_entry.delete(0, tk.END)
                    delay_seconds_entry.insert(0, "0")
    except Exception as e:
        error_message = f"Failed to load config for email '{recipient}' in Outlook Miner: {str(e)}. Using default values."
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
        start_date_entry._set_date(None)
        end_date_entry._set_date(None)
        file_number_prefix_entry.delete(0, tk.END)
        subject_keyword_entry.delete(0, tk.END)
        subject_keyword_entry.insert(0, "BILLING INVOICE")
        require_attachments_var.set(True)
        skip_forwarded_var.set(True)
        delay_seconds_entry.delete(0, tk.END)
        delay_seconds_entry.insert(0, "0")

# ============================================================================
# VALIDATION AND UTILITY FUNCTIONS
# ============================================================================

def validate_email(email):
    """
    Validate email address format using regex.

    Args:
        email (str): Email address to validate

    Returns:
        bool: True if email format is valid, False otherwise
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

def sanitize_filter_value(value):
    """
    Sanitize input for MAPI Restrict filter to escape special characters.

    Args:
        value (str): Value to sanitize

    Returns:
        str: Sanitized value with escaped special characters

    Note:
        Escapes single quotes and percent signs to prevent filter injection
    """
    if not value:
        return ""
    value = value.replace("'", "''").replace("%", "%%")
    return value

def convert_date_format(date_str):
    """
    Convert date between YYYY-MM-DD and MM/DD/YYYY formats.

    Args:
        date_str (str): Date string in either format

    Returns:
        str or None: Converted date string, or None if invalid format
    """
    if not date_str or not date_str.strip():
        return None
    try:
        parsed_date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return parsed_date.strftime("%m/%d/%Y")
    except ValueError:
        try:
            parsed_date = datetime.datetime.strptime(date_str, "%m/%d/%Y")
            return date_str
        except ValueError:
            gui_safe_log_message(f"Warning: Invalid date format '{date_str}'. Clearing date in Outlook Miner.")
            return None

# ============================================================================
# OUTLOOK INTEGRATION FUNCTIONS
# ============================================================================

def get_outlook_folder(mapi):
    """
    Navigate to and return the Outlook Sent Items folder.

    Args:
        mapi: Outlook MAPI namespace object

    Returns:
        Folder: Outlook Sent Items folder object

    Raises:
        Exception: If unable to access Sent Items folder
    """
    try:
        folder = mapi.GetDefaultFolder(5)
        if folder.Items.Count == 0:
            gui_safe_log_message("Warning: Sent Items folder is empty.")
        else:
            gui_safe_log_message(f"Sent Items folder contains {folder.Items.Count} emails.")
        return folder
    except Exception as e:
        error_message = f"Error accessing Sent Items folder in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
        raise

def log_forwarded_subject(subject):
    """Append the forwarded email subject to a separate log file."""
    try:
        with open("forwarded_emails.log", 'a', encoding='utf-8') as f:
            local_tz = pytz.timezone('US/Eastern')
            timestamp = datetime.datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S %Z")
            f.write(f"[{timestamp}] Forwarded: {subject}\n")
    except Exception as e:
        error_message = f"Error writing to forwarded_emails.log in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))

def log_forwarded_email(file_number, recipient):
    """Log forwarded email file number and recipient to the ForwardedEmails table."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                forwarded_at = datetime.datetime.now(pytz.timezone('US/Eastern')).strftime("%Y-%m-%d %H:%M:%S")
                c.execute('''INSERT OR REPLACE INTO ForwardedEmails (file_number, recipient, forwarded_at)
                             VALUES (?, ?, ?)''', (file_number, recipient.lower(), forwarded_at))
                conn.commit()
        gui_safe_log_message(f"Logged forwarded email with file number '{file_number}' to {recipient} in database.")
    except Exception as e:
        error_message = f"Error logging forwarded email to database in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))

def check_if_forwarded_db(file_number, recipient):
    """Check if the file number was previously forwarded to the recipient using the database."""
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute('''SELECT COUNT(*) FROM ForwardedEmails WHERE file_number = ? AND recipient = ?''',
                          (file_number, recipient.lower()))
                count = c.fetchone()[0]
                return count > 0
    except Exception as e:
        error_message = f"Error checking forwarded email in database for Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        return False

def extract_file_number(item, file_number_prefixes):
    """
    Extract file number from email attachment filename or subject line.

    Args:
        item: Outlook MailItem object
        file_number_prefixes (list): List of numeric prefixes to match

    Returns:
        str or None: Extracted file number if found, None otherwise

    Note:
        Searches first in attachment filenames, then in email subject
        File number format: prefix + digits (total 7 characters)
    """
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
    except Exception as e:
        error_message = f"Failed to extract file number in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
        return None

def cancel_operation():
    """Set the cancel flag to stop the scan or search operation."""
    global cancel_scan
    cancel_scan = True
    gui_safe_log_message("Cancel requested. Stopping operation in Outlook Miner.")
    cancel_button.config(state='disabled')
    root.config(cursor="")
    root.update_idletasks()

def delete_config():
    """Delete the configuration for the selected email from the Clients table."""
    init_db()
    recipient = recipient_combobox.get().strip()
    if not recipient:
        messagebox.showerror("Error", "No email selected to delete.")
        gui_safe_log_message("Error: No email selected to delete configuration in Outlook Miner.")
        return
    # Confirm deletion
    if not messagebox.askyesno("Confirm Delete", f"Delete configuration for '{recipient}'?"):
        return
    try:
        with db_lock:
            with sqlite3.connect('minerdb.db', timeout=10) as conn:
                c = conn.cursor()
                c.execute("DELETE FROM Clients WHERE recipient = ?", (recipient,))
                conn.commit()
                if c.rowcount > 0:
                    gui_queue.put((lambda: messagebox.showinfo("Success", f"Configuration for '{recipient}' deleted successfully."), ()))
                    gui_safe_log_message(f"Configuration for '{recipient}' deleted successfully in Outlook Miner.")
                else:
                    gui_queue.put((lambda: messagebox.showinfo("Info", f"No configuration found for '{recipient}' to delete."), ()))
                    gui_safe_log_message(f"No configuration found for '{recipient}' to delete in Outlook Miner.")
                emails = load_email_addresses()
                recipient_combobox['values'] = emails
                recipient_combobox.set("")
                start_date_entry._set_date(None)
                end_date_entry._set_date(None)
                file_number_prefix_entry.delete(0, tk.END)
                subject_keyword_entry.delete(0, tk.END)
                subject_keyword_entry.insert(0, "BILLING INVOICE")
                require_attachments_var.set(True)
                skip_forwarded_var.set(True)
                delay_seconds_entry.delete(0, tk.END)
                delay_seconds_entry.insert(0, "0")
    except Exception as e:
        error_message = f"Failed to delete config for '{recipient}' in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))

def show_configuration_dialog():
    """Show configuration dialog for advanced settings."""
    config_window = Toplevel(root)
    config_window.title("Configuration")
    config_window.geometry("400x250")
    config_window.configure(bg="#F5F5F5")
    config_window.transient(root)
    config_window.grab_set()
    try:
        icon_img = tk.PhotoImage(file=ICON_PNG_PATH)
        config_window.wm_iconphoto(True, icon_img)
    except tk.TclError:
        pass

    # File Number Prefixes
    tk.Label(config_window, text="File Number Prefixes (e.g., 759,123):", font=("Arial", 10), bg="#F5F5F5").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    prefix_entry = tk.Entry(config_window, width=30, font=("Arial", 10))
    prefix_entry.grid(row=0, column=1, padx=10, pady=10)
    prefix_entry.insert(0, file_number_prefix_entry.get())

    # Delay
    tk.Label(config_window, text="Delay (Sec.):", font=("Arial", 10), bg="#F5F5F5").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    delay_entry = tk.Entry(config_window, width=30, font=("Arial", 10))
    delay_entry.grid(row=1, column=1, padx=10, pady=10)
    delay_entry.insert(0, delay_seconds_entry.get())

    # Require Attachments
    tk.Label(config_window, text="Require Attachments:", font=("Arial", 10), bg="#F5F5F5").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    require_attach_var = tk.BooleanVar(value=require_attachments_var.get())
    tk.Checkbutton(config_window, variable=require_attach_var, bg="#F5F5F5").grid(row=2, column=1, padx=10, pady=10, sticky="w")

    # Skip Previously Forwarded
    tk.Label(config_window, text="Skip Previously Forwarded:", font=("Arial", 10), bg="#F5F5F5").grid(row=3, column=0, padx=10, pady=10, sticky="e")
    skip_fwd_var = tk.BooleanVar(value=skip_forwarded_var.get())
    tk.Checkbutton(config_window, variable=skip_fwd_var, bg="#F5F5F5").grid(row=3, column=1, padx=10, pady=10, sticky="w")

    def save_config_dialog():
        file_number_prefix_entry.delete(0, tk.END)
        file_number_prefix_entry.insert(0, prefix_entry.get())
        delay_seconds_entry.delete(0, tk.END)
        delay_seconds_entry.insert(0, delay_entry.get())
        require_attachments_var.set(require_attach_var.get())
        skip_forwarded_var.set(skip_fwd_var.get())
        config_window.destroy()

    # Buttons
    button_frame = tk.Frame(config_window, bg="#F5F5F5")
    button_frame.grid(row=4, column=0, columnspan=2, pady=20)
    ttk.Button(button_frame, text="Save", command=save_config_dialog).pack(side="left", padx=10)
    ttk.Button(button_frame, text="Cancel", command=config_window.destroy).pack(side="left", padx=10)

def show_main_menu():
    """Show hamburger menu with app options."""
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="Configuration...", command=show_configuration_dialog)
    menu.add_separator()
    recipient = recipient_combobox.get().strip()
    if recipient:
        menu.add_command(label=f"Delete '{recipient}'", command=delete_config)
    else:
        menu.add_command(label="Delete Email Config", state='disabled')
    # Position menu below the hamburger button
    menu.tk_popup(main_menu_button.winfo_rootx(),
                  main_menu_button.winfo_rooty() + main_menu_button.winfo_height())

def search_subjects_thread(config):
    """Search emails for the given configuration in a separate thread."""
    pythoncom.CoInitialize()
    try:
        global cancel_scan
        if cancel_scan:
            gui_safe_log_message("Search cancelled in Outlook Miner.")
            return 0, []
        subject_keyword = config['subject_keyword']
        start_date_str = config['start_date']
        end_date_str = config['end_date']
        skip_forwarded = config['skip_forwarded']
        recipient = config['recipient']
        file_number_prefix = config['file_number_prefix']
        file_number_prefixes = [prefix.strip() for prefix in file_number_prefix.split(',') if prefix.strip()] if file_number_prefix else []
        if not subject_keyword:
            gui_safe_log_message("Error: Subject Keyword is required in Outlook Miner.")
            gui_queue.put((lambda: messagebox.showerror("Error", "Subject Keyword is required."), ()))
            return 0, []
        if not start_date_str or not end_date_str:
            gui_safe_log_message("Error: Both Start Date and End Date are required in Outlook Miner.")
            gui_queue.put((lambda: messagebox.showerror("Error", "Both Start Date and End Date are required."), ()))
            return 0, []
        local_tz = pytz.timezone('US/Eastern')
        current_date_time = datetime.datetime.now(local_tz)
        gui_safe_log_message(f"Current client date and time: {current_date_time.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        try:
            start_date = datetime.datetime.strptime(start_date_str, "%m/%d/%Y")
            start_date = local_tz.localize(start_date)
            gui_safe_log_message(f"Start date set to {start_date_str}")
            end_date = datetime.datetime.strptime(end_date_str, "%m/%d/%Y")
            end_date = local_tz.localize(end_date + datetime.timedelta(days=1) - datetime.timedelta(seconds=1))
            gui_safe_log_message(f"End date set to {end_date_str}")
            if start_date > end_date:
                gui_safe_log_message("Error: Start date cannot be after end date in Outlook Miner.")
                gui_queue.put((lambda: messagebox.showerror("Error", "Start date cannot be after end date."), ()))
                return 0, []
            if end_date > current_date_time:
                gui_safe_log_message(f"Warning: End date {end_date_str} is in the future. Adjusting to current time.")
                end_date = current_date_time
        except ValueError as e:
            error_message = f"Error: Dates must be in MM/DD/YYYY format in Outlook Miner."
            gui_safe_log_message(error_message)
            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
            return 0, []
        gui_safe_log_message(f"Searching emails with subject containing '{subject_keyword}', skip forwarded: {skip_forwarded}, date range: {start_date_str} to {end_date_str}")
        try:
            max_attempts = 3
            for attempt in range(1, max_attempts + 1):
                try:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    break
                except Exception as e:
                    if attempt == max_attempts:
                        error_message = f"Failed to initialize Outlook after {max_attempts} attempts in Outlook Miner: {str(e)}"
                        gui_safe_log_message(error_message)
                        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                        return 0, []
                    gui_safe_log_message(f"Attempt {attempt}/{max_attempts} to initialize Outlook failed: {str(e)}. Retrying...")
                    time.sleep(1)
            try:
                outlook_version = outlook.Version
                gui_safe_log_message(f"Using Outlook version: {outlook_version}")
            except Exception as e:
                gui_safe_log_message(f"Could not retrieve Outlook version in Outlook Miner: {str(e)}")
            mapi = outlook.GetNamespace("MAPI")
            gui_safe_log_message(f"Accessing Outlook account: {mapi.CurrentUser.Name}")
            folder = get_outlook_folder(mapi)
            gui_safe_log_message("Connected to Outlook Sent Items folder")
            folder.Items.Sort("[SentOn]", True)
            gui_safe_log_message("Emails sorted by SentOn in descending order (most recent first)")
            try:
                items = folder.Items
                for i, item in enumerate(items, 1):
                    if i > 5:
                        break
                    try:
                        subject = item.Subject if item.Subject else "(No Subject)"
                        gui_safe_log_message(f"Sample email {i} subject: {subject}")
                    except Exception as e:
                        gui_safe_log_message(f"Failed to access subject of sample email {i} in Outlook Miner: {str(e)}")
            except Exception as e:
                gui_safe_log_message(f"Error accessing sample email subjects in Outlook Miner: {str(e)}")
            sanitized_subject = sanitize_filter_value(subject_keyword)
            restrict_filter = f"@SQL=\"urn:schemas:httpmail:subject\" ci_phrasematch '{sanitized_subject}'"
            gui_safe_log_message(f"Applying DASL filter: {restrict_filter}")
            try:
                filtered_items = folder.Items.Restrict(restrict_filter)
                total_emails = filtered_items.Count
                gui_safe_log_message(f"DASL filter applied: Found {total_emails} emails matching subject '{subject_keyword}'")
            except Exception as e:
                error_message = f"Error applying DASL filter in Outlook Miner: {str(e)}. Trying LIKE filter."
                gui_safe_log_message(error_message)
                gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                restrict_filter = f"[Subject] LIKE '%{sanitized_subject}%'"
                gui_safe_log_message(f"Applying filter: {restrict_filter}")
                try:
                    filtered_items = folder.Items.Restrict(restrict_filter)
                    total_emails = filtered_items.Count
                    gui_safe_log_message(f"LIKE filter applied: Found {total_emails} emails matching subject '{subject_keyword}'")
                except Exception as e:
                    error_message = f"Error applying LIKE filter in Outlook Miner: {str(e)}. Scanning all emails with client-side filtering."
                    gui_safe_log_message(error_message)
                    gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                    filtered_items = folder.Items
                    total_emails = filtered_items.Count
                    gui_safe_log_message(f"No filter applied: Scanning all {total_emails} emails in Sent Items folder with client-side filtering")
            emails_found = 0
            emails_scanned = 0
            gui_safe_log_message(f"Scanning {total_emails} emails in filtered set")
            matching_emails = []
            for i, item in enumerate(filtered_items, 1):
                if cancel_scan:
                    gui_safe_log_message(f"Search cancelled after scanning {emails_scanned} emails, found {emails_found} matching emails in Outlook Miner.")
                    break
                emails_scanned += 1
                if item.Class == 43:
                    try:
                        try:
                            subject = item.Subject if item.Subject else "(No Subject)"
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to access Subject in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                        if not subject or subject_keyword.upper() not in subject.upper():
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: Subject '{subject}' does not contain '{subject_keyword}'")
                            continue
                        file_number = None
                        if file_number_prefixes:
                            file_number = extract_file_number(item, file_number_prefixes)
                            if not file_number:
                                gui_safe_log_message(f"Email {i}/{total_emails} skipped: No valid file number found for prefixes {file_number_prefixes}")
                                continue
                        if skip_forwarded and file_number and check_if_forwarded_db(file_number, recipient):
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: File number '{file_number}' previously forwarded to {recipient}")
                            continue
                        try:
                            sent_on = item.SentOn
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to access SentOn in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                        if sent_on < start_date:
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: SentOn {sent_on.strftime('%Y-%m-%d %H:%M:%S')} before start date {start_date_str}")
                            continue
                        if sent_on > end_date:
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: SentOn {sent_on.strftime('%Y-%m-%d %H:%M:%S')} after end_date {end_date_str}")
                            continue
                        try:
                            matching_emails.append(f"[{sent_on.strftime('%Y-%m-%d %H:%M:%S')}] {subject}" + (f" (File Number: {file_number})" if file_number else ""))
                            emails_found += 1
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to access ReceivedTime in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                    except Exception as e:
                        error_message = f"Email {i}/{total_emails} skipped: General error during search in Outlook Miner: {str(e)}"
                        gui_safe_log_message(error_message)
                        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                        continue
                if i % 100 == 0:
                    gui_safe_log_message(f"Scanned {i}/{total_emails} emails, found {emails_found} matching emails...")
            gui_safe_log_message(f"Search completed: Scanned {emails_scanned} emails, found {emails_found} matching emails with '{subject_keyword}' in subject")
            if emails_found == 0:
                gui_safe_log_message(f"Suggestion: Verify the subject keyword '{subject_keyword}' matches the email subjects in Sent Items.")
                gui_safe_log_message(f"Suggestion: Adjust the date range ({start_date_str} to {end_date_str}) to include more emails.")
                if skip_forwarded:
                    gui_safe_log_message("Suggestion: Uncheck 'Skip Previously Forwarded Emails' to include previously forwarded emails.")
                if file_number_prefixes:
                    gui_safe_log_message(f"Suggestion: Verify the file number prefixes {file_number_prefixes} match attachment filenames or subjects.")
            return emails_scanned, matching_emails
        except Exception as e:
            error_message = f"Error during search in Outlook Miner: {str(e)}"
            gui_safe_log_message(error_message)
            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
            return 0, []
    finally:
        pythoncom.CoUninitialize()

def search_subjects():
    """Start the search operation in a separate thread."""
    gui_queue.put((lambda: cancel_button.config(state='normal'), ()))
    gui_queue.put((lambda: root.config(cursor="wait"), ()))
    root.update_idletasks()
    def process_search():
        global cancel_scan
        cancel_scan = False
        config = {
            'recipient': recipient_combobox.get().strip(),
            'start_date': start_date_entry.get(),
            'end_date': end_date_entry.get(),
            'file_number_prefix': file_number_prefix_entry.get().strip(),
            'subject_keyword': subject_keyword_entry.get().strip(),
            'require_attachments': require_attachments_var.get(),
            'skip_forwarded': skip_forwarded_var.get(),
            'delay_seconds': delay_seconds_entry.get().strip()
        }
        scanned, matching_emails = search_subjects_thread(config)
        def show_results():
            subjects_window = Toplevel(root)
            subjects_window.title("Outlook Miner - Matching Email Subjects")
            try:
                icon_img = tk.PhotoImage(file=ICON_PNG_PATH)
                subjects_window.wm_iconphoto(True, icon_img)
            except tk.TclError as e:
                gui_safe_log_message(f"Failed to load icon for subjects window in Outlook Miner: {str(e)}. Using default icon.")
            subjects_window.geometry("600x400")
            subjects_window.configure(bg="#F5F5F5")
            subjects_text = ScrolledText(subjects_window, width=80, height=20, state='normal', bg="#FFFFFF", fg="#000000", font=("Arial", 10), highlightbackground="#E0E0E0", highlightthickness=1)
            subjects_text.pack(padx=10, pady=10)
            if matching_emails:
                subjects_text.insert(tk.END, f"Found {len(matching_emails)} matching emails:\n\n")
                for email_info in matching_emails:
                    subjects_text.insert(tk.END, f"{email_info}\n")
            else:
                subjects_text.insert(tk.END, "No matching emails found.")
                gui_safe_log_message("No matching emails found in Outlook Miner. Check the log for sample email subjects to verify the subject keyword.")
            ttk.Button(subjects_window, text="Clear Log", command=lambda: subjects_text.delete(1.0, tk.END)).pack(pady=5)
            ttk.Button(subjects_window, text="Close", command=subjects_window.destroy).pack(pady=5)
        gui_queue.put((show_results, ()))
        gui_queue.put((lambda: messagebox.showinfo("Search Complete", f"Scanned {scanned} emails, found {len(matching_emails)} matching emails."), ()))
        gui_queue.put((lambda: cancel_button.config(state='disabled'), ()))
        gui_queue.put((lambda: root.config(cursor=""), ()))
    threading.Thread(target=process_search, daemon=True).start()

def show_warning_message():
    """Display a warning message box that closes after 3 seconds."""
    warning_window = Toplevel(root)
    warning_window.title("Warning")
    try:
        icon_img = tk.PhotoImage(file=ICON_PNG_PATH)
        warning_window.wm_iconphoto(True, icon_img)
    except tk.TclError as e:
        gui_safe_log_message(f"Failed to load icon for warning window in Outlook Miner: {str(e)}. Using default icon.")
    warning_window.geometry("400x100")
    warning_window.configure(bg="#F5F5F5")
    tk.Label(warning_window, text="Warning: A date range exceeding 8 days will add a 3 second delay between forwarded emails.", 
             font=("Arial", 10), bg="#F5F5F5", fg="#000000", wraplength=380).pack(pady=20)
    root.after(3000, warning_window.destroy)

def scan_and_forward_thread(config):
    """Forward emails for the given configuration in a separate thread."""
    pythoncom.CoInitialize()
    try:
        global cancel_scan
        if cancel_scan:
            gui_safe_log_message("Forward operation cancelled in Outlook Miner.")
            return 0, 0
        recipient = config['recipient']
        subject_keyword = config['subject_keyword']
        start_date_str = config['start_date']
        end_date_str = config['end_date']
        file_number_prefix = config['file_number_prefix']
        file_number_prefixes = [prefix.strip() for prefix in file_number_prefix.split(',') if prefix.strip()] if file_number_prefix else []
        require_attachments = config['require_attachments']
        skip_forwarded = config['skip_forwarded']
        delay_seconds = float(config['delay_seconds']) if config['delay_seconds'] and config['delay_seconds'].strip() else 0.0
        if not recipient or not validate_email(recipient):
            gui_safe_log_message("Error: Forward To email is invalid in Outlook Miner.")
            gui_queue.put((lambda: messagebox.showerror("Error", "Forward To email is required and must be a valid email address."), ()))
            return 0, 0
        if not subject_keyword:
            gui_safe_log_message("Error: Subject Keyword is required in Outlook Miner.")
            gui_queue.put((lambda: messagebox.showerror("Error", "Subject Keyword is required."), ()))
            return 0, 0
        if not start_date_str or not end_date_str:
            gui_safe_log_message("Error: Both Start Date and End Date are required in Outlook Miner.")
            gui_queue.put((lambda: messagebox.showerror("Error", "Both Start Date and End Date are required."), ()))
            return 0, 0
        for prefix in file_number_prefixes:
            if not re.match(r'^\d+$', prefix):
                gui_safe_log_message(f"Error: File number prefix '{prefix}' must be numeric in Outlook Miner.")
                gui_queue.put((lambda: messagebox.showerror("Error", f"File number prefix '{prefix}' must be numeric."), ()))
                return 0, 0
        local_tz = pytz.timezone('US/Eastern')
        current_date_time = datetime.datetime.now(local_tz)
        gui_safe_log_message(f"Current client date and time: {current_date_time.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        try:
            start_date = datetime.datetime.strptime(start_date_str, "%m/%d/%Y")
            start_date = local_tz.localize(start_date)
            gui_safe_log_message(f"Start date set to {start_date_str}")
            end_date = datetime.datetime.strptime(end_date_str, "%m/%d/%Y")
            end_date = local_tz.localize(end_date + datetime.timedelta(days=1) - datetime.timedelta(seconds=1))
            gui_safe_log_message(f"End date set to {end_date_str}")
            if start_date > end_date:
                gui_safe_log_message("Error: Start date cannot be after end date in Outlook Miner.")
                gui_queue.put((lambda: messagebox.showerror("Error", "Start date cannot be after end date."), ()))
                return 0, 0
            if end_date > current_date_time:
                gui_safe_log_message(f"Warning: End date {end_date_str} is in the future. Adjusting to current time.")
                end_date = current_date_time
            # Check date range and apply 3-second delay if > 8 days
            date_range_days = (end_date.date() - start_date.date()).days
            if date_range_days > 8:
                gui_safe_log_message(f"Date range of {date_range_days} days exceeds 8 days. Applying 3-second delay between forwarded emails.")
                gui_queue.put((show_warning_message, ()))
                delay_seconds = 3.0  # Override user-entered delay
        except ValueError as e:
            error_message = f"Error: Dates must be in MM/DD/YYYY format in Outlook Miner."
            gui_safe_log_message(error_message)
            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
            return 0, 0
        gui_safe_log_message(f"Filtering emails with subject containing '{subject_keyword}', forwarding to '{recipient}', require attachments: {require_attachments}, skip forwarded: {skip_forwarded}, date range: {start_date_str} to {end_date_str}, file number prefixes: {', '.join(file_number_prefixes) or 'None'}, delay: {delay_seconds} seconds")
        try:
            gui_queue.put((lambda: subject_text.delete(1.0, tk.END), ()))
            gui_safe_log_message("Cleared subject display window")
            max_attempts = 3
            for attempt in range(1, max_attempts + 1):
                try:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    break
                except Exception as e:
                    if attempt == max_attempts:
                        error_message = f"Failed to initialize Outlook after {max_attempts} attempts in Outlook Miner: {str(e)}"
                        gui_safe_log_message(error_message)
                        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                        return 0, 0
                    gui_safe_log_message(f"Attempt {attempt}/{max_attempts} to initialize Outlook failed: {str(e)}. Retrying...")
                    time.sleep(1)
            try:
                outlook_version = outlook.Version
                gui_safe_log_message(f"Using Outlook version: {outlook_version}")
            except Exception as e:
                gui_safe_log_message(f"Could not retrieve Outlook version in Outlook Miner: {str(e)}")
            mapi = outlook.GetNamespace("MAPI")
            gui_safe_log_message(f"Accessing Outlook account: {mapi.CurrentUser.Name}")
            folder = get_outlook_folder(mapi)
            gui_safe_log_message("Connected to Outlook Sent Items folder")
            folder.Items.Sort("[SentOn]", True)
            gui_safe_log_message("Emails sorted by SentOn in descending order (most recent first)")
            try:
                items = folder.Items
                for i, item in enumerate(items, 1):
                    if i > 5:
                        break
                    try:
                        subject = item.Subject if item.Subject else "(No Subject)"
                        gui_safe_log_message(f"Sample email {i} subject: {subject}")
                    except Exception as e:
                        gui_safe_log_message(f"Failed to access subject of sample email {i} in Outlook Miner: {str(e)}")
            except Exception as e:
                gui_safe_log_message(f"Error accessing sample email subjects in Outlook Miner: {str(e)}")
            sanitized_subject = sanitize_filter_value(subject_keyword)
            restrict_filter = f"@SQL=\"urn:schemas:httpmail:subject\" ci_phrasematch '{sanitized_subject}'"
            gui_safe_log_message(f"Applying DASL filter: {restrict_filter}")
            try:
                filtered_items = folder.Items.Restrict(restrict_filter)
                total_emails = filtered_items.Count
                gui_safe_log_message(f"DASL filter applied: Found {total_emails} emails matching subject '{subject_keyword}'")
            except Exception as e:
                error_message = f"Error applying DASL filter in Outlook Miner: {str(e)}. Trying LIKE filter."
                gui_safe_log_message(error_message)
                gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                restrict_filter = f"[Subject] LIKE '%{sanitized_subject}%'"
                gui_safe_log_message(f"Applying filter: {restrict_filter}")
                try:
                    filtered_items = folder.Items.Restrict(restrict_filter)
                    total_emails = filtered_items.Count
                    gui_safe_log_message(f"LIKE filter applied: Found {total_emails} emails matching subject '{subject_keyword}'")
                except Exception as e:
                    error_message = f"Error applying LIKE filter in Outlook Miner: {str(e)}. Scanning all emails with client-side filtering."
                    gui_safe_log_message(error_message)
                    gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                    filtered_items = folder.Items
                    total_emails = filtered_items.Count
                    gui_safe_log_message(f"No filter applied: Scanning all {total_emails} emails in Sent Items folder with client-side filtering")
            emails_processed = 0
            emails_scanned = 0
            gui_safe_log_message(f"Scanning {total_emails} emails in filtered set")
            for i, item in enumerate(filtered_items, 1):
                if cancel_scan:
                    gui_safe_log_message(f"Operation cancelled after scanning {emails_scanned} emails, forwarded {emails_processed} emails in Outlook Miner.")
                    break
                emails_scanned += 1
                if item.Class == 43:
                    try:
                        try:
                            subject = item.Subject if item.Subject else "(No Subject)"
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to access Subject in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                        if not subject or subject_keyword.upper() not in subject.upper():
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: Subject '{subject}' does not contain '{subject_keyword}'")
                            continue
                        file_number = None
                        if file_number_prefixes:
                            file_number = extract_file_number(item, file_number_prefixes)
                            if not file_number:
                                gui_safe_log_message(f"Email {i}/{total_emails} skipped: No valid file number found for prefixes {file_number_prefixes}")
                                continue
                        if skip_forwarded and file_number and check_if_forwarded_db(file_number, recipient):
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: File number '{file_number}' previously forwarded to {recipient}")
                            continue
                        try:
                            sent_on = item.SentOn
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to access SentOn in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                        if sent_on < start_date:
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: SentOn {sent_on.strftime('%Y-%m-%d %H:%M:%S')} before start date {start_date_str}")
                            continue
                        if sent_on > end_date:
                            gui_safe_log_message(f"Email {i}/{total_emails} skipped: SentOn {sent_on.strftime('%Y-%m-%d %H:%M:%S')} after end date {end_date_str}")
                            continue
                        try:
                            if require_attachments and item.Attachments.Count == 0:
                                gui_safe_log_message(f"Email {i}/{total_emails} skipped: No attachments")
                                continue
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to access Attachments in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                        new_subject = file_number if file_number else subject
                        try:
                            forward_email = item.Forward()
                            forward_email.To = recipient
                            forward_email.Subject = new_subject
                            forward_email.Send()
                            emails_processed += 1
                            gui_safe_log_message(f"Forwarded email {i}/{total_emails} with subject '{new_subject}' to {recipient}" + (f" (File Number: {file_number})" if file_number else ""))
                            log_forwarded_subject(new_subject)
                            gui_safe_display_subject(new_subject)
                            if file_number:
                                log_forwarded_email(file_number, recipient)
                            if delay_seconds > 0:
                                gui_safe_log_message(f"Waiting {delay_seconds} seconds before next email")
                                time.sleep(delay_seconds)
                        except Exception as e:
                            error_message = f"Email {i}/{total_emails} skipped: Failed to forward email in Outlook Miner: {str(e)}"
                            gui_safe_log_message(error_message)
                            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                            continue
                    except Exception as e:
                        error_message = f"Email {i}/{total_emails} skipped: General error processing email in Outlook Miner: {str(e)}"
                        gui_safe_log_message(error_message)
                        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
                        continue
                if i % 100 == 0:
                    gui_safe_log_message(f"Scanned {i}/{total_emails} emails, forwarded {emails_processed}...")
            if emails_processed == 0:
                gui_safe_log_message(f"No emails matched the criteria (subject containing '{subject_keyword}', date range: {start_date_str} to {end_date_str}, attachments: {require_attachments}, not previously forwarded: {skip_forwarded}).")
                if require_attachments:
                    gui_safe_log_message("Suggestion: Uncheck 'Require Attachments' to include emails without attachments.")
                if skip_forwarded:
                    gui_safe_log_message("Suggestion: Uncheck 'Skip Previously Forwarded Emails' to include previously forwarded emails.")
                if file_number_prefixes:
                    gui_safe_log_message(f"Suggestion: Verify the file number prefixes {file_number_prefixes} match attachment filenames or subjects.")
                gui_safe_log_message(f"Suggestion: Verify the subject keyword '{subject_keyword}' matches the email subjects in Sent Items.")
                gui_safe_log_message(f"Suggestion: Adjust the date range ({start_date_str} to {end_date_str}) to include more emails.")
            gui_safe_log_message(f"Success: Scanned {emails_scanned} emails, forwarded {emails_processed} emails from Sent Items")
            return emails_scanned, emails_processed
        except Exception as e:
            error_message = f"Error during forwarding in Outlook Miner: {str(e)}"
            gui_safe_log_message(error_message)
            gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))
            return 0, 0
    finally:
        pythoncom.CoUninitialize()

def scan_and_forward():
    """Start the forward operation in a separate thread, saving config automatically."""
    # Validate inputs before starting
    recipient = recipient_combobox.get().strip()
    subject_keyword = subject_keyword_entry.get().strip()
    delay_seconds_str = delay_seconds_entry.get().strip()
    start_date = start_date_entry.get().strip()
    end_date = end_date_entry.get().strip()

    if not recipient or not validate_email(recipient):
        messagebox.showerror("Error", "Forward To email is required and must be a valid email address.")
        gui_safe_log_message("Error: Forward To email is required and must be a valid email address in Outlook Miner.")
        return
    if not subject_keyword:
        messagebox.showerror("Error", "Subject Keyword is required.")
        gui_safe_log_message("Error: Subject Keyword is required in Outlook Miner.")
        return
    if not start_date or not end_date:
        messagebox.showerror("Error", "Both Start Date and End Date are required.")
        gui_safe_log_message("Error: Both Start Date and End Date are required in Outlook Miner.")
        return
    try:
        datetime.datetime.strptime(start_date, "%m/%d/%Y")
        datetime.datetime.strptime(end_date, "%m/%d/%Y")
    except ValueError:
        messagebox.showerror("Error", "Dates must be in MM/DD/YYYY format.")
        gui_safe_log_message("Error: Dates must be in MM/DD/YYYY format in Outlook Miner.")
        return
    for prefix in file_number_prefix_entry.get().strip().split(','):
        prefix = prefix.strip()
        if prefix and not re.match(r'^\d+$', prefix):
            messagebox.showerror("Error", f"File number prefix '{prefix}' must be numeric.")
            gui_safe_log_message(f"Error: File number prefix '{prefix}' must be numeric in Outlook Miner.")
            return
    delay_seconds_val = 0.0
    if delay_seconds_str:
        try:
            delay_seconds_val = float(delay_seconds_str)
            if delay_seconds_val < 0:
                raise ValueError("Delay (Sec.) must be non-negative.")
        except ValueError:
            messagebox.showerror("Error", "Delay (Sec.) must be a non-negative number.")
            gui_safe_log_message("Error: Delay (Sec.) must be a non-negative number in Outlook Miner.")
            return

    # Save configuration to database
    file_number_prefix = file_number_prefix_entry.get().strip()
    require_attachments_db = "1" if require_attachments_var.get() else "0"
    skip_forwarded_db = "1" if skip_forwarded_var.get() else "0"
    created_at = datetime.datetime.now(pytz.timezone('US/Eastern')).strftime("%Y-%m-%d %H:%M:%S")
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
                           require_attachments_db, skip_forwarded_db, str(delay_seconds_val), created_at, "", ""))
                conn.commit()
        gui_safe_log_message(f"Configuration saved for '{recipient}'")
        # Update combobox values
        emails = load_email_addresses()
        recipient_combobox['values'] = emails
        # Save last used settings
        save_last_used_email(recipient)
        save_setting('last_start_date', start_date)
        save_setting('last_end_date', end_date)
    except Exception as e:
        error_message = f"Failed to save config in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)

    gui_queue.put((lambda: cancel_button.config(state='normal'), ()))
    gui_queue.put((lambda: root.config(cursor="wait"), ()))
    root.update_idletasks()
    def process_forward():
        global cancel_scan
        cancel_scan = False
        config = {
            'recipient': recipient,
            'start_date': start_date,
            'end_date': end_date,
            'file_number_prefix': file_number_prefix,
            'subject_keyword': subject_keyword,
            'require_attachments': require_attachments_var.get(),
            'skip_forwarded': skip_forwarded_var.get(),
            'delay_seconds': delay_seconds_str
        }
        scanned, processed = scan_and_forward_thread(config)
        gui_queue.put((lambda: messagebox.showinfo("Success", f"Scanned {scanned} emails, forwarded {processed} emails from Sent Items."), ()))
        gui_queue.put((lambda: cancel_button.config(state='disabled'), ()))
        gui_queue.put((lambda: root.config(cursor=""), ()))
    threading.Thread(target=process_forward, daemon=True).start()

def initialize_log_frame(event):
    """Initialize the log frame when the Log tab is selected."""
    global log_frame_initialized, log_text
    if not log_frame_initialized:
        log_frame = ttk.Frame(notebook, style='Clam.TFrame')
        notebook.add(log_frame, text="Log")
        log_text = ScrolledText(log_frame, width=80, height=20, state='normal', bg="#FFFFFF", fg="#000000", font=("Arial", 10), highlightbackground="#E0E0E0", highlightthickness=1)
        log_text.pack(padx=10, pady=10)
        log_frame_initialized = True
        gui_safe_log_message("Log frame initialized in Outlook Miner.")

def initialize_buttons():
    """Initialize the button frame after startup."""
    global button_frame, cancel_button, main_menu_button
    button_frame = ttk.Frame(search_frame, style='Clam.TFrame')
    button_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")
    ttk.Button(button_frame, text="Preview", command=search_subjects).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Scan and Forward", command=scan_and_forward).pack(side="left", padx=5)
    cancel_button = ttk.Button(button_frame, text="Cancel", command=cancel_operation, state='disabled')
    cancel_button.pack(side="left", padx=5)
    # Add hamburger menu button to top right of main window
    main_menu_button = tk.Button(root, text="☰", font=("Arial", 12), width=3, command=show_main_menu,
                                  bg="#E0E0E0", relief="flat", cursor="hand2")
    main_menu_button.place(relx=1.0, x=-10, y=5, anchor="ne")
    status_label.destroy()
    root.after(200, process_gui_queue)
    root.update_idletasks()

def initialize_app():
    """Initialize the application with minimal resource usage."""
    try:
        gui_safe_log_message("Application initialization completed in Outlook Miner.")
    except Exception as e:
        error_message = f"Initialization failed in Outlook Miner: {str(e)}"
        gui_safe_log_message(error_message)
        gui_queue.put((lambda err=error_message: messagebox.showerror("Error", err), ()))

# ============================================================================
# COLOR SCHEME AND STYLING CONSTANTS
# ============================================================================
COLORS = {
    'primary': '#2563EB',           # Primary blue (buttons, accents)
    'primary_hover': '#1D4ED8',     # Darker blue for hover
    'primary_light': '#3B82F6',     # Lighter blue for highlights
    'header_bg': '#1E3A5F',         # Dark blue for header
    'header_text': '#FFFFFF',       # White text on header
    'bg': '#F8FAFC',                # Light gray background
    'frame_bg': '#FFFFFF',          # White frame background
    'border': '#E2E8F0',            # Light border color
    'text': '#1E293B',              # Dark text
    'text_secondary': '#64748B',    # Secondary gray text
    'input_bg': '#FFFFFF',          # White input background
    'input_border': '#CBD5E1',      # Input border
    'success': '#10B981',           # Green for success
    'warning': '#F59E0B',           # Orange for warning
}

FONTS = {
    'brand': ('Segoe UI', 16, 'bold'),
    'brand_accent': ('Segoe UI', 16),
    'heading': ('Segoe UI', 11, 'bold'),
    'label': ('Segoe UI', 10),
    'input': ('Segoe UI', 10),
    'button': ('Segoe UI', 10),
    'small': ('Segoe UI', 9),
}

# Create main GUI
root = tk.Tk()
root.title("Outlook Miner")
root.configure(bg=COLORS['bg'])
root.minsize(600, 550)
try:
    # Use PhotoImage with PNG for better quality rendering in title bar
    icon_photo = tk.PhotoImage(file=ICON_PNG_PATH)
    root.wm_iconphoto(True, icon_photo)
except tk.TclError as e:
    # Fallback to ICO if PNG fails
    try:
        root.iconbitmap(ICON_PATH)
    except tk.TclError:
        gui_safe_log_message(f"Failed to load icon for main window in Outlook Miner: {str(e)}. Using default icon.")

# Apply clam theme and configure styles
style = ttk.Style()
style.theme_use('clam')

# Configure custom styles
style.configure("App.TFrame", background=COLORS['bg'])
style.configure("Card.TFrame", background=COLORS['frame_bg'])
style.configure("Header.TFrame", background=COLORS['header_bg'])

# Tab/Notebook styling
style.configure("TNotebook", background=COLORS['bg'], borderwidth=0)
style.configure("TNotebook.Tab",
    background=COLORS['border'],
    foreground=COLORS['text'],
    padding=[20, 8],
    font=FONTS['label'])
style.map("TNotebook.Tab",
    background=[("selected", COLORS['primary']), ("active", COLORS['primary_light'])],
    foreground=[("selected", "#FFFFFF"), ("active", "#FFFFFF")])

# Button styling
style.configure("Primary.TButton",
    background=COLORS['primary'],
    foreground="#FFFFFF",
    padding=[16, 8],
    font=FONTS['button'])
style.map("Primary.TButton",
    background=[("active", COLORS['primary_hover']), ("pressed", COLORS['primary_hover'])])

style.configure("Secondary.TButton",
    background=COLORS['border'],
    foreground=COLORS['text'],
    padding=[16, 8],
    font=FONTS['button'])
style.map("Secondary.TButton",
    background=[("active", "#D1D5DB"), ("pressed", "#D1D5DB")])

# Combobox styling
style.configure("TCombobox",
    fieldbackground=COLORS['input_bg'],
    background=COLORS['input_bg'],
    foreground=COLORS['text'],
    padding=5,
    font=FONTS['input'])

# LabelFrame styling
style.configure("Card.TLabelframe",
    background=COLORS['frame_bg'],
    bordercolor=COLORS['border'],
    relief="solid",
    borderwidth=1)
style.configure("Card.TLabelframe.Label",
    background=COLORS['frame_bg'],
    foreground=COLORS['primary'],
    font=FONTS['heading'])

# ============================================================================
# HEADER / BRANDING
# ============================================================================
header_frame = tk.Frame(root, bg=COLORS['header_bg'], height=60)
header_frame.pack(fill='x', side='top')
header_frame.pack_propagate(False)

# App title with accent
title_frame = tk.Frame(header_frame, bg=COLORS['header_bg'])
title_frame.pack(side='left', padx=20, pady=12)
tk.Label(title_frame, text="Outlook", font=FONTS['brand'],
         bg=COLORS['header_bg'], fg=COLORS['header_text']).pack(side='left')
tk.Label(title_frame, text="Miner", font=FONTS['brand_accent'],
         bg=COLORS['header_bg'], fg=COLORS['primary_light']).pack(side='left')

# Loading status (temporary)
status_label = tk.Label(root, text="Loading...", font=FONTS['label'], bg=COLORS['bg'], fg=COLORS['text'])
status_label.pack(pady=20)

# Main content area
main_content = tk.Frame(root, bg=COLORS['bg'])
main_content.pack(fill='both', expand=True, padx=15, pady=10)

notebook = ttk.Notebook(main_content)
notebook.pack(pady=0, fill='both', expand=True)
notebook.bind("<<NotebookTabChanged>>", initialize_log_frame)
search_frame = ttk.Frame(notebook, style='Card.TFrame')
notebook.add(search_frame, text="  Search  ")

current_year = datetime.datetime.now().year
current_month = datetime.datetime.now().month

# Configure grid weights for proper expansion
search_frame.columnconfigure(0, weight=0)
search_frame.columnconfigure(1, weight=1)

# ============================================================================
# EMAIL SETTINGS SECTION
# ============================================================================
email_frame = ttk.LabelFrame(search_frame, text="Email Settings", style="Card.TLabelframe", padding=(15, 10))
email_frame.grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 8), sticky="ew")
email_frame.columnconfigure(1, weight=1)

tk.Label(email_frame, text="Forward To:", font=FONTS['label'],
         bg=COLORS['frame_bg'], fg=COLORS['text']).grid(row=0, column=0, padx=(0, 10), pady=8, sticky="e")
recipient_combobox = ttk.Combobox(email_frame, width=45, font=FONTS['input'])
recipient_combobox.grid(row=0, column=1, padx=0, pady=8, sticky="ew")
init_db()  # Initialize database before loading email addresses
recipient_combobox['values'] = load_email_addresses()
recipient_combobox.bind("<<ComboboxSelected>>", load_config_for_email)
# Restore last used email address
last_email = load_last_used_email()
if last_email and last_email in recipient_combobox['values']:
    recipient_combobox.set(last_email)
    # Trigger config load for the restored email (simulate selection event)
    root.after(100, lambda: load_config_for_email(None))

tk.Label(email_frame, text="Subject Keyword:", font=FONTS['label'],
         bg=COLORS['frame_bg'], fg=COLORS['text']).grid(row=1, column=0, padx=(0, 10), pady=8, sticky="e")
subject_keyword_entry = tk.Entry(email_frame, width=48, font=FONTS['input'],
                                  bg=COLORS['input_bg'], fg=COLORS['text'],
                                  highlightbackground=COLORS['input_border'], highlightthickness=1,
                                  relief='solid', bd=1)
subject_keyword_entry.grid(row=1, column=1, padx=0, pady=8, sticky="ew")
subject_keyword_entry.insert(0, "BILLING INVOICE")

# ============================================================================
# DATE RANGE SECTION
# ============================================================================
date_frame = ttk.LabelFrame(search_frame, text="Date Range", style="Card.TLabelframe", padding=(15, 10))
date_frame.grid(row=1, column=0, columnspan=2, padx=15, pady=8, sticky="ew")
date_frame.columnconfigure(1, weight=1)
date_frame.columnconfigure(3, weight=1)

tk.Label(date_frame, text="Start Date:", font=FONTS['label'],
         bg=COLORS['frame_bg'], fg=COLORS['text']).grid(row=0, column=0, padx=(0, 10), pady=8, sticky="e")
start_date_entry = DateEntry(date_frame, width=18, date_pattern='mm/dd/yyyy',
                              year=current_year, month=current_month, selectmode='day',
                              font=FONTS['input'], background=COLORS['primary'],
                              foreground="#FFFFFF", normalbackground=COLORS['input_bg'],
                              normalforeground=COLORS['text'])
start_date_entry.grid(row=0, column=1, padx=(0, 20), pady=8, sticky="w")

tk.Label(date_frame, text="End Date:", font=FONTS['label'],
         bg=COLORS['frame_bg'], fg=COLORS['text']).grid(row=0, column=2, padx=(0, 10), pady=8, sticky="e")
end_date_entry = DateEntry(date_frame, width=18, date_pattern='mm/dd/yyyy',
                            year=current_year, month=current_month, selectmode='day',
                            font=FONTS['input'], background=COLORS['primary'],
                            foreground="#FFFFFF", normalbackground=COLORS['input_bg'],
                            normalforeground=COLORS['text'])
end_date_entry.grid(row=0, column=3, padx=0, pady=8, sticky="w")

# Restore last used dates (only if no email config will be loaded)
if not last_email or last_email not in recipient_combobox['values']:
    last_start = load_setting('last_start_date')
    last_end = load_setting('last_end_date')
    if last_start:
        try:
            start_date_entry.set_date(last_start)
        except ValueError:
            pass
    if last_end:
        try:
            end_date_entry.set_date(last_end)
        except ValueError:
            pass

# Hidden configuration variables (accessed via Configuration menu)
file_number_prefix_entry = tk.Entry(search_frame)  # Hidden, stores value
file_number_prefix_entry.insert(0, "")
delay_seconds_entry = tk.Entry(search_frame)  # Hidden, stores value
delay_seconds_entry.insert(0, "0")
require_attachments_var = tk.BooleanVar(value=True)
skip_forwarded_var = tk.BooleanVar(value=True)

# ============================================================================
# RESULTS SECTION
# ============================================================================
results_frame = ttk.LabelFrame(search_frame, text="Files Sent", style="Card.TLabelframe", padding=(15, 10))
results_frame.grid(row=2, column=0, columnspan=2, padx=15, pady=8, sticky="nsew")
results_frame.columnconfigure(0, weight=1)
results_frame.rowconfigure(0, weight=1)
search_frame.rowconfigure(2, weight=1)

subject_text = ScrolledText(results_frame, width=50, height=8, state='normal',
                             bg=COLORS['input_bg'], fg=COLORS['text'],
                             font=FONTS['input'], highlightbackground=COLORS['input_border'],
                             highlightthickness=1, relief='solid', bd=1)
subject_text.grid(row=0, column=0, sticky="nsew")

root.after(100, initialize_buttons)
initialize_app()
root.mainloop()