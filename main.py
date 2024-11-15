import sys
import os
import json
import logging
import requests
import subprocess
import urllib.request
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QTextEdit, QVBoxLayout, QWidget,
    QDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QComboBox, QAction,
    QFileDialog, QCheckBox, QSystemTrayIcon, QMenu, QGraphicsDropShadowEffect, 
    QStyle, QScrollArea, QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDateTime, QUrl
from PyQt5.QtGui import QIcon, QColor, QBrush, QDesktopServices

# Constants
SETTINGS_FILE = r"C:\TSTP\Drive_Mapper\Settings\drive_settings.json"
LOG_FILE = r"C:\TSTP\Drive_Mapper\Logs\DriveManager.log"

# Ensure directories exist
os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

# Create settings file with default content if it does not exist
if not os.path.exists(SETTINGS_FILE):
    default_settings = {
        "drive_mappings": [],
        "startup_enabled": False,
        "auto_readd_enabled": False,
        "light_mode": False
    }
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(default_settings, f, indent=4)

# Create log file if it does not exist
if not os.path.exists(LOG_FILE):
    open(LOG_FILE, 'w').close()

APP_ICON = os.path.join(os.path.dirname(__file__), "app_icon.ico")

# Configure Logging
os.makedirs('logs', exist_ok=True)

# Create log files if they do not exist
for log_file in [LOG_FILE, f"logs/DriveManager_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"]:
    if not os.path.exists(log_file):
        open(log_file, 'w').close()

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Main log handler
main_handler = logging.FileHandler(LOG_FILE)
main_handler.setLevel(logging.DEBUG)
main_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
main_handler.setFormatter(main_formatter)

# Timestamped log handler
timestamped_log_file = f"logs/DriveManager_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
timestamp_handler = logging.FileHandler(timestamped_log_file)
timestamp_handler.setLevel(logging.DEBUG)
timestamp_handler.setFormatter(main_formatter)

# Stream handler for console output
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
console_handler.setFormatter(console_formatter)

logger.addHandler(main_handler)
logger.addHandler(timestamp_handler)
logger.addHandler(console_handler)

# Helper Functions
def load_settings():
    """
    Loads drive mappings and settings from the settings file.
    Migrates 'DriveLetter' to 'Drive' if necessary.
    Recreates the settings file if it is invalid.
    """
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r") as f:
                settings = json.load(f)
                drive_mappings = settings.get("drive_mappings", [])
                migrated = False
                for mapping in drive_mappings:
                    if "DriveLetter" in mapping:
                        mapping["Drive"] = mapping.pop("DriveLetter")
                        migrated = True
                if migrated:
                    save_settings(drive_mappings, settings.get("startup_enabled", False), 
                                  settings.get("auto_readd_enabled", False), 
                                  settings.get("light_mode", False))  # Save migrated settings
                    logger.info(f"Migrated 'DriveLetter' to 'Drive' in {SETTINGS_FILE}.")
                else:
                    logger.info(f"Loaded drive settings from {SETTINGS_FILE}.")
                # Load additional settings
                startup_enabled = settings.get("startup_enabled", False)
                auto_readd_enabled = settings.get("auto_readd_enabled", False)
                light_mode = settings.get("light_mode", False)
                return drive_mappings, startup_enabled, auto_readd_enabled, light_mode
        except (json.JSONDecodeError, Exception) as e:
            logger.error(f"Error loading settings: {e}. Recreating settings file.")
            QMessageBox.critical(None, "Error", f"Failed to load settings. Recreating settings file:\n{e}")
            save_settings([], False, False, False)
            return [], False, False, False
    else:
        logger.info("Settings file not found. Starting with default settings.")
        return [], False, False, False

def save_settings(drive_mappings, startup_enabled, auto_readd_enabled, light_mode):
    """
    Saves drive mappings and settings to the settings file.
    """
    try:
        current_settings = {}
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                current_settings = json.load(f)
        current_settings["drive_mappings"] = drive_mappings
        current_settings["startup_enabled"] = startup_enabled
        current_settings["auto_readd_enabled"] = auto_readd_enabled
        current_settings["light_mode"] = light_mode
        with open(SETTINGS_FILE, "w") as f:
            json.dump(current_settings, f, indent=4)
            logger.info(f"Settings saved to {SETTINGS_FILE}.")
    except Exception as e:
        logger.error(f"Error saving settings: {e}")
        QMessageBox.critical(None, "Error", f"Failed to save settings:\n{e}")

def normalize_drive_letter(drive_letter):
    """
    Ensures the drive letter is in the correct format (e.g., 'A:')
    """
    drive_letter = drive_letter.strip().upper()
    if not drive_letter.endswith(":"):
        drive_letter = f"{drive_letter}:"
    return drive_letter

def execute_cmd(command):
    """
    Executes a CMD command and returns the output and error.
    """
    try:
        completed_process = subprocess.run(
            command,
            shell=True,
            capture_output=True,
            text=True
        )
        return completed_process.stdout.strip(), completed_process.stderr.strip()
    except Exception as e:
        logger.error(f"Exception during CMD execution: {e}")
        return "", str(e)

def get_current_mapped_drives():
    """
    Retrieves currently mapped network drives using 'net use'.
    Returns a list of dictionaries with Drive and UNCPath.
    """
    stdout, stderr = execute_cmd("net use")
    drives = []
    if stdout:
        lines = stdout.splitlines()
        for line in lines:
            if line.startswith("OK") or line.startswith("Disconnected") or line.startswith("Connecting"):
                parts = line.split()
                if len(parts) >= 3:
                    drive_letter = parts[1]
                    unc_path = parts[2]
                    drives.append({"Drive": drive_letter, "UNCPath": unc_path})
    elif stderr:
        logger.error(f"Error retrieving mapped drives: {stderr}")
    return drives

def get_free_drive_letters(existing_letters=None):
    """
    Retrieves a list of free drive letters excluding those in existing_letters and currently used.
    """
    try:
        current_drives = get_current_mapped_drives()
        used_letters = [drive["Drive"].upper() for drive in current_drives]
        if existing_letters:
            used_letters.extend([dl.upper() for dl in existing_letters])
        all_letters = [f"{chr(i)}:" for i in range(ord('A'), ord('Z')+1)]
        free_letters = [letter for letter in all_letters if letter not in used_letters]
        return free_letters
    except Exception as e:
        logger.error(f"Error fetching free drive letters: {e}")
        return [f"{chr(i)}:" for i in range(ord('A'), ord('Z')+1)]

# Thread for Mapping Drives
class MapDrivesThread(QThread):
    log_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, drive_mappings, map_now=True):
        super().__init__()
        self.drive_mappings = drive_mappings
        self.map_now = map_now

    def run(self):
        self.log_signal.emit("Starting to map network drives...")
        for mapping in self.drive_mappings:
            if not self.map_now and not mapping.get("Selected", False):
                continue  # Skip if not selected for mapping and map_now is False
            drive_letter = mapping["Drive"]
            unc_path = mapping["UNCPath"]
            use_credentials = mapping.get("UseCredentials", False)
            username = mapping.get("Username", "")
            password = mapping.get("Password", "")
            self.log_signal.emit(f"Processing drive {drive_letter} -> {unc_path}...")

            # Check if drive is already mapped
            is_mapped = self.is_drive_mapped(drive_letter, unc_path)
            if is_mapped:
                self.log_signal.emit(f"Drive {drive_letter} is already mapped to {unc_path}. Skipping.")
                continue

            # Prepare net use command
            if use_credentials:
                # Note: Storing passwords in plain text is insecure.
                command = f'net use {drive_letter} "{unc_path}" "{password}" /user:{username} /persistent:no'
            else:
                command = f'net use {drive_letter} "{unc_path}" /persistent:no'

            stdout, stderr = execute_cmd(command)
            if stderr:
                # Attempt without trailing backslash
                if unc_path.endswith("\\"):
                    unc_path_retry = unc_path.rstrip("\\")
                    if use_credentials:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" "{password}" /user:{username} /persistent:no'
                    else:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" /persistent:no'
                    stdout_retry, stderr_retry = execute_cmd(command_retry)
                    if stderr_retry:
                        error_message = f"Error mapping drive {drive_letter}: {stderr_retry}"
                        self.log_signal.emit(error_message)
                        self.error_signal.emit(error_message)
                        continue
                    else:
                        success_message = f"Successfully mapped drive {drive_letter} to {unc_path_retry}."
                        self.log_signal.emit(success_message)
                else:
                    error_message = f"Error mapping drive {drive_letter}: {stderr}"
                    self.log_signal.emit(error_message)
                    self.error_signal.emit(error_message)
            else:
                success_message = f"Successfully mapped drive {drive_letter} to {unc_path}."
                self.log_signal.emit(success_message)
        self.log_signal.emit("Drive mapping process completed.")
        self.finished_signal.emit()

    def is_drive_mapped(self, drive_letter, unc_path):
        """
        Checks if a specific drive letter is mapped to the given UNC path.
        """
        current_drives = get_current_mapped_drives()
        for drive in current_drives:
            if drive["Drive"].upper() == drive_letter.upper() and drive["UNCPath"].lower() == unc_path.lower():
                return True
        return False

# Thread for Unmapping Drives
class UnmapDrivesThread(QThread):
    log_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, drive_mappings):
        super().__init__()
        self.drive_mappings = drive_mappings

    def run(self):
        self.log_signal.emit("Starting to unmap network drives...")
        for mapping in self.drive_mappings:
            if not mapping.get("Selected", False):
                continue  # Skip if not selected for unmapping
            drive_letter = mapping["Drive"]
            command = f'net use {drive_letter} /delete /y'
            stdout, stderr = execute_cmd(command)
            if stderr:
                error_message = f"Error unmapping drive {drive_letter}: {stderr}"
                self.log_signal.emit(error_message)
                self.error_signal.emit(error_message)
            else:
                success_message = f"Successfully unmapped drive {drive_letter}."
                self.log_signal.emit(success_message)
        self.log_signal.emit("Drive unmapping process completed.")
        self.finished_signal.emit()

# Thread for Checking Drives
class CheckDrivesThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, drive_mappings):
        super().__init__()
        self.drive_mappings = drive_mappings

    def run(self):
        self.log_signal.emit("Starting to check network drives...")
        for mapping in self.drive_mappings:
            drive_letter = mapping["Drive"]
            unc_path = mapping["UNCPath"]
            is_mapped = self.is_drive_mapped(drive_letter, unc_path)
            mapping["Mapped"] = "Yes" if is_mapped else "No"
            self.log_signal.emit(f"Drive {drive_letter} -> {unc_path} is {'mapped' if is_mapped else 'not mapped'}.")
        self.log_signal.emit("Drive checking process completed.")
        self.finished_signal.emit()

    def is_drive_mapped(self, drive_letter, unc_path):
        """
        Checks if a specific drive letter is mapped to the given UNC path.
        """
        current_drives = get_current_mapped_drives()
        for drive in current_drives:
            if drive["Drive"].upper() == drive_letter.upper() and drive["UNCPath"].lower() == unc_path.lower():
                return True
        return False

# Thread for Removing and Adding Drives on Startup
class ReaddDrivesThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, drive_mappings):
        super().__init__()
        self.drive_mappings = drive_mappings

    def run(self):
        self.log_signal.emit("Starting to remove and add drives on startup...")
        # Remove all drives
        for mapping in self.drive_mappings:
            drive_letter = mapping["Drive"]
            command = f'net use {drive_letter} /delete /y'
            stdout, stderr = execute_cmd(command)
            if stderr:
                self.log_signal.emit(f"Error unmapping drive {drive_letter}: {stderr}")
            else:
                self.log_signal.emit(f"Successfully unmapped drive {drive_letter}.")

        # Add all drives
        for mapping in self.drive_mappings:
            drive_letter = mapping["Drive"]
            unc_path = mapping["UNCPath"]
            use_credentials = mapping.get("UseCredentials", False)
            username = mapping.get("Username", "")
            password = mapping.get("Password", "")
            if use_credentials:
                # Note: Storing passwords in plain text is insecure.
                command = f'net use {drive_letter} "{unc_path}" "{password}" /user:{username} /persistent:no'
            else:
                command = f'net use {drive_letter} "{unc_path}" /persistent:no'

            stdout, stderr = execute_cmd(command)
            if stderr:
                # Retry without trailing backslash
                if unc_path.endswith("\\"):
                    unc_path_retry = unc_path.rstrip("\\")
                    if use_credentials:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" "{password}" /user:{username} /persistent:no'
                    else:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" /persistent:no'
                    stdout_retry, stderr_retry = execute_cmd(command_retry)
                    if stderr_retry:
                        self.log_signal.emit(f"Error mapping drive {drive_letter}: {stderr_retry}")
                        continue
                    else:
                        self.log_signal.emit(f"Successfully mapped drive {drive_letter} to {unc_path_retry}.")
                else:
                    self.log_signal.emit(f"Error mapping drive {drive_letter}: {stderr}")
            else:
                self.log_signal.emit(f"Successfully mapped drive {drive_letter} to {unc_path}.")
        self.log_signal.emit("Remove and Add Drives on startup process completed.")
        self.finished_signal.emit()

# Dialog for Adding or Editing a Drive
class AddEditDriveDialog(QDialog):
    def __init__(self, existing_drive_letters, drive_info=None, parent=None):
        super(AddEditDriveDialog, self).__init__(parent)
        self.setWindowTitle("Add Drive" if drive_info is None else "Edit Drive")
        self.setFixedSize(450, 350)
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: white;
                border: 2px solid #3c3f41;
                border-radius: 10px;
            }
            QLabel {
                color: white;
                font-weight: bold;
            }
            QLineEdit {
                background-color: #1e1e1e;
                color: white;
                border: 1px solid #555555;
                border-radius: 5px;
                padding: 5px;
            }
            QComboBox {
                background-color: #1e1e1e;
                color: white;
                border: 1px solid #555555;
                border-radius: 5px;
                padding: 5px;
            }
            QCheckBox {
                color: white;
            }
            QPushButton {
                background-color: #3c3f41;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #4e5254;
            }
        """)

        self.existing_drive_letters = [normalize_drive_letter(dl) for dl in existing_drive_letters]
        if drive_info:
            self.original_drive_letter = drive_info["Drive"]
        else:
            self.original_drive_letter = None

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Drive Letter
        drive_layout = QHBoxLayout()
        drive_label = QLabel("Drive:")
        self.drive_combo = QComboBox()
        free_letters = get_free_drive_letters(self.existing_drive_letters)
        # If editing and drive letter is unchanged, include it
        if drive_info:
            if drive_info["Drive"] not in free_letters:
                free_letters.append(drive_info["Drive"])
        if not free_letters:
            QMessageBox.critical(self, "No Available Drive Letters", "No available drive letters to assign.")
            self.reject()
        else:
            self.drive_combo.addItems(free_letters)
            if drive_info:
                index = self.drive_combo.findText(drive_info["Drive"])
                if index != -1:
                    self.drive_combo.setCurrentIndex(index)
        drive_layout.addWidget(drive_label)
        drive_layout.addWidget(self.drive_combo)
        layout.addLayout(drive_layout)

        # UNC Path
        path_layout = QHBoxLayout()
        path_label = QLabel("UNC Path:")
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText(r"e.g., \\server\share")
        if drive_info:
            self.path_input.setText(drive_info.get("UNCPath", ""))
        path_layout.addWidget(path_label)
        path_layout.addWidget(self.path_input)
        layout.addLayout(path_layout)

        # Use Different Credentials
        credentials_layout = QHBoxLayout()
        self.credentials_checkbox = QCheckBox("Use Different Credentials")
        credentials_layout.addWidget(self.credentials_checkbox)
        credentials_layout.addStretch()
        layout.addLayout(credentials_layout)

        # Username
        username_layout = QHBoxLayout()
        self.username_label = QLabel("Username:")
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Enter username")
        username_layout.addWidget(self.username_label)
        username_layout.addWidget(self.username_input)
        layout.addLayout(username_layout)

        # Password
        password_layout = QHBoxLayout()
        self.password_label = QLabel("Password:")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Enter password")
        self.password_input.setEchoMode(QLineEdit.Password)
        password_layout.addWidget(self.password_label)
        password_layout.addWidget(self.password_input)
        layout.addLayout(password_layout)

        # Initially hide credential fields
        self.username_label.hide()
        self.username_input.hide()
        self.password_label.hide()
        self.password_input.hide()

        # Connect checkbox
        self.credentials_checkbox.stateChanged.connect(self.toggle_credentials_fields)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.save_button = QPushButton("Save")
        self.cancel_button = QPushButton("Cancel")
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        # Connect Buttons
        self.save_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        # If editing and credentials are used, populate them
        if drive_info:
            self.credentials_checkbox.setChecked(drive_info.get("UseCredentials", False))
            if drive_info.get("UseCredentials", False):
                self.username_label.show()
                self.username_input.show()
                self.password_label.show()
                self.password_input.show()
                self.username_input.setText(drive_info.get("Username", ""))
                self.password_input.setText(drive_info.get("Password", ""))

    def toggle_credentials_fields(self, state):
        """
        Shows or hides the username and password fields based on the checkbox state.
        """
        if state == Qt.Checked:
            self.username_label.show()
            self.username_input.show()
            self.password_label.show()
            self.password_input.show()
        else:
            self.username_label.hide()
            self.username_input.hide()
            self.password_label.hide()
            self.password_input.hide()

    def get_drive_entry(self):
        """
        Retrieves the drive entry details from the dialog inputs.
        """
        drive_letter = self.drive_combo.currentText().strip()
        unc_path = self.path_input.text().strip()
        use_credentials = self.credentials_checkbox.isChecked()
        username = self.username_input.text().strip() if use_credentials else ""
        password = self.password_input.text().strip() if use_credentials else ""
        return {
            "Drive": drive_letter,
            "UNCPath": unc_path,
            "UseCredentials": use_credentials,
            "Username": username,
            "Password": password
        }

    def accept(self):
        """
        Validates inputs before accepting the dialog.
        """
        entry = self.get_drive_entry()
        drive_letter = entry["Drive"]
        unc_path = entry["UNCPath"]
        use_credentials = entry["UseCredentials"]
        username = entry["Username"]
        password = entry["Password"]

        if not drive_letter:
            QMessageBox.warning(self, "Invalid Input", "Please select a drive letter.")
            return
        if not unc_path.startswith("\\\\") or len(unc_path.split("\\")) < 4:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid UNC path (e.g., \\server\share).")
            return
        if use_credentials and (not username or not password):
            QMessageBox.warning(self, "Invalid Input", "Please enter both username and password.")
            return

        super().accept()

# Dialog for Editing a Drive
class EditDriveDialog(AddEditDriveDialog):
    def __init__(self, existing_drive_letters, drive_info, parent=None):
        super(EditDriveDialog, self).__init__(existing_drive_letters, drive_info, parent)
        
# Main Application Window
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TSTP Drive Mapper")
        self.powershell_script_content = ""
        self.setFixedSize(1000, 700)  # Increased width to accommodate new columns
        if os.path.exists(APP_ICON):
            self.setWindowIcon(QIcon(APP_ICON))

        # Initialize attributes
        self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode = load_settings()

        # Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Buttons Layout
        button_layout = QHBoxLayout()
        self.add_drive_button = QPushButton("Add Drive")
        self.edit_drive_button = QPushButton("Edit Drive")
        self.check_drives_button = QPushButton("Check Drives")
        self.map_drives_button = QPushButton("Map Drives")
        self.unmap_drives_button = QPushButton("Unmap Drives")
        self.remove_drive_button = QPushButton("Remove Drive")
        button_layout.addWidget(self.add_drive_button)
        button_layout.addWidget(self.edit_drive_button)
        button_layout.addWidget(self.check_drives_button)
        button_layout.addWidget(self.map_drives_button)
        button_layout.addWidget(self.unmap_drives_button)
        button_layout.addWidget(self.remove_drive_button)
        main_layout.addLayout(button_layout)

        # Drives Table
        self.drives_table = QTableWidget()
        self.drives_table.setColumnCount(7)
        self.drives_table.setHorizontalHeaderLabels(["#", "Select", "Drive", "UNC Path", "Added Date", "Mapped", "Force Auth"])
        self.drives_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # # column
        self.drives_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Select column
        self.drives_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Drive column
        self.drives_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)        # UNC Path column
        self.drives_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Added Date column
        self.drives_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Mapped column
        self.drives_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Force Auth column
        self.drives_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.drives_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.drives_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.drives_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #555555;
                border-radius: 5px;
            }
        """)
        main_layout.addWidget(self.drives_table)

        # Checkboxes
        checkbox_layout = QHBoxLayout()
        self.startup_checkbox = QCheckBox("Start on Windows Startup")
        self.auto_readd_checkbox = QCheckBox("Re-Add On Startup")
        checkbox_layout.addWidget(self.startup_checkbox)
        checkbox_layout.addWidget(self.auto_readd_checkbox)
        checkbox_layout.addStretch()
        main_layout.addLayout(checkbox_layout)

        # Log Console
        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        self.log_console.setFixedHeight(200)
        self.log_console.setStyleSheet("""
            QTextEdit {
                border: 1px solid #555555;
                border-radius: 5px;
            }
        """)
        main_layout.addWidget(self.log_console)

        central_widget.setLayout(main_layout)

        # Menu Bar
        self.create_menu()

        # System Tray
        self.create_tray_icon()

        # Connect Buttons
        self.add_drive_button.clicked.connect(self.add_drive)
        self.edit_drive_button.clicked.connect(self.edit_drive)
        self.check_drives_button.clicked.connect(self.check_drives)
        self.map_drives_button.clicked.connect(self.map_drives)
        self.unmap_drives_button.clicked.connect(self.unmap_drives)
        self.remove_drive_button.clicked.connect(self.remove_drive)

        # Connect Checkboxes
        self.startup_checkbox.stateChanged.connect(self.toggle_startup)
        self.auto_readd_checkbox.stateChanged.connect(self.toggle_auto_readd)

        # Set Checkboxes based on settings
        self.startup_checkbox.setChecked(self.startup_enabled)
        self.auto_readd_checkbox.setChecked(self.auto_readd_enabled)

        # Apply Light Mode if enabled
        if self.light_mode:
            self.apply_light_mode()

        # On startup, detect existing mapped drives and update settings
        try:
            existing_drives = get_current_mapped_drives()
            for drive in existing_drives:
                if not any(d["Drive"].upper() == drive["Drive"].upper() for d in self.drive_mappings):
                    self.drive_mappings.append({
                        "Drive": drive["Drive"],
                        "UNCPath": drive["UNCPath"],
                        "AddedDate": QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss"),
                        "Mapped": "Yes",
                        "Selected": False,
                        "UseCredentials": False,
                        "Username": "",
                        "Password": ""
                    })
                    logger.info(f"Detected existing drive: {drive['Drive']} -> {drive['UNCPath']}")

            # Save updated settings only if there are changes
            if existing_drives:
                save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)

            self.populate_drives_table()
        except Exception as e:
            logger.error(f"Error during startup drive detection: {e}")
            QMessageBox.critical(self, "Startup Error", f"An error occurred during startup:\n{e}")

        # Initialize Log
        self.update_log("Application started.")

        # If 'Re-Add On Startup' is enabled, perform the action
        if self.auto_readd_checkbox.isChecked():
            self.readd_drives()

        #Check for light or dark mode and reapply
        if self.light_mode:
            self.apply_light_mode()
        else:
            self.apply_dark_mode()

    def create_menu(self):
        menubar = self.menuBar()

        # File Menu
        file_menu = menubar.addMenu("File")

        import_action = QAction("Import", self)
        import_action.triggered.connect(self.import_settings)
        file_menu.addAction(import_action)

        export_action = QAction("Export", self)
        export_action.triggered.connect(self.export_settings)
        file_menu.addAction(export_action)

        export_ps_action = QAction("Export Drives for PowerShell Script", self)
        export_ps_action.triggered.connect(self.export_powershell_script)
        file_menu.addAction(export_ps_action)

        file_menu.addSeparator()

        save_logs_action = QAction("Save Log", self)
        save_logs_action.triggered.connect(self.save_logs)
        file_menu.addAction(save_logs_action)

        clear_logs_action = QAction("Clear Log", self)
        clear_logs_action.triggered.connect(self.clear_logs)
        file_menu.addAction(clear_logs_action)

        file_menu.addSeparator()

        toggle_console_action = QAction("Toggle Console", self)
        toggle_console_action.triggered.connect(self.toggle_console)
        file_menu.addAction(toggle_console_action)

        file_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.exit_application)
        file_menu.addAction(exit_action)

        # Settings Menu
        settings_menu = menubar.addMenu("Settings")

        add_drive_action = QAction("Add Drive", self)
        add_drive_action.triggered.connect(self.add_drive)
        settings_menu.addAction(add_drive_action)

        startup_settings_action = QAction("Startup Settings", self)
        startup_settings_action.triggered.connect(self.open_startup_settings)
        settings_menu.addAction(startup_settings_action)

        light_mode_action = QAction("Light Mode", self, checkable=True)
        light_mode_action.setChecked(self.light_mode)
        light_mode_action.triggered.connect(self.toggle_light_mode)
        settings_menu.addAction(light_mode_action)

        # Help Menu
        help_menu = menubar.addMenu("Help")

        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_page)
        help_menu.addAction(about_action)

        tutorial_action = QAction("Tutorial", self)
        tutorial_action.triggered.connect(self.show_tutorial_page)
        help_menu.addAction(tutorial_action)

        donate_action = QAction("Donate", self)
        donate_action.triggered.connect(self.show_donate_page)
        help_menu.addAction(donate_action)

        website_action = QAction("Website", self)
        website_action.triggered.connect(self.open_website)
        help_menu.addAction(website_action)

    def open_website(self):
        QDesktopServices.openUrl(QUrl("https://tstp.xyz"))
        
    def show_about_page(self):
        """
        Displays the About dialog using native PyQt5 widgets instead of embedded HTML.
        The contact information is presented as clickable buttons.
        """
        # Create a dialog
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle("About TSTP Drive Mapper")
        about_dialog.setMinimumSize(600, 700)  # Adjust size as needed

        # Main layout for the dialog
        main_layout = QVBoxLayout()
        about_dialog.setLayout(main_layout)

        # Scroll Area to handle content overflow
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_layout.addWidget(scroll_area)

        # Container widget for scroll area
        container = QWidget()
        scroll_area.setWidget(container)

        # Layout for the container
        container_layout = QVBoxLayout()
        container.setLayout(container_layout)

        # --- Title Section ---
        title_label = QLabel("About TSTP Drive Mapper")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #000000;")
        container_layout.addWidget(title_label)

        # Separator Line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #3c3f41;")
        container_layout.addWidget(separator)

        # --- Introduction Section ---
        intro_frame = QFrame()
        intro_frame.setStyleSheet("""
            QFrame {
                background-color: #3c3f41;
                border: 1px solid #3c3f41;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        intro_layout = QVBoxLayout()
        intro_frame.setLayout(intro_layout)

        intro_label = QLabel(
            "<p><strong>Version:</strong> 1.0.0</p>"
            "<p><strong>Release Date:</strong> October 1, 2024</p>"
            "<p><strong>Developer:</strong> TSTP Solutions</p>"
            "<p><strong>License:</strong> MIT License</p>"
            "<p>"
            "TSTP Drive Mapper is a comprehensive tool designed to manage network drive mappings efficiently on Windows systems. "
            "Whether you're an IT professional managing multiple network shares or a user looking to streamline your drive assignments, "
            "TSTP Drive Mapper offers an intuitive interface and robust features to simplify your workflow."
            "</p>"
        )
        intro_label.setWordWrap(True)
        intro_label.setStyleSheet("color: #ffffff; font-size: 14px;")
        intro_layout.addWidget(intro_label)

        container_layout.addWidget(intro_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Key Features Section ---
        features_frame = QFrame()
        features_frame.setStyleSheet("""
            QFrame {
                background-color: #3c3f41;
                border: 1px solid #3c3f41;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        features_layout = QVBoxLayout()
        features_frame.setLayout(features_layout)

        features_title = QLabel("Key Features")
        features_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #ffffff;")
        features_layout.addWidget(features_title)

        features_list = QLabel(
            "<ul>"
            "<li><strong>Easy Drive Management:</strong> Add, edit, and remove network drives with a user-friendly interface.</li>"
            "<li><strong>Bulk Operations:</strong> Map or unmap multiple drives simultaneously to save time.</li>"
            "<li><strong>Startup Integration:</strong> Configure the application to run at Windows startup and automatically re-add drives.</li>"
            "<li><strong>Credential Management:</strong> Securely store and manage credentials for drives requiring authentication.</li>"
            "<li><strong>Export Capabilities:</strong> Export your drive mappings as PowerShell scripts for automation and backup.</li>"
            "<li><strong>Comprehensive Logging:</strong> Keep detailed logs of all operations for auditing and troubleshooting.</li>"
            "<li><strong>Themes:</strong> Choose between light and dark modes to match your desktop preferences.</li>"
            "<li><strong>System Tray Integration:</strong> Access quick controls and settings directly from the system tray.</li>"
            "</ul>"
        )
        features_list.setWordWrap(True)
        features_list.setStyleSheet("color: #ffffff; font-size: 14px;")
        features_layout.addWidget(features_list)

        container_layout.addWidget(features_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Contact Information Section ---
        contact_frame = QFrame()
        contact_frame.setStyleSheet("""
            QFrame {
                background-color: #3c3f41;
                border: 1px solid #3c3f41;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        contact_layout = QVBoxLayout()
        contact_frame.setLayout(contact_layout)

        contact_title = QLabel("Contact Information")
        contact_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #ffffff;")
        contact_layout.addWidget(contact_title)

        # Email Button
        email_button = QPushButton("support@tstp.xyz")
        email_button.setStyleSheet("""
            QPushButton {
                background-color: #1e90ff;
                color: white;
                font-size: 14px;
                border: none;
                text-align: left;
            }
            QPushButton:hover {
                text-decoration: underline;
            }
        """)
        email_button.setCursor(Qt.PointingHandCursor)
        email_button.clicked.connect(lambda: self.open_donation_link("mailto:support@tstp.xyz"))
        contact_layout.addWidget(email_button)

        # Website Button
        website_button = QPushButton("https://tstp.xyz")
        website_button.setStyleSheet("""
            QPushButton {
                background-color: #1e90ff;
                color: white;
                font-size: 14px;
                border: none;
                text-align: left;
            }
            QPushButton:hover {
                text-decoration: underline;
            }
        """)
        website_button.setCursor(Qt.PointingHandCursor)
        website_button.clicked.connect(lambda: self.open_donation_link("https://tstp.xyz"))
        contact_layout.addWidget(website_button)

        container_layout.addWidget(contact_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Acknowledgments Section ---
        acknowledgments_frame = QFrame()
        acknowledgments_frame.setStyleSheet("""
            QFrame {
                background-color: #3c3f41;
                border: 1px solid #3c3f41;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        acknowledgments_layout = QVBoxLayout()
        acknowledgments_frame.setLayout(acknowledgments_layout)

        acknowledgments_title = QLabel("Acknowledgments")
        acknowledgments_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #ffffff;")
        acknowledgments_layout.addWidget(acknowledgments_title)

        acknowledgments_content = QLabel(
            "TSTP Drive Mapper leverages the powerful capabilities of PyQt5 for its graphical user interface and the Windows 'net use' command "
            "for managing network drives. We thank the open-source community for their invaluable contributions."
        )
        acknowledgments_content.setWordWrap(True)
        acknowledgments_content.setStyleSheet("color: #ffffff; font-size: 14px;")
        acknowledgments_layout.addWidget(acknowledgments_content)

        container_layout.addWidget(acknowledgments_frame)

        # Spacer to push content upwards
        container_layout.addStretch()

        # --- Close Button at the Bottom ---
        close_button = QPushButton("Close")
        close_button.setFixedHeight(40)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #555555;
                color: white;
                font-size: 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #777777;
            }
        """)
        close_button.clicked.connect(about_dialog.close)

        # Add Close button outside the scroll area to ensure it's always visible
        main_layout.addWidget(close_button)

        # Show the dialog
        about_dialog.exec_()

    def open_donation_link(self, url):
        """
        Opens the specified URL in the default web browser.
        """
        try:
            QDesktopServices.openUrl(QUrl(url))
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Could not open the link. Please try again.\n\nError: {str(e)}"
            )

    def show_tutorial_page(self):
        """
        Displays the Tutorial dialog using native PyQt5 widgets instead of embedded HTML.
        """
        # Create a dialog
        tutorial_dialog = QDialog(self)
        tutorial_dialog.setWindowTitle("TSTP Drive Mapper Tutorial")
        tutorial_dialog.setMinimumSize(600, 800)  # Adjust size as needed

        # Main layout for the dialog
        main_layout = QVBoxLayout()
        tutorial_dialog.setLayout(main_layout)

        # Scroll Area to handle content overflow
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_layout.addWidget(scroll_area)

        # Container widget for scroll area
        container = QWidget()
        scroll_area.setWidget(container)

        # Layout for the container
        container_layout = QVBoxLayout()
        container.setLayout(container_layout)

        # --- Title Section ---
        title_label = QLabel("TSTP Drive Mapper Tutorial")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #3c3f41;")
        container_layout.addWidget(title_label)

        # Separator Line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #dddddd;")
        container_layout.addWidget(separator)

        # --- Introduction Section ---
        intro_frame = QFrame()
        intro_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        intro_layout = QVBoxLayout()
        intro_frame.setLayout(intro_layout)

        intro_label = QLabel(
            "Welcome to the TSTP Drive Mapper tutorial. This guide is designed to help you navigate through the application's features "
            "and maximize its potential for managing your network drives effectively."
        )
        intro_label.setWordWrap(True)
        intro_label.setStyleSheet("color: #2b2b2b; font-size: 14px;")
        intro_layout.addWidget(intro_label)

        container_layout.addWidget(intro_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Getting Started Section ---
        getting_started_frame = QFrame()
        getting_started_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        getting_started_layout = QVBoxLayout()
        getting_started_frame.setLayout(getting_started_layout)

        getting_started_title = QLabel("Getting Started")
        getting_started_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        getting_started_layout.addWidget(getting_started_title)

        getting_started_content = QLabel(
            "To begin using TSTP Drive Mapper, simply launch the application. Upon launching, the main window displays a list of existing drive mappings (if any), "
            "along with options to manage them.\n\n"
            "If you're launching the application for the first time, you'll see an empty list. Start by adding your first network drive to get started."
        )
        getting_started_content.setWordWrap(True)
        getting_started_content.setStyleSheet("color: #2b2b2b; font-size: 14px;")
        getting_started_layout.addWidget(getting_started_content)

        container_layout.addWidget(getting_started_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Features and Usage Section ---
        features_usage_frame = QFrame()
        features_usage_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        features_usage_layout = QVBoxLayout()
        features_usage_frame.setLayout(features_usage_layout)

        features_usage_title = QLabel("Features and Usage")
        features_usage_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        features_usage_layout.addWidget(features_usage_title)

        # List of Features with Details
        features = [
            {
                "title": "1. Adding a New Drive",
                "content": (
                    "• Click the <strong>\"Add Drive\"</strong> button.<br>"
                    "• In the dialog, select an available drive letter from the dropdown.<br>"
                    "• Enter the UNC path (e.g., <code>\\\\server\\share</code>).<br>"
                    "• If the network share requires authentication, check <strong>\"Use Different Credentials\"</strong> and provide the necessary "
                    "username and password.<br>"
                    "• Click <strong>\"Save\"</strong> to add the drive. You will be prompted to map the drive immediately."
                )
            },
            {
                "title": "2. Editing an Existing Drive",
                "content": (
                    "• Select the drive you wish to edit by checking the corresponding checkbox in the <strong>\"Select\"</strong> column.<br>"
                    "• Click the <strong>\"Edit Drive\"</strong> button.<br>"
                    "• Modify the drive letter, UNC path, or credentials as needed.<br>"
                    "• Click <strong>\"Save\"</strong> to apply the changes. If the drive was previously mapped, it will be unmapped and remapped with the new settings."
                )
            },
            {
                "title": "3. Removing a Drive",
                "content": (
                    "• Select the drive(s) you wish to remove by checking the corresponding checkbox.<br>"
                    "• Click the <strong>\"Remove Drive\"</strong> button.<br>"
                    "• Confirm the removal in the prompt. If the drive is currently mapped, it will be unmapped before removal."
                )
            },
            {
                "title": "4. Connecting and Reconnecting Drives",
                "content": (
                    "• To establish a connection to your network drives, click the <strong>\"Connect\"</strong> button.<br>"
                    "• If a connection already exists and needs to be refreshed, the button will display <strong>\"Reconnect\"</strong>.<br>"
                    "• The application will handle the authentication process, especially if <strong>\"Force Authorization\"</strong> is enabled.<br>"
                    "• Monitor the log console for real-time updates on the connection status."
                )
            },
            {
                "title": "5. Unmapping Drives",
                "content": (
                    "• To unmap drives, select the desired drive(s) by checking the checkbox or leave all unchecked to unmap all mapped drives.<br>"
                    "• Click the <strong>\"Unmap Drives\"</strong> button. The application will handle the unmapping process and log the operations."
                )
            },
            {
                "title": "6. Checking Drive Status",
                "content": (
                    "• Click the <strong>\"Check Drives\"</strong> button to verify the current status of all drive mappings. The <strong>\"Mapped\"</strong> "
                    "column will indicate whether each drive is currently connected."
                )
            },
            {
                "title": "7. Exporting Drive Mappings",
                "content": (
                    "• Navigate to <strong>File &gt; Export as PowerShell Script</strong>.<br>"
                    "• Choose a destination to save the PowerShell script, which can be used for automation or backup purposes."
                )
            },
            {
                "title": "8. Logging and Console",
                "content": (
                    "• The log console at the bottom of the main window displays real-time logs of all operations.<br>"
                    "• Use <strong>File &gt; Save Log</strong> to export logs in various formats (TXT, JSON, XML).<br>"
                    "• Use <strong>File &gt; Clear Log</strong> to clear the log history.<br>"
                    "• Use <strong>File &gt; Toggle Console</strong> to show or hide the log console."
                )
            },
            {
                "title": "9. Settings",
                "content": (
                    "• <strong>Start on Windows Startup:</strong> Enable the application to run automatically when Windows starts.<br>"
                    "• <strong>Re-Add On Startup:</strong> Automatically remove and re-add all drive mappings when the application starts.<br>"
                    "• <strong>Light Mode:</strong> Switch between light and dark themes to suit your preference.<br>"
                    "• <strong>Force Authorization:</strong> When enabled, the application will require re-authentication for drive connections, enhancing security."
                )
            },
            {
                "title": "10. System Tray Integration",
                "content": (
                    "• The application minimizes to the system tray, allowing you to access quick controls without occupying space on your taskbar.<br>"
                    "• Right-click the tray icon to access options like opening the main window, toggling startup settings, switching themes, and exiting the application."
                )
            },
        ]

        for feature in features:
            feature_details = QFrame()
            feature_details.setStyleSheet("""
                QFrame {
                    background-color: #ffffff;
                    border: 1px solid #dddddd;
                    border-radius: 5px;
                    padding: 10px;
                    margin-top: 10px;
                }
            """)
            feature_layout = QVBoxLayout()
            feature_details.setLayout(feature_layout)

            feature_title = QLabel(feature["title"])
            feature_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #3c3f41;")
            feature_layout.addWidget(feature_title)

            feature_content = QLabel(feature["content"])
            feature_content.setWordWrap(True)
            feature_content.setStyleSheet("font-size: 14px; color: #2b2b2b;")
            feature_layout.addWidget(feature_content)

            features_usage_layout.addWidget(feature_details)

        container_layout.addWidget(features_usage_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Advanced Features Section ---
        advanced_features_frame = QFrame()
        advanced_features_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        advanced_features_layout = QVBoxLayout()
        advanced_features_frame.setLayout(advanced_features_layout)

        advanced_features_title = QLabel("Advanced Features")
        advanced_features_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        advanced_features_layout.addWidget(advanced_features_title)

        advanced_features_content = QLabel(
            "<ul>"
            "<li><strong>Bulk Operations:</strong> Quickly map or unmap multiple drives with a single action.</li>"
            "<li><strong>Credential Management:</strong> Securely manage credentials for network shares requiring authentication.</li>"
            "<li><strong>Script Export:</strong> Automate drive mappings by exporting configurations as PowerShell scripts.</li>"
            "<li><strong>Comprehensive Logging:</strong> Maintain detailed logs for auditing and troubleshooting purposes.</li>"
            "</ul>"
        )
        advanced_features_content.setWordWrap(True)
        advanced_features_content.setStyleSheet("color: #2b2b2b; font-size: 14px;")
        advanced_features_layout.addWidget(advanced_features_content)

        container_layout.addWidget(advanced_features_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Troubleshooting Section ---
        troubleshooting_frame = QFrame()
        troubleshooting_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        troubleshooting_layout = QVBoxLayout()
        troubleshooting_frame.setLayout(troubleshooting_layout)

        troubleshooting_title = QLabel("Troubleshooting")
        troubleshooting_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        troubleshooting_layout.addWidget(troubleshooting_title)

        troubleshooting_content = QLabel(
            "<ul>"
            "<li><strong>Drive Mapping Errors:</strong> Ensure that the UNC paths are correct and that you have the necessary permissions to access the network shares.</li>"
            "<li><strong>Credential Issues:</strong> Verify that the provided username and password are correct. Remember that storing passwords in plain text is insecure.</li>"
            "<li><strong>Startup Integration Failures:</strong> Ensure that the application has the necessary permissions to modify registry settings for startup.</li>"
            "<li><strong>Log Analysis:</strong> Refer to the log console and log files for detailed error messages and operational logs.</li>"
            "</ul>"
        )
        troubleshooting_content.setWordWrap(True)
        troubleshooting_content.setStyleSheet("color: #2b2b2b; font-size: 14px;")
        troubleshooting_layout.addWidget(troubleshooting_content)

        container_layout.addWidget(troubleshooting_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Best Practices Section ---
        best_practices_frame = QFrame()
        best_practices_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        best_practices_layout = QVBoxLayout()
        best_practices_frame.setLayout(best_practices_layout)

        best_practices_title = QLabel("Best Practices")
        best_practices_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        best_practices_layout.addWidget(best_practices_title)

        best_practices_content = QLabel(
            "<ul>"
            "<li>Regularly back up your drive mappings using the export feature.</li>"
            "<li>Use descriptive drive letters to easily identify network shares.</li>"
            "<li>Limit the use of different credentials to necessary shares to enhance security.</li>"
            "<li>Monitor logs to proactively identify and resolve issues.</li>"
            "</ul>"
        )
        best_practices_content.setWordWrap(True)
        best_practices_content.setStyleSheet("color: #2b2b2b; font-size: 14px;")
        best_practices_layout.addWidget(best_practices_content)

        container_layout.addWidget(best_practices_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Support Section ---
        support_frame = QFrame()
        support_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        support_layout = QVBoxLayout()
        support_frame.setLayout(support_layout)

        support_title = QLabel("Support")
        support_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        support_layout.addWidget(support_title)

        support_content = QLabel(
            "For additional support, feature requests, or to report bugs, please contact our support team at:"
        )
        support_content.setWordWrap(True)
        support_content.setStyleSheet("color: #2b2b2b; font-size: 14px;")
        support_layout.addWidget(support_content)

        # Email Button
        email_button = QPushButton("support@tstp.xyz")
        email_button.setStyleSheet("""
            QPushButton {
                background-color: #1e90ff;
                color: white;
                font-size: 14px;
                border: none;
                text-align: left;
            }
            QPushButton:hover {
                text-decoration: underline;
            }
        """)
        email_button.setCursor(Qt.PointingHandCursor)
        email_button.clicked.connect(lambda: self.open_donation_link("mailto:support@tstp.xyz"))
        support_layout.addWidget(email_button)

        # Website Button
        website_button = QPushButton("https://tstp.xyz")
        website_button.setStyleSheet("""
            QPushButton {
                background-color: #1e90ff;
                color: white;
                font-size: 14px;
                border: none;
                text-align: left;
            }
            QPushButton:hover {
                text-decoration: underline;
            }
        """)
        website_button.setCursor(Qt.PointingHandCursor)
        website_button.clicked.connect(lambda: self.open_donation_link("https://tstp.xyz"))
        support_layout.addWidget(website_button)

        container_layout.addWidget(support_frame)

        # Spacer to push content upwards
        container_layout.addStretch()

        # --- Close Button at the Bottom ---
        close_button = QPushButton("Close")
        close_button.setFixedHeight(40)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #555555;
                color: white;
                font-size: 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #777777;
            }
        """)
        close_button.clicked.connect(tutorial_dialog.close)

        # Add Close button outside the scroll area to ensure it's always visible
        main_layout.addWidget(close_button)

        # Show the dialog
        tutorial_dialog.exec_()

    def open_donation_link(self, url):
        """
        Opens the specified URL in the default web browser.
        """
        try:
            QDesktopServices.openUrl(QUrl(url))
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Could not open the link. Please try again.\n\nError: {str(e)}"
            )

    def show_donate_page(self):
        """
        Displays the Donate dialog using native PyQt5 widgets instead of embedded HTML.
        The PayPal donation button is positioned at the bottom of the dialog.
        """
        # Create a dialog
        donate_dialog = QDialog(self)
        donate_dialog.setWindowTitle("Support TSTP Drive Mapper")
        donate_dialog.setMinimumSize(600, 700)  # Adjust size as needed

        # Main layout for the dialog
        main_layout = QVBoxLayout()
        donate_dialog.setLayout(main_layout)

        # Scroll Area to handle content overflow
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_layout.addWidget(scroll_area)

        # Container widget for scroll area
        container = QWidget()
        scroll_area.setWidget(container)

        # Layout for the container
        container_layout = QVBoxLayout()
        container.setLayout(container_layout)

        # --- Title Section ---
        title_label = QLabel("Support TSTP Drive Mapper")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #3c3f41;")
        container_layout.addWidget(title_label)

        # Separator Line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        container_layout.addWidget(separator)

        # --- Introduction Section ---
        intro_label = QLabel(
            "Thank you for using TSTP Drive Mapper! We strive to provide a reliable and feature-rich tool to help you manage your network drives "
            "efficiently. If you find our application helpful, please consider supporting us to continue development and maintenance."
        )
        intro_label.setWordWrap(True)
        intro_label.setStyleSheet("font-size: 14px; color: #2b2b2b;")
        container_layout.addWidget(intro_label)

        # Spacer
        container_layout.addSpacing(10)

        # --- Why Donate Section ---
        why_donate_frame = QFrame()
        why_donate_frame.setStyleSheet("""
            QFrame {
                background-color: #f9f9f9;
                border: 1px solid #dddddd;
                border-radius: 8px;
            }
        """)
        why_donate_layout = QVBoxLayout()
        why_donate_frame.setLayout(why_donate_layout)

        why_donate_title = QLabel("Why Donate?")
        why_donate_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        why_donate_layout.addWidget(why_donate_title)

        why_donate_content = QLabel(
            "<ul>"
            "<li><strong>Development:</strong> Your contributions help us add new features and improve existing functionalities.</li>"
            "<li><strong>Maintenance:</strong> Donations assist in keeping the application updated with the latest Windows updates and security patches.</li>"
            "<li><strong>Support:</strong> They enable us to offer better support services to our users.</li>"
            "<li><strong>Community:</strong> Supporting us helps maintain a free tool that benefits the entire community.</li>"
            "</ul>"
        )
        why_donate_content.setWordWrap(True)
        why_donate_content.setStyleSheet("font-size: 14px; color: #2b2b2b;")
        why_donate_layout.addWidget(why_donate_content)

        container_layout.addWidget(why_donate_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- How to Donate Section ---
        how_to_donate_frame = QFrame()
        how_to_donate_frame.setStyleSheet("""
            QFrame {
                background-color: #f9f9f9;
                border: 1px solid #dddddd;
                border-radius: 8px;
            }
        """)
        how_to_donate_layout = QVBoxLayout()
        how_to_donate_frame.setLayout(how_to_donate_layout)

        how_to_donate_title = QLabel("How to Donate")
        how_to_donate_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        how_to_donate_layout.addWidget(how_to_donate_title)

        how_to_donate_content = QLabel(
            "We offer multiple ways to support TSTP Drive Mapper:"
        )
        how_to_donate_content.setWordWrap(True)
        how_to_donate_content.setStyleSheet("font-size: 14px; color: #2b2b2b;")
        how_to_donate_layout.addWidget(how_to_donate_content)

        # Donation Button Layout
        donation_button_layout = QVBoxLayout()

        # PayPal Donation Button
        paypal_button = QPushButton("Donate via PayPal")
        paypal_button.setFixedHeight(50)
        paypal_button.setStyleSheet("""
            QPushButton {
                background-color: #0070ba;
                color: white;
                font-size: 16px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #005f8a;
            }
        """)
        paypal_button.clicked.connect(lambda: self.open_donation_link("https://www.paypal.com/donate/?hosted_button_id=RAAYNUTMHPQQN"))
        paypal_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        donation_button_layout.addWidget(paypal_button)

        # Add button layout to how_to_donate_frame
        how_to_donate_layout.addLayout(donation_button_layout)

        container_layout.addWidget(how_to_donate_frame)

        # Spacer
        container_layout.addSpacing(20)

        # --- Thank You Section ---
        thank_you_frame = QFrame()
        thank_you_frame.setStyleSheet("""
            QFrame {
                background-color: #f9f9f9;
                border: 1px solid #dddddd;
                border-radius: 8px;
            }
        """)
        thank_you_layout = QVBoxLayout()
        thank_you_frame.setLayout(thank_you_layout)

        thank_you_title = QLabel("Thank You!")
        thank_you_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #3c3f41;")
        thank_you_layout.addWidget(thank_you_title)

        thank_you_content = QLabel(
            "Your generosity fuels our passion to create and maintain tools that make your digital life easier. Thank you for your support!"
        )
        thank_you_content.setWordWrap(True)
        thank_you_content.setStyleSheet("font-size: 14px; color: #2b2b2b;")
        thank_you_layout.addWidget(thank_you_content)

        container_layout.addWidget(thank_you_frame)

        # Spacer to push content upwards and keep buttons at the bottom
        container_layout.addStretch()

        # Show the dialog
        donate_dialog.exec_()

    def open_donation_link(self, url):
        """
        Opens the specified URL in the default web browser.
        """
        try:
            QDesktopServices.openUrl(QUrl(url))
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Could not open the link. Please try again.\n\nError: {str(e)}"
            )

    def create_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        if os.path.exists(APP_ICON):
            self.tray_icon.setIcon(QIcon(APP_ICON))
        else:
            self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setToolTip("TSTP Drive Mapper")

        tray_menu = QMenu()

        open_action = QAction("Open", self)
        open_action.triggered.connect(self.show_window)
        tray_menu.addAction(open_action)

        minimize_action = QAction("Minimize to Tray", self)
        minimize_action.triggered.connect(self.hide_window)
        tray_menu.addAction(minimize_action)

        close_to_tray_action = QAction("Close to Tray", self)
        close_to_tray_action.triggered.connect(self.hide_window)
        tray_menu.addAction(close_to_tray_action)

        tray_menu.addSeparator()

        startup_action = QAction("Start on Windows Startup", self, checkable=True)
        startup_action.setChecked(self.startup_enabled)
        startup_action.triggered.connect(self.toggle_startup_from_tray)
        tray_menu.addAction(startup_action)

        readd_action = QAction("Re-Add On Startup", self, checkable=True)
        readd_action.setChecked(self.auto_readd_enabled)
        readd_action.triggered.connect(self.toggle_readd_from_tray)
        tray_menu.addAction(readd_action)

        light_mode_action = QAction("Light Mode", self, checkable=True)
        light_mode_action.setChecked(self.light_mode)
        light_mode_action.triggered.connect(self.toggle_light_mode_from_tray)
        tray_menu.addAction(light_mode_action)

        tray_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.exit_application)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_window()

    def toggle_auto_readd(self, state):
        """
        Handles the state change of the auto_readd_checkbox.
        """
        try:
            self.auto_readd_enabled = state == Qt.Checked
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
            self.update_log(f"'Re-Add On Startup' set to {self.auto_readd_enabled}.")
        except Exception as e:
            logger.error(f"Error toggling auto re-add: {e}")
            QMessageBox.critical(self, "Auto Re-Add Toggle Error", f"An error occurred while toggling auto re-add:\n{e}")
    
    def show_window(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def hide_window(self):
        self.hide()

    def open_startup_settings(self):
        """
        Opens a dialog or settings related to startup.
        For simplicity, toggles the checkboxes.
        """
        # This can be expanded to a dedicated settings dialog
        pass  # Already handled via checkboxes and tray

    def toggle_light_mode_from_tray(self, checked):
        """
        Toggles light mode from the tray menu.
        """
        try:
            if checked:
                self.apply_light_mode()
            else:
                self.apply_dark_mode()
            self.light_mode = checked
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
        except Exception as e:
            logger.error(f"Error toggling light mode from tray: {e}")
            QMessageBox.critical(self, "Light Mode Toggle Error", f"An error occurred while toggling light mode:\n{e}")

    def toggle_light_mode(self, checked):
        """
        Toggles light mode from the settings menu.
        """
        try:
            if checked:
                self.apply_light_mode()
            else:
                self.apply_dark_mode()
            self.light_mode = checked
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
        except Exception as e:
            logger.error(f"Error toggling light mode: {e}")
            QMessageBox.critical(self, "Light Mode Toggle Error", f"An error occurred while toggling light mode:\n{e}")

    def apply_light_mode(self):
        """
        Applies the light theme to the application.
        """
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
                color: #000000;
            }
            QPushButton, QToolButton {
                background-color: #e0e0e0;
                color: #000000;
                border: 1px solid #a0a0a0;
                padding: 8px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover, QToolButton:hover {
                background-color: #d0d0d0;
            }
            QTextEdit {
                background-color: #f0f0f0;
                color: #000000;
                font-family: Consolas;
                font-size: 12px;
                border: 1px solid #a0a0a0;
                border-radius: 5px;
                padding: 10px;
            }
            QTableWidget {
                background-color: #f0f0f0;
                color: #000000;
                selection-background-color: #c0c0c0;
                border: 1px solid #a0a0a0;
                border-radius: 5px;
            }
            QTableWidget::item {
                color: #000000;
            }
            QTableWidget::item:first-child {
                background-color: #f0f0f0;
            }
            QTableWidget::item:selected {
                background-color: #ffffff;
            }
            QHeaderView::section {
                background-color: #d0d0d0;
                color: #000000;
                padding: 4px;
                border: 1px solid #a0a0a0;
                font-size: 12px;
            }
            QMenuBar {
                background-color: #d0d0d0;
                color: #000000;
            }
            QMenuBar::item:selected {
                background-color: #c0c0c0;
            }
            QMenu {
                background-color: #d0d0d0;
                color: #000000;
            }
            QMenu::item:selected {
                background-color: #c0c0c0;
            }
            QCheckBox {
                color: #000000;
                font-size: 12px;
            }
            QLabel {
                color: #000000;
                font-size: 12px;
            }
            QMessageBox {
                background-color: #ffffff;
                color: #000000;
            }
            QToolButton {
                border: none;
            }
            QPushButton#ForceConnectButton {
                background-color: #e0e0e0;
                color: #000000;
                border: 1px solid #a0a0a0;
                padding: 5px;
                border-radius: 3px;
                font-size: 11px;
            }
            QPushButton#ForceConnectButton:hover {
                background-color: #d0d0d0;
            }
        """)

    def apply_dark_mode(self):
        """
        Applies the dark theme to the application.
        """
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: white;
            }
            QPushButton, QToolButton {
                background-color: #3c3f41;
                color: white;
                border: 1px solid #555555;
                padding: 8px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover, QToolButton:hover {
                background-color: #4e5254;
            }
            QTextEdit {
                background-color: #1e1e1e;
                color: white;
                font-family: Consolas;
                font-size: 12px;
                border: 1px solid #555555;
                border-radius: 5px;
                padding: 10px;
            }
            QTableWidget {
                background-color: #1e1e1e;
                color: white;
                selection-background-color: #555555;
                border: 1px solid #555555;
                border-radius: 5px;
            }
            QTableWidget::item {
                color: white;
            }
            QTableWidget::item:first-child {
                background-color: #1e1e1e;
            }
            QTableWidget::item:selected {
                background-color: #555555;
            }
            QScrollArea::background { 
                background-color: #2b2b2b;
            }
            QHeaderView::section {
                background-color: #3c3f41;
                color: white;
                padding: 4px;
                border: 1px solid #555555;
                font-size: 12px;
            }
            QMenuBar {
                background-color: #3c3f41;
                color: white;
            }
            QMenuBar::item:selected {
                background-color: #555555;
            }
            QMenu {
                background-color: #3c3f41;
                color: white;
            }
            QMenu::item:selected {
                background-color: #555555;
            }
            QCheckBox {
                color: white;
                font-size: 12px;
            }
            QLabel {
                color: white;
                font-size: 12px;
            }
            QMessageBox {
                background-color: #2b2b2b;
                color: white;
            }
            QToolButton {
                border: none;
            }
            QPushButton#ForceConnectButton {
                background-color: #3c3f41;
                color: white;
                border: 1px solid #555555;
                padding: 5px;
                border-radius: 3px;
                font-size: 11px;
            }
            QPushButton#ForceConnectButton:hover {
                background-color: #4e5254;
            }
            QScrollArea {
                background-color: #2b2b2b;
                color: white;
            }
            QDialog {
                background-color: #2b2b2b;
                color: white;
            }
        """)

    def populate_drives_table(self):
        """
        Populates the drives table with current drive mappings.
        """
        try:
            self.drives_table.setRowCount(0)
            for index, mapping in enumerate(self.drive_mappings, start=1):
                row_position = self.drives_table.rowCount()
                self.drives_table.insertRow(row_position)

                # Row Number
                row_num_item = QTableWidgetItem(str(index))
                row_num_item.setBackground(QBrush(QColor("#3c3f41")))
                row_num_item.setForeground(QBrush(QColor("white")))
                row_num_item.setFlags(Qt.ItemIsEnabled)
                self.drives_table.setItem(row_position, 0, row_num_item)

                # Checkbox
                checkbox = QCheckBox()
                checkbox.setStyleSheet("margin-left:50%; margin-right:50%;")
                checkbox_widget = QWidget()
                checkbox_layout = QHBoxLayout(checkbox_widget)
                checkbox_layout.addWidget(checkbox)
                checkbox_layout.setAlignment(Qt.AlignCenter)
                checkbox_layout.setContentsMargins(0, 0, 0, 0)
                self.drives_table.setCellWidget(row_position, 1, checkbox_widget)
                self.drive_mappings[row_position -1]["Selected"] = False
                checkbox.stateChanged.connect(lambda state, row=row_position-1: self.update_selection(state, row))

                # Drive
                drive_item = QTableWidgetItem(mapping.get("Drive", "N/A"))
                drive_item.setForeground(QBrush(QColor("white")))
                self.drives_table.setItem(row_position, 2, drive_item)

                # UNC Path
                unc_path_item = QTableWidgetItem(mapping.get("UNCPath", "N/A"))
                unc_path_item.setForeground(QBrush(QColor("white")))
                self.drives_table.setItem(row_position, 3, unc_path_item)

                # Added Date
                added_date_item = QTableWidgetItem(mapping.get("AddedDate", "N/A"))
                added_date_item.setForeground(QBrush(QColor("white")))
                self.drives_table.setItem(row_position, 4, added_date_item)

                # Mapped Status
                mapped_item = QTableWidgetItem(mapping.get("Mapped", "No"))
                mapped_item.setForeground(QBrush(QColor("white")))
                self.drives_table.setItem(row_position, 5, mapped_item)

                # Force Auth Button
                force_connect_button = QPushButton("Reconnect" if mapping.get("Mapped", "No") == "Yes" else "Connect")
                force_connect_button.setObjectName("ForceConnectButton")
                force_connect_button.clicked.connect(lambda checked, row=row_position-1: self.force_connect(row))
                self.drives_table.setCellWidget(row_position, 6, force_connect_button)

                # Apply Shadow Effect to Button
                shadow = QGraphicsDropShadowEffect()
                shadow.setBlurRadius(5)
                shadow.setXOffset(0)
                shadow.setYOffset(0)
                force_connect_button.setGraphicsEffect(shadow)

        except Exception as e:
            logger.error(f"Error populating drives table: {e}")
            QMessageBox.critical(self, "Error", f"Failed to populate drives table:\n{e}")

    def update_selection(self, state, row):
        """
        Updates the 'Selected' status of a drive mapping based on the checkbox state.
        """
        try:
            self.drive_mappings[row]["Selected"] = True if state == Qt.Checked else False
        except IndexError:
            logger.error(f"Invalid row index {row} during selection update.")
        except Exception as e:
            logger.error(f"Error updating selection for row {row}: {e}")

    def update_drives_table_ui(self):
        """
        Updates the drives table UI after background operations.
        """
        self.populate_drives_table()
        save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
        QMessageBox.information(self, "Operation Completed", "Drive mapping status has been updated.")

    def update_log(self, message):
        """
        Appends a message to the log console and logs it.
        """
        try:
            timestamp = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
            self.log_console.append(f"[{timestamp}] {message}")
            logger.info(message)
        except Exception as e:
            logger.error(f"Error updating log: {e}")

    def add_drive(self):
        """
        Opens the Add Drive dialog and handles adding a new drive mapping.
        """
        existing_letters = [m["Drive"] for m in self.drive_mappings]
        dialog = AddEditDriveDialog(existing_letters, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            entry = dialog.get_drive_entry()
            drive_letter = entry["Drive"]
            unc_path = entry["UNCPath"]
            added_date = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
            use_credentials = entry["UseCredentials"]
            username = entry["Username"]
            password = entry["Password"]
            is_mapped = "No"  # Default

            # Ask if user wants to map now or just add to list
            reply = QMessageBox.question(
                self,
                "Map Drive",
                f"Do you want to map drive {drive_letter} to {unc_path} now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                # Execute CMD command to map the drive
                try:
                    if use_credentials:
                        # Note: Storing passwords in plain text is insecure.
                        command = f'net use {drive_letter} "{unc_path}" "{password}" /user:{username} /persistent:no'
                    else:
                        command = f'net use {drive_letter} "{unc_path}" /persistent:no'

                    stdout, stderr = execute_cmd(command)
                    if stderr:
                        # Attempt without trailing backslash
                        if unc_path.endswith("\\"):
                            unc_path_retry = unc_path.rstrip("\\")
                            if use_credentials:
                                command_retry = f'net use {drive_letter} "{unc_path_retry}" "{password}" /user:{username} /persistent:no'
                            else:
                                command_retry = f'net use {drive_letter} "{unc_path_retry}" /persistent:no'
                            stdout_retry, stderr_retry = execute_cmd(command_retry)
                            if stderr_retry:
                                error_message = f"Error mapping drive {drive_letter}: {stderr_retry}"
                                self.update_log(error_message)
                                QMessageBox.critical(self, "Mapping Error", error_message)
                                is_mapped = "No"
                            else:
                                success_message = f"Successfully mapped drive {drive_letter} to {unc_path_retry}."
                                self.update_log(success_message)
                                QMessageBox.information(self, "Drive Mapped", success_message)
                                is_mapped = "Yes"
                        else:
                            error_message = f"Error mapping drive {drive_letter}: {stderr}"
                            self.update_log(error_message)
                            QMessageBox.critical(self, "Mapping Error", error_message)
                            is_mapped = "No"
                    else:
                        success_message = f"Successfully mapped drive {drive_letter} to {unc_path}."
                        self.update_log(success_message)
                        QMessageBox.information(self, "Drive Mapped", success_message)
                        is_mapped = "Yes"
                except Exception as e:
                    error_message = f"Exception during mapping: {e}"
                    self.update_log(error_message)
                    QMessageBox.critical(self, "Mapping Error", error_message)
                    is_mapped = "No"
            else:
                self.update_log(f"Drive {drive_letter} added to the list without mapping.")

            # Add to drive mappings
            self.drive_mappings.append({
                "Drive": drive_letter,
                "UNCPath": unc_path,
                "AddedDate": added_date,
                "Mapped": is_mapped,
                "Selected": False,
                "UseCredentials": use_credentials,
                "Username": username,
                "Password": password
            })

            # Update the table
            self.populate_drives_table()
            self.check_drives()

            # Save settings
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)

    def edit_drive(self):
        """
        Opens the Edit Drive dialog for selected drive mappings.
        Corrects the row indexing and handles multiple selections.
        """
        try:
            selected_rows = [i for i in range(self.drives_table.rowCount()) if self.drive_mappings[i].get("Selected", False)]
            if not selected_rows:
                QMessageBox.information(self, "Edit Drive", "Please select at least one drive to edit.")
                return

            # Correct the index mapping if the table is in reverse order
            drive_indices = [self.drives_table.rowCount() - 1 - row for row in selected_rows]

            # Gather selected drive mappings
            drives_to_edit = [self.drive_mappings[i] for i in drive_indices]

            # Open the Multi Edit Dialog
            dialog = MultiEditDriveDialog(drives_to_edit, parent=self)
            if dialog.exec_() == QDialog.Accepted:
                edited_drives = dialog.get_drive_entries()

                for i, drive_entry in zip(drive_indices, edited_drives):
                    new_drive_letter = drive_entry["Drive"]
                    new_unc_path = drive_entry["UNCPath"]
                    new_use_credentials = drive_entry["UseCredentials"]
                    new_username = drive_entry["Username"]
                    new_password = drive_entry["Password"]

                    # If drive is mapped, unmap it first
                    if self.drive_mappings[i]["Mapped"] == "Yes":
                        command = f'net use {self.drive_mappings[i]["Drive"]} /delete /y'
                        stdout, stderr = execute_cmd(command)
                        if stderr:
                            error_message = f"Error unmapping drive {self.drive_mappings[i]['Drive']}: {stderr}"
                            self.update_log(error_message)
                            QMessageBox.critical(self, "Unmapping Error", error_message)
                            continue
                        else:
                            self.update_log(f"Successfully unmapped drive {self.drive_mappings[i]['Drive']} before editing.")

                    # Update the drive mapping
                    self.drive_mappings[i].update({
                        "Drive": new_drive_letter,
                        "UNCPath": new_unc_path,
                        "UseCredentials": new_use_credentials,
                        "Username": new_username,
                        "Password": new_password,
                        "Selected": False
                    })

                    # Attempt to map the drive with new settings
                    try:
                        if new_use_credentials:
                            # Note: Storing passwords in plain text is insecure.
                            command = f'net use {new_drive_letter} "{new_unc_path}" "{new_password}" /user:{new_username} /persistent:no'
                        else:
                            command = f'net use {new_drive_letter} "{new_unc_path}" /persistent:no'

                        stdout, stderr = execute_cmd(command)
                        if stderr:
                            # Retry without trailing backslash
                            if new_unc_path.endswith("\\"):
                                new_unc_path_retry = new_unc_path.rstrip("\\")
                                if new_use_credentials:
                                    command_retry = f'net use {new_drive_letter} "{new_unc_path_retry}" "{new_password}" /user:{new_username} /persistent:no'
                                else:
                                    command_retry = f'net use {new_drive_letter} "{new_unc_path_retry}" /persistent:no'
                                stdout_retry, stderr_retry = execute_cmd(command_retry)
                                if stderr_retry:
                                    error_message = f"Error mapping drive {new_drive_letter}: {stderr_retry}"
                                    self.update_log(error_message)
                                    QMessageBox.critical(self, "Mapping Error", error_message)
                                    self.drive_mappings[i]["Mapped"] = "No"
                                else:
                                    success_message = f"Successfully mapped drive {new_drive_letter} to {new_unc_path_retry}."
                                    self.update_log(success_message)
                                    QMessageBox.information(self, "Drive Mapped", success_message)
                                    self.drive_mappings[i]["Mapped"] = "Yes"
                            else:
                                error_message = f"Error mapping drive {new_drive_letter}: {stderr}"
                                self.update_log(error_message)
                                QMessageBox.critical(self, "Mapping Error", error_message)
                                self.drive_mappings[i]["Mapped"] = "No"
                        else:
                            success_message = f"Successfully mapped drive {new_drive_letter} to {new_unc_path}."
                            self.update_log(success_message)
                            QMessageBox.information(self, "Drive Mapped", success_message)
                            self.drive_mappings[i]["Mapped"] = "Yes"
                    except Exception as e:
                        error_message = f"Exception during mapping: {e}"
                        self.update_log(error_message)
                        QMessageBox.critical(self, "Mapping Error", error_message)
                        self.drive_mappings[i]["Mapped"] = "No"

                # Update the table
                self.populate_drives_table()
                self.check_drives()
                save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
        except Exception as e:
            logger.error(f"Error editing drive: {e}")
            QMessageBox.critical(self, "Edit Drive Error", f"An error occurred while editing the drive:\n{e}")

    def remove_drive(self):
        """
        Removes selected drive mappings from the list and unmaps them if necessary.
        """
        try:
            selected = [i for i in range(self.drives_table.rowCount()) if self.drive_mappings[i].get("Selected", False)]
            if selected:
                reply = QMessageBox.question(
                    self,
                    "Confirm Removal",
                    f"Are you sure you want to remove the selected {len(selected)} drive(s)?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    for index in sorted(selected, reverse=True):
                        drive_letter = self.drive_mappings[index]["Drive"]
                        is_mapped = self.drive_mappings[index]["Mapped"]
                        if is_mapped == "Yes":
                            # Unmap the drive
                            command = f'net use {drive_letter} /delete /y'
                            stdout, stderr = execute_cmd(command)
                            if stderr:
                                error_message = f"Error unmapping drive {drive_letter}: {stderr}"
                                self.update_log(error_message)
                                QMessageBox.critical(self, "Unmapping Error", error_message)
                            else:
                                self.update_log(f"Successfully unmapped drive {drive_letter}.")
                        # Remove from drive mappings
                        self.drive_mappings.pop(index)
                        self.drives_table.removeRow(index)
                    save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
                    self.update_log(f"Removed {len(selected)} drive(s) from the list.")
                    QMessageBox.information(self, "Drive Removed", f"Removed {len(selected)} drive(s) successfully.")
                else:
                    self.update_log("Drive removal canceled by user.")
                self.check_drives()
            else:
                QMessageBox.information(self, "No Selection", "Please select drive(s) to remove.")
        except Exception as e:
            logger.error(f"Error removing drive(s): {e}")
            QMessageBox.critical(self, "Remove Drive Error", f"An error occurred while removing drives:\n{e}")

    def check_drives(self):
        """
        Initiates a drive status check using a background thread.
        """
        try:
            self.update_log("Initiating drive status check...")
            self.check_thread = CheckDrivesThread(self.drive_mappings)
            self.check_thread.log_signal.connect(self.update_log)
            self.check_thread.finished_signal.connect(self.update_drives_table_ui)
            self.check_thread.start()
        except Exception as e:
            logger.error(f"Error initiating drive check: {e}")
            QMessageBox.critical(self, "Check Drives Error", f"An error occurred while checking drives:\n{e}")

    def map_drives(self):
        """
        Maps selected drives or all drives if none are selected.
        """
        try:
            # Determine selected drives
            selected_drives = [m for m in self.drive_mappings if m.get("Selected", False)]
            if selected_drives:
                reply = QMessageBox.question(
                    self,
                    "Map Drives",
                    f"Do you want to map the selected {len(selected_drives)} drive(s)?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    self.map_thread = MapDrivesThread(selected_drives, map_now=True)
                    self.map_thread.log_signal.connect(self.update_log)
                    self.map_thread.error_signal.connect(self.handle_mapping_error)
                    self.map_thread.finished_signal.connect(self.mapping_finished)
                    self.map_thread.start()
                    self.update_log("Started mapping selected drives in background.")
            else:
                # No specific selections, ask to map all
                reply = QMessageBox.question(
                    self,
                    "Map Drives",
                    "Do you want to map all drives in the list?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    self.map_thread = MapDrivesThread(self.drive_mappings, map_now=True)
                    self.map_thread.log_signal.connect(self.update_log)
                    self.map_thread.error_signal.connect(self.handle_mapping_error)
                    self.map_thread.finished_signal.connect(self.mapping_finished)
                    self.map_thread.start()
                    self.update_log("Started mapping all drives in background.")
        except Exception as e:
            logger.error(f"Error initiating drive mapping: {e}")
            QMessageBox.critical(self, "Map Drives Error", f"An error occurred while mapping drives:\n{e}")

    def unmap_drives(self):
        """
        Unmaps selected drives or all mapped drives if none are selected.
        """
        try:
            # Determine selected drives
            selected_drives = [m for m in self.drive_mappings if m.get("Selected", False)]
            if selected_drives:
                reply = QMessageBox.question(
                    self,
                    "Unmap Drives",
                    f"Do you want to unmap the selected {len(selected_drives)} drive(s)?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    for drive in selected_drives:
                        self.unmap_drive(drive)
                    self.update_log(f"Successfully unmapped {len(selected_drives)} selected drive(s).")
            else:
                # No specific selections, ask to unmap all
                mapped_drives = [m for m in self.drive_mappings if m.get("Mapped", "No") == "Yes"]
                if not mapped_drives:
                    QMessageBox.information(self, "Unmap Drives", "No mapped drives to unmap.")
                    return

                reply = QMessageBox.question(
                    self,
                    "Unmap Drives",
                    f"Do you want to unmap all {len(mapped_drives)} mapped drive(s)?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    for drive in mapped_drives:
                        self.unmap_drive(drive)
                    self.update_log(f"Successfully unmapped all {len(mapped_drives)} mapped drive(s).")

            # Refresh table after unmapping
            self.populate_drives_table()
            self.check_drives()

        except Exception as e:
            logger.error(f"Error initiating drive unmapping: {e}")
            QMessageBox.critical(self, "Unmap Drives Error", f"An error occurred while unmapping drives:\n{e}")

    def unmap_drive(self, drive):
        """
        Unmaps a single drive using the 'net use' command.
        """
        try:
            command = f'net use {drive["Drive"]} /delete /y'
            stdout, stderr = execute_cmd(command)

            if stderr:
                error_message = f"Error unmapping drive {drive['Drive']}: {stderr}"
                self.update_log(error_message)
                QMessageBox.critical(self, "Unmapping Error", error_message)
            else:
                success_message = f"Successfully unmapped drive {drive['Drive']}."
                self.update_log(success_message)
                drive["Mapped"] = "No"
        except Exception as e:
            error_message = f"Exception while unmapping drive {drive['Drive']}: {e}"
            self.update_log(error_message)
            QMessageBox.critical(self, "Unmapping Error", error_message)

    def handle_mapping_error(self, error_message):
        """
        Handles mapping errors by logging and displaying a message box.
        """
        self.update_log(error_message)
        QMessageBox.critical(self, "Mapping Error", error_message)

    def mapping_finished(self):
        """
        Called when mapping thread finishes. Initiates a drive check.
        """
        self.update_log("Drive mapping thread has finished.")
        self.check_drives()

    def unmapping_finished(self):
        """
        Called when unmapping thread finishes. Initiates a drive check.
        """
        self.update_log("Drive unmapping thread has finished.")
        self.check_drives()

    def import_settings(self):
        """
        Imports drive mappings from a JSON file.
        """
        try:
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Import Settings",
                "",
                "JSON Files (*.json);;All Files (*)",
                options=options
            )
            if file_path:
                with open(file_path, "r") as f:
                    settings = json.load(f)
                    imported_mappings = settings.get("drive_mappings", [])
                    for imported in imported_mappings:
                        # Handle both 'Drive' and 'DriveLetter'
                        drive_letter = normalize_drive_letter(imported.get("Drive", imported.get("DriveLetter", "")))
                        unc_path = imported.get("UNCPath", "")
                        added_date = imported.get("AddedDate", QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss"))
                        use_credentials = imported.get("UseCredentials", False)
                        username = imported.get("Username", "")
                        password = imported.get("Password", "")
                        # Check for duplicate drive letters
                        existing = next((d for d in self.drive_mappings if d["Drive"].upper() == drive_letter.upper()), None)
                        if existing:
                            # Ask user which one to keep
                            choice = QMessageBox.question(
                                self,
                                "Duplicate Drive Letter",
                                f"Drive letter {drive_letter} is already used for {existing['UNCPath']}. Do you want to replace it with {unc_path}?",
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No,
                            )
                            if choice == QMessageBox.Yes:
                                # Remove existing and add new
                                index = self.drive_mappings.index(existing)
                                self.drive_mappings.pop(index)
                                self.drives_table.removeRow(index)
                                self.drive_mappings.append({
                                    "Drive": drive_letter,
                                    "UNCPath": unc_path,
                                    "AddedDate": added_date,
                                    "Mapped": "No",
                                    "Selected": False,
                                    "UseCredentials": use_credentials,
                                    "Username": username,
                                    "Password": password
                                })
                                self.drives_table.insertRow(self.drives_table.rowCount())
                                row_position = self.drives_table.rowCount() -1

                                # Row Number
                                row_num_item = QTableWidgetItem(str(row_position +1))
                                row_num_item.setBackground(QBrush(QColor("#3c3f41")))
                                row_num_item.setForeground(QBrush(QColor("white")))
                                row_num_item.setFlags(Qt.ItemIsEnabled)
                                self.drives_table.setItem(row_position, 0, row_num_item)

                                # Checkbox
                                checkbox = QCheckBox()
                                checkbox.setStyleSheet("margin-left:50%; margin-right:50%;")
                                checkbox_widget = QWidget()
                                checkbox_layout = QHBoxLayout(checkbox_widget)
                                checkbox_layout.addWidget(checkbox)
                                checkbox_layout.setAlignment(Qt.AlignCenter)
                                checkbox_layout.setContentsMargins(0, 0, 0, 0)
                                self.drives_table.setCellWidget(row_position, 1, checkbox_widget)
                                self.drive_mappings[row_position]["Selected"] = False
                                checkbox.stateChanged.connect(lambda state, row=row_position: self.update_selection(state, row))

                                # Drive
                                drive_item = QTableWidgetItem(drive_letter)
                                drive_item.setForeground(QBrush(QColor("white")))
                                self.drives_table.setItem(row_position, 2, drive_item)

                                # UNC Path
                                unc_path_item = QTableWidgetItem(unc_path)
                                unc_path_item.setForeground(QBrush(QColor("white")))
                                self.drives_table.setItem(row_position, 3, unc_path_item)

                                # Added Date
                                added_date_item = QTableWidgetItem(added_date)
                                added_date_item.setForeground(QBrush(QColor("white")))
                                self.drives_table.setItem(row_position, 4, added_date_item)

                                # Mapped Status
                                mapped_item = QTableWidgetItem("No")
                                mapped_item.setForeground(QBrush(QColor("white")))
                                self.drives_table.setItem(row_position, 5, mapped_item)

                                # Force Auth Button
                                force_connect_button = QPushButton("Reconnect")
                                force_connect_button.setObjectName("ForceConnectButton")
                                force_connect_button.clicked.connect(lambda checked, row=row_position: self.force_connect(row))
                                self.drives_table.setCellWidget(row_position, 6, force_connect_button)

                                # Apply Shadow Effect to Button
                                shadow = QGraphicsDropShadowEffect()
                                shadow.setBlurRadius(5)
                                shadow.setXOffset(0)
                                shadow.setYOffset(0)
                                force_connect_button.setGraphicsEffect(shadow)

                                logger.info(f"Replaced drive {drive_letter} with {unc_path} from import.")
                                self.update_log(f"Replaced drive {drive_letter} with {unc_path} from import.")
                            else:
                                logger.info(f"Skipped importing drive {drive_letter} -> {unc_path}.")
                                self.update_log(f"Skipped importing drive {drive_letter} -> {unc_path}.")
                        else:
                            # Add new drive
                            self.drive_mappings.append({
                                "Drive": drive_letter,
                                "UNCPath": unc_path,
                                "AddedDate": added_date,
                                "Mapped": "No",
                                "Selected": False,
                                "UseCredentials": use_credentials,
                                "Username": username,
                                "Password": password
                            })
                            self.drives_table.insertRow(self.drives_table.rowCount())
                            row_position = self.drives_table.rowCount() -1

                            # Row Number
                            row_num_item = QTableWidgetItem(str(row_position +1))
                            row_num_item.setBackground(QBrush(QColor("#3c3f41")))
                            row_num_item.setForeground(QBrush(QColor("white")))
                            row_num_item.setFlags(Qt.ItemIsEnabled)
                            self.drives_table.setItem(row_position, 0, row_num_item)

                            # Checkbox
                            checkbox = QCheckBox()
                            checkbox.setStyleSheet("margin-left:50%; margin-right:50%;")
                            checkbox_widget = QWidget()
                            checkbox_layout = QHBoxLayout(checkbox_widget)
                            checkbox_layout.addWidget(checkbox)
                            checkbox_layout.setAlignment(Qt.AlignCenter)
                            checkbox_layout.setContentsMargins(0, 0, 0, 0)
                            self.drives_table.setCellWidget(row_position, 1, checkbox_widget)
                            self.drive_mappings[row_position]["Selected"] = False
                            checkbox.stateChanged.connect(lambda state, row=row_position: self.update_selection(state, row))

                            # Drive
                            drive_item = QTableWidgetItem(drive_letter)
                            drive_item.setForeground(QBrush(QColor("white")))
                            self.drives_table.setItem(row_position, 2, drive_item)

                            # UNC Path
                            unc_path_item = QTableWidgetItem(unc_path)
                            unc_path_item.setForeground(QBrush(QColor("white")))
                            self.drives_table.setItem(row_position, 3, unc_path_item)

                            # Added Date
                            added_date_item = QTableWidgetItem(added_date)
                            added_date_item.setForeground(QBrush(QColor("white")))
                            self.drives_table.setItem(row_position, 4, added_date_item)

                            # Mapped Status
                            mapped_item = QTableWidgetItem("No")
                            mapped_item.setForeground(QBrush(QColor("white")))
                            self.drives_table.setItem(row_position, 5, mapped_item)

                            # Force Auth Button
                            force_connect_button = QPushButton("Reconnect" if self.drive_mappings[row_position].get("Mapped", "No") == "Yes" else "Connect")
                            force_connect_button.setObjectName("ForceConnectButton")
                            force_connect_button.clicked.connect(lambda checked, row=row_position: self.force_connect(row))
                            self.drives_table.setCellWidget(row_position, 6, force_connect_button)

                            # Apply Shadow Effect to Button
                            shadow = QGraphicsDropShadowEffect()
                            shadow.setBlurRadius(5)
                            shadow.setXOffset(0)
                            shadow.setYOffset(0)
                            force_connect_button.setGraphicsEffect(shadow)

                            logger.info(f"Imported drive mapping: {drive_letter} -> {unc_path}.")
                            self.update_log(f"Imported drive mapping: {drive_letter} -> {unc_path}.")

                save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
                self.update_log(f"Imported drive mappings from {file_path}.")
                QMessageBox.information(self, "Import Successful", f"Settings imported from {file_path}.")
        except Exception as e:
            logger.error(f"Error importing settings: {e}")
            QMessageBox.critical(self, "Import Failed", f"Failed to import settings:\n{e}")

    def export_settings(self):
        """
        Exports current drive mappings to a JSON file.
        """
        try:
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Settings",
                "",
                "JSON Files (*.json);;All Files (*)",
                options=options
            )
            if file_path:
                with open(file_path, "w") as f:
                    json.dump({"drive_mappings": self.drive_mappings}, f, indent=4)
                    self.update_log(f"Exported settings to {file_path}.")
                    QMessageBox.information(self, "Export Successful", f"Settings exported to {file_path}.")
        except Exception as e:
            logger.error(f"Error exporting settings: {e}")
            QMessageBox.critical(self, "Export Failed", f"Failed to export settings:\n{e}")

    def export_powershell_script(self):
        """
        Exports the current drive mappings to a JSON file in the "Exports" folder of the Drive_Mapper directory
        and downloads or creates the PowerShell script in the Tools directory based on user preference.
        Prompts the user to open the Tools folder after export.
        """
        # Define the base directory and Exports/Tools folders
        base_dir = r"C:\TSTP\Drive_Mapper"
        exports_folder = os.path.join(base_dir, "Exports")
        tools_folder = os.path.join(base_dir, "Tools")

        # Ensure the Exports and Tools folders exist
        os.makedirs(exports_folder, exist_ok=True)
        os.makedirs(tools_folder, exist_ok=True)

        # Step 1: Save JSON Configuration
        default_config_path = os.path.join(exports_folder, "drive_mappings.json")
        config_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save PowerShell Script Configuration",
            default_config_path,
            "JSON Files (*.json);;All Files (*)"
        )

        if not config_path:
            # User canceled the save dialog
            return

        config = {
            "Drives": self.drive_mappings,
            "ReAddAtStart": self.auto_readd_enabled,
            "StartOnWindowsStart": self.startup_enabled
        }

        try:
            # Ensure the directory exists (in case user changes path)
            os.makedirs(os.path.dirname(config_path), exist_ok=True)

            # Write the JSON configuration
            with open(config_path, 'w', encoding='utf-8') as json_file:
                json.dump(config, json_file, indent=4)

            print(f"Configuration saved to {config_path}")
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error Saving Configuration",
                f"An error occurred while saving the configuration:\n{str(e)}"
            )
            return

        # Step 2: Prompt User to Download or Create PowerShell Script Locally
        user_choice = QMessageBox.question(
            self,
            "PowerShell Script Option",
            "Do you want to download the PowerShell script from the internet?\n"
            "Selecting 'No' will create the script locally from the embedded code.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        script_name = "drive_mapper_powershell.ps1"
        script_path = os.path.join(tools_folder, script_name)
        script_url = "https://www.tstp.xyz/downloads/tools/drive_mapper_powershell.ps1"

        if user_choice == QMessageBox.Yes:
            # User chose to download the script
            try:
                # Download the script from the specified URL
                with urllib.request.urlopen(script_url) as response, open(script_path, 'wb') as out_file:
                    data = response.read()
                    out_file.write(data)
                QMessageBox.information(
                    self,
                    "Download Successful",
                    f"PowerShell script downloaded successfully to:\n{script_path}"
                )
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Download Failed",
                    f"Failed to download the PowerShell script:\n{str(e)}\n\n"
                    "Attempting to create the script locally."
                )
                try:
                    # Create the script locally from embedded content
                    with open(script_path, 'w', encoding='utf-8') as ps_file:
                        ps_file.write(self.powershell_script_content)  # Ensure this attribute exists
                    QMessageBox.information(
                        self,
                        "Script Created",
                        f"PowerShell script created successfully at:\n{script_path}"
                    )
                except Exception as ex:
                    QMessageBox.critical(
                        self,
                        "Error Creating Script",
                        f"An error occurred while creating the PowerShell script locally:\n{str(ex)}"
                    )
                    return
        else:
            # User chose to create the script locally
            try:
                # Create the script locally from embedded content
                with open(script_path, 'w', encoding='utf-8') as ps_file:
                    ps_file.write(self.powershell_script_content)  # Ensure this attribute exists
                QMessageBox.information(
                    self,
                    "Script Created",
                    f"PowerShell script created successfully at:\n{script_path}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Error Creating Script",
                    f"An error occurred while creating the PowerShell script locally:\n{str(e)}"
                )
                return

        # Step 3: Prompt User to Open the Tools Folder
        open_folder = QMessageBox.question(
            self,
            "Open Tools Folder",
            "Do you want to open the Tools folder?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if open_folder == QMessageBox.Yes:
            try:
                os.startfile(tools_folder)
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Error Opening Folder",
                    f"Failed to open the Tools folder:\n{str(e)}"
                )

        # Define the PowerShell script content to create locally if download fails
        powershell_script_content = '''
        # PowerShell Script to Manage Network Drive Mappings
        # Enhanced Features:
        # - Import and export drive mappings from/to config.json
        # - Import drive mappings from an external JSON file
        # - Map, unmap, and force reconnect drives using CMD 'net use' commands
        # - Manage startup settings
        # - Implement threading to prevent GUI freezing
        # - Comprehensive error handling and logging
        # Generated on: 2024-11-14

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        # Global Variables
        $Global:ScriptPath = $MyInvocation.MyCommand.Path

        # Function to add or remove the script from Windows startup
        function Set-Startup {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [bool]$Enable
            )

            if (-not $Global:ScriptPath) {
                [System.Windows.Forms.MessageBox]::Show("Cannot determine the script path. Please run the script directly.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $shell = New-Object -ComObject WScript.Shell
            $shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\DriveMapper.lnk"

            if ($Enable) {
                if (-not (Test-Path $shortcutPath)) {
                    # Create a shortcut in the Startup folder
                    $shortcut = $shell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = "powershell.exe"
                    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Global:ScriptPath`""
                    $shortcut.WorkingDirectory = Split-Path $Global:ScriptPath
                    $shortcut.IconLocation = "shell32.dll,3"  # Standard PowerShell icon
                    $shortcut.Save()
                }
            } else {
                # Remove the shortcut if it exists
                if (Test-Path $shortcutPath) {
                    Remove-Item $shortcutPath -Force
                }
            }
        }

        # Function to export configuration to a JSON file
        function Export-Configuration {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string]$ConfigPath,

                [Parameter(Mandatory = $true)]
                [object]$ConfigData
            )
            $ConfigData | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Force
        }

        # Function to import configuration from a JSON file
        function Import-Configuration {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string]$ConfigPath
            )
            if (Test-Path $ConfigPath) {
                try {
                    return Get-Content -Path $ConfigPath | ConvertFrom-Json
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed to parse config.json. Ensure it's properly formatted.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return $null
                }
            } else {
                return @{
                    Drives = @()
                    ReAddAtStart = $false
                    StartOnWindowsStart = $false
                }
            }
        }

        # Function to import external JSON configuration
        function Import-External-Configuration {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string]$ExternalPath
            )
            if (Test-Path $ExternalPath) {
                try {
                    return Get-Content -Path $ExternalPath | ConvertFrom-Json
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed to import JSON file. Ensure it's properly formatted.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return $null
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Selected file does not exist.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return $null
            }
        }

        # Function to map a single drive using CMD 'net use' command
        function New-DriveMapping {  # Renamed from 'Map-Drive' to 'New-DriveMapping'
            param (
                [Parameter(Mandatory = $true)]
                [PSCustomObject]$DriveConfig
            )

            $driveLetter = $DriveConfig.Drive
            $uncPath = $DriveConfig.UNCPath
            $useCred = $DriveConfig.UseCredentials
            $username = $DriveConfig.Username
            $password = $DriveConfig.Password  # Renamed from 'pwd' to 'password'

            Write-Output "Processing drive ${driveLetter} -> $uncPath"

            # Prepare the 'net use' command
            if ($useCred -and $username -and $password) {
                $cmd = "net use ${driveLetter} `"$uncPath`" `"$password`" /user:`"$username`" /persistent:yes"
            } else {
                $cmd = "net use ${driveLetter} `"$uncPath`" /persistent:yes"
            }

            try {
                # Execute the 'net use' command
                $result = cmd.exe /c $cmd 2>&1
                if ($result -match "The command completed successfully") {
                    Write-Output "Successfully mapped drive ${driveLetter} to $uncPath."
                    return $true
                } else {
                    Write-Error "Failed to map drive ${driveLetter} to $uncPath. Result: $result"
                    return $false
                }
            } catch {
                Write-Error "Exception occurred while mapping drive ${driveLetter}: $_"
                return $false
            }
        }

        # Function to unmap a single drive using CMD 'net use' command
        function Remove-DriveMapping {  # Renamed from 'Unmap-Drive' to 'Remove-DriveMapping'
            param (
                [Parameter(Mandatory = $true)]
                [PSCustomObject]$DriveConfig
            )

            $driveLetter = $DriveConfig.Drive

            Write-Output "Unmapping drive ${driveLetter}"

            # Prepare the 'net use' command
            $cmd = "net use ${driveLetter} /delete /yes"

            try {
                # Execute the 'net use' command
                $result = cmd.exe /c $cmd 2>&1
                if ($result -match "The command completed successfully") {
                    Write-Output "Successfully unmapped drive ${driveLetter}."
                    return $true
                } else {
                    Write-Error "Failed to unmap drive ${driveLetter}. Result: $result"
                    return $false
                }
            } catch {
                Write-Error "Exception occurred while unmapping drive ${driveLetter}: $_"
                return $false
            }
        }

        # Function to perform drive mapping operations asynchronously using Runspaces
        function Start-DriveOperations {
            param (
                [Parameter(Mandatory = $true)]
                [string]$Operation  # "Map" or "Unmap"
            )

            # Create a RunspacePool
            $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, [Environment]::ProcessorCount)
            $runspacePool.Open()

            $jobs = @()

            foreach ($drive in $drives) {
                # Skip drives that are not selected
                if (-not $drive.Selected) {
                    continue
                }

                # Create a PowerShell instance
                $ps = [PowerShell]::Create()
                $ps.RunspacePool = $runspacePool

                if ($Operation -eq "Map") {
                    $ps.AddScript({
                        param($d)
                        New-DriveMapping -DriveConfig $d
                    }).AddArgument($drive)
                }
                elseif ($Operation -eq "Unmap") {
                    $ps.AddScript({
                        param($d)
                        Remove-DriveMapping -DriveConfig $d
                    }).AddArgument($drive)
                }
                else {
                    Write-Error "Invalid operation: $Operation"
                    continue
                }

                # Invoke asynchronously
                $job = $ps.BeginInvoke()
                $jobs += @{ PowerShell = $ps; AsyncResult = $job }
            }

            # Collect results
            foreach ($job in $jobs) {
                $job.PowerShell.EndInvoke($job.AsyncResult)
                $job.PowerShell.Dispose()
            }

            # Close the RunspacePool
            $runspacePool.Close()
            $runspacePool.Dispose()
        }

        # Function to create and display the Drive Mapping form
        function Show-DriveMappingForm {
            [CmdletBinding()]
            param ()

            # Define the path to config.json in the same directory as the script
            $scriptDirectory = (Get-Item -Path $PSCommandPath).DirectoryName
            if ($null -eq $scriptDirectory) {
                Write-Error "Failed to determine script directory."
                return
            }
            $configPath = Join-Path $scriptDirectory "config.json"
            if (-not (Test-Path $configPath)) {
                Write-Warning "Config file not found at $configPath. Creating a new one."
                $defaultConfig = @{
                    Drives = @()
                    ReAddAtStart = $false
                    StartOnWindowsStart = $false
                }
                $defaultConfig | ConvertTo-Json | Set-Content -Path $configPath
            }
            $config = Import-Configuration -ConfigPath $configPath

            # Initialize Form
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Drive Mapping Configuration"
            $form.Size = New-Object System.Drawing.Size(900, 800)
            $form.StartPosition = "CenterScreen"
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $form.MaximizeBox = $false

            # Instructions Label
            $labelInstructions = New-Object System.Windows.Forms.Label
            $labelInstructions.Text = "Configure your drive mappings below:"
            $labelInstructions.Location = New-Object System.Drawing.Point(10, 10)
            $labelInstructions.AutoSize = $true
            $form.Controls.Add($labelInstructions)

            # DataGridView for Drive Mappings
            $dataGridView = New-Object System.Windows.Forms.DataGridView
            $dataGridView.Location = New-Object System.Drawing.Point(10, 40)
            $dataGridView.Size = New-Object System.Drawing.Size(860, 400)
            $dataGridView.AutoGenerateColumns = $false
            $dataGridView.AllowUserToAddRows = $true
            $dataGridView.AllowUserToDeleteRows = $true
            $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
            $dataGridView.MultiSelect = $false
            $dataGridView.ReadOnly = $false
            $dataGridView.RowHeadersVisible = $false
            $form.Controls.Add($dataGridView)

            # Define Columns
            # Drive Column
            $colDrive = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colDrive.HeaderText = "Drive"
            $colDrive.Name = "Drive"
            $colDrive.Width = 60
            $dataGridView.Columns.Add($colDrive)

            # UNC Path Column
            $colUNCPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colUNCPath.HeaderText = "UNC Path"
            $colUNCPath.Name = "UNCPath"
            $colUNCPath.Width = 300
            $dataGridView.Columns.Add($colUNCPath)

            # Added Date Column
            $colAddedDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colAddedDate.HeaderText = "Added Date"
            $colAddedDate.Name = "AddedDate"
            $colAddedDate.Width = 150
            $dataGridView.Columns.Add($colAddedDate)

            # Mapped Column
            $colMapped = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colMapped.HeaderText = "Mapped"
            $colMapped.Name = "Mapped"
            $colMapped.Width = 60
            $colMapped.ReadOnly = $true
            $dataGridView.Columns.Add($colMapped)

            # Selected Column
            $colSelected = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
            $colSelected.HeaderText = "Selected"
            $colSelected.Name = "Selected"
            $colSelected.Width = 60
            $dataGridView.Columns.Add($colSelected)

            # Force Connect Button Column
            $colForceConnect = New-Object System.Windows.Forms.DataGridViewButtonColumn
            $colForceConnect.HeaderText = "Force Connect"
            $colForceConnect.Name = "ForceConnect"
            $colForceConnect.Text = "Force Connect"
            $colForceConnect.UseColumnTextForButtonValue = $true
            $colForceConnect.Width = 120
            $dataGridView.Columns.Add($colForceConnect)

            # Populate DataGridView with existing mappings
            if ($config.Drives -and $config.Drives.Count -gt 0) {
                foreach ($mapping in $config.Drives) {
                    $rowIndex = $dataGridView.Rows.Add()
                    $dataGridView.Rows[$rowIndex].Cells["Drive"].Value = $mapping.Drive
                    $dataGridView.Rows[$rowIndex].Cells["UNCPath"].Value = $mapping.UNCPath
                    $dataGridView.Rows[$rowIndex].Cells["AddedDate"].Value = $mapping.AddedDate
                    $dataGridView.Rows[$rowIndex].Cells["Mapped"].Value = $mapping.Mapped
                    $dataGridView.Rows[$rowIndex].Cells["Selected"].Value = $mapping.Selected
                }
            }

            # Event Handler for Force Connect Button Click
            $dataGridView.add_CellContentClick({
                param($eventSender, $e)
                if ($e.ColumnIndex -eq ($dataGridView.Columns["ForceConnect"].Index) -and $e.RowIndex -ge 0) {
                    $row = $dataGridView.Rows[$e.RowIndex]
                    $driveLetter = $row.Cells["Drive"].Value
                    $uncPath = $row.Cells["UNCPath"].Value
                    $useCredentials = $config.Drives[$e.RowIndex].UseCredentials
                    $username = $config.Drives[$e.RowIndex].Username
                    $password = $config.Drives[$e.RowIndex].Password
            
                    if ([string]::IsNullOrWhiteSpace($driveLetter) -or [string]::IsNullOrWhiteSpace($uncPath)) {
                        [System.Windows.Forms.MessageBox]::Show("Drive Letter and UNC Path cannot be empty.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }
            
                    # Prompt for credentials if required
                    if ($useCredentials -and (-not $username -or -not $password)) {
                        $credential = Get-Credential -Message "Enter credentials for mapping drive ${driveLetter}:"
                        $username = $credential.UserName
                        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
                        )
                    }
            
                    # Start mapping/unmapping in a separate job
                    $job = Start-Job -ScriptBlock {
                        param($d, $u, $useCreds, $user, $pass)
                        try {
                            # Check if the drive is already mapped
                            $checkResult = cmd.exe /c "net use ${d}" 2>&1
                            if ($checkResult -notmatch "The network connection could not be found") {
                                # Unmap the existing drive if it is already mapped
                                cmd.exe /c "net use ${d} /delete /yes" | Out-Null
                            }
            
                            # Prepare and execute the 'net use' command
                            if ($useCreds) {
                                $cmd = "net use ${d} `"$u`" `"$pass`" /user:`"$user`" /persistent:yes"
                            } else {
                                $cmd = "net use ${d} `"$u`" /persistent:yes"
                            }
            
                            # Execute the command
                            $result = cmd.exe /c $cmd 2>&1
                            if ($result -match "The command completed successfully") {
                                Write-Output "Success"
                            } else {
                                Write-Error "Failed: $result"
                            }
                        } catch {
                            Write-Error "Error while processing drive ${d}: $($_.Exception.Message)"
                        }
                    } -ArgumentList ($driveLetter, $uncPath, $useCredentials, $username, $password)
            
                    # Wait for job completion and update UI
                    Wait-Job -Job $job
                    $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
                    Remove-Job -Job $job
            
                    if ($result -eq "Success") {
                        # Update UI for success
                        $row.Cells["Mapped"].Value = "Yes"
                        $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        [System.Windows.Forms.MessageBox]::Show("Drive ${driveLetter} has been successfully connected.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        # Update UI for failure
                        $row.Cells["Mapped"].Value = "No"
                        [System.Windows.Forms.MessageBox]::Show("Failed to connect drive ${driveLetter}: $result", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
            })
            

            # Panel to hold action buttons
            $panelActions = New-Object System.Windows.Forms.Panel
            $panelActions.Size = New-Object System.Drawing.Size(860, 60)
            $panelActions.Location = New-Object System.Drawing.Point(10, 460)
            $panelActions.Anchor = "Top, Left, Right"
            $form.Controls.Add($panelActions)

            # Create Buttons
            $buttonMap = New-Object System.Windows.Forms.Button
            $buttonMap.Text = "Map Drives"
            $buttonMap.Size = New-Object System.Drawing.Size(120, 40)

            $buttonUnmap = New-Object System.Windows.Forms.Button
            $buttonUnmap.Text = "Unmap Drives"
            $buttonUnmap.Size = New-Object System.Drawing.Size(120, 40)

            $buttonImport = New-Object System.Windows.Forms.Button
            $buttonImport.Text = "Import JSON"
            $buttonImport.Size = New-Object System.Drawing.Size(120, 40)

            $buttonExport = New-Object System.Windows.Forms.Button
            $buttonExport.Text = "Export JSON"
            $buttonExport.Size = New-Object System.Drawing.Size(120, 40)

            $buttonClose = New-Object System.Windows.Forms.Button
            $buttonClose.Text = "Close"
            $buttonClose.Size = New-Object System.Drawing.Size(120, 40)

            # Add Buttons to Panel
            $panelActions.Controls.Add($buttonMap)
            $panelActions.Controls.Add($buttonUnmap)
            $panelActions.Controls.Add($buttonImport)
            $panelActions.Controls.Add($buttonExport)
            $panelActions.Controls.Add($buttonClose)

            # Arrange buttons in a single row with equal spacing using FlowLayoutPanel
            $flowLayout = New-Object System.Windows.Forms.FlowLayoutPanel
            $flowLayout.Dock = "Fill"
            $flowLayout.FlowDirection = "LeftToRight"
            $flowLayout.WrapContents = $false
            $flowLayout.AutoSize = $false
            $flowLayout.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
            $flowLayout.Controls.AddRange(@($buttonMap, $buttonUnmap, $buttonImport, $buttonExport, $buttonClose))
            $panelActions.Controls.Clear()
            $panelActions.Controls.Add($flowLayout)

            # Event Handlers for Map, Unmap, Import, Export Buttons
            $buttonMap.Add_Click({
                # Disable buttons to prevent multiple clicks
                $buttonMap.Enabled = $false
                $buttonUnmap.Enabled = $false
                $buttonImport.Enabled = $false
                $buttonExport.Enabled = $false

                # Gather selected drives
                $selectedDrives = @()
                foreach ($row in $dataGridView.Rows) {
                    if (-not $row.IsNewRow -and $row.Cells["Selected"].Value -eq $true) {
                        $index = $dataGridView.Rows.IndexOf($row)
                        $selectedDrives += @{
                            Drive = $row.Cells["Drive"].Value
                            UNCPath = $row.Cells["UNCPath"].Value
                            AddedDate = $row.Cells["AddedDate"].Value
                            Mapped = $row.Cells["Mapped"].Value
                            Selected = $row.Cells["Selected"].Value
                            UseCredentials = $config.Drives[$index].UseCredentials
                            Username = $config.Drives[$index].Username
                            Password = $config.Drives[$index].Password
                        }
                    }
                }

                if ($selectedDrives.Count -eq 0) {
                    $result = [System.Windows.Forms.MessageBox]::Show("No drives selected. Do you want to map all drives?", "No Selection", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        foreach ($row in $dataGridView.Rows) {
                            if (-not $row.IsNewRow) {
                                $index = $dataGridView.Rows.IndexOf($row)
                                $selectedDrives += @{
                                    Drive = $row.Cells["Drive"].Value
                                    UNCPath = $row.Cells["UNCPath"].Value
                                    AddedDate = $row.Cells["AddedDate"].Value
                                    Mapped = $row.Cells["Mapped"].Value
                                    Selected = $row.Cells["Selected"].Value
                                    UseCredentials = $config.Drives[$index].UseCredentials
                                    Username = $config.Drives[$index].Username
                                    Password = $config.Drives[$index].Password
                                }
                            }
                        }
                    } else {
                        # Re-enable buttons
                        $buttonMap.Enabled = $true
                        $buttonUnmap.Enabled = $true
                        $buttonImport.Enabled = $true
                        $buttonExport.Enabled = $true
                        return
                    }
                }

                # Start mapping in separate jobs to prevent GUI freezing
                foreach ($drive in $selectedDrives) {
                    Start-Job -ScriptBlock {
                        param($d, $u, $useCredentials, $username, $password)
                        if ($d -and $u) {
                            # Unmap existing drive if mapped
                            cmd.exe /c "net use $d /delete /yes" | Out-Null

                            # Prepare 'net use' command
                            if ($useCredentials) {
                                $cmd = "net use $d `"$u`" `"$password`" /user:`"$username`" /persistent:yes"
                            } else {
                                $cmd = "net use $d `"$u`" /persistent:yes"
                            }

                            # Execute the command
                            $result = cmd.exe /c $cmd 2>&1
                            if ($result -match "The command completed successfully") {
                                Write-Output "Success"
                            } else {
                                Write-Error "Failed: $result"
                            }
                        } else {
                            Write-Error "Drive Letter or UNC Path is missing."
                        }
                    } -ArgumentList ($drive.Drive, $drive.UNCPath, $drive.UseCredentials, $drive.Username, $drive.Password) | Out-Null
                }

                # Inform user that mapping has been initiated
                [System.Windows.Forms.MessageBox]::Show("Drive mapping operations have been initiated. Please wait for completion.", "Mapping Initiated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                # Monitor and handle job results
                $jobs = Get-Job | Where-Object { $_.Command -like "*net use $" }

                foreach ($job in $jobs) {
                    Wait-Job -Job $job
                    $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
                    Remove-Job -Job $job

                    if ($result -eq "Success") {
                        # Update UI with success status
                        foreach ($drive in $selectedDrives) {
                            for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                                $row = $dataGridView.Rows[$i]
                                if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                                    $row.Cells["Mapped"].Value = "Yes"
                                    $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                }
                            }
                        }
                        [System.Windows.Forms.MessageBox]::Show("Drive mappings have been successfully completed.", "Mapping Completed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        # Handle failures
                        foreach ($drive in $selectedDrives) {
                            for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                                $row = $dataGridView.Rows[$i]
                                if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                                    $row.Cells["Mapped"].Value = "No"
                                    [System.Windows.Forms.MessageBox]::Show("Failed to map drive ${drive.Drive}: $result", "Mapping Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                }
                            }
                        }
                    }
                }

                # Save current mappings to configuration
                $mappedDrives = @()
                foreach ($row in $dataGridView.Rows) {
                    if (-not $row.IsNewRow) {
                        $index = $dataGridView.Rows.IndexOf($row)
                        $mappedDrives += @{
                            Drive = $row.Cells["Drive"].Value
                            UNCPath = $row.Cells["UNCPath"].Value
                            AddedDate = $row.Cells["AddedDate"].Value
                            Mapped = $row.Cells["Mapped"].Value
                            Selected = $row.Cells["Selected"].Value
                            UseCredentials = $config.Drives[$index].UseCredentials
                            Username = $config.Drives[$index].Username
                            Password = $config.Drives[$index].Password
                        }
                    }
                }
                $newConfig = @{
                    Drives = $mappedDrives
                    ReAddAtStart = $checkboxReAddAtStart.Checked
                    StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
                }
                Export-Configuration -ConfigPath $configPath -ConfigData $newConfig

                # Set startup based on checkbox
                Set-Startup -Enable $checkboxStartOnWindowsStart.Checked

                # Re-enable buttons
                $buttonMap.Enabled = $true
                $buttonUnmap.Enabled = $true
                $buttonImport.Enabled = $true
                $buttonExport.Enabled = $true
            })

            $buttonUnmap.Add_Click({
                # Disable buttons to prevent multiple clicks
                $buttonMap.Enabled = $false
                $buttonUnmap.Enabled = $false
                $buttonImport.Enabled = $false
                $buttonExport.Enabled = $false

                # Gather selected drives
                $selectedDrives = @()
                foreach ($row in $dataGridView.Rows) {
                    if (-not $row.IsNewRow -and $row.Cells["Selected"].Value -eq $true) {
                        $index = $dataGridView.Rows.IndexOf($row)
                        $selectedDrives += @{
                            Drive = $row.Cells["Drive"].Value
                            UNCPath = $row.Cells["UNCPath"].Value
                            AddedDate = $row.Cells["AddedDate"].Value
                            Mapped = $row.Cells["Mapped"].Value
                            Selected = $row.Cells["Selected"].Value
                            UseCredentials = $config.Drives[$index].UseCredentials
                            Username = $config.Drives[$index].Username
                            Password = $config.Drives[$index].Password
                        }
                    }
                }

                if ($selectedDrives.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("Please select at least one drive to unmap.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    # Re-enable buttons
                    $buttonMap.Enabled = $true
                    $buttonUnmap.Enabled = $true
                    $buttonImport.Enabled = $true
                    $buttonExport.Enabled = $true
                    return
                }

                # Start unmapping in separate jobs to prevent GUI freezing
                foreach ($drive in $selectedDrives) {
                    Start-Job -ScriptBlock {
                        param($d)
                        $cmd = "net use $d /delete /yes"
                        $result = cmd.exe /c $cmd 2>&1
                        if ($result -match "The command completed successfully") {
                            Write-Output "Success"
                        } else {
                            Write-Error "Failed: $result"
                        }
                    } -ArgumentList ($drive.Drive) | Out-Null
                }

                # Inform user that unmapping has been initiated
                [System.Windows.Forms.MessageBox]::Show("Drive unmapping operations have been initiated. Please wait for completion.", "Unmapping Initiated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                # Monitor and handle job results
                $jobs = Get-Job | Where-Object { $_.Command -like "*net use $" }

                foreach ($job in $jobs) {
                    Wait-Job -Job $job
                    $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
                    Remove-Job -Job $job

                    if ($result -eq "Success") {
                        # Update UI with success status
                        foreach ($drive in $selectedDrives) {
                            for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                                $row = $dataGridView.Rows[$i]
                                if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                                    $row.Cells["Mapped"].Value = "No"
                                    $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                }
                            }
                        }
                        [System.Windows.Forms.MessageBox]::Show("Drive unmapping operations have been successfully completed.", "Unmapping Completed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        # Handle failures
                        foreach ($drive in $selectedDrives) {
                            for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                                $row = $dataGridView.Rows[$i]
                                if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                                    $row.Cells["Mapped"].Value = "No"
                                    [System.Windows.Forms.MessageBox]::Show("Failed to unmap drive ${drive.Drive}: $result", "Unmapping Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                }
                            }
                        }
                    }
                }

                # Save current mappings to configuration
                $mappedDrives = @()
                foreach ($row in $dataGridView.Rows) {
                    if (-not $row.IsNewRow) {
                        $index = $dataGridView.Rows.IndexOf($row)
                        $mappedDrives += @{
                            Drive = $row.Cells["Drive"].Value
                            UNCPath = $row.Cells["UNCPath"].Value
                            AddedDate = $row.Cells["AddedDate"].Value
                            Mapped = $row.Cells["Mapped"].Value
                            Selected = $row.Cells["Selected"].Value
                            UseCredentials = $config.Drives[$index].UseCredentials
                            Username = $config.Drives[$index].Username
                            Password = $config.Drives[$index].Password
                        }
                    }
                }
                $newConfig = @{
                    Drives = $mappedDrives
                    ReAddAtStart = $checkboxReAddAtStart.Checked
                    StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
                }
                Export-Configuration -ConfigPath $configPath -ConfigData $newConfig

                # Set startup based on checkbox
                Set-Startup -Enable $checkboxStartOnWindowsStart.Checked

                # Re-enable buttons
                $buttonMap.Enabled = $true
                $buttonUnmap.Enabled = $true
                $buttonImport.Enabled = $true
                $buttonExport.Enabled = $true
            })

            $buttonImport.Add_Click({
                $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $openFileDialog.InitialDirectory = $scriptDirectory
                $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
                $openFileDialog.Title = "Import Drive Mappings JSON"

                if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $externalPath = $openFileDialog.FileName
                    $externalConfig = Import-External-Configuration -ExternalPath $externalPath
                    if ($externalConfig -and $externalConfig.Drives) {
                        # Clear existing rows
                        $dataGridView.Rows.Clear()

                        # Populate DataGridView with imported mappings
                        foreach ($mapping in $externalConfig.Drives) {
                            $rowIndex = $dataGridView.Rows.Add()
                            $dataGridView.Rows[$rowIndex].Cells["Drive"].Value = $mapping.Drive
                            $dataGridView.Rows[$rowIndex].Cells["UNCPath"].Value = $mapping.UNCPath
                            $dataGridView.Rows[$rowIndex].Cells["AddedDate"].Value = $mapping.AddedDate
                            $dataGridView.Rows[$rowIndex].Cells["Mapped"].Value = $mapping.Mapped
                            $dataGridView.Rows[$rowIndex].Cells["Selected"].Value = $mapping.Selected
                        }

                        # Update other settings
                        $checkboxReAddAtStart.Checked = $externalConfig.ReAddAtStart
                        $checkboxStartOnWindowsStart.Checked = $externalConfig.StartOnWindowsStart

                        # Save the imported configuration
                        Export-Configuration -ConfigPath $configPath -ConfigData $externalConfig

                        # Update the global config variable
                        #$config = Import-Configuration -ConfigPath $configPath

                        # Set startup based on checkbox
                        Set-Startup -Enable $externalConfig.StartOnWindowsStart

                        [System.Windows.Forms.MessageBox]::Show("JSON configuration imported successfully.", "Import Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Failed to import configuration. The file may be invalid or missing required data.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
            })

            $buttonExport.Add_Click({
                $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveFileDialog.InitialDirectory = $scriptDirectory
                $saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
                $saveFileDialog.Title = "Export Drive Mappings to JSON"

                if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $exportPath = $saveFileDialog.FileName

                    # Gather all mappings
                    $mappedDrives = @()
                    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                        $currentRow = $dataGridView.Rows[$i]
                        if (-not $currentRow.IsNewRow) {
                            $mappedDrives += @{
                                Drive = $currentRow.Cells["Drive"].Value
                                UNCPath = $currentRow.Cells["UNCPath"].Value
                                AddedDate = $currentRow.Cells["AddedDate"].Value
                                Mapped = $currentRow.Cells["Mapped"].Value
                                Selected = $currentRow.Cells["Selected"].Value
                                UseCredentials = $config.Drives[$i].UseCredentials
                                Username = $config.Drives[$i].Username
                                Password = $config.Drives[$i].Password
                            }
                        }
                    }

                    $exportConfig = @{
                        Drives = $mappedDrives
                        ReAddAtStart = $checkboxReAddAtStart.Checked
                        StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
                    }

                    try {
                        Export-Configuration -ConfigPath $exportPath -ConfigData $exportConfig
                        [System.Windows.Forms.MessageBox]::Show("Configuration exported successfully to $exportPath.", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to export configuration. Error: $($_.Exception.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
            })

            $buttonClose.Add_Click({
                $form.Close()
            })

            # Panel for Checkboxes
            $panelCheckboxes = New-Object System.Windows.Forms.Panel
            $panelCheckboxes.Size = New-Object System.Drawing.Size(860, 50)
            $panelCheckboxes.Location = New-Object System.Drawing.Point(10, 530)
            $panelCheckboxes.Anchor = "Top, Left, Right"
            $form.Controls.Add($panelCheckboxes)

            # Re-add at Start Checkbox
            $checkboxReAddAtStart = New-Object System.Windows.Forms.CheckBox
            $checkboxReAddAtStart.Text = "Re-add Drives at Start"
            $checkboxReAddAtStart.AutoSize = $true
            $checkboxReAddAtStart.Checked = $config.ReAddAtStart
            $checkboxReAddAtStart.Location = New-Object System.Drawing.Point(0, 15)
            $checkboxReAddAtStart.Anchor = "Top"
            $checkboxReAddAtStart.Add_CheckedChanged({
                $config.ReAddAtStart = $checkboxReAddAtStart.Checked
                Export-Configuration -ConfigPath $configPath -ConfigData $config
            })
            $panelCheckboxes.Controls.Add($checkboxReAddAtStart)

            # Start on Windows Startup Checkbox
            $checkboxStartOnWindowsStart = New-Object System.Windows.Forms.CheckBox
            $checkboxStartOnWindowsStart.Text = "Start on Windows Startup"
            $checkboxStartOnWindowsStart.AutoSize = $true
            $checkboxStartOnWindowsStart.Checked = $config.StartOnWindowsStart
            $checkboxStartOnWindowsStart.Location = New-Object System.Drawing.Point(300, 15)  # Positioned to the right
            $checkboxStartOnWindowsStart.Anchor = "Top"
            $checkboxStartOnWindowsStart.Add_CheckedChanged({
                $config.StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
                Export-Configuration -ConfigPath $configPath -ConfigData $config
                Set-Startup -Enable $checkboxStartOnWindowsStart.Checked
            })
            $panelCheckboxes.Controls.Add($checkboxStartOnWindowsStart)

            # Arrange checkboxes in a single row with spacing using FlowLayoutPanel
            $flowCheckboxes = New-Object System.Windows.Forms.FlowLayoutPanel
            $flowCheckboxes.Dock = "Fill"
            $flowCheckboxes.FlowDirection = "LeftToRight"
            $flowCheckboxes.WrapContents = $false
            $flowCheckboxes.AutoSize = $false
            $flowCheckboxes.Padding = New-Object System.Windows.Forms.Padding(0)
            $flowCheckboxes.Controls.AddRange(@($checkboxReAddAtStart, $checkboxStartOnWindowsStart))
            $panelCheckboxes.Controls.Clear()
            $panelCheckboxes.Controls.Add($flowCheckboxes)

            # Function to re-add drives on script start if enabled
            function Invoke-ReAddDrivesAtStart {
                [CmdletBinding()]
                param (
                    [Parameter(Mandatory = $true)]
                    [string]$ConfigPath
                )
                $currentConfig = Import-Configuration -ConfigPath $ConfigPath
                if ($currentConfig.ReAddAtStart -eq $true -and $currentConfig.Drives.Count -gt 0) {
                    foreach ($mapping in $currentConfig.Drives) {
                        try {
                            $driveLetter = $mapping.Drive
                            $uncPath = $mapping.UNCPath
                            $useCred = $mapping.UseCredentials
                            $username = $mapping.Username
                            $password = $mapping.Password  # Renamed from 'pwd' to 'password'

                            # Prepare 'net use' command
                            if ($useCred -and $username -and $password) {
                                $cmd = "net use ${driveLetter} `"$uncPath`" `"$password`" /user:`"$username`" /persistent:yes"
                            } else {
                                $cmd = "net use ${driveLetter} `"$uncPath`" /persistent:yes"
                            }

                            # Execute the command
                            $result = cmd.exe /c $cmd 2>&1
                            if ($result -match "The command completed successfully") {
                                Write-Output "Successfully re-mapped drive ${driveLetter}: to $uncPath."
                                # Update the DataGridView
                                for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                                    $row = $dataGridView.Rows[$i]
                                    if ($row.Cells["Drive"].Value -eq $driveLetter -and $row.Cells["UNCPath"].Value -eq $uncPath) {
                                        $row.Cells["Mapped"].Value = "Yes"
                                        $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                    }
                                }
                            } else {
                                Write-Error "Failed to re-map drive ${driveLetter}: $result"
                            }
                        } catch {
                            Write-Output "Error re-adding drive ${driveLetter}: $($_.Exception.Message)"
                        }
                    }

                    # Save updated mappings to configuration
                    $mappedDrives = @()
                    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                        $currentRow = $dataGridView.Rows[$i]
                        if (-not $currentRow.IsNewRow) {
                            $mappedDrives += @{
                                Drive = $currentRow.Cells["Drive"].Value
                                UNCPath = $currentRow.Cells["UNCPath"].Value
                                AddedDate = $currentRow.Cells["AddedDate"].Value
                                Mapped = $currentRow.Cells["Mapped"].Value
                                Selected = $currentRow.Cells["Selected"].Value
                                UseCredentials = $config.Drives[$i].UseCredentials
                                Username = $config.Drives[$i].Username
                                Password = $config.Drives[$i].Password
                            }
                        }
                    }
                    $newConfig = @{
                        Drives = $mappedDrives
                        ReAddAtStart = $currentConfig.ReAddAtStart
                        StartOnWindowsStart = $currentConfig.StartOnWindowsStart
                    }
                    Export-Configuration -ConfigPath $ConfigPath -ConfigData $newConfig
                }
            }

            # Re-add drives at start if enabled
            Invoke-ReAddDrivesAtStart -ConfigPath $configPath

            # Show the form
            $form.ShowDialog()
        }

        # Main Execution
        Show-DriveMappingForm
        '''
        self.powershell_script_content = powershell_script_content

    def save_logs(self):
        """
        Saves the log console content to a file in selected format.
        """
        try:
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Logs",
                "",
                "Text Files (*.txt);;JSON Files (*.json);;XML Files (*.xml);;All Files (*)",
                options=options
            )
            if file_path:
                if file_path.endswith(".txt"):
                    with open(file_path, "w") as f:
                        f.write(self.log_console.toPlainText())
                elif file_path.endswith(".json"):
                    logs = self.log_console.toPlainText().split('\n')
                    log_entries = [{"timestamp": entry[:19], "message": entry[21:]} for entry in logs if len(entry) > 21]
                    with open(file_path, "w") as f:
                        json.dump(log_entries, f, indent=4)
                elif file_path.endswith(".xml"):
                    from xml.etree.ElementTree import Element, SubElement, ElementTree
                    root = Element("Logs")
                    logs = self.log_console.toPlainText().split('\n')
                    for entry in logs:
                        if len(entry) > 21:
                            log_entry = SubElement(root, "LogEntry")
                            timestamp, message = entry[:19], entry[21:]
                            SubElement(log_entry, "Timestamp").text = timestamp
                            SubElement(log_entry, "Message").text = message
                    tree = ElementTree(root)
                    tree.write(file_path)
                else:
                    with open(file_path, "w") as f:
                        f.write(self.log_console.toPlainText())
                self.update_log(f"Logs saved to {file_path}.")
                QMessageBox.information(self, "Save Logs", f"Logs saved to {file_path}.")
        except Exception as e:
            self.update_log(f"Error saving logs: {e}")
            QMessageBox.critical(self, "Save Logs Failed", f"Failed to save logs:\n{e}")

    def clear_logs(self):
        """
        Clears the log console and the log file.
        """
        try:
            reply = QMessageBox.question(
                self,
                "Confirm Clear Logs",
                "Are you sure you want to clear all logs?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.log_console.clear()
                # Also clear the main and timestamped log files
                for log_file in [LOG_FILE, timestamped_log_file]:
                    try:
                        open(log_file, 'w').close()
                    except Exception as e:
                        self.update_log(f"Error clearing log file {log_file}: {e}")
                        QMessageBox.critical(self, "Clear Logs Failed", f"Failed to clear log file {log_file}:\n{e}")
                        return
                self.update_log("Logs have been cleared.")
                QMessageBox.information(self, "Clear Logs", "Logs have been cleared successfully.")
            else:
                self.update_log("Clear logs canceled by user.")
        except Exception as e:
            self.update_log(f"Error during log clearing: {e}")
            QMessageBox.critical(self, "Clear Logs Error", f"An error occurred while clearing logs:\n{e}")

    def exit_application(self):
        """
        Exits the application after confirmation.
        """
        try:
            reply = QMessageBox.question(
                self,
                "Exit",
                "Are you sure you want to exit the application?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                QApplication.instance().quit()
            else:
                self.update_log("Exit canceled by user.")
        except Exception as e:
            logger.error(f"Error during application exit: {e}")
            QMessageBox.critical(self, "Exit Error", f"An error occurred while exiting:\n{e}")

    def toggle_console(self):
        """
        Toggles the visibility of the log console.
        """
        try:
            if self.log_console.isVisible():
                self.log_console.hide()
                self.update_log("Log console hidden.")
            else:
                self.log_console.show()
                self.update_log("Log console shown.")
        except Exception as e:
            logger.error(f"Error toggling console: {e}")
            QMessageBox.critical(self, "Toggle Console Error", f"An error occurred while toggling the console:\n{e}")

    def toggle_startup_from_tray(self, checked):
        """
        Toggles 'Start on Windows Startup' from the tray menu.
        """
        try:
            self.startup_enabled = checked
            self.set_startup(checked)
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
            self.update_log(f"'Start on Windows Startup' set to {checked}.")
        except Exception as e:
            logger.error(f"Error toggling startup from tray: {e}")
            QMessageBox.critical(self, "Startup Toggle Error", f"An error occurred while toggling startup:\n{e}")

    def toggle_readd_from_tray(self, checked):
        """
        Toggles 'Re-Add On Startup' from the tray menu.
        """
        try:
            self.auto_readd_enabled = checked
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
            self.update_log(f"'Re-Add On Startup' set to {checked}.")
        except Exception as e:
            logger.error(f"Error toggling re-add from tray: {e}")
            QMessageBox.critical(self, "Re-Add Toggle Error", f"An error occurred while toggling re-add:\n{e}")

    def toggle_startup(self, state):
        """
        Enables or disables the application to start on Windows startup.
        """
        try:
            if state == Qt.Checked:
                self.startup_enabled = True
                self.set_startup(True)
                self.update_log("Enabled 'Start on Windows Startup'.")
            else:
                self.startup_enabled = False
                self.set_startup(False)
                self.update_log("Disabled 'Start on Windows Startup'.")
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
        except Exception as e:
            logger.error(f"Error toggling startup: {e}")
            QMessageBox.critical(self, "Startup Toggle Error", f"An error occurred while toggling startup:\n{e}")

    def toggle_readd_from_menu(self, checked):
        """
        Toggles 'Re-Add On Startup' from the settings menu.
        """
        try:
            self.auto_readd_enabled = checked
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
            self.update_log(f"'Re-Add On Startup' set to {checked}.")
        except Exception as e:
            logger.error(f"Error toggling re-add from menu: {e}")
            QMessageBox.critical(self, "Re-Add Toggle Error", f"An error occurred while toggling re-add:\n{e}")

    def set_startup(self, enable):
        """
        Sets the application to run on Windows startup by modifying the registry.
        """
        try:
            import winreg
            reg_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE) as key:
                if enable:
                    script_path = os.path.abspath(sys.argv[0])
                    python_exe = sys.executable
                    winreg.SetValueEx(key, "TSTPDriveMapper", 0, winreg.REG_SZ, f'"{python_exe}" "{script_path}"')
                else:
                    try:
                        winreg.DeleteValue(key, "TSTPDriveMapper")
                    except FileNotFoundError:
                        pass
        except Exception as e:
            self.update_log(f"Error setting startup behavior: {e}")
            QMessageBox.critical(self, "Startup Error", f"Failed to set startup behavior:\n{e}")

    def save_startup_settings(self):
        """
        Saves the current startup and auto-readd settings to the settings file.
        """
        try:
            settings = {
                "startup_enabled": self.startup_enabled,
                "auto_readd_enabled": self.auto_readd_enabled,
                "light_mode": self.light_mode
            }
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r") as f:
                    current_settings = json.load(f)
            else:
                current_settings = {}
            current_settings.update(settings)
            with open(SETTINGS_FILE, "w") as f:
                json.dump(current_settings, f, indent=4)
            self.update_log("Startup settings saved.")
        except Exception as e:
            self.update_log(f"Error saving startup settings: {e}")
            QMessageBox.critical(self, "Save Settings Error", f"Failed to save startup settings:\n{e}")

    def load_startup_settings(self):
        """
        Loads startup and auto-readd settings from the settings file.
        """
        try:
            with open(SETTINGS_FILE, "r") as f:
                settings = json.load(f)
            self.startup_enabled = settings.get("startup_enabled", False)
            self.auto_readd_enabled = settings.get("auto_readd_enabled", False)
            self.light_mode = settings.get("light_mode", False)
            self.startup_checkbox.setChecked(self.startup_enabled)
            self.auto_readd_checkbox.setChecked(self.auto_readd_enabled)
            self.update_log("Startup settings loaded.")
        except Exception as e:
            self.update_log(f"Error loading startup settings: {e}")

    def readd_drives(self):
        """
        Initiates the remove and add drives process on startup using a background thread.
        """
        try:
            self.update_log("Initiating remove and add drives on startup...")
            self.readd_thread = ReaddDrivesThread(self.drive_mappings)
            self.readd_thread.log_signal.connect(self.update_log)
            self.readd_thread.finished_signal.connect(self.update_drives_table_ui)
            self.readd_thread.start()
        except Exception as e:
            logger.error(f"Error initiating remove and add drives: {e}")
            QMessageBox.critical(self, "Re-Add Drives Error", f"An error occurred while re-adding drives:\n{e}")

    def force_connect(self, row):
        """
        Forces a drive to connect using stored credentials.
        If not mapped, it changes to 'Connect' and prompts for credentials.
        """
        try:
            drive_info = self.drive_mappings[row]
            drive_letter = drive_info["Drive"]
            unc_path = drive_info["UNCPath"]

            # Disconnect the drive if it is already in use
            disconnect_command = f'net use {drive_letter} /delete /y'
            execute_cmd(disconnect_command)

            # Force Auth: Prompt for credentials if necessary
            use_credentials = drive_info.get("UseCredentials", False)
            if use_credentials:
                creds_dialog = CredentialsDialog(parent=self)
                if creds_dialog.exec_() == QDialog.Accepted:
                    username, password = creds_dialog.get_credentials()
                else:
                    self.update_log(f"Force authorization for drive {drive_letter} canceled by user.")
                    return
                command = f'net use {drive_letter} "{unc_path}" "{password}" /user:{username} /persistent:no'
            else:
                command = f'net use {drive_letter} "{unc_path}" /persistent:no'

            stdout, stderr = execute_cmd(command)
            if stderr:
                # Retry without trailing backslash
                if unc_path.endswith("\\"):
                    unc_path_retry = unc_path.rstrip("\\")
                    if use_credentials:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" "{password}" /user:{username} /persistent:no'
                    else:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" /persistent:no'
                    stdout_retry, stderr_retry = execute_cmd(command_retry)
                    if stderr_retry:
                        error_message = f"Error forcing authorization for drive {drive_letter}: {stderr_retry}"
                        self.update_log(error_message)
                        QMessageBox.critical(self, "Force Auth Error", error_message)
                    else:
                        success_message = f"Successfully forced authorization for drive {drive_letter} to {unc_path_retry}."
                        self.update_log(success_message)
                        QMessageBox.information(self, "Force Auth", success_message)
                        self.drive_mappings[row]["Mapped"] = "Yes"
                else:
                    error_message = f"Error forcing authorization for drive {drive_letter}: {stderr}"
                    self.update_log(error_message)
                    QMessageBox.critical(self, "Force Auth Error", error_message)
            else:
                success_message = f"Successfully forced authorization for drive {drive_letter} to {unc_path}."
                self.update_log(success_message)
                QMessageBox.information(self, "Force Auth", success_message)
                self.drive_mappings[row]["Mapped"] = "Yes"
        except Exception as e:
            logger.error(f"Error forcing authorization: {e}")
            QMessageBox.critical(self, "Force Auth Error", f"An error occurred while forcing authorization for the drive:\n{e}")

    def connect_drive(self, row):
        """
        Connects a drive by prompting for credentials if necessary.
        """
        try:
            drive_info = self.drive_mappings[row]
            drive_letter = drive_info["Drive"]
            unc_path = drive_info["UNCPath"]
            use_credentials = drive_info.get("UseCredentials", False)

            if use_credentials:
                # Prompt for username and password
                creds_dialog = CredentialsDialog(parent=self)
                if creds_dialog.exec_() == QDialog.Accepted:
                    username, password = creds_dialog.get_credentials()
                    command = f'net use {drive_letter} "{unc_path}" "{password}" /user:{username} /persistent:no'
                else:
                    self.update_log(f"Connect for drive {drive_letter} canceled by user.")
                    return
            else:
                command = f'net use {drive_letter} "{unc_path}" /persistent:no'

            stdout, stderr = execute_cmd(command)
            if stderr:
                # Retry without trailing backslash
                if unc_path.endswith("\\"):
                    unc_path_retry = unc_path.rstrip("\\")
                    if use_credentials:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" "{password}" /user:{username} /persistent:no'
                    else:
                        command_retry = f'net use {drive_letter} "{unc_path_retry}" /persistent:no'
                    stdout_retry, stderr_retry = execute_cmd(command_retry)
                    if stderr_retry:
                        error_message = f"Error connecting drive {drive_letter}: {stderr_retry}"
                        self.update_log(error_message)
                        QMessageBox.critical(self, "Connect Error", error_message)
                    else:
                        success_message = f"Successfully connected drive {drive_letter} to {unc_path_retry}."
                        self.update_log(success_message)
                        QMessageBox.information(self, "Connect", success_message)
                        self.drive_mappings[row]["Mapped"] = "Yes"
                else:
                    error_message = f"Error connecting drive {drive_letter}: {stderr}"
                    self.update_log(error_message)
                    QMessageBox.critical(self, "Connect Error", error_message)
            else:
                success_message = f"Successfully connected drive {drive_letter} to {unc_path}."
                self.update_log(success_message)
                QMessageBox.information(self, "Connect", success_message)
                self.drive_mappings[row]["Mapped"] = "Yes"

            # Update the table
            self.populate_drives_table()
            save_settings(self.drive_mappings, self.startup_enabled, self.auto_readd_enabled, self.light_mode)
        except Exception as e:
            logger.error(f"Error during connect: {e}")
            QMessageBox.critical(self, "Connect Error", f"An error occurred while connecting the drive:\n{e}")

    def save_startup_settings(self):
        """
        Saves the current startup and auto-readd settings to the settings file.
        """
        try:
            settings = {
                "startup_enabled": self.startup_enabled,
                "auto_readd_enabled": self.auto_readd_enabled,
                "light_mode": self.light_mode
            }
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r") as f:
                    current_settings = json.load(f)
            else:
                current_settings = {}
            current_settings.update(settings)
            with open(SETTINGS_FILE, "w") as f:
                json.dump(current_settings, f, indent=4)
            self.update_log("Startup settings saved.")
        except Exception as e:
            self.update_log(f"Error saving startup settings: {e}")
            QMessageBox.critical(self, "Save Settings Error", f"Failed to save startup settings:\n{e}")

    def show_info_dialog(self, title, message):
        """
        Displays informational dialogs for About, Tutorial, Donate, and Website.
        Adapts to the current theme (light or dark mode) and is resizable.
        """
        try:
            dialog = InfoDialog(title, message, self)

            # Apply theme-specific styles
            if self.light_mode:
                dialog.setStyleSheet("""
                    QDialog {
                        background-color: #ffffff;
                        color: white;
                    }
                    QLabel {
                        background-color: #ffffff;
                        color: black;
                        font-size: 14px;
                    }
                    QPushButton {
                        background-color: #f0f0f0;
                        color: black;
                    }
                    QScrollArea {
                        background-color: #ffffff;
                    }
                    QWidget {
                        background-color: #ffffff;
                    }
                """)
            else:
                dialog.setStyleSheet("""
                    QDialog {
                        background-color: #2b2b2b;
                        color: black;
                    }
                    QLabel {
                        background-color: #2b2b2b;
                        color: white;
                        font-size: 14px;
                    }
                    QPushButton {
                        background-color: #3c3f41;
                        color: white;
                    }
                    QScrollArea {
                        background-color: #2b2b2b;
                    }
                    QWidget {
                        background-color: #2b2b2b;
                    }
                """)

            dialog.exec_()
        except Exception as e:
            logger.error(f"Error showing info dialog: {e}")
            QMessageBox.critical(self, "Info Dialog Error", f"An error occurred while showing the info dialog:\n{e}")

class InfoDialog(QDialog):
    def __init__(self, title, html_content, parent=None):
        super(InfoDialog, self).__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(800, 600)

        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setHtml(html_content)
        self.text_edit.setReadOnly(True)  # Make it read-only
        layout.addWidget(self.text_edit)

        self.setLayout(layout)

# Dialog for Entering Credentials
class CredentialsDialog(QDialog):
    def __init__(self, parent=None):
        super(CredentialsDialog, self).__init__(parent)
        self.setWindowTitle("Enter Credentials")
        self.setFixedSize(350, 150)
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
                color: white;
                border: 2px solid #3c3f41;
                border-radius: 10px;
            }
            QLabel {
                color: white;
                font-size: 12px;
            }
            QLineEdit {
                background-color: #1e1e1e;
                color: white;
                border: 1px solid #555555;
                border-radius: 5px;
                padding: 5px;
            }
            QPushButton {
                background-color: #3c3f41;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #4e5254;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Username
        username_layout = QHBoxLayout()
        username_label = QLabel("Username:")
        self.username_input = QLineEdit()
        username_layout.addWidget(username_label)
        username_layout.addWidget(self.username_input)
        layout.addLayout(username_layout)

        # Password
        password_layout = QHBoxLayout()
        password_label = QLabel("Password:")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        password_layout.addWidget(password_label)
        password_layout.addWidget(self.password_input)
        layout.addLayout(password_layout)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Cancel")
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        # Connect Buttons
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

    def get_credentials(self):
        """
        Retrieves the entered username and password.
        """
        return self.username_input.text().strip(), self.password_input.text().strip()

class MultiEditDriveDialog(QDialog):
    def __init__(self, drives_to_edit, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Drives")
        self.setMinimumSize(600, 400)

        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_layout.addWidget(scroll_area)

        container = QWidget()
        scroll_area.setWidget(container)

        container_layout = QVBoxLayout()
        container.setLayout(container_layout)

        self.edit_sections = []

        for drive in drives_to_edit:
            frame = QFrame()
            frame.setStyleSheet("""
                QFrame {
                    background-color: #f0f0f0;
                    border: 1px solid #dddddd;
                    border-radius: 8px;
                    padding: 15px;
                    margin-bottom: 20px;
                }
            """)
            layout = QVBoxLayout()
            frame.setLayout(layout)

            # Drive Letter
            drive_letter_layout = QHBoxLayout()
            drive_letter_label = QLabel("Drive Letter:")
            drive_letter_input = QLineEdit(drive["Drive"])
            drive_letter_layout.addWidget(drive_letter_label)
            drive_letter_layout.addWidget(drive_letter_input)
            layout.addLayout(drive_letter_layout)

            # UNC Path
            unc_path_layout = QHBoxLayout()
            unc_path_label = QLabel("UNC Path:")
            unc_path_input = QLineEdit(drive["UNCPath"])
            unc_path_layout.addWidget(unc_path_label)
            unc_path_layout.addWidget(unc_path_input)
            layout.addLayout(unc_path_layout)

            # Use Credentials
            use_credentials_checkbox = QCheckBox("Use Different Credentials")
            use_credentials_checkbox.setChecked(drive["UseCredentials"])
            layout.addWidget(use_credentials_checkbox)

            # Username
            username_layout = QHBoxLayout()
            username_label = QLabel("Username:")
            username_input = QLineEdit(drive["Username"])
            username_input.setEnabled(drive["UseCredentials"])
            username_layout.addWidget(username_label)
            username_layout.addWidget(username_input)
            layout.addLayout(username_layout)

            # Password
            password_layout = QHBoxLayout()
            password_label = QLabel("Password:")
            password_input = QLineEdit(drive["Password"])
            password_input.setEchoMode(QLineEdit.Password)
            password_input.setEnabled(drive["UseCredentials"])
            password_layout.addWidget(password_label)
            password_layout.addWidget(password_input)
            layout.addLayout(password_layout)

            # Connect checkbox to enable/disable username and password fields
            use_credentials_checkbox.stateChanged.connect(
                lambda state, u=username_input, p=password_input: self.toggle_credentials(state, u, p)
            )

            self.edit_sections.append({
                "DriveLetter": drive_letter_input,
                "UNCPath": unc_path_input,
                "UseCredentials": use_credentials_checkbox,
                "Username": username_input,
                "Password": password_input,
                "OriginalDrive": drive["Drive"]
            })

            container_layout.addWidget(frame)

        # Buttons
        buttons_layout = QHBoxLayout()
        save_button = QPushButton("Save")
        save_button.setFixedHeight(40)
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        save_button.clicked.connect(self.accept)
        buttons_layout.addStretch()
        buttons_layout.addWidget(save_button)

        main_layout.addLayout(buttons_layout)

    def toggle_credentials(self, state, username_field, password_field):
        if state == Qt.Checked:
            username_field.setEnabled(True)
            password_field.setEnabled(True)
        else:
            username_field.setEnabled(False)
            password_field.setEnabled(False)

    def get_drive_entries(self):
        edited_drives = []
        for section in self.edit_sections:
            drive_entry = {
                "OriginalDrive": section["OriginalDrive"],
                "Drive": section["DriveLetter"].text(),
                "UNCPath": section["UNCPath"].text(),
                "UseCredentials": section["UseCredentials"].isChecked(),
                "Username": section["Username"].text(),
                "Password": section["Password"].text()
            }
            edited_drives.append(drive_entry)
        return edited_drives

# Main Execution
def main():
    app = QApplication(sys.argv)

    # Apply Dark Theme by default (can be toggled to Light Mode)
    # Handled via stylesheets in the MainWindow class

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
