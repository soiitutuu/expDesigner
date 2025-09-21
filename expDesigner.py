import sys
import os
import json
import ctypes
import subprocess
import logging
import time
import winreg
from typing import Any, Dict, List, Optional, Tuple

APP_NAME = "expDesigner"
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "expDesigner.config.json")
BACKUP_REG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "expDesigner-backup.reg")
LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "expDesigner.log")

logging.basicConfig(
    filename=LOG_PATH,
    filemode="a",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

def _global_exception_hook(exc_type, exc_value, exc_tb):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_tb))
    print("an unexpected error occurred. details were written to expDesigner.log.")

sys.excepthook = _global_exception_hook
logging.info("%s starting...", APP_NAME)

def is_running_as_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        logging.exception("Failed to determine admin status")
        return False

def relaunch_as_admin() -> bool:
    try:
        logging.info("Attempting to relaunch with administrative privileges via UAC...")
        if getattr(sys, "frozen", False):
            executable = sys.executable
            params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
        else:
            executable = sys.executable
            script = os.path.abspath(sys.argv[0])
            params = " ".join([f'"{script}"'] + [f'"{arg}"' for arg in sys.argv[1:]])
        ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, params, None, 1)
        logging.info(f"UAC ShellExecuteW return code: {ret}")
        return int(ret) > 32
    except Exception:
        logging.exception("UAC elevation failed")
        return False

class RegistryManager:
    ROOTS = {
        "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
        "HKCU": winreg.HKEY_CURRENT_USER,
        "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
        "HKLM": winreg.HKEY_LOCAL_MACHINE,
        "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
        "HKCR": winreg.HKEY_CLASSES_ROOT,
        "HKEY_USERS": winreg.HKEY_USERS,
        "HKU": winreg.HKEY_USERS,
        "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
        "HKCC": winreg.HKEY_CURRENT_CONFIG
    }

    @staticmethod
    def _split_path(path: str):
        if "\\" not in path:
            raise ValueError(f"Invalid registry path: {path}")
        root_name, subkey = path.split("\\", 1)
        root = RegistryManager.ROOTS.get(root_name)
        if root is None:
            raise ValueError(f"Unknown registry root in path: {path}")
        return root, subkey

    @staticmethod
    def open_key(path: str, access=winreg.KEY_READ):
        root, subkey = RegistryManager._split_path(path)
        desired = access | winreg.KEY_WOW64_64KEY
        try:
            return winreg.OpenKeyEx(root, subkey, 0, desired)
        except FileNotFoundError:
            if access & winreg.KEY_WRITE:
                return winreg.CreateKeyEx(root, subkey, 0, desired)
            raise

    @staticmethod
    def read_value(path: str, name: str, default=None):
        try:
            with RegistryManager.open_key(path, winreg.KEY_READ) as key:
                val, regtype = winreg.QueryValueEx(key, name)
                return val, regtype
        except FileNotFoundError:
            return default, None
        except OSError:
            logging.exception(f"Error reading registry value: {path}\\{name}")
            return default, None

    @staticmethod
    def read_dword(path: str, name: str, default=None):
        val, _ = RegistryManager.read_value(path, name, default)
        try:
            return int(val)
        except Exception:
            return default

    @staticmethod
    def write_value(path: str, name: str, value: Any, reg_type: int) -> bool:
        try:
            with RegistryManager.open_key(path, winreg.KEY_WRITE) as key:
                if reg_type == winreg.REG_DWORD:
                    value = int(value) & 0xFFFFFFFF
                winreg.SetValueEx(key, name, 0, reg_type, value)
            logging.info(f"Registry set: {path}\\{name} = {value} (type {reg_type})")
            return True
        except Exception:
            logging.exception(f"Error writing registry value: {path}\\{name} = {value} (type {reg_type})")
            return False

def restart_explorer():
    try:
        logging.info("Restarting Windows Explorer...")
        subprocess.call('taskkill /f /im explorer.exe & start explorer.exe', shell=True)
        logging.info("Explorer restarted")
    except Exception:
        logging.exception("Failed to restart Explorer")

def get_windows_build_number() -> int:
    try:
        val, _ = RegistryManager.read_value(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber", "0")
        return int(str(val))
    except Exception:
        return 0

def get_accent_color() -> Optional[Tuple[int, int, int]]:
    try:
        col = RegistryManager.read_dword(r"HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "ColorizationColor", None)
        if col is None:
            return None
        _a = (col >> 24) & 0xFF
        b = (col >> 16) & 0xFF
        g = (col >> 8) & 0xFF
        r = (col >> 0) & 0xFF
        return (r, g, b)
    except Exception:
        return None

def detect_windows_theme() -> str:
    v = RegistryManager.read_dword(r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
    return "dark" if v == 0 else "light"

def load_config() -> Dict[str, Any]:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            logging.exception("Failed to read config")
    return {"favorites": [], "theme": None, "compact": False, "restore_point": False}

def save_config(cfg: Dict[str, Any]):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        logging.exception("Failed to save config")

def main_app():
    app = __import__('PyQt6.QtWidgets', fromlist=['QApplication']).QApplication(sys.argv)
    from PyQt6.QtCore import Qt, QSize, QTimer, QPoint
    from PyQt6.QtGui import QPalette, QColor, QAction, QActionGroup
    from PyQt6.QtWidgets import (
        QMainWindow, QWidget, QLabel, QVBoxLayout, QHBoxLayout, QFrame,
        QListWidget, QListWidgetItem, QStackedWidget, QPushButton, QComboBox,
        QCheckBox, QStyle, QSizePolicy, QStatusBar, QScrollArea, QToolButton,
        QFileDialog, QLineEdit, QSplitter, QSpinBox, QMessageBox, QDialog,
        QPlainTextEdit
    )

    def apply_theme(app_, theme: str = "dark"):
        app_.setStyle("Fusion")
        pal = QPalette()
        accent = get_accent_color() or (10, 132, 255)
        if theme == "dark":
            pal.setColor(QPalette.ColorRole.Window, QColor(37, 37, 38))
            pal.setColor(QPalette.ColorRole.WindowText, QColor(220, 220, 220))
            pal.setColor(QPalette.ColorRole.Base, QColor(30, 30, 30))
            pal.setColor(QPalette.ColorRole.AlternateBase, QColor(45, 45, 48))
            pal.setColor(QPalette.ColorRole.ToolTipBase, QColor(220, 220, 220))
            pal.setColor(QPalette.ColorRole.ToolTipText, QColor(220, 220, 220))
            pal.setColor(QPalette.ColorRole.Text, QColor(220, 220, 220))
            pal.setColor(QPalette.ColorRole.Button, QColor(45, 45, 48))
            pal.setColor(QPalette.ColorRole.ButtonText, QColor(220, 220, 220))
            pal.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
            pal.setColor(QPalette.ColorRole.Highlight, QColor(*accent))
            pal.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
            pal.setColor(QPalette.ColorRole.PlaceholderText, QColor(160, 160, 160))
        else:
            pal = app_.palette()
            pal.setColor(QPalette.ColorRole.Highlight, QColor(*accent))
        app_.setPalette(pal)

    cfg = load_config()
    t = cfg.get("theme")
    if t not in ("dark", "light"):
        t = detect_windows_theme()
        cfg["theme"] = t
        save_config(cfg)
    apply_theme(app, t)

    def human_requires(req: Optional[str]) -> str:
        return {"explorer": "Explorer", "logoff": "Sign-out", "reboot": "Reboot"}.get(req or "", "")

    def badge_label(text: str) -> QLabel:
        lab = QLabel(text)
        lab.setStyleSheet("QLabel { color: white; background: #0078D4; border-radius: 8px; padding: 2px 6px; font-size: 11px; }")
        return lab

    class Toast(QWidget):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.WindowType.ToolTip | Qt.WindowType.FramelessWindowHint)
            self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
            self.label = QLabel("", self)
            self.label.setStyleSheet("QLabel { background: rgba(0,0,0,0.8); color: white; padding: 10px 14px; border-radius: 8px; }")
            lay = QVBoxLayout(self)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.addWidget(self.label)
            self.timer = QTimer(self)
            self.timer.timeout.connect(self.hide)

        def show_toast(self, text: str, ms: int = 2500):
            self.label.setText(text)
            self.adjustSize()
            if self.parent() and self.parent().isVisible():
                par = self.parent().geometry()
                x = par.x() + par.width() - self.width() - 20
                y = par.y() + par.height() - self.height() - 20
                self.move(QPoint(x, y))
            self.show()
            self.raise_()
            self.timer.start(ms)

    class CollapsibleSection(QWidget):
        def __init__(self, title: str, parent=None, compact: bool = False):
            super().__init__(parent)
            self.compact = compact
            self.header = QToolButton(self)
            self.header.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
            self.header.setArrowType(Qt.ArrowType.RightArrow)
            self.header.setText(title)
            self.header.setCheckable(True)
            self.header.setChecked(True)
            self.header.clicked.connect(self._on_clicked)
            self.container = QWidget(self)
            self.vlay = QVBoxLayout(self.container)
            self.vlay.setContentsMargins(4, 4, 4, 4)
            self.vlay.setSpacing(8 if not self.compact else 4)
            lay = QVBoxLayout(self)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.addWidget(self.header)
            lay.addWidget(self.container)

        def _on_clicked(self):
            checked = self.header.isChecked()
            self.container.setVisible(checked)
            self.header.setArrowType(Qt.ArrowType.DownArrow if checked else Qt.ArrowType.RightArrow)

        def add_row(self, w: QWidget):
            self.vlay.addWidget(w)

    WIN_BUILD = get_windows_build_number()

    def D(section: str, label: str, tooltip: str, path: str, name: str,
          setting_id: str, reg_type="dword", requires=None, page="Personalization",
          s_type="switch", on=1, off=0, default=None, choices=None, minBuild: Optional[int]=None,
          spin=None):
        return {
            "id": setting_id, "page": page, "section": section, "label": label, "tooltip": tooltip,
            "path": path, "name": name, "reg_type": reg_type, "requires": requires,
            "type": s_type, "on": on, "off": off, "default": default,
            "choices": choices, "minBuild": minBuild, "spin": spin
        }

    settings_schema: List[Dict[str, Any]] = []
    personalize = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    dwm = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM"
    explorer_adv = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

    settings_schema += [
        D("Theme", "Force Dark Mode (Apps)", "Forces supported apps to use Dark theme.", personalize, "AppsUseLightTheme",
          "apps_dark", requires="explorer", on=0, off=1, default=1),
        D("Theme", "Force Dark Mode (System UI)", "Sets system UI (Start, Taskbar) to Dark.", personalize, "SystemUsesLightTheme",
            "system_dark", requires="explorer", on=0, off=1, default=1),
        D("Effects", "Transparency Effects", "Enable acrylic/transparency effects.", personalize, "EnableTransparency",
          "transparency", requires="explorer", on=1, off=0, default=1),
        D("Accent", "Accent Color on Start/Taskbar", "Show accent color on Start, Taskbar and action center.",
          personalize, "ColorPrevalence", "accent_taskbar", requires="explorer", on=1, off=0, default=0),
        D("Accent", "Accent Color on Title bars and Windows", "Show accent color on title bars and window borders.",
          dwm, "ColorPrevalence", "accent_title", requires="explorer", on=1, off=0, default=0),
        D("Animations", "Taskbar Animations", "Enable/disable taskbar animations.",
          explorer_adv, "TaskbarAnimations", "taskbar_animations", requires="explorer", on=1, off=0, default=1),
        D("Gestures", "Disable Aero Shake", "Prevents windows from minimizing when shaking a title bar.",
          explorer_adv, "DisallowShaking", "disable_aero_shake", requires="explorer", on=1, off=0, default=0),
    ]

    search_path = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Search"
    feeds_path = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Feeds"
    settings_schema += [
        D("Taskbar", "Search Bar Style", "Choose between hiding, icon-only, or full search box.",
          search_path, "SearchboxTaskbarMode", "search_style", page="Taskbar & Start", s_type="combo",
          choices=[("Hide", 0), ("Icon Only", 1), ("Search Box", 2)], default=1, requires="explorer"),
        D("Taskbar", "Taskbar Size (Windows 11)", "Set taskbar icon size (requires Explorer restart).",
          explorer_adv, "TaskbarSi", "taskbar_size", page="Taskbar & Start", s_type="combo",
          choices=[("Small", 0), ("Medium", 1), ("Large", 2)], default=1, requires="explorer"),
        D("Taskbar", "Taskbar Alignment (Windows 11)", "Align taskbar icons.",
          explorer_adv, "TaskbarAl", "taskbar_align", page="Taskbar & Start", s_type="combo",
          choices=[("Left", 0), ("Center", 1)], default=1, requires="explorer"),
        D("Taskbar", "Combine Taskbar Buttons", "Show/Combine taskbar labels (Win11 may limit).",
          explorer_adv, "TaskbarGlomLevel", "taskbar_glom", page="Taskbar & Start", s_type="combo",
          choices=[("Always combine, hide labels", 0), ("Combine when taskbar is full", 1), ("Never combine", 2)], default=0, requires="explorer"),
        D("Clock", "Show Seconds in Taskbar Clock", "Adds seconds to taskbar clock (uses more resources).",
          explorer_adv, "ShowSecondsInSystemClock", "clock_seconds", page="Taskbar & Start", on=1, off=0, default=0, requires="explorer"),
        D("Buttons", "Show Task View Button", "Shows the Task View (virtual desktops) button.",
          explorer_adv, "ShowTaskViewButton", "taskview_btn", page="Taskbar & Start", on=1, off=0, default=1, requires="explorer"),
        D("Buttons", "Show Widgets Button (Windows 11)", "Shows the Widgets button on taskbar (if available).",
          explorer_adv, "TaskbarMn", "widgets_btn", page="Taskbar & Start", on=1, off=0, default=0, requires="explorer"),
        D("News", "News & Interests (Windows 10)", "Controls the News & Interests taskbar widget.",
          feeds_path, "ShellFeedsTaskbarViewMode", "news_interests", page="Taskbar & Start", s_type="combo",
          choices=[("Icon and text", 0), ("Icon only", 1), ("Off", 2)], default=2, requires="explorer"),
        D("Start Menu", "Show Most Used Apps", "Show most used apps in Start.",
          explorer_adv, "Start_TrackProgs", "start_most_used", page="Taskbar & Start", on=1, off=0, default=1, requires="explorer"),
        D("Start Menu", "Show Recently Added/Opened", "Show recently added apps and opened items.",
          explorer_adv, "Start_TrackDocs", "start_recent", page="Taskbar & Start", on=1, off=0, default=1, requires="explorer"),
    ]

    settings_schema += [
        D("Visibility", "Show Hidden Files", "Show items with the Hidden attribute.",
          explorer_adv, "Hidden", "show_hidden", page="File Explorer & UI", on=1, off=2, default=2, requires="explorer"),
        D("Visibility", "Show File Extensions", "Show known file type extensions.",
          explorer_adv, "HideFileExt", "show_ext", page="File Explorer & UI", on=0, off=1, default=1, requires="explorer"),
        D("Visibility", "Show Protected OS Files", "Show protected operating system files (be careful).",
          explorer_adv, "ShowSuperHidden", "show_superhidden", page="File Explorer & UI", on=1, off=0, default=0, requires="explorer"),
        D("Layout", "Use Compact View (Windows 11)", "Smaller spacing in File Explorer lists.",
          explorer_adv, "UseCompactMode", "compact_view", page="File Explorer & UI", on=1, off=0, default=0, requires="explorer"),
        D("Behavior", "Open File Explorer to", "Choose default start location for File Explorer.",
          explorer_adv, "LaunchTo", "explorer_launchto", page="File Explorer & UI", s_type="combo",
          choices=[("Home (Windows 11)", 0), ("This PC", 1), ("Quick Access", 2)], default=1, requires="explorer"),
        D("UI", "Show Status Bar", "Show status bar at the bottom of File Explorer.",
          explorer_adv, "ShowStatusBar", "status_bar", page="File Explorer & UI", on=1, off=0, default=1, requires="explorer"),
        D("UI", "Show Info Tips on Hover", "Show pop-up info tooltips when hovering items.",
          explorer_adv, "ShowInfoTip", "info_tips", page="File Explorer & UI", on=1, off=0, default=1, requires="explorer"),
        D("Selection", "Use Check Boxes to Select Items", "Show check boxes for item selection.",
          explorer_adv, "AutoCheckSelect", "checkbox_select", page="File Explorer & UI", on=1, off=0, default=0, requires="explorer"),
        D("Navigation", "Expand to Current Folder", "Automatically expand navigation pane to the current folder.",
          explorer_adv, "NavPaneExpandToCurrentFolder", "nav_expand", page="File Explorer & UI", on=1, off=0, default=0, requires="explorer"),
        D("Process", "Launch Folder Windows in a Separate Process", "Improves stability; uses more memory.",
          explorer_adv, "SeparateProcess", "separate_process", page="File Explorer & UI", on=1, off=0, default=0, requires="explorer"),
        D("Sync", "Show Sync Provider Notifications", "Show OneDrive/Provider notifications in Explorer.",
          explorer_adv, "ShowSyncProviderNotifications", "sync_provider", page="File Explorer & UI", on=1, off=0, default=1, requires="explorer"),
        D("Taskbar", "Show Taskbar Badges", "Show badges on taskbar buttons.",
          explorer_adv, "TaskbarBadges", "taskbar_badges", page="File Explorer & UI", on=1, off=0, default=1, requires="explorer"),
    ]

    adv_id = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
    policy_sys = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\System"
    policy_dc = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
    cloudcontent = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    location_pol = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors"
    cortana_pol = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    cdman = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    input_personal = r"HKEY_CURRENT_USER\Software\Microsoft\InputPersonalization"
    privacy_root = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Privacy"
    settings_schema += [
        D("Advertising", "Disable Advertising ID", "Disables per-user advertising identifier.",
          adv_id, "Enabled", "ads_id", page="Privacy", on=0, off=1, default=1),
        D("Activity", "Disable Activity History", "Turns off collection/upload of activity history.",
          policy_sys, "PublishUserActivities", "activity_history", page="Privacy", on=0, off=1, default=1, requires="logoff"),
        D("Activity", "Disable Activity Upload", "Blocks uploading activity history to Microsoft.",
          policy_sys, "UploadUserActivities", "activity_upload", page="Privacy", on=0, off=1, default=1, requires="logoff"),
        D("Diagnostics", "Diagnostic Data Level", "Lower is more private; availability depends on edition.",
          policy_dc, "AllowTelemetry", "telemetry", page="Privacy", s_type="combo",
          choices=[("Security (0)", 0), ("Basic (1)", 1), ("Enhanced (2)", 2), ("Full (3)", 3)], default=3),
        D("Consumer", "Disable Consumer Experience", "Prevents suggested apps and dynamic content.",
          cloudcontent, "DisableConsumerFeatures", "consumer_features", page="Privacy", on=1, off=0, default=0),
        D("Clipboard", "Disable Clipboard History", "Turns off Windows clipboard history.",
          policy_sys, "AllowClipboardHistory", "clipboard_history", page="Privacy", on=0, off=1, default=1),
        D("Location", "Disable Location", "Globally disable Windows location services.",
          location_pol, "DisableLocation", "disable_location", page="Privacy", on=1, off=0, default=0),
        D("Cortana", "Disable Cortana", "Turns off Cortana (requires sign-out).",
          cortana_pol, "AllowCortana", "cortana", page="Privacy", on=0, off=1, default=1, requires="logoff"),
        D("Suggestions", "Disable Tips, Tricks, and Suggestions", "Disables OS tips and suggestions.",
          cdman, "SubscribedContent-338393Enabled", "tips_tricks", page="Privacy", on=0, off=1, default=1),
        D("Suggestions", "Disable App Silent Installs", "Prevents suggested apps auto-install.",
          cdman, "SilentInstalledAppsEnabled", "silent_apps", page="Privacy", on=0, off=1, default=1),
        D("Suggestions", "Disable System Pane Suggestions", "Turns off suggestions in the Settings sidebar.",
          cdman, "SystemPaneSuggestionsEnabled", "pane_suggestions", page="Privacy", on=0, off=1, default=1),
        D("Spotlight", "Disable Lock Screen Spotlight", "Disables Windows Spotlight on lock screen.",
          cdman, "RotatingLockScreenEnabled", "lock_spotlight", page="Privacy", on=0, off=1, default=1),
        D("Spotlight", "Disable Lock Screen Overlay", "Disables Spotlight fun facts/trivia on lock screen.",
          cdman, "RotatingLockScreenOverlayEnabled", "lock_overlay", page="Privacy", on=0, off=1, default=1),
        D("Tailored", "Disable Tailored Experiences", "Disable personalization based on diagnostic data.",
          privacy_root, "TailoredExperiencesWithDiagnosticDataEnabled", "tailored_experiences", page="Privacy", on=0, off=1, default=1),
        D("Typing & Ink", "Disable Typing Data Collection", "Disables text input personalization.",
          input_personal, "RestrictImplicitTextCollection", "typing_collect", page="Privacy", on=1, off=0, default=0),
        D("Typing & Ink", "Disable Inking Data Collection", "Disables inking input personalization.",
          input_personal, "RestrictImplicitInkCollection", "inking_collect", page="Privacy", on=1, off=0, default=0),
    ]

    defender_pol = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows Defender"
    wu_au = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    visualfx = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
    power_fast = r"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    power_hiber = r"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power"
    graphics = r"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\GraphicsDrivers"
    gamedvr = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\GameDVR"
    mm_sysprofile = r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile"
    filesystem = r"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem"
    mem_mgmt = r"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
    desktop_cp = r"HKEY_CURRENT_USER\Control Panel\Desktop"
    mouse_cp = r"HKEY_CURRENT_USER\Control Panel\Mouse"
    settings_schema += [
        D("Security", "Disable Windows Defender (may be ignored on recent Windows)", "Legacy policy; modern Windows may ignore it.",
          defender_pol, "DisableAntiSpyware", "defender_disable", page="System & Performance", on=1, off=0, default=0, requires="reboot"),
        D("Updates", "Disable Automatic Windows Updates", "Stops automatic Windows Updates.",
          wu_au, "NoAutoUpdate", "wu_disable", page="System & Performance", on=1, off=0, default=0, requires="reboot"),
        D("Apps", "Disable Background Apps", "Prevents apps from running in the background.",
          r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications", "GlobalUserDisabled",
          "bg_apps", page="System & Performance", on=1, off=0, default=0),
        D("Boot", "Disable Fast Startup", "Can help with dual-boot or driver issues. Requires restart.",
          power_fast, "HiberbootEnabled", "fast_startup", page="System & Performance", on=0, off=1, default=1, requires="reboot"),
        D("Power", "Enable Hibernation", "Required for Fast Startup; disabling frees disk space.",
          power_hiber, "HibernateEnabled", "hibernate", page="System & Performance", on=1, off=0, default=1, requires="reboot"),
        D("Graphics", "Enable Hardware-Accelerated GPU Scheduling", "Requires supported GPU/driver and restart.",
          graphics, "HwSchMode", "hws", page="System & Performance", on=2, off=1, default=1, requires="reboot"),
        D("Gaming", "Disable Game DVR (Recording)", "Disables background recording to reduce overhead.",
          gamedvr, "AllowGameDVR", "game_dvr", page="System & Performance", on=0, off=1, default=1),
        D("Visuals", "Visual Effects", "Switch between performance vs appearance presets.",
          visualfx, "VisualFXSetting", "visualfx", page="System & Performance", s_type="combo",
          choices=[("Let Windows Decide (1)", 1), ("Best Appearance (3)", 3), ("Best Performance (2)", 2)], default=1),
        D("Network", "Network Throttling", "Disabling may help with some latency-sensitive workloads.",
          mm_sysprofile, "NetworkThrottlingIndex", "net_throttle", page="System & Performance", s_type="combo",
          choices=[("Default (10)", 10), ("Disabled (0xFFFFFFFF)", 0xFFFFFFFF)], default=10, requires="reboot"),
        D("CPU Scheduling", "System Responsiveness", "Lower dedicates more CPU to foreground apps.",
          mm_sysprofile, "SystemResponsiveness", "sys_responsiveness", page="System & Performance", s_type="combo",
          choices=[("Gaming (10)", 10), ("Default (20)", 20), ("Multimedia (75)", 75)], default=20),
        D("NTFS", "Disable Last Access Time", "Prevents updating last access timestamps on files (performance).",
          filesystem, "NtfsDisableLastAccessUpdate", "ntfs_last_access", page="System & Performance", on=1, off=0, default=1, requires="reboot"),
        D("NTFS", "Disable 8.3 Name Creation", "Disables short 8.3 names on NTFS; speeds up file ops.",
          filesystem, "NtfsDisable8dot3NameCreation", "ntfs_83", page="System & Performance", on=1, off=0, default=2, requires="reboot"),
        D("Memory", "Disable Paging Executive (Advanced)", "Keeps kernel/drivers in RAM (handle with care).",
          mem_mgmt, "DisablePagingExecutive", "disable_paging_exec", page="System & Performance", on=1, off=0, default=0, requires="reboot"),
        D("Menus", "Menu Show Delay (ms)", "Lower feels snappier; default ~400ms.",
          desktop_cp, "MenuShowDelay", "menu_delay", page="System & Performance", reg_type="sz", s_type="spin",
          spin={"min": 0, "max": 2000, "step": 10}, default="400"),
        D("Mouse", "Mouse Hover Time (ms)", "Delay before hover events; default ~400ms.",
          mouse_cp, "MouseHoverTime", "mouse_hover", page="System & Performance", reg_type="sz", s_type="spin",
          spin={"min": 0, "max": 2000, "step": 10}, default="400"),
    ]

    class SettingItem:
        def __init__(self, schema: Dict[str, Any]):
            self.s = schema
            self.widget: Optional[QWidget] = None
            self.row: Optional[QWidget] = None
            self.page_name: str = schema["page"]
            self.page_widget = None
            self.original = None
            self.current = None
            self.pending: Optional[Any] = None
            self.badge: Optional[QLabel] = None
            self.reset_btn: Optional[QToolButton] = None
            self.star_btn: Optional[QToolButton] = None

        def is_supported(self) -> bool:
            mb = self.s.get("minBuild")
            return True if not mb else (WIN_BUILD >= int(mb))

    class SettingsPageWidget(QScrollArea):
        def __init__(self, title: str, compact: bool):
            super().__init__()
            self.setWidgetResizable(True)
            self.container = QWidget(self)
            self.setWidget(self.container)
            self.layout = QVBoxLayout(self.container)
            self.layout.setContentsMargins(16, 14, 16, 14 if not compact else 6)
            self.layout.setSpacing(10 if not compact else 6)
            title_label = QLabel(title, self.container)
            title_label.setStyleSheet("font-size: 20px; font-weight: 600; margin-bottom: 6px;")
            self.layout.addWidget(title_label)
            self.actions_bar = QHBoxLayout()
            self.actions_bar.setSpacing(8)
            self.btn_reset_page = QPushButton("Reset Page")
            self.btn_reset_page.setToolTip("Reset all settings on this page to their defaults (pending until Apply).")
            self.actions_bar.addWidget(self.btn_reset_page)
            self.actions_bar.addStretch(1)
            self.layout.addLayout(self.actions_bar)
            self.sections: Dict[str, CollapsibleSection] = {}

        def get_or_create_section(self, name: str, compact: bool):
            if name not in self.sections:
                cs = CollapsibleSection(name, self.container, compact=compact)
                self.sections[name] = cs
                self.layout.addWidget(cs)
            return self.sections[name]

        def finalize(self):
            spacer = QWidget(self.container)
            spacer.setMinimumHeight(12)
            self.layout.addWidget(spacer)
            self.layout.addStretch(1)

    class PreviewDialog(QDialog):
        def __init__(self, changes: List[Tuple[object, Any, Any]], parent=None):
            super().__init__(parent)
            self.setWindowTitle("Preview Changes")
            self.resize(720, 420)
            lay = QVBoxLayout(self)
            txt = QPlainTextEdit(self)
            txt.setReadOnly(True)
            lines = []
            for item, old, new in changes:
                s = item.s
                line = f"{s['path']}\\{s['name']}  [{s['reg_type']}]  {old}  ->  {new}"
                req = human_requires(s.get("requires"))
                if req:
                    line += f"  ({req})"
                lines.append(line)
            txt.setPlainText("\n".join(lines) if lines else "No pending changes.")
            lay.addWidget(txt)
            btn = QPushButton("Close", self)
            btn.clicked.connect(self.accept)
            lay.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)

    class MainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle(APP_NAME)
            self.resize(1280, 820)
            self.compact = bool(cfg.get("compact", False))
            self.theme = cfg.get("theme", "dark")
            self.restore_point_enabled = bool(cfg.get("restore_point", False))
            self.favorites = set(cfg.get("favorites", []))
            self.toast = Toast(self)
            self.items: Dict[str, SettingItem] = {}
            self.pages_widgets: Dict[str, SettingsPageWidget] = {}
            self.nav_items: Dict[str, int] = {}
            self.pending_map: Dict[str, Any] = {}
            self.undo_stack: List[Tuple[str, Any, Any]] = []
            self._build_menu()
            splitter = QSplitter(self)
            self.setCentralWidget(splitter)
            sidebar = QWidget(self)
            s_lay = QVBoxLayout(sidebar)
            s_lay.setContentsMargins(10, 10, 10, 10)
            s_lay.setSpacing(8)
            self.search_edit = QLineEdit(sidebar)
            self.search_edit.setPlaceholderText("Search settings...")
            self.search_edit.setClearButtonEnabled(True)
            s_lay.addWidget(self.search_edit)
            self.nav_list = QListWidget(sidebar)
            self.nav_list.setUniformItemSizes(True)
            self.nav_list.setIconSize(QSize(18, 18))
            self.nav_list.setStyleSheet("""
                QListWidget { border: none; outline: 0; font-size: 14px; }
                QListWidget::item { padding: 10px 10px; margin: 2px 0; border-radius: 6px; }
                QListWidget::item:selected { background: rgba(127,127,127,0.25); }
            """)
            s_lay.addWidget(self.nav_list, 1)
            btn_bar = QHBoxLayout()
            btn_bar.setSpacing(6)
            self.btn_preview = QPushButton("Preview")
            self.btn_revert = QPushButton("Revert All")
            self.btn_apply = QPushButton("Apply (0) & Restart Explorer")
            self.btn_apply.setEnabled(False)
            btn_bar.addWidget(self.btn_preview)
            btn_bar.addWidget(self.btn_revert)
            btn_bar.addWidget(self.btn_apply, 1)
            s_lay.addLayout(btn_bar)
            self.stack = QStackedWidget(self)
            splitter.addWidget(sidebar)
            splitter.addWidget(self.stack)
            splitter.setStretchFactor(1, 1)
            splitter.setSizes([300, 900])
            status = QStatusBar(self)
            self.setStatusBar(status)
            status.showMessage("Ready")
            pages_order = ["Personalization", "Taskbar & Start", "File Explorer & UI", "Privacy", "System & Performance", "Favorites", "Search Results", "Logs"]
            self._setup_pages(pages_order)
            for s in settings_schema:
                if s.get("minBuild") and WIN_BUILD < int(s["minBuild"]):
                    continue
                item = SettingItem(s)
                self.items[s["id"]] = item
                page = self.pages_widgets[s["page"]]
                section = page.get_or_create_section(s["section"], self.compact)
                row = self._create_row_widget(item)
                item.row = row
                item.page_widget = page
                section.add_row(row)
            self._rebuild_favorites_page()
            self._build_logs_page()
            for w in self.pages_widgets.values():
                w.finalize()
                w.btn_reset_page.clicked.connect(lambda _, pg=w: self._reset_page(pg))
            self.nav_list.currentRowChanged.connect(self.stack.setCurrentIndex)
            self.search_edit.textChanged.connect(self._on_search_text)
            self.btn_apply.clicked.connect(self._apply_changes)
            self.btn_revert.clicked.connect(self._revert_all)
            self.btn_preview.clicked.connect(self._preview_changes)
            self._bind_shortcuts()
            self.refresh_from_registry()
            self.sync_timer = QTimer(self)
            self.sync_timer.timeout.connect(self._poll_registry_changes)
            self.sync_timer.start(4000)

        def _build_menu(self):
            m = self.menuBar()
            file_menu = m.addMenu("&File")
            self.act_save_profile = QAction("Save Profile (JSON)...", self)
            self.act_load_profile = QAction("Load Profile (JSON)...", self)
            self.act_export_reg = QAction("Export Pending as .reg...", self)
            self.act_quit = QAction("Exit", self)
            file_menu.addAction(self.act_save_profile)
            file_menu.addAction(self.act_load_profile)
            file_menu.addSeparator()
            file_menu.addAction(self.act_export_reg)
            file_menu.addSeparator()
            file_menu.addAction(self.act_quit)
            self.act_save_profile.triggered.connect(self._save_profile)
            self.act_load_profile.triggered.connect(self._load_profile)
            self.act_export_reg.triggered.connect(self._export_pending_reg)
            self.act_quit.triggered.connect(self.close)

            view_menu = m.addMenu("&View")
            self.act_theme_dark = QAction("Dark Theme", self, checkable=True)
            self.act_theme_light = QAction("Light Theme", self, checkable=True)
            ag = QActionGroup(self)
            ag.setExclusive(True)
            ag.addAction(self.act_theme_dark)
            ag.addAction(self.act_theme_light)
            self.act_theme_dark.setChecked(cfg.get("theme", "dark") == "dark")
            self.act_theme_light.setChecked(cfg.get("theme", "dark") == "light")
            view_menu.addAction(self.act_theme_dark)
            view_menu.addAction(self.act_theme_light)
            view_menu.addSeparator()
            self.act_compact = QAction("Compact Density", self, checkable=True)
            self.act_compact.setChecked(cfg.get("compact", False))
            view_menu.addAction(self.act_compact)
            self.act_theme_dark.toggled.connect(lambda checked: self._set_theme("dark") if checked else None)
            self.act_theme_light.toggled.connect(lambda checked: self._set_theme("light") if checked else None)
            self.act_compact.toggled.connect(self._set_compact)

            tools_menu = m.addMenu("&Tools")
            self.act_restore_point = QAction("Create restore point before Apply", self, checkable=True)
            self.act_restore_point.setChecked(bool(cfg.get("restore_point", False)))
            tools_menu.addAction(self.act_restore_point)
            self.act_restore_point.toggled.connect(self._set_restore_point)

            help_menu = m.addMenu("&Help")
            act_about = QAction("About", self)
            act_open_log = QAction("Open Log File", self)
            help_menu.addAction(act_open_log)
            help_menu.addAction(act_about)
            act_open_log.triggered.connect(self._open_log_file)
            act_about.triggered.connect(self._about)

        def _set_theme(self, theme: str):
            self.theme = theme
            cfg["theme"] = theme
            save_config(cfg)
            apply_theme(app, theme)
            self.act_theme_dark.blockSignals(True)
            self.act_theme_light.blockSignals(True)
            self.act_theme_dark.setChecked(theme == "dark")
            self.act_theme_light.setChecked(theme == "light")
            self.act_theme_dark.blockSignals(False)
            self.act_theme_light.blockSignals(False)
            for w in app.topLevelWidgets():
                w.update()
            self.toast.show_toast(f"Theme set to {theme.capitalize()}")

        def _set_compact(self, enabled: bool):
            self.compact = enabled
            cfg["compact"] = enabled
            save_config(cfg)
            self.toast.show_toast("Compact density enabled" if enabled else "Compact density disabled")

        def _set_restore_point(self, enabled: bool):
            self.restore_point_enabled = enabled
            cfg["restore_point"] = enabled
            save_config(cfg)

        def _setup_pages(self, pages_order: List[str]):
            while self.stack.count():
                w = self.stack.widget(0)
                self.stack.removeWidget(w)
                w.deleteLater()
            self.pages_widgets.clear()
            self.nav_list.clear()
            style = self.style()
            icons = {
                "Personalization": style.standardIcon(QStyle.StandardPixmap.SP_DesktopIcon),
                "Taskbar & Start": style.standardIcon(QStyle.StandardPixmap.SP_ComputerIcon),
                "File Explorer & UI": style.standardIcon(QStyle.StandardPixmap.SP_DirHomeIcon),
                "Privacy": style.standardIcon(QStyle.StandardPixmap.SP_BrowserStop),
                "System & Performance": style.standardIcon(QStyle.StandardPixmap.SP_DriveHDIcon),
                "Favorites": style.standardIcon(QStyle.StandardPixmap.SP_DirIcon),
                "Search Results": style.standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView),
                "Logs": style.standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView),
            }
            for p in pages_order:
                w = SettingsPageWidget(p, compact=self.compact)
                self.pages_widgets[p] = w
                self.stack.addWidget(w)
                item = QListWidgetItem(icons.get(p), p, self.nav_list)
                self.nav_items[p] = self.nav_list.row(item)
            self.nav_list.setCurrentRow(self.nav_items["Personalization"])

        def _create_row_widget(self, item: SettingItem) -> QWidget:
            s = item.s
            row = QWidget()
            h = QHBoxLayout(row)
            h.setContentsMargins(8, 4, 8, 4)
            h.setSpacing(10 if not self.compact else 6)
            label = QLabel(s["label"])
            label.setWordWrap(True)
            if s.get("tooltip"):
                label.setToolTip(s["tooltip"])
            badges = QHBoxLayout()
            badges.setSpacing(6)
            badge_container = QWidget()
            badge_container.setLayout(badges)
            req = human_requires(s.get("requires"))
            if req:
                b = badge_label(req)
                badges.addWidget(b)
                item.badge = b
            h.addWidget(label, 4)
            h.addWidget(badge_container, 0, Qt.AlignmentFlag.AlignLeft)
            h.addStretch(1)
            ctrl: Optional[QWidget] = None
            if s["type"] == "switch":
                ctrl = QCheckBox()
                if s.get("tooltip"):
                    ctrl.setToolTip(s["tooltip"])
                ctrl.toggled.connect(lambda checked, it=item: self._on_control_changed(it))
            elif s["type"] == "combo":
                cb = QComboBox()
                for text, val in s["choices"]:
                    cb.addItem(text, val)
                if s.get("tooltip"):
                    cb.setToolTip(s["tooltip"])
                cb.currentIndexChanged.connect(lambda _, it=item: self._on_control_changed(it))
                ctrl = cb
            elif s["type"] == "spin":
                sp = QSpinBox()
                sp.setRange(int(s["spin"]["min"]), int(s["spin"]["max"]))
                sp.setSingleStep(int(s["spin"]["step"]))
                sp.setSuffix(" ms")
                if s.get("tooltip"):
                    sp.setToolTip(s["tooltip"])
                sp.valueChanged.connect(lambda _, it=item: self._on_control_changed(it))
                ctrl = sp
            else:
                ctrl = QLabel("Unsupported")
            item.widget = ctrl
            h.addWidget(ctrl, 0, Qt.AlignmentFlag.AlignRight)
            reset_btn = QToolButton()
            reset_btn.setText("Reset")
            reset_btn.setToolTip("Reset to default (pending until Apply)")
            reset_btn.clicked.connect(lambda _, it=item: self._reset_setting(it))
            item.reset_btn = reset_btn
            h.addWidget(reset_btn, 0)
            fav_btn = QToolButton()
            fav_btn.setText("★" if s["id"] in self.favorites else "☆")
            fav_btn.setToolTip("Toggle favorite")
            fav_btn.clicked.connect(lambda _, it=item, b=fav_btn: self._toggle_favorite(it, b))
            item.star_btn = fav_btn
            h.addWidget(fav_btn, 0)
            row.setProperty("setting_id", s["id"])
            row.setToolTip(s.get("tooltip", ""))
            self._update_row_dirty_style(item, dirty=False)
            return row

        def _update_row_dirty_style(self, item: SettingItem, dirty: bool):
            if not item.row:
                return
            if dirty:
                item.row.setStyleSheet("QWidget { background: rgba(255, 196, 0, 0.12); border-radius: 6px; }")
            else:
                item.row.setStyleSheet("")

        def _widget_to_regvalue(self, item: SettingItem):
            s = item.s
            if s["type"] == "switch":
                checked = bool(item.widget.isChecked())
                return s["on"] if checked else s["off"]
            elif s["type"] == "combo":
                cb: QComboBox = item.widget
                return cb.currentData()
            elif s["type"] == "spin":
                sp: QSpinBox = item.widget
                return str(sp.value()) if s["reg_type"] == "sz" else sp.value()
            return None

        def _regvalue_to_widget(self, item: SettingItem, val):
            s = item.s
            if s["type"] == "switch":
                checked = (val == s["on"])
                item.widget.blockSignals(True)
                item.widget.setChecked(checked)
                item.widget.blockSignals(False)
            elif s["type"] == "combo":
                cb: QComboBox = item.widget
                idx = 0
                for i in range(cb.count()):
                    if cb.itemData(i) == val:
                        idx = i
                        break
                cb.blockSignals(True)
                cb.setCurrentIndex(idx)
                cb.blockSignals(False)
            elif s["type"] == "spin":
                sp: QSpinBox = item.widget
                try:
                    ival = int(val)
                except Exception:
                    ival = int(s.get("default") or 0)
                sp.blockSignals(True)
                sp.setValue(ival)
                sp.blockSignals(False)

        def refresh_from_registry(self):
            for item in self.items.values():
                if not item.is_supported():
                    if item.row:
                        item.row.setVisible(False)
                    continue
                val, _ = RegistryManager.read_value(item.s["path"], item.s["name"], item.s.get("default"))
                if val is None:
                    val = item.s.get("default")
                if item.s["reg_type"] == "dword" and isinstance(val, str):
                    try:
                        val = int(val, 0)
                    except Exception:
                        try:
                            val = int(val)
                        except Exception:
                            pass
                item.original = val
                if item.s["id"] not in self.pending_map:
                    item.current = val
                    self._regvalue_to_widget(item, val)
                    self._update_row_dirty_style(item, dirty=False)
            self._update_apply_button()

        def _poll_registry_changes(self):
            changed = 0
            for item in self.items.values():
                if item.s["id"] in self.pending_map:
                    continue
                new_val, _ = RegistryManager.read_value(item.s["path"], item.s["name"], item.s.get("default"))
                if new_val != item.current:
                    item.current = new_val
                    item.original = new_val
                    self._regvalue_to_widget(item, new_val)
                    changed += 1
                    logging.info("External change detected: %s -> %s", item.s["id"], str(new_val))
            if changed:
                self.statusBar().showMessage(f"Detected {changed} external change(s). UI synced.", 3000)

        def _on_control_changed(self, item: SettingItem):
            new_val = self._widget_to_regvalue(item)
            old_val = item.original
            item.current = new_val
            sid = item.s["id"]
            if new_val == old_val:
                if sid in self.pending_map:
                    del self.pending_map[sid]
                self._update_row_dirty_style(item, dirty=False)
            else:
                self.pending_map[sid] = new_val
                self._update_row_dirty_style(item, dirty=True)
            self.undo_stack.append((sid, old_val, new_val))
            self._update_apply_button()

        def _reset_setting(self, item: SettingItem):
            default = item.s.get("default")
            self._regvalue_to_widget(item, default)
            self._on_control_changed(item)

        def _reset_page(self, page_widget: SettingsPageWidget):
            for item in self.items.values():
                if item.page_widget is page_widget:
                    self._reset_setting(item)
            self.toast.show_toast("Page reset to defaults (pending)")

        def _toggle_favorite(self, item: SettingItem, btn: QToolButton):
            sid = item.s["id"]
            if sid in self.favorites:
                self.favorites.remove(sid)
                btn.setText("☆")
            else:
                self.favorites.add(sid)
                btn.setText("★")
            cfg["favorites"] = list(self.favorites)
            save_config(cfg)
            self._rebuild_favorites_page()

        def _rebuild_favorites_page(self):
            pg = self.pages_widgets["Favorites"]
            for sec in list(pg.sections.values()):
                sec.setParent(None)
            pg.sections.clear()
            sec = pg.get_or_create_section("Pinned Settings", self.compact)
            if not self.favorites:
                lbl = QLabel("No favorites yet. Click ☆ next to a setting to pin it.")
                sec.add_row(lbl)
            else:
                for sid in self.favorites:
                    if sid in self.items:
                        item = self.items[sid]
                        btn = QPushButton(f"Open: {item.s['page']} • {item.s['section']} • {item.s['label']}")
                        btn.setToolTip(item.s.get("tooltip", ""))
                        btn.clicked.connect(lambda _, it=item: self._jump_to_setting(it))
                        sec.add_row(btn)
            pg.finalize()

        def _jump_to_setting(self, item: SettingItem):
            page = item.s["page"]
            self.nav_list.setCurrentRow(self.nav_items[page])
            page_widget = self.pages_widgets[page]
            page_widget.ensureWidgetVisible(item.row)
            self._flash_row(item)

        def _flash_row(self, item: SettingItem):
            if not item.row:
                return
            item.row.setStyleSheet("QWidget { background: rgba(255, 255, 0, 0.25); border-radius: 6px; }")
            QTimer.singleShot(1400, lambda: self._update_row_dirty_style(item, item.s["id"] in self.pending_map))

        def _create_restore_point(self) -> bool:
            try:
                desc = f"{APP_NAME} {time.strftime('%Y-%m-%d %H:%M:%S')}"
                cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Checkpoint-Computer -Description \'{0}\' -RestorePointType \'MODIFY_SETTINGS\'"'.format(desc)
                subprocess.check_call(cmd, shell=True)
                logging.info("System restore point created: %s", desc)
                return True
            except Exception:
                logging.exception("Failed to create system restore point")
                return False

        def _export_backup_reg(self, changes: List[Tuple[object, Any, Any]]):
            try:
                lines = ["Windows Registry Editor Version 5.00", ""]
                grouped: Dict[str, List[Tuple[str, Any, str]]] = {}
                for item, old, _ in changes:
                    grouped.setdefault(item.s["path"], []).append((item.s["name"], old, item.s["reg_type"]))
                for path, entries in grouped.items():
                    lines.append(f"[{path}]")
                    for name, val, typ in entries:
                        if typ == "dword":
                            if val is None:
                                continue
                            lines.append(f'"{name}"=dword:{int(val) & 0xFFFFFFFF:08x}')
                        else:
                            sval = "" if val is None else str(val)
                            lines.append(f'"{name}"="{sval}"')
                    lines.append("")
                with open(BACKUP_REG_PATH, "w", encoding="utf-16le") as f:
                    text = "\ufeff" + "\n".join(lines)
                    f.write(text)
                logging.info("Backup .reg saved to %s", BACKUP_REG_PATH)
            except Exception:
                logging.exception("Failed to export backup .reg")

        def _apply_changes(self):
            if not self.pending_map:
                return
            changes: List[Tuple[object, Any, Any]] = []
            for sid, new_val in self.pending_map.items():
                item = self.items[sid]
                changes.append((item, item.original, new_val))
            if self.restore_point_enabled:
                self.statusBar().showMessage("Creating system restore point...", 5000)
                ok = self._create_restore_point()
                self.toast.show_toast("Restore point created" if ok else "Failed to create restore point", 3000)
            self._export_backup_reg(changes)
            failures = 0
            for item, old, new in changes:
                reg_type = winreg.REG_DWORD if item.s["reg_type"] == "dword" else winreg.REG_SZ
                ok = RegistryManager.write_value(item.s["path"], item.s["name"], new, reg_type)
                if not ok:
                    failures += 1
                else:
                    item.original = new
                    item.current = new
                    self._regvalue_to_widget(item, new)
                    self._update_row_dirty_style(item, False)
            self.pending_map.clear()
            self._update_apply_button()
            if failures:
                self.toast.show_toast(f"Applied with {failures} failure(s). See log.", 5000)
            else:
                self.toast.show_toast("Changes applied successfully.", 2500)
            restart_explorer()
            self.statusBar().showMessage("Explorer restarted", 3000)

        def _revert_all(self):
            self.refresh_from_registry()
            self.pending_map.clear()
            self._update_apply_button()
            self.toast.show_toast("All pending changes discarded")

        def _preview_changes(self):
            changes: List[Tuple[object, Any, Any]] = []
            for sid, new_val in self.pending_map.items():
                item = self.items[sid]
                changes.append((item, item.original, new_val))
            dlg = PreviewDialog(changes, self)
            dlg.exec()

        def _update_apply_button(self):
            n = len(self.pending_map)
            self.btn_apply.setText(f"Apply ({n}) & Restart Explorer")
            self.btn_apply.setEnabled(n > 0)

        def _on_search_text(self, txt: str):
            txt = (txt or "").strip().lower()
            if not txt:
                return
            pg = self.pages_widgets["Search Results"]
            for sec in list(pg.sections.values()):
                sec.setParent(None)
            pg.sections.clear()
            sec = pg.get_or_create_section(f"Results for \"{txt}\"", self.compact)
            results = []
            for it in self.items.values():
                hay = " ".join([it.s["label"], it.s.get("tooltip", ""), it.s["page"], it.s["section"], it.s["id"]]).lower()
                if txt in hay:
                    results.append(it)
            if not results:
                sec.add_row(QLabel("No matches."))
            else:
                for it in results:
                    btn = QPushButton(f"{it.s['page']} • {it.s['section']} • {it.s['label']}")
                    btn.setToolTip(it.s.get("tooltip", ""))
                    btn.clicked.connect(lambda _, ii=it: self._jump_to_setting(ii))
                    sec.add_row(btn)
            pg.finalize()
            self.nav_list.setCurrentRow(self.nav_items["Search Results"])

        def _save_profile(self):
            path, _ = QFileDialog.getSaveFileName(self, "Save Profile", "", "JSON Files (*.json)")
            if not path:
                return
            data = {sid: self._widget_to_regvalue(self.items[sid]) for sid in self.items}
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2)
                self.toast.show_toast("Profile saved")
            except Exception:
                logging.exception("Failed to save profile")
                self.toast.show_toast("Failed to save profile", 4000)

        def _load_profile(self):
            path, _ = QFileDialog.getOpenFileName(self, "Load Profile", "", "JSON Files (*.json)")
            if not path:
                return
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                applied = 0
                for sid, val in data.items():
                    if sid in self.items:
                        it = self.items[sid]
                        self._regvalue_to_widget(it, val)
                        self._on_control_changed(it)
                        applied += 1
                self.toast.show_toast(f"Profile loaded ({applied} settings applied; pending)")
            except Exception:
                logging.exception("Failed to load profile")
                self.toast.show_toast("Failed to load profile", 4000)

        def _export_pending_reg(self):
            if not self.pending_map:
                self.toast.show_toast("No pending changes to export")
                return
            path, _ = QFileDialog.getSaveFileName(self, "Export Pending as .reg", "", "Registry Files (*.reg)")
            if not path:
                return
            try:
                lines = ["Windows Registry Editor Version 5.00", ""]
                grouped: Dict[str, List[Tuple[str, Any, str]]] = {}
                for sid, new_val in self.pending_map.items():
                    it = self.items[sid]
                    grouped.setdefault(it.s["path"], []).append((it.s["name"], new_val, it.s["reg_type"]))
                for pth, entries in grouped.items():
                    lines.append(f"[{pth}]")
                    for name, val, typ in entries:
                        if typ == "dword":
                            lines.append(f'"{name}"=dword:{int(val) & 0xFFFFFFFF:08x}')
                        else:
                            sval = "" if val is None else str(val)
                            lines.append(f'"{name}"="{sval}"')
                    lines.append("")
                with open(path, "w", encoding="utf-16le") as f:
                    text = "\ufeff" + "\n".join(lines)
                    f.write(text)
                self.toast.show_toast("Exported .reg")
            except Exception:
                logging.exception("Failed to export .reg")
                self.toast.show_toast("Failed to export .reg", 4000)

        def _build_logs_page(self):
            pg = self.pages_widgets["Logs"]
            sec = pg.get_or_create_section("Application Log", self.compact)
            self.log_view = QPlainTextEdit()
            self.log_view.setReadOnly(True)
            btns = QHBoxLayout()
            btn_refresh = QPushButton("Refresh")
            btn_copy = QPushButton("Copy All")
            btns.addWidget(btn_refresh)
            btns.addWidget(btn_copy)
            btns.addStretch(1)
            sec.add_row(self.log_view)
            sec.add_row(QWidget())
            sec.vlay.addLayout(btns)
            btn_refresh.clicked.connect(self._refresh_log_view)
            btn_copy.clicked.connect(lambda: (self.log_view.selectAll(), self.log_view.copy(), self.log_view.moveCursor(self.log_view.textCursor().End)))
            self._refresh_log_view()

        def _refresh_log_view(self):
            try:
                if os.path.exists(LOG_PATH):
                    with open(LOG_PATH, "r", encoding="utf-8") as f:
                        self.log_view.setPlainText(f.read())
                else:
                    self.log_view.setPlainText("(no log yet)")
            except Exception:
                self.log_view.setPlainText("(failed to read log)")

        def _open_log_file(self):
            try:
                os.startfile(LOG_PATH)
            except Exception:
                logging.exception("Failed to open log")
                self.toast.show_toast("Failed to open log", 3000)

        def _about(self):
            QMessageBox.information(self, "About expDesigner",
                "expDesigner — Windows customization utility\n"
                "single-file, PyQt6. favorites ☆, search, profiles, backups.\n"
                "Made by VoidWither on GitHub.\n"
                "Feel free to use this code on your own projects!"
            )

        def _bind_shortcuts(self):
            act_apply = QAction(self)
            act_apply.setShortcut("Ctrl+S")
            act_apply.triggered.connect(self._apply_changes)
            self.addAction(act_apply)
            act_undo = QAction(self)
            act_undo.setShortcut("Ctrl+Z")
            act_undo.triggered.connect(self._undo_last)
            self.addAction(act_undo)
            act_search = QAction(self)
            act_search.setShortcut("Ctrl+F")
            act_search.triggered.connect(lambda: self.search_edit.setFocus(Qt.FocusReason.ShortcutFocusReason))
            self.addAction(act_search)

        def _undo_last(self):
            if not self.undo_stack:
                return
            sid, old, new = self.undo_stack.pop()
            if sid in self.items:
                it = self.items[sid]
                self._regvalue_to_widget(it, old)
                self._on_control_changed(it)
                self.toast.show_toast(f"Undid change: {it.s['label']}")

    window = MainWindow()
    window.show()
    logging.info("Application UI initialized and shown")
    return app.exec()

if __name__ == "__main__":
    if not is_running_as_admin():
        if relaunch_as_admin():
            sys.exit(0)
        else:
            print("Unable to obtain administrator privileges. Please run as administrator.")
            sys.exit(1)
    try:
        code = main_app()
        sys.exit(code)
    except Exception:
        logging.exception("Fatal error while running the application")
        raise
