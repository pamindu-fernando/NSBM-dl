#!/usr/bin/env python3
"""
Moodle File Downloader — NSBM Edition
Downloads PDF, XLSX, DOC, DOCX, TXT, PPT, PPTX files from Moodle courses.
"""

import sys
import os
import re
import time
import threading
from urllib.parse import urljoin, urlparse, unquote

import requests
from bs4 import BeautifulSoup
import email
import email.message
import email.mime.multipart
import email.mime.text
import pkg_resources

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QTextEdit, 
    QFileDialog, QMessageBox, QFrame, QTreeWidget, QTreeWidgetItem
)
from PyQt6.QtCore import Qt, pyqtSignal, QPoint, QRectF, QUrl
from PyQt6.QtGui import QFont, QCursor, QIcon, QPainter, QPainterPath, QColor, QBrush, QMouseEvent, QDesktopServices

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ─── Themes (Photoshop style) ──────────────────────────────────────────────────
# ─── Themes (Photoshop style) ──────────────────────────────────────────────────
THEMES = {
    'dark': {
        'NAME': 'dark',
        'BG_APP': '#2f2f2f',
        'BG_CARD': '#383838',
        'BG_INPUT': '#262626',
        'BORDER': '#555555',
        'TEXT_MAIN': '#e0e0e0',
        'TEXT_MUTED': '#9d9d9d',
        'ACCENT': '#1b6bcb',  
        'ACCENT_HOV': '#0054b6',
        'ICON': '☀️'  # <--- Changed to simplistic sun
    },
    'light': {
        'NAME': 'light',
        'BG_APP': '#f0f0f0',
        'BG_CARD': '#e4e4e4',
        'BG_INPUT': '#ffffff',
        'BORDER': '#c8c8c8',
        'TEXT_MAIN': '#2d2d2d',
        'TEXT_MUTED': '#6b6b6b',
        'ACCENT': '#0065cc',
        'ACCENT_HOV': '#004f9e',
        'ICON': '🌙'  # <--- Changed to simplistic moon
    }
}

ALLOWED_EXTENSIONS = {'.pdf', '.xlsx', '.doc', '.docx', '.txt', '.ppt', '.pptx'}

def get_stylesheet(theme_key):
    t = THEMES[theme_key]
    return f"""
QWidget {{
    color: {t['TEXT_MAIN']};
    font-family: "-apple-system", "Segoe UI", "Helvetica Neue", Helvetica, Arial, sans-serif;
    font-size: 9pt;
}}

/* The actual window background with curved edges is painted manually in paintEvent, but children use BG_APP */

QFrame#Header {{
    background-color: transparent;
    border-bottom: 2px solid {t['BORDER']};
}}

QFrame#Card {{
    background-color: {t['BG_CARD']};
    border: 1px solid {t['BORDER']};
    border-radius: 4px;
}}

QLabel {{
    background-color: transparent;
    border: none;
}}

QLabel#HeaderTitle {{
    font-size: 17pt;           /* Increased from 13pt */
    font-weight: 800;          /* Extra bold */
    color: #22c55e;            /* A vibrant, modern green */
    font-family: "Verdana", "Trebuchet MS", sans-serif; /* Clean, bold, terminal-style font */
}}

QLabel#HeaderSub {{
    font-size: 8pt;
    color: {t['TEXT_MUTED']};
}}

QLabel#AboutLink {{
    font-size: 9pt;
    font-weight: bold;
    padding-right: 15px;
}}
/* Style HTML link directly inside QLabel via rich text CSS subset */
QLabel#AboutLink a {{
    color: {t['ACCENT']};
    text-decoration: none;
}}

QLabel#SectionTitle {{
    font-size: 10pt;
    font-weight: 600;
    color: {t['TEXT_MAIN']};
    background-color: transparent;
    border: none;
}}

QLabel#FieldLabel {{
    color: {t['TEXT_MUTED']};
    font-size: 9pt;
}}

QLineEdit, QTextEdit {{
    background-color: {t['BG_INPUT']};
    border: 1px solid {t['BORDER']};
    border-radius: 3px;
    padding: 4px 6px;
    color: {t['TEXT_MAIN']};
    selection-background-color: {t['ACCENT']};
}}

QLineEdit:focus, QTextEdit:focus {{
    border: 1px solid {t['ACCENT']};
}}

QTreeView {{
    background-color: {t['BG_INPUT']};
    border: 1px solid {t['BORDER']};
    border-radius: 3px;
    color: {t['TEXT_MAIN']};
    outline: none;
}}

QTreeView::item {{
    padding: 4px;
}}

QTreeView::item:hover {{
    background-color: {t['BG_CARD']};
}}

QPushButton {{
    background-color: {t['BG_APP']};
    border: 1px solid {t['BORDER']};
    border-radius: 3px;
    padding: 5px 12px;
    color: {t['TEXT_MAIN']};
    font-size: 9pt;
}}

QPushButton:hover {{
    background-color: {t['BORDER']};
}}

QPushButton:pressed {{
    background-color: {t['BG_INPUT']};
}}

QPushButton:disabled {{
    color: {t['TEXT_MUTED']};
    background-color: {t['BG_CARD']};
    border: 1px solid {t['BG_CARD']};
}}

QPushButton#PrimaryBtn {{
    background-color: {t['ACCENT']};
    border: 1px solid {t['ACCENT_HOV']};
    color: white;
    font-weight: bold;
    padding: 6px;
}}

QPushButton#PrimaryBtn:hover {{
    background-color: {t['ACCENT_HOV']};
}}

QPushButton#PrimaryBtn:disabled {{
    background-color: {t['BG_CARD']};
    border: 1px solid {t['BORDER']};
    color: {t['TEXT_MUTED']};
}}

QPushButton#ThemeBtn {{
    background-color: transparent;
    border: none;
    font-size: 14pt; /* Slightly smaller to fit standard emojis */
    font-family: "Segoe UI Emoji", "-apple-system", "Segoe UI", sans-serif; /* Added Emoji font support */
    padding: 0px;
    margin: 0px;
}}

QPushButton#ThemeBtn:hover {{
    background-color: {t['BG_CARD']};
    border-radius: 16px;
}}

QPushButton#EyeBtn {{
    font-size: 8pt;
    font-weight: bold;
    padding: 2px 4px;
}}

QPushButton#WinBtn {{
    background-color: transparent;
    border: none;
    font-size: 12pt;
    font-weight: bold;
}}
QPushButton#WinBtn:hover {{
    background-color: {t['BG_CARD']};
}}
QPushButton#WinBtnClose:hover {{
    background-color: #ef4444;
    color: white;
}}
"""

# ─── Backend ──────────────────────────────────────────────────────────────────
class MoodleDownloader:
    def __init__(self, base_url, log_cb=print):
        self.base_url     = base_url.rstrip('/')
        self.session      = requests.Session()
        self.session.headers.update({
            'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                           'AppleWebKit/537.36 (KHTML, like Gecko) '
                           'Chrome/122.0.0.0 Safari/537.36'),
            'Accept-Language': 'en-US,en;q=0.9',
        })
        self.download_dir = 'moodle_downloads'
        self.log          = log_cb
        os.makedirs(self.download_dir, exist_ok=True)

    def login(self, username, password):
        login_url = urljoin(self.base_url, '/login/index.php')
        self.log(f'🔐 Connecting to {self.base_url} …')
        resp = self.session.get(login_url, timeout=20)
        resp.raise_for_status()
        soup  = BeautifulSoup(resp.text, 'lxml')
        token = soup.find('input', {'name': 'logintoken'})
        if not token:
            raise Exception('Login token not found — page structure may have changed.')
        self.session.post(login_url, timeout=20, data={
            'username': username, 'password': password,
            'logintoken': token['value']
        }).raise_for_status()
        # Verify we actually got in
        check = self.session.get(urljoin(self.base_url, '/my/'), timeout=20)
        if 'login' in check.url.lower():
            raise Exception('Login failed — check your username and password.')
        self.log('✅ Login successful!')

    def get_courses(self):
        self.log('📚 Fetching your courses …')
        resp = self.session.get(urljoin(self.base_url, '/my/'), timeout=20)
        resp.raise_for_status()
        soup  = BeautifulSoup(resp.text, 'lxml')
        links = []
        for a in soup.find_all('a', href=True):
            if '/course/view.php' in a['href']:
                full  = urljoin(self.base_url, a['href'])
                title = a.get_text(strip=True)
                if title:
                    links.append((title, full))
        unique = {url: title for title, url in links}
        final  = [(t, u) for u, t in unique.items()]
        self.log(f'🎓 Found {len(final)} course(s).')
        return final

    def download_course_files(self, course_title, course_url):
        self.log(f'\n📂  Course: {course_title}')
        folder = re.sub(r'[\\/*?:"<>|]', '_', course_title)
        dest   = os.path.join(self.download_dir, folder)
        os.makedirs(dest, exist_ok=True)
        self._visited = set()

        resp = self.session.get(course_url, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'lxml')

        for a in soup.find_all('a', href=True):
            href = a['href'].strip()
            if not href or href.startswith('#'):
                continue
            full      = urljoin(self.base_url, href)
            dedup_key = urlparse(full)._replace(fragment='').geturl()
            if dedup_key in self._visited:
                continue
            self._visited.add(dedup_key)
            path = urlparse(full).path.lower()
            if '/mod/folder/view.php' in path:
                self._process_folder(full, dest)
            elif '/mod/resource/view.php' in path:
                self._download_resource(full, dest)
            elif any(path.endswith(e) for e in ALLOWED_EXTENSIONS):
                self._download_file(full, dest)
            elif '/pluginfile.php' in path and self._is_file(full):
                self._download_file(full, dest)

    def _is_file(self, url):
        fn = os.path.basename(unquote(urlparse(url).path))
        return any(fn.lower().endswith(e) for e in ALLOWED_EXTENSIONS)

    def _process_folder(self, url, dest):
        try:
            resp  = self.session.get(url, timeout=20)
            resp.raise_for_status()
            soup  = BeautifulSoup(resp.text, 'lxml')
            title = soup.title.string.strip() if soup.title else 'Folder'
            self.log(f'  📁 Folder: {title}')
            for a in soup.find_all('a', href=True):
                href      = a['href']
                full      = urljoin(self.base_url, href)
                dedup_key = urlparse(full)._replace(fragment='').geturl()
                if dedup_key in self._visited:
                    continue
                self._visited.add(dedup_key)
                if '/pluginfile.php' in href and self._is_file(href):
                    self._download_file(full, dest)
                elif '/mod/folder/view.php' in href:
                    self._process_folder(full, dest)
        except Exception as e:
            self.log(f'  ⚠️  Folder error: {e}')

    def _download_resource(self, url, dest):
        try:
            resp = self.session.head(url, allow_redirects=True, timeout=20)
            resp.raise_for_status()
            if self._is_file(resp.url):
                self._download_file(resp.url, dest)
            else:
                resp = self.session.get(url, timeout=20)
                soup = BeautifulSoup(resp.text, 'lxml')
                for a in soup.find_all('a', href=True):
                    href = a['href']
                    if '/pluginfile.php' in href and self._is_file(href):
                        self._download_file(urljoin(self.base_url, href), dest)
                        break
        except Exception as e:
            self.log(f'  ⚠️  Resource error: {e}')

    def _download_file(self, url, dest):
        filename = None
        try:
            with self.session.get(url, stream=True, timeout=30) as r:
                r.raise_for_status()
                cd = r.headers.get('Content-Disposition', '')
                m  = re.findall('filename="?([^"]+)"?', cd)
                filename = (unquote(m[0]) if m
                            else os.path.basename(unquote(urlparse(r.url).path))
                            or 'unknown_file')
                filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
                fp = os.path.join(dest, filename)
                if os.path.exists(fp):
                    self.log(f'  ⏭️  Already exists: {filename}')
                    return
                self.log(f'  ⬇️  Downloading: {filename}')
                with open(fp, 'wb') as f:
                    for chunk in r.iter_content(8192):
                        f.write(chunk)
                self.log(f'  ✅ Saved: {filename}')
                time.sleep(0.4)
        except Exception as e:
            self.log(f'  ❌ Failed: {filename or url}  →  {e}')


# ─── GUI ──────────────────────────────────────────────────────────────────────
class MoodleApp(QWidget):
    log_signal = pyqtSignal(str)
    gui_update_signal = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.theme = 'light'

        # Frameless and translucent for custom rounded corners
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self.setWindowTitle('NSBM-dl')
        self.resize(750, 700)
        self.setMinimumSize(600, 600)

        # Variables for dragging the frameless window
        self._drag_pos = None

        # Set Application Logo
        icon_path = resource_path('logo.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.downloader = None
        self.courses = []
        self.download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
        if not os.path.exists(self.download_path):
            try:
                os.makedirs(self.download_path, exist_ok=True)
            except:
                pass # Fallback if we fail to create it initially

        self.pwd_visible = False

        self.log_signal.connect(self._append_log)
        self.gui_update_signal.connect(lambda func: func())

        self._build_ui()
        self.apply_theme()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # HEADER (Custom draggable title bar)
        self.hdr = QFrame()
        self.hdr.setObjectName("Header")
        hdr_layout = QHBoxLayout(self.hdr)
        hdr_layout.setContentsMargins(20, 10, 10, 10)

        titles_layout = QVBoxLayout()
        titles_layout.setSpacing(2)

        lbl_title = QLabel('NSBM-dl')
        lbl_title.setObjectName("HeaderTitle")
        titles_layout.addWidget(lbl_title)

        lbl_sub = QLabel('Download your files with ease . Made with ❤️ by Pamindu Fernando')
        lbl_sub.setObjectName("HeaderSub")
        titles_layout.addWidget(lbl_sub)

        hdr_layout.addLayout(titles_layout)
        hdr_layout.addStretch()

        self.lbl_about = QLabel('<a href="https://github.com/pamindu-fernando" style="text-decoration: none;">About Me</a>')
        self.lbl_about.setObjectName("AboutLink")
        self.lbl_about.setOpenExternalLinks(True)
        self.lbl_about.setCursor(Qt.CursorShape.PointingHandCursor)
        hdr_layout.addWidget(self.lbl_about)

        self.theme_btn = QPushButton()
        self.theme_btn.setObjectName("ThemeBtn")
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.setFixedSize(32, 32)
        self.theme_btn.clicked.connect(self._toggle_theme)
        hdr_layout.addWidget(self.theme_btn)
        
        # Window controls
        win_controls = QHBoxLayout()
        win_controls.setSpacing(0)
        
        min_btn = QPushButton("—")
        min_btn.setObjectName("WinBtn")
        min_btn.setFixedSize(32, 32)
        min_btn.clicked.connect(self.showMinimized)
        win_controls.addWidget(min_btn)
        
        close_btn = QPushButton("✕")
        close_btn.setObjectName("WinBtnClose")
        close_btn.setProperty("class", "WinBtn") # Reuse base style then override hover
        close_btn.setFixedSize(32, 32)
        close_btn.clicked.connect(self.close)
        win_controls.addWidget(close_btn)
        
        hdr_layout.addLayout(win_controls)
        main_layout.addWidget(self.hdr)
        
        body_layout = QVBoxLayout()
        body_layout.setContentsMargins(15, 12, 15, 12)
        body_layout.setSpacing(10)
        main_layout.addLayout(body_layout)
        
        # LOGIN CARD
        card = QFrame()
        card.setObjectName("Card")
        card_layout = QGridLayout(card)
        card_layout.setContentsMargins(15, 12, 15, 12)
        card_layout.setSpacing(6)
        
        sec_title1 = QLabel('Login')
        sec_title1.setObjectName("SectionTitle")
        card_layout.addWidget(sec_title1, 0, 0, 1, 3)
        
        lbl_save = QLabel('Save To:')
        lbl_save.setObjectName("FieldLabel")
        card_layout.addWidget(lbl_save, 1, 0)
        
        self.path_entry = QLineEdit(self.download_path)
        self.path_entry.setReadOnly(True)
        card_layout.addWidget(self.path_entry, 1, 1)
        
        btn_browse = QPushButton('Browse')
        btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_browse.clicked.connect(self._browse)
        card_layout.addWidget(btn_browse, 1, 2)
        
        lbl_url = QLabel('Nlearn URL:')
        lbl_url.setObjectName("FieldLabel")
        card_layout.addWidget(lbl_url, 2, 0)
        
        self.url_entry = QLineEdit('https://nlearn.nsbm.ac.lk')
        self.url_entry.setReadOnly(True)
        card_layout.addWidget(self.url_entry, 2, 1, 1, 2)
        
        lbl_user = QLabel('Username:')
        lbl_user.setObjectName("FieldLabel")
        card_layout.addWidget(lbl_user, 3, 0)
        
        self.user_entry = QLineEdit()
        card_layout.addWidget(self.user_entry, 3, 1, 1, 2)
        
        lbl_pwd = QLabel('Password:')
        lbl_pwd.setObjectName("FieldLabel")
        card_layout.addWidget(lbl_pwd, 4, 0)
        
        self.pass_entry = QLineEdit()
        self.pass_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_entry.setFixedHeight(28)
        card_layout.addWidget(self.pass_entry, 4, 1)
        
        self.eye_btn = QPushButton('Show')
        self.eye_btn.setObjectName("EyeBtn")
        self.eye_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.eye_btn.setFixedSize(55, 28)
        self.eye_btn.clicked.connect(self._toggle_pwd)
        card_layout.addWidget(self.eye_btn, 4, 2)
        
        # Put login button on its own row, right aligned or full width
        self.login_btn = QPushButton('Login && Fetch Modules')
        self.login_btn.setObjectName("PrimaryBtn")
        self.login_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.login_btn.clicked.connect(self._do_login)
        card_layout.addWidget(self.login_btn, 5, 1, 1, 2)
        
        body_layout.addWidget(card)
        
        # COURSE SELECTION (QTreeWidget)
        ccard = QFrame()
        ccard.setObjectName("Card")
        ccard_layout = QVBoxLayout(ccard)
        ccard_layout.setContentsMargins(15, 10, 15, 10)
        ccard_layout.setSpacing(6)
        
        c_hdr_layout = QHBoxLayout()
        sec_title2 = QLabel('Select Modules')
        sec_title2.setObjectName("SectionTitle")
        c_hdr_layout.addWidget(sec_title2)
        c_hdr_layout.addStretch()
        
        self.sel_all_btn = QPushButton('Select All')
        self.sel_all_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.sel_all_btn.setFixedWidth(80)
        self.sel_all_btn.setEnabled(False)
        self.sel_all_btn.clicked.connect(self._select_all)
        c_hdr_layout.addWidget(self.sel_all_btn)
        
        ccard_layout.addLayout(c_hdr_layout)
        
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setSelectionMode(QTreeWidget.SelectionMode.NoSelection)
        self.tree.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.tree.itemChanged.connect(self._handle_tree_check)
        ccard_layout.addWidget(self.tree)
        
        self.dl_btn = QPushButton('Download Selected')
        self.dl_btn.setObjectName("PrimaryBtn")
        self.dl_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dl_btn.setEnabled(False)
        self.dl_btn.clicked.connect(self._do_download)
        ccard_layout.addWidget(self.dl_btn)
        
        body_layout.addWidget(ccard, 2) # Stretch factor 2
        
        # LOG PANEL
        lcard = QFrame()
        lcard.setObjectName("Card")
        lcard_layout = QVBoxLayout(lcard)
        lcard_layout.setContentsMargins(15, 10, 15, 10)
        lcard_layout.setSpacing(6)
        
        l_hdr_layout = QHBoxLayout()
        sec_title3 = QLabel('Activity Log')
        sec_title3.setObjectName("SectionTitle")
        l_hdr_layout.addWidget(sec_title3)
        l_hdr_layout.addStretch()
        
        btn_clear = QPushButton('Clear')
        btn_clear.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_clear.setFixedWidth(60)
        btn_clear.clicked.connect(self._clear_log)
        l_hdr_layout.addWidget(btn_clear)
        
        lcard_layout.addLayout(l_hdr_layout)
        
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 8))
        self.log_box.setFixedHeight(110)
        lcard_layout.addWidget(self.log_box)
        
        body_layout.addWidget(lcard, 1) # Stretch factor 1

    def apply_theme(self):
        stylesheet = get_stylesheet(self.theme)
        self.setStyleSheet(stylesheet)
        self.theme_btn.setText(THEMES[self.theme]['ICON'])
        
        # Manually update the link color to match the generic text so it flips white/black
        txt_color = THEMES[self.theme]['TEXT_MAIN']
        self.lbl_about.setText(f'<a href="https://github.com/pamindu-fernando" style="color: {txt_color}; text-decoration: none;">About Me</a>')

    def _toggle_theme(self):
        self.theme = 'light' if self.theme == 'dark' else 'dark'
        self.apply_theme()
        
    def _toggle_pwd(self):
        self.pwd_visible = not self.pwd_visible
        if self.pwd_visible:
            self.pass_entry.setEchoMode(QLineEdit.EchoMode.Normal)
            self.eye_btn.setText('Hide')
        else:
            self.pass_entry.setEchoMode(QLineEdit.EchoMode.Password)
            self.eye_btn.setText('Show')
            
    def _browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Download Directory", self.download_path)
        if folder:
            self.download_path = folder
            self.path_entry.setText(folder)
            
    def log(self, msg):
        self.log_signal.emit(msg)
        
    def _append_log(self, msg):
        self.log_box.append(msg)
        # Scroll to bottom
        scrollbar = self.log_box.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
    def _clear_log(self):
        self.log_box.clear()
        
    def _do_login(self):
        url = self.url_entry.text().strip()
        user = self.user_entry.text().strip()
        pwd = self.pass_entry.text().strip()
        
        if not all([url, user, pwd]):
            QMessageBox.critical(self, 'Missing fields', 'Please fill in all fields.')
            return
            
        self.login_btn.setEnabled(False)
        self.login_btn.setText('Logging in...')
        self._clear_log()
        self.tree.clear()
        
        threading.Thread(target=self._login_worker, args=(url, user, pwd), daemon=True).start()
                
    def _login_worker(self, url, user, pwd):
        self.downloader = MoodleDownloader(url, log_cb=self.log)
        self.downloader.download_dir = self.download_path
        try:
            self.downloader.login(user, pwd)
            self.courses = self.downloader.get_courses()
            self.gui_update_signal.emit(self._populate_courses)
        except Exception as e:
            err = str(e)
            self.log(f'❌ {err}')
            self.gui_update_signal.emit(lambda: QMessageBox.critical(self, 'Login Failed', err))
        finally:
            def reset_btn():
                self.login_btn.setEnabled(True)
                self.login_btn.setText('Login && Fetch Modules')
            self.gui_update_signal.emit(reset_btn)
            
    def _cat(self, t):
        m = re.search(r'(Y[1-4]S[1-4])', t, re.I)
        return m.group(1).upper() if m else 'Other Modules'
        
    def _populate_courses(self):
        self.tree.clear()
        if not self.courses:
            item = QTreeWidgetItem(["No courses found."])
            self.tree.addTopLevelItem(item)
            return
            
        grouped = {}
        for t, u in self.courses:
            grouped.setdefault(self._cat(t), []).append((t, u))
            
        cats = sorted(grouped, key=lambda x: (x == 'Other Modules', x))
        
        self.tree.blockSignals(True)
        for cat in cats:
            root = QTreeWidgetItem([cat])
            root.setFlags(root.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            root.setCheckState(0, Qt.CheckState.Unchecked)
            
            # Make the root item slightly bolder
            font = root.font(0)
            font.setBold(True)
            root.setFont(0, font)

            self.tree.addTopLevelItem(root)
            
            for title, url in grouped[cat]:
                child = QTreeWidgetItem([title])
                child.setFlags(child.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                child.setCheckState(0, Qt.CheckState.Unchecked)
                child.setData(0, Qt.ItemDataRole.UserRole, (title, url))
                root.addChild(child)
            
            root.setExpanded(True)
            
        self.tree.blockSignals(False)
        self.sel_all_btn.setEnabled(True)
        self.log(f'✅ {len(self.courses)} course(s) loaded. Check categories or individual modules!')
        
    def _handle_tree_check(self, item, column):
        self.tree.blockSignals(True)
        # If it's a root item (category)
        if item.parent() is None:
            state = item.checkState(column)
            for i in range(item.childCount()):
                item.child(i).setCheckState(column, state)
        else:
            # It's a child item (course)
            parent = item.parent()
            all_checked = True
            any_checked = False
            for i in range(parent.childCount()):
                if parent.child(i).checkState(column) == Qt.CheckState.Checked:
                    any_checked = True
                else:
                    all_checked = False
            
            if all_checked:
                parent.setCheckState(column, Qt.CheckState.Checked)
            elif any_checked:
                parent.setCheckState(column, Qt.CheckState.PartiallyChecked)
            else:
                parent.setCheckState(column, Qt.CheckState.Unchecked)
                
        self.tree.blockSignals(False)
        self._check_dl_btn_state()

    def _check_dl_btn_state(self):
        any_checked = False
        all_checked = True
        has_items = False
        
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            if root.childCount() == 0:
                continue
            has_items = True
            for j in range(root.childCount()):
                child = root.child(j)
                if child.checkState(0) == Qt.CheckState.Checked:
                    any_checked = True
                else:
                    all_checked = False
                    
        if not has_items:
            all_checked = False
            
        self.dl_btn.setEnabled(any_checked)
        self.sel_all_btn.setText("Deselect All" if all_checked else "Select All")
        
    def _select_all(self):
        new_state = Qt.CheckState.Checked if self.sel_all_btn.text() == "Select All" else Qt.CheckState.Unchecked
        self.tree.blockSignals(True)
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            root.setCheckState(0, new_state)
            for j in range(root.childCount()):
                root.child(j).setCheckState(0, new_state)
        self.tree.blockSignals(False)
        self._check_dl_btn_state()
        
    def _do_download(self):
        selected = []
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            for j in range(root.childCount()):
                child = root.child(j)
                if child.checkState(0) == Qt.CheckState.Checked:
                    data = child.data(0, Qt.ItemDataRole.UserRole)
                    if data:
                        selected.append(data)
                        
        if not selected:
            return
            
        self.dl_btn.setEnabled(False)
        self.dl_btn.setText('Downloading...')
        self.login_btn.setEnabled(False)
        self.sel_all_btn.setEnabled(False)
        self.log(f'\n🚀 Starting download of {len(selected)} course(s) …\n')
        
        threading.Thread(target=self._dl_worker, args=(selected,), daemon=True).start()
        
    def _dl_worker(self, courses):
        try:
            for t, u in courses:
                self.downloader.download_course_files(t, u)
            self.log('\n🎉  All downloads complete!')
            self.gui_update_signal.emit(lambda: QMessageBox.information(self, 'Done!', 'All modules downloaded successfully!'))
        except Exception as e:
            err = str(e)
            self.log(f'❌ Download error: {err}')
            self.gui_update_signal.emit(lambda: QMessageBox.critical(self, 'Download Error', err))
        finally:
            self.gui_update_signal.emit(self._finish_dl)
            
    def _finish_dl(self):
        self._check_dl_btn_state()
        self.dl_btn.setText('Download Selected')
        self.login_btn.setEnabled(True)
        self.sel_all_btn.setEnabled(True)
        
        # Reveal folder natively
        if os.path.exists(self.download_path):
            if os.name == 'nt':
                os.startfile(self.download_path)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.call(['open', self.download_path])
            else:
                import subprocess
                subprocess.call(['xdg-open', self.download_path])


    # --- Custom Window Painting & Dragging ---
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Fill rounded rect with theme background
        bg_color = QColor(THEMES[self.theme]['BG_APP'])
        painter.setBrush(QBrush(bg_color))
        painter.setPen(Qt.PenStyle.NoPen)
        
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), 14.0, 14.0) # 14px border radius
        painter.drawPath(path)
        
        # Optional: draw subtle border
        painter.setPen(QColor(THEMES[self.theme]['BORDER']))
        painter.drawPath(path)
        
    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            # Only allow drag from the header area
            if self.hdr.geometry().contains(event.pos()):
                self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()

    def mouseMoveEvent(self, event: QMouseEvent):
        if self._drag_pos is not None and event.buttons() == Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event: QMouseEvent):
        self._drag_pos = None


# ─── Entry ────────────────────────────────────────────────────────────────────
def main():
    if os.name == 'nt':
        import ctypes
        myappid = 'nsbm.moodle.downloader.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        
    app = QApplication(sys.argv)
    window = MoodleApp()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()