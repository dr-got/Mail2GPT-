import os
import sys
import re
import logging
import traceback
import tempfile
import struct
from datetime import datetime
import email
from email import policy
from email.utils import parsedate_to_datetime
from email.header import decode_header

from PySide6.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout, 
                               QWidget, QTextEdit, QMessageBox, QPushButton, QHBoxLayout, QCheckBox, QFileDialog, QComboBox)
from PySide6.QtCore import Qt, QUrl, QTimer, QSettings
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from bs4 import BeautifulSoup

try:
    import extract_msg
except Exception:
    extract_msg = None

APP_VERSION = "0.4.0"

TRANSLATIONS = {
    "it": {
        "title": "Mail2GPT",
        "drop_hint": "Trascina qui un'email da Outlook oppure un file .eml/.msg\n(Se Outlook non fornisce un file, usa la cartella monitorata)",
        "language": "Lingua",
        "watch_on": "Monitoraggio ON",
        "watch_off": "Monitoraggio OFF",
        "watched_folder": "Cartella monitorata: {path}",
        "choose_folder": "Scegli cartella",
        "open_folder": "Apri cartella",
        "choose_folder_title": "Scegli cartella da monitorare",
        "app_start": "Avvio app v{version}",
        "manual_open": "Apri manualmente: {path}",
        "watch_enabled_log": "Monitoraggio cartella: ON",
        "watch_disabled_log": "Monitoraggio cartella: OFF",
        "watch_new_file": "Rilevato nuovo file in cartella monitorata: {name}",
        "watch_set": "Cartella monitorata impostata: {path}",
        "watch_set_error": "Impossibile usare la cartella scelta: {error}",
        "drop_non_local_header": "Drop con URL non-locali (link).",
        "drop_non_local_hint": "Questo tipo di Outlook può non esporre il contenuto dell'email via drag&drop. In tal caso salva/trascina come .eml oppure usa Outlook classico.",
        "drop_outlook_detected": "Drop da Outlook rilevato. Output in: {path}",
        "drop_outlook_extract_fail": "Drop da Outlook rilevato, ma non è stato possibile estrarre i contenuti.",
        "drop_formats": "Formati disponibili: {formats}",
        "drop_unrecognized": "Formato di drop non riconosciuto. Trascina un file .eml/.msg o un'email da Outlook.",
        "ignored_type": "Ignorato (tipo file non supportato): {name}",
        "processing_start": "Inizio elaborazione: {name}",
        "folder_created": "Cartella creata in: {path}",
        "attach_saved": " - Allegato salvato: {name}",
        "attach_save_fail": " - Impossibile salvare l'allegato {name}: {error}",
        "done": "Operazione completata! {count} allegati estratti con successo.",
        "fatal_error": "Errore irreversibile: {error}",
        "msg_support_missing": "Supporto .msg non disponibile: installa 'extract-msg' e ricompila l'eseguibile.",
        "md_title": "# Email",
        "md_header": "## Header",
        "md_instructions": "## Istruzioni",
        "md_instructions_text": "Leggi prima l'header, poi il body. Gli allegati sono salvati come file separati nella stessa cartella di questo .md.",
        "md_attachments": "## Allegati",
        "md_no_attachments": "- (nessun allegato)",
        "md_body": "## Body",
    },
    "en": {
        "title": "Mail2GPT",
        "drop_hint": "Drop an email from Outlook or a .eml/.msg file here\n(If Outlook doesn’t provide a file, use the watched folder)",
        "language": "Language",
        "watch_on": "Watching ON",
        "watch_off": "Watching OFF",
        "watched_folder": "Watched folder: {path}",
        "choose_folder": "Choose folder",
        "open_folder": "Open folder",
        "choose_folder_title": "Choose folder to watch",
        "app_start": "App started v{version}",
        "manual_open": "Open manually: {path}",
        "watch_enabled_log": "Folder watching: ON",
        "watch_disabled_log": "Folder watching: OFF",
        "watch_new_file": "New file detected in watched folder: {name}",
        "watch_set": "Watched folder set to: {path}",
        "watch_set_error": "Cannot use selected folder: {error}",
        "drop_non_local_header": "Drop contains non-local URLs (links).",
        "drop_non_local_hint": "This Outlook version may not expose email content via drag&drop. Save/drop as .eml or use classic Outlook.",
        "drop_outlook_detected": "Outlook drop detected. Output to: {path}",
        "drop_outlook_extract_fail": "Outlook drop detected, but content could not be extracted.",
        "drop_formats": "Available formats: {formats}",
        "drop_unrecognized": "Drop format not recognized. Drop a .eml/.msg file or an email from Outlook.",
        "ignored_type": "Ignored (unsupported file type): {name}",
        "processing_start": "Processing started: {name}",
        "folder_created": "Folder created at: {path}",
        "attach_saved": " - Attachment saved: {name}",
        "attach_save_fail": " - Cannot save attachment {name}: {error}",
        "done": "Done! {count} attachments extracted successfully.",
        "fatal_error": "Fatal error: {error}",
        "msg_support_missing": ".msg support not available: install 'extract-msg' and rebuild the executable.",
        "md_title": "# Email",
        "md_header": "## Header",
        "md_instructions": "## Instructions",
        "md_instructions_text": "Read the header first, then the body. Attachments are saved as separate files in the same folder as this .md file.",
        "md_attachments": "## Attachments",
        "md_no_attachments": "- (no attachments)",
        "md_body": "## Body",
    },
}

# Configurazione logging base
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def sanitize_filename(name):
    """Rimuove i caratteri non validi per i nomi di file e cartelle su Windows."""
    if not name:
        return "Sconosciuto"
    # Rimuove i caratteri vietati e spazi multipli
    sanitized = re.sub(r'[\\/*?:"<>|]', "", str(name))
    sanitized = re.sub(r'\s+', " ", sanitized)
    return sanitized.strip()[:100]  # Limita la lunghezza del nome

def make_output_stem(formatted_date, sender, subject):
    stem = f"{formatted_date}_{sanitize_filename(sender)}_{sanitize_filename(subject)}"
    return sanitize_filename(stem)[:120]

def parse_filegroupdescriptor(blob: bytes):
    if len(blob) < 4:
        return []
    count = struct.unpack_from("<I", blob, 0)[0]
    names = []
    entry_size = 592
    name_offset = 72
    for i in range(count):
        off = 4 + i * entry_size
        if off + entry_size > len(blob):
            break
        raw_name = blob[off + name_offset: off + entry_size]
        try:
            name = raw_name.decode("utf-16le", errors="replace").split("\x00", 1)[0].strip()
        except Exception:
            name = ""
        if name:
            names.append(name)
    return names

def parse_filegroupdescriptor_ansi(blob: bytes):
    if len(blob) < 4:
        return []
    count = struct.unpack_from("<I", blob, 0)[0]
    names = []
    entry_size = 332
    name_offset = 72
    for i in range(count):
        off = 4 + i * entry_size
        if off + entry_size > len(blob):
            break
        raw_name = blob[off + name_offset: off + entry_size]
        try:
            name = raw_name.split(b"\x00", 1)[0].decode("utf-8", errors="replace").strip()
        except Exception:
            name = ""
        if name:
            names.append(name)
    return names

def mime_has_filecontents(mime_data):
    try:
        for f in mime_data.formats():
            if 'value="FileContents"' in f:
                return True
    except Exception:
        pass
    return mime_data.hasFormat('application/x-qt-windows-mime;value="FileContents"')

def parse_text_uri_list(data: bytes):
    if not data:
        return []
    try:
        text = data.decode("utf-8", errors="replace")
    except Exception:
        try:
            text = data.decode("latin-1", errors="replace")
        except Exception:
            return []
    lines = []
    for raw in text.splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        lines.append(s)
    return lines

class EmailExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("Mail2GPT", "Mail2GPT")
        self.lang = str(self.settings.value("lang", "it"))
        if self.lang not in TRANSLATIONS:
            self.lang = "it"

        self.setWindowTitle(f"{self.t('title')} v{APP_VERSION}")
        self.resize(700, 450)
        self.setAcceptDrops(True)

        # Interfaccia grafica
        layout = QVBoxLayout()

        lang_layout = QHBoxLayout()
        self.lang_label = QLabel(self.t("language"))
        lang_layout.addWidget(self.lang_label, 0)
        self.lang_combo = QComboBox()
        self.lang_combo.addItem("Italiano", "it")
        self.lang_combo.addItem("English", "en")
        self.lang_combo.setCurrentIndex(0 if self.lang == "it" else 1)
        self.lang_combo.currentIndexChanged.connect(self.on_language_changed)
        lang_layout.addWidget(self.lang_combo, 0)
        lang_layout.addStretch(1)
        layout.addLayout(lang_layout)
        
        self.label = QLabel(
            self.t("drop_hint")
        )
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                font-size: 12pt;
                color: #555;
                background-color: #f9f9f9;
                padding: 20px;
            }
        """)
        layout.addWidget(self.label)
        
        self.watch_dir = str(self.settings.value("watch_dir", os.path.join(os.path.expanduser("~"), "Documents", "OutlookDrop")))
        self.ensure_watch_dirs()
        self.watch_state = {}

        watch_layout = QHBoxLayout()
        self.watch_enabled_checkbox = QCheckBox(self.t("watch_on"))
        self.watch_enabled_checkbox.setChecked(str(self.settings.value("watch_enabled", "true")).lower() in ["1", "true", "yes"])
        self.watch_enabled_checkbox.stateChanged.connect(self.on_watch_toggle)
        watch_layout.addWidget(self.watch_enabled_checkbox, 0)

        self.watch_label = QLabel(self.t("watched_folder", path=self.watch_dir))
        self.watch_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        watch_layout.addWidget(self.watch_label, 1)
        self.choose_watch_btn = QPushButton(self.t("choose_folder"))
        self.choose_watch_btn.clicked.connect(self.choose_watch_folder)
        watch_layout.addWidget(self.choose_watch_btn, 0)
        self.open_watch_btn = QPushButton(self.t("open_folder"))
        self.open_watch_btn.clicked.connect(self.open_watch_folder)
        watch_layout.addWidget(self.open_watch_btn, 0)
        layout.addLayout(watch_layout)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("font-family: Consolas, monospace; font-size: 10pt;")
        layout.addWidget(self.log_area)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.log(self.t("app_start", version=APP_VERSION))
        try:
            print(f"APP_VERSION={APP_VERSION}", flush=True)
        except Exception:
            pass

        self.watch_timer = QTimer(self)
        self.watch_timer.setInterval(1000)
        self.watch_timer.timeout.connect(self.poll_watch_folder)
        if self.watch_enabled_checkbox.isChecked():
            self.watch_timer.start()

    def log(self, message, is_error=False):
        """Aggiunge un messaggio al log a schermo."""
        color = "#D32F2F" if is_error else "#388E3C"
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.append(f'<span style="color:{color};">[{timestamp}] {message}</span>')
        # Scorri in basso
        scrollbar = self.log_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        QApplication.processEvents()

    def t(self, key, **kwargs):
        table = TRANSLATIONS.get(self.lang, TRANSLATIONS["it"])
        s = table.get(key, TRANSLATIONS["it"].get(key, key))
        try:
            return s.format(**kwargs)
        except Exception:
            return s

    def apply_language(self):
        self.setWindowTitle(f"{self.t('title')} v{APP_VERSION}")
        self.label.setText(self.t("drop_hint"))
        self.lang_label.setText(self.t("language"))
        self.update_watch_label()
        self.choose_watch_btn.setText(self.t("choose_folder"))
        self.open_watch_btn.setText(self.t("open_folder"))
        self.watch_enabled_checkbox.setText(self.t("watch_on") if self.watch_enabled_checkbox.isChecked() else self.t("watch_off"))

    def on_language_changed(self):
        code = self.lang_combo.currentData()
        if code not in TRANSLATIONS:
            return
        if code == self.lang:
            return
        self.lang = code
        self.settings.setValue("lang", self.lang)
        self.apply_language()

    def ensure_watch_dirs(self):
        os.makedirs(self.watch_dir, exist_ok=True)
        self.processed_dir = os.path.join(self.watch_dir, "_processed")
        os.makedirs(self.processed_dir, exist_ok=True)

    def update_watch_label(self):
        self.watch_label.setText(self.t("watched_folder", path=self.watch_dir))

    def on_watch_toggle(self):
        enabled = self.watch_enabled_checkbox.isChecked()
        self.settings.setValue("watch_enabled", "true" if enabled else "false")
        self.watch_enabled_checkbox.setText(self.t("watch_on") if enabled else self.t("watch_off"))
        if enabled:
            if not self.watch_timer.isActive():
                self.watch_timer.start()
            self.log(self.t("watch_enabled_log"))
        else:
            if self.watch_timer.isActive():
                self.watch_timer.stop()
            self.log(self.t("watch_disabled_log"))

    def open_watch_folder(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.watch_dir)
                return
        except Exception:
            pass
        self.log(self.t("manual_open", path=self.watch_dir), is_error=True)

    def choose_watch_folder(self):
        chosen = QFileDialog.getExistingDirectory(self, self.t("choose_folder_title"), self.watch_dir)
        if not chosen:
            return
        self.watch_dir = chosen
        try:
            self.ensure_watch_dirs()
        except Exception as e:
            self.log(self.t("watch_set_error", error=str(e)), is_error=True)
            return
        self.settings.setValue("watch_dir", self.watch_dir)
        self.watch_state = {}
        self.update_watch_label()
        self.log(self.t("watch_set", path=self.watch_dir))

    def poll_watch_folder(self):
        if not self.watch_enabled_checkbox.isChecked():
            return
        try:
            entries = []
            for name in os.listdir(self.watch_dir):
                if name.startswith(".") or name == "_processed":
                    continue
                p = os.path.join(self.watch_dir, name)
                if not os.path.isfile(p):
                    continue
                ext = os.path.splitext(name)[1].lower()
                if ext not in [".eml", ".msg"]:
                    continue
                try:
                    st = os.stat(p)
                except Exception:
                    continue
                entries.append((p, st.st_size, st.st_mtime))

            for p, size, mtime in entries:
                prev = self.watch_state.get(p)
                if prev is None:
                    self.watch_state[p] = (size, mtime, 0)
                    continue
                prev_size, prev_mtime, stable = prev
                if size == prev_size and mtime == prev_mtime:
                    stable += 1
                else:
                    stable = 0
                self.watch_state[p] = (size, mtime, stable)

                if stable < 1:
                    continue

                self.watch_state.pop(p, None)
                self.log(self.t("watch_new_file", name=os.path.basename(p)))
                self.process_dropped_file(p, base_dir_override=self.watch_dir)

                dst = os.path.join(self.processed_dir, os.path.basename(p))
                base, ext = os.path.splitext(dst)
                n = 1
                while os.path.exists(dst):
                    dst = f"{base}_{n}{ext}"
                    n += 1
                try:
                    os.replace(p, dst)
                except Exception:
                    pass

            known = set(self.watch_state.keys())
            current = set(p for p, _, _ in entries)
            for p in known - current:
                self.watch_state.pop(p, None)
        except Exception:
            return

    def get_outlook_drop_output_base_dir(self):
        home = os.path.expanduser("~")
        candidates = [
            os.path.join(home, "Documents", "EmailExport"),
            os.path.join(home, "Desktop", "EmailExport"),
            os.path.join(home, "EmailExport"),
            os.getcwd(),
        ]
        for c in candidates:
            try:
                os.makedirs(c, exist_ok=True)
                return c
            except Exception:
                continue
        return os.getcwd()

    def extract_outlook_virtual_files(self, mime_data, temp_dir):
        fmt_desc_w = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
        fmt_desc_a = 'application/x-qt-windows-mime;value="FileGroupDescriptor"'
        names = []
        if mime_data.hasFormat(fmt_desc_w):
            names = parse_filegroupdescriptor(bytes(mime_data.data(fmt_desc_w)))
        elif mime_data.hasFormat(fmt_desc_a):
            names = parse_filegroupdescriptor_ansi(bytes(mime_data.data(fmt_desc_a)))
        formats = []
        try:
            formats = mime_data.formats()
        except Exception:
            formats = []
        content_formats = [f for f in formats if 'value="FileContents"' in f]
        if not names and not content_formats and not mime_data.hasFormat('application/x-qt-windows-mime;value="FileContents"'):
            return []

        if not names:
            blobs = []
            seen = set()
            if mime_data.hasFormat('application/x-qt-windows-mime;value="FileContents"'):
                d = bytes(mime_data.data('application/x-qt-windows-mime;value="FileContents"'))
                if d:
                    h = hash(d)
                    if h not in seen:
                        seen.add(h)
                        blobs.append(d)
            for f in content_formats:
                d = bytes(mime_data.data(f))
                if not d:
                    continue
                h = hash(d)
                if h in seen:
                    continue
                seen.add(h)
                blobs.append(d)

            extracted_paths = []
            for i, data in enumerate(blobs):
                safe_name = f"outlook_drop_{i}"
                if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
                    safe_name += ".msg"
                else:
                    safe_name += ".eml"

                out_path = os.path.join(temp_dir, safe_name)
                with open(out_path, "wb") as wf:
                    wf.write(data)
                extracted_paths.append(out_path)
            return extracted_paths

        extracted_paths = []
        for i, name in enumerate(names):
            fmt_contents_i = f'application/x-qt-windows-mime;value="FileContents";index={i}'
            fmt_contents = 'application/x-qt-windows-mime;value="FileContents"'

            if mime_data.hasFormat(fmt_contents_i):
                data = bytes(mime_data.data(fmt_contents_i))
            elif len(names) == 1 and mime_data.hasFormat(fmt_contents):
                data = bytes(mime_data.data(fmt_contents))
            else:
                data = b""

            if not data:
                continue

            safe_name = sanitize_filename(name).replace("\\", "_").replace("/", "_")
            if not os.path.splitext(safe_name)[1]:
                if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
                    safe_name += ".msg"
                else:
                    safe_name += ".eml"

            out_path = os.path.join(temp_dir, safe_name)
            base, ext = os.path.splitext(out_path)
            n = 1
            while os.path.exists(out_path):
                out_path = f"{base}_{n}{ext}"
                n += 1

            with open(out_path, "wb") as f:
                f.write(data)
            extracted_paths.append(out_path)

        return extracted_paths

    def dragEnterEvent(self, event: QDragEnterEvent):
        # Accetta solo file o URL locali
        mime = event.mimeData()
        if mime.hasUrls() or mime.hasFormat('application/x-qt-windows-mime;value="FileGroupDescriptorW"') or mime.hasFormat('application/x-qt-windows-mime;value="FileGroupDescriptor"') or mime_has_filecontents(mime):
            event.acceptProposedAction()
            self.label.setStyleSheet("""
                QLabel {
                    border: 2px dashed #4CAF50;
                    border-radius: 10px;
                    font-size: 12pt;
                    color: #4CAF50;
                    background-color: #e8f5e9;
                    padding: 20px;
                }
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        # Ripristina lo stile originale quando si esce dalla finestra
        self.label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                font-size: 12pt;
                color: #555;
                background-color: #f9f9f9;
                padding: 20px;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        self.dragLeaveEvent(event)  # Ripristina lo stile
        mime = event.mimeData()
        did_any = False

        if mime.hasUrls():
            any_local = False
            non_local = []
            for url in mime.urls():
                if isinstance(url, QUrl) and url.isLocalFile():
                    file_path = url.toLocalFile()
                    if not file_path:
                        continue
                    any_local = True
                    did_any = True
                    self.process_dropped_file(file_path, base_dir_override=os.path.dirname(file_path))
                else:
                    try:
                        non_local.append(url.toString())
                    except Exception:
                        pass

            if did_any and any_local:
                return

            if not non_local and mime.hasFormat("text/uri-list"):
                raw_uris = parse_text_uri_list(bytes(mime.data("text/uri-list")))
                for u in raw_uris:
                    q = QUrl(u)
                    if q.isLocalFile():
                        p = q.toLocalFile()
                        if p:
                            any_local = True
                            did_any = True
                            self.process_dropped_file(p, base_dir_override=os.path.dirname(p))
                    else:
                        non_local.append(u)

                if did_any and any_local:
                    return

            if non_local:
                self.log(self.t("drop_non_local_header"), is_error=True)
                for u in non_local[:10]:
                    self.log(u, is_error=True)
                self.log(self.t("drop_non_local_hint"), is_error=True)

        if not did_any and (mime.hasFormat('application/x-qt-windows-mime;value="FileGroupDescriptorW"') or mime.hasFormat('application/x-qt-windows-mime;value="FileGroupDescriptor"') or mime_has_filecontents(mime)):
            base_dir = self.get_outlook_drop_output_base_dir()
            self.log(self.t("drop_outlook_detected", path=base_dir))
            with tempfile.TemporaryDirectory(prefix="outlook-drop-") as tmp:
                extracted = self.extract_outlook_virtual_files(mime, tmp)
                if not extracted:
                    self.log(self.t("drop_outlook_extract_fail"), is_error=True)
                    try:
                        self.log(self.t("drop_formats", formats=", ".join(mime.formats())), is_error=True)
                    except Exception:
                        pass
                    return
                for p in extracted:
                    did_any = True
                    self.process_dropped_file(p, base_dir_override=base_dir)

        if not did_any:
            try:
                self.log(self.t("drop_formats", formats=", ".join(mime.formats())), is_error=True)
            except Exception:
                pass
            self.log(self.t("drop_unrecognized"), is_error=True)

    def process_dropped_file(self, file_path, base_dir_override=None):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".eml":
            self.process_eml_file(file_path, base_dir_override=base_dir_override)
            return
        if ext == ".msg":
            self.process_msg_file(file_path, base_dir_override=base_dir_override)
            return

        try:
            with open(file_path, "rb") as f:
                head = f.read(2048)
        except Exception:
            head = b""

        if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            self.process_msg_file(file_path, base_dir_override=base_dir_override)
            return
        if head.startswith(b"From:") or b"\nFrom:" in head or head.startswith(b"Received:") or b"\nReceived:" in head:
            self.process_eml_file(file_path, base_dir_override=base_dir_override)
            return

        self.log(self.t("ignored_type", name=os.path.basename(file_path)), is_error=True)

    def process_eml_file(self, file_path, base_dir_override=None):
        try:
            filename = os.path.basename(file_path)
            self.log(self.t("processing_start", name=filename))
            
            # Lettura del file .eml
            with open(file_path, 'rb') as f:
                msg = email.message_from_binary_file(f, policy=policy.default)
            
            # Funzione di supporto per decodificare gli header
            def get_decoded_header(header_name):
                value = msg.get(header_name, '')
                if not value:
                    return ''
                decoded_parts = decode_header(value)
                result = []
                for part, encoding in decoded_parts:
                    if isinstance(part, bytes):
                        try:
                            result.append(part.decode(encoding or 'utf-8', errors='replace'))
                        except Exception:
                            result.append(part.decode('utf-8', errors='replace'))
                    else:
                        result.append(part)
                return ''.join(result)
            
            # Estrazione metadati principali
            sender = get_decoded_header('From') or "MittenteSconosciuto"
            subject = get_decoded_header('Subject') or "NessunOggetto"
            date_str = msg.get('Date', '')
            
            # Parsing robusto della data
            try:
                if date_str:
                    dt = parsedate_to_datetime(date_str)
                    formatted_date = dt.strftime("%Y-%m-%d_%H%M%S")
                    readable_date = dt.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    raise ValueError("Data assente")
            except Exception:
                # Fallback se la data non è parsabile o assente
                formatted_date = datetime.now().strftime("%Y-%m-%d_%H%M%S")
                readable_date = str(date_str) if date_str else "Data sconosciuta"
                
            # Creazione nome cartella: YYYY-MM-DD_HHMMSS_Mittente_Oggetto
            folder_name = f"{formatted_date}_{sanitize_filename(sender)}_{sanitize_filename(subject)}"
            base_dir = base_dir_override or os.path.dirname(file_path)
            output_dir = os.path.join(base_dir, folder_name)
            
            # Evita conflitti se la cartella esiste già
            counter = 1
            original_output_dir = output_dir
            while os.path.exists(output_dir):
                output_dir = f"{original_output_dir}_{counter}"
                counter += 1
                
            os.makedirs(output_dir)
            self.log(self.t("folder_created", path=output_dir))
            output_stem = make_output_stem(formatted_date, sender, subject)
            
            # 1. Gestione del corpo dell'email e degli allegati
            body_text = ""
            html_body = ""
            attachments = []
            
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # Cerca gli allegati
                if "attachment" in content_disposition or part.get_filename():
                    attachments.append(part)
                    continue
                
                # Cerca il testo normale
                if content_type == "text/plain":
                    try:
                        part_text = part.get_content()
                    except Exception:
                        payload = part.get_payload(decode=True)
                        if payload:
                            part_text = payload.decode(part.get_content_charset() or 'utf-8', errors='replace')
                        else:
                            part_text = ""
                    
                    if part_text:
                        if not html_body: # Se non abbiamo ancora l'HTML, usiamo il testo
                            body_text = part_text
                        elif not body_text: # Fallback
                            body_text = part_text

                # Cerca l'HTML
                elif content_type == "text/html":
                    try:
                        part_html = part.get_content()
                    except Exception:
                        payload = part.get_payload(decode=True)
                        if payload:
                            part_html = payload.decode(part.get_content_charset() or 'utf-8', errors='replace')
                        else:
                            part_html = ""
                            
                    if part_html:
                        html_body = part_html
            
            # Se presente l'HTML, usiamo BeautifulSoup per avere un testo più pulito
            if html_body:
                try:
                    soup = BeautifulSoup(html_body, "html.parser")
                    body_text = soup.get_text(separator="\n", strip=True)
                except Exception as e:
                    self.log(f"Errore conversione HTML (fallback al testo normale): {str(e)}", is_error=True)
            
            if not body_text:
                body_text = "Nessun contenuto testuale trovato nell'email."
                
            # 2. Estrazione degli allegati
            attachments_count = 0
            saved_attachments = []
            for part in attachments:
                att_name = part.get_filename()
                if not att_name:
                    att_name = f"allegato_sconosciuto_{attachments_count}.bin"
                else:
                    # Decodifica il nome del file se necessario
                    decoded_parts = decode_header(att_name)
                    att_name = ''
                    for p, encoding in decoded_parts:
                        if isinstance(p, bytes):
                            att_name += p.decode(encoding or 'utf-8', errors='replace')
                        else:
                            att_name += p
                
                att_name = sanitize_filename(att_name)
                
                try:
                    att_path = os.path.join(output_dir, att_name)
                    # Gestisce eventuali conflitti di nome negli allegati
                    att_counter = 1
                    base_name, ext = os.path.splitext(att_name)
                    while os.path.exists(att_path):
                        att_path = os.path.join(output_dir, f"{base_name}_{att_counter}{ext}")
                        att_counter += 1
                        
                    payload = part.get_payload(decode=True)
                    if payload:
                        with open(att_path, 'wb') as f:
                            f.write(payload)
                        attachments_count += 1
                        saved_attachments.append(os.path.basename(att_path))
                        self.log(self.t("attach_saved", name=os.path.basename(att_path)))
                except Exception as e:
                    self.log(self.t("attach_save_fail", name=att_name, error=str(e)), is_error=True)

            # 3. Scrittura del file Markdown (nome univoco)
            md_path = os.path.join(output_dir, f"{output_stem}.md")
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(self.t("md_title") + "\n\n")
                f.write(self.t("md_header") + "\n\n")
                f.write(f"- From: {sender}\n")
                f.write(f"- To: {get_decoded_header('To')}\n")
                f.write(f"- Cc: {get_decoded_header('Cc')}\n")
                f.write(f"- Subject: {subject}\n")
                f.write(f"- Date: {readable_date}\n")
                f.write(f"- Message-ID: {get_decoded_header('Message-ID')}\n\n")

                f.write(self.t("md_instructions") + "\n\n")
                f.write(self.t("md_instructions_text") + "\n\n")

                f.write(self.t("md_attachments") + "\n\n")
                if saved_attachments:
                    for a in saved_attachments:
                        f.write(f"- {a}\n")
                else:
                    f.write(self.t("md_no_attachments") + "\n")
                f.write("\n")

                f.write(self.t("md_body") + "\n\n")
                f.write(body_text)

            self.log(self.t("done", count=attachments_count))
            self.log("-" * 40)
            
        except Exception as e:
            error_trace = traceback.format_exc()
            self.log(self.t("fatal_error", error=str(e)), is_error=True)
            logging.error(error_trace)
            QMessageBox.critical(self, self.t("title"), f"{self.t('fatal_error', error=str(e))}\n\n{filename}")

    def process_msg_file(self, file_path, base_dir_override=None):
        if extract_msg is None:
            self.log(self.t("msg_support_missing"), is_error=True)
            return

        try:
            filename = os.path.basename(file_path)
            self.log(self.t("processing_start", name=filename))

            msg = extract_msg.openMsg(file_path)
            sender = msg.sender or "MittenteSconosciuto"
            subject = msg.subject or "NessunOggetto"
            date_str = msg.date

            try:
                if date_str:
                    dt = parsedate_to_datetime(date_str)
                    formatted_date = dt.strftime("%Y-%m-%d_%H%M%S")
                    readable_date = dt.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    raise ValueError("Data assente")
            except Exception:
                formatted_date = datetime.now().strftime("%Y-%m-%d_%H%M%S")
                readable_date = str(date_str) if date_str else "Data sconosciuta"

            folder_name = f"{formatted_date}_{sanitize_filename(sender)}_{sanitize_filename(subject)}"
            base_dir = base_dir_override or os.path.dirname(file_path)
            output_dir = os.path.join(base_dir, folder_name)

            counter = 1
            original_output_dir = output_dir
            while os.path.exists(output_dir):
                output_dir = f"{original_output_dir}_{counter}"
                counter += 1

            os.makedirs(output_dir)
            self.log(self.t("folder_created", path=output_dir))
            output_stem = make_output_stem(formatted_date, sender, subject)

            body_text = msg.body
            html_body = msg.htmlBody

            if html_body:
                try:
                    if isinstance(html_body, bytes):
                        html_body = html_body.decode("utf-8", errors="replace")
                    soup = BeautifulSoup(html_body, "html.parser")
                    body_text = soup.get_text(separator="\n", strip=True)
                except Exception as e:
                    self.log(f"Errore conversione HTML (fallback al testo normale): {str(e)}", is_error=True)

            if not body_text:
                body_text = "Nessun contenuto testuale trovato nell'email."
            elif isinstance(body_text, bytes):
                body_text = body_text.decode("utf-8", errors="replace")

            md_path = os.path.join(output_dir, f"{output_stem}.md")

            attachments_count = 0
            saved_attachments = []
            for attachment in msg.attachments:
                att_name = getattr(attachment, "longFilename", None) or getattr(attachment, "shortFilename", None) or "allegato_sconosciuto.bin"
                att_name = sanitize_filename(att_name)
                try:
                    att_path = os.path.join(output_dir, att_name)
                    att_counter = 1
                    base_name, ext = os.path.splitext(att_name)
                    while os.path.exists(att_path):
                        att_path = os.path.join(output_dir, f"{base_name}_{att_counter}{ext}")
                        att_counter += 1
                    with open(att_path, "wb") as f:
                        f.write(attachment.data)
                    attachments_count += 1
                    saved_attachments.append(os.path.basename(att_path))
                    self.log(self.t("attach_saved", name=os.path.basename(att_path)))
                except Exception as e:
                    self.log(self.t("attach_save_fail", name=att_name, error=str(e)), is_error=True)

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(self.t("md_title") + "\n\n")
                f.write(self.t("md_header") + "\n\n")
                f.write(f"- From: {sender}\n")
                f.write(f"- To: {msg.to or ''}\n")
                f.write(f"- Cc: {msg.cc or ''}\n")
                f.write(f"- Subject: {subject}\n")
                f.write(f"- Date: {readable_date}\n")
                f.write(f"- Message-ID: {getattr(msg, 'messageId', '')}\n\n")

                f.write(self.t("md_instructions") + "\n\n")
                f.write(self.t("md_instructions_text") + "\n\n")

                f.write(self.t("md_attachments") + "\n\n")
                if saved_attachments:
                    for a in saved_attachments:
                        f.write(f"- {a}\n")
                else:
                    f.write(self.t("md_no_attachments") + "\n")
                f.write("\n")

                f.write(self.t("md_body") + "\n\n")
                f.write(body_text)

            msg.close()
            self.log(self.t("done", count=attachments_count))
            self.log("-" * 40)

        except Exception as e:
            error_trace = traceback.format_exc()
            self.log(self.t("fatal_error", error=str(e)), is_error=True)
            logging.error(error_trace)
            QMessageBox.critical(self, self.t("title"), f"{self.t('fatal_error', error=str(e))}\n\n{filename}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion") # Stile pulito e moderno
    window = EmailExtractorApp()
    window.show()
    sys.exit(app.exec())
