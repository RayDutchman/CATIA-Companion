import sys
import shutil
import winreg
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QMessageBox, QDialog, QPushButton, QListWidget, QFileDialog,
    QAbstractItemView, QRadioButton, QButtonGroup, QLineEdit, QGroupBox,
    QListWidgetItem
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSettings


def resource_path(filename: str) -> Path:
    """
    Returns the correct path to a resource file.
    - When running as a PyInstaller .exe: looks next to the .exe
    - When running as a script: looks next to main.py
    """
    if hasattr(sys, "_MEIPASS"):
        # Running as PyInstaller bundle — look next to the .exe
        return Path(sys.executable).parent / filename
    return Path(__file__).parent / filename


# ---------------------------------------------------------------------------
# App info
# ---------------------------------------------------------------------------

APP_NAME    = "CATIA Companion"
APP_VERSION = "1.0.0"
APP_DATE    = "2026-04-03"
APP_AUTHOR  = "CHEN Weibo"
APP_CONTACT = "thucwb@gmail.com"

ABOUT_TEXT = f"""{APP_NAME} v{APP_VERSION}

A CATIA V5 productivity tool for engineering teams.
Automates drawing conversion, part export, and
installation of CATIA resources.

─────────────────────────────────────────
Developer   {APP_AUTHOR}
Contact     {APP_CONTACT}
Released    {APP_DATE}
─────────────────────────────────────────

\u00a9 2026 {APP_AUTHOR}. For internal use only."""


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # --- Window settings ---
        self.setWindowTitle("CATIA Companion")
        self.resize(600, 400)

        # --- Menu bar ---
        self._setup_menu_bar()

        # --- Central widget & layout ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # --- Placeholder content ---
        label = QLabel("Welcome to CATIA Companion")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)

        # --- Status bar ---
        self.statusBar().showMessage("Ready")

    def _setup_menu_bar(self):
        menu_bar = self.menuBar()

        # --- File ---
        file_menu = menu_bar.addMenu("File")
        file_menu.addAction(QAction("New", self))
        file_menu.addAction(QAction("Open...", self))
        file_menu.addAction(QAction("Save", self))
        file_menu.addAction(QAction("Save As...", self))
        file_menu.addSeparator()

        # --- Convert submenu ---
        convert_menu = file_menu.addMenu("Convert")
        convert_part_action = QAction("Convert CATPart/CATProduct", self)
        convert_part_action.triggered.connect(self._open_convert_part_dialog)
        convert_menu.addAction(convert_part_action)
        convert_drawing_action = QAction("Convert CATDrawing", self)
        convert_drawing_action.triggered.connect(self._open_convert_drawing_dialog)
        convert_menu.addAction(convert_drawing_action)

        # --- Export BOM ---
        export_bom_action = QAction("Export BOM from CATProduct", self)
        export_bom_action.triggered.connect(self._open_export_bom_dialog)
        file_menu.addAction(export_bom_action)

        file_menu.addSeparator()
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        # --- Edit ---
        edit_menu = menu_bar.addMenu("Edit")
        edit_menu.addAction(QAction("Undo", self))
        edit_menu.addAction(QAction("Redo", self))
        edit_menu.addSeparator()
        edit_menu.addAction(QAction("Cut", self))
        edit_menu.addAction(QAction("Copy", self))
        edit_menu.addAction(QAction("Paste", self))

        # --- Tools ---
        tools_menu = menu_bar.addMenu("Tools")
        copy_font_action = QAction("Copy Font File to CATIA folder", self)
        copy_font_action.triggered.connect(self._copy_font_to_catia)
        tools_menu.addAction(copy_font_action)
        copy_iso_action = QAction("Copy ISO.xml File to CATIA folder", self)
        copy_iso_action.triggered.connect(self._copy_iso_to_catia)
        tools_menu.addAction(copy_iso_action)
        pojie_action = QAction("PoJie", self)
        pojie_action.triggered.connect(self._pojie)
        tools_menu.addAction(pojie_action)

        # --- View ---
        view_menu = menu_bar.addMenu("View")
        view_menu.addAction(QAction("Zoom In", self))
        view_menu.addAction(QAction("Zoom Out", self))
        view_menu.addAction(QAction("Reset Zoom", self))
        view_menu.addSeparator()
        view_menu.addAction(QAction("Toggle Status Bar", self))

        # --- Help ---
        help_menu = menu_bar.addMenu("Help")
        help_menu.addAction(QAction("Documentation", self))
        about_action = QAction("About CATIA Companion", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    # ------------------------------------------------------------------ #
    # Help menu
    # ------------------------------------------------------------------ #

    def _show_about(self):
        QMessageBox.about(self, f"About {APP_NAME}", ABOUT_TEXT)

    # ------------------------------------------------------------------ #
    # File > Convert menu
    # ------------------------------------------------------------------ #

    def _open_convert_part_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="Convert CATPart/CATProduct to STEP",
            file_label="Selected CATPart/CATProduct files:",
            file_filter="CATIA Part/Product Files (*.CATPart *.CATProduct);;All Files (*)",
            no_files_msg="Please select at least one CATPart or CATProduct file.",
            conversion_fn=CATPart_to_STP,
            settings_key="CATPart"
        )
        dialog.exec()

    def _open_convert_drawing_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="Convert CATDrawing to PDF",
            file_label="Selected CATDrawing files:",
            file_filter="CATDrawing Files (*.CATDrawing);;All Files (*)",
            no_files_msg="Please select at least one CATDrawing file.",
            conversion_fn=CATDrawing_to_PDF,
            settings_key="CATDrawing"
        )
        dialog.exec()

    def _open_export_bom_dialog(self):
        dialog = ExportBOMDialog(self)
        dialog.exec()

    # ------------------------------------------------------------------ #
    # Tools menu
    # ------------------------------------------------------------------ #

    def _copy_font_to_catia(self):
        self._copy_file_to_catia(
            file_name="Changfangsong.ttf",
            relative_dest=Path("win_b64") / "resources" / "fonts" / "TrueType"
        )

    def _copy_iso_to_catia(self):
        self._copy_file_to_catia(
            file_name="ISO.xml",
            relative_dest=Path("win_b64") / "resources" / "standard" / "drafting"
        )

    def _copy_file_to_catia(self, file_name: str, relative_dest: Path):
        src_file = resource_path(file_name)

        if not src_file.exists():
            QMessageBox.warning(
                self, "File Not Found",
                f"Could not find '{file_name}' in the working folder:\n{src_file.parent}"
            )
            return

        catia_root = detect_catia_root()

        if catia_root:
            reply = QMessageBox.question(
                self, "CATIA Installation Detected",
                f"CATIA installation found at:\n{catia_root}\n\nUse this folder?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(
                self,
                "Select CATIA Installation Folder (e.g. C:\\Program Files\\Dassault Systemes\\B28)",
                ""
            )
            if not catia_root:
                return

        dest_dir = Path(catia_root) / relative_dest

        if not dest_dir.exists():
            reply = QMessageBox.question(
                self, "Folder Not Found",
                f"The target folder does not exist:\n{dest_dir}\n\nDo you want to create it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                dest_dir.mkdir(parents=True, exist_ok=True)
            else:
                return

        dest_file = dest_dir / file_name

        try:
            shutil.copy2(str(src_file), str(dest_file))
            QMessageBox.information(
                self, "Success",
                f"'{file_name}' has been copied to:\n{dest_file}"
            )
        except PermissionError:
            QMessageBox.critical(
                self, "Permission Denied",
                f"Could not copy the file. Try running the application as Administrator.\n\nTarget:\n{dest_file}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")

    def _pojie(self):
        src_dir = resource_path("Pojie")

        # Check source folder exists
        if not src_dir.exists() or not src_dir.is_dir():
            QMessageBox.warning(
                self, "Folder Not Found",
                f"Could not find the 'Pojie' folder at:\n{src_dir.parent}"
            )
            return

        # Auto-detect or manually select CATIA root
        catia_root = detect_catia_root()

        if catia_root:
            reply = QMessageBox.question(
                self, "CATIA Installation Detected",
                f"CATIA installation found at:\n{catia_root}\n\nUse this folder?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(
                self,
                "Select CATIA Installation Folder (e.g. C:\\Program Files\\Dassault Systemes\\B28)",
                ""
            )
            if not catia_root:
                return

        dest_dir = Path(catia_root) / "win_b64" / "code" / "bin"

        if not dest_dir.exists():
            QMessageBox.critical(
                self, "Folder Not Found",
                f"The target folder does not exist:\n{dest_dir}\n\nPlease check your CATIA installation."
            )
            return

        # Copy all files from Pojie folder, overwriting existing ones
        files = [f for f in src_dir.iterdir() if f.is_file()]
        if not files:
            QMessageBox.warning(self, "Empty Folder", "The 'Pojie' folder contains no files.")
            return

        try:
            copied = []
            for src_file in files:
                dest_file = dest_dir / src_file.name
                shutil.copy2(str(src_file), str(dest_file))
                copied.append(src_file.name)
                print(f"  Copied: {src_file.name} -> {dest_file}")

            QMessageBox.information(
                self, "Success",
                f"Successfully copied {len(copied)} file(s) to:\n{dest_dir}\n\n" +
                "\n".join(copied)
            )
        except PermissionError:
            QMessageBox.critical(
                self, "Permission Denied",
                f"Could not copy files. Try running the application as Administrator.\n\nTarget:\n{dest_dir}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")


# ---------------------------------------------------------------------------
# Generic Convert Dialog
# ---------------------------------------------------------------------------

class ConvertDialog(QDialog):
    def __init__(self, parent=None, title="Convert", file_label="Selected files:",
                 file_filter="All Files (*)", no_files_msg="Please select at least one file.",
                 conversion_fn=None, settings_key="default"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(520, 450)
        self._file_filter = file_filter
        self._no_files_msg = no_files_msg
        self._conversion_fn = conversion_fn

        # QSettings persists values between sessions under a key per dialog type
        self._settings = QSettings("CATIACompanion", f"ConvertDialog_{settings_key}")
        self._last_browse_dir = self._settings.value("last_browse_dir", "")
        self._last_output_dir = self._settings.value("last_output_dir", "")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # --- Input files ---
        layout.addWidget(QLabel(file_label))

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        layout.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_files)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_selected)
        btn_row.addWidget(browse_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # --- Output folder ---
        output_group = QGroupBox("Output Folder")
        output_layout = QVBoxLayout(output_group)

        self.radio_same = QRadioButton("Same folder as source files")
        self.radio_custom = QRadioButton("Choose a folder:")
        self.radio_same.setChecked(True)

        self.btn_group = QButtonGroup(self)
        self.btn_group.addButton(self.radio_same)
        self.btn_group.addButton(self.radio_custom)

        output_layout.addWidget(self.radio_same)
        output_layout.addWidget(self.radio_custom)

        folder_row = QHBoxLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("Select output folder...")
        self.folder_edit.setReadOnly(True)
        self.folder_edit.setEnabled(False)
        self.folder_browse_btn = QPushButton("Browse...")
        self.folder_browse_btn.setEnabled(False)
        self.folder_browse_btn.clicked.connect(self._browse_output_folder)
        folder_row.addWidget(self.folder_edit)
        folder_row.addWidget(self.folder_browse_btn)
        output_layout.addLayout(folder_row)

        self.radio_custom.toggled.connect(self._toggle_folder_row)

        layout.addWidget(output_group)

        # Restore last output folder if saved
        if self._last_output_dir:
            self.radio_custom.setChecked(True)
            self.folder_edit.setText(self._last_output_dir)

        # --- Confirm / Cancel ---
        action_row = QHBoxLayout()
        action_row.addStretch()
        confirm_btn = QPushButton("Confirm")
        confirm_btn.setDefault(True)
        confirm_btn.clicked.connect(self._confirm)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(confirm_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    def _toggle_folder_row(self, checked):
        self.folder_edit.setEnabled(checked)
        self.folder_browse_btn.setEnabled(checked)

    def _browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", self._last_browse_dir, self._file_filter
        )
        if files:
            # Save the directory of the first selected file
            self._last_browse_dir = str(Path(files[0]).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)
        for f in files:
            existing = [self.file_list.item(i).text() for i in range(self.file_list.count())]
            if f not in existing:
                self.file_list.addItem(f)

    def _remove_selected(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def _browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder", self._last_output_dir
        )
        if folder:
            self.folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    def _confirm(self):
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not files:
            QMessageBox.warning(self, "No Files", self._no_files_msg)
            return

        if self.radio_same.isChecked():
            output_folder = None
        else:
            output_folder = self.folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(self, "No Output Folder", "Please select an output folder.")
                return

        if self._conversion_fn:
            self._conversion_fn(files, output_folder)
        self.accept()


# ---------------------------------------------------------------------------
# CATIA installation detector
# ---------------------------------------------------------------------------

def detect_catia_root() -> str | None:
    """
    Try to find the CATIA V5 installation root from the Windows Registry.
    Looks for DEST_FOLDER under HKEY_LOCAL_MACHINE\\SOFTWARE\\Dassault Systemes\\<release>\\0.
    Returns the path string if found, or None if not detected.
    """
    registry_paths = [
        r"SOFTWARE\Dassault Systemes",
        r"SOFTWARE\WOW6432Node\Dassault Systemes",
    ]

    for reg_path in registry_paths:
        print(f"Trying registry path: HKEY_LOCAL_MACHINE\\{reg_path}")
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as ds_key:
                i = 0
                while True:
                    try:
                        release = winreg.EnumKey(ds_key, i)
                        full_key = rf"HKEY_LOCAL_MACHINE\{reg_path}\{release}\0"
                        print(f"  Trying key: {full_key}")
                        try:
                            with winreg.OpenKey(ds_key, rf"{release}\0") as release_key:
                                try:
                                    install_path, _ = winreg.QueryValueEx(release_key, "DEST_FOLDER")
                                    print(f"    Found DEST_FOLDER: {install_path}")
                                    candidate = Path(install_path)
                                    win_b64 = candidate / "win_b64"
                                    print(f"    Checking win_b64 exists: {win_b64}")
                                    if win_b64.exists():
                                        print(f"    -> Valid CATIA installation found: {candidate}")
                                        return str(candidate)
                                    else:
                                        print(f"    -> win_b64 not found, skipping.")
                                except FileNotFoundError:
                                    print(f"    -> DEST_FOLDER value not found in this key.")
                        except OSError:
                            print(f"    -> Could not open subkey \\0 under {release}.")
                        i += 1
                    except OSError:
                        break  # No more subkeys
        except OSError:
            print(f"  -> Registry path not found, skipping.")

    print("No valid CATIA installation detected.")
    return None


# ---------------------------------------------------------------------------
# Conversion function
# ---------------------------------------------------------------------------

def CATDrawing_to_PDF(file_paths: list[str], output_folder: str | None = None):
    """
    Convert CATDrawing files to PDF using pyCATIA.
    Single sheet: saved as <file>.pdf
    Multiple sheets: saved as <file>_Sheet1.pdf, <file>_Sheet2.pdf, ...
    CATIA remains visible during processing.
    """
    from pycatia import catia

    caa = catia()
    application = caa.application
    application.visible = True

    documents = application.documents

    for path in file_paths:
        src = Path(path).resolve()
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent
        dest_dir.mkdir(parents=True, exist_ok=True)

        print(f"Opening: {src}")

        documents.open(str(src))
        from pycatia.drafting_interfaces.drawing_document import DrawingDocument
        drawing_doc = DrawingDocument(application.active_document.com_object)
        drawing = drawing_doc.drawing_root

        sheets = drawing.sheets
        sheet_count = sheets.count

        dest = dest_dir / f"{src.stem}.pdf"
        drawing_doc.export_data(str(dest), "pdf")
        if not dest.exists():
            print(f"  WARNING: export_data did not create {dest}")
        else:
            print(f"  Exported {sheet_count} sheet(s) -> {dest}")

        drawing_doc.close()
        print(f"Done: {src.name}\n")


def CATPart_to_STP(file_paths: list[str], output_folder: str | None = None):
    """
    Convert CATPart/CATProduct files to STEP (.stp) using pyCATIA.
    CATIA remains visible during processing.
    """
    from pycatia import catia

    caa = catia()
    application = caa.application
    application.visible = True

    documents = application.documents

    for path in file_paths:
        src = Path(path)
        dest_dir = Path(output_folder) if output_folder else src.parent
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"{src.stem}.stp"

        print(f"Opening: {src}")

        documents.open(str(src))
        doc = application.active_document

        doc.export_data(str(dest), "stp")
        print(f"  Exported -> {dest}")

        doc.close()
        print(f"Done: {src.name}\n")


# ---------------------------------------------------------------------------
# Export BOM Dialog
# ---------------------------------------------------------------------------

# All available BOM columns
BOM_ALL_COLUMNS     = ["Level", "Part Number", "Nomenclature", "Definition", "Revision", "Source", "Material", "Quantity"]
BOM_DEFAULT_COLUMNS = ["Level", "Part Number", "Nomenclature", "Definition", "Revision", "Source", "Material", "Quantity"]

class ExportBOMDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Export BOM from CATProduct")
        self.setMinimumSize(560, 520)

        self._settings = QSettings("CATIACompanion", "ExportBOMDialog")
        self._last_browse_dir = self._settings.value("last_browse_dir", "")
        self._last_output_dir = self._settings.value("last_output_dir", "")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # --- CATProduct file ---
        layout.addWidget(QLabel("CATProduct file:"))
        file_row = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select a CATProduct file...")
        self.file_edit.setReadOnly(True)
        file_browse_btn = QPushButton("Browse...")
        file_browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self.file_edit)
        file_row.addWidget(file_browse_btn)
        layout.addLayout(file_row)

        # --- Output folder ---
        output_group = QGroupBox("Output Folder")
        output_layout = QVBoxLayout(output_group)
        self.radio_same = QRadioButton("Same folder as source file")
        self.radio_custom = QRadioButton("Choose a folder:")
        self.radio_same.setChecked(True)
        self.btn_group = QButtonGroup(self)
        self.btn_group.addButton(self.radio_same)
        self.btn_group.addButton(self.radio_custom)
        output_layout.addWidget(self.radio_same)
        output_layout.addWidget(self.radio_custom)
        folder_row = QHBoxLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("Select output folder...")
        self.folder_edit.setReadOnly(True)
        self.folder_edit.setEnabled(False)
        self.folder_browse_btn = QPushButton("Browse...")
        self.folder_browse_btn.setEnabled(False)
        self.folder_browse_btn.clicked.connect(self._browse_output_folder)
        folder_row.addWidget(self.folder_edit)
        folder_row.addWidget(self.folder_browse_btn)
        output_layout.addLayout(folder_row)
        self.radio_custom.toggled.connect(self._toggle_folder_row)
        layout.addWidget(output_group)

        # Restore last output folder
        if self._last_output_dir:
            self.radio_custom.setChecked(True)
            self.folder_edit.setText(self._last_output_dir)

        # --- Column selector ---
        col_group = QGroupBox("Columns to Export (drag to reorder)")
        col_layout = QHBoxLayout(col_group)

        # Available columns (left)
        avail_layout = QVBoxLayout()
        avail_layout.addWidget(QLabel("Available:"))
        self.avail_list = QListWidget()
        self.avail_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.avail_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        avail_layout.addWidget(self.avail_list)
        col_layout.addLayout(avail_layout)

        # Arrow buttons (center)
        arrow_layout = QVBoxLayout()
        arrow_layout.addStretch()
        add_btn = QPushButton("→")
        add_btn.setFixedWidth(36)
        add_btn.clicked.connect(self._add_column)
        remove_btn = QPushButton("←")
        remove_btn.setFixedWidth(36)
        remove_btn.clicked.connect(self._remove_column)
        up_btn = QPushButton("↑")
        up_btn.setFixedWidth(36)
        up_btn.clicked.connect(self._move_up)
        down_btn = QPushButton("↓")
        down_btn.setFixedWidth(36)
        down_btn.clicked.connect(self._move_down)
        arrow_layout.addWidget(add_btn)
        arrow_layout.addWidget(remove_btn)
        arrow_layout.addSpacing(10)
        arrow_layout.addWidget(up_btn)
        arrow_layout.addWidget(down_btn)
        arrow_layout.addStretch()
        col_layout.addLayout(arrow_layout)

        # Selected columns (right)
        selected_layout = QVBoxLayout()
        selected_layout.addWidget(QLabel("Selected:"))
        self.selected_list = QListWidget()
        self.selected_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        selected_layout.addWidget(self.selected_list)
        col_layout.addLayout(selected_layout)

        layout.addWidget(col_group)

        # Populate lists from saved settings
        saved = self._settings.value("selected_columns", BOM_DEFAULT_COLUMNS)
        if isinstance(saved, str):
            saved = [saved]
        for col in saved:
            self.selected_list.addItem(QListWidgetItem(col))
        for col in BOM_ALL_COLUMNS:
            if col not in saved:
                self.avail_list.addItem(QListWidgetItem(col))

        # --- Confirm / Cancel ---
        action_row = QHBoxLayout()
        action_row.addStretch()
        confirm_btn = QPushButton("Export")
        confirm_btn.setDefault(True)
        confirm_btn.clicked.connect(self._confirm)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(confirm_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    def _toggle_folder_row(self, checked):
        self.folder_edit.setEnabled(checked)
        self.folder_browse_btn.setEnabled(checked)

    def _browse_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select CATProduct File", self._last_browse_dir,
            "CATProduct Files (*.CATProduct);;All Files (*)"
        )
        if file:
            self.file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)

    def _browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder", self._last_output_dir
        )
        if folder:
            self.folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    def _add_column(self):
        for item in self.avail_list.selectedItems():
            self.avail_list.takeItem(self.avail_list.row(item))
            self.selected_list.addItem(QListWidgetItem(item.text()))

    def _remove_column(self):
        for item in self.selected_list.selectedItems():
            self.selected_list.takeItem(self.selected_list.row(item))
            self.avail_list.addItem(QListWidgetItem(item.text()))

    def _move_up(self):
        row = self.selected_list.currentRow()
        if row > 0:
            item = self.selected_list.takeItem(row)
            self.selected_list.insertItem(row - 1, item)
            self.selected_list.setCurrentRow(row - 1)

    def _move_down(self):
        row = self.selected_list.currentRow()
        if row < self.selected_list.count() - 1:
            item = self.selected_list.takeItem(row)
            self.selected_list.insertItem(row + 1, item)
            self.selected_list.setCurrentRow(row + 1)

    def _confirm(self):
        file_path = self.file_edit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "No File", "Please select a CATProduct file.")
            return

        selected_cols = [
            self.selected_list.item(i).text()
            for i in range(self.selected_list.count())
        ]
        if not selected_cols:
            QMessageBox.warning(self, "No Columns", "Please select at least one column to export.")
            return

        # Save column selection
        self._settings.setValue("selected_columns", selected_cols)

        if self.radio_same.isChecked():
            output_folder = None
        else:
            output_folder = self.folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(self, "No Output Folder", "Please select an output folder.")
                return

        export_bom_to_excel([file_path], output_folder, columns=selected_cols)
        self.accept()


# ---------------------------------------------------------------------------
# BOM export function
# ---------------------------------------------------------------------------

def export_bom_to_excel(file_paths: list[str], output_folder: str | None = None,
                        columns: list[str] | None = None):
    """
    Export a hierarchical BOM from CATProduct files to Excel (.xlsx).
    Columns are user-defined and ordered.
    CATIA remains visible during processing.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment
    from pycatia import catia
    from pycatia.product_structure_interfaces.product_document import ProductDocument

    if columns is None:
        columns = BOM_DEFAULT_COLUMNS

    caa = catia()
    application = caa.application
    application.visible = True

    documents = application.documents

    # Direct attribute names on the pyCATIA product object
    DIRECT_ATTR_MAP = {
        "Nomenclature": "nomenclature",
        "Revision":     "revision",
        "Definition":   "definition",
        "Source":       "source",
        "Material":     "material",
    }

    def get_property(product, name: str) -> str:
        """Read a standard CATIA property from a product or part."""
        attr = DIRECT_ATTR_MAP.get(name)
        if not attr:
            return ""

        # Build list of targets to try, in priority order
        targets = [product]
        try:
            targets.insert(0, product.reference_product)
        except Exception:
            pass

        for target in targets:
            # Try direct attribute on the product/reference_product
            try:
                value = getattr(target, attr)
                if value:
                    return str(value)
            except Exception:
                pass

            # Try via get_item("Part") for CATPart objects
            try:
                part = target.get_item("Part")
                value = getattr(part, attr)
                if value:
                    return str(value)
            except Exception:
                pass

        return ""

    def get_material(product) -> str:
        """Read the material from a product using the Chinese property name."""
        return get_property(product, "Material")

    def traverse(product, rows: list, level: int):
        """Recursively traverse the assembly tree and collect BOM rows."""
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn = name.rsplit(".", 1)[0] if "." in name else name
            print(f"    part_number failed, using instance name: {pn}")
        row = {"Level": level, "Part Number": pn}
        #print(row)#截止到这里row有两个内容，有零件号
        print(f"  {'  ' * level}[Level {level}] {pn}")

        for col in columns:
            if col in DIRECT_ATTR_MAP:
                row[col] = get_property(product, col)#到这里之后Part Number变成了空值

        rows.append(row)

        try:
            products = product.products
            count = products.count
            print(f"  {'  ' * level}  -> {count} child(ren) found")

            if count == 0:
                # This is a leaf CATPart — no children to recurse into
                return

            # Group children by part number to compute quantities
            children = {}
            for i in range(1, count + 1):
                try:
                    child = products.item(i)
                    try:
                        pn = child.part_number
                    except Exception:
                        try:
                            pn = child.reference_product.part_number
                        except Exception:
                            name = child.name
                            pn = name.rsplit(".", 1)[0] if "." in name else name
                            print(f"  {'  ' * level}  -> Using instance name as part number: {pn}")
                except Exception as e:
                    print(f"  {'  ' * level}  -> Skipping child {i}: {e}")
                    continue
                if pn not in children:
                    children[pn] = {"products": child, "qty": 0}
                children[pn]["qty"] += 1

            for pn, data in children.items():
                child_rows = []
                traverse(data["products"], child_rows, level + 1)
                if child_rows:
                    child_rows[0]["Quantity"] = data["qty"]
                rows.extend(child_rows)

        except Exception as e:
            print(f"  {'  ' * level}  -> Exception accessing children: {e}")

    for path in file_paths:
        src = Path(path).resolve()
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"{src.stem}_BOM.xlsx"

        # Check if the file is already open in Excel
        if dest.exists():
            try:
                # Try opening the file exclusively to check if it's locked
                with open(dest, "a+b"):
                    pass
            except PermissionError:
                from PySide6.QtWidgets import QMessageBox
                reply = QMessageBox.question(
                    None,
                    "File In Use",
                    f"The file is currently open in Excel:\n{dest}\n\n"
                    f"Please close it in Excel, then click Retry, or Cancel to abort.",
                    QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel
                )
                if reply == QMessageBox.StandardButton.Cancel:
                    print(f"  Aborted: {dest} is open in Excel.")
                    continue
                # Retry check
                try:
                    with open(dest, "a+b"):
                        pass
                except PermissionError:
                    QMessageBox.critical(
                        None, "Still In Use",
                        f"The file is still open. Please close it and try again.\n{dest}"
                    )
                    continue

        print(f"Opening: {src}")
        documents.open(str(src))
        product_doc = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product

        rows = []
        traverse(root_product, rows, level=0)

        # Write to Excel — plain, no formatting
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BOM"

        center = Alignment(horizontal="center")

        # Header row — bold only
        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(bold=True)

        # Data rows
        for row_idx, row in enumerate(rows, start=2):
            level = row.get("Level", 0)
            for col_idx, col_name in enumerate(columns, start=1):
                if col_name == "Level":
                    value = level
                elif col_name == "Quantity":
                    value = row.get("Quantity", 1)
                else:
                    value = row.get(col_name, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if col_name in ("Level", "Quantity"):
                    cell.alignment = center

        # Auto column widths based on content
        for col_idx, col_name in enumerate(columns, start=1):
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            max_width = len(col_name)
            for row_idx in range(2, ws.max_row + 1):
                cell_val = ws.cell(row=row_idx, column=col_idx).value
                if cell_val is not None:
                    max_width = max(max_width, len(str(cell_val)))
            ws.column_dimensions[col_letter].width = max_width

        wb.save(str(dest))
        print(f"  BOM exported -> {dest}")

        product_doc.close()
        print(f"Done: {src.name}\n")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("CATIA Companion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()