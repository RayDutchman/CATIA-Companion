import sys
import shutil
import winreg
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QMessageBox, QDialog, QPushButton, QListWidget, QFileDialog,
    QAbstractItemView, QRadioButton, QButtonGroup, QLineEdit, QGroupBox,
    QListWidgetItem, QComboBox, QCheckBox
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSettings


def resource_path(filename: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
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
# Part template properties
# ---------------------------------------------------------------------------

PART_TEMPLATE_PROPERTIES = ["物料编码", "物料名称", "中文名称", "规格型号", "物料来源", "数据状态", "存货类别", "质量", "备注"]


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CATIA Companion")
        self.resize(600, 400)
        self._setup_menu_bar()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        label = QLabel("Welcome to CATIA Companion")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
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

        convert_menu = file_menu.addMenu("Convert")
        convert_part_action = QAction("Convert CATPart/CATProduct", self)
        convert_part_action.triggered.connect(self._open_convert_part_dialog)
        convert_menu.addAction(convert_part_action)
        convert_drawing_action = QAction("Convert CATDrawing", self)
        convert_drawing_action.triggered.connect(self._open_convert_drawing_dialog)
        convert_menu.addAction(convert_drawing_action)

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
        stamp_action = QAction("刷写零件模板", self)
        stamp_action.triggered.connect(self._open_stamp_part_template_dialog)
        tools_menu.addAction(stamp_action)

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

    def _show_about(self):
        QMessageBox.about(self, f"About {APP_NAME}", ABOUT_TEXT)

    def _open_convert_part_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="Convert CATPart/CATProduct to STEP",
            file_label="Selected CATPart/CATProduct files:",
            file_filter="CATIA Part/Product Files (*.CATPart *.CATProduct);;All Files (*)",
            no_files_msg="Please select at least one CATPart or CATProduct file.",
            conversion_fn=CATPart_to_STP,
            settings_key="CATPart",
            show_prefix_option=True,
            prefix="MD_"
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
            settings_key="CATDrawing",
            show_prefix_option=True,
            prefix="DR_"
        )
        dialog.exec()

    def _open_export_bom_dialog(self):
        dialog = ExportBOMDialog(self)
        dialog.exec()

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
            QMessageBox.warning(self, "File Not Found",
                f"Could not find '{file_name}' in the working folder:\n{src_file.parent}")
            return

        catia_root = detect_catia_root()
        if catia_root:
            reply = QMessageBox.question(self, "CATIA Installation Detected",
                f"CATIA installation found at:\n{catia_root}\n\nUse this folder?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(self,
                "Select CATIA Installation Folder (e.g. C:\\Program Files\\Dassault Systemes\\B28)", "")
            if not catia_root:
                return

        dest_dir = Path(catia_root) / relative_dest
        if not dest_dir.exists():
            reply = QMessageBox.question(self, "Folder Not Found",
                f"The target folder does not exist:\n{dest_dir}\n\nDo you want to create it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                dest_dir.mkdir(parents=True, exist_ok=True)
            else:
                return

        dest_file = dest_dir / file_name
        try:
            shutil.copy2(str(src_file), str(dest_file))
            QMessageBox.information(self, "Success",
                f"'{file_name}' has been copied to:\n{dest_file}")
        except PermissionError:
            QMessageBox.critical(self, "Permission Denied",
                f"Could not copy the file. Try running the application as Administrator.\n\nTarget:\n{dest_file}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")

    def _pojie(self):
        src_dir = resource_path("Pojie")
        if not src_dir.exists() or not src_dir.is_dir():
            QMessageBox.warning(self, "Folder Not Found",
                f"Could not find the 'Pojie' folder at:\n{src_dir.parent}")
            return

        catia_root = detect_catia_root()
        if catia_root:
            reply = QMessageBox.question(self, "CATIA Installation Detected",
                f"CATIA installation found at:\n{catia_root}\n\nUse this folder?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(self,
                "Select CATIA Installation Folder (e.g. C:\\Program Files\\Dassault Systemes\\B28)", "")
            if not catia_root:
                return

        dest_dir = Path(catia_root) / "win_b64" / "code" / "bin"
        if not dest_dir.exists():
            QMessageBox.critical(self, "Folder Not Found",
                f"The target folder does not exist:\n{dest_dir}\n\nPlease check your CATIA installation.")
            return

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
            QMessageBox.information(self, "Success",
                f"Successfully copied {len(copied)} file(s) to:\n{dest_dir}\n\n" + "\n".join(copied))
        except PermissionError:
            QMessageBox.critical(self, "Permission Denied",
                f"Could not copy files. Try running the application as Administrator.\n\nTarget:\n{dest_dir}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n{e}")

    def _open_stamp_part_template_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="刷写零件模板",
            file_label="Selected CATPart files:",
            file_filter="CATIA Part Files (*.CATPart);;All Files (*)",
            no_files_msg="Please select at least one CATPart file.",
            conversion_fn=stamp_part_template,
            settings_key="StampPartTemplate"
        )
        dialog.exec()


# ---------------------------------------------------------------------------
# Generic Convert Dialog
# ---------------------------------------------------------------------------

class ConvertDialog(QDialog):
    def __init__(self, parent=None, title="Convert", file_label="Selected files:",
                 file_filter="All Files (*)", no_files_msg="Please select at least one file.",
                 conversion_fn=None, settings_key="default",
                 show_prefix_option=False, prefix=""):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(520, 450)
        self._file_filter        = file_filter
        self._no_files_msg       = no_files_msg
        self._conversion_fn      = conversion_fn
        self._show_prefix_option = show_prefix_option
        self._prefix             = prefix

        self._settings = QSettings("CATIACompanion", f"ConvertDialog_{settings_key}")
        self._last_browse_dir = self._settings.value("last_browse_dir", "")
        self._last_output_dir = self._settings.value("last_output_dir", "")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        layout.addWidget(QLabel(file_label))
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        # Restore previously saved file list
        saved_files: list = self._settings.value("saved_files", []) or []
        if isinstance(saved_files, str):
            saved_files = [saved_files]
        for f in saved_files:
            if Path(f).exists():
                self.file_list.addItem(f)

        layout.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_files)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_selected)
        remove_all_btn = QPushButton("Remove All")
        remove_all_btn.clicked.connect(self._remove_all)
        btn_row.addWidget(browse_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addWidget(remove_all_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Output folder — hidden for stamp dialog
        if settings_key != "StampPartTemplate":
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
            if self._last_output_dir:
                self.radio_custom.setChecked(True)
                self.folder_edit.setText(self._last_output_dir)
        else:
            self.radio_same  = None
            self.folder_edit = None

        # Prefix and suffix rows — shown when show_prefix_option=True
        if show_prefix_option:
            saved_add_prefix = self._settings.value("add_prefix", True)
            if isinstance(saved_add_prefix, str):
                saved_add_prefix = saved_add_prefix.lower() == "true"
            saved_prefix_value = self._settings.value("prefix_value", prefix)

            prefix_row = QHBoxLayout()
            self.prefix_checkbox = QCheckBox("Add prefix:")
            self.prefix_checkbox.setChecked(saved_add_prefix)
            self.prefix_edit = QLineEdit(saved_prefix_value)
            self.prefix_edit.setEnabled(saved_add_prefix)
            self.prefix_checkbox.toggled.connect(self.prefix_edit.setEnabled)
            prefix_row.addWidget(self.prefix_checkbox)
            prefix_row.addWidget(self.prefix_edit)
            layout.addLayout(prefix_row)

            saved_add_suffix = self._settings.value("add_suffix", False)
            if isinstance(saved_add_suffix, str):
                saved_add_suffix = saved_add_suffix.lower() == "true"
            saved_suffix_value = self._settings.value("suffix_value", "")

            suffix_row = QHBoxLayout()
            self.suffix_checkbox = QCheckBox("Add suffix:")
            self.suffix_checkbox.setChecked(saved_add_suffix)
            self.suffix_edit = QLineEdit(saved_suffix_value)
            self.suffix_edit.setEnabled(saved_add_suffix)
            self.suffix_checkbox.toggled.connect(self.suffix_edit.setEnabled)
            suffix_row.addWidget(self.suffix_checkbox)
            suffix_row.addWidget(self.suffix_edit)
            layout.addLayout(suffix_row)
        else:
            self.prefix_checkbox = None
            self.prefix_edit     = None
            self.suffix_checkbox = None
            self.suffix_edit     = None

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
            self, "Select Files", self._last_browse_dir, self._file_filter)
        if files:
            self._last_browse_dir = str(Path(files[0]).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)
        for f in files:
            existing = [self.file_list.item(i).text() for i in range(self.file_list.count())]
            if f not in existing:
                self.file_list.addItem(f)
        self._persist_file_list()

    def _remove_selected(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self._persist_file_list()

    def _remove_all(self):
        self.file_list.clear()
        self._persist_file_list()

    def _persist_file_list(self):
        """Save the current file list to QSettings so it survives dialog re-opens."""
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        self._settings.setValue("saved_files", files)

    def _browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder", self._last_output_dir)
        if folder:
            self.folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    def _confirm(self):
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not files:
            QMessageBox.warning(self, "No Files", self._no_files_msg)
            return

        if self.radio_same is None:
            output_folder = None
        elif self.radio_same.isChecked():
            output_folder = None
        else:
            output_folder = self.folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(self, "No Output Folder", "Please select an output folder.")
                return

        if self.prefix_checkbox is not None:
            prefix_value = self.prefix_edit.text() if self.prefix_checkbox.isChecked() else ""
            suffix_value = self.suffix_edit.text() if self.suffix_checkbox.isChecked() else ""
            self._settings.setValue("add_prefix", self.prefix_checkbox.isChecked())
            self._settings.setValue("prefix_value", self.prefix_edit.text())
            self._settings.setValue("add_suffix", self.suffix_checkbox.isChecked())
            self._settings.setValue("suffix_value", self.suffix_edit.text())
            self._conversion_fn(files, output_folder, prefix=prefix_value, suffix=suffix_value)
        else:
            self._conversion_fn(files, output_folder)
        self.accept()


# ---------------------------------------------------------------------------
# CATIA installation detector
# ---------------------------------------------------------------------------

def detect_catia_root() -> str | None:
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
                        print(f"  Trying key: HKEY_LOCAL_MACHINE\\{reg_path}\\{release}\\0")
                        try:
                            with winreg.OpenKey(ds_key, rf"{release}\0") as release_key:
                                try:
                                    install_path, _ = winreg.QueryValueEx(release_key, "DEST_FOLDER")
                                    candidate = Path(install_path)
                                    if (candidate / "win_b64").exists():
                                        print(f"    -> Valid CATIA installation found: {candidate}")
                                        return str(candidate)
                                except FileNotFoundError:
                                    pass
                        except OSError:
                            pass
                        i += 1
                    except OSError:
                        break
        except OSError:
            pass
    print("No valid CATIA installation detected.")
    return None


# ---------------------------------------------------------------------------
# Conversion functions
# ---------------------------------------------------------------------------

def CATDrawing_to_PDF(file_paths: list[str], output_folder: str | None = None,
                      prefix: str = "DR_", suffix: str = ""):
    """
    Convert CATDrawing files to PDF using pyCATIA.
    If prefix is non-empty, prepends it to the output filename unless it
    already starts with that prefix.
    If suffix is non-empty, appends it to the output stem unless it
    already ends with that suffix.
    """
    from pycatia import catia

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    for path in file_paths:
        src = Path(path).resolve()
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent.resolve()
        dest_dir.mkdir(parents=True, exist_ok=True)

        stem = src.stem
        if prefix and not stem.startswith(prefix):
            stem = f"{prefix}{stem}"
        if suffix and not stem.endswith(suffix):
            stem = f"{stem}{suffix}"
        out_stem = stem

        print(f"Opening: {src}")
        documents.open(str(src))
        from pycatia.drafting_interfaces.drawing_document import DrawingDocument
        drawing_doc = DrawingDocument(application.active_document.com_object)
        drawing = drawing_doc.drawing_root
        sheet_count = drawing.sheets.count

        dest = dest_dir / f"{out_stem}.pdf"
        drawing_doc.export_data(str(dest), "pdf")
        if not dest.exists():
            print(f"  WARNING: export_data did not create {dest}")
        else:
            print(f"  Exported {sheet_count} sheet(s) -> {dest}")

        drawing_doc.close()
        print(f"Done: {src.name}\n")


def CATPart_to_STP(file_paths: list[str], output_folder: str | None = None,
                   prefix: str = "MD_", suffix: str = ""):
    """
    Convert CATPart/CATProduct files to STEP (.stp) using pyCATIA.
    If prefix is non-empty, prepends it to the output filename unless it
    already starts with that prefix.
    If suffix is non-empty, appends it to the output stem unless it
    already ends with that suffix.
    """
    from pycatia import catia

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    for path in file_paths:
        src = Path(path)
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent.resolve()
        dest_dir.mkdir(parents=True, exist_ok=True)

        stem = src.stem
        if prefix and not stem.startswith(prefix):
            stem = f"{prefix}{stem}"
        if suffix and not stem.endswith(suffix):
            stem = f"{stem}{suffix}"
        out_stem = stem

        dest = dest_dir / f"{out_stem}.stp"

        print(f"Opening: {src}")
        documents.open(str(src))
        doc = application.active_document
        doc.export_data(str(dest), "stp")
        print(f"  Exported -> {dest}")
        doc.close()
        print(f"Done: {src.name}\n")


# ---------------------------------------------------------------------------
# Stamp part template function
# ---------------------------------------------------------------------------

def stamp_part_template(file_paths: list[str], output_folder: str | None = None):
    """
    For each CATPart, add the 9 standard user-defined properties if they do
    not already exist. Properties are added as strings with empty default value.
    The part is saved automatically after stamping.
    """
    from pycatia import catia
    from pycatia.mec_mod_interfaces.part_document import PartDocument

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    succeeded = []
    failed    = []

    for path in file_paths:
        src = Path(path).resolve()
        print(f"Opening: {src}")
        try:
            documents.open(str(src))
            doc        = PartDocument(application.active_document.com_object)
            product    = doc.product
            user_props = product.user_ref_properties

            existing_names: set[str] = set()
            for i in range(1, user_props.count + 1):
                try:
                    existing_names.add(user_props.item(i).name)
                except Exception:
                    pass

            added = []
            for prop_name in PART_TEMPLATE_PROPERTIES:
                if prop_name not in existing_names:
                    user_props.create_string(prop_name, "")
                    added.append(prop_name)
                    print(f"  Added property: '{prop_name}'")
                else:
                    print(f"  Skipped (already exists): '{prop_name}'")

            doc.save()
            print(f"  Saved: {src.name}")
            succeeded.append(f"{src.name} (+{len(added)} added)")

        except Exception as e:
            print(f"  ERROR processing {src.name}: {e}")
            failed.append(f"{src.name}: {e}")
        finally:
            try:
                application.active_document.close()
            except Exception:
                pass
        print()

    msg = "Stamping complete.\n\n"
    if succeeded:
        msg += "✔ Succeeded:\n" + "\n".join(f"  {s}" for s in succeeded)
    if failed:
        msg += "\n\n✘ Failed:\n" + "\n".join(f"  {f}" for f in failed)

    from PySide6.QtWidgets import QMessageBox
    if failed:
        QMessageBox.warning(None, "刷写零件模板", msg)
    else:
        QMessageBox.information(None, "刷写零件模板", msg)


# ---------------------------------------------------------------------------
# Export BOM Dialog
# ---------------------------------------------------------------------------

BOM_ALL_COLUMNS       = ["Level", "Part Number", "Nomenclature", "Definition", "Revision", "Source", "Quantity"]
BOM_DEFAULT_COLUMNS   = ["Level", "Part Number", "Nomenclature", "Definition", "Revision", "Source", "Quantity"]
BOM_PRESET_CUSTOM_COLUMNS = ["物料编码", "物料名称", "中文名称", "规格型号", "物料来源", "数据状态", "存货类别", "质量", "备注"]


class ExportBOMDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Export BOM from CATProduct")
        self.setMinimumSize(560, 580)

        self._settings = QSettings("CATIACompanion", "ExportBOMDialog")
        self._last_browse_dir = self._settings.value("last_browse_dir", "")
        self._last_output_dir = self._settings.value("last_output_dir", "")

        saved_custom = self._settings.value("custom_columns", [])
        if isinstance(saved_custom, str):
            saved_custom = [saved_custom]
        self._custom_columns: list[str] = list(saved_custom)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

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

        if self._last_output_dir:
            self.radio_custom.setChecked(True)
            self.folder_edit.setText(self._last_output_dir)

        col_group = QGroupBox("Columns to Export (drag to reorder)")
        col_outer = QVBoxLayout(col_group)
        col_layout = QHBoxLayout()

        avail_layout = QVBoxLayout()
        avail_layout.addWidget(QLabel("Available:"))
        self.avail_list = QListWidget()
        self.avail_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.avail_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        avail_layout.addWidget(self.avail_list)
        col_layout.addLayout(avail_layout)

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

        selected_layout = QVBoxLayout()
        selected_layout.addWidget(QLabel("Selected:"))
        self.selected_list = QListWidget()
        self.selected_list.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.selected_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        selected_layout.addWidget(self.selected_list)
        col_layout.addLayout(selected_layout)
        col_outer.addLayout(col_layout)

        add_custom_row = QHBoxLayout()
        self.preset_combo = QComboBox()
        self.preset_combo.addItem("— Presets —")
        for p in BOM_PRESET_CUSTOM_COLUMNS:
            self.preset_combo.addItem(p)
        self.preset_combo.currentIndexChanged.connect(self._on_preset_selected)
        add_custom_row.addWidget(self.preset_combo)

        self.custom_col_edit = QLineEdit()
        self.custom_col_edit.setPlaceholderText("Custom CATIA property name...")
        self.custom_col_edit.returnPressed.connect(self._add_custom_column)
        add_custom_row.addWidget(self.custom_col_edit)

        add_custom_btn = QPushButton("Add")
        add_custom_btn.clicked.connect(self._add_custom_column)
        add_custom_row.addWidget(add_custom_btn)

        self.delete_custom_btn = QPushButton("Delete Custom")
        self.delete_custom_btn.clicked.connect(self._delete_custom_column)
        self.delete_custom_btn.setEnabled(False)
        add_custom_row.addWidget(self.delete_custom_btn)

        col_outer.addLayout(add_custom_row)
        layout.addWidget(col_group)

        self.avail_list.itemSelectionChanged.connect(self._on_avail_selection_changed)

        saved = self._settings.value("selected_columns", BOM_DEFAULT_COLUMNS)
        if isinstance(saved, str):
            saved = [saved]
        all_known = BOM_ALL_COLUMNS + self._custom_columns
        for col in saved:
            if col in all_known:
                self.selected_list.addItem(QListWidgetItem(col))
        for col in all_known:
            if col not in saved:
                self.avail_list.addItem(QListWidgetItem(col))

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
        file, _ = QFileDialog.getOpenFileName(self, "Select CATProduct File",
            self._last_browse_dir, "CATProduct Files (*.CATProduct);;All Files (*)")
        if file:
            self.file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)

    def _browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder", self._last_output_dir)
        if folder:
            self.folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    def _on_preset_selected(self, index: int):
        if index <= 0:
            return
        label = self.preset_combo.itemText(index)
        self.custom_col_edit.setText(label)
        self.preset_combo.blockSignals(True)
        self.preset_combo.setCurrentIndex(0)
        self.preset_combo.blockSignals(False)

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

    def _add_custom_column(self):
        label = self.custom_col_edit.text().strip()
        if not label:
            return
        all_existing = (
            [self.avail_list.item(i).text() for i in range(self.avail_list.count())] +
            [self.selected_list.item(i).text() for i in range(self.selected_list.count())]
        )
        if label in all_existing:
            QMessageBox.warning(self, "Duplicate Column", f"'{label}' already exists.")
            return
        self.selected_list.addItem(QListWidgetItem(label))
        self._custom_columns.append(label)
        self._settings.setValue("custom_columns", self._custom_columns)
        self.custom_col_edit.clear()

    def _delete_custom_column(self):
        selected = self.avail_list.selectedItems()
        to_delete = [item for item in selected if item.text() in self._custom_columns]
        if not to_delete:
            return
        names = ", ".join(f"'{item.text()}'" for item in to_delete)
        reply = QMessageBox.question(self, "Delete Custom Column",
            f"Permanently delete {names}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        for item in to_delete:
            self._custom_columns.remove(item.text())
            self.avail_list.takeItem(self.avail_list.row(item))
        self._settings.setValue("custom_columns", self._custom_columns)

    def _on_avail_selection_changed(self):
        selected = self.avail_list.selectedItems()
        has_custom = any(item.text() in self._custom_columns for item in selected)
        self.delete_custom_btn.setEnabled(has_custom)

    def _confirm(self):
        file_path = self.file_edit.text().strip()
        if not file_path:
            QMessageBox.warning(self, "No File", "Please select a CATProduct file.")
            return
        selected_cols = [self.selected_list.item(i).text()
                         for i in range(self.selected_list.count())]
        if not selected_cols:
            QMessageBox.warning(self, "No Columns", "Please select at least one column to export.")
            return
        self._settings.setValue("selected_columns", selected_cols)
        if self.radio_same.isChecked():
            output_folder = None
        else:
            output_folder = self.folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(self, "No Output Folder", "Please select an output folder.")
                return
        export_bom_to_excel([file_path], output_folder, columns=selected_cols,
                            custom_columns=self._custom_columns)
        self.accept()


# ---------------------------------------------------------------------------
# BOM export function
# ---------------------------------------------------------------------------

def export_bom_to_excel(file_paths: list[str], output_folder: str | None = None,
                        columns: list[str] | None = None,
                        custom_columns: list[str] | None = None):
    """
    Export a hierarchical BOM from CATProduct files to Excel (.xlsx).
    Custom columns are read from CATIA user-defined properties (UserRefProperties).
    Each product is switched to DESIGN_MODE before reading properties.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment
    from pycatia import catia
    from pycatia.product_structure_interfaces.product_document import ProductDocument
    from pycatia.enumeration.enumeration_types import CatWorkModeType

    if columns is None:
        columns = BOM_DEFAULT_COLUMNS
    if custom_columns is None:
        custom_columns = []

    caa = catia()
    application = caa.application
    application.visible = True
    documents = application.documents

    DIRECT_ATTR_MAP = {
        "Nomenclature": "nomenclature",
        "Revision":     "revision",
        "Definition":   "definition",
        "Source":       "source",
    }

    def get_property(product, name: str) -> str:
        attr = DIRECT_ATTR_MAP.get(name)
        if not attr:
            return ""
        targets = [product]
        try:
            targets.insert(0, product.reference_product)
        except Exception:
            pass
        for target in targets:
            try:
                value = getattr(target, attr)
                if value:
                    return str(value)
            except Exception:
                pass
            try:
                part = target.get_item("Part")
                value = getattr(part, attr)
                if value:
                    return str(value)
            except Exception:
                pass
        return ""

    def get_user_property(product, name: str) -> str:
        targets = [product]
        try:
            targets.insert(0, product.reference_product)
        except Exception:
            pass
        for target in targets:
            try:
                user_props = target.user_ref_properties
                prop = user_props.item(name)
                value = prop.value
                if value is not None and str(value).strip():
                    return str(value)
            except Exception:
                pass
            try:
                part = target.get_item("Part")
                user_props = part.user_ref_properties
                prop = user_props.item(name)
                value = prop.value
                if value is not None and str(value).strip():
                    return str(value)
            except Exception:
                pass
        return ""

    def traverse(product, rows: list, level: int):
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn = name.rsplit(".", 1)[0] if "." in name else name

        try:
            product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
        except Exception as e:
            print(f"  {'  ' * level}  -> apply_work_mode failed: {e}")

        row = {"Level": level, "Part Number": pn}
        print(f"  {'  ' * level}[Level {level}] {pn}")

        for col in columns:
            if col in DIRECT_ATTR_MAP:
                row[col] = get_property(product, col)
            elif col in custom_columns:
                row[col] = get_user_property(product, col)

        rows.append(row)

        try:
            products = product.products
            count = products.count
            if count == 0:
                return
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

        if dest.exists():
            try:
                with open(dest, "a+b"):
                    pass
            except PermissionError:
                from PySide6.QtWidgets import QMessageBox
                reply = QMessageBox.question(None, "File In Use",
                    f"The file is currently open in Excel:\n{dest}\n\n"
                    f"Please close it in Excel, then click Retry, or Cancel to abort.",
                    QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel)
                if reply == QMessageBox.StandardButton.Cancel:
                    continue
                try:
                    with open(dest, "a+b"):
                        pass
                except PermissionError:
                    QMessageBox.critical(None, "Still In Use",
                        f"The file is still open. Please close it and try again.\n{dest}")
                    continue

        print(f"Opening: {src}")
        documents.open(str(src))
        product_doc = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product

        rows = []
        traverse(root_product, rows, level=0)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BOM"
        center = Alignment(horizontal="center")

        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(bold=True)

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
