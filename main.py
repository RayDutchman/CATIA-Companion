import sys
import os
import shutil
import subprocess
import winreg
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QMessageBox, QDialog, QPushButton, QListWidget, QFileDialog,
    QAbstractItemView, QRadioButton, QButtonGroup, QLineEdit, QGroupBox,
    QListWidgetItem, QComboBox, QCheckBox, QPlainTextEdit,
    QToolBar, QToolButton
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSettings, QObject, Signal

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

_log_dir = Path.home() / "CATIA_Companion" / "logs"
_log_dir.mkdir(parents=True, exist_ok=True)
_log_file = _log_dir / "catia_companion.log"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        RotatingFileHandler(_log_file, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)


class _LogSignalEmitter(QObject):
    message_logged = Signal(str)


_log_emitter = _LogSignalEmitter()


class QtLogHandler(logging.Handler):
    def emit(self, record: logging.LogRecord):
        msg = self.format(record)
        _log_emitter.message_logged.emit(msg)


_qt_log_handler = QtLogHandler()
_qt_log_handler.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
))
logging.getLogger().addHandler(_qt_log_handler)


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
        self.resize(300, 500)
        self._setup_menu_bar()
        self._setup_toolbar()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        label = QLabel("欢迎使用 CATIA Companion")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)

        # Log panel container (hidden by default)
        self._log_panel_widget = QWidget()
        log_panel_layout = QVBoxLayout(self._log_panel_widget)
        log_panel_layout.setContentsMargins(0, 0, 0, 0)
        log_panel_layout.setSpacing(4)

        # Log display area
        self._log_view = QPlainTextEdit()
        self._log_view.setReadOnly(True)
        self._log_view.setStyleSheet(
            "background-color: #1e1e1e; color: #d4d4d4;"
            " font-family: Consolas, 'Courier New', monospace; font-size: 9pt;"
        )
        self._log_view.setMinimumHeight(150)
        log_panel_layout.addWidget(self._log_view)

        # Open Log File button
        open_log_btn = QPushButton("打开日志文件")
        open_log_btn.clicked.connect(self._open_log_file)
        log_panel_layout.addWidget(open_log_btn)

        # Log file path label
        log_path_label = QLabel(f"Log: {_log_file}")
        log_path_label.setStyleSheet("color: gray; font-size: 9pt;")
        log_path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        log_panel_layout.addWidget(log_path_label)

        layout.addWidget(self._log_panel_widget)
        self._log_panel_widget.setVisible(False)

        _log_emitter.message_logged.connect(self._append_log)

        self.statusBar().showMessage("就绪")

    def _append_log(self, message: str):
        self._log_view.appendPlainText(message)
        self._log_view.verticalScrollBar().setValue(
            self._log_view.verticalScrollBar().maximum()
        )

    def _open_log_file(self):
        try:
            if sys.platform == "win32":
                os.startfile(_log_file)
            else:
                subprocess.Popen(
                    ["xdg-open", str(_log_file)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
        except Exception as e:
            QMessageBox.warning(self, "无法打开日志文件",
                f"无法打开日志文件：\n{_log_file}\n\n{e}")

    def _toggle_log_panel(self):
        self._log_panel_widget.setVisible(not self._log_panel_widget.isVisible())

    def _setup_toolbar(self):
        toolbar = QToolBar("主工具栏")
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setOrientation(Qt.Orientation.Vertical)
        self.addToolBar(Qt.ToolBarArea.LeftToolBarArea, toolbar)

        def _make_tool_button(text: str, slot) -> QToolButton:
            btn = QToolButton()
            btn.setText(text)
            btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            btn.setFixedWidth(90)
            btn.clicked.connect(slot)
            return btn

        toolbar.addWidget(_make_tool_button("从CATDrawing\n导出pdf", self._open_convert_drawing_dialog))
        toolbar.addWidget(_make_tool_button("从CATPart/\nCATProduct\n导出stp", self._open_convert_part_dialog))
        toolbar.addWidget(_make_tool_button("从CATProduct\n导出BOM", self._open_export_bom_dialog))
        toolbar.addWidget(_make_tool_button("刷写零件\n模板", self._open_stamp_part_template_dialog))
        toolbar.addWidget(_make_tool_button("查找所有\n依赖项", self._open_find_dependencies_dialog))

    def _setup_menu_bar(self):
        menu_bar = self.menuBar()

        # --- File ---
        file_menu = menu_bar.addMenu("文件")
        file_menu.addAction(QAction("新建", self))
        file_menu.addAction(QAction("打开...", self))
        file_menu.addAction(QAction("保存", self))
        file_menu.addAction(QAction("另存为...", self))
        file_menu.addSeparator()
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        # --- Export ---
        export_menu = menu_bar.addMenu("导出")
        convert_drawing_action = QAction("从CATDrawing导出pdf", self)
        convert_drawing_action.triggered.connect(self._open_convert_drawing_dialog)
        export_menu.addAction(convert_drawing_action)
        convert_part_action = QAction("从CATPart/CATProduct导出stp", self)
        convert_part_action.triggered.connect(self._open_convert_part_dialog)
        export_menu.addAction(convert_part_action)
        export_bom_action = QAction("从CATProduct导出BOM", self)
        export_bom_action.triggered.connect(self._open_export_bom_dialog)
        export_menu.addAction(export_bom_action)

        # --- Edit ---
        edit_menu = menu_bar.addMenu("编辑")
        edit_menu.addAction(QAction("撤销", self))
        edit_menu.addAction(QAction("重做", self))
        edit_menu.addSeparator()
        edit_menu.addAction(QAction("剪切", self))
        edit_menu.addAction(QAction("复制", self))
        edit_menu.addAction(QAction("粘贴", self))

        # --- Tools ---
        tools_menu = menu_bar.addMenu("工具")
        copy_font_action = QAction("复制字体文件到CATIA目录", self)
        copy_font_action.triggered.connect(self._copy_font_to_catia)
        tools_menu.addAction(copy_font_action)
        copy_iso_action = QAction("复制ISO.xml到CATIA目录", self)
        copy_iso_action.triggered.connect(self._copy_iso_to_catia)
        tools_menu.addAction(copy_iso_action)
        pojie_action = QAction("PoJie", self)
        pojie_action.triggered.connect(self._pojie)
        tools_menu.addAction(pojie_action)
        stamp_action = QAction("刷写零件模板", self)
        stamp_action.triggered.connect(self._open_stamp_part_template_dialog)
        tools_menu.addAction(stamp_action)
        find_deps_action = QAction("查找所有依赖项", self)
        find_deps_action.triggered.connect(self._open_find_dependencies_dialog)
        tools_menu.addAction(find_deps_action)

        # --- View ---
        view_menu = menu_bar.addMenu("视图")
        zoom_in_action = QAction("放大", self)
        zoom_in_action.triggered.connect(lambda: QMessageBox.information(self, "提示", "功能尚未实现"))
        view_menu.addAction(zoom_in_action)
        zoom_out_action = QAction("缩小", self)
        zoom_out_action.triggered.connect(lambda: QMessageBox.information(self, "提示", "功能尚未实现"))
        view_menu.addAction(zoom_out_action)
        zoom_reset_action = QAction("重置缩放", self)
        zoom_reset_action.triggered.connect(lambda: QMessageBox.information(self, "提示", "功能尚未实现"))
        view_menu.addAction(zoom_reset_action)
        view_menu.addSeparator()
        toggle_log_action = QAction("切换Log显示", self)
        toggle_log_action.triggered.connect(self._toggle_log_panel)
        view_menu.addAction(toggle_log_action)

        # --- Help ---
        help_menu = menu_bar.addMenu("帮助")
        help_menu.addAction(QAction("文档", self))
        about_action = QAction("关于 CATIA Companion", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _show_about(self):
        QMessageBox.about(self, f"About {APP_NAME}", ABOUT_TEXT)

    def _open_convert_part_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="将CATPart/CATProduct导出为STP",
            file_label="已选CATPart/CATProduct文件:",
            file_filter="CATIA Part/Product Files (*.CATPart *.CATProduct);;All Files (*)",
            no_files_msg="请至少选择一个CATPart或CATProduct文件。",
            conversion_fn=CATPart_to_STP,
            settings_key="CATPart",
            show_prefix_option=True,
            prefix="MD_",
            note="暂时留空"
        )
        dialog.exec()

    def _open_convert_drawing_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="将CATDrawing导出为PDF",
            file_label="已选CATDrawing文件:",
            file_filter="CATDrawing Files (*.CATDrawing);;All Files (*)",
            no_files_msg="请至少选择一个CATDrawing文件。",
            conversion_fn=CATDrawing_to_PDF,
            settings_key="CATDrawing",
            show_prefix_option=True,
            prefix="DR_",
            note="如果用于导出的CATDrawing有多页，请将CATIA设置为\u201c将多页文档保存在单向量文件中\u201d（工具->选项->常规->兼容性->图形格式->导出（另存为））"
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
            QMessageBox.warning(self, "文件未找到",
                f"在工作目录中找不到 '{file_name}'：\n{src_file.parent}")
            return

        catia_root = detect_catia_root()
        if catia_root:
            reply = QMessageBox.question(self, "检测到CATIA安装",
                f"检测到CATIA安装路径：\n{catia_root}\n\n是否使用该目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(self,
                "选择CATIA安装目录（例如 C:\\Program Files\\Dassault Systemes\\B28）", "")
            if not catia_root:
                return

        dest_dir = Path(catia_root) / relative_dest
        if not dest_dir.exists():
            reply = QMessageBox.question(self, "文件夹未找到",
                f"目标文件夹不存在：\n{dest_dir}\n\n是否要创建该文件夹？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                dest_dir.mkdir(parents=True, exist_ok=True)
            else:
                return

        dest_file = dest_dir / file_name
        try:
            shutil.copy2(str(src_file), str(dest_file))
            QMessageBox.information(self, "成功",
                f"'{file_name}' 已成功复制到：\n{dest_file}")
        except PermissionError:
            QMessageBox.critical(self, "权限不足",
                f"无法复制文件，请以管理员身份运行程序。\n\n目标路径：\n{dest_file}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生意外错误：\n{e}")

    def _pojie(self):
        src_dir = resource_path("Pojie")
        if not src_dir.exists() or not src_dir.is_dir():
            QMessageBox.warning(self, "文件夹未找到",
                f"找不到 'Pojie' 文件夹：\n{src_dir.parent}")
            return

        catia_root = detect_catia_root()
        if catia_root:
            reply = QMessageBox.question(self, "检测到CATIA安装",
                f"检测到CATIA安装路径：\n{catia_root}\n\n是否使用该目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(self,
                "选择CATIA安装目录（例如 C:\\Program Files\\Dassault Systemes\\B28）", "")
            if not catia_root:
                return

        dest_dir = Path(catia_root) / "win_b64" / "code" / "bin"
        if not dest_dir.exists():
            QMessageBox.critical(self, "文件夹未找到",
                f"目标文件夹不存在：\n{dest_dir}\n\n请检查您的CATIA安装。")
            return

        files = [f for f in src_dir.iterdir() if f.is_file()]
        if not files:
            QMessageBox.warning(self, "文件夹为空", "'Pojie' 文件夹中没有文件。")
            return

        try:
            copied = []
            for src_file in files:
                dest_file = dest_dir / src_file.name
                shutil.copy2(str(src_file), str(dest_file))
                copied.append(src_file.name)
                logger.info(f"  Copied: {src_file.name} -> {dest_file}")
            QMessageBox.information(self, "成功",
                f"已成功复制 {len(copied)} 个文件到：\n{dest_dir}\n\n" + "\n".join(copied))
        except PermissionError:
            QMessageBox.critical(self, "权限不足",
                f"无法复制文件，请以管理员身份运行程序。\n\n目标路径：\n{dest_dir}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生意外错误：\n{e}")

    def _open_stamp_part_template_dialog(self):
        dialog = ConvertDialog(
            parent=self,
            title="刷写零件模板",
            file_label="已选CATPart文件:",
            file_filter="CATIA Part Files (*.CATPart);;All Files (*)",
            no_files_msg="请至少选择一个CATPart文件。",
            conversion_fn=stamp_part_template,
            settings_key="StampPartTemplate"
        )
        dialog.exec()

    def _open_find_dependencies_dialog(self):
        dialog = FindDependenciesDialog(self)
        dialog.exec()


# ---------------------------------------------------------------------------
# Generic Convert Dialog
# ---------------------------------------------------------------------------

class ConvertDialog(QDialog):
    def __init__(self, parent=None, title="Convert", file_label="Selected files:",
                 file_filter="All Files (*)", no_files_msg="Please select at least one file.",
                 conversion_fn=None, settings_key="default",
                 show_prefix_option=False, prefix="", note: str = ""):
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
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self._browse_files)
        remove_btn = QPushButton("移除所选")
        remove_btn.clicked.connect(self._remove_selected)
        remove_all_btn = QPushButton("全部移除")
        remove_all_btn.clicked.connect(self._remove_all)
        btn_row.addWidget(browse_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addWidget(remove_all_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Output folder — hidden for stamp dialog
        if settings_key != "StampPartTemplate":
            output_group = QGroupBox("输出文件夹")
            output_layout = QVBoxLayout(output_group)
            self.radio_same = QRadioButton("与源文件相同目录")
            self.radio_custom = QRadioButton("自定义目录:")
            self.radio_same.setChecked(True)
            self.btn_group = QButtonGroup(self)
            self.btn_group.addButton(self.radio_same)
            self.btn_group.addButton(self.radio_custom)
            output_layout.addWidget(self.radio_same)
            output_layout.addWidget(self.radio_custom)
            folder_row = QHBoxLayout()
            self.folder_edit = QLineEdit()
            self.folder_edit.setPlaceholderText("选择输出文件夹...")
            self.folder_edit.setReadOnly(True)
            self.folder_edit.setEnabled(False)
            self.folder_browse_btn = QPushButton("浏览...")
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
            self.prefix_checkbox = QCheckBox("添加前缀:")
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
            self.suffix_checkbox = QCheckBox("添加后缀:")
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

        if note:
            note_label = QLabel(note)
            note_label.setWordWrap(True)
            note_label.setStyleSheet("color: gray; font-size: 11px;")
            layout.addWidget(note_label)

        action_row = QHBoxLayout()
        action_row.addStretch()
        confirm_btn = QPushButton("确认")
        confirm_btn.setDefault(True)
        confirm_btn.clicked.connect(self._confirm)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(confirm_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    def _toggle_folder_row(self, checked):
        self.folder_edit.setEnabled(checked)
        self.folder_browse_btn.setEnabled(checked)

    def _browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", self._last_browse_dir, self._file_filter)
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
            self, "选择输出文件夹", self._last_output_dir)
        if folder:
            self.folder_edit.setText(folder)
            self._last_output_dir = folder
            self._settings.setValue("last_output_dir", folder)

    def _confirm(self):
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not files:
            QMessageBox.warning(self, "未选择文件", self._no_files_msg)
            return

        if self.radio_same is None:
            output_folder = None
        elif self.radio_same.isChecked():
            output_folder = None
        else:
            output_folder = self.folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(self, "未选择输出文件夹", "请选择一个输出文件夹。")
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
        logger.debug(f"Trying registry path: HKEY_LOCAL_MACHINE\\{reg_path}")
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as ds_key:
                i = 0
                while True:
                    try:
                        release = winreg.EnumKey(ds_key, i)
                        logger.debug(f"  Trying key: HKEY_LOCAL_MACHINE\\{reg_path}\\{release}\\0")
                        try:
                            with winreg.OpenKey(ds_key, rf"{release}\0") as release_key:
                                try:
                                    install_path, _ = winreg.QueryValueEx(release_key, "DEST_FOLDER")
                                    candidate = Path(install_path)
                                    if (candidate / "win_b64").exists():
                                        logger.debug(f"    -> Valid CATIA installation found: {candidate}")
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
    logger.debug("No valid CATIA installation detected.")
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

    bulk_action = None  # "skip_all", "overwrite_all", or "cancel"

    for path in file_paths:
        if bulk_action == "cancel":
            break

        src = Path(path).resolve()
        dest_dir = Path(output_folder).resolve() if output_folder else src.parent.resolve()
        dest_dir.mkdir(parents=True, exist_ok=True)

        stem = src.stem
        if prefix and not stem.startswith(prefix):
            stem = f"{prefix}{stem}"
        if suffix and not stem.endswith(suffix):
            stem = f"{stem}{suffix}"
        out_stem = stem

        logger.info(f"Opening: {src}")
        dest = dest_dir / f"{out_stem}.pdf"

        if dest.exists():
            if bulk_action == "skip_all":
                logger.info(f"  Skipped (skip all): {dest}")
                continue
            if bulk_action == "overwrite_all":
                dest.unlink()
            else:
                msg = QMessageBox()
                msg.setWindowTitle("文件已存在")
                msg.setText(f'"{dest.name}" 已存在于输出文件夹中。')
                msg.setInformativeText(str(dest.parent))
                msg.setIcon(QMessageBox.Icon.Warning)
                skip_btn          = msg.addButton("跳过",     QMessageBox.ButtonRole.RejectRole)
                skip_all_btn      = msg.addButton("全部跳过", QMessageBox.ButtonRole.RejectRole)
                overwrite_btn     = msg.addButton("覆盖",     QMessageBox.ButtonRole.AcceptRole)
                overwrite_all_btn = msg.addButton("全部覆盖", QMessageBox.ButtonRole.AcceptRole)
                cancel_btn        = msg.addButton("取消",     QMessageBox.ButtonRole.DestructiveRole)
                msg.exec()
                clicked = msg.clickedButton()
                if clicked is cancel_btn:
                    bulk_action = "cancel"
                    break
                elif clicked is skip_all_btn:
                    bulk_action = "skip_all"
                    logger.info(f"  Skipped (skip all): {dest}")
                    continue
                elif clicked is skip_btn:
                    logger.info(f"  Skipped: {dest}")
                    continue
                elif clicked is overwrite_all_btn:
                    bulk_action = "overwrite_all"
                    dest.unlink()
                else:  # overwrite_btn
                    dest.unlink()

        documents.open(str(src))
        from pycatia.drafting_interfaces.drawing_document import DrawingDocument
        drawing_doc = DrawingDocument(application.active_document.com_object)
        drawing = drawing_doc.drawing_root
        sheet_count = drawing.sheets.count

        drawing_doc.export_data(str(dest), "pdf")
        if not dest.exists():
            logger.warning(f"  WARNING: export_data did not create {dest}")
        else:
            logger.info(f"  Exported {sheet_count} sheet(s) -> {dest}")

        drawing_doc.close()
        logger.info(f"Done: {src.name}\n")


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

    bulk_action = None  # "skip_all", "overwrite_all", or "cancel"

    for path in file_paths:
        if bulk_action == "cancel":
            break

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

        logger.info(f"Opening: {src}")
        if dest.exists():
            if bulk_action == "skip_all":
                logger.info(f"  Skipped (skip all): {dest}")
                continue
            if bulk_action == "overwrite_all":
                dest.unlink()
            else:
                msg = QMessageBox()
                msg.setWindowTitle("文件已存在")
                msg.setText(f'"{dest.name}" 已存在于输出文件夹中。')
                msg.setInformativeText(str(dest.parent))
                msg.setIcon(QMessageBox.Icon.Warning)
                skip_btn          = msg.addButton("跳过",     QMessageBox.ButtonRole.RejectRole)
                skip_all_btn      = msg.addButton("全部跳过", QMessageBox.ButtonRole.RejectRole)
                overwrite_btn     = msg.addButton("覆盖",     QMessageBox.ButtonRole.AcceptRole)
                overwrite_all_btn = msg.addButton("全部覆盖", QMessageBox.ButtonRole.AcceptRole)
                cancel_btn        = msg.addButton("取消",     QMessageBox.ButtonRole.DestructiveRole)
                msg.exec()
                clicked = msg.clickedButton()
                if clicked is cancel_btn:
                    bulk_action = "cancel"
                    break
                elif clicked is skip_all_btn:
                    bulk_action = "skip_all"
                    logger.info(f"  Skipped (skip all): {dest}")
                    continue
                elif clicked is skip_btn:
                    logger.info(f"  Skipped: {dest}")
                    continue
                elif clicked is overwrite_all_btn:
                    bulk_action = "overwrite_all"
                    dest.unlink()
                else:  # overwrite_btn
                    dest.unlink()

        logger.info(f"Opening: {src}")
        documents.open(str(src))
        doc = application.active_document
        doc.export_data(str(dest), "stp")
        logger.info(f"  Exported -> {dest}")
        doc.close()
        logger.info(f"Done: {src.name}\n")


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
        logger.info(f"Opening: {src}")
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
                    logger.info(f"  Added property: '{prop_name}'")
                else:
                    logger.info(f"  Skipped (already exists): '{prop_name}'")

            doc.save()
            logger.info(f"  Saved: {src.name}")
            succeeded.append(f"{src.name} (+{len(added)} added)")

        except Exception as e:
            logger.error(f"  ERROR processing {src.name}: {e}")
            failed.append(f"{src.name}: {e}")
        finally:
            try:
                application.active_document.close()
            except Exception:
                pass

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

BOM_ALL_COLUMNS       = ["Level", "Part Number", "Type", "Nomenclature", "Definition", "Revision", "Source", "Quantity"]
BOM_DEFAULT_COLUMNS   = ["Level", "Part Number", "Type", "Nomenclature", "Definition", "Revision", "Source", "Quantity"]
BOM_PRESET_CUSTOM_COLUMNS = ["物料编码", "物料名称", "中文名称", "规格型号", "物料来源", "数据状态", "存货类别", "质量", "备注"]


class ExportBOMDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("从CATProduct导出BOM")
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

        layout.addWidget(QLabel("CATProduct文件:"))
        file_row = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("选择一个CATProduct文件...")
        self.file_edit.setReadOnly(True)
        file_browse_btn = QPushButton("浏览...")
        file_browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self.file_edit)
        file_row.addWidget(file_browse_btn)
        layout.addLayout(file_row)

        output_group = QGroupBox("输出文件夹")
        output_layout = QVBoxLayout(output_group)
        self.radio_same = QRadioButton("与源文件相同目录")
        self.radio_custom = QRadioButton("自定义目录:")
        self.radio_same.setChecked(True)
        self.btn_group = QButtonGroup(self)
        self.btn_group.addButton(self.radio_same)
        self.btn_group.addButton(self.radio_custom)
        output_layout.addWidget(self.radio_same)
        output_layout.addWidget(self.radio_custom)
        folder_row = QHBoxLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("选择输出文件夹...")
        self.folder_edit.setReadOnly(True)
        self.folder_edit.setEnabled(False)
        self.folder_browse_btn = QPushButton("浏览...")
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

        col_group = QGroupBox("导出列（拖动以排序）")
        col_outer = QVBoxLayout(col_group)
        col_layout = QHBoxLayout()

        avail_layout = QVBoxLayout()
        avail_layout.addWidget(QLabel("可用列:"))
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
        selected_layout.addWidget(QLabel("已选列:"))
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
        self.custom_col_edit.setPlaceholderText("自定义CATIA属性名...")
        self.custom_col_edit.returnPressed.connect(self._add_custom_column)
        add_custom_row.addWidget(self.custom_col_edit)

        add_custom_btn = QPushButton("添加")
        add_custom_btn.clicked.connect(self._add_custom_column)
        add_custom_row.addWidget(add_custom_btn)

        self.delete_custom_btn = QPushButton("删除自定义")
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
        confirm_btn = QPushButton("导出")
        confirm_btn.setDefault(True)
        confirm_btn.clicked.connect(self._confirm)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(confirm_btn)
        action_row.addWidget(cancel_btn)
        layout.addLayout(action_row)

    def _toggle_folder_row(self, checked):
        self.folder_edit.setEnabled(checked)
        self.folder_browse_btn.setEnabled(checked)

    def _browse_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择CATProduct文件",
            self._last_browse_dir, "CATProduct Files (*.CATProduct);;All Files (*)")
        if file:
            self.file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)

    def _browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "选择输出文件夹", self._last_output_dir)
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
            QMessageBox.warning(self, "列名重复", f"'{label}' 已存在。")
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
        reply = QMessageBox.question(self, "删除自定义列",
            f"确认永久删除 {names}？",
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
            QMessageBox.warning(self, "未选择文件", "请选择一个CATProduct文件。")
            return
        selected_cols = [self.selected_list.item(i).text()
                         for i in range(self.selected_list.count())]
        if not selected_cols:
            QMessageBox.warning(self, "未选择列", "请至少选择一列进行导出。")
            return
        self._settings.setValue("selected_columns", selected_cols)
        if self.radio_same.isChecked():
            output_folder = None
        else:
            output_folder = self.folder_edit.text().strip()
            if not output_folder:
                QMessageBox.warning(self, "未选择输出文件夹", "请选择一个输出文件夹。")
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
    from pycatia import CatWorkModeType

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
            logger.warning(f"  {'  ' * level}  -> apply_work_mode failed: {e}")

        row = {"Level": level, "Part Number": pn}
        try:
            child_count = product.products.count
            row["Type"] = "装配体" if child_count > 0 else "零件"
        except Exception as e:
            logger.debug(f"Could not determine Type for {pn}: {e}")
            row["Type"] = ""
        logger.debug(f"  {'  ' * level}[Level {level}] {pn}")

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
                    logger.warning(f"  {'  ' * level}  -> Skipping child {i}: {e}")
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
            logger.warning(f"  {'  ' * level}  -> Exception accessing children: {e}")

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
                reply = QMessageBox.question(None, "文件正在使用",
                    f"该文件当前在Excel中已打开：\n{dest}\n\n"
                    f"请在Excel中关闭该文件，然后点击【重试】，或点击【取消】以中止。",
                    QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel)
                if reply == QMessageBox.StandardButton.Cancel:
                    continue
                try:
                    with open(dest, "a+b"):
                        pass
                except PermissionError:
                    QMessageBox.critical(None, "文件仍在使用中",
                        f"文件仍处于打开状态，请关闭后重试。\n{dest}")
                    continue

        logger.info(f"Opening: {src}")
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
                elif col_name == "Type":
                    value = row.get("Type", "")
                else:
                    value = row.get(col_name, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if col_name in ("Level", "Quantity", "Type"):
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
        logger.info(f"  BOM exported -> {dest}")
        product_doc.close()
        logger.info(f"Done: {src.name}\n")


# ---------------------------------------------------------------------------
# Find Dependencies function
# ---------------------------------------------------------------------------

def find_dependencies(target_path: str, search_dir: str, recursive: bool = True,
                      progress_callback=None) -> list[str]:
    """
    Scan search_dir for CATIA files that are referenced by target_path.
    Uses binary string matching (no CATIA COM required).
    """
    target = Path(target_path)
    search = Path(search_dir)

    try:
        target_bytes = target.read_bytes()
    except Exception as e:
        logger.error(f"Cannot read target file: {e}")
        return []

    pattern = "**/*" if recursive else "*"
    extensions = {".CATPart", ".CATProduct", ".CATDrawing"}

    results = []
    candidates = [f for ext in extensions
                  for f in search.glob(f"{pattern}{ext}") if f.is_file()]

    for i, candidate in enumerate(candidates):
        if progress_callback:
            progress_callback(i, len(candidates), str(candidate))

        # Skip the target file itself
        if candidate.resolve() == target.resolve():
            continue

        # Check if candidate filename appears in target binary content
        name_bytes = candidate.name.encode("utf-8")
        stem_bytes = candidate.stem.encode("utf-8")

        if name_bytes in target_bytes or stem_bytes in target_bytes:
            results.append(str(candidate))
            logger.info(f"  Dependency found: {candidate}")

    return results


# ---------------------------------------------------------------------------
# Find Dependencies Dialog
# ---------------------------------------------------------------------------

class FindDependenciesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("查找所有依赖项")
        self.setMinimumSize(540, 520)

        self._settings = QSettings("CATIACompanion", "FindDependenciesDialog")

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # Target file
        layout.addWidget(QLabel("目标CATIA文件（CATPart / CATProduct / CATDrawing）:"))
        target_row = QHBoxLayout()
        self._target_edit = QLineEdit()
        self._target_edit.setReadOnly(True)
        self._target_edit.setPlaceholderText("选择目标CATIA文件...")
        self._target_edit.setText(self._settings.value("last_target", ""))
        target_browse_btn = QPushButton("浏览...")
        target_browse_btn.clicked.connect(self._browse_target)
        target_row.addWidget(self._target_edit)
        target_row.addWidget(target_browse_btn)
        layout.addLayout(target_row)

        # Search directory
        layout.addWidget(QLabel("搜索范围（文件夹）:"))
        search_row = QHBoxLayout()
        self._search_edit = QLineEdit()
        self._search_edit.setReadOnly(True)
        self._search_edit.setPlaceholderText("选择搜索文件夹...")
        self._search_edit.setText(self._settings.value("last_search_dir", ""))
        search_browse_btn = QPushButton("浏览...")
        search_browse_btn.clicked.connect(self._browse_search_dir)
        search_row.addWidget(self._search_edit)
        search_row.addWidget(search_browse_btn)
        layout.addLayout(search_row)

        self._recursive_cb = QCheckBox("包含子文件夹")
        self._recursive_cb.setChecked(True)
        layout.addWidget(self._recursive_cb)

        # Action buttons
        action_row = QHBoxLayout()
        self._search_btn = QPushButton("开始搜索")
        self._search_btn.setDefault(True)
        self._search_btn.clicked.connect(self._start_search)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        action_row.addWidget(self._search_btn)
        action_row.addWidget(cancel_btn)
        action_row.addStretch()
        layout.addLayout(action_row)

        # Results area
        layout.addWidget(QLabel("找到的依赖项："))
        self._result_view = QPlainTextEdit()
        self._result_view.setReadOnly(True)
        self._result_view.setMinimumHeight(150)
        layout.addWidget(self._result_view)

        copy_btn = QPushButton("复制结果")
        copy_btn.clicked.connect(self._copy_results)
        layout.addWidget(copy_btn)

    def _browse_target(self):
        last = self._settings.value("last_target", "")
        start_dir = str(Path(last).parent) if last else ""
        file, _ = QFileDialog.getOpenFileName(
            self, "选择目标CATIA文件", start_dir,
            "CATIA Files (*.CATPart *.CATProduct *.CATDrawing);;All Files (*)")
        if file:
            self._target_edit.setText(file)
            self._settings.setValue("last_target", file)

    def _browse_search_dir(self):
        last = self._settings.value("last_search_dir", "")
        folder = QFileDialog.getExistingDirectory(self, "选择搜索文件夹", last)
        if folder:
            self._search_edit.setText(folder)
            self._settings.setValue("last_search_dir", folder)

    def _start_search(self):
        target = self._target_edit.text().strip()
        search_dir = self._search_edit.text().strip()

        if not target:
            QMessageBox.warning(self, "未选择目标文件", "请先选择一个目标CATIA文件。")
            return
        if not Path(target).exists():
            QMessageBox.warning(self, "文件不存在", f"目标文件不存在：\n{target}")
            return
        if not search_dir:
            QMessageBox.warning(self, "未选择搜索范围", "请先选择一个搜索文件夹。")
            return
        if not Path(search_dir).is_dir():
            QMessageBox.warning(self, "文件夹不存在", f"搜索文件夹不存在：\n{search_dir}")
            return

        recursive = self._recursive_cb.isChecked()
        self._search_btn.setEnabled(False)
        self._result_view.clear()
        self._result_view.appendPlainText("正在搜索...")
        QApplication.processEvents()

        def progress_callback(i, total, current_path):
            self._result_view.setPlainText(f"正在扫描 ({i + 1}/{total}):\n{current_path}")
            QApplication.processEvents()

        results = find_dependencies(target, search_dir, recursive=recursive,
                                    progress_callback=progress_callback)

        self._search_btn.setEnabled(True)

        if results:
            summary = f"搜索完成，共找到 {len(results)} 个依赖项：\n\n"
            summary += "\n".join(results)
        else:
            summary = "搜索完成，未找到任何依赖项。"
        self._result_view.setPlainText(summary)

    def _copy_results(self):
        text = self._result_view.toPlainText()
        if text:
            QApplication.clipboard().setText(text)


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
