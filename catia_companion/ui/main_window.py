"""
Main application window.

Provides:
- MainWindow – the primary QMainWindow with a grouped-button UI and menu bar.
"""

import sys
import os
import shutil
import subprocess
import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QMessageBox, QPushButton, QFileDialog, QGroupBox, QInputDialog,
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt

from catia_companion.constants import (
    APP_NAME,
    ABOUT_TEXT,
    MAIN_WINDOW_DEFAULT_WIDTH,
    MAIN_WINDOW_DEFAULT_HEIGHT,
    FONT_FILE_PATH,
    ISO_XML_FILE_PATH,
    CRACK_DIR_PATH,
)
from catia_companion.utils import resource_path, detect_catia_root
from catia_companion.logging_setup import log_signal_emitter
from catia_companion.catia.conversion import convert_drawing_to_pdf, convert_part_to_step
from catia_companion.catia.template import apply_part_template
from catia_companion.ui.log_window import LogWindow
from catia_companion.ui.convert_dialog import FileConvertDialog
from catia_companion.ui.export_bom_dialog import ExportBomDialog
from catia_companion.ui.find_deps_dialog import FindDependenciesDialog
from catia_companion.ui.bom_edit_dialog import BomEditDialog
from catia_companion.ui.help_dialog import HelpDialog

logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """Primary application window."""

    # 快速运行宏仅支持 CATScript 文件（.catvbs / .catscript），不支持 .catvba
    _MACRO_EXTENSIONS: frozenset[str] = frozenset({".catvbs", ".catscript"})

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(MAIN_WINDOW_DEFAULT_WIDTH, MAIN_WINDOW_DEFAULT_HEIGHT)

        self._log_window = LogWindow(self)
        log_signal_emitter.message_logged.connect(self._log_window.append_log)

        self._build_menu_bar()
        self._build_central_widget()
        self.statusBar().showMessage("就绪")

    # ── Central widget ─────────────────────────────────────────────────────

    def _build_central_widget(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 16)
        layout.setSpacing(12)

        # Welcome label
        welcome = QLabel(f"欢迎使用 {APP_NAME}")
        welcome.setAlignment(Qt.AlignmentFlag.AlignCenter)
        welcome.setStyleSheet("font-size: 11pt; font-weight: bold; color: #333;")
        layout.addWidget(welcome)

        # ── Export group ─────────────────────────────────────────────────
        export_group  = QGroupBox("导出")
        export_layout = QVBoxLayout(export_group)
        export_layout.setSpacing(6)

        btn_drawing = QPushButton("CATDrawing → PDF")
        btn_drawing.setToolTip("将 CATDrawing 文件批量导出为 PDF")
        btn_drawing.clicked.connect(self._open_convert_drawing_dialog)

        btn_part = QPushButton("CATPart / CATProduct → STP")
        btn_part.setToolTip("将 CATPart 或 CATProduct 文件批量导出为 STEP")
        btn_part.clicked.connect(self._open_convert_part_dialog)

        for btn in (btn_drawing, btn_part):
            export_layout.addWidget(btn)
        layout.addWidget(export_group)

        # ── BOM group ────────────────────────────────────────────────────
        bom_group  = QGroupBox("BOM")
        bom_layout = QVBoxLayout(bom_group)
        bom_layout.setSpacing(6)

        btn_bom_export = QPushButton("从 CATProduct 导出 BOM")
        btn_bom_export.setToolTip("从 CATProduct 导出 BOM 到 Excel 文件")
        btn_bom_export.clicked.connect(self._open_export_bom_dialog)

        btn_bom_edit = QPushButton("BOM 属性补全")
        btn_bom_edit.setToolTip("在表格中编辑 BOM 属性并写回 CATIA")
        btn_bom_edit.clicked.connect(self._open_bom_edit_dialog)

        for btn in (btn_bom_export, btn_bom_edit):
            bom_layout.addWidget(btn)
        layout.addWidget(bom_group)

        # ── Drawing group ────────────────────────────────────────────────
        drawing_group  = QGroupBox("图纸")
        drawing_layout = QVBoxLayout(drawing_group)
        drawing_layout.setSpacing(6)

        drawing_row = QHBoxLayout()
        drawing_row.setSpacing(6)

        btn_new_drawing = QPushButton("新建图纸")
        btn_new_drawing.setToolTip("从 CATPart/CATProduct 生成 CATDrawing 图纸")
        btn_new_drawing.clicked.connect(self._open_generate_drawing_dialog)

        btn_refresh_drawing = QPushButton("刷新图纸")
        btn_refresh_drawing.setToolTip("刷新当前活动图纸的参数信息（从对应零件/装配体同步属性）")
        btn_refresh_drawing.clicked.connect(self._open_refresh_drawing_dialog)

        drawing_row.addWidget(btn_new_drawing)
        drawing_row.addWidget(btn_refresh_drawing)
        drawing_layout.addLayout(drawing_row)
        layout.addWidget(drawing_group)

        # ── Tools group ──────────────────────────────────────────────────
        tools_group  = QGroupBox("工具")
        tools_layout = QVBoxLayout(tools_group)
        tools_layout.setSpacing(6)

        btn_font = QPushButton("复制字体文件到 CATIA 目录")
        btn_font.setToolTip("将 Changfangsong.ttf 复制到 CATIA 字体目录")
        btn_font.clicked.connect(self._copy_font_to_catia)

        btn_iso = QPushButton("复制 ISO.xml 到 CATIA 目录")
        btn_iso.setToolTip("将 ISO.xml 复制到 CATIA 标准目录")
        btn_iso.clicked.connect(self._copy_iso_to_catia)

        btn_crack = QPushButton("Crack")
        btn_crack.setToolTip("将 crack 文件夹中的文件复制到 CATIA bin 目录")
        btn_crack.clicked.connect(self._crack)

        btn_stamp = QPushButton("刷写零件模板")
        btn_stamp.setToolTip("为选中的 CATPart 添加标准用户自定义属性")
        btn_stamp.clicked.connect(self._open_stamp_part_template_dialog)

        btn_deps = QPushButton("查找所有依赖项（未实现）")
        btn_deps.setToolTip("通过 CATIA COM 查找文件的所有引用文档")
        btn_deps.clicked.connect(self._open_find_dependencies_dialog)

        for btn in (btn_font, btn_iso, btn_crack, btn_stamp, btn_deps):
            tools_layout.addWidget(btn)
        layout.addWidget(tools_group)

        layout.addStretch()

    # ── Menu bar ───────────────────────────────────────────────────────────

    def _build_menu_bar(self) -> None:
        bar = self.menuBar()

        # File
        file_menu  = bar.addMenu("文件")
        _not_implemented = lambda: QMessageBox.information(self, "提示", "功能尚未实现")
        for label, slot in (("新建", _not_implemented), ("打开...", _not_implemented),
                             ("保存", _not_implemented), ("另存为...", _not_implemented)):
            file_menu.addAction(QAction(label, self, triggered=slot))
        file_menu.addSeparator()
        file_menu.addAction(QAction("退出", self, triggered=self.close))

        # Export
        export_menu = bar.addMenu("导出")
        export_menu.addAction(QAction(
            "从CATDrawing导出pdf", self,
            triggered=self._open_convert_drawing_dialog,
        ))
        export_menu.addAction(QAction(
            "从CATPart/CATProduct导出stp", self,
            triggered=self._open_convert_part_dialog,
        ))

        # BOM
        bom_menu = bar.addMenu("BOM")
        bom_menu.addAction(QAction(
            "从CATProduct导出BOM", self,
            triggered=self._open_export_bom_dialog,
        ))
        bom_menu.addAction(QAction(
            "BOM属性补全", self, triggered=self._open_bom_edit_dialog
        ))

        # Drawing
        drawing_menu = bar.addMenu("图纸")
        drawing_menu.addAction(QAction(
            "从CATPart/CATProduct生成图纸", self,
            triggered=self._open_generate_drawing_dialog,
        ))
        drawing_menu.addAction(QAction(
            "刷新图纸信息", self,
            triggered=self._open_refresh_drawing_dialog,
        ))

        # Macro
        self._macro_menu = bar.addMenu("宏")
        self._rebuild_macro_menu()

        # Tools
        tools_menu = bar.addMenu("工具")
        for label, slot in (
            ("复制字体文件到CATIA目录",  self._copy_font_to_catia),
            ("复制ISO.xml到CATIA目录",    self._copy_iso_to_catia),
            ("Crack",                     self._crack),
            ("刷写零件模板",              self._open_stamp_part_template_dialog),
            ("查找所有依赖项（未实现）",     self._open_find_dependencies_dialog),
        ):
            tools_menu.addAction(QAction(label, self, triggered=slot))

        # View
        view_menu = bar.addMenu("视图")
        view_menu.addAction(QAction(
            "放大", self,
            triggered=lambda: QMessageBox.information(self, "提示", "功能尚未实现"),
        ))
        view_menu.addAction(QAction(
            "缩小", self,
            triggered=lambda: QMessageBox.information(self, "提示", "功能尚未实现"),
        ))
        view_menu.addAction(QAction(
            "重置缩放", self,
            triggered=lambda: self.resize(MAIN_WINDOW_DEFAULT_WIDTH, MAIN_WINDOW_DEFAULT_HEIGHT),
        ))
        view_menu.addSeparator()
        self._show_log_action = QAction("显示Log", self)
        self._show_log_action.setCheckable(True)
        self._show_log_action.toggled.connect(self._toggle_log_window)
        view_menu.addAction(self._show_log_action)

        # Help
        help_menu = bar.addMenu("帮助")
        help_menu.addAction(QAction(
            "文档", self,
            triggered=self._show_help,
        ))
        help_menu.addAction(QAction(
            f"关于 {APP_NAME}", self,
            triggered=self._show_about,
        ))

    # ── Log window ─────────────────────────────────────────────────────────

    def _toggle_log_window(self, checked: bool) -> None:
        if checked:
            self._log_window.show()
            self._log_window.raise_()
        else:
            self._log_window.hide()

    def _show_about(self) -> None:
        QMessageBox.about(self, f"About {APP_NAME}", ABOUT_TEXT)

    def _show_help(self) -> None:
        HelpDialog(self).exec()

    # ── Macro menu helpers ─────────────────────────────────────────────────

    def _macros_dir(self) -> Path:
        return resource_path("macros")

    def _rebuild_macro_menu(self) -> None:
        self._macro_menu.clear()
        macros_dir   = self._macros_dir()
        macro_files: list[Path] = []
        if macros_dir.is_dir():
            macro_files = sorted(
                f for f in macros_dir.iterdir()
                if f.is_file() and f.suffix.lower() in self._MACRO_EXTENSIONS
            )

        if macro_files:
            for macro_path in macro_files:
                self._macro_menu.addAction(QAction(
                    macro_path.name, self,
                    triggered=lambda checked=False, p=macro_path: self._run_macro(p),
                ))
            self._macro_menu.addSeparator()
        else:
            placeholder = QAction("（未找到宏文件）", self)
            placeholder.setEnabled(False)
            self._macro_menu.addAction(placeholder)
            self._macro_menu.addSeparator()

        self._macro_menu.addAction(QAction(
            "打开宏文件夹", self, triggered=self._open_macros_folder
        ))
        self._macro_menu.addAction(QAction(
            "刷新宏列表", self, triggered=self._rebuild_macro_menu
        ))

    def _open_macros_folder(self) -> None:
        macros_dir = self._macros_dir()
        macros_dir.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(str(macros_dir))
            else:
                subprocess.Popen(
                    ["xdg-open", str(macros_dir)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
        except Exception as e:
            QMessageBox.warning(
                self, "无法打开文件夹", f"无法打开宏文件夹：\n{macros_dir}\n\n{e}"
            )

    def _run_macro(self, macro_path: Path) -> None:
        if not macro_path.exists():
            QMessageBox.warning(self, "文件不存在", f"宏文件不存在：\n{macro_path}")
            return
        try:
            from pycatia import catia as _catia
            caa = _catia()
            app = caa.application
            # 不传递额外参数，直接运行宏入口函数 CATMain
            self._execute_catscript(app, macro_path, "CATMain", [])
            logger.info(f"宏执行成功：{macro_path.name}")
        except Exception as e:
            logger.error(f"宏执行失败 {macro_path.name}: {e}")
            QMessageBox.critical(
                self, "宏执行失败",
                f"运行宏时出错：\n{macro_path.name}\n\n{e}\n\n请确保CATIA已启动。",
            )

    # ── Dialog launchers ───────────────────────────────────────────────────

    def _open_convert_part_dialog(self) -> None:
        FileConvertDialog(
            parent=self,
            title="将CATPart/CATProduct导出为STP",
            file_label="已选CATPart/CATProduct文件:",
            file_filter="*.CATPart *.CATProduct (*.CATPart *.CATProduct);;All Files (*)",
            no_files_msg="请至少选择一个CATPart或CATProduct文件。",
            conversion_fn=convert_part_to_step,
            settings_key="CATPart",
            show_prefix_option=True,
            prefix="MD_",
            note="暂时留空",
        ).exec()

    def _open_convert_drawing_dialog(self) -> None:
        FileConvertDialog(
            parent=self,
            title="将CATDrawing导出为PDF",
            file_label="已选CATDrawing文件:",
            file_filter="*.CATDrawing (*.CATDrawing);;All Files (*)",
            no_files_msg="请至少选择一个CATDrawing文件。",
            conversion_fn=convert_drawing_to_pdf,
            settings_key="CATDrawing",
            show_prefix_option=True,
            prefix="DR_",
            show_update_option=True,
            note=(
                "如果用于导出的CATDrawing有多页，请将CATIA设置为"
                "\u201c将多页文档保存在单向量文件中\u201d"
                "（工具->选项->常规->兼容性->图形格式->导出（另存为））"
            ),
        ).exec()

    def _open_export_bom_dialog(self) -> None:
        ExportBomDialog(self).exec()

    def _open_bom_edit_dialog(self) -> None:
        BomEditDialog(self).exec()

    def _open_stamp_part_template_dialog(self) -> None:
        FileConvertDialog(
            parent=self,
            title="刷写零件模板",
            file_label="已选CATPart文件:",
            file_filter="*.CATPart (*.CATPart);;All Files (*)",
            no_files_msg="请至少选择一个CATPart文件。",
            conversion_fn=apply_part_template,
            settings_key="StampPartTemplate",
            show_active_doc_option=True,
        ).exec()

    def _open_find_dependencies_dialog(self) -> None:
        FindDependenciesDialog(self).exec()

    # ── Drawing generation ─────────────────────────────────────────────────

    def _drawing_templates_dir(self) -> Path:
        return resource_path("drawing_templates")

    def _open_generate_drawing_dialog(self) -> None:
        templates_dir = self._drawing_templates_dir()
        templates_dir.mkdir(parents=True, exist_ok=True)

        templates = sorted(templates_dir.glob("*.CATDrawing"))
        if not templates:
            QMessageBox.warning(
                self, "未找到模板",
                f"在以下目录中未找到任何 CATDrawing 模板文件：\n{templates_dir}\n\n"
                "请将 *.CATDrawing 模板放入该文件夹后重试。",
            )
            return

        name, ok = QInputDialog.getItem(
            self,
            "选择图纸模板",
            "请选择一个 CATDrawing 模板：",
            [t.name for t in templates],
            0,
            False,
        )
        if not ok:
            return

        template_path = templates_dir / name

        # 优先使用同名的 .catvbs 脚本；若不存在则提示用户
        catvbs_path = self._macros_dir() / "generate_drawing.catvbs"
        if not catvbs_path.exists():
            QMessageBox.warning(
                self, "宏文件未找到",
                f"未找到 CATScript 宏文件：\n{catvbs_path}\n\n"
                "请将 generate_drawing.catvbs 放入 macros 文件夹后重试。",
            )
            return
        self._run_template_macro(catvbs_path, str(template_path))

    def _open_refresh_drawing_dialog(self) -> None:
        """刷新当前活动图纸的参数信息（通过 refresh_drawing_info.catvbs 宏）。"""
        catvbs_path = self._macros_dir() / "refresh_drawing_info.catvbs"
        if not catvbs_path.exists():
            QMessageBox.warning(
                self, "宏文件未找到",
                f"未找到 CATScript 宏文件：\n{catvbs_path}\n\n"
                "请将 refresh_drawing_info.catvbs 放入 macros 文件夹后重试。",
            )
            return
        self._run_macro(catvbs_path)

    def _execute_catscript(
        self,
        app,
        macro_path: Path,
        func_name: str,
        params: list,
    ) -> None:
        """调用 CATIA SystemService.ExecuteScript 执行 CATScript 宏（.catvbs / .catscript）。

        CATIA ExecuteScript 签名::

            SystemService.ExecuteScript(iLibraryName, iLibraryType,
                                        iProgramName, iFunctionName, iParameters)

        此处使用 iLibraryType=1（目录模式）：
          - iLibraryName：宏文件所在目录
          - iProgramName：宏文件名（含扩展名）
          - iFunctionName：要调用的函数/子程序名（通常为 "CATMain"）
          - iParameters：传递给宏的参数列表
        """
        lib_dir = str(macro_path.parent)
        app.com_object.SystemService.ExecuteScript(
            lib_dir, 1, macro_path.name, func_name, params
        )

    def _run_template_macro(
        self,
        macro_path: Path,
        template_path: str,
    ) -> None:
        """通过 CATIA SystemService.ExecuteScript 运行指定的 CATScript 宏，
        并将模板文件路径作为参数传入，宏内可通过 iParameters 直接获取。
        """
        try:
            from pycatia import catia as _catia
            caa = _catia()
            app = caa.application
            # 将模板路径作为单一字符串参数传递给宏的 CATMain 函数
            self._execute_catscript(app, macro_path, "CATMain", [template_path])
            logger.info(f"宏执行成功：{macro_path.name} | 模板路径={template_path}")
        except Exception as e:
            logger.error(f"宏执行失败 {macro_path.name}: {e}")
            QMessageBox.critical(
                self, "宏执行失败",
                f"运行宏时出错：\n{macro_path.name}\n\n{e}\n\n请确保CATIA已启动。",
            )

    # ── CATIA resource file helpers ────────────────────────────────────────

    def _copy_font_to_catia(self) -> None:
        self._copy_file_to_catia(
            file_name=FONT_FILE_PATH,
            relative_dest=Path("win_b64") / "resources" / "fonts" / "TrueType",
        )

    def _copy_iso_to_catia(self) -> None:
        self._copy_file_to_catia(
            file_name=ISO_XML_FILE_PATH,
            relative_dest=Path("win_b64") / "resources" / "standard" / "drafting",
        )

    def _copy_file_to_catia(self, file_name: str, relative_dest: Path) -> None:
        src_file = resource_path(file_name)
        base_name = Path(file_name).name
        if not src_file.exists():
            QMessageBox.warning(
                self, "文件未找到",
                f"在工作目录中找不到 '{base_name}'：\n{src_file.parent}",
            )
            return

        catia_root = detect_catia_root()
        if catia_root:
            reply = QMessageBox.question(
                self, "检测到CATIA安装",
                f"检测到CATIA安装路径：\n{catia_root}\n\n是否使用该目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(
                self,
                "选择CATIA安装目录（例如 C:\\Program Files\\Dassault Systemes\\B28）",
                "",
            )
            if not catia_root:
                return

        dest_dir = Path(catia_root) / relative_dest
        if not dest_dir.exists():
            reply = QMessageBox.question(
                self, "文件夹未找到",
                f"目标文件夹不存在：\n{dest_dir}\n\n是否要创建该文件夹？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                dest_dir.mkdir(parents=True, exist_ok=True)
            else:
                return

        dest_file = dest_dir / base_name
        try:
            shutil.copy2(str(src_file), str(dest_file))
            QMessageBox.information(
                self, "成功",
                f"'{base_name}' 已成功复制到：\n{dest_file}",
            )
        except PermissionError:
            QMessageBox.critical(
                self, "权限不足",
                f"无法复制文件，请以管理员身份运行程序。\n\n目标路径：\n{dest_file}",
            )
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生意外错误：\n{e}")

    def _crack(self) -> None:
        src_dir = resource_path(CRACK_DIR_PATH)
        if not src_dir.exists() or not src_dir.is_dir():
            QMessageBox.warning(
                self, "文件夹未找到",
                f"找不到 'crack' 文件夹：\n{src_dir.parent}",
            )
            return

        catia_root = detect_catia_root()
        if catia_root:
            reply = QMessageBox.question(
                self, "检测到CATIA安装",
                f"检测到CATIA安装路径：\n{catia_root}\n\n是否使用该目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                catia_root = None

        if not catia_root:
            catia_root = QFileDialog.getExistingDirectory(
                self,
                "选择CATIA安装目录（例如 C:\\Program Files\\Dassault Systemes\\B28）",
                "",
            )
            if not catia_root:
                return

        dest_dir = Path(catia_root) / "win_b64" / "code" / "bin"
        if not dest_dir.exists():
            QMessageBox.critical(
                self, "文件夹未找到",
                f"目标文件夹不存在：\n{dest_dir}\n\n请检查您的CATIA安装。",
            )
            return

        files = [f for f in src_dir.iterdir() if f.is_file()]
        if not files:
            QMessageBox.warning(self, "文件夹为空", "'crack' 文件夹中没有文件。")
            return

        try:
            copied: list[str] = []
            for src_file in files:
                dest_file = dest_dir / src_file.name
                shutil.copy2(str(src_file), str(dest_file))
                copied.append(src_file.name)
                logger.info(f"  Copied: {src_file.name} -> {dest_file}")
            QMessageBox.information(
                self, "成功",
                f"已成功复制 {len(copied)} 个文件到：\n{dest_dir}\n\n"
                + "\n".join(copied),
            )
        except PermissionError:
            QMessageBox.critical(
                self, "权限不足",
                f"无法复制文件，请以管理员身份运行程序。\n\n目标路径：\n{dest_dir}",
            )
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生意外错误：\n{e}")
