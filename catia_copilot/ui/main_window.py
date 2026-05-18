"""
主应用程序窗口模块。

提供：
- MainWindow – 带有分组按钮 UI 和菜单栏的主 QMainWindow。
"""

import sys
import os
import re
import shutil
import subprocess
import logging
from collections.abc import Callable
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QMessageBox, QPushButton, QFileDialog, QGroupBox, QInputDialog,
    QDialog,
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QTimer

from catia_copilot.constants import (
    APP_NAME,
    ABOUT_TEXT,
    MAIN_WINDOW_DEFAULT_WIDTH,
    MAIN_WINDOW_DEFAULT_HEIGHT,
    FONT_FILE_PATH,
    ISO_XML_FILE_PATH,
    CRACK_DIR_PATH,
)
from catia_copilot.utils import resource_path, detect_catia_root, check_catia_connection, diagnose_catia_connection
from catia_copilot.logging_setup import log_signal_emitter
from catia_copilot.catia.conversion import convert_drawing_to_pdf, convert_part_to_step
from catia_copilot.catia.template import apply_part_template
from catia_copilot.ui.log_window import LogWindow
from catia_copilot.ui.convert_dialog import FileConvertDialog
from catia_copilot.ui.export_bom_dialog import ExportBomDialog
from catia_copilot.ui.find_deps_dialog import FindDependenciesDialog
from catia_copilot.ui.bom_edit_dialog import BomEditDialog
from catia_copilot.ui.mass_props_dialog import MassPropsDialog
from catia_copilot.ui.help_dialog import HelpDialog
from catia_copilot.ui.plm_sync_dialog import PlmSyncDialog

logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """主应用程序窗口。"""

    # 快速运行宏支持 CATScript（.catvbs / .catscript）和 VBA（.catvba）文件。
    _MACRO_EXTENSIONS: frozenset[str] = frozenset({".catvbs", ".catscript", ".catvba"})

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(MAIN_WINDOW_DEFAULT_WIDTH, MAIN_WINDOW_DEFAULT_HEIGHT)

        self._log_window = LogWindow(self)
        log_signal_emitter.message_logged.connect(self._log_window.append_log)

        self._build_menu_bar()
        self._build_central_widget()
        self._build_connection_indicator()
        self.statusBar().showMessage("就绪")

    # ── CATIA 连接状态指示器 ──────────────────────────────────────────────

    def _build_connection_indicator(self) -> None:
        """在状态栏右侧添加 CATIA 连接状态指示标签，并启动定时轮询。"""
        self._catia_status_label = QLabel()
        self._catia_status_label.setObjectName("catiaStatusLabel")
        self._catia_status_label.setToolTip(
            "CATIA V5 COM 连接状态（每 5 秒自动刷新）\n"
            "橙色表示 COM 对象可获取但功能测试失败，\n"
            "可通过菜单「帮助 -> CATIA 连接诊断」查看详情"
        )
        self.statusBar().addPermanentWidget(self._catia_status_label)

        # 立即检测一次，再每 5 秒轮询一次
        self._update_connection_status()
        self._connection_timer = QTimer(self)
        self._connection_timer.setInterval(5000)
        self._connection_timer.timeout.connect(self._update_connection_status)
        self._connection_timer.start()

    def _update_connection_status(self) -> None:
        """轮询 CATIA 连接状态并更新指示标签的文字和样式。"""
        status = check_catia_connection()
        if status == "connected":
            self._catia_status_label.setText("● CATIA 已连接")
            self._catia_status_label.setProperty("catiaConnected", "true")
        elif status == "broken":
            self._catia_status_label.setText("⚠ CATIA 连接异常")
            self._catia_status_label.setProperty("catiaConnected", "broken")
        else:
            self._catia_status_label.setText("● CATIA 未连接")
            self._catia_status_label.setProperty("catiaConnected", "false")
        # 强制重新应用 QSS（动态属性变化后需要刷新样式）
        self._catia_status_label.style().unpolish(self._catia_status_label)
        self._catia_status_label.style().polish(self._catia_status_label)

    def _show_catia_diagnostics(self) -> None:
        """运行 CATIA COM 详细诊断并以对话框形式呈现结果。"""
        info = diagnose_catia_connection()
        status = info["status"]
        is_elevated = bool(info.get("is_elevated"))
        catia_running = bool(info.get("catia_process_running"))

        status_text = {
            "connected":    "✅ 已连接（功能测试通过）",
            "broken":       "⚠️ 连接异常",
            "disconnected": "❌ 未连接",
        }.get(status, status)

        elevated_text = "是（管理员）" if is_elevated else "否（普通用户）"
        process_text  = "运行中" if catia_running else "未检测到"

        lines = [
            f"<b>连接状态：</b>{status_text}",
            f"<b>本程序权限：</b>{elevated_text}",
            f"<b>CNEXT.exe 进程：</b>{process_text}",
        ]

        # ── 已连接：显示连接细节 ─────────────────────────────────────────
        if status == "connected":
            if info["app_name"]:
                lines.append(f"<b>应用名称：</b>{info['app_name']}")
            if info.get("is_v5") is not None:
                lines.append(
                    "<b>产品类型：</b>CATIA V5 ✅"
                    if info["is_v5"]
                    else "<b>产品类型：</b>3DEXPERIENCE ⚠️"
                )
            if info["doc_count"] is not None:
                lines.append(f"<b>已打开文档数：</b>{info['doc_count']}")
            if info["active_doc"]:
                lines.append(f"<b>当前活动文档：</b>{info['active_doc']}")
            else:
                lines.append("<b>当前活动文档：</b>（无）")

        # ── 连接异常：区分权限不匹配方向 ────────────────────────────────
        elif status == "broken" and catia_running:
            if is_elevated:
                # 本程序管理员，CATIA 普通用户
                lines += [
                    "",
                    "<b>根本原因：</b>本程序以<b>管理员</b>权限运行，CATIA 以<b>普通用户</b>"
                    "权限运行。Windows UAC 隔离机制导致管理员进程无法看到普通用户进程注册的"
                    " ROT 对象。",
                    "<b>解决方案：</b>以<b>普通用户身份（不提权）</b>直接运行本程序。",
                ]
            else:
                lines += [
                    "",
                    "<b>根本原因：</b>CATIA 进程存在，但所有 COM 连接方式均失败。"
                    "最常见原因：CATIA 以<b>管理员</b>权限运行，而本程序以<b>普通用户</b>"
                    "权限运行（UAC ROT 隔离）。",
                    "<b>解决方案：</b>将 CATIA 改为<b>普通用户</b>权限运行（取消「以管理员身份运行」），"
                    "使两侧权限级别一致。",
                ]

        # ── 未连接 ───────────────────────────────────────────────────────
        elif status == "disconnected":
            lines += [
                "",
                "<b>原因：</b>未检测到运行中的 CATIA V5 进程。",
                "<b>建议：</b>请先启动 CATIA V5，再重试。",
            ]

        html = "<br/>".join(lines)

        msg = QMessageBox(self)
        msg.setWindowTitle("CATIA 连接诊断")
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(html)
        msg.exec()

    @staticmethod
    def _detect_crack_version_subdir(catia_root: str) -> str | None:
        """从 CATIA 安装路径末尾（如 B28、B33）推断 crack 子目录名（如 R28、R33）。

        例：``C:\\Program Files\\Dassault Systemes\\B33`` → ``"R33"``
        若路径末尾不符合 ``B\\d+`` 格式，则返回 None。
        """
        name = Path(catia_root).name.upper()
        m = re.match(r"^B(\d+)$", name)
        if m:
            return f"R{m.group(1)}"
        return None

    def _run_copy_elevated(self, operations: list[tuple[Path, Path]]) -> bool:
        """以管理员权限批量复制文件（ShellExecuteExW + WaitForSingleObject）。

        写入临时批处理文件，通过 UAC「runas」动词以管理员身份静默执行，
        同步等待完成后清理临时文件。

        :param operations: ``[(src_path, dest_path), ...]`` 复制操作列表。
        :returns: ``True`` 表示 UAC 提权已接受并等待完成；
                  ``False`` 表示用户取消 UAC 或系统调用失败。
                  返回 True **不保证**文件一定写入成功，调用方需自行验证目标文件。
        """
        import ctypes
        from ctypes import wintypes
        import tempfile

        if not operations:
            return True

        # 构建批处理文件内容（若目标目录不存在则先创建）
        lines = ["@echo off"]
        for src, dest in operations:
            parent = str(dest.parent)
            lines.append(f'if not exist "{parent}" mkdir "{parent}"')
            lines.append(f'copy /Y "{src}" "{dest}"')
        bat_content = "\r\n".join(lines) + "\r\n"

        # 写入临时 .bat 文件（CATIA 路径通常为 ASCII，GBK 可安全表示）
        bat_path = ""
        try:
            fd, bat_path = tempfile.mkstemp(suffix=".bat")
            with os.fdopen(fd, "w", encoding="gbk", errors="replace") as f:
                f.write(bat_content)
        except Exception as exc:
            logger.warning(f"创建临时批处理文件失败：{exc}")
            return False

        # SHELLEXECUTEINFOW 结构体（Windows SDK 定义）
        SEE_MASK_NOCLOSEPROCESS = 0x00000040
        SW_HIDE = 0

        class _SHELLEXECUTEINFOW(ctypes.Structure):
            _fields_ = [
                ("cbSize",         wintypes.DWORD),
                ("fMask",          wintypes.ULONG),
                ("hwnd",           wintypes.HWND),
                ("lpVerb",         wintypes.LPCWSTR),
                ("lpFile",         wintypes.LPCWSTR),
                ("lpParameters",   wintypes.LPCWSTR),
                ("lpDirectory",    wintypes.LPCWSTR),
                ("nShow",          ctypes.c_int),
                ("hInstApp",       wintypes.HINSTANCE),
                ("lpIDList",       ctypes.c_void_p),
                ("lpClass",        wintypes.LPCWSTR),
                ("hkeyClass",      wintypes.HKEY),
                ("dwHotKey",       wintypes.DWORD),
                ("hIconOrMonitor", wintypes.HANDLE),   # union hIcon/hMonitor
                ("hProcess",       wintypes.HANDLE),
            ]

        sei = _SHELLEXECUTEINFOW()
        sei.cbSize = ctypes.sizeof(sei)
        sei.fMask = SEE_MASK_NOCLOSEPROCESS
        sei.hwnd = None
        sei.lpVerb = "runas"
        sei.lpFile = "cmd.exe"
        sei.lpParameters = f'/c "{bat_path}"'
        sei.lpDirectory = None
        sei.nShow = SW_HIDE

        accepted = False
        try:
            ok = bool(ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei)))
            if ok and sei.hProcess:
                # 等待 cmd.exe 执行完毕（最多 60 秒）
                WAIT_TIMEOUT = 0x00000102
                ret = ctypes.windll.kernel32.WaitForSingleObject(sei.hProcess, 60000)
                ctypes.windll.kernel32.CloseHandle(sei.hProcess)
                if ret == WAIT_TIMEOUT:
                    logger.warning("提权复制操作等待超时（60 秒），请手动确认结果。")
                accepted = True
            else:
                logger.info("ShellExecuteExW 返回失败，用户可能取消了 UAC 提权。")
        except Exception as exc:
            logger.warning(f"ShellExecuteExW 调用异常：{exc}")
        finally:
            try:
                if bat_path:
                    os.unlink(bat_path)
            except Exception:
                pass

        return accepted

    # ── 中央控件区域 ──────────────────────────────────────────────────────

    def _build_central_widget(self) -> None:
        """构建主窗口的中央控件区域。"""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 16)
        layout.setSpacing(12)

        # 欢迎标签
        welcome = QLabel(f"欢迎使用 {APP_NAME}")
        welcome.setAlignment(Qt.AlignmentFlag.AlignCenter)
        welcome.setStyleSheet("font-size: 11pt; font-weight: bold; color: #333;")
        layout.addWidget(welcome)

        # ── 导出功能组 ────────────────────────────────────────────────────
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

        # ── BOM 功能组 ────────────────────────────────────────────────────
        bom_group  = QGroupBox("BOM")
        bom_layout = QVBoxLayout(bom_group)
        bom_layout.setSpacing(6)

        btn_bom_export = QPushButton("从 CATProduct 导出 BOM")
        btn_bom_export.setToolTip("从 CATProduct 导出 BOM 到 Excel 文件")
        btn_bom_export.clicked.connect(self._open_export_bom_dialog)

        btn_bom_edit = QPushButton("BOM 属性补全")
        btn_bom_edit.setToolTip("在表格中编辑 BOM 属性并写回 CATIA")
        btn_bom_edit.clicked.connect(self._open_bom_edit_dialog)

        btn_mass_props = QPushButton("重量、重心、惯量统计")
        btn_mass_props.setToolTip(
            "遍历产品树，读取零件质量/重心/转动惯量，计算装配体总质量特性并导出"
        )
        btn_mass_props.clicked.connect(self._open_mass_props_dialog)

        btn_plm_sync = QPushButton("同步 BOM 到 PLM")
        btn_plm_sync.setToolTip(
            "将当前 CATIA 产品结构（BOM）同步到 DocdokuPLM 服务端，\n"
            "自动创建零件、写入属性并上传 STEP 几何文件"
        )
        btn_plm_sync.clicked.connect(self._open_plm_sync_dialog)

        for btn in (btn_bom_export, btn_bom_edit, btn_mass_props, btn_plm_sync):
            bom_layout.addWidget(btn)
        layout.addWidget(bom_group)

        # ── 图纸功能组 ────────────────────────────────────────────────────
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

        # ── 工具功能组 ────────────────────────────────────────────────────
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

        # 快速装配：紧固件 + 托板螺母 并列一行
        assembly_row = QHBoxLayout()
        assembly_row.setSpacing(6)

        btn_fastener = QPushButton("快速装配紧固件")
        btn_fastener.setToolTip("在装配体中连续放置紧固件实例")
        btn_fastener.clicked.connect(self._open_fastener_assembly_dialog)

        btn_nut_plate = QPushButton("快速装配托板螺母")
        btn_nut_plate.setToolTip("在装配体中连续放置托板螺母实例")
        btn_nut_plate.clicked.connect(self._open_nut_plate_assembly_dialog)

        assembly_row.addWidget(btn_fastener)
        assembly_row.addWidget(btn_nut_plate)
        tools_layout.addLayout(assembly_row)
        layout.addWidget(tools_group)

        layout.addStretch()

    # ── 菜单栏 ────────────────────────────────────────────────────────────

    def _show_not_implemented(self) -> None:
        """显示"功能尚未实现"提示对话框。"""
        QMessageBox.information(self, "提示", "功能尚未实现")

    def _build_menu_bar(self) -> None:
        """构建应用程序菜单栏。"""
        bar = self.menuBar()

        # 文件菜单
        file_menu  = bar.addMenu("文件")
        for label, slot in (("新建", self._show_not_implemented), ("打开...", self._show_not_implemented),
                             ("保存", self._show_not_implemented), ("另存为...", self._show_not_implemented)):
            file_menu.addAction(QAction(label, self, triggered=slot))
        file_menu.addSeparator()
        file_menu.addAction(QAction("退出", self, triggered=self.close))

        # 导出菜单
        export_menu = bar.addMenu("导出")
        export_menu.addAction(QAction(
            "从CATDrawing导出pdf", self,
            triggered=self._open_convert_drawing_dialog,
        ))
        export_menu.addAction(QAction(
            "从CATPart/CATProduct导出stp", self,
            triggered=self._open_convert_part_dialog,
        ))

        # BOM 菜单
        bom_menu = bar.addMenu("BOM")
        bom_menu.addAction(QAction(
            "从CATProduct导出BOM", self,
            triggered=self._open_export_bom_dialog,
        ))
        bom_menu.addAction(QAction(
            "BOM属性补全", self, triggered=self._open_bom_edit_dialog
        ))
        bom_menu.addAction(QAction(
            "同步BOM到PLM", self, triggered=self._open_plm_sync_dialog
        ))

        # 图纸菜单
        drawing_menu = bar.addMenu("图纸")
        drawing_menu.addAction(QAction(
            "从CATPart/CATProduct生成图纸", self,
            triggered=self._open_generate_drawing_dialog,
        ))
        drawing_menu.addAction(QAction(
            "刷新图纸信息", self,
            triggered=self._open_refresh_drawing_dialog,
        ))

        # 宏菜单
        self._macro_menu = bar.addMenu("宏")
        self._rebuild_macro_menu()

        # 工具菜单
        tools_menu = bar.addMenu("工具")
        for label, slot in (
            ("复制字体文件到CATIA目录",  self._copy_font_to_catia),
            ("复制ISO.xml到CATIA目录",    self._copy_iso_to_catia),
            ("Crack",                     self._crack),
            ("刷写零件模板",              self._open_stamp_part_template_dialog),
            ("查找所有依赖项（未实现）",     self._open_find_dependencies_dialog),
            ("快速装配紧固件",              self._open_fastener_assembly_dialog),
            ("快速装配托板螺母",            self._open_nut_plate_assembly_dialog),
        ):
            tools_menu.addAction(QAction(label, self, triggered=slot))

        # 视图菜单
        view_menu = bar.addMenu("视图")
        view_menu.addAction(QAction(
            "放大", self,
            triggered=self._show_not_implemented,
        ))
        view_menu.addAction(QAction(
            "缩小", self,
            triggered=self._show_not_implemented,
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

        # 帮助菜单
        help_menu = bar.addMenu("帮助")
        help_menu.addAction(QAction(
            "文档", self,
            triggered=self._show_help,
        ))
        help_menu.addAction(QAction(
            "CATIA 连接诊断", self,
            triggered=self._show_catia_diagnostics,
        ))
        help_menu.addAction(QAction(
            f"关于 {APP_NAME}", self,
            triggered=self._show_about,
        ))

    # ── 日志窗口 ──────────────────────────────────────────────────────────

    def _toggle_log_window(self, checked: bool) -> None:
        """切换日志窗口的显示/隐藏状态。"""
        if checked:
            self._log_window.show()
            self._log_window.raise_()
        else:
            self._log_window.hide()

    def _show_about(self) -> None:
        """显示关于对话框。"""
        QMessageBox.about(self, f"About {APP_NAME}", ABOUT_TEXT)

    def _show_help(self) -> None:
        """显示帮助文档对话框。"""
        self._show_dialog("_dlg_help", lambda: HelpDialog(self))

    # ── 非模态对话框管理 ──────────────────────────────────────────────────

    def _show_dialog(self, attr: str, factory: Callable[[], QDialog]) -> None:
        """以非模态方式打开对话框，若已存在则将其置于前台。

        所有通过此方法打开的对话框均被设置为独立顶级窗口，使其在 Windows
        任务栏中拥有独立条目并可单独最小化，同时仍与主窗口归属同一应用程序分组。

        关键点：仅调用 setWindowFlags() 无法让对话框在任务栏独立显示，因为
        dialog 仍持有 Qt 父窗口引用，Qt 在创建原生窗口时会将父窗口作为 Win32
        "Owner 窗口"传给 CreateWindowEx。Windows 规定有 Owner 的窗口不会在
        任务栏单独出现。必须通过 setParent(None, flags) 同时清除父引用并设置
        窗口类型，确保原生窗口创建时无 Owner，从而获得独立的任务栏条目。

        :param attr: 用于在 MainWindow 上缓存对话框实例的属性名。
        :param factory: 无参可调用对象，返回新的 QDialog 实例。
        """
        dlg = getattr(self, attr, None)
        if dlg is None:
            dlg = factory()
            # setParent(None, flags) 同时：
            #   1. 清除 Qt 父引用（使 Win32 原生窗口创建时无 Owner）
            #   2. 设置 Window 类型和标准按钮标志
            # 这是使对话框出现在任务栏的必要条件；单独调用 setWindowFlags()
            # 无效，因为那不会断开已有的 Qt 父子关系和 Win32 Owner 关联。
            dlg.setParent(
                None,
                Qt.WindowType.Window
                | Qt.WindowType.WindowTitleHint
                | Qt.WindowType.WindowSystemMenuHint
                | Qt.WindowType.WindowCloseButtonHint
                | Qt.WindowType.WindowMaximizeButtonHint
                | Qt.WindowType.WindowMinimizeButtonHint,
            )
            dlg.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
            dlg.destroyed.connect(lambda _=None, a=attr: setattr(self, a, None))
            setattr(self, attr, dlg)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def closeEvent(self, event) -> None:  # noqa: N802
        """主窗口关闭时，同时关闭所有通过 _show_dialog 打开的子窗口。

        由于子窗口通过 ``setParent(None, ...)`` 清除了 Qt 父引用（以获得独立
        的任务栏条目），Qt 的默认父子关闭机制对其无效，需在此手动关闭。
        所有子窗口均设有 ``WA_DeleteOnClose``，close() 会触发其销毁和清理。
        """
        for attr, value in list(vars(self).items()):
            if attr.startswith("_dlg_") and isinstance(value, QDialog):
                value.close()
        super().closeEvent(event)

    # ── 宏菜单辅助方法 ────────────────────────────────────────────────────

    def _macros_dir(self) -> Path:
        """返回宏文件夹路径。"""
        return resource_path("macros")

    def _rebuild_macro_menu(self) -> None:
        """重建宏菜单，扫描 macros 文件夹并添加菜单项。"""
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
            from catia_copilot.catia.connection import get_catia_v5_application as _catia
            caa = _catia()
            app = caa.application
            if macro_path.suffix.lower() == ".catvba":
                self._execute_catvba(app, macro_path, "CATMain", [])
            else:
                # .catvbs / .catscript — CATScript 目录模式
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
        self._show_dialog("_dlg_convert_part", lambda: FileConvertDialog(
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
        ))

    def _open_convert_drawing_dialog(self) -> None:
        self._show_dialog("_dlg_convert_drawing", lambda: FileConvertDialog(
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
        ))

    def _open_export_bom_dialog(self) -> None:
        self._show_dialog("_dlg_export_bom", lambda: ExportBomDialog(self))

    def _open_bom_edit_dialog(self) -> None:
        self._show_dialog("_dlg_bom_edit", lambda: BomEditDialog(self))

    def _open_mass_props_dialog(self) -> None:
        self._show_dialog("_dlg_mass_props", lambda: MassPropsDialog(self))

    def _open_plm_sync_dialog(self) -> None:
        def factory():
            dlg = PlmSyncDialog(self)
            # 同步运行期间暂停 CATIA 连接检查，避免 COM 调用持有 GIL 阻塞 Worker 线程
            dlg.sync_started.connect(self._connection_timer.stop)
            dlg.sync_done.connect(self._connection_timer.start)
            return dlg
        self._show_dialog("_dlg_plm_sync", factory)

    def _open_stamp_part_template_dialog(self) -> None:
        self._show_dialog("_dlg_stamp_template", lambda: FileConvertDialog(
            parent=self,
            title="刷写零件模板",
            file_label="已选CATPart文件:",
            file_filter="*.CATPart (*.CATPart);;All Files (*)",
            no_files_msg="请至少选择一个CATPart文件。",
            conversion_fn=apply_part_template,
            settings_key="StampPartTemplate",
            show_active_doc_option=True,
        ))

    def _open_find_dependencies_dialog(self) -> None:
        self._show_dialog("_dlg_find_deps", lambda: FindDependenciesDialog(self))

    def _open_fastener_assembly_dialog(self) -> None:
        """直接运行 macros 文件夹中的 fastener_assembly.catvba VBA 宏。"""
        catvba_path = self._macros_dir() / "fastener_assembly.catvba"
        if not catvba_path.exists():
            QMessageBox.warning(
                self, "宏文件未找到",
                f"未找到 VBA 宏文件：\n{catvba_path}\n\n"
                "请在 CATIA VBA 编辑器中按照 macros/fastener_assembly.txt 创建宏，\n"
                "将 VBA 项目导出为 fastener_assembly.catvba，\n"
                "并放入 macros 文件夹后重试。",
            )
            return
        self._run_macro(catvba_path)

    def _open_nut_plate_assembly_dialog(self) -> None:
        """直接运行 macros 文件夹中的 nut_plate_assembly.catvba VBA 宏。"""
        catvba_path = self._macros_dir() / "nut_plate_assembly.catvba"
        if not catvba_path.exists():
            QMessageBox.warning(
                self, "宏文件未找到",
                f"未找到 VBA 宏文件：\n{catvba_path}\n\n"
                "请在 CATIA VBA 编辑器中按照 macros/nut_plate_assembly.txt 创建宏，\n"
                "将 VBA 项目导出为 nut_plate_assembly.catvba，\n"
                "并放入 macros 文件夹后重试。",
            )
            return
        self._run_macro(catvba_path)

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

    def _execute_catvba(
        self,
        app,
        macro_path: Path,
        func_name: str,
        params: list,
    ) -> None:
        """调用 CATIA SystemService.ExecuteScript 执行 VBA 宏（.catvba）。

        此处使用 iLibraryType=2（VBA 项目文件模式）：
          - iLibraryName：.catvba 文件完整路径
          - iProgramName：VBA 模块名（中文 CATIA 默认为 "模块1"，英文/法语等环境为 "Module1"）
          - iFunctionName：要调用的函数/子程序名（通常为 "CATMain"）
          - iParameters：传递给宏的参数列表

        为兼容不同语言的 CATIA 安装，依次尝试 "模块1"（中文）和 "Module1"（英文/法语），
        任一成功即返回；两者均失败时抛出最后一次的异常。
        """
        last_exc: Exception | None = None
        for module_name in ("模块1", "Module1"):
            try:
                app.com_object.SystemService.ExecuteScript(
                    str(macro_path), 2, module_name, func_name, params
                )
                return
            except Exception as e:
                last_exc = e
        raise last_exc  # type: ignore[misc]

    def _run_template_macro(
        self,
        macro_path: Path,
        template_path: str,
    ) -> None:
        """通过 CATIA SystemService.ExecuteScript 运行指定的 CATScript 宏，
        并将模板文件路径作为参数传入，宏内可通过 iParameters 直接获取。
        """
        try:
            from catia_copilot.catia.connection import get_catia_v5_application as _catia
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
            reply = QMessageBox.question(
                self, "权限不足",
                f"无法直接复制文件（权限不足）。\n\n"
                f"目标路径：\n{dest_file}\n\n"
                f"是否通过 UAC 提权以管理员身份重试？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                accepted = self._run_copy_elevated([(src_file, dest_file)])
                if accepted:
                    if dest_file.exists():
                        QMessageBox.information(
                            self, "成功",
                            f"'{base_name}' 已成功复制到：\n{dest_file}",
                        )
                    else:
                        QMessageBox.warning(
                            self, "结果未知",
                            f"提权复制已执行，但无法确认文件是否成功写入。\n"
                            f"请手动确认：\n{dest_file}",
                        )
                else:
                    QMessageBox.information(self, "已取消", "用户取消了 UAC 提权，文件未复制。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生意外错误：\n{e}")

    def _crack(self) -> None:
        base_src_dir = resource_path(CRACK_DIR_PATH)
        if not base_src_dir.exists() or not base_src_dir.is_dir():
            QMessageBox.warning(
                self, "文件夹未找到",
                f"找不到 'crack' 文件夹：\n{base_src_dir.parent}",
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

        # 按版本推断专属子目录（如安装路径末尾为 B28 → crack/R28）
        version_subdir = self._detect_crack_version_subdir(catia_root)
        src_dir = base_src_dir
        if version_subdir:
            versioned_dir = base_src_dir / version_subdir
            if versioned_dir.is_dir():
                src_dir = versioned_dir
                logger.info(f"使用版本专属 crack 目录：{src_dir}")
            else:
                reply = QMessageBox.question(
                    self, "找不到版本专属目录",
                    f"未找到版本专属 crack 子目录：\n{versioned_dir}\n\n"
                    f"是否改用通用 crack 根目录中的文件？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.No:
                    return
                # src_dir 保持为 base_src_dir（通用目录）

        dest_dir = Path(catia_root) / "win_b64" / "code" / "bin"
        if not dest_dir.exists():
            QMessageBox.critical(
                self, "文件夹未找到",
                f"目标文件夹不存在：\n{dest_dir}\n\n请检查您的CATIA安装。",
            )
            return

        files = [f for f in src_dir.iterdir() if f.is_file()]
        if not files:
            QMessageBox.warning(
                self, "文件夹为空",
                f"crack 源目录中没有文件：\n{src_dir}",
            )
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
            reply = QMessageBox.question(
                self, "权限不足",
                f"无法直接复制文件（权限不足）。\n\n"
                f"目标路径：\n{dest_dir}\n\n"
                f"是否通过 UAC 提权以管理员身份重试？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                ops = [(f, dest_dir / f.name) for f in files]
                accepted = self._run_copy_elevated(ops)
                if accepted:
                    success_count = sum(1 for _, dst in ops if dst.exists())
                    if success_count == len(ops):
                        QMessageBox.information(
                            self, "成功",
                            f"已成功复制 {success_count} 个文件到：\n{dest_dir}\n\n"
                            + "\n".join(f.name for f in files),
                        )
                    else:
                        QMessageBox.warning(
                            self, "部分完成",
                            f"提权复制已执行，但仅确认 {success_count}/{len(ops)} 个文件写入成功。\n"
                            f"目标路径：\n{dest_dir}\n\n请手动确认复制结果。",
                        )
                else:
                    QMessageBox.information(self, "已取消", "用户取消了 UAC 提权，文件未复制。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生意外错误：\n{e}")
