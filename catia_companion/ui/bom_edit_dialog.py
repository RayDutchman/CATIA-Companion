"""
BOM 编辑对话框模块。

提供：
- BomEditDialog – 可编辑表格，用于完成 BOM 属性并通过 COM 写回 CATIA。
"""

import copy
import os
import logging
import subprocess
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel, QLineEdit,
    QPushButton, QTreeWidget, QTreeWidgetItem, QHeaderView, QAbstractItemView,
    QComboBox, QCheckBox, QGroupBox, QMessageBox, QApplication,
    QFileDialog, QProgressDialog, QRadioButton, QButtonGroup, QStyledItemDelegate,
    QMenu, QWidgetAction,
)
from PySide6.QtGui import QColor, QPen, QPainter, QPixmap
from PySide6.QtCore import Qt, QSettings

from catia_companion.constants import (
    PRESET_USER_REF_PROPERTIES,
    BOM_EDIT_COLUMN_ORDER,
    BOM_COLUMN_DISPLAY_NAMES,
    BOM_READONLY_COLUMNS,
    SOURCE_TO_DISPLAY,
    SOURCE_OPTIONS,
    PART_NUMBER_VALID_PATTERN,
    FILENAME_NOT_FOUND,
)
from catia_companion.catia.bom_collect import collect_bom_rows, flatten_bom_to_summary
from catia_companion.catia.bom_write import write_bom_to_catia
from catia_companion.utils import read_catia_thumbnail

logger = logging.getLogger(__name__)

# 自定义 UserRole 用于 QTreeWidgetItem：标记行为锁定（不可读/未找到）
_ITEM_LOCKED_ROLE: int = Qt.ItemDataRole.UserRole + 1


def _is_catia_com_error(exc: Exception) -> bool:
    """Return True if *exc* is a ``pywintypes.com_error`` from the CATIA COM layer.

    This distinguishes deliberate user-cancel signals (CATIA raises a COM error
    when the user clicks Cancel or No in its own SaveAs dialog) from genuine
    OS-level failures such as disk-full or permission-denied, which are plain
    Python exceptions and must always be reported to the user.
    """
    try:
        import pywintypes  # noqa: PLC0415
        return isinstance(exc, pywintypes.com_error)
    except ImportError:
        return False


def _find_catia_doc_by_path(docs, path: Path) -> object | None:
    """返回解析路径与 *path* 匹配的 CATIA 文档对象，如果未找到则返回 ``None``。

    参数：
        docs: CATIA 文档集合
        path: 要匹配的文件路径

    返回：
        匹配的 CATIA 文档对象，或 None
    """
    for i in range(1, docs.count + 1):
        try:
            d = docs.item(i)
            if Path(d.full_name).resolve() == path:
                return d
        except Exception:
            pass
    return None


def _find_catia_doc_by_part_number(docs, pn: str) -> object | None:
    """返回根产品零件编号与 *pn* 匹配的第一个已打开 CATIA 文档对象。

    当零件尚未保存到磁盘时（无文件路径），可通过零件编号定位已在 CATIA
    中打开的文档。如果零件编号不可用则回退到按文档名（不含扩展名）匹配。
    未找到时返回 ``None``。

    参数：
        docs: CATIA 文档集合
        pn:   要匹配的零件编号

    返回：
        匹配的 CATIA 文档对象，或 None
    """
    for i in range(1, docs.count + 1):
        try:
            d = docs.item(i)
            try:
                doc_pn = d.com_object.Product.PartNumber
            except Exception:
                # 非零件/产品类文档（如工程图）无 .Product；回退到文档名茎
                doc_pn = Path(d.name).stem
            if doc_pn == pn:
                return d
        except Exception:
            pass
    return None



class _BomTreeDelegate(QStyledItemDelegate):
    """Per-column read-only enforcement for the BOM QTreeWidget.

    QTreeWidgetItem flags are row-wide; this delegate returns ``None`` from
    :meth:`createEditor` for any column whose internal name belongs to
    :data:`~catia_companion.constants.BOM_READONLY_COLUMNS`, and also for
    any row that has been marked locked (file not found / unreadable).
    """

    def __init__(self, cols_fn, tree: QTreeWidget) -> None:
        super().__init__(tree)
        self._cols_fn = cols_fn  # callable: () -> list[str]

    def createEditor(self, parent, option, index):
        tree = self.parent()
        item = tree.itemFromIndex(index)
        if item is not None and item.data(0, _ITEM_LOCKED_ROLE):
            return None
        col_name = self._cols_fn()[index.column()]
        if col_name in BOM_READONLY_COLUMNS:
            return None
        return super().createEditor(parent, option, index)


class _BomTreeWidget(QTreeWidget):
    """QTreeWidget that draws Windows-Regedit-style dotted connector lines.

    Qt's default Windows/Fusion styles omit the vertical guide lines that
    connect parent and child nodes.  This subclass overrides
    :meth:`drawBranches` to paint 1-pixel-on / 1-pixel-off dotted lines
    (keyed on absolute viewport coordinates so vertical guides remain
    phase-consistent across consecutive rows).
    """

    _LINE_COLOR = QColor("#a0aab4")

    def drawBranches(self, painter: QPainter, rect, index) -> None:
        # 首先调用父类的 drawBranches，让 Qt 绘制默认的展开/折叠箭头指示器。
        super().drawBranches(painter, rect, index)

        indent = self.indentation()  # 获取每一层级的缩进宽度（像素）。
        model  = self.model()        # 获取当前树控件关联的数据模型。

        # 从当前节点向上遍历到根节点，依次记录每一层的祖先节点是否还有下一个兄弟节点
        # （即：在同一层级中，该节点下方是否还有其他节点）。
        has_next: list[bool] = []  # 存储各层级"是否有下一个兄弟"的布尔值列表。
        tmp = index                # 从当前节点的索引开始向上遍历。
        while True:
            par = tmp.parent()  # 获取当前节点的父节点索引。
            # 如果父节点有效（即不是根节点），则获取父节点下的子节点总数；
            # 否则（当前节点本身就是顶层节点）获取顶层节点总数。
            cnt = model.rowCount(par) if par.isValid() else model.rowCount()
            # 如果当前节点的行号小于兄弟节点总数减一，说明它后面还有兄弟节点。
            has_next.append(tmp.row() < cnt - 1)
            if not par.isValid():  # 已到达顶层节点，停止向上遍历。
                break
            tmp = par  # 继续向上，处理父节点。
        # 翻转列表：使 has_next[0] 对应最顶层祖先，has_next[-1] 对应当前节点自身。
        has_next.reverse()

        depth = len(has_next) - 1  # 当前节点的深度：顶层节点为 0，其子节点为 1，以此类推。

        # 顶层节点（depth == 0）不需要绘制任何连接线，直接返回。
        if depth == 0:
            return

        # 计算当前行在垂直方向上的中点 y 坐标，用于绘制水平横线和连接角。
        mid_y = (rect.top() + rect.bottom()) // 2

        pen = QPen(self._LINE_COLOR, 1, Qt.PenStyle.SolidLine)  # 创建 1 像素宽的实线画笔，颜色为类属性中定义的连接线颜色。
        pen.setDashPattern([1.0, 1.0])  # 将画笔设为点状虚线：1 像素绘制、1 像素间隔交替。
        # 根据当前行顶部的绝对 y 坐标对虚线相位进行对齐，
        # 确保相邻行之间的竖向连接线点阵在视觉上连续、不错位。
        pen.setDashOffset(rect.top() % 2)

        painter.save()  # 保存当前画笔状态，避免影响其他控件的绘制。
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, False)  # 关闭抗锯齿，保持像素级对齐，使点状虚线清晰。
        painter.setPen(pen)  # 应用上面配置好的点状虚线画笔。

        # 遍历当前节点的所有祖先层（除最近一级直接父层外），
        # 如果该层的祖先节点下方还有兄弟节点（has_next[d] 为 True），
        # 则在该层对应的 x 列绘制一条贯穿整行高度的竖线（表示该分支尚未结束）。
        for d in range(depth - 1):
            if has_next[d + 1]:  # 该祖先层仍有后续兄弟节点，需要绘制连续竖线。
                x = rect.left() + d * indent + indent // 2  # 计算该祖先层连接线的 x 坐标（列中心）。
                painter.drawLine(x, rect.top(), x, rect.bottom())  # 绘制贯通整行的竖线。

        # 处理当前节点的直接父层（最近一级）连接线，分两种情况：
        #   T 型连接符（├─）：当前节点后面还有兄弟节点 → 绘制贯通整行的竖线 + 水平横线。
        #   L 型连接符（└─）：当前节点是最后一个子节点 → 仅绘制上半段竖线（转角）+ 水平横线。
        x     = rect.left() + (depth - 1) * indent + indent // 2  # 直接父层连接线的 x 坐标（列中心）。
        x_end = rect.left() + depth * indent                       # 水平横线的终点 x 坐标（当前节点内容列的左边缘）。
        if has_next[-1]:  # T 型：当前节点后面还有兄弟节点。
            painter.drawLine(x, rect.top(), x, rect.bottom())  # 绘制贯通整行高度的竖线（T 型竖边）。
        else:             # L 型：当前节点是最后一个子节点。
            painter.drawLine(x, rect.top(), x, mid_y)          # 仅绘制从行顶到行中点的上半段竖线（L 型转角）。
        painter.drawLine(x, mid_y, x_end, mid_y)               # 绘制从竖线底部延伸到内容区左边缘的水平横线。

        painter.restore()  # 恢复之前保存的画笔状态，避免影响后续其他控件的绘制。


class _FileRenameDialog(QDialog):
    """Dialog for renaming or moving a single CATIA file via CATIA SaveAs.

    Lets the user change the file stem (name without extension) and/or the
    target directory independently.  Validates the new stem with
    :data:`~catia_companion.constants.PART_NUMBER_VALID_PATTERN` and creates
    the target directory on demand.
    """

    def __init__(self, current_fp: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("编辑文件名/路径")
        self.setMinimumWidth(540)
        self._current_fp = current_fp
        self._p          = Path(current_fp)

        layout = QFormLayout(self)
        layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        layout.setSpacing(8)

        # Current path (read-only)
        cur_label = QLabel(current_fp)
        cur_label.setWordWrap(True)
        cur_label.setStyleSheet("color: #555;")
        layout.addRow("当前路径：", cur_label)

        # New filename (stem only; extension is preserved automatically)
        self._name_edit = QLineEdit(self._p.stem)
        layout.addRow(f"新文件名（不含扩展名 {self._p.suffix}）：", self._name_edit)

        # New directory (with browse button)
        dir_widget = QWidget()
        dir_layout = QHBoxLayout(dir_widget)
        dir_layout.setContentsMargins(0, 0, 0, 0)
        self._dir_edit = QLineEdit(str(self._p.parent))
        dir_btn        = QPushButton("浏览…")
        dir_btn.setFixedWidth(64)
        dir_btn.clicked.connect(self._browse_dir)
        dir_layout.addWidget(self._dir_edit)
        dir_layout.addWidget(dir_btn)
        layout.addRow("新目录：", dir_widget)

        # Path preview
        self._preview_label = QLabel()
        self._preview_label.setWordWrap(True)
        self._preview_label.setStyleSheet("color: #333; font-style: italic;")
        layout.addRow("新路径预览：", self._preview_label)

        self._name_edit.textChanged.connect(self._update_preview)
        self._dir_edit.textChanged.connect(self._update_preview)
        self._update_preview()

        # Buttons
        btn_widget = QWidget()
        btn_row    = QHBoxLayout(btn_widget)
        btn_row.setContentsMargins(0, 0, 0, 0)
        btn_row.addStretch()
        ok_btn     = QPushButton("确认")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self._validate_and_accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addRow(btn_widget)

    # ── Properties ─────────────────────────────────────────────────────────

    @property
    def new_stem(self) -> str:
        return self._name_edit.text().strip()

    @property
    def new_dir(self) -> str:
        return self._dir_edit.text().strip()

    @property
    def new_path(self) -> str:
        stem      = self.new_stem or self._p.stem
        directory = self.new_dir  or str(self._p.parent)
        return str(Path(directory) / (stem + self._p.suffix))

    # ── Slots ───────────────────────────────────────────────────────────────

    def _update_preview(self) -> None:
        self._preview_label.setText(self.new_path)

    def _browse_dir(self) -> None:
        d = QFileDialog.getExistingDirectory(
            self, "选择目标目录",
            self._dir_edit.text() or str(self._p.parent),
        )
        if d:
            self._dir_edit.setText(d)

    def _validate_and_accept(self) -> None:
        stem = self.new_stem or self._p.stem
        if stem != self._p.stem and not PART_NUMBER_VALID_PATTERN.fullmatch(stem):
            QMessageBox.warning(
                self, "文件名含非法字符",
                f"文件名 「{stem}」 含有非法字符。\n"
                "不允许：控制字符、非ASCII字符，以及Windows文件名禁用字符"
                "（\\ / : * ? \" < > |）。",
            )
            return
        new_p = Path(self.new_path)
        if new_p.resolve() == self._p.resolve():
            QMessageBox.warning(self, "路径未改变", "新路径与当前路径相同，无需操作。")
            return
        dest_dir = new_p.parent
        if not dest_dir.exists():
            ret = QMessageBox.question(
                self, "目录不存在",
                f"目标目录不存在：\n{dest_dir}\n\n是否创建该目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                QMessageBox.critical(self, "创建目录失败", f"无法创建目录：\n{exc}")
                return
        self.accept()


class BomEditDialog(QDialog):
    """Editable BOM table for completing and writing back product properties.

    - 文件名 / 层级 / 类型 / 数量 are read-only (structural).
    - 零件编号 is editable with duplicate-detection.
    - 源 (Source) uses a QComboBox (未知 / 自制 / 外购).
    - Rows sharing the same Part Number are linked and update together.
    - "应用" writes changes back without closing; "完成" writes and closes.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("BOM属性补全")
        self.setMinimumSize(900, 600)
        self.resize(1100, 700)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowMinimizeButtonHint
        )

        # ── Settings ─────────────────────────────────────────────────────────
        # Share ExportBomDialog's custom-column config
        self._export_settings = QSettings("CATIACompanion", "ExportBOMDialog")
        self._last_browse_dir = self._export_settings.value("last_browse_dir", "")

        saved_custom = self._export_settings.value("custom_columns", [])
        if isinstance(saved_custom, str):
            saved_custom = [saved_custom]
        self._custom_columns: list[str] = list(saved_custom)

        # BomEditDialog-specific settings
        self._edit_settings  = QSettings("CATIACompanion", "BomEditDialog")
        saved_visible        = self._edit_settings.value("visible_preset_columns", [])
        if isinstance(saved_visible, str):
            saved_visible = [saved_visible]
        self._visible_preset_cols: list[str] = [
            c for c in saved_visible if c in PRESET_USER_REF_PROPERTIES
        ]

        self._summarize: bool = self._edit_settings.value("summarize", False, type=bool)
        self._summary_include_assemblies: bool = self._edit_settings.value(
            "summary_include_assemblies", False, type=bool
        )
        self._summary_sort_column: str = self._edit_settings.value(
            "summary_sort_column", "Part Number"
        )

        # All custom columns (including all presets) so that pre-loading from
        # CATIA covers every column regardless of current visibility.
        self._all_custom_columns: list[str] = list(dict.fromkeys(
            self._custom_columns + list(PRESET_USER_REF_PROPERTIES)
        ))

        self._show_filepath_col: bool = self._edit_settings.value(
            "show_filepath_column", False, type=bool,
        )
        self._show_filename_col: bool = self._edit_settings.value(
            "show_filename_column", True, type=bool,
        )

        self._columns: list[str] = self._build_visible_columns()

        # ── State ─────────────────────────────────────────────────────────────
        # {original_pn: {col_name: value}}  (canonical data, Source as display label)
        self._canonical_data: dict[str, dict[str, str]] = {}
        # Snapshot at last load/apply for dirty-only write-back
        self._snapshot_data: dict[str, dict[str, str]] = {}
        # {original_pn: {col_name, ...}}  – fields changed since last write-back
        self._modified_keys: dict[str, set[str]] = {}
        # All BOM rows in traversal order
        self._rows: list[dict] = []
        # Guard against re-entrant change handling
        self._is_updating: bool = False
        # Parallel list: self._item_by_row[i] is the QTreeWidgetItem for self._rows[i]
        self._item_by_row: list[QTreeWidgetItem] = []
        # True once BOM has been successfully loaded at least once
        self._bom_loaded: bool = False
        # Raw (hierarchical) BOM rows as returned by collect_bom_rows(); used to
        # reconstruct the display rows when the user toggles the BOM type.
        self._raw_rows: list[dict] = []

        # ── Layout ────────────────────────────────────────────────────────────
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # Source selection
        self._use_active_chk = QCheckBox("使用当前CATIA活动文档（不选择文件）")
        self._use_active_chk.toggled.connect(self._toggle_file_row)
        layout.addWidget(self._use_active_chk)

        file_row = QHBoxLayout()
        self._file_edit       = QLineEdit()
        self._file_edit.setPlaceholderText("选择一个CATProduct文件...")
        self._file_edit.setReadOnly(True)
        self._file_browse_btn = QPushButton("浏览...")
        self._file_browse_btn.clicked.connect(self._browse_file)
        self._load_btn        = QPushButton("加载BOM")
        self._load_btn.clicked.connect(self._load_bom)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(self._file_browse_btn)
        file_row.addWidget(self._load_btn)
        layout.addLayout(file_row)

        # ── BOM type + display options (single compact group) ────────────────
        display_group  = QGroupBox("BOM类型与显示选项")
        display_layout = QVBoxLayout(display_group)
        display_layout.setSpacing(4)
        display_layout.setContentsMargins(8, 6, 8, 6)

        # Row 1: radio buttons + filepath checkbox on the same line
        bom_type_row = QHBoxLayout()
        self._bom_type_btn_group = QButtonGroup(self)
        self._radio_hierarchical = QRadioButton("层级BOM")
        self._radio_summary_bom  = QRadioButton("汇总BOM")
        if self._summarize:
            self._radio_summary_bom.setChecked(True)
        else:
            self._radio_hierarchical.setChecked(True)
        self._bom_type_btn_group.addButton(self._radio_hierarchical)
        self._bom_type_btn_group.addButton(self._radio_summary_bom)
        self._radio_summary_bom.toggled.connect(self._on_bom_type_changed)
        bom_type_row.addWidget(self._radio_hierarchical)
        bom_type_row.addWidget(self._radio_summary_bom)

        self._summary_opts_widget = QWidget()
        summary_opts_layout = QHBoxLayout(self._summary_opts_widget)
        summary_opts_layout.setContentsMargins(0, 0, 0, 0)
        summary_opts_layout.setSpacing(8)

        self._include_assemblies_chk = QCheckBox("包含产品和部件（子装配体）")
        self._include_assemblies_chk.setToolTip(
            "勾选后，汇总BOM中也会列出产品和部件（子装配体），而不仅限于零件。"
        )
        self._include_assemblies_chk.setChecked(self._summary_include_assemblies)
        self._include_assemblies_chk.toggled.connect(self._on_include_assemblies_toggled)
        summary_opts_layout.addWidget(self._include_assemblies_chk)
        summary_opts_layout.addSpacing(8)
        summary_opts_layout.addWidget(QLabel("排序列:"))
        self._sort_col_combo = QComboBox()
        _sort_cols = list(BOM_EDIT_COLUMN_ORDER) + [
            c for c in PRESET_USER_REF_PROPERTIES if c not in BOM_EDIT_COLUMN_ORDER
        ] + [
            c for c in self._custom_columns
            if c not in BOM_EDIT_COLUMN_ORDER and c not in PRESET_USER_REF_PROPERTIES
        ]
        for col in _sort_cols:
            self._sort_col_combo.addItem(BOM_COLUMN_DISPLAY_NAMES.get(col, col), col)
        sort_saved_idx = self._sort_col_combo.findData(self._summary_sort_column)
        if sort_saved_idx >= 0:
            self._sort_col_combo.setCurrentIndex(sort_saved_idx)
        self._sort_col_combo.currentIndexChanged.connect(self._on_sort_col_changed)
        summary_opts_layout.addWidget(self._sort_col_combo)

        self._summary_opts_widget.setVisible(self._summarize)
        bom_type_row.addWidget(self._summary_opts_widget)
        bom_type_row.addStretch()
        display_layout.addLayout(bom_type_row)

        layout.addWidget(display_group)

        hint = QLabel(
            "层级 / 类型 / 数量 为结构属性，不可编辑，"
            "零件编号可编辑但不能与其他行冲突，"
            "文件名/路径可编辑。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # Preset column visibility checkboxes
        preset_group  = QGroupBox("自定义属性列（勾选以显示）")
        preset_layout = QHBoxLayout(preset_group)
        preset_layout.setSpacing(12)
        self._preset_checkboxes: dict[str, QCheckBox] = {}
        # "Filename" is a built-in column but can be toggled like a preset
        fn_cb = QCheckBox(BOM_COLUMN_DISPLAY_NAMES.get("Filename", "Filename"))
        fn_cb.setChecked(self._show_filename_col)
        fn_cb.toggled.connect(self._on_preset_col_toggled)
        preset_layout.addWidget(fn_cb)
        self._preset_checkboxes["Filename"] = fn_cb
        # "显示完整路径" follows immediately after the Filename checkbox
        self._filepath_chk = QCheckBox("显示完整路径")
        self._filepath_chk.setToolTip("勾选后文件名列将显示文件完整路径（含目录），而非仅文件名")
        self._filepath_chk.setChecked(self._show_filepath_col)
        self._filepath_chk.toggled.connect(self._on_show_filepath_toggled)
        preset_layout.addWidget(self._filepath_chk)
        for col_name in PRESET_USER_REF_PROPERTIES:
            cb = QCheckBox(col_name)
            cb.setChecked(col_name in self._visible_preset_cols)
            cb.toggled.connect(self._on_preset_col_toggled)
            preset_layout.addWidget(cb)
            self._preset_checkboxes[col_name] = cb
        layout.addWidget(preset_group)

        # BOM tree widget (replaces QTableWidget; tree handles expand/collapse natively)
        self._table = _BomTreeWidget()
        self._table.setHeaderLabels(self._display_headers())
        hdr = self._table.header()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        hdr.setSectionsMovable(True)
        hdr.setFixedHeight(28)
        self._table.setUniformRowHeights(True)
        self._table.setRootIsDecorated(True)
        self._table.setSortingEnabled(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setIndentation(16)
        self._table.itemChanged.connect(self._on_item_changed)
        _delegate = _BomTreeDelegate(lambda: self._columns, self._table)
        self._table.setItemDelegate(_delegate)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._on_tree_context_menu)
        layout.addWidget(self._table, 1)

        # Bottom buttons
        btn_row = QHBoxLayout()

        autofit_btn = QPushButton("自适应列宽")
        autofit_btn.setToolTip("根据内容自动调整所有列的宽度")
        autofit_btn.clicked.connect(self._autofit_columns)
        btn_row.addWidget(autofit_btn)

        expand_btn = QPushButton("全部展开")
        expand_btn.setToolTip("展开结构树中的所有节点")
        expand_btn.clicked.connect(self._table.expandAll)
        btn_row.addWidget(expand_btn)

        collapse_btn = QPushButton("全部折叠")
        collapse_btn.setToolTip("折叠结构树中的所有节点")
        collapse_btn.clicked.connect(self._table.collapseAll)
        btn_row.addWidget(collapse_btn)

        self._rename_btn = QPushButton("按零件编号修改文件名")
        self._rename_btn.setEnabled(False)
        self._rename_btn.clicked.connect(self._rename_by_part_number)
        btn_row.addWidget(self._rename_btn)

        self._rename_file_btn = QPushButton("编辑选中文件名/路径")
        self._rename_file_btn.setToolTip("为选中行的文件执行重命名或移动（通过CATIA另存为）")
        self._rename_file_btn.setEnabled(False)
        self._rename_file_btn.clicked.connect(self._rename_selected_file)
        btn_row.addWidget(self._rename_file_btn)
        btn_row.addStretch()

        self._save_btn   = QPushButton("应用")
        self._save_btn.setEnabled(False)
        self._save_btn.clicked.connect(self._apply_changes)

        self._finish_btn = QPushButton("完成")
        self._finish_btn.setDefault(True)
        self._finish_btn.setEnabled(False)
        self._finish_btn.clicked.connect(self._finish_and_close)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)

        btn_row.addWidget(self._save_btn)
        btn_row.addWidget(self._finish_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    # ── Source toggle ─────────────────────────────────────────────────────────

    def _toggle_file_row(self, use_active: bool) -> None:
        self._file_edit.setEnabled(not use_active)
        self._file_browse_btn.setEnabled(not use_active)

    # ── BOM type toggle ───────────────────────────────────────────────────────

    def _on_bom_type_changed(self, summary_checked: bool) -> None:
        self._summarize = summary_checked
        self._edit_settings.setValue("summarize", summary_checked)
        self._summary_opts_widget.setVisible(summary_checked)
        # If BOM is already loaded, re-derive display rows from the raw rows and repopulate
        if self._raw_rows:
            self._rows = (
                flatten_bom_to_summary(
                    self._raw_rows,
                    include_assemblies=self._summary_include_assemblies,
                    sort_column=self._summary_sort_column or None,
                )
                if summary_checked else self._raw_rows
            )
            self._rebuild_columns_and_repopulate()

    def _on_include_assemblies_toggled(self, checked: bool) -> None:
        self._summary_include_assemblies = checked
        self._edit_settings.setValue("summary_include_assemblies", checked)
        # Rebuild summary display if BOM is loaded and summary mode is active
        if self._summarize and self._raw_rows:
            self._rows = flatten_bom_to_summary(
                self._raw_rows,
                include_assemblies=checked,
                sort_column=self._summary_sort_column or None,
            )
            # When assemblies are included show the Type column; otherwise hide it
            self._rebuild_columns_and_repopulate()

    def _on_sort_col_changed(self, _index: int) -> None:
        col = self._sort_col_combo.currentData()
        if col:
            self._summary_sort_column = col
            self._edit_settings.setValue("summary_sort_column", col)
            # Re-sort the currently displayed summary rows if applicable
            if self._summarize and self._raw_rows:
                self._rows = flatten_bom_to_summary(
                    self._raw_rows,
                    include_assemblies=self._summary_include_assemblies,
                    sort_column=col,
                )
                self._populate_table()

    # ── Table helpers ─────────────────────────────────────────────────────────

    def _autofit_columns(self) -> None:
        """Resize all columns to fit their content, with a minimum width."""
        # QTreeWidget has resizeColumnToContents(int) not resizeColumnsToContents()
        min_width = 60
        for col in range(self._table.columnCount()):
            self._table.resizeColumnToContents(col)
            if self._table.columnWidth(col) < min_width:
                self._table.setColumnWidth(col, min_width)

    def _rebuild_columns_and_repopulate(self) -> None:
        """Rebuild the visible column list, update headers, and repopulate if rows are loaded."""
        self._columns = self._build_visible_columns()
        self._table.setHeaderLabels(self._display_headers())
        if self._rows:
            self._populate_table()
            for col in range(self._table.columnCount()):
                self._table.resizeColumnToContents(col)

    # ── Preset column helpers ─────────────────────────────────────────────────

    def _display_headers(self) -> list[str]:
        """Return display header labels for the current column list.

        When "文件名列显示完整路径" is active the Filename column header is
        shown as "完整路径" so users can tell what they're looking at.
        """
        result = []
        for c in self._columns:
            if c == "Filename" and self._show_filepath_col:
                result.append("完整路径")
            else:
                result.append(BOM_COLUMN_DISPLAY_NAMES.get(c, c))
        return result

    def _build_visible_columns(self) -> list[str]:
        base = list(BOM_EDIT_COLUMN_ORDER)
        if not self._show_filename_col:
            base = [c for c in base if c != "Filename"]
        if self._summarize:
            # In summary mode Level has no meaning; also hide Type unless assemblies shown
            cols_to_hide = {"Level"}
            if not self._summary_include_assemblies:
                cols_to_hide.add("Type")
            base = [c for c in base if c not in cols_to_hide]
        visible_preset = [
            c for c in PRESET_USER_REF_PROPERTIES if c in self._visible_preset_cols
        ]
        other_custom   = [
            c for c in self._custom_columns
            if c not in BOM_EDIT_COLUMN_ORDER and c not in PRESET_USER_REF_PROPERTIES
        ]
        return base + visible_preset + other_custom

    def _on_preset_col_toggled(self) -> None:
        # "Filename" checkbox controls the built-in filename column visibility
        if "Filename" in self._preset_checkboxes:
            new_show_fn = self._preset_checkboxes["Filename"].isChecked()
            if new_show_fn != self._show_filename_col:
                self._show_filename_col = new_show_fn
                self._edit_settings.setValue("show_filename_column", self._show_filename_col)
        self._visible_preset_cols = [
            name for name, cb in self._preset_checkboxes.items()
            if name != "Filename" and cb.isChecked()
        ]
        self._edit_settings.setValue("visible_preset_columns", self._visible_preset_cols)
        self._rebuild_columns_and_repopulate()

    def _on_show_filepath_toggled(self, checked: bool) -> None:
        self._show_filepath_col = checked
        self._edit_settings.setValue("show_filepath_column", checked)
        self._rebuild_columns_and_repopulate()

    # ── File picker ───────────────────────────────────────────────────────────

    def _browse_file(self) -> None:
        file, _ = QFileDialog.getOpenFileName(
            self, "选择CATProduct文件",
            self._last_browse_dir,
            "*.CATProduct (*.CATProduct);;All Files (*)",
        )
        if file:
            self._file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._export_settings.setValue("last_browse_dir", self._last_browse_dir)

    # ── Load BOM ──────────────────────────────────────────────────────────────

    def _load_bom(self) -> None:
        if self._use_active_chk.isChecked():
            file_path = None
        else:
            file_path = self._file_edit.text().strip()
            if not file_path:
                QMessageBox.warning(self, "未选择文件", "请先选择一个CATProduct文件。")
                return
            if not Path(file_path).exists():
                QMessageBox.warning(self, "文件不存在", f"文件不存在：\n{file_path}")
                return

        self._load_btn.setEnabled(False)
        self._load_btn.setText("加载中…")
        QApplication.processEvents()

        progress = QProgressDialog("正在加载BOM，请稍候…", None, 0, 0, self)
        progress.setWindowTitle("加载BOM")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(300)
        progress.setValue(0)

        def _on_row_collected(count: int) -> None:
            progress.setLabelText(f"正在加载BOM，请稍候… 已读取 {count} 个节点")
            progress.repaint()
            QApplication.processEvents()

        try:
            all_read_cols = list(dict.fromkeys(
                BOM_EDIT_COLUMN_ORDER
                + [c for c in self._all_custom_columns if c not in BOM_EDIT_COLUMN_ORDER]
            ))
            rows = collect_bom_rows(
                file_path, all_read_cols, self._all_custom_columns,
                progress_callback=_on_row_collected,
            )
        except Exception as e:
            progress.close()
            logger.error(f"Failed to load BOM for edit: {e}")
            QMessageBox.critical(
                self, "加载失败",
                f"加载BOM时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            self._load_btn.setEnabled(True)
            self._load_btn.setText("加载BOM")
            return
        finally:
            progress.close()

        self._load_btn.setEnabled(True)
        self._load_btn.setText("重新加载BOM")

        # Always save the raw hierarchical rows so we can switch modes later
        self._raw_rows = rows

        # In summary mode collapse the hierarchy into unique parts
        display_rows = (
            flatten_bom_to_summary(
                rows,
                include_assemblies=self._summary_include_assemblies,
                sort_column=self._summary_sort_column or None,
            )
            if self._summarize else rows
        )

        self._rows = display_rows

        # Build PN-keyed canonical data from the raw rows (first occurrence wins).
        # Using raw rows ensures all parts are indexed regardless of current mode.
        all_data_cols = list(dict.fromkeys(
            BOM_EDIT_COLUMN_ORDER
            + [c for c in self._all_custom_columns if c not in BOM_EDIT_COLUMN_ORDER]
        ))
        self._canonical_data = {}
        for row in rows:
            pn = str(row.get("Part Number", ""))
            if pn and pn not in self._canonical_data:
                data: dict[str, str] = {}
                for col in all_data_cols:
                    val = str(row.get(col, ""))
                    if col == "Source":
                        val = SOURCE_TO_DISPLAY.get(val, val)
                    data[col] = val
                self._canonical_data[pn] = data

        self._snapshot_data  = copy.deepcopy(self._canonical_data)
        self._modified_keys.clear()

        saved_widths = (
            [self._table.columnWidth(i) for i in range(self._table.columnCount())]
            if self._bom_loaded else []
        )

        self._populate_table()
        if not self._bom_loaded:
            for _c in range(self._table.columnCount()):
                self._table.resizeColumnToContents(_c)
            self._bom_loaded = True
        else:
            for i, w in enumerate(saved_widths):
                if i < self._table.columnCount():
                    self._table.setColumnWidth(i, w)

        self._save_btn.setEnabled(True)
        self._finish_btn.setEnabled(True)
        self._rename_btn.setEnabled(True)
        self._rename_file_btn.setEnabled(True)

    def _populate_table(self) -> None:
        self._is_updating = True
        self._table.blockSignals(True)

        self._table.clear()                          # removes all items; headers persist
        self._table.setHeaderLabels(self._display_headers())
        self._item_by_row = []

        # parent_stack: list of (level, item_or_None)
        # The sentinel at position 0 represents the invisible root (level −1).
        parent_stack: list[tuple[int, QTreeWidgetItem | None]] = [(-1, None)]

        for row_idx, row_data in enumerate(self._rows):
            level = 0 if self._summarize else int(row_data.get("Level", 0))

            # Pop until the top of the stack has a level strictly below ours
            while len(parent_stack) > 1 and parent_stack[-1][0] >= level:
                parent_stack.pop()

            parent_item = parent_stack[-1][1]
            item = QTreeWidgetItem()
            # Store row_idx in UserRole of column 0 for reverse lookup
            item.setData(0, Qt.ItemDataRole.UserRole, row_idx)

            if parent_item is None:
                self._table.addTopLevelItem(item)
            else:
                parent_item.addChild(item)

            parent_stack.append((level, item))
            self._item_by_row.append(item)

            pn         = str(row_data.get("Part Number", ""))
            not_found  = bool(row_data.get("_not_found"))
            unreadable = bool(row_data.get("_unreadable"))
            row_locked = unreadable or not_found

            for col_idx, col_name in enumerate(self._columns):

                # Source → QComboBox (overlay widget; not stored as item text)
                if col_name == "Source":
                    raw    = str(row_data.get("Source", ""))
                    pn_val = self._canonical_data.get(pn, {}).get(
                        "Source", SOURCE_TO_DISPLAY.get(raw, raw)
                    )
                    if pn_val not in SOURCE_OPTIONS:
                        pn_val = SOURCE_TO_DISPLAY.get(pn_val, SOURCE_OPTIONS[0])
                    combo = QComboBox()
                    combo.blockSignals(True)
                    combo.addItems(SOURCE_OPTIONS)
                    combo.setCurrentText(pn_val)
                    combo.blockSignals(False)
                    if row_locked:
                        combo.setEnabled(False)
                    else:
                        combo.currentTextChanged.connect(
                            lambda text, r=row_idx: self._on_source_changed(r, text)
                        )
                    self._table.setItemWidget(item, col_idx, combo)
                    continue

                # All other columns → item text
                if col_name == "Quantity":
                    value = str(row_data.get("Quantity", "1"))
                elif col_name == "Filename":
                    fp = str(row_data.get("_filepath", ""))
                    fn = str(row_data.get("Filename", ""))
                    value = (fp if fp else fn) if self._show_filepath_col else fn
                elif col_name == "Filepath":
                    value = str(row_data.get("_filepath", ""))
                elif col_name in BOM_READONLY_COLUMNS:
                    value = str(row_data.get(col_name, ""))
                else:
                    value = str(
                        self._canonical_data.get(pn, {}).get(
                            col_name, row_data.get(col_name, "")
                        )
                    )
                item.setText(col_idx, value)

                if col_name == "Filename":
                    fp = str(row_data.get("_filepath", ""))
                    fn = str(row_data.get("Filename", ""))
                    if fp:
                        if self._show_filepath_col:
                            if fn and fn != FILENAME_NOT_FOUND:
                                item.setToolTip(col_idx, fn)
                        else:
                            item.setToolTip(col_idx, fp)

            # Non-locked rows: allow in-place editing (delegate blocks read-only columns)
            if not row_locked:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                item.setData(0, _ITEM_LOCKED_ROLE, False)
            else:
                grey = QColor(160, 160, 160)
                bg   = QColor(250, 245, 245) if not_found else QColor(245, 245, 245)
                item.setData(0, _ITEM_LOCKED_ROLE, True)
                for ci in range(len(self._columns)):
                    item.setForeground(ci, grey)
                    item.setBackground(ci, bg)
                fn_col = self._columns.index("Filename") if "Filename" in self._columns else -1
                if fn_col >= 0:
                    tip = (
                        "该零件/装配体的文件未被CATIA检索到，行内容不可编辑。"
                        if not_found else
                        "该零件/装配体处于轻量化模式，无法读取属性。"
                    )
                    item.setToolTip(fn_col, tip)

        self._table.expandAll()
        self._table.blockSignals(False)
        self._is_updating = False

    # ── Tree helpers ──────────────────────────────────────────────────────────

    def _iter_all_items(self):
        """Yield every QTreeWidgetItem in DFS (pre-order) traversal."""
        def _walk(parent: QTreeWidgetItem):
            yield parent
            for i in range(parent.childCount()):
                yield from _walk(parent.child(i))
        for i in range(self._table.topLevelItemCount()):
            yield from _walk(self._table.topLevelItem(i))

    # ── Source combo change ───────────────────────────────────────────────────

    def _on_source_changed(self, row_idx: int, text: str) -> None:
        if self._is_updating:
            return
        if "Source" not in self._columns:
            return
        src_col_idx = self._columns.index("Source")

        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        direct_rows = selected_row_indices if row_idx in selected_row_indices else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            pn = str(self._rows[r].get("Part Number", ""))
            if pn:
                pns_to_update.add(pn)

        for pn in pns_to_update:
            if pn in self._canonical_data:
                self._canonical_data[pn]["Source"] = text
                self._modified_keys.setdefault(pn, set()).add("Source")

        self._is_updating = True
        for other_item in self._iter_all_items():
            other_row_idx = other_item.data(0, Qt.ItemDataRole.UserRole)
            if other_row_idx is None or other_row_idx == row_idx:
                continue
            other_pn = str(self._rows[other_row_idx].get("Part Number", ""))
            if other_pn in pns_to_update:
                combo = self._table.itemWidget(other_item, src_col_idx)
                if isinstance(combo, QComboBox) and combo.currentText() != text:
                    combo.blockSignals(True)
                    combo.setCurrentText(text)
                    combo.blockSignals(False)
        self._is_updating = False

    # ── Regular cell edit ─────────────────────────────────────────────────────

    def _on_item_changed(self, item: QTreeWidgetItem, col_idx: int) -> None:
        if self._is_updating:
            return
        row_idx = item.data(0, Qt.ItemDataRole.UserRole)
        if row_idx is None:
            return
        col_name = self._columns[col_idx]

        if col_name in BOM_READONLY_COLUMNS or col_name == "Source":
            return

        new_value = item.text(col_idx)
        pn        = str(self._rows[row_idx].get("Part Number", ""))

        if col_name == "Part Number":
            # ── Empty / whitespace-only PN ────────────────────────────────────
            if not new_value.strip():
                QMessageBox.warning(
                    self, "零件编号不能为空",
                    "零件编号不能为空或仅含空格，请输入有效的零件编号。",
                )
                self._is_updating = True
                item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                self._is_updating = False
                return

            # ── Strip leading/trailing whitespace silently ────────────────────
            if new_value != new_value.strip():
                new_value = new_value.strip()
                self._is_updating = True
                item.setText(col_idx, new_value)
                self._is_updating = False

            # ── Character validity ────────────────────────────────────────────
            if not PART_NUMBER_VALID_PATTERN.fullmatch(new_value):
                QMessageBox.warning(
                    self, "零件编号含非法字符",
                    f"零件编号 \"{new_value}\" 含有非法字符。\n"
                    "不允许：控制字符、非ASCII字符，以及Windows文件名禁用字符"
                    "（\\ / : * ? \" < > |）。",
                )
                self._is_updating = True
                item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                self._is_updating = False
                return

            # ── Conflict with current canonical values ────────────────────────
            for other_pn, data in self._canonical_data.items():
                if other_pn == pn:
                    continue
                if data.get("Part Number", other_pn) == new_value:
                    QMessageBox.warning(
                        self, "零件编号冲突",
                        f"零件编号 \"{new_value}\" 与 \"{other_pn}\" "
                        f"的当前零件编号冲突，不允许修改。",
                    )
                    self._is_updating = True
                    item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                    self._is_updating = False
                    return

            # ── Conflict with snapshot (what CATIA currently holds) ───────────
            for other_pn, data in self._snapshot_data.items():
                if other_pn == pn:
                    continue
                if data.get("Part Number", other_pn) == new_value:
                    QMessageBox.warning(
                        self, "零件编号冲突",
                        f"零件编号 \"{new_value}\" 与 \"{other_pn}\" "
                        f"的原始零件编号冲突，不允许修改。",
                    )
                    self._is_updating = True
                    item.setText(col_idx, self._canonical_data.get(pn, {}).get("Part Number", pn))
                    self._is_updating = False
                    return

        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        direct_rows = selected_row_indices if row_idx in selected_row_indices else {row_idx}

        pns_to_update: set[str] = set()
        for r in direct_rows:
            r_pn = str(self._rows[r].get("Part Number", ""))
            if r_pn:
                pns_to_update.add(r_pn)
                if r_pn in self._canonical_data:
                    self._canonical_data[r_pn][col_name] = new_value
                    self._modified_keys.setdefault(r_pn, set()).add(col_name)

        self._is_updating = True
        for other_item in self._iter_all_items():
            if other_item is item:
                continue
            other_row_idx = other_item.data(0, Qt.ItemDataRole.UserRole)
            if other_row_idx is None:
                continue
            other_pn = str(self._rows[other_row_idx].get("Part Number", ""))
            if other_pn in pns_to_update:
                if other_item.text(col_idx) != new_value:
                    other_item.setText(col_idx, new_value)
        self._is_updating = False

    # ── Write-back ────────────────────────────────────────────────────────────

    def _rename_by_part_number(self) -> None:
        """SaveAs each CATIA file using its Part Number as the filename."""
        if self._modified_keys:
            ret = QMessageBox.question(
                self, "存在未回传的修改",
                "检测到BOM属性尚未写回CATIA。\n\n"
                "必须先将修改写回CATIA，才能确保零件编号与CATIA文件一致。\n\n"
                "是否立即执行写回？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
            self._write_back(close_on_success=False)
            # If write-back failed or was only partial, modified_keys is still
            # non-empty – stop here so the user can fix the issue first.
            if self._modified_keys:
                return
            # Write-back cleared all modifications; fall through to rename.

        to_rename: list[tuple[str, str]] = []
        seen_fps:  set[str] = set()
        for row in self._rows:
            fp = str(row.get("_filepath", ""))
            if not fp or fp in seen_fps:
                continue
            seen_fps.add(fp)
            orig_pn = str(row.get("Part Number", ""))
            pn      = str(self._canonical_data.get(orig_pn, {}).get("Part Number", orig_pn))
            if pn and Path(fp).stem != pn:
                to_rename.append((fp, pn))

        if not to_rename:
            QMessageBox.information(self, "无需改名", "所有文件名已与零件编号一致。")
            return

        delete_old = (
            QMessageBox.question(
                self, "是否删除旧文件",
                "另存为完成后，是否删除旧文件？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            ) == QMessageBox.StandardButton.Yes
        )

        QMessageBox.information(self, "请在CATIA中继续操作", "准备就绪，请在CATIA中确认后续操作。")

        renamed_count = 0

        for fp, pn in reversed(to_rename):
            if not PART_NUMBER_VALID_PATTERN.fullmatch(pn):
                QMessageBox.warning(
                    self, "零件编号含非法字符",
                    f"零件编号 「{pn}」 含有非法字符。\n"
                    "不允许：控制字符、非ASCII字符，以及Windows文件名禁用字符"
                    "（\\ / : * ? \" < > |）。\n请在表格中修改此零件编号后重试。",
                )
                continue

            if not Path(fp).exists():
                continue

            ext    = Path(fp).suffix
            new_fp = str(Path(fp).parent / (pn + ext))
            target_existed_before = Path(new_fp).exists()

            try:
                from pycatia import catia as _pycatia
                caa         = _pycatia()
                application = caa.application
                application.visible = True
                documents   = application.documents
                src         = Path(fp).resolve()

                target_doc = _find_catia_doc_by_path(documents, src)
                if target_doc is None:
                    documents.open(str(src))
                    target_doc = _find_catia_doc_by_path(documents, src)

                if target_doc is None:
                    QMessageBox.warning(
                        self, "无法找到文档",
                        f"无法在CATIA中找到或打开文档：\n{fp}",
                    )
                    continue

                target_doc.com_object.SaveAs(new_fp)

                if delete_old and Path(fp).resolve() != Path(new_fp).resolve():
                    try:
                        os.remove(fp)
                    except Exception as del_err:
                        logger.warning(f"Failed to delete old file {fp}: {del_err}")

                for row in self._rows:
                    if str(row.get("_filepath", "")) == fp:
                        row["_filepath"] = new_fp
                        row["Filename"]  = pn
                for row in self._raw_rows:
                    if str(row.get("_filepath", "")) == fp:
                        row["_filepath"] = new_fp
                        row["Filename"]  = pn
                renamed_count += 1

            except Exception as e:
                # Only treat the exception as a user-initiated cancel when:
                #   1. The exception came from the CATIA COM layer (pywintypes.com_error),
                #      which is what CATIA raises when the user clicks Cancel or No in
                #      its own SaveAs dialog – NOT for OS-level failures.
                #   2. The source file is still intact.
                #   3. The target file was either pre-existing or was never created.
                # Any other exception (OSError, PermissionError, non-COM errors like
                # disk full) must be surfaced to the user as a real failure.
                if _is_catia_com_error(e) and Path(fp).exists() and (
                    target_existed_before or not Path(new_fp).exists()
                ):
                    logger.info(
                        f"SaveAs skipped for {Path(fp).name} "
                        "(user cancelled or declined overwrite in CATIA)"
                    )
                    continue
                QMessageBox.warning(
                    self, "另存为失败", f"文件「{Path(fp).name}」另存为失败：\n{e}"
                )

        if renamed_count > 0:
            QMessageBox.information(
                self, "改名完成",
                f"已成功将 {renamed_count} 个文件通过CATIA另存为功能改名。",
            )
            self._populate_table()

    def _rename_selected_file(self) -> None:
        """Rename or move the file for a single selected BOM row via CATIA SaveAs."""
        selected_row_indices = {
            it.data(0, Qt.ItemDataRole.UserRole)
            for it in self._table.selectedItems()
            if it.data(0, Qt.ItemDataRole.UserRole) is not None
        }
        if len(selected_row_indices) != 1:
            QMessageBox.warning(
                self, "请选择单行",
                "请在表格中选中恰好一行，再执行此操作。",
            )
            return

        row_idx  = next(iter(selected_row_indices))
        row_data = self._rows[row_idx]
        fp       = str(row_data.get("_filepath", ""))

        if not fp or row_data.get("_not_found"):
            QMessageBox.warning(self, "无有效路径", "该行没有可用的文件路径，无法执行重命名/移动。")
            return
        if not Path(fp).exists():
            QMessageBox.warning(self, "文件不存在", f"文件不存在：\n{fp}")
            return

        # Require attribute write-back before renaming to keep file content consistent.
        orig_pn = str(row_data.get("Part Number", ""))
        if orig_pn in self._modified_keys:
            ret = QMessageBox.question(
                self, "存在未写回的属性修改",
                f"零件「{orig_pn}」的属性尚未写回CATIA。\n\n"
                "必须先将修改写回CATIA，才能确保文件内容与表格一致。\n\n"
                "是否立即执行写回？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
            self._write_back(close_on_success=False)
            # Only proceed if write-back actually cleared the modifications;
            # if it failed (error dialog shown), modified_keys still has the entry.
            if orig_pn in self._modified_keys:
                return

        dlg = _FileRenameDialog(fp, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_fp                = dlg.new_path
        target_existed_before = Path(new_fp).exists()

        delete_old = (
            QMessageBox.question(
                self, "是否删除旧文件",
                f"另存为完成后，是否删除旧文件？\n\n旧文件：{fp}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            ) == QMessageBox.StandardButton.Yes
        )

        try:
            from pycatia import catia as _pycatia
            caa         = _pycatia()
            application = caa.application
            application.visible = True
            documents   = application.documents
            src         = Path(fp).resolve()

            target_doc = _find_catia_doc_by_path(documents, src)
            if target_doc is None:
                documents.open(str(src))
                target_doc = _find_catia_doc_by_path(documents, src)

            if target_doc is None:
                QMessageBox.warning(
                    self, "无法找到文档",
                    f"无法在CATIA中找到或打开文档：\n{fp}",
                )
                return

            target_doc.com_object.SaveAs(new_fp)

            if delete_old and Path(fp).resolve() != Path(new_fp).resolve():
                try:
                    os.remove(fp)
                except Exception as del_err:
                    logger.warning(f"Failed to delete old file {fp}: {del_err}")

            new_stem = Path(new_fp).stem
            for row in self._rows:
                if str(row.get("_filepath", "")) == fp:
                    row["_filepath"] = new_fp
                    row["Filename"]  = new_stem
            for row in self._raw_rows:
                if str(row.get("_filepath", "")) == fp:
                    row["_filepath"] = new_fp
                    row["Filename"]  = new_stem
            self._populate_table()
            QMessageBox.information(
                self, "操作成功",
                f"文件已成功另存为：\n{new_fp}",
            )

        except Exception as e:
            # Only treat the exception as a user-initiated cancel when:
            #   1. The exception came from the CATIA COM layer (pywintypes.com_error),
            #      which is what CATIA raises when the user clicks Cancel or No in
            #      its own SaveAs dialog – NOT for OS-level failures.
            #   2. The source file is still intact.
            #   3. The target file was either pre-existing or was never created.
            # Any other exception (OSError, PermissionError, non-COM errors like
            # disk full) must be surfaced to the user as a real failure.
            if _is_catia_com_error(e) and Path(fp).exists() and (
                target_existed_before or not Path(new_fp).exists()
            ):
                # Most likely the user clicked Cancel or No in CATIA's own SaveAs
                # prompt – treat this as a deliberate skip with no error dialog.
                logger.info(
                    f"SaveAs skipped for {Path(fp).name} "
                    f"(user cancelled or declined overwrite in CATIA; exception: {e})"
                )
                return
            QMessageBox.warning(self, "另存为失败", f"文件操作失败：\n{e}")

    def _write_back(self, *, close_on_success: bool) -> None:
        """Write only the changed fields back to CATIA."""
        if self._use_active_chk.isChecked():
            file_path = None
        else:
            file_path = self._file_edit.text().strip()
            if not file_path:
                QMessageBox.warning(self, "未选择文件", "请选择一个CATProduct文件。")
                return

        # dirty_data must be keyed by the *current* CATIA PN, which may differ
        # from the internal canonical key (orig_pn) when a PN rename was already
        # written back in a previous write-back operation.  We keep pn_remap to
        # go back from current_pn → orig_pn for the post-write snapshot update.
        dirty_data: dict[str, dict[str, str]] = {}
        pn_remap:   dict[str, str]            = {}  # current_pn → orig_pn
        for pn, dirty_cols in self._modified_keys.items():
            if pn not in self._canonical_data:
                continue
            changed = {
                col: self._canonical_data[pn][col]
                for col in dirty_cols if col in self._canonical_data[pn]
            }
            if changed:
                # Use the snapshot PN (= what CATIA currently holds for this
                # node) as the lookup key for the traversal.  If the PN has
                # never been written back, the snapshot value equals orig_pn.
                current_pn = self._snapshot_data.get(pn, {}).get(
                    "Part Number", pn
                )
                dirty_data[current_pn] = changed
                pn_remap[current_pn]   = pn

        if not dirty_data:
            if close_on_success:
                self.accept()
            else:
                QMessageBox.information(self, "无更改", "没有检测到任何修改，无需写回。")
            return

        self._save_btn.setEnabled(False)
        self._finish_btn.setEnabled(False)
        QApplication.processEvents()

        progress = QProgressDialog("正在写回CATIA，请稍候…", None, 0, 0, self)
        progress.setWindowTitle("写回CATIA")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(300)
        progress.setValue(0)

        def _on_node_written(count: int) -> None:
            progress.setLabelText(f"正在写回CATIA，请稍候… 已处理 {count} 个节点")
            progress.repaint()
            QApplication.processEvents()

        try:
            write_bom_to_catia(file_path, dirty_data, self._all_custom_columns,
                               _on_node_written)
        except Exception as e:
            progress.close()
            logger.error(f"Failed to write BOM back to CATIA: {e}")
            self._save_btn.setEnabled(True)
            self._finish_btn.setEnabled(True)
            QMessageBox.critical(
                self, "写回失败",
                f"写回CATIA时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            return
        finally:
            progress.close()

        for current_pn, changed in dirty_data.items():
            pn = pn_remap.get(current_pn, current_pn)
            if pn in self._snapshot_data:
                self._snapshot_data[pn].update(changed)
            if pn in self._modified_keys:
                self._modified_keys[pn] -= set(changed.keys())
                if not self._modified_keys[pn]:
                    del self._modified_keys[pn]

        self._save_btn.setEnabled(True)
        self._finish_btn.setEnabled(True)

        if close_on_success:
            QMessageBox.information(
                self, "完成",
                "BOM属性已成功写回CATIA，请在CATIA中手动保存文件。",
            )
            self.accept()
        else:
            QMessageBox.information(
                self, "应用成功",
                "BOM属性已成功写回CATIA，请在CATIA中手动保存文件。",
            )

    def _apply_changes(self) -> None:
        """Write changes back to CATIA and keep the dialog open."""
        self._write_back(close_on_success=False)

    def _finish_and_close(self) -> None:
        """Write changes back to CATIA and close the dialog."""
        self._write_back(close_on_success=True)

    # ── Right-click context menu ──────────────────────────────────────────────

    def _on_tree_context_menu(self, pos) -> None:
        """Show a context menu for the right-clicked BOM row.

        If a thumbnail is embedded in the backing file it is shown at the top
        of the menu as a non-interactive image widget.
        """
        item = self._table.itemAt(pos)
        if item is None:
            return
        row_idx = item.data(0, Qt.ItemDataRole.UserRole)
        if row_idx is None:
            return

        row_data     = self._rows[row_idx]
        fp           = str(row_data.get("_filepath", ""))
        fp_path      = Path(fp) if fp else None
        is_component = row_data.get("Type") == "部件"
        not_found    = bool(row_data.get("_not_found"))
        pn           = str(row_data.get("Part Number", ""))

        # Ensure the right-clicked row is selected so that the downstream
        # helpers (_rename_selected_file etc.) can find it.
        if not item.isSelected():
            self._table.clearSelection()
            item.setSelected(True)

        menu = QMenu(self)

        # ── Embedded thumbnail (shown inline at the top when available) ───────
        # Conditions where we skip thumbnail extraction:
        #   • 部件: filepath belongs to the parent product, not this component
        #   • not_found: CATIA couldn't resolve the file
        #   • file doesn't exist on disk (unsaved or missing)
        if fp and not is_component and not not_found and fp_path is not None and fp_path.exists():
            img_bytes = read_catia_thumbnail(fp)
            if img_bytes:
                pixmap = QPixmap()
                loaded = pixmap.loadFromData(img_bytes)
                if loaded and not pixmap.isNull():
                    thumb_label = QLabel()
                    thumb_label.setPixmap(pixmap)
                    thumb_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    thumb_label.setContentsMargins(6, 4, 6, 4)
                    thumb_action = QWidgetAction(menu)
                    thumb_action.setDefaultWidget(thumb_label)
                    menu.addAction(thumb_action)
                    menu.addSeparator()

        # ── 打开路径 ──────────────────────────────────────────────────────────
        act_open_path = menu.addAction("打开路径")
        path_available = bool(fp) and fp_path is not None and (
            fp_path.exists() or fp_path.parent.exists()
        )
        act_open_path.setEnabled(path_available)

        # ── 在CATIA中打开 ─────────────────────────────────────────────────────
        # Locate by Part Number only.  Exclude:
        #   • 部件 – their Part Number identifies the parent product
        #   • not_found (断链接, light-red rows) – CATIA couldn't resolve the
        #     reference; such rows may not have a valid Part Number in CATIA's
        #     session and "opening" them is meaningless
        act_open_catia = menu.addAction("在CATIA中打开")
        act_open_catia.setEnabled(not is_component and not not_found and bool(pn))

        menu.addSeparator()

        # ── 编辑文件名/路径 ───────────────────────────────────────────────────
        act_edit_path = menu.addAction("编辑文件名/路径")
        act_edit_path.setEnabled(
            bool(fp) and not not_found
            and fp_path is not None and fp_path.exists()
        )

        action = menu.exec(self._table.viewport().mapToGlobal(pos))

        if action == act_open_path:
            self._open_path(fp)
        elif action == act_open_catia:
            self._open_in_catia(pn, fp)
        elif action == act_edit_path:
            self._rename_selected_file()

    def _open_path(self, fp: str) -> None:
        """Open the folder containing *fp* in Windows Explorer, highlighting the file."""
        p = Path(fp).resolve()
        try:
            if p.exists():
                # Quote the path so Explorer handles spaces in directory names.
                subprocess.Popen(f'explorer /select,"{p}"', shell=True)
            elif p.parent.exists():
                subprocess.Popen(f'explorer "{p.parent}"', shell=True)
        except Exception as exc:
            logger.warning(f"Failed to open path in Explorer: {exc}")

    def _open_in_catia(self, orig_pn: str, fp: str = "") -> None:
        """Activate the already-open CATIA document for the row identified by
        *orig_pn* (the Part Number as loaded from CATIA, used as the key into
        ``_snapshot_data``) and optionally *fp* (the backing file path).

        Lookup strategy (dual-track):
        1. If *fp* resolves to an existing file, try to match by full path first.
           This is immune to Part Number edits that have not yet been written back.
        2. Fall back to matching by Part Number, using the value from
           ``_snapshot_data`` – which is updated on every write-back and therefore
           always reflects the PN that CATIA currently holds, even after a rename.

        The caller must ensure that broken-link (not_found) rows are excluded
        from triggering this method.
        """
        try:
            from pycatia import catia as _pycatia
            caa         = _pycatia()
            application = caa.application
            application.visible = True
            documents   = application.documents

            doc: object | None = None

            # ── 1st track: match by file path (not affected by PN edits) ─────
            if fp:
                fp_path = Path(fp)
                if fp_path.exists():
                    doc = _find_catia_doc_by_path(documents, fp_path.resolve())

            # ── 2nd track: match by snapshot PN (reflects last write-back) ───
            snapshot_pn = self._snapshot_data.get(orig_pn, {}).get(
                "Part Number", orig_pn
            )
            if snapshot_pn != orig_pn:
                logger.debug(
                    "_open_in_catia: falling back to snapshot PN %r (orig %r)",
                    snapshot_pn, orig_pn,
                )
            if doc is None:
                doc = _find_catia_doc_by_part_number(documents, snapshot_pn)

            if doc is None:
                QMessageBox.warning(
                    self, "在CATIA中打开失败",
                    f"无法在CATIA中找到零件编号为 \"{snapshot_pn}\" 的文档。\n\n"
                    f"请确认该零件已在CATIA中打开。",
                )
                return

            doc.activate()
        except Exception as e:
            QMessageBox.warning(self, "在CATIA中打开失败", f"无法在CATIA中打开文件：\n{e}")
