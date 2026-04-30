"""
质量特性汇总对话框模块。

提供：
- MassPropsDialog – 遍历产品树，展示每个零件实例的质量/重心/转动惯量，
                    支持：
                      • 手动编辑重量（等比缩放惯量，联动同型号零件）
                      • 层级BOM / 汇总BOM 切换
                      • kg / g 单位切换（重量与转动惯量）
                      • 文件名 / 零件编号 / 术语 / 版本列可隐藏
                      • 计算装配体总质量特性并导出 Excel
"""

import logging
import math
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTreeWidgetItem, QHeaderView, QAbstractItemView,
    QCheckBox, QGroupBox, QMessageBox, QApplication,
    QFileDialog, QProgressDialog, QLineEdit, QGridLayout, QFrame,
    QRadioButton, QButtonGroup, QWidget, QComboBox,
    QStyledItemDelegate,
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QSettings

from catia_copilot.constants import (
    MASS_PROPS_COLUMNS,
    MASS_PROPS_COLUMN_DISPLAY_NAMES,
    MASS_PROPS_HIDEABLE_COLUMNS,
    MASS_PROPS_READONLY_COLUMNS,
    FILENAME_NOT_FOUND,
    FILENAME_UNSAVED,
)
from catia_copilot.catia.mass_props_collect import collect_mass_props_rows, _row_inertia_to_root, recompute_product_rows
from catia_copilot.catia.mass_props_calc import rollup_mass_properties
from catia_copilot.ui.bom_widgets import _BomTreeWidget

logger = logging.getLogger(__name__)

# UserRole for row index (maps to self._rows)
_ROW_IDX_ROLE = Qt.ItemDataRole.UserRole
# UserRole+1 for "locked" flag
_ITEM_LOCKED_ROLE = Qt.ItemDataRole.UserRole + 1

# Inertia column → (row, col) in the 3×3 tensor
_INERTIA_IDX: dict[str, tuple[int, int]] = {
    "Ixx": (0, 0), "Iyy": (1, 1), "Izz": (2, 2),
    "Ixy": (0, 1), "Ixz": (0, 2), "Iyz": (1, 2),
}

# Sortable columns in summary BOM mode
_SUMMARY_SORT_COLUMNS: list[str] = [
    "Part Number", "Nomenclature", "Revision", "Filename", "Weight",
    "CogX", "CogY", "CogZ",
]


class _MassPropsDelegate(QStyledItemDelegate):
    """只允许对未锁定零件行的"Weight"列进行编辑，其余列一律只读。"""

    def __init__(self, cols_fn, tree) -> None:
        super().__init__(tree)
        self._cols_fn = cols_fn  # callable: () -> list[str]

    def createEditor(self, parent, option, index):
        tree = self.parent()
        item = tree.itemFromIndex(index)
        if item is None:
            return None
        if item.data(0, _ITEM_LOCKED_ROLE):
            return None
        cols = self._cols_fn()
        if index.column() >= len(cols):
            return None
        col_name = cols[index.column()]
        if col_name in MASS_PROPS_READONLY_COLUMNS:
            return None
        return super().createEditor(parent, option, index)


def _fmt(value) -> str:
    """数值 → 字符串，None → '—'（不含单位换算）。"""
    if value is None:
        return "—"
    try:
        v = float(value)
        if math.isclose(v, round(v), rel_tol=0.0, abs_tol=1e-6):
            return f"{v:.0f}"
        if abs(v) >= 1e4 or (v != 0.0 and abs(v) < 0.001):
            return f"{v:.3e}"
        return f"{v:.3f}"
    except (TypeError, ValueError):
        return str(value)


class MassPropsDialog(QDialog):
    """质量特性汇总对话框。

    - 遍历 CATProduct 树，每个节点（零件/产品/部件实例）单独显示一行（层级BOM模式）。
      Weight / CogX / CogY / CogZ / Ixx–Iyz 均在根产品坐标系下显示，与装配位置有关。
    - 汇总BOM模式：相同零件编号的零件实例合并为一行，并显示数量（Quantity）；
      仅列出零件（不含产品和部件）；Weight / CogX / CogY / CogZ / Ixx–Iyz
      在零件自身坐标系下显示，与装配位置无关。
    - 仅零件节点的"重量"列可编辑；修改后等比缩放该行惯量，
      并同步更新所有相同零件编号的行（及 _rows 中全部同PN数据）。
    - 单位可在 kg/g 间切换（影响重量列与转动惯量列的显示和导出）。
    - "计算"按钮汇总装配体总质量特性（考虑位姿变换）。
    - "导出表格"将当前数据（含汇总行）写入 Excel。
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("质量特性汇总")
        self.setMinimumSize(1100, 650)
        self.resize(1300, 750)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowMinimizeButtonHint
        )

        self._settings = QSettings("CATIACompanion", "MassPropsDialog")
        self._last_browse_dir: str = self._settings.value("last_browse_dir", "")

        # ── 持久化显示选项 ─────────────────────────────────────────────────
        saved_hid = self._settings.value("visible_hideable_cols", list(MASS_PROPS_HIDEABLE_COLUMNS))
        if saved_hid is None:
            saved_hid = list(MASS_PROPS_HIDEABLE_COLUMNS)
        elif isinstance(saved_hid, str):
            saved_hid = [saved_hid]
        else:
            saved_hid = list(saved_hid)
        self._visible_hideable_cols: set[str] = {
            c for c in saved_hid if c in MASS_PROPS_HIDEABLE_COLUMNS
        }

        self._summarize: bool = self._settings.value("summarize", False, type=bool)
        self._unit: str = self._settings.value("unit", "g")
        # 内部单位为 g / g·mm²；根据所选单位制设置换算因子：
        #   "g"     → mass: ×1,      inertia: ×1         (g,     g·mm²)
        #   "kg"    → mass: ×0.001,  inertia: ×0.001     (kg,    kg·mm²)
        #   "kg_m2" → mass: ×0.001,  inertia: ×1e-9      (kg,    kg·m²)
        self._unit_factor, self._inertia_unit_factor = self._calc_unit_factors(self._unit)

        # ── 汇总BOM专用选项 ───────────────────────────────────────────────────
        self._summary_sort_column: str = self._settings.value(
            "summary_sort_column", ""
        )

        # ── 内部状态 ──────────────────────────────────────────────────────
        self._rows: list[dict] = []
        # display_row_idx → QTreeWidgetItem
        self._item_by_row: list[QTreeWidgetItem] = []
        # Part Number → list[QTreeWidgetItem] (all visible items with that PN)
        self._pn_to_items: dict[str, list[QTreeWidgetItem]] = {}
        self._is_updating: bool = False
        self._rollup_result: dict | None = None
        self._loaded: bool = False
        self._col_widths: dict[str, int] = {}

        # columns is rebuilt whenever visibility/mode changes
        self._columns: list[str] = self._build_columns()

        self._build_ui()

    # ── Column management ──────────────────────────────────────────────────

    @staticmethod
    def _calc_unit_factors(unit: str) -> tuple[float, float]:
        """根据单位制字符串返回 (mass_factor, inertia_factor)。

        Returns:
            mass_factor:    g → 显示单位的换算因子（重量列）
            inertia_factor: g·mm² → 显示单位的换算因子（惯量列）
        """
        if unit == "g":
            return 1.0, 1.0
        if unit == "kg_m2":
            return 1e-3, 1e-9
        # "kg" (kg / kg·mm²)
        return 1e-3, 1e-3

    def _weight_unit_label(self) -> str:
        """返回重量列的单位标签字符串。"""
        return "g" if self._unit == "g" else "kg"

    def _inertia_unit_label(self) -> str:
        """返回惯量列的单位标签字符串。"""
        if self._unit == "g":
            return "g·mm²"
        if self._unit == "kg_m2":
            return "kg·m²"
        return "kg·mm²"

    def _build_columns(self) -> list[str]:
        """根据当前可见性设置和 BOM 模式，构建列名列表。

        层级BOM：Level 在第 0 列（承载树形装饰线），# 在第 1 列。
        汇总BOM：无 Level 列，# 在第 0 列（无装饰线需求），增加 Quantity 列。
        """
        if self._summarize:
            base = ["#", "Type"]
            for c in MASS_PROPS_HIDEABLE_COLUMNS:
                if c in self._visible_hideable_cols:
                    base.append(c)
            base += ["Quantity", "Weight", "CogX", "CogY", "CogZ",
                     "Ixx", "Iyy", "Izz", "Ixy", "Ixz", "Iyz"]
        else:
            base = ["Level", "#", "Type"]
            for c in MASS_PROPS_HIDEABLE_COLUMNS:
                if c in self._visible_hideable_cols:
                    base.append(c)
            base += ["Weight", "CogX", "CogY", "CogZ",
                     "Ixx", "Iyy", "Izz", "Ixy", "Ixz", "Iyz"]
        return base

    def _column_header(self, col_name: str) -> str:
        """返回列名的中文显示名（含当前单位后缀）。"""
        if col_name == "Weight":
            return f"重量 ({self._weight_unit_label()})"
        if col_name in _INERTIA_IDX:
            return f"{col_name} ({self._inertia_unit_label()})"
        return MASS_PROPS_COLUMN_DISPLAY_NAMES.get(col_name, col_name)

    def _display_headers(self) -> list[str]:
        return [self._column_header(c) for c in self._columns]

    def _fmt_mass_val(self, value) -> str:
        """将质量原始值（g）乘以 _unit_factor 并格式化为字符串（重量列专用）。"""
        if value is None:
            return "—"
        try:
            v = float(value) * self._unit_factor
            if math.isclose(v, round(v), rel_tol=0.0, abs_tol=1e-6):
                return f"{v:.0f}"
            if abs(v) >= 1e4 or (v != 0.0 and abs(v) < 0.001):
                return f"{v:.3e}"
            return f"{v:.3f}"
        except (TypeError, ValueError):
            return str(value)

    def _fmt_inertia_val(self, value) -> str:
        """将惯量原始值（g·mm²）乘以 _inertia_unit_factor 并格式化为字符串（惯量列专用）。"""
        if value is None:
            return "—"
        try:
            v = float(value) * self._inertia_unit_factor
            if math.isclose(v, round(v), rel_tol=0.0, abs_tol=1e-6):
                return f"{v:.0f}"
            if abs(v) >= 1e4 or (v != 0.0 and abs(v) < 0.001):
                return f"{v:.3e}"
            return f"{v:.3f}"
        except (TypeError, ValueError):
            return str(value)

    # ── UI 构建 ────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── 数据来源选择 ────────────────────────────────────────────────────
        self._use_active_chk = QCheckBox("使用当前CATIA活动文档（不选择文件）")
        self._use_active_chk.toggled.connect(self._toggle_file_row)
        layout.addWidget(self._use_active_chk)

        file_row = QHBoxLayout()
        self._file_edit = QLineEdit()
        self._file_edit.setPlaceholderText("选择一个 CATProduct 文件…")
        self._file_edit.setReadOnly(True)
        self._file_browse_btn = QPushButton("浏览…")
        self._file_browse_btn.clicked.connect(self._browse_file)
        self._load_btn = QPushButton("加载")
        self._load_btn.clicked.connect(self._load_data)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(self._file_browse_btn)
        file_row.addWidget(self._load_btn)
        layout.addLayout(file_row)

        # ── 显示选项（BOM类型 + 单位 + 列可见性）──────────────────────────
        opts_group = QGroupBox("显示选项")
        opts_main = QVBoxLayout(opts_group)
        opts_main.setSpacing(6)
        opts_main.setContentsMargins(8, 6, 8, 6)

        # 第一行：BOM 类型 + 单位
        row1 = QHBoxLayout()
        row1.setSpacing(16)

        self._bom_type_group = QButtonGroup(self)
        self._radio_hier = QRadioButton("层级BOM")
        self._radio_summ = QRadioButton("汇总BOM")
        self._radio_hier.setChecked(not self._summarize)
        self._radio_summ.setChecked(self._summarize)
        self._bom_type_group.addButton(self._radio_hier)
        self._bom_type_group.addButton(self._radio_summ)
        self._radio_summ.toggled.connect(self._on_bom_type_changed)
        row1.addWidget(self._radio_hier)
        row1.addWidget(self._radio_summ)

        row1.addSpacing(24)

        unit_lbl = QLabel("单位：")
        row1.addWidget(unit_lbl)
        self._unit_group = QButtonGroup(self)
        self._radio_kg    = QRadioButton("kg / kg·mm²")
        self._radio_g     = QRadioButton("g / g·mm²")
        self._radio_kg_m2 = QRadioButton("kg / kg·m²")
        self._radio_kg.setChecked(self._unit == "kg")
        self._radio_g.setChecked(self._unit == "g")
        self._radio_kg_m2.setChecked(self._unit == "kg_m2")
        self._unit_group.addButton(self._radio_kg)
        self._unit_group.addButton(self._radio_g)
        self._unit_group.addButton(self._radio_kg_m2)
        self._radio_kg.toggled.connect(self._on_unit_changed)
        self._radio_g.toggled.connect(self._on_unit_changed)
        self._radio_kg_m2.toggled.connect(self._on_unit_changed)
        row1.addWidget(self._radio_kg)
        row1.addWidget(self._radio_g)
        row1.addWidget(self._radio_kg_m2)

        # 汇总BOM专用选项（排序列）
        self._summary_opts_widget = QWidget()
        summary_opts_layout = QHBoxLayout(self._summary_opts_widget)
        summary_opts_layout.setContentsMargins(0, 0, 0, 0)
        summary_opts_layout.setSpacing(8)
        summary_opts_layout.addSpacing(16)

        summary_opts_layout.addWidget(QLabel("排序列:"))
        self._sort_col_combo = QComboBox()
        self._sort_col_combo.addItem("（不排序）", "")
        for col in _SUMMARY_SORT_COLUMNS:
            self._sort_col_combo.addItem(MASS_PROPS_COLUMN_DISPLAY_NAMES.get(col, col), col)
        saved_sort_idx = self._sort_col_combo.findData(self._summary_sort_column)
        if saved_sort_idx >= 0:
            self._sort_col_combo.setCurrentIndex(saved_sort_idx)
        self._sort_col_combo.currentIndexChanged.connect(self._on_sort_col_changed)
        summary_opts_layout.addWidget(self._sort_col_combo)

        self._summary_opts_widget.setVisible(self._summarize)
        row1.addWidget(self._summary_opts_widget)
        row1.addStretch()

        opts_main.addLayout(row1)

        # 第二行：可隐藏列复选框
        row2 = QHBoxLayout()
        row2.setSpacing(12)
        lbl = QLabel("显示列：")
        row2.addWidget(lbl)
        self._hid_col_checks: dict[str, QCheckBox] = {}
        for col_name in MASS_PROPS_HIDEABLE_COLUMNS:
            cb = QCheckBox(MASS_PROPS_COLUMN_DISPLAY_NAMES.get(col_name, col_name))
            cb.setChecked(col_name in self._visible_hideable_cols)
            cb.setProperty("col_name", col_name)
            cb.toggled.connect(self._on_col_visibility_changed)
            row2.addWidget(cb)
            self._hid_col_checks[col_name] = cb
        row2.addStretch()
        opts_main.addLayout(row2)

        layout.addWidget(opts_group)

        # ── BOM说明标签 ─────────────────────────────────────────────────────
        self._bom_desc_lbl = QLabel(self._bom_desc_text())
        self._bom_desc_lbl.setWordWrap(True)
        self._bom_desc_lbl.setStyleSheet(
            "QLabel { background-color: #EEF4FC; border: 1px solid #B8D0F0;"
            " border-radius: 4px; padding: 4px 8px; color: #2B4C7E; font-size: 11px; }"
        )
        layout.addWidget(self._bom_desc_lbl)

        # ── 树形表格 ────────────────────────────────────────────────────────
        self._table = _BomTreeWidget()
        self._table.setColumnCount(len(self._columns))
        self._table.setHeaderLabels(self._display_headers())
        hdr = self._table.header()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        hdr.setSectionsMovable(False)
        hdr.setFixedHeight(28)
        self._table.setUniformRowHeights(True)
        self._table.setRootIsDecorated(True)
        self._table.setSortingEnabled(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setAlternatingRowColors(True)
        self._table.setIndentation(16)
        self._table.setStyleSheet("QTreeWidget::item { min-height: 24px; }")
        self._table.setItemDelegate(_MassPropsDelegate(lambda: self._columns, self._table))
        self._table.itemChanged.connect(self._on_item_changed)
        hdr.sectionResized.connect(self._on_section_resized)
        layout.addWidget(self._table, 1)

        # ── 汇总面板 ────────────────────────────────────────────────────────
        summary_group  = QGroupBox("汇总结果（基于根产品坐标系）")
        summary_layout = QGridLayout(summary_group)
        summary_layout.setSpacing(8)
        summary_layout.setContentsMargins(10, 8, 10, 8)

        def _lbl(text: str) -> QLabel:
            lb = QLabel(text)
            lb.setStyleSheet("font-weight: bold;")
            return lb

        def _val_lbl() -> QLabel:
            lb = QLabel("—")
            lb.setMinimumWidth(120)
            lb.setStyleSheet("font-family: monospace;")
            return lb

        summary_layout.addWidget(_lbl("总质量："), 0, 0)
        self._lbl_weight = _val_lbl()
        summary_layout.addWidget(self._lbl_weight, 0, 1)

        summary_layout.addWidget(_lbl("总重心 X："), 0, 2)
        self._lbl_cx = _val_lbl()
        summary_layout.addWidget(self._lbl_cx, 0, 3)

        summary_layout.addWidget(_lbl("总重心 Y："), 0, 4)
        self._lbl_cy = _val_lbl()
        summary_layout.addWidget(self._lbl_cy, 0, 5)

        summary_layout.addWidget(_lbl("总重心 Z："), 0, 6)
        self._lbl_cz = _val_lbl()
        summary_layout.addWidget(self._lbl_cz, 0, 7)

        inertia_entries = [
            ("Ixx:", "lbl_ixx"), ("Iyy:", "lbl_iyy"), ("Izz:", "lbl_izz"),
            ("Ixy:", "lbl_ixy"), ("Ixz:", "lbl_ixz"), ("Iyz:", "lbl_iyz"),
        ]
        for i, (text, attr) in enumerate(inertia_entries):
            r_i = 1 + i // 4
            c_i = (i % 4) * 2
            summary_layout.addWidget(_lbl(text), r_i, c_i)
            lbl = _val_lbl()
            setattr(self, f"_{attr}", lbl)
            summary_layout.addWidget(lbl, r_i, c_i + 1)

        layout.addWidget(summary_group)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        # ── 底部按钮行 ──────────────────────────────────────────────────────
        btn_row = QHBoxLayout()

        autofit_btn = QPushButton("自适应列宽")
        autofit_btn.clicked.connect(self._autofit_columns)
        btn_row.addWidget(autofit_btn)

        expand_btn = QPushButton("全部展开")
        expand_btn.clicked.connect(self._table.expandAll)
        btn_row.addWidget(expand_btn)

        collapse_btn = QPushButton("全部折叠")
        collapse_btn.clicked.connect(self._table.collapseAll)
        btn_row.addWidget(collapse_btn)

        btn_row.addStretch()

        self._calc_btn = QPushButton("计算")
        self._calc_btn.setToolTip("汇总装配体总质量特性（质量 / 重心 / 转动惯量）")
        self._calc_btn.setEnabled(False)
        self._calc_btn.clicked.connect(self._calculate)
        btn_row.addWidget(self._calc_btn)

        self._export_btn = QPushButton("导出表格")
        self._export_btn.setToolTip("将当前表格（含汇总行）导出为 Excel（.xlsx）文件")
        self._export_btn.setEnabled(False)
        self._export_btn.clicked.connect(self._export_table)
        btn_row.addWidget(self._export_btn)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.reject)
        btn_row.addWidget(close_btn)

        layout.addLayout(btn_row)

    # ── BOM 说明文字 ───────────────────────────────────────────────────────

    def _bom_desc_text(self) -> str:
        """返回当前 BOM 模式对应的说明文字。"""
        if self._summarize:
            return (
                "【汇总BOM】按零件编号合并，仅列出零件（不含产品和部件）。"
                "Weight / CogX / CogY / CogZ / Ixx–Iyz "
                "在零件自身坐标系下显示，与装配位置无关。"
                "底部「汇总结果」在根产品坐标系下计算。"
            )
        return (
            "【层级BOM】展示零件节点和产品/部件节点。"
            "Weight / CogX / CogY / CogZ / Ixx–Iyz "
            "在根产品坐标系下显示，与零件的装配位置有关。"
            "底部「汇总结果」在根产品坐标系下计算。"
        )

    # ── 文件/活动文档切换 ──────────────────────────────────────────────────

    def _toggle_file_row(self, use_active: bool) -> None:
        self._file_edit.setEnabled(not use_active)
        self._file_browse_btn.setEnabled(not use_active)

    def _browse_file(self) -> None:
        file, _ = QFileDialog.getOpenFileName(
            self, "选择CATProduct文件",
            self._last_browse_dir,
            "*.CATProduct (*.CATProduct);;All Files (*)",
        )
        if file:
            self._file_edit.setText(file)
            self._last_browse_dir = str(Path(file).parent)
            self._settings.setValue("last_browse_dir", self._last_browse_dir)

    # ── 显示选项响应 ───────────────────────────────────────────────────────

    def _on_bom_type_changed(self, checked: bool) -> None:
        self._summarize = self._radio_summ.isChecked()
        self._settings.setValue("summarize", self._summarize)
        self._summary_opts_widget.setVisible(self._summarize)
        self._bom_desc_lbl.setText(self._bom_desc_text())
        self._rebuild_columns_and_table()

    def _on_unit_changed(self, checked: bool) -> None:
        if self._radio_g.isChecked():
            self._unit = "g"
        elif self._radio_kg_m2.isChecked():
            self._unit = "kg_m2"
        else:
            self._unit = "kg"
        self._unit_factor, self._inertia_unit_factor = self._calc_unit_factors(self._unit)
        self._settings.setValue("unit", self._unit)
        if self._rows:
            self._refresh_unit_display()

    def _on_col_visibility_changed(self, checked: bool) -> None:
        for col_name, cb in self._hid_col_checks.items():
            if cb.isChecked():
                self._visible_hideable_cols.add(col_name)
            else:
                self._visible_hideable_cols.discard(col_name)
        self._settings.setValue("visible_hideable_cols",
                                list(self._visible_hideable_cols))
        self._rebuild_columns_and_table()

    def _on_sort_col_changed(self, _index: int) -> None:
        col = self._sort_col_combo.currentData()
        self._summary_sort_column = col or ""
        self._settings.setValue("summary_sort_column", self._summary_sort_column)
        if self._summarize and self._rows:
            self._populate_table()

    def _rebuild_columns_and_table(self) -> None:
        """重建列列表并重新填充表格（保留列宽）。"""
        if self._rows:
            # Save current column widths before rebuilding
            for col_idx, col_name in enumerate(self._columns):
                self._col_widths[col_name] = self._table.columnWidth(col_idx)
        self._columns = self._build_columns()
        self._populate_table()
        # Restore column widths
        for col_idx, col_name in enumerate(self._columns):
            if col_name in self._col_widths:
                self._table.setColumnWidth(col_idx, self._col_widths[col_name])

    def _refresh_unit_display(self) -> None:
        """仅更新列标题和重量/惯量单元格的显示值（单位切换时调用，避免全量重建）。"""
        # Update headers
        self._table.setHeaderLabels(self._display_headers())

        mass_col_indices: list[tuple[str, int]] = []
        for col_name in ("Weight",) + tuple(_INERTIA_IDX.keys()):
            if col_name in self._columns:
                mass_col_indices.append((col_name, self._columns.index(col_name)))

        if not mass_col_indices:
            return

        self._is_updating = True
        display_rows = self._get_display_rows()
        for di, row_data in enumerate(display_rows):
            if di >= len(self._item_by_row):
                break
            item = self._item_by_row[di]
            node_type = str(row_data.get("Type", ""))
            # 刷新所有有数据的节点（零件、产品、部件）
            if not any(row_data.get(c) is not None
                       for c in ("Weight",) + tuple(_INERTIA_IDX.keys())):
                continue
            for col_name, col_idx in mass_col_indices:
                raw = row_data.get(col_name)
                if raw is not None:
                    if col_name == "Weight":
                        item.setText(col_idx, self._fmt_mass_val(raw))
                    else:
                        item.setText(col_idx, self._fmt_inertia_val(raw))
        self._is_updating = False

        # Update summary labels if result is available
        if self._rollup_result:
            self._update_summary_labels(self._rollup_result)

    # ── 加载数据 ───────────────────────────────────────────────────────────

    def _load_data(self) -> None:
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

        progress = QProgressDialog("正在加载产品树，请稍候…", None, 0, 0, self)
        progress.setWindowTitle("加载质量特性")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(300)
        progress.setValue(0)

        def _on_row_collected(count: int) -> None:
            progress.setLabelText(f"正在加载产品树，请稍候… 已读取 {count} 个节点")
            progress.repaint()
            QApplication.processEvents()

        try:
            rows = collect_mass_props_rows(file_path, progress_callback=_on_row_collected)
        except Exception as e:
            progress.close()
            logger.error(f"加载质量特性失败: {e}")
            QMessageBox.critical(
                self, "加载失败",
                f"加载产品树时出错：\n{e}\n\n请确保CATIA已启动。",
            )
            self._load_btn.setEnabled(True)
            self._load_btn.setText("加载")
            return
        finally:
            progress.close()

        self._load_btn.setEnabled(True)
        self._load_btn.setText("重新加载")

        # Save column widths before repopulating
        if self._loaded:
            for col_idx, col_name in enumerate(self._columns):
                self._col_widths[col_name] = self._table.columnWidth(col_idx)

        self._rows = rows
        self._rollup_result = None
        self._clear_summary_labels()
        self._columns = self._build_columns()
        self._populate_table()

        if not self._loaded:
            # Auto-fit on first load
            for _c, col_name in enumerate(self._columns):
                if col_name == "#":
                    self._table.setColumnWidth(_c, 40)
                    self._col_widths[col_name] = 40
                else:
                    self._table.resizeColumnToContents(_c)
                    self._col_widths[col_name] = self._table.columnWidth(_c)
            self._loaded = True
        else:
            for col_idx, col_name in enumerate(self._columns):
                if col_name in self._col_widths:
                    self._table.setColumnWidth(col_idx, self._col_widths[col_name])

        self._calc_btn.setEnabled(True)
        self._export_btn.setEnabled(True)

        failed_count = sum(1 for r in rows if r.get("_meas_failed") and r.get("Type") == "零件")
        if failed_count:
            QMessageBox.information(
                self, "部分零件测量失败",
                f"有 {failed_count} 个零件节点无法完成质量特性测量（显示橙色背景）。\n\n"
                "可能原因：\n"
                "  • 零件文档未加载到CATIA会话中\n"
                "  • 零件无有效的保持测量的「惯量包络体.1」\n"
                "  • 零件未成功运行「创建质量关系」宏（MP_* 参数不存在）\n\n"
                "未能测量的零件不参与最终汇总计算。",
            )

        # 加载完成后自动计算汇总结果
        self._calculate()

    # ── Display row builders ───────────────────────────────────────────────

    def _get_display_rows(self) -> list[dict]:
        """返回当前模式下应显示的行列表。"""
        if self._summarize:
            return self._build_summary_rows()
        # 层级BOM：展示全部节点（零件、产品、部件），使用根产品坐标系下的值。
        # 每行附加 _rows_idx，指向 self._rows 中的原始索引，
        # 以确保 _make_item / _on_item_changed 能正确回写数据。
        # 对零件行，将 _root_mp 中的根坐标系 COG / 惯量值覆盖显示字段；
        # 产品/部件行的显示字段已由 _post_process_rows() 写入根坐标系汇总值。
        result = []
        for i, row in enumerate(self._rows):
            r = dict(row)
            r["_rows_idx"] = i
            if r.get("Type") == "零件":
                rmp = r.get("_root_mp")
                if rmp:
                    cog = rmp.get("cog", [None, None, None])
                    r["CogX"] = cog[0]
                    r["CogY"] = cog[1]
                    r["CogZ"] = cog[2]
                    I = rmp.get("inertia")
                    if I:
                        r["Ixx"] = I[0][0]
                        r["Iyy"] = I[1][1]
                        r["Izz"] = I[2][2]
                        r["Ixy"] = I[0][1]
                        r["Ixz"] = I[0][2]
                        r["Iyz"] = I[1][2]
            result.append(r)
        return result

    def _build_summary_rows(self) -> list[dict]:
        """汇总模式：将相同零件编号的行合并，增加 Quantity 字段。

        每个唯一 PN 保留第一次出现的行数据（含 _rows 中的索引），
        Quantity = 该 PN 在 _rows 中出现的实例数量。
        """
        seen_pn: dict[str, dict] = {}    # pn → canonical row copy
        qty: dict[str, int] = {}
        order: list[str] = []

        for i, row in enumerate(self._rows):
            pn = str(row.get("Part Number", ""))
            if not pn:
                pn = str(row.get("Filename", "")) or "(未分组)"
            if pn not in seen_pn:
                r = dict(row)
                r["_rows_idx"] = i   # link back to canonical _rows entry
                seen_pn[pn] = r
                qty[pn] = 1
                order.append(pn)
            else:
                qty[pn] += 1

        result = []
        for pn in order:
            r = dict(seen_pn[pn])
            r["Quantity"] = qty[pn]
            result.append(r)

        # 仅保留零件行（汇总BOM不显示产品和部件）
        result = [r for r in result if r.get("Type") == "零件"]

        # 按排序列排序
        if self._summary_sort_column:
            col = self._summary_sort_column
            result.sort(key=lambda r: str(r.get(col, "") or ""))

        return result

    # ── 填充表格 ───────────────────────────────────────────────────────────

    def _populate_table(self) -> None:
        self._is_updating = True
        self._table.blockSignals(True)

        self._table.clear()
        self._table.setColumnCount(len(self._columns))
        self._table.setHeaderLabels(self._display_headers())
        self._table.setRootIsDecorated(not self._summarize)
        self._item_by_row = []
        self._pn_to_items.clear()

        display_rows = self._get_display_rows()

        if self._summarize:
            self._populate_flat(display_rows)
        else:
            self._populate_tree(display_rows)

        self._table.expandAll()
        self._table.blockSignals(False)
        self._is_updating = False

    def _make_item(self, row_idx: int, row_data: dict) -> QTreeWidgetItem:
        """构建并填充一行的 QTreeWidgetItem。

        row_idx: 对应的 self._rows 索引（汇总模式用 _rows_idx 字段）。
        """
        item = QTreeWidgetItem()
        item.setData(0, _ROW_IDX_ROLE, row_idx)

        pn          = str(row_data.get("Part Number", ""))
        not_found   = bool(row_data.get("_not_found"))
        no_file     = bool(row_data.get("_no_file"))
        unreadable  = bool(row_data.get("_unreadable"))
        meas_failed = bool(row_data.get("_meas_failed"))
        node_type   = str(row_data.get("Type", ""))
        row_locked  = unreadable or not_found or meas_failed

        if pn:
            self._pn_to_items.setdefault(pn, []).append(item)

        seq_no = str(len(self._item_by_row) + 1)

        for col_idx, col_name in enumerate(self._columns):
            if col_name == "#":
                item.setText(col_idx, seq_no)
            elif col_name == "Level":
                item.setText(col_idx, str(row_data.get("Level", 0)))
            elif col_name == "Filename":
                fp = str(row_data.get("_filepath", ""))
                fn = str(row_data.get("Filename", ""))
                if no_file:
                    value = FILENAME_UNSAVED
                else:
                    value = Path(fp).name if fp else fn
                item.setText(col_idx, value)
                if no_file:
                    pass  # tooltip 由下方 no_file 块统一设置
                elif fp:
                    item.setToolTip(col_idx, fp)
            elif col_name == "Quantity":
                item.setText(col_idx, str(row_data.get("Quantity", 1)))
            elif col_name == "Weight":
                raw = row_data.get("Weight")
                if raw is None:
                    item.setText(col_idx, "—" if node_type == "零件" else "")
                else:
                    item.setText(col_idx, self._fmt_mass_val(raw))
            elif col_name in _INERTIA_IDX or col_name in ("CogX", "CogY", "CogZ"):
                raw = row_data.get(col_name)
                if raw is None:
                    item.setText(col_idx, "—" if node_type == "零件" else "")
                else:
                    if col_name in _INERTIA_IDX:
                        item.setText(col_idx, self._fmt_inertia_val(raw))
                    else:
                        item.setText(col_idx, _fmt(raw))
            else:
                item.setText(col_idx, str(row_data.get(col_name, "")))

        # Editability: only unlocked part rows, Weight column only
        if node_type == "零件" and not row_locked:
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            item.setData(0, _ITEM_LOCKED_ROLE, False)
        else:
            item.setData(0, _ITEM_LOCKED_ROLE, True)

        # Row colouring
        if row_locked:
            grey = QColor(160, 160, 160)
            if not_found:
                bg  = QColor(255, 205, 205)
                tip = "该零件/装配体的文件未被CATIA检索到，行内容不可编辑。"
            elif meas_failed:
                bg  = QColor(255, 210, 160)
                tip = "该零件的质量特性测量失败，行内容不可编辑。"
            else:
                bg  = QColor(245, 245, 245)
                tip = "该零件/装配体处于轻量化模式，无法读取属性。"
            for ci in range(len(self._columns)):
                item.setForeground(ci, grey)
                item.setBackground(ci, bg)
                item.setToolTip(ci, tip)
        elif no_file:
            bg_unsaved = QColor(255, 245, 180)
            no_file_tip = "该零件尚未保存到磁盘，质量特性数据可能不完整。"
            for ci in range(len(self._columns)):
                item.setBackground(ci, bg_unsaved)
                item.setToolTip(ci, no_file_tip)
        elif node_type in ("产品", "部件"):
            bg = QColor(240, 242, 245)
            for ci in range(len(self._columns)):
                item.setBackground(ci, bg)

        self._item_by_row.append(item)
        return item

    def _populate_flat(self, display_rows: list[dict]) -> None:
        """汇总BOM模式：所有行为顶级项（无树形层级）。"""
        for di, row_data in enumerate(display_rows):
            rows_idx = row_data.get("_rows_idx", di)
            item = self._make_item(rows_idx, row_data)
            self._table.addTopLevelItem(item)

    def _populate_tree(self, display_rows: list[dict]) -> None:
        """层级BOM模式：按 Level 构建树形结构。"""
        parent_stack: list[tuple[int, QTreeWidgetItem | None]] = [(-1, None)]

        for di, row_data in enumerate(display_rows):
            level = int(row_data.get("Level", 0))
            # 使用 _rows_idx（若存在）映射回 self._rows，保持与 _populate_flat 一致
            rows_idx = row_data.get("_rows_idx", di)

            while len(parent_stack) > 1 and parent_stack[-1][0] >= level:
                parent_stack.pop()

            parent_item = parent_stack[-1][1]
            item = self._make_item(rows_idx, row_data)

            if parent_item is None:
                self._table.addTopLevelItem(item)
            else:
                parent_item.addChild(item)

            parent_stack.append((level, item))

    # ── 单元格编辑 ─────────────────────────────────────────────────────────

    def _on_item_changed(self, item: QTreeWidgetItem, col_idx: int) -> None:
        if self._is_updating:
            return
        row_idx = item.data(0, _ROW_IDX_ROLE)
        if row_idx is None:
            return

        col_name = self._columns[col_idx]
        if col_name != "Weight":
            return

        if item.data(0, _ITEM_LOCKED_ROLE):
            return

        row_data = self._rows[row_idx]
        if row_data.get("Type") != "零件":
            return

        new_text = item.text(col_idx).strip()
        try:
            # 输入值为当前显示单位；除以 _unit_factor 还原到内部单位（g）
            new_display_val = float(new_text)
            new_weight_stored = new_display_val / self._unit_factor
        except (ValueError, TypeError):
            self._is_updating = True
            item.setText(col_idx, self._fmt_mass_val(row_data.get("Weight")))
            self._is_updating = False
            return

        if new_weight_stored < 0.0:
            QMessageBox.warning(
                self, "重量不合法",
                "重量不能为负数，请输入大于或等于 0 的值。",
            )
            self._is_updating = True
            item.setText(col_idx, self._fmt_mass_val(row_data.get("Weight")))
            self._is_updating = False
            return

        pn = str(row_data.get("Part Number", ""))

        # ── Update ALL _rows entries with the same PN ──────────────────────
        for r in self._rows:
            if str(r.get("Part Number", "")) != pn or r.get("Type") != "零件":
                continue
            old_w = r.get("Weight")
            try:
                old_w_f = float(old_w) if old_w is not None else 0.0
            except (ValueError, TypeError):
                old_w_f = 0.0
            scale = (new_weight_stored / old_w_f) if old_w_f > 0.0 else 1.0
            r["Weight"] = new_weight_stored
            mp = r.get("_mass_props")
            if mp:
                mp["weight"] = new_weight_stored
                if scale != 1.0 and old_w_f > 0.0:
                    orig_i = mp.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)])
                    mp["inertia"] = [[orig_i[ir][ic] * scale for ic in range(3)]
                                     for ir in range(3)]
                    # 同步更新 _root_mp 中的惯量（缩放后重新旋转到根坐标系）
                    I_root = _row_inertia_to_root(r)
                    rmp = r.get("_root_mp")
                    if rmp is not None:
                        rmp["inertia"] = I_root
                        rmp["weight"]  = new_weight_stored
                    # 更新行级惯量显示字段（零件自身坐标系）
                    I_local_new = mp.get("inertia", [[0.0] * 3 for _ in range(3)])
                    for ic_name, (ir2, ic2) in _INERTIA_IDX.items():
                        r[ic_name] = I_local_new[ir2][ic2]
                else:
                    # 仅更新质量，惯量不变（显示字段为局部坐标系值，无需重新计算）
                    rmp = r.get("_root_mp")
                    if rmp is not None:
                        rmp["weight"] = new_weight_stored
            else:
                # 无 _mass_props，直接从行字段读取并缩放
                if scale != 1.0 and old_w_f > 0.0:
                    for ic_name in _INERTIA_IDX:
                        cur = r.get(ic_name)
                        if cur is not None:
                            r[ic_name] = float(cur) * scale

        # ── Update visible tree items with the same PN ─────────────────────
        self._is_updating = True
        w_idx = self._columns.index("Weight") if "Weight" in self._columns else -1

        for vis_item in self._pn_to_items.get(pn, []):
            vis_row_idx = vis_item.data(0, _ROW_IDX_ROLE)
            if vis_row_idx is None:
                continue
            vis_row = self._rows[vis_row_idx]
            if vis_row.get("Type") != "零件":
                continue
            if w_idx >= 0:
                vis_item.setText(w_idx, self._fmt_mass_val(vis_row.get("Weight")))
            for ic_name, (ir, ic) in _INERTIA_IDX.items():
                if ic_name in self._columns:
                    ic_idx = self._columns.index(ic_name)
                    if self._summarize:
                        # 汇总BOM：显示零件自身坐标系值
                        raw_i = vis_row.get(ic_name)
                    else:
                        # 层级BOM：显示根产品坐标系值
                        rmp = vis_row.get("_root_mp")
                        raw_i = (
                            rmp["inertia"][ir][ic]
                            if rmp and rmp.get("inertia")
                            else vis_row.get(ic_name)
                        )
                    if raw_i is not None:
                        vis_item.setText(ic_idx, self._fmt_inertia_val(raw_i))

        self._is_updating = False
        self._rollup_result = None
        self._clear_summary_labels()

    # ── 计算 ───────────────────────────────────────────────────────────────

    def _calculate(self) -> None:
        if not self._rows:
            return
        # 先重新计算产品/部件行（使用更新后的 _root_mp），刷新表格
        recompute_product_rows(self._rows)
        self._refresh_product_items()
        try:
            result = rollup_mass_properties(self._rows)
        except Exception as e:
            logger.error(f"质量特性计算失败: {e}")
            QMessageBox.critical(self, "计算失败", f"计算总质量特性时出错：\n{e}")
            return
        self._rollup_result = result
        self._update_summary_labels(result)

    def _refresh_product_items(self) -> None:
        """刷新树形表格中所有产品/部件行的显示值（仅层级BOM模式有效）。

        在 _calculate() 调用 recompute_product_rows() 更新 self._rows 后，
        调用本方法将新的汇总值写回对应的 QTreeWidgetItem，以保持表格与数据同步。
        汇总BOM不含产品/部件行，故直接返回。
        """
        if self._summarize:
            return
        self._is_updating = True
        try:
            for item in self._item_by_row:
                row_idx = item.data(0, _ROW_IDX_ROLE)
                if row_idx is None:
                    continue
                row_data = self._rows[row_idx]
                if row_data.get("Type") not in ("产品", "部件"):
                    continue
                for col_idx, col_name in enumerate(self._columns):
                    if col_name == "Weight":
                        raw = row_data.get("Weight")
                        item.setText(col_idx, self._fmt_mass_val(raw) if raw is not None else "")
                    elif col_name in ("CogX", "CogY", "CogZ"):
                        raw = row_data.get(col_name)
                        item.setText(col_idx, _fmt(raw) if raw is not None else "")
                    elif col_name in _INERTIA_IDX:
                        raw = row_data.get(col_name)
                        item.setText(col_idx, self._fmt_inertia_val(raw) if raw is not None else "")
        finally:
            self._is_updating = False

    def _clear_summary_labels(self) -> None:
        for lbl in (self._lbl_weight, self._lbl_cx, self._lbl_cy, self._lbl_cz,
                    self._lbl_ixx, self._lbl_iyy, self._lbl_izz,
                    self._lbl_ixy, self._lbl_ixz, self._lbl_iyz):
            lbl.setText("—")

    def _update_summary_labels(self, result: dict) -> None:
        unit_lbl     = self._weight_unit_label()
        inertia_unit = self._inertia_unit_label()
        w_val = result.get("total_weight", 0.0)
        self._lbl_weight.setText(f"{self._fmt_mass_val(w_val)} {unit_lbl}")
        cog = result.get("cog", [0.0, 0.0, 0.0])
        self._lbl_cx.setText(f"{_fmt(cog[0])} mm")
        self._lbl_cy.setText(f"{_fmt(cog[1])} mm")
        self._lbl_cz.setText(f"{_fmt(cog[2])} mm")
        I = result.get("inertia", [[0.0] * 3 for _ in range(3)])
        self._lbl_ixx.setText(f"{self._fmt_inertia_val(I[0][0])} {inertia_unit}")
        self._lbl_iyy.setText(f"{self._fmt_inertia_val(I[1][1])} {inertia_unit}")
        self._lbl_izz.setText(f"{self._fmt_inertia_val(I[2][2])} {inertia_unit}")
        self._lbl_ixy.setText(f"{self._fmt_inertia_val(I[0][1])} {inertia_unit}")
        self._lbl_ixz.setText(f"{self._fmt_inertia_val(I[0][2])} {inertia_unit}")
        self._lbl_iyz.setText(f"{self._fmt_inertia_val(I[1][2])} {inertia_unit}")

    # ── 导出 ───────────────────────────────────────────────────────────────

    def _export_table(self) -> None:
        if not self._rows:
            QMessageBox.warning(self, "无数据", "请先加载产品树数据。")
            return

        default_dir = self._last_browse_dir or ""
        dest, _ = QFileDialog.getSaveFileName(
            self, "导出质量特性表格",
            str(Path(default_dir) / "质量特性.xlsx"),
            "Excel 文件 (*.xlsx);;CSV 文件 (*.csv);;所有文件 (*)",
        )
        if not dest:
            return

        dest_path = Path(dest)
        suffix = dest_path.suffix.lower()
        if suffix not in (".xlsx", ".csv"):
            dest_path = dest_path.with_suffix(".xlsx")
            suffix = ".xlsx"

        try:
            if suffix == ".csv":
                self._do_export_csv(dest_path)
            else:
                self._do_export(str(dest_path))
            QMessageBox.information(self, "导出成功", f"文件已保存到：\n{dest_path}")
        except Exception as e:
            logger.error(f"导出失败: {e}")
            QMessageBox.critical(self, "导出失败", f"导出时出错：\n{e}")

    def _do_export(self, dest: str) -> None:
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from catia_copilot.utils import estimate_column_width

        # Export columns (omit internal "#" column)
        export_cols = [c for c in self._columns if c != "#"]

        wb  = openpyxl.Workbook()
        ws  = wb.active
        ws.title = "质量特性"

        center      = Alignment(horizontal="center", vertical="center")
        header_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
        thin_side   = Side(style="thin")
        thin_border = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side,
        )

        # Table header
        for ci, col_name in enumerate(export_cols, start=1):
            cell = ws.cell(row=1, column=ci, value=self._column_header(col_name))
            cell.font   = Font(bold=True)
            cell.fill   = header_fill
            cell.border = thin_border

        # Data rows
        display_rows = self._get_display_rows()
        for ri, row_data in enumerate(display_rows, start=2):
            for ci, col_name in enumerate(export_cols, start=1):
                raw = row_data.get(col_name)
                if raw is None:
                    value = ""
                elif col_name == "Weight":
                    try:
                        value = float(raw) * self._unit_factor
                    except (TypeError, ValueError):
                        value = ""
                elif col_name in _INERTIA_IDX:
                    try:
                        value = float(raw) * self._inertia_unit_factor
                    except (TypeError, ValueError):
                        value = ""
                elif col_name in ("CogX", "CogY", "CogZ"):
                    try:
                        value = float(raw)
                    except (TypeError, ValueError):
                        value = ""
                else:
                    value = raw
                cell = ws.cell(row=ri, column=ci, value=value)
                cell.border = thin_border
                if col_name == "Level":
                    cell.alignment = center

        # Summary row (if calculated)
        if self._rollup_result:
            summary_row_idx = len(display_rows) + 2
            cog = self._rollup_result.get("cog", [0.0, 0.0, 0.0])
            I   = self._rollup_result.get("inertia", [[0.0] * 3 for _ in range(3)])
            w   = self._rollup_result.get("total_weight", 0.0)
            summary = {
                "Part Number":  "总计 (根产品)",
                "Weight":       w * self._unit_factor,
                "CogX":         cog[0],
                "CogY":         cog[1],
                "CogZ":         cog[2],
                "Ixx":          I[0][0] * self._inertia_unit_factor,
                "Iyy":          I[1][1] * self._inertia_unit_factor,
                "Izz":          I[2][2] * self._inertia_unit_factor,
                "Ixy":          I[0][1] * self._inertia_unit_factor,
                "Ixz":          I[0][2] * self._inertia_unit_factor,
                "Iyz":          I[1][2] * self._inertia_unit_factor,
            }
            summary_fill = PatternFill(fill_type="solid", fgColor="C6EFCE")
            for ci, col_name in enumerate(export_cols, start=1):
                val  = summary.get(col_name, "")
                cell = ws.cell(row=summary_row_idx, column=ci, value=val)
                cell.font   = Font(bold=True)
                cell.fill   = summary_fill
                cell.border = thin_border

        ws.freeze_panes = "A2"

        # Auto-fit column widths
        for ci, col_name in enumerate(export_cols, start=1):
            col_letter = ws.cell(row=1, column=ci).column_letter
            header     = self._column_header(col_name)
            max_width  = max(estimate_column_width(header), 8)
            for row_i in range(2, ws.max_row + 1):
                cv = ws.cell(row=row_i, column=ci).value
                if cv is not None:
                    max_width = max(max_width, estimate_column_width(str(cv)))
            ws.column_dimensions[col_letter].width = max_width + 2

        wb.save(dest)

    def _do_export_csv(self, dest: Path) -> None:
        """将当前表格数据（含汇总行）写入 UTF-8 with BOM 的 CSV 文件。"""
        import csv

        export_cols = [c for c in self._columns if c != "#"]
        display_rows = self._get_display_rows()

        def _cell_value(col_name: str, raw) -> str:
            if raw is None:
                return ""
            if col_name == "Weight":
                try:
                    return str(float(raw) * self._unit_factor)
                except (TypeError, ValueError):
                    return ""
            if col_name in _INERTIA_IDX:
                try:
                    return str(float(raw) * self._inertia_unit_factor)
                except (TypeError, ValueError):
                    return ""
            return str(raw)

        with open(dest, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([self._column_header(c) for c in export_cols])
            for row_data in display_rows:
                writer.writerow([
                    _cell_value(c, row_data.get(c)) for c in export_cols
                ])
            if self._rollup_result:
                cog = self._rollup_result.get("cog", [0.0, 0.0, 0.0])
                I   = self._rollup_result.get("inertia", [[0.0] * 3 for _ in range(3)])
                w   = self._rollup_result.get("total_weight", 0.0)
                summary = {
                    "Part Number":  "总计 (根产品)",
                    "Weight":       str(w * self._unit_factor),
                    "CogX":         str(cog[0]),
                    "CogY":         str(cog[1]),
                    "CogZ":         str(cog[2]),
                    "Ixx":          str(I[0][0] * self._inertia_unit_factor),
                    "Iyy":          str(I[1][1] * self._inertia_unit_factor),
                    "Izz":          str(I[2][2] * self._inertia_unit_factor),
                    "Ixy":          str(I[0][1] * self._inertia_unit_factor),
                    "Ixz":          str(I[0][2] * self._inertia_unit_factor),
                    "Iyz":          str(I[1][2] * self._inertia_unit_factor),
                }
                writer.writerow([summary.get(c, "") for c in export_cols])
        logger.info(f"质量特性表格已导出 (csv) -> {dest}")

    def _autofit_columns(self) -> None:
        min_width = 60
        for col_idx in range(len(self._columns)):
            self._table.resizeColumnToContents(col_idx)
            if self._table.columnWidth(col_idx) < min_width:
                self._table.setColumnWidth(col_idx, min_width)
        for col_idx, col_name in enumerate(self._columns):
            self._col_widths[col_name] = self._table.columnWidth(col_idx)

    def _on_section_resized(self, logical_index: int, _old: int, new_size: int) -> None:
        if logical_index < len(self._columns):
            self._col_widths[self._columns[logical_index]] = new_size
