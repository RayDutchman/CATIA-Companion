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
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTreeWidgetItem, QHeaderView, QAbstractItemView,
    QCheckBox, QGroupBox, QMessageBox, QApplication,
    QFileDialog, QProgressDialog, QLineEdit, QGridLayout, QFrame,
    QRadioButton, QButtonGroup, QWidget,
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QSettings

from catia_copilot.constants import (
    MASS_PROPS_COLUMNS,
    MASS_PROPS_COLUMN_DISPLAY_NAMES,
    MASS_PROPS_HIDEABLE_COLUMNS,
    FILENAME_NOT_FOUND,
)
from catia_copilot.catia.mass_props_collect import collect_mass_props_rows
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


def _fmt(value, digits: int = 4) -> str:
    """数值 → 字符串，None → '—'（不含单位换算）。"""
    if value is None:
        return "—"
    try:
        return f"{float(value):.{digits}f}"
    except (TypeError, ValueError):
        return str(value)


class MassPropsDialog(QDialog):
    """质量特性汇总对话框。

    - 遍历 CATProduct 树，每个零件实例单独显示一行（层级BOM模式）。
    - 汇总BOM模式：相同零件编号的实例合并为一行，并显示数量。
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
        self._unit: str = self._settings.value("unit", "kg")
        self._unit_factor: float = 1000.0 if self._unit == "g" else 1.0

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
        if self._unit == "g":
            if col_name == "Weight":
                return "重量 (g)"
            if col_name in _INERTIA_IDX:
                return f"{col_name} (g·mm²)"
        return MASS_PROPS_COLUMN_DISPLAY_NAMES.get(col_name, col_name)

    def _display_headers(self) -> list[str]:
        return [self._column_header(c) for c in self._columns]

    def _fmt_mass_val(self, value) -> str:
        """将质量/惯量原始值（kg / kg·mm²）转换并格式化为当前单位的字符串。"""
        if value is None:
            return "—"
        try:
            return f"{float(value) * self._unit_factor:.4f}"
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
        self._radio_kg = QRadioButton("kg / kg·mm²")
        self._radio_g = QRadioButton("g / g·mm²")
        self._radio_kg.setChecked(self._unit == "kg")
        self._radio_g.setChecked(self._unit == "g")
        self._unit_group.addButton(self._radio_kg)
        self._unit_group.addButton(self._radio_g)
        self._radio_kg.toggled.connect(self._on_unit_changed)
        row1.addWidget(self._radio_kg)
        row1.addWidget(self._radio_g)
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
        self._rebuild_columns_and_table()

    def _on_unit_changed(self, checked: bool) -> None:
        self._unit = "g" if self._radio_g.isChecked() else "kg"
        self._unit_factor = 1000.0 if self._unit == "g" else 1.0
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
            if node_type not in ("零件",):
                continue
            for col_name, col_idx in mass_col_indices:
                raw = row_data.get(col_name)
                if raw is not None:
                    item.setText(col_idx, self._fmt_mass_val(raw))
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

        failed_count = sum(1 for r in rows if r.get("_spa_failed") and r.get("Type") == "零件")
        if failed_count:
            QMessageBox.information(
                self, "部分零件测量失败",
                f"有 {failed_count} 个零件节点无法完成SPA质量测量（显示橙色背景）。\n\n"
                "可能原因：\n"
                "  • 零件文档未加载到CATIA会话中\n"
                "  • CATIA SPA工作台不可用\n"
                "  • 零件无几何实体\n\n"
                "未能测量的零件不参与最终汇总计算。",
            )

    # ── Display row builders ───────────────────────────────────────────────

    def _get_display_rows(self) -> list[dict]:
        """返回当前模式下应显示的行列表。"""
        if self._summarize:
            return self._build_summary_rows()
        return self._rows

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

        pn        = str(row_data.get("Part Number", ""))
        not_found = bool(row_data.get("_not_found"))
        unreadable = bool(row_data.get("_unreadable"))
        spa_failed = bool(row_data.get("_spa_failed"))
        node_type  = str(row_data.get("Type", ""))
        row_locked = unreadable or not_found

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
                value = Path(fp).name if fp else fn
                item.setText(col_idx, value)
                if fp:
                    item.setToolTip(col_idx, fp)
            elif col_name == "Quantity":
                item.setText(col_idx, str(row_data.get("Quantity", 1)))
            elif col_name == "Weight":
                raw = row_data.get("Weight")
                if raw is None or node_type not in ("零件",):
                    item.setText(col_idx, "" if node_type in ("产品", "部件") else "—")
                else:
                    item.setText(col_idx, self._fmt_mass_val(raw))
            elif col_name in _INERTIA_IDX or col_name in ("CogX", "CogY", "CogZ"):
                raw = row_data.get(col_name)
                if raw is None or node_type not in ("零件",):
                    item.setText(col_idx, "" if node_type in ("产品", "部件") else "—")
                else:
                    if col_name in _INERTIA_IDX:
                        item.setText(col_idx, self._fmt_mass_val(raw))
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
            bg   = QColor(255, 205, 205) if not_found else QColor(245, 245, 245)
            for ci in range(len(self._columns)):
                item.setForeground(ci, grey)
                item.setBackground(ci, bg)
        elif spa_failed and node_type == "零件":
            bg = QColor(255, 210, 160)
            for ci in range(len(self._columns)):
                item.setBackground(ci, bg)
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

        for rows_idx, row_data in enumerate(display_rows):
            level = int(row_data.get("Level", 0))
            # In hierarchical mode display_rows IS self._rows, so rows_idx == _rows index

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
            # Input is in current unit; convert back to kg for storage
            new_display_val = float(new_text)
            new_weight_kg = new_display_val / self._unit_factor
        except (ValueError, TypeError):
            self._is_updating = True
            item.setText(col_idx, self._fmt_mass_val(row_data.get("Weight")))
            self._is_updating = False
            return

        if new_weight_kg < 0.0:
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
            scale = (new_weight_kg / old_w_f) if old_w_f > 0.0 else 1.0
            r["Weight"] = new_weight_kg
            mp = r.get("_mass_props")
            if mp:
                mp["weight"] = new_weight_kg
                if scale != 1.0 and old_w_f > 0.0:
                    orig_i = mp.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)])
                    mp["inertia"] = [[orig_i[ri][ci] * scale for ci in range(3)]
                                     for ri in range(3)]
            # Update row-level inertia fields too
            mp_cur = r.get("_mass_props") or {}
            inertia_cur = mp_cur.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)])
            for ic_name, (ri, ci) in _INERTIA_IDX.items():
                r[ic_name] = inertia_cur[ri][ci]

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
            mp_vis = vis_row.get("_mass_props") or {}
            inertia_vis = mp_vis.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)])
            for ic_name, (ri, ci) in _INERTIA_IDX.items():
                if ic_name in self._columns:
                    ic_idx = self._columns.index(ic_name)
                    vis_item.setText(ic_idx, self._fmt_mass_val(inertia_vis[ri][ci]))

        self._is_updating = False
        self._rollup_result = None
        self._clear_summary_labels()

    # ── 计算 ───────────────────────────────────────────────────────────────

    def _calculate(self) -> None:
        if not self._rows:
            return
        try:
            result = rollup_mass_properties(self._rows)
        except Exception as e:
            logger.error(f"质量特性计算失败: {e}")
            QMessageBox.critical(self, "计算失败", f"计算总质量特性时出错：\n{e}")
            return
        self._rollup_result = result
        self._update_summary_labels(result)

    def _clear_summary_labels(self) -> None:
        for lbl in (self._lbl_weight, self._lbl_cx, self._lbl_cy, self._lbl_cz,
                    self._lbl_ixx, self._lbl_iyy, self._lbl_izz,
                    self._lbl_ixy, self._lbl_ixz, self._lbl_iyz):
            lbl.setText("—")

    def _update_summary_labels(self, result: dict) -> None:
        unit_lbl = self._unit  # "kg" or "g"
        inertia_unit = f"{unit_lbl}·mm²"
        w_val = result.get("total_weight", 0.0)
        self._lbl_weight.setText(f"{w_val * self._unit_factor:.4f} {unit_lbl}")
        cog = result.get("cog", [0.0, 0.0, 0.0])
        self._lbl_cx.setText(f"{cog[0]:.4f} mm")
        self._lbl_cy.setText(f"{cog[1]:.4f} mm")
        self._lbl_cz.setText(f"{cog[2]:.4f} mm")
        I = result.get("inertia", [[0.0] * 3 for _ in range(3)])
        self._lbl_ixx.setText(f"{I[0][0] * self._unit_factor:.4f} {inertia_unit}")
        self._lbl_iyy.setText(f"{I[1][1] * self._unit_factor:.4f} {inertia_unit}")
        self._lbl_izz.setText(f"{I[2][2] * self._unit_factor:.4f} {inertia_unit}")
        self._lbl_ixy.setText(f"{I[0][1] * self._unit_factor:.4f} {inertia_unit}")
        self._lbl_ixz.setText(f"{I[0][2] * self._unit_factor:.4f} {inertia_unit}")
        self._lbl_iyz.setText(f"{I[1][2] * self._unit_factor:.4f} {inertia_unit}")

    # ── 导出 ───────────────────────────────────────────────────────────────

    def _export_table(self) -> None:
        if not self._rows:
            QMessageBox.warning(self, "无数据", "请先加载产品树数据。")
            return

        default_dir = self._last_browse_dir or ""
        dest, _ = QFileDialog.getSaveFileName(
            self, "导出质量特性表格",
            str(Path(default_dir) / "质量特性.xlsx"),
            "Excel 文件 (*.xlsx);;所有文件 (*)",
        )
        if not dest:
            return

        try:
            self._do_export(dest)
            QMessageBox.information(self, "导出成功", f"文件已保存到：\n{dest}")
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
                        value = float(raw) * self._unit_factor
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
                "Ixx":          I[0][0] * self._unit_factor,
                "Iyy":          I[1][1] * self._unit_factor,
                "Izz":          I[2][2] * self._unit_factor,
                "Ixy":          I[0][1] * self._unit_factor,
                "Ixz":          I[0][2] * self._unit_factor,
                "Iyz":          I[1][2] * self._unit_factor,
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

    # ── 列宽辅助 ───────────────────────────────────────────────────────────

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
