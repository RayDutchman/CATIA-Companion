"""
质量特性汇总对话框模块。

提供：
- MassPropsDialog – 遍历产品树，展示每个零件实例的质量/重心/转动惯量，
                    支持手动编辑重量（自动等比缩放惯量并联动同型号零件），
                    计算装配体总质量特性并导出为 Excel 表格。
"""

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTreeWidgetItem, QHeaderView, QAbstractItemView,
    QCheckBox, QGroupBox, QMessageBox, QApplication,
    QFileDialog, QProgressDialog, QLineEdit, QGridLayout, QFrame,
)
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QSettings

from catia_copilot.constants import (
    MASS_PROPS_COLUMNS,
    MASS_PROPS_COLUMN_DISPLAY_NAMES,
    MASS_PROPS_READONLY_COLUMNS,
    FILENAME_NOT_FOUND,
)
from catia_copilot.catia.mass_props_collect import collect_mass_props_rows
from catia_copilot.catia.mass_props_calc import rollup_mass_properties
from catia_copilot.ui.bom_widgets import _BomTreeWidget

logger = logging.getLogger(__name__)

# UserRole for row index (reuse same pattern as BomEditDialog)
_ROW_IDX_ROLE = Qt.ItemDataRole.UserRole
# UserRole+1 for "locked" flag
_ITEM_LOCKED_ROLE = Qt.ItemDataRole.UserRole + 1


def _fmt(value, digits: int = 4) -> str:
    """将数值格式化为字符串，None → '—'。"""
    if value is None:
        return "—"
    try:
        return f"{float(value):.{digits}f}"
    except (TypeError, ValueError):
        return str(value)


class MassPropsDialog(QDialog):
    """质量特性汇总对话框。

    - 遍历 CATProduct 树，每个零件实例单独显示一行。
    - 仅零件节点的"重量"列可编辑；修改后等比缩放该行惯量，
      并同步更新所有相同零件编号的行。
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

        # 内部状态
        self._rows: list[dict] = []
        self._item_by_row: list[QTreeWidgetItem] = []
        self._pn_to_items: dict[str, list[QTreeWidgetItem]] = {}
        self._is_updating: bool = False
        self._rollup_result: dict | None = None
        self._loaded: bool = False
        self._col_widths: dict[str, int] = {}

        # 固定列顺序（始终显示所有列）
        self._columns: list[str] = (
            ["#", "Level"] + [c for c in MASS_PROPS_COLUMNS if c != "Level"]
        )

        self._build_ui()

    # ── UI 构建 ────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # 数据来源选择行
        self._use_active_chk = QCheckBox("使用当前CATIA活动文档（不选择文件）")
        self._use_active_chk.toggled.connect(self._toggle_file_row)
        layout.addWidget(self._use_active_chk)

        file_row = QHBoxLayout()
        self._file_edit       = QLineEdit()
        self._file_edit.setPlaceholderText("选择一个 CATProduct 文件…")
        self._file_edit.setReadOnly(True)
        self._file_browse_btn = QPushButton("浏览…")
        self._file_browse_btn.clicked.connect(self._browse_file)
        self._load_btn        = QPushButton("加载")
        self._load_btn.clicked.connect(self._load_data)
        file_row.addWidget(self._file_edit)
        file_row.addWidget(self._file_browse_btn)
        file_row.addWidget(self._load_btn)
        layout.addLayout(file_row)

        hint = QLabel(
            "仅零件节点（类型=零件）的「重量」列可编辑；修改后将等比缩放该零件的转动惯量，"
            "并同步更新装配体中所有相同零件编号的行。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(hint)

        # 树形表格
        self._table = _BomTreeWidget()
        headers = [MASS_PROPS_COLUMN_DISPLAY_NAMES.get(c, c) for c in self._columns]
        self._table.setColumnCount(len(headers))
        self._table.setHeaderLabels(headers)
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

        # 总质量
        summary_layout.addWidget(_lbl("总质量："), 0, 0)
        self._lbl_weight = _val_lbl()
        summary_layout.addWidget(self._lbl_weight, 0, 1)

        # 重心坐标
        summary_layout.addWidget(_lbl("总重心 X："), 0, 2)
        self._lbl_cx = _val_lbl()
        summary_layout.addWidget(self._lbl_cx, 0, 3)

        summary_layout.addWidget(_lbl("总重心 Y："), 0, 4)
        self._lbl_cy = _val_lbl()
        summary_layout.addWidget(self._lbl_cy, 0, 5)

        summary_layout.addWidget(_lbl("总重心 Z："), 0, 6)
        self._lbl_cz = _val_lbl()
        summary_layout.addWidget(self._lbl_cz, 0, 7)

        # 转动惯量
        inertia_cols = [
            ("Ixx:", "lbl_ixx"), ("Iyy:", "lbl_iyy"), ("Izz:", "lbl_izz"),
            ("Ixy:", "lbl_ixy"), ("Ixz:", "lbl_ixz"), ("Iyz:", "lbl_iyz"),
        ]
        for i, (text, attr) in enumerate(inertia_cols):
            row_i = 1 + i // 4
            col_i = (i % 4) * 2
            summary_layout.addWidget(_lbl(text), row_i, col_i)
            lbl = _val_lbl()
            setattr(self, f"_{attr}", lbl)
            summary_layout.addWidget(lbl, row_i, col_i + 1)

        layout.addWidget(summary_group)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        # 底部按钮行
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
            rows = collect_mass_props_rows(
                file_path,
                progress_callback=_on_row_collected,
            )
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

        # 保存列宽
        if self._loaded:
            for col_idx, col_name in enumerate(self._columns):
                self._col_widths[col_name] = self._table.columnWidth(col_idx)

        self._rows = rows
        self._rollup_result = None
        self._clear_summary_labels()
        self._populate_table()

        if not self._loaded:
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

        # 有 SPA 测量失败的行时给出提示
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

    # ── 填充表格 ───────────────────────────────────────────────────────────

    def _populate_table(self) -> None:
        self._is_updating = True
        self._table.blockSignals(True)

        self._table.clear()
        headers = [MASS_PROPS_COLUMN_DISPLAY_NAMES.get(c, c) for c in self._columns]
        self._table.setColumnCount(len(headers))
        self._table.setHeaderLabels(headers)
        self._item_by_row = []
        self._pn_to_items.clear()

        # parent_stack: (level, QTreeWidgetItem | None)
        parent_stack: list[tuple[int, QTreeWidgetItem | None]] = [(-1, None)]

        weight_col_idx = self._columns.index("Weight") if "Weight" in self._columns else -1

        for row_idx, row_data in enumerate(self._rows):
            level = int(row_data.get("Level", 0))

            while len(parent_stack) > 1 and parent_stack[-1][0] >= level:
                parent_stack.pop()

            parent_item = parent_stack[-1][1]
            item = QTreeWidgetItem()
            item.setData(0, _ROW_IDX_ROLE, row_idx)

            if parent_item is None:
                self._table.addTopLevelItem(item)
            else:
                parent_item.addChild(item)

            parent_stack.append((level, item))
            self._item_by_row.append(item)

            pn        = str(row_data.get("Part Number", ""))
            not_found = bool(row_data.get("_not_found"))
            unreadable = bool(row_data.get("_unreadable"))
            spa_failed = bool(row_data.get("_spa_failed"))
            node_type  = str(row_data.get("Type", ""))
            row_locked = unreadable or not_found

            # 构建零件编号→树形项映射（用于联动更新）
            if pn:
                self._pn_to_items.setdefault(pn, []).append(item)

            # 填充各列
            for col_idx, col_name in enumerate(self._columns):
                if col_name == "#":
                    item.setText(col_idx, str(row_idx + 1))
                elif col_name == "Level":
                    item.setText(col_idx, str(row_data.get("Level", 0)))
                elif col_name == "Filename":
                    fp = str(row_data.get("_filepath", ""))
                    fn = str(row_data.get("Filename", ""))
                    value = Path(fp).name if fp else fn
                    item.setText(col_idx, value)
                    if fp:
                        item.setToolTip(col_idx, fp)
                elif col_name in ("Weight", "CogX", "CogY", "CogZ",
                                  "Ixx", "Iyy", "Izz", "Ixy", "Ixz", "Iyz"):
                    raw = row_data.get(col_name)
                    if raw is None or node_type not in ("零件",):
                        item.setText(col_idx, "" if node_type in ("产品", "部件") else "—")
                    else:
                        item.setText(col_idx, _fmt(raw))
                else:
                    item.setText(col_idx, str(row_data.get(col_name, "")))

            # 可编辑性：仅零件行且未锁定的 Weight 列可编辑
            if node_type == "零件" and not row_locked:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                item.setData(0, _ITEM_LOCKED_ROLE, False)
            else:
                item.setData(0, _ITEM_LOCKED_ROLE, True)

            # 行着色
            if row_locked:
                grey = QColor(160, 160, 160)
                bg   = QColor(255, 205, 205) if not_found else QColor(245, 245, 245)
                for ci in range(len(self._columns)):
                    item.setForeground(ci, grey)
                    item.setBackground(ci, bg)
            elif spa_failed and node_type == "零件":
                # SPA 测量失败：橙色背景
                bg = QColor(255, 210, 160)
                for ci in range(len(self._columns)):
                    item.setBackground(ci, bg)
            elif node_type in ("产品", "部件"):
                # 组/子装配体：淡灰色背景
                bg = QColor(240, 242, 245)
                for ci in range(len(self._columns)):
                    item.setBackground(ci, bg)

        self._table.expandAll()
        self._table.blockSignals(False)
        self._is_updating = False

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

        # 检查锁定标志
        if item.data(0, _ITEM_LOCKED_ROLE):
            return

        row_data = self._rows[row_idx]
        if row_data.get("Type") != "零件":
            return

        new_text = item.text(col_idx).strip()
        try:
            new_weight = float(new_text)
        except (ValueError, TypeError):
            # 恢复原值
            self._is_updating = True
            old_val = row_data.get("Weight")
            item.setText(col_idx, _fmt(old_val))
            self._is_updating = False
            return

        if new_weight < 0.0:
            QMessageBox.warning(
                self, "重量不合法",
                "重量不能为负数，请输入大于或等于 0 的值。",
            )
            self._is_updating = True
            item.setText(col_idx, _fmt(row_data.get("Weight")))
            self._is_updating = False
            return

        # 计算比例系数
        old_weight = row_data.get("Weight")
        try:
            old_float = float(old_weight) if old_weight is not None else 0.0
        except (ValueError, TypeError):
            old_float = 0.0

        scale = (new_weight / old_float) if old_float > 0.0 else 1.0

        # 更新本行的原始数据
        row_data["Weight"] = new_weight

        # 同步 _mass_props 中的 weight（仅覆盖本行用于 rollup）
        mp = row_data.get("_mass_props")
        if mp:
            mp["weight"] = new_weight
            if scale != 1.0 and old_float > 0.0:
                orig_inertia = mp.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)])
                mp["inertia"] = [[orig_inertia[r][c] * scale for c in range(3)]
                                  for r in range(3)]

        # 同步惯量列
        inertia = mp.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)]) if mp else [[0.0, 0.0, 0.0] for _ in range(3)]
        inertia_map = {
            "Ixx": inertia[0][0], "Iyy": inertia[1][1], "Izz": inertia[2][2],
            "Ixy": inertia[0][1], "Ixz": inertia[0][2], "Iyz": inertia[1][2],
        }
        self._is_updating = True
        for ic_name, ic_val in inertia_map.items():
            if ic_name in self._columns:
                ic_idx = self._columns.index(ic_name)
                new_val = (ic_val * scale) if (scale != 1.0 and old_float > 0.0) else ic_val
                row_data[ic_name] = new_val
                item.setText(ic_idx, _fmt(new_val))

        # 联动同型号零件（相同 Part Number 的其他行）
        pn = str(row_data.get("Part Number", ""))
        if pn and pn in self._pn_to_items:
            for other_item in self._pn_to_items[pn]:
                other_row_idx = other_item.data(0, _ROW_IDX_ROLE)
                if other_row_idx is None or other_row_idx == row_idx:
                    continue
                other_row = self._rows[other_row_idx]
                if other_row.get("Type") != "零件":
                    continue

                # 更新 Weight
                other_row["Weight"] = new_weight
                if "Weight" in self._columns:
                    w_idx = self._columns.index("Weight")
                    other_item.setText(w_idx, _fmt(new_weight))

                # 更新 _mass_props 及惯量
                other_mp = other_row.get("_mass_props")
                if other_mp:
                    other_old_w = other_mp.get("weight", 0.0) or 0.0
                    other_scale = (new_weight / other_old_w) if other_old_w > 0.0 else 1.0
                    other_mp["weight"] = new_weight
                    if other_scale != 1.0 and other_old_w > 0.0:
                        orig_i = other_mp.get("inertia", [[0.0]*3]*3)
                        other_mp["inertia"] = [
                            [orig_i[r][c] * other_scale for c in range(3)]
                            for r in range(3)
                        ]
                    other_inertia = other_mp.get("inertia", [[0.0]*3]*3)
                    other_imap = {
                        "Ixx": other_inertia[0][0], "Iyy": other_inertia[1][1],
                        "Izz": other_inertia[2][2], "Ixy": other_inertia[0][1],
                        "Ixz": other_inertia[0][2], "Iyz": other_inertia[1][2],
                    }
                    for ic_name, ic_val in other_imap.items():
                        other_row[ic_name] = ic_val
                        if ic_name in self._columns:
                            ic_idx = self._columns.index(ic_name)
                            other_item.setText(ic_idx, _fmt(ic_val))

        self._is_updating = False

        # 重置汇总结果（重量已变，需重新计算）
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
        self._lbl_weight.setText(f"{result.get('total_weight', 0.0):.4f} kg")
        cog = result.get("cog", [0.0, 0.0, 0.0])
        self._lbl_cx.setText(f"{cog[0]:.4f} mm")
        self._lbl_cy.setText(f"{cog[1]:.4f} mm")
        self._lbl_cz.setText(f"{cog[2]:.4f} mm")
        I = result.get("inertia", [[0.0]*3]*3)
        self._lbl_ixx.setText(f"{I[0][0]:.4f}")
        self._lbl_iyy.setText(f"{I[1][1]:.4f}")
        self._lbl_izz.setText(f"{I[2][2]:.4f}")
        self._lbl_ixy.setText(f"{I[0][1]:.4f}")
        self._lbl_ixz.setText(f"{I[0][2]:.4f}")
        self._lbl_iyz.setText(f"{I[1][2]:.4f}")

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

        # 导出列（排除内部列 "#" 和矩阵缓存列）
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

        # 表头
        for ci, col_name in enumerate(export_cols, start=1):
            cell       = ws.cell(row=1, column=ci,
                                 value=MASS_PROPS_COLUMN_DISPLAY_NAMES.get(col_name, col_name))
            cell.font  = Font(bold=True)
            cell.fill  = header_fill
            cell.border = thin_border

        # 数据行
        for ri, row_data in enumerate(self._rows, start=2):
            for ci, col_name in enumerate(export_cols, start=1):
                raw = row_data.get(col_name)
                if raw is None:
                    value = ""
                elif col_name in ("Weight", "CogX", "CogY", "CogZ",
                                   "Ixx", "Iyy", "Izz", "Ixy", "Ixz", "Iyz"):
                    try:
                        value = float(raw)
                    except (TypeError, ValueError):
                        value = ""
                else:
                    value = raw
                cell = ws.cell(row=ri, column=ci, value=value)
                cell.border = thin_border
                if col_name in ("Level",):
                    cell.alignment = center

        # 汇总行（如果已计算）
        if self._rollup_result:
            sep_row = len(self._rows) + 2
            # 空行
            summary_row = sep_row + 1
            cog     = self._rollup_result.get("cog", [0.0, 0.0, 0.0])
            I       = self._rollup_result.get("inertia", [[0.0]*3]*3)
            summary = {
                "Part Number":  "总计 (根产品)",
                "Weight":       self._rollup_result.get("total_weight", 0.0),
                "CogX":         cog[0],
                "CogY":         cog[1],
                "CogZ":         cog[2],
                "Ixx":          I[0][0],
                "Iyy":          I[1][1],
                "Izz":          I[2][2],
                "Ixy":          I[0][1],
                "Ixz":          I[0][2],
                "Iyz":          I[1][2],
            }
            summary_fill = PatternFill(fill_type="solid", fgColor="C6EFCE")
            for ci, col_name in enumerate(export_cols, start=1):
                val  = summary.get(col_name, "")
                cell = ws.cell(row=summary_row, column=ci, value=val)
                cell.font   = Font(bold=True)
                cell.fill   = summary_fill
                cell.border = thin_border

        # 冻结表头
        ws.freeze_panes = "A2"

        # 自适应列宽
        for ci, col_name in enumerate(export_cols, start=1):
            col_letter = ws.cell(row=1, column=ci).column_letter
            header     = MASS_PROPS_COLUMN_DISPLAY_NAMES.get(col_name, col_name)
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
