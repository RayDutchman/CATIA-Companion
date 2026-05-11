"""集中定义 UI 行状态颜色令牌。

所有颜色值均使用 ``"#rrggbb"`` 十六进制字符串，VS Code 内置调色板会在
悬停时自动激活，方便直接在此文件中调色。修改此文件即可同步更新整个应用
的配色，无需分别修改各 dialog 文件。

配色语义 & 触发条件速查（背景色均为柔和色调，避免长时间阅读疲劳）：

  颜色         常量名               触发 flag / 条件
  ──────────── ──────────────────── ─────────────────────────────────────────
  深橙（前景） MODIFIED_FG          字段已在 UI 中修改但尚未写回 CATIA
  中灰（前景） ROW_LOCKED_FG        行处于锁定状态（_not_found / _unreadable /
                                   _meas_failed，或 BOM 中 not_found /
                                   is_lightweight）时，叠加于行背景色之上
  红           ROW_NOT_FOUND_BG     row_data["_not_found"] = True
                                   （文件未被 CATIA 检索到）
  灰           ROW_LIGHTWEIGHT_BG   row_data["_unreadable"] = True
                                   （轻量化模式，无法读取属性）
  黄           ROW_UNSAVED_BG       row_data["_no_file"] = True
                                   （零件尚未保存到磁盘）
  琥珀         ROW_MEAS_FAILED_BG   row_data["_meas_failed"] = True
                                   （质量特性测量失败；仅 mass_props_dialog）
  蓝灰         ROW_PRODUCT_BG       row_data["Type"] in ("产品", "部件")
                                   （产品/部件汇总行；仅 mass_props_dialog）
  薰衣草（背景）EXCL_BG             row_data["_excluded"] = True
                                   （行已被手动排除，不参与汇总计算）
  深紫（前景） EXCL_FG              同上，与 EXCL_BG 同时生效
  蓝           MIRROR_BG            row_data["_is_mirror"] = True
                                   （对称件虚拟镜像行；仅 mass_props_dialog）
"""

from PySide6.QtGui import QColor

# ── BOM 编辑 dialog：已修改但未写回 CATIA 的字段视觉样式 ──────────────────────
# 触发：字段值与原始 CATIA 属性不同（_modified_keys 中存在该字段）
MODIFIED_FG          = QColor("#c05800")   # 深橙色：已修改字段的文字色
MODIFIED_COMBO_STYLE = "QComboBox { font-weight: bold; color: #c05800; }"  # 下拉框样式

# ── 共用行状态颜色（BOM 编辑 + 质量特性两个 dialog 均使用） ──────────────────
# ROW_LOCKED_FG 总是与下方某一背景色同时叠加，不单独出现
ROW_LOCKED_FG      = QColor("#909090")   # 中灰前景：行锁定时覆盖文字色
ROW_NOT_FOUND_BG   = QColor("#ffcccc")   # 红背景：_not_found=True（文件未被检索到）
ROW_LIGHTWEIGHT_BG = QColor("#ebebeb")   # 灰背景：_unreadable=True（轻量化模式）
ROW_UNSAVED_BG     = QColor("#fff9c4")   # 黄背景：_no_file=True（未保存到磁盘）

# ── 质量特性 dialog：额外行状态颜色 ───────────────────────────────────────────
ROW_MEAS_FAILED_BG = QColor("#ffe0b3")   # 琥珀背景：_meas_failed=True（测量失败）
ROW_PRODUCT_BG     = QColor("#dde8f5")   # 蓝灰背景：Type in ("产品","部件")（汇总行）
EXCL_BG            = QColor("#d8d4f0")   # 薰衣草背景：_excluded=True（已排除行）
EXCL_FG            = QColor("#5858a0")   # 深紫前景：_excluded=True，与 EXCL_BG 同时生效
MIRROR_BG          = QColor("#c8e4ff")   # 蓝背景：_is_mirror=True（对称件虚拟行）

# ── BOM 树控件：层级连接线颜色 ─────────────────────────────────────────────────
WIDGET_LINE_COLOR  = QColor("#a0aab4")   # 树形控件层级连接线颜色
