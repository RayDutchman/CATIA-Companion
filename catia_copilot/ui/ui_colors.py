"""集中定义 UI 行状态颜色令牌。

所有颜色值均使用 ``"#rrggbb"`` 十六进制字符串，VS Code 内置调色板会在
悬停时自动激活，方便直接在此文件中调色。修改此文件即可同步更新整个应用
的配色，无需分别修改各 dialog 文件。
"""

from PySide6.QtGui import QColor

# ── BOM 编辑 dialog：已修改但未写回 CATIA 的字段视觉样式 ──────────────────────
MODIFIED_FG          = QColor("#b85c00")   # 深橙色：已修改字段的文字色
MODIFIED_COMBO_STYLE = "QComboBox { font-weight: bold; color: #b85c00; }"  # 下拉框样式

# ── 共用行状态颜色（BOM 编辑 + 质量特性两个 dialog 均使用） ──────────────────
ROW_LOCKED_FG      = QColor("#a0a0a0")   # 锁定行文字色（灰色）
ROW_NOT_FOUND_BG   = QColor("#ffcdcd")   # 文件未被 CATIA 检索到，行背景（浅红）
ROW_LIGHTWEIGHT_BG = QColor("#f5f5f5")   # 轻量化模式行背景（浅灰）
ROW_UNSAVED_BG     = QColor("#fff5b4")   # 未保存到磁盘行背景（浅黄）

# ── 质量特性 dialog：额外行状态颜色 ───────────────────────────────────────────
ROW_MEAS_FAILED_BG = QColor("#ffd2a0")   # 测量失败行背景（浅橙）
ROW_PRODUCT_BG     = QColor("#f0f2f5")   # 产品/部件汇总行背景（浅蓝灰）
EXCL_BG            = QColor("#d8d8e8")   # 排除行背景（浅灰紫，区别于红/橙/黄等异常色）
EXCL_FG            = QColor("#828296")   # 排除行文字色（灰紫）
MIRROR_BG          = QColor("#e6f0ff")   # 对称件虚拟行背景（浅蓝）

# ── BOM 树控件：层级连接线颜色 ─────────────────────────────────────────────────
WIDGET_LINE_COLOR  = QColor("#a0aab4")   # 树形控件层级连接线颜色
