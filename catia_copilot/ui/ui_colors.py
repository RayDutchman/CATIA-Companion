"""集中定义 UI 行状态颜色令牌。

所有颜色值均使用 ``"#rrggbb"`` 十六进制字符串，VS Code 内置调色板会在
悬停时自动激活，方便直接在此文件中调色。修改此文件即可同步更新整个应用
的配色，无需分别修改各 dialog 文件。

配色语义对照（背景色均为柔和色调，避免长时间阅读疲劳）：
  ● 红     ROW_NOT_FOUND_BG   — 错误：文件未被检索到
  ● 琥珀   ROW_MEAS_FAILED_BG — 警告：质量特性测量失败
  ● 黄     ROW_UNSAVED_BG     — 提示：文件尚未保存到磁盘
  ● 灰     ROW_LIGHTWEIGHT_BG — 中性：轻量化模式（不可读）
  ● 蓝灰   ROW_PRODUCT_BG     — 结构：产品/部件汇总行
  ● 薰衣草  EXCL_BG            — 禁用：已排除行
  ● 蓝     MIRROR_BG          — 虚拟：对称件镜像行
"""

from PySide6.QtGui import QColor

# ── BOM 编辑 dialog：已修改但未写回 CATIA 的字段视觉样式 ──────────────────────
MODIFIED_FG          = QColor("#c05800")   # 深橙色：已修改字段的文字色
MODIFIED_COMBO_STYLE = "QComboBox { font-weight: bold; color: #c05800; }"  # 下拉框样式

# ── 共用行状态颜色（BOM 编辑 + 质量特性两个 dialog 均使用） ──────────────────
ROW_LOCKED_FG      = QColor("#909090")   # 锁定行文字色（中灰，与默认文字色对比明确）
ROW_NOT_FOUND_BG   = QColor("#ffcccc")   # 红：文件未被 CATIA 检索到
ROW_LIGHTWEIGHT_BG = QColor("#ebebeb")   # 灰：轻量化模式（不可读属性）
ROW_UNSAVED_BG     = QColor("#fff9c4")   # 黄：文件尚未保存到磁盘

# ── 质量特性 dialog：额外行状态颜色 ───────────────────────────────────────────
ROW_MEAS_FAILED_BG = QColor("#ffe0b3")   # 琥珀：质量特性测量失败（介于红和黄之间）
ROW_PRODUCT_BG     = QColor("#dde8f5")   # 蓝灰：产品/部件汇总行（结构层次感）
EXCL_BG            = QColor("#d8d4f0")   # 薰衣草：已排除行背景（色相与红/橙/黄完全分离）
EXCL_FG            = QColor("#5858a0")   # 深紫：已排除行文字色（与背景形成可读对比）
MIRROR_BG          = QColor("#c8e4ff")   # 蓝：对称件虚拟行（比 ROW_PRODUCT_BG 更饱和）

# ── BOM 树控件：层级连接线颜色 ─────────────────────────────────────────────────
WIDGET_LINE_COLOR  = QColor("#a0aab4")   # 树形控件层级连接线颜色
