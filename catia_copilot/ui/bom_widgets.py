"""BOM 树控件：自定义委托与带连接线的 QTreeWidget。"""

from PySide6.QtWidgets import QTreeWidget, QStyledItemDelegate
from PySide6.QtGui import QColor, QPen, QPainter
from PySide6.QtCore import Qt

from catia_copilot.constants import BOM_READONLY_COLUMNS

# 自定义 UserRole 用于 QTreeWidgetItem：标记行为锁定（不可读/未找到）
_ITEM_LOCKED_ROLE: int = Qt.ItemDataRole.UserRole + 1


class _BomTreeDelegate(QStyledItemDelegate):
    """BOM QTreeWidget 的逐列只读强制委托。

    QTreeWidgetItem 的 flags 是按行设置的；此委托对内部名称属于
    :data:`~catia_copilot.constants.BOM_READONLY_COLUMNS` 的列，
    以及被标记为锁定的行（文件未找到/不可读）从 :meth:`createEditor`
    返回 ``None``，从而阻止编辑。
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
    """绘制 Windows 注册表编辑器风格点状连接线的 QTreeWidget。

    Qt 的默认 Windows/Fusion 样式省略了连接父节点和子节点的垂直导向线。
    此子类重写 :meth:`drawBranches` 以绘制 1像素实线/1像素空白的点状虚线
    （基于绝对视口坐标，以确保垂直导向线在连续行之间保持相位一致）。
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
