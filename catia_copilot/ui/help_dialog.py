"""
帮助对话框 – 在可滚动的富文本窗口中显示用户文档。
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QTextBrowser, QPushButton, QHBoxLayout,
)
from PySide6.QtCore import Qt

from catia_copilot.constants import APP_NAME, APP_VERSION, APP_AUTHOR, APP_CONTACT


_HELP_HTML = f"""\
<h2>{APP_NAME} v{APP_VERSION} — 帮助文档</h2>

<h3>概述</h3>
<p>
{APP_NAME} 是一款面向工程团队的 CATIA V5 辅助工具，旨在简化日常操作，提升工作效率。
支持图纸与零件的批量导出、BOM 管理、宏脚本快捷运行，以及 CATIA 资源文件的一键部署。
</p>

<hr/>
<h3>运行环境要求</h3>
<ul>
  <li>操作系统：Windows 10 / 11</li>
  <li>已安装 CATIA V5 R28（文件导出等功能需要 CATIA 处于运行状态）</li>
</ul>

<hr/>
<h3>功能说明</h3>

<h4>一、导出</h4>
<table border="0" cellpadding="4">
<tr><td><b>CATDrawing → PDF</b></td>
    <td>批量将 CATDrawing 文件导出为 PDF。<br/>
    支持为输出文件添加自定义前缀（默认 DR_）。<br/>
    <i>注意：多页图纸请在 CATIA 中设置"将多页文档保存在单向量文件中"
    （工具 → 选项 → 常规 → 兼容性 → 图形格式 → 导出）。</i></td></tr>
<tr><td><b>CATPart / CATProduct → STP</b></td>
    <td>批量将 CATPart 或 CATProduct 文件导出为 STEP 格式。<br/>
    支持为输出文件添加自定义前缀（默认 MD_）。</td></tr>
<tr><td><b>从 CATProduct 导出 BOM</b></td>
    <td>从当前打开的 CATProduct 文件中提取完整 BOM 信息，并导出至 Excel (.xlsx)。<br/>
    可选择需要包含的列、自定义列、以及导出层级。</td></tr>
</table>

<h4>二、编辑</h4>
<table border="0" cellpadding="4">
<tr><td><b>BOM 属性补全</b></td>
    <td>加载当前 CATProduct 的 BOM 属性到表格中，可直接编辑零件编号、术语、
    定义、版本、来源等字段，以及自定义的用户属性（物料编码、物料名称、
    规格型号等）。修改完成后可一键写回 CATIA。<br/>
    <i>同一文件的属性修改会自动联动更新。</i></td></tr>
<tr><td><b>新建图纸</b></td>
    <td>根据 drawing_templates 文件夹中的 CATDrawing 模板，在 CATIA 中为当前
    活动的 CATPart 或 CATProduct 生成新图纸。<br/>
    <i>需在 CATIA 中打开目标零件/装配体，并将 *.CATDrawing 模板放入
    drawing_templates 文件夹。</i></td></tr>
<tr><td><b>刷新图纸</b></td>
    <td>将 CATIA 中当前活动 CATDrawing 图纸的参数（零件编号、术语、版本及
    自定义属性）与对应的零件/装配体同步刷新。<br/>
    <i>需在 CATIA 中同时打开目标图纸和对应零件/装配体文档。</i></td></tr>
</table>

<h4>三、工具</h4>
<table border="0" cellpadding="4">
<tr><td><b>复制字体文件到 CATIA 目录</b></td>
    <td>将 ChangFangSong.ttf 字体文件复制到 CATIA 的 TrueType 字体目录。
    程序会自动检测 CATIA 安装路径，也可手动选择。</td></tr>
<tr><td><b>复制 ISO.xml 到 CATIA 目录</b></td>
    <td>将 ISO.xml 标准文件复制到 CATIA 的 drafting 标准目录，
    用于设置制图标准。</td></tr>
<tr><td><b>刷写零件模板</b></td>
    <td>为选中的 CATPart 文件批量添加标准用户自定义属性
    （物料编码、物料名称、规格型号、物料来源、数据状态、
    存货类别、重量、备注）。</td></tr>
<tr><td><b>宏</b></td>
    <td>自动扫描 macros 文件夹中的 .catvbs / .catscript 文件，
    可直接在菜单中运行。支持打开宏文件夹和刷新宏列表。</td></tr>
<tr><td><b>紧固件快速装配</b></td>
    <td>使用 VBA 宏快速批量装配紧固件到产品孔位。<br/>
    支持自动对齐孔轴线、定位紧固件中心，以及装配后即时翻转方向。<br/>
    <i>需要在 CATIA 中打开紧固件 CATPart 文件和目标 CATProduct 文件。</i></td></tr>
<tr><td><b>托板螺母快速装配</b></td>
    <td>使用 VBA 宏快速批量装配托板螺母到产品孔位。<br/>
    通过选择托板螺母两个铆钉孔确定参考几何，再依次选择安装孔完成批量装配，支持即时翻转方向。<br/>
    <i>需要在 CATIA 中打开托板螺母 CATPart 文件和目标 CATProduct 文件。</i></td></tr>
</table>

<h4>四、视图</h4>
<table border="0" cellpadding="4">
<tr><td><b>显示 Log</b></td>
    <td>打开日志窗口，查看操作记录和错误信息，方便排查问题。</td></tr>
</table>

<hr/>
<h3>常见问题</h3>
<table border="0" cellpadding="4">
<tr><td><b>Q: 提示无法连接 CATIA？</b></td>
    <td>A: 请确认 CATIA V5 已启动并处于运行状态。程序通过 COM 自动化接口
    与 CATIA 通信，需要先打开 CATIA。</td></tr>
<tr><td><b>Q: 复制文件提示权限不足？</b></td>
    <td>A: CATIA 通常安装在 Program Files 目录，需要管理员权限才能写入。
    请右键以管理员身份运行本程序。</td></tr>
<tr><td><b>Q: BOM 导出的 Excel 打开后乱码？</b></td>
    <td>A: 导出使用 UTF-8 编码，请确保使用较新版本的 Excel 打开。</td></tr>
<tr><td><b>Q: 如何添加自定义宏？</b></td>
    <td>A: 点击菜单"宏 → 打开宏文件夹"，将 .catvbs 或 .catscript
    文件放入该文件夹，然后点击"刷新宏列表"即可。</td></tr>
</table>

<hr/>
<p style="color: #888;">
开发者：{APP_AUTHOR} | 联系方式：{APP_CONTACT}<br/>
仅供内部使用，请勿外传。
</p>
"""


class HelpDialog(QDialog):
    """Scrollable help dialog with rich-text documentation."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} — 帮助文档")
        self.resize(700, 560)
        self.setMinimumSize(480, 360)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml(_HELP_HTML)
        layout.addWidget(browser)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        layout.addLayout(btn_layout)
