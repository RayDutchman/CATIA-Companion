# CATIA Companion

> 一款面向工程团队的 CATIA V5 辅助工具，旨在简化日常操作、提升工作效率。

**版本：** 1.1.0 &nbsp;|&nbsp; **发布日期：** 2026-04-10 &nbsp;|&nbsp; **作者：** CHEN Weibo

---

## 功能一览

### 导出

| 功能 | 说明 |
|------|------|
| **CATDrawing → PDF** | 批量将 CATDrawing 文件导出为 PDF，支持自定义文件前缀 |
| **CATPart / CATProduct → STP** | 批量将 CATPart 或 CATProduct 文件导出为 STEP 格式 |
| **从 CATProduct 导出 BOM** | 从 CATProduct 中提取完整 BOM 信息并导出至 Excel (.xlsx) |

### 编辑

| 功能 | 说明 |
|------|------|
| **BOM 属性补全** | 在表格中编辑 BOM 属性（零件编号、术语、定义、版本、来源及自定义用户属性），一键写回 CATIA |
| **新建图纸** | 根据 `drawing_templates` 文件夹中的 CATDrawing 模板，在 CATIA 中为当前活动零件/装配体生成新图纸 |
| **刷新图纸** | 将当前活动 CATDrawing 图纸的参数与对应零件/装配体同步刷新（零件编号、术语、版本及自定义属性） |

### 工具

| 功能 | 说明 |
|------|------|
| **复制字体文件到 CATIA 目录** | 将 ChangFangSong.ttf 一键复制到 CATIA TrueType 字体目录 |
| **复制 ISO.xml 到 CATIA 目录** | 将 ISO.xml 标准文件一键复制到 CATIA drafting 标准目录 |
| **刷写零件模板** | 为 CATPart 批量添加标准用户自定义属性（物料编码、物料名称等） |
| **宏管理** | 自动扫描 macros 文件夹中的 `.catvbs` / `.catscript` 文件，可直接运行 |

### 其他

- **日志窗口** — 查看操作记录与错误信息
- **帮助文档** — 内置帮助文档，在菜单"帮助 → 文档"中查看

---

## 运行环境要求

- **操作系统：** Windows 10 / 11
- **Python：** 3.10 或更高版本
- **CATIA V5：** 文件导出等功能需要 CATIA 处于运行状态（通过 COM 自动化接口通信）

---

## 安装 / 开发环境搭建

```bash
# 1. 克隆仓库
git clone https://github.com/RayDutchman/CATIA-Companion.git
cd CATIA-Companion

# 2. 创建并激活虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 3. 安装依赖
pip install -r requirements.txt
```

### 运行

```bash
python main.py
```

---

## 打包为 Windows 可执行文件

```bash
# 前置依赖
pip install pyinstaller

# 打包
pyinstaller build.spec

# 输出目录
# dist\CATIA Companion\CATIA Companion.exe
```

ISO.xml、ChangFangSong.ttf 等资源文件会由 spec 配置自动复制到输出目录。

---

## 项目结构

```
CATIA-Companion/
├── main.py                          # 应用入口
├── catia_companion/
│   ├── constants.py                 # 常量与配置
│   ├── logging_setup.py             # 日志初始化
│   ├── utils.py                     # 工具函数
│   ├── catia/                       # CATIA COM 自动化逻辑
│   │   ├── conversion.py            #   图纸/零件导出
│   │   ├── template.py              #   零件模板刷写
│   │   ├── bom_collect.py           #   BOM 数据采集
│   │   ├── bom_export.py            #   BOM 导出 Excel
│   │   ├── bom_write.py             #   BOM 属性写回 CATIA
│   │   └── dependencies.py          #   依赖查找（开发中）
│   └── ui/                          # PySide6 界面
│       ├── main_window.py           #   主窗口
│       ├── convert_dialog.py        #   文件转换对话框
│       ├── export_bom_dialog.py     #   BOM 导出对话框
│       ├── bom_edit_dialog.py       #   BOM 编辑对话框
│       ├── find_deps_dialog.py      #   依赖查找对话框
│       ├── help_dialog.py           #   帮助文档对话框
│       ├── log_window.py            #   日志窗口
│       └── style.qss               #   QSS 样式表
├── build.spec                       # PyInstaller 打包配置
├── requirements.txt                 # Python 依赖
├── pyproject.toml                   # 项目元数据
├── ISO.xml                          # CATIA 制图标准文件
└── ChangFangSong.ttf                # 仿宋字体文件
```

---

## 依赖

| 包 | 用途 |
|------|------|
| [PySide6](https://pypi.org/project/PySide6/) | Qt 6 GUI 框架 |
| [pycatia](https://pypi.org/project/pycatia/) | CATIA V5 COM 自动化 |
| [openpyxl](https://pypi.org/project/openpyxl/) | Excel 文件读写 |

---

## 联系方式

- **开发者：** CHEN Weibo
- **邮箱：** thucwb@gmail.com

> 仅供内部使用，请勿外传。
