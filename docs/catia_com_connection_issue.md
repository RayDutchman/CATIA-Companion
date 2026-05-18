# CATIA V5 COM 连接失败问题归纳

## 环境

- CATIA V5 R33，以**普通用户**身份运行
- Python / CATIA-Copilot，以**普通用户**身份运行
- Windows 11，pywin32

---

## 症状

程序启动后状态栏始终显示"未连接"或"连接断开"，点击任何需要 CATIA 的功能时，程序会**新建一个 CATIA 实例**，而不是连接到已经运行的那个。

控制台日志中可见：

```
GetActiveObject 失败: (-2147221005, '无效的类字符串', None, None)
```

---

## 根本原因

### 原因 1：ProgID `CATIA.Application` 不存在于注册表（直接原因）

错误码 `-2147221005` 对应 `CO_E_CLASSSTRING`（`0x800401F3`），含义是：

> Windows 无法在 `HKCR`（`HKEY_CLASSES_ROOT`）中找到名为 `CATIA.Application` 的 ProgID。

`GetActiveObject` 的工作流程是：**ProgID → CLSID → 在 ROT 中查找实例**。
第一步（ProgID 解析）就失败了，后续步骤完全没有执行，与 CATIA 是否正在运行无关。

CATIA V5 R33 安装后未写入 `HKCR\CATIA.Application` 条目（早期版本会写入，R33 改变了注册方式），导致原始代码的唯一连接路径失效。

### 原因 2：两套独立的 COM 连接实现，修复未同步（深层原因）

`utils.py`（负责状态检测）和 `connection.py`（负责业务操作）各自维护了一份 ROT 枚举逻辑，互不调用。

后来在 `utils.py` 中修复了 ROT 枚举（加了 `QueryInterface(IID_IDispatch)`），状态栏恢复显示"已连接"，但 `connection.py` 中的那份从未同步修复——业务操作仍然失败，程序走到 `_launch_catia_v5()` 分支，启动了新实例。

### 原因 3：`uac_admin=True` 会使问题更严重

`build.spec` 中原本设置了 `uac_admin=True`，打包后的 exe 会以管理员身份启动。

Windows UAC 会隔离不同完整性级别进程的 ROT：**管理员进程无法看到普通用户进程注册的 COM 对象**。CATIA 以普通用户运行，程序以管理员运行，则 ROT 对程序完全不可见，三阶段连接策略全部失败。

---

## 修复方案

### 1. `utils.py`：新增 `get_catia_v5_com_dispatch()`

封装完整的**三阶段连接策略**，作为全局唯一的 COM 连接入口：

| 阶段 | 方式 | 解决的问题 |
|------|------|-----------|
| 1 | 遍历注册表中所有 `CATIA*` ProgID + 经典备选（`CATIA.Application`、`CNEXT.Application`） | 兼容 ProgID 存在的场景 |
| 2 | 已知 CLSID `{87FD6F40-E252-11D5-8040-0001B5FA1031}` 直连 | 绕过 ProgID 缺失（`CO_E_CLASSSTRING`）问题 |
| 3 | ROT 枚举 + `QueryInterface(IID_IDispatch)` | 解决 `GetObject` 返回 `PyIUnknown` 无法直接 `Dispatch` 的问题 |

### 2. `connection.py`：删除重复实现，统一委托

删除 `connection.py` 中独立的 `_find_catia_v5_in_rot()`、`_is_catia_v5_dispatch()` 等函数，`_get_v5_com_object()` 直接调用 `utils.get_catia_v5_com_dispatch()`，确保单一真相来源。

### 3. `build.spec`：将 `uac_admin` 改为 `False`

```python
# 修改前
uac_admin=True

# 修改后
uac_admin=False  # 以普通用户权限运行，与 CATIA V5 保持相同权限级别
```

---

## 关键知识点

### `CO_E_CLASSSTRING` vs `E_ACCESSDENIED`

| 错误码 | 十六进制 | 含义 | 常见场景 |
|--------|----------|------|---------|
| `-2147221005` | `0x800401F3` | `CO_E_CLASSSTRING`：ProgID 不存在 | CATIA V5 R33 未注册 ProgID |
| `-2147024891` | `0x80070005` | `E_ACCESSDENIED`：访问被拒绝 | 本进程权限低于 CATIA 进程（如 CATIA 管理员、本程序普通用户） |
| `-2147221021` | `0x800401E3` | `MK_E_UNAVAILABLE`：对象未运行 | ProgID/CLSID 存在，但 CATIA 未启动 |

### ROT 枚举的 `QueryInterface` 陷阱

`rot.GetObject(moniker)` 返回的是 `PyIUnknown`，直接传给 `win32com.client.Dispatch()` 会失败。
必须先调用 `obj.QueryInterface(pythoncom.IID_IDispatch)` 取得 `IDispatch` 接口，再包装：

```python
idisp = obj.QueryInterface(pythoncom.IID_IDispatch)
dispatch = win32com.client.dynamic.Dispatch(idisp)
```

### Windows UAC ROT 隔离规则

- 管理员进程注册的 COM 对象 → 普通用户进程**可见**（向下可见）
- 普通用户进程注册的 COM 对象 → 管理员进程**不可见**（向上隔离）

因此，若 CATIA 以普通用户运行，本程序也必须以普通用户运行，才能通过 ROT 找到 CATIA。

---

## 涉及文件

| 文件 | 变更内容 |
|------|---------|
| `catia_copilot/utils.py` | 新增 `get_catia_v5_com_dispatch()`；修复 `_find_catia_v5_in_rot()` 加 `QueryInterface`；新增三阶段策略相关辅助函数 |
| `catia_copilot/catia/connection.py` | 删除重复 COM 逻辑；`_get_v5_com_object()` 改为委托 `utils.get_catia_v5_com_dispatch()` |
| `build.spec` | `uac_admin=True` → `False` |
