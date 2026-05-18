# CATIA V5 COM 连接失败问题归纳

## 环境

- CATIA V5 R33（或 R28），以**普通用户**身份运行
- Python / CATIA-Copilot，以**普通用户**身份运行
- Windows 11，pywin32

---

## 症状

程序启动后状态栏始终显示"未连接"或"连接断开"，点击任何需要 CATIA 的功能时，可能程序会**新建一个 CATIA 实例**，而不是连接到已经运行的那个。

控制台日志中可见：

```
GetActiveObject 失败: (-2147221005, '无效的类字符串', None, None)
```

---

## 根本原因

### 原因 1：不同 CATIA 版本使用不同的 CLSID（直接原因）

**验证结果（2026-05-18）：**

| 版本 | ProgID | CLSID |
|------|--------|-------|
| **R33** | `CATIA.Application` ✅ 存在 | `{87fd6f40-e252-11d5-8040-0010b5fa1031}` |
| **R28（推测）** | `CATIA.Application` ✅ 存在 | `{87FD6F40-E252-11D5-8040-0001B5FA1031}` |
| **差异** | 相同（后装版本覆盖） | **最后8位不同**：`0010b5fa1031` vs `0001b5fa1031` |

**问题分析：**

错误码 `-2147221005` (`CO_E_CLASSSTRING`) 通常表示 ProgID 或 CLSID 无效。但实际验证显示：

1. ✅ R33 **完整注册了** `HKCR\CATIA.Application` ProgID
2. ✅ R33 的 CLSID `{...0010b5fa1031}` 注册完整，LocalServer32 指向 CNEXT.exe
3. ✅ 直接调用 `GetActiveObject('CATIA.Application')` **可以成功连接** R33

**真正的问题：**

- 旧代码中硬编码的 CLSID `{87FD6F40-E252-11D5-8040-0001B5FA1031}` 是 **R28（或更早版本）的 CLSID**
- 当 R28 和 R33 共存时，阶段2可以通过 R28 的 CLSID 连接到 R28
- 卸载 R28 后，阶段2失败（R33 用的是不同的 CLSID `0010b5fa1031`）
- 如果阶段1（ProgID）或阶段3（ROT枚举）也失败，则完全无法连接

**重要发现：**

`CATIA.Application` ProgID 在注册表中只能有一个实例，后安装的版本会覆盖前一个版本的指向。因此：
- R28 后装 → ProgID 指向 R28 的 CLSID
- R33 后装 → ProgID 指向 R33 的 CLSID

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
| 2 | 已知 CLSID 列表直连（R28: `{...0001B5FA1031}`, R33: `{...0010B5FA1031}`） | 兼容不同 CATIA 版本的 CLSID 差异 |
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

## 连接检测优化（2026-05-18）

### 4 态状态机

`check_catia_connection()` 返回以下 4 种状态：

| 状态 | 含义 |
|------|------|
| `"connected"` | COM 对象可获取，功能测试（`app.Name`）通过 |
| `"broken"` | CATIA 进程存在，但所有 COM 连接方式均失败（无 `E_ACCESSDENIED`） |
| `"access_denied"` | `GetActiveObject` 明确返回 `E_ACCESSDENIED (-2147024891)` |
| `"disconnected"` | 未检测到运行中的 CATIA 进程 |

### 性能优化

**进程检查前置**：第一步通过 `tasklist` 检查 `CNEXT.exe` 是否存在，不存在则立即返回 `"disconnected"`，跳过所有 COM 探测。

**快速路径 `_last_working_key`**：记录上次成功连接的 ProgID 或 CLSID，下次轮询优先用该 key 快速尝试。命中时完全跳过注册表扫描和全量 CLSID 探测，稳定连接状态下每 5 秒仅执行 `tasklist + 1 次 GetActiveObject + app.Name`。

**注册表缓存 `_cached_registry_progids`**：ProgID 由安装程序写入，运行期不变，永久缓存合理。只扫描一次注册表，日志只打印一次。

**`is_admin()` 缓存**：进程属性绝对不变，`_is_admin_cache` 在进程生命周期内只调用 Win32 API 一次。

---

## 文件复制权限提升（2026-05-18）

### 背景

复制字体/ISO 标准文件/crack 文件到 CATIA 安装目录（`Program Files\Dassault Systemes\B28\...`）需要管理员权限。将整个程序以管理员运行会导致 COM 连接失败（UAC ROT 隔离），不能作为解决方案。

### 方案：ShellExecuteExW + WaitForSingleObject

实现 `_run_copy_elevated(operations)` 方法：

1. 将所有文件复制操作写入临时 `.bat` 批处理文件（`copy /Y` 命令）
2. 通过 `ShellExecuteExW`（`runas` 动词 + `SEE_MASK_NOCLOSEPROCESS`）以管理员权限静默运行 `cmd.exe /c <bat>`
3. `WaitForSingleObject(hProcess, 60000)` 同步等待完成（最多 60 秒）
4. 关闭进程句柄，删除临时批处理文件
5. 调用方通过检查目标文件是否存在来验证结果

**优点**：主程序仍以普通用户运行（COM 连接正常），仅对文件复制操作临时提权；不需要整体重启程序。

**触发条件**：`shutil.copy2` 抛出 `PermissionError` 时，弹出询问对话框，用户确认后调用 `_run_copy_elevated()`。

---

## crack 版本子目录约定（2026-05-18）

`crack/` 目录下按 CATIA 版本分子目录：

```
crack/
├── R28/    ← CATIA V5 R28 的 crack 文件
├── R33/    ← CATIA V5 R33 的 crack 文件
└── （可根据需要添加 R29、R30 等）
```

**版本推断逻辑**（`_detect_crack_version_subdir`）：

- `detect_catia_root()` 返回形如 `C:\Program Files\Dassault Systemes\B33` 的路径
- 取路径末尾 `B33` → 匹配 `^B(\d+)$` → 对应子目录 `R33`
- 若版本专属子目录不存在，提示用户，并询问是否改用 `crack/` 根目录中的文件（向后兼容）

---

## 验证结果 (2026-05-18)

### 测试环境
- CATIA V5 R33 重新安装（纯R33，无R28）
- Windows Python 3.13 + pywin32
- CATIA 以普通用户身份运行

### COM 注册情况
```
HKCR\CATIA.Application
  └─ CLSID: {87fd6f40-e252-11d5-8040-0010b5fa1031}

HKCR\CLSID\{87fd6f40-e252-11d5-8040-0010b5fa1031}
  ├─ (Default): OLE CATIA.Application
  ├─ LocalServer32: "C:\...\CNEXT.exe" -env CATIA_P3.V5-6R2023.B33 ...
  ├─ ProgID: CATIA.Application.1
  └─ VersionIndependentProgID: CATIA.Application
```

### 连接测试结果
| 测试 | 方法 | 结果 |
|------|------|------|
| 1 | `GetActiveObject('CATIA.Application')` | ✅ **成功** |
| 2 | `GetActiveObject('{87fd6f40-e252-11d5-8040-0010b5fa1031}')` | ✅ **成功** |
| 3 | `GetActiveObject('{87FD6F40-E252-11D5-8040-0001B5FA1031}')` | ❌ 失败 (MK_E_UNAVAILABLE) |
| 4 | ROT 枚举 | ❌ 未找到 CATIA 对象* |

\* **注意**：ROT 枚举未找到对象，但 `GetActiveObject` 成功，说明 CATIA 通过其他机制注册了 COM 对象，而不是传统的 ROT。这是 CATIA 的正常行为。

### 结论
1. ✅ **R33 完整注册了 ProgID 和 CLSID**
2. ✅ **真正问题：不同版本使用不同 CLSID**

### 解决方案验证
- 阶段1（ProgID）：✅ 可以直接连接 R33
- 阶段2（CLSID 列表）：需要同时包含 `0001b5fa1031` 和 `0010b5fa1031` 才能兼容两个版本
- 阶段3（ROT）：对 CATIA 无效，但保留作为其他场景的兼容方案

---

## 关键知识点

### `CO_E_CLASSSTRING` vs `E_ACCESSDENIED`

| 错误码 | 十六进制 | 含义 | 常见场景 |
|--------|----------|------|---------|
| `-2147221005` | `0x800401F3` | `CO_E_CLASSSTRING`：ProgID/CLSID 无效或对象未在 ROT 注册 | 使用了错误版本的 CLSID，或 CATIA 未启动 |
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

**推论**：不能通过将整个程序以管理员运行来解决文件复制权限问题——这会导致 COM 连接失败。应使用 `_run_copy_elevated()` 仅对文件复制操作单独提权。

---

## 涉及文件

| 文件 | 变更内容 |
|------|---------|
| `catia_copilot/utils.py` | 新增 `get_catia_v5_com_dispatch()`；`check_catia_connection()` 重写（进程检查前置、`_last_working_key` 快速路径、注册表永久缓存）；`is_admin()` 进程级缓存；`_find_catia_progids_in_registry()` 永久缓存 |
| `catia_copilot/catia/connection.py` | 删除重复 COM 逻辑；`_get_v5_com_object()` 改为委托 `utils.get_catia_v5_com_dispatch()` |
| `catia_copilot/ui/main_window.py` | 新增 `_run_copy_elevated()`（ShellExecuteExW 提权复制）；新增 `_detect_crack_version_subdir()`；`_copy_file_to_catia` PermissionError 改为提权重试；`_crack()` 支持版本子目录 + 提权重试 |
| `catia_copilot/ui/help_dialog.py` | 删除 gen_py 相关 FAQ 条目；broken 状态说明改为"初始化/注册异常" |
| `build.spec` | `uac_admin=True` → `False` |
| `crack/R28/`、`crack/R33/` | 新建版本专属 crack 子目录（各含 `.gitkeep`） |


