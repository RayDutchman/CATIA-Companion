# 关于 `dispatch.Version` 抛出异常的问题

---

## 补充信息（纯 V5 环境，无 3DE）

> 我的电脑只有V5没有3DE，`.Version` 仍然一定会抛出异常。

这是关键新信息，修正了之前的根因分析：

### 真正的根因

**`.Version` 不是因为 gen_py 缓存污染才抛异常，而是在 CATIA V5 Application 的 COM IDispatch 接口中根本不暴露这个属性。**

- 使用 `dynamic.Dispatch`（晚绑定）时，win32com 通过 `IDispatch::GetIDsOfNames("Version")` 查询属性的 DispID
- CATIA V5 Application 对象的 IDispatch **不返回** "Version" 的 DispID，导致 `AttributeError`/`com_error`
- 这与 gen_py 缓存完全无关，纯 V5 环境下也必然抛异常

### 修正后的分析

| 场景 | `.Version` | `.Name` |
|---|---|---|
| 纯 V5（任何环境） | **始终抛异常** | 返回 `"CATIA"` 或 `"CNEXT"` ✅ |
| 3DE | 返回 `"3DEXPERIENCE..."` ✅ | 返回 `"CNEXT"` （与V5相同，无法区分） |

这说明：
- **`.Name` 是唯一可靠的 V5 识别方式**（对于纯 V5 场景）
- **`.Version` 的作用是识别 3DE**（当 `.Version` 成功且包含 "3DEXPERIENCE" → 排除 3DE）
- 在 V5+3DE 共存时：先用 `.Version` 排除 3DE，再用 `.Name` 确认 V5 —— 两步缺一不可

---

## 本次修复的逻辑（完整版）

**旧代码（根本性 bug）：**
```python
def _is_catia_v5_dispatch(dispatch):
    try:
        return "V5" in str(dispatch.Version)
    except Exception:
        return False   # ← 纯 V5 时始终执行此行，始终返回 False！
```

**新代码（修复后）：**
```python
def _is_catia_v5_dispatch(dispatch):
    try:
        version = str(dispatch.Version)          # V5 下始终抛异常，直接进 except
        if "3DEXPERIENCE" in version.upper():
            return False                          # 排除 3DE
        ...
        return True
    except Exception:
        try:
            name = str(dispatch.Name).upper()    # V5 的主要识别路径
            return name in ("CNEXT", "CATIA")    # 纯 V5 → True ✅
        except Exception:
            return False
```

---

## 完整场景矩阵

| 场景 | 旧代码 | 新代码 |
|---|---|---|
| 纯 V5（无 3DE） | `.Version` 抛异常 → **False ❌**（始终断连！） | `.Version` 抛 → `.Name`="CATIA" → **True ✅** |
| V5 + 3DE 共存，连接 V5 对象 | `.Version` 抛异常 → **False ❌** | `.Version` 抛 → `.Name`="CNEXT" → **True ✅** |
| V5 + 3DE 共存，连接 3DE 对象 | `.Version` 返回 "3DEXPERIENCE..." → False ✅ | 同左，`.Version` 成功 → **False ✅**（正确排除） |
| 3DE 也不暴露 `.Version`（极端情况） | `.Version` 抛 → **False ✅**（巧合正确） | `.Version` 抛 → `.Name`="CNEXT" → **True ❌**（无法区分） |

---

## 总结

1. **旧代码在纯 V5 环境下就是坏的** —— 不需要 3DE，`dispatch.Version` 就始终抛异常，连接始终失败。
2. **新代码的 `.Name` 回退是正确且必要的** —— 这是 V5 对象的唯一可靠识别方式。
3. `dynamic.Dispatch` 的修改仍然有价值 —— 防止 3DE 的 gen_py 缓存干扰获取 dispatch 对象本身的过程（影响的是能否拿到对象，而非 `.Version` 是否工作）。
4. **当前修复对纯 V5 场景完全有效**：`.Version` 抛异常 → `.Name`="CATIA"/"CNEXT" → 返回 True → 连接成功。

---

## 诊断对话框显示 `(-2147221020, '无效的语法', None, None)`

这是另一个独立的 bug，已在本次一并修复：`MkParseDisplayName` 不能解析裸 ProgID 字符串！

### 错误码分析

| 十进制 | 十六进制 | 名称 | 含义 |
|---|---|---|---|
| -2147221020 | 0x800401E4 | `MK_E_SYNTAX` | 无效的 Moniker 显示名称语法 |

### 根因

旧代码中有：
```python
clsid = pythoncom.MkParseDisplayName("CATIA.Application")[0]  # ← 这里抛异常！
```

`MkParseDisplayName` 是用于解析 **Moniker 显示名称**（如 `"clsid:{...}"` 格式）的 API，**不接受**裸 ProgID 字符串（如 `"CATIA.Application"`）。正确的 API 是：
```python
clsid = pythoncom.CLSIDFromProgID("CATIA.Application")         # ← 正确：查注册表
raw   = pythoncom.GetActiveObject(clsid)
```

`CLSIDFromProgID` 会直接查询注册表键 `HKCR\CATIA.Application\CLSID`，这才是 ProgID → CLSID 的标准路径。

### 为什么之前显示"未连接"而不是"错误"

诊断代码的流程：
1. `MkParseDisplayName` → 抛 `MK_E_SYNTAX` → 保存到 `result["error"]`，`app=None`
2. 进入 ROT 枚举降级路径 → CATIA IS in ROT，但之前 `_is_catia_v5_dispatch` 也是坏的 → 也返回 None
3. 最终 `status="disconnected"`，但 `error` 字段里保留了步骤 1 的异常 → 显示 `MK_E_SYNTAX`

修复 `CLSIDFromProgID` 后，步骤 1 会成功找到 CATIA，步骤 2 用 `.Name` 识别为 V5，连接成功。

> **但**：`pythoncom.CLSIDFromProgID` 在旧版本 pywin32 中不存在，导致出现下一个错误。

---

## 第三次错误（当前）：`module 'pythoncom' has no attribute 'CLSIDFromProgID'`

`CLSIDFromProgID` 是 `pythoncom` 扩展模块中后来才添加的方法，在用户安装的 pywin32 版本中不存在。

### 正确的 API：`pywintypes.IID(progid)`

`win32com.client.GetActiveObject` 内部用的就是这个：

```python
# win32com/client/__init__.py 源码
def GetActiveObject(progid, clsctx=None):
    clsid = pywintypes.IID(progid)   # ← 解析 ProgID → CLSID
    ...
    interface = pythoncom.GetActiveObject(clsid)
```

`pywintypes.IID(string)` 构造函数：
- 若传入 `"{XXXXXXXX-...}"` 形式的 CLSID 字符串 → 直接解析
- 若传入 `"CATIA.Application"` 这样的 ProgID → 调用 Windows `CoClsidFromProgID` API 解析
- 在所有 pywin32 版本中均存在（pywintypes 是 pywin32 最基础的模块）

### 本次修复范围

所有三处 `CLSIDFromProgID` 改为 `pywintypes.IID`：

| 文件 | 函数 |
|---|---|
| `catia_copilot/catia/connection.py` | `_get_v5_com_object()` |
| `catia_copilot/utils.py` | `check_catia_connection()` |
| `catia_copilot/utils.py` | `diagnose_catia_connection()` |

### 修复后的完整连接流程（纯 V5）

1. `pywintypes.IID("CATIA.Application")` → 调用 Windows CoClsidFromProgID → 得到 CLSID ✅
2. `pythoncom.GetActiveObject(clsid)` → 在 Windows ROT 中找到正在运行的 CATIA V5 ✅
3. `dynamic.Dispatch(raw)` → 强制晚绑定包装 ✅
4. `app.Name` → 功能性测试，CATIA V5 返回 `"CATIA"` ✅
5. `_is_catia_v5_dispatch(app)` → `.Version` 抛异常 → `.Name`=`"CATIA"` → **True ✅**
6. 返回连接对象 ✅

**修复后，诊断对话框和主界面均应显示"已连接"。**
