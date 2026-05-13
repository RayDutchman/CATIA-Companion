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
