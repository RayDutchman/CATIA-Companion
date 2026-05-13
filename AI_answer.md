# 关于 `dispatch.Version` 抛出异常的问题

## 问题

> 以前尝试过 `dispatch.Version`，确实会抛出 exception，所以本 session 所做的工作还有用吗？

## 结论：有用，而且修复方向完全正确

`dispatch.Version` 抛出异常，**正是**本 session 修复所针对的核心问题。修复逻辑分两层：

---

## 第一层：`.Version` 异常时回退到 `.Name` 检查

**旧代码（有 bug）：**
```python
def _is_catia_v5_dispatch(dispatch):
    try:
        return "V5" in str(dispatch.Version)
    except Exception:
        return False   # ← 抛异常时直接返回 False，误判为"不是 V5"
```

**新代码（修复后）：**
```python
def _is_catia_v5_dispatch(dispatch):
    try:
        return "V5" in str(dispatch.Version)
    except Exception:
        try:
            return str(dispatch.Name).upper() in ("CNEXT", "CATIA")  # ← 回退检查
        except Exception:
            return False
```

CATIA V5 的进程名是 `CNEXT`，`dispatch.Name` 在 V5 环境下**不会**抛异常，可以可靠地区分 V5 与 3DE。

---

## 第二层：用 `dynamic.Dispatch` 绕过 gen_py 缓存

`.Version` 抛异常的**根本原因**是 gen_py 早绑定缓存污染。当 3DEXPERIENCE 安装后，它的 typelib 写入了 gen_py 缓存，导致 win32com 用 3DE 的接口定义去调用 V5 对象，`.Version` 访问失败。

修复方法：所有 ROT 枚举和 `GetActiveObject` 调用改用 `win32com.client.dynamic.Dispatch`（晚绑定），完全绕过 gen_py 缓存。这样 `.Version` 有机会正常工作；即使仍然失败，第一层 `.Name` 回退也能兜底。

---

## 完整调用链对比

| 场景 | 旧代码结果 | 新代码结果 |
|---|---|---|
| 纯 V5 环境，gen_py 干净 | `.Version` 成功 → True ✅ | `.Version` 成功 → True ✅ |
| 纯 V5 环境，gen_py 被 3DE 污染 | `.Version` 抛异常 → **False ❌**（误判！） | `.Version` 抛异常 → `.Name`="CNEXT" → **True ✅** |
| 纯 V5 + dynamic.Dispatch | `.Version` 抛异常 → **False ❌** | `.Version` 可能成功 → True ✅；失败则回退 `.Name` ✅ |
| 3DE 环境 | `.Version` 返回非"V5"字符串 → False ✅ | 同左 ✅ |

---

## 总结

- `dispatch.Version` 抛异常不是偶发问题，而是"V5 + 3DE 共存"场景下的**必然现象**。
- 旧代码把这个异常当作"不是 V5"处理，导致状态栏显示"CATIA 未连接"。
- 本 session 的修复**专门针对这个异常场景**，通过 `.Name` 回退 + `dynamic.Dispatch` 双重保护，确保 V5 能被正确识别。
- **本 session 所做的工作完全有用，且是必要的修复。**
