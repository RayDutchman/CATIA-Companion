"""
例程：测试 SPAWorkbench.GetMeasurable(ref) 的全部常用测量接口
=============================================================

官方文档参考（方法1）
--------------------
- Measurable 接口完整参考（CATIA V5 Automation API）：
  https://catiadoc.free.fr/online/SPAug_C2/CAADriaSPAMeasurable.htm

- SPAWorkbench（InertiaWorkbench）接口参考：
  https://catiadoc.free.fr/online/SPAug_C2/CAADriaSPAWorkbench.htm

- SPA 自动化 API 总目录：
  https://catiadoc.free.fr/online/SPAug_C2/SPAugC2js.htm

- CATIA V5 Automation 所有模块索引：
  https://catiadoc.free.fr/online/CAAScdBase/CAAScdBase_toc.htm

运行前提
--------
- 已安装 pywin32（``pip install pywin32``）
- CATIA V5 正在运行，已打开一个含实体几何的 CATPart 文档

win32com ByRef Sub 调用规则（后期绑定）
----------------------------------------
CATIA Measurable 的测量方法均为 VBA Sub 风格（ByRef 输出参数）。
在 Python win32com 后期绑定中：
  - ByRef 输出参数以**函数返回值**形式给出，不需要传入任何数组
  - 返回值可能是单个值、tuple 或 list，本例程统一转为 list 处理

正确写法：
    cog     = meas.GetCOGPosition()   # → (Gx, Gy, Gz)  [mm]
    inertia = meas.GetInertia()       # → (Ixx,Ixy,...,Izz)  [kg·m²]，9 元素

错误写法（会报 TypeError）：
    arr = [0.0] * 3
    meas.GetCOGPosition(arr)          # ✗

Measurable 常用成员速查
-----------------------
属性：
  Mass      kg        体积质量（需赋材料）
  Volume    mm³       体积
  Area      mm²       表面积
  Density   kg/m³     密度（需赋材料，否则报错）

方法（返回值单位）：
  GetCOGPosition()         → (Gx, Gy, Gz)               [mm]
  GetInertia()             → 9元素，3×3行主序惯量矩阵    [kg·m²]
  GetPrincipalInertia()    → (I1, I2, I3) 三个主惯量     [kg·m²]
  GetPrincipalAxes()       → 9元素，三主轴方向矩阵       [无量纲]
  GetPoint()               → (x, y, z)                   [mm]
  GetDirection()           → (dx, dy, dz)                [无量纲]
  GetAxis()                → (ox,oy,oz, dx,dy,dz)        [mm / 无量纲]
  GetPlane()               → (nx,ny,nz, ox,oy,oz)        [无量纲 / mm]
"""

import sys
import win32com.client


def _p(*args, **kwargs):
    """带 flush 的 print，确保每行在 Windows 控制台立即显示。"""
    print(*args, **kwargs)
    sys.stdout.flush()


def _to_list(result) -> list:
    """把 COM 返回值统一转为 list（兼容 tuple / list / 单值）。"""
    if result is None:
        return []
    if hasattr(result, "__iter__") and not isinstance(result, str):
        return list(result)
    return [result]


# ─────────────────────────────────────────────────────────────────────────────
# 测试直接属性
# ─────────────────────────────────────────────────────────────────────────────

def test_properties(meas) -> None:
    _p("\n── 直接属性 ─────────────────────────────────────────────────────────────")
    for attr, unit in [("Mass", "kg"), ("Volume", "mm³"), ("Area", "mm²"), ("Density", "kg/m³")]:
        try:
            val = getattr(meas, attr)
            _p(f"  {attr:<12} = {val}  [{unit}]")
        except Exception as exc:
            _p(f"  {attr:<12}   [不可用] {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# 测试 ByRef Sub 测量方法
# ─────────────────────────────────────────────────────────────────────────────

def test_cog(meas) -> None:
    _p("\n── GetCOGPosition()  →  (Gx, Gy, Gz) [mm] ─────────────────────────────")
    try:
        vals = _to_list(meas.GetCOGPosition())
        labels = ["Gx", "Gy", "Gz"]
        for k, v in enumerate(vals):
            lbl = labels[k] if k < len(labels) else f"[{k}]"
            _p(f"  {lbl} = {v} mm")
        if not vals:
            _p("  [警告] 返回值为空")
    except Exception as exc:
        _p(f"  [错误] {exc}")


def test_inertia(meas) -> None:
    _p("\n── GetInertia()  →  3×3 惯量矩阵 [kg·m²]（行主序）────────────────────")
    try:
        vals = _to_list(meas.GetInertia())
        labels = ["Ixx", "Ixy", "Ixz", "Iyx", "Iyy", "Iyz", "Izx", "Izy", "Izz"]
        for k, v in enumerate(vals):
            lbl = labels[k] if k < len(labels) else f"I[{k}]"
            _p(f"  {lbl:<5} = {v}  kg·m²")
        if not vals:
            _p("  [警告] 返回值为空")
    except Exception as exc:
        _p(f"  [错误] {exc}")


def test_principal_inertia(meas) -> None:
    _p("\n── GetPrincipalInertia()  →  (I1, I2, I3) [kg·m²] ─────────────────────")
    try:
        vals = _to_list(meas.GetPrincipalInertia())
        for k, v in enumerate(vals):
            _p(f"  I_principal[{k}] = {v}  kg·m²")
        if not vals:
            _p("  [警告] 返回值为空")
    except Exception as exc:
        _p(f"  [错误] {exc}")


def test_principal_axes(meas) -> None:
    _p("\n── GetPrincipalAxes()  →  三主轴方向 3×3 矩阵（行主序）────────────────")
    try:
        vals = _to_list(meas.GetPrincipalAxes())
        labels = [
            "Axis1.x", "Axis1.y", "Axis1.z",
            "Axis2.x", "Axis2.y", "Axis2.z",
            "Axis3.x", "Axis3.y", "Axis3.z",
        ]
        for k, v in enumerate(vals):
            lbl = labels[k] if k < len(labels) else f"axes[{k}]"
            _p(f"  {lbl:<10} = {v}")
        if not vals:
            _p("  [警告] 返回值为空")
    except Exception as exc:
        _p(f"  [错误] {exc}")


def test_point(meas) -> None:
    """对点特征（如工作点）调用 GetPoint()；对实体 Body 通常报错，属正常现象。"""
    _p("\n── GetPoint()  →  (x, y, z) [mm]（仅适用于点特征）────────────────────")
    try:
        vals = _to_list(meas.GetPoint())
        labels = ["x", "y", "z"]
        for k, v in enumerate(vals):
            lbl = labels[k] if k < len(labels) else f"[{k}]"
            _p(f"  {lbl} = {v} mm")
        if not vals:
            _p("  [警告] 返回值为空")
    except Exception as exc:
        _p(f"  [不适用于此几何类型] {exc}")


def test_direction(meas) -> None:
    """对方向特征调用 GetDirection()；对实体 Body 通常报错，属正常现象。"""
    _p("\n── GetDirection()  →  (dx, dy, dz)（仅适用于方向/线特征）─────────────")
    try:
        vals = _to_list(meas.GetDirection())
        labels = ["dx", "dy", "dz"]
        for k, v in enumerate(vals):
            lbl = labels[k] if k < len(labels) else f"[{k}]"
            _p(f"  {lbl} = {v}")
        if not vals:
            _p("  [警告] 返回值为空")
    except Exception as exc:
        _p(f"  [不适用于此几何类型] {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

def main():
    _p("=" * 64)
    _p("  CATIA Measurable 接口测试例程")
    _p("=" * 64)

    # ── 1. 连接 CATIA ─────────────────────────────────────────────────────────
    _p("\n[1/5] 连接 CATIA Application ...")
    catia = win32com.client.Dispatch("CATIA.Application")
    _p("      OK")

    # ── 2. 获取活动 PartDocument ───────────────────────────────────────────────
    _p("\n[2/5] 获取活动 PartDocument ...")
    try:
        doc = catia.ActiveDocument
    except Exception as exc:
        _p(f"[错误] 无活动文档：{exc}")
        return
    try:
        part = doc.Part
    except Exception as exc:
        _p(f"[错误] 活动文档不是 CATPart：{exc}")
        return
    _p(f"      文档：{doc.Name}")

    # ── 3. 获取 SPAWorkbench ──────────────────────────────────────────────────
    _p("\n[3/5] 获取 SPAWorkbench ...")
    try:
        spa = doc.GetWorkbench("SPAWorkbench")
        _p("      OK")
    except Exception as exc:
        _p(f"[错误] 无法获取 SPAWorkbench：{exc}")
        return

    # ── 4. 创建对第一个 Body 的引用 ───────────────────────────────────────────
    _p("\n[4/5] 创建 Body.1 引用 ...")
    try:
        body = part.Bodies.Item(1)
        ref  = part.CreateReferenceFromObject(body)
        _p(f"      Body 名称：{body.Name}")
    except Exception as exc:
        _p(f"[错误] 无法创建引用：{exc}")
        return

    # ── 5. 获取 Measurable 对象 ───────────────────────────────────────────────
    _p("\n[5/5] 调用 spa.GetMeasurable(ref) ...")
    try:
        meas = spa.GetMeasurable(ref)
        _p("      OK  →  Measurable 对象已获取")
    except Exception as exc:
        _p(f"[错误] GetMeasurable 失败：{exc}")
        return

    # ── 6. 测量 ───────────────────────────────────────────────────────────────
    _p("\n" + "=" * 64)
    _p("  开始测量（Body.1）")
    _p("=" * 64)

    test_properties(meas)
    test_cog(meas)
    test_inertia(meas)
    test_principal_inertia(meas)
    test_principal_axes(meas)
    test_point(meas)
    test_direction(meas)

    _p("\n" + "=" * 64)
    _p("  测试完成")
    _p("=" * 64)


if __name__ == "__main__":
    main()
