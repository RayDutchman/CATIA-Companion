"""
例程：枚举 SPA.GetMeasurable(Ref) 返回的 Measurable 对象的所有 COM 成员
==========================================================================

本文件演示 **方法2**——通过 win32com 后期绑定在运行时枚举
``SPAWorkbench.GetMeasurable(ref)`` 所返回的 ``Measurable`` COM 对象的
全部属性与方法（使用 ITypeInfo.GetFuncDesc / GetVarDesc 遍历 COM TypeLib）。

官方文档参考（方法1）
--------------------
- Measurable 接口参考（CATIA V5 Automation API）：
  https://catiadoc.free.fr/online/SPAug_C2/SPAugC2js.htm

- SPAWorkbench（InertiaWorkbench）接口参考：
  https://catiadoc.free.fr/online/SPAug_C2/CAADriaSPAWorkbench.htm

- CATIA V5 Automation 文档主索引：
  https://catiadoc.free.fr/online/CAAScdBase/CAAScdBase_toc.htm

- Dassault 官方 3DS Developer 网页（需登录）：
  https://www.3ds.com/support/documentation/

运行前提
--------
- 已安装 pycatia（``pip install pycatia``）与 pywin32（``pip install pywin32``）
- CATIA V5 正在运行，且已打开一个含实体几何的 CATPart 文档

枚举原理
--------
win32com 后期绑定（late-binding）通过 ``IDispatch::GetTypeInfo`` 获取
COM 对象的类型信息（TypeInfo），再遍历 ``FuncDesc`` 和 ``VarDesc``
可以列出接口暴露的全部函数与属性，无需提前调用 makepy。

常见 Measurable 成员汇总（V5R20+）
-----------------------------------
属性（Property）：
  Mass        → float，质量，单位 kg
  Volume      → float，体积，单位 mm³
  Area        → float，表面积，单位 mm²
  Density     → float，密度，单位 kg/m³（仅实体可用）

方法（Method / Sub）：
  GetCOGPosition(oCoords)
      ByRef Sub，oCoords 需预声明为 Array(2)，
      返回后 oCoords(0/1/2) = Gx/Gy/Gz，单位 mm

  GetInertia(oInertia)
      ByRef Sub，oInertia 需预声明为 Array(8)，
      返回 3×3 行主序惯量矩阵，单位 kg·m²
      索引：0=Ixx,1=Ixy,2=Ixz,3=Iyx,4=Iyy,5=Iyz,6=Izx,7=Izy,8=Izz

  GetPrincipalInertia(oInertia)
      ByRef Sub，oInertia = Array(2)，返回三个主惯量，单位 kg·m²

  GetPrincipalAxes(oAxes)
      ByRef Sub，oAxes = Array(8)，返回三个主轴方向 3×3 矩阵

  GetLengthBetween(iRef, oLength)
      与另一引用的最短距离，单位 mm

  GetAngleBetween(iRef, oAngle)
      与另一引用的夹角，单位 rad

  GetMinimumDistancePoints(iRef, oPoint1, oPoint2)
      两个引用之间最近点坐标，各含 3 个 float，单位 mm

  GetDirection(oDirection)
      方向向量 Array(2)

  GetPoint(oPoint)
      点坐标 Array(2)，单位 mm

  GetAxis(oAxis)
      轴线（点+方向）Array(5)

  GetPlane(oPlane)
      平面（法向+原点）Array(5)

win32com ByRef Sub 调用注意事项
--------------------------------
CATIA COM 方法中大量使用 ByRef 输出参数（VBA Sub 惯例）。
在 Python win32com 后期绑定中，ByRef 参数作为 **Python 函数的返回值**
返回，而不是修改传入的参数本身。

正确写法（Python）：
    import win32com.client
    catia = win32com.client.Dispatch("CATIA.Application")
    spa   = catia.ActiveDocument.GetWorkbench("SPAWorkbench")
    ref   = catia.ActiveDocument.Part.CreateReferenceFromObject(
                catia.ActiveDocument.Part.Bodies.Item(1))
    meas  = spa.GetMeasurable(ref)

    # Sub 风格（无返回值）→ win32com 把 ByRef 参数当返回值返回
    result = meas.GetCOGPosition()        # 返回 (Gx, Gy, Gz) tuple 或 list
    gx, gy, gz = result                   # 单位 mm

    inertia = meas.GetInertia()           # 返回 9 元素 tuple，单位 kg·m²
    # inertia[0]=Ixx, [1]=Ixy, [2]=Ixz, ...

错误写法（会报 TypeError）：
    arr = [0.0] * 3
    meas.GetCOGPosition(arr)   # ✗ 后期绑定不支持传入 ByRef 数组
"""

import win32com.client


# ─────────────────────────────────────────────────────────────────────────────
# 1. 通过 ITypeInfo 枚举 COM 对象的全部成员
# ─────────────────────────────────────────────────────────────────────────────

# INVOKEKIND 常量（来自 OAIdl.h）
_INVOKE_FUNC     = 1
_INVOKE_PROPGET  = 2
_INVOKE_PROPPUT  = 4
_INVOKE_PROPPUTREF = 8

# TYPEKIND 枚举中我们只关心 DISPATCH / INTERFACE
_TYPEKIND_DISPATCH  = 4

# VARTYPE 常量（简化映射，仅列常用值）
_VT_MAP = {
    0: "VT_EMPTY", 1: "VT_NULL", 2: "VT_I2", 3: "VT_I4",
    4: "VT_R4",    5: "VT_R8",   8: "VT_BSTR", 9: "VT_DISPATCH",
    10: "VT_ERROR", 11: "VT_BOOL", 12: "VT_VARIANT", 13: "VT_UNKNOWN",
    16: "VT_I1",   17: "VT_UI1",  18: "VT_UI2", 19: "VT_UI4",
    20: "VT_I8",   21: "VT_UI8",  22: "VT_INT", 23: "VT_UINT",
    24: "VT_VOID", 25: "VT_HRESULT", 26: "VT_PTR", 27: "VT_SAFEARRAY",
    28: "VT_CARRAY", 29: "VT_USERDEFINED", 30: "VT_LPSTR", 31: "VT_LPWSTR",
    64: "VT_FILETIME",
}
_VT_BYREF = 0x4000
_VT_ARRAY = 0x2000


def _vt_name(vt: int) -> str:
    """把 VARTYPE 整数转成可读字符串，含 BYREF/ARRAY 修饰。"""
    prefix = ""
    if vt & _VT_BYREF:
        prefix += "ByRef "
        vt = vt & ~_VT_BYREF
    if vt & _VT_ARRAY:
        prefix += "Array "
        vt = vt & ~_VT_ARRAY
    return prefix + _VT_MAP.get(vt, f"VT_{vt:#06x}")


def enumerate_com_members(com_obj) -> list[dict]:
    """
    通过 ITypeInfo 枚举 COM 对象暴露的全部函数与属性。

    Parameters
    ----------
    com_obj : win32com Dispatch 对象

    Returns
    -------
    list[dict]  每个成员字典，键：
                  name, kind, ret_type, params, memid
    """
    members = []
    try:
        ti = com_obj._oleobj_.GetTypeInfo(0)
    except Exception as exc:
        print(f"[错误] 无法获取 TypeInfo：{exc}")
        return members

    try:
        ta = ti.GetTypeAttr()
    except Exception as exc:
        print(f"[错误] 无法获取 TypeAttr：{exc}")
        return members

    # ── 遍历函数 ──────────────────────────────────────────────────────────────
    for i in range(ta.cFuncs):
        try:
            fd = ti.GetFuncDesc(i)
            names = ti.GetNames(fd.memid, fd.cParams + 1)  # [funcname, param1, ...]
            func_name = names[0] if names else f"<func{i}>"
            param_names = names[1:]

            ik = fd.invkind
            if ik == _INVOKE_FUNC:
                kind = "Method"
            elif ik == _INVOKE_PROPGET:
                kind = "PropGet"
            elif ik in (_INVOKE_PROPPUT, _INVOKE_PROPPUTREF):
                kind = "PropPut"
            else:
                kind = f"Invoke({ik})"

            ret_vt  = fd.retType.vt
            ret_str = _vt_name(ret_vt)

            params_info = []
            for j, ep in enumerate(fd.args):
                p_name = param_names[j] if j < len(param_names) else f"p{j}"
                p_type = _vt_name(ep.vt)
                params_info.append(f"{p_name}: {p_type}")

            members.append({
                "name":     func_name,
                "kind":     kind,
                "ret_type": ret_str,
                "params":   params_info,
                "memid":    fd.memid,
            })
        except Exception:
            continue

    # ── 遍历变量/常量（属性也可能以 VarDesc 形式暴露）────────────────────────
    for i in range(ta.cVars):
        try:
            vd = ti.GetVarDesc(i)
            names = ti.GetNames(vd.memid, 1)
            var_name = names[0] if names else f"<var{i}>"
            var_type = _vt_name(vd.elemdescVar.vt)
            members.append({
                "name":     var_name,
                "kind":     "VarDesc",
                "ret_type": var_type,
                "params":   [],
                "memid":    vd.memid,
            })
        except Exception:
            continue

    return members


def print_members(members: list[dict], title: str = "") -> None:
    """格式化打印 COM 成员列表。"""
    print(f"\n{'=' * 72}")
    if title:
        print(f"  COM 对象：{title}")
    print(f"  共 {len(members)} 个成员")
    print(f"{'=' * 72}")
    print(f"  {'种类':<10} {'成员名':<36} {'返回类型':<18} 参数")
    print(f"  {'-'*10} {'-'*36} {'-'*18} {'-'*30}")
    for m in members:
        params_str = ", ".join(m["params"]) if m["params"] else "—"
        print(f"  {m['kind']:<10} {m['name']:<36} {m['ret_type']:<18} {params_str}")
    print(f"{'=' * 72}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 2. 实际调用 Measurable 成员并打印测量结果
# ─────────────────────────────────────────────────────────────────────────────

def measure_body(meas) -> None:
    """
    调用 Measurable 的全部常用成员，打印测量结果。

    win32com 后期绑定中，ByRef Sub 的输出参数以函数返回值形式给出：
      result = meas.GetCOGPosition()   →  (Gx, Gy, Gz)  [mm]
      result = meas.GetInertia()       →  (I00..I22)     [kg·m²]
    """
    # ── 直接属性 ──────────────────────────────────────────────────────────────
    for attr in ("Mass", "Volume", "Area", "Density"):
        try:
            val = getattr(meas, attr)
            print(f"  {attr:<16} = {val}")
        except Exception as exc:
            print(f"  {attr:<16}   [不可用: {exc}]")

    print()

    # ── GetCOGPosition：返回 (Gx, Gy, Gz) [mm] ───────────────────────────────
    try:
        result = meas.GetCOGPosition()
        # result 可能是 tuple、list 或单个值；统一转为 list
        cog = list(result) if hasattr(result, "__iter__") else [result]
        labels = ["Gx", "Gy", "Gz"]
        for k, v in enumerate(cog):
            lbl = labels[k] if k < len(labels) else f"cog[{k}]"
            print(f"  {'GetCOGPosition':<16}   {lbl} = {v} mm")
    except Exception as exc:
        print(f"  GetCOGPosition    [不可用: {exc}]")

    print()

    # ── GetInertia：返回 9 元素 tuple，3×3 行主序 [kg·m²] ────────────────────
    try:
        result = meas.GetInertia()
        inertia = list(result) if hasattr(result, "__iter__") else [result]
        labels = [
            "Ixx", "Ixy", "Ixz",
            "Iyx", "Iyy", "Iyz",
            "Izx", "Izy", "Izz",
        ]
        for k, v in enumerate(inertia):
            lbl = labels[k] if k < len(labels) else f"I[{k}]"
            print(f"  {'GetInertia':<16}   {lbl} = {v} kg·m²")
    except Exception as exc:
        print(f"  GetInertia        [不可用: {exc}]")

    print()

    # ── GetPrincipalInertia：三个主惯量 [kg·m²] ──────────────────────────────
    try:
        result = meas.GetPrincipalInertia()
        vals = list(result) if hasattr(result, "__iter__") else [result]
        for k, v in enumerate(vals):
            print(f"  {'GetPrincipalInertia':<16}  I_principal[{k}] = {v} kg·m²")
    except Exception as exc:
        print(f"  GetPrincipalInertia [不可用: {exc}]")

    print()

    # ── GetPrincipalAxes：3×3 主轴方向矩阵 ───────────────────────────────────
    try:
        result = meas.GetPrincipalAxes()
        vals = list(result) if hasattr(result, "__iter__") else [result]
        axis_labels = [
            "Axis1.x", "Axis1.y", "Axis1.z",
            "Axis2.x", "Axis2.y", "Axis2.z",
            "Axis3.x", "Axis3.y", "Axis3.z",
        ]
        for k, v in enumerate(vals):
            lbl = axis_labels[k] if k < len(axis_labels) else f"axes[{k}]"
            print(f"  {'GetPrincipalAxes':<16}  {lbl} = {v}")
    except Exception as exc:
        print(f"  GetPrincipalAxes  [不可用: {exc}]")


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # ── 1. 连接 CATIA ─────────────────────────────────────────────────────────
    print("正在连接 CATIA Application ...")
    catia = win32com.client.Dispatch("CATIA.Application")

    # ── 2. 获取活动 PartDocument ───────────────────────────────────────────────
    try:
        doc = catia.ActiveDocument
    except Exception:
        print("[错误] 没有打开的活动文档，请先在 CATIA 中打开一个 CATPart 文件。")
        return

    try:
        part = doc.Part
    except Exception:
        print("[错误] 活动文档不是 PartDocument（.CATPart），请切换到零件文档。")
        return

    print(f"  活动文档：{doc.Name}")

    # ── 3. 获取 SPAWorkbench ──────────────────────────────────────────────────
    try:
        spa = doc.GetWorkbench("SPAWorkbench")
    except Exception as exc:
        print(f"[错误] 无法获取 SPAWorkbench：{exc}")
        return

    # ── 4. 创建对第一个 Body 的引用 ───────────────────────────────────────────
    try:
        body = part.Bodies.Item(1)
        ref  = part.CreateReferenceFromObject(body)
    except Exception as exc:
        print(f"[错误] 无法创建 Body 引用（文档是否含实体几何？）：{exc}")
        return

    # ── 5. 获取 Measurable 对象 ───────────────────────────────────────────────
    try:
        meas = spa.GetMeasurable(ref)
    except Exception as exc:
        print(f"[错误] spa.GetMeasurable(ref) 失败：{exc}")
        return

    print(f"  成功获取 Measurable 对象：{meas}")

    # ── 6. 枚举并打印 COM TypeInfo 中的全部成员 ───────────────────────────────
    print("\n[步骤 1] 通过 ITypeInfo 枚举 Measurable 的所有 COM 成员 ...")
    members = enumerate_com_members(meas)
    if members:
        print_members(members, title="Measurable (SPAWorkbench.GetMeasurable)")
    else:
        print("  [警告] TypeInfo 枚举未返回任何成员（可能是纯 IDispatch 对象）。")
        print("  提示：在 CATIA 目录下运行 makepy 可生成早期绑定，dir(meas) 即可看到成员。")

    # ── 7. 实际调用所有常用 Measurable 成员 ──────────────────────────────────
    print("[步骤 2] 实际调用 Measurable 常用成员 ...")
    measure_body(meas)

    # ── 8. 备用：dir() 枚举（早期绑定或 makepy 后才有效）────────────────────
    print("[步骤 3] dir(meas) 输出（后期绑定下通常只显示通用 Dispatch 属性）：")
    for name in sorted(dir(meas)):
        if not name.startswith("_"):
            print(f"  {name}")


if __name__ == "__main__":
    main()
