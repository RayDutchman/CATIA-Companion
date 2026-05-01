"""
例程：遍历产品子孙节点，读取每个实例的隐藏状态
=================================================

本文件演示如何通过 win32com 后期绑定（late-binding）接口，
遍历 CATIA 产品树的所有子孙节点，并读取每个节点在父装配中
的实例级可见性（显示 / 隐藏）。

运行前提
--------
- 已安装 pycatia（``pip install pycatia``）与 pywin32（``pip install pywin32``）
- CATIA V5 / V6 处于运行状态，并已打开一个产品文档（.CATProduct）

win32com 后期绑定说明
---------------------
本项目使用 ``win32com.client.Dispatch`` 而非 ``EnsureDispatch`` 访问 COM 对象，
即"后期绑定（late-binding）"模式：

  早绑定（early binding）：
    win32com.client.gencache.EnsureDispatch("CATIA.Application")
    - 从类型库生成 Python 包装类，存放于 %LOCALAPPDATA%\\Temp\\gen_py\\
    - 调用时按预编译方法签名传参，类型检查在 Python 层完成
    - 缺点：若 gen_py 缓存与当前 CATIA 版本不匹配，可能导致 COM 连接异常

  后期绑定（late binding）：
    win32com.client.Dispatch("CATIA.Application")   # 或 pycatia.catia()
    - 通过 IDispatch::GetIDsOfNames + IDispatch::Invoke 在运行时解析方法
    - 无需 gen_py 缓存，不受版本不匹配影响
    - ByRef 出参行为不同（见下文）

后期绑定对 ByRef 出参的处理
----------------------------
CATIA 的许多 COM 方法（如 ``VisPropertySet.GetShow``、
``Position.GetComponents``）在 VBA 中声明为带 ``ByRef`` 参数的 ``Sub``。
在后期绑定模式下，win32com 无法从类型库得知某参数是 ByRef 出参，
因此会将 COM 调用的 ByRef 修改后的参数值以 **Python 返回值** 的形式返回，
而非写回传入的对象：

  # ✗ 错误：传入 VARIANT(VT_BYREF|VT_I4, 0)
  #         win32com 在构造 DISPPARAMS 时尝试对其调用 int()，抛出 TypeError
  from win32com.client import VARIANT
  from pythoncom import VT_BYREF, VT_I4
  show_var = VARIANT(VT_BYREF | VT_I4, 0)
  sel.VisProperties.GetShow(show_var)          # TypeError!

  # ✓ 正确：传入普通整数 0 作为 ByRef 参数的"占位输入值"
  #         win32com 将修改后的值作为 Python 返回值给出
  result = sel.VisProperties.GetShow(0)
  # result 可能是整数，也可能是 (int,) 形式的元组（取最后一个元素）
  show_val = result[-1] if isinstance(result, tuple) else result
  hidden = bool(show_val)                      # catVisNoShow=1 → True（隐藏）

可见性值含义
------------
  catVisShow   = 0  → 可见（visible）
  catVisNoShow = 1  → 隐藏（hidden）
"""

from __future__ import annotations


def _get_catia_application():
    """获取 CATIA Application COM 对象（后期绑定）。"""
    from pycatia import catia
    caa = catia()
    return caa.application


def is_instance_hidden(product, application) -> bool:
    """检测产品节点实例在父装配中是否处于隐藏状态。

    原理：将该节点的 COM 对象添加到 ActiveDocument.Selection，然后通过
    Selection.VisProperties.GetShow() 读取实例级可见性。此方法读取的是
    "当前实例"在父装配环境下的显示状态，与零件文档自身的属性无关。

    注意：此调用会短暂修改 CATIA 的当前选择集，读取完毕后立即清空以还原。

    参数
    ----
    product     : pycatia Product 包装对象（节点实例）
    application : CATIA Application COM 对象

    返回
    ----
    True  → 实例处于隐藏状态（catVisNoShow = 1）
    False → 实例处于可见状态（catVisShow = 0），或读取失败时保守返回 False
    """
    com = product.com_object
    sel = None
    try:
        sel = application.com_object.ActiveDocument.Selection
        sel.Clear()
        sel.Add(com)

        # 后期绑定 ByRef 模式：传入 0 作为占位输入值，返回值即为出参的修改结果。
        # 绝对不要传入 VARIANT(VT_BYREF|VT_I4, 0)——win32com 会在序列化 DISPPARAMS
        # 时对其调用 int()，导致 TypeError。
        result = sel.VisProperties.GetShow(0)

        # 部分 CATIA 版本可能将返回值包装为 tuple，取最后一个元素
        show_val = result[-1] if isinstance(result, tuple) else result
        return bool(show_val) if show_val is not None else False

    except Exception:
        # 读取失败（如节点类型不支持 VisProperties）→ 保守视为可见
        return False
    finally:
        # 无论是否异常，都清空选择集，还原 CATIA 状态
        try:
            if sel is not None:
                sel.Clear()
        except Exception:
            pass


def traverse_and_print_hidden(product, application, level: int = 0) -> None:
    """递归遍历产品子孙节点，打印每个节点的名称、层级和隐藏状态。

    参数
    ----
    product     : pycatia Product 包装对象（起始节点）
    application : CATIA Application COM 对象
    level       : 当前递归深度（根节点为 0）
    """
    try:
        pn = product.part_number
    except Exception:
        name = product.name
        pn = name.rsplit(".", 1)[0] if "." in name else name

    # 根节点（level=0）是虚拟根，本身没有父装配上下文，跳过可见性检测
    if level >= 1:
        hidden = is_instance_hidden(product, application)
        status = "隐藏" if hidden else "可见"
    else:
        status = "根节点（跳过可见性检测）"

    indent = "  " * level
    print(f"{indent}[level={level}] {pn}  →  {status}")

    # 递归遍历子节点
    try:
        count = product.products.count
        for i in range(1, count + 1):
            try:
                child = product.products.item(i)
                traverse_and_print_hidden(child, application, level + 1)
            except Exception as e:
                print(f"{indent}  [!] 遍历子节点 {i} 失败: {e}")
    except Exception:
        pass  # 叶子零件无 .products，跳过


def collect_hidden_states(
    product,
    application,
    level: int = 0,
    result: list[dict] | None = None,
) -> list[dict]:
    """递归遍历产品子孙节点，收集每个节点的名称和隐藏状态，返回列表。

    返回值中每个元素为：
        {
          "level":   int,   # 层级深度（根节点为 0）
          "name":    str,   # 零件号（part_number）
          "hidden":  bool,  # True = 隐藏，False = 可见
        }
    """
    if result is None:
        result = []

    try:
        pn = product.part_number
    except Exception:
        name = product.name
        pn = name.rsplit(".", 1)[0] if "." in name else name

    if level >= 1:
        hidden = is_instance_hidden(product, application)
    else:
        hidden = False  # 根节点不检测

    result.append({"level": level, "name": pn, "hidden": hidden})

    try:
        count = product.products.count
        for i in range(1, count + 1):
            try:
                child = product.products.item(i)
                collect_hidden_states(child, application, level + 1, result)
            except Exception:
                pass
    except Exception:
        pass

    return result


# ---------------------------------------------------------------------------
# 主入口：直接运行本脚本即可演示遍历与隐藏状态读取
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from pycatia.product_structure_interfaces.product_document import ProductDocument

    app = _get_catia_application()

    # 使用当前 CATIA 活动文档（需为已打开的 .CATProduct）
    product_doc = ProductDocument(app.active_document.com_object)
    root = product_doc.product

    print("=== 遍历产品树：显示每个节点的可见状态 ===\n")
    traverse_and_print_hidden(root, app, level=0)

    print("\n=== 汇总：隐藏节点列表 ===")
    states = collect_hidden_states(root, app)
    hidden_nodes = [s for s in states if s["hidden"]]
    if hidden_nodes:
        for s in hidden_nodes:
            print(f"  level={s['level']}  {s['name']}")
    else:
        print("  （无隐藏节点）")
