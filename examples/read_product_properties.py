"""
例程：读取单个零件/产品的内置属性与用户自定义属性
=====================================================

本文件演示如何通过 pycatia / win32com 后期绑定接口，对 **单个** 产品节点
（无需遍历子树）读取两类属性：

1. **内置属性**（CATIA "Properties" 对话框 → "Product" 选项卡）
   - Nomenclature（命名）
   - Revision（修订版本）
   - Definition（定义）
   - Source（来源）

   这些属性通过 pycatia ``Product`` 对象的同名 Python 属性直接读取。
   当产品实例没有自己的 ``ReferenceProduct``（例如根产品）时，直接读取
   ``product`` 本身；否则优先读取 ``reference_product``，再回退到实例。
   这与 ``bom_collect._get_prop()`` 的逻辑完全一致。

2. **用户自定义属性**（CATIA "Properties" 对话框 → "User Defined" / "Properties"
   选项卡中手动添加的属性）
   通过 ``product.user_ref_properties.item(name)`` 读取。同样优先尝试
   ``reference_product``，再回退到实例，与 ``bom_collect._get_user_prop()`` 一致。

运行前提
--------
- 已安装 pycatia（``pip install pycatia``）
- CATIA V5/V6 处于运行状态，并已打开至少一个产品或零件文档

与 bom_collect.py 的对应关系
-----------------------------
``bom_collect.py`` 中的 ``_get_prop()`` 和 ``_get_user_prop()`` 都是嵌套函数，
不对外暴露。本例程将同等逻辑提取为可独立调用的顶层函数，便于在调试或独立脚本
中复用。
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# 内置属性字段映射（与 bom_collect.DIRECT_ATTR_MAP 完全对应）
# ---------------------------------------------------------------------------
_DIRECT_ATTR_MAP: dict[str, str] = {
    "Nomenclature": "nomenclature",
    "Revision":     "revision",
    "Definition":   "definition",
    "Source":       "source",
}

# 全部内置属性字段名（可直接传给 read_builtin_properties()）
BUILTIN_PROPERTY_NAMES: list[str] = list(_DIRECT_ATTR_MAP.keys())


def _candidate_targets(product):
    """返回读取属性时要依次尝试的 COM 目标列表。

    优先尝试 ``reference_product``（引用产品，对应文档内的定义），
    再回退到实例本身。这与 bom_collect 内部的处理逻辑完全一致。
    """
    targets = [product]
    try:
        targets.insert(0, product.reference_product)
    except Exception:
        pass
    return targets


def read_builtin_properties(product, names: list[str] | None = None) -> dict[str, str]:
    """读取 CATIA 产品的内置属性（Nomenclature / Revision / Definition / Source）。

    与 ``bom_collect._get_prop()`` 逻辑相同，但作为独立函数暴露。

    参数
    ----
    product : pycatia Product 包装对象
    names   : 要读取的属性字段名列表；默认读取全部四个内置字段。

    返回
    ----
    dict，key 为字段名，value 为字符串（空字符串表示未填写或读取失败）。
    """
    if names is None:
        names = BUILTIN_PROPERTY_NAMES

    targets = _candidate_targets(product)
    result: dict[str, str] = {}

    for name in names:
        attr = _DIRECT_ATTR_MAP.get(name)
        if not attr:
            result[name] = ""
            continue
        value = ""
        for target in targets:
            try:
                v = getattr(target, attr)
                if v is not None:
                    value = str(v)
                    break
            except Exception:
                pass
        result[name] = value

    return result


def read_user_defined_properties(
    product,
    names: list[str] | None = None,
) -> dict[str, str]:
    """读取 CATIA 产品的用户自定义属性（User-Defined Properties）。

    与 ``bom_collect._get_user_prop()`` 逻辑相同，但作为独立函数暴露。

    当 ``names`` 为 ``None`` 时，自动枚举所有用户自定义属性。
    当 ``names`` 为具体列表时，只读取指定字段（未找到的字段返回空字符串）。

    参数
    ----
    product : pycatia Product 包装对象
    names   : 要读取的属性名列表；传 ``None`` 则枚举全部用户自定义属性。

    返回
    ----
    dict，key 为属性名，value 为字符串。
    """
    targets = _candidate_targets(product)

    if names is None:
        # 枚举所有用户自定义属性（以第一个可用 target 的 user_ref_properties 为准）
        all_props: dict[str, str] = {}
        for target in targets:
            try:
                urp = target.user_ref_properties
                for i in range(1, urp.count + 1):
                    try:
                        prop = urp.item(i)
                        v    = prop.value
                        all_props[prop.name] = str(v) if v is not None else ""
                    except Exception:
                        pass
                # 成功枚举则不再尝试下一个 target
                break
            except Exception:
                pass
        return all_props

    # 按指定名称逐一读取
    result: dict[str, str] = {}
    for name in names:
        value = ""
        for target in targets:
            try:
                prop  = target.user_ref_properties.item(name)
                v     = prop.value
                if v is not None and str(v).strip():
                    value = str(v)
                    break
            except Exception:
                pass
        result[name] = value
    return result


def print_all_properties(product, indent: str = "") -> None:
    """打印单个产品节点的所有内置属性与用户自定义属性。

    参数
    ----
    product : pycatia Product 包装对象
    indent  : 打印时的缩进前缀（默认空字符串）
    """
    try:
        pn = product.part_number
    except Exception:
        name = product.name
        pn   = name.rsplit(".", 1)[0] if "." in name else name

    print(f"{indent}零件号（Part Number）: {pn}")

    # ── 内置属性 ──────────────────────────────────────────────────────────
    print(f"{indent}【内置属性】")
    builtin = read_builtin_properties(product)
    for k, v in builtin.items():
        display = v if v else "（未填写）"
        print(f"{indent}  {k}: {display}")

    # ── 用户自定义属性 ────────────────────────────────────────────────────
    print(f"{indent}【用户自定义属性】")
    user_props = read_user_defined_properties(product)
    if user_props:
        for k, v in user_props.items():
            display = v if v else "（空值）"
            print(f"{indent}  {k}: {display}")
    else:
        print(f"{indent}  （无用户自定义属性）")


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from pycatia import catia
    from pycatia.product_structure_interfaces.product_document import ProductDocument

    caa = catia()
    app = caa.application

    # 获取当前活动文档
    # bom_collect.py 的做法：直接用 ProductDocument 包装 com_object，
    # 该方法对 .CATPart 和 .CATProduct 均有效，无需探测文档类型。
    active_doc = app.active_document
    com_doc    = active_doc.com_object

    # 通过文件扩展名区分零件 / 产品，仅用于展示信息（不影响后续读取）
    doc_name = active_doc.full_name
    if doc_name.lower().endswith(".catpart"):
        print(f"当前文档为 .CATPart：{doc_name}")
    else:
        print(f"当前文档为 .CATProduct：{doc_name}")

    root = ProductDocument(com_doc).product

    print("=" * 55)
    print("读取根产品的内置属性与用户自定义属性")
    print("=" * 55)
    print_all_properties(root)

    # ── 可选：指定属性名读取 ─────────────────────────────────────────────
    print()
    print("─" * 55)
    print("【按名称读取指定内置属性示例】")
    partial = read_builtin_properties(root, names=["Nomenclature", "Revision"])
    for k, v in partial.items():
        print(f"  {k}: {v if v else '（未填写）'}")

    print()
    print("─" * 55)
    print("【按名称读取指定用户自定义属性示例（属性不存在时返回空字符串）】")
    # 将下方列表替换为你的产品实际拥有的用户自定义属性名
    custom_names = ["Material", "Supplier", "DrawingNumber"]
    custom = read_user_defined_properties(root, names=custom_names)
    for k, v in custom.items():
        print(f"  {k}: {v if v else '（未找到或为空）'}")
