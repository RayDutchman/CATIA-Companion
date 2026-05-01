"""
例程：读取当前活动文档（零件或产品）的全部参数
=================================================

本文件演示如何通过 win32com 后期绑定（late-binding）接口，
枚举并读取 CATIA 当前活动文档中的 **所有参数**（Parameters 集合），
包括参数名称、类型与数值。

适用文档类型
------------
- PartDocument（.CATPart）       —— Part.Parameters
- ProductDocument（.CATProduct）—— Product.Parameters

运行前提
--------
- 已安装 pycatia（``pip install pycatia``）与 pywin32（``pip install pywin32``）
- CATIA V5 / V6 处于运行状态，并已打开至少一个零件或产品文档

win32com 后期绑定与参数读取说明
--------------------------------
CATIA 参数对象（Parameter）有两种常用的取值方式，行为差异如下：

  param.ValueAsString()
      返回 **含单位的格式化字符串**，例如 "10.5mm"、"2.3kg"。
      这是 CATIA 内部按用户显示格式拼接的字符串，单位已嵌入其中。
      本例程使用此方法读取并打印参数值。

  param.Value
      返回 **裸 SI 数值**（Python float），例如 0.0105（单位 m）、
      0.0023（单位 kg）。CATIA 内部始终以 SI 单位（kg / m / kg·m²）
      存储数值，Value 直接暴露该原始浮点数，不携带任何单位信息。
      mass_props_collect.py 正是通过 param.Value 读取 Keep 测量参数，
      因为后续代码需要对数值做换算，刻意避免解析带单位的字符串。

  param.UserUnit（RealParam 专属）
      理论上返回用户为该参数设置的显示单位字符串（如 "mm"、"kg"）。
      但在 win32com **后期绑定**（late-binding）模式下，CATIA 的
      Parameters 集合以基类 Parameter 接口暴露所有参数，无法直接
      访问 RealParam 子类的 UserUnit 属性，调用通常会抛出 COM 异常
      并返回空串，因此本例程不再尝试读取该属性。

综上：
- 若只需**显示**参数值，用 ValueAsString()，单位已包含在字符串里。
- 若需**数值计算**，用 Value，得到裸 SI float，再自行按需换算。
- 不要依赖 UserUnit 在后期绑定下正常工作。

其他常用属性
------------
  param.Name          # 完整参数名，例如 "Part1\\质量"
  param.UserAccessMode# 0=读写, 1=只读, 2=隐藏
  param.Comment       # 备注字符串（可能为空）

注意事项
--------
1. COM 索引从 1 开始（不是 0）。

2. 后期绑定下，ByRef 出参以 Python 返回值形式给出（详见
   examples/traverse_hidden_state.py），本例程无需处理此问题。

3. CATIA 参数按"扁平化"方式存放于 Parameters 集合——无论参数
   属于哪个特征，均可从文档级 Parameters 统一枚举到。

4. 若要读取 **产品树** 中某个子节点（子产品/零件实例）的参数，
   需要先拿到该节点的底层 PartDocument 或 sub-ProductDocument，
   再访问其 Parameters 集合（参见文末扩展示例）。
"""

import win32com.client


# ─────────────────────────────────────────────────────────────────────────────
# 辅助：从参数对象安全读取带单位的字符串表示
# ─────────────────────────────────────────────────────────────────────────────
def _value_str(param) -> str:
    """返回 ValueAsString() 的结果（含单位），失败时退回 str(param.Value)。

    ValueAsString() 返回 CATIA 格式化后的字符串，单位已嵌入其中，例如
    "10.5mm"、"2.3kg"。注意：这与 param.Value 返回的裸 SI float 不同——
    param.Value 不含单位，适合做数值计算（mass_props_collect.py 即使用
    param.Value 以便后续统一换算）。
    """
    try:
        return param.ValueAsString()
    except Exception:
        try:
            return str(param.Value)
        except Exception:
            return "<不可读>"


# ─────────────────────────────────────────────────────────────────────────────
# 核心：枚举 Parameters 集合，打印所有参数
# ─────────────────────────────────────────────────────────────────────────────
def print_all_parameters(params_com, title: str = "") -> list[dict]:
    """
    枚举 CATIA Parameters 集合，打印并返回所有参数信息。

    Parameters
    ----------
    params_com : win32com Parameters COM 对象
    title      : 显示在表头的文档名称（可选）

    Returns
    -------
    list[dict]  每个参数的字典，键为：
                  name, value_str, value_raw, comment

    注意：不再返回 user_unit。在 win32com 后期绑定模式下，Parameters 集合
    以基类 Parameter 接口暴露参数，无法访问 RealParam 子类的 UserUnit 属性，
    调用通常会抛出 COM 异常。单位信息已包含在 value_str（ValueAsString()）
    的字符串里，无需单独的单位列。
    """
    try:
        count = params_com.Count
    except Exception as exc:
        print(f"[错误] 无法读取 Parameters.Count：{exc}")
        return []

    print(f"\n{'=' * 60}")
    if title:
        print(f"  文档：{title}")
    print(f"  共找到 {count} 个参数")
    print(f"{'=' * 60}")
    # 并列打印两种读取方式：
    #   ValueAsString() → 含单位的格式化字符串，如 "10.5mm"
    #   param.Value     → 裸 SI float，如 0.0105（无单位，适合数值计算）
    print(f"  {'#':<5} {'参数名':<45} {'ValueAsString()':<25} {'param.Value（SI裸值）'}")
    print(f"  {'-' * 5} {'-' * 45} {'-' * 25} {'-' * 20}")

    results = []
    for i in range(1, count + 1):          # COM 集合索引从 1 开始
        try:
            param = params_com.Item(i)
            name    = str(param.Name)
            val_str = _value_str(param)

            # 尝试读取注释（部分参数没有 Comment 属性）
            try:
                comment = str(param.Comment)
            except Exception:
                comment = ""

            # 尝试读取原始数值（float / int / str）
            try:
                val_raw = param.Value
            except Exception:
                val_raw = None

            print(f"  {i:<5} {name:<45} {val_str:<25} {val_raw}")

            results.append({
                "name":      name,
                "value_str": val_str,
                "value_raw": val_raw,
                "comment":   comment,
            })

        except Exception as exc:
            print(f"  {i:<5} [读取第 {i} 个参数时出错: {exc}]")

    print(f"{'=' * 60}\n")
    return results


# ─────────────────────────────────────────────────────────────────────────────
# 主流程：连接 CATIA，读取活动文档的全部参数
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # ── 1. 连接 CATIA（后期绑定）──────────────────────────────────────────────
    print("正在连接 CATIA Application ...")
    catia = win32com.client.Dispatch("CATIA.Application")
    print(f"  CATIA 版本：{catia.Version}")

    # ── 2. 获取活动文档 ────────────────────────────────────────────────────────
    try:
        doc = catia.ActiveDocument
    except Exception:
        print("[错误] 没有打开的活动文档，请先在 CATIA 中打开一个零件或产品文档。")
        return

    doc_name = str(doc.Name)
    doc_type = str(type(doc))
    print(f"  活动文档：{doc_name}  (COM 类型: {doc_type})")

    # ── 3. 根据文档类型获取 Parameters 集合 ───────────────────────────────────
    # CATIA 文档分两种：
    #   PartDocument    → doc.Part.Parameters
    #   ProductDocument → doc.Product.Parameters
    #
    # 判断方式：尝试访问 doc.Part；若抛出 AttributeError/com_error 则为 Product。
    params_com = None
    try:
        part = doc.Part          # 若是 PartDocument，此行成功
        params_com = part.Parameters
        print("  文档类型：PartDocument（零件）")
    except Exception:
        pass

    if params_com is None:
        try:
            product = doc.Product   # ProductDocument
            params_com = product.Parameters
            print("  文档类型：ProductDocument（产品/装配体）")
        except Exception as exc:
            print(f"[错误] 无法获取 Parameters 集合：{exc}")
            return

    # ── 4. 枚举并打印全部参数 ──────────────────────────────────────────────────
    rows = print_all_parameters(params_com, title=doc_name)

    # ── 5. 示例：按名称查找特定参数 ───────────────────────────────────────────
    search_keyword = "质量"    # 修改为你想查找的关键词
    matches = [r for r in rows if search_keyword in r["name"]]
    if matches:
        print(f"包含关键词 '{search_keyword}' 的参数：")
        for m in matches:
        print(f"  {m['name']} = {m['value_str']}  (SI裸值: {m['value_raw']})")
    else:
        print(f"未找到包含关键词 '{search_keyword}' 的参数。")


# ─────────────────────────────────────────────────────────────────────────────
# 扩展示例：读取产品树中指定子零件的参数
# ─────────────────────────────────────────────────────────────────────────────
def read_part_parameters_by_occurrence(catia, root_product, part_number_name: str):
    """
    在产品树中找到 PartNumber == part_number_name 的第一个子产品/零件，
    打开对应的 PartDocument，并枚举其全部参数。

    参数
    ----
    catia             : CATIA Application COM 对象
    root_product      : 根产品 COM 对象（ProductDocument.Product）
    part_number_name  : 要查找的零件编号字符串（如 "Part100"）
    """
    products = root_product.Products
    for i in range(1, products.Count + 1):
        child = products.Item(i)
        if str(child.PartNumber) == part_number_name:
            # 子节点的引用文档
            try:
                ref_doc = child.ReferenceProduct.Parent
                params_com = None
                try:
                    params_com = ref_doc.Part.Parameters
                except Exception:
                    params_com = ref_doc.Product.Parameters
                print_all_parameters(params_com, title=part_number_name)
                return
            except Exception as exc:
                print(f"[错误] 读取 {part_number_name} 的参数失败：{exc}")
                return
    print(f"[警告] 在根产品直接子节点中未找到 PartNumber='{part_number_name}'")


if __name__ == "__main__":
    main()
