"""
例程：读取当前活动文档（零件或产品）的全部参数
=================================================

本文件演示如何通过 win32com 后期绑定（late-binding）接口，
枚举并读取 CATIA 当前活动文档中的 **所有参数**（Parameters 集合），
包括参数名称、类型、数值以及单位。

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
CATIA 参数对象（Parameter）的常用属性：

  param.Name          # 完整参数名，例如 "Part1\\质量"
  param.Value         # 当前值（Python float / int / str，取决于参数类型）
  param.UserAccessMode# 0=读写, 1=只读, 2=隐藏
  param.Comment       # 备注字符串（可能为空）

此外，若需获取 **带单位的字符串**（如 "10.5mm"），可调用：
  param.ValueAsString()  # 返回含单位的字符串

对于 RealParam（实数类型），还支持：
  param.UserUnit       # 用户单位字符串，如 "mm"、"kg"
  param.Magnitude      # 量纲字符串，如 "LENGTH"、"MASS"

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
    """返回参数的字符串表示（含单位）。失败时退回 str(param.Value)。"""
    try:
        return param.ValueAsString()
    except Exception:
        try:
            return str(param.Value)
        except Exception:
            return "<不可读>"


def _user_unit(param) -> str:
    """返回参数的用户单位（仅 RealParam 有效）。无单位时返回空串。"""
    try:
        return str(param.UserUnit)
    except Exception:
        return ""


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
                  name, value_str, value_raw, user_unit, comment
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
    print(f"  {'#':<5} {'参数名':<45} {'值（含单位）':<20} {'用户单位'}")
    print(f"  {'-' * 5} {'-' * 45} {'-' * 20} {'-' * 10}")

    results = []
    for i in range(1, count + 1):          # COM 集合索引从 1 开始
        try:
            param = params_com.Item(i)
            name       = str(param.Name)
            val_str    = _value_str(param)
            unit       = _user_unit(param)

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

            print(f"  {i:<5} {name:<45} {val_str:<20} {unit}")

            results.append({
                "name":       name,
                "value_str":  val_str,
                "value_raw":  val_raw,
                "user_unit":  unit,
                "comment":    comment,
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
            print(f"  {m['name']} = {m['value_str']}  ({m['user_unit']})")
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
