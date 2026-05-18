"""
探针脚本：找出 CATIA Product/Part 对象上"描述"字段的正确访问方式
===================================================================

CATIA V5 COM 自动化 API 中，产品的"描述"属性：
  - IDL 属性名：Description（Product 对象上的读写属性）
  - CAA/C++ 内部名：V_description
  - pycatia 封装后：product.description（推测）

本脚本通过多种方式穷举尝试，打印每种方式的结果，帮助确认哪种有效。

运行前提
--------
- CATIA V5 已启动并打开了一个 CATProduct 文档
- 至少有一个节点在 CATIA 的"属性"对话框里填写了"描述"文本（否则全部为空，
  无法区分"读取成功但为空"和"读取失败"）
- 已安装 pycatia 和 pywin32
"""

import win32com.client


def probe_description(product_com, label: str = "") -> None:
    """对单个 CATIA Product COM 对象，穷举各种"描述"读取方式并打印结果。"""
    print(f"\n{'─' * 60}")
    print(f"  节点：{label}")
    print(f"{'─' * 60}")

    # ── 1. 直接 COM 属性（IDL 标准名） ───────────────────────────────────────
    candidates_com = [
        "Description",
        "V_description",
        "V_Description",
        "description",
        "DescriptionRef",
    ]
    print("\n[A] 直接 COM 属性（product_com.<name>）：")
    for name in candidates_com:
        try:
            val = getattr(product_com, name)
            print(f"    product_com.{name:<20} = {val!r}")
        except AttributeError:
            print(f"    product_com.{name:<20} → AttributeError（属性不存在）")
        except Exception as e:
            print(f"    product_com.{name:<20} → {type(e).__name__}: {e}")

    # ── 2. pycatia Product 对象包装 ──────────────────────────────────────────
    print("\n[B] pycatia Product 包装（需要 pycatia 可用）：")
    try:
        from pycatia.product_structure_interfaces.product import Product as PyProduct
        py_product = PyProduct(product_com)
        for name in ["description", "Description", "v_description"]:
            try:
                val = getattr(py_product, name)
                print(f"    py_product.{name:<20} = {val!r}")
            except AttributeError:
                print(f"    py_product.{name:<20} → AttributeError")
            except Exception as e:
                print(f"    py_product.{name:<20} → {type(e).__name__}: {e}")
    except Exception as e:
        print(f"    无法创建 pycatia Product 包装：{e}")

    # ── 3. reference_product 的 COM 属性 ────────────────────────────────────
    print("\n[C] reference_product COM 属性：")
    try:
        ref = product_com.ReferenceProduct
        for name in candidates_com:
            try:
                val = getattr(ref, name)
                print(f"    ref.{name:<24} = {val!r}")
            except AttributeError:
                print(f"    ref.{name:<24} → AttributeError")
            except Exception as e:
                print(f"    ref.{name:<24} → {type(e).__name__}: {e}")
    except Exception as e:
        print(f"    无法访问 ReferenceProduct：{e}")

    # ── 4. 通过 Parameters 集合查找含 "desc"/"描述" 的参数 ──────────────────
    print("\n[D] Parameters 集合中含 'desc'/'描述' 关键词的参数：")
    found_params = False
    for target_com, src_name in [(product_com, "product_com"),
                                  (getattr(product_com, "ReferenceProduct", None), "ref")]:
        if target_com is None:
            continue
        try:
            params = target_com.Parameters
            count = params.Count
            for i in range(1, count + 1):
                try:
                    p = params.Item(i)
                    pname = str(p.Name).lower()
                    if "desc" in pname or "描述" in pname:
                        print(f"    [{src_name}] 参数 #{i}: {p.Name!r} = {p.ValueAsString()!r}")
                        found_params = True
                except Exception:
                    pass
        except Exception as e:
            print(f"    [{src_name}] 无法访问 Parameters：{e}")
    if not found_params:
        print("    （未找到含 'desc'/'描述' 的参数）")

    # ── 5. 列出 COM 对象所有可用属性（TYPEINFO，若可用） ────────────────────
    print("\n[E] COM TypeInfo 属性列表（仅含 'desc' / 'V_' 前缀）：")
    try:
        type_info = product_com._oleobj_.GetTypeInfo(0)
        type_attr = type_info.GetTypeAttr()
        for i in range(type_attr.cVars):
            var_desc = type_info.GetVarDesc(i)
            name = type_info.GetNames(var_desc.memid)[0]
            nl = name.lower()
            if "desc" in nl or nl.startswith("v_"):
                print(f"    TypeInfo 属性：{name}")
        for i in range(type_attr.cFuncs):
            func_desc = type_info.GetFuncDesc(i)
            name = type_info.GetNames(func_desc.memid)[0]
            nl = name.lower()
            if "desc" in nl or nl.startswith("v_"):
                print(f"    TypeInfo 方法：{name}")
    except Exception as e:
        print(f"    TypeInfo 不可用：{e}")


def main():
    print("正在连接 CATIA Application ...")
    catia = win32com.client.Dispatch("CATIA.Application")

    try:
        doc = catia.ActiveDocument
    except Exception:
        print("[错误] 无活动文档，请先在 CATIA 中打开一个 CATProduct。")
        return

    print(f"活动文档：{doc.Name}")

    # 获取根产品
    try:
        root_com = doc.Product
    except Exception as e:
        print(f"[错误] 无法获取 doc.Product：{e}")
        return

    # 先探测根产品本身
    probe_description(root_com, label=f"根产品 ({root_com.PartNumber})")

    # 再探测第一个子节点（如果有）
    try:
        children = root_com.Products
        if children.Count >= 1:
            child_com = children.Item(1)
            probe_description(child_com, label=f"子节点 1 ({child_com.PartNumber})")
    except Exception as e:
        print(f"\n无法读取子节点：{e}")

    print("\n\n探测完成。")
    print("请查看上方 [A][B][C] 中哪一行返回了你在 CATIA 属性对话框里填写的描述文本，")
    print("那就是正确的属性名，告知我即可。")


if __name__ == "__main__":
    main()
