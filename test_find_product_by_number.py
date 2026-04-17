"""
测试 product.products.com_object.FindProductByNumber(pn) 是否可用。

用法（在装有 CATIA V5 的 Windows 机器上，先打开一个 CATProduct 文档）：
    python test_find_product_by_number.py [PartNumber]

若不传参数，脚本会自动取产品树中第一个子节点的 PartNumber 来测试。
"""

import sys

# ── 1. 连接 CATIA ────────────────────────────────────────────────────────────
try:
    from pycatia import catia
    from pycatia.product_structure_interfaces.product_document import ProductDocument
except ImportError:
    print("[ERROR] pycatia 未安装，请先执行: pip install pycatia")
    sys.exit(1)

print("正在连接 CATIA …")
try:
    caa = catia()
    application = caa.application
except Exception as e:
    print(f"[ERROR] 无法连接 CATIA（请确保 CATIA 已运行）: {e}")
    sys.exit(1)

# ── 2. 获取当前活跃的 ProductDocument ────────────────────────────────────────
try:
    active_doc = application.active_document
    product_doc = ProductDocument(active_doc.com_object)
    root = product_doc.product
    print(f"[OK] 当前文档: {active_doc.name}")
    print(f"[OK] 根产品 PartNumber: {root.part_number}")
except Exception as e:
    print(f"[ERROR] 无法获取活跃 ProductDocument（请确保已打开 .CATProduct）: {e}")
    sys.exit(1)

# ── 3. 确定要查找的 PartNumber ────────────────────────────────────────────────
if len(sys.argv) > 1:
    target_pn = sys.argv[1]
    print(f"[INFO] 使用命令行指定的 PartNumber: {target_pn!r}")
else:
    # 自动取第一个子产品的 PN
    try:
        count = root.products.count
        if count == 0:
            print("[WARN] 根产品没有子节点，改用根产品自身的 PartNumber 进行测试")
            target_pn = root.part_number
        else:
            first_child = root.products.item(1)
            target_pn = first_child.part_number
            print(f"[INFO] 自动选取第一个子节点的 PartNumber: {target_pn!r}")
    except Exception as e:
        print(f"[ERROR] 无法读取子节点: {e}")
        sys.exit(1)

# ── 4. 检查 COM 对象是否暴露 FindProductByNumber ──────────────────────────────
products_com = root.products.com_object      # 根产品的 Products 集合 COM 对象
products_root_com = root.com_object          # 根产品自身的 COM 对象

print("\n─── 内省 COM 对象上的方法 ───────────────────────────────────────────────")

# 尝试通过 dir() 检查（COM 对象往往不支持 dir，但值得一试）
found_on_products = "FindProductByNumber" in dir(products_com)
found_on_root     = "FindProductByNumber" in dir(products_root_com)
print(f"  dir(products.com_object) 中含 FindProductByNumber : {found_on_products}")
print(f"  dir(root.com_object)     中含 FindProductByNumber : {found_on_root}")

# 用 hasattr 确认属性是否真实可调用
has_on_products = hasattr(products_com, "FindProductByNumber")
has_on_root     = hasattr(products_root_com, "FindProductByNumber")
print(f"  hasattr(products.com_object, 'FindProductByNumber'): {has_on_products}")
print(f"  hasattr(root.com_object,     'FindProductByNumber'): {has_on_root}")

# ── 5. 实际调用测试 ───────────────────────────────────────────────────────────
print(f"\n─── 实际调用 FindProductByNumber({target_pn!r}) ──────────────────────────")

# 5-a: 在 Products 集合上调用
print("\n[尝试 1] products.com_object.FindProductByNumber(pn) …")
try:
    result = products_com.FindProductByNumber(target_pn)
    print(f"  ✓ 成功！返回值: {result}")
except AttributeError as e:
    print(f"  ✗ AttributeError（方法不存在）: {e}")
except Exception as e:
    print(f"  ✗ 其他异常: {type(e).__name__}: {e}")

# 5-b: 在根产品 COM 对象本身上调用
print("\n[尝试 2] root.com_object.FindProductByNumber(pn) …")
try:
    result = products_root_com.FindProductByNumber(target_pn)
    print(f"  ✓ 成功！返回值: {result}")
except AttributeError as e:
    print(f"  ✗ AttributeError（方法不存在）: {e}")
except Exception as e:
    print(f"  ✗ 其他异常: {type(e).__name__}: {e}")

# 5-c: 在 Products 集合上用 Item(pn) 按名称查找（CATIA 官方支持按 Name 索引）
print(f"\n[尝试 3] products.com_object.Item(pn) 按 PartNumber 作为 Name 查找 …")
try:
    result = products_com.Item(target_pn)
    print(f"  ✓ 成功！返回值: {result}")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

# 5-d: 通过 pycatia 封装层的 item(pn) 按名称查找
print(f"\n[尝试 4] root.products.item(pn) (pycatia 封装) …")
try:
    result = root.products.item(target_pn)
    print(f"  ✓ 成功！返回值: {result}")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

print("\n─── 测试完毕 ────────────────────────────────────────────────────────────")
