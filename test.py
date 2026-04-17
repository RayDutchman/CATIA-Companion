"""
Diagnostic script: print 3 COM filepath accessors for every node
in the active CATIA product tree.

Run from the repo root with CATIA open and a CATProduct as the active doc:
    python test.py

Output columns per node
-----------------------
  LEVEL   : indent depth (0 = root)
  PN      : product.part_number (or .name fallback)
  TYPE    : 产品 / 零件 / 部件 (same logic as bom_collect._traverse)
  A0      : product.reference_product.com_object.Parent.FullName
  A1      : product.com_object.ReferenceProduct.Parent.FullName
  A2      : product.com_object.Parent.FullName
  CHOSEN  : which accessor won (0/1/2) – same as get_product_filepath()
  IS_OWN  : True when chosen path != parent node's path (standalone file)
            False when == parent path (embedded 部件)

Findings (2026-04-17)
---------------------
  For standalone 产品/零件:
    A0, A1, A2 all succeed and return the node's own file path.
  For embedded 部件:
    A0 succeeds → returns PARENT's file path (correct for 部件 detection)
    A1 succeeds → returns PARENT's file path (same result)
    A2 raises AttributeError – the COM object has no .Parent on embedded nodes.
  Conclusion: A2 is never useful and was removed first. Then A0 was also
  dropped because A1 (pure COM, no pycatia wrapper) is simpler and more
  robust. Production code now uses only A1:
      product.com_object.ReferenceProduct.Parent.FullName
"""

import sys

try:
    from pycatia import catia
    from pycatia.product_structure_interfaces.product_document import ProductDocument
except ImportError:
    sys.exit("pycatia is not installed – run: pip install pycatia")


def _try(fn, product):
    try:
        return fn(product), None
    except Exception as exc:
        return None, f"ERR: {type(exc).__name__}: {exc}"


def _get_filepath(product):
    """Mirror of bom_collect.get_product_filepath() – return (path, accessor_idx)."""
    accessors = [
        lambda p: p.reference_product.com_object.Parent.FullName,
        lambda p: p.com_object.ReferenceProduct.Parent.FullName,
        lambda p: p.com_object.Parent.FullName,
    ]
    for i, fn in enumerate(accessors):
        val, _ = _try(fn, product)
        if val:
            return val, i
    return "", -1


def _traverse(product, level=0, parent_filepath="", rows=None):
    if rows is None:
        rows = []

    try:
        pn = product.part_number
    except Exception:
        name = product.name
        pn = name.rsplit(".", 1)[0] if "." in name else name

    # All 3 raw accessor values
    a0, e0 = _try(lambda p: p.reference_product.com_object.Parent.FullName, product)
    a1, e1 = _try(lambda p: p.com_object.ReferenceProduct.Parent.FullName, product)
    a2, e2 = _try(lambda p: p.com_object.Parent.FullName, product)

    chosen_path, chosen_idx = _get_filepath(product)
    is_own = bool(chosen_path) and chosen_path != parent_filepath

    # Determine node type (same logic as bom_collect)
    try:
        child_count = product.products.count
    except Exception:
        child_count = 0

    if chosen_path and chosen_path == parent_filepath:
        node_type = "部件"
    elif child_count > 0:
        node_type = "产品"
    else:
        node_type = "零件"

    rows.append({
        "level":        level,
        "pn":           pn,
        "type":         node_type,
        "a0":           a0 or e0,
        "a1":           a1 or e1,
        "a2":           a2 or e2,
        "chosen_idx":   chosen_idx,
        "chosen_path":  chosen_path,
        "parent_path":  parent_filepath,
        "is_own":       is_own,
    })

    for i in range(1, child_count + 1):
        try:
            child = product.products.item(i)
            _traverse(child, level + 1, chosen_path, rows)
        except Exception as exc:
            rows.append({
                "level":      level + 1,
                "pn":         f"<error: {exc}>",
                "type":       "?",
                "a0": "", "a1": "", "a2": "",
                "chosen_idx": -1, "chosen_path": "", "parent_path": "",
                "is_own": False,
            })

    return rows


def main():
    caa = catia()
    app = caa.application
    try:
        doc = app.active_document
    except Exception as exc:
        sys.exit(f"No active CATIA document: {exc}")

    try:
        product_doc = ProductDocument(doc.com_object)
        root = product_doc.product
    except Exception as exc:
        sys.exit(f"Active document is not a CATProduct: {exc}")

    print(f"\nActive document: {doc.full_name}\n")

    rows = _traverse(root)

    # ── Pretty-print ─────────────────────────────────────────────────────────
    sep = "─" * 160
    print(sep)
    print(f"{'LVL':>3}  {'TYPE':>4}  {'WIN?':>5}  {'A-IDX':>5}  {'PN':<30}  {'CHOSEN PATH'}")
    print(sep)

    for r in rows:
        indent = "  " * r["level"]
        print(f"{r['level']:>3}  {r['type']:>4}  {str(r['is_own']):>5}  "
              f"{r['chosen_idx']:>5}  {indent}{r['pn']:<{max(1,30-len(indent))}}  "
              f"{r['chosen_path']}")

    print(sep)
    print("\n── Raw accessor breakdown ──\n")
    for r in rows:
        indent = "  " * r["level"]
        print(f"{indent}[{r['level']}] {r['pn']}  (type={r['type']}, is_own={r['is_own']})")
        print(f"{indent}  A0 (ref_product.co.Parent.FullName):  {r['a0']}")
        print(f"{indent}  A1 (co.RefProduct.Parent.FullName):   {r['a1']}")
        print(f"{indent}  A2 (co.Parent.FullName):              {r['a2']}")
        print(f"{indent}  parent_path:                          {r['parent_path']}")
        print()


if __name__ == "__main__":
    main()
