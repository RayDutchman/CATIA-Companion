"""
BOM data-collection helpers.

Provides:
- get_product_filepath()  – resolve the backing file path of a CATIA product
- collect_bom_rows()      – traverse a product tree and return a list of row dicts
"""

import logging
from collections.abc import Callable
from pathlib import Path

from catia_companion.constants import FILENAME_NOT_FOUND

logger = logging.getLogger(__name__)


def get_product_filepath(product) -> str:
    """Return the full file path of the CATIA document backing *product*.

    Tries several COM access patterns in order and returns an empty string if
    none succeed.
    """
    for accessor in (
        lambda p: p.reference_product.com_object.Parent.FullName,
        lambda p: p.com_object.ReferenceProduct.Parent.FullName,
        lambda p: p.com_object.Parent.FullName,
    ):
        try:
            return accessor(product)
        except Exception:
            pass
    return ""


def collect_bom_rows(
    file_path: str | None,
    columns: list[str],
    custom_columns: list[str],
    progress_callback: Callable[[int], None] | None = None,
) -> list[dict]:
    """Return a list of row dicts representing the hierarchical BOM.

    Parameters
    ----------
    file_path:
        Path to a ``.CATProduct`` file.  Pass ``None`` to use the currently
        active CATIA document without opening or closing any file.
    columns:
        The column names (internal) to read for each product node.
    custom_columns:
        Column names that are user-defined properties (read via
        ``UserRefProperties``).
    progress_callback:
        Optional callable invoked with the current row count after each node
        is appended to the result list.  May raise an exception to abort the
        traversal (e.g. when the user cancels).
    """
    from pycatia import catia, CatWorkModeType
    from pycatia.product_structure_interfaces.product_document import ProductDocument

    DIRECT_ATTR_MAP: dict[str, str] = {
        "Nomenclature": "nomenclature",
        "Revision":     "revision",
        "Definition":   "definition",
        "Source":       "source",
    }

    def _get_prop(product, name: str) -> str:
        attr = DIRECT_ATTR_MAP.get(name)
        if not attr:
            return ""
        targets = [product]
        try:
            targets.insert(0, product.reference_product)
        except Exception:
            pass
        for target in targets:
            try:
                value = getattr(target, attr)
                if value is not None:
                    return str(value)
            except Exception:
                pass
        return ""

    def _get_user_prop(product, name: str) -> str:
        targets = [product]
        try:
            targets.insert(0, product.reference_product)
        except Exception:
            pass
        for target in targets:
            try:
                prop  = target.user_ref_properties.item(name)
                value = prop.value
                if value is not None and str(value).strip():
                    return str(value)
            except Exception:
                pass
        return ""

    _total_count: list[int] = [0]

    def _traverse(product, rows: list, level: int, parent_filepath: str = "") -> None:
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

        filepath  = get_product_filepath(product)
        not_found = not bool(filepath)
        is_readable = True
        try:
            product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
        except Exception:
            is_readable = False

        row: dict = {
            "Level":        level,
            "Part Number":  pn,
            "Filename":     Path(filepath).stem if filepath else FILENAME_NOT_FOUND,
            "_filepath":    filepath,
            "_not_found":   not_found,
            "_unreadable":  not is_readable,
        }

        try:
            child_count = product.products.count
            if filepath and filepath == parent_filepath:
                # The child shares the same backing file as its parent, which
                # means it is an embedded sub-assembly (部件) rather than a
                # standalone product (产品) or leaf part (零件).
                row["Type"] = "部件"
            elif child_count > 0:
                row["Type"] = "产品"
            else:
                row["Type"] = "零件"
        except Exception:
            row["Type"] = ""

        for col in columns:
            if col in DIRECT_ATTR_MAP:
                row[col] = _get_prop(product, col)
            elif col in custom_columns:
                row[col] = _get_user_prop(product, col)

        rows.append(row)
        _total_count[0] += 1
        if progress_callback is not None:
            progress_callback(_total_count[0])

        try:
            products  = product.products
            count     = products.count
            if count == 0:
                return
            children: dict = {}
            for i in range(1, count + 1):
                try:
                    child = products.item(i)
                    try:
                        cpn = child.part_number
                    except Exception:
                        try:
                            cpn = child.reference_product.part_number
                        except Exception:
                            n   = child.name
                            cpn = n.rsplit(".", 1)[0] if "." in n else n
                except Exception:
                    continue
                if cpn not in children:
                    children[cpn] = {"product": child, "qty": 0}
                children[cpn]["qty"] += 1

            for cpn, data in children.items():
                child_rows: list = []
                _traverse(data["product"], child_rows, level + 1,
                          parent_filepath=filepath)
                if child_rows:
                    child_rows[0]["Quantity"] = data["qty"]
                rows.extend(child_rows)
        except Exception:
            pass

    # ── CATIA connection ────────────────────────────────────────────────────
    caa         = catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    if file_path is None:
        product_doc  = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product
        rows: list[dict] = []
        _traverse(root_product, rows, level=0)
        return rows

    src = Path(file_path).resolve()
    already_open: set[Path] = set()
    for i in range(1, documents.count + 1):
        try:
            already_open.add(Path(documents.item(i).full_name).resolve())
        except Exception:
            pass

    if src not in already_open:
        documents.open(str(src))

    target_doc = None
    for i in range(1, documents.count + 1):
        try:
            doc = documents.item(i)
            if Path(doc.full_name).resolve() == src:
                target_doc = doc
                break
        except Exception:
            pass
    if target_doc is None:
        raise RuntimeError(f"无法在CATIA中找到文档：{src}")

    product_doc  = ProductDocument(target_doc.com_object)
    root_product = product_doc.product
    rows = []
    _traverse(root_product, rows, level=0)
    return rows
