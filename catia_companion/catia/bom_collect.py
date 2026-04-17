"""
BOM data-collection helpers.

Provides:
- get_product_filepath()     – resolve the backing file path of a CATIA product
- collect_bom_rows()         – traverse a product tree and return a list of row dicts
- flatten_bom_to_summary()   – collapse a hierarchical BOM into a flat summary
                               (unique parts with cumulative quantities)
"""

import logging
from collections.abc import Callable
from pathlib import Path

from catia_companion.constants import FILENAME_NOT_FOUND

logger = logging.getLogger(__name__)


def get_product_filepath(product) -> str:
    """Return the full file path of the CATIA document backing *product*.

    Uses ``com_object.ReferenceProduct.Parent.FullName`` – a pure COM path
    that works for both standalone products/parts and embedded 部件 (which
    have no own file and return their parent's path).  Returns an empty
    string on failure.
    """
    try:
        return product.com_object.ReferenceProduct.Parent.FullName
    except Exception:
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

    # Cache properties by filepath to avoid redundant DESIGN_MODE switches and
    # COM property reads for the same physical document referenced multiple
    # times in the assembly tree (e.g. the same fastener used 50 times).
    # NOTE: this dict is local to each collect_bom_rows() call, so it is
    # discarded after the traversal and never shared across invocations.
    _props_cache: dict[str, dict] = {}

    def _traverse(product, rows: list, level: int, parent_filepath: str = "") -> None:
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

        filepath  = get_product_filepath(product)
        not_found = not bool(filepath)

        # Use the cache to skip DESIGN_MODE + property reads for repeated files.
        cached = bool(filepath) and filepath in _props_cache
        is_readable = True

        if not cached:
            try:
                product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
            except Exception:
                is_readable = False

            props: dict = {}
            for col in columns:
                if col in DIRECT_ATTR_MAP:
                    props[col] = _get_prop(product, col)
                elif col in custom_columns:
                    props[col] = _get_user_prop(product, col)
            props["_is_readable"] = is_readable

            if filepath:
                _props_cache[filepath] = props
        else:
            props       = _props_cache[filepath]
            is_readable = bool(props.get("_is_readable", True))

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
            if col in DIRECT_ATTR_MAP or col in custom_columns:
                row[col] = props.get(col, "")

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


def flatten_bom_to_summary(
    rows: list[dict],
    include_assemblies: bool = False,
    sort_column: str | None = None,
) -> list[dict]:
    """Collapse a hierarchical BOM into a flat summary BOM.

    Each unique part (identified by its backing filepath when available, or by
    Part Number otherwise) appears exactly once in the result.  The
    ``Quantity`` value is the *total* count across the whole assembly tree,
    computed by multiplying the per-level quantities along every path from the
    root to that part.

    The root row (Level == 0) is always excluded from the result.

    Parameters
    ----------
    rows:
        The hierarchical BOM rows returned by :func:`collect_bom_rows`.
    include_assemblies:
        When ``False`` (default) rows whose ``Type`` is ``"产品"`` or
        ``"部件"`` are omitted so that only leaf parts appear in the summary.
        Set to ``True`` to include sub-assemblies and assemblies as well.
    sort_column:
        Internal column name to sort the result by.  Sorting is case-
        insensitive string comparison.  Defaults to ``"Part Number"`` when
        ``None``.

    Returns
    -------
    list[dict]
        Flat list of row dicts.  Each dict contains the same keys as the input
        rows except that ``Level`` is removed and ``Quantity`` reflects the
        total accumulated count.
    """
    if not rows:
        return []

    # ── Step 1: compute absolute quantity for every row ──────────────────────
    # We walk the rows in traversal order.  A stack tracks (level, cum_qty)
    # for each ancestor on the current path.
    # cum_qty[i] = cumulative quantity multiplier up to and including level i.
    cum_qty_stack: list[tuple[int, int]] = []  # (level, cumulative_qty)

    # absolute_qtys[i] = total count of rows[i] in the whole assembly
    absolute_qtys: list[int] = []

    for row in rows:
        level = row.get("Level", 0)
        qty   = int(row.get("Quantity", 1) or 1)

        # Pop stack entries that belong to a sibling or a higher-level ancestor
        while cum_qty_stack and cum_qty_stack[-1][0] >= level:
            cum_qty_stack.pop()

        # Parent's cumulative multiplier (1 if this is the root)
        parent_cum = cum_qty_stack[-1][1] if cum_qty_stack else 1

        abs_qty = parent_cum * qty
        absolute_qtys.append(abs_qty)
        cum_qty_stack.append((level, abs_qty))

    # ── Step 2: deduplicate by filepath (preferred) or Part Number ───────────
    # Key: filepath if non-empty, else Part Number.
    # For each key we keep the first row's attributes and accumulate quantity.
    seen_order:  list[str]       = []   # insertion-ordered keys
    summary:     dict[str, dict] = {}   # key → merged row dict
    key_to_qty:  dict[str, int]  = {}   # key → accumulated total qty

    _assembly_types = {"产品", "部件"}

    for row, abs_qty in zip(rows, absolute_qtys):
        level = row.get("Level", 0)

        # Always skip the root assembly (level 0 – the top-level product itself)
        if level == 0:
            continue

        # Optionally skip sub-assemblies and assemblies
        if not include_assemblies and row.get("Type", "") in _assembly_types:
            continue

        fp  = row.get("_filepath", "")
        key = fp if fp else str(row.get("Part Number", ""))
        if not key:
            continue

        if key not in summary:
            seen_order.append(key)
            merged = {k: v for k, v in row.items() if k != "Level"}
            merged["Quantity"] = abs_qty
            summary[key]       = merged
            key_to_qty[key]    = abs_qty
        else:
            key_to_qty[key]          += abs_qty
            summary[key]["Quantity"]  = key_to_qty[key]

    # ── Step 3: sort and return ───────────────────────────────────────────────
    result    = [summary[k] for k in seen_order]
    sort_key  = sort_column if sort_column else "Part Number"
    result.sort(key=lambda r: str(r.get(sort_key, "")).lower())
    return result
