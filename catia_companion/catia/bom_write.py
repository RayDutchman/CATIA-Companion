"""
BOM write-back to CATIA.

Provides:
- write_bom_to_catia() – write edited properties back via COM, using a
                         filepath-based fast path when available and falling
                         back to a pruned tree traversal for embedded
                         sub-assemblies (部件) or parts whose backing file is
                         unknown.
"""

import logging
from collections.abc import Callable
from pathlib import Path

from catia_companion.constants import (
    BOM_READONLY_COLUMNS,
    SOURCE_FROM_DISPLAY,
)

logger = logging.getLogger(__name__)


def write_bom_to_catia(
    file_path: str | None,
    pn_data: dict[str, dict[str, str]],
    custom_columns: list[str],
    pn_to_filepath: dict[str, str] | None = None,
    progress_callback: Callable[[int], None] | None = None,
) -> None:
    """Write edited BOM properties back to CATIA via COM.

    Parameters
    ----------
    file_path:
        Path to the ``.CATProduct`` root assembly file, or ``None`` to use the
        currently active CATIA document.  Only needed when the tree-traversal
        fallback path is required (see *pn_to_filepath*).
    pn_data:
        Mapping from original Part Number → ``{column_name: new_value}``.
        Only changed fields need to be included.
    custom_columns:
        Column names that are user-defined properties (written via
        ``UserRefProperties``).
    pn_to_filepath:
        Optional mapping from original Part Number → absolute path of the
        backing CATIA document for that part.  When provided, parts are
        written by opening each file directly — O(dirty parts) — instead of
        traversing the whole product tree — O(all nodes).  Parts absent from
        this mapping (e.g. embedded sub-assemblies of type 部件) fall back to
        the tree-traversal path.
    progress_callback:
        Optional callable invoked with the running processed-item count after
        each part is written.  May raise an exception to abort.
    """
    from pycatia import catia, CatWorkModeType
    from pycatia.product_structure_interfaces.product_document import ProductDocument

    WRITABLE_DIRECT: dict[str, str] = {
        "Nomenclature": "nomenclature",
        "Revision":     "revision",
        "Definition":   "definition",
        "Source":       "source",
    }

    def _set_prop(product, name: str, value: str) -> None:
        attr = WRITABLE_DIRECT.get(name)
        if not attr:
            return
        targets: list = []
        try:
            targets.append(product.reference_product)
        except Exception:
            pass
        targets.append(product)
        for target in targets:
            try:
                setattr(
                    target,
                    attr,
                    int(SOURCE_FROM_DISPLAY.get(value, value))
                    if name == "Source" else value,
                )
                return
            except Exception:
                continue

    def _set_user_prop(product, name: str, value: str) -> None:
        targets: list = []
        try:
            targets.append(product.reference_product)
        except Exception:
            pass
        targets.append(product)
        # Try to update an existing property first
        for target in targets:
            try:
                target.user_ref_properties.item(name).value = value
                return
            except Exception:
                pass
        # Property does not exist – create it on the first available target
        for target in targets:
            try:
                target.user_ref_properties.create_string(name, value)
                return
            except Exception:
                continue

    def _apply_to_product(product, data: dict[str, str]) -> None:
        """Apply the changed fields in *data* to *product*."""
        try:
            product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
        except Exception:
            pass
        for col, value in data.items():
            if col in BOM_READONLY_COLUMNS:
                continue
            if col == "Part Number":
                try:
                    product.part_number = value
                except Exception:
                    try:
                        product.com_object.PartNumber = value
                    except Exception:
                        pass
            elif col in WRITABLE_DIRECT:
                _set_prop(product, col, value)
            elif col in custom_columns:
                _set_user_prop(product, col, value)

    _total_count: list[int] = [0]

    def _inc_progress() -> None:
        _total_count[0] += 1
        if progress_callback is not None:
            progress_callback(_total_count[0])

    # ── CATIA connection ────────────────────────────────────────────────────
    caa         = catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    def _find_open_doc(fp_path: Path):
        """Return the already-open CATIA document whose path matches *fp_path*."""
        for i in range(1, documents.count + 1):
            try:
                d = documents.item(i)
                if Path(d.full_name).resolve() == fp_path:
                    return d
            except Exception:
                pass
        return None

    # ── Fast path: write directly to each backing file ──────────────────────
    # Group dirty PNs by filepath so each document is opened at most once.
    handled_pns: set[str] = set()

    if pn_to_filepath:
        fp_to_pns: dict[str, list[str]] = {}
        for pn, fp in pn_to_filepath.items():
            if pn in pn_data and fp:
                fp_to_pns.setdefault(fp, []).append(pn)

        for fp, pns in fp_to_pns.items():
            try:
                fp_path = Path(fp).resolve()
                doc = _find_open_doc(fp_path)
                if doc is None:
                    documents.open(str(fp_path))
                    doc = _find_open_doc(fp_path)
                if doc is None:
                    logger.warning(
                        "无法找到或打开文件：%s，将通过树遍历处理 %s", fp, pns
                    )
                    continue

                product_doc  = ProductDocument(doc.com_object)
                root_product = product_doc.product

                for pn in pns:
                    _apply_to_product(root_product, pn_data[pn])
                    handled_pns.add(pn)
                    _inc_progress()

            except Exception as exc:
                logger.warning(
                    "直接写入 %s 失败（%s），将通过树遍历处理 %s", fp, exc, pns
                )

    # ── Fallback: pruned tree traversal for remaining PNs ───────────────────
    remaining: dict[str, dict[str, str]] = {
        pn: data for pn, data in pn_data.items() if pn not in handled_pns
    }

    if not remaining:
        logger.info(
            "Write-back complete via fast path (%d part(s) written)",
            len(handled_pns),
        )
        return

    # Mutable set used for early-exit pruning; entries are discarded as each
    # PN is written so that subtrees with no remaining targets are skipped.
    remaining_pns: set[str] = set(remaining)

    def _traverse_write(product) -> None:
        if not remaining_pns:
            return

        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

        if pn in remaining_pns:
            _apply_to_product(product, remaining[pn])
            remaining_pns.discard(pn)

        _inc_progress()

        if not remaining_pns:
            return

        try:
            count = product.products.count
            for i in range(1, count + 1):
                if not remaining_pns:
                    break
                try:
                    _traverse_write(product.products.item(i))
                except Exception:
                    pass
        except Exception:
            pass

    if file_path is None:
        product_doc  = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product
        _traverse_write(root_product)
        logger.info(
            "Write-back complete for active document "
            "(fast path: %d, traversal: %d)",
            len(handled_pns), len(remaining),
        )
        return

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
    _traverse_write(root_product)
    logger.info(
        "Write-back complete for %s (fast path: %d, traversal: %d)",
        src.name, len(handled_pns), len(remaining),
    )
