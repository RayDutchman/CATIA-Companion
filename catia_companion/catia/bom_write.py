"""
BOM write-back to CATIA.

Provides:
- write_bom_to_catia() – traverse the product tree and write edited properties
                         back via COM
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
    progress_callback: Callable[[int], None] | None = None,
) -> None:
    """Write edited BOM properties back to CATIA via COM.

    Parameters
    ----------
    file_path:
        Path to the ``.CATProduct`` file that was edited, or ``None`` to use
        the currently active CATIA document without opening or saving anything.
    pn_data:
        Mapping from original Part Number → ``{column_name: new_value}``.
        Only changed fields need to be included.
    custom_columns:
        Column names that are user-defined properties (written via
        ``UserRefProperties``).
    progress_callback:
        Optional callable invoked with the current node count after each node
        is visited during the traversal.  May raise an exception to abort.
        The traversal is post-order (children before parents), so deeper
        levels are written to CATIA before their parent levels.
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

    _total_count: list[int] = [0]

    # Track backing filepaths that have already been written so that repeated
    # instances of the same physical document (e.g. the same fastener used
    # 50 times) are skipped together with their entire sub-tree.  This mirrors
    # the _props_cache optimization in collect_bom_rows and keeps the write-back
    # node count consistent with the read node count.
    # NOTE: nodes without a filepath (embedded sub-assemblies / 部件) are
    # always processed because they share the parent file but may represent
    # structurally distinct sub-trees.
    _written_fps: set[str] = set()

    # Mutable copy of the dirty-PN set used for early-exit: once every dirty
    # part has been written we can stop traversing the rest of the tree.
    remaining_pns: set[str] = set(pn_data.keys())

    def _traverse_write(product, parent_filepath: str = "") -> None:
        # Early exit: nothing left to write.
        if not remaining_pns:
            return

        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

        # Resolve the backing filepath for this node.
        filepath = ""
        for accessor in (
            lambda p: p.reference_product.com_object.Parent.FullName,
            lambda p: p.com_object.ReferenceProduct.Parent.FullName,
            lambda p: p.com_object.Parent.FullName,
        ):
            try:
                filepath = accessor(product)
                break
            except Exception:
                pass

        # A node is an embedded 部件 (no own file) when its resolved filepath
        # is identical to its parent's filepath – the same logic used by
        # collect_bom_rows to set Type=="部件".  Such nodes must NOT be
        # de-duplicated via _written_fps: all siblings under the same 组件
        # resolve to the same parent path, so only the first one would ever
        # be visited if the guard were applied to them.
        _is_own_file = bool(filepath) and filepath != parent_filepath

        # If we have already processed this file (written its properties and
        # recursed into its children), skip the whole sub-tree.  Only apply
        # this guard for nodes that have their own backing file.
        if _is_own_file and filepath in _written_fps:
            return

        # ── Recurse into children FIRST (post-order / bottom-up) ────────────
        # This guarantees that deeper levels (e.g. level 6) are written to
        # CATIA before their parent levels (e.g. level 5).  The parent node's
        # PN remains in remaining_pns throughout child processing, so the
        # early-exit break inside the loop only fires when truly nothing is
        # left to write (i.e. the parent itself is also not dirty).
        try:
            count = product.products.count
            for i in range(1, count + 1):
                if not remaining_pns:
                    break
                try:
                    _traverse_write(product.products.item(i),
                                    parent_filepath=filepath)
                except Exception:
                    pass
        except Exception:
            pass

        if pn in pn_data:
            try:
                product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
            except Exception:
                pass
            for col, value in pn_data[pn].items():
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
            remaining_pns.discard(pn)

        _total_count[0] += 1
        if progress_callback is not None:
            progress_callback(_total_count[0])

        # Mark this filepath as done after its sub-tree has been fully
        # traversed so that future identical references are skipped entirely.
        # Only standalone-file nodes are recorded; embedded 部件 nodes must
        # not pollute the set with the parent document's path.
        if filepath and _is_own_file:
            _written_fps.add(filepath)

    # ── CATIA connection ────────────────────────────────────────────────────
    caa         = catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    if file_path is None:
        product_doc  = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product
        _traverse_write(root_product, parent_filepath="")
        logger.info("Write-back complete for active document (not saved)")
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
    _traverse_write(root_product, parent_filepath="")
    logger.info(
        f"Write-back complete for {src.name} "
        "(not saved; user must save manually in CATIA)"
    )
