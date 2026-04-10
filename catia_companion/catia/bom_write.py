"""
BOM write-back to CATIA.

Provides:
- write_bom_to_catia() – traverse the product tree and write edited properties
                         back via COM
"""

import logging
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
    """
    from catia_companion.catia.connection import connect_to_catia
    from pycatia import CatWorkModeType
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

    def _traverse_write(product) -> None:
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

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

        try:
            count = product.products.count
            for i in range(1, count + 1):
                try:
                    _traverse_write(product.products.item(i))
                except Exception:
                    pass
        except Exception:
            pass

    # ── CATIA connection ────────────────────────────────────────────────────
    caa         = connect_to_catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    if file_path is None:
        product_doc  = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product
        _traverse_write(root_product)
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
    _traverse_write(root_product)
    logger.info(
        f"Write-back complete for {src.name} "
        "(not saved; user must save manually in CATIA)"
    )
