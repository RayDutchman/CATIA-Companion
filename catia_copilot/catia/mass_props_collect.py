"""
质量特性数据收集模块。

提供：
- collect_mass_props_rows() – 遍历产品树，读取每个零件实例的质量/重心/转动惯量，
                              不对兄弟零件进行数量合并（每个实例单独记录一行）。
"""

import logging
from collections.abc import Callable
from pathlib import Path

from catia_copilot.constants import FILENAME_NOT_FOUND

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pure-Python 4×4 matrix helpers (no numpy dependency)
# ---------------------------------------------------------------------------

def _identity_4x4() -> list[list[float]]:
    """返回 4×4 单位矩阵。"""
    return [
        [1.0, 0.0, 0.0, 0.0],
        [0.0, 1.0, 0.0, 0.0],
        [0.0, 0.0, 1.0, 0.0],
        [0.0, 0.0, 0.0, 1.0],
    ]


def _mat4_mul(A: list[list[float]], B: list[list[float]]) -> list[list[float]]:
    """4×4 矩阵乘法，返回 A @ B。"""
    C = [[0.0] * 4 for _ in range(4)]
    for i in range(4):
        for j in range(4):
            for k in range(4):
                C[i][j] += A[i][k] * B[k][j]
    return C


def _position_to_mat4(product_com) -> list[list[float]]:
    """从 product.com_object.Position.GetComponents() 读取位置，返回 4×4 矩阵。

    CATIA Position.GetComponents 返回 12 个分量，存储约定（列主序）：
      [0..2]  = 旋转矩阵第一列 (R[:,0])
      [3..5]  = 旋转矩阵第二列 (R[:,1])
      [6..8]  = 旋转矩阵第三列 (R[:,2])
      [9..11] = 平移向量 T

    变换关系：P_parent = R @ P_local + T
    """
    try:
        components = [0.0] * 12
        product_com.Position.GetComponents(components)
    except Exception:
        try:
            result = product_com.Position.GetComponents()
            if hasattr(result, '__len__') and len(result) == 12:
                components = list(result)
            else:
                return _identity_4x4()
        except Exception:
            return _identity_4x4()

    return [
        [components[0], components[3], components[6], components[9]],
        [components[1], components[4], components[7], components[10]],
        [components[2], components[5], components[8], components[11]],
        [0.0,           0.0,           0.0,           1.0          ],
    ]


# ---------------------------------------------------------------------------
# SPA measurement helpers
# ---------------------------------------------------------------------------

def _measure_part_mass_props(doc_com, part_com) -> dict | None:
    """通过 CATIA SPA 工作台测量零件的质量特性。

    策略（按优先级）：
    1. 直接测量整个 Part（最可靠；适合单体和多体零件）。
    2. 若整体测量失败，逐一测量各 Body（通过 CreateReferenceFromObject 创建引用）。

    返回字典：
      {
        "weight":  float,          # 总质量，kg
        "cog":     [x, y, z],      # 重心坐标（零件局部坐标系），mm
        "inertia": [[Ixx,Ixy,Ixz],
                    [Iyx,Iyy,Iyz],
                    [Izx,Izy,Izz]], # 重心处转动惯量（零件局部坐标轴），kg·mm²
      }
    若测量失败则返回 None。
    """
    try:
        spa = doc_com.GetWorkbench("SPAWorkbench")
    except Exception as e:
        logger.debug(f"无法获取 SPAWorkbench: {e}")
        return None

    total_mass = 0.0
    weighted_cog = [0.0, 0.0, 0.0]
    # 转动惯量（在零件原点处，零件局部坐标轴）
    I_at_origin = [[0.0] * 3 for _ in range(3)]

    def _accumulate_body(measurable_target):
        nonlocal total_mass
        # 质量
        try:
            mass = measurable_target.GetMass()
        except Exception:
            return

        if mass is None or mass <= 0.0:
            return

        # 重心
        try:
            cog_result = [0.0, 0.0, 0.0]
            measurable_target.GetCOG(cog_result)
            cog = cog_result
        except Exception:
            try:
                cog_result = measurable_target.GetCOG()
                cog = list(cog_result) if cog_result else [0.0, 0.0, 0.0]
            except Exception:
                cog = [0.0, 0.0, 0.0]

        # 转动惯量（在重心处，SPA 默认行为）
        try:
            inertia_arr = [0.0] * 9
            measurable_target.GetInertiaMatrix(inertia_arr)
            inertia = inertia_arr
        except Exception:
            try:
                inertia_result = measurable_target.GetInertiaMatrix()
                inertia = list(inertia_result) if inertia_result else [0.0] * 9
            except Exception:
                inertia = [0.0] * 9

        # 将 9 元素行主序数组转为 3×3 矩阵
        I_cog = [
            [inertia[0], inertia[1], inertia[2]],
            [inertia[3], inertia[4], inertia[5]],
            [inertia[6], inertia[7], inertia[8]],
        ]

        # 平行轴定理：将转动惯量从重心移到零件原点
        r = cog  # [x, y, z]
        r2 = r[0]**2 + r[1]**2 + r[2]**2
        # I_origin = I_cog + m*(r²·E - r⊗r)
        for row in range(3):
            for col in range(3):
                delta = (1.0 if row == col else 0.0) * r2 - r[row] * r[col]
                I_at_origin[row][col] += I_cog[row][col] + mass * delta

        # 累计质量和加权重心
        total_mass += mass
        for j in range(3):
            weighted_cog[j] += mass * cog[j]

    # ── 策略1：直接测量整个 Part ─────────────────────────────────────────
    whole_part_ok = False
    try:
        meas = spa.GetMeasurable(part_com)
        _accumulate_body(meas)
        if total_mass > 0.0:
            whole_part_ok = True
    except Exception as e:
        logger.debug(f"无法直接测量 Part: {e}")

    # ── 策略2：逐一测量 Bodies（整体测量失败时的备用）──────────────────
    if not whole_part_ok:
        body_count = 0
        bodies = None
        try:
            bodies = part_com.Bodies
            body_count = bodies.Count
        except Exception as e:
            logger.debug(f"无法访问 Bodies: {e}")

        if body_count > 0 and bodies is not None:
            for i in range(1, body_count + 1):
                try:
                    body = bodies.Item(i)
                    # 优先通过 CreateReferenceFromObject 获取正确类型
                    try:
                        ref = part_com.CreateReferenceFromObject(body)
                        meas = spa.GetMeasurable(ref)
                    except Exception as ref_err:
                        logger.debug(f"CreateReferenceFromObject/GetMeasurable(ref) 失败: {ref_err}，改用直接传入 Body")
                        meas = spa.GetMeasurable(body)
                    _accumulate_body(meas)
                except Exception as e:
                    logger.debug(f"无法测量 Body {i}: {e}")

    if total_mass <= 0.0:
        return {
            "weight": 0.0,
            "cog":    [0.0, 0.0, 0.0],
            "inertia": [[0.0] * 3 for _ in range(3)],
        }

    # 计算零件重心（局部坐标系）
    part_cog = [weighted_cog[j] / total_mass for j in range(3)]

    # 平行轴定理：将转动惯量从原点移回重心
    r = part_cog
    r2 = r[0]**2 + r[1]**2 + r[2]**2
    I_at_cog = [[0.0] * 3 for _ in range(3)]
    for row in range(3):
        for col in range(3):
            delta = (1.0 if row == col else 0.0) * r2 - r[row] * r[col]
            I_at_cog[row][col] = I_at_origin[row][col] - total_mass * delta

    return {
        "weight":  total_mass,
        "cog":     part_cog,
        "inertia": I_at_cog,
    }


# ---------------------------------------------------------------------------
# Main collection function
# ---------------------------------------------------------------------------

def collect_mass_props_rows(
    file_path: str | None,
    progress_callback: Callable[[int], None] | None = None,
) -> list[dict]:
    """遍历产品树，返回每个节点的质量特性行列表。

    与 collect_bom_rows() 的关键区别：
      - **不对兄弟零件去重**——每个实例单独输出一行。
      - 仅对类型为"零件"的叶子节点调用 SPA 测量；部件/产品节点跳过测量。
      - 每行额外包含 ``_placement`` 字段（4×4 列表），为该实例到根坐标系的变换矩阵。

    参数：
        file_path:
            ``.CATProduct`` 文件路径，或 ``None`` 表示使用当前 CATIA 活动文档。
        progress_callback:
            可选回调，每追加一行后调用，传入当前行数。可通过抛出异常中止遍历。

    返回：
        行字典列表，每行含以下键：
          Level, Type, Part Number, Filename, Nomenclature, Revision,
          Weight, CogX, CogY, CogZ, Ixx, Iyy, Izz, Ixy, Ixz, Iyz,
          _filepath, _placement, _not_found, _unreadable, _spa_failed
    """
    from pycatia import catia, CatWorkModeType
    from pycatia.product_structure_interfaces.product_document import ProductDocument

    _total_count: int = 0
    # 以文件路径为键缓存质量特性测量结果，避免同一零件多实例重复测量
    _mass_cache: dict[str, dict] = {}

    def _get_prop(product, name: str) -> str:
        """读取直接属性（Nomenclature / Revision）。"""
        attr_map = {"Nomenclature": "nomenclature", "Revision": "revision"}
        attr = attr_map.get(name)
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

    def _traverse(
        product,
        rows: list,
        level: int,
        parent_filepath: str,
        parent_mat4: list[list[float]],
        documents,
    ) -> None:
        nonlocal _total_count

        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

        # 解析文件路径
        try:
            filepath = product.com_object.ReferenceProduct.Parent.FullName
        except Exception:
            filepath = ""

        not_found = not bool(filepath)

        # 判断节点类型
        try:
            child_count = product.products.count
        except Exception:
            child_count = 0

        is_embedded = (bool(filepath) and bool(parent_filepath)
                       and filepath == parent_filepath)
        if not_found:
            node_type = ""
        elif is_embedded:
            node_type = "部件"
        elif child_count > 0:
            node_type = "产品"
        else:
            node_type = "零件"

        # 计算本节点到根的累积变换矩阵
        local_mat4 = _position_to_mat4(product.com_object)
        abs_mat4   = _mat4_mul(parent_mat4, local_mat4)

        # 读取属性（仅 Nomenclature 和 Revision）
        is_readable = True
        nomenclature = ""
        revision     = ""

        if not not_found:
            try:
                current_mode = product.get_work_mode()
                if current_mode != CatWorkModeType.DESIGN_MODE:
                    product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
            except Exception:
                try:
                    product.apply_work_mode(CatWorkModeType.DESIGN_MODE)
                except Exception:
                    is_readable = False

            if is_readable:
                nomenclature = _get_prop(product, "Nomenclature")
                revision     = _get_prop(product, "Revision")

        # SPA 测量（仅叶子零件）
        mass_props: dict | None = None
        spa_failed = False

        if node_type == "零件" and is_readable and filepath:
            if filepath in _mass_cache:
                mass_props = _mass_cache[filepath]
            else:
                # 查找对应的文档
                target_doc = None
                fp_resolved = Path(filepath).resolve()
                for i in range(1, documents.count + 1):
                    try:
                        doc = documents.item(i)
                        if Path(doc.full_name).resolve() == fp_resolved:
                            target_doc = doc
                            break
                    except Exception:
                        pass

                if target_doc is not None:
                    try:
                        # 尝试获取 Part 对象（PartDocument 接口）
                        part_doc_com  = target_doc.com_object
                        part_com      = part_doc_com.Part
                        mass_props    = _measure_part_mass_props(part_doc_com, part_com)
                    except Exception as e:
                        logger.debug(f"无法测量零件 {filepath}: {e}")
                        mass_props  = None
                        spa_failed  = True

                    _mass_cache[filepath] = mass_props
                else:
                    logger.debug(f"找不到已打开的文档: {filepath}")
                    spa_failed = True

        if mass_props is None:
            spa_failed = spa_failed or (node_type == "零件" and is_readable and not not_found)

        # 组装行数据
        mp = mass_props or {}
        cog = mp.get("cog", [0.0, 0.0, 0.0])
        inertia = mp.get("inertia", [[0.0]*3 for _ in range(3)])

        row: dict = {
            "Level":        level,
            "Type":         node_type,
            "Part Number":  pn,
            "Filename":     Path(filepath).stem if filepath else FILENAME_NOT_FOUND,
            "Nomenclature": nomenclature,
            "Revision":     revision,
            "Weight":       mp.get("weight", None),
            "CogX":         cog[0] if mp else None,
            "CogY":         cog[1] if mp else None,
            "CogZ":         cog[2] if mp else None,
            "Ixx":          inertia[0][0] if mp else None,
            "Iyy":          inertia[1][1] if mp else None,
            "Izz":          inertia[2][2] if mp else None,
            "Ixy":          inertia[0][1] if mp else None,
            "Ixz":          inertia[0][2] if mp else None,
            "Iyz":          inertia[1][2] if mp else None,
            "_filepath":    filepath,
            "_placement":   abs_mat4,
            "_not_found":   not_found,
            "_unreadable":  not is_readable,
            "_spa_failed":  spa_failed,
            "_mass_props":  mass_props,  # 原始测量值，供联动修改时使用
        }

        rows.append(row)
        _total_count += 1
        if progress_callback is not None:
            progress_callback(_total_count)

        # 递归子节点（不去重，每个实例单独遍历）
        try:
            count = product.products.count
            for i in range(1, count + 1):
                try:
                    child = product.products.item(i)
                    _traverse(child, rows, level + 1,
                              parent_filepath=filepath,
                              parent_mat4=abs_mat4,
                              documents=documents)
                except Exception as e:
                    logger.debug(f"遍历子节点 {i} 失败: {e}")
        except Exception:
            pass

    # ── CATIA 连接 ──────────────────────────────────────────────────────────
    caa         = catia()
    application = caa.application
    application.visible = True
    documents   = application.documents

    if file_path is None:
        product_doc  = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product
        rows: list[dict] = []
        _traverse(root_product, rows, level=0, parent_filepath="",
                  parent_mat4=_identity_4x4(), documents=documents)
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
    _traverse(root_product, rows, level=0, parent_filepath="",
              parent_mat4=_identity_4x4(), documents=documents)
    return rows
