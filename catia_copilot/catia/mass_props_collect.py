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
    """通过 CATIA SPA 工作台逐一测量零件各 Body 的质量特性。

    策略：仅使用 per-body 路径，通过 ``part_com.CreateReferenceFromObject(body)``
    创建引用后调用 ``spa.GetMeasurable(ref)``。

    当材质定义在零件（Part）层面而非单个 Body 层面时，对 Body 引用调用
    ``GetMass()`` 会抛出 CATIA COM 错误（body 不拥有材质属性）。
    此时回退为体积法：
      1. ``GetVolume(body_ref)``  →  该 Body 的体积（纯几何量，始终可用）
      2. ``GetCOG(body_ref)``     →  该 Body 的几何中心（等价于质心，密度均匀）
      3. 由零件整体推算密度：``density = part_mass / part_vol``
         （密度是材质常量，推算结果精确，与包络体无关）
      4. ``body_mass = density × body_vol``

    如此即可正确排除"包络体"：包络体对应的 Body 不被纳入质量求和，
    而零件整体数据仅用于推算密度，不直接贡献最终质量。

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
    # ── DEBUG: 记录传入对象的 COM 类型名 ─────────────────────────────────
    try:
        logger.debug(f"[SPA] doc_com 类型: {type(doc_com).__name__}")
    except Exception as e:
        logger.debug(f"[SPA] 无法获取 doc_com 类型: {e}")
    try:
        logger.debug(f"[SPA] part_com 类型: {type(part_com).__name__}")
    except Exception as e:
        logger.debug(f"[SPA] 无法获取 part_com 类型: {e}")

    try:
        spa = doc_com.GetWorkbench("SPAWorkbench")
        logger.debug(f"[SPA] GetWorkbench('SPAWorkbench') 成功，类型: {type(spa).__name__}")
    except Exception as e:
        logger.debug(f"[SPA] 无法获取 SPAWorkbench: {e}")
        return None

    # ── 获取 Bodies 集合 ──────────────────────────────────────────────────
    bodies = None
    body_count = 0
    try:
        bodies = part_com.Bodies
        body_count = bodies.Count
        logger.debug(f"[SPA] Bodies.Count = {body_count}")
    except Exception as e:
        logger.debug(f"[SPA] 无法访问 Bodies: {e}")

    if body_count == 0 or bodies is None:
        logger.debug("[SPA] 没有找到任何 Body，测量中止")
        return None

    # ── 预推算零件密度（材质定义在 Part 层时 per-body GetMass 失败的备用路径）
    # density = part_mass / part_volume，单位 kg/mm³
    # 仅用于推算密度，不直接参与最终质量累加，因此包络体不影响结果。
    _part_density: float | None = None
    try:
        meas_whole = spa.GetMeasurable(part_com)
        logger.debug(f"[SPA] 整体 GetMeasurable(part_com) 成功（仅用于密度推算）")
        _w_mass: float | None = None
        try:
            _w_mass = meas_whole.Mass
            logger.debug(f"[SPA] 整体 .Mass = {_w_mass}")
        except Exception as _me1:
            logger.debug(f"[SPA] 整体 .Mass 失败: {_me1}，尝试 GetMass()")
            try:
                _w_mass = meas_whole.GetMass()
                logger.debug(f"[SPA] 整体 GetMass() = {_w_mass}")
            except Exception as _me2:
                logger.debug(f"[SPA] 整体 GetMass() 也失败: {_me2}")
        _w_vol: float | None = None
        try:
            _w_vol_arr = [0.0]
            meas_whole.GetVolume(_w_vol_arr)
            _w_vol = _w_vol_arr[0]
            logger.debug(f"[SPA] 整体 GetVolume(arr) = {_w_vol}")
        except Exception as _ve1:
            logger.debug(f"[SPA] 整体 GetVolume(arr) 失败: {_ve1}，尝试无参")
            try:
                _w_vol = meas_whole.GetVolume()
                logger.debug(f"[SPA] 整体 GetVolume() 无参 = {_w_vol}")
            except Exception as _ve2:
                logger.debug(f"[SPA] 整体 GetVolume() 无参也失败: {_ve2}")
        if _w_mass and _w_mass > 0.0 and _w_vol and _w_vol > 0.0:
            _part_density = _w_mass / _w_vol
            logger.debug(f"[SPA] 推算零件密度: {_part_density:.6e} kg/mm³")
    except Exception as _de:
        logger.debug(f"[SPA] 整体测量（密度推算）失败: {_de}")

    total_mass = 0.0
    weighted_cog = [0.0, 0.0, 0.0]
    # 转动惯量（在零件原点处，零件局部坐标轴）
    I_at_origin = [[0.0] * 3 for _ in range(3)]

    def _get_cog(meas, label: str) -> list[float]:
        """尝试从 Measurable 读取重心坐标，返回 [x, y, z]（失败时返回全零）。"""
        try:
            cog_arr = [0.0, 0.0, 0.0]
            meas.GetCOG(cog_arr)
            logger.debug(f"[SPA]   {label} GetCOG(arr) = {cog_arr}")
            return list(cog_arr)
        except Exception as e1:
            logger.debug(f"[SPA]   {label} GetCOG(arr) 失败: {e1}，尝试无参")
        try:
            cog_result = meas.GetCOG()
            cog = list(cog_result) if cog_result else [0.0, 0.0, 0.0]
            logger.debug(f"[SPA]   {label} GetCOG() 无参 = {cog}")
            return cog
        except Exception as e2:
            logger.debug(f"[SPA]   {label} GetCOG() 无参也失败: {e2}")
        return [0.0, 0.0, 0.0]

    def _get_inertia_matrix(meas, label: str) -> list[float]:
        """尝试从 Measurable 读取转动惯量矩阵（9 元素），失败时返回全零。"""
        try:
            inertia_arr = [0.0] * 9
            meas.GetInertiaMatrix(inertia_arr)
            logger.debug(f"[SPA]   {label} GetInertiaMatrix(arr) = {inertia_arr}")
            return list(inertia_arr)
        except Exception as e1:
            logger.debug(f"[SPA]   {label} GetInertiaMatrix(arr) 失败: {e1}，尝试无参")
        try:
            inertia_result = meas.GetInertiaMatrix()
            inertia = list(inertia_result) if inertia_result else [0.0] * 9
            logger.debug(f"[SPA]   {label} GetInertiaMatrix() 无参 = {inertia}")
            return inertia
        except Exception as e2:
            logger.debug(f"[SPA]   {label} GetInertiaMatrix() 无参也失败: {e2}")
        return [0.0] * 9

    def _accumulate(mass: float, cog: list[float], inertia_9: list[float]):
        """将一个 Body 的质量特性累积到 total_mass / weighted_cog / I_at_origin。"""
        nonlocal total_mass
        I_cog = [
            [inertia_9[0], inertia_9[1], inertia_9[2]],
            [inertia_9[3], inertia_9[4], inertia_9[5]],
            [inertia_9[6], inertia_9[7], inertia_9[8]],
        ]
        r = cog
        r2 = r[0]**2 + r[1]**2 + r[2]**2
        for row in range(3):
            for col in range(3):
                delta = (1.0 if row == col else 0.0) * r2 - r[row] * r[col]
                I_at_origin[row][col] += I_cog[row][col] + mass * delta
        total_mass += mass
        for j in range(3):
            weighted_cog[j] += mass * cog[j]

    # ── 逐一测量各 Body ───────────────────────────────────────────────────
    for i in range(1, body_count + 1):
        try:
            body = bodies.Item(i)
            try:
                body_name = body.Name
            except Exception:
                body_name = f"Body_{i}"
            label = f"Body_{i}({body_name})"
            logger.debug(f"[SPA] {label} COM 类型: {type(body).__name__}")

            # 通过 CreateReferenceFromObject 创建正确类型的引用
            ref = None
            try:
                ref = part_com.CreateReferenceFromObject(body)
                logger.debug(f"[SPA]   {label} CreateReferenceFromObject 成功，类型: {type(ref).__name__}")
            except Exception as ref_err:
                logger.debug(f"[SPA]   {label} CreateReferenceFromObject 失败: {ref_err}")

            if ref is None:
                logger.debug(f"[SPA]   {label} ref 为 None，跳过")
                continue

            try:
                meas = spa.GetMeasurable(ref)
                logger.debug(f"[SPA]   {label} GetMeasurable(ref) 成功，类型: {type(meas).__name__}")
            except Exception as meas_err:
                logger.debug(f"[SPA]   {label} GetMeasurable(ref) 失败: {meas_err}")
                continue

            # ── 尝试直接读取质量（先 .Mass 属性，再 GetMass() 方法）─────
            mass: float | None = None
            try:
                mass = meas.Mass
                logger.debug(f"[SPA]   {label} .Mass = {mass}")
            except Exception as ma_err:
                logger.debug(f"[SPA]   {label} .Mass 失败: {ma_err}，尝试 GetMass()")
                try:
                    mass = meas.GetMass()
                    logger.debug(f"[SPA]   {label} GetMass() = {mass}")
                except Exception as gm_err:
                    logger.debug(f"[SPA]   {label} GetMass() 也失败（材质可能定义在 Part 层）: {gm_err}")

            if mass is None or mass <= 0.0:
                # ── 回退：体积 × 密度 ───────────────────────────────────
                logger.debug(f"[SPA]   {label} 尝试体积法回退")
                vol: float | None = None
                try:
                    vol_arr = [0.0]
                    meas.GetVolume(vol_arr)
                    vol = vol_arr[0]
                    logger.debug(f"[SPA]   {label} GetVolume(arr) = {vol}")
                except Exception as ve1:
                    logger.debug(f"[SPA]   {label} GetVolume(arr) 失败: {ve1}，尝试无参")
                    try:
                        vol = meas.GetVolume()
                        logger.debug(f"[SPA]   {label} GetVolume() 无参 = {vol}")
                    except Exception as ve2:
                        logger.debug(f"[SPA]   {label} GetVolume() 无参也失败: {ve2}")

                if vol is not None and vol > 0.0 and _part_density is not None and _part_density > 0.0:
                    mass = _part_density * vol
                    logger.debug(f"[SPA]   {label} 体积法质量 = {_part_density:.6e} × {vol:.6g} = {mass:.6g} kg")
                else:
                    logger.debug(
                        f"[SPA]   {label} 体积法也无法获取质量"
                        f"（vol={vol}, density={_part_density}），跳过"
                    )
                    continue

            # ── 读取重心 ───────────────────────────────────────────────
            cog = _get_cog(meas, label)

            # ── 读取转动惯量 ───────────────────────────────────────────
            # 当材质在 Part 层时，Body ref 的 GetInertiaMatrix 可能也失败；
            # 此时使用质量 × 零转动惯量（点质量近似）并依赖平行轴定理贡献。
            inertia_9 = _get_inertia_matrix(meas, label)

            _accumulate(mass, cog, inertia_9)

        except Exception as e:
            logger.debug(f"[SPA] 访问 Body {i} 失败: {e}")

    logger.debug(f"[SPA] 所有 Body 累计质量: {total_mass}")

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

    logger.debug(f"[SPA] 最终结果: weight={total_mass}, cog={part_cog}")
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

        # 判断节点类型 —— 依据支持文件扩展名而非子节点数量
        # （子节点数为0的 CATProduct 也应被识别为"产品"，而非"零件"）
        is_embedded = (bool(filepath) and bool(parent_filepath)
                       and filepath == parent_filepath)
        if not_found:
            node_type = ""
        elif is_embedded:
            node_type = "部件"
        else:
            ext = Path(filepath).suffix.lower()
            if ext == ".catpart":
                node_type = "零件"
            else:
                # .catproduct 或其他未知扩展名统一视为"产品"
                node_type = "产品"

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
