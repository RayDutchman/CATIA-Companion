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


def _position_to_mat4(product) -> list[list[float]]:
    """从 pycatia Product 包装对象的 position.get_components() 读取位置，返回 4×4 矩阵。

    接收 pycatia Product 包装对象（非 com_object），无参调用
    ``product.position.get_components()``，捕获返回的 12 个 double 分量。
    存储约定（列主序）：
      arr[0..2]  = 旋转矩阵第一列 (R[:,0])
      arr[3..5]  = 旋转矩阵第二列 (R[:,1])
      arr[6..8]  = 旋转矩阵第三列 (R[:,2])
      arr[9..11] = 平移向量 T

    变换关系：P_parent = R @ P_local + T

    若调用失败，返回 4×4 单位矩阵（视作零件位于父坐标系原点）。
    """
    try:
        product_name = getattr(product, "name", repr(product))
    except Exception:
        product_name = repr(product)

    try:
        arr = product.position.get_components()
    except Exception as e:
        logger.debug(f"[MAT4] {product_name}: get_components 失败: {e}，返回单位矩阵")
        return _identity_4x4()

    if arr is None or len(arr) < 12:
        logger.debug(
            f"[MAT4] {product_name}: arr 无效（len={len(arr) if arr else None}），返回单位矩阵"
        )
        return _identity_4x4()

    mat = [
        [arr[0], arr[3], arr[6], arr[9]],
        [arr[1], arr[4], arr[7], arr[10]],
        [arr[2], arr[5], arr[8], arr[11]],
        [0.0,    0.0,    0.0,    1.0   ],
    ]
    logger.debug(
        f"[MAT4] {product_name}: R[0]={mat[0][:3]}, T={[mat[0][3], mat[1][3], mat[2][3]]}"
    )
    return mat


# ---------------------------------------------------------------------------
# 质量特性读取辅助函数
# ---------------------------------------------------------------------------


def _try_mp_params(part_com, label: str = "") -> dict | None:
    """读取由 create_inertia_relations.catvbs 写入的 MP_* 用户参数。

    参数单位与程序内部单位一致（无需转换）：
      - MP_Mass_g                   → g（直接使用）
      - MP_COGx/y/z_mm              → mm（直接使用）
      - MP_Ixx/yy/zz/xy/xz/yz_gmm2 → g·mm²（直接使用）

    返回与 _measure_part_mass_props 相同结构的字典，或 None（参数不存在或质量≤0）。
    """
    tag = f"[MP] {label} " if label else "[MP] "
    try:
        params = part_com.Parameters

        def _get(name: str) -> float | None:
            try:
                return float(params.Item(name).Value)
            except Exception:
                return None

        mass_g = _get("MP_Mass_g")
        if mass_g is None or mass_g <= 0.0:
            logger.debug(f"{tag}MP_Mass_g 不存在或为零，跳过")
            return None

        cogx = _get("MP_COGx_mm") or 0.0
        cogy = _get("MP_COGy_mm") or 0.0
        cogz = _get("MP_COGz_mm") or 0.0
        ixx  = _get("MP_Ixx_gxmm2") or 0.0
        iyy  = _get("MP_Iyy_gxmm2") or 0.0
        izz  = _get("MP_Izz_gxmm2") or 0.0
        ixy  = _get("MP_Ixy_gxmm2") or 0.0
        ixz  = _get("MP_Ixz_gxmm2") or 0.0
        iyz  = _get("MP_Iyz_gxmm2") or 0.0

        return {
            "weight":  mass_g,
            "cog":     [cogx, cogy, cogz],
            "inertia": [
                [ixx, ixy, ixz],
                [ixy, iyy, iyz],
                [ixz, iyz, izz],
            ],
        }
    except Exception as e:
        logger.debug(f"{tag}MP_* 参数读取异常: {e}")
        return None


def _run_inertia_vbs_and_read(
    doc_com, part_com, label: str = "", part_number: str = ""
) -> dict | None:
    """若 MP_* 参数不存在，自动运行 create_inertia_relations.catvbs，
    再读取写入的 MP_* 参数。

    VBS 脚本路径：``macros/create_inertia_relations.catvbs``（相对于项目根）。
    若零件没有 "惯量包络体.1\\质量"（Keep 测量）参数，脚本会静默退出，不弹框。

    参数：
        doc_com:     COM 对象（PartDocument 层）。
        part_com:    COM 对象（Part 层），用于读取 MP_* 参数。
        label:       调试标签，用于 debug 日志。
        part_number: 零件号（PartNumber），作为 iParameter[0] 传给 VBS，
                     VBS 按此在已打开文档中定位并激活目标零件文档。
                     为空时 VBS 使用当前活动文档。
    """
    tag = f"[MP] {label} " if label else "[MP] "

    try:
        from catia_copilot.utils import resource_path
        vbs_path = resource_path("macros/create_inertia_relations.catvbs")
        if not vbs_path.is_file():
            logger.debug(f"{tag}找不到 VBS 文件: {vbs_path}，跳过")
            return None

        logger.debug(f"{tag}激活零件文档并运行 {vbs_path.name}，part_number={part_number!r}")
        doc_com.Activate()
        # 将零件号作为 iParameter 传给 VBS，让脚本按 PartNumber 定位目标文档。
        # 使用 PartNumber 而非文件路径，可避免受零件未保存（路径无效）的影响。
        # ExecuteScript 参数：(脚本目录, 语言类型=1表示VBScript, 脚本文件名, 入口Sub名, 参数数组)
        doc_com.Application.SystemService.ExecuteScript(
            str(vbs_path.parent), 1, vbs_path.name, "CATMain", [part_number]
        )

        result = _try_mp_params(part_com, label)
        if result is not None:
            logger.debug(f"{tag}VBS 执行后 MP_* 参数读取成功")
        else:
            logger.debug(f"{tag}VBS 执行后 MP_* 参数仍不可用（零件可能无 Keep 惯量测量）")
        return result
    except Exception as e:
        logger.debug(f"{tag}VBS 路径失败: {e}")
        return None


def _measure_part_mass_props(doc_com, part_com, part_number: str = "") -> dict | None:
    """测量零件质量特性。

    所有返回值均使用 **g / mm / g·mm²** 单位制（与 create_inertia_relations.catvbs 一致）。

    路径优先级：
      1. **MP_* 参数**：直接读取由 ``create_inertia_relations.catvbs`` 写入的
         ``MP_Mass_g``、``MP_COGx/y/z_mm``、``MP_Ixx/yy/zz/xy/xz/yz_gmm2`` 参数。
      2. **VBS 自动绑定**：若路径 1 失败，自动运行 VBS 脚本尝试创建 MP_* 参数后
         再读取。若零件无 Keep 惯量测量，脚本会静默退出，不阻塞。

    参数：
        doc_com:     COM 对象（PartDocument 层）。
        part_com:    COM 对象（Part 层）。
        part_number: 零件号（PartNumber），传给 VBS 脚本以定位目标文档。

    返回字典：
      {
        "weight":  float,          # 总质量，g
        "cog":     [x, y, z],      # 重心坐标（零件局部坐标系），mm
        "inertia": [[Ixx,Ixy,Ixz],
                    [Iyx,Iyy,Iyz],
                    [Izx,Izy,Izz]], # 重心处转动惯量（零件局部坐标轴），g·mm²
      }
    若所有路径均失败则返回 None。
    """
    # ── 路径 1：读取 MP_* 用户参数（由 create_inertia_relations.catvbs 写入）──────
    _mp = _try_mp_params(part_com, "直接读取")
    if _mp is not None:
        return _mp

    # ── 路径 2：自动运行 VBS 绑定脚本（要求零件已有 Keep 测量）─────────────────────
    _mp = _run_inertia_vbs_and_read(doc_com, part_com, "VBS绑定", part_number)
    if _mp is not None:
        return _mp

    return None


# ---------------------------------------------------------------------------
# Post-processing helpers
# ---------------------------------------------------------------------------


def _post_process_rows(rows: list[dict]) -> None:
    """对遍历后的行列表进行两轮后处理，使显示字段反映根产品坐标系。

    第一轮（零件行）
        利用 ``_placement``（零件局部→根的变换矩阵）将局部坐标系的重心和
        转动惯量旋转变换到根坐标系，更新 ``CogX/Y/Z`` 及 ``Ixx``-``Iyz``
        显示字段，并将变换结果缓存到 ``_root_mp`` 中供父级汇总和编辑后
        回写使用。

    第二轮（产品/部件行）
        收集该节点子树内所有零件的根坐标系质量特性，按标准刚体力学汇总
        （平行轴定理），计算该节点在根坐标系下的总质量、总重心和总转动
        惯量，并更新到显示字段中。

        若子树内所有零件均测量失败，则该节点的显示字段保持为 ``None``
        （显示为 "—"）。
    """
    n = len(rows)

    # ── 第一轮：零件行 → 变换到根坐标系 ─────────────────────────────────────
    for row in rows:
        if row.get("Type") != "零件":
            continue
        mp = row.get("_mass_props")
        if not mp:
            continue
        placement = row.get("_placement")
        if placement is None:
            continue

        R = [[placement[i][j] for j in range(3)] for i in range(3)]
        T = [placement[0][3], placement[1][3], placement[2][3]]

        # 重心变换：r_root = R @ r_local + T
        cog_local = mp.get("cog", [0.0, 0.0, 0.0])
        cog_root  = [
            sum(R[i][k] * cog_local[k] for k in range(3)) + T[i]
            for i in range(3)
        ]

        # 惯量旋转：I_root = R @ I_local @ R^T（在重心处，仅旋转轴方向）
        I_local = mp.get("inertia", [[0.0] * 3 for _ in range(3)])
        RT = [[R[j][i] for j in range(3)] for i in range(3)]
        RI = [
            [sum(R[i][k] * I_local[k][j] for k in range(3)) for j in range(3)]
            for i in range(3)
        ]
        I_root = [
            [sum(RI[i][k] * RT[k][j] for k in range(3)) for j in range(3)]
            for i in range(3)
        ]

        # 更新显示字段为根坐标系值
        row["CogX"] = cog_root[0]
        row["CogY"] = cog_root[1]
        row["CogZ"] = cog_root[2]
        row["Ixx"]  = I_root[0][0]
        row["Iyy"]  = I_root[1][1]
        row["Izz"]  = I_root[2][2]
        row["Ixy"]  = I_root[0][1]
        row["Ixz"]  = I_root[0][2]
        row["Iyz"]  = I_root[1][2]

        # 缓存根坐标系数据，供父级汇总及编辑后重新写入使用
        row["_root_mp"] = {
            "weight":  mp.get("weight", 0.0),
            "cog":     cog_root,
            "inertia": I_root,
        }

    # ── 第二轮：产品/部件行 → 汇总子孙零件 ───────────────────────────────────
    for i in range(n):
        row = rows[i]
        if row.get("Type") not in ("产品", "部件"):
            continue

        level = int(row.get("Level", 0))

        # 收集当前节点子树内所有已成功测量的零件的根坐标系质量特性
        child_parts: list[dict] = []
        for j in range(i + 1, n):
            desc = rows[j]
            if int(desc.get("Level", 0)) <= level:
                break  # 已超出子树范围
            rmp = desc.get("_root_mp")
            if rmp and float(rmp.get("weight", 0.0)) > 0.0:
                child_parts.append(rmp)

        if not child_parts:
            continue

        # 汇总（与 rollup_mass_properties 算法一致）
        M_total   = 0.0
        sum_mr    = [0.0, 0.0, 0.0]
        I_at_orig = [[0.0] * 3 for _ in range(3)]

        for rmp in child_parts:
            m  = float(rmp.get("weight", 0.0))
            if m <= 0.0:
                continue
            r  = rmp.get("cog", [0.0, 0.0, 0.0])
            Ic = rmp.get("inertia", [[0.0] * 3 for _ in range(3)])
            # 平行轴定理：从零件重心移到根原点
            r2 = sum(r[k] ** 2 for k in range(3))
            for ii in range(3):
                for jj in range(3):
                    delta = (1.0 if ii == jj else 0.0) * r2 - r[ii] * r[jj]
                    I_at_orig[ii][jj] += Ic[ii][jj] + m * delta
            M_total += m
            for k in range(3):
                sum_mr[k] += m * r[k]

        if M_total <= 0.0:
            continue

        cog_total = [sum_mr[k] / M_total for k in range(3)]

        # 平行轴定理：从根原点移回汇总重心
        rc  = cog_total
        rc2 = sum(rc[k] ** 2 for k in range(3))
        I_final = [[0.0] * 3 for _ in range(3)]
        for ii in range(3):
            for jj in range(3):
                delta = (1.0 if ii == jj else 0.0) * rc2 - rc[ii] * rc[jj]
                I_final[ii][jj] = I_at_orig[ii][jj] - M_total * delta

        row["Weight"] = M_total
        row["CogX"]   = cog_total[0]
        row["CogY"]   = cog_total[1]
        row["CogZ"]   = cog_total[2]
        row["Ixx"]    = I_final[0][0]
        row["Iyy"]    = I_final[1][1]
        row["Izz"]    = I_final[2][2]
        row["Ixy"]    = I_final[0][1]
        row["Ixz"]    = I_final[0][2]
        row["Iyz"]    = I_final[1][2]


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
        local_mat4 = _position_to_mat4(product)
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
                        mass_props    = _measure_part_mass_props(part_doc_com, part_com, pn)
                    except Exception as e:
                        logger.debug(f"无法测量零件 {filepath}: {e}")
                        mass_props  = None
                        spa_failed  = True

                    # ── DEBUG: 记录每个零件的测量结果概要 ─────────────────────
                    if mass_props is not None:
                        _mp_dbg = mass_props
                        logger.debug(
                            f"[TRAV] {pn} 测量成功: "
                            f"weight={_mp_dbg.get('weight')}g, "
                            f"cog={[round(v,3) for v in _mp_dbg.get('cog',[0,0,0])]}, "
                            f"Ixx={_mp_dbg.get('inertia',[[0]])[0][0]:.3g}g·mm²"
                        )
                    else:
                        logger.debug(f"[TRAV] {pn} 所有测量路径均失败")

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
        _post_process_rows(rows)
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
    _post_process_rows(rows)
    return rows
