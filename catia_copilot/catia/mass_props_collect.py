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
        ixx  = _get("MP_Ixx_gmm2") or 0.0
        iyy  = _get("MP_Iyy_gmm2") or 0.0
        izz  = _get("MP_Izz_gmm2") or 0.0
        ixy  = _get("MP_Ixy_gmm2") or 0.0
        ixz  = _get("MP_Ixz_gmm2") or 0.0
        iyz  = _get("MP_Iyz_gmm2") or 0.0

        logger.debug(
            f"{tag}MP_* 参数读取成功: "
            f"mass={mass_g}g, cog=({cogx},{cogy},{cogz})mm"
        )
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
    doc_com, part_com, label: str = "", filepath: str = ""
) -> dict | None:
    """若 MP_* 参数不存在，自动运行 create_inertia_relations.catvbs，
    再读取写入的 MP_* 参数。

    VBS 脚本路径：``macros/create_inertia_relations.catvbs``（相对于项目根）。
    若零件没有 "惯量包络体.1\\质量"（Keep 测量）参数，脚本会静默退出，不弹框。

    参数：
        doc_com:  COM 对象（PartDocument 层）。
        part_com: COM 对象（Part 层），用于读取 MP_* 参数。
        label:    调试标签，用于 debug 日志。
        filepath: 零件文档的磁盘路径，作为 iParameter[0] 传给 VBS，
                  VBS 可据此定位并激活目标文档；为空时 VBS 使用当前活动文档。
    """
    tag = f"[MP] {label} " if label else "[MP] "

    try:
        from catia_copilot.utils import resource_path
        vbs_path = resource_path("macros/create_inertia_relations.catvbs")
        if not vbs_path.is_file():
            logger.debug(f"{tag}找不到 VBS 文件: {vbs_path}，跳过")
            return None

        logger.debug(f"{tag}激活零件文档并运行 {vbs_path.name}，filepath={filepath!r}")
        doc_com.Activate()
        # 将零件文件路径作为 iParameter 传给 VBS，让脚本可自行定位目标文档
        doc_com.Application.SystemService.ExecuteScript(
            str(vbs_path.parent), 1, vbs_path.name, "CATMain", [filepath]
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


def _measure_part_mass_props(doc_com, part_com, filepath: str = "") -> dict | None:
    """测量零件质量特性。

    所有返回值均使用 **g / mm / g·mm²** 单位制（与 create_inertia_relations.catvbs 一致）。

    路径优先级：
      1. **MP_* 参数**：直接读取由 ``create_inertia_relations.catvbs`` 写入的
         ``MP_Mass_g``、``MP_COGx/y/z_mm``、``MP_Ixx/yy/zz/xy/xz/yz_gmm2`` 参数。
      2. **VBS 自动绑定**：若路径 1 失败，自动运行 VBS 脚本尝试创建 MP_* 参数后
         再读取。若零件无 Keep 惯量测量，脚本会静默退出，不阻塞。
      3. **SPA 逐 Body 测量**：对每个 Body 依次尝试：
         A. ``spa.GetMeasurable(body)`` — 直接传 Body 对象；
         B. ``spa.GetMeasurable(CreateReferenceFromObject(body))`` — 通过引用。
         读取质量时依次尝试：``.Volume`` → ``GetVolume(arr)`` → ``GetVolume()``，
         再根据推算密度换算为质量；或直接读取 ``.Mass`` / ``GetMass()``。
         SPA 返回值（kg/kg·mm²）在返回前统一换算为 g/g·mm²（×1000）。

    参数：
        doc_com:  COM 对象（PartDocument 层）。
        part_com: COM 对象（Part 层）。
        filepath: 零件文档的磁盘路径，传给 VBS 脚本以定位目标文档。

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
    # ── DEBUG: 记录传入对象的 COM 类型名 ─────────────────────────────────
    try:
        logger.debug(f"[SPA] doc_com 类型: {type(doc_com).__name__}")
    except Exception as e:
        logger.debug(f"[SPA] 无法获取 doc_com 类型: {e}")
    try:
        logger.debug(f"[SPA] part_com 类型: {type(part_com).__name__}")
    except Exception as e:
        logger.debug(f"[SPA] 无法获取 part_com 类型: {e}")

    # ── 路径 1：读取 MP_* 用户参数（由 create_inertia_relations.catvbs 写入）──────
    _mp = _try_mp_params(part_com, "直接读取")
    if _mp is not None:
        return _mp

    # ── 路径 2：自动运行 VBS 绑定脚本（要求零件已有 Keep 测量）─────────────────────
    _mp = _run_inertia_vbs_and_read(doc_com, part_com, "VBS绑定", filepath)
    if _mp is not None:
        return _mp

    # ── 路径 3：SPA 逐 Body 测量（兜底）─────────────────────────────────────────
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

    # ── 辅助：获取一个 Measurable 对象（直接传对象或通过 Reference）─────────
    def _get_measurable(obj, obj_label: str):
        """尝试获取 obj 的 Measurable：先直接传对象，再通过 CreateReferenceFromObject。
        返回 (measurable, source_label)，失败时返回 (None, None)。
        """
        # 路径 A：直接传 Body/Part 对象（CATIA V5 脚本推荐方式）
        try:
            m = spa.GetMeasurable(obj)
            logger.debug(f"[SPA]   {obj_label} GetMeasurable(直接) 成功，类型: {type(m).__name__}")
            return m, "直接"
        except Exception as e_direct:
            logger.debug(f"[SPA]   {obj_label} GetMeasurable(直接) 失败: {e_direct}")

        # 路径 B：通过 CreateReferenceFromObject 创建引用再测量
        try:
            ref = part_com.CreateReferenceFromObject(obj)
            logger.debug(f"[SPA]   {obj_label} CreateReferenceFromObject 成功，类型: {type(ref).__name__}")
            m = spa.GetMeasurable(ref)
            logger.debug(f"[SPA]   {obj_label} GetMeasurable(ref) 成功，类型: {type(m).__name__}")
            return m, "ref"
        except Exception as e_ref:
            logger.debug(f"[SPA]   {obj_label} GetMeasurable(ref) 失败: {e_ref}")

        return None, None

    def _read_mass(meas, label: str) -> float | None:
        """从 Measurable 读取质量：先 .Mass 属性，再 GetMass() 方法。"""
        try:
            v = meas.Mass
            logger.debug(f"[SPA]   {label} .Mass = {v}")
            return v
        except Exception as e1:
            logger.debug(f"[SPA]   {label} .Mass 失败: {e1}，尝试 GetMass()")
        try:
            v = meas.GetMass()
            logger.debug(f"[SPA]   {label} GetMass() = {v}")
            return v
        except Exception as e2:
            logger.debug(f"[SPA]   {label} GetMass() 也失败: {e2}")
        return None

    def _read_volume(meas, label: str) -> float | None:
        """从 Measurable 读取体积：先 .Volume 属性，再传数组，再无参调用。"""
        try:
            v = meas.Volume
            logger.debug(f"[SPA]   {label} .Volume = {v}")
            return v
        except Exception as e0:
            logger.debug(f"[SPA]   {label} .Volume 失败: {e0}，尝试 GetVolume(arr)")
        try:
            arr = [0.0]
            meas.GetVolume(arr)
            logger.debug(f"[SPA]   {label} GetVolume(arr) = {arr[0]}")
            return arr[0]
        except Exception as e1:
            logger.debug(f"[SPA]   {label} GetVolume(arr) 失败: {e1}，尝试无参")
        try:
            v = meas.GetVolume()
            logger.debug(f"[SPA]   {label} GetVolume() 无参 = {v}")
            return v
        except Exception as e2:
            logger.debug(f"[SPA]   {label} GetVolume() 无参也失败: {e2}")
        return None

    def _read_density(meas, label: str) -> float | None:
        """从 Measurable 直接读取密度：先 .Density 属性，再 GetDensity() 方法。
        材质已应用时 SPA 可直接返回密度，无需通过 Mass/Volume 推算。
        """
        try:
            v = meas.Density
            if v is not None and v > 0.0:
                logger.debug(f"[SPA]   {label} .Density = {v}")
                return float(v)
        except Exception as e1:
            logger.debug(f"[SPA]   {label} .Density 失败: {e1}，尝试 GetDensity()")
        try:
            v = meas.GetDensity()
            if v is not None and v > 0.0:
                logger.debug(f"[SPA]   {label} GetDensity() = {v}")
                return float(v)
        except Exception as e2:
            logger.debug(f"[SPA]   {label} GetDensity() 也失败: {e2}")
        return None

    # ── 预推算零件密度（材质定义在 Part 层时 per-body GetMass 失败的备用路径）
    # 优先顺序：
    #   1. meas.Density / GetDensity()  —— SPA 直接返回材质密度
    #   2. mass / volume 比值            —— 通过质量和体积推算
    #   3. Part.Analyze                  —— Part 级 Analyze 接口推算
    # density 单位 kg/mm³
    _part_density: float | None = None
    _density_candidates = [("part_com", part_com)] + [
        (f"Body_{j}", bodies.Item(j)) for j in range(1, body_count + 1)
    ]
    for _dc_label, _dc_obj in _density_candidates:
        try:
            _mw, _src = _get_measurable(_dc_obj, f"密度探针({_dc_label})")
            if _mw is None:
                continue
            # 路径 1：直接读密度
            _direct_d = _read_density(_mw, f"密度探针({_dc_label})")
            if _direct_d and _direct_d > 0.0:
                _part_density = _direct_d
                logger.debug(
                    f"[SPA] 推算零件密度（来源={_dc_label}/{_src}，直接读取）:"
                    f" {_part_density:.6e} kg/mm³"
                )
                break
            # 路径 2：质量/体积比
            _wm = _read_mass(_mw, f"密度探针({_dc_label})")
            _wv = _read_volume(_mw, f"密度探针({_dc_label})")
            if _wm and _wm > 0.0 and _wv and _wv > 0.0:
                _part_density = _wm / _wv
                logger.debug(
                    f"[SPA] 推算零件密度（来源={_dc_label}/{_src}，质量/体积）:"
                    f" {_part_density:.6e} kg/mm³"
                )
                break
        except Exception as _de:
            logger.debug(f"[SPA] 密度探针({_dc_label}) 失败: {_de}")

    # 路径 3：Part.Analyze 接口
    if _part_density is None:
        try:
            _analyze = part_com.Analyze
            _pa_mass = _analyze.Mass
            _pa_vol  = _analyze.Volume
            if _pa_mass and _pa_mass > 0.0 and _pa_vol and _pa_vol > 0.0:
                _part_density = _pa_mass / _pa_vol
                logger.debug(
                    f"[SPA] 推算零件密度（Part.Analyze）: {_part_density:.6e} kg/mm³"
                )
        except Exception as _pae:
            logger.debug(f"[SPA] Part.Analyze 密度推算失败: {_pae}")

    if _part_density is None:
        logger.debug("[SPA] 无法推算零件密度，体积法回退将不可用")

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

            meas, meas_src = _get_measurable(body, label)
            if meas is None:
                logger.debug(f"[SPA]   {label} 无法获取 Measurable，跳过")
                continue

            # ── 读取质量 ─────────────────────────────────────────────
            mass = _read_mass(meas, label)

            if mass is None or mass <= 0.0:
                # ── 回退：体积 × 密度 ───────────────────────────────────
                logger.debug(f"[SPA]   {label} 质量获取失败，尝试体积法回退")
                vol = _read_volume(meas, label)

                # 密度优先使用全局推算值；若仍为 None，尝试从本 Body 的
                # Measurable 直接读取（材质仅应用于单个 Body 的场景）
                effective_density = _part_density
                if effective_density is None:
                    effective_density = _read_density(meas, label)

                if vol is not None and vol > 0.0 and effective_density is not None and effective_density > 0.0:
                    mass = effective_density * vol
                    logger.debug(
                        f"[SPA]   {label} 体积法质量"
                        f" = {effective_density:.6e} × {vol:.6g} = {mass:.6g} kg"
                    )
                else:
                    logger.debug(
                        f"[SPA]   {label} 体积法也无法获取质量"
                        f"（vol={vol}, density={effective_density}），尝试 GetCOG 诊断探针"
                    )
                    # ── GetCOG 诊断探针 ─────────────────────────────────────
                    # GetCOG 是纯几何量（几何重心），不依赖材质定义，可能在
                    # GetMass / GetVolume 均失败的环境下仍然有效。
                    # 此处仅探测并记录结果，以便下一步判断 SPA Measurable
                    # 接口是否可用于几何属性。
                    diag_cog = _get_cog(meas, label)
                    if any(v != 0.0 for v in diag_cog):
                        logger.debug(
                            f"[SPA]   {label} GetCOG 探针成功: {diag_cog}"
                            f"（但无质量数据，跳过累积）"
                        )
                    else:
                        logger.debug(
                            f"[SPA]   {label} GetCOG 探针返回全零或失败，跳过"
                        )
                    continue

            # ── 读取重心 ───────────────────────────────────────────────
            cog = _get_cog(meas, label)

            # ── 读取转动惯量 ───────────────────────────────────────────
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
    # SPA 返回值单位：质量 kg、坐标 mm、惯量 kg·mm²。
    # 统一换算为程序内部单位 g / mm / g·mm²（质量和惯量各乘以 1000，坐标不变）。
    return {
        "weight":  total_mass * 1000.0,
        "cog":     part_cog,
        "inertia": [[I_at_cog[r][c] * 1000.0 for c in range(3)] for r in range(3)],
    }


# ---------------------------------------------------------------------------
# Post-processing helpers
# ---------------------------------------------------------------------------

def _row_inertia_to_root(row: dict) -> list[list[float]]:
    """从行的 ``_mass_props`` 局部惯量和 ``_placement`` 变换矩阵，
    计算并返回根坐标系下的惯量张量（3×3 列表）。

    若无有效 placement，直接返回局部惯量。
    """
    mp      = row.get("_mass_props") or {}
    I_local = mp.get("inertia", [[0.0] * 3 for _ in range(3)])
    placement = row.get("_placement")
    if placement is None:
        return I_local
    R  = [[placement[i][j] for j in range(3)] for i in range(3)]
    RT = [[R[j][i] for j in range(3)] for i in range(3)]
    RI = [
        [sum(R[i][k] * I_local[k][j] for k in range(3)) for j in range(3)]
        for i in range(3)
    ]
    return [
        [sum(RI[i][k] * RT[k][j] for k in range(3)) for j in range(3)]
        for i in range(3)
    ]


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
                        mass_props    = _measure_part_mass_props(part_doc_com, part_com, filepath)
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
