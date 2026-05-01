"""
重量、重心、惯量统计数据收集模块。

提供：
- collect_mass_props_rows() – 遍历产品树，读取每个零件实例的质量/重心/转动惯量，
                              不对兄弟零件进行数量合并（每个实例单独记录一行）。

数据流概述
----------
1. collect_mass_props_rows() 打开或复用已打开的 .CATProduct 文档，
   调用内部递归函数 _traverse() 深度优先遍历整棵产品树。

2. _traverse() 对每个节点：
   a. 判断节点类型（零件 / 部件 / 产品）；
   b. 通过 _position_to_mat4() 读取该节点相对父节点的局部变换矩阵，
      与父节点的累积矩阵相乘，得到"局部→根"的绝对变换矩阵（_placement）；
   c. 若节点为叶子零件，调用 _measure_part_mass_props() 测量质量特性
      （重心坐标和转动惯量在零件局部坐标系下给出），并写入行字典。

3. _post_process_rows() 对收集到的行列表进行两轮后处理：
   · 第一轮：用 _placement 中的旋转矩阵 R 和平移向量 T，
     将零件局部坐标系下的重心和转动惯量变换到根产品坐标系。
   · 第二轮：对每个产品 / 部件节点，按平行轴定理汇总子孙零件的质量特性，
     计算该节点在根坐标系下的总质量、总重心和总转动惯量。

质量特性读取
-----------
依次读取 CATIA SPA "测量惯量 + 保持测量" 写入的 "惯量包络体.1" 至
"惯量包络体.{MAX_INERTIA_INDEX}" Keep 参数，在零件级按平行轴定理汇总后存储：
  惯量包络体.N\\质量    → 质量，CATIA 原始单位 kg（已为 SI，直接存储）
  惯量包络体.N\\Gx/Gy/Gz → 重心坐标，CATIA 原始单位 mm（÷1000 换算为 m 后存储）
  惯量包络体.N\\IoxG/IoyG/IozG/IxyG/IxzG/IyzG → 转动惯量分量，CATIA 原始单位 kg·m²（已为 SI，直接存储）

单位制（内部存储，全程 SI）
--------------------------
  质量   ：kg
  长度   ：m
  惯量   ：kg·m²
整个流程以 SI 为基准，UI 显示时按用户选择换算到 g/mm/g·mm² 等实用单位。
"""

import gzip
import json
import logging
import math
from collections.abc import Callable
from pathlib import Path



from catia_copilot.constants import FILENAME_NOT_FOUND, FILENAME_UNSAVED

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 可调参数：惯量包络体编号上限
# ---------------------------------------------------------------------------

# 每个零件最多读取"惯量包络体.1"到"惯量包络体.MAX_INERTIA_INDEX"的保持测量。
# 编号不要求连续；所有编号在此范围内存在的测量均会被读取并在零件级汇总。
MAX_INERTIA_INDEX: int = 50


# ---------------------------------------------------------------------------
# 纯 Python 4×4 齐次变换矩阵辅助函数（不依赖 numpy）
#
# 矩阵布局（行主序，4 行 4 列）：
#   [ R[0][0]  R[0][1]  R[0][2]  Tx ]
#   [ R[1][0]  R[1][1]  R[1][2]  Ty ]
#   [ R[2][0]  R[2][1]  R[2][2]  Tz ]
#   [    0        0        0      1  ]
# 其中 R 为 3×3 旋转矩阵，T = (Tx, Ty, Tz) 为平移向量。
#
# 变换关系：P_parent = R @ P_local + T
# 累积（父×子）：M_abs = M_parent @ M_local
# ---------------------------------------------------------------------------

def _identity_4x4() -> list[list[float]]:
    """返回 4×4 单位矩阵（对应"无旋转、无平移"的恒等变换）。"""
    return [
        [1.0, 0.0, 0.0, 0.0],
        [0.0, 1.0, 0.0, 0.0],
        [0.0, 0.0, 1.0, 0.0],
        [0.0, 0.0, 0.0, 1.0],
    ]


def _mat4_mul(A: list[list[float]], B: list[list[float]]) -> list[list[float]]:
    """4×4 矩阵乘法，返回 C = A @ B。

    用于将两个齐次变换矩阵复合：若 A 描述"父→祖父"变换，
    B 描述"子→父"变换，则 A @ B 描述"子→祖父"变换。
    """
    C = [[0.0] * 4 for _ in range(4)]
    for i in range(4):
        for j in range(4):
            for k in range(4):
                C[i][j] += A[i][k] * B[k][j]
    return C


def _row_inertia_to_root(row: dict) -> list[list[float]]:
    """将行的 _mass_props.inertia（零件局部坐标系）旋转变换到根坐标系。

    公式：I_root = R @ I_local @ R^T
    其中 R 为从行的 _placement 矩阵中提取的 3×3 旋转子矩阵。

    若 _placement 或 _mass_props 缺失，返回 3×3 零矩阵。
    """
    placement = row.get("_placement")
    mp = row.get("_mass_props")
    if placement is None or mp is None:
        return [[0.0] * 3 for _ in range(3)]
    I_local = mp.get("inertia", [[0.0] * 3 for _ in range(3)])
    R  = [[placement[i][j] for j in range(3)] for i in range(3)]
    RT = [[R[j][i] for j in range(3)] for i in range(3)]          # R^T
    RI = [
        [sum(R[i][k] * I_local[k][j] for k in range(3)) for j in range(3)]
        for i in range(3)
    ]
    I_root = [
        [sum(RI[i][k] * RT[k][j] for k in range(3)) for j in range(3)]
        for i in range(3)
    ]
    return I_root


def _position_to_mat4(product) -> list[list[float]]:
    """从 pycatia Product 包装对象的 position.get_components() 读取位置，返回 4×4 变换矩阵。

    接收 pycatia Product 包装对象（非 com_object），无参调用
    ``product.position.get_components()``，捕获返回的 12 个 double 分量。

    CATIA Position.GetComponents 数组布局（**列主序**，共 12 个元素）：
      arr[ 0.. 2] = X 轴方向向量（旋转矩阵第 1 列）
      arr[ 3.. 5] = Y 轴方向向量（旋转矩阵第 2 列）
      arr[ 6.. 8] = Z 轴方向向量（旋转矩阵第 3 列）
      arr[ 9..11] = 原点平移向量 T = (Tx, Ty, Tz)

    组装为行主序 4×4 矩阵：
      mat[i][j] 对应旋转矩阵第 i 行、第 j 列，即 arr[j*3 + i]
      mat[i][3] 对应平移分量 arr[9 + i]

    变换含义：P_parent = R @ P_local + T

    若调用失败或返回值无效，返回 4×4 单位矩阵（等价于零件位于父坐标系原点，无旋转）。
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

    # 将列主序 12 元素数组重新排列为行主序 4×4 矩阵：
    #   第 0 行 = [arr[0], arr[3], arr[6], arr[ 9]/1000]  ← X 分量
    #   第 1 行 = [arr[1], arr[4], arr[7], arr[10]/1000]  ← Y 分量
    #   第 2 行 = [arr[2], arr[5], arr[8], arr[11]/1000]  ← Z 分量
    #   第 3 行 = [    0,      0,      0,         1    ]  ← 齐次行
    # 注：CATIA Position.GetComponents 返回的平移分量单位为 mm，
    #     此处除以 1000 将其转换为 m，与内部 SI 单位制（m）保持一致。
    mat = [
        [arr[0], arr[3], arr[6], arr[9]  / 1000.0],
        [arr[1], arr[4], arr[7], arr[10] / 1000.0],
        [arr[2], arr[5], arr[8], arr[11] / 1000.0],
        [0.0,    0.0,    0.0,    1.0               ],
    ]
    logger.debug(
        f"[MAT4] {product_name}: R[0]={mat[0][:3]}, T={[mat[0][3], mat[1][3], mat[2][3]]}"
    )
    return mat


# ---------------------------------------------------------------------------
# 质量特性读取辅助函数
# ---------------------------------------------------------------------------


def _read_keep_inertia_params(
    part_com,
    part_number: str = "",
    label: str = "",
    read_mode: str = "all",
) -> dict | None:
    """读取 CATIA SPA Keep 测量写入的"惯量包络体.1"至"惯量包络体.MAX_INERTIA_INDEX"参数，
    并在零件级按平行轴定理汇总为单一质量特性。

    先决条件：零件已在 SPA（惯量分析）中执行"测量惯量"并勾选"保持测量"，
    使参数树中出现 "惯量包络体.N\\质量"、"惯量包络体.N\\Gx" 等字段（N ≥ 1）。
    **必须单独打开零件文件再建立测量**——在产品环境下建立的测量其参考系为产品坐标系，
    导致结果不正确，此类测量将不被读取。

    读取策略（对每个编号 N，依次尝试以下前缀，取第一个能读到有效质量的前缀）：
      1. "{part_number}\\惯量包络体.N\\"  ← CATIA 以零件编号作为顶层命名空间
      2. "惯量包络体.N\\"                  ← 当前文档上下文回退前缀
    编号不要求连续，所有 1 ≤ N ≤ MAX_INERTIA_INDEX 中存在的测量均会被读取（取决于 read_mode）。

    read_mode 参数控制读取哪些惯量包络体编号：
      "first" — 仅读取"惯量包络体.1"（固定取编号 1 的测量结果）。
      "last"  — 扫描所有编号，仅返回编号最大的有效测量结果。
      "all"   — 读取全部有效编号并按平行轴定理汇总（默认行为）。

    CATIA Keep 参数的原始单位（注意坐标为 mm，非 m）：
      质量                            CATIA 原始: kg   → 内部存储: kg  （无需换算）
      Gx / Gy / Gz                    CATIA 原始: mm   → 内部存储: m   （÷ 1 000）
      IoxG / IoyG / IozG              CATIA 原始: kg·m² → 内部存储: kg·m²（无需换算）
      IxyG / IxzG / IyzG              CATIA 原始: kg·m² → 内部存储: kg·m²（无需换算）

    零件级汇总算法（标准刚体力学，均在零件局部坐标系下）：
      1. 累积各测量的质量及"惯量移到局部坐标原点"的贡献。
      2. 计算总重心：r_c = Σ(m_i · r_i) / M。
      3. 平行轴定理从原点移回总重心，得汇总惯量张量。

    CATIA Keep 参数中亦可选读取密度字段：
      密度                            CATIA 原始: kg/m³ → 内部存储: kg/m³（无需换算）
      当单个测量内材料不统一时 CATIA 返回 -1；跨多个惯量包络体密度不一致时同样返回 -1。

    返回值结构（内部 SI 单位）：
      {
        "weight":  float,               # 总质量，kg
        "cog":     [x, y, z],           # 总重心，m（零件局部坐标系，已由 mm 换算）
        "inertia": [[Ixx, Ixy, Ixz],    # 总重心处转动惯量张量（3×3 对称矩阵），kg·m²
                    [Ixy, Iyy, Iyz],
                    [Ixz, Iyz, Izz]],
        "density": float | None,        # 密度，kg/m³；-1.0 表示不统一；None 表示无密度数据
      }
    若所有编号均未找到有效质量，则返回 None。
    """
    tag = f"[MP] {label} " if label else "[MP] "
    try:
        params = part_com.Parameters

        def _get(prefix: str, name: str) -> float | None:
            try:
                return float(params.Item(prefix + name).Value)
            except Exception:
                return None

        # ── 确定需要扫描的编号范围 ────────────────────────────────────────────────
        if read_mode == "first":
            check_indices = [1]
        else:
            # "last" 和 "all" 均需扫描全范围，以找到所有或编号最大的有效测量
            check_indices = list(range(1, MAX_INERTIA_INDEX + 1))

        # ── 逐编号读取，收集所有有效测量 ──────────────────────────────────────────
        measurements: list[dict] = []
        for idx in check_indices:
            envelope_name = f"惯量包络体.{idx}"
            probe_prefix = (f"{part_number}\\{envelope_name}\\" if part_number
                            else f"{envelope_name}\\")

            mass_si = _get(probe_prefix, "质量")
            if mass_si is None or mass_si <= 0.0:
                continue  # 该编号不存在，跳过

            prefix_ok = probe_prefix

            gx_si  = _get(prefix_ok, "Gx")
            gy_si  = _get(prefix_ok, "Gy")
            gz_si  = _get(prefix_ok, "Gz")
            ixx_si = _get(prefix_ok, "IoxG")
            iyy_si = _get(prefix_ok, "IoyG")
            izz_si = _get(prefix_ok, "IozG")
            ixy_si = _get(prefix_ok, "IxyG")
            ixz_si = _get(prefix_ok, "IxzG")
            iyz_si = _get(prefix_ok, "IyzG")

            # 惯量分量允许为 0（球对称体），但不允许任意分量读取失败
            if any(v is None for v in (gx_si, gy_si, gz_si,
                                       ixx_si, iyy_si, izz_si,
                                       ixy_si, ixz_si, iyz_si)):
                logger.debug(f"{tag}{envelope_name} 部分参数缺失，跳过该测量")
                continue

            # 密度（可选参数）：CATIA 原始单位 kg/m³，不一致时返回 -1
            density_raw = _get(prefix_ok, "密度")

            measurements.append({
                "weight": mass_si,
                # Gx/Gy/Gz 由 CATIA 以 mm 存储，÷1000 换算为内部 SI 单位（m）
                "cog": [gx_si / 1000.0, gy_si / 1000.0, gz_si / 1000.0],
                "inertia": [
                    [ixx_si, ixy_si, ixz_si],
                    [ixy_si, iyy_si, iyz_si],
                    [ixz_si, iyz_si, izz_si],
                ],
                "density": density_raw,  # None：无密度参数；-1.0：CATIA 报材料不统一；>0：kg/m³
            })

        if not measurements:
            logger.debug(f"{tag}未找到任何有效的惯量包络体参数，返回 None")
            return None

        # "last" 模式：仅保留编号最大的有效测量（已按升序扫描，取最后一个）
        if read_mode == "last":
            measurements = [measurements[-1]]

        if len(measurements) == 1:
            # 仅一个测量，无需汇总，直接返回（density 已含在 measurements[0] 中）
            return measurements[0]

        # ── 零件级汇总：平行轴定理（均在零件局部坐标系下）────────────────────────
        M_total   = 0.0
        sum_mr    = [0.0, 0.0, 0.0]
        I_at_orig = [[0.0] * 3 for _ in range(3)]

        for meas in measurements:
            m  = float(meas["weight"])
            r  = meas["cog"]
            Ic = meas["inertia"]
            r2 = sum(r[k] ** 2 for k in range(3))
            for ii in range(3):
                for jj in range(3):
                    delta = (1.0 if ii == jj else 0.0) * r2 - r[ii] * r[jj]
                    I_at_orig[ii][jj] += Ic[ii][jj] + m * delta
            M_total += m
            for k in range(3):
                sum_mr[k] += m * r[k]

        cog_total = [sum_mr[k] / M_total for k in range(3)]

        rc  = cog_total
        rc2 = sum(rc[k] ** 2 for k in range(3))
        I_final = [[0.0] * 3 for _ in range(3)]
        for ii in range(3):
            for jj in range(3):
                delta = (1.0 if ii == jj else 0.0) * rc2 - rc[ii] * rc[jj]
                I_final[ii][jj] = I_at_orig[ii][jj] - M_total * delta

        # ── 跨多个惯量包络体的密度汇总 ────────────────────────────────────────
        # 规则：任意一个测量报"不统一"（-1）→ 整体为 -1；
        #       所有有效密度值（>0）不完全相同 → 整体为 -1（多材料）；
        #       所有有效密度值相同 → 取该值；无任何密度数据 → None。
        agg_density: float | None = None
        has_inconsistent = False
        valid_densities: list[float] = []
        for meas in measurements:
            d = meas.get("density")
            if d is None:
                continue
            if d < 0:
                has_inconsistent = True
            else:
                valid_densities.append(d)
        if has_inconsistent:
            agg_density = -1.0
        elif valid_densities:
            # 判断各密度值是否一致（相对误差 < 1e-9）
            d0 = valid_densities[0]
            if all(math.isclose(d, d0, rel_tol=1e-9) for d in valid_densities[1:]):
                agg_density = d0
            else:
                agg_density = -1.0

        logger.debug(
            f"{tag}汇总 {len(measurements)} 个惯量包络体测量: "
            f"weight={M_total:.4g} kg, cog={[round(v,4) for v in cog_total]} m, "
            f"density={agg_density} kg/m³"
        )
        return {"weight": M_total, "cog": cog_total, "inertia": I_final, "density": agg_density}

    except Exception as e:
        logger.debug(f"{tag}惯量包络体参数读取异常: {e}")
        return None


def _measure_part_mass_props(
    part_com,
    part_number: str = "",
    read_mode: str = "all",
) -> dict | None:
    """测量零件质量特性。

    所有返回值均使用 **SI 单位制（kg / m / kg·m²）**。

    先决条件：
      零件已在 SPA 中执行"测量惯量"并勾选"保持测量"，
      从而在参数树中生成 "惯量包络体.N\\质量" 等 Keep 参数（N ≥ 1）。
      **必须单独打开零件文件再建立测量**——在产品环境下建立的测量使用产品坐标系，
      将导致结果不正确。

    参数：
        part_com:    COM 对象（Part 层）。
        part_number: 零件编号（PartNumber），用于构造参数前缀。
        read_mode:   控制读取哪些惯量包络体（"first"/"last"/"all"，默认 "all"）。

    返回字典：
      {
        "weight":  float,          # 总质量，kg
        "cog":     [x, y, z],      # 重心坐标（零件局部坐标系），m
        "inertia": [[Ixx,Ixy,Ixz],
                    [Iyx,Iyy,Iyz],
                    [Izx,Izy,Izz]], # 重心处转动惯量（零件局部坐标轴），kg·m²
      }
    若所有惯量包络体参数均不存在（零件未执行 Keep 测量）则返回 None。
    """
    return _read_keep_inertia_params(part_com, part_number, read_mode=read_mode)


# ---------------------------------------------------------------------------
# 后处理辅助函数
# ---------------------------------------------------------------------------


def _rollup_one_product(child_parts: list[dict]) -> dict | None:
    """对单个产品/部件节点，按平行轴定理汇总子树内所有零件的根坐标系质量特性。

    参数：
        child_parts: 该节点子树内所有零件的 ``_root_mp`` 字典列表，
                     每个元素含 weight（kg）、cog（m 列表）、inertia（3×3 列表）。

    返回字典（若总质量 > 0）：
        {"weight": M_total, "cog": [x, y, z], "inertia": [[3×3]]}
    若所有子零件质量均 ≤ 0，返回 None。

    算法（标准刚体力学，所有计算均在根坐标系下）：
      步骤 1+2：累积质量、质量×重心，同时将各零件重心处惯量移至根坐标原点。
        I_i_at_O = I_i + m_i * (|r_i|² * E - r_i ⊗ r_i)
      步骤 3：计算总重心 r_c = Σ(m_i * r_i) / M
      步骤 4：以平行轴定理从根原点移回总重心。
        I_final = I_total_at_O - M * (|r_c|² * E - r_c ⊗ r_c)
    """
    M_total   = 0.0
    sum_mr    = [0.0, 0.0, 0.0]
    I_at_orig = [[0.0] * 3 for _ in range(3)]

    for rmp in child_parts:
        m = float(rmp.get("weight", 0.0))
        if m <= 0.0:
            continue
        r  = rmp.get("cog",     [0.0, 0.0, 0.0])
        Ic = rmp.get("inertia", [[0.0] * 3 for _ in range(3)])

        # 平行轴定理：将零件重心处惯量移到根坐标原点
        r2 = sum(r[k] ** 2 for k in range(3))
        for ii in range(3):
            for jj in range(3):
                delta = (1.0 if ii == jj else 0.0) * r2 - r[ii] * r[jj]
                I_at_orig[ii][jj] += Ic[ii][jj] + m * delta

        M_total += m
        for k in range(3):
            sum_mr[k] += m * r[k]

    if M_total <= 0.0:
        return None

    # 计算总重心
    cog_total = [sum_mr[k] / M_total for k in range(3)]

    # 平行轴定理：从根原点移回总重心
    rc  = cog_total
    rc2 = sum(rc[k] ** 2 for k in range(3))
    I_final = [[0.0] * 3 for _ in range(3)]
    for ii in range(3):
        for jj in range(3):
            delta = (1.0 if ii == jj else 0.0) * rc2 - rc[ii] * rc[jj]
            I_final[ii][jj] = I_at_orig[ii][jj] - M_total * delta

    return {"weight": M_total, "cog": cog_total, "inertia": I_final}


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

    # ── 第一轮：零件行 → 将局部坐标系质量特性变换到根产品坐标系 ──────────────────────
    #
    # 每个零件行的 _placement 字段存储了"零件局部→根"的 4×4 齐次变换矩阵，
    # 其中左上 3×3 块为旋转矩阵 R，右上 3×1 列为平移向量 T。
    #
    # （一）重心坐标变换
    #   零件局部坐标系下的重心 r_local，变换到根坐标系：
    #     r_root = R @ r_local + T
    #
    # （二）转动惯量旋转变换
    #   转动惯量张量在不同坐标系下通过相似变换（旋转）互换：
    #     I_root = R @ I_local @ R^T
    #   注意：此处仅做坐标轴旋转，不做平移（平移修正由第二轮平行轴定理完成）。
    #   I_local 是在零件重心处、沿零件局部坐标轴方向的惯量；
    #   I_root  是在零件重心处、沿根产品坐标轴方向的惯量。
    for row in rows:
        if row.get("Type") != "零件":
            continue
        mp = row.get("_mass_props")
        if not mp:
            continue
        placement = row.get("_placement")
        if placement is None:
            continue

        # 从 4×4 矩阵中提取 3×3 旋转矩阵 R 和平移向量 T
        R = [[placement[i][j] for j in range(3)] for i in range(3)]
        T = [placement[0][3], placement[1][3], placement[2][3]]

        # ── （一）重心坐标变换：r_root[i] = Σ_k(R[i][k] * r_local[k]) + T[i] ──
        cog_local = mp.get("cog", [0.0, 0.0, 0.0])
        cog_root  = [
            sum(R[i][k] * cog_local[k] for k in range(3)) + T[i]
            for i in range(3)
        ]

        # ── （二）惯量旋转：I_root = R @ I_local @ R^T ────────────────────────
        #   分两步计算：先求 RI = R @ I_local，再求 RI @ R^T
        I_local = mp.get("inertia", [[0.0] * 3 for _ in range(3)])
        RT = [[R[j][i] for j in range(3)] for i in range(3)]  # R^T（转置）
        RI = [
            [sum(R[i][k] * I_local[k][j] for k in range(3)) for j in range(3)]
            for i in range(3)
        ]
        I_root = [
            [sum(RI[i][k] * RT[k][j] for k in range(3)) for j in range(3)]
            for i in range(3)
        ]

        # 将零件自身坐标系下的值写入显示字段（汇总BOM展示及后续缩放用）
        # 注：汇总BOM 的单行显示以零件自身坐标系为准；
        #     层级BOM 在 _get_display_rows() 中会以 _root_mp 值覆盖这些字段；
        #     根坐标系数据缓存于 _root_mp，供第二轮汇总和底部计算面板使用。
        row["CogX"] = cog_local[0]
        row["CogY"] = cog_local[1]
        row["CogZ"] = cog_local[2]
        row["Ixx"]  = I_local[0][0]
        row["Iyy"]  = I_local[1][1]
        row["Izz"]  = I_local[2][2]
        row["Ixy"]  = I_local[0][1]
        row["Ixz"]  = I_local[0][2]
        row["Iyz"]  = I_local[1][2]

        # 缓存根坐标系数据到 _root_mp，供第二轮汇总及底部计算面板使用
        row["_root_mp"] = {
            "weight":  mp.get("weight", 0.0),
            "cog":     cog_root,
            "inertia": I_root,
        }

    # ── 第二轮：产品/部件行 → 按平行轴定理汇总子孙零件的质量特性 ──────────────────
    # 算法由 _rollup_one_product() 实现（详见其文档字符串）。
    for i in range(n):
        row = rows[i]
        if row.get("Type") not in ("产品", "部件"):
            continue

        level = int(row.get("Level", 0))

        # 收集当前节点子树内所有已成功测量零件的根坐标系质量特性（_root_mp）
        # 子树范围：行索引 i+1 开始，直到遇到层级 ≤ 当前层级的行为止
        child_parts: list[dict] = []
        for j in range(i + 1, n):
            desc = rows[j]
            if int(desc.get("Level", 0)) <= level:
                break  # 已超出子树范围，停止遍历
            rmp = desc.get("_root_mp")
            if rmp and float(rmp.get("weight", 0.0)) > 0.0:
                child_parts.append(rmp)

        if not child_parts:
            continue  # 子树内无有效零件质量数据，跳过本节点

        result = _rollup_one_product(child_parts)
        if result is None:
            continue

        # 将汇总结果写入本节点的显示字段
        row["Weight"] = result["weight"]
        row["CogX"]   = result["cog"][0]
        row["CogY"]   = result["cog"][1]
        row["CogZ"]   = result["cog"][2]
        row["Ixx"]    = result["inertia"][0][0]
        row["Iyy"]    = result["inertia"][1][1]
        row["Izz"]    = result["inertia"][2][2]
        row["Ixy"]    = result["inertia"][0][1]
        row["Ixz"]    = result["inertia"][0][2]
        row["Iyz"]    = result["inertia"][1][2]


def recompute_product_rows(rows: list[dict]) -> None:
    """重新计算所有产品/部件行的汇总质量特性。

    与 ``_post_process_rows()`` 第二轮逻辑相同，但可在初始加载后独立调用——
    例如用户在对话框中手动修改了零件重量（同时更新了 ``_root_mp``），
    点击"计算"按钮后需要刷新产品/部件行的汇总结果。

    处理流程（与 _post_process_rows 第二轮完全一致）：
      · 遍历 rows，对每个产品/部件节点，收集子树内全部零件的 ``_root_mp``；
      · 按平行轴定理汇总质量、重心和转动惯量；
      · 将结果写回该节点的显示字段（Weight / CogX/Y/Z / Ixx–Iyz）。
    """
    n = len(rows)
    for i in range(n):
        row = rows[i]
        if row.get("Type") not in ("产品", "部件"):
            continue

        level = int(row.get("Level", 0))

        # 收集子树内全部零件的根坐标系质量特性
        child_parts: list[dict] = []
        for j in range(i + 1, n):
            desc = rows[j]
            if int(desc.get("Level", 0)) <= level:
                break
            rmp = desc.get("_root_mp")
            if rmp and float(rmp.get("weight", 0.0)) > 0.0:
                child_parts.append(rmp)

        if not child_parts:
            continue

        result = _rollup_one_product(child_parts)
        if result is None:
            continue

        # 写回显示字段
        row["Weight"] = result["weight"]
        row["CogX"]   = result["cog"][0]
        row["CogY"]   = result["cog"][1]
        row["CogZ"]   = result["cog"][2]
        row["Ixx"]    = result["inertia"][0][0]
        row["Iyy"]    = result["inertia"][1][1]
        row["Izz"]    = result["inertia"][2][2]
        row["Ixy"]    = result["inertia"][0][1]
        row["Ixz"]    = result["inertia"][0][2]
        row["Iyz"]    = result["inertia"][1][2]


# ---------------------------------------------------------------------------
# 二进制序列化 / 反序列化（保存与载入行数据）
#
# 格式：将行列表序列化为 JSON 字符串，再用 gzip 压缩后写入二进制文件（.mpd）。
# 用记事本等文本编辑器打开只显示乱码；只能通过本模块的接口读取。
# ---------------------------------------------------------------------------

# 序列化时跳过的内部字段：_root_mp 可由 _post_process_rows() 重新计算，
# _rows_idx 是显示层临时注入的索引，均无需持久化。
_SERIALIZE_SKIP: frozenset[str] = frozenset({"_root_mp", "_rows_idx"})


def save_rows(rows: list[dict], file_path: str) -> None:
    """将行数据序列化为压缩二进制文件（.mpd）。

    内部以 JSON 序列化行数据后用 gzip 压缩，写为二进制文件。
    用记事本等文本工具打开无法读取有效内容。

    序列化时跳过 ``_root_mp``（加载后可由 :func:`_post_process_rows` 重新计算）
    和 ``_rows_idx``（仅供显示层使用）。其余所有字段均原样写出。

    参数：
        rows:      ``collect_mass_props_rows()`` 或 :func:`load_rows`
                   返回的行列表。
        file_path: 目标文件路径（不存在则创建，已存在则覆盖）。
    """
    serializable = [
        {k: v for k, v in row.items() if k not in _SERIALIZE_SKIP}
        for row in rows
    ]
    payload = json.dumps(serializable, ensure_ascii=False).encode("utf-8")
    with gzip.open(file_path, "wb") as f:
        f.write(payload)


def load_rows(file_path: str) -> list[dict]:
    """从压缩二进制文件（.mpd）反序列化行数据，并重建运行时缓存字段。

    读取由 :func:`save_rows` 保存的文件，恢复行列表后调用
    :func:`_post_process_rows` 重新计算 ``_root_mp`` 及产品/部件汇总字段，
    与从 CATIA 现场加载后的状态完全等价。

    参数：
        file_path: 要读取的文件路径。

    返回：
        经过后处理的行列表（包含 ``_root_mp`` 及汇总显示字段）。
    """
    with gzip.open(file_path, "rb") as f:
        rows: list[dict] = json.loads(f.read().decode("utf-8"))
    _post_process_rows(rows)
    return rows


# ---------------------------------------------------------------------------
# 主收集函数
# ---------------------------------------------------------------------------

def _compute_root_mp_from_placement(
    placement: list[list[float]],
    mass_props: dict,
) -> dict:
    """利用 4×4 变换矩阵将零件局部坐标系下的质量特性变换到根坐标系。

    从 *placement*（零件局部→根的 4×4 齐次变换矩阵）中提取 3×3 旋转矩阵 R
    和平移向量 T，对 *mass_props* 中的重心坐标和转动惯量张量执行坐标变换：
      - 重心坐标：r_root = R @ r_local + T
      - 惯量张量：I_root = R @ I_local @ R^T

    返回字典格式与 ``_mass_props`` / ``_root_mp`` 字段一致（内部 SI 单位）。
    """
    R  = [[placement[i][j] for j in range(3)] for i in range(3)]
    T  = [placement[i][3] for i in range(3)]
    cog_local = mass_props.get("cog", [0.0, 0.0, 0.0])
    cog_root  = [
        sum(R[i][k] * cog_local[k] for k in range(3)) + T[i]
        for i in range(3)
    ]
    I_local = mass_props.get("inertia", [[0.0] * 3 for _ in range(3)])
    RT = [[R[j][i] for j in range(3)] for i in range(3)]
    RI = [
        [sum(R[i][k] * I_local[k][j] for k in range(3)) for j in range(3)]
        for i in range(3)
    ]
    I_root = [
        [sum(RI[i][k] * RT[k][j] for k in range(3)) for j in range(3)]
        for i in range(3)
    ]
    return {
        "weight":  mass_props.get("weight", 0.0),
        "cog":     cog_root,
        "inertia": I_root,
    }


def remeasure_part_mass_props(
    filepath: str,
    part_number: str = "",
    read_mode: str = "all",
) -> dict | None:
    """通过 CATIA COM 接口重新读取指定零件的质量特性（惯量包络体 Keep 测量）。

    在 CATIA 当前已打开的文档中查找与 *filepath* 匹配的零件文档，再调用
    :func:`_measure_part_mass_props` 读取 Keep 测量参数。适用于用户在 CATIA 中
    补充或更改惯量包络体后，无需重新遍历整棵产品树即可刷新单个零件的质量特性。

    参数：
        filepath:    零件文档的磁盘完整路径（若文档尚未在 CATIA 中打开则返回 None）。
        part_number: 零件编号（PartNumber），用于构造 Keep 参数前缀。
        read_mode:   控制读取哪些惯量包络体（"first"/"last"/"all"）。

    返回：
        成功时返回质量特性字典（内部 SI 单位，与 :func:`collect_mass_props_rows`
        相同格式）；找不到文档或读取失败时返回 None。
    """
    from pycatia import catia  # 运行时导入，避免无 CATIA 环境时报错
    try:
        caa = catia()
        application = caa.application
        application.visible = True
        documents = application.documents

        fp_resolved = Path(filepath).resolve()
        target_doc = None
        doc_count = documents.count  # 缓存文档数量，减少重复 COM 属性访问
        for i in range(1, doc_count + 1):
            try:
                doc = documents.item(i)
                if Path(doc.full_name).resolve() == fp_resolved:
                    target_doc = doc
                    break
            except Exception:
                pass

        if target_doc is None:
            logger.debug(f"[REMEAS] 找不到已打开的文档: {filepath}")
            return None

        part_com = target_doc.com_object.Part
        return _measure_part_mass_props(part_com, part_number, read_mode=read_mode)
    except Exception as e:
        logger.debug(f"[REMEAS] 重新读取质量特性失败 ({filepath}): {e}")
        return None


def collect_mass_props_rows(
    file_path: str | None,
    progress_callback: Callable[[int], None] | None = None,
    read_mode: str = "all",
    skip_hidden: bool = False,
) -> list[dict]:
    """遍历产品树，返回每个节点的质量特性行列表。

    与 collect_bom_rows() 的关键区别：
      - **不对兄弟零件去重**——每个实例单独输出一行。
      - 仅对类型为"零件"的叶子节点执行质量特性测量（通过 MP_* 用户参数或 VBS 绑定脚本），
        部件/产品节点跳过测量，其质量由后处理阶段按平行轴定理汇总子树获得。
      - 每行额外包含 ``_placement`` 字段（4×4 列表），为该实例到根坐标系的变换矩阵。

    参数：
        file_path:
            ``.CATProduct`` 文件路径，或 ``None`` 表示使用当前 CATIA 活动文档。
        progress_callback:
            可选回调，每追加一行后调用，传入当前行数。可通过抛出异常中止遍历。
        read_mode:
            控制读取哪些惯量包络体（传递给 ``_measure_part_mass_props``）：
            "first" — 仅读取惯量包络体.1；
            "last"  — 读取编号最大的惯量包络体；
            "all"   — 全部读取并按平行轴定理汇总（默认）。
        skip_hidden:
            若为 True，则跳过处于隐藏状态的节点：
            零件隐藏时不读取该行；产品/部件隐藏时连同其全部子孙一并跳过。

    返回：
        行字典列表，每行含以下键：
          Level, Type, Part Number, Filename, Nomenclature, Revision,
          Density, Weight, CogX, CogY, CogZ, Ixx, Iyy, Izz, Ixy, Ixz, Iyz,
          _filepath, _placement, _not_found, _no_file, _unreadable, _meas_failed
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

    def _is_hidden(product, pn: str = "") -> bool:
        """检查产品实例（occurrence / 树节点）在父装配中是否处于隐藏状态。

        通过 ActiveDocument.Selection 读取实例级可见性：
          1. 清空选择集
          2. 将当前节点的 COM 对象加入选择集
          3. 从 Selection.VisProperties 读取 GetShow() 结果

        此方式会临时修改 CATIA 当前选择集（副作用），finally 块中自动清空，
        以将影响降至最低。读取失败则保守地视为可见，返回 False。

        返回：catVisNoShow=1（隐藏）→ True；catVisShow=0（可见）→ False。
        """
        tag = pn or "<unknown>"
        com = product.com_object
        sel = None
        try:
            sel = application.com_object.ActiveDocument.Selection
            sel.Clear()
            sel.Add(com)
            # In win32com late-binding (IDispatch), ByRef out-params are returned
            # as Python return values when you pass an initial plain int.
            # Passing a VARIANT(VT_BYREF|VT_I4,…) causes win32com to attempt
            # int(variant) during dispatch argument marshalling which raises TypeError.
            result = sel.VisProperties.GetShow(0)
            # result may be the show-state int directly, or a tuple ending with it
            if isinstance(result, tuple):
                show_val = result[-1]
            else:
                show_val = result
            hidden = bool(show_val) if show_val is not None else False
            logger.debug(f"[VIS] {tag}: Selection.VisProperties.GetShow()={show_val} → hidden={hidden}")
            return hidden
        except Exception as e:
            logger.debug(f"[VIS] {tag}: Selection.VisProperties.GetShow() 不可用 ({e})，视为可见")
        finally:
            try:
                if sel is not None:
                    sel.Clear()
            except Exception:
                pass

        return False

    def _traverse(
        product,
        rows: list,
        level: int,
        parent_filepath: str,
        parent_mat4: list[list[float]],
    ) -> None:
        """递归遍历产品树，将每个节点的质量特性信息追加到 rows。

        参数：
            product:         当前节点的 pycatia Product 对象。
            rows:            行字典列表，结果追加于此。
            level:           当前节点的层级深度（根节点为 0）。
            parent_filepath: 父节点的文件路径（用于判断"嵌入式部件"）。
            parent_mat4:     父节点到根的累积 4×4 变换矩阵。
        """
        nonlocal _total_count

        # 读取零件编号（PartNumber）；失败时退而使用名称去掉扩展名
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

        # 可见性探测：仅当用户勾选"忽略隐藏的节点"（skip_hidden=True）时才发起
        # COM 调用（Selection.VisProperties.GetShow）；skip_hidden=False 时完全不
        # 调用 _is_hidden()，从而避免任何多余的 COM 开销。
        # 根节点（level=0）的实例是虚拟根，不存在 parent 上下文，跳过探测。
        if level >= 1 and skip_hidden:
            if _is_hidden(product, pn):
                return

        # 解析本节点对应的磁盘文件路径（通过 COM ReferenceProduct.Parent.FullName）
        try:
            filepath = product.com_object.ReferenceProduct.Parent.FullName
        except Exception:
            filepath = ""

        # filepath 为空 → CATIA 无法解析该节点的文件引用（引用丢失或文档未载入）
        not_found = not bool(filepath)
        # filepath 非空但磁盘上不存在 → 文件尚未保存到磁盘（仍在 CATIA 内存中）
        no_file   = bool(filepath) and not Path(filepath).exists()

        # ── 判断节点类型 ──────────────────────────────────────────────────────
        # 规则：
        #   1. filepath 为空             → 类型为 ""（未知/缺失）
        #   2. 与父节点 filepath 相同    → "部件"（零件特征在同一文件中定义，即嵌入式子结构）
        #   3. .catpart 文件             → "零件"（叶子节点，需进行质量测量）
        #   4. .catproduct 或其他        → "产品"（中间装配节点，质量由子树汇总）
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

        # ── 计算本节点到根坐标系的累积变换矩阵 ──────────────────────────────
        # local_mat4：本节点相对父节点的局部变换（由 CATIA Position 读取）
        # abs_mat4  ：本节点到根的绝对变换 = parent_mat4 @ local_mat4
        # 此矩阵存入 _placement 字段，后续 _post_process_rows 用它将局部质量特性变换到根系
        local_mat4 = _position_to_mat4(product)
        abs_mat4   = _mat4_mul(parent_mat4, local_mat4)

        # ── 读取 Nomenclature / Revision 属性 ─────────────────────────────────
        is_readable = True
        nomenclature = ""
        revision     = ""

        if not not_found:
            # 确保节点处于"设计模式"（非可视化/缓存模式），否则属性读取可能失败
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

        # ── 质量特性测量（仅对叶子零件节点）────────────────────────────────────
        # 依次读取"惯量包络体.1"至"惯量包络体.MAX_INERTIA_INDEX"的 Keep 测量参数，
        # 在零件级汇总后存储（零件须已执行 SPA 保持测量）。
        mass_props: dict | None = None
        meas_failed = False

        if node_type == "零件" and is_readable and filepath:
            if filepath in _mass_cache:
                # 同一文件路径已测量过（多实例复用），直接取缓存，避免重复耗时测量
                mass_props = _mass_cache[filepath]
            else:
                try:
                    # ReferenceProduct.Parent 就是该零件的 PartDocument COM 对象，
                    # 无需遍历 Documents 集合按路径查找，直接取 .Part 即可。
                    part_doc_com = product.com_object.ReferenceProduct.Parent
                    part_com     = part_doc_com.Part
                    mass_props   = _measure_part_mass_props(part_com, pn, read_mode=read_mode)
                except Exception as e:
                    logger.debug(f"无法测量零件 {filepath}: {e}")
                    mass_props  = None
                    meas_failed = True

                if mass_props is not None:
                    logger.debug(
                        f"[TRAV] {pn} 测量成功: "
                        f"weight={mass_props.get('weight')}g, "
                        f"cog={[round(v,3) for v in mass_props.get('cog',[0,0,0])]}, "
                        f"Ixx={mass_props.get('inertia',[[0]])[0][0]:.3g}g·mm²"
                    )
                else:
                    logger.debug(f"[TRAV] {pn} 惯量包络体参数不存在或读取失败")

                # 写入缓存（即使测量失败也缓存 None，防止重复尝试）
                _mass_cache[filepath] = mass_props

        # 若零件本应可测但最终无数据，标记 meas_failed（无论是找不到文档还是读参数失败）
        if mass_props is None:
            meas_failed = meas_failed or (node_type == "零件" and is_readable and not not_found)

        # ── 组装行字典 ─────────────────────────────────────────────────────────
        # CogX/Y/Z 和 Ixx 等此处存储的是零件局部坐标系下的原始测量值；
        # _post_process_rows() 将在遍历结束后统一将其变换到根坐标系。
        mp = mass_props or {}
        cog = mp.get("cog", [0.0, 0.0, 0.0])
        inertia = mp.get("inertia", [[0.0]*3 for _ in range(3)])

        row: dict = {
            "Level":        level,
            "Type":         node_type,
            "Part Number":  pn,
            # Filename 三态：文件路径为空 → "未检索到"；路径非空但磁盘不存在 → "未保存"；正常 → 文件名（不含扩展名）
            "Filename":     (FILENAME_UNSAVED   if no_file
                             else Path(filepath).stem if filepath
                             else FILENAME_NOT_FOUND),
            "Nomenclature": nomenclature,
            "Revision":     revision,
            "Density":      mp.get("density", None),  # kg/m³；-1.0 表示不统一；None 表示无密度数据
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
            "_placement":   abs_mat4,   # 零件局部坐标系 → 根产品坐标系的 4×4 变换矩阵
            "_not_found":   not_found,  # True：CATIA 无法解析文件引用（路径丢失）
            "_no_file":     no_file,    # True：路径有效但文件尚未保存到磁盘
            "_unreadable":  not is_readable,
            "_meas_failed": meas_failed,  # True：零件文档可访问但惯量包络体参数不存在
            "_mass_props":  mass_props,   # 原始测量值，供联动修改时使用
        }

        rows.append(row)
        _total_count += 1
        if progress_callback is not None:
            progress_callback(_total_count)

        # 递归遍历子节点
        # 注意：不跳过重复实例——同一文件多次出现时每个实例单独记录一行，
        # 质量特性通过 _mass_cache 共享，不会重复测量。
        try:
            count = product.products.count
            for i in range(1, count + 1):
                try:
                    child = product.products.item(i)
                    _traverse(child, rows, level + 1,
                              parent_filepath=filepath,
                              parent_mat4=abs_mat4)
                except Exception as e:
                    logger.debug(f"遍历子节点 {i} 失败: {e}")
        except Exception:
            pass

    # ── CATIA 连接与文档处理 ─────────────────────────────────────────────────
    caa         = catia()
    application = caa.application
    application.visible = True  # 确保 CATIA 窗口可见，避免后台静默状态下 COM 调用挂起
    documents   = application.documents

    if file_path is None:
        # 使用当前 CATIA 活动文档（不做文件操作，直接读取）
        product_doc  = ProductDocument(application.active_document.com_object)
        root_product = product_doc.product
        rows: list[dict] = []
        # 根节点的父矩阵为单位矩阵（无变换），从第 0 层开始遍历
        _traverse(root_product, rows, level=0, parent_filepath="",
                  parent_mat4=_identity_4x4())
        _post_process_rows(rows)
        # 遍历过程中 VBS 可能激活了各子零件文档；恢复活动文档为根产品
        try:
            product_doc.com_object.Activate()
        except Exception as e:
            logger.debug(f"恢复根文档激活状态失败（无害）: {e}")
        return rows

    src = Path(file_path).resolve()

    # 记录 CATIA 中已打开的所有文档路径，避免重复打开同一文件
    already_open: set[Path] = set()
    for i in range(1, documents.count + 1):
        try:
            already_open.add(Path(documents.item(i).full_name).resolve())
        except Exception:
            pass

    if src not in already_open:
        documents.open(str(src))

    # 在已打开文档列表中查找与目标路径匹配的文档对象
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
    # 根节点的父矩阵为单位矩阵，从第 0 层开始遍历
    _traverse(root_product, rows, level=0, parent_filepath="",
              parent_mat4=_identity_4x4())
    _post_process_rows(rows)
    # 遍历过程中 VBS 可能激活了各子零件文档；恢复活动文档为根产品
    try:
        target_doc.com_object.Activate()
    except Exception as e:
        logger.debug(f"恢复根文档激活状态失败（无害）: {e}")
    return rows
