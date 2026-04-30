"""
质量特性数据收集模块。

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
直接读取 CATIA SPA "测量惯量 + 保持测量" 写入的 "惯量包络体.1" Keep 参数：
  惯量包络体.1\\质量    → 质量，SI 原始值（kg）
  惯量包络体.1\\Gx/Gy/Gz → 重心坐标，SI 原始值（m）
  惯量包络体.1\\IoxG/IoyG/IozG/IxyG/IxzG/IyzG → 转动惯量分量，SI 原始值（kg·m²）
CATIA Keep 参数以 SI 单位存储，程序直接使用，无需换算。

单位制（内部存储）
------------------
  质量   ：kg
  长度   ：m
  惯量   ：kg·m²
整个流程使用 SI 单位，显示时再由 UI 层按用户选择换算。
"""

import logging
from collections.abc import Callable
from pathlib import Path

from catia_copilot.constants import FILENAME_NOT_FOUND, FILENAME_UNSAVED

logger = logging.getLogger(__name__)


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


def _read_keep_inertia_params(part_com, part_number: str = "", label: str = "") -> dict | None:
    """直接读取 CATIA SPA Keep 测量写入的"惯量包络体.1"参数。

    先决条件：零件已在 SPA（惯量分析）中执行"测量惯量"并勾选"保持测量"，
    使参数树中出现 "惯量包络体.1\\质量"、"惯量包络体.1\\Gx" 等字段。

    前缀策略（依次尝试，取第一个能读到有效质量的前缀）：
      1. "{part_number}\\惯量包络体.1\\"  ← CATIA 以零件号作为顶层命名空间
      2. "惯量包络体.1\\"                  ← 当前文档上下文回退前缀

    CATIA Keep 参数的原始单位为 SI 制，程序直接使用，无需换算：
      质量                            SI: kg   （内部存储单位）
      Gx / Gy / Gz                    SI: m    （内部存储单位）
      IoxG / IoyG / IozG              SI: kg·m² （内部存储单位）
      IxyG / IxzG / IyzG              SI: kg·m² （内部存储单位）

    返回值结构：
      {
        "weight":  float,               # 质量，kg
        "cog":     [x, y, z],           # 重心，m（零件局部坐标系）
        "inertia": [[Ixx, Ixy, Ixz],    # 转动惯量张量（3×3 对称矩阵），kg·m²
                    [Ixy, Iyy, Iyz],
                    [Ixz, Iyz, Izz]],
      }
    若"惯量包络体.1\\质量"不存在、值 ≤ 0，或任意惯量分量缺失，则返回 None。
    """
    tag = f"[MP] {label} " if label else "[MP] "
    try:
        params = part_com.Parameters

        def _get(prefix: str, name: str) -> float | None:
            try:
                return float(params.Item(prefix + name).Value)
            except Exception:
                return None

        # 按前缀依次尝试，取第一个能读到有效质量的前缀
        prefixes = []
        if part_number:
            prefixes.append(f"{part_number}\\惯量包络体.1\\")
        prefixes.append("惯量包络体.1\\")

        prefix_ok = None
        mass_si = None
        for pfix in prefixes:
            v = _get(pfix, "质量")
            if v is not None and v > 0.0:
                prefix_ok = pfix
                mass_si = v
                break

        if prefix_ok is None:
            logger.debug(f"{tag}未找到 惯量包络体.1\\质量，返回 None")
            return None

        def _req(name: str) -> float | None:
            return _get(prefix_ok, name)

        gx_si  = _req("Gx");   gy_si  = _req("Gy");   gz_si  = _req("Gz")
        ixx_si = _req("IoxG"); iyy_si = _req("IoyG"); izz_si = _req("IozG")
        ixy_si = _req("IxyG"); ixz_si = _req("IxzG"); iyz_si = _req("IyzG")

        # 惯量分量允许为 0（球对称体），但不允许任意分量读取失败
        if any(v is None for v in (gx_si, gy_si, gz_si,
                                   ixx_si, iyy_si, izz_si,
                                   ixy_si, ixz_si, iyz_si)):
            logger.debug(f"{tag}部分 惯量包络体.1 参数缺失，返回 None")
            return None

        return {
            "weight": mass_si,
            "cog":    [gx_si, gy_si, gz_si],
            "inertia": [
                [ixx_si, ixy_si, ixz_si],
                [ixy_si, iyy_si, iyz_si],
                [ixz_si, iyz_si, izz_si],
            ],
        }
    except Exception as e:
        logger.debug(f"{tag}惯量包络体.1 参数读取异常: {e}")
        return None


def _measure_part_mass_props(part_com, part_number: str = "") -> dict | None:
    """测量零件质量特性。

    所有返回值均使用 **SI 单位制（kg / m / kg·m²）**。

    先决条件：
      零件已在 SPA 中执行"测量惯量"并勾选"保持测量"，
      从而在参数树中生成 "惯量包络体.1\\质量" 等 Keep 参数。

    参数：
        part_com:    COM 对象（Part 层）。
        part_number: 零件号（PartNumber），用于构造参数前缀。

    返回字典：
      {
        "weight":  float,          # 总质量，kg
        "cog":     [x, y, z],      # 重心坐标（零件局部坐标系），m
        "inertia": [[Ixx,Ixy,Ixz],
                    [Iyx,Iyy,Iyz],
                    [Izx,Izy,Izz]], # 重心处转动惯量（零件局部坐标轴），kg·m²
      }
    若"惯量包络体.1"参数不存在（零件未执行 Keep 测量）则返回 None。
    """
    return _read_keep_inertia_params(part_com, part_number)


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
    #
    # 算法步骤（所有计算均在根坐标系下进行）：
    #
    # 步骤 1：汇总总质量及质量×重心
    #   M = Σ m_i
    #   Σ(m_i * r_i)  （分 x/y/z 三分量分别累加）
    #
    # 步骤 2：以平行轴定理将各零件惯量移到根坐标原点
    #   I_i_at_O = I_i_cog + m_i * (|r_i|² * E - r_i ⊗ r_i)
    #     其中：E 为 3×3 单位矩阵；r_i ⊗ r_i 为外积（秩-1 矩阵）
    #   I_total_at_O = Σ I_i_at_O
    #
    # 步骤 3：计算总重心
    #   r_c = Σ(m_i * r_i) / M
    #
    # 步骤 4：以平行轴定理从根原点移回总重心
    #   I_final = I_total_at_O - M * (|r_c|² * E - r_c ⊗ r_c)
    #
    # 注：步骤 2 和步骤 4 都用到了平行轴定理（Steiner 定理）。
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
            # 子树内无有效零件质量数据，跳过本节点
            continue

        # ── 步骤 1 + 2：累积质量、质量×重心，同时将各零件惯量移到根坐标原点 ──
        M_total   = 0.0        # 总质量，g
        sum_mr    = [0.0, 0.0, 0.0]   # Σ(m_i * r_i)，g·mm
        # I_at_orig：所有零件惯量移至根原点后的总和，g·mm²
        I_at_orig = [[0.0] * 3 for _ in range(3)]

        for rmp in child_parts:
            m  = float(rmp.get("weight", 0.0))
            if m <= 0.0:
                continue
            r  = rmp.get("cog", [0.0, 0.0, 0.0])    # 零件重心（根坐标系），mm
            Ic = rmp.get("inertia", [[0.0] * 3 for _ in range(3)])  # 零件重心处惯量

            # 平行轴定理（从零件重心 → 根坐标原点）：
            #   I_i_at_O[ii][jj] = Ic[ii][jj] + m * (|r|² * δ[ii][jj] - r[ii]*r[jj])
            #   δ 为 Kronecker delta，即当 ii==jj 时为 1，否则为 0
            r2 = sum(r[k] ** 2 for k in range(3))  # |r|²
            for ii in range(3):
                for jj in range(3):
                    delta = (1.0 if ii == jj else 0.0) * r2 - r[ii] * r[jj]
                    I_at_orig[ii][jj] += Ic[ii][jj] + m * delta

            M_total += m
            for k in range(3):
                sum_mr[k] += m * r[k]

        if M_total <= 0.0:
            continue

        # ── 步骤 3：计算总重心 r_c = Σ(m_i * r_i) / M ───────────────────────
        cog_total = [sum_mr[k] / M_total for k in range(3)]

        # ── 步骤 4：以平行轴定理从根原点移回总重心 ─────────────────────────────
        #   I_final[ii][jj] = I_at_orig[ii][jj] - M * (|r_c|² * δ[ii][jj] - r_c[ii]*r_c[jj])
        rc  = cog_total
        rc2 = sum(rc[k] ** 2 for k in range(3))  # |r_c|²
        I_final = [[0.0] * 3 for _ in range(3)]
        for ii in range(3):
            for jj in range(3):
                delta = (1.0 if ii == jj else 0.0) * rc2 - rc[ii] * rc[jj]
                I_final[ii][jj] = I_at_orig[ii][jj] - M_total * delta

        # 将汇总结果写入本节点的显示字段
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

        # 步骤 1 + 2：累积质量、质量×重心，同时将各零件惯量移到根坐标原点
        M_total   = 0.0
        sum_mr    = [0.0, 0.0, 0.0]
        I_at_orig = [[0.0] * 3 for _ in range(3)]

        for rmp in child_parts:
            m = float(rmp.get("weight", 0.0))
            if m <= 0.0:
                continue
            r  = rmp.get("cog", [0.0, 0.0, 0.0])
            Ic = rmp.get("inertia", [[0.0] * 3 for _ in range(3)])

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

        # 步骤 3：计算总重心
        cog_total = [sum_mr[k] / M_total for k in range(3)]

        # 步骤 4：从根原点移回总重心
        rc  = cog_total
        rc2 = sum(rc[k] ** 2 for k in range(3))
        I_final = [[0.0] * 3 for _ in range(3)]
        for ii in range(3):
            for jj in range(3):
                delta = (1.0 if ii == jj else 0.0) * rc2 - rc[ii] * rc[jj]
                I_final[ii][jj] = I_at_orig[ii][jj] - M_total * delta

        # 写回显示字段
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
      - 仅对类型为"零件"的叶子节点执行质量特性测量（通过 MP_* 用户参数或 VBS 绑定脚本），
        部件/产品节点跳过测量，其质量由后处理阶段按平行轴定理汇总子树获得。
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

    def _traverse(
        product,
        rows: list,
        level: int,
        parent_filepath: str,
        parent_mat4: list[list[float]],
        documents,
    ) -> None:
        """递归遍历产品树，将每个节点的质量特性信息追加到 rows。

        参数：
            product:          当前节点的 pycatia Product 对象。
            rows:             行字典列表，结果追加于此。
            level:            当前节点的层级深度（根节点为 0）。
            parent_filepath:  父节点的文件路径（用于判断"嵌入式部件"）。
            parent_mat4:      父节点到根的累积 4×4 变换矩阵。
            documents:        CATIA Application.Documents 集合，用于查找零件文档。
        """
        nonlocal _total_count

        # 读取零件号（PartNumber）；失败时退而使用名称去掉扩展名
        try:
            pn = product.part_number
        except Exception:
            name = product.name
            pn   = name.rsplit(".", 1)[0] if "." in name else name

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
        # 直接读取"惯量包络体.1" Keep 测量参数（零件须已执行 SPA 保持测量）。
        mass_props: dict | None = None
        meas_failed = False

        if node_type == "零件" and is_readable and filepath:
            if filepath in _mass_cache:
                # 同一文件路径已测量过（多实例复用），直接取缓存，避免重复耗时测量
                mass_props = _mass_cache[filepath]
            else:
                # 在已打开的文档集合中查找与本零件路径匹配的文档
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
                        # 通过 COM 获取 Part 对象（仅 PartDocument 拥有 .Part 属性）
                        part_doc_com  = target_doc.com_object
                        part_com      = part_doc_com.Part
                        mass_props    = _measure_part_mass_props(part_com, pn)
                    except Exception as e:
                        logger.debug(f"无法测量零件 {filepath}: {e}")
                        mass_props  = None
                        meas_failed  = True

                    # ── DEBUG：记录每个零件的测量结果概要 ────────────────────
                    if mass_props is not None:
                        _mp_dbg = mass_props
                        logger.debug(
                            f"[TRAV] {pn} 测量成功: "
                            f"weight={_mp_dbg.get('weight')}g, "
                            f"cog={[round(v,3) for v in _mp_dbg.get('cog',[0,0,0])]}, "
                            f"Ixx={_mp_dbg.get('inertia',[[0]])[0][0]:.3g}g·mm²"
                        )
                    else:
                        logger.debug(f"[TRAV] {pn} 惯量包络体.1 参数不存在或读取失败")

                    # 写入缓存（即使测量失败也缓存 None，防止重复尝试）
                    _mass_cache[filepath] = mass_props
                else:
                    # 文档未在 CATIA 中打开，无法完成测量
                    logger.debug(f"找不到已打开的文档: {filepath}")
                    meas_failed = True

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
            "_meas_failed": meas_failed,  # True：零件文档可访问但 惯量包络体.1 参数不存在
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
                              parent_mat4=abs_mat4,
                              documents=documents)
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
                  parent_mat4=_identity_4x4(), documents=documents)
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
              parent_mat4=_identity_4x4(), documents=documents)
    _post_process_rows(rows)
    # 遍历过程中 VBS 可能激活了各子零件文档；恢复活动文档为根产品
    try:
        target_doc.com_object.Activate()
    except Exception as e:
        logger.debug(f"恢复根文档激活状态失败（无害）: {e}")
    return rows
