"""
质量特性汇总计算模块。

提供：
- rollup_mass_properties() – 将每个零件实例的质量特性（含位置变换矩阵）
                             按标准刚体力学汇总，计算装配体根产品的总质量、
                             总重心和总转动惯量。
"""


def _mat3_mul(A: list[list[float]], B: list[list[float]]) -> list[list[float]]:
    """3×3 矩阵乘法，返回 A @ B。"""
    C = [[0.0] * 3 for _ in range(3)]
    for i in range(3):
        for j in range(3):
            for k in range(3):
                C[i][j] += A[i][k] * B[k][j]
    return C


def _mat3_transpose(A: list[list[float]]) -> list[list[float]]:
    """返回 3×3 矩阵的转置。"""
    return [[A[j][i] for j in range(3)] for i in range(3)]


def rollup_mass_properties(rows: list[dict]) -> dict:
    """汇总装配体各零件实例的质量特性，计算根产品的总质量特性。

    算法（标准刚体力学）：
    1. 对每个零件实例，从 ``_placement`` 提取 3×3 旋转矩阵 R 和平移向量 T。
    2. 将局部重心变换到根坐标系：``r_root = R @ r_local + T``。
    3. 将局部转动惯量张量旋转到根坐标系：``I_root_at_local_cog = R @ I_local @ R^T``。
    4. 用平行轴定理将转动惯量从局部重心移到根原点：
       ``I_root_at_origin = I_root_at_local_cog + m*(|r_root|²·E - r_root⊗r_root)``
    5. 累加所有实例：``M_total = Σm_i``，``Σ(m_i·r_i)``，``I_origin_total = Σ(I_i_at_origin)``。
    6. 计算总重心：``CoG_total = Σ(m_i·r_i) / M_total``。
    7. 用平行轴定理从根原点移回总重心：
       ``I_final = I_origin_total - M_total*(|CoG_total|²·E - CoG_total⊗CoG_total)``。

    参数：
        rows:
            ``collect_mass_props_rows()`` 返回的行列表。
            仅处理 ``Type == "零件"`` 且 ``_mass_props`` 不为 ``None`` 的行。

    返回：
        字典 ``{"total_weight": float, "cog": [x, y, z], "inertia": [[3×3]]}``，
        单位分别为 kg、mm、kg·mm²。若无有效零件则返回全零结果。
    """
    M_total      = 0.0
    sum_mr       = [0.0, 0.0, 0.0]      # Σ(m_i · r_i)
    I_at_origin  = [[0.0] * 3 for _ in range(3)]  # Σ I_i 转换到根原点

    for row in rows:
        if row.get("Type") != "零件":
            continue
        mp = row.get("_mass_props")
        if not mp:
            continue
        mass = float(mp.get("weight", 0.0))
        if mass <= 0.0:
            continue

        # 用户可能已修改 Weight 列，优先使用行中的覆盖值
        weight_override = row.get("Weight")
        if weight_override is not None:
            try:
                override_val = float(weight_override)
                if override_val > 0.0 and override_val != mass:
                    # 按比例缩放转动惯量
                    scale = override_val / mass
                    mass = override_val
                    orig_inertia = mp.get("inertia", [[0.0, 0.0, 0.0] for _ in range(3)])
                    mp = {
                        "weight":  mass,
                        "cog":     mp.get("cog", [0.0, 0.0, 0.0]),
                        "inertia": [[orig_inertia[r][c] * scale for c in range(3)]
                                    for r in range(3)],
                    }
            except (TypeError, ValueError):
                pass

        # 1. 提取变换矩阵中的 R（3×3）和 T（3×1）
        placement = row.get("_placement")
        if placement:
            R = [[placement[i][j] for j in range(3)] for i in range(3)]
            T = [placement[0][3], placement[1][3], placement[2][3]]
        else:
            R = [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]]
            T = [0.0, 0.0, 0.0]

        # 2. 将局部重心变换到根坐标系
        r_local = mp.get("cog", [0.0, 0.0, 0.0])
        r_root  = [
            sum(R[i][k] * r_local[k] for k in range(3)) + T[i]
            for i in range(3)
        ]

        # 3. 旋转局部转动惯量张量到根坐标系（仍在局部重心处）
        I_local = mp.get("inertia", [[0.0]*3 for _ in range(3)])
        RT      = _mat3_transpose(R)
        I_root_at_local_cog = _mat3_mul(R, _mat3_mul(I_local, RT))

        # 4. 平行轴定理：从局部重心移到根原点
        r2 = sum(r_root[k] ** 2 for k in range(3))
        I_root_at_origin_i = [[0.0] * 3 for _ in range(3)]
        for ii in range(3):
            for jj in range(3):
                delta = (1.0 if ii == jj else 0.0) * r2 - r_root[ii] * r_root[jj]
                I_root_at_origin_i[ii][jj] = (
                    I_root_at_local_cog[ii][jj] + mass * delta
                )

        # 5. 累加
        M_total += mass
        for k in range(3):
            sum_mr[k] += mass * r_root[k]
        for ii in range(3):
            for jj in range(3):
                I_at_origin[ii][jj] += I_root_at_origin_i[ii][jj]

    if M_total <= 0.0:
        return {
            "total_weight": 0.0,
            "cog":          [0.0, 0.0, 0.0],
            "inertia":      [[0.0] * 3 for _ in range(3)],
        }

    # 6. 计算总重心
    cog_total = [sum_mr[k] / M_total for k in range(3)]

    # 7. 平行轴定理：从根原点移回总重心
    r  = cog_total
    r2 = sum(r[k] ** 2 for k in range(3))
    I_final = [[0.0] * 3 for _ in range(3)]
    for ii in range(3):
        for jj in range(3):
            delta = (1.0 if ii == jj else 0.0) * r2 - r[ii] * r[jj]
            I_final[ii][jj] = I_at_origin[ii][jj] - M_total * delta

    return {
        "total_weight": M_total,
        "cog":          cog_total,
        "inertia":      I_final,
    }
