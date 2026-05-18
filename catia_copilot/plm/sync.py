"""
CATIA BOM → DocdokuPLM 同步逻辑。

入口函数：sync_bom_to_plm()
  - 从当前活动 CATIA 文档读取完整产品结构（BOM）
  - 后序深度优先遍历（子节点先于父节点同步）
  - 每个节点：create_part → update_iteration（属性 + 子组件列表）→ checkin
  - 可选：导出 STEP 并上传几何文件
  - 返回 SyncResult（汇总创建数、跳过数、失败数）

注意：所有 CATIA COM 调用必须在主线程中完成，BOM 数据提取后
可在后台线程中执行 PLM 网络请求。本模块 sync_bom_to_plm() 负责
BOM 提取，调用方负责线程调度。
"""

import logging
from dataclasses import dataclass, field
from typing import Any

logger = logging.getLogger(__name__)


# ── 数据结构 ──────────────────────────────────────────────────────────────────

@dataclass
class BomNode:
    """BOM 树中的一个节点（对应 CATIA 零件或组件）。"""
    part_number: str          # CATIA Part Number（PLM 零件编号）
    # 所有属性统一存入 attrs，键名与 CATIA 列名一致：
    #   内置属性使用英文列名（Nomenclature / Definition / Revision / Source / Description）
    #   自定义属性使用 UserRefProperty 键名（中文，如"材料"/"重量"等）
    attrs: dict[str, str] = field(default_factory=dict)
    children: list["BomNode"] = field(default_factory=list)
    # 叶子节点（零件）保存 pycatia Part 对象引用，用于导出 STEP（预留）
    _catia_ref: Any = field(default=None, repr=False)


@dataclass
class SyncResult:
    """同步操作汇总结果。"""
    created: int = 0
    skipped: int = 0
    failed: int = 0
    errors: list[str] = field(default_factory=list)

    @property
    def total(self) -> int:
        return self.created + self.skipped + self.failed

    def summary(self) -> str:
        lines = [
            f"同步完成：共 {self.total} 个节点",
            f"  ✓ 新建：{self.created}",
            f"  → 已存在（跳过）：{self.skipped}",
            f"  ✗ 失败：{self.failed}",
        ]
        if self.errors:
            lines.append("\n失败详情：")
            for e in self.errors[:10]:
                lines.append(f"  · {e}")
            if len(self.errors) > 10:
                lines.append(f"  … 共 {len(self.errors)} 条错误")
        return "\n".join(lines)


# ── BOM 提取（CATIA COM，须在主线程调用） ──────────────────────────────────────

# 需要从 CATIA 读取的属性列（内置英文列名 + 用户自定义中文属性名）
# 顺序与 constants.py BOM_ALL_COLUMNS 中的内置属性对齐
_BUILTIN_ATTR_COLS: list[str] = ["Nomenclature", "Definition", "Revision", "Source", "Description"]
_CUSTOM_COLS:       list[str] = ["零件类型", "设计状态", "材料", "重量", "物料编码", "存货类别", "规格型号", "备注"]
_ALL_ATTR_COLS:     list[str] = _BUILTIN_ATTR_COLS + _CUSTOM_COLS

# PLM instanceAttributes 时跳过的列：结构性列 + 已作为 name 字段写入的 Nomenclature
_STRUCTURAL_COLS: frozenset[str] = frozenset({
    "Level", "Type", "Filename", "Filepath", "Part Number", "Quantity",
    "Nomenclature",   # 已作为零件名称（name 字段）写入，无需重复存为属性
})


def extract_bom(progress_callback=None) -> BomNode | None:
    """从当前活动 CATIA 文档提取 BOM 树。

    须在主线程中调用（COM 线程亲和性）。

    委托 bom_collect.collect_bom_rows() 完成实际产品树遍历（包含
    reference_product.part_number 处理、design_mode 切换、属性缓存等），
    再通过 _rows_to_bom_tree() 将平面行列表还原为 BomNode 树。

    参数：
        progress_callback: 可选回调 (message: str)，用于更新进度提示

    返回：
        BomNode 根节点，或 None（无活动文档或出错时）
    """
    from catia_copilot.catia.bom_collect import collect_bom_rows

    def _cb(msg: str) -> None:
        if progress_callback:
            progress_callback(msg)
        logger.debug(msg)

    _cb("正在读取 BOM……")

    # 行数进度回调：将当前行数转成文字透传
    def _row_cb(count: int) -> None:
        _cb(f"  已读取 {count} 个节点……")

    try:
        rows = collect_bom_rows(
            file_path=None,
            columns=_ALL_ATTR_COLS,        # 内置属性 + 自定义属性
            custom_columns=_CUSTOM_COLS,   # 仅自定义属性（用于 UserRefProperties 读取）
            progress_callback=_row_cb,
        )
    except Exception as exc:
        logger.error(f"BOM 提取失败：{exc}")
        _cb(f"BOM 提取失败：{exc}")
        return None

    if not rows:
        logger.warning("BOM 为空，无活动文档或文档无产品结构")
        return None

    _cb(f"BOM 读取完成，共 {len(rows)} 个节点，正在构建树……")
    return _rows_to_bom_tree(rows)


def _rows_to_bom_tree(rows: list[dict]) -> BomNode | None:
    """将 collect_bom_rows() 返回的平面层级行列表转换为 BomNode 树。

    行按前序（父先于子）排列，Level 字段表示深度（根节点为 0）。
    使用栈恢复父子关系。
    所有属性列直接存入 node.attrs，键名 = CATIA 列名。
    Source 数值字符串（"0"/"1"/"2"）在此处转换为中文显示名。
    """
    if not rows:
        return None

    from catia_copilot.constants import SOURCE_TO_DISPLAY

    root: BomNode | None = None
    stack: list[BomNode] = []

    for row in rows:
        level = int(row.get("Level", 0))
        pn    = str(row.get("Part Number") or "").strip()
        if not pn:
            pn = str(row.get("Filename") or "UNKNOWN").strip()

        node = BomNode(part_number=pn)

        # 将所有属性列存入 attrs（跳过结构性列）
        for col in _ALL_ATTR_COLS:
            val = str(row.get(col) or "").strip()
            if col == "Source":
                val = SOURCE_TO_DISPLAY.get(val, val)
            if val:
                node.attrs[col] = val

        if level == 0:
            root  = node
            stack = [node]
        else:
            while len(stack) > level:
                stack.pop()
            if stack:
                stack[-1].children.append(node)
            stack.append(node)

    return root


# ── PLM 同步（可在后台线程调用） ──────────────────────────────────────────────

def sync_bom_to_plm(
    bom_root: BomNode,
    client,
    workspace: str,
    upload_step: bool = False,
    progress_callback=None,
) -> SyncResult:
    """将 BOM 树同步到 DocdokuPLM。

    本函数不涉及 CATIA COM 调用，可在后台线程中安全执行。

    参数：
        bom_root:          extract_bom() 返回的根节点
        client:            已登录的 PlmApiClient 实例
        workspace:         PLM 工作区名称
        upload_step:       是否导出并上传 STEP 几何文件（暂未实现，预留接口）
        progress_callback: 可选回调 (message: str)

    返回：
        SyncResult 汇总
    """
    from catia_copilot.plm.api_client import PlmApiError

    result = SyncResult()

    def _cb(msg: str) -> None:
        if progress_callback:
            progress_callback(msg)
        logger.debug(msg)

    # 确保模板存在（失败不阻断同步，改为警告并以 tpl_id=None 继续）
    tpl_id: str | None = None
    try:
        tpl_id = client.ensure_part_template(workspace)
    except PlmApiError as exc:
        logger.warning(f"模板初始化失败（将以无模板方式继续）：{exc}")
        _cb(f"警告：模板初始化失败，将以无模板方式继续 — {exc}")

    # 后序遍历：子节点先于父节点处理
    _sync_node(bom_root, client, workspace, tpl_id, result, _cb)
    return result


def _sync_node(
    node: BomNode,
    client,
    workspace: str,
    tpl_id: str | None,
    result: SyncResult,
    cb,
) -> tuple[str, str] | None:
    """递归同步单个 BOM 节点，返回 (part_number, version) 或 None（失败时）。"""
    from catia_copilot.plm.api_client import PlmApiError

    # 1. 递归处理子节点（后序）
    child_components = []
    for child in node.children:
        ref = _sync_node(child, client, workspace, tpl_id, result, cb)
        if ref:
            pn, ver = ref
            child_components.append({"component": {"number": pn, "version": ver}})

    pn = node.part_number
    cb(f"同步：{pn}")

    # PLM 零件名称 = Nomenclature（中文名称）；若为空则回退到零件编号
    plm_name = node.attrs.get("Nomenclature") or pn

    # 2. 创建或获取零件；同时确定当前可写迭代号
    #    - 新建：create_part 返回 WIP 状态，迭代号固定为 1
    #    - 已存在：需先 checkout 产生新迭代，再更新属性
    iteration     = 1
    needs_checkout = False
    try:
        part_number, version = client.create_part(
            workspace, pn, plm_name,
            node.attrs.get("Description", ""),
            tpl_id,
        )
        result.created += 1
    except PlmApiError as exc:
        if exc.status_code == 409:
            try:
                part_number, version, iteration = client._get_latest_version(workspace, pn)
                result.skipped += 1
                needs_checkout = True
            except PlmApiError as exc2:
                result.failed += 1
                result.errors.append(f"{pn}: 查询现有版本失败 — {exc2}")
                cb(f"  ✗ {pn}: 查询现有版本失败 — {exc2}")
                return None
        else:
            result.failed += 1
            result.errors.append(f"{pn}: 创建失败 — {exc}")
            cb(f"  ✗ {pn}: 创建失败 — {exc}")
            return None

    # 3. 已存在零件：checkout 产生新迭代后才能写属性
    if needs_checkout:
        try:
            iteration = client.checkout_part(workspace, part_number, version)
        except PlmApiError as exc:
            logger.warning(f"Checkout 失败（{pn}），跳过属性更新：{exc}")
            return part_number, version

    # 4. 更新属性：直接使用 node.attrs，跳过结构性列
    attr_values = {
        k: v for k, v in node.attrs.items()
        if k not in _STRUCTURAL_COLS and v
    }

    try:
        client.update_iteration(
            workspace, part_number, version, iteration,
            attr_values, child_components,
        )
    except PlmApiError as exc:
        logger.warning(f"属性更新失败（{pn}）：{exc}")

    # 5. Check In
    try:
        client.checkin_part(workspace, part_number, version)
    except PlmApiError as exc:
        logger.warning(f"Check In 失败（{pn}）：{exc}")

    return part_number, version
