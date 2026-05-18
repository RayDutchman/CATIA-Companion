"""
DocdokuPLM REST API 客户端。

仅使用标准库（urllib），不引入任何第三方依赖。

典型用法：
    client = PlmApiClient("http://localhost:8001/docdoku-plm-server-rest/api")
    client.login("admin", "password")
    tpl_id = client.ensure_part_template("Workspace_0")
    pn, version = client.create_part("Workspace_0", "PART-001", "描述", tpl_id)
    client.update_iteration("Workspace_0", pn, version, 1, {"材料": "铝合金", "重量": "1.23"}, [])
    client.checkin_part("Workspace_0", pn, version)
"""

import base64
import http.cookiejar
import json
import logging
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

logger = logging.getLogger(__name__)


class PlmApiError(Exception):
    """DocdokuPLM API 调用异常，携带 HTTP 状态码（status_code）。"""

    def __init__(self, message: str, status_code: int = 0):
        super().__init__(message)
        self.status_code = status_code


class PlmApiClient:
    """DocdokuPLM REST API 客户端（无状态，JWT 存储于实例）。"""

    # 零件属性模板 ID（固定）
    TEMPLATE_ID = "CATIA_Standard"

    # 模板字段定义：(字段名, 类型)
    # 类型：TEXT 或 NUMBER
    _TEMPLATE_ATTRS = [
        # CATIA 内置属性
        ("中文名称",  "TEXT"),
        ("版本",      "TEXT"),
        ("定义",      "TEXT"),
        ("来源",      "TEXT"),
        # 用户自定义属性
        ("零件类型",  "TEXT"),
        ("设计状态",  "TEXT"),
        ("材料",      "TEXT"),
        ("重量",      "NUMBER"),
        ("物料编码",  "TEXT"),
        ("存货类别",  "TEXT"),
        ("规格型号",  "TEXT"),
        ("备注",      "TEXT"),
    ]

    def __init__(self, base_url: str):
        """
        参数：
            base_url: DocdokuPLM REST API 根地址，例如
                      "http://localhost:8001/docdoku-plm-server-rest/api"
        """
        self._base = base_url.rstrip("/")
        self._token: str | None = None
        self._basic_auth: str | None = None   # base64(login:password)，Basic Auth 兜底

        # Cookie jar：自动接收 Set-Cookie（JSESSIONID 等）并在后续请求中回传
        self._cj = http.cookiejar.CookieJar()
        self._opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self._cj)
        )

    # ── 内部辅助 ─────────────────────────────────────────────────────────────

    def _headers(self, extra: dict | None = None) -> dict:
        """构造请求头，优先附加 JWT Bearer，次选 Basic Auth。"""
        h = {"Content-Type": "application/json", "Accept": "application/json"}
        if self._token:
            h["Authorization"] = f"Bearer {self._token}"
        elif self._basic_auth:
            h["Authorization"] = f"Basic {self._basic_auth}"
        if extra:
            h.update(extra)
        return h

    def _request(
        self,
        method: str,
        path: str,
        body: Any = None,
        *,
        expect_json: bool = True,
        extra_headers: dict | None = None,
        return_response_headers: bool = False,
    ) -> Any:
        """发送 HTTP 请求，返回解析后的 JSON（或 None）。

        参数：
            method:                 HTTP 方法（GET / POST / PUT / DELETE）
            path:                   相对于 base_url 的路径（以 / 开头）
            body:                   请求体，自动序列化为 JSON
            expect_json:            为 False 时不解析响应体，直接返回 None
            extra_headers:          额外请求头
            return_response_headers: 为 True 时返回 (data, headers_dict) 二元组
        """
        url = self._base + path
        data = json.dumps(body).encode() if body is not None else None
        headers = self._headers(extra_headers)
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        logger.debug(f"PLM {method} {url}")
        try:
            with self._opener.open(req, timeout=30) as resp:
                raw = resp.read()
                resp_headers = dict(resp.headers)
                parsed = json.loads(raw) if (expect_json and raw) else None
                if return_response_headers:
                    return parsed, resp_headers
                return parsed
        except urllib.error.HTTPError as exc:
            body_text = ""
            try:
                body_text = exc.read().decode(errors="replace")
            except Exception:
                pass
            raise PlmApiError(
                f"{method} {path} 失败 [{exc.code}]: {body_text[:200]}",
                status_code=exc.code,
            ) from exc
        except urllib.error.URLError as exc:
            raise PlmApiError(f"网络错误（{exc.reason}）：{url}") from exc

    # ── 认证 ─────────────────────────────────────────────────────────────────

    def login(self, login: str, password: str) -> None:
        """登录并保存认证凭据。

        认证优先级：
        1. 若响应体包含 ``jwt``/``token`` 字段 → Bearer JWT
        2. 若响应头包含 ``Authorization`` / ``X-Auth-Token`` → Bearer JWT（来自响应头）
        3. 若 cookie jar 收到会话 cookie（JSESSIONID 等）→ 会话 cookie（自动处理）
        4. 上述均无 → 回落至 Basic Auth（base64 编码的 login:password）

        参数：
            login:    用户名
            password: 密码
        """
        result, resp_headers = self._request(
            "POST", "/auth/login",
            {"login": login, "password": password},
            return_response_headers=True,
        )

        # 诊断：打印登录响应体和关键响应头，便于排查 401
        logger.debug(f"PLM 登录响应体：{result}")
        logger.debug(f"PLM 登录响应头：{ {k: v for k, v in (resp_headers or {}).items()} }")

        # 1. 响应头 `jwt:` —— DocdokuPLM Payara 版本的标准做法
        #    Access-Control-Expose-Headers 明确声明了 jwt 头
        jwt_header = (resp_headers or {}).get("jwt") or (resp_headers or {}).get("Jwt")
        if jwt_header:
            self._token = jwt_header.strip()
            logger.info("PLM 登录成功（JWT Bearer，来自响应头 jwt）")
            return

        # 2. 响应体 jwt/token 字段（部分旧版本）
        token = (result or {}).get("jwt") or (result or {}).get("token")
        if token:
            self._token = token
            logger.info("PLM 登录成功（JWT Bearer，来自响应体）")
            return

        # 3. 响应头 Authorization / X-Auth-Token
        auth_header = (resp_headers or {}).get("Authorization") or \
                      (resp_headers or {}).get("authorization")
        if auth_header and auth_header.lower().startswith("bearer "):
            self._token = auth_header[7:].strip()
            logger.info("PLM 登录成功（JWT Bearer，来自响应头 Authorization）")
            return
        xauth = (resp_headers or {}).get("X-Auth-Token") or \
                (resp_headers or {}).get("x-auth-token")
        if xauth:
            self._token = xauth.strip()
            logger.info("PLM 登录成功（JWT Bearer，来自响应头 X-Auth-Token）")
            return

        # 3. 检查 cookie jar 是否收到会话 cookie
        has_session_cookie = any(True for _ in self._cj)
        if has_session_cookie:
            cookies_info = [(c.name, c.domain) for c in self._cj]
            logger.info(f"PLM 登录成功（Cookie 会话）：{cookies_info}")
            return

        # 4. 兜底：使用 Basic Auth（每次请求带 Authorization: Basic）
        self._basic_auth = base64.b64encode(
            f"{login}:{password}".encode()
        ).decode()
        logger.info("PLM 登录成功（Basic Auth 兜底）")

    # ── 零件模板 ──────────────────────────────────────────────────────────────

    def ensure_part_template(self, workspace: str) -> str:
        """确保零件属性模板 CATIA_Standard 存在，不存在则创建。

        返回模板 ID（即 TEMPLATE_ID）。
        """
        # DocdokuPLM REST 端点：/workspaces/{ws}/part-templates
        path = f"/workspaces/{urllib.parse.quote(workspace)}/part-templates"
        # 检查是否已存在
        try:
            templates = self._request("GET", path) or []
            for tpl in templates:
                if tpl.get("id") == self.TEMPLATE_ID:
                    logger.debug(f"PLM 模板 '{self.TEMPLATE_ID}' 已存在")
                    return self.TEMPLATE_ID
        except PlmApiError:
            pass

        # 创建模板
        attr_templates = [
            {
                "name": name,
                "type": atype,
                "mandatory": False,
                "locked": False,
            }
            for name, atype in self._TEMPLATE_ATTRS
        ]
        self._request("POST", path, {
            "id": self.TEMPLATE_ID,
            "mask": "",
            "attributeTemplates": attr_templates,
        })
        logger.info(f"PLM 模板 '{self.TEMPLATE_ID}' 创建成功")
        return self.TEMPLATE_ID

    # ── 零件 CRUD ─────────────────────────────────────────────────────────────

    def create_part(
        self,
        workspace: str,
        part_number: str,
        description: str,
        template_id: str | None = None,
    ) -> tuple[str, str]:
        """创建零件，返回 (零件号, 版本)。

        若零件已存在（HTTP 409），则直接返回现有零件的版本号（不报错）。

        参数：
            template_id: 零件模板 ID；传 None 时不携带 templateId 字段
        """
        ws = urllib.parse.quote(workspace)
        path = f"/workspaces/{ws}/parts"
        payload: dict = {
            "number": part_number,
            "name": part_number,
            "description": description,
        }
        if template_id is not None:
            payload["templateId"] = template_id
        try:
            result = self._request("POST", path, payload)
            version = (result or {}).get("version", "A")
            logger.info(f"PLM 零件已创建：{part_number}-{version}")
            return part_number, version
        except PlmApiError as exc:
            if exc.status_code == 409:
                # 标准重复冲突
                logger.debug(f"PLM 零件已存在（409），跳过创建：{part_number}")
                return self._get_latest_version(workspace, part_number)
            if exc.status_code == 400 and (
                "不唯一" in str(exc) or "unique" in str(exc).lower()
            ):
                # DocdokuPLM 部分版本对重复零件返回 400 而非 409
                logger.debug(f"PLM 零件已存在（400 不唯一），跳过创建：{part_number}")
                return self._get_latest_version(workspace, part_number)
            raise

    def _get_latest_version(self, workspace: str, part_number: str) -> tuple[str, str]:
        """获取已存在零件的最新版本号。

        DocdokuPLM 端点须带版本后缀：/parts/{pn}-{ver}
        实际上零件版本几乎都从 A 开始，直接尝试 -A。
        """
        ws = urllib.parse.quote(workspace)
        pn = urllib.parse.quote(part_number)
        # 尝试常见版本序列 A→B→C
        for ver in ("A", "B", "C"):
            try:
                result = self._request("GET", f"/workspaces/{ws}/parts/{pn}-{ver}") or {}
                version = result.get("version", ver)
                return part_number, version
            except PlmApiError as exc:
                if exc.status_code == 404:
                    continue
                raise
        # 兜底：返回 A（后续操作失败时会有对应错误）
        return part_number, "A"

    def update_iteration(
        self,
        workspace: str,
        part_number: str,
        version: str,
        iteration: int,
        attr_values: dict[str, str],
        components: list[dict],
    ) -> None:
        """更新零件迭代的属性和子组件列表。

        参数：
            attr_values:  字段名→值 映射，重量字段自动转 NUMBER 类型
            components:   子组件列表，每项为 {"component": {"number": ..., "version": ...}}
        """
        ws = urllib.parse.quote(workspace)
        pn = urllib.parse.quote(part_number)
        path = f"/workspaces/{ws}/parts/{pn}-{version}/iterations/{iteration}"

        # 构造 instanceAttributes
        number_fields = {name for name, atype in self._TEMPLATE_ATTRS if atype == "NUMBER"}
        instance_attrs = []
        for name, value in attr_values.items():
            atype = "NUMBER" if name in number_fields else "TEXT"
            instance_attrs.append({"type": atype, "name": name, "value": value})

        self._request("PUT", path, {
            "instanceAttributes": instance_attrs,
            "components": components,
        })
        logger.debug(f"PLM 属性更新：{part_number}-{version} iter{iteration}")

    def checkin_part(self, workspace: str, part_number: str, version: str) -> None:
        """Check In 零件（锁定当前迭代）。"""
        ws = urllib.parse.quote(workspace)
        pn = urllib.parse.quote(part_number)
        self._request(
            "PUT",
            f"/workspaces/{ws}/parts/{pn}-{version}/checkin",
            expect_json=False,
        )
        logger.debug(f"PLM Check In：{part_number}-{version}")

    def upload_step(
        self,
        workspace: str,
        part_number: str,
        version: str,
        iteration: int,
        step_path: str,
    ) -> None:
        """将 STEP 文件作为几何文件上传到零件迭代。

        参数：
            step_path: 本地 .stp 文件的绝对路径
        """
        import os
        ws = urllib.parse.quote(workspace)
        pn = urllib.parse.quote(part_number)
        url = (
            f"{self._base}/workspaces/{ws}/parts/{pn}-{version}"
            f"/iterations/{iteration}/geometry"
        )

        filename = os.path.basename(step_path)
        boundary = "----PlmUploadBoundary7f3a9b2c"
        with open(step_path, "rb") as f:
            file_data = f.read()

        body_parts = [
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="upload"; filename="{filename}"\r\n'.encode(),
            b"Content-Type: application/octet-stream\r\n\r\n",
            file_data,
            f"\r\n--{boundary}--\r\n".encode(),
        ]
        body = b"".join(body_parts)

        headers = self._headers({
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        })
        # Content-Type 已被 _headers 设为 application/json，需要覆盖
        del headers["Content-Type"]
        headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"

        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        logger.debug(f"PLM STEP 上传：{filename} → {part_number}-{version} iter{iteration}")
        try:
            with self._opener.open(req, timeout=120):
                pass
        except urllib.error.HTTPError as exc:
            body_text = ""
            try:
                body_text = exc.read().decode(errors="replace")
            except Exception:
                pass
            raise PlmApiError(
                f"STEP 上传失败 [{exc.code}]: {body_text[:200]}",
                status_code=exc.code,
            ) from exc
        except urllib.error.URLError as exc:
            raise PlmApiError(f"网络错误（{exc.reason}）：{url}") from exc
        logger.info(f"PLM STEP 上传成功：{filename}")
