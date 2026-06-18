# ============================================
# 请求工具模块 - 封装 requests，统一发送 HTTP 请求
# 存放路径: common/requests_util.py
# ============================================

import requests
from common.log_util import log


class RequestsUtil:
    """
    HTTP 请求工具类。
    基于 requests.Session() 实现，自动保持 Cookie / Token，
    自动记录每次请求和响应的日志，Token 过期时自动重新登录。
    """

    def __init__(self, base_url="", timeout=30):
        """初始化会话，绑定基础URL和超时时间"""
        self.base_url = base_url.rstrip("/")  # 去掉末尾斜杠，避免拼接出双斜杠
        self.timeout = timeout
        self.session = requests.Session()  # 使用 Session 保持会话状态
        # 登录信息（外部设置，用于 Token 过期时自动重新登录）
        self._login_path = None
        self._login_body = None
        # 标记：是否正在重试（防止无限循环）
        self._is_retrying = False
        # 公共请求头
        self.session.headers.update({
            "Content-Type": "application/json",
            "User-Agent": "APITestFramework/1.0",
        })

    # ---- 外部可调用的快捷方法 ----
    def get(self, path, **kwargs):
        """发送 GET 请求"""
        return self._request("GET", path, **kwargs)

    def post(self, path, **kwargs):
        """发送 POST 请求"""
        return self._request("POST", path, **kwargs)

    def put(self, path, **kwargs):
        """发送 PUT 请求"""
        return self._request("PUT", path, **kwargs)

    def delete(self, path, **kwargs):
        """发送 DELETE 请求"""
        return self._request("DELETE", path, **kwargs)

    # ---- 核心：统一的请求发送逻辑 ----
    def _request(self, method, path, **kwargs):
        """
        内部统一方法：拼接URL、发送请求、记录日志、返回响应对象。
        所有对外方法最终都调用这里。
        Token 过期时（401）自动重新登录并重试一次。
        """
        url = self.base_url + path
        timeout = kwargs.pop("timeout", self.timeout)

        # ---------- 请求日志 ----------
        log.info(f">>> {method} {url}")
        if "json" in kwargs:
            log.info(f">>> 请求体: {kwargs['json']}")
        if "params" in kwargs:
            log.info(f">>> 请求参数: {kwargs['params']}")

        try:
            response = self.session.request(
                method=method,
                url=url,
                timeout=timeout,
                **kwargs,
            )
        except requests.Timeout:
            log.error(f"<<< 请求超时! ({timeout}s)")
            raise
        except requests.ConnectionError:
            log.error(f"<<< 无法连接到服务器: {url}")
            raise

        # ---------- Token 过期自动刷新 ----------
        # 收到 401 且不是登录接口本身且还没重试过 → 自动重新登录
        if (response.status_code == 401
                and path != self._login_path
                and not self._is_retrying
                and self._login_path is not None):
            log.warning("<<< Token 已过期 (401)，正在自动重新登录...")
            self._re_login()
            log.info(">>> 重试原请求...")
            self._is_retrying = True
            response = self._request(method, path, **kwargs)
            self._is_retrying = False

        # ---------- 响应日志 ----------
        log.info(f"<<< 状态码: {response.status_code}")
        try:
            # 尝试打印 JSON 格式的响应体（方便排查问题）
            log.info(f"<<< 响应体: {response.json()}")
        except ValueError:
            log.info(f"<<< 响应体(非JSON): {response.text[:500]}")

        return response

    # ---- 辅助方法 ----
    def update_headers(self, headers: dict):
        """动态更新请求头（例如添加 Authorization token）"""
        self.session.headers.update(headers)
        log.info(f"已更新请求头: {headers}")

    def set_login_info(self, login_path: str, login_body: dict):
        """
        设置登录信息，用于 Token 过期时自动重新登录。

        参数：
        - login_path: 登录接口路径，如 "/api/v1/login"
        - login_body: 登录请求体，如 {"username":"admin","password":"123456"}

        调用后，框架会在收到 401 时自动用此信息重新登录并重试。
        """
        self._login_path = login_path
        self._login_body = login_body
        log.info(f"已设置自动重新登录信息，登录路径: {login_path}")

    def _re_login(self):
        """
        Token 过期时自动重新登录。
        用预先设置好的账号密码重新调登录接口，拿到新 Token 后更新请求头。
        """
        if self._login_path is None or self._login_body is None:
            log.error("未设置登录信息，无法自动重新登录！请先调用 set_login_info()")
            return

        try:
            resp = self.session.post(
                self.base_url + self._login_path,
                json=self._login_body,
                timeout=self.timeout,
            )
            if resp.status_code == 200:
                token = resp.json().get("token", "")
                if token:
                    self.update_headers({"Authorization": f"Bearer {token}"})
                    log.info(f"Token 已刷新: {token[:20]}...")
                else:
                    log.error("重新登录失败！响应中未找到 token 字段")
            else:
                log.error(f"重新登录失败！状态码: {resp.status_code}")
        except Exception as e:
            log.error(f"重新登录异常: {e}")
