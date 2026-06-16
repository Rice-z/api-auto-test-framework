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
    也会自动记录每次请求和响应的日志。
    """

    def __init__(self, base_url="", timeout=30):
        """初始化会话，绑定基础URL和超时时间"""
        self.base_url = base_url.rstrip("/")  # 去掉末尾斜杠，避免拼接出双斜杠
        self.timeout = timeout
        self.session = requests.Session()  # 使用 Session 保持会话状态
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
