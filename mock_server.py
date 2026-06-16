# ============================================
# 本地 Mock 服务器 — 模拟被测系统
# 存放路径: mock_server.py（项目根目录）
# ============================================
# 作用：
#   1. 不需要外部网络就能运行测试
#   2. 精确控制接口行为（成功、401、400 等场景）
#   3. 演示框架和被测系统的交互过程
# 启动方式：
#   python mock_server.py      → 默认监听 http://localhost:8899
#   python mock_server.py 9999 → 自定义端口
# ============================================

import json
import sys
from http.server import HTTPServer, BaseHTTPRequestHandler


class MockAPIHandler(BaseHTTPRequestHandler):
    """
    模拟 API 服务器的请求处理器。
    每个 do_XXX 方法对应一种 HTTP 方法。
    """

    # ---- 工具方法 ----
    def _read_body(self):
        """读取请求体 JSON"""
        try:
            length = int(self.headers.get("Content-Length", 0))
            if length == 0:
                return {}
            return json.loads(self.rfile.read(length))
        except (json.JSONDecodeError, ValueError):
            return {}

    def _send_json(self, data: dict, status=200):
        """发送 JSON 响应"""
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _log_request(self):
        """打印请求日志（方便观察测试过程）"""
        print(f"  [{self.command}] {self.path}")

    # ---- HTTP 方法处理 ----

    def do_POST(self):
        """处理 POST 请求"""
        self._log_request()
        body = self._read_body()

        # ========== 登录接口 ==========
        if self.path == "/api/v1/login":
            username = body.get("username", "")
            password = body.get("password", "")

            if not username:
                self._send_json({"msg": "用户名不能为空"}, status=400)
            elif not password:
                self._send_json({"msg": "密码不能为空"}, status=400)
            elif password != "123456":
                self._send_json({"msg": "用户名或密码错误"}, status=401)
            else:
                self._send_json({
                    "msg": "登录成功",
                    "token": "mock_token_abc123",
                    "username": username,
                }, status=200)

        else:
            # 其他 POST 接口：回显请求体（模拟创建资源）
            self._send_json({
                "msg": "创建成功",
                "data": body,
                "id": 101,
            }, status=200)

    def do_GET(self):
        """处理 GET 请求"""
        self._log_request()

        # ========== 需要认证的接口 ==========
        auth_header = self.headers.get("Authorization", "")

        if self.path == "/api/v1/users":
            if not auth_header:
                self._send_json({"msg": "未授权，请先登录"}, status=401)
                return
            self._send_json({
                "users": [
                    {"id": 1, "name": "张三", "email": "zhangsan@example.com"},
                    {"id": 2, "name": "李四", "email": "lisi@example.com"},
                ]
            }, status=200)

        elif self.path == "/api/v1/user/profile":
            if not auth_header:
                self._send_json({"msg": "未授权，请先登录"}, status=401)
                return
            self._send_json({
                "id": 1,
                "username": "admin",
                "role": "管理员",
                "token_used": auth_header,  # 回显使用的 Token，方便验证
            }, status=200)

        else:
            self._send_json({"msg": "接口不存在"}, status=404)

    def do_PUT(self):
        """处理 PUT 请求"""
        self._log_request()
        body = self._read_body()
        self._send_json({
            "msg": "更新成功",
            "data": body,
        }, status=200)

    def do_DELETE(self):
        """处理 DELETE 请求"""
        self._log_request()
        self._send_json({"msg": "删除成功"}, status=200)

    # 屏蔽访问日志，保持控制台清爽
    def log_message(self, format, *args):
        pass


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8899
    server = HTTPServer(("localhost", port), MockAPIHandler)
    print("=" * 55)
    print(f"  Mock API 服务已启动")
    print(f"  地址: http://localhost:{port}")
    print(f"  按 Ctrl+C 停止服务")
    print("=" * 55)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
        server.shutdown()


if __name__ == "__main__":
    main()
