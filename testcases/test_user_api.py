# ============================================
# 用户接口测试用例 — 演示登录后调用认证接口
# 存放路径: testcases/test_user_api.py
# ============================================
# 这个文件演示了框架的扩展性：
#   - 不需要重新登录（conftest.py 全局登录已处理）
#   - req_client 自动携带 Authorization 请求头
#   - Token 过期时自动重新登录并重试
#   - 只需关注业务逻辑和断言

import pytest
import yaml
import os
from common.assert_util import AssertUtil
from common.log_util import log

# ==================== 加载测试数据 ====================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_FILE = os.path.join(BASE_DIR, "data", "user_data.yaml")

with open(DATA_FILE, encoding="utf-8") as f:
    user_data = yaml.safe_load(f)["test_cases"]


# ==================== 测试类 ====================

class TestUserApi:
    """
    用户接口测试类。
    演示如何快速为新接口添加测试用例。
    """

    @pytest.mark.parametrize("case", user_data, ids=[c["name"] for c in user_data])
    def test_user_api(self, case, req_client, login_token):
        """
        用户接口参数化测试。

        参数说明：
        - case:        测试数据（path / method / expected）
        - req_client:  全局请求客户端（登录后自动携带 Token）
        - login_token: 登录 fixture，确保先登录再执行
        """
        log.info(f"========== 执行用例: {case['name']} ==========")

        method = case.get("method", "GET")

        # ---- Token 过期自动刷新测试 ----
        if case["expected"]["check"] == "auto_refresh":
            # 1. 保存当前有效的 Token
            old_auth = req_client.session.headers.get("Authorization", "")

            # 2. 故意注入一个过期 Token
            req_client.update_headers({"Authorization": "Bearer expired_token"})
            log.info(f"模拟 Token 过期，已注入过期 Token: Bearer expired_token")

            # 3. 发请求 → Mock 返回 401 → 框架自动重新登录 → 重试 → 拿到 200
            if method == "GET":
                resp = req_client.get(case["path"])
            else:
                resp = req_client.post(case["path"], json=case.get("data", {}))

            # 4. 断言自动刷新后请求成功
            AssertUtil.assert_status_code(resp, case["expected"]["code"])
            resp_json = resp.json()
            assert resp_json.get("msg") == "Token 有效，请求成功", (
                f"Token 自动刷新后请求失败！响应: {resp_json}"
            )

            # 5. 验证 Token 已经被刷新（不再是 expired_token）
            new_auth = req_client.session.headers.get("Authorization", "")
            assert "expired_token" not in new_auth, (
                "Token 未被刷新！请求头中仍包含过期 Token！"
            )
            log.info(f"✓ Token 自动刷新验证通过！新 Token: {new_auth[:30]}...")

        else:
            # ---- 普通用例 ----
            # 1. 发送请求
            if method == "GET":
                resp = req_client.get(case["path"])
            elif method == "POST":
                resp = req_client.post(case["path"], json=case.get("data", {}))
            else:
                raise ValueError(f"不支持的请求方法: {method}")

            # 2. 断言状态码
            AssertUtil.assert_status_code(resp, case["expected"]["code"])

            # 3. 根据 check 类型做业务断言
            check_type = case["expected"]["check"]
            resp_json = resp.json()

            if check_type == "has_users":
                # 验证返回了用户列表
                users = resp_json.get("users", [])
                assert len(users) > 0, "用户列表为空！"
                log.info(f"✓ 获取到 {len(users)} 个用户: {users}")

            elif check_type == "has_profile":
                # 验证返回了用户信息
                username = resp_json.get("username", "")
                assert username, "用户信息中缺少 username 字段！"
                log.info(f"✓ 用户信息: username={username}, role={resp_json.get('role')}")

        log.info(f"========== 用例通过: {case['name']} ==========")
