# ============================================
# 登录接口测试用例
# 存放路径: testcases/test_login.py
# ============================================

import pytest
import yaml
import os
from common.assert_util import AssertUtil
from common.log_util import log

# ==================== 加载测试数据 ====================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_FILE = os.path.join(BASE_DIR, "data", "login_data.yaml")

with open(DATA_FILE, encoding="utf-8") as f:
    login_data = yaml.safe_load(f)["test_cases"]


# ==================== 测试类 ====================

class TestLogin:
    """
    登录接口测试类。
    每个 test_ 开头的方法会被 pytest 自动识别为一条测试用例。
    """

    @pytest.mark.parametrize("case", login_data, ids=[c["name"] for c in login_data])
    def test_login(self, case, req_client):
        """
        登录接口参数化测试。
        一条用例覆盖多种场景：成功、密码错误、空用户名、空密码。

        参数说明：
        - case:      由 parametrize 传入，包含 name / data / expected
        - req_client: conftest.py 中定义的全局请求客户端
        """
        log.info(f"========== 执行用例: {case['name']} ==========")

        # 1. 发送请求
        resp = req_client.post(
            "/api/v1/login",               # 登录接口路径
            json=case["data"],              # 请求体：username + password
        )

        # 2. 断言 HTTP 状态码
        AssertUtil.assert_status_code(resp, case["expected"]["code"])

        # 3. 断言响应消息字段
        AssertUtil.assert_contains(resp, "msg", case["expected"]["msg"])

        # 4. 如果是成功场景，额外断言 token 不为空
        if case["expected"]["code"] == 200:
            AssertUtil.assert_not_empty(resp, "token")

        log.info(f"========== 用例通过: {case['name']} ==========")
