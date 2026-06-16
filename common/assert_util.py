# ============================================
# 断言工具模块 - 封装常用的结果判断
# 存放路径: common/assert_util.py
# ============================================

from common.log_util import log


class AssertUtil:
    """断言工具类：把常用的判断逻辑封装起来，让测试用例更简洁"""

    @staticmethod
    def assert_status_code(response, expected_code: int):
        """
        断言 HTTP 状态码是否一致。
        这是最基础的断言，每个接口测试都会用到。
        """
        actual = response.status_code
        assert actual == expected_code, (
            f"状态码断言失败！预期: {expected_code}，实际: {actual}"
        )
        log.info(f"✓ 状态码断言通过: {actual} == {expected_code}")

    @staticmethod
    def assert_contains(response, key: str, expected_value):
        """
        断言响应 JSON 中某个字段的值是否等于预期值。
        比如检查 {"msg": "登录成功"} 中的 msg 字段。
        """
        try:
            resp_json = response.json()
        except ValueError:
            raise AssertionError("响应体不是合法的 JSON 格式，无法进行字段断言")

        actual_value = resp_json.get(key)
        assert actual_value == expected_value, (
            f"字段 [{key}] 断言失败！预期: {expected_value}，实际: {actual_value}"
        )
        log.info(f"✓ 字段断言通过: resp['{key}'] == {expected_value}")

    @staticmethod
    def assert_not_empty(response, key: str):
        """
        断言响应 JSON 中某个字段的值不为空。
        比如检查登录成功后 token 不为空。
        """
        try:
            resp_json = response.json()
        except ValueError:
            raise AssertionError("响应体不是合法的 JSON 格式，无法进行字段断言")

        actual_value = resp_json.get(key)
        assert actual_value is not None and actual_value != "", (
            f"字段 [{key}] 值为空！实际: {actual_value}"
        )
        log.info(f"✓ 字段非空断言通过: resp['{key}'] = {actual_value}")
