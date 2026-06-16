# ============================================
# pytest 全局夹具 - 自动登录、加载配置
# 存放路径: conftest.py（项目根目录）
# ============================================
# conftest.py 是 pytest 的特殊文件，放在项目根目录，
# 里面定义的 fixture 会被同目录下所有测试用例自动发现和使用。

import pytest
import yaml
import os
from common.requests_util import RequestsUtil
from common.log_util import log

# ==================== 加载配置 ====================

# 获取 conftest.py 所在目录（即项目根目录）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 读取全局配置文件
with open(os.path.join(BASE_DIR, "config", "settings.yaml"), encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)

log.info(f"已加载全局配置，目标地址: {CONFIG['base_url']}")


# ==================== Fixture 定义 ====================

@pytest.fixture(scope="session")
def base_url():
    """
    返回被测系统的基础URL。
    scope="session" 表示整个测试会话只执行一次。
    """
    return CONFIG["base_url"]


@pytest.fixture(scope="session")
def req_client():
    """
    创建并返回一个全局的请求客户端（Session 级别）。
    整个测试过程共用同一个 Session，自动保持 Cookie。
    """
    client = RequestsUtil(
        base_url=CONFIG["base_url"],
        timeout=CONFIG.get("timeout", 30),
    )
    log.info("全局请求客户端已创建")
    return client


@pytest.fixture(scope="session")
def login_token(req_client):
    """
    【核心】全局自动登录。
    在测试开始前，用配置文件的账号密码调用登录接口，
    拿到 token 后写入请求客户端的公共请求头，
    后续所有接口请求自动携带这个 token。
    """
    username = CONFIG["account"]["username"]
    password = CONFIG["account"]["password"]
    login_path = CONFIG["login_path"]

    log.info(f"===== 开始全局自动登录，账号: {username} =====")

    # 发送登录请求
    resp = req_client.post(
        login_path,
        json={"username": username, "password": password},
    )

    # 校验登录是否成功
    assert resp.status_code == 200, f"全局登录失败！状态码: {resp.status_code}"

    # 从响应中提取 token
    token = resp.json().get("token", "")
    assert token, "全局登录失败！响应中未找到 token 字段"

    # 将 token 写入公共请求头，后续所有请求自动携带
    req_client.update_headers({"Authorization": f"Bearer {token}"})

    log.info(f"===== 全局登录成功，Token 已设置: {token[:20]}... =====")

    return token


@pytest.fixture
def load_yaml_data(request):
    """
    从 data/ 目录加载 YAML 测试数据文件。
    用法：在测试函数上加装饰器 @pytest.mark.parametrize
    这里作为 fixture 供测试用例调用 load_yaml_data 方法。
    """
    file_name = request.param  # 通过 parametrize 传入文件名
    file_path = os.path.join(BASE_DIR, "data", file_name)
    with open(file_path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data
