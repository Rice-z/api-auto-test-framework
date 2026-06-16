# 接口自动化测试框架

基于 Python + pytest + requests 的接口自动化测试框架，支持数据驱动测试、全局自动登录、Mock 离线测试和 HTML 报告自动生成。

## 技术栈

`Python` `pytest` `requests` `YAML` `pytest-html`

## 项目结构

```
├── config/              # 配置层：被测地址、账号等全局配置
│   └── settings.yaml
├── common/              # 工具层：可复用的底层封装
│   ├── requests_util.py   HTTP 请求封装（Session 管理、自动日志）
│   ├── assert_util.py     断言工具（状态码、字段校验）
│   └── log_util.py        日志工具（单例模式、双输出）
├── data/                # 数据层：测试数据，与代码分离
│   ├── login_data.yaml    登录接口测试数据（4 条用例）
│   └── user_data.yaml     用户接口测试数据（2 条用例）
├── testcases/           # 用例层：测试执行
│   ├── test_login.py      登录接口测试
│   └── test_user_api.py   用户接口测试（认证后调用）
├── conftest.py          # pytest 全局夹具：自动登录、依赖注入
├── run.py               # 一键运行入口
├── mock_server.py       # Mock API 服务器（离线测试）
└── requirements.txt     # 依赖清单
```

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动 Mock 服务器（保持窗口运行）
python mock_server.py

# 3. 运行测试
python run.py -v

# 4. 查看报告
# 浏览器打开 reports/ 目录下最新的 HTML 文件
```

## 核心特性

- **分层架构**：配置层 / 工具层 / 数据层 / 用例层，关注点分离
- **数据驱动测试**：YAML 管理测试数据 + pytest parametrize 参数化，加用例只加数据不改代码
- **全局自动登录**：conftest.py 的 session 级别 fixture，整个测试只登录一次，Token 注入后所有请求自动携带
- **HTTP 会话封装**：基于 requests.Session，自动保持 Cookie、动态更新请求头、内置日志与异常捕获
- **Mock 离线测试**：基于 Python 内置 http.server 搭建 Mock API 服务器，不依赖外部网络
- **自动生成报告**：pytest-html 生成 HTML 报告，同时输出 JUnit XML 便于 CI/CD 集成

## 测试覆盖

| 模块 | 用例 | 覆盖场景 |
|------|------|---------|
| 登录接口 | 4 条 | 登录成功 / 密码错误 / 用户名为空 / 密码为空 |
| 用户接口 | 2 条 | 获取用户列表 / 获取用户信息（需要 Token 认证） |

## 示例运行结果

```
testcases/test_login.py::test_login[登录成功-正确账号密码] PASSED
testcases/test_login.py::test_login[登录失败-密码错误]     PASSED
testcases/test_login.py::test_login[登录失败-用户名为空]   PASSED
testcases/test_login.py::test_login[登录失败-密码为空]     PASSED
testcases/test_user_api.py::test_user_api[获取用户列表]   PASSED
testcases/test_user_api.py::test_user_api[获取用户信息]   PASSED

===================== 6 passed in 14.58s =====================
```
