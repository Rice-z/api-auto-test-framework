# ============================================
# 一键运行入口 - 执行用例并生成简易测试报告
# 存放路径: run.py（项目根目录）
# ============================================
# 运行方式：
#   python run.py          → 执行全部用例并生成报告
#   python run.py -v       → 详细输出模式
#   python run.py -s       → 同时打印控制台日志（不捕获print/log）
# ============================================

import sys
import os
import time
from datetime import datetime

import pytest


def main():
    """程序入口：解析命令行参数，调用 pytest 执行测试"""

    # ---- 报告输出目录 ----
    report_dir = "./reports"
    os.makedirs(report_dir, exist_ok=True)

    # 报告文件名带时间戳，避免覆盖
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_report = os.path.join(report_dir, f"report_{timestamp}.html")
    xml_report  = os.path.join(report_dir, f"report_{timestamp}.xml")

    # ---- 构造 pytest 命令行参数 ----
    # 用户可以在 python run.py 后面追加参数，比如 -v -s
    user_args = sys.argv[1:]

    pytest_args = [
        "testcases/",                          # 测试用例目录
        f"--html={html_report}",               # 生成 HTML 报告（需要 pip install pytest-html）
        "--self-contained-html",               # HTML 报告内嵌 CSS，方便发给别人看
        f"--junitxml={xml_report}",            # 生成 XML 报告（方便 CI 工具解析）
        "--tb=short",                           # 失败时简化回溯信息
    ] + user_args  # 把用户传入的参数追加到末尾

    # ---- 打印启动信息 ----
    print("=" * 60)
    print("  接口自动化测试框架")
    print(f"  开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # ---- 执行测试 ----
    start = time.time()
    exit_code = pytest.main(pytest_args)
    elapsed = time.time() - start

    # ---- 打印汇总 ----
    print("=" * 60)
    print(f"  执行完毕! 耗时: {elapsed:.2f} 秒")
    print(f"  HTML 报告: {html_report}")
    print(f"  XML  报告: {xml_report}")
    print(f"  退出码: {exit_code} (0=全部通过, 非0=有失败)")
    print("=" * 60)

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
