# 双击这个文件：自动启动 Mock 服务器 → 跑测试 → 打开报告
import subprocess
import webbrowser
import time
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))
python = r".venv\Scripts\python"

print("=" * 55)
print("  接口自动化测试框架 — 一键演示")
print("=" * 55)

# 1. 启动 Mock 服务器（后台）
print("\n[1/3] 启动 Mock 服务器...")
mock = subprocess.Popen([python, "mock_server.py"],
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE)
time.sleep(2)
print("  Mock 服务器已启动: http://localhost:8899")

# 2. 运行测试
print("\n[2/3] 运行测试用例...")
print("-" * 55)
result = subprocess.run([python, "run.py", "-v", "-s"],
                        encoding="utf-8", errors="replace")
print("-" * 55)

# 3. 打开报告
print("\n[3/3] 打开测试报告...")
reports_dir = "reports"
if os.path.exists(reports_dir):
    html_files = sorted(
        [f for f in os.listdir(reports_dir) if f.endswith(".html")],
        reverse=True
    )
    if html_files:
        report_path = os.path.abspath(os.path.join(reports_dir, html_files[0]))
        webbrowser.open(report_path)
        print(f"  报告: {report_path}")

# 4. 结束
mock.terminate()
print("\n" + "=" * 55)
if result.returncode == 0:
    print("  6 passed! 全部通过！")
else:
    print(f"  有失败，退出码: {result.returncode}")
print("=" * 55)
input("\n按回车键退出...")
