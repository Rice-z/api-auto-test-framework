import subprocess, os

os.chdir(os.path.dirname(os.path.abspath(__file__)))
subprocess.run([r".venv\Scripts\python", "mock_server.py"])
input("按回车键退出...")
