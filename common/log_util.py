# ============================================
# 日志工具模块 - 统一控制台和文件输出
# 存放路径: common/log_util.py
# ============================================

import logging
import os
from datetime import datetime


class LogUtil:
    """日志工具类：同时输出到控制台和日志文件"""

    _instance = None  # 单例实例

    def __new__(cls, *args, **kwargs):
        """单例模式：确保全局只有一个日志对象"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, log_dir="./reports", level=logging.INFO):
        """初始化日志器，设置日志级别、格式和输出目标"""
        # 防止重复初始化
        if hasattr(self, "_initialized"):
            return
        self._initialized = True

        # 创建日志目录
        os.makedirs(log_dir, exist_ok=True)

        # 创建日志器
        self.logger = logging.getLogger("APITest")
        self.logger.setLevel(level)

        # 日志格式：时间 [级别] 文件名:行号 - 消息
        fmt = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(filename)s:%(lineno)d - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        # ---- 控制台输出 ----
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(fmt)
        self.logger.addHandler(console_handler)

        # ---- 文件输出（按日期命名） ----
        today = datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(log_dir, f"test_{today}.log")
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)  # 文件记录更详细的日志
        file_handler.setFormatter(fmt)
        self.logger.addHandler(file_handler)

    def info(self, msg):
        self.logger.info(msg)

    def debug(self, msg):
        self.logger.debug(msg)

    def warning(self, msg):
        self.logger.warning(msg)

    def error(self, msg):
        self.logger.error(msg)


# 模块级单例，其他模块直接 import 使用
log = LogUtil()
