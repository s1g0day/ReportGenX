# -*- coding: utf-8 -*-
"""
@Description: 日志配置模块
统一的日志记录器配置，支持：
- 控制台输出
- 文件输出（带日志轮转）
- 环境变量配置日志级别
"""

import logging
import sys
import os
from pathlib import Path
from logging.handlers import RotatingFileHandler


def get_log_level_from_env() -> int:
    """
    从环境变量获取日志级别
    
    环境变量 LOG_LEVEL 可选值：
    - DEBUG: 调试信息
    - INFO: 一般信息（默认）
    - WARNING: 警告信息
    - ERROR: 错误信息
    - CRITICAL: 严重错误
    
    Returns:
        日志级别常量
    """
    level_name = os.getenv('LOG_LEVEL', 'INFO').upper()
    level_map = {
        'DEBUG': logging.DEBUG,
        'INFO': logging.INFO,
        'WARNING': logging.WARNING,
        'ERROR': logging.ERROR,
        'CRITICAL': logging.CRITICAL
    }
    return level_map.get(level_name, logging.INFO)


def get_log_directory() -> Path:
    """
    获取日志目录路径
    
    优先级:
    1. REPORTGENX_USERDATA_DIR 环境变量 (生产环境由 Electron 设置)
    2. PyInstaller frozen 模式: 从可执行文件位置推导
    3. 开发环境: 从当前文件位置推导
    
    Returns:
        日志目录的 Path 对象
    """
    # 优先检查环境变量（生产环境 Electron 传递 userData 路径）
    userdata_dir = os.getenv('REPORTGENX_USERDATA_DIR')
    if userdata_dir:
        log_dir = Path(userdata_dir) / 'output' / 'logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        return log_dir

    # 检测是否为 PyInstaller 打包环境
    if getattr(sys, 'frozen', False):
        # 打包后: resources/backend/dist/api/api.exe
        # 需要往上跳两级到 resources/backend/
        exe_dir = Path(sys.executable).parent  # dist/api/
        backend_dir = exe_dir.parent.parent    # backend/
    else:
        # 开发环境: backend/core/logger.py
        backend_dir = Path(__file__).parent.parent
    
    log_dir = backend_dir / 'output' / 'logs'
    
    # 确保目录存在
    log_dir.mkdir(parents=True, exist_ok=True)
    
    return log_dir


def setup_logger(name: str, level: int = None) -> logging.Logger:
    """
    配置并返回日志记录器
    
    Args:
        name: 日志记录器名称
        level: 日志级别（可选，默认从环境变量读取）
        
    Returns:
        配置好的日志记录器
    """
    logger = logging.getLogger(name)
    
    # 避免重复添加 handler
    if logger.handlers:
        return logger
    
    # 从环境变量获取日志级别
    if level is None:
        level = get_log_level_from_env()
    
    logger.setLevel(level)
    
    # 创建格式化器
    formatter = logging.Formatter(
        '[%(asctime)s] [%(name)s] [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 1. 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 2. 文件处理器（带日志轮转）
    try:
        log_dir = get_log_directory()
        log_file = log_dir / f'{name}.log'
        
        # RotatingFileHandler: 单个文件最大 10MB，保留 5 个备份
        file_handler = RotatingFileHandler(
            filename=str(log_file),
            maxBytes=10 * 1024 * 1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        # 如果文件日志创建失败，只使用控制台输出
        logger.warning(f"Failed to create file handler: {e}")
    
    return logger


# 创建默认日志记录器
default_logger = setup_logger('ReportGenX')

