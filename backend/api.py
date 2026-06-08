import base64
import hashlib
import io
import importlib
import json
import os
import re
import shutil
import socket
import sys
import time
import uuid
import urllib.request
import zipfile
import subprocess
import traceback
from contextlib import asynccontextmanager
from datetime import datetime
from typing import Annotated, Any, Optional
from urllib.parse import quote, urlparse

import tldextract
import uvicorn
import yaml
from PIL import Image
from fastapi import APIRouter, Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from starlette.middleware.base import BaseHTTPMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from starlette.exceptions import HTTPException as StarletteHTTPException

API_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(API_DIR)
while PROJECT_ROOT in sys.path:
    sys.path.remove(PROJECT_ROOT)
sys.path.insert(0, PROJECT_ROOT)

import backend.core as backend_core_package
from backend.core.data_reader_db import DbDataReader
from backend.core.handler_registry import HandlerRegistry
from backend.core.logger import setup_logger
from backend.core.report_merger import ReportMerger
from backend.core.template_manager import TemplateManager
from backend.core.exceptions import TemplateNotFoundError

_LEGACY_CORE_ALIAS_MODULES = (
    "base_handler",
    "data_reader_db",
    "document_editor",
    "document_image_processor",
    "exceptions",
    "handler_registry",
    "handler_utils",
    "logger",
    "report_merger",
    "summary_generator",
    "template_manager",
)

logger = setup_logger('API')


def _ensure_legacy_core_import_aliases(
    use_alias_fallback: bool = False,
    fail_on_missing_core: bool = True,
) -> None:
    """Enable opt-in legacy aliasing only when top-level ``core`` package is unavailable."""

    try:
        importlib.import_module("core")
        logger.info("Detected top-level core SDK package; skip legacy core aliasing")
        return
    except Exception:
        if not use_alias_fallback:
            message = "Top-level core SDK not found and legacy alias fallback is disabled"
            if fail_on_missing_core:
                raise RuntimeError(
                    f"{message}. Set plugin_runtime.use_legacy_core_alias=true for rollback."
                )
            logger.warning(message)
            return

    logger.warning("Top-level core SDK unavailable; enabling legacy core alias fallback")
    sys.modules.setdefault("core", backend_core_package)

    for module_name in _LEGACY_CORE_ALIAS_MODULES:
        backend_module_name = f"backend.core.{module_name}"
        try:
            module_obj = importlib.import_module(backend_module_name)
        except Exception as exc:
            logger.warning(
                "Failed to alias %s as core.%s: %s",
                backend_module_name,
                module_name,
                exc,
            )
            continue

        sys.modules.setdefault(f"core.{module_name}", module_obj)

if getattr(sys, 'frozen', False):
    # PyInstaller 打包后的路径处理
    # 目录模式: resources/backend/dist/api/api.exe
    # 需要往上跳两级到 resources/backend/
    EXE_PATH = sys.executable
    EXE_DIR = os.path.dirname(EXE_PATH)  # dist/api/
    base_dir = os.path.dirname(os.path.dirname(EXE_DIR))  # backend/
else:
    base_dir = API_DIR

BASE_DIR = base_dir

# ========== User Data Directory Resolution ==========
# Production (frozen): Electron 通过 --userdata-dir 传入 OS 用户数据目录
#   --> 所有可变数据写入 AppData，安装目录保持只读
# Development (not frozen): 直接使用 backend/ 目录，行为不变
# Fallback: AppData 不可用时回退到 BASE_DIR（只读，不 crash）

def _parse_userdata_dir_arg() -> Optional[str]:
    """从 sys.argv 解析 --userdata-dir= 参数。"""
    for arg in sys.argv:
        if arg.startswith('--userdata-dir='):
            val = arg.split('=', 1)[1].strip()
            if val:
                return val
    return None

_USERDATA_DIR_ARG = _parse_userdata_dir_arg()
_IS_FROZEN = getattr(sys, 'frozen', False)

# 仅在生产模式下启用 AppData
_USERDATA_ENABLED = bool(_IS_FROZEN and _USERDATA_DIR_ARG)
USER_DATA_DIR: Optional[str] = _USERDATA_DIR_ARG if _USERDATA_ENABLED else None

def _ensure_dir(path: str) -> None:
    """确保目录存在（幂等）。"""
    os.makedirs(path, exist_ok=True)

def _init_userdata_dir() -> bool:
    """
    初始化 AppData 目录结构并执行首次启动迁移。
    
    Returns:
        True 表示 AppData 初始化成功，False 表示回退到 BASE_DIR。
    """
    if not USER_DATA_DIR:
        return False
    
    try:
        # 创建目录结构
        _ensure_dir(os.path.join(USER_DATA_DIR, 'data'))
        _ensure_dir(os.path.join(USER_DATA_DIR, 'output', 'report'))
        _ensure_dir(os.path.join(USER_DATA_DIR, 'output', 'temp'))
        _ensure_dir(os.path.join(USER_DATA_DIR, 'output', 'logs'))
        _ensure_dir(os.path.join(USER_DATA_DIR, 'templates'))
        
        # 首次启动迁移：若 AppData 下无 config.yaml，从安装目录复制种子数据
        appdata_config = os.path.join(USER_DATA_DIR, 'config.yaml')
        if not os.path.exists(appdata_config):
            logger.info("First launch detected — migrating seed data to AppData...")
            seeds = [
                ('config.yaml', 'config.yaml'),
                (os.path.join('data', 'combined.db'), os.path.join('data', 'combined.db')),
                ('shared-config.json', 'shared-config.json'),
            ]
            for src_rel, dst_rel in seeds:
                src_path = os.path.join(BASE_DIR, src_rel)
                dst_path = os.path.join(USER_DATA_DIR, dst_rel)
                if not os.path.exists(dst_path) and os.path.exists(src_path):
                    _ensure_dir(os.path.dirname(dst_path))
                    shutil.copy2(src_path, dst_path)
                    logger.info(f"  Copied {src_rel} -> AppData")
                elif not os.path.exists(src_path):
                    logger.warning(f"  Seed file not found, skipped: {src_path}")
            logger.info("First-launch migration complete.")
        
        return True
    except Exception as e:
        logger.error(f"Failed to initialize AppData directory: {e}")
        logger.warning("Falling back to install directory (read-only mode).")
        return False

# 执行 AppData 初始化；失败时关闭 AppData 模式
if _USERDATA_ENABLED:
    if not _init_userdata_dir():
        _USERDATA_ENABLED = False
        USER_DATA_DIR = None

# 可变数据根目录：生产环境 = AppData，开发环境 = backend/
DATA_ROOT = USER_DATA_DIR if _USERDATA_ENABLED else BASE_DIR

logger.info(f"DATA_ROOT: {DATA_ROOT} (frozen={_IS_FROZEN}, appdata={_USERDATA_ENABLED})")

# ========== Path Definitions ==========
# 注意：内置模板仍在安装目录（只读），其他可变数据使用 DATA_ROOT

CONF_PATH = os.path.join(DATA_ROOT, "config.yaml")
SHARED_CONF_PATH = os.path.join(DATA_ROOT, "shared-config.json")

def load_config():
    if not os.path.exists(CONF_PATH):
        raise FileNotFoundError(f"Config file not found at {CONF_PATH}")
    with open(CONF_PATH, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

def _normalize_shared_config(raw: dict[str, Any]) -> dict[str, Any]:
    server = raw.get("server", {}) if isinstance(raw, dict) else {}
    app_meta = raw.get("app", {}) if isinstance(raw, dict) else {}
    security = raw.get("security", {}) if isinstance(raw, dict) else {}
    paths = raw.get("paths", {}) if isinstance(raw, dict) else {}
    plugin_runtime = raw.get("plugin_runtime", {}) if isinstance(raw, dict) else {}

    host = server.get("host", "127.0.0.1")
    if host == "localhost":
        host = "127.0.0.1"

    port = server.get("port", 8000)
    if not isinstance(port, int) or port < 1 or port > 65535:
        port = 8000

    runtime_mode = str(plugin_runtime.get("mode", "hybrid")).lower()
    if runtime_mode not in {"legacy", "hybrid", "descriptor", "isolated"}:
        runtime_mode = "hybrid"

    use_legacy_core_alias = plugin_runtime.get("use_legacy_core_alias", False)
    if not isinstance(use_legacy_core_alias, bool):
        use_legacy_core_alias = False

    force_legacy_templates = plugin_runtime.get("force_legacy_templates", [])
    if not isinstance(force_legacy_templates, list):
        force_legacy_templates = []

    subprocess_strategy = str(plugin_runtime.get("subprocess_strategy", "hybrid")).lower()
    if subprocess_strategy not in {"descriptor", "legacy", "hybrid"}:
        subprocess_strategy = "hybrid"

    try:
        subprocess_timeout_seconds = float(plugin_runtime.get("subprocess_timeout_seconds", 120))
    except (TypeError, ValueError):
        subprocess_timeout_seconds = 120.0
    if subprocess_timeout_seconds <= 0:
        subprocess_timeout_seconds = 120.0
    if subprocess_timeout_seconds > 600:
        subprocess_timeout_seconds = 600.0

    def _normalize_template_list(value: Any) -> list[str]:
        if not isinstance(value, list):
            return []
        return [str(item) for item in value if item]

    def _normalize_percent(value: Any, default: float) -> float:
        try:
            parsed = float(value)
        except (TypeError, ValueError):
            parsed = default
        if parsed < 0:
            return 0.0
        if parsed > 100:
            return 100.0
        return parsed

    isolated_enabled_templates = _normalize_template_list(plugin_runtime.get("isolated_enabled_templates", []))
    isolated_disabled_templates = _normalize_template_list(plugin_runtime.get("isolated_disabled_templates", []))
    isolated_rollout_percent = _normalize_percent(plugin_runtime.get("isolated_rollout_percent", 0), 0.0)

    template_rollout_raw = plugin_runtime.get("isolated_template_rollout", {})
    isolated_template_rollout: dict[str, float] = {}
    if isinstance(template_rollout_raw, dict):
        for key, value in template_rollout_raw.items():
            template_key = str(key).strip()
            if not template_key:
                continue
            isolated_template_rollout[template_key] = _normalize_percent(value, 0.0)

    isolated_fallback_mode = str(plugin_runtime.get("isolated_fallback_mode", "hybrid")).lower()
    if isolated_fallback_mode not in {"legacy", "hybrid", "descriptor"}:
        isolated_fallback_mode = "hybrid"

    metrics_emit_every_n = plugin_runtime.get("metrics_emit_every_n", 50)
    if not isinstance(metrics_emit_every_n, int) or metrics_emit_every_n <= 0:
        metrics_emit_every_n = 50
    if metrics_emit_every_n > 10000:
        metrics_emit_every_n = 10000

    return {
        "server": {
            "host": host if isinstance(host, str) and host else "127.0.0.1",
            "port": port,
        },
        "app": {
            "version": str(app_meta.get("version", "")),
        },
        "security": {
            "external_protocols": security.get("external_protocols", ["https:"]),
            "external_hosts": security.get("external_hosts", ["github.com", "www.github.com"]),
        },
        "paths": {
            "open_folder_allowlist": paths.get(
                "open_folder_allowlist",
                [os.path.join("output", "report"), os.path.join("output", "temp"), "output"],
            )
        },
        "plugin_runtime": {
            "mode": runtime_mode,
            "use_legacy_core_alias": use_legacy_core_alias,
            "force_legacy_templates": [str(item) for item in force_legacy_templates if item],
            "subprocess_strategy": subprocess_strategy,
            "subprocess_timeout_seconds": subprocess_timeout_seconds,
            "isolated_enabled_templates": isolated_enabled_templates,
            "isolated_disabled_templates": isolated_disabled_templates,
            "isolated_rollout_percent": isolated_rollout_percent,
            "isolated_template_rollout": isolated_template_rollout,
            "isolated_fallback_mode": isolated_fallback_mode,
            "metrics_emit_every_n": metrics_emit_every_n,
        }
    }


def load_shared_config():
    """加载共享配置（服务器端口、安全策略、路径白名单等）。"""
    if os.path.exists(SHARED_CONF_PATH):
        with open(SHARED_CONF_PATH, 'r', encoding='utf-8') as f:
            return _normalize_shared_config(json.load(f))

    return _normalize_shared_config({})


def persist_shared_config(next_config: dict[str, Any]) -> dict[str, Any]:
    """Normalize and persist shared runtime config to disk."""
    normalized = _normalize_shared_config(next_config)
    with open(SHARED_CONF_PATH, 'w', encoding='utf-8') as f:
        json.dump(normalized, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return normalized


def get_web_ui_config() -> dict[str, bool]:
    """Read web_ui config from shared-config.json with safe defaults."""
    global _web_ui_enabled
    try:
        if os.path.exists(SHARED_CONF_PATH):
            with open(SHARED_CONF_PATH, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            web_ui = raw.get("web_ui", {}) if isinstance(raw, dict) else {}
            enabled = web_ui.get("enabled", False) if isinstance(web_ui, dict) else False
            _web_ui_enabled = bool(enabled)
        else:
            _web_ui_enabled = False
    except Exception as exc:
        logger.warning(f"Failed to read web_ui config, defaulting to disabled: {exc}")
        _web_ui_enabled = False
    return {"enabled": _web_ui_enabled}


def set_web_ui_config(enabled: bool) -> dict[str, bool]:
    """Update web_ui.enabled in shared-config.json atomically."""
    global _web_ui_enabled
    try:
        enabled_bool = bool(enabled)
        _web_ui_enabled = enabled_bool

        if os.path.exists(SHARED_CONF_PATH):
            with open(SHARED_CONF_PATH, 'r', encoding='utf-8') as f:
                current = json.load(f)
        else:
            current = {}

        if not isinstance(current, dict):
            current = {}

        current["web_ui"] = {"enabled": enabled_bool}

        tmp_path = SHARED_CONF_PATH + ".tmp"
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json.dump(current, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp_path, SHARED_CONF_PATH)

        return {"enabled": enabled_bool}
    except Exception as exc:
        logger.error(f"Failed to persist web_ui config: {exc}")
        raise


config = load_config()
shared_config = load_shared_config()
_web_ui_enabled = False
APP_API_TOKEN = str(os.getenv("APP_API_TOKEN", "") or "")
_ensure_legacy_core_import_aliases(
    use_alias_fallback=bool(
        shared_config.get("plugin_runtime", {}).get("use_legacy_core_alias", False)
    ),
    fail_on_missing_core=True,
)

# @service/ endpoint mapping — templates reference shared framework services
# via "@service/name" in schema.yaml behavior endpoints instead of hardcoding
# framework API URLs. Add entries here when new shared endpoints are created.
SERVICE_MAP = {
    "@service/process-url": "/api/process-url",
    "@service/vulnerability-lookup": "/api/vulnerability",
}

# 延迟初始化：提升 macOS 启动速度
# 这些变量在首次访问时才会加载数据
_db_reader = None
_cached_vuln_list = None
_cached_vulnerabilities = None
_cached_icp_infos = None
_template_manager = None

TEMPLATES_BASE_DIR = os.path.join(BASE_DIR, "templates")  # 内置模板（安装目录，只读）
TEMPLATES_DIR = TEMPLATES_BASE_DIR  # 模板根目录
TEMPLATES_DELETED_DIR = os.path.join(TEMPLATES_BASE_DIR, "_deleted")  # 已删除模板备份目录（隐藏）

# 用户自定义模板目录（AppData，可读写）
USER_TEMPLATES_DIR: Optional[str] = (
    os.path.join(USER_DATA_DIR, "templates") if _USERDATA_ENABLED else None
)


def get_db_reader():
    """延迟初始化数据库读取器"""
    global _db_reader
    if _db_reader is None:
        _db_reader = DbDataReader(
            db_path=os.path.join(DATA_ROOT, config["vul_or_icp"])
        )
    return _db_reader


def get_cached_vulnerabilities():
    """延迟加载漏洞缓存"""
    global _cached_vuln_list, _cached_vulnerabilities
    if _cached_vuln_list is None or _cached_vulnerabilities is None:
        return reload_vulnerabilities_cache()
    return _cached_vuln_list, _cached_vulnerabilities


def get_cached_icp_infos():
    """延迟加载 ICP 备案缓存"""
    global _cached_icp_infos
    if _cached_icp_infos is None:
        return reload_icp_cache()
    return _cached_icp_infos


def reload_vulnerabilities_cache():
    """重新加载漏洞缓存"""
    global _cached_vuln_list, _cached_vulnerabilities
    _cached_vuln_list, _cached_vulnerabilities = get_db_reader().read_vulnerabilities_from_db()
    return _cached_vuln_list, _cached_vulnerabilities


def reload_icp_cache():
    """重新加载 ICP 缓存"""
    global _cached_icp_infos
    _cached_icp_infos = get_db_reader().read_Icp_from_db()
    return _cached_icp_infos


def _assert_template_handler_alignment(template_manager: TemplateManager) -> None:
    """校验模板与处理器注册一致性，避免运行期出现缺失处理器。

    支持两种处理器模式：
    1. HandlerRegistry 注册（旧类继承模式）
    2. PLUGIN descriptor（新纯函数模式）
    """
    import sys

    template_ids = set(template_manager.template_ids)
    registered_ids = set(HandlerRegistry.list_registered())

    missing_handlers = sorted(template_ids - registered_ids)

    # PLUGIN descriptor 是有效的替代注册方式
    still_missing = []
    for tid in missing_handlers:
        module_name = f"templates.{tid}.handler"
        module = sys.modules.get(module_name)
        if module is not None and getattr(module, "PLUGIN", None) is not None:
            logger.debug(f"Template '{tid}' uses PLUGIN descriptor mode")
            continue
        still_missing.append(tid)

    extra_handlers = sorted(registered_ids - template_ids)

    if extra_handlers:
        logger.warning(f"Detected extra handlers without matching templates: {extra_handlers}")

    if still_missing:
        logger.critical(f"Template handlers missing for templates: {still_missing}")
        raise RuntimeError(
            "Template handler registry mismatch. "
            f"Missing handlers for templates: {', '.join(still_missing)}"
        )


def get_template_manager():
    """
    延迟初始化模板管理器
    
    模板管理器会自动扫描 templates/ 目录并动态加载所有 handler.py
    无需手动导入（阶段 1：任务 1.4）
    """
    global _template_manager
    if _template_manager is None:
        _template_manager = TemplateManager(
            TEMPLATES_DIR, config,
            user_templates_dir=USER_TEMPLATES_DIR
        )
        _assert_template_handler_alignment(_template_manager)
        
        # 动态挂载模板路由
        for template_id, router in _template_manager.get_template_routers().items():
            # Security Fix: 强制路由隔离，防止模板覆盖系统API或越权
            # 所有模板定义的 API 必须通过 /api/plugin/{template_id}/ 访问
            safe_prefix = f"/api/plugin/{template_id}"
            app.include_router(
                router, 
                prefix=safe_prefix,
                tags=[f"Plugin: {template_id}"]
            )
            logger.info(f"Mounted router for template: {template_id} at {safe_prefix}")
            
        # 输出已注册的 Handler（由动态加载自动注册）
        logger.info(f"Registered handlers: {HandlerRegistry.list_registered()}")
    return _template_manager


# ========== FastAPI 依赖注入类型别名 ==========
# 使用 Annotated 模式减少重复的 get_xxx() 调用
DbReaderDep = Annotated[DbDataReader, Depends(get_db_reader)]
TemplateManagerDep = Annotated[TemplateManager, Depends(get_template_manager)]


def get_vulnerabilities_cache() -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    """依赖注入：获取漏洞缓存"""
    return get_cached_vulnerabilities()


def get_icp_cache() -> dict[str, dict[str, Any]]:
    """依赖注入：获取 ICP 缓存"""
    return get_cached_icp_infos()


VulnCacheDep = Annotated[
    tuple[list[dict[str, Any]], dict[str, dict[str, Any]]],
    Depends(get_vulnerabilities_cache),
]
IcpCacheDep = Annotated[dict[str, dict[str, Any]], Depends(get_icp_cache)]


# ========== 通用响应辅助函数 ==========
def success_response(message: str = "操作成功", **kwargs) -> dict[str, Any]:
    """生成成功响应"""
    return {"success": True, "message": message, **kwargs}


def error_response(message: str, detail: Optional[str] = None) -> dict[str, Any]:
    """生成错误响应"""
    return {"success": False, "message": message, "detail": detail}


def handle_db_result(success: bool, msg: str, reload_cache_fn=None) -> dict[str, Any]:
    """
    统一处理数据库操作结果
    
    Args:
        success: 操作是否成功
        msg: 操作消息
        reload_cache_fn: 成功后需要调用的缓存刷新函数
    """
    if success:
        if reload_cache_fn:
            reload_cache_fn()
        return success_response(msg)
    raise HTTPException(status_code=400, detail=msg)


def _normalize_version_string(raw: Any) -> str:
    """归一化版本号：去除前缀 V/v 并返回字符串。"""
    if raw is None:
        return ""
    version = str(raw).strip()
    return version[1:] if version.lower().startswith("v") else version


def _is_newer_version(current: str, latest: str) -> bool:
    """Return True if latest is strictly newer than current (semver comparison)."""
    if not current or not latest:
        return False
    try:
        cur = [int(x) for x in current.split(".")]
        lat = [int(x) for x in latest.split(".")]
    except (ValueError, TypeError):
        return False
    return lat > cur


def _raise_internal_error(action: str, exc: Exception) -> None:
    """统一兜底异常日志与 HTTP 500。"""
    logger.error(f"{action} failed: {traceback.format_exc()}")
    raise HTTPException(status_code=500, detail=f"{action} failed") from exc


def _build_error_payload(message: str, detail: Optional[str] = None, **extra: Any) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "success": False,
        "message": message,
        "error": message,
        "detail": detail,
    }
    if extra:
        payload.update(extra)
    return payload


def _error_response(status_code: int, message: str, detail: Optional[str] = None, **extra: Any) -> JSONResponse:
    return JSONResponse(status_code=status_code, content=_build_error_payload(message, detail, **extra))


def _is_token_protected_get_path(path: str) -> bool:
    if path in {"/api/backup-db", "/api/health-auth"}:
        return True

    # 模板导出是敏感读取（可获取模板源码/资源）
    if re.match(r"^/api/templates/[^/]+/export$", path):
        return True

    return False


def _requires_app_token(method: str, path: str) -> bool:
    """Only protect sensitive API requests; keep read-only GET APIs open by default."""
    normalized_method = (method or "").upper()
    if normalized_method == "OPTIONS":
        return False

    if normalized_method in {"POST", "PUT", "PATCH", "DELETE"} and path.startswith("/api/"):
        return True

    if normalized_method == "GET" and _is_token_protected_get_path(path):
        return True

    return False


class ConfigResponse(BaseModel):
    """返回给前端的初始化配置信息"""
    version: str
    supplierName: str
    hazard_types: list[str]
    unit_types: list[str]
    industries: list[str]
    vulnerabilities_list: list[dict[str, str]]


class VersionResponse(BaseModel):
    """版本一致性检查响应"""
    backend_version: str
    shared_version: str
    is_synced: bool


class UrlProcessRequest(BaseModel):
    url: str

class UrlProcessResponse(BaseModel):
    url: str
    domain: str
    ip: str
    icp_info: Optional[dict[str, Any]] = None

class CustomVulnRequest(BaseModel):
    id: Optional[str] = None  # 编辑时使用，新增时为空
    name: str
    category: Optional[str] = ""
    port: Optional[str] = ""
    level: str
    basis: Optional[str] = ""
    description: str
    impact: Optional[str] = ""
    suggestion: str

class UploadImageRequest(BaseModel):
    image_base64: str
    filename: Optional[str] = "image.png"

class UploadImageResponse(BaseModel):
    file_path: str
    url: str

class MergeRequest(BaseModel):
    files: list[str]
    output_filename: Optional[str] = "Merged_Report.docx"

class MergeResponse(BaseModel):
    success: bool
    message: str
    file_path: str
    download_url: str

class DeleteFileRequest(BaseModel):
    path: str

class IcpEntryRequest(BaseModel):
    unitName: Optional[str] = ""
    natureName: Optional[str] = ""
    domain: Optional[str] = ""
    mainLicence: Optional[str] = ""
    serviceLicence: Optional[str] = ""
    updateRecordTime: Optional[str] = ""

class BatchDeleteRequest(BaseModel):
    ids: list[str]

class UpdateConfigRequest(BaseModel):
    supplierName: str


class PluginRuntimeConfigRequest(BaseModel):
    mode: Optional[str] = None
    use_legacy_core_alias: Optional[bool] = None
    force_legacy_templates: Optional[list[str]] = None
    subprocess_strategy: Optional[str] = None
    subprocess_timeout_seconds: Optional[float] = None
    isolated_enabled_templates: Optional[list[str]] = None
    isolated_disabled_templates: Optional[list[str]] = None
    isolated_rollout_percent: Optional[float] = None
    isolated_template_rollout: Optional[dict[str, float]] = None
    isolated_fallback_mode: Optional[str] = None
    metrics_emit_every_n: Optional[int] = None


class WebUIConfigRequest(BaseModel):
    enabled: bool

class OpenFolderRequest(BaseModel):
    """打开文件夹请求"""
    path: Optional[str] = None

class ListReportsRequest(BaseModel):
    """列出报告请求（预留分页/过滤）"""
    page: Optional[int] = 1
    limit: Optional[int] = 100
    folder: Optional[str] = None

# ========== Lifespan 生命周期管理 (替代废弃的 @app.on_event) ==========
@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    应用生命周期管理器
    - startup: 初始化模板管理器
    - shutdown: 清理资源
    """
    # Startup
    get_template_manager()
    get_web_ui_config()
    logger.info("Template system initialized")
    yield
    # Shutdown (如有需要可在此添加清理逻辑)
    logger.info("Application shutting down")


app = FastAPI(lifespan=lifespan)

# 域路由（首批拆分：config / vulnerabilities / icp）
config_router = APIRouter(tags=["Config"])
vuln_router = APIRouter(tags=["Vulnerability"])
icp_router = APIRouter(tags=["ICP"])


# ========== 全局异常处理器 (统一错误响应) ==========
@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(request: Request, exc: StarletteHTTPException):
    """处理 HTTP 异常，返回统一格式"""
    detail_text = str(exc.detail)
    return _error_response(exc.status_code, detail_text, detail_text)


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """处理请求验证错误，返回统一格式"""
    errors = exc.errors()
    error_messages = [f"{err['loc']}: {err['msg']}" for err in errors]
    return _error_response(422, "Validation Error", "; ".join(error_messages))


@app.exception_handler(Exception)
async def general_exception_handler(request: Request, exc: Exception):
    """处理未捕获的异常"""
    logger.error(f"Unhandled exception: {traceback.format_exc()}")
    return _error_response(500, "Internal Server Error", str(exc) if os.getenv("DEBUG") else None)


class WebUIBlockMiddleware(BaseHTTPMiddleware):
    """Block /ui requests when web_ui.enabled is False."""

    async def dispatch(self, request: Request, call_next):
        if request.url.path.startswith("/ui") and not _web_ui_enabled:
            return JSONResponse({"detail": "Web UI is disabled"}, status_code=404)
        return await call_next(request)


@app.middleware("http")
async def app_token_middleware(request: Request, call_next):
    """Protect sensitive APIs with per-launch token when token is configured."""
    if APP_API_TOKEN and _requires_app_token(request.method, request.url.path):
        provided = request.headers.get("X-App-Token", "") or request.query_params.get("app_token", "")
        if provided != APP_API_TOKEN:
            return _error_response(403, "Forbidden", "Invalid application token")

    return await call_next(request)


app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1",
        "http://localhost",
        "null",
    ],
    allow_origin_regex=r"^http://(127\.0\.0\.1|localhost)(:\d+)?$",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OUTPUT_DIR = os.path.join(DATA_ROOT, "output", "report")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
app.mount("/reports", StaticFiles(directory=OUTPUT_DIR), name="reports")

TEMP_DIR = os.path.join(DATA_ROOT, "output", "temp")
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)
app.mount("/temp", StaticFiles(directory=TEMP_DIR), name="temp")

SRC_DIR = os.path.join(BASE_DIR, "..", "src")
if os.path.exists(SRC_DIR):
    app.mount("/ui", StaticFiles(directory=SRC_DIR, html=True), name="frontend")

app.add_middleware(WebUIBlockMiddleware)


def _is_subpath(path_to_check: str, base_path: str) -> bool:
    """安全判断：path_to_check 是否位于 base_path 内。"""
    try:
        target_real = os.path.realpath(path_to_check)
        base_real = os.path.realpath(base_path)
        return os.path.commonpath([target_real, base_real]) == base_real
    except ValueError:
        return False


def _build_open_folder_allowlist() -> list[str]:
    defaults = [
        os.path.join("output", "report"),
        os.path.join("output", "temp"),
        "output",
    ]
    configured = shared_config.get("paths", {}).get("open_folder_allowlist", defaults)
    if not isinstance(configured, list):
        configured = defaults
    resolved: list[str] = []
    # 相对路径基于 DATA_ROOT 解析（生产环境 = AppData，开发环境 = backend/）
    for relative_path in configured:
        if not isinstance(relative_path, str):
            continue
        resolved.append(os.path.realpath(os.path.join(DATA_ROOT, relative_path)))

    if not resolved:
        resolved = [os.path.realpath(OUTPUT_DIR)]

    return resolved


OPEN_FOLDER_ALLOWLIST = _build_open_folder_allowlist()


def _is_allowed_open_folder(path_to_check: str) -> bool:
    return any(_is_subpath(path_to_check, allowed_base) for allowed_base in OPEN_FOLDER_ALLOWLIST)


@app.get("/")
def read_root():
    return {"message": "ReportGenX Backend is Running", "version": config["version"]}


@app.get("/api/health")
def health_check():
    return {"success": True, "status": "ok"}


@app.get("/api/health-auth")
def health_check_auth():
    """Token-protected liveness endpoint for Electron main process handshake."""
    return {"success": True, "status": "ok", "auth": True}

@config_router.get("/api/config", response_model=ConfigResponse)
def get_config():
    """获取初始化配置数据，填充前端下拉框"""
    def clean_list(lst):
        if not lst:
            return []
        return [str(item) for item in lst if item]

    return {
        "version": config["version"],
        "supplierName": config["supplierName"],
        "hazard_types": clean_list(config.get("hazard_type", [])),
        "unit_types": clean_list(config.get("unitType", [])),
        "industries": clean_list(config.get("industry", [])),
        "vulnerabilities_list": get_cached_vulnerabilities()[0]
    }

@config_router.get("/api/frontend-config")
def get_frontend_config():
    """获取前端全局配置（风险等级颜色等共享配置）"""
    return {
        "risk_levels": config.get("risk_levels", []),
        "operating_systems": config.get("operating_systems", []),
        "version": config["version"]
    }

@config_router.get("/api/version", response_model=VersionResponse)
def get_version_info():
    """返回后端与共享配置版本，用于发布前与运行时一致性校验。"""
    backend_version = _normalize_version_string(config.get("version", ""))
    shared_version = _normalize_version_string(shared_config.get("app", {}).get("version", ""))
    return {
        "backend_version": backend_version,
        "shared_version": shared_version,
        "is_synced": bool(backend_version and shared_version and backend_version == shared_version),
    }

_update_check_cache = {"last_check": 0, "data": None}
UPDATE_CHECK_TTL = 3600

@config_router.get("/api/check-update")
def check_update():
    """Check for new version on GitHub Releases. Cached for 1 hour."""
    global _update_check_cache
    now = time.time()
    if _update_check_cache["last_check"] and (now - _update_check_cache["last_check"]) < UPDATE_CHECK_TTL:
        return _update_check_cache["data"]

    current_version = _normalize_version_string(config.get("version", ""))
    try:
        req = urllib.request.Request(
            "https://api.github.com/repos/s1g0day/ReportGenX/releases/latest",
            headers={"Accept": "application/vnd.github+json", "User-Agent": "ReportGenX"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            release = json.loads(resp.read())
        tag = release.get("tag_name", "")
        latest_version = _normalize_version_string(tag)
        result = {
            "has_update": _is_newer_version(current_version, latest_version),
            "current_version": current_version,
            "latest_version": latest_version,
            "download_url": release.get("html_url", "https://github.com/s1g0day/ReportGenX/releases"),
            "release_notes": release.get("body", ""),
            "error": None,
        }
    except Exception as e:
        result = {"has_update": False, "current_version": current_version, "latest_version": "", "download_url": "", "release_notes": "", "error": str(e)}

    _update_check_cache = {"last_check": now, "data": result}
    return result


@config_router.get("/api/plugin-runtime-config")
def get_plugin_runtime_config():
    return {
        "success": True,
        "plugin_runtime": shared_config.get("plugin_runtime", {}),
    }


@config_router.post("/api/plugin-runtime-config")
def update_plugin_runtime_config(req: PluginRuntimeConfigRequest):
    """更新 plugin_runtime 配置（用于隔离模式灰度开关）。"""
    global shared_config

    try:
        updates = req.model_dump(exclude_none=True)
        current = dict(shared_config)
        plugin_runtime_current = current.get("plugin_runtime", {})
        if not isinstance(plugin_runtime_current, dict):
            plugin_runtime_current = {}

        plugin_runtime_next = dict(plugin_runtime_current)
        plugin_runtime_next.update(updates)
        current["plugin_runtime"] = plugin_runtime_next

        shared_config = persist_shared_config(current)
        return {
            "success": True,
            "message": "Plugin runtime config updated",
            "plugin_runtime": shared_config.get("plugin_runtime", {}),
        }
    except Exception as exc:
        _raise_internal_error("update plugin runtime config", exc)


@config_router.get("/api/web-ui-config")
def get_web_ui_config_endpoint():
    """返回 web_ui 配置（enabled 状态）。"""
    return get_web_ui_config()


@config_router.post("/api/web-ui-config")
def update_web_ui_config_endpoint(req: WebUIConfigRequest):
    """更新 web_ui.enabled 配置。"""
    try:
        result = set_web_ui_config(req.enabled)
        return {"success": True, "enabled": result["enabled"]}
    except Exception as exc:
        _raise_internal_error("update web ui config", exc)


@config_router.get("/api/backup-db")
def backup_database():
    """下载数据库备份"""
    db_path = os.path.join(DATA_ROOT, config["vul_or_icp"])
    if os.path.exists(db_path):
        filename = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        return FileResponse(path=db_path, filename=filename, media_type='application/x-sqlite3')
    raise HTTPException(status_code=404, detail="Database file not found")

@vuln_router.get("/api/vulnerability/{id_or_name}")
def get_vulnerability_detail(id_or_name: str):
    """根据 ID 或名称获取漏洞详情"""
    _, cached_vulns = get_cached_vulnerabilities()
    if id_or_name in cached_vulns:
        return cached_vulns[id_or_name]
    
    for v in cached_vulns.values():
        val_name = v.get('Vuln_Name')
        if val_name and (val_name == id_or_name or val_name.lower() == id_or_name.lower()):
            return v
    
    desc, sol = get_db_reader().get_vulnerability_info(id_or_name)
    if desc:
        return {
            "Vuln_Description": desc,
            "Repair_suggestions": sol,
            "Vuln_Name": id_or_name,
            "Vuln_data_source": "legacy_fallback"
        }

    return {"error": "Vulnerability not found"}

def _process_url_value(url_text: str) -> dict[str, Any]:
    """处理输入的 URL/IP，解析域名和 IP，查找匹配的 ICP 备案信息。"""
    text = (url_text or "").strip()
    result = {
        "url": text,
        "domain": "",
        "ip": "",
        "icp_info": None
    }

    if not text:
        return result
    
    ip_pattern = r'^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    
    if re.match(ip_pattern, text):
        result["ip"] = text
    else:
        extracted = tldextract.extract(text)
        if extracted.domain and extracted.suffix:
            result["domain"] = f"{extracted.domain}.{extracted.suffix}"
            full_domain = f"{extracted.subdomain}.{extracted.domain}.{extracted.suffix}" if extracted.subdomain else result["domain"]
            
            try:
                result["ip"] = socket.gethostbyname(full_domain)
            except (socket.gaierror, socket.herror, OSError):
                pass
        elif text.startswith('http'):
            try:
                parsed = urlparse(text)
                netloc = parsed.netloc
                host = netloc.split(':')[0] if ':' in netloc else netloc
                result["domain"] = host
                if re.match(ip_pattern, host):
                    result["ip"] = host
            except (ValueError, AttributeError):
                pass

    if result["domain"]:
        domain_key = result["domain"].lower()
        cached_icp = get_cached_icp_infos()
        logger.debug(f"Looking for domain: {domain_key}")
        logger.debug(f"Available domains (first 10): {list(cached_icp.keys())[:10]}")
        
        if domain_key in cached_icp:
            result["icp_info"] = cached_icp[domain_key]
            logger.debug(f"Direct match found!")
        else:
            parts = domain_key.split('.')
            if len(parts) > 2:
                root_domain = f"{parts[-2]}.{parts[-1]}"
                logger.debug(f"Trying root domain: {root_domain}")
                if root_domain in cached_icp:
                    result["icp_info"] = cached_icp[root_domain]
                    logger.debug(f"Root domain match found!")
            
            if not result["icp_info"]:
                for cached_domain in cached_icp.keys():
                    if domain_key.endswith('.' + cached_domain) or domain_key == cached_domain:
                        result["icp_info"] = cached_icp[cached_domain]
                        logger.debug(f"Suffix match found: {cached_domain}")
                        break
        
        if not result["icp_info"]:
            logger.debug(f"No match found for: {domain_key}")
        else:
            info = result['icp_info']
            logger.debug(f"Found ICP info for: {domain_key}")
            if isinstance(info, dict):
                logger.debug(f"unitName bytes: {repr(str(info.get('unitName', '')).encode('utf-8'))}")

    return result


@app.post("/api/process-url", response_model=UrlProcessResponse)
def process_url(req: UrlProcessRequest):
    """POST: 处理输入的 URL/IP，解析域名和 IP，查找匹配的 ICP 备案信息。"""
    return _process_url_value(req.url)


@app.get("/api/process-url", response_model=UrlProcessResponse)
def process_url_get(url: str = ""):
    """GET 兼容入口：支持 query 参数 url，避免旧客户端触发 405。"""
    return _process_url_value(url)

@app.post("/api/upload-image", response_model=UploadImageResponse)
def upload_image(req: UploadImageRequest):
    """接收 Base64 图片，保存为临时文件，返回绝对路径 (Security Fixed)"""
    try:
        if ',' in req.image_base64:
            header, encoded = req.image_base64.split(',', 1)
        else:
            encoded = req.image_base64

        data = base64.b64decode(encoded)
        
        # 安全检查 1: 文件大小限制 (10MB)
        if len(data) > 10 * 1024 * 1024:
            raise HTTPException(status_code=400, detail="Image size exceeds 10MB limit")

        # 安全检查 2: 验证图片格式与完整性
        try:
            with Image.open(io.BytesIO(data)) as img:
                img.verify()  # 验证文件结构是否损坏
                
                # 重新打开以读取属性（verify后需重新打开）
                with Image.open(io.BytesIO(data)) as valid_img:
                    img_format = valid_img.format
                    if img_format is None:
                        raise HTTPException(status_code=400, detail="Unable to determine image format")
                    fmt = img_format.lower()
                    
                    # 白名单格式检查
                    ALLOWED_FORMATS = {'png', 'jpeg', 'jpg', 'gif', 'bmp'}
                    if fmt not in ALLOWED_FORMATS:
                        raise HTTPException(status_code=400, detail=f"Unsupported image format: {fmt}")
                    
                    # 像素炸弹防护 (限制分辨率 10000x10000)
                    if valid_img.width * valid_img.height > 100000000:
                         raise HTTPException(status_code=400, detail="Image dimensions too large")
                        
                    # 确定安全的文件后缀
                    ext = f".{fmt}" if fmt != 'jpeg' else '.jpg'
        except Exception as e:
            if isinstance(e, HTTPException):
                raise e
            logger.warning(f"Invalid image upload attempt: {str(e)}")
            raise HTTPException(status_code=400, detail="Invalid image file")

        # 使用生成的安全文件名
        filename = f"{uuid.uuid4()}{ext}"
        file_path = os.path.join(TEMP_DIR, filename)
        
        with open(file_path, "wb") as f:
            f.write(data)
            
        return {
            "file_path": file_path,
            "url": f"/temp/{filename}"
        }
    except Exception as e:
        if isinstance(e, HTTPException):
            raise e
        logger.error(f"Upload failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@vuln_router.get("/api/vulnerabilities")
def get_vulnerabilities():
    """获取所有漏洞列表及其详情"""
    try:
        # 显式重载缓存，避免数据不一致
        _, cached_vulns = reload_vulnerabilities_cache()
        if not cached_vulns:
             return []
        
        # 将字典的值转换为列表返回
        return list(cached_vulns.values())
    except Exception as e:
        _raise_internal_error("fetch vulnerabilities", e)

@vuln_router.post("/api/vulnerabilities")
def add_vulnerability(vuln: CustomVulnRequest, db: DbReaderDep):
    """添加新漏洞"""
    # 移除 id 字段（如果存在），因为新增时不需要
    data = vuln.model_dump(exclude={'id'})
    success, msg = db.add_vulnerability_to_db(data)
    return handle_db_result(success, msg, reload_vulnerabilities_cache)

@vuln_router.put("/api/vulnerabilities/{Vuln_id}")
def update_vulnerability(Vuln_id: str, vuln: CustomVulnRequest, db: DbReaderDep):
    """更新漏洞信息"""
    # 移除 id 字段（使用 URL 路径参数中的 Vuln_id）
    data = vuln.model_dump(exclude={'id'})
    success, msg = db.update_vulnerability_in_db(Vuln_id, data)
    return handle_db_result(success, msg, reload_vulnerabilities_cache)

@vuln_router.delete("/api/vulnerabilities/{Vuln_id}")
def delete_vulnerability(Vuln_id: str, db: DbReaderDep):
    """删除漏洞"""
    success, msg = db.delete_vulnerability_from_db(Vuln_id)
    return handle_db_result(success, msg, reload_vulnerabilities_cache)

@config_router.post("/api/open-folder")
def open_folder(req: OpenFolderRequest):
    """打开文件所在目录（带路径白名单校验）。"""
    requested_path = req.path
    if not requested_path or requested_path == "default":
        requested_path = OUTPUT_DIR

    normalized_path = os.path.realpath(os.path.abspath(os.path.normpath(requested_path)))

    if os.path.isfile(normalized_path):
        target_dir = os.path.dirname(normalized_path)
    elif os.path.isdir(normalized_path):
        target_dir = normalized_path
    else:
        # 兼容“文件已删除但目录仍在”的场景
        parent_dir = os.path.dirname(normalized_path)
        if not os.path.isdir(parent_dir):
            raise HTTPException(status_code=404, detail="Path does not exist")
        target_dir = parent_dir

    if not _is_allowed_open_folder(target_dir):
        raise HTTPException(status_code=403, detail="Path is not allowed")

    try:
        if os.name == "nt":
            os.startfile(target_dir)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target_dir])
        else:
            subprocess.Popen(["xdg-open", target_dir])
        return success_response("Folder opened", path=target_dir)
    except OSError as exc:
        _raise_internal_error("open folder", exc)

@app.post("/api/merge-reports", response_model=MergeResponse)
def merge_reports(req: MergeRequest):
    """合并多个报告"""
    try:
        # 确定输出路径
        # 默认放到 output/report/合并报告 下
        MERGE_DIR = os.path.join(OUTPUT_DIR, "Combined")
        if not os.path.exists(MERGE_DIR):
            os.makedirs(MERGE_DIR)
            
        # 如果未指定文件名或为空，生成一个带时间戳的
        filename = req.output_filename or f"Merged_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
        filename = os.path.basename(filename)
        filename = re.sub(r'[<>:"/\\|?*]+', '_', filename).strip()
        if not filename:
            filename = f"Merged_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
        if not filename.endswith('.docx'):
            filename += '.docx'
            
        output_path = os.path.realpath(os.path.join(MERGE_DIR, filename))
        if not _is_subpath(output_path, MERGE_DIR):
            raise HTTPException(status_code=400, detail="Invalid output filename")
        
        success, msg = ReportMerger.merge_reports(req.files, output_path)
        
        if success:
            return {
                "success": True,
                "message": "合并成功",
                "file_path": output_path,
                "download_url": "" # 暂不需要下载URL，本地打开即可
            }
        else:
            return {
                "success": False,
                "message": msg,
                "file_path": "",
                "download_url": ""
            }
    except Exception as e:
        _raise_internal_error("merge reports", e)

@app.post("/api/list-reports")
def list_reports(req: Optional[ListReportsRequest] = None):
    """列出生成的报告文件，供选择合并"""
    # req 预留用于分页或过滤
    try:
        # 扫描 OUTPUT_DIR (及其第一级子文件夹?)
        # 假设报告都存在于 output/report/下面，可能按公司名分了文件夹
        # 我们深度遍历一下，但不要太深
        
        report_files = []
        for root, dirs, files in os.walk(OUTPUT_DIR):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~$'):
                    full_path = os.path.join(root, file)
                    stat = os.stat(full_path)
                    report_files.append({
                        "path": full_path,
                        "name": file,
                        "mtime": stat.st_mtime,
                        "date": datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                        "folder": os.path.basename(root)
                    })
        
        # 按修改时间倒序
        report_files.sort(key=lambda x: x['mtime'], reverse=True)
        return report_files
    except Exception as e:
        _raise_internal_error("list reports", e)

@app.post("/api/delete-report")
def delete_report(req: DeleteFileRequest):
    """删除指定的报告文件"""
    try:
        file_path = req.path
        # 安全检查：确保文件在 output 目录下
        abs_path = os.path.realpath(os.path.abspath(file_path))

        if not _is_subpath(abs_path, OUTPUT_DIR):
             return {"success": False, "message": "无法删除该目录下的文件 (Permission Denied)"}

        if os.path.exists(abs_path) and os.path.isfile(abs_path):
            os.remove(abs_path)
            return {"success": True, "message": "文件已删除"}
        else:
            return {"success": False, "message": "文件不存在或已被删除"}
    except Exception as e:
        _raise_internal_error("delete report", e)

@icp_router.get("/api/icp-columns")
def get_icp_columns():
    """获取 ICP 表的所有字段名"""
    return get_db_reader().get_table_columns("icp_info")

@icp_router.get("/api/icp-list")
def list_icp_entries(db: DbReaderDep):
    """获取所有 ICP 信息 (直接读取数据库)"""
    return db.read_icp_raw_list()

@icp_router.delete("/api/icp-entry/{vuln_id}")
def delete_icp_entry(vuln_id: str, db: DbReaderDep):
    """删除指定的 ICP 信息 (根据 Vuln_id)"""
    success, msg = db.delete_icp_entry(vuln_id)
    if success:
        reload_icp_cache()
        return success_response(msg)
    if "未找到" in msg:
        return error_response(msg)
    raise HTTPException(status_code=500, detail=msg)

@icp_router.post("/api/icp-entry")
def add_icp_entry(req: IcpEntryRequest, db: DbReaderDep):
    """新增 ICP 信息"""
    success, msg = db.add_icp_entry(req.model_dump())
    return handle_db_result(success, msg)

@icp_router.put("/api/icp-entry/{vuln_id}")
def update_icp_entry(vuln_id: str, req: IcpEntryRequest, db: DbReaderDep):
    """更新 ICP 信息"""
    success, msg = db.update_icp_entry(vuln_id, req.model_dump())
    return handle_db_result(success, msg)

@icp_router.post("/api/icp-batch-delete")
def batch_delete_icp(req: BatchDeleteRequest, db: DbReaderDep):
    """批量删除 ICP 信息"""
    success, msg = db.batch_delete_icp(req.ids)
    return handle_db_result(success, msg)

@icp_router.post("/api/icp-import")
async def import_icp(file: UploadFile = File(...), overwrite: bool = Form(default=False)):
    """Import ICP entries from Excel (.xlsx/.xls). Uses single-connection transaction."""
    import openpyxl
    import sqlite3
    from io import BytesIO

    CH_TO_EN = {
        'domain': 'domain', '域名': 'domain',
        'unitName': 'unitName', '单位名称': 'unitName',
        'natureName': 'natureName', '单位性质': 'natureName', '性质': 'natureName',
        'mainLicence': 'mainLicence', '备案号': 'mainLicence', 'ICP备案号': 'mainLicence',
        'serviceLicence': 'serviceLicence', '许可证号': 'serviceLicence', 'ICP许可证号': 'serviceLicence',
        'updateRecordTime': 'updateRecordTime', '更新时间': 'updateRecordTime',
    }

    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(BytesIO(contents), read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return {"success": False, "error": "Empty file"}

        headers = [str(h).strip() if h else '' for h in rows[0]]
        column_map = {}
        for i, h in enumerate(headers):
            en = CH_TO_EN.get(h, h)
            if en and en not in column_map:
                column_map[i] = en

        if 'domain' not in column_map.values():
            return {"success": False, "error": "Excel must have a 'domain' (域名) column"}

        db = get_db_reader()
        db_path = db.db_path
        conn = sqlite3.connect(db_path)
        db._ensure_column_exists(conn, "icp_info", "Vuln_id")
        cursor = conn.cursor()
        imported = 0
        skipped = 0
        errors = []

        ICP_FIELDS = ['Vuln_id', 'unitName', 'natureName', 'domain', 'mainLicence', 'serviceLicence', 'updateRecordTime']

        try:
            conn.execute("BEGIN")
            for row_idx, row in enumerate(rows[1:], start=2):
                entry = {}
                for col_idx, en_name in column_map.items():
                    val = row[col_idx] if col_idx < len(row) else ''
                    entry[en_name] = str(val).strip() if val else ''

                domain = entry.get('domain', '')
                if not domain:
                    errors.append(f"Row {row_idx}: missing domain, skipped")
                    skipped += 1
                    continue

                vuln_id = hashlib.sha256(domain.encode()).hexdigest()[:16]

                cursor.execute("SELECT 1 FROM icp_info WHERE Vuln_id = ?", (vuln_id,))
                exists = cursor.fetchone() is not None

                if exists:
                    if overwrite:
                        updates = ', '.join([f"{f}=?" for f in ICP_FIELDS if f != 'Vuln_id'])
                        vals = [entry.get(f, '') for f in ICP_FIELDS if f != 'Vuln_id'] + [vuln_id]
                        cursor.execute(f"UPDATE icp_info SET {updates} WHERE Vuln_id=?", vals)
                        imported += 1
                    else:
                        skipped += 1
                else:
                    placeholders = ','.join(['?'] * len(ICP_FIELDS))
                    col_str = ','.join(ICP_FIELDS)
                    vals = [vuln_id] + [entry.get(f, '') for f in ICP_FIELDS if f != 'Vuln_id']
                    cursor.execute(f"INSERT INTO icp_info ({col_str}) VALUES ({placeholders})", vals)
                    imported += 1

            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

        reload_icp_cache()
        return {"success": True, "imported": imported, "skipped": skipped, "errors": errors}
    except Exception as e:
        logger.error(f"ICP import failed: {e}")
        return {"success": False, "error": str(e)}


@vuln_router.post("/api/vulnerabilities/import")
async def import_vulnerabilities(file: UploadFile = File(...), overwrite: bool = Form(default=False)):
    """Import vulnerabilities from Excel (.xlsx/.xls). Single-connection transaction."""
    import openpyxl
    import sqlite3
    from io import BytesIO

    CH_TO_EN = {
        'name': 'name', '漏洞名称': 'name',
        'category': 'category', '漏洞分类': 'category',
        'port': 'port', '默认端口': 'port',
        'level': 'level', '风险级别': 'level',
        'basis': 'basis', '定级依据': 'basis',
        'description': 'description', '漏洞描述': 'description',
        'impact': 'impact', '漏洞危害': 'impact',
        'suggestion': 'suggestion', '加固建议': 'suggestion', '修复建议': 'suggestion',
    }

    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(BytesIO(contents), read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return {"success": False, "error": "Empty file"}

        headers = [str(h).strip() if h else '' for h in rows[0]]
        column_map = {}
        for i, h in enumerate(headers):
            en = CH_TO_EN.get(h, h)
            if en and en not in column_map:
                column_map[i] = en

        if 'name' not in column_map.values():
            return {"success": False, "error": "Excel must have a 'name' (漏洞名称) column"}

        db = get_db_reader()
        db_path = db.db_path
        conn = sqlite3.connect(db_path)
        db._ensure_column_exists(conn, "vulnerabilities", "Vuln_id")
        db._ensure_column_exists(conn, "vulnerabilities", "Class_basis")
        db._ensure_column_exists(conn, "vulnerabilities", "Vuln_Hazards")
        cursor = conn.cursor()
        imported = 0
        skipped = 0
        errors = []

        VULN_FIELDS = ['Vuln_id', 'Vuln_Name', 'Vuln_Class', 'Default_port',
                       'Risk_Level', 'Class_basis', 'Vuln_Description',
                       'Vuln_Hazards', 'Repair_suggestions']
        FIELD_MAP = {
            'name': 'Vuln_Name', 'category': 'Vuln_Class', 'port': 'Default_port',
            'level': 'Risk_Level', 'basis': 'Class_basis', 'description': 'Vuln_Description',
            'impact': 'Vuln_Hazards', 'suggestion': 'Repair_suggestions',
        }

        try:
            conn.execute("BEGIN")
            for row_idx, row in enumerate(rows[1:], start=2):
                entry = {}
                for col_idx, en_name in column_map.items():
                    val = row[col_idx] if col_idx < len(row) else ''
                    entry[en_name] = str(val).strip() if val else ''

                name = entry.get('name', '')
                if not name:
                    errors.append(f"Row {row_idx}: missing name, skipped")
                    skipped += 1
                    continue

                vuln_id = hashlib.md5(name.encode('utf-8')).hexdigest()

                cursor.execute("SELECT 1 FROM vulnerabilities WHERE Vuln_id = ?", (vuln_id,))
                exists = cursor.fetchone() is not None

                if exists:
                    if overwrite:
                        sets = []
                        vals = []
                        for en_key, db_col in FIELD_MAP.items():
                            sets.append(f"{db_col}=?")
                            vals.append(entry.get(en_key, ''))
                        vals.append(vuln_id)
                        cursor.execute(f"UPDATE vulnerabilities SET {', '.join(sets)} WHERE Vuln_id=?", vals)
                        imported += 1
                    else:
                        skipped += 1
                else:
                    col_str = ','.join(VULN_FIELDS)
                    placeholders = ','.join(['?'] * len(VULN_FIELDS))
                    vals = [vuln_id, name, entry.get('category', ''), entry.get('port', ''),
                            entry.get('level', ''), entry.get('basis', ''), entry.get('description', ''),
                            entry.get('impact', ''), entry.get('suggestion', '')]
                    cursor.execute(f"INSERT INTO vulnerabilities ({col_str}) VALUES ({placeholders})", vals)
                    imported += 1

            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

        reload_vulnerabilities_cache()
        return {"success": True, "imported": imported, "skipped": skipped, "errors": errors}
    except Exception as e:
        logger.error(f"Vulnerability import failed: {e}")
        return {"success": False, "error": str(e)}

@vuln_router.get("/api/vulnerabilities/template")
def download_vuln_template():
    """Download vulnerability import Excel template."""
    import openpyxl
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vuln Template"

    headers = ['漏洞名称', '漏洞分类', '默认端口', '风险级别', '定级依据', '漏洞描述', '漏洞危害', '修复建议']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    ws.cell(row=2, column=1, value='SQL注入漏洞')
    ws.cell(row=2, column=2, value='Web安全')
    ws.cell(row=2, column=3, value='80')
    ws.cell(row=2, column=4, value='高危')
    ws.cell(row=2, column=5, value='OWASP Top 10')
    ws.cell(row=2, column=6, value='可对数据库进行非法操作')
    ws.cell(row=2, column=7, value='导致数据泄露')
    ws.cell(row=2, column=8, value='使用参数化查询')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=vulnerability_template.xlsx"},
    )

@vuln_router.get("/api/vulnerabilities/export")
def export_vulnerabilities():
    """Export vulnerabilities to Excel."""
    import openpyxl
    from io import BytesIO

    _, vulns = get_cached_vulnerabilities()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vulnerabilities"

    headers = ['漏洞名称', '漏洞分类', '默认端口', '风险级别', '定级依据', '漏洞描述', '漏洞危害', '修复建议']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    db_keys = ['Vuln_Name', 'Vuln_Class', 'Default_port', 'Risk_Level',
               'Class_basis', 'Vuln_Description', 'Vuln_Hazards', 'Repair_suggestions']
    for row_idx, (_, v) in enumerate(vulns.items(), start=2):
        for col_idx, key in enumerate(db_keys, 1):
            ws.cell(row=row_idx, column=col_idx, value=v.get(key, '') or '')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=vulnerabilities.xlsx"},
    )

@icp_router.get("/api/icp-export")
def export_icp_entries():
    """Export ICP entries to Excel."""
    import openpyxl
    from io import BytesIO

    icp_list = get_db_reader().read_icp_raw_list()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ICP Data"

    headers = ['domain', 'unitName', 'natureName', 'mainLicence', 'serviceLicence', 'updateRecordTime']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    for row_idx, entry in enumerate(icp_list, start=2):
        for col_idx, key in enumerate(headers, 1):
            ws.cell(row=row_idx, column=col_idx, value=entry.get(key, '') or '')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=icp_data.xlsx"},
    )
def download_icp_template():
    """Download ICP import Excel template."""
    import openpyxl
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ICP Template"

    headers = ['domain', 'unitName', 'natureName', 'mainLicence', 'serviceLicence', 'updateRecordTime']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    ws.cell(row=2, column=1, value='example.com')
    ws.cell(row=2, column=2, value='示例公司')
    ws.cell(row=2, column=3, value='企业')
    ws.cell(row=2, column=4, value='京ICP备12345678号')
    ws.cell(row=2, column=5, value='京ICP证123456号')
    ws.cell(row=2, column=6, value='2026-01-01')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=icp_template.xlsx"},
    )

@config_router.post("/api/update-config")
def update_config(req: UpdateConfigRequest):
    """更新配置文件中的 supplierName"""
    try:
        # Load current YAML
        with open(CONF_PATH, 'r', encoding='utf-8') as f:
            current_conf = yaml.safe_load(f)
        
        # Update value
        current_conf['supplierName'] = req.supplierName
        
        # Write back to YAML
        with open(CONF_PATH, 'w', encoding='utf-8') as f:
            yaml.safe_dump(current_conf, f, allow_unicode=True, sort_keys=False)
            
        # Update memory config
        config['supplierName'] = req.supplierName
        
        # 更新模板管理器的配置
        get_template_manager().update_config(config)
        
        return {"success": True, "message": "配置已更新"}
    except Exception as e:
        _raise_internal_error("update config", e)


# ========== 模板相关 API ==========

@app.get("/api/services")
def get_service_map():
    """Expose @service/ → URL mapping for frontend behavior endpoint resolution.

    Templates use '@service/name' prefixes in schema.yaml behavior endpoints
    instead of hardcoding framework API URLs. This endpoint provides the mapping
    so the frontend form-renderer can resolve them at runtime.
    """
    return {
        "success": True,
        "services": SERVICE_MAP,
    }

@app.get("/api/templates")
def get_templates(include_details: bool = False):
    """
    获取所有可用模板列表
    
    Args:
        include_details: 是否包含详细信息（文件大小、字段数等）
    """
    if include_details:
        # 返回详细信息
        templates = [
            get_template_manager().get_template_details(t["id"]) 
            for t in get_template_manager().get_template_list()
        ]
        # 过滤掉 None 值（模板不存在的情况）
        templates = [t for t in templates if t is not None]
    else:
        # 返回简要信息
        templates = get_template_manager().get_template_list()
    
    return {
        "templates": templates,
        "default_template": get_template_manager().default_template_id
    }

@app.get("/api/templates/{template_id}/schema")
def get_template_schema(template_id: str):
    """获取指定模板的完整 schema"""
    schema = get_template_manager().get_template_schema(template_id)
    if not schema:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    return schema

@app.get("/api/templates/{template_id}/versions")
def get_template_versions(template_id: str):
    """获取模板的所有版本"""
    versions = get_template_manager().get_template_versions(template_id)
    return {"template_id": template_id, "versions": versions}

@app.get("/api/templates/{template_id}/data-sources")
def get_template_data_sources(template_id: str):
    """获取模板所需的数据源数据"""
    # 准备数据库数据
    vuln_list, _ = get_cached_vulnerabilities()
    db_data = {
        "vulnerabilities": vuln_list,
        "icp_cache": list(get_cached_icp_infos().values())
    }
    
    resolved = get_template_manager().resolve_data_sources(template_id, db_data)
    if not resolved:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    
    return resolved

@app.get("/api/templates/{template_id}/widgets/{filename}")
def get_template_widget(template_id: str, filename: str):
    """Serve template widget files (JS / CSS)."""
    tm = get_template_manager()
    widget_path = os.path.join(tm.get_template_dir(template_id), "widgets", filename)
    if ".." in filename or not os.path.exists(widget_path):
        raise HTTPException(status_code=404, detail="Widget not found")
    ext = os.path.splitext(filename)[1].lower()
    media_map = {".js": "application/javascript", ".css": "text/css"}
    return FileResponse(path=widget_path, media_type=media_map.get(ext, "application/octet-stream"))

@app.get("/api/widgets/shared/{filename}")
def get_shared_widget(filename: str):
    """Serve shared widget files (JS / CSS) from the common widgets directory."""
    shared_dir = os.path.join(BASE_DIR, "widgets")
    widget_path = os.path.join(shared_dir, filename)
    if ".." in filename or not os.path.exists(widget_path):
        raise HTTPException(status_code=404, detail="Shared widget not found")
    ext = os.path.splitext(filename)[1].lower()
    media_map = {".js": "application/javascript", ".css": "text/css"}
    return FileResponse(path=widget_path, media_type=media_map.get(ext, "application/octet-stream"))

@app.post("/api/templates/{template_id}/validate")
def validate_template_data(template_id: str, data: dict[str, Any]):
    """验证表单数据"""
    is_valid, errors = get_template_manager().validate_report_data(template_id, data)
    return {
        "valid": is_valid,
        "errors": errors
    }


def _build_runtime_execution_config() -> dict[str, Any]:
    """Merge host config with shared plugin runtime switches for execution."""
    execution_config = dict(config)
    execution_config["plugin_runtime"] = shared_config.get("plugin_runtime", {})
    return execution_config


def _get_plugin_runtime_class():
    """Lazily import runtime to avoid package resolution edge cases."""
    module = importlib.import_module("backend.plugin_host.runtime")
    return module.PluginRuntime


def _reload_templates_with_alignment_check() -> tuple[TemplateManager, list[dict[str, Any]], set[str], set[str], list[str], bool, Optional[str]]:
    """Reload templates and enforce handler alignment.

    Returns:
        (tm, templates, handlers_before, handlers_after, new_handlers, requires_restart, warning_message)
    """
    tm = get_template_manager()
    handlers_before = set(HandlerRegistry.list_registered())

    tm.reload_templates()
    _assert_template_handler_alignment(tm)

    templates = tm.get_template_list()
    handlers_after = set(HandlerRegistry.list_registered())
    new_handlers = sorted(list(handlers_after - handlers_before))
    routers = tm.get_template_routers()
    new_routes = [h for h in new_handlers if h in routers]

    if new_routes:
        warning_message = f"检测到新增带路由的模板 ({', '.join(new_routes)})，需要重启应用才能生效"
        requires_restart = True
    else:
        warning_message = None
        requires_restart = False

    return tm, templates, handlers_before, handlers_after, new_handlers, requires_restart, warning_message

@app.post("/api/templates/{template_id}/generate")
def generate_template_report(template_id: str, data: dict[str, Any]):
    """
    使用指定模板生成报告 (新接口，使用 Handler)
    """
    try:
        result = _get_plugin_runtime_class().execute(
            template_id=template_id,
            data=data,
            output_dir=OUTPUT_DIR,
            template_manager=get_template_manager(),
            config=_build_runtime_execution_config(),
        )
        
        if result["success"]:
            report_path = os.path.realpath(result["report_path"])
            if not _is_subpath(report_path, OUTPUT_DIR):
                raise HTTPException(status_code=400, detail="Report path escaped output directory")
            relative_report_path = os.path.relpath(report_path, OUTPUT_DIR).replace(os.sep, "/")
            download_url = f"/reports/{quote(relative_report_path, safe='/')}"
            
            return {
                "success": True,
                "report_path": report_path,
                "download_url": download_url,
                "message": result["message"]
            }
        return _build_error_payload(
            result.get("message", "报告生成失败"),
            None,
            report_path="",
            download_url="",
            errors=result.get("errors", []),
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Generate report error: {traceback.format_exc()}")
        return _build_error_payload(
            "报告生成异常",
            str(e),
            report_path="",
            download_url="",
            errors=[],
        )

@app.get("/api/templates/{template_id}/preview")
def get_template_preview_config(template_id: str):
    """获取模板预览配置"""
    schema = get_template_manager().get_template_schema(template_id)
    if not schema:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    
    return {
        "template_id": template_id,
        "name": schema.get("name", ""),
        "preview": schema.get("preview", {})
    }

@app.post("/api/templates/reload")
def reload_templates():
    """
    重新加载所有模板（支持热加载）
    
    注意：
    - 模板的 Handler 和 Schema 可以热加载
    - 新增的 API 路由需要重启应用才能生效（FastAPI 限制）
    """
    try:
        _, templates, _, handlers_after, new_handlers, requires_restart, warning_message = _reload_templates_with_alignment_check()

        return {
            "success": True,
            "message": "模板已重新加载",
            "warning": warning_message,
            "loaded_count": len(templates),
            "templates": templates,
            "handlers": list(handlers_after),
            "new_handlers": new_handlers,
            "requires_restart": requires_restart
        }
    except Exception as e:
        _raise_internal_error("reload templates", e)


@app.get("/api/templates/{template_id}/details")
def get_template_details(template_id: str):
    """获取模板详细信息"""
    details = get_template_manager().get_template_details(template_id)
    if not details:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    return details


@app.put("/api/templates/{template_id}/set-default")
def set_default_template(template_id: str):
    """将指定模板设置为默认模板"""
    try:
        result = get_template_manager().set_default_template(template_id)
        return result
    except TemplateNotFoundError:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    except Exception as e:
        _raise_internal_error("set default template", e)


@app.get("/api/templates/{template_id}/check-deps")
def check_template_dependencies(template_id: str):
    """
    检查模板依赖（解决问题 9：模板依赖管理缺失）
    
    Args:
        template_id: 模板ID
        
    Returns:
        {
            "template_id": str,
            "dependencies_satisfied": bool,
            "missing_dependencies": List[str],
            "message": str
        }
    """
    tm = get_template_manager()
    satisfied, missing = tm.check_dependencies(template_id)
    
    return {
        "template_id": template_id,
        "dependencies_satisfied": satisfied,
        "missing_dependencies": missing,
        "message": "所有依赖已满足" if satisfied else f"缺少依赖: {', '.join(missing)}"
    }


@app.get("/api/templates/{template_id}/export")
def export_template(template_id: str):
    """导出模板为压缩包"""
    tm = get_template_manager()
    template_dir = tm.get_template_dir(template_id)
    if not os.path.exists(template_dir):
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    
    # 创建内存中的 ZIP 文件
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(template_dir):
            # 排除 __pycache__ 等目录
            dirs[:] = [d for d in dirs if not d.startswith('__')]
            
            for file in files:
                if file.endswith('.pyc'):
                    continue
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, template_dir)
                zf.write(file_path, f"{template_id}/{arcname}")
    
    zip_buffer.seek(0)
    
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={template_id}_template.zip"}
    )


@app.post("/api/templates/batch-export")
def batch_export_templates(template_ids: list[str]):
    """批量导出多个模板为一个压缩包"""
    if not template_ids:
        raise HTTPException(status_code=400, detail="No template IDs provided")
    
    tm = get_template_manager()
    # 创建内存中的 ZIP 文件
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for template_id in template_ids:
            template_dir = tm.get_template_dir(template_id)
            if not os.path.exists(template_dir):
                logger.warning(f"Template not found: {template_id}")
                continue
            
            # 将每个模板添加到 ZIP 中
            for root, dirs, files in os.walk(template_dir):
                # 排除 __pycache__ 等目录
                dirs[:] = [d for d in dirs if not d.startswith('__')]
                
                for file in files:
                    if file.endswith('.pyc'):
                        continue
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, template_dir)
                    zf.write(file_path, f"{template_id}/{arcname}".replace("\\", "/"))
    
    zip_buffer.seek(0)
    
    # 生成文件名（带时间戳）
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"templates_batch_{timestamp}.zip"
    
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def _detect_templates_in_zip(names: list[str]) -> list[str]:
    """
    检测 ZIP 中的模板结构
    返回模板 ID 列表
    """
    template_ids = set()
    
    for name in names:
        parts = name.replace('\\', '/').split('/')
        if len(parts) >= 2 and parts[1] == 'schema.yaml':
            # 找到 template_id/schema.yaml
            template_ids.add(parts[0])
    
    return sorted(list(template_ids))


def _import_single_template(zf, template_id: str, all_names: list[str], overwrite: bool) -> dict[str, Any]:
    """
    从 ZIP 中导入单个模板 (Security Hardened)
    
    导入目标目录：
    - 生产模式（AppData 启用）：导入到 USER_TEMPLATES_DIR，跨重装持久保留
    - 开发模式 / AppData 未启用：导入到 TEMPLATES_DIR
    """
    # 确定导入目标目录
    import_base_dir = USER_TEMPLATES_DIR if _USERDATA_ENABLED and USER_TEMPLATES_DIR else TEMPLATES_DIR

    # 过滤出属于该模板的文件
    template_files = [n for n in all_names if n.startswith(f"{template_id}/")]

    if not template_files:
        return {"success": False, "reason": "No files found"}

    # 检查必需文件
    schema_found = any(n == f"{template_id}/schema.yaml" for n in template_files)
    if not schema_found:
        return {"success": False, "reason": "schema.yaml not found"}

    # 检查是否已存在（同时检查内置目录和用户目录）
    target_dir = os.path.join(import_base_dir, template_id)
    builtin_dir = os.path.join(TEMPLATES_DIR, template_id)
    for check_dir in [target_dir, builtin_dir]:
        if check_dir == target_dir:
            continue
        if os.path.exists(check_dir):
            return {"success": False, "reason": f"Template '{template_id}' already exists as built-in template"}

    backup_dir: Optional[str] = None
    if os.path.exists(target_dir):
        if not overwrite:
            return {"success": False, "reason": "Already exists (overwrite=false)"}
        # 备份现有模板
        backup_dir = f"{target_dir}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        shutil.move(target_dir, backup_dir)

    # [Security Fix] Zip Slip 防御 & 安全解压
    extracted_files = []
    try:
        for file_name in template_files:
            # 1. 检查文件名是否包含路径遍历字符
            normalized = file_name.replace('\\', '/')
            if '..' in normalized or normalized.startswith('/'):
                raise ValueError(f"Malicious path detected: {file_name}")

            # 2. 规范化目标路径
            target_path = os.path.join(import_base_dir, file_name)
            if not os.path.abspath(target_path).startswith(os.path.abspath(import_base_dir)):
                raise ValueError(f"Path traversal attempt: {file_name}")

            # 3. 解压
            zf.extract(file_name, import_base_dir)
            extracted_files.append(target_path)

            # 4. [Security Fix] 如果是 .py 文件，立即进行静态审计
            if file_name.endswith('.py'):
                try:
                    tm = get_template_manager()
                    tm.audit_code_security(template_id, target_path)
                except Exception as audit_err:
                    logger.error(f"Audit failed during import for {file_name}: {audit_err}")
                    raise ValueError(f"Security Policy Violation: {str(audit_err)}")

    except Exception as e:
        # 发生错误，清理残局
        if os.path.exists(target_dir):
            shutil.rmtree(target_dir)
        if backup_dir and os.path.exists(backup_dir):
            shutil.move(backup_dir, target_dir)
        return {"success": False, "reason": str(e)}

    return {
        "success": True,
        "template_id": template_id,
        "target_dir": target_dir,
        "backup_dir": backup_dir,
    }


@app.post("/api/templates/import")
async def import_template(file: UploadFile = File(...), overwrite: bool = Form(default=False)):
    """
    导入模板压缩包
    支持三种格式：
    1. 单个模板 ZIP（template_id/schema.yaml）
    2. 多模板综合 ZIP（template1/..., template2/...）
    3. 通过多次调用导入多个文件
    """
    if not file.filename or not file.filename.endswith('.zip'):
        raise HTTPException(status_code=400, detail="Only ZIP files are supported")
    
    try:
        content = await file.read()
        zip_buffer = io.BytesIO(content)
        
        with zipfile.ZipFile(zip_buffer, 'r') as zf:
            names = zf.namelist()
            if not names:
                raise HTTPException(status_code=400, detail="Empty ZIP file")
            
            # 检测 ZIP 结构：单模板 or 多模板
            template_ids = _detect_templates_in_zip(names)
            
            if not template_ids:
                raise HTTPException(status_code=400, detail="No valid templates found in ZIP")
            
            imported = []
            replaced = []
            skipped = []
            errors = []
            import_contexts: list[dict[str, Any]] = []
            
            for template_id in template_ids:
                try:
                    result = _import_single_template(zf, template_id, names, overwrite)
                    if result['success']:
                        imported.append(template_id)
                        import_contexts.append(result)
                        if result.get('backup_dir'):
                            replaced.append(template_id)
                    else:
                        skipped.append({'id': template_id, 'reason': result['reason']})
                except Exception as e:
                    errors.append({'id': template_id, 'error': str(e)})
            
            # 重新加载模板
            if imported:
                try:
                    _reload_templates_with_alignment_check()
                except Exception as reload_err:
                    # reload 失败时回滚导入变更，恢复旧模板目录
                    for ctx in import_contexts:
                        target_dir = ctx.get("target_dir")
                        backup_dir = ctx.get("backup_dir")
                        if isinstance(target_dir, str) and os.path.exists(target_dir):
                            shutil.rmtree(target_dir)
                        if isinstance(backup_dir, str) and backup_dir and isinstance(target_dir, str) and target_dir and os.path.exists(backup_dir):
                            shutil.move(backup_dir, target_dir)

                    # 再次尝试恢复模板注册状态
                    try:
                        _reload_templates_with_alignment_check()
                    except Exception:
                        logger.error("Failed to restore template registry after import rollback")

                    raise HTTPException(status_code=500, detail=f"Import rollback applied: {str(reload_err)}")

                # reload 成功后清理备份目录
                for ctx in import_contexts:
                    backup_dir = ctx.get("backup_dir")
                    if isinstance(backup_dir, str) and backup_dir and os.path.exists(backup_dir):
                        shutil.rmtree(backup_dir)
            
            return {
                "success": len(imported) > 0,
                "imported": imported,
                "replaced": replaced,
                "skipped": skipped,
                "errors": errors,
                "message": f"Successfully imported {len(imported)} template(s)"
            }
            
    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="Invalid ZIP file")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Import failed: {str(e)}")


@app.post("/api/templates/batch-import")
async def batch_import_templates(files: list[UploadFile] = File(...), overwrite: bool = Form(default=False)):
    """
    批量导入多个模板文件
    支持选择多个 ZIP 文件同时上传
    """
    if not files:
        raise HTTPException(status_code=400, detail="No files provided")
    
    results = []
    
    for file in files:
        try:
            # 调用单文件导入逻辑
            result = await import_template(file, overwrite)
            results.append({
                "filename": file.filename,
                "success": result["success"],
                "imported": result.get("imported", []),
                "message": result.get("message", "")
            })
        except HTTPException as e:
            results.append({
                "filename": file.filename,
                "success": False,
                "message": str(e.detail),
                "detail": str(e.detail),
            })
        except Exception as e:
            results.append({
                "filename": file.filename,
                "success": False,
                "message": str(e),
                "detail": str(e),
            })
    
    total_imported = sum(len(r.get("imported", [])) for r in results if r["success"])
    
    return {
        "success": total_imported > 0,
        "total_files": len(files),
        "total_imported": total_imported,
        "results": results
    }


@app.delete("/api/templates/{template_id}")
def delete_template(template_id: str, backup: bool = True):
    """删除模板（支持内置模板备份删除和用户模板彻底删除）"""
    tm = get_template_manager()

    # 保护检查：扫描内置目录和用户目录的 runtime.yaml
    protected_templates = []
    dirs_to_check = [TEMPLATES_DIR]
    if USER_TEMPLATES_DIR and os.path.isdir(USER_TEMPLATES_DIR):
        dirs_to_check.append(USER_TEMPLATES_DIR)
    for tid in tm.template_ids:
        for base_dir in dirs_to_check:
            runtime_path = os.path.join(base_dir, tid, "runtime.yaml")
            if os.path.exists(runtime_path):
                with open(runtime_path, 'r', encoding='utf-8') as f:
                    runtime_config = yaml.safe_load(f) or {}
                    if runtime_config.get('protected'):
                        protected_templates.append(tid)
                break  # found runtime.yaml, no need to check other dirs
    if template_id in protected_templates:
        raise HTTPException(status_code=403, detail=f"Cannot delete protected template: {template_id}")

    # 优先使用 TemplateManager.delete_template（支持用户模板）
    if template_id in tm.user_template_ids:
        success, msg = tm.delete_template(template_id)
        if success:
            return {"success": True, "message": msg}
        raise HTTPException(status_code=500, detail=msg)

    # 内置模板：沿用旧逻辑（备份到 _deleted 目录）
    template_dir = os.path.join(TEMPLATES_DIR, template_id)
    if not os.path.exists(template_dir):
        raise HTTPException(status_code=404, detail=f"Template not found: {template_id}")
    
    try:
        if backup:
            # 确保 deleted 目录存在
            os.makedirs(TEMPLATES_DELETED_DIR, exist_ok=True)
            # 移动到 deleted 目录
            deleted_path = os.path.join(TEMPLATES_DELETED_DIR, template_id)
            # 如果已存在同名已删除模板，添加时间戳
            if os.path.exists(deleted_path):
                deleted_path = os.path.join(TEMPLATES_DELETED_DIR, f"{template_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            shutil.move(template_dir, deleted_path)
        else:
            shutil.rmtree(template_dir)
        
        # 重新加载模板
        tm.reload_templates()
        
        return {
            "success": True,
            "message": f"Template '{template_id}' deleted successfully",
            "backed_up": backup
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Delete failed: {str(e)}")


# 将按域拆分的路由挂载回主应用（保持原路径不变）
app.include_router(config_router)
app.include_router(vuln_router)
app.include_router(icp_router)


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    # 从共享配置读取服务器设置
    server_host = shared_config.get("server", {}).get("host", "127.0.0.1")
    server_port = shared_config.get("server", {}).get("port", 8000)
    # workers=1 确保单进程运行，避免 Windows 上产生多个进程
    # reload=False 禁用热重载（打包后不需要）
    uvicorn.run(app, host=server_host, port=server_port, workers=1, reload=False)

