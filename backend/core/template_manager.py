# -*- coding: utf-8 -*-
"""
@Createtime: 2026-01-24
@description: 模板管理器 - 负责加载、验证和管理报告模板
支持版本管理、数据源解析、动态表单生成
"""

import os
import json
import yaml
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime
from functools import lru_cache
import re
from .logger import setup_logger
from .schema_models import (
    ALLOWED_FIELD_TYPES,
    ALLOWED_DATA_SOURCE_TYPES,
    ALLOWED_ACTION_TYPES,
    Behavior,
    BehaviorAction,
    DataSourceDef,
    FieldDefinition,
    FieldGroup,
    PreviewField,
    TemplateInfo,
    ValidationRule,
    validate_template_id,
)

from .exceptions import (
    TemplateNotFoundError,
    TemplateLoadError,
    InvalidTemplateIdError,
    SchemaParseError,
    DependencyError,
    PathTraversalError
)

# 初始化日志记录器
logger = setup_logger('TemplateManager')


def validate_path_safety(path: str, base_dir: str) -> bool:
    """
    验证路径安全性，防止路径遍历攻击（解决问题 12：安全风险）
    
    Args:
        path: 要验证的路径
        base_dir: 基础目录
        
    Returns:
        bool: 路径是否安全
        
    Examples:
        >>> validate_path_safety("templates/my_template", "templates")
        True
        >>> validate_path_safety("../../../etc/passwd", "templates")
        False
    """
    try:
        # 规范化路径
        abs_path = os.path.abspath(os.path.join(base_dir, path))
        abs_base = os.path.abspath(base_dir)
        
        # 检查路径是否在基础目录内
        return abs_path.startswith(abs_base)
    except Exception:
        return False




class TemplateManager:
    """模板管理器"""
    
    # 排除的目录：以 _ 或 . 开头的目录，以及特定系统目录
    EXCLUDED_DIRS = {'_deleted', '_backup', '__pycache__', '.git', '.vscode', '.idea'}
    
    def __init__(self, templates_dir: str, config: Optional[Dict[str, Any]] = None,
                 user_templates_dir: Optional[str] = None):
        """
        初始化模板管理器
        
        Args:
            templates_dir: 内置模板根目录路径（安装目录，只读）
            config: 全局配置 (用于解析数据源)
            user_templates_dir: 用户自定义模板目录（AppData，可读写），可选
        """
        self.templates_dir = templates_dir
        self.user_templates_dir = user_templates_dir
        self.config = config if config is not None else {}
        self._templates: Dict[str, TemplateInfo] = {}
        self._template_versions: Dict[str, List[str]] = {}  # {template_id: [versions]}
        self._template_routers: Dict[str, Any] = {}  # 新增：存储模板路由
        self._template_source_dirs: Dict[str, str] = {}  # {template_id: source_directory}
        self._load_all_templates()
    
    def _load_all_templates(self):
        """扫描并加载所有模板（内置 + 用户自定义）"""
        self._scan_templates_dir(self.templates_dir, is_user=False)

        if self.user_templates_dir and os.path.isdir(self.user_templates_dir):
            logger.info(f"Scanning user templates directory: {self.user_templates_dir}")
            self._scan_templates_dir(self.user_templates_dir, is_user=True)

        self._load_order_overrides()
    
    def _scan_templates_dir(self, templates_base: str, is_user: bool = False):
        """扫描指定目录下的模板"""
        if not os.path.exists(templates_base):
            if not is_user:
                logger.warning(f"Templates directory not found: {templates_base}")
            return
        
        for item in os.listdir(templates_base):
            # 跳过排除目录和隐藏目录（以 _ 或 . 开头）
            if item in self.EXCLUDED_DIRS or item.startswith(('_', '.')):
                logger.debug(f"Skipping excluded directory: {item}")
                continue
            
            # 安全检查：防止路径遍历（解决问题 12）
            if not validate_path_safety(item, templates_base):
                logger.warning(f"Skipping unsafe path: {item}")
                continue
            
            template_path = os.path.join(templates_base, item)
            if not os.path.isdir(template_path):
                continue
            
            # 与已加载的内置模板冲突时：用户模板跳过并警告
            if is_user and item in self._templates:
                logger.warning(
                    f"User template '{item}' conflicts with built-in template — "
                    f"user template skipped. Rename the user template folder to use it."
                )
                continue
            
            schema_path = os.path.join(template_path, "schema.yaml")
            if os.path.exists(schema_path):
                self._template_source_dirs[item] = template_path
                self._load_template(item, schema_path)
    
    def _load_template(self, template_id: str, schema_path: str):
        """
        加载单个模板的 schema
        
        Args:
            template_id: 模板ID (目录名)
            schema_path: schema.yaml 文件路径
        """
        # 验证模板 ID 是否符合 Python 模块命名规范
        if not validate_template_id(template_id):
            logger.error(f"Invalid template ID format: {template_id}")
            return
        
        template_dir = os.path.dirname(schema_path)
        
        try:
            from .schema_loader import SchemaLoader
            
            # Use SchemaLoader to parse schema.yaml into TemplateInfo
            template_info = SchemaLoader.load_schema(template_dir)
            
            self._templates[template_info.id] = template_info
            
            # 版本管理
            if template_info.id not in self._template_versions:
                self._template_versions[template_info.id] = []
            if template_info.version not in self._template_versions[template_info.id]:
                self._template_versions[template_info.id].append(template_info.version)
            
            logger.info(f"Loaded template: {template_info.id} v{template_info.version} ({template_info.name})")
            
            # 动态加载 handler（阶段 1：任务 1.1）
            self._load_handler(template_info.id, template_dir)
            
        except FileNotFoundError as e:
            error = TemplateLoadError(template_id, str(e))
            logger.error(str(error))
        except KeyError as e:
            error = TemplateLoadError(template_id, f"Missing required field in schema: {str(e)}")
            logger.error(str(error))
        except ValueError as e:
            error = TemplateLoadError(template_id, f"Invalid value in schema: {str(e)}")
            logger.error(str(error))
        except Exception as e:
            error = TemplateLoadError(template_id, f"Unexpected error: {str(e)}")
            logger.error(str(error))
            import traceback
            traceback.print_exc()
    
    def audit_code_security(self, template_id: str, file_path: str):
        """
        [Security Fix] 静态审计代码安全性
        检查 handler.py 是否包含禁止的导入或危险函数调用
        """
        import ast
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                tree = ast.parse(f.read())
            
            for node in ast.walk(tree):
                # 检查 import
                if isinstance(node, (ast.Import, ast.ImportFrom)):
                    # 解析模块名
                    module_name = ""
                    if isinstance(node, ast.Import):
                        for alias in node.names:
                            module_name = alias.name.split('.')[0]
                            self._check_module_name(template_id, module_name)
                    elif isinstance(node, ast.ImportFrom):
                        if node.module:
                            module_name = node.module.split('.')[0]
                            self._check_module_name(template_id, module_name)

                # 检查禁止的内置函数调用 (eval, exec, etc.)
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
                    if node.func.id in {'eval', 'exec', 'compile', 'globals', 'locals'}:
                         raise ValueError(f"Security Alert: Usage of forbidden function '{node.func.id}' in handler.py")

        except Exception as e:
            # 任何解析错误或安全违规都视为审计失败
            logger.error(f"Code audit failed for {template_id}: {e}")
            raise ValueError(f"Rejected malicious code: {str(e)}")

    def _check_module_name(self, template_id: str, module_name: str):
        """检查模块名是否在白名单中"""
        # 放行 core.* 内部模块
        if module_name == 'core': 
            return
            
        if module_name not in self.ALLOWED_PACKAGES:
             # 特殊处理子模块
             if '.' in module_name:
                 base_mod = module_name.split('.')[0]
                 if base_mod in self.ALLOWED_PACKAGES:
                     return
             
             raise ValueError(f"Security Alert: Import of unapproved module '{module_name}' is forbidden")

    def _load_handler(self, template_id: str, template_dir: str):
        """
        动态加载模板的 handler.py（阶段 1：任务 1.1）
        
        功能：
        1. 使用 importlib 动态导入 handler.py
        2. HandlerRegistry.register() 注册到 HandlerRegistry
        
        Args:
            template_id: 模板ID
            template_dir: 模板目录路径
        """
        handler_path = os.path.join(template_dir, "handler.py")
        if not os.path.exists(handler_path):
            logger.warning(f"Handler not found for template: {template_id}")
            return
        
        # [Security Fix] 加载前先执行静态代码审计
        try:
            self.audit_code_security(template_id, handler_path)
        except ValueError as e:
            logger.critical(f"Security audit failed for {template_id}, loading blocked: {e}")
            return # 阻止加载

        try:
            import importlib.util
            import sys
            
            # 动态加载模块
            module_name = f"templates.{template_id}.handler"
            spec = importlib.util.spec_from_file_location(module_name, handler_path)
            if spec is None or spec.loader is None:
                logger.error(f"Failed to create module spec for {template_id}")
                return
            
            module = importlib.util.module_from_spec(spec)
            
            # 添加到 sys.modules，避免重复加载
            sys.modules[module_name] = module
            
            # 执行模块（触发 HandlerRegistry.register()）
            spec.loader.exec_module(module)
            
            # 收集模板路由（如果有）
            if hasattr(module, 'router'):
                self._template_routers[template_id] = module.router
                logger.info(f"Collected router for template: {template_id}")
            
            logger.info(f"Dynamically loaded handler for: {template_id}")
            
        except Exception as e:
            logger.error(f"Failed to load handler for {template_id}: {e}")
    
    def get_template_routers(self) -> Dict[str, Any]:
        """获取所有模板的路由"""
        return self._template_routers

    def get_template(self, template_id: str, raise_if_not_found: bool = False) -> Optional[TemplateInfo]:
        """
        获取指定模板信息
        
        Args:
            template_id: 模板ID
            raise_if_not_found: 如果为 True，模板不存在时抛出异常
            
        Returns:
            模板信息，如果不存在且 raise_if_not_found=False 则返回 None
            
        Raises:
            TemplateNotFoundError: 当 raise_if_not_found=True 且模板不存在时
        """
        template = self._templates.get(template_id)
        if template is None and raise_if_not_found:
            raise TemplateNotFoundError(template_id)
        return template
    
    def get_template_list(self) -> List[Dict[str, Any]]:
        """获取所有模板的简要列表（按 order 排序）"""
        # 先按 order 排序模板
        sorted_templates = sorted(self._templates.values(), key=lambda t: t.order)
        
        result = []
        for t in sorted_templates:
            template_dir = self.get_template_dir(t.id)
            result.append({
                "id": t.id,
                "name": t.name,
                "description": t.description,
                "icon": t.icon,
                "version": t.version,
                "author": t.author,
                "update_time": t.update_time,
                "order": t.order,
                "has_widgets": os.path.isdir(os.path.join(template_dir, "widgets"))
            })
        return result
    
    def set_default_template(self, template_id: str) -> Dict[str, Any]:
        """
        将指定模板设置为默认模板（order=0），其他模板 order 递增 1。
        
        Args:
            template_id: 模板ID
            
        Returns:
            dict with success, message, and template_id
            
        Raises:
            TemplateNotFoundError: 当模板不存在时
        """
        template = self.get_template(template_id, raise_if_not_found=True)
        
        if template.order == 0:
            return {
                "success": True,
                "message": f"Template '{template_id}' is already the default",
                "template_id": template_id
            }
        
        # 将所有模板 order 递增 1，再将目标模板设为 0
        for t in self._templates.values():
            t.order += 1
        
        template.order = 0
        
        logger.info(f"Set default template: {template_id}")
        self._persist_order()
        return {
            "success": True,
            "message": f"Template '{template_id}' set as default",
            "template_id": template_id
        }

    def _load_order_overrides(self):
        """Load persisted order values from template_order.json."""
        try:
            order_file = os.path.join(os.path.dirname(self.templates_dir), "data", "template_order.json")
            if not os.path.exists(order_file):
                return
            with open(order_file, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            if not isinstance(saved, dict):
                return
            for tid, order in saved.items():
                if tid in self._templates:
                    try:
                        self._templates[tid].order = int(order)
                    except (ValueError, TypeError):
                        continue
        except Exception:
            pass  # safe fallback

    def _persist_order(self):
        """Persist current template order to template_order.json atomically."""
        try:
            order_file = os.path.join(os.path.dirname(self.templates_dir), "data", "template_order.json")
            os.makedirs(os.path.dirname(order_file), exist_ok=True)
            order_map = {tid: getattr(t, 'order', 999) for tid, t in self._templates.items()}
            tmp_path = order_file + ".tmp"
            with open(tmp_path, 'w', encoding='utf-8') as f:
                json.dump(order_map, f, ensure_ascii=False, indent=2)
                f.write("\n")
            os.replace(tmp_path, order_file)
        except Exception as exc:
            logger.warning(f"Failed to persist template order: {exc}")

    def get_template_versions(self, template_id: str) -> List[str]:
        """获取模板的所有版本"""
        return self._template_versions.get(template_id, [])
    
    def compare_versions(self, version1: str, version2: str) -> int:
        """
        比较两个版本号（解决问题 11：模板版本管理）
        
        Args:
            version1: 版本号1 (如 "1.2.3")
            version2: 版本号2 (如 "1.2.4")
            
        Returns:
            -1: version1 < version2
             0: version1 == version2
             1: version1 > version2
        """
        try:
            v1_parts = [int(x) for x in version1.split('.')]
            v2_parts = [int(x) for x in version2.split('.')]
            
            # 补齐长度
            max_len = max(len(v1_parts), len(v2_parts))
            v1_parts.extend([0] * (max_len - len(v1_parts)))
            v2_parts.extend([0] * (max_len - len(v2_parts)))
            
            for v1, v2 in zip(v1_parts, v2_parts):
                if v1 < v2:
                    return -1
                elif v1 > v2:
                    return 1
            return 0
        except (ValueError, AttributeError):
            # 如果版本号格式不正确，按字符串比较
            if version1 < version2:
                return -1
            elif version1 > version2:
                return 1
            return 0
    
    def check_version_conflict(self, template_id: str, new_version: str) -> Tuple[bool, str]:
        """
        检查版本冲突（解决问题 11：模板版本管理）
        
        Args:
            template_id: 模板ID
            new_version: 新版本号
            
        Returns:
            (是否有冲突, 冲突信息)
        """
        existing_versions = self.get_template_versions(template_id)
        
        if not existing_versions:
            return False, ""
        
        # 检查是否已存在相同版本
        if new_version in existing_versions:
            return True, f"Version {new_version} already exists"
        
        # 检查新版本是否比现有版本旧
        current_template = self._templates.get(template_id)
        if current_template:
            current_version = current_template.version
            if self.compare_versions(new_version, current_version) < 0:
                return True, f"New version {new_version} is older than current version {current_version}"
        
        return False, ""
    
    def _serialize_field(self, field: FieldDefinition) -> Dict[str, Any]:
        """
        将字段定义序列化为字典
        
        Args:
            field: 字段定义对象
            
        Returns:
            字段字典（包含所有字段和额外字段）
        """
        return field.model_dump()
    
    def _serialize_field_group(self, group: FieldGroup) -> Dict[str, Any]:
        """
        将字段分组序列化为字典
        
        Args:
            group: 字段分组对象
            
        Returns:
            字段分组字典
        """
        return group.model_dump()
    
    def _serialize_data_source(self, ds: DataSourceDef) -> Dict[str, Any]:
        """
        将数据源定义序列化为字典
        
        Args:
            ds: 数据源定义对象
            
        Returns:
            数据源字典
        """
        return ds.model_dump()
    
    def _serialize_behavior(self, behavior: Behavior) -> Dict[str, Any]:
        """
        将行为定义序列化为字典
        
        Args:
            behavior: 行为定义对象
            
        Returns:
             行为字典（含 trigger.fields 数组等额外字段）
        """
        # Use model_dump() to preserve extra fields like trigger.fields
        # that are not mapped to explicit Behavior model attributes
        data = behavior.model_dump()
        # Backward compat: actions use model_dump() which now includes extra fields
        data["actions"] = [a.model_dump() for a in behavior.actions]
        return data
    
    @lru_cache(maxsize=128)
    def _get_cached_schema(self, template_id: str, version: str) -> Optional[Dict[str, Any]]:
        """
        获取缓存的模板 schema（解决问题 13：性能优化）
        
        使用 LRU 缓存避免重复解析 schema
        """
        return self.get_template_schema(template_id)
    
    def get_template_schema(self, template_id: str) -> Optional[Dict[str, Any]]:
        """
        获取模板的完整 schema (用于前端渲染表单)
        """
        template = self._templates.get(template_id)
        if not template:
            return None
        
        result = {
            "id": template.id,
            "name": template.name,
            "description": template.description,
            "version": template.version,
            "icon": template.icon,
            "author": template.author,
            "field_groups": [self._serialize_field_group(g) for g in template.field_groups],
            "fields": [self._serialize_field(f) for f in template.fields],
            "data_sources": [self._serialize_data_source(ds) for ds in template.data_sources],
            "behaviors": [self._serialize_behavior(b) for b in template.behaviors],
            "validation": {
                "rules": [
                    {
                        "fields": r.fields,
                        "rule": r.rule,
                        "message": r.message
                    }
                    for r in template.validation_rules
                ]
            },
            "output": template.output_config,
            "preview": template.preview_config,
            "dependent_fields": template.dependent_fields,
            "summary_configs": template.summary_configs
        }

        # Inject raw schema fields not captured by TemplateInfo (e.g. vuln_save)
        template_path = self.get_template_dir(template_id)
        result["has_widgets"] = os.path.isdir(os.path.join(template_path, "widgets"))
        schema_path = os.path.join(template_path, "schema.yaml")
        if os.path.exists(schema_path):
            try:
                with open(schema_path, 'r', encoding='utf-8') as f:
                    raw_schema = yaml.safe_load(f)
                if raw_schema and 'vuln_save' in raw_schema:
                    result['vuln_save'] = raw_schema['vuln_save']
            except Exception as e:
                logger.warning(f"Failed to read raw schema for {template_id}: {e}")

        return result
    
    def get_template_dir(self, template_id: str) -> str:
        """获取模板的实际源目录（内置或用户目录）。
        
        优先返回用户模板目录路径（存储在 _template_source_dirs 中为完整路径），
        若未找到则回退到内置模板目录拼接。
        
        Args:
            template_id: 模板ID
            
        Returns:
            模板的源目录完整路径
        """
        if template_id in self._template_source_dirs:
            return self._template_source_dirs[template_id]
        return os.path.join(self.templates_dir, template_id)

    def _get_template_source_dir(self, template_id: str) -> Optional[str]:
        """获取模板的实际源目录（内置或用户目录）。"""
        return self._template_source_dirs.get(template_id, self.templates_dir)
    
    def get_template_file_path(self, template_id: str) -> Optional[str]:
        """获取模板 docx 文件的完整路径"""
        template = self._templates.get(template_id)
        if not template:
            return None
        
        template_dir = self._get_template_source_dir(template_id)
        if not template_dir:
            return None
        template_file = os.path.join(template_dir, template.template_file)
        
        if os.path.exists(template_file):
            return template_file
        return None
    
    def resolve_data_sources(self, template_id: str, 
                             db_data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        解析模板所需的数据源
        
        Args:
            template_id: 模板ID
            db_data: 数据库数据 (如漏洞列表、ICP缓存等)
            
        Returns:
            {source_id: data, ...}
        """
        template = self._templates.get(template_id)
        if not template:
            return {}
        
        db_data = db_data if db_data is not None else {}
        resolved = {}
        
        for ds in template.data_sources:
            if ds.type == 'config' and ds.config_key:
                # 从全局配置读取
                resolved[ds.id] = self.config.get(ds.config_key, [])
            elif ds.type == 'database':
                # 从传入的数据库数据读取
                resolved[ds.id] = db_data.get(ds.id, [])
            elif ds.type == 'api':
                # API 类型由前端处理
                resolved[ds.id] = {"endpoint": ds.endpoint}
        
        # 处理字段级别的 source 属性 (如 config.risk_levels, config.supplierName)
        for field in template.fields:
            if field.source and field.source.startswith('config.'):
                config_key = field.source.replace('config.', '')
                if field.source not in resolved:
                    resolved[field.source] = self.config.get(config_key, [])
        
        # 处理嵌套 columns 中的 source（需要读取原始 schema）
        template_path = self._get_template_source_dir(template_id)
        if template_path is None:
            template_path = os.path.join(self.templates_dir, template_id)
        schema_path = os.path.join(template_path, "schema.yaml")
        if os.path.exists(schema_path):
            try:
                with open(schema_path, 'r', encoding='utf-8') as f:
                    raw_schema = yaml.safe_load(f)
                for field_data in raw_schema.get('fields', []):
                    columns = field_data.get('columns', [])
                    for col in columns:
                        if isinstance(col, dict):
                            src = col.get('source', '')
                            if src.startswith('config.') and src not in resolved:
                                config_key = src.replace('config.', '')
                                resolved[src] = self.config.get(config_key, [])
            except Exception as e:
                logger.warning(f"Failed to parse columns from schema: {e}")
        
        return resolved
    
    def validate_report_data(self, template_id: str, data: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """
        验证报告数据是否符合模板要求
        
        Returns:
            (是否有效, 错误信息列表)
        """
        template = self._templates.get(template_id)
        if not template:
            return False, [f"Template not found: {template_id}"]
        
        errors = []
        
        # 检查必填字段
        for field_def in template.fields:
            if field_def.required:
                value = data.get(field_def.key, "")
                if not value or (isinstance(value, str) and not value.strip()):
                    errors.append(f"字段 '{field_def.label}' 为必填项")
            
            # 检查字段验证规则
            if field_def.validation:
                pattern = field_def.validation.get('pattern')
                if pattern:
                    value = data.get(field_def.key, "")
                    if value and not re.match(pattern, str(value)):
                        errors.append(field_def.validation.get('message', f"字段 '{field_def.label}' 格式不正确"))
        
        # 检查全局验证规则
        for rule in template.validation_rules:
            if rule.rule == 'required':
                for field_key in rule.fields:
                    value = data.get(field_key, "")
                    if not value or (isinstance(value, str) and not value.strip()):
                        errors.append(rule.message)
                        break
        
        return len(errors) == 0, errors
    
    def build_replacements(self, template_id: str, data: Dict[str, Any], 
                          extra: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
        """
        根据模板定义和数据构建替换字典
        
        Args:
            template_id: 模板ID
            data: 前端提交的数据
            extra: 额外的替换项 (如 supplierName, reportTime)
            
        Returns:
            替换字典 {"#key#": "value", ...}
        """
        template = self._templates.get(template_id)
        if not template:
            return {}
        
        replacements = {}
        
        # 从模板字段构建
        for field_def in template.fields:
            # 优先使用模板中定义的占位符
            if field_def.template_placeholder:
                key = field_def.template_placeholder
            else:
                key = f"#{field_def.key}#"
            
            value = data.get(field_def.key)
            if value is None:
                value = field_def.default if field_def.default != 'today' else ''
            
            # 特殊处理：跳过图片类型和复杂类型，由专门的处理器处理
            if field_def.type in ('image', 'image_list', 'grouped_image_list'):
                continue
            if isinstance(value, list):
                # 图片列表等复杂类型，跳过文本替换
                continue
            else:
                replacements[key] = str(value) if value is not None else ""
        
        # 合并额外数据
        if extra:
            for k, v in extra.items():
                if not k.startswith("#"):
                    k = f"#{k}#"
                replacements[k] = str(v) if v is not None else ""
        
        return replacements
    
    def generate_output_path(self, template_id: str, data: Dict[str, Any], 
                            base_output_dir: str) -> str:
        """
        根据模板配置生成输出路径
        
        Args:
            template_id: 模板ID
            data: 报告数据
            base_output_dir: 基础输出目录
            
        Returns:
            完整的输出文件路径
        """
        template = self._templates.get(template_id)
        if not template:
            return os.path.join(base_output_dir, "report.docx")
        
        output_config = template.output_config
        
        # 解析文件名模式
        filename_pattern = output_config.get('filename_pattern', 'report_{date}.docx')
        output_dir_pattern = output_config.get('output_dir', '')
        
        # 替换变量
        now = datetime.now()
        replacements = {
            'date': now.strftime('%Y-%m-%d'),
            'datetime': now.strftime('%Y%m%d_%H%M%S'),
            'timestamp': str(int(now.timestamp()))
        }
        replacements.update(data)
        
        # 解析文件名
        filename = filename_pattern
        for key, value in replacements.items():
            filename = filename.replace(f'{{{key}}}', str(value) if value else '')
        
        # 清理文件名中的非法字符
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # 解析输出目录
        if output_dir_pattern:
            output_dir = output_dir_pattern
            for key, value in replacements.items():
                output_dir = output_dir.replace(f'{{{key}}}', str(value) if value else '')
            output_dir = os.path.join(base_output_dir, output_dir)
        else:
            output_dir = base_output_dir
        
        # 确保目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        return os.path.join(output_dir, filename)
    
    # Security: 允许的依赖包白名单
    ALLOWED_PACKAGES = {
        # 标准库 (Standard Library)
        'abc', 'argparse', 'ast', 'base64', 'code', 'collections', 'contextlib', 'copy', 
        'csv', 'ctypes', 'dataclasses', 'datetime', 'email', 'enum', 'errno', 'fcntl', 
        'filecmp', 'fnmatch', 'functools', 'glob', 'hashlib', 'importlib', 'io', 'itertools', 
        'json', 'locale', 'logging', 'math', 'multiprocessing', 'operator', 'os', 'pathlib', 
        'pickle', 'platform', 'plistlib', 'pprint', 'random', 're', 'shlex', 'shutil', 
        'signal', 'socket', 'sqlite3', 'stat', 'string', 'struct', 'subprocess', 'sys', 
        'sysconfig', 'tempfile', 'textwrap', 'threading', 'time', 'traceback', 'typing', 
        'unittest', 'urllib', 'uuid', 'warnings', 'weakref', 'xml', 'zipfile',

        # 第三方库 (Third Party)
        'requests',            # HTTP 请求
        'pandas',              # 数据处理
        'numpy',               # 数值计算
        'openpyxl',            # Excel 处理
        'python-docx', 'docx', # Word 处理
        'docxcompose',         # Word 合并
        'pillow', 'PIL',       # 图片处理
        'lxml',                # XML/HTML 解析
        'pyyaml', 'yaml',      # YAML 解析
        'beautifulsoup4', 'bs4', # 网页解析
        'matplotlib',          # 图表绘制
        'tldextract',          # 域名解析
        'uvicorn',             # ASGI 服务器
        'fastapi',             # Web 框架
        'pydantic',            # 数据验证
        'packaging',           # 版本处理
        'PyInstaller',         # 打包工具

        # 项目内部包 (Project Internal)
        'core',                # SDK facade layer
        'backend',             # Core implementation layer (GenerationContext etc.)

        # 常用底层依赖
        'urllib3', 'certifi', 'idna', 'charset-normalizer', 'six',
        'python-dateutil', 'pytz', 'typing-extensions', 'click', 'colorama'
    }

    def check_dependencies(self, template_id: str, raise_on_missing: bool = False) -> Tuple[bool, List[str]]:
        """
        检查模板依赖是否满足（解决问题 9：模板依赖管理缺失）
        Security Fix: 增加依赖白名单检查，防止恶意依赖引入
        
        Args:
            template_id: 模板ID
            raise_on_missing: 如果为 True，依赖缺失时抛出异常
            
        Returns:
            (是否满足, 缺失的依赖列表)
            
        Raises:
            TemplateNotFoundError: 模板不存在
            DependencyError: 当 raise_on_missing=True 且依赖缺失时
        """
        template = self._templates.get(template_id)
        if not template:
            if raise_on_missing:
                raise TemplateNotFoundError(template_id)
            return False, [f"Template not found: {template_id}"]
        
        # 从模板目录读取 schema.yaml 获取依赖声明
        template_path = self._get_template_source_dir(template_id)
        if template_path is None:
            template_path = os.path.join(self.templates_dir, template_id)
        schema_path = os.path.join(template_path, "schema.yaml")
        
        if not os.path.exists(schema_path):
            return True, []
        
        try:
            with open(schema_path, 'r', encoding='utf-8') as f:
                schema = yaml.safe_load(f)
            
            dependencies = schema.get('dependencies', [])
            if not dependencies:
                return True, []
            
            missing = []
            for dep in dependencies:
                # 格式解析：requests>=2.28.0 -> requests
                pkg_name = dep.split('>=')[0].split('==')[0].split('<')[0].split('[')[0].strip().lower()
                
                # Security Check: 白名单验证
                if pkg_name not in self.ALLOWED_PACKAGES:
                    logger.warning(f"Security Alert: Template {template_id} requests unapproved dependency '{pkg_name}'")
                    if raise_on_missing:
                         raise DependencyError(template_id, [f"Unapproved dependency: {pkg_name}"])
                    missing.append(f"{dep} (Unapproved)")
                    continue

                try:
                    # 简单检查：尝试导入包名
                    __import__(pkg_name)
                except ImportError:
                    missing.append(dep)
            
            if missing:
                logger.warning(f"Template {template_id} has missing dependencies: {missing}")
                if raise_on_missing:
                    raise DependencyError(template_id, missing)
                return False, missing
            
            return True, []
            
        except DependencyError:
            raise
        except Exception as e:
            error_msg = f"Error checking dependencies: {str(e)}"
            logger.error(f"Failed to check dependencies for {template_id}: {e}")
            if raise_on_missing:
                raise TemplateLoadError(template_id, error_msg)
            return False, [error_msg]
    
    def reload_templates(self):
        """重新加载所有模板"""
        import sys
        from .base_handler import HandlerRegistry
        
        # 清理 HandlerRegistry（解决问题 7：重复注册问题）
        HandlerRegistry.clear()
        logger.info("Cleared HandlerRegistry")
        
        # 清空路由缓存
        self._template_routers.clear()
        
        # 清理 sys.modules 中的旧模块（解决问题 6：sys.modules 缓存问题）
        modules_to_remove = []
        for module_name in sys.modules.keys():
            if module_name.startswith('templates.') and '.handler' in module_name:
                modules_to_remove.append(module_name)
        
        for module_name in modules_to_remove:
            del sys.modules[module_name]
            logger.info(f"Removed old module from cache: {module_name}")
        
        # 清理 LRU 缓存（解决问题 13：性能优化）
        self._get_cached_schema.cache_clear()
        logger.info("Cleared schema cache")
        
        self._templates.clear()
        self._template_versions.clear()
        self._template_source_dirs.clear()
        self._load_all_templates()
    
    def get_template_details(self, template_id: str) -> Optional[Dict[str, Any]]:
        """
        获取模板的详细信息（包括文件大小、字段数量等）
        
        Args:
            template_id: 模板ID
            
        Returns:
            模板详细信息字典，如果模板不存在则返回 None
        """
        template = self._templates.get(template_id)
        if not template:
            return None
        
        template_path = self._get_template_source_dir(template_id)
        if template_path is None:
            template_path = os.path.join(self.templates_dir, template_id)
        schema_path = os.path.join(template_path, "schema.yaml")
        docx_path = os.path.join(template_path, "template.docx")
        
        # 获取文件大小
        schema_size = os.path.getsize(schema_path) if os.path.exists(schema_path) else 0
        docx_size = os.path.getsize(docx_path) if os.path.exists(docx_path) else 0
        total_size = schema_size + docx_size
        
        # 获取文件修改时间
        mtime = os.path.getmtime(schema_path) if os.path.exists(schema_path) else 0
        update_time = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S") if mtime else ""
        
        return {
            "id": template.id,
            "name": template.name,
            "description": template.description,
            "icon": template.icon,
            "version": template.version,
            "author": template.author,
            "create_time": template.create_time,
            "update_time": update_time,
            "order": template.order,
            "field_count": len(template.fields),
            "group_count": len(template.field_groups),
            "file_size": total_size,
            "file_size_mb": round(total_size / 1024 / 1024, 2),
            "has_schema": os.path.exists(schema_path),
            "has_docx": os.path.exists(docx_path),
            "has_widgets": os.path.isdir(os.path.join(template_path, "widgets")),
            "is_default": template.id == self.default_template_id
        }
    
    def delete_template(self, template_id: str) -> Tuple[bool, str]:
        """
        删除指定模板（仅允许删除用户自定义模板）
        
        Args:
            template_id: 模板ID
            
        Returns:
            (成功标志, 消息)
        """
        import shutil
        
        # 检查模板是否存在
        if template_id not in self._templates:
            return False, f"模板不存在: {template_id}"
        
        # 内置模板不可删除
        source_dir = self._template_source_dirs.get(template_id)
        if source_dir is None or os.path.dirname(os.path.abspath(source_dir)) == os.path.abspath(self.templates_dir):
            return False, "内置模板不可删除"
        
        # 防止删除默认模板
        if template_id == self.default_template_id and len(self._templates) == 1:
            return False, "无法删除唯一的模板"
        
        template_path = source_dir
        
        try:
            # 删除模板目录
            if os.path.exists(template_path):
                shutil.rmtree(template_path)
                logger.info(f"Deleted user template directory: {template_path}")
            
            # 从内存中移除
            del self._templates[template_id]
            if template_id in self._template_versions:
                del self._template_versions[template_id]
            if template_id in self._template_source_dirs:
                del self._template_source_dirs[template_id]
            
            return True, f"模板 {template_id} 已删除"
        except Exception as e:
            logger.error(f"Failed to delete template {template_id}: {str(e)}")
            return False, f"删除失败: {str(e)}"
    
    def update_config(self, config: Dict[str, Any]):
        """更新全局配置"""
        self.config = config
    
    @property
    def template_ids(self) -> List[str]:
        """获取所有模板ID"""
        return list(self._templates.keys())
    
    @property
    def user_template_ids(self) -> List[str]:
        """获取用户自定义模板ID列表"""
        if not self.user_templates_dir:
            return []
        user_dir = os.path.abspath(self.user_templates_dir)
        return [
            tid for tid, src in self._template_source_dirs.items()
            if os.path.dirname(os.path.abspath(src)) == user_dir
        ]
    
    @property
    def default_template_id(self) -> Optional[str]:
        """获取默认模板ID (order 最小的模板)"""
        if self._templates:
            # 按 order 排序，返回第一个
            sorted_templates = sorted(self._templates.values(), key=lambda t: t.order)
            return sorted_templates[0].id if sorted_templates else None
        return None
