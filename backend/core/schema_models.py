# -*- coding: utf-8 -*-
"""
Pydantic schema models for template definitions.

Extracted from template_manager.py to break a circular dependency
between schema_loader.py and template_manager.py.

All model classes, shared constants, and the validate_template_id
function live here so both schema_loader and template_manager can
import them without creating import cycles.
"""

import re
from typing import Any, Dict, List

from pydantic import BaseModel, Field, field_validator

from .logger import setup_logger

logger = setup_logger("SchemaModels")


# =============================================================================
# Shared constants
# =============================================================================

# 允许的字段类型
ALLOWED_FIELD_TYPES = {
    'text', 'select', 'textarea', 'date', 'image', 'image_list',
    'searchable_select', 'checkbox', 'checkbox_group', 'number',
    'array', 'widget', 'grouped_image_list'
}

# 允许的数据源类型
ALLOWED_DATA_SOURCE_TYPES = {'database', 'config', 'api'}

# 允许的行为动作类型
ALLOWED_ACTION_TYPES = {'api_call', 'compute', 'set_value', 'risk_compute', 'trigger_behavior'}


# =============================================================================
# Validation helpers
# =============================================================================

def validate_template_id(template_id: str) -> bool:
    """
    验证模板 ID 是否符合 Python 模块命名规范

    规则：
    - 只允许字母、数字、下划线
    - 不能以数字开头
    - 推荐使用 snake_case 命名（如 my_template）

    Args:
        template_id: 模板 ID

    Returns:
        bool: 是否有效

    Examples:
        >>> validate_template_id("my_template")
        True
        >>> validate_template_id("my-template")
        False
        >>> validate_template_id("123test")
        False
    """
    if not template_id:
        return False

    # 只允许字母、数字、下划线，且不能以数字开头
    pattern = r'^[a-zA-Z_][a-zA-Z0-9_]*$'
    return bool(re.match(pattern, template_id))


# =============================================================================
# Pydantic model classes
# =============================================================================

class FieldDefinition(BaseModel):
    """字段定义 (Pydantic 模型，带运行时验证)"""
    key: str = Field(..., description="字段键名 (对应模板中的占位符，不含 #)")
    label: str = Field(..., description="显示标签")
    type: str = Field(..., description="字段类型")
    required: bool = Field(default=False, description="是否必填")
    default: Any = Field(default="", description="默认值")
    placeholder: str = Field(default="", description="输入提示")
    help_text: str = Field(default="", description="帮助文本")
    options: List[Any] = Field(default_factory=list, description="下拉选项")
    source: str = Field(default="", description="数据源")
    max_count: int = Field(default=5, ge=1, le=100, description="最大数量")
    group: str = Field(default="default", description="字段分组")
    order: int = Field(default=0, ge=0, description="显示顺序")
    readonly: bool = Field(default=False, description="是否只读")
    computed: bool = Field(default=False, description="是否计算字段")
    compute_from: str = Field(default="", description="计算来源字段")
    compute_rule: Dict[str, Any] = Field(default_factory=dict, description="计算规则")
    on_change: str = Field(default="", description="值变化时触发的行为ID")
    validation: Dict[str, Any] = Field(default_factory=dict, description="验证规则")
    auto_generate: bool = Field(default=False, description="是否自动生成")
    auto_generate_rule: str = Field(default="", description="自动生成规则")
    auto_fill_from: str = Field(default="", description="自动填充来源")
    fill_field: str = Field(default="", description="填充的字段名")
    template_placeholder: str = Field(default="", description="模板中的占位符")
    inline_group: str = Field(default="", description="内联分组")
    rows: int = Field(default=3, ge=1, le=50, description="textarea 行数")
    accept: str = Field(default="", description="文件接受类型")
    max_size_mb: int = Field(default=5, ge=1, le=100, description="最大文件大小MB")
    paste_enabled: bool = Field(default=False, description="是否支持粘贴")
    with_description: bool = Field(default=False, description="图片是否带描述")
    description_placeholder: str = Field(default="", description="描述输入提示")
    search_placeholder: str = Field(default="", description="搜索提示")
    display_field: str = Field(default="", description="显示字段")
    value_field: str = Field(default="", description="值字段")
    editable_config: bool = Field(default=False, description="是否可编辑配置")
    save_to_config: bool = Field(default=False, description="是否保存到配置")
    presets: Dict[str, Any] = Field(default_factory=dict, description="预设值")
    show_if: Dict[str, Any] = Field(default_factory=dict, description="条件显示逻辑")
    transform: str = Field(default="", description="数据转换规则")
    columns: List[Dict[str, Any]] = Field(default_factory=list, description="列定义(复杂类型)")
    widget: str = Field(default="", description="Widget 文件名")
    model_config = {"extra": "allow"}  # 允许额外字段，保持向后兼容

    @field_validator('type')
    @classmethod
    def validate_field_type(cls, v: str) -> str:
        """验证字段类型是否在允许列表中"""
        if v not in ALLOWED_FIELD_TYPES:
            logger.warning(f"Unknown field type: {v}, allowed types: {ALLOWED_FIELD_TYPES}")
            # 不抛出异常，只警告，保持向后兼容
        return v

    @field_validator('key')
    @classmethod
    def validate_key_format(cls, v: str) -> str:
        """验证字段键名格式"""
        if not v:
            raise ValueError("Field key cannot be empty")
        if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', v):
            logger.warning(f"Field key '{v}' does not follow snake_case convention")
        return v


class FieldGroup(BaseModel):
    """字段分组"""
    id: str
    name: str
    icon: str = ""
    order: int = Field(default=0, ge=0)
    collapsed: bool = False


class DataSourceDef(BaseModel):
    """数据源定义"""
    id: str
    type: str
    description: str = ""
    config_key: str = ""
    endpoint: str = ""

    @field_validator('type')
    @classmethod
    def validate_type(cls, v: str) -> str:
        if v not in ALLOWED_DATA_SOURCE_TYPES:
            raise ValueError(f"Invalid data source type: {v}, allowed: {ALLOWED_DATA_SOURCE_TYPES}")
        return v


class BehaviorAction(BaseModel):
    """行为动作"""
    model_config = {"extra": "allow"}
    type: str
    endpoint: str = ""
    params: Dict[str, Any] = Field(default_factory=dict)
    result_mapping: Dict[str, Any] = Field(default_factory=dict)
    target: str = ""
    rules: Dict[str, Any] = Field(default_factory=dict)
    expression: str = ""  # compute 类型的表达式

    @field_validator('type')
    @classmethod
    def validate_type(cls, v: str) -> str:
        if v not in ALLOWED_ACTION_TYPES:
            logger.warning(f"Unknown action type: {v}")
        return v


class Behavior(BaseModel):
    """行为定义"""
    model_config = {"extra": "allow"}
    id: str
    trigger_field: str = ""
    trigger_event: str = "change"
    actions: List[BehaviorAction] = Field(default_factory=list)


class ValidationRule(BaseModel):
    """验证规则"""
    fields: List[str]
    rule: str
    message: str


class PreviewField(BaseModel):
    """预览字段"""
    key: str
    label: str


class TemplateInfo(BaseModel):
    """模板信息"""
    model_config = {"extra": "allow"}
    id: str
    name: str
    description: str = ""
    version: str = "1.0.0"
    order: int = Field(default=999, ge=0)
    template_file: str = ""
    icon: str = ""
    author: str = ""
    create_time: str = ""
    update_time: str = ""
    fields: List[FieldDefinition] = Field(default_factory=list)
    field_groups: List[FieldGroup] = Field(default_factory=list)
    data_sources: List[DataSourceDef] = Field(default_factory=list)
    behaviors: List[Behavior] = Field(default_factory=list)
    validation_rules: List[ValidationRule] = Field(default_factory=list)
    output_config: Dict[str, Any] = Field(default_factory=dict)
    preview_config: Dict[str, Any] = Field(default_factory=dict)
    dependent_fields: Dict[str, Any] = Field(default_factory=dict)
    summary_configs: Dict[str, Any] = Field(default_factory=dict)

    @field_validator('id')
    @classmethod
    def validate_template_id_field(cls, v: str) -> str:
        """验证模板 ID 格式 (snake_case)"""
        if not validate_template_id(v):
            raise ValueError(f"Invalid template ID format: {v}. Must be snake_case (e.g., my_template)")
        return v

    @field_validator('version')
    @classmethod
    def validate_version_format(cls, v: str) -> str:
        """验证版本号格式"""
        if not re.match(r'^\d+\.\d+\.\d+$', v):
            logger.warning(f"Version '{v}' does not follow semantic versioning (x.y.z)")
        return v
