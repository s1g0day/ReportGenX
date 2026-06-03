# 模板开发完整指南

> ReportGenX 模板系统开发文档
> 版本: 2.0.0 | 更新日期: 2026-06-04 | 适用版本: 0.21.1

---

## 目录

1. [快速开始](#0-快速开始)
2. [架构概述](#1-架构概述)
3. [PLUGIN 描述符规范](#2-plugin-描述符规范)
4. [GenerationContext API](#3-generationcontext-api)
5. [Schema 规范](#4-schema-规范)
6. [runtime.yaml 参考](#5-runtimeyaml-参考)
7. [Widget 系统](#6-widget-系统)
8. [Core SDK 门面](#7-core-sdk-门面)
9. [最佳实践](#8-最佳实践)
10. [常见问题](#9-常见问题)

---

## 0. 快速开始

5 分钟创建一个最小模板。

### 目录结构

```
backend/templates/my_template/
├── schema.yaml        # 表单定义 (必需)
├── handler.py         # 生成逻辑 (必需)
├── template.docx      # Word 模板 (可选)
└── runtime.yaml       # 日志配置 (推荐)
```

### handler.py

```python
from backend.core.generation_context import GenerationContext

PLUGIN = {"id": "my_template", "execute": None}

def execute(data, output_dir, template_manager, config, template_id):
    template_dir = template_manager.get_template_dir(template_id)
    info = template_manager.get_template_info(template_id)
    ctx = GenerationContext(template_dir, info, config, output_dir)

    doc = ctx.load_document()
    ctx.replace_text({
        "#title#": data.get("title", ""),
        "#content#": data.get("content", ""),
    })
    path = ctx.save_document(doc, ctx.build_output_path(data.get("unit_name", "unknown"), "report.docx"))
    ctx.postprocess(path, data, log_prefix="my_template")
    return {"success": True, "report_path": path, "message": "生成成功", "errors": []}

PLUGIN["execute"] = execute
```

### schema.yaml

```yaml
id: my_template
name: 我的模板
version: "1.0.0"
fields:
  - id: title
    type: text
    label: 标题
    required: true
  - id: content
    type: textarea
    label: 内容
output_config:
  filename: "{{unit_name}}_report.docx"
```

### 加载模板

```bash
curl -X POST http://127.0.0.1:8000/api/templates/reload
```

前端刷新模板列表，选择 `my_template` 即可。

> 详细说明见后续章节。

---

## 1. 架构概述

### 模板生命周期

```
preprocess()  →  validate()  →  generate(data, ctx)  →  postprocess (ctx.postprocess)
    │                 │                  │                         │
    填充默认值      校验数据          使用 ctx 生成文档           TXT + SQLite 日志
```

### 关键组件

| 组件 | 位置 | 职责 |
|------|------|------|
| `TemplateManager` | `backend/core/template_manager.py` | 扫描/加载/校验模板 |
| `SchemaLoader` | `backend/core/schema_loader.py` | YAML → Pydantic 解析 |
| `PluginRuntime` | `backend/plugin_host/runtime.py` | 执行调度 (descriptor/hybrid/legacy/isolated) |
| `GenerationContext` | `backend/core/generation_context.py` | 模板服务注入层 |
| `AppFormRenderer` | `src/js/form-renderer.js` | 前端动态表单引擎 |

### 执行链路

```
POST /api/templates/{id}/generate
  → PluginRuntime.execute(template_id, data, ...)
    → 解析 PLUGIN 描述符
    → 调用 execute(data, output_dir, template_manager, config, template_id)
      → 创建 GenerationContext
      → preprocess() → generate() → ctx.postprocess()
    → 返回 {"success", "report_path", "message", "errors", "execution_meta"}
```

---

## 2. PLUGIN 描述符规范

### 标准模式（当前所有 6 个模板均使用）

每个模板在 `handler.py` 中导出模块级 `PLUGIN` 字典：

```python
PLUGIN = {
    "id": "template_id",       # 必须与目录名一致
    "execute": execute,        # 可调用对象 (函数)
}
```

### execute() 函数签名

```python
def execute(
    data: Dict[str, Any],
    output_dir: str,
    template_manager: Any,
    config: Optional[Dict[str, Any]] = None,
    template_id: str = "template_name",
) -> Dict[str, Any]:
```

**返回值规范**:

```python
{
    "success": bool,
    "report_path": str,        # 生成文件的绝对路径
    "message": str,             # 成功/失败消息
    "errors": List[str],        # 校验错误列表 (成功时为空)
}
```

### execute() 内部最佳实践

```python
def execute(data, output_dir, template_manager, config, template_id="my_template"):
    from backend.core.generation_context import GenerationContext
    from backend.core.logger import setup_logger
    from backend.core.schema_loader import SchemaLoader

    template_dir = template_manager.get_template_dir(template_id)
    template_info = SchemaLoader.load_schema(template_dir)
    runtime_cfg = SchemaLoader.load_runtime(template_dir)

    logger = setup_logger('TemplateLogger')
    ctx = GenerationContext(template_dir, template_info, config, output_dir, logger)

    # 1. Preprocess
    processed = preprocess(data, config or {})

    # 2. Validate
    is_valid, errors = validate(processed, config or {}, template_info)
    if not is_valid:
        return {"success": False, "report_path": "", "message": "; ".join(errors), "errors": errors}

    # 3. Generate
    success, path, msg = generate(processed, ctx)

    # 4. Postprocess
    if success:
        ctx.postprocess(path, processed, **runtime_cfg_params)

    return {"success": success, "report_path": path, "message": msg, "errors": [] if success else errors}
```

> ⚠️ **不再使用** `BaseTemplateHandler` 类继承 + `HandlerRegistry.register()`。当前 6 个模板全部基于 PLUGIN descriptor 模式。与 `BaseTemplateHandler` 类名保留仅作为旧模板兼容。

---

## 3. GenerationContext API

`GenerationContext` 是注入到 `generate()` 函数的服务上下文。模板通过 `ctx` 访问所有框架能力。

### 文档操作

| 方法 | 说明 |
|------|------|
| `ctx.load_document() → Document` | 加载 template.docx |
| `ctx.doc` (property) | 懒加载获取当前文档对象 |
| `ctx.editor` (property) | 获取 `DocumentEditor` 实例 |
| `ctx.img_processor` (property) | 获取 `DocumentImageProcessor` 实例 |

### 文本替换

```python
ctx.replace_text(replacements: Dict[str, str], enable_risk_color=False, risk_key=None)
# 批量替换占位符。例: {"#name#": "张三", "#date#": "2026-05-30"}

ctx.replace_text_colored(replacements: Dict[str, str])
# 替换文本并启用风险等级颜色
```

### 图片处理

```python
ctx.process_single_image('#placeholder#', image_data, fallback_text='（未提供）')
# 处理单张图片占位符。image_data 可以是路径字符串或 {'path': '...'} 字典

ctx.process_image_list('#placeholder#', images: List, keyword=None)
# 处理图片列表。images 可以是 [str, ...] 或 [{'path': '...', 'description': '...'}, ...]
# keyword: 可选，用于定位表格单元格中的占位符
```

### 表格操作

```python
ctx.populate_table(header_text, data_rows, row_builder_func, keep_header_rows=1, clear_indent=False)
# 通过表头文本定位表格，用数据行填充
# row_builder_func(row_cells, data_item): 回调函数，填写每行数据

ctx.clear_paragraph_indent(para)
# 清除段落首行缩进
```

### 目录

```python
ctx.insert_toc('#toc#', '目  录')
# 在占位符位置插入 Word TOC 域
```

### 输出与保存

```python
ctx.save(filename: str) → str
# 保存文档到 output_dir/filename，自动处理同名冲突。返回最终路径

ctx.save_document(doc, output_path) → str
# 保存指定文档，处理同名冲突

ctx.build_output_path(unit_name, filename) → str
# 构建输出路径: output_dir/unit_name/filename，自动清理非法字符
```

### 替换字典构建

```python
ctx.build_replacements(data: Dict, extra: Optional[Dict] = None) → Dict[str, str]
# 根据 template_info.fields 自动构建 #key# → value 映射
# 自动跳过 image/list 类型字段
# extra: 额外键值对，键自动加 # 前缀（如果没有的话）
```

### 工具方法

```python
ctx.get_date(fmt='%Y-%m-%d') → str              # 当前日期
ctx.gen_id(prefix='RPT', random_length=4) → str  # 生成报告 ID, 如 'RPT-20260530-ABCD'
ctx.sanitize_filename(name) → str                # 清理非法文件名字符
ctx.create_output_dir(base_dir, sub_dir) → str   # 创建并返回目录路径
```

### 漏洞库查询

```python
ctx.lookup_vulnerability(vuln_id) → Dict      # 按 ID 或名称查漏洞，返回完整记录
ctx.get_vulnerability_name(vuln_id) → str     # 获取漏洞显示名称
```

### 摘要生成

```python
ctx.summarize_count(items, type_key, type_names, template_zero, template_single, template_multi) → str
# 生成计数型摘要文本

ctx.summarize_data(items, type_key, count_key, template_zero, template_with_data) → (str, int)
# 生成数据量型摘要文本

ctx.summary_templates → SummaryTemplates
# 预定义的摘要模板配置
```

### 后处理（日志）

```python
ctx.postprocess(output_path, data, log_prefix, log_fields, db_table, db_name, db_field_map)
# 写入 TXT 日志 + SQLite 记录
# 参数通常从 runtime.yaml 读取
```

### Fallback 报告

```python
ctx.generate_fallback(data) → (bool, str, str)
# 当 template.docx 缺失时，根据 schema 字段定义自动生成简单文档
```

---

## 4. Schema 规范

### 完整 YAML 结构

```yaml
# ── 基本信息 ──
id: template_id           # 唯一标识 (必需，与目录名一致)
name: 模板名称            # 显示名称 (必需)
description: 描述         # (可选)
version: "1.0.0"         # (推荐)
icon: "📄"               # 图标 emoji (可选)
author: 作者             # (可选)
order: 1                 # 排序权重
create_time: "2026-01-01"
update_time: "2026-01-01"

template_file: template.docx   # Word 模板文件名

# ── 依赖声明 (可选) ──
dependencies:
  - requests>=2.28.0

# ── 字段分组 ──
field_groups:
  - id: basic
    name: 基础信息
    icon: "📋"           # 分组图标 (可选)
    order: 1             # 显示排序
    collapsed: false     # 默认是否折叠

# ── 数据源 ──
data_sources:
  - id: vulns
    type: database       # database | config | api
    description: 漏洞库
  - id: risk_levels
    type: config
    config_key: risk_levels

# ── 字段定义 ──
fields:
  - key: field_key       # 字段键名 (必需，唯一)
    label: 字段标签      # 显示标签 (必需)
    type: text           # 字段类型 (见下方)
    required: false
    default: ""          # "today" 表示当前日期
    placeholder: ""
    group: basic         # 所属分组 id
    order: 1
    readonly: false
    source: vulns        # 数据源引用 (type=select/searchable_select 时)
    options: []          # 静态选项列表
    template_placeholder: "#fieldKey#"  # 模板中的占位符

# ── 行为定义 ──
behaviors:
  - id: behavior_id
    trigger:
      field: trigger_field
      event: change      # change | data_changed | manual
    actions:
      - type: compute    # compute | api_call | set_value
        target: target_field
        rules:
          value1: result1

# ── 验证规则 ──
validation:
  rules:
    - fields: [field1, field2]
      rule: required
      message: 错误提示

# ── 输出配置 ──
output:
  filename_pattern: "{title}_{date}.docx"
  output_dir: "output/{unit_name}"
  log_to_db: true
  log_table: template_records

# ── 预览配置 ──
preview:
  enabled: true
  fields:
    - key: field_key
      label: 显示标签
```

### 字段类型

| 类型 | 说明 | 特有属性 |
|------|------|----------|
| `text` | 单行文本 | `placeholder` |
| `textarea` | 多行文本 | `rows` |
| `select` | 下拉选择 | `options`, `source` |
| `searchable_select` | 可搜索下拉 | `search_placeholder`, `source`, `display_field`, `value_field` |
| `date` | 日期选择器 | `default: "today"`, `format` |
| `image` | 单张图片上传 | `accept`, `max_size_mb`, `paste_enabled` |
| `image_list` | 多张图片上传 | `max_count`, `with_description`, `description_placeholder` |
| `grouped_image_list` | 分组图片列表 | `groups`, `max_per_group` |
| `checkbox` | 单个复选框 | `checked_value`, `unchecked_value` |
| `checkbox_group` | 复选框组 | `options` |
| `number` | 数字输入 | `min`, `max`, `step` |
| `array` | 动态数组（可增删行） | `item_schema`, `min_items`, `max_items` |
| `widget` | 自定义 HTML 组件 | `widget_file`（从 `widgets/` 目录加载） |
| `hidden` | 隐藏字段 | - |

### 选项格式

```yaml
# 简单格式
options:
  - 选项1
  - 选项2

# 对象格式 (推荐)
options:
  - value: opt1
    label: 选项一
  - value: opt2
    label: 选项二
```

### 行为类型

| 类型 | 说明 |
|------|------|
| `compute` | 根据规则计算目标字段值 |
| `api_call` | 调用后端 API 获取数据填充 |
| `set_value` | 直接设置目标字段值 |

### 数据源类型

| 类型 | 说明 | 配置键 |
|------|------|--------|
| `database` | 从 SQLite 数据库读取 | - |
| `config` | 从 config.yaml 读取 | `config_key` |
| `api` | 调用外部 API | `endpoint` |

### 漏洞库保存配置 (`vuln_save`)

当模板包含漏洞表单字段时，可声明 `vuln_save` 映射，自动将生成报告中的漏洞数据写入漏洞库：

```yaml
vuln_save:
  name_field: vuln_name          # 漏洞名称字段 key（用于去重判断）
  mapping:
    name: vuln_name              # 漏洞库列 → 表单字段 key
    description: description
    suggestion: repair_suggestion
    level: hazard_level
```

- `name_field` 指定漏洞名称对应的表单字段，用于判断是否已存在同名漏洞
- `mapping` 中键为漏洞库列名，值为表单字段 key
- 不声明 `vuln_save` 则跳过自动入库
- 参考：`backend/templates/single_vuln_report/schema.yaml`、`backend/templates/vuln_report/schema.yaml`

### 字段联动 (`dependent_fields`)

声明字段之间的联动关系，支持模板字符串、自动生成和预计算：

```yaml
dependent_fields:
  # 简单模板替换 — 当一个字段变化时，自动填充另一个字段
  entry_text:
    trigger_field: target_name
    template: "通过前期信息收集，查询${target_name}备案域名..."

  # 自动生成 + 预计算
  report_conclusion:
    trigger_fields: [internet_vulns, intranet_vulns, controlled_servers]
    auto_generate: true
    computed_template: "${target_name}存在有效高危漏洞${_effective_high_vulns}个..."
    pre_compute:
      count_vulns:
        source: [internet_vulns, intranet_vulns]
        level_field: vuln_level
      output:
        _effective_high_vulns: count_high_critical_vulns
```

- `trigger_field` / `trigger_fields`：联动触发源字段
- `template`：使用 `${fieldKey}` 引用表单字段值
- `auto_generate: true`：字段值变化时自动重新计算
- `pre_compute`：在模板替换前执行预计算（如统计、聚合），结果注入到 `computed_template` 的变量上下文
- `output`：预计算结果变量名映射（`_` 前缀变量为内部变量）
- 参考：`backend/templates/intranet_vuln/schema.yaml`、`backend/templates/Attack_Defense/schema.yaml`

---

## 5. runtime.yaml 参考

每个模板可包含 `runtime.yaml` 来配置后处理行为：

```yaml
log_prefix: template_name      # TXT 日志文件前缀
log_fields:                    # 要记录到 TXT 日志的字段列表
  - report_date
  - unit_name
  - vuln_name

db_table: template_records     # SQLite 表名
db_fields:                     # 列名 → data key 映射
  report_date: report_date
  unit_name: unit_name
```

- `log_prefix` 和 `log_fields` 配合：写入 `{date}_{log_prefix}_output.txt`
- `db_table` 和 `db_fields` 配合：写入 `{date}_output.db` 的 SQLite 表
- 所有字段均为可选 — 不提供则不记录

---

## 6. Widget 系统

模板可以在 `widgets/` 目录下放置自定义 JS/CSS 文件，扩展前端表单功能。

### 目录结构

```
backend/templates/my_template/
└── widgets/
    ├── vuln_list.js    # Widget JS (工厂函数)
    └── style.css       # Widget 样式 (可选)
```

### Widget JS 规范

Widget 通过 `window.__widgetFactories` 注册：

```javascript
// widgets/vuln_list.js
(function() {
    window.__widgetFactories = window.__widgetFactories || {};

    window.__widgetFactories['my_widget'] = function(container, callbacks) {
        // callbacks 提供:
        //   callbacks.getData()       → 获取当前表单数据
        //   callbacks.setData(data)   → 更新表单数据
        //   callbacks.getFormValue(key)   → 获取指定字段值
        //   callbacks.setFormValue(key, val) → 设置指定字段值
        //   callbacks.uploadImage(file)     → 上传图片
        //   callbacks.apiRequest(opts)      → 调用后端 API
        //   callbacks.dataSources           → 模板数据源
        //   callbacks.getConfig()           → 获取全局配置
        //   callbacks.toast(msg, type)      → 显示提示
        //   callbacks.openImagePreview(url) → 打开图片预览

        container.innerHTML = `<div class="my-widget">...</div>`;
    };
})();
```

### 在 schema 中声明

```yaml
fields:
  - key: vuln_table
    label: 漏洞列表
    type: widget
    group: detail
    order: 1
    widget_file: "vuln_list.js"
```

### Widget 文件服务

模板级 Widget 通过 `GET /api/templates/{template_id}/widgets/<file>` 加载。

### 共享 Widget（`backend/widgets/`）

项目级别共享 Widget 目录 `backend/widgets/` 存放跨模板复用的通用组件，当前包含：

```
backend/widgets/
├── vuln_list.js    # 通用漏洞列表 Widget（~45KB, schema-driven）
└── style.css       # 共享样式
```

- 共享 Widget 同样通过 `window.__widgetFactories` 注册
- Schema 中 `widget_file` 查找顺序：先查模板的 `widgets/` 目录，再查 `backend/widgets/`
- `vuln_list.js` 合并了 4 个模板中近似的漏洞列表代码为单一共享实现
- 共享 Widget 在打包时通过 `extraResources` 包含到应用目录

---

## 7. Core SDK 门面

`core/` 包是模板可导入的 SDK 门面层，从 `backend.core.*` 重新导出稳定 API。

```python
from core import (
    gen_report_id,           # 生成报告 ID → 'RPT-20260530-ABCD'
    set_default_dates,       # 填充日期默认值
    set_supplier_defaults,   # 填充供应商默认值
    GenerationContext,       # 生成上下文 (通常由 execute() 内创建)
    SummaryGenerator,        # 摘要生成器
    SummaryTemplates,        # 摘要模板配置
)
```

> **设计规则**: `core/` 是纯 re-export 门面 — 不应在此添加独立实现。所有业务逻辑实现在 `backend/core/`。

---

## 8. 最佳实践

### Schema 设计

- ✅ 字段 `key` 使用 snake_case 命名
- ✅ 为每个字段指定 `group` 和 `order` 确保正确分组排序
- ✅ 必填字段设置 `required: true`，在 `validation` 中声明规则
- ✅ 使用 `template_placeholder` 明确占位符格式 (如 `#fieldName#`)
- ✅ 复用已有数据源 (`config.risk_levels`, `config.hazard_type` 等)

### Handler 开发

- ✅ 将纯业务逻辑（preprocess/validate/generate）写为独立函数
- ✅ `execute()` 只负责流程编排：创建 ctx → preprocess → validate → generate → postprocess
- ✅ 在 `generate()` 中通过 `ctx` 访问所有框架服务
- ✅ 异常处理完善 — 捕获异常并返回明确的错误消息
- ✅ **无论是否有图片，都应调用 `ctx.process_single_image()` / `ctx.process_image_list()` 来清理占位符**

### 模板文件

- ✅ Word 模板中使用一致格式的占位符：`#fieldKey#`
- ✅ 图片占位符放在独立段落中
- ✅ 保持模板结构清晰，避免过度样式化

---

## 9. 常见问题

### Q: 新模板不显示？

1. 检查 `schema.yaml` YAML 语法是否正确
2. 检查 `handler.py` 是否有 Python 语法错误
3. 确认目录名仅含字母、数字、下划线
4. 调用 `POST /api/templates/reload` 热加载

### Q: 字段不显示？

1. 检查字段的 `group` 是否在 `field_groups` 中定义
2. 检查 `order` 是否设置正确

### Q: 报告生成失败？

1. 检查 `template.docx` 是否存在，或确保代码可处理 fallback 情况
2. 检查占位符格式与 `template_placeholder` 是否一致
3. 查看后端控制台错误日志

### Q: PLUGIN descriptor 和旧 BaseTemplateHandler 的区别？

`BaseTemplateHandler` 是旧的类继承模式，已不再推荐。当前所有 6 个模板使用 PLUGIN descriptor + `execute()` 函数模式：
- 不再继承任何类
- 通过 `GenerationContext(ctx)` 获取框架服务
- `PLUGIN` 字典告诉 `PluginRuntime` 如何调用模板

---

## 现有模板参考

| 模板 ID | 名称 | 字段数 | 复杂度 | 查看 |
|---------|------|--------|--------|------|
| `vuln_report` | 漏洞报告 | ~18 | 中 | `backend/templates/vuln_report/` |
| `intrusion_report` | 入侵痕迹报告 | ~25 | 中 | `backend/templates/intrusion_report/` |
| `penetration_test` | 渗透测试报告 | ~30+ | 高 | `backend/templates/penetration_test/` |
| `Attack_Defense` | 攻防演练报告 | ~40+ | 高 | `backend/templates/Attack_Defense/` |
| `single_vuln_report` | 单个漏洞报告 | ~10 | 低 | `backend/templates/single_vuln_report/` |
| `intranet_vuln` | 内网测试报告 | ~40+ | 高 | `backend/templates/intranet_vuln/` |

推荐从 `single_vuln_report` 开始阅读 — 代码最简洁、模式最清晰。
