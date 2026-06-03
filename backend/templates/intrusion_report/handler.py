# -*- coding: utf-8 -*-
"""
@Createtime: 2026-01-26
@Updatetime: 2026-05-29
@description: 入侵痕迹报告处理器 - pure function interface with GenerationContext injection.

Template exposes:
    preprocess(data, config) -> dict
    validate(data, config, template_info) -> (bool, list)
    generate(data, ctx) -> (bool, str, str)

No class inheritance. All SDK services accessed via ctx (GenerationContext).
"""

import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from core import gen_report_id, set_default_dates, set_supplier_defaults


# ═══════════════════════════════════════════════════════════════════
# Constants
# ═══════════════════════════════════════════════════════════════════

SEVERITY_LEVEL_MAP = {
    'critical': '严重', 'high': '高危', 'medium': '中危', 'low': '低危',
    '严重': '严重', '高危': '高危', '中危': '中危', '低危': '低危',
}

INTRUSION_TYPE_MAP = {
    'webshell': 'Webshell植入', 'backdoor': '后门程序', 'malware': '恶意软件',
    'data_theft': '数据窃取', 'privilege_escalation': '权限提升',
    'lateral_movement': '横向移动', 'crypto_mining': '挖矿木马',
    'ransomware': '勒索软件', 'other': '其他',
}

# ═══════════════════════════════════════════════════════════════════
# Pure functions — template business logic
# ═══════════════════════════════════════════════════════════════════

def preprocess(data: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Data preprocessing.

    1. Generate report ID
    2. Convert intrusion type display value
    3. Convert severity level
    4. Resolve attack method (custom input or vuln lookup)
    5. Set default dates
    6. Set supplier / analyst defaults
    """
    processed = data.copy()

    # 1. Generate report ID
    if not processed.get('report_id'):
        processed['report_id'] = gen_report_id(prefix="IR", use_sequence=False)

    # 2. Convert intrusion type
    intrusion_type = processed.get('intrusion_type', '')
    processed['intrusion_type_display'] = INTRUSION_TYPE_MAP.get(intrusion_type, intrusion_type)

    # 3. Convert severity level
    severity = processed.get('severity_level', '')
    processed['severity_level'] = SEVERITY_LEVEL_MAP.get(severity, severity)

    # 4. Resolve attack method
    attack_method_custom = processed.get('attack_method_custom', '').strip()
    attack_method_id = processed.get('attack_method', '')
    if attack_method_custom:
        processed['attack_method_display'] = attack_method_custom
    elif attack_method_id:
        # Resolved by the execute() adapter via ctx before preprocess is called;
        # fall back to the raw id if not pre-resolved.
        processed['attack_method_display'] = processed.get('attack_method_display', attack_method_id)
    else:
        processed['attack_method_display'] = ''

    # 5. Set default dates
    set_default_dates(processed, ['discovery_time', 'report_time'])

    # 6. Set analyst defaults
    set_supplier_defaults(processed, config, ['analyst_name'])

    return processed


def validate(
    data: Dict[str, Any],
    config: Dict[str, Any],
    template_info: Any,
) -> Tuple[bool, List[str]]:
    """Data validation using template field definitions."""
    errors = []

    if template_info:
        for field_def in template_info.fields:
            if field_def.required:
                value = data.get(field_def.key, "")
                if not value or (isinstance(value, str) and not value.strip()):
                    errors.append(f"Field '{field_def.label}' is required")

        for rule in template_info.validation_rules:
            if rule.rule == 'required':
                for field_key in rule.fields:
                    value = data.get(field_key, "")
                    if not value or (isinstance(value, str) and not value.strip()):
                        errors.append(rule.message)
                        break

    return len(errors) == 0, errors


def generate(data: Dict[str, Any], ctx: Any) -> Tuple[bool, str, str]:
    """
    Generate intrusion report.

    Args:
        data: Preprocessed form data
        ctx: GenerationContext providing all framework services

    Returns:
        (success, output_path, message)
    """
    try:
        # 1. Load document (falls back to generate_fallback if template missing)
        doc = ctx.load_document()
        if not doc:
            return ctx.generate_fallback(data)

        # 2. Build replacements
        extra_replacements = {
            '#intrusion_type#': data.get('intrusion_type_display', ''),
            '#attack_method#': data.get('attack_method_display', ''),
            '#supplierName#': data.get('supplier_name') or ctx.config.get('supplierName', ''),
            '#reportTime#': datetime.now().strftime("%Y-%m-%d"),
        }
        replacements = ctx.build_replacements(data, extra_replacements)

        # 3. Replace text
        ctx.replace_text(replacements)

        # 4. Process evidence images
        evidence_images = data.get('evidence_images', [])
        log_evidence_images = data.get('log_evidence', [])
        ctx.process_image_list('#evidence_images#', evidence_images)
        ctx.process_image_list('#log_evidence#', log_evidence_images)

        # 5. Save report
        unit_name = data.get('unit_name', 'Unknown')
        intrusion_type = data.get('intrusion_type_display', data.get('intrusion_type', '入侵痕迹'))
        severity_level = data.get('severity_level', '高危')
        filename = f"【入侵痕迹报告】{unit_name}存在{intrusion_type}【{severity_level}】.docx"

        output_path = ctx.build_output_path(unit_name, filename)
        final_path = ctx.save_document(doc, output_path)

        return True, final_path, "Intrusion report generated successfully"

    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, "", f"Report generation failed: {str(e)}"


# ═══════════════════════════════════════════════════════════════════
# Runtime adapter — bridges PluginRuntime descriptor protocol
# ═══════════════════════════════════════════════════════════════════

def execute(
    data: Dict[str, Any],
    output_dir: str,
    template_manager: Any,
    config: Optional[Dict[str, Any]] = None,
    template_id: str = "intrusion_report",
) -> Dict[str, Any]:
    """Descriptor execution entrypoint for plugin runtime."""
    from backend.core.generation_context import GenerationContext
    from backend.core.logger import setup_logger
    from backend.core.schema_loader import SchemaLoader

    template_dir = template_manager.get_template_dir(template_id)
    template_info = SchemaLoader.load_schema(template_dir)
    runtime_cfg = SchemaLoader.load_runtime(template_dir)

    logger = setup_logger('IntrusionReport')
    ctx = GenerationContext(template_dir, template_info, config, output_dir, logger)

    # Resolve attack method via ctx before preprocess (eliminates DbDataReader in handler)
    pre_data = dict(data or {})
    if not pre_data.get('attack_method_custom', '').strip() and pre_data.get('attack_method'):
        vuln_name = ctx.get_vulnerability_name(pre_data['attack_method'])
        if vuln_name:
            pre_data['attack_method_display'] = vuln_name

    # Preprocess
    processed = preprocess(pre_data, config or {})

    # Validate
    is_valid, errors = validate(processed, config or {}, template_info)
    if not is_valid:
        return {
            "success": False,
            "report_path": "",
            "message": "Data validation failed: " + "; ".join(errors),
            "errors": errors,
        }

    # Generate
    success, path, msg = generate(processed, ctx)

    # Postprocess
    if success:
        ctx.postprocess(
            path, processed,
            log_prefix=runtime_cfg.get('log_prefix', ''),
            log_fields=runtime_cfg.get('log_fields', []),
            db_table=runtime_cfg.get('db_table', ''),
            db_name=f"{datetime.now().strftime('%Y-%m-%d')}_output.db",
            db_field_map=runtime_cfg.get('db_fields', {}),
        )

    return {
        "success": success,
        "report_path": path,
        "message": msg,
        "errors": errors if not success else [],
    }


PLUGIN = {
    "id": "intrusion_report",
    "execute": execute,
}
