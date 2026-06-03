# -*- coding: utf-8 -*-
"""
@Createtime: 2026-01-24
@Updatetime: 2026-05-29
@description: 漏洞报告处理器 - pure function interface with GenerationContext injection.

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

ALERT_LEVEL_MAP = {
    "高危": "2级",
    "中危": "3级",
    "低危": "4级",
    "信息性": "5级",
}

# ═══════════════════════════════════════════════════════════════════
# Pure functions — template business logic
# ═══════════════════════════════════════════════════════════════════

def preprocess(data: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Data preprocessing.

    1. Generate vulnerability ID (if empty)
    2. Calculate alert level
    3. Set default discovery date
    4. Set supplier defaults
    5. Set city/region defaults
    6. Build report name
    7. Combine vuln description + harm
    8. Set report time
    """
    processed = data.copy()

    # 1. Generate vulnerability ID
    if not processed.get('vulnerability_id'):
        processed['vulnerability_id'] = gen_report_id(prefix="YHBH", use_sequence=True)

    # 2. Calculate alert level
    hazard_level = processed.get('hazard_level', '高危')
    processed['alert_level'] = ALERT_LEVEL_MAP.get(hazard_level, '2级')

    # 3. Set discovery date
    set_default_dates(processed, ['discovery_date'])

    # 4. Set supplier defaults
    set_supplier_defaults(processed, config)

    # 5. Set city/region defaults
    if not processed.get('city'):
        processed['city'] = config.get('city', '北京')
    if not processed.get('region'):
        processed['region'] = config.get('region', '海淀区')

    # 6. Build report name
    unit_name = processed.get('unit_name', '')
    website_name = processed.get('website_name', '')
    vul_name = processed.get('vul_name', '')
    report_name = f"{unit_name}{website_name}存在{vul_name}漏洞".replace("漏洞漏洞", "漏洞")
    processed['report_name'] = report_name

    # 7. Combine vuln description + harm
    vul_description = processed.get('vul_description', '')
    vul_harm = processed.get('vul_harm', '')
    if vul_harm:
        processed['vul_description_full'] = f"{vul_description}{vul_harm}"
    else:
        processed['vul_description_full'] = vul_description

    # 8. Set report time
    processed['report_time'] = datetime.now().strftime("%Y-%m-%d")

    return processed


def validate(
    data: Dict[str, Any],
    config: Dict[str, Any],
    template_info: Any,
) -> Tuple[bool, List[str]]:
    """
    Data validation — vulnerability report specific checks.
    """
    errors = []

    # Check required fields from template info
    if template_info:
        for field_def in template_info.fields:
            if field_def.required:
                value = data.get(field_def.key, "")
                if not value or (isinstance(value, str) and not value.strip()):
                    errors.append(f"Field '{field_def.label}' is required")

        # Check global validation rules
        for rule in template_info.validation_rules:
            if rule.rule == 'required':
                for field_key in rule.fields:
                    value = data.get(field_key, "")
                    if not value or (isinstance(value, str) and not value.strip()):
                        errors.append(rule.message)
                        break

    # Custom validation
    if not data.get('vul_name') and not data.get('vul_name_select'):
        errors.append("Please select or enter vulnerability name")

    return len(errors) == 0, errors


def generate(data: Dict[str, Any], ctx: Any) -> Tuple[bool, str, str]:
    """
    Generate vulnerability report.

    Args:
        data: Preprocessed form data
        ctx: GenerationContext providing all framework services

    Returns:
        (success, output_path, message)
    """
    try:
        # 1. Load document
        doc = ctx.load_document()
        if not doc:
            return False, "", "Template file loading failed"

        # 2. Build replacements
        extra_replacements = {
            "#supplierName#": data.get('supplier_name', ctx.config.get('supplierName', '')),
            "#reportTime#": data.get('report_time', datetime.now().strftime("%Y-%m-%d")),
            "#reportName#": data.get('report_name', ''),
            "#vulDescription#": data.get('vul_description_full', data.get('vul_description', '')),
            "#customerCompanyName#": data.get('unit_name', ''),
            "#target#": data.get('url', ''),
        }
        replacements = ctx.build_replacements(data, extra_replacements)

        # 3. Replace text
        ctx.replace_text(replacements)

        # 4. Process images
        ctx.process_single_image('#screenshotoffiling#', data.get('icp_screenshot'))
        ctx.process_image_list('#evidenceScreenshot#', data.get('vuln_evidence_images', []))

        # 5. Save report
        region = data.get('region', '')
        hazard_type = data.get('hazard_type', '漏洞报告')
        report_name = data.get('report_name', 'Report')
        hazard_level = data.get('hazard_level', '高危')
        filename = f"【{region}】【{hazard_type}】{report_name}【{hazard_level}】.docx"

        unit_name = data.get('unit_name', 'Unknown')
        output_path = ctx.build_output_path(unit_name, filename)
        final_path = ctx.save_document(doc, output_path)

        return True, final_path, "Report generated successfully"

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
    template_id: str = "vuln_report",
) -> Dict[str, Any]:
    """Descriptor execution entrypoint for plugin runtime."""
    from backend.core.generation_context import GenerationContext
    from backend.core.logger import setup_logger
    from backend.core.schema_loader import SchemaLoader

    template_dir = template_manager.get_template_dir(template_id)
    template_info = SchemaLoader.load_schema(template_dir)
    runtime_cfg = SchemaLoader.load_runtime(template_dir)

    logger = setup_logger('VulnReport')
    ctx = GenerationContext(template_dir, template_info, config, output_dir, logger)

    # Preprocess
    processed = preprocess(data, config or {})

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
    "id": "vuln_report",
    "execute": execute,
}
