# -*- coding: utf-8 -*-
"""
SchemaLoader - Shared YAML -> Pydantic parsing utility.

Extracts schema loading logic from template_manager.py into a reusable module
used by both TemplateManager and BaseTemplateHandler.

Based on _load_template() logic from template_manager.py (lines 319-458).
"""

import os
from typing import Any, Dict, Optional

import yaml

from .logger import setup_logger
from .schema_models import (
    Behavior,
    BehaviorAction,
    DataSourceDef,
    FieldDefinition,
    FieldGroup,
    TemplateInfo,
    ValidationRule,
)

logger = setup_logger("SchemaLoader")


class SchemaLoader:
    """Reusable YAML -> Pydantic parsing for template schemas and runtime configs."""

    @staticmethod
    def load_schema(template_dir: str) -> TemplateInfo:
        """
        Read schema.yaml from template_dir, parse to TemplateInfo (Pydantic model).

        Args:
            template_dir: Absolute path to template directory (e.g. backend/templates/my_template)

        Returns:
            TemplateInfo Pydantic model with all fields, groups, behaviors, etc.

        Raises:
            FileNotFoundError: If schema.yaml does not exist
            yaml.YAMLError: If schema.yaml is malformed
            ValueError: If schema data is invalid
        """
        schema_path = os.path.join(template_dir, "schema.yaml")

        if not os.path.exists(schema_path):
            raise FileNotFoundError(f"schema.yaml not found in {template_dir}")

        with open(schema_path, "r", encoding="utf-8") as f:
            schema = yaml.safe_load(f)

        # Parse field groups
        field_groups = []
        for idx, group_data in enumerate(schema.get("field_groups", [])):
            group_data.setdefault("id", f"group_{idx}")
            group_data.setdefault("order", idx)
            try:
                group = FieldGroup(**group_data)
                field_groups.append(group)
            except Exception as e:
                logger.warning(f"Invalid field group in {template_dir}: {e}")
        field_groups.sort(key=lambda x: x.order)

        # Parse data sources
        data_sources = []
        for ds_data in schema.get("data_sources", []):
            try:
                ds = DataSourceDef(**ds_data)
                data_sources.append(ds)
            except Exception as e:
                logger.warning(f"Invalid data source in {template_dir}: {e}")

        # Parse field definitions
        fields = []
        for idx, field_data in enumerate(schema.get("fields", [])):
            field_data.setdefault("order", idx)
            try:
                field_def = FieldDefinition(**field_data)
                fields.append(field_def)
            except Exception as e:
                logger.warning(
                    f"Invalid field '{field_data.get('key', 'unknown')}' in {template_dir}: {e}"
                )
        fields.sort(key=lambda x: x.order)

        # Parse behaviors
        behaviors = []
        for beh_data in schema.get("behaviors", []):
            trigger = beh_data.get("trigger", {})
            actions = []
            for act_data in beh_data.get("actions", []):
                try:
                    action = BehaviorAction(**act_data)
                    actions.append(action)
                except Exception as e:
                    logger.warning(f"Invalid action in {template_dir}: {e}")

            try:
                behavior = Behavior(
                    id=beh_data.get("id", ""),
                    trigger_field=trigger.get("field", ""),
                    trigger_event=trigger.get("event", "change"),
                    actions=actions,
                )
                behaviors.append(behavior)
            except Exception as e:
                logger.warning(f"Invalid behavior in {template_dir}: {e}")

        # Parse validation rules
        validation_rules = []
        validation_data = schema.get("validation", {})
        for rule_data in validation_data.get("rules", []):
            try:
                rule = ValidationRule(**rule_data)
                validation_rules.append(rule)
            except Exception as e:
                logger.warning(f"Invalid validation rule in {template_dir}: {e}")

        template_id = schema.get("id", os.path.basename(template_dir))

        # Create TemplateInfo
        template_info = TemplateInfo(
            id=template_id,
            name=schema.get("name", template_id),
            description=schema.get("description", ""),
            version=schema.get("version", "1.0.0"),
            order=schema.get("order", 999),
            template_file=schema.get("template_file", "template.docx"),
            icon=schema.get("icon", ""),
            author=schema.get("author", ""),
            create_time=schema.get("create_time", ""),
            update_time=schema.get("update_time", ""),
            fields=fields,
            field_groups=field_groups,
            data_sources=data_sources,
            behaviors=behaviors,
            validation_rules=validation_rules,
            output_config=schema.get("output", {}),
            preview_config=schema.get("preview", {}),
            dependent_fields=schema.get("dependent_fields", {}),
            summary_configs=schema.get("summary_configs", {}),
        )

        return template_info

    @staticmethod
    def load_runtime(template_dir: str) -> Dict[str, Any]:
        """
        Read runtime.yaml from template_dir, return dict with handler config.

        Args:
            template_dir: Absolute path to template directory

        Returns:
            Dict with keys: log_prefix, log_fields, db_table, db_fields

        Raises:
            FileNotFoundError: If runtime.yaml does not exist
        """
        runtime_path = os.path.join(template_dir, "runtime.yaml")

        if not os.path.exists(runtime_path):
            raise FileNotFoundError(
                f"runtime.yaml not found in {template_dir}. "
                f"Template must have its own runtime.yaml with handler configuration."
            )

        with open(runtime_path, "r", encoding="utf-8") as f:
            runtime = yaml.safe_load(f) or {}

        return runtime

    @staticmethod
    def get_template_path(template_dir: str, template_info: TemplateInfo) -> str:
        """
        Resolve the template.docx file path.

        Args:
            template_dir: Absolute path to template directory
            template_info: TemplateInfo object with template_file field

        Returns:
            Absolute path to template.docx, or raises FileNotFoundError
        """
        template_file = os.path.join(template_dir, template_info.template_file)

        if not os.path.exists(template_file):
            raise FileNotFoundError(
                f"Template file '{template_info.template_file}' not found in {template_dir}"
            )

        return template_file
