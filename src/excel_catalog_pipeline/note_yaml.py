"""YAML presentation shared by proxy-note writers."""

from typing import Any

import yaml


class SourcePathDumper(yaml.SafeDumper):
    """Keep sourceFullPath strings single-quoted, including embedded apostrophes."""

    def represent_mapping(
        self, tag: str, mapping: Any, flow_style: bool | None = None
    ) -> yaml.MappingNode:
        node = super().represent_mapping(tag, mapping, flow_style)
        for key, value in node.value:
            if (
                key.value == "sourceFullPath"
                and isinstance(value, yaml.ScalarNode)
                and value.tag == "tag:yaml.org,2002:str"
            ):
                value.style = "'"
        return node
