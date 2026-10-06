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


# The base exporter preserves timestamp strings without implicit YAML date types.
# Proxy notes use NoteDumper to apply their quoted JST timestamp presentation.
SourcePathDumper.yaml_implicit_resolvers = {
    key: [(tag, pattern) for tag, pattern in values if tag != "tag:yaml.org,2002:timestamp"]
    for key, values in SourcePathDumper.yaml_implicit_resolvers.items()
}
