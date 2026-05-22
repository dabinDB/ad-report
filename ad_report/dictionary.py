from __future__ import annotations

import re
from pathlib import Path
from typing import Any


def load_yaml(path: str | Path) -> dict[str, Any]:
    import yaml

    with Path(path).open("r", encoding="utf-8") as file:
        return yaml.safe_load(file) or {}


def normalize_token(value: Any) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"[\s_\-()/\[\]{}:.,%]+", "", text).lower()


class StandardDictionary:
    def __init__(self, config: dict[str, Any]):
        self.config = config
        self.time_dimensions = config.get("time_dimensions", {})
        self.categorical_dimensions = config.get("categorical_dimensions", {})
        self.metrics = config.get("metrics", {})
        self.dimensions = {**self.time_dimensions, **self.categorical_dimensions}
        self._synonym_index = self._build_synonym_index()

    @property
    def dimension_names(self) -> list[str]:
        return list(self.dimensions.keys())

    @property
    def metric_names(self) -> list[str]:
        return list(self.metrics.keys())

    def _build_synonym_index(self) -> dict[str, str]:
        index: dict[str, str] = {}
        for name, spec in {**self.dimensions, **self.metrics}.items():
            candidates = [name, *spec.get("synonyms", [])]
            for candidate in candidates:
                index[normalize_token(candidate)] = name
        return index

    def match(self, value: Any) -> str | None:
        token = normalize_token(value)
        if not token:
            return None
        if token in self._synonym_index:
            return self._synonym_index[token]
        for synonym, standard_name in self._synonym_index.items():
            if synonym and (synonym in token or token in synonym):
                return standard_name
        return None

    def is_dimension(self, name: str) -> bool:
        return name in self.dimensions

    def is_metric(self, name: str) -> bool:
        return name in self.metrics

    def metric_spec(self, name: str) -> dict[str, Any]:
        return self.metrics.get(name, {})

    def derived_dimensions(self) -> dict[str, dict[str, Any]]:
        return {
            name: spec
            for name, spec in self.time_dimensions.items()
            if spec.get("derive_from")
        }
