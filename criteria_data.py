#! /usr/bin/env python3
from functools import lru_cache
from pathlib import Path

import yaml

DATA_FILE = Path(__file__).with_name("criteria_tables.yaml")


def _error(message: str) -> ValueError:
    return ValueError(f"Invalid criteria data in {DATA_FILE.name}: {message}")


@lru_cache(maxsize=1)
def load_tables() -> dict:
    raw = yaml.safe_load(DATA_FILE.read_text(encoding="utf-8"))

    if not isinstance(raw, dict) or not raw:
        raise _error("top-level structure must be a non-empty mapping")

    tables = {}

    for mission_name, mission_data in raw.items():
        if not isinstance(mission_name, str) or not mission_name.strip():
            raise _error("mission names must be non-empty strings")
        if not isinstance(mission_data, dict):
            raise _error(f"'{mission_name}' must map to an object")

        note = mission_data.get("note")
        groups = mission_data.get("groups")

        if not isinstance(note, str) or not note.strip():
            raise _error(f"'{mission_name}.note' must be a non-empty string")
        if not isinstance(groups, list) or not groups:
            raise _error(f"'{mission_name}.groups' must be a non-empty list")

        normalized_groups = []
        for group_idx, group in enumerate(groups):
            path = f"{mission_name}.groups[{group_idx}]"
            if not isinstance(group, dict):
                raise _error(f"'{path}' must be an object")

            label = group.get("label")
            rows = group.get("rows")
            if not isinstance(label, str) or not label.strip():
                raise _error(f"'{path}.label' must be a non-empty string")
            if not isinstance(rows, list) or not rows:
                raise _error(f"'{path}.rows' must be a non-empty list")

            normalized_rows = []
            for row_idx, row in enumerate(rows):
                row_path = f"{path}.rows[{row_idx}]"
                if not isinstance(row, dict):
                    raise _error(f"'{row_path}' must be an object")

                criterion = row.get("criterion")
                levels = row.get("levels")
                is_promotion = row.get("promotion_relevant")

                if not isinstance(criterion, str) or not criterion.strip():
                    raise _error(f"'{row_path}.criterion' must be a non-empty string")
                if not isinstance(levels, list) or len(levels) != 4 or not all(isinstance(x, str) for x in levels):
                    raise _error(f"'{row_path}.levels' must be a list of exactly 4 strings")
                if not isinstance(is_promotion, bool):
                    raise _error(f"'{row_path}.promotion_relevant' must be a boolean")

                normalized_rows.append((criterion, levels, is_promotion))

            normalized_groups.append((label, normalized_rows))

        tables[mission_name] = {
            "note": note,
            "groups": normalized_groups,
        }

    return tables
