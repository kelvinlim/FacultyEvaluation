#! /usr/bin/env python3

"""Validate criteria_tables.yaml structure without generating outputs."""

from criteria_data import load_tables


def main() -> None:
    tables = load_tables()
    mission_count = len(tables)
    group_count = sum(len(data["groups"]) for data in tables.values())
    row_count = sum(len(rows) for data in tables.values() for _, rows in data["groups"])

    print(f"Criteria data valid: {mission_count} missions, {group_count} groups, {row_count} rows")


if __name__ == "__main__":
    main()
