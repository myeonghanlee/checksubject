#!/usr/bin/env python3
"""Build a searchable JSON catalog from the supplied reference Markdown tables."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data" / "고등학교교과및과목검증기준.md"
OUTPUT = ROOT / "data" / "reference_catalog.json"


def clean_course(value: str) -> str:
    value = value.strip()
    # The user instructed that terminal asterisk footnote markers are not part of the name.
    return re.sub(r"\*{1,2}$", "", value).rstrip()


def cell_items(value: str) -> list[str]:
    if not value or value.strip() in {"—", "-"}:
        return []
    values = []
    for item in re.split(r"<br\s*/?>", value, flags=re.IGNORECASE):
        course = clean_course(item)
        if course and course not in {"—", "-"}:
            values.append(course)
    return values


def cells(line: str) -> list[str]:
    return [part.strip() for part in line.strip().strip("|").split("|")]


def table_lines(lines: list[str], start: int, end: int) -> list[list[str]]:
    return [cells(line) for line in lines[start:end] if line.lstrip().startswith("|") and not re.match(r"\|\s*:?-{2,}", line)]


def main() -> None:
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    catalog: list[dict[str, str]] = []

    # Table 1: general education courses by subject group and course-selection type.
    t1_start = next(i for i, line in enumerate(lines) if line.startswith("| 교과(군) | 공통 과목 |"))
    t1_end = next(i for i in range(t1_start + 1, len(lines)) if lines[i].startswith("## 표 2"))
    for row in table_lines(lines, t1_start, t1_end):
        if len(row) < 5 or row[0] == "교과(군)":
            continue
        group = row[0]
        for col_idx, course_type in zip(range(1, 5), ["공통 과목", "일반 선택", "진로 선택", "융합 선택"]):
            for course in cell_items(row[col_idx]):
                catalog.append({"name": course, "table": "표 1", "classification": "보통교과", "group": group, "category": course_type, "context": ""})

    # Table 2: special-purpose high school elective courses by stream/group.
    t2_start = next(i for i, line in enumerate(lines) if line.startswith("| 계열 | 교과(군) | 특목고 선택 과목 |"))
    t2_end = next(i for i in range(t2_start + 1, len(lines)) if lines[i].startswith("## 표 3"))
    for row in table_lines(lines, t2_start, t2_end):
        if len(row) < 3 or row[0] == "계열":
            continue
        stream, group = row[0], row[1]
        for course in cell_items(row[2]):
            catalog.append({"name": course, "table": "표 2", "classification": "특목고 선택 과목", "group": group, "category": "특목고 선택 과목", "context": stream})

    # Table 3: vocational/professional courses, including course segment.
    t3_start = next(i for i, line in enumerate(lines) if line.startswith("| 교과(군) | 전문 공통 |"))
    t3_end = next((i for i in range(t3_start + 1, len(lines)) if lines[i].startswith("## ")), len(lines))
    course_categories = [(1, "전문 공통"), (2, "전공 일반"), (3, "전공 실무")]
    for row in table_lines(lines, t3_start, t3_end):
        if len(row) < 4 or row[0] == "교과(군)":
            continue
        group = row[0]
        for col_idx, course_type in course_categories:
            for course in cell_items(row[col_idx]):
                catalog.append({"name": course, "table": "표 3", "classification": "전문 교과", "group": group, "category": course_type, "context": ""})

    OUTPUT.write_text(json.dumps(catalog, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    by_table = {table: sum(1 for entry in catalog if entry["table"] == table) for table in ("표 1", "표 2", "표 3")}
    print(f"Wrote {len(catalog)} catalog entries to {OUTPUT}")
    print("Entries by table:", by_table)
    print("Unique normalized names:", len({entry['name'] for entry in catalog}))


if __name__ == "__main__":
    main()
