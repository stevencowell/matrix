"""Convert teacher allocation CSV files to the JSON structure used by the Matrix app."""

import argparse
import csv
import json
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple, Union

TAG_PATTERN = re.compile(r"^(?:[A-Za-z]{2,}|S\d+)\s*:\s*")
YEAR_PATTERN = re.compile(r"Year\s*(\d{1,2})", re.IGNORECASE)
ROW_PATTERN = re.compile(r"Row\s*(\d+)", re.IGNORECASE)
LINE_PATTERN = re.compile(r"Line\s*(\d+)", re.IGNORECASE)


@dataclass
class TeacherRow:
    name: str
    allowance_periods: float
    line_values: List[str]


def read_csv_rows(path: Path) -> List[List[str]]:
    with path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.reader(handle)
        return [row for row in reader]


def strip_tags(value: str) -> str:
    text = value.strip().replace("\u2013", "-").replace("\u2014", "-")
    while True:
        match = TAG_PATTERN.match(text)
        if not match:
            break
        text = text[match.end():].lstrip()
    return re.sub(r"\s+", " ", text).strip()


def split_subject_codes(value: str) -> List[str]:
    if not value or value.strip() == "":
        return []
    cleaned = strip_tags(value)
    if not cleaned:
        return []
    parts = re.split(r"\s*[/|]\s*", cleaned)
    codes: List[str] = []
    for part in parts:
        code = strip_tags(part).upper()
        if code:
            codes.append(code)
    return codes


def find_teacher_section(rows: Sequence[Sequence[str]]) -> Optional[int]:
    for index, row in enumerate(rows):
        if row and any("teacher matrix" in (cell or "").strip().lower() for cell in row):
            return index
    return None


def parse_year_sections(rows: Sequence[Sequence[str]], teacher_start: int) -> Tuple[Dict[str, List[Set[str]]], List[str]]:
    year_subjects: Dict[str, List[Set[str]]] = {}
    current_year: Optional[str] = None
    line_names: List[str] = []
    line_start_index: Optional[int] = None

    for row in rows[:teacher_start]:
        stripped = [cell.strip() for cell in row]
        if not any(stripped):
            continue

        year_cell = next((cell for cell in stripped if YEAR_PATTERN.search(cell)), None)
        if year_cell:
            match = YEAR_PATTERN.search(year_cell)
            if match:
                current_year = f"Year {int(match.group(1))}"
            else:
                current_year = year_cell.strip()
            line_names = []
            line_start_index = None
            continue

        if current_year and not line_names:
            line_candidates: List[Tuple[int, str]] = [
                (idx, cell.strip())
                for idx, cell in enumerate(row)
                if cell and LINE_PATTERN.match(cell.strip())
            ]
            if line_candidates:
                line_start_index = line_candidates[0][0]
                line_names = [name for _, name in line_candidates]
                year_subjects[current_year] = [set() for _ in line_names]
            continue

        if current_year and line_names and line_start_index is not None:
            if line_start_index - 1 >= 0:
                row_label = row[line_start_index - 1].strip()
            else:
                row_label = ""
            if not ROW_PATTERN.match(row_label or ""):
                continue

            for idx, _ in enumerate(line_names):
                column = line_start_index + idx
                cell = row[column] if column < len(row) else ""
                for code in split_subject_codes(cell):
                    year_subjects[current_year][idx].add(code)

    return year_subjects, line_names


def parse_teacher_rows(rows: Sequence[Sequence[str]], header_index: int) -> Tuple[List[str], List[str], List[TeacherRow]]:
    header_row: Sequence[str] = rows[header_index]
    line_columns: List[int] = []
    line_names: List[str] = []

    for idx, cell in enumerate(header_row):
        if cell and LINE_PATTERN.match(cell.strip()):
            line_columns.append(idx)
            line_names.append(cell.strip())

    teachers: List[str] = []
    parsed_rows: List[TeacherRow] = []

    for row in rows[header_index + 1:]:
        if not row or all((cell or "").strip() == "" for cell in row):
            continue

        name = (row[0] or "").strip()
        if not name:
            continue

        teachers.append(name)
        allowance_raw = row[2] if len(row) > 2 else ""
        allowance_value = extract_number(allowance_raw)
        line_values: List[str] = []
        for column in line_columns:
            cell = row[column] if column < len(row) else ""
            line_values.append(cell)
        parsed_rows.append(TeacherRow(name=name, allowance_periods=allowance_value, line_values=line_values))

    return teachers, line_names, parsed_rows


def extract_number(value: str) -> float:
    if not value:
        return 0.0
    cleaned = strip_tags(value).replace(",", "").strip()
    if not cleaned:
        return 0.0
    match = re.search(r"[-+]?\d+(?:\.\d+)?", cleaned)
    if not match:
        return 0.0
    try:
        return float(match.group(0))
    except ValueError:
        return 0.0


def build_subject_mappings(
    subjects_by_year: Dict[str, List[Set[str]]]
) -> Tuple[Dict[str, str], Dict[str, int]]:
    year_mapping: Dict[str, str] = {}
    line_mapping: Dict[str, int] = {}
    for year, line_sets in subjects_by_year.items():
        for index, codes in enumerate(line_sets):
            for code in codes:
                if not code:
                    continue
                year_mapping.setdefault(code, year)
                line_mapping.setdefault(code, index)
    return year_mapping, line_mapping


def infer_year_from_code(code: str) -> Optional[str]:
    match = re.match(r"(\d{1,2})", code)
    if not match:
        return None
    year_number = int(match.group(1))
    if year_number <= 0:
        return None
    return f"Year {year_number}"


def create_teacher_load_settings(rows: Sequence[TeacherRow]) -> Dict[str, Dict[str, Union[float, str]]]:
    settings: Dict[str, Dict[str, Union[float, str]]] = {}
    for row in rows:
        settings[row.name] = {
            "fte": "1.0",
            "periodAllowance": max(0.0, row.allowance_periods),
            "assemblyFullCount": 0,
            "assemblyShortCount": 0,
            "additionalMinutes": 0,
        }
    return settings


def build_allocations(
    teacher_rows: Sequence[TeacherRow],
) -> Tuple[Dict[str, Union[List[str], str]], Set[str], Dict[str, int]]:
    allocations: Dict[str, Union[List[str], str]] = {}
    subjects: Set[str] = set()
    subject_line_mapping: Dict[str, int] = {}

    for teacher_index, row in enumerate(teacher_rows):
        for line_index, cell in enumerate(row.line_values):
            codes = split_subject_codes(cell)
            if not codes:
                continue
            key = f"{line_index}-{teacher_index}"
            subjects.update(codes)
            for code in codes:
                if code not in subject_line_mapping:
                    subject_line_mapping[code] = line_index
            allocations[key] = codes if len(codes) > 1 else codes[0]

    return allocations, subjects, subject_line_mapping


def merge_subject_sets(*sets: Iterable[Set[str]]) -> Set[str]:
    combined: Set[str] = set()
    for subject_set in sets:
        combined.update(subject_set)
    return combined


def flatten_year_subjects(subjects_by_year: Dict[str, List[Set[str]]]) -> Set[str]:
    combined: Set[str] = set()
    for line_sets in subjects_by_year.values():
        for codes in line_sets:
            combined.update(codes)
    return combined


def convert_csv_to_payload(csv_path: Path) -> Dict[str, object]:
    rows = read_csv_rows(csv_path)
    teacher_section = find_teacher_section(rows)
    if teacher_section is None:
        raise ValueError("Could not locate the Teacher Matrix section in the CSV file.")

    subjects_by_year, matrix_line_names = parse_year_sections(rows, teacher_section)

    header_index = None
    for offset in range(teacher_section, len(rows)):
        row = rows[offset]
        if row and row[0].strip().lower() == "teacher":
            header_index = offset
            break
    if header_index is None:
        raise ValueError("Teacher header row not found after the Teacher Matrix marker.")

    teachers, teacher_line_names, teacher_rows = parse_teacher_rows(rows, header_index)
    if not teachers:
        raise ValueError("No teacher rows found in the CSV file.")

    allocations, allocation_subjects, allocation_line_mapping = build_allocations(teacher_rows)

    matrix_subjects = flatten_year_subjects(subjects_by_year)
    subject_year_mapping, matrix_line_mapping = build_subject_mappings(subjects_by_year)
    all_subjects = merge_subject_sets(allocation_subjects, matrix_subjects)

    for code in allocation_subjects:
        if code not in subject_year_mapping:
            inferred = infer_year_from_code(code)
            if inferred:
                subject_year_mapping[code] = inferred

    for code, line_index in matrix_line_mapping.items():
        allocation_line_mapping.setdefault(code, line_index)

    teacher_load_settings = create_teacher_load_settings(teacher_rows)

    lines = teacher_line_names if teacher_line_names else matrix_line_names

    payload: Dict[str, object] = {
        "allocations": allocations,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "subjects": sorted(all_subjects),
        "teachers": list(teachers),
        "lines": list(lines),
        "csvData": rows,
        "subjectLineMapping": allocation_line_mapping,
        "subjectSplits": {},
        "subjectYearMapping": subject_year_mapping,
        "teacherLoadSettings": teacher_load_settings,
    }

    return payload


def default_output_path(input_path: Path) -> Path:
    today = datetime.now().strftime("%Y-%m-%d")
    return input_path.with_name(f"faculty-allocations-{today}.json")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a timetable matrix CSV into the Matrix JSON format.")
    parser.add_argument("csv_file", type=Path, help="Path to the CSV file to convert.")
    parser.add_argument("-o", "--output", type=Path, help="Path for the generated JSON file.")
    parser.add_argument("--pretty", action="store_true", help="Pretty-print the JSON output.")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    csv_path: Path = args.csv_file
    if not csv_path.exists():
        raise SystemExit(f"CSV file not found: {csv_path}")

    payload = convert_csv_to_payload(csv_path)

    output_path = args.output or default_output_path(csv_path)
    json_kwargs = {"ensure_ascii": False}
    if args.pretty:
        json_kwargs["indent"] = 2
        json_kwargs["sort_keys"] = False
    else:
        json_kwargs["separators"] = (",", ":")

    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, **json_kwargs)
        handle.write("\n")

    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
