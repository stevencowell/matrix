import argparse
import json
import pathlib
import re
from collections import defaultdict

STANDARD_PERIOD_MINUTES = 59
WEDNESDAY_PERIOD_MINUTES = 38
WEDNESDAY_LINES = {
    "weda1",
    "weda5",
    "weda6",
    "wedb1",
    "wedb5",
    "wedb6",
}

YEAR_PERIOD_ALLOCATION = {
    12: 7,
    11: 7,
    10: 5,
    9: 6,
}
YEAR8_ELECTIVE_PERIODS = 3
YEAR8_TECH_MANDATORY_PERIODS = 5
YEAR7_TECH_MANDATORY_PERIODS = 5

FTE_BASE_PERIODS_MAPPING = {
    "1.0": 37.95,
    "0.8": 30.35,
    "0.6": 22.77,
    "0.4": 15.17,
    "0.2": 11.8,
}

ASSEMBLY_FULL_PERIOD_EQUIVALENT = 0.62
ASSEMBLY_FULL_MINUTES = STANDARD_PERIOD_MINUTES * ASSEMBLY_FULL_PERIOD_EQUIVALENT
ASSEMBLY_SHORT_MINUTES = 19


def is_wednesday_line(label: str) -> bool:
    return label and label.strip().lower() in WEDNESDAY_LINES


def is_wednesday_subject(code: str) -> bool:
    return bool(code and "_wed" in code.lower())


def normalize_subject_code_for_periods(subject_code: str) -> str:
    if not subject_code:
        return ""
    normalized = str(subject_code).upper()
    normalized = re.sub(r"\s+", " ", normalized).strip()
    normalized = re.sub(r"^S[12]\s*:?\s*", "", normalized)
    normalized = re.sub(r"^(?:YEAR|YR)\s*", "", normalized)
    return normalized


def build_split_lookup(subject_splits):
    lookup = {}
    base_map = defaultdict(list)
    for base, entries in subject_splits.items():
        if not isinstance(entries, list):
            continue
        for entry in entries:
            code = entry.get("code")
            periods = entry.get("periods")
            if not code:
                continue
            lookup[code] = {
                "base": base,
                "periods": periods,
                "index": entry.get("index"),
                "total": entry.get("totalSplits"),
            }
            base_map[base].append(code)
    return lookup, base_map


def parse_period_value(value):
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return 0.0


def get_subject_period_value(subject_code, split_lookup):
    if not subject_code:
        return 0.0

    if is_wednesday_subject(subject_code):
        return 1.0

    split_info = split_lookup.get(subject_code)
    if split_info:
        return parse_period_value(split_info.get("periods"))

    normalized = normalize_subject_code_for_periods(subject_code)
    if not normalized:
        return 0.0

    match = re.match(r"^(\d{1,2})", normalized)
    if not match:
        return 0.0

    year = int(match.group(1))
    if year in YEAR_PERIOD_ALLOCATION:
        return float(YEAR_PERIOD_ALLOCATION[year])

    remainder = normalized[len(match.group(1)) :].strip()
    compact = re.sub(r"\s+", "", remainder)

    if year == 8:
        tech_keyword = bool(re.search(r"\bTECH\b", remainder))
        mandatory_keyword = bool(re.search(r"(^|\s)MAND(?:ATORY)?\b", remainder))
        is_tech_mandatory = (
            compact.startswith("TM") or tech_keyword or mandatory_keyword or "TECHMAND" in compact
        )
        return float(YEAR8_TECH_MANDATORY_PERIODS if is_tech_mandatory else YEAR8_ELECTIVE_PERIODS)

    if year == 7:
        return float(YEAR7_TECH_MANDATORY_PERIODS)

    return 0.0


def compute_line_minutes(lines):
    minutes = []
    for label in lines:
        minutes.append(
            WEDNESDAY_PERIOD_MINUTES if is_wednesday_line(label) else STANDARD_PERIOD_MINUTES
        )
    return minutes


def build_allocation_records(data, split_lookup):
    records = []
    teachers = data["teachers"]
    lines = data["lines"]
    line_minutes = compute_line_minutes(lines)

    for key, subjects in data["allocations"].items():
        line_idx_str, teacher_idx_str = key.split("-")
        line_idx = int(line_idx_str)
        teacher_idx = int(teacher_idx_str)
        line_label = lines[line_idx]
        line_min = line_minutes[line_idx]
        subject_list = subjects if isinstance(subjects, list) else [subjects]
        for subject in subject_list:
            period_value = get_subject_period_value(subject, split_lookup)
            minutes = period_value * line_min
            records.append(
                {
                    "teacher_index": teacher_idx,
                    "teacher": teachers[teacher_idx],
                    "line_index": line_idx,
                    "line_label": line_label,
                    "subject": subject,
                    "period_value": period_value,
                    "line_minutes": line_min,
                    "minutes": minutes,
                }
            )
    return records


def summarize_teacher(records, teacher_name):
    teacher_records = [rec for rec in records if rec["teacher"] == teacher_name]
    total_minutes = sum(rec["minutes"] for rec in teacher_records)
    return teacher_records, total_minutes


def compute_allowance_minutes(settings):
    period_allowance = settings.get("periodAllowance", 0) * STANDARD_PERIOD_MINUTES
    assembly_full = settings.get("assemblyFullCount", 0) * ASSEMBLY_FULL_MINUTES
    assembly_short = settings.get("assemblyShortCount", 0) * ASSEMBLY_SHORT_MINUTES
    additional = settings.get("additionalMinutes", 0)
    return period_allowance + assembly_full + assembly_short + additional


def compute_base_minutes(fte_value):
    normalized = str(fte_value).strip()
    base_periods = FTE_BASE_PERIODS_MAPPING.get(normalized)
    if base_periods is None:
        try:
            numeric = float(normalized)
            normalized_numeric = f"{numeric:.1f}"
            base_periods = FTE_BASE_PERIODS_MAPPING.get(
                normalized_numeric, FTE_BASE_PERIODS_MAPPING["1.0"]
            )
        except ValueError:
            base_periods = FTE_BASE_PERIODS_MAPPING["1.0"]
    return base_periods * STANDARD_PERIOD_MINUTES


def format_teacher_summary(name, records, load_settings):
    teacher_records, teaching_minutes = summarize_teacher(records, name)

    load_settings = load_settings or {}
    allowance_minutes = compute_allowance_minutes(load_settings)
    total_load_minutes = teaching_minutes + allowance_minutes
    base_minutes = compute_base_minutes(load_settings.get("fte", "1.0"))
    balance_minutes = base_minutes - total_load_minutes

    period_allowance = load_settings.get("periodAllowance", 0) * STANDARD_PERIOD_MINUTES
    assembly_full = load_settings.get("assemblyFullCount", 0) * ASSEMBLY_FULL_MINUTES
    assembly_short = load_settings.get("assemblyShortCount", 0) * ASSEMBLY_SHORT_MINUTES
    additional = load_settings.get("additionalMinutes", 0)

    lines = [
        f"### {name}",
        "",
        "| Line | Subject | Periods | Line Minutes | Load (min) |",
        "| - | - | - | - | - |",
    ]
    for rec in teacher_records:
        subject_display = rec["subject"].replace("▸", "->")
        lines.append(
            f"| {rec['line_label']} | {subject_display} | "
            f"{rec['period_value']:.2f} | {rec['line_minutes']:.0f} | {rec['minutes']:.2f} |"
        )
    if not teacher_records:
        lines.append("| — | — | — | — | — |")

    lines.extend(
        [
            "",
            "**Teaching minutes:** {:.2f}".format(teaching_minutes),
            "**Allowances:**",
            f"- Period allowance: {period_allowance:.2f}",
            f"- Full assemblies: {assembly_full:.2f}",
            f"- TLC / short assemblies: {assembly_short:.2f}",
            f"- Additional minutes: {additional:.2f}",
            "**Total allowance minutes:** {:.2f}".format(allowance_minutes),
            "",
            "**Total load:** {:.2f}".format(total_load_minutes),
            "**Capacity (base minutes):** {:.2f}".format(base_minutes),
            "**Balance (capacity - load):** {:.2f}".format(balance_minutes),
            "",
            "---",
            "",
        ]
    )

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Verify teacher load calculations.")
    parser.add_argument(
        "--data",
        default=r"C:\Users\scowell1\Downloads\matrix-main (10)\matrix-main\faculty-allocations with Wed Classes.json",
        help="Path to the allocation JSON file.",
    )
    parser.add_argument(
        "--teacher",
        help="Optional teacher name to filter. If omitted, all teachers are processed.",
    )
    parser.add_argument(
        "--output",
        default="teacher_load_report.md",
        help="Output file for the formatted report (markdown).",
    )
    args = parser.parse_args()

    data_path = pathlib.Path(args.data)
    data = json.loads(data_path.read_text(encoding="utf-8"))

    split_lookup, _ = build_split_lookup(data.get("subjectSplits", {}))
    allocation_records = build_allocation_records(data, split_lookup)

    report_sections = []
    teachers = data.get("teachers", [])
    load_settings_map = data.get("teacherLoadSettings", {})

    target_teachers = [args.teacher] if args.teacher else teachers

    for teacher in target_teachers:
        summary = format_teacher_summary(
            teacher,
            allocation_records,
            load_settings_map.get(teacher, {}),
        )
        report_sections.append(summary)

    report_text = "\n".join(report_sections).strip() + "\n"
    pathlib.Path(args.output).write_text(report_text, encoding="utf-8")
    print(f"Wrote report to {args.output}")


if __name__ == "__main__":
    main()
