"""업로드된 XLSX의 A·D·H 실제 셀 값으로 구역예배 출결을 집계한다."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterable

from models import HeaderCheck
from utils import AppError, open_workbook_from_bytes, safe_cell_text
from worship_models import (
    WORSHIP_REGIONS,
    WorshipAnalysisResult,
    WorshipCounts,
    WorshipExcludedRow,
    WorshipRegionResult,
    WorshipSheetCandidate,
    WorshipValidationDifference,
)

HEADER_ROW = 1
FACE_TO_FACE_VALUE = "대면모임"
ZOOM_VALUE = "줌"
PHONE_VALUE = "통화"
ALLOWED_VALUES = {FACE_TO_FACE_VALUE, ZOOM_VALUE, PHONE_VALUE}
HEADER_RULES: dict[str, tuple[str, ...]] = {
    "A": ("지역",),
    "D": ("이름", "성명", "이름구분"),
    "H": ("참여방법", "참여 방법", "모임방법", "모임 방법", "참석방법", "참석 방법", "출결", "구역예배"),
}


@dataclass(frozen=True)
class WorshipDataRow:
    row_number: int
    region_raw: object
    name_raw: object
    attendance_raw: object
    region: str
    name: str
    attendance: str
    completely_blank: bool


def normalize_exact_text(value: object) -> str:
    """셀 값을 문자열로 바꾸고 앞뒤 공백만 제거한다."""

    return "" if value is None else str(value).strip()


def validate_worship_headers(worksheet, *, header_row: int = HEADER_ROW) -> list[HeaderCheck]:
    """A·D·H 고정 열이 존재하는지 확인하고 헤더명 차이는 경고로 남긴다."""

    checks: list[HeaderCheck] = []
    for column, expected_values in HEADER_RULES.items():
        column_index = ord(column) - ord("A") + 1
        if worksheet.max_column < column_index:
            checks.append(
                HeaderCheck(
                    sheet_name=worksheet.title,
                    column=column,
                    expected_values=expected_values,
                    actual_value="",
                    status="오류",
                    description="필수 열 위치가 워크시트 범위를 벗어났습니다.",
                    readable=False,
                )
            )
            continue
        actual = safe_cell_text(worksheet.cell(header_row, column_index).value)
        if actual in expected_values:
            status = "일치"
            description = "기대 헤더와 일치합니다."
        elif actual:
            status = "경고"
            description = "열 위치는 읽을 수 있으나 헤더 이름이 다릅니다. 지정된 열 위치로 분석합니다."
        else:
            status = "경고"
            description = "헤더는 공란이지만 지정된 열 위치의 실제 셀 값으로 분석합니다."
        checks.append(
            HeaderCheck(
                sheet_name=worksheet.title,
                column=column,
                expected_values=expected_values,
                actual_value=actual,
                status=status,
                description=description,
                readable=True,
            )
        )
    return checks


def discover_worship_sheets(
    file_bytes: bytes,
    *,
    header_row: int = HEADER_ROW,
) -> list[WorshipSheetCandidate]:
    """워크북의 모든 시트를 검사해 A·D·H 기반 분석 후보를 반환한다."""

    workbook = open_workbook_from_bytes(file_bytes, data_only=True)
    candidates: list[WorshipSheetCandidate] = []
    for worksheet in workbook.worksheets:
        checks = validate_worship_headers(worksheet, header_row=header_row)
        name_rows = attendance_rows = blank_attendance_rows = 0
        for row_number in range(header_row + 1, worksheet.max_row + 1):
            name = safe_cell_text(worksheet.cell(row_number, 4).value)
            if not name:
                continue
            name_rows += 1
            attendance = safe_cell_text(worksheet.cell(row_number, 8).value)
            if attendance:
                attendance_rows += 1
            else:
                blank_attendance_rows += 1
        reasons: list[str] = []
        if not all(check.readable for check in checks):
            reasons.append("A·D·H 필수 열을 읽을 수 없음")
        if worksheet.max_row <= header_row:
            reasons.append("데이터 행이 없음")
        if name_rows == 0:
            reasons.append("D열 이름이 있는 데이터 행이 없음")
        candidates.append(
            WorshipSheetCandidate(
                name=worksheet.title,
                max_row=worksheet.max_row,
                name_rows=name_rows,
                attendance_rows=attendance_rows,
                blank_attendance_rows=blank_attendance_rows,
                is_analyzable=not reasons,
                header_checks=checks,
                exclusion_reason="; ".join(reasons),
            )
        )
    analyzable = [item for item in candidates if item.is_analyzable]
    if analyzable:
        preferred = [item for item in analyzable if item.attendance_rows]
        (preferred or analyzable)[-1].recommended = True
    return candidates


def _read_rows(worksheet, *, header_row: int) -> list[WorshipDataRow]:
    rows: list[WorshipDataRow] = []
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        values = [
            worksheet.cell(row_number, column).value
            for column in range(1, max(8, worksheet.max_column) + 1)
        ]
        region_raw = worksheet.cell(row_number, 1).value
        name_raw = worksheet.cell(row_number, 4).value
        attendance_raw = worksheet.cell(row_number, 8).value
        rows.append(
            WorshipDataRow(
                row_number=row_number,
                region_raw=region_raw,
                name_raw=name_raw,
                attendance_raw=attendance_raw,
                region=normalize_exact_text(region_raw),
                name=normalize_exact_text(name_raw),
                attendance=normalize_exact_text(attendance_raw),
                completely_blank=all(not normalize_exact_text(value) for value in values),
            )
        )
    return rows


def _is_countable(row: WorshipDataRow) -> bool:
    return (
        not row.completely_blank
        and bool(row.name)
        and row.region in WORSHIP_REGIONS
        and row.attendance in ALLOWED_VALUES
    )


def _increment(counts: WorshipCounts, attendance: str) -> None:
    if attendance == FACE_TO_FACE_VALUE:
        counts.face_to_face += 1
    elif attendance == ZOOM_VALUE:
        counts.zoom += 1
    elif attendance == PHONE_VALUE:
        counts.phone += 1


def aggregate_by_row_iteration(rows: Iterable[WorshipDataRow]) -> dict[str, WorshipCounts]:
    aggregates = {region: WorshipCounts() for region in WORSHIP_REGIONS}
    for row in rows:
        if _is_countable(row):
            _increment(aggregates[row.region], row.attendance)
    return aggregates


def aggregate_by_conditional_count(rows: Iterable[WorshipDataRow]) -> dict[str, WorshipCounts]:
    source = list(rows)
    aggregates: dict[str, WorshipCounts] = {}
    for region in WORSHIP_REGIONS:
        regional = [row for row in source if row.name and row.region == region]
        aggregates[region] = WorshipCounts(
            face_to_face=sum(row.attendance == FACE_TO_FACE_VALUE for row in regional),
            zoom=sum(row.attendance == ZOOM_VALUE for row in regional),
            phone=sum(row.attendance == PHONE_VALUE for row in regional),
        )
    return aggregates


def _validate(
    first: dict[str, WorshipCounts],
    second: dict[str, WorshipCounts],
) -> list[WorshipValidationDifference]:
    differences: list[WorshipValidationDifference] = []
    for region in WORSHIP_REGIONS:
        left = first[region].validation_values()
        right = second[region].validation_values()
        for metric in left:
            if left[metric] != right[metric]:
                differences.append(
                    WorshipValidationDifference(region, metric, left[metric], right[metric])
                )
    return differences


def _sum_counts(aggregates: dict[str, WorshipCounts]) -> WorshipCounts:
    return WorshipCounts(
        face_to_face=sum(item.face_to_face for item in aggregates.values()),
        zoom=sum(item.zoom for item in aggregates.values()),
        phone=sum(item.phone for item in aggregates.values()),
    )


def _excluded_rows(rows: Iterable[WorshipDataRow], sheet_name: str) -> list[WorshipExcludedRow]:
    excluded: list[WorshipExcludedRow] = []
    for row in rows:
        reason = ""
        if row.completely_blank:
            reason = "완전히 빈 행"
        elif not row.name:
            reason = "D열 이름 공란"
        elif not row.region:
            reason = "A열 지역 공란"
        elif row.region not in WORSHIP_REGIONS:
            reason = "분석 대상 외 지역"
        elif not row.attendance:
            reason = "H열 참여방식 공란"
        elif row.attendance not in ALLOWED_VALUES:
            reason = "예상하지 못한 H열 값"
        if reason:
            excluded.append(
                WorshipExcludedRow(
                    sheet_name=sheet_name,
                    row_number=row.row_number,
                    region=row.region,
                    name=row.name,
                    attendance=row.attendance,
                    reason=reason,
                )
            )
    return excluded


def _normalize_rosters(rosters: dict[str, int | None] | None) -> dict[str, int | None]:
    rosters = rosters or {}
    normalized: dict[str, int | None] = {}
    for region in WORSHIP_REGIONS:
        value = rosters.get(region)
        if value is None:
            normalized[region] = None
            continue
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise AppError(f"{region} 재적은 0 이상의 정수로 입력해 주세요.")
        normalized[region] = value
    return normalized


def _formula_without_cache_rows(raw_worksheet, data_worksheet, rows: list[WorshipDataRow]) -> list[int]:
    missing: list[int] = []
    for row in rows:
        for column in (1, 4, 8):
            raw_cell = raw_worksheet.cell(row.row_number, column)
            cached_cell = data_worksheet.cell(row.row_number, column)
            if raw_cell.data_type == "f" and cached_cell.value is None:
                missing.append(row.row_number)
                break
    return missing


def analyze_worship_sheet(
    file_bytes: bytes,
    selected_sheet: str,
    *,
    rosters: dict[str, int | None] | None = None,
    report_title: str = "구역예배 성인",
    source_filename: str = "업로드.xlsx",
    header_row: int = HEADER_ROW,
) -> WorshipAnalysisResult:
    """선택 시트의 실제 셀 값을 두 방식으로 집계하고 재적을 결합한다."""

    candidates = discover_worship_sheets(file_bytes, header_row=header_row)
    candidate = next((item for item in candidates if item.name == selected_sheet), None)
    if candidate is None or not candidate.is_analyzable:
        raise AppError("분석 가능한 구역예배 시트를 선택해 주세요.")

    data_workbook = open_workbook_from_bytes(file_bytes, data_only=True)
    raw_workbook = open_workbook_from_bytes(file_bytes, data_only=False)
    worksheet = data_workbook[selected_sheet]
    raw_worksheet = raw_workbook[selected_sheet]
    rows = _read_rows(worksheet, header_row=header_row)
    if not any(row.name and not row.completely_blank for row in rows):
        raise AppError("D열 이름이 있는 데이터 행을 찾지 못했습니다.")

    first = aggregate_by_row_iteration(rows)
    second = aggregate_by_conditional_count(rows)
    normalized_rosters = _normalize_rosters(rosters)
    region_results = [
        WorshipRegionResult(region, first[region], normalized_rosters[region])
        for region in WORSHIP_REGIONS
    ]
    hidden_rows = [
        row.row_number
        for row in rows
        if worksheet.row_dimensions[row.row_number].hidden
    ]
    formula_missing = _formula_without_cache_rows(raw_worksheet, worksheet, rows)
    title = report_title.strip() or "구역예배 성인"

    return WorshipAnalysisResult(
        source_filename=source_filename,
        selected_sheet=selected_sheet,
        report_title=title,
        candidates=candidates,
        header_checks=validate_worship_headers(worksheet, header_row=header_row),
        region_results=region_results,
        total_counts=_sum_counts(first),
        validation_differences=_validate(first, second),
        excluded_rows=_excluded_rows(rows, selected_sheet),
        header_row=header_row,
        data_start_row=header_row + 1,
        last_data_row=worksheet.max_row,
        hidden_row_count=len(hidden_rows),
        filter_applied=bool(worksheet.auto_filter.ref),
        formula_without_cached_value_count=len(set(formula_missing)),
        analyzed_at=datetime.now().astimezone(),
    )

