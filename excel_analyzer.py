"""선택한 시험 시트의 실제 셀 값을 집계하고 독립 검산한다."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable

from models import (
    AnalysisResult,
    ExcludedRow,
    HeaderCheck,
    RegionResult,
    SpecialItem,
    StatusCounts,
    UnexpectedValue,
    ValidationDifference,
)
from sheet_detector import HEADER_ROW, discover_candidate_sheets, validate_candidate_headers
from utils import AppError, compress_row_numbers, open_workbook_from_bytes

REGULAR_EXAM_VALUES = {"정규응시", "정규응시(타지파)"}
FACE_TO_FACE_VALUE = "대면응시"
ONE_TO_ONE_VALUE = "일대일응시"
WRITTEN_VALUE = "서면응시"
INFORMAL_VALUE = "비공식응시"
ABSENT_VALUE = "미응시"
ALLOWED_STATUS_VALUES = {
    *REGULAR_EXAM_VALUES,
    FACE_TO_FACE_VALUE,
    ONE_TO_ONE_VALUE,
    WRITTEN_VALUE,
    INFORMAL_VALUE,
    ABSENT_VALUE,
    "",
}
MANUAL_REGION = "대학"
RULE_VERSION = "1.0"


@dataclass(frozen=True)
class DataRow:
    """A·D·I 열에서 읽은 하나의 원본 데이터 행."""

    row_number: int
    region_raw: object
    name_raw: object
    status_raw: object
    region: str
    name: str
    status: str
    completely_blank: bool


def normalize_exact_text(value: object) -> str:
    """셀 값을 문자열로 변환하고 앞뒤 공백만 제거한다."""

    if value is None:
        return ""
    return str(value).strip()


def is_blank(value: object) -> bool:
    """None, 빈 문자열, 공백 문자만 있는 문자열인지 판정한다."""

    return normalize_exact_text(value) == ""


def _read_data_rows(worksheet, *, header_row: int) -> list[DataRow]:
    """헤더 다음 행부터 워크시트의 실제 셀 값을 수정 없이 읽는다."""

    rows: list[DataRow] = []
    max_column = max(9, worksheet.max_column)
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        row_values = [
            worksheet.cell(row_number, column).value
            for column in range(1, max_column + 1)
        ]
        region_raw = worksheet.cell(row_number, 1).value
        name_raw = worksheet.cell(row_number, 4).value
        status_raw = worksheet.cell(row_number, 9).value
        rows.append(
            DataRow(
                row_number=row_number,
                region_raw=region_raw,
                name_raw=name_raw,
                status_raw=status_raw,
                region=normalize_exact_text(region_raw),
                name=normalize_exact_text(name_raw),
                status=normalize_exact_text(status_raw),
                completely_blank=all(is_blank(value) for value in row_values),
            )
        )
    return rows


def _is_eligible(row: DataRow) -> bool:
    """지역별 자동 집계 대상이 되는 행인지 판정한다."""

    return (
        not row.completely_blank
        and bool(row.name)
        and bool(row.region)
        and row.region != MANUAL_REGION
    )


def _increment_status(counts: StatusCounts, status: str) -> None:
    """정확히 일치하는 상태 하나만 해당 집계 필드에 더한다."""

    if status == "정규응시":
        counts.regular_standard += 1
    elif status == "정규응시(타지파)":
        counts.regular_other_tribe += 1
    elif status == FACE_TO_FACE_VALUE:
        counts.face_to_face += 1
    elif status == ONE_TO_ONE_VALUE:
        counts.one_to_one += 1
    elif status == WRITTEN_VALUE:
        counts.written += 1
    elif status == INFORMAL_VALUE:
        counts.informal += 1
    elif status == ABSENT_VALUE:
        counts.explicit_absent += 1
    elif status == "":
        counts.blank_status += 1
    else:
        counts.unexpected += 1


def aggregate_by_row_iteration(rows: Iterable[DataRow]) -> dict[str, StatusCounts]:
    """방식 1: 모든 대상 행을 한 번씩 순회해 지역별 상태를 집계한다."""

    aggregates: dict[str, StatusCounts] = {}
    for row in rows:
        if not _is_eligible(row):
            continue
        counts = aggregates.setdefault(row.region, StatusCounts())
        _increment_status(counts, row.status)
    return aggregates


def aggregate_by_conditional_count(
    rows: Iterable[DataRow],
    region_order: Iterable[str] | None = None,
) -> dict[str, StatusCounts]:
    """방식 2: 지역과 각 상태 조건을 독립 필터링해 다시 카운트한다."""

    eligible_rows = [row for row in rows if _is_eligible(row)]
    regions = list(region_order or dict.fromkeys(row.region for row in eligible_rows))
    regions = [region for region in regions if region and region != MANUAL_REGION]
    aggregates: dict[str, StatusCounts] = {}
    for region in regions:
        regional_rows = [row for row in eligible_rows if row.region == region]
        statuses = [row.status for row in regional_rows]
        aggregates[region] = StatusCounts(
            regular_standard=sum(status == "정규응시" for status in statuses),
            regular_other_tribe=sum(
                status == "정규응시(타지파)" for status in statuses
            ),
            face_to_face=sum(status == FACE_TO_FACE_VALUE for status in statuses),
            one_to_one=sum(status == ONE_TO_ONE_VALUE for status in statuses),
            written=sum(status == WRITTEN_VALUE for status in statuses),
            informal=sum(status == INFORMAL_VALUE for status in statuses),
            explicit_absent=sum(status == ABSENT_VALUE for status in statuses),
            blank_status=sum(status == "" for status in statuses),
            unexpected=sum(status not in ALLOWED_STATUS_VALUES for status in statuses),
        )
    return aggregates


def _sum_counts(aggregates: dict[str, StatusCounts]) -> StatusCounts:
    """대학과 지역 공란이 이미 제외된 지역 집계를 전체 합계로 더한다."""

    total = StatusCounts()
    for counts in aggregates.values():
        total.regular_standard += counts.regular_standard
        total.regular_other_tribe += counts.regular_other_tribe
        total.face_to_face += counts.face_to_face
        total.one_to_one += counts.one_to_one
        total.written += counts.written
        total.informal += counts.informal
        total.explicit_absent += counts.explicit_absent
        total.blank_status += counts.blank_status
        total.unexpected += counts.unexpected
    return total


def validate_aggregates(
    row_iteration: dict[str, StatusCounts],
    conditional_count: dict[str, StatusCounts],
) -> list[ValidationDifference]:
    """두 집계의 지역별·항목별 값과 전체 합계를 모두 비교한다."""

    differences: list[ValidationDifference] = []
    region_names = list(dict.fromkeys([*row_iteration, *conditional_count]))
    for region in region_names:
        first = row_iteration.get(region, StatusCounts()).validation_values()
        second = conditional_count.get(region, StatusCounts()).validation_values()
        for metric in first:
            if first[metric] != second[metric]:
                differences.append(
                    ValidationDifference(region, metric, first[metric], second[metric])
                )

    first_total = _sum_counts(row_iteration).validation_values()
    second_total = _sum_counts(conditional_count).validation_values()
    for metric in first_total:
        if first_total[metric] != second_total[metric]:
            differences.append(
                ValidationDifference(
                    "전체 합계", metric, first_total[metric], second_total[metric]
                )
            )
    return differences


def build_region_results(
    aggregates: dict[str, StatusCounts],
    region_order: Iterable[str],
) -> list[RegionResult]:
    """원본 발견 순서를 유지하고 대학은 수기 입력용 공란 행으로 만든다."""

    results: list[RegionResult] = []
    seen: set[str] = set()
    for region in region_order:
        if not region or region in seen:
            continue
        seen.add(region)
        if region == MANUAL_REGION:
            results.append(RegionResult(region=region, is_manual_region=True))
        elif region in aggregates:
            results.append(RegionResult(region=region, counts=aggregates[region]))
    return results


def collect_unexpected_values(
    rows: Iterable[DataRow],
    *,
    sheet_name: str,
) -> list[UnexpectedValue]:
    """허용 목록에 없는 상태값을 원본값·지역·행번호별로 요약한다."""

    groups: dict[tuple[str, str], list[DataRow]] = defaultdict(list)
    for row in rows:
        if _is_eligible(row) and row.status not in ALLOWED_STATUS_VALUES:
            groups[(str(row.status_raw), row.status)].append(row)

    unexpected_values: list[UnexpectedValue] = []
    for (original, normalized), grouped_rows in groups.items():
        details = [
            ExcludedRow(
                sheet_name=sheet_name,
                row_number=row.row_number,
                region=row.region,
                name=row.name,
                status=original,
                reason="예상하지 못한 시험현황 값",
            )
            for row in grouped_rows
        ]
        unexpected_values.append(
            UnexpectedValue(
                original_value=original,
                normalized_value=normalized,
                count=len(grouped_rows),
                region_counts=dict(Counter(row.region for row in grouped_rows)),
                row_numbers=[row.row_number for row in grouped_rows],
                details=details,
            )
        )
    return unexpected_values


def collect_excluded_rows(
    rows: Iterable[DataRow],
    *,
    sheet_name: str,
) -> list[ExcludedRow]:
    """집계에서 제외되는 각 행을 하나의 우선 사유와 함께 기록한다."""

    excluded: list[ExcludedRow] = []
    for row in rows:
        reason = ""
        if row.completely_blank:
            reason = "완전히 빈 행"
        elif not row.name:
            reason = "이름 공란"
        elif not row.region:
            reason = "지역 공란"
        elif row.region == MANUAL_REGION:
            reason = "대학 지역"
        elif row.status not in ALLOWED_STATUS_VALUES:
            reason = "예상하지 못한 시험현황 값"
        if reason:
            excluded.append(
                ExcludedRow(
                    sheet_name=sheet_name,
                    row_number=row.row_number,
                    region=row.region,
                    name=row.name,
                    status=normalize_exact_text(row.status_raw),
                    reason=reason,
                )
            )
    return excluded


def _formula_without_cache_rows(raw_worksheet, data_worksheet, rows: list[DataRow]) -> list[int]:
    """분석 열의 수식에 저장된 계산값이 없는 행을 찾는다."""

    missing: list[int] = []
    for row in rows:
        for column in (1, 4, 9):
            raw_cell = raw_worksheet.cell(row.row_number, column)
            cached_cell = data_worksheet.cell(row.row_number, column)
            if raw_cell.data_type == "f" and cached_cell.value is None:
                missing.append(row.row_number)
                break
    return missing


def _special_items(
    *,
    rows: list[DataRow],
    excluded_rows: list[ExcludedRow],
    unexpected_values: list[UnexpectedValue],
    header_checks: list[HeaderCheck],
    candidates,
    all_sheet_names: list[str],
    selected_sheet: str,
    hidden_rows: list[int],
    filter_applied: bool,
    formula_missing_rows: list[int],
) -> list[SpecialItem]:
    """화면과 결과 파일에 기록할 분석 특이사항 목록을 만든다."""

    items: list[SpecialItem] = []
    blank_status_rows = [row.row_number for row in rows if _is_eligible(row) and not row.status]
    if blank_status_rows:
        items.append(
            SpecialItem(
                "시험현황 공란",
                "공란",
                len(blank_status_rows),
                compress_row_numbers(blank_status_rows),
                "이름은 있으나 시험현황이 공란이어서 미응시자에 포함했습니다.",
            )
        )

    for unexpected in unexpected_values:
        items.append(
            SpecialItem(
                "예상하지 못한 시험현황 값",
                unexpected.normalized_value,
                unexpected.count,
                compress_row_numbers(unexpected.row_numbers),
                "허용된 응시 유형과 미응시자 집계에서 제외했습니다.",
            )
        )

    reason_to_category = {
        "이름 공란": "이름 공란",
        "지역 공란": "지역 공란",
        "대학 지역": "대학 지역 제외",
        "완전히 빈 행": "완전히 빈 행",
    }
    for reason, category in reason_to_category.items():
        matching = [row.row_number for row in excluded_rows if row.reason == reason]
        if matching:
            description = {
                "이름 공란": "D열 이름이 공란이어서 집계에서 제외했습니다.",
                "지역 공란": "A열 지역이 공란이어서 지역별 집계에서 제외했습니다.",
                "대학 지역": "대학은 별도 시트에서 수기로 입력하므로 모든 합계에서 제외했습니다.",
                "완전히 빈 행": "실제 값이 없는 행을 집계에서 제외했습니다.",
            }[reason]
            items.append(
                SpecialItem(
                    category,
                    reason,
                    len(matching),
                    compress_row_numbers(matching),
                    description,
                )
            )

    if formula_missing_rows:
        items.append(
            SpecialItem(
                "수식 저장 결과값 없음",
                "저장된 계산값 없음",
                len(set(formula_missing_rows)),
                compress_row_numbers(formula_missing_rows),
                "data_only=True로 읽을 저장된 계산값이 없어 해당 셀을 공란으로 처리했습니다.",
            )
        )
    if hidden_rows:
        items.append(
            SpecialItem(
                "숨겨진 행 포함",
                "포함",
                len(hidden_rows),
                compress_row_numbers(hidden_rows),
                "숨김 여부와 관계없이 실제 셀 값이 있는 행을 검토했습니다.",
            )
        )
    if filter_applied:
        items.append(
            SpecialItem(
                "필터 범위 행 포함",
                "포함",
                len(rows),
                compress_row_numbers([row.row_number for row in rows]),
                "자동 필터 표시 여부와 관계없이 워크시트의 데이터 행을 검토했습니다.",
            )
        )
    for check in header_checks:
        if check.status == "경고":
            items.append(
                SpecialItem(
                    "헤더 이름 불일치",
                    f"{check.column}열: {check.actual_value}",
                    1,
                    str(HEADER_ROW),
                    check.description,
                )
            )

    candidate_names = {candidate.name for candidate in candidates if candidate.is_analyzable}
    excluded_sheets = [name for name in all_sheet_names if name not in candidate_names]
    if excluded_sheets:
        items.append(
            SpecialItem(
                "후보에서 제외된 시트",
                ", ".join(excluded_sheets),
                len(excluded_sheets),
                "",
                "시트명 규칙, 필수 열 또는 데이터 행 조건을 만족하지 않았습니다.",
            )
        )
    items.append(
        SpecialItem(
            "선택한 분석 시트",
            selected_sheet,
            1,
            "",
            "사용자가 최종 선택한 시트의 실제 셀 값을 분석했습니다.",
        )
    )
    return items


def analyze_selected_sheet(
    file_bytes: bytes,
    selected_sheet: str,
    *,
    source_filename: str = "업로드.xlsx",
    header_row: int = HEADER_ROW,
) -> AnalysisResult:
    """선택한 시트를 분석하고 두 독립 방식으로 검산한 전체 결과를 반환한다."""

    candidates = discover_candidate_sheets(file_bytes, header_row=header_row)
    if not any(candidate.is_analyzable for candidate in candidates):
        raise AppError("분석 가능한 시험 시트를 찾지 못했습니다.")

    data_workbook = open_workbook_from_bytes(file_bytes, data_only=True)
    raw_workbook = open_workbook_from_bytes(file_bytes, data_only=False)
    if selected_sheet not in data_workbook.sheetnames:
        raise AppError("선택한 시트를 찾지 못했습니다.")

    selected_candidate = next(
        (candidate for candidate in candidates if candidate.name == selected_sheet), None
    )
    if selected_candidate is None or not selected_candidate.is_analyzable:
        raise AppError("선택한 시트를 찾지 못했습니다.")

    worksheet = data_workbook[selected_sheet]
    raw_worksheet = raw_workbook[selected_sheet]
    header_checks = validate_candidate_headers(worksheet, header_row=header_row)
    if not header_checks:
        raise AppError("헤더를 확인할 수 없습니다.")
    if not all(check.readable for check in header_checks):
        raise AppError("필수 열을 읽을 수 없습니다.")

    rows = _read_data_rows(worksheet, header_row=header_row)
    if not any(row.name and not row.completely_blank for row in rows):
        raise AppError("이름이 있는 데이터 행을 찾지 못했습니다.")

    region_order = list(
        dict.fromkeys(
            row.region
            for row in rows
            if row.name and row.region and not row.completely_blank
        )
    )
    if not any(region != MANUAL_REGION for region in region_order):
        raise AppError("집계 대상 지역을 찾지 못했습니다.")

    row_iteration = aggregate_by_row_iteration(rows)
    conditional_count = aggregate_by_conditional_count(rows, region_order)
    differences = validate_aggregates(row_iteration, conditional_count)
    region_results = build_region_results(row_iteration, region_order)
    unexpected_values = collect_unexpected_values(rows, sheet_name=selected_sheet)
    excluded_rows = collect_excluded_rows(rows, sheet_name=selected_sheet)
    hidden_rows = [
        row.row_number
        for row in rows
        if worksheet.row_dimensions[row.row_number].hidden
    ]
    formula_missing_rows = _formula_without_cache_rows(raw_worksheet, worksheet, rows)
    special_items = _special_items(
        rows=rows,
        excluded_rows=excluded_rows,
        unexpected_values=unexpected_values,
        header_checks=header_checks,
        candidates=candidates,
        all_sheet_names=data_workbook.sheetnames,
        selected_sheet=selected_sheet,
        hidden_rows=hidden_rows,
        filter_applied=bool(worksheet.auto_filter.ref),
        formula_missing_rows=formula_missing_rows,
    )

    return AnalysisResult(
        source_filename=source_filename,
        selected_sheet=selected_sheet,
        candidates=candidates,
        header_checks=header_checks,
        region_results=region_results,
        status_counts_by_region=row_iteration,
        total_counts=_sum_counts(row_iteration),
        validation_differences=differences,
        unexpected_values=unexpected_values,
        special_items=special_items,
        excluded_rows=excluded_rows,
        header_row=header_row,
        data_start_row=header_row + 1,
        last_data_row=worksheet.max_row,
        hidden_row_count=len(hidden_rows),
        filter_applied=bool(worksheet.auto_filter.ref),
        formula_without_cached_value_count=len(set(formula_missing_rows)),
        analyzed_at=datetime.now().astimezone(),
        rule_version=RULE_VERSION,
    )
