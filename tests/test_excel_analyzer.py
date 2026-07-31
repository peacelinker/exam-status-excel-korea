"""상태 정규화, 지역별 집계, 제외, 두 방식 검산 테스트."""

from __future__ import annotations

from copy import deepcopy

import pytest

from excel_analyzer import (
    DataRow,
    aggregate_by_conditional_count,
    aggregate_by_row_iteration,
    analyze_selected_sheet,
    build_region_results,
    is_blank,
    normalize_exact_text,
    validate_aggregates,
)
from models import StatusCounts
from utils import AppError, compress_row_numbers


@pytest.mark.parametrize("value", [None, "", " ", "\t\n"])
def test_blank_variants(value):
    assert is_blank(value)


def test_normalize_exact_text_trims_outer_spaces():
    assert normalize_exact_text("  정규응시  ") == "정규응시"


def test_normalize_exact_text_preserves_internal_spaces():
    assert normalize_exact_text(" 정규 응시 ") == "정규 응시"


def test_normalize_numeric_name_is_not_blank():
    assert normalize_exact_text(1234) == "1234"


def test_region_order_follows_first_appearance(sample_result):
    assert [item.region for item in sample_result.region_results] == ["서대문", "마포", "대학", "신촌"]


def test_regular_standard_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].regular_standard == 1


def test_regular_other_tribe_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].regular_other_tribe == 1


def test_regular_total_is_sum_of_two_exact_values(sample_result):
    assert sample_result.status_counts_by_region["서대문"].regular_total == 2


def test_face_to_face_is_separate_from_regular(sample_result):
    counts = sample_result.status_counts_by_region["서대문"]
    assert counts.face_to_face == 1
    assert counts.regular_total == 2


def test_one_to_one_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].one_to_one == 1


def test_written_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].written == 1


def test_informal_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].informal == 1


def test_total_exam_sums_five_types_once(sample_result):
    assert sample_result.status_counts_by_region["서대문"].total_exam == 6


def test_explicit_absent_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].explicit_absent == 1


def test_blank_status_count(sample_result):
    assert sample_result.status_counts_by_region["서대문"].blank_status == 1


def test_final_absent_is_explicit_plus_blank(sample_result):
    counts = sample_result.status_counts_by_region["서대문"]
    assert counts.absent_total == counts.explicit_absent + counts.blank_status == 2


def test_unexpected_value_is_not_added_to_exam_or_absent(sample_result):
    counts = sample_result.status_counts_by_region["마포"]
    assert counts.unexpected == 1
    assert counts.total_exam == 4
    assert counts.absent_total == 1


def test_internal_space_value_remains_unexpected(sample_result):
    assert any(item.normalized_value == "정규 응시" for item in sample_result.unexpected_values)


def test_outer_spaces_are_removed_before_counting(sample_result):
    assert sample_result.status_counts_by_region["마포"].regular_standard == 1


def test_blank_name_row_is_excluded(sample_result):
    assert any(row.reason == "이름 공란" and row.row_number == 12 for row in sample_result.excluded_rows)


def test_blank_region_row_is_excluded(sample_result):
    assert any(row.reason == "지역 공란" and row.row_number == 13 for row in sample_result.excluded_rows)


def test_manual_university_region_is_present(sample_result):
    university = next(item for item in sample_result.region_results if item.region == "대학")
    assert university.is_manual_region


def test_manual_university_numeric_cells_are_none(sample_result):
    university = next(item for item in sample_result.region_results if item.region == "대학")
    assert all(value is None for key, value in university.as_row().items() if key != "지역")


def test_manual_university_is_excluded_from_totals(sample_result):
    assert sample_result.total_counts.regular_total == 5


def test_target_column_is_blank_for_all_regions(sample_result):
    assert all(row["전체 응시목표"] is None for row in sample_result.aggregate_rows())


def test_ratio_column_is_blank_for_all_regions(sample_result):
    assert all(row["전도 재적대비 %"] is None for row in sample_result.aggregate_rows())


def test_hidden_row_is_included(sample_result):
    assert sample_result.hidden_row_count == 1
    assert sample_result.status_counts_by_region["신촌"].regular_standard == 1


def test_filter_range_rows_are_included(sample_result):
    assert sample_result.filter_applied
    assert sample_result.total_counts.total_exam == 11


def test_formula_without_cached_value_warns(sample_result):
    assert sample_result.formula_without_cached_value_count == 2
    assert any(item.category == "수식 저장 결과값 없음" for item in sample_result.special_items)


def test_numeric_name_row_is_included(sample_result):
    assert sample_result.status_counts_by_region["마포"].written == 1


def test_same_row_is_not_counted_twice(sample_result):
    counts = sample_result.total_counts
    assert counts.total_exam == (
        counts.regular_total
        + counts.face_to_face
        + counts.one_to_one
        + counts.written
        + counts.informal
    )


def test_row_iteration_and_conditional_count_match():
    rows = [
        DataRow(2, "A", "N1", "정규응시", "A", "N1", "정규응시", False),
        DataRow(3, "A", "N2", None, "A", "N2", "", False),
        DataRow(4, "B", "N3", "기타", "B", "N3", "기타", False),
    ]
    first = aggregate_by_row_iteration(rows)
    second = aggregate_by_conditional_count(rows, ["A", "B"])
    assert validate_aggregates(first, second) == []


def test_intentional_validation_mismatch_is_reported():
    first = {"A": StatusCounts(regular_standard=1)}
    second = {"A": StatusCounts(regular_standard=2)}
    differences = validate_aggregates(first, second)
    assert any(item.region == "A" and item.metric == "정규응시" for item in differences)
    assert any(item.region == "전체 합계" for item in differences)


def test_build_region_results_keeps_manual_placeholder():
    results = build_region_results({"서대문": StatusCounts(regular_standard=1)}, ["서대문", "대학"])
    assert results[1].region == "대학" and results[1].is_manual_region


def test_compress_row_numbers():
    assert compress_row_numbers([4, 5, 6, 8, 10, 11, 12]) == "4-6, 8, 10-12"


def test_unexpected_details_contain_sheet_and_row(sample_result):
    detail = sample_result.unexpected_values[0].details[0]
    assert detail.sheet_name == "0719"
    assert detail.row_number == 11


def test_completely_blank_row_is_recorded(sample_result):
    assert any(row.reason == "완전히 빈 행" for row in sample_result.excluded_rows)


def test_selected_sheet_missing_raises_korean_error(sample_workbook_bytes):
    with pytest.raises(AppError, match="선택한 시트"):
        analyze_selected_sheet(sample_workbook_bytes, "9999")


def test_no_candidate_sheet_raises_korean_error(workbook_factory):
    data = workbook_factory([("일반시트", [("서대문", "익명", "정규응시")], None)])
    with pytest.raises(AppError, match="분석 가능한"):
        analyze_selected_sheet(data, "일반시트")


def test_only_university_has_no_aggregate_region(workbook_factory):
    data = workbook_factory([("0719", [("대학", "익명", "정규응시")], None)])
    with pytest.raises(AppError, match="집계 대상 지역"):
        analyze_selected_sheet(data, "0719")


def test_header_name_mismatch_is_retained_as_warning(workbook_factory):
    data = workbook_factory([("0719", [("서대문", "익명", "정규응시")], ("구역", "대상자", "응시상태"))])
    result = analyze_selected_sheet(data, "0719")
    assert all(check.status == "경고" for check in result.header_checks)
    assert any(item.category == "헤더 이름 불일치" for item in result.special_items)


def test_result_validation_property_is_true(sample_result):
    assert sample_result.validation_passed


def test_deepcopy_does_not_change_sample_result(sample_result):
    copied = deepcopy(sample_result)
    copied.total_counts.regular_standard += 1
    assert sample_result.total_counts.regular_standard == 3
