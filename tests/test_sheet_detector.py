"""후보 시트 탐색과 헤더 검증 테스트."""

from __future__ import annotations

import pytest

from sheet_detector import discover_candidate_sheets, recommend_latest_sheet
from utils import InvalidExcelError


def _candidate(candidates, name):
    return next(item for item in candidates if item.name == name)


def test_discovers_region_total_sheet(workbook_factory):
    data = workbook_factory([("지역전체", [("서대문", "익명", "정규응시")], None)])
    assert _candidate(discover_candidate_sheets(data), "지역전체").is_analyzable


def test_discovers_previous_exam_sheet(workbook_factory):
    data = workbook_factory([("직전시험", [("서대문", "익명", "정규응시")], None)])
    assert _candidate(discover_candidate_sheets(data), "직전시험").candidate_type == "직전시험"


def test_discovers_four_digit_sheet(workbook_factory):
    data = workbook_factory([("1207", [("서대문", "익명", "정규응시")], None)])
    assert _candidate(discover_candidate_sheets(data), "1207").candidate_type == "날짜형 숫자 시트"


@pytest.mark.parametrize("name", ["지역전체 사본", "지역전체 복사본", "직전시험 사본"])
def test_excludes_copy_sheet_names(workbook_factory, name):
    data = workbook_factory([(name, [("서대문", "익명", "정규응시")], None)])
    assert discover_candidate_sheets(data) == []


@pytest.mark.parametrize("name", ["0719사본", "07190", "7-19", "시험0719"])
def test_excludes_non_exact_four_digit_names(workbook_factory, name):
    data = workbook_factory([(name, [("서대문", "익명", "정규응시")], None)])
    assert discover_candidate_sheets(data) == []


def test_missing_required_header_makes_candidate_unavailable(workbook_factory):
    data = workbook_factory([("0719", [("서대문", "익명", "정규응시")], ("지역", "이름", None))])
    candidate = _candidate(discover_candidate_sheets(data), "0719")
    assert not candidate.is_analyzable
    assert "필수 헤더" in candidate.exclusion_reason


def test_different_nonblank_header_warns_but_allows_analysis(workbook_factory):
    data = workbook_factory([("0719", [("서대문", "익명", "정규응시")], ("구역", "대상자", "응시상태"))])
    candidate = _candidate(discover_candidate_sheets(data), "0719")
    assert candidate.is_analyzable
    assert all(check.status == "경고" for check in candidate.header_checks)


def test_sheet_without_name_rows_is_excluded(workbook_factory):
    data = workbook_factory([("0719", [("서대문", None, "정규응시")], None)])
    candidate = _candidate(discover_candidate_sheets(data), "0719")
    assert not candidate.is_analyzable
    assert candidate.name_rows == 0


def test_multiple_candidates_remain_selectable(sample_workbook_bytes):
    candidates = [item for item in discover_candidate_sheets(sample_workbook_bytes) if item.is_analyzable]
    assert {item.name for item in candidates} >= {"지역전체", "직전시험", "0719", "0726"}


def test_recommends_latest_valid_mmdd_sheet(sample_workbook_bytes):
    recommendation = recommend_latest_sheet(discover_candidate_sheets(sample_workbook_bytes))
    assert recommendation is not None
    assert recommendation.name == "0726"


def test_candidate_statistics_count_status_blanks(sample_workbook_bytes):
    candidate = _candidate(discover_candidate_sheets(sample_workbook_bytes), "0719")
    assert candidate.name_rows > candidate.status_rows
    assert candidate.blank_status_rows >= 1


def test_corrupted_excel_raises_korean_error():
    with pytest.raises(InvalidExcelError, match="손상"):
        discover_candidate_sheets(b"not-an-xlsx")
