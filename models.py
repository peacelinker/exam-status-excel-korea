"""시험 응시 현황 집계에서 사용하는 데이터 모델."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


@dataclass(frozen=True)
class HeaderCheck:
    """필수 열의 기대 헤더와 실제 헤더를 비교한 결과."""

    sheet_name: str
    column: str
    expected_values: tuple[str, ...]
    actual_value: str
    status: str
    description: str
    readable: bool = True

    @property
    def expected_label(self) -> str:
        """화면과 보고서에 표시할 기대 헤더 문자열을 반환한다."""

        return ", ".join(self.expected_values)


@dataclass
class SheetCandidate:
    """자동 탐색된 시험 시트 후보의 구조 요약."""

    name: str
    candidate_type: str
    max_row: int
    name_rows: int
    status_rows: int
    blank_status_rows: int
    is_analyzable: bool
    header_checks: list[HeaderCheck] = field(default_factory=list)
    exclusion_reason: str = ""
    recommended: bool = False


@dataclass
class StatusCounts:
    """하나의 지역에서 발견된 시험현황 상태별 건수."""

    regular_standard: int = 0
    regular_other_tribe: int = 0
    face_to_face: int = 0
    one_to_one: int = 0
    written: int = 0
    informal: int = 0
    explicit_absent: int = 0
    blank_status: int = 0
    unexpected: int = 0

    @property
    def regular_total(self) -> int:
        """정규응시와 정규응시(타지파)의 합계를 반환한다."""

        return self.regular_standard + self.regular_other_tribe

    @property
    def total_exam(self) -> int:
        """다섯 응시 유형의 전체응시 합계를 반환한다."""

        return (
            self.regular_total
            + self.face_to_face
            + self.one_to_one
            + self.written
            + self.informal
        )

    @property
    def absent_total(self) -> int:
        """명시적 미응시와 시험현황 공란의 합계를 반환한다."""

        return self.explicit_absent + self.blank_status

    def validation_values(self) -> dict[str, int]:
        """독립 검산에 사용하는 모든 파생값을 사전으로 반환한다."""

        return {
            "정규응시": self.regular_total,
            "정규응시(일반)": self.regular_standard,
            "정규응시(타지파)": self.regular_other_tribe,
            "대면응시": self.face_to_face,
            "일대일응시": self.one_to_one,
            "서면응시": self.written,
            "비공식응시": self.informal,
            "전체응시": self.total_exam,
            "명시적 미응시": self.explicit_absent,
            "시험현황 공란": self.blank_status,
            "최종 미응시자": self.absent_total,
            "예상하지 못한 값": self.unexpected,
        }


@dataclass
class RegionResult:
    """화면과 집계결과 시트에 표시할 지역별 결과."""

    region: str
    counts: StatusCounts = field(default_factory=StatusCounts)
    is_manual_region: bool = False

    def as_row(self) -> dict[str, Any]:
        """사용자 표의 고정 컬럼 순서에 맞는 행을 반환한다."""

        if self.is_manual_region:
            return {
                "지역": self.region,
                "전체 응시목표": None,
                "정규응시": None,
                "대면응시": None,
                "일대일응시": None,
                "서면응시": None,
                "비공식응시": None,
                "전체응시": None,
                "미응시자": None,
                "전도 재적대비 %": None,
            }
        return {
            "지역": self.region,
            "전체 응시목표": None,
            "정규응시": self.counts.regular_total,
            "대면응시": self.counts.face_to_face,
            "일대일응시": self.counts.one_to_one,
            "서면응시": self.counts.written,
            "비공식응시": self.counts.informal,
            "전체응시": self.counts.total_exam,
            "미응시자": self.counts.absent_total,
            "전도 재적대비 %": None,
        }


@dataclass(frozen=True)
class ValidationDifference:
    """두 독립 집계 방식 사이에서 발견된 차이."""

    region: str
    metric: str
    row_iteration: int
    conditional_count: int


@dataclass(frozen=True)
class SpecialItem:
    """분석 과정에서 사용자 확인이 필요한 특이사항."""

    category: str
    value: str
    count: int
    row_range: str
    description: str


@dataclass(frozen=True)
class ExcludedRow:
    """집계에서 제외되었으며 감사 목적으로 기록하는 행."""

    sheet_name: str
    row_number: int
    region: str
    name: str
    status: str
    reason: str


@dataclass
class UnexpectedValue:
    """허용 목록에 없는 시험현황 값의 요약과 상세."""

    original_value: str
    normalized_value: str
    count: int
    region_counts: dict[str, int]
    row_numbers: list[int]
    details: list[ExcludedRow]


@dataclass
class AnalysisResult:
    """한 시트에 대한 전체 분석·검산·내보내기 정보."""

    source_filename: str
    selected_sheet: str
    candidates: list[SheetCandidate]
    header_checks: list[HeaderCheck]
    region_results: list[RegionResult]
    status_counts_by_region: dict[str, StatusCounts]
    total_counts: StatusCounts
    validation_differences: list[ValidationDifference]
    unexpected_values: list[UnexpectedValue]
    special_items: list[SpecialItem]
    excluded_rows: list[ExcludedRow]
    header_row: int
    data_start_row: int
    last_data_row: int
    hidden_row_count: int
    filter_applied: bool
    formula_without_cached_value_count: int
    analyzed_at: datetime
    rule_version: str = "1.0"

    @property
    def validation_passed(self) -> bool:
        """검산 차이가 하나도 없으면 참을 반환한다."""

        return not self.validation_differences

    @property
    def candidate_names(self) -> list[str]:
        """분석정보 기록용 후보 시트명 목록을 반환한다."""

        return [candidate.name for candidate in self.candidates if candidate.is_analyzable]

    def aggregate_rows(self, include_total: bool = True) -> list[dict[str, Any]]:
        """지역별 결과와 선택적인 합계 행을 표 행으로 반환한다."""

        rows = [result.as_row() for result in self.region_results]
        if include_total:
            rows.append(
                {
                    "지역": "합계",
                    "전체 응시목표": None,
                    "정규응시": self.total_counts.regular_total,
                    "대면응시": self.total_counts.face_to_face,
                    "일대일응시": self.total_counts.one_to_one,
                    "서면응시": self.total_counts.written,
                    "비공식응시": self.total_counts.informal,
                    "전체응시": self.total_counts.total_exam,
                    "미응시자": self.total_counts.absent_total,
                    "전도 재적대비 %": None,
                }
            )
        return rows
