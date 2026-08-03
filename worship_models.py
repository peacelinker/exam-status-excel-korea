"""구역예배 출결 분석에서 사용하는 데이터 모델."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from models import HeaderCheck


WORSHIP_REGIONS = ["서대문", "마포", "합정", "새신", "신촌", "홍대", "소성"]


@dataclass
class WorshipSheetCandidate:
    """A·D·H 열을 읽을 수 있는 구역예배 시트 후보."""

    name: str
    max_row: int
    name_rows: int
    attendance_rows: int
    blank_attendance_rows: int
    is_analyzable: bool
    header_checks: list[HeaderCheck] = field(default_factory=list)
    exclusion_reason: str = ""
    recommended: bool = False


@dataclass
class WorshipCounts:
    """한 지역의 구역예배 참여 방식별 실제 행 수."""

    face_to_face: int = 0
    zoom: int = 0
    phone: int = 0

    @property
    def total(self) -> int:
        return self.face_to_face + self.zoom + self.phone

    def validation_values(self) -> dict[str, int]:
        return {
            "대면": self.face_to_face,
            "줌": self.zoom,
            "전화": self.phone,
            "전체": self.total,
        }


@dataclass
class WorshipRegionResult:
    """고정 지역 한 곳의 집계값과 자동 계수 후 사용자가 확인한 재적."""

    region: str
    counts: WorshipCounts = field(default_factory=WorshipCounts)
    roster: int | None = None

    def _percent(self, count: int) -> float | None:
        if self.roster is None or self.roster == 0:
            return None
        return count / self.roster

    @property
    def face_percent(self) -> float | None:
        return self._percent(self.counts.face_to_face)

    @property
    def zoom_percent(self) -> float | None:
        return self._percent(self.counts.zoom)

    @property
    def phone_percent(self) -> float | None:
        return self._percent(self.counts.phone)

    @property
    def attendance_percent(self) -> float | None:
        return self._percent(self.counts.total)

    @property
    def absent(self) -> int | None:
        if self.roster is None:
            return None
        return self.roster - self.counts.total

    def as_row(self) -> dict[str, Any]:
        return {
            "재적": self.roster,
            "지역": self.region,
            "대면": self.counts.face_to_face,
            "퍼센트": self.face_percent,
            "줌": self.counts.zoom,
            "줌 퍼센트": self.zoom_percent,
            "전화": self.counts.phone,
            "전화 퍼센트": self.phone_percent,
            "전체": self.counts.total,
            "미참여": self.absent,
            "출결 재적대비 %": self.attendance_percent,
        }


@dataclass(frozen=True)
class WorshipExcludedRow:
    """집계하지 않은 원본 행과 그 이유."""

    sheet_name: str
    row_number: int
    region: str
    name: str
    attendance: str
    reason: str


@dataclass(frozen=True)
class WorshipValidationDifference:
    """두 독립 집계 방식 사이의 차이."""

    region: str
    metric: str
    row_iteration: int
    conditional_count: int


@dataclass
class WorshipAnalysisResult:
    """선택한 구역예배 시트의 분석·검산·내보내기 정보."""

    source_filename: str
    selected_sheet: str
    report_title: str
    candidates: list[WorshipSheetCandidate]
    header_checks: list[HeaderCheck]
    region_results: list[WorshipRegionResult]
    total_counts: WorshipCounts
    validation_differences: list[WorshipValidationDifference]
    excluded_rows: list[WorshipExcludedRow]
    header_row: int
    data_start_row: int
    last_data_row: int
    hidden_row_count: int
    filter_applied: bool
    formula_without_cached_value_count: int
    analyzed_at: datetime
    rule_version: str = "1.1"

    @property
    def validation_passed(self) -> bool:
        return not self.validation_differences

    @property
    def all_rosters_complete(self) -> bool:
        return all(item.roster is not None for item in self.region_results)

    @property
    def total_roster(self) -> int | None:
        if not self.all_rosters_complete:
            return None
        return sum(item.roster or 0 for item in self.region_results)

    @property
    def total_attendance_percent(self) -> float | None:
        roster = self.total_roster
        if roster is None or roster == 0:
            return None
        return self.total_counts.total / roster

    @property
    def total_absent(self) -> int | None:
        roster = self.total_roster
        if roster is None:
            return None
        return roster - self.total_counts.total

    def total_row(self) -> dict[str, Any]:
        roster = self.total_roster

        def percent(count: int) -> float | None:
            if roster is None or roster == 0:
                return None
            return count / roster

        return {
            "재적": roster,
            "지역": "전체",
            "대면": self.total_counts.face_to_face,
            "퍼센트": percent(self.total_counts.face_to_face),
            "줌": self.total_counts.zoom,
            "줌 퍼센트": percent(self.total_counts.zoom),
            "전화": self.total_counts.phone,
            "전화 퍼센트": percent(self.total_counts.phone),
            "전체": self.total_counts.total,
            "미참여": self.total_absent,
            "출결 재적대비 %": self.total_attendance_percent,
        }

    def aggregate_rows(self, *, include_total: bool = True) -> list[dict[str, Any]]:
        rows = [item.as_row() for item in self.region_results]
        if include_total:
            rows.append(self.total_row())
        return rows

