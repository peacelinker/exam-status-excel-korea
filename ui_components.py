"""Streamlit 화면을 구성하는 재사용 가능한 한국어 UI 컴포넌트."""

from __future__ import annotations

from html import escape
from pathlib import Path

import streamlit as st

from models import AnalysisResult, HeaderCheck, SheetCandidate
from utils import compress_row_numbers


def load_css() -> None:
    """assets/styles.css를 읽어 Streamlit 화면에 주입한다."""

    css_path = Path(__file__).parent / "assets" / "styles.css"
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


def render_header() -> None:
    """앱 제목, 설명, 개인정보 처리 안내가 있는 히어로를 표시한다."""

    st.markdown(
        """
        <header class="app-hero">
          <div class="eyebrow">EXCEL WORKFLOW · 실제 셀 데이터 기반</div>
          <h1>시험·구역예배 엑셀 자동 집계기</h1>
          <p class="hero-description">엑셀의 실제 셀 데이터를 기준으로 시험 응시 또는 구역예배 출결을 지역별로 집계하고 결과 엑셀 파일을 생성합니다.</p>
          <div class="privacy-note">원본 파일은 수정하지 않으며, 업로드된 바이트는 현재 앱 실행 메모리에서만 처리합니다.</div>
        </header>
        """,
        unsafe_allow_html=True,
    )


def render_progress(active_step: int) -> None:
    """업로드·선택·검산·다운로드의 진행 상태를 표시한다."""

    labels = ["파일 업로드", "시트 선택", "집계·검산", "결과 다운로드"]
    items = []
    for index, label in enumerate(labels, start=1):
        state = "done" if index < active_step else "active" if index == active_step else "waiting"
        mark = "✓" if state == "done" else str(index)
        items.append(
            f'<div class="progress-item {state}"><span>{mark}</span><strong>{escape(label)}</strong></div>'
        )
    st.markdown(f'<div class="progress-grid">{"".join(items)}</div>', unsafe_allow_html=True)


def render_step_card(number: int, title: str, description: str) -> None:
    """각 단계의 번호, 제목, 설명을 카드형 섹션 헤더로 표시한다."""

    st.markdown(
        f"""
        <div class="step-heading">
          <span class="step-number">{number}</span>
          <div><h2>{escape(title)}</h2><p>{escape(description)}</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sheet_candidates(candidates: list[SheetCandidate]) -> None:
    """후보 시트별 구조 검사 결과를 표로 표시한다."""

    rows = []
    for candidate in candidates:
        header_status = ", ".join(
            f"{check.column}:{check.status}" for check in candidate.header_checks
        )
        rows.append(
            {
                "시트명": candidate.name,
                "후보 유형": candidate.candidate_type,
                "전체 사용 행": candidate.max_row,
                "이름 있음": candidate.name_rows,
                "시험현황 있음": candidate.status_rows,
                "시험현황 공란": candidate.blank_status_rows,
                "분석 가능": "가능" if candidate.is_analyzable else "제외",
                "헤더 확인": header_status,
                "제외 사유": candidate.exclusion_reason,
                "추천": "추천" if candidate.recommended else "",
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)


def render_header_checks(checks: list[HeaderCheck]) -> None:
    """선택한 시트의 A·D·I 필수 열 확인 결과를 표시한다."""

    rows = [
        {
            "시트명": check.sheet_name,
            "열": check.column,
            "기대값": check.expected_label,
            "실제값": check.actual_value,
            "판정": check.status,
            "설명": check.description,
        }
        for check in checks
    ]
    st.dataframe(rows, use_container_width=True, hide_index=True)


def _kpi(label: str, value: str, detail: str, tone: str = "blue") -> str:
    """KPI 카드 하나의 안전한 HTML을 반환한다."""

    return (
        f'<div class="kpi-card {tone}"><span>{escape(label)}</span>'
        f'<strong>{escape(value)}</strong><small>{escape(detail)}</small></div>'
    )


def render_summary_cards(result: AnalysisResult) -> None:
    """전체응시, 미응시자, 지역 수, 예상 밖 값 KPI를 표시한다."""

    ordinary_regions = sum(not item.is_manual_region for item in result.region_results)
    cards = [
        _kpi("전체응시", f"{result.total_counts.total_exam:,}명", "5개 응시 유형 합계", "blue"),
        _kpi("최종 미응시자", f"{result.total_counts.absent_total:,}명", "명시적 미응시 + 상태 공란", "orange"),
        _kpi("자동 집계 지역", f"{ordinary_regions:,}개", "대학·지역 공란 제외", "navy"),
        _kpi("예상 밖 상태", f"{result.total_counts.unexpected:,}건", "자동 집계에서 제외", "red"),
    ]
    st.markdown(f'<div class="kpi-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def render_region_results(result: AnalysisResult) -> None:
    """대학 공란 행과 합계 행을 포함한 지역별 결과 표를 표시한다."""

    st.dataframe(result.aggregate_rows(include_total=True), use_container_width=True, hide_index=True)
    st.caption("전체 응시목표와 전도 재적대비 %는 자동 계산하지 않으며 실제 빈 칸으로 유지합니다.")


def render_validation_result(result: AnalysisResult) -> None:
    """두 독립 집계 방식의 일치 여부와 차이 상세를 표시한다."""

    if result.validation_passed:
        st.success("두 가지 독립 집계 방식의 결과가 모두 일치합니다.", icon="✅")
        return
    st.error("두 가지 검산 결과가 일치하지 않습니다. 결과 파일을 생성하지 않았습니다.", icon="🚫")
    st.dataframe(
        [
            {
                "지역": difference.region,
                "항목": difference.metric,
                "방식 1: 행 순회": difference.row_iteration,
                "방식 2: 조건별 카운트": difference.conditional_count,
            }
            for difference in result.validation_differences
        ],
        use_container_width=True,
        hide_index=True,
    )


def render_special_items(result: AnalysisResult) -> None:
    """예상하지 못한 값, 특이사항, 제외행을 요약·상세 형태로 표시한다."""

    st.markdown("#### 예상하지 못한 값")
    if not result.unexpected_values:
        st.info("예상하지 못한 시험현황 값이 없습니다.", icon="ℹ️")
    else:
        st.dataframe(
            [
                {
                    "원본 값": item.original_value,
                    "정규화한 값": item.normalized_value,
                    "건수": item.count,
                    "지역별 건수": ", ".join(
                        f"{region} {count}" for region, count in item.region_counts.items()
                    ),
                    "행번호 범위": compress_row_numbers(item.row_numbers),
                }
                for item in result.unexpected_values
            ],
            use_container_width=True,
            hide_index=True,
        )
        with st.expander("예상하지 못한 값 상세 행 보기"):
            st.dataframe(
                [
                    {
                        "시트명": row.sheet_name,
                        "행번호": row.row_number,
                        "지역": row.region,
                        "이름": row.name,
                        "시험현황 원본 값": row.status,
                    }
                    for item in result.unexpected_values
                    for row in item.details
                ],
                use_container_width=True,
                hide_index=True,
            )

    st.markdown("#### 특이사항 및 제외행")
    st.dataframe(
        [
            {
                "구분": item.category,
                "값": item.value,
                "건수": item.count,
                "행번호 범위": item.row_range,
                "설명": item.description,
            }
            for item in result.special_items
        ],
        use_container_width=True,
        hide_index=True,
    )
    with st.expander(f"제외행 상세 보기 · {len(result.excluded_rows):,}건"):
        st.dataframe(
            [
                {
                    "시트명": row.sheet_name,
                    "행번호": row.row_number,
                    "지역": row.region,
                    "이름": row.name,
                    "시험현황": row.status,
                    "제외 사유": row.reason,
                }
                for row in result.excluded_rows
            ],
            use_container_width=True,
            hide_index=True,
        )


def render_download_section(
    *,
    xlsx_bytes: bytes,
    xlsx_filename: str,
    csv_bytes: bytes,
    csv_filename: str,
) -> None:
    """검산 성공 뒤에만 XLSX와 CSV 다운로드 버튼을 표시한다."""

    st.markdown(
        '<div class="download-callout"><strong>결과 파일이 준비되었습니다.</strong><span>원본과 별개의 새 파일이며, 원본 바이트는 변경되지 않았습니다.</span></div>',
        unsafe_allow_html=True,
    )
    first, second = st.columns(2)
    with first:
        st.download_button(
            "붙여넣기용 결과 엑셀 다운로드",
            data=xlsx_bytes,
            file_name=xlsx_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    with second:
        st.download_button(
            "결과 CSV 다운로드",
            data=csv_bytes,
            file_name=csv_filename,
            mime="text/csv; charset=utf-8",
            use_container_width=True,
        )

