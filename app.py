"""시험 응시 현황 엑셀 자동 집계기 Streamlit 진입점."""

from __future__ import annotations

import hashlib
import logging

import streamlit as st

from excel_analyzer import analyze_selected_sheet
from report_exporter import create_csv_bytes, create_result_workbook
from sheet_detector import discover_candidate_sheets
from ui_components import (
    load_css,
    render_download_section,
    render_header,
    render_header_checks,
    render_progress,
    render_region_results,
    render_sheet_candidates,
    render_special_items,
    render_step_card,
    render_summary_cards,
    render_validation_result,
)
from utils import (
    AppError,
    ensure_xlsx_filename,
    log_safe_exception,
    safe_filename_part,
)

LOGGER = logging.getLogger("exam_status_app")


def _clear_stale_result(analysis_key: str) -> None:
    """업로드 파일이나 선택 시트가 바뀌면 이전 분석 결과를 제거한다."""

    bundle = st.session_state.get("analysis_bundle")
    if bundle and bundle.get("key") != analysis_key:
        st.session_state.pop("analysis_bundle", None)


def _render_conditions() -> None:
    """사용자가 실행 전에 확인할 핵심 집계 규칙을 카드로 표시한다."""

    first, second, third, fourth = st.columns(4)
    first.metric("지역 열", "A열", "앞뒤 공백만 제거")
    second.metric("이름 열", "D열", "공란 행 제외")
    third.metric("시험현황 열", "I열", "정확히 일치하는 값")
    fourth.metric("검산", "2가지", "행 순회 + 조건별 카운트")
    st.caption(
        "대학은 수기 입력용 공란 행만 유지하고 모든 합계에서 제외합니다. "
        "숨겨진 행과 필터 범위 행도 실제 값이 있으면 포함합니다."
    )


def _render_absence_summary(result) -> None:
    """명시적 미응시, 상태 공란, 최종 미응시자를 나눠 표시한다."""

    first, second, third = st.columns(3)
    first.metric("명시적 미응시", f"{result.total_counts.explicit_absent:,}명")
    second.metric("시험현황 공란", f"{result.total_counts.blank_status:,}명")
    third.metric("최종 미응시자", f"{result.total_counts.absent_total:,}명")


def main() -> None:
    """업로드부터 검산·다운로드까지 전체 화면 흐름을 실행한다."""

    st.set_page_config(
        page_title="시험 응시 현황 엑셀 자동 집계기",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    load_css()
    render_header()

    bundle = st.session_state.get("analysis_bundle")
    render_progress(4 if bundle and bundle.get("result") else 1)

    render_step_card(
        1,
        "엑셀 파일 업로드",
        ".xlsx 파일만 허용합니다. 원본 바이트는 읽기 전용으로 사용하며 새 결과 파일을 만듭니다.",
    )
    uploaded_file = st.file_uploader(
        "시험 현황 엑셀 선택",
        type=["xlsx"],
        accept_multiple_files=False,
        help="개인정보가 포함된 파일은 비공개로 배포한 앱에서 사용하는 것을 권장합니다.",
        label_visibility="collapsed",
    )

    if uploaded_file is None:
        st.button("분석 실행", disabled=True, use_container_width=True)
        st.info("엑셀 파일을 업로드하면 후보 시트 탐색과 헤더 확인을 시작합니다.", icon="📎")
        return

    try:
        ensure_xlsx_filename(uploaded_file.name)
        file_bytes = uploaded_file.getvalue()
        file_digest = hashlib.sha256(file_bytes).hexdigest()
        candidates = discover_candidate_sheets(file_bytes)
    except AppError as exc:
        st.error(str(exc), icon="🚫")
        st.button("분석 실행", disabled=True, use_container_width=True)
        return
    except Exception as exc:  # 일반 화면에는 traceback을 노출하지 않는다.
        log_safe_exception(LOGGER, exc)
        st.error("엑셀 파일이 손상되었습니다.", icon="🚫")
        st.button("분석 실행", disabled=True, use_container_width=True)
        return

    render_step_card(
        2,
        "시험 시트 자동 탐색 결과",
        "지역전체, 직전시험, 숫자 4자리 시트를 찾고 실제 A·D·I 열과 데이터 행을 확인했습니다.",
    )
    if candidates:
        render_sheet_candidates(candidates)
    analyzable = [candidate for candidate in candidates if candidate.is_analyzable]
    if not analyzable:
        st.error("분석 가능한 시험 시트를 찾지 못했습니다.", icon="🚫")
        st.button("분석 실행", disabled=True, use_container_width=True)
        return

    recommended_index = next(
        (index for index, candidate in enumerate(analyzable) if candidate.recommended),
        0,
    )
    selected_sheet = st.selectbox(
        "분석할 최신 시험 시트 선택",
        options=[candidate.name for candidate in analyzable],
        index=recommended_index,
        help="추천은 초기 선택값일 뿐이며 원하는 후보를 직접 선택할 수 있습니다.",
    )
    selected_candidate = next(
        candidate for candidate in analyzable if candidate.name == selected_sheet
    )
    if selected_candidate.recommended:
        st.caption("현재 선택은 시트명 규칙을 기준으로 한 최신 추천 후보입니다. 필요하면 다른 시트를 선택하세요.")

    render_step_card(
        3,
        "집계 조건과 헤더 확인",
        "분석 실행 전에 고정 열 위치와 실제 헤더, 제외 규칙을 확인합니다.",
    )
    _render_conditions()
    render_header_checks(selected_candidate.header_checks)

    analysis_key = f"{file_digest}:{selected_sheet}"
    _clear_stale_result(analysis_key)
    analyze_clicked = st.button(
        "선택한 시트 분석 실행",
        type="primary",
        use_container_width=True,
        disabled=False,
    )
    if analyze_clicked:
        try:
            with st.spinner("실제 셀 값을 읽고 두 가지 방식으로 집계·검산하고 있습니다..."):
                original_digest = hashlib.sha256(file_bytes).hexdigest()
                result = analyze_selected_sheet(
                    file_bytes,
                    selected_sheet,
                    source_filename=uploaded_file.name,
                )
                if hashlib.sha256(file_bytes).hexdigest() != original_digest:
                    raise AppError("원본 엑셀 바이트가 변경되었습니다.")

                xlsx_bytes = b""
                csv_bytes = b""
                if result.validation_passed:
                    xlsx_bytes = create_result_workbook(result)
                    csv_bytes = create_csv_bytes(result)

                timestamp = result.analyzed_at.strftime("%Y%m%d_%H%M")
                sheet_part = safe_filename_part(result.selected_sheet)
                st.session_state["analysis_bundle"] = {
                    "key": analysis_key,
                    "result": result,
                    "xlsx_bytes": xlsx_bytes,
                    "csv_bytes": csv_bytes,
                    "xlsx_filename": f"시험응시현황_집계결과_{sheet_part}_{timestamp}.xlsx",
                    "csv_filename": f"시험응시현황_집계결과_{sheet_part}.csv",
                }
        except AppError as exc:
            st.session_state.pop("analysis_bundle", None)
            st.error(str(exc), icon="🚫")
        except Exception as exc:  # 일반 화면에는 traceback을 노출하지 않는다.
            st.session_state.pop("analysis_bundle", None)
            log_safe_exception(LOGGER, exc)
            st.error("분석 중 오류가 발생했습니다. 파일 구조와 필수 열을 확인해 주세요.", icon="🚫")

    bundle = st.session_state.get("analysis_bundle")
    if not bundle or bundle.get("key") != analysis_key:
        return

    result = bundle["result"]
    render_progress(4 if result.validation_passed else 3)
    render_step_card(4, "전체 집계 요약", "선택 시트의 자동 집계 핵심 수치입니다.")
    render_summary_cards(result)

    render_step_card(5, "지역별 집계 결과", "원본에서 발견된 지역 순서와 실제 집계값을 표시합니다.")
    render_region_results(result)

    render_step_card(6, "상태값별 독립 검산", "행 단위 순회와 조건별 카운트 결과를 항목별로 비교합니다.")
    render_validation_result(result)

    render_step_card(7, "미응시자 집계 결과", "명시적 미응시와 시험현황 공란을 분리해 보여줍니다.")
    _render_absence_summary(result)

    render_step_card(8, "예상하지 못한 값과 제외행", "자동 보정하지 않은 값과 집계에서 제외된 행을 확인합니다.")
    render_special_items(result)

    if result.formula_without_cached_value_count:
        st.warning("수식 셀의 저장된 계산값이 없습니다. 특이사항에서 해당 행 범위를 확인하세요.", icon="⚠️")

    if result.validation_passed:
        render_step_card(9, "결과 파일 다운로드", "새 XLSX 보고서와 UTF-8 BOM CSV를 내려받을 수 있습니다.")
        render_download_section(
            xlsx_bytes=bundle["xlsx_bytes"],
            xlsx_filename=bundle["xlsx_filename"],
            csv_bytes=bundle["csv_bytes"],
            csv_filename=bundle["csv_filename"],
        )


if __name__ == "__main__":
    main()
