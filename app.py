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
from worship_analyzer import analyze_worship_sheet, count_worship_rosters, discover_worship_sheets
from worship_exporter import create_worship_csv, create_worship_workbook
from worship_models import WORSHIP_REGIONS

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


def _render_exam_analysis() -> None:
    """기존 시험 분석의 업로드부터 검산·다운로드까지 화면을 실행한다."""

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
                    "xlsx_filename": f"시험응시현황_붙여넣기양식_{sheet_part}_{timestamp}.xlsx",
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
        render_step_card(9, "결과 파일 다운로드", "두 번째 참고 양식과 같은 붙여넣기용 XLSX와 CSV를 내려받을 수 있습니다.")
        render_download_section(
            xlsx_bytes=bundle["xlsx_bytes"],
            xlsx_filename=bundle["xlsx_filename"],
            csv_bytes=bundle["csv_bytes"],
            csv_filename=bundle["csv_filename"],
        )


def _clear_stale_worship_result(analysis_key: str) -> None:
    bundle = st.session_state.get("worship_bundle")
    if bundle and bundle.get("key") != analysis_key:
        st.session_state.pop("worship_bundle", None)


def _parse_roster_inputs(raw_values: dict[str, str]) -> dict[str, int | None]:
    """지역별 재적 입력을 공란 또는 0 이상의 정수로 변환한다."""

    rosters: dict[str, int | None] = {}
    for region, raw_value in raw_values.items():
        value = raw_value.strip().replace(",", "")
        if not value:
            rosters[region] = None
            continue
        if not value.isdigit():
            raise AppError(f"{region} 재적은 공란 또는 0 이상의 정수로 입력해 주세요.")
        rosters[region] = int(value)
    return rosters


def _render_worship_candidates(candidates) -> None:
    rows = []
    for candidate in candidates:
        rows.append(
            {
                "시트명": candidate.name,
                "전체 사용 행": candidate.max_row,
                "이름 있음": candidate.name_rows,
                "H열 값 있음": candidate.attendance_rows,
                "H열 공란": candidate.blank_attendance_rows,
                "분석 가능": "가능" if candidate.is_analyzable else "제외",
                "제외 사유": candidate.exclusion_reason,
                "추천": "추천" if candidate.recommended else "",
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)


def _format_percent(value) -> str:
    return "" if value is None else f"{value:.1%}"


def _render_worship_result_table(result) -> None:
    rows = []
    for item in result.aggregate_rows(include_total=True):
        rows.append(
            {
                "재적": "" if item["재적"] is None else item["재적"],
                "지역": item["지역"],
                "대면": item["대면"],
                "퍼센트(대면)": _format_percent(item["퍼센트"]),
                "줌": item["줌"],
                "퍼센트(줌)": _format_percent(item["줌 퍼센트"]),
                "전화": item["전화"],
                "퍼센트(전화)": _format_percent(item["전화 퍼센트"]),
                "전체": item["전체"],
                "미참여": "" if item["미참여"] is None else item["미참여"],
                "출결 재적대비 %": _format_percent(item["출결 재적대비 %"]),
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)


def _render_worship_analysis() -> None:
    """구역예배 A·D·H 실제 셀 집계와 자동 재적 확인 흐름을 실행한다."""

    bundle = st.session_state.get("worship_bundle")
    render_progress(4 if bundle and bundle.get("result") else 1)
    render_step_card(
        1,
        "구역예배 엑셀 업로드",
        "원본 XLSX의 A열 지역·D열 이름·H열 참여방식을 직접 읽습니다. 원본 파일은 수정하지 않습니다.",
    )
    uploaded_file = st.file_uploader(
        "구역예배 현황 엑셀 선택",
        type=["xlsx"],
        accept_multiple_files=False,
        key="worship_uploader",
        label_visibility="collapsed",
    )
    if uploaded_file is None:
        st.button("구역예배 분석 실행", disabled=True, use_container_width=True)
        st.info("엑셀 파일을 업로드하면 모든 시트의 A·D·H 열을 확인합니다.", icon="📎")
        return

    try:
        ensure_xlsx_filename(uploaded_file.name)
        file_bytes = uploaded_file.getvalue()
        file_digest = hashlib.sha256(file_bytes).hexdigest()
        candidates = discover_worship_sheets(file_bytes)
    except AppError as exc:
        st.error(str(exc), icon="🚫")
        return
    except Exception as exc:
        log_safe_exception(LOGGER, exc)
        st.error("엑셀 파일이 손상되었거나 읽을 수 없습니다.", icon="🚫")
        return

    render_step_card(
        2,
        "구역예배 시트 선택",
        "‘지역전체’ 또는 0726처럼 이름이 정확히 네 자리 숫자인 시트만 표시합니다.",
    )
    _render_worship_candidates(candidates)
    analyzable = [candidate for candidate in candidates if candidate.is_analyzable]
    if not analyzable:
        st.error("분석 가능한 구역예배 시트를 찾지 못했습니다.", icon="🚫")
        return
    recommended_index = next(
        (index for index, candidate in enumerate(analyzable) if candidate.recommended),
        0,
    )
    selected_sheet = st.selectbox(
        "분석할 구역예배 시트",
        options=[candidate.name for candidate in analyzable],
        index=recommended_index,
        key="worship_sheet",
    )
    selected_candidate = next(item for item in analyzable if item.name == selected_sheet)
    render_header_checks(selected_candidate.header_checks)

    try:
        auto_rosters = count_worship_rosters(file_bytes, selected_sheet)
    except AppError as exc:
        st.error(str(exc), icon="🚫")
        return

    render_step_card(
        3,
        "지역별 재적 자동 입력 및 확인",
        "A열 지역과 D열 이름이 있는 실제 행 수를 지역원 총원으로 넣었습니다. 필요하면 수정하거나 비워 둘 수 있습니다.",
    )
    st.caption("집계 기준: H열이 정확히 ‘대면모임’이면 대면, ‘줌’이면 줌, ‘통화’이면 전화")
    st.caption("재적 기준: H열 값과 관계없이 A열이 해당 지역이고 D열 이름이 있는 행 수")
    raw_rosters: dict[str, str] = {}
    first_row = st.columns(4)
    second_row = st.columns(3)
    for index, region in enumerate(WORSHIP_REGIONS):
        columns = first_row if index < 4 else second_row
        column_index = index if index < 4 else index - 4
        with columns[column_index]:
            raw_rosters[region] = st.text_input(
                f"{region} 재적",
                value=str(auto_rosters[region]),
                placeholder="공란 가능",
                key=f"worship_roster:{file_digest}:{selected_sheet}:{region}",
            )

    report_title = f"{selected_sheet} 구역예배 성인"
    analysis_key = f"worship:{file_digest}:{selected_sheet}:{tuple(raw_rosters.items())}"
    _clear_stale_worship_result(analysis_key)

    analyze_clicked = st.button(
        "선택한 시트의 실제 셀 값 분석 실행",
        type="primary",
        use_container_width=True,
    )
    if analyze_clicked:
        try:
            rosters = _parse_roster_inputs(raw_rosters)
            with st.spinner("A·D·H 실제 셀 값을 행별로 읽고 독립 검산하고 있습니다..."):
                original_digest = hashlib.sha256(file_bytes).hexdigest()
                result = analyze_worship_sheet(
                    file_bytes,
                    selected_sheet,
                    rosters=rosters,
                    report_title=report_title,
                    source_filename=uploaded_file.name,
                )
                if hashlib.sha256(file_bytes).hexdigest() != original_digest:
                    raise AppError("원본 엑셀 바이트가 변경되었습니다.")
                xlsx_bytes = create_worship_workbook(result)
                csv_bytes = create_worship_csv(result)
                timestamp = result.analyzed_at.strftime("%Y%m%d_%H%M")
                sheet_part = safe_filename_part(result.selected_sheet)
                st.session_state["worship_bundle"] = {
                    "key": analysis_key,
                    "result": result,
                    "xlsx_bytes": xlsx_bytes,
                    "csv_bytes": csv_bytes,
                    "xlsx_filename": f"구역예배_지역별집계_{sheet_part}_{timestamp}.xlsx",
                    "csv_filename": f"구역예배_지역별집계_{sheet_part}.csv",
                }
        except AppError as exc:
            st.session_state.pop("worship_bundle", None)
            st.error(str(exc), icon="🚫")
        except Exception as exc:
            st.session_state.pop("worship_bundle", None)
            log_safe_exception(LOGGER, exc)
            st.error("구역예배 분석 중 오류가 발생했습니다. 파일 구조를 확인해 주세요.", icon="🚫")

    bundle = st.session_state.get("worship_bundle")
    if not bundle or bundle.get("key") != analysis_key:
        return
    result = bundle["result"]
    render_progress(4 if result.validation_passed else 3)
    render_step_card(4, "구역예배 지역별 결과", "이미지와 같은 고정 지역 순서로 실제 집계값을 표시합니다.")
    summary_columns = st.columns(4)
    summary_columns[0].metric("전체 참여", f"{result.total_counts.total:,}명")
    summary_columns[1].metric("대면", f"{result.total_counts.face_to_face:,}명")
    summary_columns[2].metric("줌", f"{result.total_counts.zoom:,}명")
    summary_columns[3].metric("전화", f"{result.total_counts.phone:,}명")
    _render_worship_result_table(result)

    if any(item.roster is None for item in result.region_results):
        st.info("재적이 공란인 지역과 전체 행의 퍼센트·미참여는 빈칸으로 유지했습니다.", icon="ℹ️")
    over_roster = [item for item in result.region_results if item.roster is not None and item.counts.total > item.roster]
    if over_roster:
        st.warning(
            "참여 인원이 재적보다 많은 지역이 있습니다: "
            + ", ".join(f"{item.region}({item.counts.total}/{item.roster})" for item in over_roster),
            icon="⚠️",
        )
    if result.validation_passed:
        st.success("행 순회 방식과 조건별 카운트 방식의 지역별 결과가 모두 일치합니다.", icon="✅")
    else:
        st.error("독립 검산 결과가 일치하지 않아 결과 파일을 생성하지 않았습니다.", icon="🚫")

    render_step_card(5, "제외행 확인", "공란·대상 외 지역·예상하지 못한 H열 값은 자동 보정하지 않고 별도로 표시합니다.")
    if result.excluded_rows:
        st.dataframe(
            [
                {
                    "시트명": item.sheet_name,
                    "행번호": item.row_number,
                    "지역": item.region,
                    "이름": item.name,
                    "H열 값": item.attendance,
                    "제외 사유": item.reason,
                }
                for item in result.excluded_rows
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("제외된 행이 없습니다.", icon="ℹ️")

    if result.validation_passed:
        render_step_card(6, "결과 파일 다운로드", "재적을 첫 열에 넣은 11열 XLSX와 CSV를 내려받을 수 있습니다.")
        first, second = st.columns(2)
        with first:
            st.download_button(
                "구역예배 결과 엑셀 다운로드",
                data=bundle["xlsx_bytes"],
                file_name=bundle["xlsx_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        with second:
            st.download_button(
                "구역예배 결과 CSV 다운로드",
                data=bundle["csv_bytes"],
                file_name=bundle["csv_filename"],
                mime="text/csv; charset=utf-8",
                use_container_width=True,
            )


def main() -> None:
    """시험과 구역예배 분석 모드를 한 앱에서 제공한다."""

    st.set_page_config(
        page_title="시험·구역예배 엑셀 자동 집계기",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    load_css()
    render_header()
    analysis_type = st.radio(
        "분석 유형",
        options=["시험", "구역예배"],
        horizontal=True,
        key="analysis_type",
    )
    if analysis_type == "시험":
        _render_exam_analysis()
    else:
        _render_worship_analysis()


if __name__ == "__main__":
    main()

