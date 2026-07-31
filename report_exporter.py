"""검산을 통과한 분석 결과를 새 XLSX와 CSV 바이트로 생성한다."""

from __future__ import annotations

import csv
from io import BytesIO, StringIO

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from models import AnalysisResult
from utils import AppError, compress_row_numbers

AGGREGATE_HEADERS = [
    "지역",
    "전체 응시목표",
    "정규응시",
    "대면응시",
    "일대일응시",
    "서면응시",
    "비공식응시",
    "전체응시",
    "미응시자",
    "전도 재적대비 %",
]
SHEET_NAMES = [
    "집계결과",
    "전체합계",
    "상태값검산",
    "헤더확인",
    "특이사항",
    "제외행",
    "분석정보",
]

NAVY = "16324F"
BLUE = "2563EB"
PALE_BLUE = "EAF2FF"
PALE_GRAY = "F5F7FA"
GRID = "DCE2EA"
WHITE = "FFFFFF"
SUCCESS = "DDF5E7"
WARNING = "FFF1D6"
TOTAL = "FFF4B8"
THIN_BORDER = Border(bottom=Side(style="thin", color=GRID))


def _sheet(workbook: Workbook, name: str):
    """기존 시트를 재사용하거나 새 시트를 만든다."""

    if name in workbook.sheetnames:
        return workbook[name]
    return workbook.create_sheet(name)


def _title(worksheet, text: str, last_column: int) -> None:
    """워크시트 상단에 일관된 보고서 제목 띠를 적용한다."""

    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    cell = worksheet.cell(1, 1, text)
    cell.font = Font(name="맑은 고딕", size=16, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(vertical="center", horizontal="left")
    worksheet.row_dimensions[1].height = 30
    worksheet.sheet_view.showGridLines = False


def _header(worksheet, row: int, headers: list[str]) -> None:
    """표 헤더에 배경색, 중앙 정렬, 하단 구분선을 적용한다."""

    for column, value in enumerate(headers, start=1):
        cell = worksheet.cell(row, column, value)
        cell.font = Font(name="맑은 고딕", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    worksheet.row_dimensions[row].height = 28


def _format_body(worksheet, start_row: int, end_row: int, columns: int) -> None:
    """본문의 글꼴, 정렬, 교대 배경, 숫자 형식을 정리한다."""

    for row in range(start_row, end_row + 1):
        fill = PatternFill("solid", fgColor=PALE_GRAY if row % 2 == 0 else WHITE)
        for column in range(1, columns + 1):
            cell = worksheet.cell(row, column)
            cell.font = Font(name="맑은 고딕", size=10, color="273142")
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = Alignment(
                horizontal="right" if isinstance(cell.value, (int, float)) else "left",
                vertical="center",
                wrap_text=column == 1,
            )
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"


def _autosize(worksheet, *, max_width: int = 48) -> None:
    """내용 길이에 맞춰 열 너비를 조정하되 과도한 폭은 제한한다."""

    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        length = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=8,
        )
        worksheet.column_dimensions[column_letter].width = min(max(length + 3, 10), max_width)


def write_aggregate_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """지역별 고정 컬럼 결과, 대학 공란 행, 전체 합계를 작성한다."""

    worksheet = _sheet(workbook, "집계결과")
    _title(worksheet, "시험 응시 현황 지역별 집계 결과", len(AGGREGATE_HEADERS))
    _header(worksheet, 2, AGGREGATE_HEADERS)
    row_number = 3
    manual_rows: list[int] = []
    for region_result in result.region_results:
        values = list(region_result.as_row().values())
        for column, value in enumerate(values, start=1):
            worksheet.cell(row_number, column, value)
        if region_result.is_manual_region:
            manual_rows.append(row_number)
            worksheet.cell(row_number, 2).comment = Comment(
                "대학 지역은 별도 시트에서 수기로 입력합니다. 자동 집계 수치와 전체 합계에서 제외됩니다.",
                "자동 집계기",
            )
        row_number += 1

    total_values = list(result.aggregate_rows(include_total=True)[-1].values())
    for column, value in enumerate(total_values, start=1):
        cell = worksheet.cell(row_number, column, value)
        cell.font = Font(name="맑은 고딕", bold=True, color=NAVY)
        cell.fill = PatternFill("solid", fgColor=TOTAL)
        cell.border = THIN_BORDER
        if isinstance(value, (int, float)):
            cell.number_format = "#,##0"

    _format_body(worksheet, 3, max(3, row_number - 1), len(AGGREGATE_HEADERS))
    for manual_row in manual_rows:
        for column in range(1, len(AGGREGATE_HEADERS) + 1):
            worksheet.cell(manual_row, column).fill = PatternFill("solid", fgColor=WARNING)
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter.ref = f"A2:J{row_number}"
    worksheet.column_dimensions["A"].width = 16
    for column in range(2, 11):
        worksheet.column_dimensions[get_column_letter(column)].width = 17


def write_total_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """전체 합계와 제외·경고 건수를 세로형 요약 표로 작성한다."""

    worksheet = _sheet(workbook, "전체합계")
    _title(worksheet, "전체 집계 요약", 2)
    _header(worksheet, 2, ["항목", "값"])
    metrics = [
        ("정규응시 합계", result.total_counts.regular_total),
        ("대면응시 합계", result.total_counts.face_to_face),
        ("일대일응시 합계", result.total_counts.one_to_one),
        ("서면응시 합계", result.total_counts.written),
        ("비공식응시 합계", result.total_counts.informal),
        ("전체응시 합계", result.total_counts.total_exam),
        ("명시적 미응시 합계", result.total_counts.explicit_absent),
        ("시험현황 공란 합계", result.total_counts.blank_status),
        ("최종 미응시자 합계", result.total_counts.absent_total),
        ("예상하지 못한 상태값 행 수", result.total_counts.unexpected),
        (
            "대학 지역 제외 행 수",
            sum(row.reason == "대학 지역" for row in result.excluded_rows),
        ),
        (
            "지역 공란 제외 행 수",
            sum(row.reason == "지역 공란" for row in result.excluded_rows),
        ),
    ]
    for row, values in enumerate(metrics, start=3):
        worksheet.cell(row, 1, values[0])
        worksheet.cell(row, 2, values[1])
    _format_body(worksheet, 3, 2 + len(metrics), 2)
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter.ref = f"A2:B{2 + len(metrics)}"
    _autosize(worksheet)


def write_status_validation_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """각 지역에서 발견된 원시 상태값과 파생 합계를 작성한다."""

    worksheet = _sheet(workbook, "상태값검산")
    headers = [
        "지역",
        "정규응시",
        "정규응시(타지파)",
        "정규응시 합계",
        "대면응시",
        "일대일응시",
        "서면응시",
        "비공식응시",
        "명시적 미응시",
        "시험현황 공란",
        "예상하지 못한 값",
        "전체 데이터 행",
    ]
    _title(worksheet, "상태값별 검산", len(headers))
    _header(worksheet, 2, headers)
    row_number = 3
    for region_result in result.region_results:
        if region_result.is_manual_region:
            continue
        counts = region_result.counts
        values = [
            region_result.region,
            counts.regular_standard,
            counts.regular_other_tribe,
            counts.regular_total,
            counts.face_to_face,
            counts.one_to_one,
            counts.written,
            counts.informal,
            counts.explicit_absent,
            counts.blank_status,
            counts.unexpected,
            counts.total_exam
            + counts.explicit_absent
            + counts.blank_status
            + counts.unexpected,
        ]
        for column, value in enumerate(values, start=1):
            worksheet.cell(row_number, column, value)
        row_number += 1
    if row_number > 3:
        _format_body(worksheet, 3, row_number - 1, len(headers))
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter.ref = f"A2:L{max(2, row_number - 1)}"
    _autosize(worksheet, max_width=24)


def write_header_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """헤더 기대값과 실제값, 판정, 설명을 작성한다."""

    worksheet = _sheet(workbook, "헤더확인")
    headers = ["시트명", "열", "기대값", "실제값", "판정", "설명"]
    _title(worksheet, "필수 헤더 확인 결과", len(headers))
    _header(worksheet, 2, headers)
    for row, check in enumerate(result.header_checks, start=3):
        values = [
            check.sheet_name,
            check.column,
            check.expected_label,
            check.actual_value,
            check.status,
            check.description,
        ]
        for column, value in enumerate(values, start=1):
            worksheet.cell(row, column, value)
    _format_body(worksheet, 3, 2 + len(result.header_checks), len(headers))
    for row, check in enumerate(result.header_checks, start=3):
        worksheet.cell(row, 5).fill = PatternFill(
            "solid", fgColor=SUCCESS if check.status == "일치" else WARNING
        )
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter.ref = f"A2:F{2 + len(result.header_checks)}"
    _autosize(worksheet)


def write_special_items_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """공란, 예상 밖 상태, 구조 경고 등 특이사항을 작성한다."""

    worksheet = _sheet(workbook, "특이사항")
    headers = ["구분", "값", "건수", "행번호 범위", "설명"]
    _title(worksheet, "특이사항 및 분석 경고", len(headers))
    _header(worksheet, 2, headers)
    for row, item in enumerate(result.special_items, start=3):
        values = [item.category, item.value, item.count, item.row_range, item.description]
        for column, value in enumerate(values, start=1):
            worksheet.cell(row, column, value)
    end_row = 2 + len(result.special_items)
    if result.special_items:
        _format_body(worksheet, 3, end_row, len(headers))
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter.ref = f"A2:E{max(2, end_row)}"
    _autosize(worksheet)


def write_excluded_rows_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """집계에서 제외된 행의 위치와 사유를 감사 표로 작성한다."""

    worksheet = _sheet(workbook, "제외행")
    headers = ["시트명", "행번호", "지역", "이름", "시험현황", "제외 사유"]
    _title(worksheet, "집계 제외 행", len(headers))
    _header(worksheet, 2, headers)
    for row, item in enumerate(result.excluded_rows, start=3):
        values = [
            item.sheet_name,
            item.row_number,
            item.region,
            item.name,
            item.status,
            item.reason,
        ]
        for column, value in enumerate(values, start=1):
            worksheet.cell(row, column, value)
    end_row = 2 + len(result.excluded_rows)
    if result.excluded_rows:
        _format_body(worksheet, 3, end_row, len(headers))
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter.ref = f"A2:F{max(2, end_row)}"
    _autosize(worksheet)


def write_analysis_info_sheet(workbook: Workbook, result: AnalysisResult) -> None:
    """개인 이름 없이 파일·시트·분석 범위·검산 정보를 작성한다."""

    worksheet = _sheet(workbook, "분석정보")
    _title(worksheet, "분석 실행 정보", 2)
    _header(worksheet, 2, ["항목", "값"])
    info = [
        ("분석 일시", result.analyzed_at.isoformat(timespec="seconds")),
        ("원본 파일명", result.source_filename),
        ("선택한 시트명", result.selected_sheet),
        ("후보 시트 목록", ", ".join(result.candidate_names)),
        ("사용한 헤더 행", result.header_row),
        ("데이터 시작 행", result.data_start_row),
        ("마지막 데이터 행", result.last_data_row),
        ("숨겨진 행 수", result.hidden_row_count),
        ("필터 설정 여부", "예" if result.filter_applied else "아니요"),
        (
            "검산 결과",
            "일치" if result.validation_passed else "불일치",
        ),
        ("집계 규칙 버전", result.rule_version),
        (
            "수식 저장 결과값 없음",
            result.formula_without_cached_value_count,
        ),
    ]
    for row, values in enumerate(info, start=3):
        worksheet.cell(row, 1, values[0])
        worksheet.cell(row, 2, values[1])
    _format_body(worksheet, 3, 2 + len(info), 2)
    _autosize(worksheet)


def create_result_workbook(result: AnalysisResult) -> bytes:
    """검산 성공 결과를 7개 시트의 새 XLSX 바이트로 생성한다."""

    if not result.validation_passed:
        raise AppError("두 가지 검산 결과가 일치하지 않습니다.")
    try:
        workbook = Workbook()
        workbook.active.title = "집계결과"
        write_aggregate_sheet(workbook, result)
        write_total_sheet(workbook, result)
        write_status_validation_sheet(workbook, result)
        write_header_sheet(workbook, result)
        write_special_items_sheet(workbook, result)
        write_excluded_rows_sheet(workbook, result)
        write_analysis_info_sheet(workbook, result)
        stream = BytesIO()
        workbook.save(stream)
        return stream.getvalue()
    except AppError:
        raise
    except Exception as exc:
        raise AppError("결과 엑셀 생성에 실패했습니다.") from exc


def create_csv_bytes(result: AnalysisResult) -> bytes:
    """집계결과와 같은 컬럼의 UTF-8 BOM CSV 바이트를 생성한다."""

    if not result.validation_passed:
        raise AppError("두 가지 검산 결과가 일치하지 않습니다.")
    stream = StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=AGGREGATE_HEADERS, extrasaction="ignore")
    writer.writeheader()
    for row in result.aggregate_rows(include_total=True):
        writer.writerow({key: "" if value is None else value for key, value in row.items()})
    return stream.getvalue().encode("utf-8-sig")
