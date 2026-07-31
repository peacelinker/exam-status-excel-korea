"""파일 검증, 문자열 처리, 안전한 오류 기록을 위한 공통 함수."""

from __future__ import annotations

import logging
import re
import traceback
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook


class AppError(Exception):
    """사용자에게 한국어로 안내할 수 있는 예상된 앱 오류."""


class InvalidExcelError(AppError):
    """업로드된 파일이 정상적인 XLSX가 아닐 때 발생하는 오류."""


def ensure_xlsx_filename(filename: str) -> None:
    """파일 확장자가 .xlsx인지 검사하고 아니면 사용자 오류를 발생시킨다."""

    if not filename.lower().endswith(".xlsx"):
        raise InvalidExcelError(".xlsx 엑셀 파일만 업로드할 수 있습니다.")


def open_workbook_from_bytes(file_bytes: bytes, *, data_only: bool = True) -> Workbook:
    """원본 바이트를 수정하지 않고 메모리에서 워크북을 연다."""

    try:
        return load_workbook(
            filename=BytesIO(file_bytes),
            data_only=data_only,
            read_only=False,
        )
    except (BadZipFile, InvalidFileException, OSError, ValueError, KeyError) as exc:
        raise InvalidExcelError("엑셀 파일이 손상되었습니다.") from exc


def compress_row_numbers(row_numbers: list[int]) -> str:
    """행번호 목록을 4-15, 18, 21-30 형식으로 압축한다."""

    values = sorted(set(row_numbers))
    if not values:
        return ""
    ranges: list[str] = []
    start = previous = values[0]
    for value in values[1:]:
        if value == previous + 1:
            previous = value
            continue
        ranges.append(str(start) if start == previous else f"{start}-{previous}")
        start = previous = value
    ranges.append(str(start) if start == previous else f"{start}-{previous}")
    return ", ".join(ranges)


def safe_filename_part(value: str) -> str:
    """다운로드 파일명에 사용할 수 없는 문자를 밑줄로 치환한다."""

    cleaned = re.sub(r"[\\/:*?\"<>|\x00-\x1f]", "_", value.strip())
    return cleaned or "분석시트"


def safe_cell_text(value: object) -> str:
    """셀 값을 표시용 문자열로 변환하고 앞뒤 공백만 제거한다."""

    if value is None:
        return ""
    return str(value).strip()


def log_safe_exception(logger: logging.Logger, exc: BaseException) -> None:
    """개인정보나 셀 값 없이 오류 종류와 마지막 코드 위치만 기록한다."""

    frames = traceback.extract_tb(exc.__traceback__) if exc.__traceback__ else []
    if frames:
        last = frames[-1]
        location = f"{Path(last.filename).name}:{last.lineno}"
    else:
        location = "알 수 없는 위치"
    logger.error("오류 종류=%s, 위치=%s", type(exc).__name__, location)
