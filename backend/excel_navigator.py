"""
Excel 셀 탐색 모듈
ExcelReader의 navigate_to_cell을 호출합니다.
"""

from typing import Dict


def navigate_to_cell(workbook_name: str, sheet_name: str, cell_ref: str) -> Dict:
    """Excel에서 지정된 셀로 이동합니다."""
    from backend.excel_reader import get_excel_reader
    reader = get_excel_reader()
    return reader.navigate_to_cell(workbook_name, sheet_name, cell_ref)
