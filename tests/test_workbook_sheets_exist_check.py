import pytest

from isp_workbook_parser import Parser
from isp_workbook_parser.parser import TableConfigError


def test_end_row_not_on_sheet_throws_error():
    with pytest.raises(TableConfigError):
        Parser(
            "tests/test_data/2024-isp-inputs-and-assumptions-workbook-missing-sheets.xlsx"
        )
