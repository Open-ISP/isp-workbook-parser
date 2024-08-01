import pytest

from isp_assumptions_parser.parser import Parser
from isp_assumptions_parser.parser import TableConfigError

workbook = Parser("../workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")


def test_end_row_runs_into_another_table_throws_error():
    table_config = {
        "tab": "Network Capability",
        "junk_rows_at_start": 1,
        "column_name_corrections": {},
        "start_row": 6,
        "end_row": 34,
        "range": "B:J",
    }
    error_message = (
        "The first column of the table contains na values indicating the table end row is "
        "incorrectly specified."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_end_row_runs_into_notes_throws_error():
    table_config = {
        "tab": "Network Capability",
        "junk_rows_at_start": 1,
        "column_name_corrections": {},
        "start_row": 6,
        "end_row": 22,
        "range": "B:J",
    }
    error_message = "The first column of the table contains the sub string 'Notes:'."
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_end_row_too_soon_throws_error():
    table_config = {
        "tab": "Network Capability",
        "junk_rows_at_start": 1,
        "column_name_corrections": {},
        "start_row": 6,
        "end_row": 20,
        "range": "B:J",
    }
    error_message = "There is data in the row after the defined table end."
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_end_column_too_soon_throws_error():
    table_config = {
        "tab": "Network Capability",
        "junk_rows_at_start": 1,
        "column_name_corrections": {},
        "start_row": 6,
        "end_row": 21,
        "range": "B:I",
    }
    error_message = (
        "There is data in the column adjacent to the last column in the table."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_good_config_throws_no_error():
    table_config = {
        "tab": "Network Capability",
        "junk_rows_at_start": 1,
        "column_name_corrections": {},
        "start_row": 6,
        "end_row": 21,
        "range": "B:J",
    }
    workbook.get_table_from_config(table_config)
