import pytest

from isp_assumptions_parser.parser import Parser
from isp_assumptions_parser.parser import TableConfigError
from isp_assumptions_parser.config_model import Table

workbook = Parser("workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")


def test_end_row_runs_into_another_table_throws_error():
    table_config = Table(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=34,
        column_range="B:J",
    )
    error_message = (
        "The first column of the table DUMMY contains na values indicating the table end row is "
        "incorrectly specified."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_end_row_runs_into_notes_throws_error():
    table_config = Table(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=22,
        column_range="B:J",
    )
    error_message = (
        "The first column of the table DUMMY contains the sub string 'Notes:'."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_end_row_too_soon_throws_error():
    table_config = Table(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=20,
        column_range="B:J",
    )
    error_message = (
        "There is data in the row after the defined table end for table DUMMY."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_end_column_too_soon_throws_error():
    table_config = Table(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=21,
        column_range="B:I",
    )
    error_message = (
        "There is data in the column adjacent to the last column in the table DUMMY."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook.get_table_from_config(table_config)


def test_good_config_throws_no_error():
    table_config = Table(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=21,
        column_range="B:J",
    )
    workbook.get_table_from_config(table_config)