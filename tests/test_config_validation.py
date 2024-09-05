import pytest

from isp_workbook_parser.parser import TableConfigError
from isp_workbook_parser.config_model import TableConfig


def test_end_row_not_on_sheet_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Aggregated energy storages",
        header_rows=95,
        end_row=200,
        column_range="B:AF",
    )
    error_message = (
        f"The end_row for table {table_config.name} is not within the excel sheet."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_first_header_row_not_on_sheet_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Aggregated energy storages",
        header_rows=200,
        end_row=95,
        column_range="B:AF",
    )
    error_message = f"The first header row for table {table_config.name} is not within the excel sheet."
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_first_column_not_on_sheet_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Aggregated energy storages",
        header_rows=10,
        end_row=23,
        column_range="AI:AF",
    )
    error_message = (
        f"The first column for table {table_config.name} is not within the excel sheet."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_last_column_not_on_sheet_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Aggregated energy storages",
        header_rows=10,
        end_row=23,
        column_range="B:AI",
    )
    error_message = (
        f"The last column for table {table_config.name} is not within the excel sheet."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_end_row_runs_into_another_table_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Aggregated energy storages",
        header_rows=10,
        end_row=23,
        column_range="B:AF",
    )
    error_message = (
        "The first column of the table DUMMY contains na values indicating the table end row is "
        "incorrectly specified."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_end_row_runs_into_notes_throws_error(workbook_v6):
    table_config = TableConfig(
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
        workbook_v6.get_table_from_config(table_config)


def test_first_header_row_too_late_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[7, 8],
        end_row=20,
        column_range="B:J",
    )
    error_message = f"There is data or a header above the first header row for table {table_config.name}."
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_end_row_too_soon_throws_error(workbook_v6):
    table_config = TableConfig(
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
        workbook_v6.get_table_from_config(table_config)


def test_end_column_too_soon_throws_error(workbook_v6):
    table_config = TableConfig(
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
        workbook_v6.get_table_from_config(table_config)


def test_start_column_too_far_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=21,
        column_range="C:J",
    )
    error_message = (
        "There is data in the column adjacent to the first column in the table DUMMY."
    )
    with pytest.raises(TableConfigError, match=error_message):
        workbook_v6.get_table_from_config(table_config)


def test_good_config_throws_no_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[6, 7],
        end_row=21,
        column_range="B:J",
    )
    workbook_v6.get_table_from_config(table_config)
