import re

import pytest

from isp_workbook_parser.config_model import TableConfig
from isp_workbook_parser.parser import TableConfigError


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


def test_last_column_empty_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Generation limits",
        header_rows=8,
        end_row=52,
        column_range="B:E",
    )
    error_message = f"The last column of the table {table_config.name} is empty."
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
        sheet_name="Generator Reliability Settings",
        header_rows=[20, 21],
        end_row=28,
        column_range="B:G",
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


def test_duplicate_column_names_throws_error(workbook_v6):
    table_config = TableConfig(
        name="DUMMY",
        sheet_name="Network Capability",
        header_rows=[7, 8],
        end_row=20,
        column_range="B:J",
    )
    error_message = "There are duplicate column names in the table DUMMY."
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


def test_incorrect_table_name_throws_error(workbook_v6):
    error_message = "The table_name (affine_heat_rates_new_entrant) provided is not in the config for this workbook version. Did you mean 'affine_heat_rates_new_entrants'?"
    error_message = re.escape(error_message)
    with pytest.raises(ValueError, match=error_message):
        workbook_v6.get_table("affine_heat_rates_new_entrant")
