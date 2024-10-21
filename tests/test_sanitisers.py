from pathlib import Path

import pandas as pd

from isp_workbook_parser.sanitisers import (
    _remove_series_double_whitespaces,
    _remove_series_notes_after_values,
    _remove_series_thousands_commas,
    _remove_series_trailing_asterisks,
    _remove_series_trailing_footnotes,
    _replace_series_newlines_with_whitespace,
    _strip_series_whitespaces,
    _values_casting_and_sanitisation,
)


def test_sanitisation_on_flow_path_transfer_capability():
    unsanitised = pd.read_csv(Path("tests", "test_data", "unsanitised.csv"))
    expected = pd.read_csv(Path("tests", "test_data", "sanitised.csv"))
    # handle carriage return on Windows
    for col in unsanitised.columns:
        where_str = unsanitised[col].apply(lambda x: isinstance(x, str))
        unsanitised.loc[where_str, col] = unsanitised.loc[where_str, col].str.replace(
            r"\r", "", regex=True
        )
    test_sanitised = _values_casting_and_sanitisation(unsanitised)
    pd.testing.assert_frame_equal(test_sanitised, expected, check_dtype=False)


def test_replace_series_newlines_with_whitespace(sample_series):
    result = _replace_series_newlines_with_whitespace(sample_series)
    expected = pd.Series(
        [
            "First line Second line",
            "This  is  a  test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote1",
            "  leading and trailing  ",
            "1,234,567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_double_whitespaces(sample_series):
    result = _remove_series_double_whitespaces(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
            "This is a test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote1",
            " leading and trailing ",
            "1,234,567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_trailing_asterisks(sample_series):
    result = _remove_series_trailing_asterisks(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
            "This  is  a  test",
            "Value with ",
            "SomeUnitA5",
            "An actual footnote1",
            "  leading and trailing  ",
            "1,234,567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_trailing_footnotes(sample_series):
    result = _remove_series_trailing_footnotes(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
            "This  is  a  test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote",
            "  leading and trailing  ",
            "1,234,567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_strip_series_whitespaces(sample_series):
    result = _strip_series_whitespaces(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
            "This  is  a  test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote1",
            "leading and trailing",
            "1,234,567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_thousands_commas(sample_series):
    result = _remove_series_thousands_commas(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
            "This  is  a  test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote1",
            "  leading and trailing  ",
            "1234567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_notes_after_values(sample_series):
    result = _remove_series_notes_after_values(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
            "This  is  a  test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote1",
            "  leading and trailing  ",
            "1,234,567",
            "50.0",
            "35.5",
            "42.0",
            "999.9",
            "2024-25 data for year",  # No change for financial year case
        ]
    )
    pd.testing.assert_series_equal(result, expected)
