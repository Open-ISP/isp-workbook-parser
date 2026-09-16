# Copyright (C) 2026 University of New South Wales
#
# This file is free software; you can redistribute it and/or modify it
# under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.

from pathlib import Path

import pandas as pd

from isp_workbook_parser.sanitisers import (
    _extract_numeric_value_millions,
    _remove_series_bracketed_footnotes,
    _remove_series_double_whitespaces,
    _remove_series_notes_after_values,
    _remove_series_thousands_commas,
    _remove_series_trailing_asterisks,
    _remove_series_trailing_footnotes,
    _replace_series_newlines_with_whitespace,
    _strip_series_whitespaces,
    _values_casting_and_sanitisation,
    _where_multiple_values_with_notes,
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_extract_numeric_value_millions(sample_series):
    result = _extract_numeric_value_millions(sample_series)
    expected = pd.Series(
        [
            "First line\nSecond line",
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
            1_000_000,
            1_000_000,
            1_000_000,
            "$ 1.0.0 M",
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
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
            "$ 1 M",
            "$ 1.0 M",
            "$ 1,,.0 M",
            "$ 1.0.0 M",
        ]
    )
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_notes_after_values_with_special_characters():
    unsanitised = pd.Series(
        [
            (
                "4758 (Marinus Link Pty Ltd and TasNetworks have advised that $534 million, "
                "in $2023, of this amount relates to approved early works and other incurred "
                "costs that should be excluded from the cost estimate for the 2026 ISP in "
                "accordance with the AER's CBA Guidelines. AEMO has removed this from the "
                "estimate of $5035 million in $2023, and has then adjusted to $2025.)"
            ),
            (
                "7035 (Transgrid has advised $565 million of this amount relates to approved "
                "early works and other incurred costs that should be excluded from the total "
                "cost estimate of $7600 million for the 2026 ISP.)"
            ),
            (
                "2431 (This figure reflects the estimate from Option 2 with a portion costed "
                "at Class 5b removed.)"
            ),
            "1749.5 (only part of this figure is included)",
        ]
    )
    result = _remove_series_notes_after_values(unsanitised)
    expected = pd.Series(["4758", "7035", "2431", "1749.5"])
    pd.testing.assert_series_equal(result, expected)


def test_remove_series_bracketed_footnotes():
    unsanitised = pd.Series(
        [
            "750[footnote14]",
            "750[footnote15]",
            "1234[footnote 3]",
            "750",
        ]
    )
    result = _remove_series_bracketed_footnotes(unsanitised)
    expected = pd.Series(["750", "750", "1234", "750"])
    pd.testing.assert_series_equal(result, expected)


# (cell, does `_where_multiple_values_with_notes` fire, result after
# `_remove_series_notes_after_values`). Cells reach these functions with newlines
# already replaced by whitespace, so the two-value cells are written that way here.
MULTIPLE_VALUE_CASES = [
    # -- Two values, each with a note: the whole cell is kept, as no single value can
    # stand in for it. Four further cells in the workbooks share the shape of the
    # first case ("201 (with Marinus Link) 434 (without Marinus Link)" and
    # "0 (with Marinus Link) 0 (without Marinus Link)" in 6.0 REZ Augmentations
    # Options, "250 (generation) 325 (pump)" in 6.0 Storage properties and
    # "350 (Summer) 362 (Winter)" in 7.3 Gas System Properties).
    (  # 6.0 Flow Path Augmentation options J34
        "930 (NSW works) 964 (QLD works)",
        True,
        "930 (NSW works) 964 (QLD works)",
    ),
    (  # 6.0 Flow Path Augmentation options M80, approximation symbol on the first value
        "~90 (underground cable) 0 (HVAC new easement)",
        True,
        "~90 (underground cable) 0 (HVAC new easement)",
    ),
    (  # 6.0 Flow Path Augmentation options M81, and on the second value
        "0 (underground cable) ~94 (HVAC new easement)",
        True,
        "0 (underground cable) ~94 (HVAC new easement)",
    ),
    # Punctuation other than whitespace may separate the two values.
    ("930 (NSW works), 964 (QLD works)", True, "930 (NSW works), 964 (QLD works)"),
    # The second value need not carry a note of its own.
    ("250 (generation) 325", True, "250 (generation) 325"),
    # -- One value plus a note: still cut down to the value, so the column casts to a
    # numeric type.
    ("685 (with QNI Minor)", False, "685"),  # 6.0 Network Capability
    ("0.16 (apply from 5,400 MW)", False, "0.16"),  # 6.0 Build limits, decimal value
    (  # 6.0 Network Capability. The digit after the note is a footnote reference, not
        # a second value, so the separator between them may not contain letters.
        "400 (with VNI SIPS) - Note 8 (Snowy 2.0 generation or pump load <= 660 "
        "- Note 11)",
        False,
        "400",
    ),
    # A second bracketed note is not a second value either.
    ("400 (with VNI SIPS) (Note 4)", False, "400"),  # 6.0 Network Capability
    # Currency symbols and thousands commas in the note (the case that motivated
    # discarding the whole note rather than matching its characters).
    ("5325 (based on provided cost of $5,035 in $2023)", False, "5325"),  # 7.3
    # A note containing parentheses does not make the cell multi-value: the note is
    # the first parenthesis-free bracketed run, so the "2023" below is part of the
    # note rather than a second value.
    ("5325 (provided cost ($5,035) 2023)", False, "5325"),
    # -- Two values, but not in the shape substitution 1 matches, so the predicate has
    # nothing to guard against: no note is stripped and the cell comes through whole.
    ("930(NSW works) 964 (QLD works)", False, "930(NSW works) 964 (QLD works)"),
    (
        "about 930 (NSW works) 964 (QLD works)",
        False,
        "about 930 (NSW works) 964 (QLD works)",
    ),
    # -- A known gap. Two values delimited by hyphens rather than bracketed notes are
    # still cut down to the first, by substitution 2, which this predicate does not
    # guard. Whether the workbooks hold cells of this shape has not been checked.
    ("930 - NSW works 964 - QLD works", False, "930"),
    # -- Nothing to flag.
    ("930 (NSW works)", False, "930"),
    ("930", False, "930"),
    ("", False, ""),
]


def test_multiple_values_with_notes_detection_and_sanitisation():
    """Cells holding two values are detected and kept whole; cells holding one value
    and a note are still cut down to that value."""
    cells, fires, sanitised = (
        list(field) for field in zip(*MULTIPLE_VALUE_CASES, strict=True)
    )
    unsanitised = pd.Series(cells)
    pd.testing.assert_series_equal(
        _where_multiple_values_with_notes(unsanitised), pd.Series(fires)
    )
    pd.testing.assert_series_equal(
        _remove_series_notes_after_values(unsanitised), pd.Series(sanitised)
    )


def test_where_multiple_values_with_notes_on_mixed_and_index_input():
    """Columns reaching the sanitisers hold a mix of strings, numbers and nulls, and
    the sanitisers are also applied to a `pandas.Index` of column names."""
    series = pd.Series(["250 (generation) 325 (pump)", "250 (generation)", 42.0, None])
    pd.testing.assert_series_equal(
        _where_multiple_values_with_notes(series),
        pd.Series([True, False, False, False]),
    )
    index = pd.Index(["250 (generation) 325 (pump)", "250 (generation)"])
    assert list(_where_multiple_values_with_notes(index)) == [True, False]


def test_values_casting_and_sanitisation_leaves_multi_value_column_as_text():
    """A column containing a multi-value cell cannot be cast to a numeric type, which
    is the signal to consumers that the cell holds more than one value."""
    df = pd.DataFrame({"capacity": ["250 (generation) 325 (pump)", "500 (generation)"]})
    result = _values_casting_and_sanitisation(df)
    assert result["capacity"].dtype == "object"
    assert result["capacity"].tolist() == ["250 (generation) 325 (pump)", "500"]
