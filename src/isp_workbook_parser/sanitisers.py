import pandas as pd


def _column_name_sanitiser(columns: pd.Index | pd.Series) -> pd.Index | pd.Series:
    """
    Sanitises column names by:
    1. Removing 'versioning' from column names introduced by `mangle_dupe_cols` in
    pandas parser, e.g. 'Generator.1' is sanitised to 'Generator'
    2. Stripping leading and trailing whitespaces
    3. Remove any newlines
    4. Removes duplicated whitespaces
    5. Removes trailing numbers that are footnotes
    """
    columns = columns.astype(str)
    columns = columns.str.replace(r"\.[\.\d]+$", "", regex=True)
    columns = columns.str.strip()
    columns = _replace_series_newlines_with_whitespace(columns)
    columns = _remove_series_double_whitespaces(columns)
    columns = _remove_column_name_trailing_footnotes(columns)
    return columns


def _remove_column_name_trailing_footnotes(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes footnotes by replacing a single trailing digit not preceded by
    a whitespace (e.g. name of a unit) or another digit (i.e. footnotes are assumed
    to be single digit) with an empty string"""
    return series.str.replace(r"(?<=[^\s\d])\d$", "", regex=True)


def _values_casting_and_sanitisation(df: pd.DataFrame) -> pd.DataFrame:
    """Attempts to convert `pd.DataFrame` values to numeric types. If this fails,
    sanitises strings in the same column and then re-attempts casting to a numeric type.

    String sanitisation is only applied to string values in columns that cannot be
    converted to a numeric data type, as applying string methods to non-string values
    will return `pd.NA`
    """
    df = _replace_dataframe_hyphens_with_na(df)
    for object_col in df.dtypes[df.dtypes == "object"].keys():
        try:
            df.loc[:, object_col] = pd.to_numeric(df[object_col])
        except (ValueError, TypeError):
            where_str_values = df[object_col].apply(lambda x: isinstance(x, str))
            if not df.loc[where_str_values, object_col].empty:
                for series_func in (
                    _replace_series_newlines_with_whitespace,
                    _remove_series_double_whitespaces,
                    _strip_series_whitespaces,
                    _remove_series_trailing_asterisks,
                    _remove_series_trailing_footnotes,
                    _remove_series_thousands_commas,
                    _remove_series_notes_after_values,
                ):
                    df.loc[where_str_values, object_col] = series_func(df[object_col])
            # re-attempt conversion following sanitisation
            try:
                df[object_col] = pd.to_numeric(df[object_col])
            except (ValueError, TypeError):
                pass
    return df


def _replace_dataframe_hyphens_with_na(df: pd.DataFrame) -> pd.DataFrame:
    """Replaces any hyphen values with a `pandas.NA`"""
    return df.replace("-", pd.NA, regex=False)


def _replace_series_newlines_with_whitespace(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Replaces newlines in a `pandas.Series` or `pandas.Index` with a whitespace"""
    return series.str.replace(r"\n", " ", regex=True)


def _remove_series_double_whitespaces(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes any duplicated whitespaces in a `pandas.Series` or `pandas.Index`"""
    return series.str.replace(r"\s\s", " ", regex=True)


def _remove_series_trailing_asterisks(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """
    Replaces trailing asterisks with an empty string in a `pandas.Series`
    or `pandas.Index`
    """
    return series.str.replace(r"\*$", "", regex=True)


def _remove_series_trailing_footnotes(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes footnotes in a `pandas.Series` or `pandas.Index`

    This is done by replacing a single trailing digit NOT preceded by a whitespace
    (e.g. name of a unit), another digit (i.e. footnotes are assumed
    to be single digit), a capital letter, underscore, hyphen or hash (e.g. DUID),
    hat/caret (e.g. indicating 'to the power of' in constraints) or a decimal point
    (e.g. Snowy 2.0) with an empty string
    """
    return series.str.replace(r"(?<=[^A-Z\s\d\.\_\-\#\^])\d$", "", regex=True)


def _strip_series_whitespaces(series: pd.Index | pd.Series) -> pd.Index | pd.Series:
    """Strips trailing and leading whitespaces in a `pandas.Series` or `pandas.Index`"""
    return series.str.strip(" ")


def _remove_series_thousands_commas(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes thousands commas (i.e. commas preceded by and following digits)
    in a `pandas.Series` or `pandas.Index`"""
    return series.str.replace(r"(?<=[0-9]),(?=[0-9]{1,3})", "", regex=True)


def _remove_series_notes_after_values(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes notes after numeric values in a `pandas.Series` or `pandas.Index`

    This is done using three regular expression substitutions:
        1. Capture a value (digits and decimal points) followed by one or more sequences
            of text preceded by an opening parenthesis. Retain the captured group.
        2. Capture a value (digits and decimal points) followed by one or more sequences
            of text preceded by a hyphen (with or without a space between the value
            and the hyphen), BUT not where a hyphen is used to denote a financial year
            (e.g. 2024-25). Retain the captured group.
        3. Replace any hyphen followed by one or more sequences of text preceded by a
            hyphen with an empty string.
    """
    series = series.str.replace(
        r"^([0-9\.]+)\s(?:(\([\w\s\.\<\=\-]+\)?\s?)+)", r"\1", regex=True
    )
    series = series.str.replace(
        r"^(?![0-9]{4}\-[0-9]{2,4})([0-9\.]+)\s?(?:(\-[\w\s\.\<\=\-\(\)]+)+)",
        r"\1",
        regex=True,
    )
    series = series.str.replace(r"^\-\s?(?:(\([\w\s\.\<\=\-\(\)]+)+)", "", regex=True)
    return series
