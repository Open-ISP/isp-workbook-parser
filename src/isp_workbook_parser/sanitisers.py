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
    columns = _remove_series_trailing_footnotes_column_names(columns)
    return columns


def _replace_dataframe_hyphens_with_na(df: pd.DataFrame) -> pd.DataFrame:
    """Replaces any hyphen values with a `pandas.NA`"""
    return df.replace("-", pd.NA, regex=False)


def _replace_series_newlines_with_whitespace(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Replaces newlines in `pandas.Series` and `pandas.Index` with a whitespace"""
    return series.str.replace(r"\n", " ", regex=True)


def _remove_series_trailing_footnotes(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes footnotes by replacing a single trailing digit not preceded by
    a whitespace (e.g. name of a unit), another digit (i.e. footnotes are assumed
    to be single digit) or a capital letter (e.g. DUID) with an empty string"""
    return series.str.replace(r"(?<=[^A-Z\s\d])\d$", "", regex=True)


def _remove_series_trailing_footnotes_column_names(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes footnotes by replacing a single trailing digit not preceded by
    a whitespace (e.g. name of a unit) or another digit (i.e. footnotes are assumed
    to be single digit) with an empty string"""
    return series.str.replace(r"(?<=[^\s\d])\d$", "", regex=True)


def _remove_series_double_whitespaces(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Removes any duplicated whitespaces"""
    return series.str.replace(r"\s\s", " ", regex=True)


def _remove_series_trailing_asterisks(
    series: pd.Index | pd.Series,
) -> pd.Index | pd.Series:
    """Replaces trailing asterisks with an empty string"""
    return series.str.replace(r"\*$", "", regex=True)


def _remove_series_commas(series: pd.Index | pd.Series) -> pd.Index | pd.Series:
    """Replaces commas with an empty string"""
    return series.str.replace(r"\,", "", regex=True)
