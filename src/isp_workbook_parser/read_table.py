import pandas as pd
from isp_workbook_parser import TableConfig
import numpy as np


def read_table(workbook_file: pd.ExcelFile, table: TableConfig) -> pd.DataFrame:
    """Parses a table given a YAML config for the table

    If `table.header_rows` is an integer, the table is parsed directly.

    If `table.header_rows` is a list of integers:
        1. `table.header_rows` is checked to ensure that the list is composed of
        consecutive and increasing integers
        2. The table is parsed, with all columns read in as objects
        3. The highest header (column of DataFrame) is forward filled
        4. The intermediate headers (first rows of DataFrame), if they exist, are
        forward filled
        5. The last header is added to the list of headers without forward filling.
        Any NaNs will mean the previous header names are applied
        6. Headers are joined with '-' as a separator
        7. New headers are applied, NAs in the DataFrame are forward filled and the
        header rows in the table are dropped

    Args:
        workbook_file: pandas ExcelFile object
        table: Parsed table config

    Returns:
        Table as a pandas DataFrame

    """

    def _ffill_highest_header(initial_header: pd.Series) -> pd.Series:
        """
        Forward fills the highest header row (parsed as columns) for processing
        a multi-header table
        """
        initial_header[initial_header.str.contains("Unnamed")] = pd.NA
        ffill_initial_header = initial_header.ffill().reset_index(drop=True).fillna("")
        return ffill_initial_header

    def _ffill_intermediate_header_row(intermediate_header: pd.Series) -> pd.Series:
        """
        Forward fills intermediate header rows (parsed as rows) for processing
        a multi-header table
        """
        _ffill_intermediate_header = (
            intermediate_header.ffill().reset_index(drop=True).fillna("").astype(str)
        )
        return _ffill_intermediate_header

    def _build_cleaned_dataframe(
        df_initial: pd.DataFrame, header_rows_in_table: int, new_headers: pd.Series
    ) -> pd.DataFrame:
        """
        Builds a cleaned DataFrame with the merged headers by:
        1. Dropping the header rows in the table
        2. Applying the merged headers as the columns of the DataFrame
        3. Forward fill values across columns (handles merged value cells)
        """
        df_cleaned = df_initial.iloc[header_rows_in_table:, :]
        df_cleaned.columns = new_headers
        df_cleaned = df_cleaned.ffill(axis=1)
        return df_cleaned

    if type(table.header_rows) is int:
        df = pd.read_excel(
            workbook_file,
            sheet_name=table.sheet_name,
            header=(table.header_rows - 1),
            usecols=table.column_range,
            nrows=(table.end_row - table.header_rows),
        )
        return df
    else:
        df_initial = pd.read_excel(
            workbook_file,
            sheet_name=table.sheet_name,
            header=(table.header_rows[0] - 1),
            usecols=table.column_range,
            nrows=(table.end_row - table.header_rows[0]),
            # do not parse dtypes
            dtype="object",
        )
        # check that header_rows list is sorted
        assert sorted(table.header_rows) == table.header_rows
        # check that the header_rows are adjacent
        assert set(np.diff(table.header_rows)) == set([1])
        header_rows_in_table = table.header_rows[-1] - table.header_rows[0]
        initial_header = pd.Series(df_initial.columns)
        ffilled_initial_header = _ffill_highest_header(initial_header)
        # for multiple header rows
        filled_headers = []
        # ffill intermediate header rows
        for i in range(0, header_rows_in_table - 1):
            filled_headers.append(_ffill_intermediate_header_row(df_initial.iloc[i, :]))
        # do not ffill last header row
        filled_headers.append(
            df_initial.iloc[header_rows_in_table - 1, :]
            .reset_index(drop=True)
            .fillna("")
            .astype(str)
        )
        # add separators manually - ignore any "" entries
        for series in filled_headers:
            series[series != ""] = " - " + series[series != ""]
        merged_headers = ffilled_initial_header.str.cat(filled_headers)
        df_cleaned = _build_cleaned_dataframe(
            df_initial, header_rows_in_table, merged_headers
        )
        return df_cleaned
