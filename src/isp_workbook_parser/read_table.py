import pandas as pd
from isp_workbook_parser import TableConfig
from typing import Union, List
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

    Examples:

    The example below reads the "Existing Generators Summary" table from the 2024
    version 6 workbook.

    >>> from isp_workbook_parser import TableConfig, read_table

    >>> config = TableConfig(
    ... name='existing_generators_summary',
    ... sheet_name='Summary Mapping',
    ... header_rows=[4, 5, 6],
    ... end_row=258,
    ... column_range="B:AC",
    ... )

    >>> existing_gen_summary = read_table(
    ... workbook_file='workbooks/6.0/2024-isp-inputs-and-assumptions-workbook.xlsx',
    ... table=config,
    ... )
    >>> existing_gen_summary.head()
      Existing generator  ... Connection cost_Partial outage_Technology
    0          Bayswater  ...                            Black Coal NSW
    1            Eraring  ...                            Black Coal NSW
    2           Mt Piper  ...                            Black Coal NSW
    3      Vales Point B  ...                            Black Coal NSW
    4          Callide B  ...                            Black Coal QLD
    <BLANKLINE>
    [5 rows x 28 columns]

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
        4. Reset the DataFrame index
        """
        df_cleaned = df_initial.iloc[header_rows_in_table:, :]
        df_cleaned.columns = new_headers
        df_cleaned = df_cleaned.ffill(axis=1).reset_index(drop=True)
        return df_cleaned

    def _skip_rows_in_dataframe(
        df: pd.DataFrame, config_skip_rows: Union[int, List[int]], last_header_row: int
    ) -> pd.DataFrame:
        """
        Drop rows specified by `skip_rows` by applying an offset from the header and
        dropping based on index values
        """
        df_reset_index = df.reset_index(drop=True)
        if isinstance(config_skip_rows, int):
            skip_rows = [config_skip_rows - last_header_row - 1]
        else:
            skip_rows = np.subtract(config_skip_rows, last_header_row + 1)
        dropped = df_reset_index.drop(index=skip_rows).reset_index(drop=True)
        return dropped

    if type(table.header_rows) is int:
        df = pd.read_excel(
            workbook_file,
            sheet_name=table.sheet_name,
            header=(table.header_rows - 1),
            usecols=table.column_range,
            nrows=(table.end_row - table.header_rows),
        )
        if table.skip_rows:
            df = _skip_rows_in_dataframe(df, table.skip_rows, table.header_rows)
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
        # start processing multiple header rows
        header_rows_in_table = table.header_rows[-1] - table.header_rows[0]
        initial_header = pd.Series(df_initial.columns)
        ffilled_initial_header = _ffill_highest_header(initial_header)
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
            series[series != ""] = "_" + series[series != ""]
        merged_headers = ffilled_initial_header.str.cat(filled_headers)
        df_cleaned = _build_cleaned_dataframe(
            df_initial, header_rows_in_table, merged_headers
        )
        if table.skip_rows:
            df_cleaned = _skip_rows_in_dataframe(
                df_cleaned, table.skip_rows, table.header_rows[-1]
            )
        return df_cleaned
