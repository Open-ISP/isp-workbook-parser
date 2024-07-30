from pathlib import Path

import pandas as pd
import openpyxl

from .assumptions_workbook_settings import table_configs


class Parser:
    """Extracts ISP inputs and assumptions data from the IASR workbbook.

    The primary intend usage is to extract all the tables and save as csv files using the save_tables method, but
    inidividual tables can be extracted as pd.DataFrames using get_table, and tables without pre-defined config can
    be extracted with user specified config using get_table_from_config. For a list of tables with pre-defined config
    compatible the version of the workbook provided use get_table_names.

    Examples

    Create a Parser instance for a particular workbook (will also check config is available for workbook version).

    >>> workbook = Parser("workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")

    Save all the tables with available config to the directory example_output as csv files.

    >>> workbook.save_tables('example_output')
                        Generator  ... Auxiliary load (%)
    2                   Bayswater  ...     Black Coal NSW
    3                     Eraring  ...     Black Coal NSW
    4                    Mt Piper  ...     Black Coal NSW
    5               Vales Point B  ...     Black Coal NSW
    6                   Callide B  ...     Black Coal QLD
    ..                        ...  ...                ...
    247  Stockyard Hill Wind Farm  ...               Wind
    248          Waubra Wind Farm  ...               Wind
    249    Yaloak South Wind Farm  ...               Wind
    250          Yambuk Wind Farm  ...               Wind
    251          Yendon Wind Farm  ...               Wind
    <BLANKLINE>
    [250 rows x 27 columns]
    """

    def __init__(self, file_path):
        self.file_path = self._make_path_object(file_path)
        self.file = pd.ExcelFile(self.file_path)
        self.workbook_version = self._get_version()
        self._check_version_is_supported()
        self.table_names = table_configs[self.workbook_version].keys()

    @staticmethod
    def _make_path_object(path):
        if not isinstance(path, Path):
            path = Path(path)
        return path

    def _get_version(self):
        sheet = openpyxl.load_workbook(self.file_path)["Change Log"]
        last_value = None
        for cell in sheet["B"]:
            if cell.value is not None:
                last_value = cell.value
        return str(last_value)

    def _check_version_is_supported(self):
        if self.workbook_version not in table_configs.keys():
            raise ValueError(
                f"The workbook version {self.workbook_version} is not supported."
            )

    def _check_data_end_where_expected(self, end_row, range):
        pass

    def _read_table(self, tab, start_row, end_row, range):
        nrows = end_row - start_row + 1
        data = pd.read_excel(
            self.file,
            sheet_name=tab,
            usecols=range,
            skiprows=start_row - 1,
            nrows=nrows,
        )
        return data

    @staticmethod
    def _clean_table(data, column_name_corrections, junk_rows_at_start):
        data = data.rename(columns=column_name_corrections)
        data = data.iloc[junk_rows_at_start:]
        return data

    def get_table_names(self) -> list[str]:
        """Returns the list of tables that there is config for extracting from the workbook version provided.

        Examples:
        >>> workbook = Parser("workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")

        >>> workbook.get_table_names()

        Returns:
            List of the tables that there is configuration information for extracting from the workbook.
        """
        return self.table_names

    def get_table_from_config(self, table_config):
        """Will write docs once we have finalised the config data format."""
        data = self._read_table(
            table_config["tab"],
            table_config["start_row"],
            table_config["end_row"],
            table_config["range"],
        )
        data = self._clean_table(
            data,
            table_config["column_name_corrections"],
            table_config["junk_rows_at_start"],
        )
        return data

    def get_table(self, table_name: str) -> pd.DataFrame:
        """Retrieves a table from the assumptions workbook and returns as pd.DataFrame.

        Examples
        >>> workbook = Parser("workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")

        >>> workbook.get_table('existing_generators_summary')
                            Generator  ... Auxiliary load (%)
        2                   Bayswater  ...     Black Coal NSW
        3                     Eraring  ...     Black Coal NSW
        4                    Mt Piper  ...     Black Coal NSW
        5               Vales Point B  ...     Black Coal NSW
        6                   Callide B  ...     Black Coal QLD
        ..                        ...  ...                ...
        247  Stockyard Hill Wind Farm  ...               Wind
        248          Waubra Wind Farm  ...               Wind
        249    Yaloak South Wind Farm  ...               Wind
        250          Yambuk Wind Farm  ...               Wind
        251          Yendon Wind Farm  ...               Wind
        <BLANKLINE>
        [250 rows x 27 columns]

        Args:
            table_name: string specifying the table to retrieve.
        """
        if not isinstance(table_name, str):
            ValueError("The parameter table_name must be provided as a string.")

        if table_name not in self.table_names:
            ValueError(
                "The table_name provided is not in the config for this workbook version."
            )

        table_config = table_configs[self.workbook_version][table_name]
        data = self.get_table_from_config(table_config)
        return data

    def save_tables(
        self, directory: str | Path, tables: list[str] | str = "all"
    ) -> None:
        """Saves tables from the assumptions workbook to the specified directory as CSV files.

        Examples:
        >>> workbook = Parser("workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")

        >>> workbook.save_tables(
        ... directory="example_output",
        ... tables=['existing_generators_summary']
        ... )

        Args:
            tables: A list of strings specifying which tables to extract from the workbook and save, or the
                str 'all', which will result in all the tables the isp-assumptions-parser has config for being
                saved.
            directory: A str specifying the path to the directory or a pathlib Path object.

        Returns:
            None
        """
        directory = self._make_path_object(directory)

        if not directory.is_dir():
            ValueError("The path provided is not a directory.")

        if not (isinstance(tables, str) or isinstance(tables, list)):
            ValueError("The parameter tables must be provided as str or list[str].")

        if isinstance(tables, str) and tables != "all":
            ValueError(
                "If the parameter tables is provided as a str it must \n",
                f"have the value 'all' but '{tables}' was provided.",
            )

        if tables == "all":
            tables = self.get_table_names()

        for table_name in tables:
            table = self.get_table(table_name)
            save_path = directory / Path(f"{table_name}.csv")
            table.to_csv(save_path)
