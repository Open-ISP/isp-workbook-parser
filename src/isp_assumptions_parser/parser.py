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
    """

    def __init__(self, file_path: str | Path):
        self.file_path = self._make_path_object(file_path)
        self.file = pd.ExcelFile(self.file_path)
        self.openpyxl_file = openpyxl.load_workbook(self.file_path)
        self.workbook_version = self._get_version()
        self._check_version_is_supported()
        self.table_names = table_configs[self.workbook_version].keys()

    @staticmethod
    def _make_path_object(path: str | Path):
        if not isinstance(path, Path):
            path = Path(path)
        return path

    def _get_version(self):
        sheet = self.openpyxl_file["Change Log"]
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

    def _check_data_ends_where_expected(
        self, tab: str, end_row: int, range: str, name: str
    ):
        first_column = range.split(":")[0]
        first_col_index = openpyxl.utils.column_index_from_string(first_column)
        second_col_index = first_col_index + 1
        # We check that value in the second column is blank because sometime the row after the first column will
        # contain notes on the data.
        value_in_second_column_after_last_row = (
            self.openpyxl_file[tab].cell(row=end_row + 1, column=second_col_index).value
        )
        if value_in_second_column_after_last_row is not None:
            if name is None:
                error_message = "There is data in the row after the defined table end."
            else:
                error_message = f"There is data in the row after the defined table end for table {name}."

            raise TableConfigError(error_message)

    @staticmethod
    def _check_for_over_run_into_another_table(data: pd.DataFrame, name: str):
        if data[data.columns[0]].isna().any():
            if name is None:
                error_message = (
                    "The first column of the table contains na values indicating the table end row is "
                    "incorrectly specified."
                )
            else:
                error_message = (
                    f"The first column of the table {name} contains na values indicating the table end "
                    f"row is incorrectly specified."
                )

            raise TableConfigError(error_message)

    @staticmethod
    def _check_for_over_run_into_notes(data: pd.DataFrame, name: str):
        notes_sub_strings = ["Notes:", "Note:", "Source:", "Sources:"]
        for sub_string in notes_sub_strings:
            if data[data.columns[0]].str.contains(sub_string).any():
                if name is None:
                    error_message = f"The first column of the table contains the sub string '{sub_string}'."
                else:
                    error_message = f"The first column of the table {name} contains the sub string '{sub_string}'."

                raise TableConfigError(error_message)

    @staticmethod
    def _check_last_column_not_empty(data: pd.DataFrame, name: str):
        if data[data.columns[-1]].isna().all():
            if name is None:
                error_message = "The last column of the table is empty."
            else:
                error_message = f"The last column of the table {name} is empty."

            raise TableConfigError(error_message)

    def _check_for_missed_column_on_right_hand_side_of_table(
        self, tab: str, start_row: int, end_row: int, range: int, name: str
    ):
        last_column = range.split(":")[1]
        last_col_index = openpyxl.utils.column_index_from_string(last_column)
        column_next_to_last_column = openpyxl.utils.get_column_letter(
            last_col_index + 1
        )
        column_next_to_last_column = (
            column_next_to_last_column + ":" + column_next_to_last_column
        )
        range_error = False
        try:
            data = self._read_table(
                tab,
                start_row,
                end_row,
                column_next_to_last_column,
            )
            if not data[data.columns[0]].isna().all():
                range_error = True
        except pd.errors.ParserError:
            range_error = False

        if range_error:
            if name is None:
                error_message = "There is data in the column adjacent to the last column in the table."
            else:
                error_message = f"There is data in the column adjacent to the last column in the table {name}."

            raise TableConfigError(error_message)

    def _read_table(self, tab: str, start_row: int, end_row: int, range: str):
        nrows = end_row - start_row
        data = pd.read_excel(
            io=self.file,
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

    def get_table_from_config(
        self, table_config, name: str = None, config_checks: bool = True
    ) -> pd.DataFrame:
        """Retrieves a table from the assumptions workbook using the config provided and returns as pd.DataFrame.

        Examples:

        Args:
            table_config:
            name: name of the table as a string, used when processing multiple tables so errors raised when checking
                the config can specify which table the check failed for.
            config_checks: A boolean that specifies whether to check the table config by checking if the data
                starts and ends where expected and the workbook header matches the config header.

        """
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
        if config_checks:
            self._check_data_ends_where_expected(
                table_config["tab"],
                table_config["end_row"],
                table_config["range"],
                name=name,
            )
            self._check_for_missed_column_on_right_hand_side_of_table(
                table_config["tab"],
                table_config["start_row"],
                table_config["end_row"],
                table_config["range"],
                name,
            )
            self._check_for_over_run_into_another_table(data, name=name)
            self._check_for_over_run_into_notes(data, name=name)
            self._check_last_column_not_empty(data, name=name)
        return data

    def get_table(self, table_name: str, config_checks: bool = True) -> pd.DataFrame:
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
            config_checks: A boolean that specifies whether to check the tabe config by checking if the data
                starts and ends where expected and the workbook header matches the config header.
        """
        if not isinstance(table_name, str):
            ValueError("The parameter table_name must be provided as a string.")

        if table_name not in self.table_names:
            ValueError(
                "The table_name provided is not in the config for this workbook version."
            )

        table_config = table_configs[self.workbook_version][table_name]
        data = self.get_table_from_config(
            table_config, name=table_name, config_checks=config_checks
        )
        return data

    def save_tables(
        self,
        directory: str | Path,
        tables: list[str] | str = "all",
        config_checks: bool = True,
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
            config_checks: A boolean that specifies whether to check the tabe config by checking if the data
                starts and ends where expected and the workbook header matches the config header.

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
            table = self.get_table(table_name, config_checks=config_checks)
            save_path = directory / Path(f"{table_name}.csv")
            table.to_csv(save_path)


class TableConfigError(Exception):
    """Raise for table configuration failing check."""
