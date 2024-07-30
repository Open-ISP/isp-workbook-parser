from pathlib import Path

import pandas as pd
import openpyxl

from assumptions_workbook_settings import table_configs


class AssumptionsWorkbookInterface:
    """
    Examples
    >>> workbook = AssumptionsWorkbookInterface("workbooks/5.2/2023 IASR Assumptions Workbook.xlsx")

    >>> workbook.get_data('existing_generators_summary')
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
        if not isinstance(file_path, Path):
            self.file_path = Path(file_path)
        else:
            self.file_path = file_path

        self.file = pd.ExcelFile(self.file_path)
        self.workbook_version = self._get_version()
        self._check_version_is_supported()

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

    def get_data(self, table_name):
        table_config = table_configs[self.workbook_version][table_name]
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
