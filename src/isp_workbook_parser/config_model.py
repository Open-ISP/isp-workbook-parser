from pathlib import Path
from typing import List, Optional

import yaml
from pydantic import BaseModel


class TableConfig(BaseModel):
    """A `Pydantic` class for storing the location of a table within an Excel Workbook, which is referred to as a table config throughout this package.

    The `Pydantic` class verifies the type of each element of the configuration.

    Examples:

    A TableConfig instance can be manually defined:

    >>> table_config = TableConfig(
    ... name='existing_generators_summary',
    ... sheet_name='Summary Mapping',
    ... header_rows=[4, 5, 6],
    ... end_row=258,
    ... column_range="B:AC",
    ... )

    Or created from YAML file, which can contain one or many table config definitions:

    >>> table_configs = load_yaml(Path("src/isp_table_configs/6.0/capacity_factors.yaml"))

    >>> print(table_configs)
    {'wind_high_capacity_factors': TableConfig(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'wind_medium_capacity_factors': TableConfig(name='wind_medium_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='T:AJ', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'solar_pv_capacity_factors': TableConfig(name='solar_pv_capacity_factors', sheet_name='Capacity Factors ', header_rows=[50, 51, 52], end_row=91, column_range='B:R', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'solar_thermal_15hrstorage_capacity_factors': TableConfig(name='solar_thermal_15hrstorage_capacity_factors', sheet_name='Capacity Factors ', header_rows=[50, 51, 52], end_row=91, column_range='T:AJ', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'offshore_wind_fixed_capacity_factors': TableConfig(name='offshore_wind_fixed_capacity_factors', sheet_name='Capacity Factors ', header_rows=[93, 94, 95], end_row=102, column_range='B:R', skip_rows=102, columns_with_merged_rows=None, forward_fill_values=True), 'offshore_wind_floating_capacity_factors': TableConfig(name='offshore_wind_floating_capacity_factors', sheet_name='Capacity Factors ', header_rows=[93, 94, 95], end_row=102, column_range='T:AJ', skip_rows=102, columns_with_merged_rows=None, forward_fill_values=True)}

    Attributes:
        name: the table name
        sheet_name: the sheet where the table is located.
        header_rows: an `int` specifying the row with the table column names, or if the table header is
            defined over multiple rows, then a list of the row numbers sorted in ascending order.
        end_row: the last row of table data.
        column_range: the columns over which the table is defined in the alphabetical format, i.e. 'B:F'
        skip_rows: optional, an `int` specifying a row to skip, or a list of `int` corresponding to
            row numbers to skip.
        columns_with_merged_rows: optional, a `str` specifying a column with merged rows
            in alphabetical format (e.g. "B") or a list of `str` corresponding to columns
            in alphabetical format with merged rows (e.g. ["B", "D"]).
        forward_fill_values: optional, a `bool` specifying whether table values should be
            forward filled. Default `True` since this functionality is needed to handle
            merged cells. Should be set to `False` where there are empty columns
    """

    name: str
    sheet_name: str
    header_rows: int | List[int]
    end_row: int
    column_range: str
    skip_rows: Optional[int | List[int]] = None
    columns_with_merged_rows: Optional[str | List[str]] = None
    forward_fill_values: bool = True


def load_yaml(path: Path) -> dict[str, TableConfig]:
    """Loads the YAML file specified by the path returning a dict of `TableConfig`s.

    Each table config defined in a YAML file is converted to a `TableConfig` and stored in the dictionary using its name
    as the key value.

    Examples:

    >>> path_to_yaml = Path("src/isp_table_configs/6.0/capacity_factors.yaml")

    The contents of the YAML file should look like:

    >>> print(open(path_to_yaml).read())
    wind_high_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [7, 8, 9]
      end_row: 48
      column_range: "B:R"
    <BLANKLINE>
    wind_medium_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [7, 8, 9]
      end_row: 48
      column_range: "T:AJ"
    <BLANKLINE>
    solar_pv_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [50, 51, 52]
      end_row: 91
      column_range: "B:R"
    <BLANKLINE>
    solar_thermal_15hrstorage_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [50, 51, 52]
      end_row: 91
      column_range: "T:AJ"
    <BLANKLINE>
    offshore_wind_fixed_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [93, 94, 95]
      end_row: 102
      column_range: "B:R"
      skip_rows: 102
    <BLANKLINE>
    offshore_wind_floating_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [93, 94, 95]
      end_row: 102
      column_range: "T:AJ"
      skip_rows: 102
    <BLANKLINE>

    When read using `load_yaml` it will be converted to a dictionary contain `TableConfig` instances:

    >>> print(load_yaml(path_to_yaml))
    {'wind_high_capacity_factors': TableConfig(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'wind_medium_capacity_factors': TableConfig(name='wind_medium_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='T:AJ', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'solar_pv_capacity_factors': TableConfig(name='solar_pv_capacity_factors', sheet_name='Capacity Factors ', header_rows=[50, 51, 52], end_row=91, column_range='B:R', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'solar_thermal_15hrstorage_capacity_factors': TableConfig(name='solar_thermal_15hrstorage_capacity_factors', sheet_name='Capacity Factors ', header_rows=[50, 51, 52], end_row=91, column_range='T:AJ', skip_rows=None, columns_with_merged_rows=None, forward_fill_values=True), 'offshore_wind_fixed_capacity_factors': TableConfig(name='offshore_wind_fixed_capacity_factors', sheet_name='Capacity Factors ', header_rows=[93, 94, 95], end_row=102, column_range='B:R', skip_rows=102, columns_with_merged_rows=None, forward_fill_values=True), 'offshore_wind_floating_capacity_factors': TableConfig(name='offshore_wind_floating_capacity_factors', sheet_name='Capacity Factors ', header_rows=[93, 94, 95], end_row=102, column_range='T:AJ', skip_rows=102, columns_with_merged_rows=None, forward_fill_values=True)}

    Args:
        path: pathlib Path instance specifying the location of the YAML file.

    """
    with open(path, "r") as f:
        config = yaml.safe_load(f)
        f.close()
    if config is not None:
        tables = {name: TableConfig(name=name, **config[name]) for name in config}
    else:
        tables = {}
    return tables
