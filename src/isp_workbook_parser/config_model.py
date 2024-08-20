import yaml

from pydantic import BaseModel
from typing import List
from pathlib import Path


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
    {'wind_high_capacity_factors': TableConfig(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R')}

    Attributes:
        name: the table name
        sheet_name: the sheet where the table is located.
        header_rows: an `int` specifying the row with the table column names, or if the table header is
            defined over multiple rows, then a list of the row numbers sorted in ascending order.
        end_row: the last row of table data.
        column_range: the columns over which the table is defined in the alphabetical format, i.e. 'B:F'
    """

    name: str
    sheet_name: str
    header_rows: int | List[int]
    end_row: int
    column_range: str


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

    When read using `load_yaml` it will be converted to a dictionary contain `TableConfig` instances:

    >>> print(load_yaml(path_to_yaml))
    {'wind_high_capacity_factors': TableConfig(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R')}

    Args:
        path: pathlib Path instance specifying the location of the YAML file.

    """
    with open(path, "r") as f:
        config = yaml.safe_load(f)
        f.close()
    tables = {name: TableConfig(name=name, **config[name]) for name in config}
    return tables
