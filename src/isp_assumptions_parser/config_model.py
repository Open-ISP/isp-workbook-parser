import yaml

from pydantic import BaseModel
from typing import List
from pathlib import Path


class Table(BaseModel):
    """A pydantic class for storing the location of a table within an Excel Workbook, referred to as the table config.

    The pydantic class is used so that the type of each element of the configuration is validated.

    Examples:

    A Table instance can be manually defined:

    >>> table_config = Table(
    ... name='existing_generators_summary',
    ... sheet_name='Summary Mapping',
    ... header_rows=[4, 5, 6],
    ... end_row=256,
    ... column_range="B:AC",
    ... )

    Or created from YAML file, which can contain one or many table config definitions:

    >>> table_configs = load_yaml(Path("config/5.2/capacity_factors.yaml"))

    >>> print(table_configs)
    {'wind_high_capacity_factors': Table(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R')}

    Attributes:
        name: str, the table name
        sheet_name: str, the sheet where the table is located.
        header_rows: int | List[int], an int specifying the row with the table column names, or if the table header is
            defined over multiple rows then a list of these rows sorted in ascending order.
        end_row: int, the last row of table data.
        column_range: str, the columns over which the table is defined in the alphabetical format, i.e. 'B:F'
    """

    name: str
    sheet_name: str
    header_rows: int | List[int]
    end_row: int
    column_range: str


def load_yaml(path: Path) -> dict[Table]:
    """Loads the YAML file specified by the path returning a dict of Table objects.

    Each table config defined in YAML file is converted to Table object and stored in the dictionary using its name
    as the key value.

    Examples:

    >>> path_to_yaml = Path("config/5.2/capacity_factors.yaml")

    The contents of the YAML file should look like:

    >>> print(open(path_to_yaml).read())
    wind_high_capacity_factors:
      sheet_name: "Capacity Factors "
      header_rows: [7, 8, 9]
      end_row: 48
      column_range: "B:R"
    <BLANKLINE>

    When read using load_yaml it will be converted to a dictionary contain Table instances:

    >>> print(load_yaml(path_to_yaml))
    {'wind_high_capacity_factors': Table(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R')}

    Args:
        path: pathlib Path instance specifying the location of the YAML file.

    """
    with open(path, "r") as f:
        config = yaml.safe_load(f)
        f.close()
    tables = {name: Table(name=name, **config[name]) for name in config}
    return tables
