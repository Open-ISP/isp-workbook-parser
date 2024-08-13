import yaml
from pprint import pprint

from pydantic import BaseModel
from typing import List
from pathlib import Path


class Table(BaseModel):
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

    >>> pprint(load_yaml(path_to_yaml))
    {'wind_high_capacity_factors': Table(name='wind_high_capacity_factors', sheet_name='Capacity Factors ', header_rows=[7, 8, 9], end_row=48, column_range='B:R')}

    Args:
        path: pathlib Path instance specifying the location of the YAML file.

    """
    with open(path, "r") as f:
        config = yaml.safe_load(f)
        f.close()
    tables = {name: Table(name=name, **config[name]) for name in config}
    return tables
