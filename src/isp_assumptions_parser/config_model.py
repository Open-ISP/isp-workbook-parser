import yaml

from pydantic import BaseModel
from typing import List
from pathlib import Path


class Table(BaseModel):
    name: str
    sheet_name: str
    header_rows: int | List[int]
    end_row: int
    column_range: str
    header_names: List[str]


def load_yaml(path: Path) -> List[Table]:
    with open(path, "r") as f:
        config = yaml.safe_load(f)
        f.close()
    tables = [Table(name=name, **config[name]) for name in config]
    return tables
