from pydantic import BaseModel
from typing import List


class Table(BaseModel):
    name: str
    sheet_name: str
    header_row: int
    end_row: int
    column_range: str
    header_names: List[str]
