import pandas as pd

from .config_model import TableConfig as TableConfig
from .config_model import load_yaml as load_yaml
from .parser import Parser
from .read_table import read_table as read_table

__all__ = ["Parser", "TableConfig", "load_yaml", "read_table"]

pd.set_option("future.no_silent_downcasting", True)
