# Copyright (C) 2026 University of New South Wales
#
# This file is free software; you can redistribute it and/or modify it
# under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.

import pandas as pd

from .config_model import TableConfig, load_yaml
from .parser import Parser
from .read_table import read_table

__all__ = ["Parser", "TableConfig", "load_yaml", "read_table"]

pd.set_option("future.no_silent_downcasting", True)  # noqa: FBT003
