# Copyright (C) 2026 University of New South Wales
#
# This file is free software; you can redistribute it and/or modify it
# under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.

import pytest

from isp_workbook_parser import Parser
from isp_workbook_parser.parser import TableConfigError


def test_end_row_not_on_sheet_throws_error():
    with pytest.raises(TableConfigError):
        Parser(
            "tests/test_data/2024-isp-inputs-and-assumptions-workbook-missing-sheets.xlsx"
        )
