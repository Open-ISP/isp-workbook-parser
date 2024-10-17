import pandas as pd
import pytest

from isp_workbook_parser import Parser


@pytest.fixture(scope="session")
def workbook_v6() -> Parser:
    workbook = Parser("workbooks/6.0/2024-isp-inputs-and-assumptions-workbook.xlsx")
    return workbook


@pytest.fixture(scope="module")
def sample_series():
    return pd.Series(
        [
            "First line\nSecond line",
            "This  is  a  test",
            "Value with *",
            "SomeUnitA5",
            "An actual footnote1",
            "  leading and trailing  ",
            "1,234,567",
            "50.0 - note",
            "35.5 (comment)",
            "42.0 - note (additional info)",
            "999.9 (info) - second note",
            "2024-25 data for year",
        ]
    )
