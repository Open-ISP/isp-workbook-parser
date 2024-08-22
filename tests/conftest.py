import pytest

from isp_workbook_parser import Parser


@pytest.fixture(scope="session")
def workbook_v6() -> Parser:
    workbook = Parser("workbooks/6.0/2024-isp-inputs-and-assumptions-workbook.xlsx")
    return workbook
