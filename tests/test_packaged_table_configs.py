import pytest
from isp_workbook_parser import Parser
from pathlib import Path

workbook_path = Path("workbooks")


@pytest.mark.parametrize("workbook_version_folder", workbook_path.iterdir())
def test_packaged_table_configs_for_each_version(workbook_version_folder: Path):
    xl_file = [file for file in workbook_version_folder.glob("*.xls*")]
    assert (
        len(xl_file) == 1
    ), f"There should only be one Excel workbook in each version sub-directory, got {xl_file}"
    workbook_name = xl_file.pop()
    workbook = Parser(workbook_name)
    configs = workbook.get_table_names()
    for config in configs:
        workbook.get_table(config, config_checks=True)
