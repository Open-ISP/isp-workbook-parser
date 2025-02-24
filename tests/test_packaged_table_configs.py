from pathlib import Path

import pytest

from isp_workbook_parser import Parser

workbook_path = Path("workbooks")


@pytest.mark.parametrize("workbook_version_folder", workbook_path.iterdir())
def test_packaged_table_configs_for_each_version(workbook_version_folder: Path):
    xl_file = [file for file in workbook_version_folder.glob("[!.]*.xls*")]
    assert len(xl_file) == 1, (
        f"There should only be one Excel workbook in each version sub-directory, got {xl_file}"
    )
    workbook_name = xl_file.pop()
    workbook = Parser(workbook_name)

    # Check that configs don't look at the same tables.
    configs = workbook.table_configs
    sheet_and_header_combos = [
        c.sheet_name + str(c.header_rows) + c.column_range for c in configs.values()
    ]
    table_names = [c.name for c in configs.values()]
    duplicate_configs = []
    for index, value in enumerate(sheet_and_header_combos):
        if sheet_and_header_combos.count(value) > 1:
            duplicate_configs.append(table_names[index])
    if len(duplicate_configs) > 0:
        print(duplicate_configs)
    assert len(duplicate_configs) == 0

    save_dir = Path(f"example_output/{workbook.workbook_version}")
    save_dir.mkdir(parents=True, exist_ok=True)
    error_tables = {}
    for table_name in workbook.table_configs:
        try:
            table = workbook.get_table(table_name)
            save_path = save_dir / Path(f"{table_name}.csv")
            table.to_csv(save_path, index=False)
        except Exception as e:
            error_tables[table_name] = e
    if error_tables:
        error_str = ""
        for key in error_tables:
            error_str += str(error_tables[key]) + "\n"
        raise TableLoadError(error_str)


class TableLoadError(Exception):
    """Exception to throw if table loading fails"""
