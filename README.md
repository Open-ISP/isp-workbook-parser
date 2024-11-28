# AEMO Integrated System Plan Assumptions Workbook Parser
[![PyPI version](https://badge.fury.io/py/isp-workbook-parser.svg)](https://badge.fury.io/py/isp-workbook-parser)
[![Continuous Integration and Deployment](https://github.com/Open-ISP/isp-workbook-parser/actions/workflows/cicd.yml/badge.svg)](https://github.com/Open-ISP/isp-workbook-parser/actions/workflows/cicd.yml)
[![codecov](https://codecov.io/github/Open-ISP/isp-workbook-parser/graph/badge.svg?token=BUGWITKZV1)](https://codecov.io/github/Open-ISP/isp-workbook-parser)
[![pre-commit.ci status](https://results.pre-commit.ci/badge/github/Open-ISP/isp-workbook-parser/main.svg)](https://results.pre-commit.ci/latest/github/Open-ISP/isp-workbook-parser/main)
[![UV](https://camo.githubusercontent.com/4ab8b0cb96c66d58f1763826bbaa0002c7e4aea0c91721bdda3395b986fe30f2/68747470733a2f2f696d672e736869656c64732e696f2f656e64706f696e743f75726c3d68747470733a2f2f7261772e67697468756275736572636f6e74656e742e636f6d2f61737472616c2d73682f75762f6d61696e2f6173736574732f62616467652f76302e6a736f6e)](https://github.com/astral-sh/uv)

A Python package for reading data from the Inputs, Assumptions and Scenarios Report (IASR) Microsoft Excel workbook
published by the Australian Energy Market Operator for use in their Integrated System Plan modelling.

## Table of contents

- [Installation](#installation)
- [How the package works](#how-the-package-works)
- [Table configurations](#table-configurations)
- [Examples](#examples)
    - [Bulk export](#bulk-export)
    - [List tables with configuration files](#list-tables-with-configuration-files)
    - [Get table as DataFrame](#get-table-as-dataframe)
    - [Get table with custom configuration](#get-table-with-custom-configuration)
- [Contributing](#contributing)
- [License](#license)

## Installation

```bash
pip install isp-workbook-parser
```

## How the package works

1. Load a workbook using `Parser` (see examples below).
   - While we do not include workbooks with the package distribution, you can find the versions for which table configurations are written within `workbooks/<version>`.
2. Table configuration files for data tables are located in `src/config/<version>`
   - These specify the name, location, columns and data range of tables to be extracted from a particular workbook version. Optionally, rows to skip and not read in (e.g. where AEMO has formatted a row with a strike through to indicate that the data is no longer being used) and columns with merged rows can also be specified and handled.
   - These are included with the package distributions.
3. `Parser` loads the MS Excel workbook and, by default, will check if the version of the workbook is supported by seeing if configuration files are included in the package for that version.
4. If they are, `Parser` can use these configuration files to parse the data tables and save them as CSVs.

> [!NOTE]
> This package makes some opinionated decisions when processing tables. For example,
> multiple header row tables are reduced to a single header, data in merged cells is inferred from surrounding cells,
> and notes and footnotes are dropped (amonst other ways in which the data is sanitised).
> For more detail, refer to the docstring and code in [`read_table.py`](https://github.com/Open-ISP/isp-workbook-parser/blob/main/src/isp_workbook_parser/read_table.py)
> and [`sanitisers.py`](https://github.com/Open-ISP/isp-workbook-parser/blob/main/src/isp_workbook_parser/sanitisers.py).

## Table configurations

<details>
<summary>Table configuration file attributes</summary>
<br>

- `name`: the table name
- `sheet_name`: the sheet where the table is located
  - N.B. there may be spaces at the end of sheet names in the workbook
- `header_rows`: this specifies the Excel row(s) with table column names
  - A single row of table column names (e.g. `6`)
  - Or a list of row numbers for the table header sorted in ascending order (e.g. `[6, 7, 8]`)
- `end_row`: the last row of table data
- `column_range`: the Excel column range of the table in alphabetical/Excel format, e.g. `"B:F"`
- `skip_rows`: optional, Excel row(s) in the table that should not be read in
  - A single row (e.g. `15`)
  - Or a list of rows  (e.g. `[15, 16]`)
- `columns_with_merged_rows`: optional, Excel column(s) with merged rows
  - A single column in alphabetical format (e.g. `"B"`),
  - Or a list of columns in alphabetical format (e.g. `["B", "D"]`)
- `forward_fill_values`: optional, specifies whether table values should be forward filled
  - Default `True` to handle merged cells in tables
  - Should be set to `False` where there are empty columns

</details>

### Adding table configuration files to this package

Refer to the [contributing instructions](https://github.com/Open-ISP/isp-workbook-parser/blob/main/CONTRIBUTING.md) for details on how to contribute table configuration (YAML) files to this repository and package.

## Examples

### Bulk export

Export all the data tables the package has a config file for to CSV files.

```python
from isp_workbook_parser import Parser

workbook = Parser("<path/to/workbook>/2024-isp-inputs-and-assumptions-workbook.xlsx")

workbook.save_tables('<path/to/output directory>')
```

### List tables with configuration files

Return a dictionary of table names, with lists of tables names stored under a key which is their sheet name in the workbook.
For a given workbook version, this only returns tables the package has a configuration file for.

```python
from isp_workbook_parser import Parser

workbook = Parser("<path/to/workbook>/2024-isp-inputs-and-assumptions-workbook.xlsx")

names = workbook.get_table_names()

names['Build limits']
```

### Get table as DataFrame

Get a single table as a pandas `DataFrame`.

```python
from isp_workbook_parser import Parser

workbook = Parser("<path/to/workbook>/2024-isp-inputs-and-assumptions-workbook.xlsx")

table = workbook.get_table("retirement_costs")
```

### Get table with custom configuration

Get a table by directly providing the table config.

```python
from isp_workbook_parser import Parser, TableConfig

workbook = Parser("<path/to/workbook>/2024-isp-inputs-and-assumptions-workbook.xlsx")

table_config = TableConfig(
  name="table_name",
  sheet_name="sheet_name",
  header_rows=5,
  end_row=21,
  column_range="B:J",
)

workbook.get_table_from_config(table_config)
```

## Contributing

Interested in contributing to the source code or adding table configurations? Check out the [contributing instructions](https://github.com/Open-ISP/isp-workbook-parser/blob/main/CONTRIBUTING.md), which also includes steps to install `isp-workbook-parser` for development.

Please note that this project is released with a [Code of Conduct](https://github.com/Open-ISP/isp-workbook-parser/blob/main/CONDUCT.md). By contributing to this project, you agree to abide by its terms.

## License

`isp-workbook-parser` was created as a part of the [OpenISP project](https://github.com/Open-ISP). It is licensed under the terms of [GNU GPL-3.0-or-later](https://github.com/Open-ISP/isp-workbook-parser/blob/main/LICENSE) licences.
