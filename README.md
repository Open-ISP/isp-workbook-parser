# AEMO Integrated System Plan Assumptions Workbook Parser

A Python package for reading data from the Inputs, Assumptions and Scenarios Report (IASR) Microsoft Excel workbook
published by the Australian Energy Market Operator for use in their Integrated System Plan modelling.

## Table of contents

- [How the package works](#how-the-package-works)
- [Table configurations](#table-configurations)
- [Examples](#examples)
    - [Bulk export](#bulk-export)
    - [List tables with configuration files](#list-tables-with-configuration-files)
    - [Get table as DataFrame](#get-table-as-dataframe)
    - [Get table with custom configuration](#get-table-with-custom-configuration)


## How the package works

1. Load a workbook using `Parser` (see examples below).
   - While we do not include workbooks with the package distribution, you can find the versions for which table configurations are written within `workbooks/<version>`.
2. Table configuration files for data tables are located in `src/config/<version>`
   - These specify the name, location, columns and data range of tables to be extracted from a particular workbook version.
   - These are included with the package distributions.
3. `Parser` loads the MS Excel workbook and, by default, will check if the version of the workbook is supported by seeing if configuration files are included in the package for that version.
4. If they are, `Parser` can use these configuration files to parse the data tables and save them as CSVs.

> [!NOTE]
> This package makes some opinionated decisions when processing tables with multiple header rows, including how tables are reduced to a single header and how data in merged cells is handled. For more detail, refer to the docstring and code in [`read_table.py`](src/isp_workbook_parser/read_table.py).

## Table configurations

Tables are defined in the configuration files using the following attributes:

- `name`: the table name
- `sheet_name`: the sheet where the table is located
   - N.B. there may be spaces at the end of sheet names in the workbook
- `header_rows`: this specifies the row(s) with table column names
   - If there is a single row of table column names, then provide a single number (i.e. an `int` like `6`)
   - If the table header is defined over multiple rows, then provide a list of row numbers sorted in ascending order (e.g. `[6, 7, 8]`)
- `end_row`: the last row of table data
- `column_range`: the column range of the table in alphabetical/Excel format, e.g. `"B:F"`

## Examples

### Bulk export

Export all the data tables the parser has a config file for to CSV files.

```python
from isp_workbook_parser import Parser

workbook = Parser("<path/to/workbook>/2024-isp-inputs-and-assumptions-workbook.xlsx")

workbook.save_tables('<path/to/output directory>')
```

### List tables with configuration files

List all the tables the package has a configuration file for (for the given workbook version).

```python
from isp_workbook_parser import Parser

workbook = Parser("<path/to/workbook>/2024-isp-inputs-and-assumptions-workbook.xlsx")

workbook.get_table_names()
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

## License
`isp-workbook-parser` was created as a part of the [OpenISP project](https://github.com/Open-ISP). It is licensed under the terms of [GNU GPL-3.0-or-later](LICENSE) licences.
