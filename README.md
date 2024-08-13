# AEMO Integrated System Plan Assumptions Parser

A Python package for reading data from the Inputs, Assumptions and Scenarios Report (IASR) MS Excel workbook
published by the Australian Energy Market Operator for use in their Integrated System Plan modelling.

## How it works

1. Workbooks versions are saved in `workbooks/<version>`
2. YAML configuration files for data tables are located in `config/<version>`
   - These specify the name, location, columns and data range of tables to be extracted from a particular workbook version
3. The parser loads the MS Excel workbook and uses the available config files to parse data tables and save the as CSVs

## Examples

## Bulk export

Export all the data tables the parser has a config file for to CSV files.

```python
from isp_assumptions_parser import Parser

workbook = Parser("<path/to/workbook>/2023 IASR Assumptions Workbook.xlsx")

workbook.save_tables('<path/to/output directory>')
```

## List tables with config

List all the tables the parser has a config file for (for the given workbook version).

```python
from isp_assumptions_parser import Parser

workbook = Parser("<path/to/workbook>/2023 IASR Assumptions Workbook.xlsx")

workbook.get_table_names()
```

## Get table as DataFrame

Get a single table as a pandas `DataFrame`.

```python
from isp_assumptions_parser import Parser

workbook = Parser("<path/to/workbook>/2023 IASR Assumptions Workbook.xlsx")

table = workbook.get_table("existing_generators_summary")
```

## Get table with custom config

Get a table by directly providing the config.

```python
# Will add once we settle on config format.
```
