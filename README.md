# ISP Assumptions Parser 

A Python program for reading data from the Assumptions, Inputs, and Scenarios Report (AISR) MS Excel Workbook
published by the Australian Energy Market Operator for use in the Integrated System Plan modelling.

# Examples 

## Bulk export

Export all the data tables the parser has config for to csv files.

```python
from isp_assumptions_parser import Parser

workbook = Parser("<path/to/workbook>/2023 IASR Assumptions Workbook.xlsx")

workbook.save_tables('<path/to/output directory>')
```

## List tables with config

List all the tables the parser has config for (for the given workbook version).

```python
from isp_assumptions_parser import Parser

workbook = Parser("<path/to/workbook>/2023 IASR Assumptions Workbook.xlsx")

workbook.get_table_names()
```

## Get table as DataFrame

Get a single table as a pd.DataFrame.

```python
from isp_assumptions_parser import Parser

workbook = Parser("<path/to/workbook>/2023 IASR Assumptions Workbook.xlsx")

table = workbook.get_table("existing_generators_summary")
```

## Get table with custom config

Get a table by directly providing config.

```python
# Will add once we settle on config format.
```
