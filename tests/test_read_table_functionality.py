from isp_workbook_parser.config_model import TableConfig


def test_skip_single_row_in_single_header_row_table(workbook_v6):
    table_config = TableConfig(
        name="build_cost_current_policies",
        sheet_name="Build costs",
        header_rows=15,
        end_row=30,
        column_range="B:AI",
        skip_rows=30,
    )
    df = workbook_v6.get_table_from_config(table_config)
    assert len(df) == (table_config.end_row - table_config.header_rows - 1)
    assert df[df.Technology.str.contains("Hydrogen")].empty


def test_skip_multiple_rows_in_single_header_row_table(workbook_v6):
    table_config = TableConfig(
        name="existing_generator_maintenance_rates",
        sheet_name="Maintenance",
        header_rows=7,
        end_row=19,
        column_range="B:D",
        skip_rows=[8, 9, 19],
    )
    df = workbook_v6.get_table_from_config(table_config)
    assert len(df) == (table_config.end_row - table_config.header_rows - 3)
    assert df[df["Generator type"].str.contains("Hydrogen")].empty
    assert df[df["Generator type"].str.contains("Coal")].empty


def test_skip_multiple_rows_in_multiple_header_row_table(workbook_v6):
    table_config = TableConfig(
        name="wind_high_capacity_factors",
        sheet_name="Capacity Factors ",
        header_rows=[7, 8, 9],
        end_row=48,
        column_range="B:R",
        # Victoria
        skip_rows=(list(range(29, 35)) + [48]),
    )
    df = workbook_v6.get_table_from_config(table_config)
    assert len(df) == (table_config.end_row - table_config.header_rows[-1] - 7)
    assert df[df["Wind High_REZ ID"].str.contains("V")].empty
