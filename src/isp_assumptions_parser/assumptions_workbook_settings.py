column_name_corrections = {
    "5.2": {
        "generator_summary_data": {
            "Existing generator": "Generator",
            "Committed generator": "Generator",
            "Anticipated projects": "Generator",
            "Batteries": "Generator",
            "Additional projects": "Generator",
            "New entrants": "Generator",
            "Forced outage rate": "Forced outage rate - Full outage (% of time) - Until 2022",
            "Unnamed: 13": "Forced outage rate - Full outage (% of time) - Post 2022",
            "Unnamed: 14": "Forced outage rate - Partial outage (% of time) - Until 2022",
            "Unnamed: 15": "Forced outage rate - Partial outage (% of time) - Post 2022",
            "Mean time to repair": "Mean time to repair - Full outage (% of time) - Until 2022",
            "Unnamed: 17": "Mean time to repair - Full outage (% of time) - Post 2022",
            "Unnamed: 18": "Mean time to repair - Partial outage (% of time) - Until 2022",
            "Unnamed: 19": "Mean time to repair - Partial outage (% of time) - Post 2022",
        }
    }
}

table_configs = {
    "5.2": {
        # Summary mapping
        "existing_generators_summary": {
            "tab": "Summary Mapping",
            "junk_rows_at_start": 2,
            "column_name_corrections": column_name_corrections["5.2"][
                "generator_summary_data"
            ],
            "start_row": 4,
            "end_row": 256,
            "range": "B:AB",
        },
        "committed_generators_summary": {
            "tab": "Summary Mapping",
            "junk_rows_at_start": 2,
            "column_name_corrections": column_name_corrections["5.2"][
                "generator_summary_data"
            ],
            "start_row": 259,
            "end_row": 286,
            "range": "B:AB",
        },
        "anticipated_generator_location": {
            "tab": "Summary Mapping",
            "junk_rows_at_start": 2,
            "column_name_corrections": column_name_corrections["5.2"][
                "generator_summary_data"
            ],
            "start_row": 288,
            "end_row": 305,
            "range": "B:AB",
        },
        "battery_generator_summary": {
            "tab": "Summary Mapping",
            "junk_rows_at_start": 2,
            "column_name_corrections": column_name_corrections["5.2"][
                "generator_summary_data"
            ],
            "start_row": 307,
            "end_row": 352,
            "range": "B:AB",
        },
        "additional_generator_summary": {
            "tab": "Summary Mapping",
            "junk_rows_at_start": 2,
            "column_name_corrections": column_name_corrections["5.2"][
                "generator_summary_data"
            ],
            "start_row": 355,
            "end_row": 369,
            "range": "B:AB",
        },
        "new_generator_summary": {
            "tab": "Summary Mapping",
            "junk_rows_at_start": 2,
            "column_name_corrections": column_name_corrections["5.2"][
                "generator_summary_data"
            ],
            "start_row": 372,
            "end_row": 662,
            "range": "B:AG",
        },
        # Build costs
        "build_costs_current_policies": {
            "tab": "Build costs",
            "junk_rows_at_start": 0,
            "column_name_corrections": {},
            "start_row": 15,
            "end_row": 30,
            "range": "B:AI",
        },
        "build_costs_global_nze_post_2050": {
            "tab": "Build costs",
            "junk_rows_at_start": 0,
            "column_name_corrections": {},
            "start_row": 34,
            "end_row": 49,
            "range": "B:AI",
        },
        "build_costs_global_nze_by_2050": {
            "tab": "Build costs",
            "junk_rows_at_start": 0,
            "column_name_corrections": {},
            "start_row": 53,
            "end_row": 68,
            "range": "B:AI",
        },
        "build_costs_pumped_hydro_energy_storage": {
            "tab": "Build costs",
            "junk_rows_at_start": 0,
            "column_name_corrections": {},
            "start_row": 53,
            "end_row": 68,
            "range": "B:AI",
        },
        # Locational Cost Factors
        "locational_cost_factors": {
            "tab": "Locational Cost Factors",
            "junk_rows_at_start": 0,
            "column_name_corrections": {},
            "start_row": 11,
            "end_row": 26,
            "range": "B:G",
        },
        "technology_cost_breakdown_ratios": {
            "tab": "Locational Cost Factors",
            "junk_rows_at_start": 0,
            "column_name_corrections": {},
            "start_row": 33,
            "end_row": 48,
            "range": "B:F",
        },
        # Lead time and project life
        # Network Capability
        "flow_path_transfer_capability": {
            "tab": "Network Capability",
            "junk_rows_at_start": 1,
            "column_name_corrections": {},
            "start_row": 6,
            "end_row": 21,
            "range": "B:E",
        },
    }
}
