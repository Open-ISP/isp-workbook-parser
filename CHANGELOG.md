# Changelog

Notable changes to `isp-workbook-parser`, for people using the package
(new workbook versions, renamed/removed/added tables, and new or changed public config/API)

## [2.9.0] - 2026-09-15

### Added

- Support for IASR workbook version 7.8 (final 2026 ISP workbook).
    - Includes new "Hybrid site limits" sheet and 7 new tables.
    - More details of workbook changes here:
  [workbook_content_changes_7.8.md](https://github.com/Open-ISP/project-planning/blob/main/ISP2026/workbook_content_changes_7.8.md).
- `skip_checks` table config option: Allows you to opt a specific table out from specific validation checks, (rather than disabling all checks for it).

### Changed / Fixed

- **Table renames** (impacts version 7.8 of the workbook, as well as 7.3/7.5 retrospectivelly).
  - `*_reduced_energy_efficiency` → `*_lower_energy_efficiency`
  - `flow_path_augmentation_cost_slower_growth_CNSW-NNSW` →
    `flow_path_augmentation_costs_slower_growth_CNSW-NNSW`
  - Two clashing table names resolved (one of tables was previously being shadowed/ignored):
    - `water_for_hydrogen` →
    `desalination_electricity_demand_for_h2`
    - `embedded_storage_consultant_scenario_mapping`
    → `aggregated_storage_consultant_scenario_mapping`
- Fixed a validation issue that resulted in some tables dropping rows when a table's second column happened to be blank
  - (affected `marginal_loss_factors_existing_generators` in 7.8).
- Fixed sanitiser handling of `[footnote]`-style markers and notes containing `$` amounts (which were previously partly sanitising the values).
- Cells holding two distinct values (e.g. `"250 (generation) 325 (pump)"`) are intentionally returned as text (instead of choosing, for example, the first value).
- Dropped Python 3.9 support; added Python 3.14.
