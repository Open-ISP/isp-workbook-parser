# Contributing

Contributions are welcome, and they are greatly appreciated! Every little bit
helps, and credit will always be given.

## Types of Contributions

## Adding table configuration files (YAMLs)

Table configuration files for each workbook version are bundled with this package. They are located within `src/isp_table_configs/<workbook version number>`. To add table configuration files:

1. Follow the instructions in "Get Started!" below to create a new branch
2. Commit the new table configurations
3. Open a pull request, and if it passes config check tests, a maintainer will merge them in

A YAML file can contain many tables, which are defined as objects. For example, we can define a table called "retirement costs" as follows:
```yaml
retirement_costs:
  sheet_name: "Retirement"
  header_rows: 9
  end_row: 16
  column_range: "H:I"
```

While we have no strict rules on how table configurations should be named and organised, we encourage the following:

- Group tables of a common theme in one `.yaml` file
  - For example, all tables that contain generation capacity factor data should ideally be defined within a single YAML file, e.g. `capacity_factors.yaml`
- Name individual tables (e.g. `retirement_cost` as above) with sufficient detail such that another user using `Parser.get_tables()` can infer what data the table contains

### Report Bugs

If you are reporting a bug, please include:

* Your operating system name and version.
* Any details about your local setup that might be helpful in troubleshooting.
* Detailed steps to reproduce the bug.

### Fix Bugs

Look through the GitHub issues for bugs. Anything tagged with "bug" and "help
wanted" is open to whoever wants to implement it.

### Implement Features

Look through the GitHub issues for features. Anything tagged with "enhancement"
and "help wanted" is open to whoever wants to implement it.

### Write Documentation

You can never have enough documentation! Please feel free to contribute to any
part of the documentation, such as the official docs, docstrings, or even
on the web in blog posts, articles, and such.

### Submit Feedback

If you are proposing a feature:

* Explain in detail how it would work.
* Keep the scope as narrow as possible, to make it easier to implement.
* Remember that this is a volunteer-driven project, and that contributions
  are welcome :)

## Get Started!

Ready to contribute? Here's how to set up `isp-workbook-parser` for local development.

1. Download a copy of `isp-workbook-parser` locally.
2. Install [`rye`](https://rye.astral.sh/).
3. Install `isp-workbook-parser` using `rye` by running `rye sync` in the project directory.
4. Install the `pre-commit` git hook scripts that `isp-workbook-parser` uses by running the following code using `rye`:

      ```console
      $ rye run pre-commit install
      ```

5. Use `git` (or similar) to create a branch for local development and make your changes:

    ```console
    $ git checkout -b name-of-your-bugfix-or-feature
    ```

6. When you're done making changes, check that your changes conform to any code formatting requirements and pass any tests. That is:
    - Use the the `ruff` formatter bundled with `rye` by running `rye fmt`
    - Run tests by running `rye test`

7. Commit your changes and open a pull request.

## Pull Request Guidelines

Before you submit a pull request, check that it meets these guidelines:

1. The pull request should include additional tests if appropriate.
2. If the pull request adds functionality, the docstrings/README/docs should be updated.
3. The pull request should work for all currently supported operating systems and versions of Python.

## Code of Conduct

Please note that the `isp-workbook-parser` project is released with a
[Code of Conduct](CONDUCT.md). By contributing to this project you agree to abide by its terms.
