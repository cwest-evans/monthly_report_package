# Monthly Pulse Package Runner

Python-based monthly reporting package runner for generating and emailing Monthly Pulse Excel packages from SQL Server staged reporting data.

The process currently uses SQL Server stored procedures to initialize a monthly run, stage reporting data, claim recipient/package work items, generate formatted Excel workbooks, zip the package, and send through Microsoft Graph.

## Current Status

This project is operational but still being cleaned up.

The current runner can:

* Initialize a monthly Pulse run
* Stage package data through SQL Server stored procedures
* Claim pending recipients from the run queue
* Generate Excel workbooks with multiple reporting tabs
* Apply formatting from the Excel format specification workbook
* Zip generated packages
* Send packages through Microsoft Graph
* Mark recipients as successful or failed in SQL Server

The current goal is to reduce manual SQL execution and move toward a repeatable command-based monthly workflow.

## Project Files

Current core files:

```text
config.py
excel_formatter.py
graph_mailer.py
pulse_runner.py
report_pdf.py
report_template.html
```

Important support files:

```text
specs/excel_package_column_formats.xlsx
.env
.gitignore
```

Generated/runtime files should not be committed:

```text
.venv/
.env
output_data/
output_test/
test_inputs/
*.parquet
*.xlsx.zip
__pycache__/
.idea/
```

## Environment Variables

The runner expects credentials and settings to be provided through environment variables.

Required SQL variables:

```env
SQL_SERVER=
SQL_DATABASE=Viewpoint
SQL_USERNAME=
SQL_PASSWORD=
SQL_DRIVER=ODBC Driver 17 for SQL Server
```

Required Microsoft Graph variables:

```env
TENANT_ID=
CLIENT_ID=
CLIENT_SECRET=
GRAPH_SENDER_UPN=
```

Runtime mode variables:

```env
ENV=DEV
DEV_OVERRIDE_TO=
```

When `ENV=DEV` and `DEV_OVERRIDE_TO` is populated, emails are redirected to the override address instead of the real recipient. This is used for dry runs.

## Current Monthly Workflow

The current process is still partly SQL-driven.

### 1. Confirm source data is refreshed

Before running the package, confirm the latest source data exists for the intended snapshot month.

Example:

```sql
SELECT MAX(CAST(snapshot_month AS date)) AS Latest_WIP_Snapshot
FROM dbo.EGC_table_dashboard_WIP;
```

### 2. Initialize the run

```sql
EXEC dbo.EGC_sp_Pulse_Run_Init
    @SnapshotMonth = 'YYYY-MM-01',
    @AsOfMonth     = 'YYYY-MM-01',
    @FiscalYear    = YYYY;
```

### 3. Confirm the run queue

```sql
DECLARE @RunId int =
(
    SELECT RunId
    FROM dbo.EGC_Pulse_table_Run
    WHERE SnapshotMonth = 'YYYY-MM-01'
);

SELECT Status, COUNT(*) AS Cnt
FROM dbo.EGC_Pulse_table_RunRecipient
WHERE RunId = @RunId
GROUP BY Status;
```

### 4. Limit recipients if needed

For a dry fire or controlled release, manually set everyone else to `SKIPPED` and the intended recipients to `PENDING`.

### 5. Stage the run

```sql
DECLARE @RunId int =
(
    SELECT RunId
    FROM dbo.EGC_Pulse_table_Run
    WHERE SnapshotMonth = 'YYYY-MM-01'
);

EXEC dbo.EGC_sp_Pulse_Run_StageAll @RunId = @RunId;
```

### 6. Run the Python runner

```powershell
python pulse_runner.py
```

In DEV mode:

```powershell
$env:ENV="DEV"
$env:DEV_OVERRIDE_TO="your.email@evans-gc.com"
python .\pulse_runner.py
```

In PROD mode:

```powershell
$env:ENV="PROD"
$env:DEV_OVERRIDE_TO=""
python .\pulse_runner.py
```

## Current Output Tabs

The generated workbook currently includes:

1. Rolling 18 Backlog
2. WIP
3. Income Statement (Month)
4. Income Statement (YTD)
5. Job Revenue (WIP MoM Latest)
6. Awards (Month)
7. Awards (YTD)

Older Job Revenue Month/YTD result sets are still returned by SQL but are intentionally skipped by the Python runner.

## Known Cleanup Items

### High Priority

* Move manual SQL queue setup into Python
* Add command-line arguments for snapshot month, as-of month, fiscal year, run mode, and recipient list
* Add a true dry-run mode that generates files without sending emails
* Add a validate-only mode that checks source data, run queue, and stage row counts
* Remove or gate temporary parquet snapshot generation
* Add structured logging
* Add a clear monthly runbook under `docs/`

### Medium Priority

* Move source files into a package structure under `src/monthly_pulse/`
* Split operational SQL into organized files under `sql/`
* Add a `requirements.txt`
* Add basic unit tests for Excel formatting
* Add a changelog

## Desired Future Workflow

The goal is to make the monthly process executable through one or two clear commands.

Example validation command:

```powershell
python pulse_runner.py --snapshot-month 2026-04-01 --as-of-month 2026-04-01 --validate-only
```

Example dry run:

```powershell
python pulse_runner.py --snapshot-month 2026-04-01 --as-of-month 2026-04-01 --env DEV --override-to your.email@evans-gc.com --recipients jmillard@evans-gc.com,erech@evans-gc.com
```

Example production send:

```powershell
python pulse_runner.py --snapshot-month 2026-04-01 --as-of-month 2026-04-01 --env PROD --recipients jmillard@evans-gc.com,erech@evans-gc.com --send
```

## Git Notes

Do not commit secrets or runtime artifacts.

Safe to commit:

```text
*.py
README.md
.gitignore
.env.example
requirements.txt
specs/
docs/
sql/
```

Do not commit:

```text
.env
.venv/
output_data/
test_inputs/
*.parquet
*.xlsx.zip
.idea/
__pycache__/
```
