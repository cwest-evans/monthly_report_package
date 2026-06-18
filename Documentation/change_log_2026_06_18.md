# Monthly Pulse Runner Change Log

Date: 2026-06-18
Project: Monthly Pulse Package Runner
Repo: `monthly-pulse-runner`

## Summary

Today’s work focused on stabilizing the Monthly Pulse package project, getting it into Git/GitHub, and reducing the amount of manual SQL required to run monthly packages.

The project moved from a mostly manual SQL-driven process toward a Python-controlled command-line workflow. SQL Server stored procedures still handle the heavy data staging and package data retrieval, but Python is now becoming the main operator interface for monthly runs.

## Major Improvements Completed

### 1. Git repository initialized and pushed to GitHub

A local Git repository was initialized for the project and connected to the GitHub repository:

```text
https://github.com/cwest-evans/monthly-pulse-runner.git
```

The local branch was renamed to `main` and pushed to GitHub.

### 2. Git hygiene added

A `.gitignore` file was created to prevent local/runtime files from being committed.

Ignored items include:

```text
.env
.venv/
.idea/
output_data/
output_test/
test_inputs/
*.parquet
*.xlsx.zip
__pycache__/
*.log
tree.txt
```

This protects secrets, virtual environment files, generated reports, temporary parquet snapshots, and local IDE settings.

### 3. Environment example added

An `.env.example` file was generated from the local `.env` structure with values removed.

The `.gitignore` originally ignored `.env.example` because of broad `.env.*` rules. This was corrected by explicitly allowing `.env.example`.

Purpose:

* Keep `.env` private
* Document required environment variables
* Make the project easier to set up on another machine

Key environment variable groups:

```text
SQL_SERVER
SQL_DATABASE
SQL_USERNAME
SQL_PASSWORD
SQL_DRIVER

TENANT_ID
CLIENT_ID
CLIENT_SECRET
GRAPH_SENDER_UPN

ENV
DEV_OVERRIDE_TO
```

### 4. README added

A README was created to document the current project purpose, operating model, environment variables, current workflow, output tabs, known cleanup items, and desired future workflow.

The README establishes the project as an operational Monthly Pulse package runner that is actively being improved.

### 5. `pulse_runner.py` converted into a CLI-controlled runner

The largest functional improvement was replacing the old hardcoded/manual operation pattern with command-line arguments.

New CLI options added:

```text
--snapshot-month
--as-of-month
--fiscal-year
--only-recipients
--validate-only
--skip-stage
--write-test-snapshot
--max-attempts
```

This allows controlled monthly runs like:

```powershell
python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --only-recipients jmillard@evans-gc.com,erech@evans-gc.com
```

### 6. Recipient limiting moved into Python

Before today, limiting a run to a few recipients required manually updating `EGC_Pulse_table_RunRecipient` through SQL.

That logic now exists in Python through the `--only-recipients` argument.

Current behavior:

* Listed recipients are reset to `PENDING`
* All other recipients for the run are marked `SKIPPED`
* Queue counts are printed after limiting

This significantly reduces the risk of accidentally sending a controlled run to the wrong audience.

### 7. Month selection moved into Python

Before today, specific month runs required manually calling:

```sql
EXEC dbo.EGC_sp_Pulse_Run_Init
  @SnapshotMonth = 'YYYY-MM-01',
  @AsOfMonth     = 'YYYY-MM-01',
  @FiscalYear    = YYYY;
```

The runner can now call `EGC_sp_Pulse_Run_Init` with those values based on CLI arguments.

This reduces confusion around which snapshot/as-of month is being processed.

### 8. Validation mode added

A `--validate-only` mode was added.

This mode:

* Initializes or retrieves the run
* Optionally limits recipients
* Stages data unless `--skip-stage` is used
* Prints validation output
* Exits before claiming recipients or sending emails

This gives a safe pre-flight workflow before real sends.

Example:

```powershell
python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --only-recipients jmillard@evans-gc.com,erech@evans-gc.com --validate-only
```

### 9. Stage and queue validation added to runner output

The runner now prints useful validation details during execution, including:

* Run metadata
* Recipient status counts
* Actionable queue rows
* Stage row counts
* Latest WIP snapshot
* WIP stage row count
* WIP MoM row count
* Awards Month row count
* Awards YTD row count

This gives the operator much more visibility directly from the CLI.

### 10. Temporary parquet snapshot behavior made optional

The previous runner wrote parquet snapshots into `test_inputs/sample1` as a temporary report-testing step.

That behavior is now controlled by:

```text
--write-test-snapshot
```

This keeps production runs cleaner while preserving the option for report testing.

### 11. Real controlled send executed for May 2026

A real CLI run was executed for:

```text
SnapshotMonth: 2026-05-01
AsOfMonth: 2026-05-01
FiscalYear: 2026
Recipients:
  - jmillard@evans-gc.com
  - erech@evans-gc.com
```

The run was configured to send to their real email addresses by setting:

```powershell
$env:ENV="PROD"
Remove-Item Env:\DEV_OVERRIDE_TO -ErrorAction SilentlyContinue
```

Then running:

```powershell
python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --only-recipients jmillard@evans-gc.com,erech@evans-gc.com
```

### 12. Type-checking concern resolved

A type-checking concern was identified around `token` being typed as `str | None` while `send_mail()` expects `str`.

A guard was added before calling `send_mail()` so the runner fails clearly if a token is unexpectedly missing.

Purpose:

* Satisfy type checkers
* Protect future refactors
* Prevent accidental `None` values from being passed into the Graph mailer

## Important Design Decision

The project should continue moving toward this separation of responsibilities:

```text
Python = operator interface and orchestration
SQL Server = source-of-truth staging and heavy data processing
Excel formatter = workbook presentation layer
Graph mailer = delivery layer
```

The goal is not necessarily to remove SQL Server stored procedures. The goal is to remove the need for a person to manually execute scattered SQL commands every month.

## Current Improved Workflow

### Validate only

```powershell
$env:ENV="PROD"
Remove-Item Env:\DEV_OVERRIDE_TO -ErrorAction SilentlyContinue

python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --only-recipients jmillard@evans-gc.com,erech@evans-gc.com --validate-only
```

### Real send

```powershell
$env:ENV="PROD"
Remove-Item Env:\DEV_OVERRIDE_TO -ErrorAction SilentlyContinue

python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --only-recipients jmillard@evans-gc.com,erech@evans-gc.com
```

### DEV override send

```powershell
$env:ENV="DEV"
$env:DEV_OVERRIDE_TO="cwest@evans-gc.com"

python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --only-recipients jmillard@evans-gc.com,erech@evans-gc.com
```

## Commits Made

Known commits from today included:

```text
Initial monthly pulse runner baseline
Add monthly pulse runner README
Add environment variable example
Add CLI controls to monthly pulse runner
Resolve type-checking concern in pulse runner
```

## Remaining Cleanup Items

### High Priority

* Add a formal monthly runbook under `Documentation/`
* Add `requirements.txt`
* Add a `sql/` folder with validation/stored procedure reference scripts
* Add a file-only / no-send mode if not already fully retained in the local runner
* Add stronger validation for missing recipients in `--only-recipients`
* Add confirmation output before real PROD sends

### Medium Priority

* Move code into a package structure, likely `src/monthly_pulse/`
* Create a `scripts/` folder for common PowerShell commands
* Add structured logging to a `logs/` folder
* Add tests for Excel formatting behavior
* Decide whether `report_pdf.py`, `report_template.html`, `report_assets/`, and `Documentation/` should stay in this repo or be split later

### Future Ideal CLI

The long-term goal is for the monthly process to be as simple as:

```powershell
python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --validate-only
```

Then:

```powershell
python .\pulse_runner.py --snapshot-month 2026-05-01 --as-of-month 2026-05-01 --fiscal-year 2026 --send
```

Eventually, this should support a clean, documented monthly process that can be repeated without relying on memory or one-off SQL snippets.
