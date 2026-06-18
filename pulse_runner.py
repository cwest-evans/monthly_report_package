import argparse
import os
import traceback
import zipfile
from pathlib import Path

import pandas as pd
import pyodbc

from config import SQL, GRAPH, ENV, DEV_OVERRIDE_TO
from graph_mailer import get_access_token, send_mail
from excel_formatter import load_format_spec_from_xlsx, apply_formats


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output_data"

FORMAT_SPEC_PATHS = [
    BASE_DIR / "specs" / "excel_package_column_formats.xlsx",
    BASE_DIR / "excel_package_column_formats.xlsx",
]

SHEETS = [
    "Rolling 18 Backlog",
    "WIP",
    "Income Statement (Month)",
    "Income Statement (YTD)",
    "Job Revenue (WIP MoM Latest)",
    "Awards (Month)",
    "Awards (YTD)",
]

PARQUET_NAMES = [
    "backlog",
    "wip",
    "is_month",
    "is_ytd",
    "rev_wip_mom_latest",
    "awards_month",
    "awards_ytd",
]


def find_format_spec_path() -> Path:
    for path in FORMAT_SPEC_PATHS:
        if path.exists():
            return path

    searched = "\n".join(f"  - {p}" for p in FORMAT_SPEC_PATHS)
    raise FileNotFoundError(f"Could not find Excel format spec. Checked:\n{searched}")


FORMAT_SPEC = load_format_spec_from_xlsx(str(find_format_spec_path()))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the EGC Monthly Pulse package process.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument(
        "--snapshot-month",
        help="Snapshot month for the Pulse run, e.g. 2026-04-01. If omitted, SQL proc default/latest behavior is used.",
    )

    parser.add_argument(
        "--as-of-month",
        help="As-of month for the Pulse run, e.g. 2026-04-01. If omitted, SQL proc default/latest behavior is used.",
    )

    parser.add_argument(
        "--fiscal-year",
        type=int,
        help="Fiscal year for the Pulse run, e.g. 2026. If omitted, SQL proc default/latest behavior is used.",
    )

    parser.add_argument(
        "--only-recipients",
        help=(
            "Comma-separated recipient email list to process. "
            "All other RunRecipient rows for the run will be marked SKIPPED."
        ),
    )

    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Initialize, optionally limit recipients, stage, print validation counts, and exit without claiming/sending packages.",
    )

    parser.add_argument(
        "--skip-stage",
        action="store_true",
        help="Skip StageAll. Use only when the run has already been staged intentionally.",
    )

    parser.add_argument(
        "--write-test-snapshot",
        action="store_true",
        help="Write parquet snapshots to test_inputs/sample1 for report testing.",
    )

    parser.add_argument(
        "--max-attempts",
        type=int,
        default=3,
        help="Maximum attempts allowed by the recipient claim stored procedure.",
    )

    return parser.parse_args()


def parse_recipient_list(raw: str | None) -> list[str]:
    if not raw:
        return []

    return [
        email.strip().lower()
        for email in raw.split(",")
        if email.strip()
    ]


def get_conn() -> pyodbc.Connection:
    conn_str = (
        f"DRIVER={{{SQL['driver']}}};"
        f"SERVER={SQL['server']};"
        f"DATABASE={SQL['database']};"
        f"UID={SQL['username']};"
        f"PWD={SQL['password']};"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str, autocommit=True)


def read_all_resultsets(cursor) -> list[pd.DataFrame]:
    dfs: list[pd.DataFrame] = []

    while True:
        if cursor.description is not None:
            cols = [c[0] for c in cursor.description]
            rows = cursor.fetchall()
            dfs.append(pd.DataFrame.from_records(rows, columns=cols))

        if not cursor.nextset():
            break

    return dfs


def print_df(title: str, df: pd.DataFrame | None) -> None:
    print(f"\n{title}")
    print("-" * len(title))

    if df is None or df.empty:
        print("(no rows)")
    else:
        print(df.to_string(index=False))


def run_init(
    cursor,
    snapshot_month: str | None = None,
    as_of_month: str | None = None,
    fiscal_year: int | None = None,
) -> dict:
    """
    Initializes or retrieves a Pulse run.

    If month arguments are supplied, the run is anchored to that month.
    If omitted, the SQL stored procedure uses its default/latest-month behavior.
    """
    if snapshot_month or as_of_month or fiscal_year:
        cursor.execute(
            """
            EXEC dbo.EGC_sp_Pulse_Run_Init
                @SnapshotMonth = ?,
                @AsOfMonth     = ?,
                @FiscalYear    = ?;
            """,
            snapshot_month,
            as_of_month,
            fiscal_year,
        )
    else:
        cursor.execute("EXEC dbo.EGC_sp_Pulse_Run_Init;")

    dfs = read_all_resultsets(cursor)

    if not dfs or dfs[0].empty:
        raise RuntimeError("EGC_sp_Pulse_Run_Init did not return a Run row.")

    return dfs[0].iloc[0].to_dict()


def stage_all(cursor, run_id: int) -> None:
    cursor.execute("EXEC dbo.EGC_sp_Pulse_Run_StageAll ?;", run_id)
    _ = read_all_resultsets(cursor)


def claim_next(cursor, run_id: int, max_attempts: int = 3) -> dict | None:
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_Recipient_ClaimNext ?, ?;",
        run_id,
        max_attempts,
    )

    dfs = read_all_resultsets(cursor)

    if not dfs or dfs[0].empty:
        return None

    return dfs[0].iloc[0].to_dict()


def get_package(cursor, run_id: int, email: str, div_key: str, mkt_key: str) -> list[pd.DataFrame]:
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_GetRecipientPackage_Staged ?, ?, ?, ?;",
        run_id,
        email,
        div_key,
        mkt_key,
    )

    dfs = read_all_resultsets(cursor)

    # SQL currently returns 9 result sets.
    # The old Job Revenue Month/YTD result sets are still returned but intentionally skipped.
    while len(dfs) < 9:
        dfs.append(pd.DataFrame())

    keep_idxs = [0, 1, 4, 5, 6, 7, 8]
    return [dfs[i] for i in keep_idxs]


def mark_success(cursor, run_id: int, email: str, div_key: str, mkt_key: str, filepath: str) -> None:
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_Recipient_MarkSuccess ?, ?, ?, ?, ?;",
        run_id,
        email,
        div_key,
        mkt_key,
        filepath,
    )


def mark_failed(cursor, run_id: int, email: str, div_key: str, mkt_key: str, err: str) -> None:
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_Recipient_MarkFailed ?, ?, ?, ?, ?;",
        run_id,
        email,
        div_key,
        mkt_key,
        err[:4000],
    )


def limit_run_recipients(cursor, run_id: int, recipients: list[str]) -> None:
    """
    Limits the run queue to a specific recipient list.

    All non-listed RunRecipient rows are marked SKIPPED.
    Listed recipients are reset to PENDING.
    """
    if not recipients:
        return

    cursor.execute("IF OBJECT_ID('tempdb..#OnlyRecipients') IS NOT NULL DROP TABLE #OnlyRecipients;")
    cursor.execute("CREATE TABLE #OnlyRecipients (RecipientEmail nvarchar(255) NOT NULL PRIMARY KEY);")

    cursor.executemany(
        "INSERT INTO #OnlyRecipients (RecipientEmail) VALUES (?);",
        [(email,) for email in recipients],
    )

    cursor.execute(
        """
        UPDATE rr
        SET rr.Status = 'SKIPPED',
            rr.UpdatedAt = SYSDATETIME()
        FROM dbo.EGC_Pulse_table_RunRecipient rr
        WHERE rr.RunId = ?
          AND LOWER(rr.RecipientEmail) NOT IN
          (
              SELECT RecipientEmail
              FROM #OnlyRecipients
          );
        """,
        run_id,
    )

    cursor.execute(
        """
        UPDATE rr
        SET rr.AttemptCount = 0,
            rr.Status = 'PENDING',
            rr.ErrorMessage = NULL,
            rr.LastAttemptAt = NULL,
            rr.CompletedAt = NULL,
            rr.FilePathOrUrl = NULL,
            rr.UpdatedAt = SYSDATETIME()
        FROM dbo.EGC_Pulse_table_RunRecipient rr
        WHERE rr.RunId = ?
          AND LOWER(rr.RecipientEmail) IN
          (
              SELECT RecipientEmail
              FROM #OnlyRecipients
          );
        """,
        run_id,
    )

    cursor.execute(
        """
        SELECT
            Status,
            COUNT(*) AS Cnt
        FROM dbo.EGC_Pulse_table_RunRecipient
        WHERE RunId = ?
        GROUP BY Status
        ORDER BY Status;
        """,
        run_id,
    )

    dfs = read_all_resultsets(cursor)
    print_df("Recipient queue after limiting", dfs[0] if dfs else pd.DataFrame())


def validate_run(cursor, run_id: int) -> None:
    """
    Prints key validation checks for the run.
    """
    cursor.execute(
        """
        SELECT
            RunId,
            SnapshotMonth,
            AsOfMonth,
            FiscalYear,
            StageStatus,
            StagedAt,
            CreatedAt,
            UpdatedAt
        FROM dbo.EGC_Pulse_table_Run
        WHERE RunId = ?;
        """,
        run_id,
    )

    dfs = read_all_resultsets(cursor)
    print_df("Run", dfs[0] if dfs else pd.DataFrame())

    cursor.execute(
        """
        SELECT
            Status,
            COUNT(*) AS Cnt
        FROM dbo.EGC_Pulse_table_RunRecipient
        WHERE RunId = ?
        GROUP BY Status
        ORDER BY Status;
        """,
        run_id,
    )

    dfs = read_all_resultsets(cursor)
    print_df("Recipient status counts", dfs[0] if dfs else pd.DataFrame())

    cursor.execute(
        """
        SELECT
            RecipientEmail,
            DivisionKey,
            MarketSegmentKey,
            Division,
            MarketSegment,
            Status
        FROM dbo.EGC_Pulse_table_RunRecipient
        WHERE RunId = ?
          AND Status IN ('PENDING', 'IN_PROGRESS', 'FAILED')
        ORDER BY Status, RecipientEmail, DivisionKey, MarketSegmentKey;
        """,
        run_id,
    )

    dfs = read_all_resultsets(cursor)
    print_df("Actionable recipient queue", dfs[0] if dfs else pd.DataFrame())

    cursor.execute(
        """
        SELECT
            Latest_WIP_Snapshot =
                (SELECT MAX(CAST(snapshot_month AS date)) FROM dbo.EGC_table_dashboard_WIP),

            WIP_Rows =
                (SELECT COUNT(*) FROM dbo.EGC_Pulse_stage_WIP WHERE RunId = ?),

            WIP_MoM_Rows =
                (SELECT COUNT(*) FROM dbo.EGC_Pulse_stage_JobRevenue_WIPMoM_Latest WHERE RunId = ?),

            Awards_Month_Rows =
                (SELECT COUNT(*) FROM dbo.EGC_Pulse_stage_Awards_Month WHERE RunId = ?),

            Awards_YTD_Rows =
                (SELECT COUNT(*) FROM dbo.EGC_Pulse_stage_Awards_YTD WHERE RunId = ?);
        """,
        run_id,
        run_id,
        run_id,
        run_id,
    )

    dfs = read_all_resultsets(cursor)
    print_df("Stage row counts", dfs[0] if dfs else pd.DataFrame())


def zip_file(src_path: str | Path) -> str:
    src_path = str(src_path)
    zip_path = src_path + ".zip"

    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(src_path, arcname=os.path.basename(src_path))

    return zip_path


def safe_sheet_name(name: str) -> str:
    bad = ["\\", "/", "*", "[", "]", ":", "?"]

    for ch in bad:
        name = name.replace(ch, "-")

    return name[:31]


def write_excel(filepath: str | Path, dfs: list[pd.DataFrame]) -> None:
    filepath = Path(filepath)
    filepath.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        for i, df in enumerate(dfs):
            sheet = safe_sheet_name(SHEETS[i] if i < len(SHEETS) else f"Sheet{i + 1}")

            if df is None or df.empty:
                pd.DataFrame({"(no rows)": []}).to_excel(writer, sheet_name=sheet, index=False)
            else:
                df.to_excel(writer, sheet_name=sheet, index=False)

        apply_formats(writer.book, FORMAT_SPEC, autosize=True)


def resolve_to_address(actual_email: str) -> str:
    # DEV safety: never email real users when override is populated.
    if (ENV or "").upper() == "DEV" and DEV_OVERRIDE_TO:
        return DEV_OVERRIDE_TO

    return actual_email


def write_test_snapshot(dfs: list[pd.DataFrame]) -> None:
    parquet_dir = BASE_DIR / "test_inputs" / "sample1"
    parquet_dir.mkdir(parents=True, exist_ok=True)

    for name, df in zip(PARQUET_NAMES, dfs):
        if df is None:
            df = pd.DataFrame()

        df.to_parquet(parquet_dir / f"{name}.parquet", index=False)

    print(f"Wrote parquet snapshot to: {parquet_dir}")


def main() -> None:
    args = parse_args()
    only_recipients = parse_recipient_list(args.only_recipients)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    token = None

    if not args.validate_only:
        token = get_access_token(
            GRAPH["tenant_id"],
            GRAPH["client_id"],
            GRAPH["client_secret"],
        )

    with get_conn() as conn:
        cur = conn.cursor()

        run = run_init(
            cur,
            snapshot_month=args.snapshot_month,
            as_of_month=args.as_of_month,
            fiscal_year=args.fiscal_year,
        )

        run_id = int(run["RunId"])
        as_of_str = pd.to_datetime(run["AsOfMonth"]).strftime("%Y-%m-%d")

        print(f"RunId={run_id} AsOfMonth={as_of_str} ENV={ENV}")

        if only_recipients:
            print("Limiting run to recipients:")

            for recipient in only_recipients:
                print(f"  - {recipient}")

            limit_run_recipients(cur, run_id, only_recipients)

        if args.skip_stage:
            print("Skipping StageAll because --skip-stage was supplied.")
        else:
            print("Staging run data (one-time)...")
            stage_all(cur, run_id)
            print("Staging complete.")

        validate_run(cur, run_id)

        if args.validate_only:
            print("Validate-only mode complete. No recipients were claimed and no packages were sent.")
            return

        while True:
            job = claim_next(cur, run_id, max_attempts=args.max_attempts)

            if job is None:
                print("No more recipients to process.")
                break

            email = job["RecipientEmail"]
            div_key = job["DivisionKey"]
            mkt_key = job["MarketSegmentKey"]

            div_label = "ALL" if div_key == "*" else div_key
            mkt_label = "ALL" if mkt_key == "*" else mkt_key

            try:
                dfs = get_package(cur, run_id, email, div_key, mkt_key)

                if args.write_test_snapshot:
                    write_test_snapshot(dfs)

                safe_email = str(email).replace("@", "_at_").replace(".", "_")
                fname = f"Monthly Pulse - {safe_email} - {as_of_str} - {div_label} - {mkt_label}.xlsx"
                filepath = OUTPUT_DIR / fname

                write_excel(filepath, dfs)
                zip_path = zip_file(filepath)

                to_addr = resolve_to_address(email)
                subject = f"Monthly Pulse Check ({as_of_str}) - {div_label} | {mkt_label}"
                body = f"""
                <p>Attached is your Monthly Pulse Check package for <b>{as_of_str}</b>.</p>
                <p><b>Scope:</b> Division={div_label}, Market={mkt_label}</p>
                <p><i>Note:</i> The file is zipped to keep the attachment size small.</p>
                """

                size_mb = os.path.getsize(zip_path) / (1024 * 1024)

                if size_mb >= 3.0:
                    raise ValueError(
                        f"Zipped attachment is {size_mb:.2f} MB "
                        "(still too large for simple Graph send)."
                    )

                send_mail(
                    token=token,
                    sender_upn=GRAPH["sender_upn"],
                    to_addrs=[to_addr],
                    subject=subject,
                    html_body=body,
                    attachment_path=zip_path,
                )

                mark_success(cur, run_id, email, div_key, mkt_key, zip_path)
                print(f"SUCCESS: {email} [{div_key}|{mkt_key}] -> sent to {to_addr}")

            except Exception:
                err = traceback.format_exc()
                mark_failed(cur, run_id, email, div_key, mkt_key, err)
                print(f"FAILED: {email} [{div_key}|{mkt_key}]\n{err}")


if __name__ == "__main__":
    main()