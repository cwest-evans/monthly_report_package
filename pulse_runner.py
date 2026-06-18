import os
import traceback
import pyodbc
import pandas as pd
import zipfile
from pathlib import Path

from config import SQL, GRAPH, ENV, DEV_OVERRIDE_TO
from graph_mailer import get_access_token, send_mail
from excel_formatter import load_format_spec_from_xlsx, apply_formats

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = r"output_data"

FORMAT_SPEC_PATH = BASE_DIR / "specs" / "excel_package_column_formats.xlsx"
FORMAT_SPEC = load_format_spec_from_xlsx(str(FORMAT_SPEC_PATH))

SHEETS = [
    "Rolling 18 Backlog",
    "WIP",
    "Income Statement (Month)",
    "Income Statement (YTD)",
    "Job Revenue (WIP MoM Latest)",
    "Awards (Month)",
    "Awards (YTD)",
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


def run_init(cursor) -> dict:
    cursor.execute("EXEC dbo.EGC_sp_Pulse_Run_Init;")
    df = read_all_resultsets(cursor)[0]
    return df.iloc[0].to_dict()

def stage_all(cursor, run_id: int):
    cursor.execute("EXEC dbo.EGC_sp_Pulse_Run_StageAll ?;", run_id)
    # optional: consume any resultsets so the cursor is clean
    _ = read_all_resultsets(cursor)


def claim_next(cursor, run_id: int, max_attempts: int = 3) -> dict | None:
    cursor.execute("EXEC dbo.EGC_sp_Pulse_Recipient_ClaimNext ?, ?;", run_id, max_attempts)
    dfs = read_all_resultsets(cursor)
    if not dfs or dfs[0].empty:
        return None
    return dfs[0].iloc[0].to_dict()



def get_package(cursor, run_id: int, email: str, div_key: str, mkt_key: str) -> list[pd.DataFrame]:
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_GetRecipientPackage_Staged ?, ?, ?, ?;",
        run_id, email, div_key, mkt_key
    )
    dfs = read_all_resultsets(cursor)

    # SQL now returns 9 result sets; pad if needed
    while len(dfs) < 9:
        dfs.append(pd.DataFrame())

    # Disable Job Revenue Month/YTD (result sets 3 and 4 => zero-based idx 2 and 3)
    keep_idxs = [0, 1, 4, 5, 6, 7, 8]
    return [dfs[i] for i in keep_idxs]


def mark_success(cursor, run_id: int, email: str, div_key: str, mkt_key: str, filepath: str):
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_Recipient_MarkSuccess ?, ?, ?, ?, ?;",
        run_id, email, div_key, mkt_key, filepath
    )


def mark_failed(cursor, run_id: int, email: str, div_key: str, mkt_key: str, err: str):
    cursor.execute(
        "EXEC dbo.EGC_sp_Pulse_Recipient_MarkFailed ?, ?, ?, ?, ?;",
        run_id, email, div_key, mkt_key, err[:4000]
    )


def zip_file(src_path: str) -> str:
    """
    Zips src_path into src_path + '.zip' and returns the zip path.
    If zip already exists, it will be overwritten.
    """
    zip_path = src_path + ".zip"
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(src_path, arcname=os.path.basename(src_path))

    return zip_path

def safe_sheet_name(name: str) -> str:
    bad = ['\\', '/', '*', '[', ']', ':', '?']
    for ch in bad:
        name = name.replace(ch, "-")
    return name[:31]


def write_excel(filepath: str, dfs: list[pd.DataFrame]):
    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        for i, df in enumerate(dfs):
            sheet = safe_sheet_name(SHEETS[i] if i < len(SHEETS) else f"Sheet{i+1}")
            if df is None or df.empty:
                pd.DataFrame({"(no rows)": []}).to_excel(writer, sheet_name=sheet, index=False)
            else:
                df.to_excel(writer, sheet_name=sheet, index=False)

        # Apply formatting AFTER sheets are written
        apply_formats(writer.book, FORMAT_SPEC, autosize=True)


def resolve_to_address(actual_email: str) -> str:
    # DEV safety: never email real users
    if (ENV or "").upper() == "DEV" and DEV_OVERRIDE_TO:
        return DEV_OVERRIDE_TO
    return actual_email


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    token = get_access_token(GRAPH["tenant_id"], GRAPH["client_id"], GRAPH["client_secret"])

    with get_conn() as conn:
        cur = conn.cursor()

        run = run_init(cur)
        run_id = int(run["RunId"])
        as_of_str = pd.to_datetime(run["AsOfMonth"]).strftime("%Y-%m-%d")

        print(f"RunId={run_id} AsOfMonth={as_of_str} ENV={ENV}")

        # NEW: stage once per run (heavy work)
        print("Staging run data (one-time)...")
        stage_all(cur, run_id)
        print("Staging complete.")

        while True:
            job = claim_next(cur, run_id)
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

                # TEMP: dump parquet snapshot for report testing (one-time)
                from pathlib import Path

                PARQUET_DIR = Path("test_inputs") / "sample1"
                PARQUET_DIR.mkdir(parents=True, exist_ok=True)

                parquet_names = [
                    "backlog",
                    "wip",
                    "is_month",
                    "is_ytd",
                    "rev_wip_mom_latest",
                    "awards_month",
                    "awards_ytd",
                ]

                # Only dump once (so you don't overwrite every loop)
                if not (PARQUET_DIR / "backlog.parquet").exists():
                    for name, df in zip(parquet_names, dfs):
                        if df is None:
                            df = pd.DataFrame()
                        df.to_parquet(PARQUET_DIR / f"{name}.parquet", index=False)
                    print(f"Wrote parquet snapshot to: {PARQUET_DIR}")

                fname = f"Monthly Pulse - {email} - {as_of_str}.xlsx"
                filepath = os.path.join(OUTPUT_DIR, fname)

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
                    raise ValueError(f"Zipped attachment is {size_mb:.2f} MB (still too large for simple Graph send).")

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
