# config.py
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

def env(key: str, default: str | None = None) -> str:
    v = os.getenv(key, default)
    if v is None:
        raise RuntimeError(f"Missing required env var: {key}")
    return v

SQL = {
    "server": env("SQL_SERVER"),
    "database": env("SQL_DATABASE", "Viewpoint"),
    "username": env("SQL_USERNAME"),
    "password": env("SQL_PASSWORD"),
    "driver": env("SQL_DRIVER", "ODBC Driver 17 for SQL Server"),
}

GRAPH = {
    "tenant_id": env("TENANT_ID"),
    "client_id": env("CLIENT_ID"),
    "client_secret": env("CLIENT_SECRET"),
    "sender_upn": env("GRAPH_SENDER_UPN"),
}

ENV = env("ENV", "PROD").upper()
DEV_OVERRIDE_TO = os.getenv("DEV_OVERRIDE_TO")

