import os
import time
from pathlib import Path

import pyodbc
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

BASE_DIR = Path(__file__).resolve().parents[1]
load_dotenv(BASE_DIR / ".env")


def _strip_outer_braces(value: str) -> str:
    value = value.strip()
    if value.startswith("{") and value.endswith("}"):
        return value[1:-1]
    return value


def _format_odbc_value(value: str, force_braces: bool = False) -> str:
    value = str(value)
    if value.startswith("{") and value.endswith("}"):
        return value
    if force_braces or any(char in value for char in ";{}") or value != value.strip():
        return "{" + value.replace("}", "}}") + "}"
    return value


db_driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
db_server = os.getenv("DB_SERVER", "localhost")
db_port = os.getenv("DB_PORT", "1433")
db_database = os.getenv("DB_DATABASE", "master")
db_user = os.getenv("DB_USER", "")
db_password = os.getenv("DB_PASSWORD", "")
db_encrypt = os.getenv("DB_ENCRYPT", "")
db_trust_server_certificate = os.getenv("DB_TRUST_SERVER_CERTIFICATE", "")
db_connection_timeout = os.getenv("DB_CONNECTION_TIMEOUT", "15")
db_odbc_extra = os.getenv("DB_ODBC_EXTRA", "")
db_connect_retries = int(os.getenv("DB_CONNECT_RETRIES", "5"))
db_connect_retry_delay = float(os.getenv("DB_CONNECT_RETRY_DELAY", "2"))


def build_odbc_connection_string(mask_password: bool = False) -> str:
    parts = [
        f"DRIVER={_format_odbc_value(_strip_outer_braces(db_driver), force_braces=True)}",
        f"SERVER={_format_odbc_value(db_server)},{db_port}",
        f"DATABASE={_format_odbc_value(db_database)}",
    ]

    if db_user and db_password:
        password = "***" if mask_password else db_password
        parts.extend(
            [
                f"UID={_format_odbc_value(db_user)}",
                f"PWD={_format_odbc_value(password)}",
            ]
        )
    else:
        parts.append("Trusted_Connection=yes")

    if db_encrypt:
        parts.append(f"Encrypt={db_encrypt}")
    if db_trust_server_certificate:
        parts.append(f"TrustServerCertificate={db_trust_server_certificate}")
    if db_connection_timeout:
        parts.append(f"Connection Timeout={db_connection_timeout}")
    if db_odbc_extra:
        parts.extend(part for part in db_odbc_extra.strip().strip(";").split(";") if part)

    return ";".join(parts) + ";"


conn_str = build_odbc_connection_string()


def _is_retryable_connection_error(exc: Exception) -> bool:
    error_text = str(exc).lower()
    retryable_markers = (
        "08001",
        "10061",
        "tcp",
        "timeout",
        "tempo limite",
        "login timeout",
        "logon expirou",
        "server was not found",
        "servidor nao foi encontrado",
        "actively refused",
        "recusou ativamente",
    )
    return any(marker in error_text for marker in retryable_markers)


def connect_with_retry():
    attempts = max(1, db_connect_retries)
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            return pyodbc.connect(conn_str, timeout=int(db_connection_timeout))
        except pyodbc.Error as exc:
            last_error = exc
            if attempt >= attempts or not _is_retryable_connection_error(exc):
                raise
            time.sleep(db_connect_retry_delay)

    raise last_error


engine = create_engine(
    "mssql+pyodbc://",
    creator=connect_with_retry,
    fast_executemany=True,
    pool_pre_ping=True,
)

SessionLocal = sessionmaker(bind=engine)

def get_session():
    return SessionLocal()
