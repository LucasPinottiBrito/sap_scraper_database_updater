import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

load_dotenv()

db_driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
db_server = os.getenv("DB_SERVER", "localhost")
db_port = os.getenv("DB_PORT", "1433")
db_database = os.getenv("DB_DATABASE", "master")
db_user = os.getenv("DB_USER", "")
db_password = os.getenv("DB_PASSWORD", "")

conn_str = f"DRIVER={db_driver};SERVER={db_server},{db_port};DATABASE={db_database};"

if db_user and db_password:
    conn_str += f"UID={db_user};PWD={db_password};"
else:
    conn_str += "Trusted_Connection=yes;"

odbc_url = conn_str.replace(" ", "+")
engine = create_engine(
    f"mssql+pyodbc:///?odbc_connect={odbc_url}",
    fast_executemany=True,
)

SessionLocal = sessionmaker(bind=engine)

def get_session():
    return SessionLocal()
