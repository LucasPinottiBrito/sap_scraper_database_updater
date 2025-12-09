from databricks import sql
import os
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

class DatabricksConnector:
    def __init__(self):
        self.server_hostname = os.getenv("DATABRICKS_URL").replace("https://", "")
        self.access_token = os.getenv("DATABRICKS_TOKEN")
        self.connection = None
        self.cursor = None

    def connect(self):
        self.connection = sql.connect(
            server_hostname=self.server_hostname,
            http_path="/sql/1.0/warehouses/6628c75182b117ed",
            access_token=self.access_token
        )
        self.cursor = self.connection.cursor()

    def execute_query(self, query):
        if not self.cursor:
            raise Exception("Connection not established. Call connect() first.")
        self.cursor.execute(query)
        result = self.cursor.fetchall()
        df_pd = pd.DataFrame(result, columns=[col[0] for col in self.cursor.description])
        return df_pd

    def get_data_from_anlage(self, anlage: str):
        query = f"""
        SELECT *
        FROM bronze.anlage_data
        WHERE anlage = '{anlage}'
        """
        return self.execute_query(query)
    
    
    def close(self):
        if self.cursor:
            self.cursor.close()
        if self.connection:
            self.connection.close()