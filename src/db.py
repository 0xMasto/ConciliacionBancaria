import pandas as pd
from sqlalchemy import create_engine
import urllib.parse

# Parámetros de conexión CORRECTOS para SQL Server
server = "10.10.0.23"
database = "Nodum_Prod"
username = "userpbi"
password = "Zo2_Zud9_K1"
driver = "ODBC Driver 17 for SQL Server"

# Crear string de conexión
params = urllib.parse.quote_plus(
    f"DRIVER={{{driver}}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password}"
)

# Engine CORRECTO para SQL Server
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

def obtener_df_sql_por_orden():
    query = """
        SELECT TOP 10 *
        FROM Nodum_Prod.dbo.cpf_contaux
        WHERE fec_doc = CONVERT(VARCHAR, '22-10-2025')
    """
    
    try:
        df = pd.read_sql(query, engine)
        print(f"✅ Leídos {len(df)} registros")
        print(df.head())
        return df
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

# Ejecutar la función
obtener_df_sql_por_orden()