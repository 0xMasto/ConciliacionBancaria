import pandas as pd
from sqlalchemy import create_engine
import socket

import psycopg2


# Configuración para PostgreSQL
POSTGRES_CONFIG = {
    "host": "10.10.1.162",
    "database": "postgres",
    "user": "postgres", 
    "password": "postgres",
    "port": "54322"
}

# Crear engine para PostgreSQL
engine = create_engine(
    f"postgresql+psycopg2://{POSTGRES_CONFIG['user']}:{POSTGRES_CONFIG['password']}@"
    f"{POSTGRES_CONFIG['host']}:{POSTGRES_CONFIG['port']}/{POSTGRES_CONFIG['database']}"
)

def obtener_df_bd():
    query = """
        SELECT *
        FROM conciliacion.m_cpf_contaux t
        WHERE t.conciliado = FALSE
        LIMIT 100
        
    """
    
    try:
        df = pd.read_sql(query, engine)
        print(f"✅ Leídos {len(df)} registros de la base de datos")
        return df
    except Exception as e:
        print(f"❌ Error en la consulta: {e}")
        return None



# Ejecutar la función
if __name__ == "__main__":
    df = obtener_df_bd()
    if df is not None:
        print(df)