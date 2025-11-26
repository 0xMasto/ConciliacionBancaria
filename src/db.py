import pandas as pd
from sqlalchemy import create_engine, text
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

def obtener_df_bd(cod_tit: str) -> pd.DataFrame | None:
    """
    Devuelve los registros NO conciliados de conciliacion.m_cpf_contaux
    filtrando por:
      - conciliado = false
      - trim(cod_aux) = 'bancos'
      - trim(cod_tit) = cod_tit (string)
    """

    sql = """
        SELECT *
        FROM conciliacion.m_cpf_contaux t
        WHERE t.conciliado = FALSE
          AND trim(t.cod_aux) = 'bancos'
          AND trim(t.cod_tit) = :cod_tit
    """

    try:
        df = pd.read_sql(text(sql), engine, params={"cod_tit": cod_tit})
        print(f"📥 Leídos {len(df)} registros de BD para cod_tit={cod_tit}")
        return df
    except Exception as e:
        print(f"❌ Error leyendo BD: {e}")
        return None
    
# Ejecutar la función
if __name__ == "__main__":
    df = obtener_df_bd()
    if df is not None:
        print(df)