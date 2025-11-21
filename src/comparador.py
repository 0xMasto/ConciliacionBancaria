# comparador.py
import pandas as pd
from typing import Tuple

def _normalizar_excel(df_excel: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara el DataFrame del Excel:
    - Detecta columnas de Fecha, Débito y Crédito.
    - Calcula Monto_Excel = Crédito - Débito.
    - Crea columnas normalizadas para comparación: Fecha_norm, Monto_norm.
    """
    df = df_excel.copy()

    # Mapeo de nombres tolerante a acentos y mayúsculas
    cols_lower = {c.lower(): c for c in df.columns}

    # Fecha
    fecha_col = None
    for key in ["fecha"]:
        if key in cols_lower:
            fecha_col = cols_lower[key]
            break
    if not fecha_col:
        raise ValueError("No se encontró la columna 'Fecha' en el DataFrame del Excel.")

    # Débito / Crédito
    deb_col = None
    cred_col = None
    for key in ["débito", "debito"]:
        if key in cols_lower:
            deb_col = cols_lower[key]
            break

    for key in ["crédito", "credito"]:
        if key in cols_lower:
            cred_col = cols_lower[key]
            break

    if not deb_col and not cred_col:
        raise ValueError("No se encontraron columnas 'Débito' o 'Crédito' en el Excel.")

    # Cálculo de monto
    deb = df[deb_col].fillna(0) if deb_col else 0
    cred = df[cred_col].fillna(0) if cred_col else 0
    df["Monto_Excel"] = pd.to_numeric(cred, errors="coerce").fillna(0) - pd.to_numeric(deb, errors="coerce").fillna(0)

    # Normalizar fecha y monto para comparación
    df["Fecha_norm"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date
    df["Monto_norm"] = df["Monto_Excel"].round(2)  # redondeo a 2 decimales por seguridad

    # Filtrar filas válidas
    df = df[df["Fecha_norm"].notna() & df["Monto_norm"].notna()]

    return df


def _normalizar_bd(df_bd: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara el DataFrame de la base de datos:
    - Usa fec_doc como fecha.
    - Usa imp_mov_mn como monto.
    - Mantiene nro_trans.
    - Crea columnas normalizadas para comparación: Fecha_norm, Monto_norm.
    """
    required_cols = {"fec_doc", "imp_mov_mn", "nro_trans"}
    faltantes = required_cols - set(df_bd.columns)
    if faltantes:
        raise ValueError(f"En el DataFrame de BD faltan columnas requeridas: {faltantes}")

    df = df_bd.copy()

    df["Fecha_norm"] = pd.to_datetime(df["fec_doc"], errors="coerce").dt.date
    df["Monto_norm"] = pd.to_numeric(df["imp_mov_mn"], errors="coerce").round(2)

    df = df[df["Fecha_norm"].notna() & df["Monto_norm"].notna()]

    return df


def comparar(df_excel: pd.DataFrame, df_bd: pd.DataFrame) -> pd.DataFrame:
    """
    Compara movimientos del Excel contra la BD por (Fecha, Monto).

    Coincidencia:
        Excel: Fecha_norm, Monto_norm  (derivado de Fecha + Crédito/Débito)
        BD:    Fecha_norm, Monto_norm  (fec_doc + imp_mov_mn)

    Devuelve un DataFrame con:
        - Fecha_Excel
        - Monto_Excel
        - Fecha_BD
        - Monto_BD
        - nro_trans
    """
    df_excel_norm = _normalizar_excel(df_excel)
    df_bd_norm = _normalizar_bd(df_bd)

    # Merge inner por llaves normalizadas (many-to-many si se repiten)
    merged = df_excel_norm.merge(
        df_bd_norm,
        how="inner",
        on=["Fecha_norm", "Monto_norm"],
        suffixes=("_excel", "_bd")
    )

    # Armar resultado final con las columnas pedidas
    resultado = pd.DataFrame({
        "Fecha_Excel": merged["Fecha_norm"],
        "Monto_Excel": merged["Monto_Excel"],
        "Fecha_BD": merged["Fecha_norm"],
        "Monto_BD": merged["imp_mov_mn"],
        "nro_trans": merged["nro_trans"],
    })

    return resultado.reset_index(drop=True)


def comparar_y_exportar(
    df_excel: pd.DataFrame,
    df_bd: pd.DataFrame,
    ruta_salida: str
) -> Tuple[pd.DataFrame, str]:
    """
    Compara Excel vs BD y exporta las coincidencias a un archivo Excel.

    - df_excel: DataFrame con movimientos del estado de cuenta.
    - df_bd:    DataFrame con la tabla m_cpf_contaux (incluyendo fec_doc, imp_mov_mn, nro_trans).
    - ruta_salida: ruta del archivo .xlsx a crear.

    Devuelve:
        (df_resultado, ruta_salida)
    """
    resultado = comparar(df_excel, df_bd)

    if resultado.empty:
        # Exportamos igual, pero vacío, para tener un rastro
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resultado.to_excel(writer, sheet_name="Coincidencias", index=False)
    else:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resultado.to_excel(writer, sheet_name="Coincidencias", index=False)

    return resultado, ruta_salida


# Prueba rápida manual (ejemplo, podés borrar esto si querés)
if __name__ == "__main__":
    # Ejemplo mínimo de estructura (solo para test manual)
    df_excel_demo = pd.DataFrame({
        "Fecha": ["2025-01-01", "2025-01-02"],
        "Crédito": [1000, None],
        "Débito": [None, 500],
    })
    df_bd_demo = pd.DataFrame({
        "fec_doc": ["2025-01-01", "2025-01-02"],
        "imp_mov_mn": [1000, -500],
        "nro_trans": [1, 2],
    })

    res, path = comparar_y_exportar(df_excel_demo, df_bd_demo, "comparacion_demo.xlsx")
    print(res)
    print("Exportado a:", path)
