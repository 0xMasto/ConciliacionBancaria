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
    df["Monto_Excel"] = (
        pd.to_numeric(cred, errors="coerce").fillna(0)
        - pd.to_numeric(deb, errors="coerce").fillna(0)
    )

    df["Fecha"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date

    # Normalización interna para comparar
    df["Fecha_norm"] = df["Fecha"]
    df["Monto_norm"] = df["Monto_Excel"].round(2)

    # Filtrar filas válidas
    df = df[df["Fecha_norm"].notna() & df["Monto_norm"].notna()]

    return df


def _normalizar_bd(df_bd: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara el DataFrame de la base de datos:
    - Usa fec_doc como fecha.
    - Usa imp_mov_mo como monto.
    - Mantiene nro_trans.
    - Crea columnas normalizadas para comparación:
        Fecha_norm, Monto_norm, Fecha_BD, Monto_BD.
    """
    required_cols = {"fec_doc", "imp_mov_mo", "nro_trans"}
    faltantes = required_cols - set(df_bd.columns)
    if faltantes:
        raise ValueError(f"En el DataFrame de BD faltan columnas requeridas: {faltantes}")

    df = df_bd.copy()

    df["Fecha_BD"] = pd.to_datetime(df["fec_doc"], errors="coerce").dt.date
    df["Monto_BD"] = pd.to_numeric(df["imp_mov_mo"], errors="coerce").round(2)

    df["Fecha_norm"] = df["Fecha_BD"]
    df["Monto_norm"] = df["Monto_BD"]

    df = df[df["Fecha_norm"].notna() & df["Monto_norm"].notna()]

    return df

def comparar(df_excel: pd.DataFrame, df_bd: pd.DataFrame) -> pd.DataFrame:
    """
    Compara movimientos del Excel contra la BD por (Fecha, Monto),
    y devuelve TODAS las filas del Excel con columnas adicionales:

    - Descripcion (unificada)
    - Fecha_Excel, Monto_Excel
    - Fecha_norm, Monto_norm
    - Fecha_BD, Monto_BD, nro_trans
    - Encontrado (True/False)
    """

    df_excel_norm = _normalizar_excel(df_excel)
    df_bd_norm = _normalizar_bd(df_bd)

    # Left join → TODO el Excel, y trae datos BD si hay coincidencia
    merged = df_excel_norm.merge(
        df_bd_norm[["Fecha_norm", "Monto_norm", "Fecha_BD", "Monto_BD", "nro_trans"]],
        how="left",
        on=["Fecha_norm", "Monto_norm"],
        indicator=True
    )

    # Partimos del Excel normalizado
    resultado = df_excel_norm.copy()

    # Agregamos columnas BD
    resultado["Fecha_BD"] = merged["Fecha_BD"]
    resultado["Monto_BD"] = merged["Monto_BD"]
    resultado["nro_trans"] = merged["nro_trans"]
    resultado["Encontrado"] = merged["_merge"].eq("both")

    # -------------------------------------
    # 🔥 ELIMINAR COLUMNAS NO NECESARIAS
    # -------------------------------------
    # Trabajamos en lower para no depender de mayúsculas/acentos
    drop_targets_lower = {
        "número de documento",
        "numero de documento",
        "asunto",
        "dependencia",
        "débito",
        "debito",
        "crédito",
        "credito",
        "saldo",
        "referencia",
        "destino",
    }

    cols_a_eliminar = [
        c for c in resultado.columns
        if c.lower() in drop_targets_lower
    ]

    resultado = resultado.drop(columns=cols_a_eliminar, errors="ignore")

    # -------------------------------------
    # 🏷 RENOMBRAR DESCRIPCIÓN / CONCEPTO
    # -------------------------------------
    # Queremos una sola columna "Descripcion"
    candidatos_desc = [
        "Descripcion",
        "Descripción",
        "descripción",
        "descripcion",
        "Concepto",
        "concepto",
    ]

    desc_encontrada = None
    for cand in candidatos_desc:
        if cand in resultado.columns:
            desc_encontrada = cand
            break

    if desc_encontrada and desc_encontrada != "Descripcion":
        resultado = resultado.rename(columns={desc_encontrada: "Descripcion"})

    # -------------------------------------
    # 🗓 RENOMBRAR FECHA A Fecha_Excel
    # -------------------------------------
    if "Fecha" in resultado.columns:
        resultado = resultado.rename(columns={"Fecha": "Fecha_Excel"})

    # -------------------------------------
    # 📐 REORDENAR COLUMNAS
    # -------------------------------------
    orden_preferido = []
    # 1) Descripcion
    if "Descripcion" in resultado.columns:
        orden_preferido.append("Descripcion")
    # 2) Fecha_Excel y Monto_Excel
    if "Fecha_Excel" in resultado.columns:
        orden_preferido.append("Fecha_Excel")
    if "Monto_Excel" in resultado.columns:
        orden_preferido.append("Monto_Excel")
    # 3) columnas internas
    for c in ["Fecha_norm", "Monto_norm"]:
        if c in resultado.columns:
            orden_preferido.append(c)
    # 4) columnas de BD
    for c in ["Fecha_BD", "Monto_BD"]:
        if c in resultado.columns:
            orden_preferido.append(c)
    # 5) identificador + flag
    for c in ["nro_trans", "Encontrado"]:
        if c in resultado.columns:
            orden_preferido.append(c)

    # Agregar el resto de columnas que queden, respetando el orden original
    restantes = [c for c in resultado.columns if c not in orden_preferido]
    nuevo_orden = orden_preferido + restantes

    resultado = resultado[nuevo_orden]

    return resultado.reset_index(drop=True)


def comparar_y_exportar(
    df_excel: pd.DataFrame,
    df_bd: pd.DataFrame,
    ruta_salida: str
) -> Tuple[pd.DataFrame, str]:
    """
    Compara Excel vs BD y exporta TODOS los movimientos del Excel,
    indicando si se encontró coincidencia en la BD.

    - df_excel: DataFrame con movimientos del estado de cuenta.
    - df_bd:    DataFrame con la tabla m_cpf_contaux (incluyendo fec_doc, imp_mov_mo, nro_trans).
    - ruta_salida: ruta del archivo .xlsx a crear.

    Devuelve:
        (df_resultado, ruta_salida)
    """
    resultado = comparar(df_excel, df_bd)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        resultado.to_excel(writer, sheet_name="Comparacion", index=False)

    return resultado, ruta_salida

