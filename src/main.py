# main.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import pandas as pd

# Módulos propios
import lectorItau
import lectorBrou
import db


class ComparadorApp:
    def __init__(self, root):
        self.root = root
        self.df_excel = None
        self.df_bd = None
        self.df_comparacion = None
        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        self.root.title("Comparador de Estados de Cuenta vs Base de Datos")
        self.root.geometry("800x550")
        self.root.resizable(True, True)

        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        # --- Selección de archivo ---
        file_box = ttk.LabelFrame(main, text="1) Selección de archivo", padding=12)
        file_box.pack(fill="x", pady=(0, 10))

        ttk.Label(file_box, text="Archivo:").grid(row=0, column=0, sticky="w")
        self.entrada_archivo = ttk.Entry(file_box, width=70)
        self.entrada_archivo.grid(row=0, column=1, padx=6, pady=4, sticky="we")
        ttk.Button(file_box, text="Examinar...", command=self.seleccionar_archivo).grid(row=0, column=2, padx=4)

        ttk.Label(file_box, text="Tipo de archivo:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.combo_tipo = ttk.Combobox(file_box, values=["Itaú", "BROU"], state="readonly", width=18)
        self.combo_tipo.grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))
        self.combo_tipo.set("Itaú")

        # --- Acciones ---
        actions = ttk.LabelFrame(main, text="2) Procesamiento y comparación", padding=12)
        actions.pack(fill="x", pady=(0, 10))

        ttk.Button(actions, text="Procesar y Comparar", command=self.procesar_y_comparar).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(actions, text="Exportar resultados", command=self.exportar_resultados).grid(row=0, column=1, padx=4, pady=4)

        # --- Resultados / Log ---
        results = ttk.LabelFrame(main, text="Resultados", padding=12)
        results.pack(fill="both", expand=True)

        self.text_resultados = tk.Text(results, height=18, wrap="word")
        self.text_resultados.pack(fill="both", expand=True)
        self.text_resultados.configure(state="disabled")

    # -------------- Utilidades --------------
    def log(self, mensaje: str):
        self.text_resultados.configure(state="normal")
        self.text_resultados.insert(tk.END, f"{mensaje}\n")
        self.text_resultados.see(tk.END)
        self.text_resultados.configure(state="disabled")
        self.root.update_idletasks()

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de estado de cuenta",
            filetypes=[("Archivos Excel", "*.xls *.xlsx"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            self.entrada_archivo.delete(0, tk.END)
            self.entrada_archivo.insert(0, ruta)

    # -------------- Flujo principal --------------
    def procesar_y_comparar(self):
        """Flujo completo: leer Excel -> preparar -> leer BD -> comparar."""
        if not self.procesar_archivo():
            return
        if not self.consultar_bd():
            return
        self.comparar_datos()

    def procesar_archivo(self) -> bool:
        ruta = self.entrada_archivo.get().strip()
        tipo = (self.combo_tipo.get() or "").strip()

        if not ruta or not os.path.exists(ruta):
            messagebox.showerror("Error", "Seleccioná un archivo válido.")
            return False

        self.log(f"📁 Procesando archivo: {ruta}")
        try:
            if tipo == "Itaú":
                self.df_excel = lectorItau.procesar_itau(ruta)
            elif tipo == "BROU":
                self.df_excel = lectorBrou.procesar_brou(ruta)
            else:
                messagebox.showwarning("Atención", "Seleccioná el tipo de archivo (Itaú o BROU).")
                return False

            if self.df_excel is None or self.df_excel.empty:
                self.log("⚠️ El lector devolvió un DataFrame vacío.")
                return False

            self._preparar_datos_excel()
            self.log(f"✅ Archivo {tipo} procesado ({len(self.df_excel)} filas).")
            return True

        except Exception as e:
            self.log(f"❌ Error procesando archivo: {e}")
            messagebox.showerror("Error", f"Ocurrió un error procesando el archivo:\n{e}")
            return False

    def _preparar_datos_excel(self):
        """Normaliza las columnas del Excel para comparación: Monto, Monto_Abs y Fecha (solo fecha)."""
        df = self.df_excel.copy()

        # Asegurar existencia de columnas (en Itaú/BROU ya vienen normalizadas)
        cols = {c.lower(): c for c in df.columns}
        deb_col = cols.get("débito") or cols.get("debito")
        cred_col = cols.get("crédito") or cols.get("credito")
        fecha_col = cols.get("fecha")

        # Calcula Monto = Crédito - Débito
        if deb_col and cred_col:
            df["Monto"] = (df[cred_col].fillna(0) - df[deb_col].fillna(0)).astype(float)
        elif "Monto" in df.columns:
            df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce")
        else:
            raise ValueError("No se pudo determinar la columna Monto a partir de Débito/Crédito.")

        df = df[df["Monto"].notna() & (df["Monto"] != 0)].copy()
        df["Monto_Abs"] = df["Monto"].abs()

        # Normaliza fecha a date
        if not fecha_col:
            raise ValueError("No se encontró la columna 'Fecha' en el Excel procesado.")
        df["Fecha"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date

        self.df_excel = df

    def consultar_bd(self) -> bool:
        """Lee la tabla de BD y estandariza las columnas clave para la comparación."""
        self.log("🔌 Conectando a la base de datos...")
        try:
            # opcional: probar_conexion para log amigable
            if hasattr(db, "probar_conexion"):
                if not db.probar_conexion():
                    self.log("❌ No hay conexión con la base.")
                    return False

            self.log("🔍 Consultando tabla conciliacion.m_cpf_contaux...")
            self.df_bd = db.obtener_df_bd()  # <- usa tu db.py

            if self.df_bd is None or self.df_bd.empty:
                self.log("⚠️ La tabla de BD no devolvió registros.")
                return False

            # Estandarizar columnas: monto y fecha (tolerante a nombres)
            self._normalizar_columnas_bd()
            self.log(f"✅ BD consultada ({len(self.df_bd)} filas).")
            return True

        except Exception as e:
            self.log(f"❌ Error consultando BD: {e}")
            messagebox.showerror("Error BD", f"Error consultando la base de datos:\n{e}")
            return False

    def _normalizar_columnas_bd(self):
        """Intenta identificar columnas de importe y fecha en la BD."""
        df = self.df_bd.copy()

        # Posibles nombres de monto
        candidatos_monto = [
            "imp_neto", "importe_neto", "importe", "monto", "amount", "importe_total", "val_neto", "imp_total"
        ]
        # Posibles nombres de fecha
        candidatos_fecha = [
            "fec_doc", "fecha", "fecha_doc", "fch_doc", "fecha_mov", "fch_mov"
        ]

        df_cols_lower = {c.lower(): c for c in df.columns}

        monto_col = next((df_cols_lower[c] for c in df_cols_lower if c in candidatos_monto), None)
        fecha_col = next((df_cols_lower[c] for c in df_cols_lower if c in candidatos_fecha), None)

        if not monto_col:
            raise ValueError(
                "No pude identificar la columna de monto en la BD. "
                "Ajustá la lista de candidatos en _normalizar_columnas_bd()."
            )
        if not fecha_col:
            raise ValueError(
                "No pude identificar la columna de fecha en la BD. "
                "Ajustá la lista de candidatos en _normalizar_columnas_bd()."
            )

        # Convierte tipos
        df["_Monto_BD"] = pd.to_numeric(df[monto_col], errors="coerce").abs()
        df["_Fecha_BD"] = pd.to_datetime(df[fecha_col], errors="coerce").dt.date

        # Conserva columnas útiles + las normalizadas
        self.df_bd = df

        self.log(f"🧭 Columnas BD detectadas -> monto: '{monto_col}'  |  fecha: '{fecha_col}'")

    def comparar_datos(self):
        """Compara Excel vs BD por (Fecha, Monto_Abs)."""
        if self.df_excel is None or self.df_excel.empty:
            messagebox.showwarning("Advertencia", "Primero procesá el archivo.")
            return
        if self.df_bd is None or self.df_bd.empty:
            messagebox.showwarning("Advertencia", "Primero consultá la BD.")
            return

        self.log("🔄 Comparando datos (Fecha, Monto absoluto)...")
        try:
            # Merge por llaves de comparación
            excel_cmp = self.df_excel[["Fecha", "Monto", "Monto_Abs", "Concepto"]].copy() if "Concepto" in self.df_excel.columns \
                else self.df_excel[["Fecha", "Monto", "Monto_Abs"]].copy()
            excel_cmp.rename(columns={"Fecha": "_Fecha_BD", "Monto_Abs": "_Monto_BD"}, inplace=True)

            # Many-to-many merge permite contar coincidencias
            merged = excel_cmp.merge(
                self.df_bd[["_Fecha_BD", "_Monto_BD"]],
                how="left", on=["_Fecha_BD", "_Monto_BD"], indicator=True
            )

            merged["Estado_BD"] = merged["_merge"].map({"both": "ENCONTRADO", "left_only": "NO ENCONTRADO", "right_only": "NO ENCONTRADO"})
            resumen = merged["Estado_BD"].value_counts()

            encontrados = int(resumen.get("ENCONTRADO", 0))
            total = len(merged)
            no_encontrados = total - encontrados
            tasa = (encontrados / total * 100) if total else 0.0

            self.log("📊 RESULTADOS DE COMPARACIÓN")
            self.log(f"✅ Encontrados en BD: {encontrados}")
            self.log(f"❌ No encontrados en BD: {no_encontrados}")
            self.log(f"📈 Tasa de coincidencia: {tasa:.1f}%")

            # DataFrame “bonito” de comparación
            cols_out = ["_Fecha_BD", "Monto", "Estado_BD"]
            if "Concepto" in merged.columns:
                cols_out.insert(1, "Concepto")

            self.df_comparacion = merged[cols_out].rename(columns={
                "_Fecha_BD": "Fecha_Excel",
                "Monto": "Monto_Excel"
            })

        except Exception as e:
            self.log(f"❌ Error en comparación: {e}")
            messagebox.showerror("Error", f"Error comparando datos:\n{e}")

    def exportar_resultados(self):
        """Exporta la comparación a un Excel (2 hojas: Comparacion y Resumen)."""
        if self.df_comparacion is None or self.df_comparacion.empty:
            messagebox.showwarning("Advertencia", "No hay resultados para exportar.")
            return

        try:
            filename = filedialog.asksaveasfilename(
                title="Guardar resultados como...",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
            )
            if not filename:
                return

            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                self.df_comparacion.to_excel(writer, sheet_name="Comparacion", index=False)

                encontrados = (self.df_comparacion["Estado_BD"] == "ENCONTRADO").sum()
                total = len(self.df_comparacion)
                resumen_df = pd.DataFrame({
                    "Total_Registros": [total],
                    "Encontrados_BD": [encontrados],
                    "No_Encontrados_BD": [total - encontrados],
                    "Tasa_Coincidencia": [f"{(encontrados / total * 100) if total else 0:.1f}%"]
                })
                resumen_df.to_excel(writer, sheet_name="Resumen", index=False)

            self.log(f"💾 Resultados exportados a: {filename}")
            messagebox.showinfo("Éxito", f"Resultados exportados a:\n{filename}")

        except Exception as e:
            self.log(f"❌ Error exportando resultados: {e}")
            messagebox.showerror("Error", f"Error exportando resultados:\n{e}")


def main():
    root = tk.Tk()
    # (Opcional) estilo nativo en Windows
    try:
        root.call("source", "sun-valley.tcl")
        root.call("set_theme", "light")
    except Exception:
        pass
    app = ComparadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
