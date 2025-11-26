# main.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime  # <-- agregado
import pandas as pd

# Módulos propios
import lectorItau
import lectorBrou
import db
import comparador


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
        ttk.Button(actions, text="Exportar coincidencias", command=self.exportar_comparacion).grid(row=0, column=1, padx=4, pady=4)

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
        if not self.procesar_archivo():
            return
        if not self.consultar_bd():
            return
        self.comparar_datos()

    # ----------------- LECTURA EXCEL -----------------
    def procesar_archivo(self):
        ruta = self.entrada_archivo.get().strip()
        tipo = self.combo_tipo.get().strip()

        if not ruta or not os.path.exists(ruta):
            messagebox.showerror("Error", "Seleccioná un archivo válido.")
            return False

        self.log(f"📁 Procesando archivo: {ruta}")

        try:
            if tipo == "Itaú":
                self.df_excel = lectorItau.procesar_itau(ruta)
            else:
                self.df_excel = lectorBrou.procesar_brou(ruta)

            if self.df_excel is None or self.df_excel.empty:
                self.log("⚠️ El lector devolvió un DataFrame vacío.")
                return False

            self.log(f"✅ Archivo {tipo} procesado ({len(self.df_excel)} filas)")
            return True

        except Exception as e:
            self.log(f"❌ Error procesando archivo: {e}")
            messagebox.showerror("Error", str(e))
            return False
        
    # ----------------- LECTURA BD -----------------
    def consultar_bd(self):
        self.log("🔌 Conectando a la base de datos...")

        try:
            if hasattr(db, "probar_conexion") and not db.probar_conexion():
                self.log("❌ No hay conexión con la base.")
                return False

            banco = self.combo_tipo.get().strip().lower()

            # Logica banco -> cod_tit
            if banco == "itaú" or banco == "itau":
                cod_tit = "113"
            elif banco == "brou":
                cod_tit = "001"
            else:
                messagebox.showerror("Error", f"Banco desconocido: {banco}")
                return False

            # Llamar ahora sí al método nuevo
            self.df_bd = db.obtener_df_bd(cod_tit)

            if self.df_bd is None or self.df_bd.empty:
                self.log(f"⚠️ La BD no devolvió registros para cod_tit={cod_tit}.")
                return False

            self.log(f"✅ BD cargada ({len(self.df_bd)} filas) para banco {banco.upper()} con cod_tit={cod_tit}")
            return True

        except Exception as e:
            self.log(f"❌ Error consultando BD: {e}")
            messagebox.showerror("Error", str(e))
            return False

    # ----------------- COMPARACIÓN -----------------
    def comparar_datos(self):
        try:
            self.log("🔄 Comparando datos...")

            self.df_comparacion = comparador.comparar(self.df_excel, self.df_bd)

            self.log(f"📊 Coincidencias encontradas: {len(self.df_comparacion)}")

            if len(self.df_comparacion) == 0:
                self.log("⚠️ No hubo coincidencias (fecha + monto).")
            else:
                self.log("✅ Comparación completada.")

        except Exception as e:
            self.log(f"❌ Error en comparación: {e}")
            messagebox.showerror("Error", str(e))

    # ----------------- EXPORTACIÓN -----------------
    def exportar_comparacion(self):
        if self.df_comparacion is None:
            messagebox.showwarning("Advertencia", "No hay datos para exportar.")
            return

        try:
            # Banco para el nombre
            banco = self.combo_tipo.get().strip() or "Banco"
            banco_sanitizado = banco.replace(" ", "").upper()  # Itaú -> ITAÚ (queda ITAÚ pero en nombre de archivo se acepta)
            # Fecha actual
            fecha_str = datetime.now().strftime("%Y_%m_%d_%H%M%S")

            # Directorio base: mismo que el archivo de entrada (si existe) o cwd
            ruta_entrada = self.entrada_archivo.get().strip()
            if ruta_entrada:
                base_dir = os.path.dirname(ruta_entrada)
            else:
                base_dir = os.getcwd()

            nombre_archivo = f"ConciliacionBancaria_{banco_sanitizado}_{fecha_str}.xlsx"
            ruta = os.path.join(base_dir, nombre_archivo)

            # Llamamos al comparador para generar y exportar las coincidencias
            comparador.comparar_y_exportar(self.df_excel, self.df_bd, ruta)

            self.log(f"💾 Archivo exportado: {ruta}")
            messagebox.showinfo("Éxito", f"Archivo exportado:\n{ruta}")

        except Exception as e:
            self.log(f"❌ Error exportando: {e}")
            messagebox.showerror("Error", str(e))


def main():
    root = tk.Tk()
    app = ComparadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
