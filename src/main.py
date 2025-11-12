import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

# Importa tus lectores
import lectorItau
import lectorBrou

def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo de estado de cuenta",
        filetypes=[("Archivos Excel", "*.xls *.xlsx"), ("Todos los archivos", "*.*")]
    )
    if ruta:
        entrada_archivo.delete(0, tk.END)
        entrada_archivo.insert(0, ruta)

def procesar_archivo():
    ruta = entrada_archivo.get()
    tipo = combo_tipo.get()

    if not ruta or not os.path.exists(ruta):
        messagebox.showerror("Error", "Seleccioná un archivo válido.")
        return

    if tipo == "Itaú":
        try:
            lectorItau.procesar_itau(ruta)
            messagebox.showinfo("Éxito", "Archivo Itaú procesado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Error procesando Itaú:\n{e}")

    elif tipo == "BROU":
        try:
            lectorBrou.procesar_brou(ruta)
            messagebox.showinfo("Éxito", "Archivo BROU procesado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Error procesando BROU:\n{e}")

    else:
        messagebox.showwarning("Atención", "Seleccioná el tipo de archivo (BROU o Itaú).")

# ----- Interfaz gráfica -----
root = tk.Tk()
root.title("Lector de Estados de Cuenta")
root.geometry("450x250")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

# Selector de archivo
ttk.Label(frame, text="Archivo:").grid(row=0, column=0, sticky="w", pady=10)
entrada_archivo = ttk.Entry(frame, width=40)
entrada_archivo.grid(row=0, column=1, padx=5)
ttk.Button(frame, text="Examinar...", command=seleccionar_archivo).grid(row=0, column=2)

# Selector de tipo
ttk.Label(frame, text="Tipo de archivo:").grid(row=1, column=0, sticky="w", pady=10)
combo_tipo = ttk.Combobox(frame, values=["Itaú", "BROU"], state="readonly", width=15)
combo_tipo.grid(row=1, column=1, sticky="w")
combo_tipo.set("Itaú")

# Botón procesar
ttk.Button(frame, text="Procesar", command=procesar_archivo).grid(row=3, column=1, pady=20)

root.mainloop()
