import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from qvd import qvd_reader  # Suponiendo que tienes una función qvd_reader para leer archivos QVD

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Lector QVDs")

        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        self.btn_buscar = tk.Button(self.frame, text="Buscar archivo", command=self.buscar_archivo)
        self.btn_buscar.pack(pady=10)

        self.btn_guardar = tk.Button(self.frame, text="Guardar como XLSX", command=self.guardar_xlsx, state=tk.DISABLED)
        self.btn_guardar.pack(pady=10)

        self.tabla = ttk.Treeview(self.frame)
        self.tabla.pack(pady=10)

    def buscar_archivo(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos QVD", "*.qvd")])
        if ruta_archivo:
            try:
                self.df = qvd_reader.read(ruta_archivo)  # Suponiendo que tienes una función read_qvd para leer archivos QVD
                columnas = self.df.columns.tolist()  # Obtener los nombres de las columnas del DataFrame
                self.configurar_tabla(columnas)
                self.mostrar_datos()
                self.btn_guardar.config(state=tk.NORMAL)  # Habilitar el botón de Guardar
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

    def configurar_tabla(self, columnas):
        self.tabla["columns"] = tuple(columnas)
        for columna in columnas:
            self.tabla.heading(columna, text=columna)

    def mostrar_datos(self):
        self.limpiar_tabla()
        for index, row_data in self.df.iterrows():
            self.tabla.insert("", tk.END, values=tuple(row_data))

    def limpiar_tabla(self):
        for row in self.tabla.get_children():
            self.tabla.delete(row)

    def guardar_xlsx(self):
        ruta_guardar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos XLSX", "*.xlsx")])
        if ruta_guardar:
            try:
                self.df.to_excel(ruta_guardar, index=False)  # Guardar DataFrame en archivo XLSX
                messagebox.showinfo("Guardado", "Archivo XLSX guardado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
