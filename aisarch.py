import pandas as pd
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox
from tkinter.ttk import Combobox
from tempfile import NamedTemporaryFile
import shutil
import os

class AisladorColumnaGUI:
    def __init__(self, master):
        self.master = master
        master.title("Aislador de Columna")
        master.geometry("690x150")

        # Colors
        background_color = "#E0E0E0"
        button_color = "#4CAF50"

        master.configure(bg=background_color)

        # File Selection Section
        self.label_origen = Label(master, text="Archivo Excel de origen:", bg=background_color)
        self.label_origen.grid(row=0, column=0, padx=10, pady=5)

        self.entry_origen = Entry(master, width=30)
        self.entry_origen.grid(row=0, column=1, padx=10, pady=5)

        self.button_examinar = Button(master, text="Examinar", command=self.seleccionar_archivo_origen, bg=button_color)
        self.button_examinar.grid(row=0, column=2, padx=10, pady=5)

        # Column Selection Section
        self.label_columna = Label(master, text="Seleccionar Columna:", bg=background_color)
        self.label_columna.grid(row=1, column=0, padx=10, pady=5)

        self.combo_columna = Combobox(master, width=27, state="readonly")
        self.combo_columna.grid(row=1, column=1, padx=10, pady=5)

        # Destination Section
        self.label_destino = Label(master, text="Nombre predeterminado para el archivo Excel de destino:", bg=background_color)
        self.label_destino.grid(row=2, column=0, padx=10, pady=5)

        self.entry_destino = Entry(master, width=30)
        self.entry_destino.grid(row=2, column=1, padx=10, pady=5)
        self.entry_destino.insert(0, "aislado")

        # Button to Isolate Column
        self.button_aislar = Button(master, text="Aislar Columna", command=self.aislar_columna, bg=button_color)
        self.button_aislar.grid(row=3, column=1, pady=10)

    def seleccionar_archivo_origen(self):
        archivo_origen = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        self.entry_origen.delete(0, "end")
        self.entry_origen.insert(0, archivo_origen)

        # Load available columns into the ComboBox
        try:
            df = pd.read_excel(archivo_origen)
            columnas_disponibles = list(df.columns)
            self.combo_columna["values"] = columnas_disponibles
        except FileNotFoundError:
            messagebox.showerror("Error", "Archivo no encontrado. Asegúrate de que el archivo exista y vuelve a intentarlo.")
            self.combo_columna.set("")

    def aislar_columna(self):
        archivo_origen = self.entry_origen.get()
        columna_seleccionada = self.combo_columna.get()
        nombre_predeterminado = self.entry_destino.get()
        extension_destino = ".xlsx"

        try:
            df = pd.read_excel(archivo_origen)
        except FileNotFoundError:
           messagebox.showerror("Error", "Archivo no encontrado. Asegúrate de que el archivo exista y vuelve a intentarlo.")
           return

           # Check if the selected column exists in the DataFrame
           if columna_seleccionada not in df.columns:
               messagebox.showerror("Error", f"La columna '{columna_seleccionada}' no existe en el archivo de origen.")
               return

        # Crear un DataFrame con solo la columna seleccionada
        df_resultado = pd.DataFrame({columna_seleccionada: df[columna_seleccionada]})
        
        # Agregar la extensión ".pdf" solo a los datos de la columna seleccionada
        df_resultado[columna_seleccionada] = df_resultado[columna_seleccionada].astype(str) + ".pdf"

        # Obtener la ruta del archivo original y construir la ruta para el nuevo archivo
        ruta_original = os.path.dirname(archivo_origen)
        destino_final = os.path.join(ruta_original, f"{nombre_predeterminado}{extension_destino}")

           # Guardar el DataFrame resultante como un archivo Excel
        df_resultado.to_excel(destino_final, index=False)

        messagebox.showinfo("Éxito", f"La extensión '.pdf' se ha agregado a la columna '{columna_seleccionada}'. El archivo '{destino_final}' ha sido creado.")

root = Tk()
app = AisladorColumnaGUI(root)
root.mainloop()