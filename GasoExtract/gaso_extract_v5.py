import re
import pandas as pd
from pdfminer.high_level import extract_text
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.constants import *
from ttkbootstrap import Frame, Button, Label, Treeview
from tkinter.ttk import Progressbar
import os

def parse_pdf(file_path):
    text = extract_text(file_path)

    # Patrones para capturar: ID, PLACAS, TOTAL VEHÍCULO
    vehicle_pattern = re.compile(
        r'(\d{1,5})\s+.*?\n.*?\n([\w\d-]{3,15})\n.*?Total vehículo - \d+ Cargas\n([\d.,]+)\n([\d.,]+)', re.DOTALL)

    results = []
    for match in vehicle_pattern.findall(text):
        unit_id, placa, litros, total = match
        results.append({
            'Unidad': unit_id.strip(),
            'Placas': placa.strip(),
            'Litros Totales': float(litros.replace('.', '').replace(',', '.').strip()),
            'Total Gastado (MXN)': float(total.replace('.', '').replace(',', '.').strip())
        })

    return results


def open_pdf():
    file_path = filedialog.askopenfilename(
        title="Selecciona un archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if not file_path:
        return

    loading_label.config(text="Cargando...")
    root.update_idletasks()

    data = parse_pdf(file_path)

    loading_label.config(text="")
    progress_bar.stop()

    if not data:
        messagebox.showwarning("Sin datos", "No se encontraron datos en el archivo.")
        return

    df = pd.DataFrame(data)

    # Guardar como Excel
    output_file = os.path.splitext(file_path)[0] + "_resumen.xlsx"
    df.to_excel(output_file, index=False)

    # Mostrar en la tabla
    for row in tree.get_children():
        tree.delete(row)

    for _, row in df.iterrows():
        tree.insert("", "end", values=(row['Unidad'], row['Placas'], row['Litros Totales'], row['Total Gastado (MXN)']))

    messagebox.showinfo("Éxito", f"Resumen generado y guardado en:\n{output_file}")


# ---------- Interfaz ----------
root = tk.Tk()
root.title("Resumen de Gasolina")
root.geometry("850x500")
root.iconbitmap("gas.ico")

style = Style("cosmo")

frame = Frame(root, padding=20)
frame.pack(fill=BOTH, expand=True)

Label(frame, text="Selecciona un PDF para generar el resumen", font=("Open Sans", 12)).pack(pady=10)
Button(frame, text="Abrir PDF", bootstyle=PRIMARY, command=open_pdf).pack(pady=5)

loading_label = Label(frame, text="", font=("Open Sans", 10))
loading_label.pack(pady=5)

progress_bar = Progressbar(frame, mode='indeterminate', length=300)
progress_bar.pack(pady=5)

# Tabla
cols = ("Unidad", "Placas", "Litros Totales", "Total Gastado (MXN)")
tree = Treeview(frame, columns=cols, show="headings", bootstyle=INFO)

for col in cols:
    tree.heading(col, text=col)
    tree.column(col, anchor="center")

tree.pack(fill=BOTH, expand=True, pady=10)

root.mainloop()
