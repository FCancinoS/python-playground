import pandas as pd
import json
from shapely.geometry import Point, shape
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import os
import matplotlib.pyplot as plt
import math


imei_mapping = {
    '864032050292919': '50',
    '864032050263878': '51',
    '864032050286788': '39',
    '864032050292752': '41',
    '864032050292554': '46',
    '864032050286259': '45',
    '864032050287448': '44',
    '864032050253705': '48',
    '864032050253770': '47',
    '864032050254638': '42',
    '864032050253804': '49',
    '864032050258589': '40',
    '864032050286192': 'SN1', 
    '864032050253812': 'SN2',
    '864032050254703': 'SN3'
}


def segmentar_hora(dt):
    return dt.replace(minute=0, second=0) if dt.minute < 30 else dt.replace(minute=30, second=0)

def detectar_sector(lat, lon, poligonos):
    punto = Point(lon, lat)
    for nombre, poligono in poligonos:
        if poligono.contains(punto):
            return nombre
    return "Sin datos"

def procesar_archivos(lista_excels, path_geojson, path_guardado):
    try:
        with open(path_geojson, "r", encoding="utf-8") as f:
            geojson_data = json.load(f)

        poligonos = [(feature.get("properties", {}).get("name") or feature.get("name"), shape(feature["geometry"])) for feature in geojson_data["features"]]

        graficos_datos = []

        with pd.ExcelWriter(path_guardado, engine='openpyxl') as writer:
            start_row = 0

            for path_excel in lista_excels:
                df = pd.read_excel(path_excel, header=1, engine='openpyxl' if path_excel.endswith('.xlsx') else 'xlrd')

                imei = str(df['IMEI'].dropna().iloc[0])
                solid_device = imei_mapping.get(imei, "Desconocido")
                l6 = imei[-6:] if imei else "N/A"

                encabezado = f"Unidad: {solid_device} - IMEI: {imei} - L6: {l6}"

                df = df[df["Coordinates"].notnull() & df["Position Time"].notnull()]
                df[['Lat', 'Lon']] = df['Coordinates'].str.split(',', expand=True).astype(float)
                df["Hora"] = pd.to_datetime(df["Position Time"])
                df["Segmento"] = df["Hora"].apply(segmentar_hora)
                df["Sector"] = df.apply(lambda row: detectar_sector(row["Lat"], row["Lon"], poligonos), axis=1)

                conteo = df.groupby(["Segmento", "Sector"]).size().reset_index(name="Conteo")
                mayores = conteo.sort_values("Conteo", ascending=False).drop_duplicates("Segmento").sort_values("Segmento")
                mayores['Unidad'] = solid_device
                graficos_datos.append(mayores)

                pd.DataFrame([[encabezado]]).to_excel(writer, sheet_name="Resumen", startrow=start_row, index=False, header=False)
                start_row += 1

                mayores[["Segmento", "Sector"]].to_excel(writer, sheet_name="Resumen", startrow=start_row, index=False)
                start_row += len(mayores) + 2

        # Graficas
        if graficos_datos:
            df_grafica = pd.concat(graficos_datos)
            sectores_unicos = sorted(df_grafica['Sector'].unique())
            sector_to_num = {sector: i for i, sector in enumerate(sectores_unicos)}
            df_grafica['Sector_Num'] = df_grafica['Sector'].map(sector_to_num)

            unidades = sorted(df_grafica['Unidad'].unique())
            unidades_por_grafica = math.ceil(len(unidades) / 3)

            for i in range(0, len(unidades), unidades_por_grafica):
                unidades_grupo = unidades[i:i + unidades_por_grafica]
                plt.figure(figsize=(12, 6))

                for unidad in unidades_grupo:
                    df_unidad = df_grafica[df_grafica['Unidad'] == unidad]
                    plt.plot(df_unidad['Segmento'], df_unidad['Sector_Num'], marker='o', label=f"Unidad {unidad}")

                plt.yticks(list(sector_to_num.values()), list(sector_to_num.keys()))
                plt.xlabel("Segmento horario")
                plt.ylabel("Sector")
                plt.title("Sectores por unidad")
                plt.legend()
                plt.grid(True)

                nombre_grafica = os.path.splitext(path_guardado)[0] + f'_grafica_{(i // unidades_por_grafica) + 1}.png'
                plt.savefig(nombre_grafica, bbox_inches='tight')
                plt.close()

        messagebox.showinfo("Éxito", "El análisis fue completado correctamente.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def seleccionar_excel():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if files:
        entrada_var.set(';'.join(files))

def seleccionar_geojson():
    file = filedialog.askopenfilename(filetypes=[("GeoJSON files", "*.geojson")])
    if file:
        geojson_var.set(file)

def guardar_como():
    file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file:
        salida_var.set(file)

def ejecutar():
    if not entrada_var.get() or not geojson_var.get() or not salida_var.get():
        messagebox.showwarning("Campos incompletos", "Selecciona todos los archivos necesarios.")
        return
    procesar_archivos(entrada_var.get().split(';'), geojson_var.get(), salida_var.get())

# Interfaz
app = ttkb.Window(themename="flatly")
app.title("Sectorizer - GeoJSON Personalizado")

try:
    icon_path = os.path.join(os.path.dirname(__file__), "../logo.ico")
    app.iconbitmap(icon_path)
except Exception:
    pass

app.geometry("500x500")
entrada_var = tk.StringVar()
geojson_var = tk.StringVar()
salida_var = tk.StringVar()

ttkb.Label(app, text="Archivo Excel de entrada:").pack(pady=(10, 0))
ttkb.Entry(app, textvariable=entrada_var, width=20).pack()
ttkb.Button(app, text="Seleccionar Excel", command=seleccionar_excel, bootstyle=INFO).pack(pady=(5, 10))

ttkb.Label(app, text="Archivo GeoJSON de sectores:").pack()
ttkb.Entry(app, textvariable=geojson_var, width=20).pack()
ttkb.Button(app, text="Seleccionar GeoJSON", command=seleccionar_geojson, bootstyle=INFO).pack(pady=(5, 10))

ttkb.Label(app, text="Guardar resultado como:").pack()
ttkb.Button(app, text="Seleccionar destino", command=guardar_como, bootstyle=INFO).pack(pady=(5, 10))
ttkb.Entry(app, textvariable=salida_var, width=20).pack()

ttkb.Button(app, text="Analizar", command=ejecutar, bootstyle=SUCCESS).pack(pady=10)

ttkb.Label(app, text="fcancino solutions", font=("Arial", 9), foreground="gray").pack(side="bottom", pady=5)

app.mainloop()
