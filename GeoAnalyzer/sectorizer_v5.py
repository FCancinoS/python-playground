import pandas as pd
import json
from shapely.geometry import Point, shape
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import os
from math import cos, radians

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
    '864032050286192': '43', 
    '864032050253812': '26',
    '864032050254703': 'A',
    '864032050255577': 'B',
    '864032050286036': 'C',
    '864032050292877': 'D',
    '864032050258084': 'E'
}

BUFFER_METROS = 30.0     # tolerancia de borde
NEAREST_METROS = 100.0   # asignación automática al más cercano
SIN_SECTOR = "Fuera de sectores"
def _normalize_columns(df):
    # quita espacios invisibles y bordes
    df.columns = (df.columns.astype(str)
                  .str.replace(r'[\u200b\u200c\u200d\u00a0]', '', regex=True)
                  .str.strip())
    return df

def _pick(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _parse_coords_series(s):
    # limpia espacios y NBSP/ZWSP
    s = (s.astype(str)
           .str.replace(r'[\s\u200b\u200c\u200d\u00a0]', '', regex=True))
    mask = s.str.match(r'^-?\d+(\.\d+)?,-?\d+(\.\d+)?$')
    return s, mask

def segmentar_hora(dt):
    return dt.replace(minute=0, second=0) if dt.minute < 30 else dt.replace(minute=30, second=0)

def detectar_sector(lat, lon, poligonos):
    """
    1) poligono.covers(punto)  -> incluye frontera
    2) punto con buffer ~30 m  -> intersecta
    3) polígono más cercano si distancia ≤ 100 m
    """
    punto = Point(lon, lat)

    # 1) Dentro o en frontera
    for nombre, poligono in poligonos:
        try:
            if poligono.covers(punto):
                return nombre
        except Exception:
            continue

    # 2) Tolerancia pequeña (buffer ~30 m) en grados aprox a la latitud
    deg_lat_m = 111_000.0
    deg_lon_m = max(111_000.0 * cos(radians(lat)), 1e-6)
    buf_lon = BUFFER_METROS / deg_lon_m
    buf_lat = BUFFER_METROS / deg_lat_m
    punto_buf = punto.buffer(max(buf_lon, buf_lat))

    for nombre, poligono in poligonos:
        try:
            if poligono.intersects(punto_buf):
                return nombre
        except Exception:
            continue

    # 3) Polígono más cercano (si suficientemente cerca)
    mejor_nombre, mejor_m = None, float("inf")
    for nombre, poligono in poligonos:
        try:
            d_deg = poligono.distance(punto)   # en grados (no geodésico, pero OK para distancias cortas)
            d_m   = d_deg * 111_000.0
            if d_m < mejor_m:
                mejor_nombre, mejor_m = nombre, d_m
        except Exception:
            continue

    if mejor_nombre is not None and mejor_m <= NEAREST_METROS:
        return mejor_nombre

    return SIN_SECTOR

def procesar_archivos(lista_excels, path_geojson, path_guardado):
    try:
        # Asegurar carpeta destino
        os.makedirs(os.path.dirname(path_guardado), exist_ok=True)

        # Cargar GeoJSON (polígonos)
        with open(path_geojson, "r", encoding="utf-8") as f:
            geojson_data = json.load(f)

        poligonos = []
        for feature in geojson_data["features"]:
            nombre = feature.get("properties", {}).get("name") or feature.get("name")
            geom = shape(feature["geometry"])
            poligonos.append((nombre, geom))

        with pd.ExcelWriter(path_guardado, engine='openpyxl') as writer:
            # --- Semilla de la hoja 'Resumen' para que siempre exista ---
            pd.DataFrame([["Reporte de Sectores"]]).to_excel(
                writer, sheet_name="Resumen", startrow=0, index=False, header=False
            )
            start_row = 2  # siguiente fila libre en 'Resumen'

            for path_excel in lista_excels:
                # --- LECTURA ROBUSTA ---
                try:
                    df = pd.read_excel(
                        path_excel,
                        header=1,
                        engine=('openpyxl' if path_excel.lower().endswith('.xlsx') else 'xlrd')
                    )
                except Exception:
                    df = pd.read_excel(
                        path_excel,
                        header=0,
                        engine=('openpyxl' if path_excel.lower().endswith('.xlsx') else 'xlrd')
                    )

                # Normalizar encabezados
                df = _normalize_columns(df)

                # Buscar columnas clave
                col_time   = _pick(df, ["Position Time", "PositionTime", "GPS Time", "Time"])
                col_coords = _pick(df, ["Coordinates", "Coordinate", "Lat/Lon", "LatLon", "Location"])
                col_imei   = _pick(df, ["IMEI", "Imei", "imei"])

                # IMEI / unidad
                imei = "N/A"
                if col_imei and df[col_imei].notna().any():
                    imei = str(df[col_imei].dropna().iloc[0])
                solid_device = imei_mapping.get(imei, "Desconocido")
                l6 = imei[-6:] if imei and imei != "N/A" else "N/A"
                encabezado = f"Unidad: {solid_device} - IMEI: {imei} - L6: {l6}"

                # Validaciones de columnas
                if not col_time or not col_coords:
                    pd.DataFrame([[encabezado], ["Sin columnas de tiempo/coords detectadas"]]) \
                        .to_excel(writer, sheet_name="Resumen", startrow=start_row, index=False, header=False)
                    start_row += 3
                    continue

                # Filtrar no nulos
                df = df[df[col_time].notna() & df[col_coords].notna()]
                if df.empty:
                    pd.DataFrame([[encabezado], ["Sin filas con tiempo y coordenadas"]]) \
                        .to_excel(writer, sheet_name="Resumen", startrow=start_row, index=False, header=False)
                    start_row += 3
                    continue

                # Limpiar y validar coords
                coords, mask = _parse_coords_series(df[col_coords])
                df = df[mask].copy()
                if df.empty:
                    pd.DataFrame([[encabezado], ["Coordenadas con formato inválido (no 'lat,lon')"]]) \
                        .to_excel(writer, sheet_name="Resumen", startrow=start_row, index=False, header=False)
                    start_row += 3
                    continue

                # Split a floats
                df[['Lat', 'Lon']] = coords.str.split(',', expand=True).astype(float)

                # Parseo de fecha/hora
                df["Hora"] = pd.to_datetime(df[col_time], errors='coerce')
                df = df[df["Hora"].notna()]
                if df.empty:
                    pd.DataFrame([[encabezado], ["No se pudieron parsear fechas/horas (NaT)"]]) \
                        .to_excel(writer, sheet_name="Resumen", startrow=start_row, index=False, header=False)
                    start_row += 3
                    continue

                # Segmento y sector
                df["Segmento"] = df["Hora"].apply(segmentar_hora)
                df["Sector"] = df.apply(lambda r: detectar_sector(r["Lat"], r["Lon"], poligonos), axis=1)      

                conteo = df.groupby(["Segmento", "Sector"]).size().reset_index(name="Conteo")
                reales = conteo[conteo["Sector"] != SIN_SECTOR]
                sind   = conteo[conteo["Sector"] == SIN_SECTOR]

                if not reales.empty:
                    mayores = (reales.sort_values(["Segmento", "Conteo"], ascending=[True, False])
                                    .drop_duplicates("Segmento", keep="first"))
                    seg_con_reales = set(mayores["Segmento"])
                    faltantes = sind[~sind["Segmento"].isin(seg_con_reales)]
                    if not faltantes.empty:
                        faltantes = (faltantes.sort_values(["Segmento", "Conteo"], ascending=[True, False])
                                            .drop_duplicates("Segmento", keep="first"))
                        mayores = pd.concat([mayores, faltantes], ignore_index=True).sort_values("Segmento")
                else:
                    mayores = (sind.sort_values(["Segmento", "Conteo"], ascending=[True, False])
                                    .drop_duplicates("Segmento", keep="first"))

                mayores['Unidad'] = solid_device

                # --- ESCRITURA EN 'Resumen' (¡lo que te faltaba!) ---
                pd.DataFrame([[encabezado]]).to_excel(
                    writer, sheet_name="Resumen", startrow=start_row, index=False, header=False
                )
                start_row += 1

                mayores[["Segmento", "Sector"]].to_excel(
                    writer, sheet_name="Resumen", startrow=start_row, index=False
                )
                start_row += len(mayores) + 2  # separador entre bloques

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
