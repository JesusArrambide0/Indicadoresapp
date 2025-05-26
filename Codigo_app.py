import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# Cargar datos desde Excel
archivo_excel = "AppInfo.xlsx"
df = pd.read_excel(archivo_excel)

# Convertir las columnas de tiempo a datetime
df["Call Start Time"] = pd.to_datetime(df["Call Start Time"])
df["Call End Time"] = pd.to_datetime(df["Call End Time"])
df["Talk Time (s)"] = (df["Call End Time"] - df["Call Start Time"]).dt.total_seconds()

# Normalizar nombres de agentes
df["Agent Name"] = df["Agent Name"].str.strip().str.upper()

# Calcular fechas mínimas y máximas para usar en el widget de fecha
fecha_inicio_default = df["Call Start Time"].min()
fecha_fin_default = df["Call Start Time"].max()

# Convertir a datetime.date para evitar errores en date_input
if isinstance(fecha_inicio_default, pd.Timestamp):
    fecha_inicio_default = fecha_inicio_default.date()
if isinstance(fecha_fin_default, pd.Timestamp):
    fecha_fin_default = fecha_fin_default.date()

# Widget para seleccionar rango de fechas
rango_fechas = st.date_input(
    "Selecciona el rango de fechas:",
    value=(fecha_inicio_default, fecha_fin_default),
    min_value=fecha_inicio_default,
    max_value=fecha_fin_default
)

# Si el usuario seleccionó una fecha única (no un rango)
if isinstance(rango_fechas, tuple):
    fecha_inicio, fecha_fin = rango_fechas
else:
    fecha_inicio = fecha_fin = rango_fechas

# Filtrar datos por el rango de fechas
df_filtrado = df[
    (df["Call Start Time"].dt.date >= fecha_inicio) &
    (df["Call Start Time"].dt.date <= fecha_fin)
]

# Mostrar métricas clave
st.subheader("Métricas generales")

llamadas_totales = len(df_filtrado)
llamadas_entrantes = len(df_filtrado[df_filtrado["Call Type"] == "Inbound"])
llamadas_salientes = len(df_filtrado[df_filtrado["Call Type"] == "Outbound"])
duracion_promedio = df_filtrado["Talk Time (s)"].mean() / 60 if llamadas_totales > 0 else 0

st.metric("Llamadas Totales", llamadas_totales)
st.metric("Entrantes", llamadas_entrantes)
st.metric("Salientes", llamadas_salientes)
st.metric("Duración Promedio (min)", round(duracion_promedio, 2))

# Mostrar datos si el usuario lo desea
if st.checkbox("Mostrar datos filtrados"):
    st.dataframe(df_filtrado)
