import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Análisis de Llamadas", layout="wide")

st.title("📞 Análisis de productividad de llamadas")

# Cargar archivo
archivo = st.file_uploader("Sube el archivo Excel", type=["xlsx"])
if archivo is not None:
    df = pd.read_excel(archivo)

    # Mostrar columnas disponibles para depuración
    st.write("Columnas detectadas:", df.columns.tolist())

    columnas_requeridas = ["Agent Name", "Call Start Time", "Call End Time", "Call Type"]
    faltantes = [col for col in columnas_requeridas if col not in df.columns]

    if faltantes:
        st.error(f"Faltan columnas necesarias en el archivo: {', '.join(faltantes)}")
        st.stop()

    # Convertir fechas
    df["Call Start Time"] = pd.to_datetime(df["Call Start Time"], errors="coerce")
    df["Call End Time"] = pd.to_datetime(df["Call End Time"], errors="coerce")

    if df["Call Start Time"].isna().all():
        st.error("Todos los valores en 'Call Start Time' son inválidos o están vacíos.")
        st.stop()

    # Crear columnas adicionales
    df["Fecha"] = df["Call Start Time"].dt.date
    df["Hora"] = df["Call Start Time"].dt.hour
    df["Duración (min)"] = (df["Call End Time"] - df["Call Start Time"]).dt.total_seconds() / 60
    df["Duración (min)"] = df["Duración (min)"].fillna(0)

    # Normalizar nombres
    df["Agent Name"] = df["Agent Name"].str.upper().str.strip()

    # Rango de fechas
    if df["Fecha"].dropna().empty:
        st.error("No se encontraron fechas válidas para filtrar.")
        st.stop()

    fecha_min, fecha_max = df["Fecha"].min(), df["Fecha"].max()
    fecha_inicio = st.date_input("Desde", fecha_min, min_value=fecha_min, max_value=fecha_max)
    fecha_fin = st.date_input("Hasta", fecha_max, min_value=fecha_min, max_value=fecha_max)

    if fecha_inicio > fecha_fin:
        st.warning("La fecha de inicio no puede ser mayor que la fecha de fin.")
        st.stop()

    df_filtrado = df[(df["Fecha"] >= fecha_inicio) & (df["Fecha"] <= fecha_fin)]

    if df_filtrado.empty:
        st.warning("No hay datos en el rango de fechas seleccionado.")
        st.stop()

    st.subheader("Resumen general")

    total_llamadas = df_filtrado.shape[0]
    total_entrantes = df_filtrado[df_filtrado["Call Type"].str.lower() == "incoming"].shape[0]
    total_salientes = df_filtrado[df_filtrado["Call Type"].str.lower() == "outgoing"].shape[0]
    duracion_promedio = df_filtrado["Duración (min)"].mean()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de llamadas", total_llamadas)
    col2.metric("Entrantes", total_entrantes)
    col3.metric("Salientes", total_salientes)
    col4.metric("Duración promedio (min)", f"{duracion_promedio:.2f}")

    st.divider()
    st.subheader("Llamadas por agente")

    llamadas_por_agente = df_filtrado.groupby("Agent Name").agg({
        "Duración (min)": ["count", "sum", "mean"]
    }).reset_index()
    llamadas_por_agente.columns = ["Agente", "Cantidad", "Total min", "Promedio min"]
    llamadas_por_agente = llamadas_por_agente.sort_values("Cantidad", ascending=False)

    st.dataframe(llamadas_por_agente, use_container_width=True)

    st.divider()
    st.subheader("Distribución por hora del día")

    llamadas_por_hora = df_filtrado.groupby("Hora").size()
    fig, ax = plt.subplots()
    llamadas_por_hora.plot(kind="bar", ax=ax)
    ax.set_xlabel("Hora del día")
    ax.set_ylabel("Número de llamadas")
    ax.set_title("Llamadas por hora")
    st.pyplot(fig)

    st.divider()
    st.subheader("Tendencia de llamadas por día")

    llamadas_por_fecha = df_filtrado.groupby("Fecha").size()
    fig, ax = plt.subplots()
    llamadas_por_fecha.plot(ax=ax)
    ax.set_xlabel("Fecha")
    ax.set_ylabel("Número de llamadas")
    ax.set_title("Tendencia diaria")
    st.pyplot(fig)
