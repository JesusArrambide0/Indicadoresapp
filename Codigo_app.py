import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import numpy as np

# Cargar datos
@st.cache_data
def cargar_datos():
    df = pd.read_excel("AppInfo.xlsx")

    # Limpiar y convertir columnas
    df['Call Start Time'] = pd.to_datetime(df['Call Start Time'])
    df['Call End Time'] = pd.to_datetime(df['Call End Time'])
    df['Talk Time'] = pd.to_timedelta(df['Talk Time'])

    # Normalizar nombres de agentes
    df['Agent Name'] = df['Agent Name'].str.strip().str.title()

    return df

df = cargar_datos()

# Filtro por fecha
st.sidebar.date_input("Fecha inicio", key="fecha_inicio", value=datetime(2025, 5, 19))
st.sidebar.date_input("Fecha fin", key="fecha_fin", value=datetime(2025, 5, 24))
fecha_inicio = pd.to_datetime(st.session_state.fecha_inicio)
fecha_fin = pd.to_datetime(st.session_state.fecha_fin)

df_filtrado = df[(df['Call Start Time'].dt.date >= fecha_inicio.date()) &
                 (df['Call Start Time'].dt.date <= fecha_fin.date())]

# Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Resumen General",
    "🕒 Distribución Hora/Día",
    "❌ Llamadas Perdidas",
    "🧑‍💼 Por Agente",
    "🔥 Heatmaps"
])

# 📊 TAB 1: Resumen General
with tab1:
    st.title("Resumen General")

    total_llamadas = len(df_filtrado)
    agentes = df_filtrado['Agent Name'].nunique()
    duracion_promedio = df_filtrado['Talk Time'].mean()

    col1, col2, col3 = st.columns(3)
    col1.metric("Total de llamadas", total_llamadas)
    col2.metric("Agentes únicos", agentes)
    col3.metric("Duración promedio", str(duracion_promedio))

    # Llamadas por agente
    st.subheader("Llamadas por agente")
    llamadas_por_agente = df_filtrado['Agent Name'].value_counts()
    fig1, ax1 = plt.subplots()
    llamadas_por_agente.plot(kind='bar', ax=ax1)
    plt.xticks(rotation=0)
    st.pyplot(fig1)

# 🕒 TAB 2: Distribución por Hora y Día
with tab2:
    st.title("Distribución por hora y día")

    df_filtrado['Hora'] = df_filtrado['Call Start Time'].dt.hour
    df_filtrado['Día'] = df_filtrado['Call Start Time'].dt.day_name()

    fig2, ax2 = plt.subplots()
    sns.histplot(data=df_filtrado, x='Hora', hue='Call Type', multiple='stack', ax=ax2)
    st.pyplot(fig2)

    fig3, ax3 = plt.subplots()
    orden_dias = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    sns.countplot(data=df_filtrado, x='Día', order=orden_dias, hue='Call Type', ax=ax3)
    plt.xticks(rotation=45)
    st.pyplot(fig3)

# ❌ TAB 3: Llamadas Perdidas
with tab3:
    st.title("Llamadas Perdidas")

    perdidas = df_filtrado[df_filtrado['Call Type'].str.contains("Missed", na=False)]
    total_perdidas = len(perdidas)

    st.metric("Llamadas Perdidas", total_perdidas)

    if total_perdidas > 0:
        st.subheader("Detalle de llamadas perdidas")
        st.dataframe(perdidas[['Call Start Time', 'Agent Name']])

# 🧑‍💼 TAB 4: Análisis por Agente
with tab4:
    st.title("Análisis por Agente")

    duracion_promedio = df_filtrado.groupby('Agent Name')['Talk Time'].mean()
    fig4, ax4 = plt.subplots()
    duracion_promedio.dt.total_seconds().div(60).plot(kind='bar', ax=ax4, color='skyblue')
    plt.ylabel("Duración promedio (minutos)")
    plt.xticks(rotation=0)
    st.pyplot(fig4)

# 🔥 TAB 5: Heatmap de llamadas por hora y día
with tab5:
    st.title("Heatmap de llamadas")

    heatmap_data = df_filtrado.copy()
    heatmap_data['Hora'] = heatmap_data['Call Start Time'].dt.hour
    heatmap_data['Día'] = heatmap_data['Call Start Time'].dt.day_name()

    pivot_table = heatmap_data.pivot_table(index='Día', columns='Hora', values='Call Type', aggfunc='count').fillna(0)

    # Ordenar días
    pivot_table = pivot_table.reindex(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])

    fig5, ax5 = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_table, cmap="YlGnBu", annot=True, fmt=".0f", ax=ax5)
    plt.title("Llamadas por Hora y Día")
    st.pyplot(fig5)
