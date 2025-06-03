import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Análisis de Estados de Agentes", layout="wide")

st.title("📊 Análisis de Estados de Agentes")

# Leer el archivo desde la misma carpeta
archivo = "Estadosinfo.xlsx"
df = pd.read_excel(archivo)

# Normalizar nombres de columnas
df.columns = df.columns.str.strip()

# Asegurar que las columnas clave existen
columnas_esperadas = ['Agent Name', 'State Transition Time', 'Agent State', 'Reason', 'Duration']
if not all(col in df.columns for col in columnas_esperadas):
    st.error("El archivo no contiene todas las columnas necesarias.")
    st.stop()

# Renombrar columnas para facilitar el trabajo
df = df.rename(columns={
    'Agent Name': 'Agente',
    'State Transition Time': 'FechaHora',
    'Agent State': 'Estado',
    'Reason': 'Motivo',
    'Duration': 'Duración'
})

# Convertir fechas y obtener la fecha (sin hora)
df['FechaHora'] = pd.to_datetime(df['FechaHora'])
df['Fecha'] = df['FechaHora'].dt.date
df['Hora'] = df['FechaHora'].dt.time

# Convertir duración a horas
df['Duración'] = pd.to_timedelta(df['Duración'])
df['DuraciónHoras'] = df['Duración'].dt.total_seconds() / 3600

# Filtros de fecha y agente
fechas_min = df['Fecha'].min()
fechas_max = df['Fecha'].max()

st.sidebar.header("Filtros")
rango_fechas = st.sidebar.date_input("Selecciona rango de fechas", [fechas_min, fechas_max], min_value=fechas_min, max_value=fechas_max)
if len(rango_fechas) != 2:
    st.error("Selecciona un rango de fechas válido.")
    st.stop()

agentes = df['Agente'].unique().tolist()
agente_seleccionado = st.sidebar.selectbox("Selecciona agente (opcional)", options=["Todos"] + agentes)

# Filtrar datos según filtros
df_filtrado = df[(df['Fecha'] >= rango_fechas[0]) & (df['Fecha'] <= rango_fechas[1])]
if agente_seleccionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado['Agente'] == agente_seleccionado]

# --- NUEVO: Cálculo promedio Talk Time y Ring Time ---

# Supongamos que tienes otro DataFrame o en df_filtrado las columnas 'Talk Time' y 'Ring Time'
# En caso contrario, deberás cargar otro archivo o ajustar esta parte
df_tiempos = df_filtrado.copy()  # Cambia si tienes otro DF específico para talk/ring time

if 'Talk Time' in df_tiempos.columns:
    df_tiempos['Talk Time'] = pd.to_timedelta(df_tiempos['Talk Time'], errors='coerce')
    talktime_seconds = df_tiempos['Talk Time'].dt.total_seconds().dropna()
    promedio_talktime = talktime_seconds.mean()
    st.write(f"Promedio de Talk Time para {agente_seleccionado}: **{promedio_talktime:.2f} segundos**")
else:
    st.warning("No se encontró la columna 'Talk Time' en el archivo.")

if 'Ring Time' in df_tiempos.columns:
    df_tiempos['Ring Time'] = pd.to_timedelta(df_tiempos['Ring Time'], errors='coerce')
    ringtime_seconds = df_tiempos['Ring Time'].dt.total_seconds().dropna()
    promedio_ringtime = ringtime_seconds.mean()
    st.write(f"Promedio de Ring Time para {agente_seleccionado}: **{promedio_ringtime:.2f} segundos**")
else:
    st.warning("No se encontró la columna 'Ring Time' en el archivo.")

# Calcular el primer "Logged-in" del día por agente
logged = df_filtrado[df_filtrado['Estado'].str.lower() == 'logged-in'].copy()
primer_logged = logged.sort_values(by='FechaHora').groupby(['Agente', 'Fecha']).first().reset_index()
primer_logged['Hora Entrada'] = primer_logged['FechaHora'].dt.time

# Reglas de horario esperado
horarios = {
    'Jonathan Alejandro Zúñiga': 12,
    'Jesús Armando Arrambide': 8,
    'Maria Teresa Loredo Morales': 10,
    'Jorge Cesar Flores Rivera': 8
}

# Función para verificar si hubo retraso
def es_retraso(row):
    esperado = horarios.get(row['Agente'], 8)
    return row['FechaHora'].hour >= esperado

if not primer_logged.empty:
    primer_logged['Retraso'] = primer_logged.apply(es_retraso, axis=1)
else:
    primer_logged['Retraso'] = pd.Series(dtype=bool)

# Tiempo total por estado por agente y día
tiempo_por_estado = df_filtrado.groupby(['Agente', 'Fecha', 'Estado'])['DuraciónHoras'].sum().reset_index()

# Tabla pivote para mostrar el tiempo distribuido por estado
tiempo_pivot = tiempo_por_estado.pivot_table(index=['Agente', 'Fecha'], columns='Estado', values='DuraciónHoras', fill_value=0).reset_index()

# Resumen general
resumen_agente = df_filtrado.groupby('Agente')['DuraciónHoras'].sum().reset_index(name='Total de Horas')
resumen_agente = resumen_agente.sort_values(by='Total de Horas', ascending=False)

# Mostrar tablas
st.subheader("📌 Resumen de tiempo total por agente")
st.dataframe(resumen_agente, use_container_width=True)

st.subheader("🕓 Primer ingreso (Logged-in) y retrasos")
if not primer_logged.empty:
    styled_df = primer_logged[['Agente', 'Fecha', 'Hora Entrada', 'Retraso']].style.applymap(
        lambda x: 'background-color: #ff9999; font-weight: bold;' if x else '',
        subset=['Retraso']
    )
    st.dataframe(styled_df, use_container_width=True)
else:
    st.warning("No se encontraron registros de primer Logged-in para los filtros aplicados.")

st.subheader("⏱️ Tiempo invertido por estado por día")
st.dataframe(tiempo_pivot, use_container_width=True)

# Gráfica: Tiempo invertido por estado para el agente o todos
fig = px.bar(tiempo_por_estado, x='Estado', y='DuraciónHoras', color='Estado',
             labels={'DuraciónHoras': 'Horas', 'Estado': 'Estado'}, title="Tiempo total invertido por Estado")
fig.update_traces(texttemplate='%{y:.2f}', textposition='outside')
fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
st.plotly_chart(fig, use_container_width=True)

# Análisis porcentual de tiempo por estado
total_horas = tiempo_por_estado['DuraciónHoras'].sum()
if total_horas > 0:
    resumen_estado = tiempo_por_estado.groupby('Estado')['DuraciónHoras'].sum().reset_index()
    resumen_estado['Porcentaje'] = 100 * resumen_estado['DuraciónHoras'] / total_horas
    resumen_estado = resumen_estado.sort_values(by='Porcentaje', ascending=False)
    st.subheader("📈 Análisis porcentual de tiempo invertido por estado")
    st.dataframe(resumen_estado.style.format({'Porcentaje': '{:.2f}%'}), use_container_width=True)
else:
    st.info("No hay datos para mostrar el análisis porcentual.")
