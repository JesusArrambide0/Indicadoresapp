import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Cargar datos
archivo_excel = "AppInfo.xlsx"
df = pd.read_excel(archivo_excel, engine="openpyxl")

# Normalizar nombres
mapeo_a_nombre_completo = {
    "Jorge": "Jorge Cesar Flores Rivera",
    "Maria": "Maria Teresa Loredo Morales",
    "Jonathan": "Jonathan Alejandro Zúñiga",
}
df["Agent Name"] = df["Agent Name"].replace(mapeo_a_nombre_completo)

# Preprocesamiento
df["Call Start Time"] = pd.to_datetime(df["Call Start Time"], errors="coerce")
df["Talk Time"] = pd.to_timedelta(df["Talk Time"], errors="coerce")
df["Fecha"] = df["Call Start Time"].dt.date
df["Hora"] = df["Call Start Time"].dt.hour
df["DíaSemana_En"] = df["Call Start Time"].dt.day_name()

dias_traducidos = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
}
df["DíaSemana"] = df["DíaSemana_En"].map(dias_traducidos)

# Llamadas perdidas
df["LlamadaPerdida"] = df["Talk Time"] == pd.Timedelta("0:00:00")

def agentes_por_horario(hora):
    if 8 <= hora < 10:
        return ["Jorge Cesar Flores Rivera"]
    elif 10 <= hora < 12:
        return ["Jorge Cesar Flores Rivera", "Maria Teresa Loredo Morales"]
    elif 12 <= hora < 16:
        return ["Jorge Cesar Flores Rivera", "Maria Teresa Loredo Morales", "Jonathan Alejandro Zúñiga"]
    elif 16 <= hora < 18:
        return ["Jonathan Alejandro Zúñiga", "Maria Teresa Loredo Morales"]
    elif 18 <= hora < 20:
        return ["Jonathan Alejandro Zúñiga"]
    else:
        return []

filas = []
for _, row in df.iterrows():
    if row["LlamadaPerdida"]:
        agentes = agentes_por_horario(row["Hora"])
        if agentes:
            for agente in agentes:
                filas.append({**row, "AgenteFinal": agente})
        else:
            if pd.notna(row["Agent Name"]):
                filas.append({**row, "AgenteFinal": row["Agent Name"]})
    else:
        if pd.notna(row["Agent Name"]):
            filas.append({**row, "AgenteFinal": row["Agent Name"]})

df_expandido = pd.DataFrame(filas)
df_expandido = df_expandido[df_expandido["AgenteFinal"].notna()]

# Título principal
st.title("Análisis Integral de Productividad y Llamadas")

# Filtro por fechas
fechas_disponibles = df["Fecha"].sort_values().unique()
fecha_inicio_default = fechas_disponibles[0]
fecha_fin_default = fechas_disponibles[-1]

rango_fechas = st.date_input(
    "Selecciona el rango de fechas:",
    value=(fecha_inicio_default, fecha_fin_default),
    min_value=fecha_inicio_default,
    max_value=fecha_fin_default
)

# Manejo robusto del input para rango de fechas
if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
    fecha_inicio, fecha_fin = rango_fechas
elif isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 1:
    fecha_inicio = fecha_fin = rango_fechas[0]
elif hasattr(rango_fechas, "date"):  # pd.Timestamp u objeto similar
    fecha_inicio = fecha_fin = rango_fechas.date()
elif isinstance(rango_fechas, str):
    fecha_inicio = fecha_fin = pd.to_datetime(rango_fechas).date()
else:
    fecha_inicio, fecha_fin = fecha_inicio_default, fecha_fin_default

df_filtrado = df[(df["Fecha"] >= fecha_inicio) & (df["Fecha"] <= fecha_fin)].copy()
df_expandido_filtrado = df_expandido[(df_expandido["Fecha"] >= fecha_inicio) & (df_expandido["Fecha"] <= fecha_fin)].copy()

# Tablas y agrupaciones
detalle = df_expandido_filtrado.groupby(["AgenteFinal", "Fecha"]).agg(
    LlamadasTotales=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum")
).reset_index()
detalle["LlamadasAtendidas"] = detalle["LlamadasTotales"] - detalle["LlamadasPerdidas"]
detalle["Productividad (%)"] = (detalle["LlamadasAtendidas"] / detalle["LlamadasTotales"] * 100).round(2)

resumen_diario_todos = df_expandido_filtrado.groupby("Fecha").agg(
    LlamadasTotales=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum")
).reset_index()
resumen_diario_todos["LlamadasAtendidas"] = resumen_diario_todos["LlamadasTotales"] - resumen_diario_todos["LlamadasPerdidas"]
resumen_diario_todos["Productividad (%)"] = (resumen_diario_todos["LlamadasAtendidas"] / resumen_diario_todos["LlamadasTotales"] * 100).round(2)

df_productividad = df_filtrado.groupby(df_filtrado["Call Start Time"].dt.date).agg(
    LlamadasRecibidas=("Talk Time", "count"),
    LlamadasPerdidas=("Talk Time", lambda x: (x == pd.Timedelta("0:00:00")).sum())
).reset_index().rename(columns={"Call Start Time": "Fecha"})

df_productividad["Productividad (%)"] = ((df_productividad["LlamadasRecibidas"] - df_productividad["LlamadasPerdidas"]) / df_productividad["LlamadasRecibidas"] * 100).round(2)
df_productividad["Tasa de Abandono (%)"] = 100 - df_productividad["Productividad (%)"]
df_productividad["DíaSemana"] = pd.to_datetime(df_productividad["Fecha"]).dt.day_name().map(dias_traducidos)

# Heatmap de llamadas recibidas
pivot_llamadas = df_filtrado.pivot_table(index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0)
pivot_llamadas = pivot_llamadas.reindex(columns=dias_traducidos.keys(), fill_value=0)
pivot_llamadas.columns = [dias_traducidos[d] for d in pivot_llamadas.columns]
pivot_llamadas.index = [f"{h}:00" for h in pivot_llamadas.index]

# Heatmap de llamadas perdidas
pivot_perdidas = df_filtrado[df_filtrado["LlamadaPerdida"]].pivot_table(
    index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0
)
pivot_perdidas = pivot_perdidas.reindex(columns=dias_traducidos.keys(), fill_value=0)
pivot_perdidas.columns = [dias_traducidos[d] for d in pivot_perdidas.columns]
pivot_perdidas.index = [f"{h}:00" for h in pivot_perdidas.index]

# Talk time promedio por agente
talktime_por_agente = df_filtrado[df_filtrado["Talk Time"] > pd.Timedelta(0)].groupby("Agent Name")["Talk Time"].agg(["mean", "count"])
talktime_por_agente["mean_minutes"] = talktime_por_agente["mean"].dt.total_seconds() / 60

# Alerta de picos (en la barra lateral)
st.sidebar.header("🚨 Alertas de Picos")
llamadas_dia = df_filtrado.groupby("Fecha").agg(
    Total=("Talk Time", "count"),
    Perdidas=("LlamadaPerdida", "sum")
).reset_index()

media_recibidas = llamadas_dia["Total"].mean()
std_recibidas = llamadas_dia["Total"].std()

media_perdidas = llamadas_dia["Perdidas"].mean()
std_perdidas = llamadas_dia["Perdidas"].std()

alertas_recibidas = llamadas_dia[llamadas_dia["Total"] > media_recibidas + 2 * std_recibidas]
alertas_perdidas = llamadas_dia[llamadas_dia["Perdidas"] > media_perdidas + 2 * std_perdidas]

if not alertas_recibidas.empty or not alertas_perdidas.empty:
    st.sidebar.write("📅 Fechas con alertas:")
    for fecha in pd.concat([alertas_recibidas["Fecha"], alertas_perdidas["Fecha"]]).unique():
        st.sidebar.write(f"🔺 {fecha}")
else:
    st.sidebar.success("Sin alertas detectadas.")

# Pestañas principales
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Detalle por Programador",
    "Resumen Diario Total",
    "Heatmap Llamadas",
    "Productividad General",
    "Heatmap Llamadas Perdidas",
    "Distribución Talk Time",
    "Alerta de Picos"
])

with tab1:
    st.header("📋 Detalle Diario por Programador")
    agente_seleccionado = st.selectbox("Selecciona Programador", options=detalle["AgenteFinal"].unique())
    df_agente = detalle[detalle["AgenteFinal"] == agente_seleccionado].sort_values("Fecha")
    st.dataframe(df_agente.style.format({"Productividad (%)": "{:.2f}"}))

with tab2:
    st.header("📊 Resumen Diario Total")
    st.dataframe(resumen_diario_todos.style.format({"Productividad (%)": "{:.2f}"}))

with tab3:
    st.header("📈 Distribución de llamadas por hora y día")
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_llamadas, cmap="YlGnBu", annot=True, fmt="d", ax=ax)
    ax.set_xlabel("Día de la Semana")
    ax.set_ylabel("Hora del Día")
    st.pyplot(fig)

with tab4:
    st.header("📈 Productividad y Tasa de Abandono Diaria")
    st.dataframe(df_productividad[["Fecha", "LlamadasRecibidas", "LlamadasPerdidas", "Productividad (%)", "Tasa de Abandono (%)", "DíaSemana"]])

with tab5:
    st.header("📉 Llamadas Perdidas por Hora y Día")
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_perdidas, cmap="OrRd", annot=True, fmt="d", ax=ax)
    ax.set_xlabel("Día de la Semana")
    ax.set_ylabel("Hora del Día")
    st.pyplot(fig)

with tab6:
    st.header("⏱️ Duración promedio de llamadas (minutos) por agente")
    st.dataframe(talktime_por_agente[["mean_minutes", "count"]].rename(columns={"mean_minutes": "Duración Promedio (min)","count": "Número de llamadas"}))

with tab7:
    st.header("🚨 Alertas de Picos")
    if not alertas_recibidas.empty:
        st.write("🔺 Picos en llamadas recibidas:")
        st.dataframe(alertas_recibidas)
    if not alertas_perdidas.empty:
        st.write("🔺 Picos en llamadas perdidas:")
        st.dataframe(alertas_perdidas)
    if alertas_recibidas.empty and alertas_perdidas.empty:
        st.write("No hay alertas de picos en el rango seleccionado.")
