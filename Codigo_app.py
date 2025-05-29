import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Cargar archivo
archivo_excel = "AppInfo.xlsx"
df = pd.read_excel(archivo_excel, engine="openpyxl")

# Normalizar nombres
mapeo_a_nombre_completo = {
    "Jorge": "Jorge Cesar Flores Rivera",
    "Maria": "Maria Teresa Loredo Morales",
    "Jonathan": "Jonathan Alejandro Zúñiga",
}
df["Agent Name"] = df["Agent Name"].replace(mapeo_a_nombre_completo)

# Procesamiento inicial
df["Call Start Time"] = pd.to_datetime(df["Call Start Time"], errors="coerce")

# Eliminar filas con fechas inválidas
df = df.dropna(subset=["Call Start Time"])

df["Talk Time"] = pd.to_timedelta(df["Talk Time"], errors="coerce")
df["Fecha"] = df["Call Start Time"].dt.date
df["Hora"] = df["Call Start Time"].dt.hour
df["DíaSemana_En"] = df["Call Start Time"].dt.day_name()

# Traducción de días
dias_traducidos = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
}
df["DíaSemana"] = df["DíaSemana_En"].map(dias_traducidos)

# Identificar llamadas perdidas por Talk Time
df["LlamadaPerdida"] = df["Talk Time"] == pd.Timedelta("0:00:00")

# Función para asignar agente según horario
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

# Expandir filas para llamadas perdidas
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

st.title("Análisis Integral de Productividad y Llamadas")

# Filtro por rango de fechas
fecha_min = df["Fecha"].min()
fecha_max = df["Fecha"].max()
fecha_inicio, fecha_fin = st.date_input("Selecciona un rango de fechas:", [fecha_min, fecha_max])

# Filtrar dataframes por fechas
df_filtrado = df[(df["Fecha"] >= fecha_inicio) & (df["Fecha"] <= fecha_fin)]
df_expandido_filtrado = df_expandido[(df_expandido["Fecha"] >= fecha_inicio) & (df_expandido["Fecha"] <= fecha_fin)]

# Productividad general diaria
df_productividad = df_filtrado.groupby("Fecha").agg(
    LlamadasRecibidas=("Talk Time", "count"),
    LlamadasPerdidas=("Talk Time", lambda x: (x == pd.Timedelta("0:00:00")).sum())
).reset_index()

df_productividad["Productividad (%)"] = (
    (df_productividad["LlamadasRecibidas"] - df_productividad["LlamadasPerdidas"]) / df_productividad["LlamadasRecibidas"] * 100
).round(2)
df_productividad["Tasa de Abandono (%)"] = 100 - df_productividad["Productividad (%)"]
df_productividad["DíaSemana"] = pd.to_datetime(df_productividad["Fecha"]).dt.day_name().map(dias_traducidos)

# Detalle diario por programador
detalle = df_expandido_filtrado.groupby(["AgenteFinal", "Fecha"]).agg(
    LlamadasTotales=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum"),
    TalkTimeTotal=("Talk Time", "sum")
).reset_index()
detalle["LlamadasAtendidas"] = detalle["LlamadasTotales"] - detalle["LlamadasPerdidas"]
detalle["Productividad (%)"] = (detalle["LlamadasAtendidas"] / detalle["LlamadasTotales"] * 100).round(2)
detalle["Promedio Talk Time (seg)"] = (detalle["TalkTimeTotal"].dt.total_seconds() / detalle["LlamadasAtendidas"]).round(2)

# Días válidos sin domingo y en orden
dias_validos = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
dias_validos_es = [dias_traducidos[d] for d in dias_validos]

# Heatmap llamadas entrantes
pivot_table = df_filtrado[df_filtrado["DíaSemana_En"].isin(dias_validos)].pivot_table(
    index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0
)
pivot_table = pivot_table.reindex(columns=dias_validos, fill_value=0)
pivot_table.columns = [dias_traducidos[d] for d in pivot_table.columns]
pivot_table = pivot_table.sort_index(ascending=True)

# Invertir índice de horas para empezar desde 8am hacia abajo
horas_ordenadas = list(range(8, 21))  # 8am a 20pm
pivot_table = pivot_table.reindex(horas_ordenadas[::-1], fill_value=0)
pivot_table.index = [f"{h}:00" for h in pivot_table.index]

# Preparación del pivot table para heatmap llamadas perdidas (reordenado y con índice legible)
dias_validos = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
dias_traducidos = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
}
horas_ordenadas = list(range(8, 21))  # Desde 8am hasta 20pm

pivot_perdidas = df_expandido_filtrado[
    (df_expandido_filtrado["DíaSemana_En"].isin(dias_validos)) & (df_expandido_filtrado["LlamadaPerdida"])
]

pivot_table_perdidas = pivot_perdidas.pivot_table(
    index="Hora",
    columns="DíaSemana_En",
    aggfunc="size",
    fill_value=0
)

# Reordenar columnas y filas
pivot_table_perdidas = pivot_table_perdidas.reindex(columns=dias_validos, fill_value=0)
pivot_table_perdidas.columns = [dias_traducidos[d] for d in pivot_table_perdidas.columns]
pivot_table_perdidas = pivot_table_perdidas.reindex(horas_ordenadas[::-1], fill_value=0)  # invertir orden para mostrar de 8am hacia abajo
pivot_table_perdidas.index = [f"{h}:00" for h in pivot_table_perdidas.index]
# Tabs para navegación
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Productividad General",
    "Detalle por Programador",
    "Heatmap Llamadas",
    "Heatmap Llamadas Perdidas",
    "Distribución & Alertas"
])

with tab1:
    st.header("Productividad y Tasa de Abandono Diaria")
    st.dataframe(df_productividad[["Fecha", "LlamadasRecibidas", "LlamadasPerdidas", "Productividad (%)", "Tasa de Abandono (%)", "DíaSemana"]])

with tab2:
    st.header("Detalle Diario por Programador")
    agente_seleccionado_detalle = st.selectbox("Selecciona Programador", options=detalle["AgenteFinal"].unique(), key="detalle_agente")
    df_agente = detalle[detalle["AgenteFinal"] == agente_seleccionado_detalle].sort_values("Fecha")
    st.dataframe(df_agente.style.format({"Productividad (%)": "{:.2f}", "Promedio Talk Time (seg)": "{:.2f}"}))

with tab3:
    st.header("Distribución de llamadas por hora y día")
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_table, cmap="YlGnBu", annot=True, fmt="d", ax=ax)
    ax.set_xlabel("Día de la Semana")
    ax.set_ylabel("Hora del Día")
    ax.invert_yaxis()
    st.pyplot(fig)

with tab4:
    st.header("Distribución de llamadas perdidas por hora y día")
    fig2, ax2 = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_table_perdidas, cmap="OrRd", annot=True, fmt="d", ax=ax2)
    ax2.set_xlabel("Día de la Semana")
    ax2.set_ylabel("Hora del Día")
    ax2.invert_yaxis()
    st.pyplot(fig2)

with tab5:
    st.header("Distribución y Promedio del Tiempo de Conversación por Agente")
    agente_seleccionado = st.selectbox("Selecciona un agente para ver distribución de Talk Time:", options=df_expandido_filtrado["AgenteFinal"].unique())
    df_agente_talktime = df_expandido_filtrado[(df_expandido_filtrado["AgenteFinal"] == agente_seleccionado) & (~df_expandido_filtrado["LlamadaPerdida"])]

    if not df_agente_talktime.empty:
        talktime_seconds = df_agente_talktime["Talk Time"].dt.total_seconds()
        fig_hist, ax_hist = plt.subplots()
        ax_hist.hist(talktime_seconds, bins=30, color='skyblue', edgecolor='black')
        ax_hist.set_title(f"Distribución de Talk Time (segundos) - {agente_seleccionado}")
        ax_hist.set_xlabel("Segundos")
        ax_hist.set_ylabel("Frecuencia")
        st.pyplot(fig_hist)

        promedio_talktime = talktime_seconds.mean()
        st.write(f"Promedio de Talk Time para {agente_seleccionado}: **{promedio_talktime:.2f} segundos**")
    else:
        st.write(f"No hay llamadas atendidas para {agente_seleccionado} en el rango seleccionado.")

    st.header("Alertas de picos o pérdidas de llamadas")

    resumen_agente_hora = df_expandido_filtrado.groupby(["AgenteFinal", "Hora"]).agg(
        TotalLlamadas=("Talk Time", "count"),
        LlamadasPerdidas=("LlamadaPerdida", "sum")
    ).reset_index()

    alertas = []

    for agente in resumen_agente_hora["AgenteFinal"].unique():
        data_agente = resumen_agente_hora[resumen_agente_hora["AgenteFinal"] == agente]

        mean_total = data_agente["TotalLlamadas"].mean()
        std_total = data_agente["TotalLlamadas"].std()
        umbral_total = mean_total + 2 * std_total

        for _, row in data_agente.iterrows():
            if row["TotalLlamadas"] > umbral_total:
                alertas.append(f"Alerta: {agente} tuvo un pico de llamadas ({row['TotalLlamadas']}) a la hora {row['Hora']}:00")

        # Verificar si hubo hora sin llamadas entre horas con llamadas
        horas_con_llamadas = sorted(data_agente[data_agente["TotalLlamadas"] > 0]["Hora"].unique())
        for i in range(len(horas_con_llamadas) - 1):
            if horas_con_llamadas[i+1] - horas_con_llamadas[i] > 1:
                faltante = horas_con_llamadas[i] + 1
                alertas.append(f"Alerta: {agente} no tuvo llamadas a la hora {faltante}:00 entre horas con actividad.")

    if alertas:
        for alerta in alertas:
            st.warning(alerta)
    else:
        st.success("No se detectaron alertas de picos o pérdidas de llamadas.")

