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
df["Ring Time"] = pd.to_timedelta(df["Ring Time"], errors="coerce")
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

# Invertir índice de horas
horas_ordenadas = list(range(8, 21))
pivot_table = pivot_table.reindex(horas_ordenadas[::-1], fill_value=0)
pivot_table.index = [f"{h}:00" for h in pivot_table.index]

# Heatmap llamadas perdidas
pivot_perdidas = df_expandido_filtrado[
    (df_expandido_filtrado["DíaSemana_En"].isin(dias_validos)) & (df_expandido_filtrado["LlamadaPerdida"])
]
pivot_table_perdidas = pivot_perdidas.pivot_table(
    index="Hora",
    columns="DíaSemana_En",
    aggfunc="size",
    fill_value=0
)
pivot_table_perdidas = pivot_table_perdidas.reindex(columns=dias_validos, fill_value=0)
pivot_table_perdidas.columns = [dias_traducidos[d] for d in pivot_table_perdidas.columns]
pivot_table_perdidas = pivot_table_perdidas.reindex(horas_ordenadas[::-1], fill_value=0)
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

    def color_fila_tab1(row):
        valor = row["Productividad (%)"]
        if valor >= 97:
            color = "background-color: #28a745; color: white;"
        elif 90 <= valor < 97:
            color = "background-color: #fff3cd; color: black;"
        else:
            color = "background-color: #dc3545; color: white;"
        return [color] * len(row)

    styled_df = df_productividad[["Fecha", "LlamadasRecibidas", "LlamadasPerdidas", "Productividad (%)", "Tasa de Abandono (%)", "DíaSemana"]].style.apply(color_fila_tab1, axis=1).format({"Productividad (%)": "{:.2f}", "Tasa de Abandono (%)": "{:.2f}"})
    st.dataframe(styled_df)

    dias_cumplen = df_productividad[df_productividad["Productividad (%)"] >= 97].shape[0]
    total_dias = df_productividad.shape[0]
    porcentaje_cumplen = (dias_cumplen / total_dias * 100) if total_dias > 0 else 0

    st.markdown(f"### Cumplimiento de Meta")
    st.markdown(f"- **Días que cumplen meta (≥97% Productividad):** {dias_cumplen} días")
    st.markdown(f"- **Porcentaje de cumplimiento:** {porcentaje_cumplen:.2f} %")

with tab2:
    st.header("Detalle Diario por Programador")
    agente_seleccionado = st.selectbox("Selecciona Programador:", detalle["AgenteFinal"].unique())
    detalle_agente = detalle[detalle["AgenteFinal"] == agente_seleccionado].sort_values("Fecha")

    def color_fila_detalle(row):
        valor = row["Productividad (%)"]
        if valor >= 97:
            color = "background-color: #28a745; color: white;"
        elif 90 <= valor < 97:
            color = "background-color: #fff3cd; color: black;"
        else:
            color = "background-color: #dc3545; color: white;"
        return [color] * len(row)

    detalle_tabla = detalle_agente[["Fecha", "LlamadasTotales", "LlamadasPerdidas", "LlamadasAtendidas", "Productividad (%)", "Promedio Talk Time (seg)"]]
    styled_detalle = detalle_tabla.style.apply(color_fila_detalle, axis=1).format({"Productividad (%)": "{:.2f}", "Promedio Talk Time (seg)": "{:.2f}"})
    st.dataframe(styled_detalle)

    dias_cumplen_prog = detalle_agente[detalle_agente["Productividad (%)"] >= 97].shape[0]
    total_dias_prog = detalle_agente.shape[0]
    porcentaje_cumplen_prog = (dias_cumplen_prog / total_dias_prog * 100) if total_dias_prog > 0 else 0

    st.markdown(f"### Cumplimiento de Meta para {agente_seleccionado}")
    st.markdown(f"- **Días que cumplen meta (≥97% Productividad):** {dias_cumplen_prog} días")
    st.markdown(f"- **Porcentaje de cumplimiento:** {porcentaje_cumplen_prog:.2f} %")

    st.markdown("---")
    st.subheader("📊 Promedio de Productividad por Agente (Fechas Filtradas)")
    promedio_productividad = (
        detalle.groupby("AgenteFinal")["Productividad (%)"]
        .mean()
        .round(2)
        .sort_values(ascending=False)
    )
    st.dataframe(
        promedio_productividad.reset_index().rename(columns={"Productividad (%)": "Promedio Productividad (%)"}),
        use_container_width=True
    )

with tab3:
    st.header("Heatmap de Llamadas Entrantes")
    plt.figure(figsize=(10, 6))
    sns.heatmap(pivot_table, annot=True, fmt="d", cmap="YlGnBu")
    st.pyplot(plt.gcf())

with tab4:
    st.header("Heatmap de Llamadas Perdidas")
    plt.figure(figsize=(10, 6))
    sns.heatmap(pivot_table_perdidas, annot=True, fmt="d", cmap="OrRd")
    st.pyplot(plt.gcf())

with tab5:
    st.header("Análisis detallado por Agente")

    agente_seleccionado = st.selectbox("Selecciona Agente", options=df_expandido_filtrado["AgenteFinal"].unique(), key="tab5_agente")

    df_agente_talktime = df_expandido_filtrado[
        (df_expandido_filtrado["AgenteFinal"] == agente_seleccionado) & (~df_expandido_filtrado["LlamadaPerdida"])
    ]

    if not df_agente_talktime.empty and df_agente_talktime["Talk Time"].notna().any():
        talktime_seconds = df_agente_talktime["Talk Time"].dt.total_seconds()

        fig_hist_talk, ax_hist_talk = plt.subplots()
        ax_hist_talk.hist(talktime_seconds, bins=30, color='skyblue', edgecolor='black')
        ax_hist_talk.set_title(f"Distribución de Talk Time (segundos) - {agente_seleccionado}")
        ax_hist_talk.set_xlabel("Segundos")
        ax_hist_talk.set_ylabel("Frecuencia")
        st.pyplot(fig_hist_talk)

        promedio_talktime = talktime_seconds.mean()
        st.write(f"Promedio de Talk Time para {agente_seleccionado}: **{promedio_talktime:.2f} segundos**")
    else:
        st.write(f"No hay datos de Talk Time para {agente_seleccionado} en el período seleccionado.")
