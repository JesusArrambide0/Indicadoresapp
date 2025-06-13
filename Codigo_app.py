import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Cargar archivo
archivo_excel = "AppInfo.xlsx"
df = pd.read_excel(archivo_excel, engine="openpyxl")

# Normalizar nombres (con reemplazo de Maria Teresa por Gabriela Lizeth)
mapeo_a_nombre_completo = {
    "Jorge": "Jorge Cesar Flores Rivera",
    "Maria": "Gabriela Lizeth Hernandez",
    "Jonathan": "Jonathan Alejandro Zúñiga",
}
df["Agent Name"] = df["Agent Name"].replace(mapeo_a_nombre_completo)

# Procesamiento inicial
df["Call Start Time"] = pd.to_datetime(df["Call Start Time"], format="%d/%m/%y %H:%M:%S", errors="coerce")
df = df.dropna(subset=["Call Start Time"])
df["Talk Time"] = pd.to_timedelta(df["Talk Time"], errors="coerce")
df["Ring Time"] = pd.to_timedelta(df["Ring Time"], errors="coerce")
df["Fecha"] = df["Call Start Time"].dt.date
df["Hora"] = df["Call Start Time"].dt.hour
df["DíaSemana_En"] = df["Call Start Time"].dt.day_name()

dias_traducidos = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
}
df["DíaSemana"] = df["DíaSemana_En"].map(dias_traducidos)

# Definir llamada perdida: Talk Time 0 y Agent Name vacío o NaN
df["LlamadaPerdida"] = (
    (df["Talk Time"] == pd.Timedelta(0)) &
    (df["Agent Name"].isna() | (df["Agent Name"].str.strip() == ""))
)

def agentes_por_horario(hora):
    if 8 <= hora < 10:
        return ["Jorge Cesar Flores Rivera"]
    elif 10 <= hora < 12:
        return ["Jorge Cesar Flores Rivera", "Gabriela Lizeth Hernandez"]
    elif 12 <= hora < 16:
        return ["Jorge Cesar Flores Rivera", "Gabriela Lizeth Hernandez", "Jonathan Alejandro Zúñiga"]
    elif 16 <= hora < 18:
        return ["Jonathan Alejandro Zúñiga", "Gabriela Lizeth Hernandez"]
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

st.title("Análisis Integral de Productividad y Llamadas")

# Rango de fechas por defecto
fecha_min = df["Fecha"].min()
fecha_max = df["Fecha"].max()

fecha_seleccion = st.date_input("Selecciona un rango de fechas:", [fecha_min, fecha_max])
if isinstance(fecha_seleccion, tuple) or isinstance(fecha_seleccion, list):
    fecha_inicio, fecha_fin = fecha_seleccion
else:
    fecha_inicio = fecha_fin = fecha_seleccion

df_filtrado = df[(df["Fecha"] >= fecha_inicio) & (df["Fecha"] <= fecha_fin)]
df_expandido_filtrado = df_expandido[(df_expandido["Fecha"] >= fecha_inicio) & (df_expandido["Fecha"] <= fecha_fin)]

# Productividad General con Llamadas Perdidas basadas en la condición exacta
df_productividad = df_filtrado.groupby("Fecha").agg(
    LlamadasRecibidas=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum")
).reset_index()

df_productividad["Productividad (%)"] = (
    (df_productividad["LlamadasRecibidas"] - df_productividad["LlamadasPerdidas"]) /
    df_productividad["LlamadasRecibidas"] * 100
).round(2)
df_productividad["Tasa de Abandono (%)"] = 100 - df_productividad["Productividad (%)"]
df_productividad["DíaSemana"] = pd.to_datetime(df_productividad["Fecha"]).dt.day_name().map(dias_traducidos)

# Detalle por Programador con misma lógica de llamadas perdidas
detalle = df_expandido_filtrado.groupby(["AgenteFinal", "Fecha"]).agg(
    LlamadasTotales=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum"),
    TalkTimeTotal=("Talk Time", "sum"),
    RingTimeTotal=("Ring Time", "sum")
).reset_index()

detalle["LlamadasAtendidas"] = detalle["LlamadasTotales"] - detalle["LlamadasPerdidas"]
detalle["Productividad (%)"] = (detalle["LlamadasAtendidas"] / detalle["LlamadasTotales"] * 100).round(2)
detalle["Promedio Talk Time (seg)"] = (detalle["TalkTimeTotal"].dt.total_seconds() / detalle["LlamadasAtendidas"]).round(2)
detalle["Promedio Ring Time (seg)"] = (detalle["RingTimeTotal"].dt.total_seconds() / detalle["LlamadasTotales"]).round(2)

dias_validos = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
dias_validos_es = [dias_traducidos[d] for d in dias_validos]

pivot_table = df_filtrado[df_filtrado["DíaSemana_En"].isin(dias_validos)].pivot_table(
    index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0
)
pivot_table = pivot_table.reindex(columns=dias_validos, fill_value=0)
pivot_table.columns = [dias_traducidos[d] for d in pivot_table.columns]
pivot_table = pivot_table.sort_index(ascending=True)

horas_ordenadas = list(range(8, 21))
pivot_table = pivot_table.reindex(horas_ordenadas[::-1], fill_value=0)  # Invierte el orden aquí
pivot_table.index = [f"{h}:00" for h in pivot_table.index]

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
pivot_table_perdidas = pivot_table_perdidas.reindex(horas_ordenadas[::-1], fill_value=0)  # Invierte el orden aquí también
pivot_table_perdidas.index = [f"{h}:00" for h in pivot_table_perdidas.index]

# Tabs
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

    st.markdown(f"**{dias_cumplen}** días cumplen con productividad ≥ 97% de un total de **{total_dias}** días ({porcentaje_cumplen:.2f}%)")
    
     # Gráfico líneas productividad
    fig, ax = plt.subplots(figsize=(10, 4))
    sns.lineplot(data=df_productividad, x="Fecha", y="Productividad (%)", marker="o", ax=ax)
    ax.axhline(97, color="green", linestyle="--", label="Meta 97%")
    ax.axhline(90, color="orange", linestyle="--", label="Alerta 90%")
    ax.set_ylim(0, 105)
    ax.set_ylabel("Productividad (%)")
    ax.set_title("Productividad diaria")
    ax.legend()
    ax.grid(True, linestyle="--", alpha=0.6)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

with tab2:
    st.subheader("Detalle de llamadas por Programador")
    agentes_unicos_tab2 = sorted(detalle["AgenteFinal"].unique())
    agente_tab2_seleccionado = st.selectbox("Selecciona un programador para análisis", agentes_unicos_tab2)

    detalle_filtrado = detalle[detalle["AgenteFinal"] == agente_tab2_seleccionado]

    # Tabla resumen filtrada y coloreada
    def color_fila_tab2(row):
        valor = row["Productividad (%)"]
        if valor >= 97:
            color = "background-color: #28a745; color: white;"
        elif 90 <= valor < 97:
            color = "background-color: #fff3cd; color: black;"
        else:
            color = "background-color: #dc3545; color: white;"
        return [color] * len(row)

    tabla_detalle = detalle_filtrado[[
        "AgenteFinal", "Fecha", "LlamadasTotales", "LlamadasPerdidas",
        "Productividad (%)", "Promedio Talk Time (seg)", "Promedio Ring Time (seg)"
    ]].sort_values(by=["AgenteFinal", "Fecha"])

    styled_detalle = tabla_detalle.style.apply(color_fila_tab2, axis=1).format({
        "Productividad (%)": "{:.2f}",
        "Promedio Talk Time (seg)": "{:.1f}",
        "Promedio Ring Time (seg)": "{:.1f}"
    })
    
    st.dataframe(styled_detalle)

with tab3:
    st.subheader("Heatmap de llamadas por hora y día")
    fig, ax_heat = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_table, annot=True, fmt="d", cmap="YlGnBu", cbar=True, ax=ax_heat)
    ax_heat.set_title("Cantidad de llamadas atendidas por hora y día (Lunes a Sábado)")
    ax_heat.set_xlabel("Día de la semana")
    ax_heat.set_ylabel("Hora del día")
    st.pyplot(fig)

with tab4:
    st.subheader("Heatmap de llamadas perdidas por hora y día")
    fig2, ax_heat_perdidas = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_table_perdidas, annot=True, fmt="d", cmap="YlOrRd", cbar=True, ax=ax_heat_perdidas)
    ax_heat_perdidas.set_title("Cantidad de llamadas perdidas por hora y día (Lunes a Sábado)")
    ax_heat_perdidas.set_xlabel("Día de la semana")
    ax_heat_perdidas.set_ylabel("Hora del día")
    st.pyplot(fig2)

with tab5:
    st.header("Distribución y Alertas")

    agentes_unicos_tab5 = sorted(df_expandido_filtrado["AgenteFinal"].unique())
    agente_tab5_seleccionado = st.selectbox("Selecciona un agente para análisis", agentes_unicos_tab5)

    df_agente = df_expandido_filtrado[df_expandido_filtrado["AgenteFinal"] == agente_tab5_seleccionado]
    df_agente_talktime = df_agente[~df_agente["LlamadaPerdida"]].copy()
    df_agente_talktime["TalkTime_seg"] = df_agente_talktime["Talk Time"].dt.total_seconds()
    df_agente["RingTime_seg"] = df_agente["Ring Time"].dt.total_seconds()

    # Histograma Talk Time
    fig_dist, ax_dist = plt.subplots(figsize=(10, 4))
    sns.histplot(df_agente_talktime["TalkTime_seg"], bins=30, kde=True, ax=ax_dist)
    ax_dist.set_xlabel("Duración de llamada (segundos)")
    ax_dist.set_title(f"Distribución Duración de Llamadas Atendidas - {agente_tab5_seleccionado}")
    st.pyplot(fig_dist)

    promedio_talk = df_agente_talktime["TalkTime_seg"].mean()
    st.markdown(f"**Promedio Talk Time:** {promedio_talk:.2f} segundos")

    # Histograma Ring Time
    fig_ring, ax_ring = plt.subplots(figsize=(10, 4))
    sns.histplot(df_agente["RingTime_seg"], bins=30, kde=True, ax=ax_ring, color="orange")
    ax_ring.set_xlabel("Ring Time (segundos)")
    ax_ring.set_title(f"Distribución de Ring Time - {agente_tab5_seleccionado}")
    st.pyplot(fig_ring)

    promedio_ring = df_agente["RingTime_seg"].mean()
    st.markdown(f"**Promedio Ring Time:** {promedio_ring:.2f} segundos")

    # Alertas simple (ejemplo: días con productividad menor a 90% para agente seleccionado)
    detalle_agente = detalle[(detalle["AgenteFinal"] == agente_tab5_seleccionado) & (detalle["Fecha"] >= fecha_inicio) & (detalle["Fecha"] <= fecha_fin)]
    dias_alerta = detalle_agente[detalle_agente["Productividad (%)"] < 90]

    if not dias_alerta.empty:
        st.warning(f"Días con productividad < 90% para {agente_tab5_seleccionado}:")
        st.dataframe(dias_alerta[["Fecha", "Productividad (%)", "LlamadasTotales", "LlamadasPerdidas"]])
    else:
        st.success(f"No se detectaron días con productividad menor a 90% para {agente_tab5_seleccionado}.")
