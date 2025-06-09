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
# Interpretar "Call Start Time" en formato dd/mm/aa hh:mm:ss
df["Call Start Time"] = pd.to_datetime(df["Call Start Time"], format="%d/%m/%y %H:%M:%S", errors="coerce")

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

# Función para asignar agente según horario (con Maria Teresa cambiado a Gabriela)
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

# Heatmap llamadas perdidas - SOLO LLAMADAS ÚNICAS
pivot_perdidas_df = df_expandido_filtrado[
    (df_expandido_filtrado["DíaSemana_En"].isin(dias_validos)) & (df_expandido_filtrado["LlamadaPerdida"])
]

# Aquí eliminamos duplicados para contar solo llamadas únicas:
# Asumiendo que 'Call Start Time' y 'AgenteFinal' definen unicidad
pivot_perdidas_df_unique = pivot_perdidas_df.drop_duplicates(subset=["Call Start Time", "AgenteFinal"])

pivot_table_perdidas = pivot_perdidas_df_unique.pivot_table(
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

    styled_df = df_productividad[["Fecha", "LlamadasRecibidas", "LlamadasPerdidas", "Productividad (%)", "Tasa de Abandono (%)", "DíaSemana"]].style.apply(color_fila_tab1, axis=1)
    st.dataframe(styled_df, use_container_width=True)

with tab2:
    st.header("Detalle Diario por Programador")
    st.dataframe(detalle, use_container_width=True)

with tab3:
    st.header("Heatmap de Llamadas Entrantes")
    fig, ax = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_table, cmap="viridis", annot=True, fmt="d", ax=ax)
    ax.invert_yaxis()
    ax.set_title("Llamadas entrantes por hora y día")
    st.pyplot(fig)

with tab4:
    st.header("Heatmap de Llamadas Perdidas (solo llamadas únicas)")
    fig2, ax2 = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_table_perdidas, cmap="rocket_r", annot=True, fmt="d", ax=ax2)
    ax2.invert_yaxis()
    ax2.set_title("Llamadas perdidas únicas por hora y día")
    st.pyplot(fig2)

with tab5:
    st.header("Distribución y Alertas")
    st.write("Implementa tus gráficos o análisis adicionales aquí.")
    st.subheader("Distribución de 'Talk Time' por agente")
    agentes_dist = st.multiselect("Selecciona agentes para distribución:", options=df_filtrado["Agent Name"].unique(), default=df_filtrado["Agent Name"].unique())

    df_dist_filtrado = df_filtrado[df_filtrado["Agent Name"].isin(agentes_dist) & (df_filtrado["Talk Time"] > pd.Timedelta(0))]

    if not df_dist_filtrado.empty:
        # Mostrar tiempo promedio Talk Time por agente
        promedio_talk = df_dist_filtrado.groupby("Agent Name")["Talk Time"].mean()
        promedio_talk_min = promedio_talk.dt.total_seconds() / 60
        st.markdown("### Tiempo promedio de 'Talk Time' por agente (minutos)")
        st.dataframe(promedio_talk_min.round(2).to_frame())

        fig5, ax5 = plt.subplots(figsize=(10, 5))
        for agente in agentes_dist:
            sns.histplot(df_dist_filtrado[df_dist_filtrado["Agent Name"] == agente]["Talk Time"].dt.total_seconds() / 60, bins=30, kde=True, label=agente, ax=ax5)
        ax5.set_xlabel("Duración de llamada (minutos)")
        ax5.set_title("Distribución de duración de llamadas")
        ax5.legend()
        st.pyplot(fig5)
    else:
        st.write("No hay datos de 'Talk Time' para los agentes seleccionados en el rango de fechas.")

    st.subheader("Distribución de 'Ring Time' por agente")
    agentes_ring = st.multiselect("Selecciona agentes para distribución de 'Ring Time':", options=df_filtrado["Agent Name"].unique(), default=df_filtrado["Agent Name"].unique())

    df_ring_filtrado = df_filtrado[df_filtrado["Agent Name"].isin(agentes_ring) & (df_filtrado["Ring Time"] > pd.Timedelta(0))]

    if not df_ring_filtrado.empty:
        fig6, ax6 = plt.subplots(figsize=(10, 5))
        for agente in agentes_ring:
            sns.histplot(df_ring_filtrado[df_ring_filtrado["Agent Name"] == agente]["Ring Time"].dt.total_seconds() / 60, bins=30, kde=True, label=agente, ax=ax6)
        ax6.set_xlabel("Ring Time (minutos)")
        ax6.set_title("Distribución de 'Ring Time'")
        ax6.legend()
        st.pyplot(fig6)

        promedio_ring = df_ring_filtrado.groupby("Agent Name")["Ring Time"].mean()
        promedio_ring_min = promedio_ring.dt.total_seconds() / 60
        st.markdown("### Promedio de 'Ring Time' por agente (minutos)")
        st.dataframe(promedio_ring_min.round(2).to_frame())
    else:
        st.write("No hay datos de 'Ring Time' para los agentes seleccionados en el rango de fechas.")
