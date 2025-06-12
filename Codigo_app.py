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

# Expandir filas para llamadas perdidas con asignación por agente según horario
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

fecha_min = df["Fecha"].min()
fecha_max = df["Fecha"].max()
fecha_inicio, fecha_fin = st.date_input("Selecciona un rango de fechas:", [fecha_min, fecha_max])

df_filtrado = df[(df["Fecha"] >= fecha_inicio) & (df["Fecha"] <= fecha_fin)]
df_expandido_filtrado = df_expandido[(df_expandido["Fecha"] >= fecha_inicio) & (df_expandido["Fecha"] <= fecha_fin)]

# --- AJUSTE PRODUCTIVIDAD GENERAL: contar llamadas perdidas únicas (sin duplicados)
llamadas_perdidas_unicas = df_filtrado[df_filtrado["LlamadaPerdida"]].drop_duplicates(subset=["Call Start Time", "Talk Time"])
num_llamadas_perdidas_unicas_por_fecha = llamadas_perdidas_unicas.groupby("Fecha").size()

df_productividad = df_filtrado.groupby("Fecha").agg(
    LlamadasRecibidas=("Talk Time", "count"),
).reset_index()

df_productividad["LlamadasPerdidas"] = df_productividad["Fecha"].map(num_llamadas_perdidas_unicas_por_fecha).fillna(0).astype(int)

df_productividad["Productividad (%)"] = (
    (df_productividad["LlamadasRecibidas"] - df_productividad["LlamadasPerdidas"]) / df_productividad["LlamadasRecibidas"] * 100
).round(2)
df_productividad["Tasa de Abandono (%)"] = 100 - df_productividad["Productividad (%)"]
df_productividad["DíaSemana"] = pd.to_datetime(df_productividad["Fecha"]).dt.day_name().map(dias_traducidos)

# Detalle diario por programador con llamadas perdidas asignadas por agente según horarios (ya expandido)
detalle = df_expandido_filtrado.groupby(["AgenteFinal", "Fecha"]).agg(
    LlamadasTotales=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum"),
    TalkTimeTotal=("Talk Time", "sum")
).reset_index()
detalle["LlamadasAtendidas"] = detalle["LlamadasTotales"] - detalle["LlamadasPerdidas"]
detalle["Productividad (%)"] = (detalle["LlamadasAtendidas"] / detalle["LlamadasTotales"] * 100).round(2)
detalle["Promedio Talk Time (seg)"] = (detalle["TalkTimeTotal"].dt.total_seconds() / detalle["LlamadasAtendidas"]).round(2)

# (el resto del código sigue igual...)

# Días válidos sin domingo y en orden
dias_validos = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
dias_validos_es = [dias_traducidos[d] for d in dias_validos]

pivot_table = df_filtrado[df_filtrado["DíaSemana_En"].isin(dias_validos)].pivot_table(
    index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0
)
pivot_table = pivot_table.reindex(columns=dias_validos, fill_value=0)
pivot_table.columns = [dias_traducidos[d] for d in pivot_table.columns]
pivot_table = pivot_table.sort_index(ascending=True)

horas_ordenadas = list(range(8, 21))
pivot_table = pivot_table.reindex(horas_ordenadas[::-1], fill_value=0)
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
pivot_table_perdidas = pivot_table_perdidas.reindex(horas_ordenadas[::-1], fill_value=0)
pivot_table_perdidas.index = [f"{h}:00" for h in pivot_table_perdidas.index]

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
    st.pyplot(fig)

with tab2:
    st.header("Detalle Diario por Programador")

    def color_fila_tab2(row):
        valor = row["Productividad (%)"]
        if valor >= 97:
            color = "background-color: #28a745; color: white;"
        elif 90 <= valor < 97:
            color = "background-color: #fff3cd; color: black;"
        else:
            color = "background-color: #dc3545; color: white;"
        return [color] * len(row)

    styled_detalle = detalle[["AgenteFinal", "Fecha", "LlamadasTotales", "LlamadasPerdidas", "LlamadasAtendidas", "Productividad (%)", "Promedio Talk Time (seg)"]].style.apply(color_fila_tab2, axis=1).format({"Productividad (%)": "{:.2f}", "Promedio Talk Time (seg)": "{:.2f}"})
    st.dataframe(styled_detalle)

    agentes_seleccionados = st.multiselect("Selecciona programadores para ver gráfica:", options=detalle["AgenteFinal"].unique(), default=detalle["AgenteFinal"].unique())

    if agentes_seleccionados:
        fig2, ax2 = plt.subplots(figsize=(10, 4))
        for agente in agentes_seleccionados:
            df_plot = detalle[detalle["AgenteFinal"] == agente]
            ax2.plot(df_plot["Fecha"], df_plot["Productividad (%)"], marker="o", label=agente)
        ax2.axhline(97, color="green", linestyle="--", label="Meta 97%")
        ax2.axhline(90, color="orange", linestyle="--", label="Alerta 90%")
        ax2.set_ylim(0, 105)
        ax2.set_ylabel("Productividad (%)")
        ax2.set_title("Productividad diaria por programador")
        ax2.legend()
        st.pyplot(fig2)

with tab3:
    st.header("Heatmap Llamadas Entrantes (Horas vs Día)")

    fig3, ax3 = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_table, annot=True, fmt="d", cmap="YlGnBu", ax=ax3)
    ax3.set_xlabel("Día de la semana")
    ax3.set_ylabel("Hora del día")
    st.pyplot(fig3)

with tab4:
    st.header("Heatmap Llamadas Perdidas (Horas vs Día)")

    fig4, ax4 = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_table_perdidas, annot=True, fmt="d", cmap="YlOrRd", ax=ax4)
    ax4.set_xlabel("Día de la semana")
    ax4.set_ylabel("Hora del día")
    st.pyplot(fig4)

with tab5:
    st.header("Distribución de Duración de Llamadas y Alertas")

    st.subheader("Distribución de 'Talk Time' por agente")
    agentes_dist = st.multiselect("Selecciona agentes para distribución:", options=df_filtrado["Agent Name"].unique(), default=df_filtrado["Agent Name"].unique())

    df_dist_filtrado = df_filtrado[df_filtrado["Agent Name"].isin(agentes_dist) & (df_filtrado["Talk Time"] > pd.Timedelta(0))]

    if not df_dist_filtrado.empty:
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
