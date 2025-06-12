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

# Filtro por agentes para la pestaña detalle
agentes_unicos = sorted(detalle["AgenteFinal"].unique())
agentes_seleccionados = st.sidebar.multiselect("Selecciona Agentes para Detalle", agentes_unicos, default=agentes_unicos)

detalle_filtrado = detalle[detalle["AgenteFinal"].isin(agentes_seleccionados)]

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

# Dentro de Tab 2: filtro de agentes + boxplot Talk Time + tabla filtrada
with tab2:
    st.header("Detalle Diario por Programador")

    # Selector de agentes solo aquí (tab2)
    agentes_unicos = sorted(detalle["AgenteFinal"].unique())
    agentes_seleccionados = st.multiselect("Selecciona Agentes para Detalle", agentes_unicos, default=agentes_unicos)

    detalle_filtrado = detalle[detalle["AgenteFinal"].isin(agentes_seleccionados)]

    st.subheader("Boxplot de Talk Time por Agente")
    df_box = df_expandido_filtrado[
        (df_expandido_filtrado["AgenteFinal"].isin(agentes_seleccionados)) &
        (df_expandido_filtrado["Talk Time"] > pd.Timedelta(0))
    ].copy()
    df_box["TalkTime_seg"] = df_box["Talk Time"].dt.total_seconds()

    fig_box, ax_box = plt.subplots(figsize=(10, 5))
    sns.boxplot(x="AgenteFinal", y="TalkTime_seg", data=df_box, ax=ax_box)
    ax_box.set_ylabel("Talk Time (segundos)")
    ax_box.set_xlabel("Agente")
    ax_box.set_title("Distribución de Talk Time por Agente")
    st.pyplot(fig_box)

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

    detalle_mostrar = detalle_filtrado.sort_values(["AgenteFinal", "Fecha"])
    styled_detalle = detalle_mostrar.style.apply(color_fila_tab2, axis=1).format({
        "Productividad (%)": "{:.2f}",
        "Promedio Talk Time (seg)": "{:.1f}"
    })
    st.dataframe(styled_detalle)
with tab3:
    st.header("Heatmap de Llamadas Recibidas")

    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(pivot_table, annot=True, fmt="d", cmap="Blues", ax=ax)
    ax.set_title("Llamadas Recibidas por Hora y Día")
    st.pyplot(fig)

with tab4:
    st.header("Heatmap de Llamadas Perdidas")

    fig2, ax2 = plt.subplots(figsize=(10, 8))
    sns.heatmap(pivot_table_perdidas, annot=True, fmt="d", cmap="Reds", ax=ax2)
    ax2.set_title("Llamadas Perdidas por Hora y Día")
    st.pyplot(fig2)

# Dentro de Tab 5: análisis adicional para Ring Time
with tab5:
    st.header("Distribución de Duración de Llamadas y Alertas")

    st.subheader("Distribución de Duración (Talk Time en segundos)")
    duracion_seg = df_filtrado["Talk Time"].dt.total_seconds().dropna()
    fig3, ax3 = plt.subplots()
    sns.histplot(duracion_seg, bins=30, ax=ax3)
    ax3.set_xlabel("Duración (segundos)")
    ax3.set_ylabel("Cantidad de llamadas")
    st.pyplot(fig3)

    st.subheader("Distribución de Duración (Ring Time en segundos)")
    ring_time_seg = df_filtrado["Ring Time"].dt.total_seconds().dropna()
    fig4, ax4 = plt.subplots()
    sns.histplot(ring_time_seg, bins=30, color='orange', ax=ax4)
    ax4.set_xlabel("Ring Time (segundos)")
    ax4.set_ylabel("Cantidad de llamadas")
    st.pyplot(fig4)

    st.subheader("Promedio Ring Time por Día")
    promedio_ring_por_dia = df_filtrado.groupby("Fecha")["Ring Time"].mean().dt.total_seconds()
    st.line_chart(promedio_ring_por_dia)

    # Alertas por picos - ejemplo sencillo de llamadas por hora (Talk Time)
    llamadas_por_hora = df_filtrado.groupby("Hora").size()
    umbral_pico = llamadas_por_hora.mean() + 2 * llamadas_por_hora.std()
    picos = llamadas_por_hora[llamadas_por_hora > umbral_pico]

    if not picos.empty:
        st.warning(f"Se detectaron picos inusuales de llamadas en las horas: {', '.join(map(str, picos.index))}")
    else:
        st.info("No se detectaron picos inusuales de llamadas.")
