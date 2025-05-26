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

# Asignar agente según horario
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
    return []

# Expandir filas
df["AgenteFinal"] = df.apply(
    lambda row: row["Agent Name"] if not row["LlamadaPerdida"]
    else agentes_por_horario(row["Hora"])[0] if agentes_por_horario(row["Hora"])
    else row["Agent Name"], axis=1
)
df_expandido = df[df["AgenteFinal"].notna()].copy()

# Interfaz Streamlit
st.title("Análisis Integral de Productividad y Llamadas")

# Rango de fechas
fecha_min, fecha_max = df["Fecha"].min(), df["Fecha"].max()
fecha_inicio, fecha_fin = st.date_input("Selecciona un rango de fechas:", [fecha_min, fecha_max])

# Filtrado
df_filtrado = df[(df["Fecha"] >= fecha_inicio) & (df["Fecha"] <= fecha_fin)]
df_expandido_filtrado = df_expandido[(df_expandido["Fecha"] >= fecha_inicio) & (df_expandido["Fecha"] <= fecha_fin)]

# Productividad diaria
df_productividad = df_filtrado.groupby("Fecha").agg(
    LlamadasRecibidas=("Talk Time", "count"),
    LlamadasPerdidas=("Talk Time", lambda x: (x == pd.Timedelta("0:00:00")).sum())
).reset_index()
df_productividad["Productividad (%)"] = (
    (df_productividad["LlamadasRecibidas"] - df_productividad["LlamadasPerdidas"]) / df_productividad["LlamadasRecibidas"] * 100).round(2)
df_productividad["Tasa de Abandono (%)"] = 100 - df_productividad["Productividad (%)"]
df_productividad["DíaSemana"] = pd.to_datetime(df_productividad["Fecha"]).dt.day_name().map(dias_traducidos)

# Detalle diario por agente
detalle = df_expandido_filtrado.groupby(["AgenteFinal", "Fecha"]).agg(
    LlamadasTotales=("Talk Time", "count"),
    LlamadasPerdidas=("LlamadaPerdida", "sum"),
    TalkTimeTotal=("Talk Time", "sum")
).reset_index()
detalle["LlamadasAtendidas"] = detalle["LlamadasTotales"] - detalle["LlamadasPerdidas"]
detalle["Productividad (%)"] = (detalle["LlamadasAtendidas"] / detalle["LlamadasTotales"] * 100).round(2)
detalle["Promedio Talk Time (seg)"] = (detalle["TalkTimeTotal"].dt.total_seconds() / detalle["LlamadasAtendidas"]).round(2)

# Heatmaps
dias_validos = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
dias_validos_es = [dias_traducidos[d] for d in dias_validos]
horas_ordenadas = list(range(8, 21))

pivot = df_filtrado[df_filtrado["DíaSemana_En"].isin(dias_validos)].pivot_table(
    index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0)
pivot = pivot.reindex(columns=dias_validos).rename(columns=dias_traducidos).reindex(horas_ordenadas[::-1])
pivot.index = [f"{h}:00" for h in pivot.index]

pivot_perdidas = df_expandido_filtrado[(df_expandido_filtrado["DíaSemana_En"].isin(dias_validos)) & (df_expandido_filtrado["LlamadaPerdida"])]
pivot_perdidas = pivot_perdidas.pivot_table(index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0)
pivot_perdidas = pivot_perdidas.reindex(columns=dias_validos).rename(columns=dias_traducidos).reindex(horas_ordenadas[::-1])
pivot_perdidas.index = [f"{h}:00" for h in pivot_perdidas.index]

# Tabs
with st.tabs(["Productividad General", "Detalle por Programador", "Heatmap Llamadas", "Heatmap Llamadas Perdidas", "Distribución & Alertas"]) as (tab1, tab2, tab3, tab4, tab5):
    with tab1:
        st.header("Productividad y Tasa de Abandono Diaria")
        st.dataframe(df_productividad)

    with tab2:
        st.header("Detalle Diario por Programador")
        agente = st.selectbox("Selecciona Programador", detalle["AgenteFinal"].unique())
        st.dataframe(detalle[detalle["AgenteFinal"] == agente])

    with tab3:
        st.header("Distribución de llamadas por hora y día")
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.heatmap(pivot, cmap="YlGnBu", annot=True, fmt="d", ax=ax)
        ax.set_xlabel("Día de la Semana")
        ax.set_ylabel("Hora del Día")
        st.pyplot(fig)

    with tab4:
        st.header("Distribución de llamadas perdidas por hora y día")
        fig2, ax2 = plt.subplots(figsize=(10, 6))
        sns.heatmap(pivot_perdidas, cmap="OrRd", annot=True, fmt="d", ax=ax2)
        ax2.set_xlabel("Día de la Semana")
        ax2.set_ylabel("Hora del Día")
        st.pyplot(fig2)

    with tab5:
        st.header("Distribución y Promedio del Tiempo de Conversación")
        agente = st.selectbox("Selecciona agente", df_expandido_filtrado["AgenteFinal"].unique(), key="dist")
        df_agente = df_expandido_filtrado[(df_expandido_filtrado["AgenteFinal"] == agente) & (~df_expandido_filtrado["LlamadaPerdida"])]

        if not df_agente.empty:
            tiempos = df_agente["Talk Time"].dt.total_seconds()
            fig_hist, ax_hist = plt.subplots()
            ax_hist.hist(tiempos, bins=30, color='skyblue', edgecolor='black')
            ax_hist.set_title(f"Distribución de Talk Time - {agente}")
            ax_hist.set_xlabel("Segundos")
            ax_hist.set_ylabel("Frecuencia")
            st.pyplot(fig_hist)
            st.write(f"Promedio de Talk Time para {agente}: **{tiempos.mean():.2f} segundos**")
        else:
            st.write(f"No hay llamadas atendidas para {agente} en el rango seleccionado.")

        st.header("Alertas de Picos o Pérdidas")
        resumen = df_expandido_filtrado.groupby(["AgenteFinal", "Hora"]).agg(
            TotalLlamadas=("Talk Time", "count"),
            LlamadasPerdidas=("LlamadaPerdida", "sum")
        ).reset_index()

        for agente in resumen["AgenteFinal"].unique():
            datos = resumen[resumen["AgenteFinal"] == agente]
            umbral_llamadas = datos["TotalLlamadas"].mean() + 2 * datos["TotalLlamadas"].std()
            umbral_perdidas = datos["LlamadasPerdidas"].mean() + 2 * datos["LlamadasPerdidas"].std()

            for _, fila in datos.iterrows():
                if fila["TotalLlamadas"] > umbral_llamadas:
                    st.warning(f"⚠️ Pico de llamadas para {agente} a las {fila['Hora']}:00 - {int(fila['TotalLlamadas'])} llamadas.")
                if fila["LlamadasPerdidas"] > umbral_perdidas:
                    st.warning(f"⚠️ Pico de llamadas perdidas para {agente} a las {fila['Hora']}:00 - {int(fila['LlamadasPerdidas'])} perdidas.")
