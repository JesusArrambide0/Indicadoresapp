import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns

archivo_excel = "App Info.xlsx"
df = pd.read_excel(archivo_excel, engine="openpyxl")

# Normalizar nombres
mapeo_a_nombre_completo = {
    "Jorge": "Jorge Cesar Flores Rivera",
    "Maria": "Maria Teresa Loredo Morales",
    "Jonathan": "Jonathan Alejandro Zúñiga",
}
df["Agent Name"] = df["Agent Name"].replace(mapeo_a_nombre_completo)

df["Call Start Time"] = pd.to_datetime(df["Call Start Time"], errors="coerce")
df["Talk Time"] = pd.to_timedelta(df["Talk Time"], errors="coerce")
df["Fecha"] = df["Call Start Time"].dt.date
df["Hora"] = df["Call Start Time"].dt.hour
df["DíaSemana_En"] = df["Call Start Time"].dt.day_name()

dias_traducidos = {
    "Monday": "Lunes",
    "Tuesday": "Martes",
    "Wednesday": "Miércoles",
    "Thursday": "Jueves",
    "Friday": "Viernes",
    "Saturday": "Sábado",
    "Sunday": "Domingo"
}
df["DíaSemana"] = df["DíaSemana_En"].map(dias_traducidos)

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

st.title("Análisis Integral de Productividad y Llamadas")

# Filtro global por día (incluye "Todos" para no filtrar)
opciones_dias = ["Todos"] + list(dias_traducidos.values())
dia_filtrado = st.selectbox("Filtrar por Día de la Semana:", opciones_dias)

# Filtrar dataframes según día seleccionado
if dia_filtrado != "Todos":
    df_filtrado = df[df["DíaSemana"] == dia_filtrado]
    df_expandido_filtrado = df_expandido[df_expandido["DíaSemana"] == dia_filtrado]
else:
    df_filtrado = df.copy()
    df_expandido_filtrado = df_expandido.copy()

# Recalcular datos con filtro aplicado

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

# Heatmap
pivot_table = df_filtrado.pivot_table(index="Hora", columns="DíaSemana_En", aggfunc="size", fill_value=0)
dias_ordenados_en = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
pivot_table = pivot_table.reindex(columns=dias_ordenados_en, fill_value=0)
pivot_table.columns = [dias_traducidos[d] for d in pivot_table.columns]
pivot_table.index = [f"{h}:00" for h in pivot_table.index]

tab1, tab2, tab3, tab4 = st.tabs(["Detalle por Programador", "Resumen Diario Total", "Heatmap Llamadas", "Productividad General"])

with tab1:
    st.header("Detalle Diario por Programador")
    agente_seleccionado = st.selectbox("Selecciona Programador", options=detalle["AgenteFinal"].unique())
    df_agente = detalle[detalle["AgenteFinal"] == agente_seleccionado].sort_values("Fecha")
    st.dataframe(df_agente.style.format({"Productividad (%)": "{:.2f}"}))

with tab2:
    st.header("Resumen Diario Total (Todos los Programadores)")
    st.dataframe(resumen_diario_todos.style.format({"Productividad (%)": "{:.2f}"}))

with tab3:
    st.header("Distribución de llamadas por hora y día")
    # Si se quiere, aquí se puede eliminar el filtro de día específico para heatmap
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_table, cmap="YlGnBu", annot=True, fmt="d", ax=ax)
    ax.set_xlabel("Día de la Semana")
    ax.set_ylabel("Hora del Día")
    st.pyplot(fig)

with tab4:
    st.header("Productividad y Tasa de Abandono Diaria")
    st.dataframe(df_productividad[["Fecha", "LlamadasRecibidas", "LlamadasPerdidas", "Productividad (%)", "Tasa de Abandono (%)", "DíaSemana"]])
