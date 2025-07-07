
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import matplotlib.pyplot as plt
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Calculadora Horas Extras Fertrac",
    layout="wide",
    page_icon="🕒"
)

# Mostrar el logo y título
st.image("logo_fertrac.png", width=200)
st.markdown(
    "<h2 style='color:#f37021;'>Bienvenido a la herramienta de cálculo de horas extras de Fertrac.</h2>"
    "<p>Por favor, sube los archivos requeridos para comenzar.</p>",
    unsafe_allow_html=True
)

# Subida de archivos
col1, col2, col3 = st.columns(3)

with col1:
    input_file = st.file_uploader("📤 Subir archivo de datos (input_datos.xlsx)", type=["xlsx"], key="input")

with col2:
    empleados_file = st.file_uploader("📤 Subir base de empleados", type=["xlsx"], key="empleados")

with col3:
    porcentaje_file = st.file_uploader("📤 Subir factores de horas extras (%)", type=["xlsx"], key="porcentaje")

if input_file and empleados_file and porcentaje_file:
    df_input = pd.read_excel(input_file)
    df_empleados = pd.read_excel(empleados_file)
    df_porcentaje = pd.read_excel(porcentaje_file)

    # Limpieza de columnas
    df_input.columns = df_input.columns.str.upper().str.strip()
    df_empleados.columns = df_empleados.columns.str.upper().str.strip()
    df_porcentaje.columns = df_porcentaje.columns.str.upper().str.strip()

    # Merge con empleados
    df = df_input.merge(df_empleados, left_on="CÉDULA", right_on="CEDULA", how="left")

    # Conversión de horas
    df["HRA INGRESO"] = pd.to_datetime(df["HRA INGRESO"].astype(str)).dt.time
    df["HORA SALIDA"] = pd.to_datetime(df["HORA SALIDA"].astype(str)).dt.time

    df["HORAS TRABAJADAS"] = (
        pd.to_datetime(df["HORA SALIDA"].astype(str)) - pd.to_datetime(df["HRA INGRESO"].astype(str))
    ).dt.total_seconds() / 3600

    df["HORAS EXTRA"] = df["HORAS TRABAJADAS"].clip(lower=0) - 8
    df["HORAS EXTRA"] = df["HORAS EXTRA"].clip(lower=0)

    df["TIPO EXTRA"] = "Extra Diurna"
    factor_map = dict(zip(df_porcentaje["TIPO HORA EXTRA"].str.upper(), df_porcentaje["FACTOR"]))
    df["FACTOR"] = df["TIPO EXTRA"].map(lambda x: factor_map.get(x.upper(), 1.0))

    df["IMPORTE HORA"] = df["SALARIO BASICO"] / (30 * 8)
    df["VALOR EXTRA"] = df["HORAS EXTRA"] * df["IMPORTE HORA"] * df["FACTOR"]
    df["VALOR TOTAL A PAGAR"] = df["VALOR EXTRA"] + df["COMISIÓN O BONIFICACIÓN"]

    # Mostrar tabla
    st.subheader("📊 Resultados del cálculo")
    st.dataframe(df)

    # Gráfico por persona
    st.subheader("👤 Horas extra por empleado")
    fig1, ax1 = plt.subplots(figsize=(10, 4))
    df.groupby("NOMBRE")["HORAS EXTRA"].sum().plot(kind='bar', ax=ax1, color='#f37021')
    ax1.set_ylabel("Horas Extra")
    st.pyplot(fig1)

    # Gráfico por área
    st.subheader("🏢 Horas extra por área")
    fig2, ax2 = plt.subplots(figsize=(8, 4))
    df.groupby("AREA")["HORAS EXTRA"].sum().plot(kind='barh', ax=ax2, color='#f37021')
    ax2.set_xlabel("Horas Extra")
    st.pyplot(fig2)

    # Descargar archivo
    hoy = date.today().isoformat()
    output_filename = f"resultado_pagos_{hoy}.xlsx"
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    st.download_button(
        label="📥 Descargar archivo Excel",
        data=towrite,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
