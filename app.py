import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, time
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(
    page_title="Calculadora Horas Extras Fertrac",
    layout="wide",
    page_icon="🕒"
)

st.image("logo_fertrac.png", width=200)
st.markdown(
    "<h2 style='color:#f37021;'>Bienvenido a la herramienta de cálculo de horas extras de Fertrac.</h2>"
    "<p>Por favor, sube los archivos requeridos para comenzar.</p>",
    unsafe_allow_html=True
)

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

    # Normalizar columnas
    df_input.columns = df_input.columns.str.upper().str.strip()
    df_empleados.columns = df_empleados.columns.str.upper().str.strip()
    df_porcentaje.columns = df_porcentaje.columns.str.upper().str.strip()

    # Merge con empleados
    df = df_input.merge(df_empleados, left_on="CÉDULA", right_on="CEDULA", how="left")

    # Procesamiento de fechas y horas
    df["FECHA"] = pd.to_datetime(df["FECHA"])
    df["DIA_NUM"] = df["FECHA"].dt.weekday  # 0=Lun, 5=Sáb, 6=Dom

    df["HRA INGRESO"] = pd.to_datetime(df["HRA INGRESO"].astype(str)).dt.time
    df["HORA SALIDA"] = pd.to_datetime(df["HORA SALIDA"].astype(str)).dt.time

    def convertir_a_datetime(fecha, hora):
        return pd.to_datetime(fecha.astype(str) + ' ' + hora.astype(str))

    df["DT_INGRESO"] = convertir_a_datetime(df["FECHA"], df["HRA INGRESO"])
    df["DT_SALIDA"] = convertir_a_datetime(df["FECHA"], df["HORA SALIDA"])

    # Definir horarios de fin según turno
    def obtener_hora_fin_jornada(row):
        """Retorna la hora de fin de jornada según el turno y día"""
        turno = row.get("TURNO", "TURNO 1").upper()
        dia = row["DIA_NUM"]
        
        if dia == 5:  # Sábado
            return datetime.combine(row["FECHA"], time(12, 0))
        else:  # Lunes a viernes
            if turno == "TURNO 2":
                return datetime.combine(row["FECHA"], time(22, 0))  # 10 PM
            else:  # TURNO 1 (por defecto)
                return datetime.combine(row["FECHA"], time(18, 0))  # 6 PM

    def calcular_horas_extras_y_recargo(row):
        """Calcula horas extras y recargo nocturno por separado"""
        ingreso = row["DT_INGRESO"]
        salida = row["DT_SALIDA"]
        fin_jornada = obtener_hora_fin_jornada(row)
        
        # Definir límites de recargo nocturno (7 PM - 9 PM)
        inicio_nocturno = datetime.combine(row["FECHA"], time(19, 0))
        fin_nocturno = datetime.combine(row["FECHA"], time(21, 0))
        
        horas_extra = 0
        recargo_nocturno = 0
        
        # Si salió después de la jornada
        if salida > fin_jornada:
            tiempo_despues_jornada = (salida - fin_jornada).total_seconds() / 3600
            
            # Separar en horas extras y recargo nocturno
            if salida > inicio_nocturno:
                # Parte antes de 7 PM = horas extras normales
                if fin_jornada < inicio_nocturno:
                    horas_extra = (inicio_nocturno - fin_jornada).total_seconds() / 3600
                
                # Parte entre 7 PM y 9 PM = recargo nocturno
                inicio_recargo = max(fin_jornada, inicio_nocturno)
                fin_recargo = min(salida, fin_nocturno)
                if fin_recargo > inicio_recargo:
                    recargo_nocturno = (fin_recargo - inicio_recargo).total_seconds() / 3600
                
                # Parte después de 9 PM = horas extras normales
                if salida > fin_nocturno:
                    horas_extra += (salida - fin_nocturno).total_seconds() / 3600
            else:
                # Todo es hora extra normal (antes de 7 PM)
                horas_extra = tiempo_despues_jornada
        
        return max(horas_extra, 0), max(recargo_nocturno, 0)

    def calcular_trabajo_real(row):
        """Calcula horas trabajadas descontando 1 hora de almuerzo"""
        ingreso = row["DT_INGRESO"]
        salida = row["DT_SALIDA"]
        dia = row["DIA_NUM"]

        total = (salida - ingreso).total_seconds() / 3600
        
        # Descontar 1 hora de almuerzo en días de lunes a viernes
        if dia < 5:  # Lunes a viernes
            return total - 1
        else:  # Sábados y domingos
            return total

    # Aplicar cálculos
    df["HORAS TRABAJADAS"] = df.apply(calcular_trabajo_real, axis=1)
    
    # Calcular horas extras y recargo nocturno
    resultados = df.apply(calcular_horas_extras_y_recargo, axis=1)
    df["HORAS EXTRA"] = [r[0] for r in resultados]
    df["RECARGO NOCTURNO"] = [r[1] for r in resultados]

    # Asignar tipos de hora extra
    df["TIPO EXTRA"] = df.apply(
        lambda row: "Recargo Nocturno" if row["RECARGO NOCTURNO"] > 0 else "Extra Diurna",
        axis=1
    )

    # Mapear factores
    factor_map = dict(zip(df_porcentaje["TIPO HORA EXTRA"].str.upper(), df_porcentaje["FACTOR"]))
    
    # Calcular valores monetarios
    # IMPORTANTE: Dividir por 220 horas (aplicable desde 2do semestre 2025)
    df["IMPORTE HORA"] = df["SALARIO BASICO"] / 220
    
    # Calcular valor de horas extras normales
    df["VALOR EXTRA DIURNA"] = df.apply(
        lambda row: row["HORAS EXTRA"] * row["IMPORTE HORA"] * factor_map.get("EXTRA DIURNA", 1.25),
        axis=1
    )
    
    # Calcular valor de recargo nocturno
    df["VALOR RECARGO NOCTURNO"] = df.apply(
        lambda row: row["RECARGO NOCTURNO"] * row["IMPORTE HORA"] * factor_map.get("RECARGO NOCTURNO", 1.35),
        axis=1
    )
    
    # Valor total de extras
    df["VALOR EXTRA"] = df["VALOR EXTRA DIURNA"] + df["VALOR RECARGO NOCTURNO"]
    df["VALOR TOTAL A PAGAR"] = df["VALOR EXTRA"] + df["COMISIÓN O BONIFICACIÓN"]

    # Redondear valores
    df["IMPORTE HORA"] = df["IMPORTE HORA"].round(2)
    df["VALOR EXTRA"] = df["VALOR EXTRA"].round(2)
    df["VALOR EXTRA DIURNA"] = df["VALOR EXTRA DIURNA"].round(2)
    df["VALOR RECARGO NOCTURNO"] = df["VALOR RECARGO NOCTURNO"].round(2)
    df["VALOR TOTAL A PAGAR"] = df["VALOR TOTAL A PAGAR"].round(2)
    df["HORAS TRABAJADAS"] = df["HORAS TRABAJADAS"].round(2)
    df["HORAS EXTRA"] = df["HORAS EXTRA"].round(2)
    df["RECARGO NOCTURNO"] = df["RECARGO NOCTURNO"].round(2)

    # Resumen informativo
    st.info(f"""
    ✅ **Configuración aplicada:**
    - División por **220 horas** mensuales (vigente desde 2do semestre 2025)
    - Descuento de **1 hora de almuerzo** en días de lunes a viernes
    - **Recargo nocturno** de 7 PM a 9 PM con factor diferenciado
    - Soporte para **Turno 1** (hasta 6 PM) y **Turno 2** (hasta 10 PM)
    """)

    # Agregar columna de mes y año para análisis temporal
    df["MES"] = df["FECHA"].dt.to_period('M')
    df["MES_NOMBRE"] = df["FECHA"].dt.strftime('%B %Y')

    # FILTRO POR ÁREA
    st.subheader("🔍 Filtros de visualización")
    col_filtro1, col_filtro2 = st.columns(2)
    
    with col_filtro1:
        areas_disponibles = ["Todas las áreas"] + sorted(df["AREA"].unique().tolist())
        area_seleccionada = st.selectbox("Seleccionar área:", areas_disponibles)
    
    with col_filtro2:
        meses_disponibles = ["Todos los meses"] + sorted(df["MES_NOMBRE"].unique().tolist())
        mes_seleccionado = st.selectbox("Seleccionar mes:", meses_disponibles)
    
    # Aplicar filtros
    df_filtrado = df.copy()
    if area_seleccionada != "Todas las áreas":
        df_filtrado = df_filtrado[df_filtrado["AREA"] == area_seleccionada]
    if mes_seleccionado != "Todos los meses":
        df_filtrado = df_filtrado[df_filtrado["MES_NOMBRE"] == mes_seleccionado]

    st.subheader("📊 Resultados del cálculo")
    
    # Columnas a mostrar
    columnas_mostrar = [
        "FECHA", "NOMBRE", "CÉDULA", "AREA", "TURNO", 
        "HRA INGRESO", "HORA SALIDA", "HORAS TRABAJADAS",
        "HORAS EXTRA", "RECARGO NOCTURNO", "ACTIVIDAD DESARROLLADA",
        "SALARIO BASICO", "IMPORTE HORA", "VALOR EXTRA DIURNA",
        "VALOR RECARGO NOCTURNO", "VALOR EXTRA", "VALOR TOTAL A PAGAR"
    ]
    
    # Filtrar solo las columnas que existen
    columnas_disponibles = [col for col in columnas_mostrar if col in df_filtrado.columns]
    st.dataframe(df_filtrado[columnas_disponibles], use_container_width=True)

    # Visualizaciones
    col_viz1, col_viz2 = st.columns(2)
    
    with col_viz1:
        st.subheader("👤 Horas extra por empleado")
        fig1, ax1 = plt.subplots(figsize=(10, 6))
        empleado_stats = df_filtrado.groupby("NOMBRE").agg({
            "HORAS EXTRA": "sum",
            "RECARGO NOCTURNO": "sum"
        })
        if not empleado_stats.empty:
            empleado_stats.plot(kind='bar', ax=ax1, color=['#f37021', '#ff9966'], stacked=True)
            ax1.set_ylabel("Horas")
            ax1.set_xlabel("Empleado")
            ax1.legend(["Horas Extra Diurnas", "Recargo Nocturno"])
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig1)
        else:
            st.info("No hay datos para mostrar con los filtros seleccionados")

    with col_viz2:
        st.subheader("🏢 Horas extra por área")
        fig2, ax2 = plt.subplots(figsize=(10, 6))
        area_stats = df_filtrado.groupby("AREA").agg({
            "HORAS EXTRA": "sum",
            "RECARGO NOCTURNO": "sum"
        })
        if not area_stats.empty:
            area_stats.plot(kind='barh', ax=ax2, color=['#f37021', '#ff9966'], stacked=True)
            ax2.set_xlabel("Horas")
            ax2.set_ylabel("Área")
            ax2.legend(["Horas Extra Diurnas", "Recargo Nocturno"])
            plt.tight_layout()
            st.pyplot(fig2)
        else:
            st.info("No hay datos para mostrar con los filtros seleccionados")

    # COMPARATIVO MENSUAL
    st.subheader("📅 Comparativo mensual")
    
    # Agrupar por mes
    comparativo_mensual = df.groupby("MES_NOMBRE").agg({
        "HORAS EXTRA": "sum",
        "RECARGO NOCTURNO": "sum",
        "VALOR EXTRA": "sum",
        "VALOR TOTAL A PAGAR": "sum"
    }).reset_index()
    
    # Ordenar por fecha
    comparativo_mensual["ORDEN"] = pd.to_datetime(comparativo_mensual["MES_NOMBRE"], format='%B %Y')
    comparativo_mensual = comparativo_mensual.sort_values("ORDEN")
    
    col_comp1, col_comp2 = st.columns(2)
    
    with col_comp1:
        st.markdown("#### Horas por mes")
        fig3, ax3 = plt.subplots(figsize=(10, 6))
        x = range(len(comparativo_mensual))
        width = 0.35
        ax3.bar([i - width/2 for i in x], comparativo_mensual["HORAS EXTRA"], 
                width, label='Horas Extra Diurnas', color='#f37021')
        ax3.bar([i + width/2 for i in x], comparativo_mensual["RECARGO NOCTURNO"], 
                width, label='Recargo Nocturno', color='#ff9966')
        ax3.set_xlabel('Mes')
        ax3.set_ylabel('Horas')
        ax3.set_xticks(x)
        ax3.set_xticklabels(comparativo_mensual["MES_NOMBRE"], rotation=45, ha='right')
        ax3.legend()
        plt.tight_layout()
        st.pyplot(fig3)
    
    with col_comp2:
        st.markdown("#### Costos por mes")
        fig4, ax4 = plt.subplots(figsize=(10, 6))
        ax4.plot(comparativo_mensual["MES_NOMBRE"], comparativo_mensual["VALOR_EXTRA"], 
                marker='o', linewidth=2, markersize=8, color='#f37021', label='Valor Extras')
        ax4.plot(comparativo_mensual["MES_NOMBRE"], comparativo_mensual["VALOR TOTAL A PAGAR"], 
                marker='s', linewidth=2, markersize=8, color='#ff9966', label='Total a Pagar')
        ax4.set_xlabel('Mes')
        ax4.set_ylabel('Valor ($)')
        ax4.legend()
        plt.xticks(rotation=45, ha='right')
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig4)
    
    # Tabla comparativa mensual
    st.markdown("#### Tabla comparativa mensual")
    tabla_comparativa = comparativo_mensual[["MES_NOMBRE", "HORAS EXTRA", "RECARGO NOCTURNO", 
                                             "VALOR EXTRA", "VALOR TOTAL A PAGAR"]].copy()
    tabla_comparativa.columns = ["Mes", "Horas Extra", "Recargo Nocturno", 
                                  "Valor Extras ($)", "Total a Pagar ($)"]
    
    # Formatear valores monetarios
    tabla_comparativa["Valor Extras ($)"] = tabla_comparativa["Valor Extras ($)"].apply(lambda x: f"${x:,.2f}")
    tabla_comparativa["Total a Pagar ($)"] = tabla_comparativa["Total a Pagar ($)"].apply(lambda x: f"${x:,.2f}")
    tabla_comparativa["Horas Extra"] = tabla_comparativa["Horas Extra"].apply(lambda x: f"{x:.2f}")
    tabla_comparativa["Recargo Nocturno"] = tabla_comparativa["Recargo Nocturno"].apply(lambda x: f"{x:.2f}")
    
    st.dataframe(tabla_comparativa, use_container_width=True, hide_index=True)

    # Resumen estadístico (con datos filtrados)
    st.subheader("📈 Resumen general")
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
    
    with col_stat1:
        st.metric("Total Horas Extra Diurnas", f"{df_filtrado['HORAS EXTRA'].sum():.2f}")
    
    with col_stat2:
        st.metric("Total Recargo Nocturno", f"{df_filtrado['RECARGO NOCTURNO'].sum():.2f}")
    
    with col_stat3:
        st.metric("Total a Pagar Extras", f"${df_filtrado['VALOR EXTRA'].sum():,.2f}")
    
    with col_stat4:
        st.metric("Total General", f"${df_filtrado['VALOR TOTAL A PAGAR'].sum():,.2f}")

    # Descarga
    hoy = date.today().isoformat()
    output_filename = f"resultado_pagos_{hoy}.xlsx"
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    st.download_button(
        label="📥 Descargar archivo Excel con resultados",
        data=towrite,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
