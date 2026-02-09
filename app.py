import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, time, timedelta
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

col1, col2 = st.columns(2)

with col1:
    input_file = st.file_uploader("📤 Subir archivo de datos (input_datos.xlsx)", type=["xlsx"], key="input")
    empleados_file = st.file_uploader("📤 Subir base de empleados", type=["xlsx"], key="empleados")

with col2:
    porcentaje_file = st.file_uploader("📤 Subir factores de horas extras (%)", type=["xlsx"], key="porcentaje")
    turnos_file = st.file_uploader("📤 Subir configuración de turnos", type=["xlsx"], key="turnos")

if input_file and empleados_file and porcentaje_file and turnos_file:
    df_input = pd.read_excel(input_file)
    df_empleados = pd.read_excel(empleados_file)
    df_porcentaje = pd.read_excel(porcentaje_file)
    df_turnos = pd.read_excel(turnos_file)

    # Normalizar columnas
    df_input.columns = df_input.columns.str.upper().str.strip()
    df_empleados.columns = df_empleados.columns.str.upper().str.strip()
    df_porcentaje.columns = df_porcentaje.columns.str.upper().str.strip()
    df_turnos.columns = df_turnos.columns.str.upper().str.strip()

    # Normalizar nombres de columnas de cédula
    # El archivo de input puede tener "CÉDULA" y el de empleados "CEDULA"
    if "CÉDULA" in df_input.columns:
        cedula_input = "CÉDULA"
    elif "CEDULA" in df_input.columns:
        cedula_input = "CEDULA"
    else:
        st.error("⚠️ Error: No se encontró columna de cédula en archivo de datos")
        st.info("La columna debe llamarse 'CÉDULA' o 'CEDULA'")
        st.stop()
    
    if "CEDULA" in df_empleados.columns:
        cedula_empleados = "CEDULA"
    elif "CÉDULA" in df_empleados.columns:
        cedula_empleados = "CÉDULA"
    else:
        st.error("⚠️ Error: No se encontró columna de cédula en base de empleados")
        st.info("La columna debe llamarse 'CEDULA' o 'CÉDULA'")
        st.stop()
    
    # Convertir cédulas a string para evitar problemas de tipo
    df_input[cedula_input] = df_input[cedula_input].astype(str).str.strip()
    df_empleados[cedula_empleados] = df_empleados[cedula_empleados].astype(str).str.strip()
    
    # Merge con empleados
    df = df_input.merge(df_empleados, left_on=cedula_input, right_on=cedula_empleados, how="left")
    
    # Validar que existan las columnas necesarias del archivo de empleados
    columnas_requeridas = ["NOMBRE", "AREA", "SALARIO BASICO"]
    columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
    
    if columnas_faltantes:
        st.error(f"⚠️ Error: Faltan columnas en el archivo de empleados: {', '.join(columnas_faltantes)}")
        st.info(f"""
        El archivo de empleados debe tener las siguientes columnas:
        - CEDULA o CÉDULA
        - NOMBRE
        - AREA
        - SALARIO BASICO
        - CARGO (opcional, pero recomendado)
        """)
        st.stop()
    
    # Verificar si existe CARGO (opcional pero recomendado)
    if "CARGO" not in df.columns:
        st.warning("⚠️ Advertencia: No se encontró columna 'CARGO' en el archivo de empleados. Se dejará vacía.")
        df["CARGO"] = ""
    
    # COMISIÓN O BONIFICACIÓN debe venir del archivo input_datos
    # Si no existe en input_datos, llenar con 0
    if "COMISIÓN O BONIFICACIÓN" not in df.columns:
        if "COMISION O BONIFICACION" not in df.columns:
            st.warning("⚠️ Advertencia: No se encontró columna 'COMISIÓN O BONIFICACIÓN' en el archivo de datos de entrada. Se asumirá valor $0 para todos.")
            df["COMISIÓN O BONIFICACIÓN"] = 0
        else:
            df["COMISIÓN O BONIFICACIÓN"] = df["COMISION O BONIFICACION"]
    
    # Agregar columna OBSERVACIONES si no existe en el input
    if "OBSERVACIONES" not in df.columns and "OBSERVACION" not in df.columns:
        df["OBSERVACIONES"] = ""  # Columna vacía por defecto
    elif "OBSERVACION" in df.columns:
        df["OBSERVACIONES"] = df["OBSERVACION"]  # Renombrar si existe con otro nombre
    
    # Verificar que se hayan encontrado coincidencias
    # Primero verificar si existe la columna NOMBRE después del merge
    if "NOMBRE" in df.columns:
        empleados_sin_info = df[df["NOMBRE"].isna()]
        if len(empleados_sin_info) > 0:
            st.warning(f"⚠️ Advertencia: {len(empleados_sin_info)} registros no tienen información de empleado")
            st.warning(f"Cédulas sin coincidencia: {empleados_sin_info[cedula_input].unique().tolist()[:5]}")
            st.info("Verifica que las cédulas en ambos archivos coincidan exactamente")
    else:
        st.error("⚠️ Error: No se encontró la columna NOMBRE en el archivo de empleados")
        st.info("El archivo de empleados debe tener una columna llamada 'NOMBRE' o 'Nombre'")
        st.stop()

    # Procesamiento de fechas y horas
    try:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors='coerce')
        
        # Verificar que las fechas se convirtieron correctamente
        if df["FECHA"].isna().any():
            fechas_problema = df[df["FECHA"].isna()].index.tolist()
            st.error(f"⚠️ Error: Algunas fechas no tienen el formato correcto en las filas: {fechas_problema[:5]}")
            st.info("Las fechas deben estar en formato: YYYY-MM-DD (ej: 2025-02-04) o DD/MM/YYYY")
            st.stop()
        
        df["DIA_NUM"] = df["FECHA"].dt.weekday  # 0=Lun, 5=Sáb, 6=Dom
    except Exception as e:
        st.error(f"⚠️ Error al procesar fechas: {str(e)}")
        st.info("Verifica que la columna FECHA tenga fechas válidas")
        st.stop()

    try:
        # Función para convertir horas de cualquier formato a time
        def convertir_a_time(valor):
            if pd.isna(valor):
                return None
            # Si ya es un objeto time, devolverlo
            if isinstance(valor, time):
                return valor
            # Si es datetime, extraer solo la hora
            if isinstance(valor, datetime):
                return valor.time()
            # Si es string, intentar parsearlo
            if isinstance(valor, str):
                try:
                    # Intentar parsear como HH:MM o HH:MM:SS
                    parsed = pd.to_datetime(valor, format='%H:%M', errors='coerce')
                    if pd.isna(parsed):
                        parsed = pd.to_datetime(valor, format='%H:%M:%SS', errors='coerce')
                    if not pd.isna(parsed):
                        return parsed.time()
                except:
                    pass
            return None
        
        # Convertir horas usando la función robusta
        df["HRA INGRESO"] = df["HRA INGRESO"].apply(convertir_a_time)
        df["HORA SALIDA"] = df["HORA SALIDA"].apply(convertir_a_time)
        
        # Verificar que las conversiones fueron exitosas
        if df["HRA INGRESO"].isna().any() or df["HORA SALIDA"].isna().any():
            st.error("⚠️ Error: Algunas horas de entrada o salida no tienen el formato correcto")
            st.info("Las horas pueden estar en formato Excel (tiempo) o texto HH:MM (ej: 08:00, 14:00, 18:30)")
            st.stop()
    except Exception as e:
        st.error(f"⚠️ Error al procesar horas: {str(e)}")
        st.info("Verifica el formato de las columnas HRA INGRESO y HORA SALIDA")
        st.stop()

    def convertir_a_datetime(fecha, hora):
        return pd.to_datetime(fecha.astype(str) + ' ' + hora.astype(str))

    df["DT_INGRESO"] = convertir_a_datetime(df["FECHA"], df["HRA INGRESO"])
    df["DT_SALIDA"] = convertir_a_datetime(df["FECHA"], df["HORA SALIDA"])

    # Crear diccionario de configuración de turnos
    turnos_config = {}
    for _, row in df_turnos.iterrows():
        turno_nombre = str(row["TURNO"]).upper().strip()
        
        # Convertir horas de manera segura (acepta time, datetime, o string)
        def safe_time_convert(val):
            if pd.isna(val):
                return None
            if isinstance(val, time):
                return val
            if isinstance(val, datetime):
                return val.time()
            if isinstance(val, str):
                try:
                    return datetime.strptime(val, '%H:%M').time()
                except:
                    try:
                        return datetime.strptime(val, '%H:%M:%S').time()
                    except:
                        return None
            return None
        
        entrada = safe_time_convert(row["ENTRADA"])
        salida = safe_time_convert(row["SALIDA"])
        
        if entrada and salida:
            turnos_config[turno_nombre] = {
                "entrada": entrada,
                "salida": salida
            }

    # Agregar info del turno al dataframe
    def asignar_turno_info(row):
        turno = str(row.get("TURNO", "")).upper().strip()
        config = turnos_config.get(turno, {})
        return pd.Series({
            "TURNO ENTRADA": config.get("entrada"),
            "TURNO SALIDA": config.get("salida")
        })

    df[["TURNO ENTRADA", "TURNO SALIDA"]] = df.apply(asignar_turno_info, axis=1)

    # Crear diccionario de porcentajes (factores)
    porcentajes = {}
    for _, row in df_porcentaje.iterrows():
        tipo_key = str(row.get("TIPO", "")).strip().upper()
        if tipo_key:
            porcentajes[tipo_key] = {
                "ORDINARIA": row.get("% ORDINARIA", 0) / 100 if pd.notna(row.get("% ORDINARIA")) else 0,
                "EXTRA DIURNA": row.get("% EXTRA DIURNA", 0) / 100 if pd.notna(row.get("% EXTRA DIURNA")) else 0,
                "EXTRA NOCTURNA": row.get("% EXTRA NOCTURNA", 0) / 100 if pd.notna(row.get("% EXTRA NOCTURNA")) else 0,
                "RECARGO NOCTURNO": row.get("% RECARGO NOCTURNO", 0) / 100 if pd.notna(row.get("% RECARGO NOCTURNO")) else 0,
                "RECARGO DOMINICAL": row.get("% RECARGO DOMINICAL", 0) / 100 if pd.notna(row.get("% RECARGO DOMINICAL")) else 0
            }

    # Calcular valor hora ordinaria
    df["Valor Ordinario Hora"] = (df["SALARIO BASICO"] + df["COMISIÓN O BONIFICACIÓN"]) / 240

    # Función principal de cálculo de horas extras
    def calcular_horas_extras_y_recargo(row):
        try:
            ingreso = row["DT_INGRESO"]
            salida = row["DT_SALIDA"]
            turno_entrada = row["TURNO ENTRADA"]
            turno_salida = row["TURNO SALIDA"]
            dia_num = row["DIA_NUM"]
            
            # Si la hora real de salida es menor que la de entrada, pasó a otro día
            if salida.time() < ingreso.time():
                salida += timedelta(days=1)
            
            # Convertir turno a datetime para comparar
            turno_entrada_dt = datetime.combine(ingreso.date(), turno_entrada) if turno_entrada else None
            turno_salida_dt = datetime.combine(ingreso.date(), turno_salida) if turno_salida else None
            
            if turno_salida_dt and turno_entrada_dt and turno_salida_dt < turno_entrada_dt:
                turno_salida_dt += timedelta(days=1)
            
            # Horas trabajadas total
            horas_trabajadas = (salida - ingreso).total_seconds() / 3600
            
            # Definir hora nocturna (22:00 a 06:00)
            medianoche = datetime.combine(row["FECHA"], time(0, 0)) + timedelta(days=1)
            inicio_nocturno = datetime.combine(row["FECHA"], time(22, 0))
            fin_nocturno = medianoche.replace(hour=6, minute=0, second=0)
            
            # Determinar tipo de día
            tipo_dia = "DOMINICAL/FESTIVO" if dia_num == 6 else "LABORAL"
            factor = porcentajes.get(tipo_dia, porcentajes.get("LABORAL", {}))
            
            # Inicializar
            horas_extra_diurna = 0
            horas_extra_nocturna = 0
            horas_recargo_nocturno = 0
            
            # Calcular horas extra si hay turno definido
            if turno_salida_dt and turno_entrada_dt:
                # Si trabajó más del turno
                if salida > turno_salida_dt:
                    tiempo_extra = salida - turno_salida_dt
                    
                    # Dividir en diurna y nocturna
                    inicio_extra = turno_salida_dt
                    fin_extra = salida
                    
                    # Iterar por cada hora trabajada extra
                    current = inicio_extra
                    while current < fin_extra:
                        next_hour = min(current + timedelta(hours=1), fin_extra)
                        horas = (next_hour - current).total_seconds() / 3600
                        
                        # Verificar si está en horario nocturno
                        es_nocturno = (current.time() >= time(22, 0) or current.time() < time(6, 0))
                        
                        if es_nocturno:
                            horas_extra_nocturna += horas
                        else:
                            horas_extra_diurna += horas
                        
                        current = next_hour
                
                # Calcular recargo nocturno dentro del turno normal
                inicio_trabajo = max(ingreso, turno_entrada_dt)
                fin_trabajo = min(salida, turno_salida_dt)
                
                current = inicio_trabajo
                while current < fin_trabajo:
                    next_hour = min(current + timedelta(hours=1), fin_trabajo)
                    horas = (next_hour - current).total_seconds() / 3600
                    
                    # Verificar si está en horario nocturno
                    es_nocturno = (current.time() >= time(22, 0) or current.time() < time(6, 0))
                    
                    if es_nocturno:
                        horas_recargo_nocturno += horas
                    
                    current = next_hour
            
            valor_hora = row["Valor Ordinario Hora"]
            
            return pd.Series({
                "HORAS TRABAJADAS": horas_trabajadas,
                "Cant. HORAS EXTRA DIURNA": horas_extra_diurna,
                "VALOR EXTRA DIURNA": horas_extra_diurna * valor_hora * (1 + factor.get("EXTRA DIURNA", 0)),
                "Cant. HORAS EXTRA NOCTURNA": horas_extra_nocturna,
                "VALOR EXTRA NOCTURNA": horas_extra_nocturna * valor_hora * (1 + factor.get("EXTRA NOCTURNA", 0)),
                "Cant. RECARGO NOCTURNO": horas_recargo_nocturno,
                "VALOR RECARGO NOCTURNO": horas_recargo_nocturno * valor_hora * factor.get("RECARGO NOCTURNO", 0),
                "TOTAL HORAS EXTRA": horas_extra_diurna + horas_extra_nocturna + horas_recargo_nocturno
            })
        except Exception as e:
            st.error(f"Error en fila: {e}")
            return pd.Series({
                "HORAS TRABAJADAS": 0,
                "Cant. HORAS EXTRA DIURNA": 0,
                "VALOR EXTRA DIURNA": 0,
                "Cant. HORAS EXTRA NOCTURNA": 0,
                "VALOR EXTRA NOCTURNA": 0,
                "Cant. RECARGO NOCTURNO": 0,
                "VALOR RECARGO NOCTURNO": 0,
                "TOTAL HORAS EXTRA": 0
            })

    # Aplicar cálculos
    resultados = df.apply(calcular_horas_extras_y_recargo, axis=1)
    df = pd.concat([df, resultados], axis=1)

    # Calcular valor total a pagar
    df["VALOR TOTAL A PAGAR"] = (
        df["VALOR EXTRA DIURNA"] + 
        df["VALOR EXTRA NOCTURNA"] + 
        df["VALOR RECARGO NOCTURNO"]
    )

    # Agregar columnas de mes
    df["MES"] = df["FECHA"].dt.month
    df["MES_NOMBRE"] = df["FECHA"].dt.strftime('%B')

    # Calcular TOTAL BASE LIQUIDACION
    df["TOTAL BASE LIQUIDACION"] = df["SALARIO BASICO"] + df["COMISIÓN O BONIFICACIÓN"]

    # Extraer horas reales como time objects para el Excel
    df["hora_real_INGRESO"] = df["DT_INGRESO"].dt.time
    df["hora_real_SALIDA"] = df["DT_SALIDA"].dt.time

    # Renombrar Observacion si es necesario
    if "OBSERVACIONES" in df.columns:
        df["Observacion"] = df["OBSERVACIONES"]
    else:
        df["Observacion"] = ""

    # Mostrar resultados
    st.success("✅ Cálculo completado exitosamente")
    
    # Resumen
    col_res1, col_res2, col_res3 = st.columns(3)
    with col_res1:
        st.metric("Total Empleados", df[cedula_input].nunique())
    with col_res2:
        st.metric("Total Horas Extras", f"{df['TOTAL HORAS EXTRA'].sum():.2f}")
    with col_res3:
        st.metric("Total a Pagar", f"${df['VALOR TOTAL A PAGAR'].sum():,.2f}")

    # Vista previa
    st.subheader("📋 Vista Previa de Resultados")
    st.dataframe(df.head(10))

    # Gráficos
    st.subheader("📊 Análisis Visual")
    
    tab1, tab2, tab3, tab4 = st.tabs(["👥 Por Empleado", "🏢 Por Área", "📅 Horas Mensuales", "💰 Costos Mensuales"])
    
    with tab1:
        # Horas por empleado
        fig1, ax1 = plt.subplots(figsize=(10, 6))
        por_empleado = df.groupby("NOMBRE")["TOTAL HORAS EXTRA"].sum().sort_values(ascending=False).head(10)
        por_empleado.plot(kind='barh', ax=ax1, color='#f37021')
        ax1.set_xlabel('Horas Extras')
        ax1.set_title('Top 10 Empleados por Horas Extras')
        plt.tight_layout()
        st.pyplot(fig1)
        
        # Botón de descarga
        buf1 = BytesIO()
        fig1.savefig(buf1, format='png', dpi=300, bbox_inches='tight')
        buf1.seek(0)
        st.download_button("📥 Descargar Gráfico", data=buf1, file_name="empleados.png", mime="image/png")
    
    with tab2:
        # Por área
        fig2, ax2 = plt.subplots(figsize=(10, 6))
        por_area = df.groupby("AREA")["VALOR TOTAL A PAGAR"].sum().sort_values(ascending=False)
        por_area.plot(kind='bar', ax=ax2, color='#f37021')
        ax2.set_ylabel('Valor Total ($)')
        ax2.set_title('Costos por Área')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        st.pyplot(fig2)
        
        buf2 = BytesIO()
        fig2.savefig(buf2, format='png', dpi=300, bbox_inches='tight')
        buf2.seek(0)
        st.download_button("📥 Descargar Gráfico", data=buf2, file_name="areas.png", mime="image/png")
    
    with tab3:
        # Horas mensuales
        fig3, ax3 = plt.subplots(figsize=(10, 6))
        por_mes_horas = df.groupby("MES_NOMBRE")["TOTAL HORAS EXTRA"].sum()
        por_mes_horas.plot(kind='line', ax=ax3, marker='o', color='#f37021', linewidth=2)
        ax3.set_ylabel('Horas Extras')
        ax3.set_title('Horas Extras por Mes')
        ax3.grid(True, alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig3)
        
        buf3 = BytesIO()
        fig3.savefig(buf3, format='png', dpi=300, bbox_inches='tight')
        buf3.seek(0)
        st.download_button("📥 Descargar Gráfico", data=buf3, file_name="horas_mensuales.png", mime="image/png")
    
    with tab4:
        # Costos mensuales
        fig4, ax4 = plt.subplots(figsize=(10, 6))
        por_mes_valor = df.groupby("MES_NOMBRE")["VALOR TOTAL A PAGAR"].sum()
        por_mes_valor.plot(kind='bar', ax=ax4, color='#f37021')
        ax4.set_ylabel('Valor Total ($)')
        ax4.set_title('Costos por Mes')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        st.pyplot(fig4)
        
        buf4 = BytesIO()
        fig4.savefig(buf4, format='png', dpi=300, bbox_inches='tight')
        buf4.seek(0)
        st.download_button("📥 Descargar Gráfico", data=buf4, file_name="costos_mensuales.png", mime="image/png")

    # Descarga de Excel
    st.subheader("📥 Descargar Resultados")
    
    col_desc1, col_desc2 = st.columns([2, 1])
    
    with col_desc1:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        
        # Fecha para nombre archivo
        fecha_actual = datetime.now().strftime("%Y%m%d")
        output_filename = f"HorasExtras_Fertrac_{fecha_actual}.xlsx"
        
        # Encabezados
        headers = [
            "CÉDULA", "NOMBRE", "CARGO", "AREA", "SALARIO BASICO",
            "COMISIÓN O BONIFICACIÓN", "TOTAL BASE LIQUIDACION", "IMPORTE HORA",
            "FECHA", "TURNO", "TURNO ENTRADA", "TURNO SALIDA",
            "HRA INGRESO", "HORA SALIDA", "ACTIVIDAD DESARROLLADA",
            "HORAS TRABAJADAS", "HORAS EXTRA DIURNA", "VALOR EXTRA DIURNA",
            "HORAS EXTRA NOCTURNA", "VALOR EXTRA NOCTURNA",
            "RECARGO NOCTURNO", "VALOR RECARGO NOCTURNO",
            "TOTAL HORAS EXTRA", "VALOR TOTAL A PAGAR",
            "MES", "MES_NOMBRE", "OBSERVACIONES"
        ]
        
        ws.append(headers)
        
        # Formatear encabezados
        header_font = Font(bold=True, size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        
        # Mapeo de columnas
        column_mapping = {
            "CÉDULA": cedula_input,
            "NOMBRE": "NOMBRE",
            "CARGO": "CARGO",
            "AREA": "AREA",
            "SALARIO BASICO": "SALARIO BASICO",
            "COMISIÓN O BONIFICACIÓN": "COMISIÓN O BONIFICACIÓN",
            "TOTAL BASE LIQUIDACION": "TOTAL BASE LIQUIDACION",
            "IMPORTE HORA": "Valor Ordinario Hora",
            "FECHA": "FECHA",
            "TURNO": "TURNO",
            "TURNO ENTRADA": "TURNO ENTRADA",
            "TURNO SALIDA": "TURNO SALIDA",
            "HRA INGRESO": "hora_real_INGRESO",
            "HORA SALIDA": "hora_real_SALIDA",
            "ACTIVIDAD DESARROLLADA": "ACTIVIDAD DESARROLLADA",
            "HORAS TRABAJADAS": "HORAS TRABAJADAS",
            "HORAS EXTRA DIURNA": "Cant. HORAS EXTRA DIURNA",
            "VALOR EXTRA DIURNA": "VALOR EXTRA DIURNA",
            "HORAS EXTRA NOCTURNA": "Cant. HORAS EXTRA NOCTURNA",
            "VALOR EXTRA NOCTURNA": "VALOR EXTRA NOCTURNA",
            "RECARGO NOCTURNO": "Cant. RECARGO NOCTURNO",
            "VALOR RECARGO NOCTURNO": "VALOR RECARGO NOCTURNO",
            "TOTAL HORAS EXTRA": "TOTAL HORAS EXTRA",
            "VALOR TOTAL A PAGAR": "VALOR TOTAL A PAGAR",
            "MES": "MES",
            "MES_NOMBRE": "MES_NOMBRE",
            "OBSERVACIONES": "Observacion"
        }
        
        # Escribir datos
        for idx, row in df.iterrows():
            row_data = []
            for header in headers:
                df_col = column_mapping.get(header, header)
                
                if df_col in df.columns:
                    value = row[df_col]
                    
                    if pd.isna(value):
                        row_data.append("")
                    elif isinstance(value, (time, datetime)):
                        if isinstance(value, datetime):
                            row_data.append(value.strftime('%Y-%m-%d %H:%M'))
                        else:
                            row_data.append(value.strftime('%H:%M'))
                    elif isinstance(value, (np.int64, np.int32, np.int16, np.int8)):
                        row_data.append(int(value))
                    elif isinstance(value, (np.float64, np.float32, np.float16)):
                        row_data.append(float(value))
                    elif isinstance(value, (pd.Period,)):
                        row_data.append(str(value))
                    elif hasattr(value, 'item'):
                        row_data.append(value.item())
                    else:
                        try:
                            if isinstance(value, (int, float, str, bool)):
                                row_data.append(value)
                            else:
                                row_data.append(str(value))
                        except:
                            row_data.append(str(value))
                else:
                    row_data.append("")
            
            ws.append(row_data)
        
        # Formato celdas
        for row_idx in range(2, ws.max_row + 1):
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
                header_name = headers[col_idx - 1]
                if "VALOR" in header_name or "SALARIO" in header_name or "TOTAL BASE" in header_name or "IMPORTE" in header_name:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'
                elif "HORAS" in header_name:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'
        
        # Anchos de columna
        column_widths = {
            'A': 12, 'B': 25, 'C': 20, 'D': 15, 'E': 15,
            'F': 18, 'G': 18, 'H': 15, 'I': 12, 'J': 22,
            'K': 12, 'L': 12, 'M': 12, 'N': 12, 'O': 30,
            'P': 12, 'Q': 12, 'R': 15, 'S': 12, 'T': 15,
            'U': 12, 'V': 15, 'W': 12, 'X': 18, 'Y': 10,
            'Z': 15, 'AA': 30
        }
        
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # Guardar
        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL (.XLSX) ⬇️",
            data=excel_buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel",
            use_container_width=True
        )
        
        st.warning("""
        ⚠️ **IMPORTANTE:** Si el archivo se descarga como `.csv`:
        
        **Solución rápida:**
        1. Descarga el archivo (aunque sea .csv)
        2. Renómbralo: cambia `.csv` por `.xlsx`
        3. Ábrelo en Excel - funcionará perfectamente
        
        El contenido **SÍ es Excel válido**, solo la extensión está incorrecta.
        """)
        
        st.info("""
        💡 **¿Por qué pasa esto?**
        
        Es un problema conocido de algunos navegadores (Chrome/Edge en Windows) con Streamlit.
        El archivo generado es 100% Excel, pero el navegador le pone extensión .csv al descargar.
        """)
    
    with col_desc2:
        st.markdown("### 📊 Todas las gráficas ya tienen su botón")
        st.markdown("""
        Cada gráfico tiene su propio botón de descarga:
        - Gráfico de empleados
        - Gráfico de áreas  
        - Gráfico de horas mensuales
        - Gráfico de costos mensuales
        """)
