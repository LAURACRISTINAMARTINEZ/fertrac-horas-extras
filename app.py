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
    
    # Agregar columna OBSERVACIONES si no existe en el input
    if "OBSERVACIONES" not in df.columns and "OBSERVACION" not in df.columns:
        df["OBSERVACIONES"] = ""  # Columna vacía por defecto
    elif "OBSERVACION" in df.columns:
        df["OBSERVACIONES"] = df["OBSERVACION"]  # Renombrar si existe con otro nombre
    
    # Verificar que se hayan encontrado coincidencias
    empleados_sin_info = df[df["NOMBRE"].isna()]
    if len(empleados_sin_info) > 0:
        st.warning(f"⚠️ Advertencia: {len(empleados_sin_info)} registros no tienen información de empleado")
        st.warning(f"Cédulas sin coincidencia: {empleados_sin_info[cedula_input].unique().tolist()[:5]}")
        st.info("Verifica que las cédulas en ambos archivos coincidan exactamente")

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
                        parsed = pd.to_datetime(valor, format='%H:%M:%S', errors='coerce')
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
        try:
            # Función para convertir a time
            def a_time(valor):
                if isinstance(valor, time):
                    return valor
                if isinstance(valor, datetime):
                    return valor.time()
                if isinstance(valor, str):
                    parsed = pd.to_datetime(valor, format='%H:%M', errors='coerce')
                    if pd.isna(parsed):
                        parsed = pd.to_datetime(valor, format='%H:%M:%S', errors='coerce')
                    if not pd.isna(parsed):
                        return parsed.time()
                return None
            
            hora_entrada = a_time(row["HORA ENTRADA"])
            hora_salida = a_time(row["HORA SALIDA"])
            
            if hora_entrada is None or hora_salida is None:
                st.warning(f"⚠️ Advertencia: Turno '{turno_nombre}' tiene formato de hora inválido, se omitirá")
                continue
            
            turnos_config[turno_nombre] = {
                "entrada": hora_entrada,
                "salida": hora_salida
            }
        except Exception as e:
            st.error(f"Error procesando turno {turno_nombre}: {e}")
            continue
    
    # Verificar que se cargaron turnos
    if not turnos_config:
        st.error("⚠️ No se pudo cargar ningún turno del archivo de configuración")
        st.stop()
    
    # Mostrar configuración de turnos cargada
    st.success(f"""
    ✅ **Configuración de turnos cargada:**
    {chr(10).join([f"- {turno}: {config['entrada'].strftime('%H:%M')} a {config['salida'].strftime('%H:%M')}" 
                   for turno, config in turnos_config.items()])}
    """)

    # Definir horarios según turno (usando configuración)
    def obtener_horarios_turno(row):
        """Retorna las horas de entrada y salida según la configuración del turno"""
        turno = row.get("TURNO", "TURNO 1").upper().strip()
        dia = row["DIA_NUM"]
        
        # Si el turno existe en la configuración, usar esos horarios
        if turno in turnos_config:
            hora_entrada = datetime.combine(row["FECHA"], turnos_config[turno]["entrada"])
            hora_salida = datetime.combine(row["FECHA"], turnos_config[turno]["salida"])
        else:
            # Si no está configurado, usar valores por defecto de TURNO 1
            st.warning(f"⚠️ Turno '{turno}' no encontrado en configuración. Usando TURNO 1 por defecto.")
            hora_entrada = datetime.combine(row["FECHA"], time(8, 0))
            hora_salida = datetime.combine(row["FECHA"], time(18, 0))
        
        # Ajuste para sábados: jornada termina a las 12 PM
        if dia == 5:
            hora_salida = datetime.combine(row["FECHA"], time(12, 0))
        
        return hora_entrada, hora_salida

    def calcular_horas_extras_y_recargo(row):
        """
        Calcula horas extras diurnas, nocturnas y recargo nocturno por separado
        
        LÓGICA:
        - Horario nocturno: 7 PM (19:00) - 6 AM (06:00)
        - Horario diurno: 6 AM (06:00) - 7 PM (19:00)
        
        - Extra Diurna: Horas fuera de jornada en horario diurno
        - Extra Nocturna: Horas fuera de jornada en horario nocturno  
        - Recargo Nocturno: Horas DENTRO de jornada en horario nocturno
        """
        ingreso_real = row["DT_INGRESO"]
        salida_real = row["DT_SALIDA"]
        
        # Obtener horarios del turno
        hora_entrada_esperada, hora_salida_esperada = obtener_horarios_turno(row)
        
        # Límites de horario nocturno (7 PM - 6 AM)
        inicio_noche = datetime.combine(row["FECHA"], time(19, 0))  # 7 PM
        fin_noche = datetime.combine(row["FECHA"], time(6, 0))      # 6 AM del día siguiente
        
        # Si fin_noche es 6 AM, ajustar al día siguiente
        if fin_noche < inicio_noche:
            fin_noche = fin_noche.replace(day=fin_noche.day + 1)
        
        horas_extra_diurna = 0
        horas_extra_nocturna = 0
        recargo_nocturno = 0
        
        # ==== CALCULAR RECARGO NOCTURNO (horas ordinarias en horario nocturno) ====
        # Solo se paga si la jornada NORMAL cae en horario nocturno
        
        # Inicio de trabajo normal vs inicio noche
        inicio_trabajo = max(hora_entrada_esperada, ingreso_real)
        fin_trabajo = min(hora_salida_esperada, salida_real)
        
        # Calcular intersección de jornada normal con horario nocturno (7 PM - 6 AM)
        if inicio_trabajo < fin_trabajo:
            # Verificar si hay intersección con 7 PM - 11:59 PM del mismo día
            if fin_trabajo > inicio_noche:
                inicio_recargo = max(inicio_trabajo, inicio_noche)
                fin_recargo = fin_trabajo
                recargo_nocturno += (fin_recargo - inicio_recargo).total_seconds() / 3600
            
            # Verificar si hay intersección con 12 AM - 6 AM (si la jornada cruza medianoche)
            medianoche = datetime.combine(row["FECHA"], time(0, 0)).replace(day=row["FECHA"].day + 1)
            fin_noche_dia_sig = datetime.combine(row["FECHA"], time(6, 0)).replace(day=row["FECHA"].day + 1)
            
            if fin_trabajo > medianoche and inicio_trabajo < fin_noche_dia_sig:
                inicio_recargo = max(inicio_trabajo, medianoche)
                fin_recargo = min(fin_trabajo, fin_noche_dia_sig)
                if fin_recargo > inicio_recargo:
                    recargo_nocturno += (fin_recargo - inicio_recargo).total_seconds() / 3600
        
        # ==== CALCULAR HORAS EXTRA (fuera de jornada) ====
        
        # 1. LLEGÓ ANTES de la hora de entrada
        if ingreso_real < hora_entrada_esperada:
            tiempo_antes = (hora_entrada_esperada - ingreso_real).total_seconds() / 3600
            
            # Verificar si esas horas son diurnas o nocturnas
            # De ingreso_real hasta hora_entrada_esperada
            if ingreso_real >= inicio_noche or ingreso_real < fin_noche:
                # Empezó en horario nocturno
                # Calcular cuánto fue nocturno y cuánto diurno
                if hora_entrada_esperada <= fin_noche:
                    # Todo fue nocturno
                    horas_extra_nocturna += tiempo_antes
                else:
                    # Parte nocturno, parte diurno
                    if ingreso_real < fin_noche:
                        # Desde ingreso hasta 6 AM = nocturno
                        horas_extra_nocturna += (fin_noche - ingreso_real).total_seconds() / 3600
                        # Desde 6 AM hasta entrada = diurno
                        horas_extra_diurna += (hora_entrada_esperada - fin_noche).total_seconds() / 3600
                    else:
                        # Todo diurno
                        horas_extra_diurna += tiempo_antes
            else:
                # Empezó en horario diurno
                horas_extra_diurna += tiempo_antes
        
        # 2. SALIÓ DESPUÉS de la hora de salida
        if salida_real > hora_salida_esperada:
            tiempo_despues = (salida_real - hora_salida_esperada).total_seconds() / 3600
            
            # Verificar si esas horas son diurnas o nocturnas
            # De hora_salida_esperada hasta salida_real
            
            # Caso simple: verificar cada segmento
            tiempo_actual = hora_salida_esperada
            
            while tiempo_actual < salida_real:
                # Determinar si este momento es nocturno o diurno
                hora_del_dia = tiempo_actual.time()
                
                # Nocturno: 19:00 - 23:59 y 00:00 - 05:59
                if hora_del_dia >= time(19, 0) or hora_del_dia < time(6, 0):
                    # Es nocturno
                    # Encontrar hasta cuándo es nocturno
                    if hora_del_dia >= time(19, 0):
                        # Estamos entre 7 PM y medianoche
                        fin_segmento_noche = datetime.combine(tiempo_actual.date(), time(23, 59, 59))
                    else:
                        # Estamos entre medianoche y 6 AM
                        fin_segmento_noche = datetime.combine(tiempo_actual.date(), time(6, 0))
                    
                    fin_segmento = min(fin_segmento_noche, salida_real)
                    horas_extra_nocturna += (fin_segmento - tiempo_actual).total_seconds() / 3600
                    tiempo_actual = fin_segmento
                else:
                    # Es diurno (6 AM - 7 PM)
                    fin_segmento_dia = datetime.combine(tiempo_actual.date(), time(19, 0))
                    fin_segmento = min(fin_segmento_dia, salida_real)
                    horas_extra_diurna += (fin_segmento - tiempo_actual).total_seconds() / 3600
                    tiempo_actual = fin_segmento
                
                # Evitar loop infinito
                if tiempo_actual >= salida_real:
                    break
                if fin_segmento == tiempo_actual:
                    tiempo_actual = tiempo_actual.replace(second=tiempo_actual.second + 1)
        
        return max(horas_extra_diurna, 0), max(horas_extra_nocturna, 0), max(recargo_nocturno, 0)

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

    # Agregar columnas con los horarios del turno para visualización
    def obtener_horarios_turno_para_mostrar(row):
        """Retorna los horarios del turno como strings para mostrar en tabla"""
        turno = row.get("TURNO", "TURNO 1").upper().strip()
        dia = row["DIA_NUM"]
        
        if turno in turnos_config:
            hora_entrada = turnos_config[turno]["entrada"]
            hora_salida = turnos_config[turno]["salida"]
        else:
            hora_entrada = time(8, 0)
            hora_salida = time(18, 0)
        
        # Ajuste para sábados
        if dia == 5:
            hora_salida = time(12, 0)
        
        return hora_entrada.strftime('%H:%M'), hora_salida.strftime('%H:%M')
    
    # Aplicar para obtener horarios del turno
    horarios_turno = df.apply(obtener_horarios_turno_para_mostrar, axis=1)
    df["TURNO ENTRADA"] = [h[0] for h in horarios_turno]
    df["TURNO SALIDA"] = [h[1] for h in horarios_turno]
    
    # Aplicar cálculos
    df["HORAS TRABAJADAS"] = df.apply(calcular_trabajo_real, axis=1)
    
    # Calcular horas extras diurnas, nocturnas y recargo nocturno
    resultados = df.apply(calcular_horas_extras_y_recargo, axis=1)
    df["HORAS EXTRA DIURNA"] = [r[0] for r in resultados]
    df["HORAS EXTRA NOCTURNA"] = [r[1] for r in resultados]
    df["RECARGO NOCTURNO"] = [r[2] for r in resultados]
    
    # Total de horas extra (diurnas + nocturnas)
    df["TOTAL HORAS EXTRA"] = df["HORAS EXTRA DIURNA"] + df["HORAS EXTRA NOCTURNA"]

    # Asignar tipos de hora extra
    df["TIPO EXTRA"] = df.apply(
        lambda row: "Recargo Nocturno" if row["RECARGO NOCTURNO"] > 0 else 
                    ("Extra Nocturna" if row["HORAS EXTRA NOCTURNA"] > 0 else "Extra Diurna"),
        axis=1
    )

    # Mapear factores
    factor_map = dict(zip(df_porcentaje["TIPO HORA EXTRA"].str.upper(), df_porcentaje["FACTOR"]))
    
    # Calcular valores monetarios
    # IMPORTANTE: Dividir por 220 horas (aplicable desde 2do semestre 2025)
    df["IMPORTE HORA"] = df["SALARIO BASICO"] / 220
    
    # Calcular valor de cada tipo de hora
    df["VALOR EXTRA DIURNA"] = df.apply(
        lambda row: row["HORAS EXTRA DIURNA"] * row["IMPORTE HORA"] * factor_map.get("EXTRA DIURNA", 1.25),
        axis=1
    )
    
    df["VALOR EXTRA NOCTURNA"] = df.apply(
        lambda row: row["HORAS EXTRA NOCTURNA"] * row["IMPORTE HORA"] * factor_map.get("EXTRA NOCTURNA", 1.25),
        axis=1
    )
    
    df["VALOR RECARGO NOCTURNO"] = df.apply(
        lambda row: row["RECARGO NOCTURNO"] * row["IMPORTE HORA"] * factor_map.get("RECARGO NOCTURNO", 1.35),
        axis=1
    )
    
    # Valor total de extras
    df["VALOR EXTRA"] = df["VALOR EXTRA DIURNA"] + df["VALOR EXTRA NOCTURNA"] + df["VALOR RECARGO NOCTURNO"]
    df["VALOR TOTAL A PAGAR"] = df["VALOR EXTRA"] + df["COMISIÓN O BONIFICACIÓN"]

    # Redondear valores
    df["IMPORTE HORA"] = df["IMPORTE HORA"].round(2)
    df["VALOR EXTRA DIURNA"] = df["VALOR EXTRA DIURNA"].round(2)
    df["VALOR EXTRA NOCTURNA"] = df["VALOR EXTRA NOCTURNA"].round(2)
    df["VALOR RECARGO NOCTURNO"] = df["VALOR RECARGO NOCTURNO"].round(2)
    df["VALOR EXTRA"] = df["VALOR EXTRA"].round(2)
    df["VALOR TOTAL A PAGAR"] = df["VALOR TOTAL A PAGAR"].round(2)
    df["HORAS TRABAJADAS"] = df["HORAS TRABAJADAS"].round(2)
    df["HORAS EXTRA DIURNA"] = df["HORAS EXTRA DIURNA"].round(2)
    df["HORAS EXTRA NOCTURNA"] = df["HORAS EXTRA NOCTURNA"].round(2)
    df["RECARGO NOCTURNO"] = df["RECARGO NOCTURNO"].round(2)
    df["TOTAL HORAS EXTRA"] = df["TOTAL HORAS EXTRA"].round(2)

    # Resumen informativo
    st.info(f"""
    ✅ **Configuración aplicada:**
    - División por **220 horas** mensuales (vigente desde 2do semestre 2025)
    - Descuento de **1 hora de almuerzo** en días de lunes a viernes
    - **Horario nocturno:** 7:00 PM a 6:00 AM (recargo aplicable a jornada ordinaria en este horario)
    - **Horas extra diurnas:** Fuera de jornada en horario diurno (6:00 AM - 7:00 PM)
    - **Horas extra nocturnas:** Fuera de jornada en horario nocturno (7:00 PM - 6:00 AM)
    - **Turnos configurables** según archivo de turnos cargado
    """)

    # Agregar columna de mes y año para análisis temporal
    df["MES"] = df["FECHA"].dt.to_period('M')
    
    # Crear nombres de mes en ESPAÑOL
    meses_espanol = {
        'January': 'Enero',
        'February': 'Febrero',
        'March': 'Marzo',
        'April': 'Abril',
        'May': 'Mayo',
        'June': 'Junio',
        'July': 'Julio',
        'August': 'Agosto',
        'September': 'Septiembre',
        'October': 'Octubre',
        'November': 'Noviembre',
        'December': 'Diciembre'
    }
    
    # Primero crear en inglés y luego traducir
    df["MES_NOMBRE_TEMP"] = df["FECHA"].dt.strftime('%B %Y')
    
    # Traducir a español
    def traducir_mes(mes_nombre):
        for ingles, espanol in meses_espanol.items():
            if ingles in mes_nombre:
                return mes_nombre.replace(ingles, espanol)
        return mes_nombre
    
    df["MES_NOMBRE"] = df["MES_NOMBRE_TEMP"].apply(traducir_mes)
    df = df.drop("MES_NOMBRE_TEMP", axis=1)

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
    
    # Columnas a mostrar con desglose detallado de los 3 tipos
    columnas_mostrar = [
        "FECHA", "NOMBRE", "CÉDULA", "AREA", 
        "TURNO", "TURNO ENTRADA", "TURNO SALIDA",           # Horarios del turno
        "HRA INGRESO", "HORA SALIDA", "HORAS TRABAJADAS",
        "HORAS EXTRA DIURNA", "VALOR EXTRA DIURNA",         # Extra diurna
        "HORAS EXTRA NOCTURNA", "VALOR EXTRA NOCTURNA",     # Extra nocturna  
        "RECARGO NOCTURNO", "VALOR RECARGO NOCTURNO",       # Recargo nocturno
        "TOTAL HORAS EXTRA",                                 # Total de extras
        "ACTIVIDAD DESARROLLADA",
        "VALOR EXTRA",                                       # Total solo extras
        "COMISIÓN O BONIFICACIÓN", 
        "VALOR TOTAL A PAGAR"                                # Total con bonificación
    ]
    
    # Filtrar solo las columnas que existen
    columnas_disponibles = [col for col in columnas_mostrar if col in df_filtrado.columns]
    
    # Crear copia para formatear la visualización
    df_display = df_filtrado[columnas_disponibles].copy()
    
    # Renombrar columnas para mayor claridad
    renombrar = {
        "VALOR EXTRA": "VALOR TOTAL EXTRA"
        # VALOR TOTAL A PAGAR se queda con su nombre original
    }
    df_display = df_display.rename(columns=renombrar)
    
    # Formatear valores monetarios para mejor visualización
    columnas_dinero = ["VALOR EXTRA DIURNA", "VALOR EXTRA NOCTURNA", "VALOR RECARGO NOCTURNO", 
                       "VALOR TOTAL EXTRA", "COMISIÓN O BONIFICACIÓN", "VALOR TOTAL A PAGAR"]
    
    for col in columnas_dinero:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")
    
    # Mostrar tabla con formato mejorado
    st.dataframe(df_display, use_container_width=True)

    # Visualizaciones
    col_viz1, col_viz2 = st.columns(2)
    
    with col_viz1:
        st.subheader("👤 Horas extra por empleado")
        fig1, ax1 = plt.subplots(figsize=(10, 6))
        empleado_stats = df_filtrado.groupby("NOMBRE").agg({
            "HORAS EXTRA DIURNA": "sum",
            "HORAS EXTRA NOCTURNA": "sum",
            "RECARGO NOCTURNO": "sum"
        })
        if not empleado_stats.empty:
            empleado_stats.plot(kind='bar', ax=ax1, color=['#f37021', '#ff6600', '#ff9966'], stacked=True)
            ax1.set_ylabel("Horas")
            ax1.set_xlabel("Empleado")
            ax1.legend(["Extra Diurna", "Extra Nocturna", "Recargo Nocturno"])
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig1)
            
            # Botón para descargar gráfico
            buf1 = BytesIO()
            fig1.savefig(buf1, format='png', dpi=300, bbox_inches='tight')
            buf1.seek(0)
            st.download_button(
                label="📊 Descargar gráfico empleados",
                data=buf1,
                file_name=f"grafico_empleados_{date.today().isoformat()}.png",
                mime="image/png",
                key="download_empleados"
            )
        else:
            st.info("No hay datos para mostrar con los filtros seleccionados")

    with col_viz2:
        st.subheader("🏢 Horas extra por área")
        fig2, ax2 = plt.subplots(figsize=(10, 6))
        area_stats = df_filtrado.groupby("AREA").agg({
            "HORAS EXTRA DIURNA": "sum",
            "HORAS EXTRA NOCTURNA": "sum",
            "RECARGO NOCTURNO": "sum"
        })
        if not area_stats.empty:
            area_stats.plot(kind='barh', ax=ax2, color=['#f37021', '#ff6600', '#ff9966'], stacked=True)
            ax2.set_xlabel("Horas")
            ax2.set_ylabel("Área")
            ax2.legend(["Extra Diurna", "Extra Nocturna", "Recargo Nocturno"])
            plt.tight_layout()
            st.pyplot(fig2)
            
            # Botón para descargar gráfico
            buf2 = BytesIO()
            fig2.savefig(buf2, format='png', dpi=300, bbox_inches='tight')
            buf2.seek(0)
            st.download_button(
                label="📊 Descargar gráfico áreas",
                data=buf2,
                file_name=f"grafico_areas_{date.today().isoformat()}.png",
                mime="image/png",
                key="download_areas"
            )
        else:
            st.info("No hay datos para mostrar con los filtros seleccionados")

    # COMPARATIVO MENSUAL
    st.subheader("📅 Comparativo mensual")
    
    # Agrupar por mes
    comparativo_mensual = df.groupby("MES_NOMBRE").agg({
        "HORAS EXTRA DIURNA": "sum",
        "HORAS EXTRA NOCTURNA": "sum",
        "RECARGO NOCTURNO": "sum",
        "VALOR EXTRA": "sum",
        "VALOR TOTAL A PAGAR": "sum"
    }).reset_index()
    
    # Ordenar por fecha
    # Crear columna temporal con fecha para ordenar
    def extraer_fecha(mes_nombre):
        # De "Febrero 2025" extraer año y mes
        partes = mes_nombre.split()
        if len(partes) == 2:
            mes_texto, anio = partes
            meses = {
                'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
                'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
                'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
            }
            mes_num = meses.get(mes_texto, 1)
            return datetime(int(anio), mes_num, 1)
        return datetime(2025, 1, 1)
    
    comparativo_mensual["ORDEN"] = comparativo_mensual["MES_NOMBRE"].apply(extraer_fecha)
    comparativo_mensual = comparativo_mensual.sort_values("ORDEN")
    
    col_comp1, col_comp2 = st.columns(2)
    
    with col_comp1:
        st.markdown("#### Horas por mes")
        fig3, ax3 = plt.subplots(figsize=(10, 6))
        x = range(len(comparativo_mensual))
        width = 0.25
        
        ax3.bar([i - width for i in x], comparativo_mensual["HORAS EXTRA DIURNA"], 
                width, label='Extra Diurna', color='#f37021')
        ax3.bar(x, comparativo_mensual["HORAS EXTRA NOCTURNA"], 
                width, label='Extra Nocturna', color='#ff6600')
        ax3.bar([i + width for i in x], comparativo_mensual["RECARGO NOCTURNO"], 
                width, label='Recargo Nocturno', color='#ff9966')
        
        ax3.set_xlabel('Mes')
        ax3.set_ylabel('Horas')
        ax3.set_xticks(x)
        ax3.set_xticklabels(comparativo_mensual["MES_NOMBRE"], rotation=45, ha='right')
        ax3.legend()
        plt.tight_layout()
        st.pyplot(fig3)
        
        # Botón para descargar gráfico de horas
        buf3 = BytesIO()
        fig3.savefig(buf3, format='png', dpi=300, bbox_inches='tight')
        buf3.seek(0)
        st.download_button(
            label="📊 Descargar gráfico de horas",
            data=buf3,
            file_name=f"grafico_horas_mensual_{date.today().isoformat()}.png",
            mime="image/png"
        )
    
    with col_comp2:
        st.markdown("#### Costos por mes")
        fig4, ax4 = plt.subplots(figsize=(10, 6))
        
        # Gráfico de barras agrupadas en lugar de líneas
        x = range(len(comparativo_mensual))
        width = 0.35
        
        ax4.bar([i - width/2 for i in x], comparativo_mensual["VALOR EXTRA"], 
                width, label='Valor Extras', color='#f37021')
        ax4.bar([i + width/2 for i in x], comparativo_mensual["VALOR TOTAL A PAGAR"], 
                width, label='Total a Pagar', color='#ff9966')
        
        ax4.set_xlabel('Mes')
        ax4.set_ylabel('Valor ($)')
        ax4.set_xticks(x)
        ax4.set_xticklabels(comparativo_mensual["MES_NOMBRE"], rotation=45, ha='right')
        ax4.legend()
        plt.grid(True, alpha=0.3, axis='y')
        plt.tight_layout()
        st.pyplot(fig4)
        
        # Botón para descargar gráfico de costos
        buf4 = BytesIO()
        fig4.savefig(buf4, format='png', dpi=300, bbox_inches='tight')
        buf4.seek(0)
        st.download_button(
            label="📊 Descargar gráfico de costos",
            data=buf4,
            file_name=f"grafico_costos_mensual_{date.today().isoformat()}.png",
            mime="image/png"
        )
    
    # Tabla comparativa mensual
    st.markdown("#### Tabla comparativa mensual")
    tabla_comparativa = comparativo_mensual[["MES_NOMBRE", "HORAS EXTRA DIURNA", "HORAS EXTRA NOCTURNA",
                                             "RECARGO NOCTURNO", "VALOR EXTRA", "VALOR TOTAL A PAGAR"]].copy()
    tabla_comparativa.columns = ["Mes", "H. Extra Diurna", "H. Extra Nocturna", "H. Recargo Nocturno",
                                  "Valor Extras ($)", "Total a Pagar ($)"]
    
    # Formatear valores monetarios
    tabla_comparativa["Valor Extras ($)"] = tabla_comparativa["Valor Extras ($)"].apply(lambda x: f"${x:,.2f}")
    tabla_comparativa["Total a Pagar ($)"] = tabla_comparativa["Total a Pagar ($)"].apply(lambda x: f"${x:,.2f}")
    tabla_comparativa["H. Extra Diurna"] = tabla_comparativa["H. Extra Diurna"].apply(lambda x: f"{x:.2f}")
    tabla_comparativa["H. Extra Nocturna"] = tabla_comparativa["H. Extra Nocturna"].apply(lambda x: f"{x:.2f}")
    tabla_comparativa["H. Recargo Nocturno"] = tabla_comparativa["H. Recargo Nocturno"].apply(lambda x: f"{x:.2f}")
    
    st.dataframe(tabla_comparativa, use_container_width=True, hide_index=True)

    # Resumen estadístico (con datos filtrados)
    st.subheader("📈 Resumen general")
    col_stat1, col_stat2, col_stat3, col_stat4, col_stat5, col_stat6 = st.columns(6)
    
    with col_stat1:
        st.metric("H. Extra Diurnas", f"{df_filtrado['HORAS EXTRA DIURNA'].sum():.2f}")
    
    with col_stat2:
        st.metric("H. Extra Nocturnas", f"{df_filtrado['HORAS EXTRA NOCTURNA'].sum():.2f}")
    
    with col_stat3:
        st.metric("H. Recargo Nocturno", f"{df_filtrado['RECARGO NOCTURNO'].sum():.2f}")
    
    with col_stat4:
        st.metric("Total Extras ($)", f"${df_filtrado['VALOR EXTRA'].sum():,.2f}")
    
    with col_stat5:
        st.metric("Total Bonificaciones ($)", f"${df_filtrado['COMISIÓN O BONIFICACIÓN'].sum():,.2f}")
    
    with col_stat6:
        st.metric("Total General ($)", f"${df_filtrado['VALOR TOTAL A PAGAR'].sum():,.2f}")

    # Descarga de resultados
    st.subheader("💾 Descargar resultados")
    
    col_desc1, col_desc2 = st.columns(2)
    
    with col_desc1:
        st.markdown("### 📥 Descargar datos completos en Excel")
        st.markdown("Incluye todos los cálculos y resultados detallados")
        
        # Crear el archivo Excel en memoria usando pandas directamente
        hoy = date.today().isoformat()
        output_filename = f"resultado_pagos_{hoy}.xlsx"
        
        # Preparar DataFrame para exportar con columnas en el orden exacto solicitado
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        
        # Crear workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        
        # Definir encabezados en el orden EXACTO solicitado
        headers = [
            "CÉDULA", "NOMBRE", "AREA", "SALARIO BASICO", "COMISIÓN O BONIFICACIÓN",
            "TOTAL BASE LIQUIDACION", "Valor Ordinario Hora", "FECHA", "TURNO",
            "TURNO ENTRADA", "TURNO SALIDA", "DT_INGRESO", "DT_SALIDA",
            "ACTIVIDAD DESARROLLADA", "HORAS TRABAJADAS",
            "Cant. HORAS EXTRA DIURNA", "VALOR EXTRA DIURNA",
            "Cant. HORAS EXTRA NOCTURNA", "VALOR EXTRA NOCTURNA",
            "Cant. RECARGO NOCTURNO", "VALOR RECARGO NOCTURNO",
            "TOTAL HORAS EXTRA", "VALOR TOTAL A PAGAR",
            "MES", "MES_NOMBRE", "Observacion"
        ]
        
        ws.append(headers)
        
        # Formatear encabezados - BLANCO Y NEGRO (sin colores)
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
        
        # Mapeo de columnas del DataFrame a los encabezados
        column_mapping = {
            "CÉDULA": "CÉDULA",
            "NOMBRE": "NOMBRE",
            "AREA": "AREA",
            "SALARIO BASICO": "SALARIO BASICO",
            "COMISIÓN O BONIFICACIÓN": "COMISIÓN O BONIFICACIÓN",
            "TOTAL BASE LIQUIDACION": "VALOR TOTAL A PAGAR",  # Calculado
            "Valor Ordinario Hora": "IMPORTE HORA",
            "FECHA": "FECHA",
            "TURNO": "TURNO",
            "TURNO ENTRADA": "TURNO ENTRADA",
            "TURNO SALIDA": "TURNO SALIDA",
            "DT_INGRESO": "HRA INGRESO",
            "DT_SALIDA": "HORA SALIDA",
            "ACTIVIDAD DESARROLLADA": "ACTIVIDAD DESARROLLADA",
            "HORAS TRABAJADAS": "HORAS TRABAJADAS",
            "Cant. HORAS EXTRA DIURNA": "HORAS EXTRA DIURNA",
            "VALOR EXTRA DIURNA": "VALOR EXTRA DIURNA",
            "Cant. HORAS EXTRA NOCTURNA": "HORAS EXTRA NOCTURNA",
            "VALOR EXTRA NOCTURNA": "VALOR EXTRA NOCTURNA",
            "Cant. RECARGO NOCTURNO": "RECARGO NOCTURNO",
            "VALOR RECARGO NOCTURNO": "VALOR RECARGO NOCTURNO",
            "TOTAL HORAS EXTRA": "TOTAL HORAS EXTRA",
            "VALOR TOTAL A PAGAR": "VALOR TOTAL A PAGAR",
            "MES": "MES",
            "MES_NOMBRE": "MES_NOMBRE",
            "Observacion": "OBSERVACIONES"  # Nueva columna
        }
        
        # Escribir datos fila por fila
        for idx, row in df.iterrows():
            row_data = []
            for header in headers:
                df_col = column_mapping.get(header, header)
                
                if df_col in df.columns:
                    value = row[df_col]
                    
                    # Convertir valores según tipo a tipos básicos de Python
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
                    elif hasattr(value, 'item'):  # numpy types
                        row_data.append(value.item())
                    else:
                        # Convertir a string si es cualquier otro tipo
                        try:
                            # Intentar convertir a tipo básico
                            if isinstance(value, (int, float, str, bool)):
                                row_data.append(value)
                            else:
                                row_data.append(str(value))
                        except:
                            row_data.append(str(value))
                else:
                    # Columna no existe, poner vacío
                    row_data.append("")
            
            ws.append(row_data)
        
        # Aplicar formato a todas las celdas
        for row_idx in range(2, ws.max_row + 1):
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # Aplicar formato numérico según la columna
                header_name = headers[col_idx - 1]
                if "VALOR" in header_name or "SALARIO" in header_name or "TOTAL BASE" in header_name or "Valor Ordinario" in header_name:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'
                elif "Cant." in header_name or "HORAS" in header_name:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'
        
        # Ajustar anchos de columnas
        column_widths = {
            'A': 12,  # CÉDULA
            'B': 25,  # NOMBRE
            'C': 15,  # AREA
            'D': 15,  # SALARIO BASICO
            'E': 18,  # COMISIÓN
            'F': 18,  # TOTAL BASE
            'G': 15,  # Valor Hora
            'H': 12,  # FECHA
            'I': 22,  # TURNO
            'J': 12,  # TURNO ENTRADA
            'K': 12,  # TURNO SALIDA
            'L': 12,  # DT_INGRESO
            'M': 12,  # DT_SALIDA
            'N': 30,  # ACTIVIDAD
            'O': 12,  # HORAS TRAB
            'P': 12,  # Cant Extra Diurna
            'Q': 15,  # Valor Extra Diurna
            'R': 12,  # Cant Extra Nocturna
            'S': 15,  # Valor Extra Nocturna
            'T': 12,  # Cant Recargo
            'U': 15,  # Valor Recargo
            'V': 12,  # Total Horas
            'W': 18,  # Valor Total
            'X': 10,  # MES
            'Y': 15,  # MES_NOMBRE
            'Z': 30,  # Observacion
        }
        
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # Guardar en buffer
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
