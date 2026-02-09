[README.md](https://github.com/user-attachments/files/25186877/README.md)
# 🕒 Calculadora de Horas Extras - Fertrac

Sistema automatizado para el cálculo de horas extras, recargos nocturnos y liquidación de nómina para empleados de Fertrac.

## 📋 Tabla de Contenidos

- [Descripción General](#descripción-general)
- [Características Principales](#características-principales)
- [Requisitos del Sistema](#requisitos-del-sistema)
- [Instalación](#instalación)
- [Estructura de Archivos de Entrada](#estructura-de-archivos-de-entrada)
- [Lógica de Cálculo](#lógica-de-cálculo)
- [Uso de la Aplicación](#uso-de-la-aplicación)
- [Archivo de Salida](#archivo-de-salida)
- [Preguntas Frecuentes](#preguntas-frecuentes)

---

## 📖 Descripción General

Esta aplicación web desarrollada en Streamlit permite calcular automáticamente:
- Horas extras diurnas y nocturnas
- Recargos nocturnos
- Valores monetarios de cada tipo de hora
- Liquidación total por empleado

El sistema está diseñado específicamente para cumplir con la normativa laboral colombiana y los requerimientos internos de Fertrac.

---

## ✨ Características Principales

### 📊 Cálculos Automatizados
- **Precisión decimal completa**: Los cálculos mantienen precisión matemática hasta el último paso
- **Tres tipos de horas**: Extra Diurna, Extra Nocturna y Recargo Nocturno
- **Turnos configurables**: Soporte para múltiples turnos de trabajo
- **Base de liquidación correcta**: Usa Salario Básico + Comisiones/Bonificaciones

### 📈 Visualizaciones
- Gráficos por empleado
- Gráficos por área
- Comparativos mensuales
- Resúmenes estadísticos

### 💾 Exportación
- Archivo Excel con resultados completos
- Gráficos descargables en PNG
- Formato profesional listo para nómina

---

## 🖥️ Requisitos del Sistema

### Software Necesario
```
Python 3.8+
streamlit >= 1.28.0
pandas >= 2.0.0
numpy >= 1.24.0
openpyxl >= 3.1.0
matplotlib >= 3.7.0
```

### Archivos Requeridos
La aplicación necesita 4 archivos Excel de entrada:
1. **input_datos.xlsx** - Datos diarios de asistencia
2. **base_empleados.xlsx** - Información de empleados
3. **factores_horas_extras.xlsx** - Factores multiplicadores
4. **configuracion_turnos.xlsx** - Definición de turnos

---

## 📂 Estructura de Archivos de Entrada

### 1. input_datos.xlsx
Contiene los registros diarios de asistencia de cada empleado.

**Columnas obligatorias:**
| Columna | Tipo | Descripción |
|---------|------|-------------|
| CÉDULA o CEDULA | Texto | Número de identificación del empleado |
| FECHA | Fecha | Fecha del registro (YYYY-MM-DD o DD/MM/YYYY) |
| HRA INGRESO | Hora | Hora de entrada real (HH:MM) |
| HORA SALIDA | Hora | Hora de salida real (HH:MM) |
| TURNO | Texto | Nombre del turno asignado |
| ACTIVIDAD DESARROLLADA | Texto | Descripción de la actividad |
| COMISIÓN O BONIFICACIÓN | Número | Valor de comisiones del día (puede ser 0) |

**Columnas opcionales:**
- OBSERVACIONES o OBSERVACION

**Ejemplo:**
```
CÉDULA      | FECHA      | HRA INGRESO | HORA SALIDA | TURNO   | ACTIVIDAD                    | COMISIÓN O BONIFICACIÓN
2964423     | 2026-01-15 | 08:45      | 19:30       | TURNO 1 | Alistamiento y cargue        | 0
78760275    | 2026-01-15 | 08:05      | 19:30       | TURNO 1 | Apoyo en bodega              | 50000
```

---

### 2. base_empleados.xlsx
Contiene la información base de cada empleado.

**Columnas obligatorias:**
| Columna | Tipo | Descripción |
|---------|------|-------------|
| CEDULA o CÉDULA | Texto | Número de identificación (debe coincidir con input_datos) |
| NOMBRE | Texto | Nombre completo del empleado |
| AREA | Texto | Área de trabajo (Bodega, Ventas, etc.) |
| SALARIO BASICO | Número | Salario base mensual |

**Columnas opcionales:**
| Columna | Tipo | Descripción |
|---------|------|-------------|
| CARGO o Cargo | Texto | Cargo del empleado (se incluye en reporte final) |

**Ejemplo:**
```
CEDULA   | NOMBRE                      | AREA    | CARGO                    | SALARIO BASICO
2964423  | CARLOS JULIO CANGREJO       | Bodega  | ASISTENTE DE BODEGA      | 3000000
78760275 | CIRO SAMITH SOTO PATERNINA  | Bodega  | AUXILIAR DE BODEGA       | 2400000
```

⚠️ **Importante**: Las cédulas en ambos archivos deben coincidir exactamente (sin espacios extra).

---

### 3. factores_horas_extras.xlsx
Define los multiplicadores para cada tipo de hora extra.

**Columnas obligatorias:**
| Columna | Tipo | Descripción |
|---------|------|-------------|
| TIPO HORA EXTRA | Texto | Tipo de hora (EXTRA DIURNA, EXTRA NOCTURNA, RECARGO NOCTURNO) |
| FACTOR | Número | Multiplicador a aplicar |

**Ejemplo:**
```
TIPO HORA EXTRA    | FACTOR
EXTRA DIURNA       | 1.25
EXTRA NOCTURNA     | 1.75
RECARGO NOCTURNO   | 1.35
```

**Factores comunes según normativa colombiana:**
- Extra Diurna: 1.25 (25% adicional)
- Extra Nocturna: 1.75 (75% adicional)
- Recargo Nocturno: 1.35 (35% adicional)

---

### 4. configuracion_turnos.xlsx
Define los horarios de cada turno de trabajo.

**Columnas obligatorias:**
| Columna | Tipo | Descripción |
|---------|------|-------------|
| TURNO | Texto | Nombre del turno |
| HORA ENTRADA | Hora | Hora de inicio del turno (HH:MM) |
| HORA SALIDA | Hora | Hora de fin del turno (HH:MM) |

**Ejemplo:**
```
TURNO      | HORA ENTRADA | HORA SALIDA
TURNO 1    | 08:45        | 18:10
TURNO 2    | 07:00        | 17:00
TURNO 3    | 06:00        | 14:00
```

⚠️ **Nota sobre sábados**: El sistema ajusta automáticamente los sábados para terminar a las 12:00 PM.

---

## 🧮 Lógica de Cálculo

### Conceptos Fundamentales

#### 1. Horarios
- **Horario Diurno**: 6:00 AM - 7:00 PM (06:00 - 19:00)
- **Horario Nocturno**: 7:00 PM - 6:00 AM (19:00 - 06:00)

#### 2. Tipos de Horas

**HORAS EXTRA DIURNAS**
- **Definición**: Horas trabajadas **fuera de la jornada laboral** en horario diurno
- **Cuándo se generan**:
  - Llegó antes de la hora de entrada del turno (en horario diurno)
  - Salió después de la hora de salida del turno (en horario diurno)
- **Factor**: 1.25 (configurable)

**HORAS EXTRA NOCTURNAS**
- **Definición**: Horas trabajadas **fuera de la jornada laboral** en horario nocturno
- **Cuándo se generan**:
  - Llegó antes de la hora de entrada del turno (en horario nocturno)
  - Salió después de la hora de salida del turno (en horario nocturno)
- **Factor**: 1.75 (configurable)

**RECARGO NOCTURNO**
- **Definición**: Horas trabajadas **dentro de la jornada laboral ordinaria** que caen en horario nocturno
- **Cuándo se genera**:
  - La jornada normal del turno incluye horas entre 7:00 PM y 6:00 AM
  - Ejemplo: Turno de 6:00 PM a 2:00 AM → las horas de 7:00 PM a 2:00 AM tienen recargo nocturno
- **Factor**: 1.35 (configurable)

---

### Proceso de Cálculo Paso a Paso

#### Paso 1: Determinar Jornada Laboral
```
Jornada del Turno:
- Hora Entrada: Según configuración del turno
- Hora Salida: Según configuración del turno
- Ajuste Sábados: Si es sábado, hora salida = 12:00 PM
```

#### Paso 2: Calcular Horas Trabajadas
```
Horas Trabajadas = (Hora Salida Real - Hora Ingreso Real)
- Lunes a Viernes: Se descuenta 1 hora de almuerzo
- Sábados y Domingos: No se descuenta almuerzo
```

#### Paso 3: Identificar Horas Extra
```
SI (Hora Ingreso Real < Hora Entrada Turno):
    Tiempo Antes = Hora Entrada Turno - Hora Ingreso Real
    
    SI (Hora está en rango nocturno 19:00-06:00):
        → Horas Extra Nocturnas += Tiempo Antes
    SI NO:
        → Horas Extra Diurnas += Tiempo Antes

SI (Hora Salida Real > Hora Salida Turno):
    Tiempo Después = Hora Salida Real - Hora Salida Turno
    
    Para cada segmento de tiempo:
        SI (Hora está en rango nocturno 19:00-06:00):
            → Horas Extra Nocturnas += Segmento
        SI NO:
            → Horas Extra Diurnas += Segmento
```

#### Paso 4: Identificar Recargo Nocturno
```
Jornada Efectiva = Intersección de:
    - Horas realmente trabajadas dentro del turno
    - Horario nocturno (19:00-06:00)

Recargo Nocturno = Horas de Jornada Efectiva en horario nocturno
```

#### Paso 5: Cálculo Monetario
```
Base de Liquidación = Salario Básico + Comisión o Bonificación

Valor Hora Ordinaria = Base de Liquidación / 220 horas

Valor Extra Diurna = Horas Extra Diurnas × Valor Hora × 1.25
Valor Extra Nocturna = Horas Extra Nocturnas × Valor Hora × 1.75
Valor Recargo Nocturno = Recargo Nocturno × Valor Hora × 1.35

Valor Total Extras = Suma de todos los valores anteriores
Valor Total a Pagar = Valor Total Extras + Comisión o Bonificación
```

---

### Ejemplos Prácticos

#### Ejemplo 1: Llegada Anticipada en Horario Diurno
```
Turno: 08:45 - 18:10
Ingreso Real: 07:30
Salida Real: 18:10

Cálculo:
- Llegó 1.25 horas antes (07:30 a 08:45)
- Horario: 07:30 está en horario diurno (después de 06:00)
- Resultado: 1.25 horas extra diurnas
```

#### Ejemplo 2: Salida Tarde en Horario Nocturno
```
Turno: 08:45 - 18:10
Ingreso Real: 08:45
Salida Real: 20:00

Cálculo:
- Salió 1.83 horas después (18:10 a 20:00)
- 18:10 a 19:00 = 0.83 horas (horario diurno)
- 19:00 a 20:00 = 1.00 hora (horario nocturno)
- Resultado: 0.83 horas extra diurnas + 1.00 hora extra nocturna
```

#### Ejemplo 3: Turno con Recargo Nocturno
```
Turno: 14:00 - 22:00
Ingreso Real: 14:00
Salida Real: 22:00

Cálculo de Recargo:
- Jornada del turno: 14:00 a 22:00 (8 horas)
- Horario nocturno empieza a las 19:00
- 19:00 a 22:00 = 3 horas dentro de jornada en horario nocturno
- Resultado: 3 horas de recargo nocturno (NO son extras, son parte de la jornada normal)
```

#### Ejemplo 4: Combinación de Extra y Recargo
```
Turno: 14:00 - 22:00
Ingreso Real: 13:00
Salida Real: 23:30

Cálculo:
1. Extra Diurna:
   - 13:00 a 14:00 = 1 hora (llegó antes, en horario diurno)
   
2. Recargo Nocturno:
   - 19:00 a 22:00 = 3 horas (jornada normal en horario nocturno)
   
3. Extra Nocturna:
   - 22:00 a 23:30 = 1.5 horas (salió tarde, en horario nocturno)

Resultado Final:
- 1 hora extra diurna
- 1.5 horas extra nocturna
- 3 horas recargo nocturno
```

---

### Precisión de Cálculos

⚠️ **MUY IMPORTANTE**: El sistema mantiene **precisión decimal completa** durante todos los cálculos.

**Ejemplo de la importancia de la precisión:**
```
❌ INCORRECTO (redondeando antes de calcular):
Horas Extra = 0.8333 → redondeo → 0.83
Valor = 0.83 × $20,000 × 1.25 = $20,750

✅ CORRECTO (manteniendo precisión):
Horas Extra = 0.8333333...
Valor = 0.8333333 × $20,000 × 1.25 = $20,833.33

Diferencia: $83.33 (por cada registro)
```

**Cómo lo hace el sistema:**
1. Calcula horas con precisión completa (ej: 0.8333333...)
2. Calcula valores monetarios con esa precisión
3. **Solo al final** redondea a 2 decimales para mostrar
4. En Excel exporta los valores exactos (sin pérdida de precisión)

---

## 🚀 Uso de la Aplicación

### Paso 1: Acceder a la Aplicación
Abre tu navegador y ve a la URL de la aplicación Streamlit.

### Paso 2: Cargar Archivos
1. **Archivo de datos**: Sube `input_datos.xlsx`
2. **Base de empleados**: Sube `base_empleados.xlsx`
3. **Factores de horas extras**: Sube `factores_horas_extras.xlsx`
4. **Configuración de turnos**: Sube `configuracion_turnos.xlsx`

### Paso 3: Revisión Automática
El sistema valida automáticamente:
- ✅ Existencia de columnas requeridas
- ✅ Formatos de fecha y hora
- ✅ Coincidencia de cédulas entre archivos
- ✅ Configuración de turnos

### Paso 4: Ver Resultados
- **Tabla principal**: Resultados detallados por registro
- **Filtros**: Por área y por mes
- **Gráficos**: 
  - Horas extra por empleado
  - Horas extra por área
  - Comparativo mensual de horas
  - Comparativo mensual de costos

### Paso 5: Descargar Resultados
- **Excel completo**: Botón "DESCARGAR EXCEL"
- **Gráficos individuales**: Cada gráfico tiene su botón de descarga

---

## 📊 Archivo de Salida

### Estructura del Excel Exportado

El archivo `resultado_pagos_YYYY-MM-DD.xlsx` contiene las siguientes columnas:

| # | Columna | Descripción |
|---|---------|-------------|
| 1 | CÉDULA | Identificación del empleado |
| 2 | NOMBRE | Nombre completo |
| 3 | CARGO | Cargo del empleado |
| 4 | AREA | Área de trabajo |
| 5 | SALARIO BASICO | Salario base mensual |
| 6 | COMISIÓN O BONIFICACIÓN | Comisiones/bonificaciones del día |
| 7 | TOTAL BASE LIQUIDACION | Salario Básico + Comisión |
| 8 | Valor Ordinario Hora | Valor de la hora ordinaria (Base/220) |
| 9 | FECHA | Fecha del registro |
| 10 | TURNO | Turno asignado |
| 11 | TURNO ENTRADA | Hora de inicio del turno |
| 12 | TURNO SALIDA | Hora de fin del turno |
| 13 | hora_real_INGRESO | Hora real de entrada |
| 14 | hora_real_SALIDA | Hora real de salida |
| 15 | ACTIVIDAD DESARROLLADA | Descripción de actividad |
| 16 | HORAS TRABAJADAS | Total de horas trabajadas |
| 17 | Cant. HORAS EXTRA DIURNA | Cantidad de horas extra diurnas |
| 18 | VALOR EXTRA DIURNA | Valor monetario extra diurna |
| 19 | Cant. HORAS EXTRA NOCTURNA | Cantidad de horas extra nocturnas |
| 20 | VALOR EXTRA NOCTURNA | Valor monetario extra nocturna |
| 21 | Cant. RECARGO NOCTURNO | Cantidad de horas con recargo |
| 22 | VALOR RECARGO NOCTURNO | Valor monetario recargo nocturno |
| 23 | TOTAL HORAS EXTRA | Suma de horas extra (diurna + nocturna) |
| 24 | VALOR TOTAL A PAGAR | Valor total del día |
| 25 | MES | Período (formato: 2026-01) |
| 26 | MES_NOMBRE | Nombre del mes (ej: Enero 2026) |
| 27 | Observacion | Observaciones adicionales |

### Formato del Archivo
- **Encabezados**: Negrita, centrados, con bordes
- **Valores monetarios**: Formato #,##0.00 (ej: $1,234.56)
- **Horas**: Formato 0.00 (ej: 1.25)
- **Fechas**: Formato YYYY-MM-DD HH:MM
- **Columnas ajustadas**: Anchos optimizados para lectura

---

## ❓ Preguntas Frecuentes

### ¿Qué pasa si un empleado no tiene comisión?
El valor de COMISIÓN O BONIFICACIÓN puede ser 0. Si la columna no existe en el archivo de entrada, se asume $0 automáticamente.

### ¿Cómo maneja el sistema los sábados?
Los sábados se ajustan automáticamente para terminar a las 12:00 PM, independientemente de la hora de salida configurada en el turno.

### ¿Puedo usar nombres de columnas en minúsculas?
Sí, el sistema normaliza automáticamente todos los nombres de columnas a MAYÚSCULAS, por lo que "cargo", "Cargo" y "CARGO" son equivalentes.

### ¿Qué pasa si hay empleados sin información en la base de empleados?
El sistema muestra una advertencia con las cédulas que no tienen coincidencia, pero continúa procesando los registros que sí tienen información.

### ¿Por qué el archivo se descarga como .csv en lugar de .xlsx?
Es un problema conocido de algunos navegadores (Chrome/Edge en Windows). El archivo **SÍ es un Excel válido**, solo necesitas renombrarlo de `.csv` a `.xlsx`.

### ¿Cómo sé si mis cálculos son correctos?
El sistema incluye validaciones automáticas y muestra:
- Total de horas por categoría
- Valores monetarios detallados
- Gráficos comparativos

Puedes verificar manualmente algunos registros usando las fórmulas descritas en la sección "Lógica de Cálculo".

### ¿Qué normativa laboral usa el sistema?
El sistema está configurado para:
- **División de horas mensuales**: 220 horas (vigente desde 2do semestre 2025 en Colombia)
- **Descuento de almuerzo**: 1 hora en días laborales (lunes a viernes)
- **Horario nocturno**: 7:00 PM a 6:00 AM (según Código Sustantivo del Trabajo de Colombia)

### ¿Puedo cambiar los factores de horas extras?
Sí, solo necesitas editar el archivo `factores_horas_extras.xlsx` con los nuevos multiplicadores y volver a cargar los archivos.

---

## 🔧 Configuración y Personalización

### Cambiar Horarios Nocturnos
Si necesitas cambiar el rango de horario nocturno (actualmente 19:00 - 06:00), debes modificar la función `es_horario_nocturno()` en el código fuente.

### Agregar Nuevos Turnos
Solo necesitas agregar nuevas filas en el archivo `configuracion_turnos.xlsx` con el nombre del turno y sus horarios.

### Modificar el Descuento de Almuerzo
El descuento actual es de 1 hora en días de lunes a viernes. Para cambiarlo, modifica la función `calcular_trabajo_real()` en el código.

---

## 📞 Soporte

Para reportar problemas o solicitar mejoras:
1. Verifica que tus archivos de entrada cumplan con la estructura requerida
2. Revisa los mensajes de error o advertencia que muestra la aplicación
3. Consulta esta documentación para entender la lógica de cálculo
4. Contacta al equipo de desarrollo con:
   - Descripción del problema
   - Archivos de ejemplo (sin datos sensibles)
   - Capturas de pantalla de los errores

---

## 📝 Notas de Versión

### Versión Actual: 2.0
- ✅ Cálculo con precisión decimal completa
- ✅ Base de liquidación correcta (Salario + Comisiones)
- ✅ Soporte para columna CARGO
- ✅ Nombres de columna flexibles (hora_real_INGRESO/SALIDA)
- ✅ Validación mejorada de archivos de entrada
- ✅ Exportación Excel profesional
- ✅ Gráficos descargables

### Mejoras Futuras Planificadas
- [ ] Soporte para múltiples períodos en un solo archivo
- [ ] Exportación PDF de reportes
- [ ] Dashboard interactivo con filtros avanzados
- [ ] Historial de procesamiento

---

## 👥 Créditos

**Desarrollado para**: Fertrac  
**Tecnologías**: Python, Streamlit, Pandas, Openpyxl, Matplotlib  
**Última actualización**: Febrero 2026

---

## 📄 Licencia

Este software es propiedad de Fertrac y es para uso interno exclusivo.

---

**¿Necesitas ayuda?** Consulta la sección de [Preguntas Frecuentes](#-preguntas-frecuentes) o contacta al equipo de TI.
