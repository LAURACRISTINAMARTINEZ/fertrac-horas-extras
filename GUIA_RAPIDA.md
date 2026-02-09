# 📖 Guía Rápida de Uso - Calculadora de Horas Extras

## 🎯 Inicio Rápido (5 minutos)

### 1️⃣ Preparar tus archivos Excel

Necesitas 4 archivos:

#### ✅ input_datos.xlsx
```
Columnas mínimas:
- CÉDULA
- FECHA
- HRA INGRESO
- HORA SALIDA
- TURNO
- ACTIVIDAD DESARROLLADA
- COMISIÓN O BONIFICACIÓN
```

#### ✅ base_empleados.xlsx
```
Columnas mínimas:
- CEDULA
- NOMBRE
- AREA
- SALARIO BASICO
- CARGO (opcional)
```

#### ✅ factores_horas_extras.xlsx
```
TIPO HORA EXTRA    | FACTOR
EXTRA DIURNA       | 1.25
EXTRA NOCTURNA     | 1.75
RECARGO NOCTURNO   | 1.35
```

#### ✅ configuracion_turnos.xlsx
```
TURNO   | HORA ENTRADA | HORA SALIDA
TURNO 1 | 08:45        | 18:10
TURNO 2 | 07:00        | 17:00
```

---

### 2️⃣ Subir archivos a la aplicación

1. Abre la aplicación en tu navegador
2. Arrastra o selecciona cada archivo en su casilla correspondiente
3. Espera a que se validen automáticamente

---

### 3️⃣ Revisar resultados

La aplicación mostrará:
- ✅ Configuración aplicada
- ✅ Tabla con todos los cálculos
- ✅ Gráficos por empleado y área
- ✅ Resumen mensual

---

### 4️⃣ Descargar Excel

1. Desplázate hasta "Descargar resultados"
2. Haz clic en "DESCARGAR EXCEL (.XLSX)"
3. Si se descarga como .csv, solo renómbralo a .xlsx

---

## ⚠️ Errores Comunes y Soluciones

### ❌ "No se encontró columna de cédula"
**Solución**: Verifica que tu archivo tenga una columna llamada "CÉDULA" o "CEDULA"

### ❌ "Algunas fechas no tienen el formato correcto"
**Solución**: Las fechas deben estar en formato:
- YYYY-MM-DD (2026-01-15)
- DD/MM/YYYY (15/01/2026)

### ❌ "Algunas horas no tienen el formato correcto"
**Solución**: Las horas deben estar en formato HH:MM (08:45, 19:30)

### ❌ "X registros no tienen información de empleado"
**Solución**: Las cédulas en input_datos.xlsx deben existir en base_empleados.xlsx

### ❌ "El archivo se descarga como .csv"
**Solución**: 
1. Descarga el archivo aunque sea .csv
2. Renómbralo y cambia la extensión a .xlsx
3. Ábrelo en Excel normalmente

---

## 💡 Consejos Útiles

### ✨ Para mejores resultados:
- ✅ Elimina espacios extra en las cédulas
- ✅ Usa formatos de celda apropiados en Excel (fecha para fechas, hora para horas)
- ✅ Verifica que todas las cédulas existan en ambos archivos
- ✅ Asegúrate de que los nombres de turnos coincidan exactamente

### 📊 Para analizar datos:
- Usa los filtros por área y mes
- Descarga los gráficos individuales
- Revisa los totales en el resumen general

### 💾 Para exportar:
- El Excel incluye TODOS los datos calculados
- Los gráficos se descargan en formato PNG de alta calidad
- Puedes importar el Excel directamente a tu sistema de nómina

---

## 🔢 Entendiendo los Resultados

### Tipos de Horas

**HORAS EXTRA DIURNA** (6 AM - 7 PM)
- Llegaste antes del turno (en horario diurno)
- Saliste después del turno (en horario diurno)
- Se paga con factor 1.25

**HORAS EXTRA NOCTURNA** (7 PM - 6 AM)
- Llegaste antes del turno (en horario nocturno)
- Saliste después del turno (en horario nocturno)
- Se paga con factor 1.75

**RECARGO NOCTURNO** (7 PM - 6 AM)
- Horas NORMALES del turno que caen en horario nocturno
- NO son horas extra, son parte de la jornada ordinaria
- Se paga con factor 1.35

### Ejemplo Visual

```
Turno configurado: 8:45 AM - 6:10 PM
Horario real: 7:30 AM - 8:00 PM

├─ 7:30 - 8:45 (1.25h) → EXTRA DIURNA
├─ 8:45 - 6:10 (8.42h) → HORAS NORMALES
│  └─ Se descuenta 1h almuerzo = 7.42h trabajadas
└─ 6:10 - 7:00 (0.83h) → EXTRA DIURNA
   7:00 - 8:00 (1.00h) → EXTRA NOCTURNA
```

---

## 📞 ¿Necesitas Ayuda?

### Antes de reportar un problema:
1. ✅ Verifica que tus archivos tengan las columnas correctas
2. ✅ Lee los mensajes de error que muestra la aplicación
3. ✅ Consulta esta guía rápida
4. ✅ Revisa el README.md completo para más detalles

### Para reportar:
- Describe exactamente qué estabas haciendo
- Incluye el mensaje de error completo
- Si es posible, adjunta ejemplos (sin datos sensibles)

---

## 🎓 Aprende Más

Para entender la lógica completa de cálculo, consulta:
- **README.md**: Documentación técnica completa
- **Sección "Lógica de Cálculo"**: Ejemplos detallados paso a paso

---

**Última actualización**: Febrero 2026  
**Versión**: 2.0
