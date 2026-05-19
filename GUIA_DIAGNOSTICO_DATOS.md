# 🔍 GUÍA PASO A PASO — Diagnosticar y Corregir Datos Incompletos

## 🎯 Tu Problema Actual

- ✗ Nombres llegan como IDs de Kobo (`WOTR202626` en lugar del nombre real)
- ✗ Campos vacíos: Teléfono, Zona, Email, Educación, etc.
- ✗ No sabes qué campos están mal mapeados

---

## 📋 DIAGNÓSTICO (3 pasos fáciles)

### PASO 1: Ver qué datos realmente tiene Kobo

**En tu Sheet:**
1. Click en menú **📊 PASO A PASO**
2. Click en **🔍 Ver Primer Registro Kobo**

**Resultado que verás:**
```
📋 PRIMER REGISTRO KOBO — QUÉ DATOS TIENES
═════════════════════════════════════════════

✅ CAMPOS CON DATOS (12):
  • nombre_completo_del_la_participante
    → Juan Carlos Pérez López
  • numero_de_dpi_opcional
    → 1234567890123
  • edad
    → 28
  • ... más campos

❌ CAMPOS VACÍOS (8):
  • correo_electronico_opcional
  • lugar_de_residencia
  • ... más campos
```

**¿Qué significa?** Los campos con ✅ tienen datos. Los que están vacíos no.

---

### PASO 2: Comparar mapeos (si código vs realidad)

**En tu Sheet:**
1. Click en menú **📊 PASO A PASO**
2. Click en **⚖️ Comparar Mapeos**

**Resultado que verás:**
```
⚖️ COMPARACIÓN: CÓDIGO vs TU KOBO
═════════════════════════════════════════════

✅ Nombre            OK
✅ DPI               OK
✅ Edad              OK
✅ Teléfono          OK
❌ Email             NO EXISTE
❌ Educación         NO EXISTE
❌ Laboral           NO EXISTE

✅ 4 correctos  |  ❌ 3 faltantes

💡 SOLUCIÓN:
1. Haz click en "📋 Exportar Muestra Kobo"
2. Mira la nueva hoja "INSPECCION_KOBO"
3. Copia los nombres EXACTOS de los campos
4. Actualiza Paso_a_Paso.gs línea ~590
```

---

### PASO 3: Ver datos reales en tu Sheet

**En tu Sheet:**
1. Click en menú **📊 PASO A PASO**
2. Click en **📋 Exportar Muestra Kobo**

**Resultado:** Se crea nueva hoja **INSPECCION_KOBO** con 10 registros reales de Kobo.

Ahora puedes:
- Ver exactamente qué campos tiene tu formulario
- Ver qué datos llegaron
- Copiar nombres EXACTOS de campos

---

## 🔧 SOLUCIONAR PROBLEMA

### Si campos están vacíos o mal mapeados:

**Ejemplo:** "Email" llega vacío en el sistema, pero en tu Kobo se llama "correo_electronico"

#### Opción 1: Ver en la hoja INSPECCION_KOBO

1. Haz los 3 pasos de diagnóstico arriba
2. Abre la hoja **INSPECCION_KOBO**
3. Busca la columna que tiene los emails reales
4. Copia el nombre EXACTO de esa columna

#### Opción 2: Actualizar el script

Si encontraste que un campo se llama diferente:

1. Abre **Paso_a_Paso.gs** en el Editor de Scripts
2. Busca la línea ~590 donde dice `fila[M.EMAIL-1] = r['correo_electronico_opcional']`
3. Reemplaza `'correo_electronico_opcional'` con el nombre EXACTO que encontraste

**Ejemplo:**
```javascript
// ANTES (no funciona):
fila[M.EMAIL-1] = r['correo_electronico_opcional'] || '';

// DESPUÉS (funciona):
fila[M.EMAIL-1] = r['email_participante'] || '';
```

4. Click en **Guardar**
5. Vuelve a tu Sheet
6. Click en **📊 PASO A PASO → 🔄✅ Sincronizar Kobo (con validaciones)**

---

## 📊 CAMPOS MAPEADOS ACTUALMENTE

Estos son los campos que el código busca en Kobo. Si tu Kobo tiene nombres diferentes, debes actualizar:

```javascript
// Línea ~590 en Paso_a_Paso.gs

fila[M.NOMBRE-1]     = r['nombre_completo_del_la_participante']    // Nombre
fila[M.DPI-1]        = r['numero_de_dpi_opcional']                 // DPI
fila[M.EDAD-1]       = r['edad']                                   // Edad
fila[M.GENERO-1]     = r['genero']                                 // Género
fila[M.TELEFONO-1]   = r['numero_de_telefono']                     // Teléfono
fila[M.ZONA-1]       = r['lugar_de_residencia']                    // Zona
fila[M.EMAIL-1]      = r['correo_electronico_opcional']            // Email
fila[M.EDUCACION-1]  = r['cual_es_el_ultimo_grado_que_completaste']// Educación
fila[M.LABORAL-1]    = r['cual_es_tu_situacion_laboral_actual']    // Laboral
fila[M.FORTALEZAS-1] = r['que_sabes_hacer_bien']                   // Fortalezas
fila[M.OBJETIVO-1]   = r['que_tipo_de_empleo_estas_buscando_especificamente'] // Objetivo
fila[M.PERFIL-1]     = r['perfil_asignado']                        // Perfil
fila[M.PRIORIDAD-1]  = r['prioridad_caso']                         // Prioridad
fila[M.PUNTAJE-1]    = r['puntaje_total_60']                       // Puntaje
fila[M.DIM1-1]       = r['dimension_1_capital_educativo']          // Dim1
fila[M.DIM2-1]       = r['dimension_2_capital_laboral']            // Dim2
fila[M.DIM3-1]       = r['dimension_3_habilidades_digitales']      // Dim3
fila[M.DIM4-1]       = r['dimension_4_claridad_vocacional']        // Dim4
fila[M.DIM5-1]       = r['dimension_5_barreras_estructurales']     // Dim5
fila[M.DIM6-1]       = r['dimension_6_red_apoyo']                  // Dim6
```

---

## ✅ VERIFICAR QUE FUNCIONÓ

Después de actualizar mapeos:

1. **Click:** 📊 PASO A PASO → 🔄✅ Sincronizar Kobo (con validaciones)
2. **Click:** 📊 PASO A PASO → 📊 Analizar Calidad de Datos

Deberías ver:
```
✅ CAMPOS COMPLETADOS:
  Nombre ............ 100% (50/50) ✅
  DPI ............... 100% (50/50) ✅
  Teléfono .......... 100% (50/50) ✅
  Email ............ 95% (47/50) ✅
```

Si algunos todavía están en 0%, repite el diagnóstico.

---

## 🆘 CASOS ESPECIALES

### P: "Mi Kobo tiene campos diferentes a los que busca el código"

**Solución:**
1. Abre la hoja **INSPECCION_KOBO** (si no existe, click en 📋 Exportar Muestra Kobo)
2. Busca la columna que corresponde a cada campo del sistema
3. Copia el nombre EXACTO
4. Actualiza Paso_a_Paso.gs línea ~590

### P: "¿Cómo sé cuál columna en INSPECCION_KOBO es cuál?"

Mira el primer registro (fila 2) y busca:
- Nombre real → busca el que tiene un nombre completo
- DPI → 13 números
- Edad → número entre 15-70
- Teléfono → número que empieza con +502 o 2-7
- Email → contiene @

### P: "Algunas filas vienen vacías, ¿qué hago?"

El sistema now rechaza participantes incompletos. Esto es **BUENO** porque:
- ✅ No crea registros "fantasma" con datos vacíos
- ✅ Te notifica en el Log cuáles fueron rechazados
- ✅ Mantiene datos limpios

Para agregarloss:
1. Completa los datos en Kobo
2. Vuelve a sincronizar

---

## 📋 CHECKLIST

Antes de sincronizar, verifica:

- [ ] He hecho click en 🔍 Ver Primer Registro Kobo
- [ ] He hecho click en ⚖️ Comparar Mapeos
- [ ] He visto la hoja INSPECCION_KOBO
- [ ] He identificado campos mal mapeados
- [ ] He actualizado Paso_a_Paso.gs si necesario
- [ ] He guardado el script
- [ ] Ahora voy a hacer click en 🔄✅ Sincronizar Kobo (con validaciones)
- [ ] Voy a revisar los datos en el Sheet principal

---

## 📞 RESUMEN RÁPIDO

```
¿Datos incompletos?
    ↓
Click: 🔍 Ver Primer Registro Kobo
    ↓
¿Campos vacíos?
    ↓
Click: ⚖️ Comparar Mapeos
    ↓
¿Mapeos incorrectos?
    ↓
Click: 📋 Exportar Muestra Kobo
    ↓
Identifica campos en INSPECCION_KOBO
    ↓
Actualiza Paso_a_Paso.gs línea ~590
    ↓
Click: 🔄✅ Sincronizar Kobo (con validaciones)
    ↓
Click: 📊 Analizar Calidad de Datos
    ↓
✅ Datos ahora están limpios y completos
```

---

**Versión:** 1.0  
**Fecha:** 2026-05-19  
**Sistema:** Paso a Paso v8.7+ con Herramientas de Diagnóstico
