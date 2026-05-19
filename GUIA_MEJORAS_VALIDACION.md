# 🚀 GUÍA DE MEJORAS: Validación y Diagnóstico de Datos

**Versión:** 1.0  
**Fecha:** 2026-05-19  
**Objetivo:** Resolver problemas de datos incompletos y campos vacíos en sincronización Kobo

---

## 📋 ¿Qué Mejoramos?

### ✅ Problema 1: Datos llegan incompletos
**Antes:** No sabías cuáles campos se perdían en la sincronización  
**Ahora:** Sistema valida automáticamente que tengas los datos críticos (Nombre, DPI, Teléfono)

### ✅ Problema 2: Campos no mapean bien
**Antes:** Nombres hardcodeados de campos pueden no coincidir con tu Kobo  
**Ahora:** Función nueva inspecciona exactamente qué campos tiene tu formulario

### ✅ Problema 3: Sin diagnóstico
**Antes:** Debías revisar manualmente si había problemas  
**Ahora:** Panel automático de calidad de datos te muestra % de completitud

---

## 🎯 3 NUEVAS FUNCIONES

### 1️⃣ **Inspeccionar Campos Kobo**
**Menú:** 📊 PASO A PASO → 🔬 Inspeccionar Campos Kobo

**¿Qué hace?**
- Lee tu formulario Kobo en vivo
- Lista TODOS los campos disponibles
- Muestra cuáles están mapeados ✅ y cuáles faltan ❌
- Incluye sugerencias de mapeo

**Cuándo usarlo:**
- ✅ Cuando datos llegan vacíos
- ✅ Cuando agregues nuevos campos al formulario
- ✅ Para confirmar que Kobo está bien configurado

**Resultado:**
```
📊 CAMPOS DISPONIBLES EN KOBO
================================

🔵 CAMPOS CRÍTICOS (requeridos):
  NOMBRE            → nombre_completo_del_la_participante ✅
  DPI               → numero_de_dpi_opcional ✅
  TELEFONO          → numero_de_telefono ✅
  EDAD              → edad ✅
  GENERO            → genero ✅
  EMAIL             → correo_electronico_opcional ✅

🟢 TODOS LOS CAMPOS (37 disponibles):
  • _uuid
  • _submission_time
  • nombre_completo_del_la_participante
  • numero_de_dpi_opcional
  ... (más campos)
```

---

### 2️⃣ **Analizar Calidad de Datos**
**Menú:** 📊 PASO A PASO → 📊 Analizar Calidad de Datos

**¿Qué hace?**
- Escanea TODOS tus participantes actuales
- Calcula % de completitud por campo
- Muestra registros incompletos
- Genera recomendaciones

**Cuándo usarlo:**
- ✅ Cada semana (auditoría)
- ✅ Después de una sincronización grande
- ✅ Cuando sospechas datos incompletos

**Resultado:**
```
📊 DIAGNÓSTICO DE CALIDAD DE DATOS
═════════════════════════════════════════

Total de participantes: 245

✅ CAMPOS COMPLETADOS:
  Nombre ............ 98.4% (241/245)
  DPI ............... 95.1% (233/245)
  Teléfono .......... 96.7% (237/245)
  Email ............ 78.4% (192/245)

🟢 REGISTROS COMPLETOS (datos críticos): 93.9% (230/245)

🔴 REGISTROS INCOMPLETOS (15):
  Fila 23: Juan Pérez → Falta: DPI
  Fila 45: María López → Falta: Teléfono
  Fila 67: Carlos → Falta: Nombre, Teléfono
  ... (12 más)

💡 RECOMENDACIONES:
  • Completar DPI faltantes (crítico para identificación)
  • Agregar teléfono para contacto
```

---

### 3️⃣ **Sincronizar Kobo con Validaciones**
**Menú:** 📊 PASO A PASO → 🔄✅ Sincronizar Kobo (con validaciones)

**¿Qué hace?**
- Descarga datos de Kobo (como antes)
- **NUEVO:** Valida ANTES de agregar
- Rechaza participantes con datos incompletos
- Registra por qué fueron rechazados
- Muestra reporte detallado

**Cuándo usarlo:**
- ✅ En lugar del botón normal "Sincronizar Kobo"
- ✅ Cuando quieras evitar datos "fantasma"
- ✅ Para mantener calidad de datos

**Resultado:**
```
✅ Sincronización completada
📥 42 nuevos participantes agregados
📊 87 total disponible en Kobo

🔴 3 rechazados (datos incompletos)
Revisa la hoja "Log" para detalles.
```

**Reporte en la hoja "Log":**
```
Timestamp | Tipo | Detalles | Usuario
----------|------|----------|--------
2026-05-19 14:30 | ⚠️ VALIDACIÓN | [IDER123] ❌ FALTA: Teléfono | usuario@email.com
2026-05-19 14:30 | ⚠️ VALIDACIÓN | [IDER124] ❌ FALTA: DPI | usuario@email.com
```

---

## 🔧 CAMPOS VALIDADOS AUTOMÁTICAMENTE

### Obligatorios (bloquea agregar si faltan)
- ✅ **Nombre Completo** → `nombre_completo_del_la_participante`
- ✅ **DPI** → `numero_de_dpi_opcional`
- ✅ **Teléfono** → `numero_de_telefono`

### Importantes (advertencia si faltan)
- ⚠️ **Edad** → `edad`
- ⚠️ **Email** → `correo_electronico_opcional`
- ⚠️ **Educación** → `cual_es_el_ultimo_grado_que_completaste`

---

## 📊 REPORTE EN HOJA "LOG"

El sistema ahora registra automáticamente:
- Cada validación exitosa
- Cada rechazo con motivo
- Diagnósticos de campos
- Errores de sincronización

**Acceder:** Haz click en la pestaña "Log" en tu Spreadsheet

---

## 💡 TROUBLESHOOTING

### P: "¿Qué hago si campos llegan vacíos?"
**R:** 
1. Click en 🔬 **Inspeccionar Campos Kobo**
2. Compara los nombres con tu formulario Kobo
3. Si no coinciden, edita los nombres en el código Paso_a_Paso.gs (líneas 576-603)
4. Intenta sincronizar de nuevo

### P: "¿Puedo ver qué participantes tienen datos incompletos?"
**R:** Sí, usa 📊 **Analizar Calidad de Datos** — te muestra exactamente cuáles filas tienen qué falta.

### P: "¿Puedo hacer que sea obligatorio llenar más campos?"
**R:** Sí, edita la función `validarDatosRequeridos()` en MEJORAS_VALIDACION_DATOS.gs y agrega más campos obligatorios.

### P: "¿Puedo cambiar los nombres de campos de Kobo?"
**R:** Sí, en la función `sincronizarConValidaciones()` (línea 280+), modifica los nombres exactos según tu formulario Kobo.

---

## 📈 FLUJO RECOMENDADO

### PRIMERA VEZ (Setup)
1. 🔬 Inspeccionar Campos Kobo → verifica que todos coincidan
2. 🔄✅ Sincronizar Kobo (con validaciones) → primera carga
3. 📊 Analizar Calidad de Datos → diagnóstico inicial

### RUTINA SEMANAL
1. 📊 Analizar Calidad de Datos → revisar % completitud
2. 🔄✅ Sincronizar Kobo (con validaciones) → si hay nuevos
3. Revisar hoja "Log" → ver si hay rechazos

### SI HAY PROBLEMAS
1. 🔬 Inspeccionar Campos Kobo → confirmar mapeo
2. ⚙️ Configuración → ajustar nombres si necesario
3. 🔄✅ Sincronizar Kobo (con validaciones) → reintentar

---

## 🎓 EJEMPLO PRÁCTICO

### Escenario: Te llegan 50 participantes nuevos pero datos vacíos

**Paso 1:** Click en 🔬 Inspeccionar Campos Kobo
```
Resultado: EMAIL muestra "❌ NO ENCONTRADO"
Razón: En tu Kobo el campo se llama "email_participante" no "correo_electronico_opcional"
```

**Paso 2:** Edita Paso_a_Paso.gs línea 590
```javascript
// Antes:
fila[M.EMAIL-1] = r['correo_electronico_opcional'] || '';

// Después:
fila[M.EMAIL-1] = r['email_participante'] || '';
```

**Paso 3:** Click en 🔄✅ Sincronizar Kobo (con validaciones)
```
Resultado: ✅ Todos los 50 se agregan correctamente con emails
```

**Paso 4:** Click en 📊 Analizar Calidad de Datos
```
Resultado: Email ............ 100% (50/50) ✅
```

---

## 📞 ¿PREGUNTAS?

Para más detalles sobre el sistema:
- 📖 **SISTEMA_PASO_A_PASO.md** — Arquitectura completa
- 🚀 **GUIA_INSTALACION_V3.md** — Instalación paso a paso
- 💡 **IDEAS_MEJORAS.md** — Próximas features

---

**Última actualización:** 2026-05-19  
**Versión sistema:** v8.7+ con mejoras validación
