# 🔍 AUDITORÍA DEL SISTEMA PASO A PASO
**Fecha:** 2026-05-19  
**Rama:** `claude/review-system-status-j0uHH`  
**Versión analizada:** v8.2

---

## ⚠️ PROBLEMAS CRÍTICOS ENCONTRADOS

### 1. **SECRETOS EXPUESTOS EN CÓDIGO** 🔴
**Severidad:** CRÍTICA  
**Ubicación:** `Paso_a_Paso.gs`, línea 9

```javascript
KOBO_API_KEY: "64cc018b88067397addd36b09288be8b6539cf39",
KOBO_ASSET_ID: "abHRWdRnPhKwzPQBajc7RZ",
```

**Riesgo:**
- API Key visible en repositorio público
- Cualquiera con acceso al código puede hacer requests a Kobo
- La clave ya está comprometida (está en GitHub)

**Solución:**
- Mover a `PropertiesService` (con interfaz de configuración)
- Implementar setter seguro de credenciales
- Generar nueva API Key en Kobo

---

### 2. **FALTA DE VALIDACIONES DE ESTADOS** 🟠
**Severidad:** ALTA  
**Ubicación:** `manejarEdicion()` (línea 463)

**Problema:** No hay validación de flujo de estado. Actualmente se puede:
- Pasar de "Orientación" directo a "Completado"
- Cambiar a estados inválidos
- No se valida que campos requeridos estén completados

**Impacto:** Datos inconsistentes en expedientes

**Solución requerida:**
```javascript
FLUJO_ESTADOS_PERMITIDOS = {
  'Orientación': ['Mentoría', 'Inactivo'],
  'Mentoría': ['Formación', 'Cierre', 'Derivación', 'Inactivo'],
  'Formación': ['Cierre', 'Derivación', 'Inactivo'],
  'Cierre': [], // Final
  'Derivación': [],
  'Inactivo': ['Orientación', 'Mentoría']
}
```

---

### 3. **CAMPOS REQUERIDOS NO VALIDADOS** 🟠
**Severidad:** ALTA  
**Ubicación:** `manejarEdicion()` (línea 463+)

**Problema:** Cuando cambias de estado, no se valida:
- Estado = "Mentoría" → Falta validar "Responsable_Actual"
- Estado = "Cierre" → Falta validar "Tipo_Cierre" y "Motivo"
- Estado = "Derivación" → Falta validar "Derivado_A"

**Impacto:** Mentoría sin responsable asignado, cierre sin motivo

---

### 4. **SINCRONIZACIÓN SHEETS→DOCS FRÁGIL** 🟡
**Severidad:** MEDIA  
**Ubicación:** `sincronizarConDocumento()` (línea 512)

```javascript
body.replaceText(campo.p, campo.pre + (valorNuevo || '—'));
```

**Problemas:**
- `replaceText()` reemplaza TODAS las instancias del patrón
- Si "Nombre:" aparece en HISTORIAL, lo reemplaza mal
- Sin manejo de excepciones si el doc fue eliminado
- No valida que el regex sea único

**Solución:** Usar `ElementType.PARAGRAPH` con búsqueda más específica

---

### 5. **FUNCIÓN `guardarDerivacion()` INCOMPLETA** 🟡
**Severidad:** MEDIA  
**Ubicación:** línea 630

```javascript
sheet.appendRow([
  new Date(),
  id,
  nombre,
  organizacion,  // ← Debería ser 'Tipo' según header de instalar()
  motivo,        // ← Debería ser 'Destino'
  resultado,     // ← Debería ser 'Motivo'
  '',            // ← Debería ser 'Estado'
  ''             // ← Debería ser 'Responsable'
]);
```

**Impacto:** Datos desalineados con estructura de hoja Derivaciones

---

### 6. **SIN MANEJO DE ERRORES EN TRIGGERS** 🟡
**Severidad:** MEDIA  
**Ubicación:** Múltiples funciones

```javascript
} catch(e) {
  // Error silencioso
}
```

**Problema:** 
- No hay logging de errores en triggers
- Imposible debuggear problemas de sincronización
- Los usuarios no saben si algo falló

**Solución:**
- Crear función `registrarError(nombreFuncion, error)`
- Guardar en hoja "Logs" con fecha/hora
- Email al admin si hay error crítico

---

### 7. **COLUMNAS DESINCRONIZADAS** 🟡
**Severidad:** MEDIA  
**Ubicaciones:** Múltiples

**Problema:**
- En `instalar()` se crean 26 columnas
- En `CONFIG.COL` hay referencias a 26 columnas
- Pero en `verFicha()` línea 776 se leen 26: `hoja.getRange(rango.getRow(), 1, 1, 26)`
- La estructura no está documentada, los nombres cambian entre funciones

**Impacto:** Si alguien añade columna, todo se rompe

---

### 8. **SIN VALIDACIÓN DE INPUT EN FORMULARIOS HTML** 🟡
**Severidad:** MEDIA  
**Ubicación:** `abrirFormDerivacion()` (línea 605)

```javascript
'<input id="motivo" placeholder="...">'
```

**Riesgo:**
- Sin sanitización de HTML
- Vulnerable a XSS si se vuelve a mostrar
- Sin validación de largo máximo
- Sin escape de caracteres especiales

---

### 9. **NO HAY PRUEBA DE CONEXIÓN A KOBO** 🟡
**Severidad:** MEDIA

**Problema:**
- La función `probarKobo()` existe en el menú pero no está implementada
- Sin forma de validar credenciales antes de instalar

---

### 10. **FALTA LOGGING CENTRALIZADO** 🟡
**Severidad:** MEDIA

**Problema:**
- Algunos cambios se guardan en hoja "Log"
- Otros se guardan en Google Docs
- Otros se pierden
- Sin consistencia en qué se registra

---

## ✅ COSAS QUE FUNCIONAN BIEN

1. ✅ Estructura general del sistema (Maestro + Derivaciones + Analytics + Log + Dashboard)
2. ✅ Integración con Kobo (importación de datos)
3. ✅ Creación automática de carpetas y expedientes en Drive
4. ✅ Triggers bien configurados (automáticos y instalables)
5. ✅ Sincronización unidireccional Sheets→Docs básica funciona
6. ✅ Dashboard automático se actualiza
7. ✅ Historial de cambios se registra
8. ✅ Interfaz UI con HTML Service es moderna
9. ✅ Colores por perfil funcionan
10. ✅ Emojis hacen la interfaz más amigable

---

## 📋 PLAN DE ACCIÓN PRIORIZADO

### **FASE 1: SEGURIDAD (1-2 días)**
- [ ] Mover API Key a PropertiesService
- [ ] Implementar formulario seguro de configuración
- [ ] Sanitizar inputs en formularios HTML
- [ ] Regenerar API Key en Kobo (la actual está comprometida)

### **FASE 2: VALIDACIONES (2-3 días)**
- [ ] Implementar validación de flujo de estados
- [ ] Añadir validación de campos requeridos por estado
- [ ] Crear función de validación centralizada
- [ ] Mejorar mensajes de error al usuario

### **FASE 3: MEJORAS TÉCNICAS (2-3 días)**
- [ ] Centralizar logging de errores
- [ ] Refactorizar `sincronizarConDocumento()` para mayor robustez
- [ ] Implementar `probarKobo()`
- [ ] Documentar estructura de columnas
- [ ] Completar función `guardarDerivacion()`

### **FASE 4: TESTING (1-2 días)**
- [ ] Crear función de auto-test
- [ ] Validar flujo completo (Kobo → Sheet → Doc)
- [ ] Probar todos los triggers
- [ ] Verificar sincronización en ambas direcciones

---

## 🎯 RECOMENDACIÓN INMEDIATA

**Comenzar por FASE 1 (Seguridad)** porque:
1. La API Key comprometida es riesgo inmediato
2. Tomar 1-2 horas
3. Es requisito para cualquier deployment

Después **FASE 2 (Validaciones)** porque:
1. Evita datos inconsistentes
2. Reduce errores en expedientes
3. Mejora confiabilidad del sistema

---

## 📊 MÉTRICAS DEL CÓDIGO ACTUAL

| Métrica | Valor | Interpretación |
|---------|-------|-----------------|
| Total de funciones | 40 | Bien documentado |
| Líneas de código | 1,592 | Manejable |
| Funciones con try-catch | 35/40 | 87.5% — Bueno |
| Funciones con validación | 12/40 | 30% — **Necesita mejorar** |
| Líneas de logging | ~20 | Muy bajo |
| Documentación | Buena | README completo |

---

## 🚀 PRÓXIMOS PASOS

1. Revisar y validar este análisis
2. Comenzar con FASE 1 (Seguridad)
3. Crear test suite
4. Preparar para productivo
5. Documentar cambios en CHANGELOG.md

**Estimado total:** 8-12 horas de trabajo

---

*Análisis completado con Claude en sesión de review de sistema*
