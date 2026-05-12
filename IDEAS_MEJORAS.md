# IDEAS Y MEJORAS — SISTEMA PASO A PASO

## 🏆 TIER 1 — Fácil + Alto Impacto

### ✅ 1. Sincronización Sheets → Docs (IMPLEMENTADO v7.1+)
**Qué hace:** Cuando cambias cualquier campo en una fila del Sheet, se actualiza automáticamente el Google Doc del participante.

**Campos sincronizados:**
- Nombre, DPI, Edad, Género, Teléfono, Zona
- Educación, Situación Laboral, Fortalezas, Objetivo Laboral
- Perfil, Prioridad, Estado

**Extras:**
- Si cambia el Nombre → también cambia el título del documento
- Cada cambio queda registrado en la sección "HISTORIAL DE CAMBIOS" del documento
- Todo queda en la hoja "Log" del Sheet

---

### 2. Validaciones inteligentes en cambios de estado
**Problema:** Hoy puedes cambiar de "Orientación" directamente a "Completado" sin pasos intermedios.

**Solución:** Reglas de flujo de trabajo:
```
Orientación → Mentoría ✅
Orientación → Completado ❌ (requiere pasar por Mentoría)
Mentoría → Formación ✅
Mentoría → Completado ✅
Formación → Completado ✅
Cualquier estado → Cierre ✅ (con motivo requerido)
```

**Cómo:** En `onEdit()`, si cambia la columna ESTADO, validar el flujo antes de permitir.

**Tiempo estimado:** 1.5 horas

---

### 3. Campos requeridos por estado
**Lógica:**
- Si Estado = "Mentoría" → requerir "Responsable_Actual"
- Si Estado = "Cierre" → requerir "Tipo_Cierre" y "Motivo_Cierre"
- Si Estado = "Completado" → requerir "Resultado_Final"

**Cómo:** En `onEdit()`, verificar campos relacionados.

**Tiempo estimado:** 1 hora

---

### 4. Semáforo visual automático
**Qué hace:** Columna adicional con emoji según urgencia calculada automáticamente:
```
🔴 Crítico: prioridad CRÍTICA + sin contacto > 15 días
🟠 Urgente: prioridad ALTA + sin contacto > 21 días
🟡 Atención: sin contacto > 30 días
🟢 Al día: contactado en los últimos 14 días
```

**Cómo:** `onEdit()` o trigger diario actualiza esta columna.

**Tiempo estimado:** 1 hora

---

## 🚀 TIER 2 — Medio + Muy Alto Impacto

### 5. Tabla de evolución en el documento
**Qué hace:** El documento lleva una tabla cronológica de todos los cambios:
```
HISTORIAL DE CAMBIOS
─────────────────────────────────────────────
12/05/2026 14:30 | Estado    | Orientación → Mentoría
11/05/2026 09:15 | Perfil    | Perfil C → Perfil B
10/05/2026 16:00 | Prioridad | BAJO → ALTO
```

**Cómo:** `sincronizarConDocumento()` encuentra la sección HISTORIAL y appends.

**Tiempo estimado:** 2 horas

---

### 6. Gráficos dinámicos en el documento
**Qué hace:** Regenerar una tabla visual de dimensiones en el Doc:
```
DIAGNÓSTICO
─────────────────────────────────────────────
Educativo:   ████████░░   7/10
Laboral:     ████░░░░░░   4/10
Digital:     ██████░░░░   6/10
Vocacional:  █████░░░░░   5/10
Barreras:    ███░░░░░░░   3/10
Red Apoyo:   ██████████  10/10
```

**Cómo:** Cuando cambia cualquier dimensión, regenerar esta sección en el Doc.

**Tiempo estimado:** 2 horas

---

### 7. Recordatorios en Google Calendar
**Qué hace:** Cuando cambia a "Mentoría" → crear evento de seguimiento en 7 días.
**Cuando cierra un caso** → crear evento de revisión de resultados en 30 días.

**Cómo:**
```javascript
CalendarApp.getDefaultCalendar().createEvent(
  'Seguimiento: ' + nombre,
  fechaSeguimiento,
  new Date(fechaSeguimiento.getTime() + 30 * 60 * 1000),
  { description: docUrl }
);
```

**Tiempo estimado:** 2 horas

---

### 8. Plan de Acción estructurado
**Qué hace:** Convertir "Plan de Acción" en un formulario interactivo dentro del sidebar:
```
+ Agregar Acción:
  [ Descripción       ]
  [ Responsable       ]
  [ Fecha límite      ]
  [💾 Agregar]
```
Se guarda en el Doc como tabla y en una nueva columna del Sheet.

**Tiempo estimado:** 3 horas

---

## 💡 TIER 3 — Avanzado + Premium

### 9. Exportar a PDF profesional
**Qué hace:** Botón "📄 Generar PDF" en el menú. Genera un PDF del perfil del participante, formateado y profesional, guardado en su carpeta Drive.

**Cómo:**
```javascript
const pdf = DriveApp.getFileById(docId).getAs('application/pdf');
carpeta.createFile(pdf).setName('Perfil_' + nombre + '_' + fecha + '.pdf');
```

**Tiempo estimado:** 2 horas

---

### 10. Sincronización bidireccional (Doc → Sheet)
**Qué hace:** Si el encargado escribe en el Doc (ej: nuevas notas de sesión), se captura en el Sheet automáticamente.

**Cómo:** Trigger `onDocEdit` no existe, pero se puede hacer con un trigger de tiempo que compare versiones del documento.

**Complejidad:** Alta — 4 horas

---

### 11. Notificación al participante por email
**Qué hace:** Cuando el estado cambia a "Mentoría" o "Formación", enviar email personalizado al participante:
```
Hola [Nombre],

Tienes una actualización en tu proceso con Creamos.
Tu estado ha avanzado a: Mentoría ✅

Tu mentor/a: [Responsable_Actual]
Próxima sesión: [Proxima_Sesion]

¡Seguimos contigo! 🌱
Equipo Creamos
```

**Tiempo estimado:** 2 horas (requiere columna EMAIL llenada)

---

### 12. Dashboard web como app independiente
**Qué hace:** Una URL web pública (dentro de la organización) que muestra el dashboard en tiempo real. No requiere acceso al Sheet.

**Cómo:**
```javascript
// Publicar como Web App
function doGet() {
  return HtmlService.createTemplateFromFile('dashboard').evaluate();
}
```

**Tiempo estimado:** 5-6 horas

---

### 13. Múltiples formularios Kobo
**Qué hace:** Soportar más de un formulario Kobo. Ej:
- Formulario de Registro → Sheet "Maestro"
- Formulario de Seguimiento mensual → actualizar filas existentes
- Encuesta de satisfacción → Sheet "Satisfacción"

**Tiempo estimado:** 3 horas

---

### 14. Re-diagnóstico y seguimiento de progreso
**Qué hace:** Si el mismo participante llena el formulario Kobo por segunda vez, comparar puntajes:
```
PROGRESO DEL PARTICIPANTE
─────────────────────────
Dimensión    | Antes | Ahora | Cambio
Educativo:   |   4   |   7   |  ↑ +3
Laboral:     |   3   |   5   |  ↑ +2
Digital:     |   6   |   6   |  → 0
```

**Tiempo estimado:** 4 horas

---

### 15. WhatsApp automation (via Make/Zapier)
**Qué hace:** Cuando cambia el estado, enviar mensaje de WhatsApp al participante.

**Cómo:** Apps Script llama a un webhook de Make/Zapier que envía el WhatsApp.

**Requiere:** Cuenta en Make (gratis) + número de WhatsApp Business

**Tiempo estimado:** 2 horas de configuración

---

## 📊 IDEAS DE DISEÑO

### Dashboard en hoja "Dashboard"
- Actualización automática al sincronizar
- Por perfil (A/B/C/D)
- Por estado (Orientación/Mentoría/Formación/Cierre)
- Por zona geográfica
- Puntaje promedio / máximo / mínimo
- Evolución mensual (nuevos por mes)

### Sidebar del participante
- Foto del participante (si está subida en Drive)
- Timeline visual de cambios de estado
- Gráfico de radar más grande y colorido
- Botones de acción: "Cambiar estado", "Registrar derivación"
- Contador de días en el estado actual

### Colores por prioridad (además de perfil)
- Borde de fila en rojo brillante para casos CRÍTICO
- Ícono en la primera columna: 🔴 🟠 🟡 🟢

---

## 🗓️ ROADMAP SUGERIDO

| Versión | Features                                              | Estado     |
|---------|-------------------------------------------------------|------------|
| v7.1    | Fix triggers + sync Sheets→Docs básico                | ✅ HECHO    |
| v7.2    | Sync TODOS los campos + historial en Doc              | 🔄 En progreso |
| v8.0    | Validaciones + semáforo + calendario + PDF            | 📋 Planeado |
| v9.0    | Dashboard web + múltiples Kobo + re-diagnóstico       | 💡 Futuro   |

---

*Sistema Paso a Paso — Creamos Guatemala*
*Última actualización: 2026-05-12*
