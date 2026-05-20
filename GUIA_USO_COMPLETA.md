# 📚 GUÍA DE USO COMPLETA — Sistema Paso a Paso v8.5

## 🎯 ¿Qué es este sistema?

Sistema integrado para gestionar participantes de programas de empleabilidad y formación. Importa datos de múltiples fuentes (Kobo, Google Sheets), los organiza, crea expedientes automáticos en Drive, y genera reportes de impacto.

---

## ⚡ INSTALACIÓN RÁPIDA (5 minutos)

### 1️⃣ **Descarga el código**
Desde: https://github.com/adrian-9856/DP_paso_a_paso

Archivo: `Paso_a_Paso.gs`

### 2️⃣ **Pega en Google Sheets**
1. Abre tu Google Sheet **"Paso a Paso - Sistema Maestro"**
2. Herramientas > Editor de secuencias de comandos
3. Selecciona TODO (Ctrl+A) y borra
4. Pega el código de `Paso_a_Paso.gs`
5. Presiona Ctrl+S para guardar
6. Recarga la página (F5)

### 3️⃣ **Instala el sistema**
Verás un menú: `📊 PASO A PASO`

Haz clic en: `📥 Instalar Sistema`

El sistema creará automáticamente:
- Hoja **Maestro** (participantes principales)
- Hoja **Derivaciones** (personas en otros programas)
- Hoja **Log** (registro de cambios)
- Hoja **Dashboard** (estadísticas)
- Hoja **Analytics** (análisis detallado)

### 4️⃣ **Configura credenciales** (IMPORTANTE)
`📊 PASO A PASO > ⚙️ Configuración`

Ingresa:
- **Email para alertas**: tu@email.com (donde recibir notificaciones)
- **ID de carpeta Drive**: para guardar expedientes (obtén de la URL)
- **API Key Kobo** (si usas Kobo): desde Kobo > Cuenta > Seguridad

---

## 📥 IMPORTAR DATOS — 4 OPCIONES

### OPCIÓN 1: Desde **DP_Empleabilidad** (Formación/RRHH)
Personas en programas de formación técnica.

```
📊 PASO A PASO > 🏢 Importar DP_Empleabilidad
```

Esto:
- ✅ Lee datos de tu sheet externo (DP_Empleabilidad)
- ✅ Detecta nuevo/existentes (no duplica)
- ✅ Crea derivaciones automáticas
- ⏭️ Siguiente: "📁 Crear Expedientes Drive"

**Mapeo de columnas:**
| Tu hoja | → | Sistema |
|---------|---|---------|
| Fecha de ingreso | → | Fecha_Registro |
| Creamos ID | → | ID_Creamos |
| Nombre completo | → | Nombre_Completo |
| Teléfono | → | Teléfono |
| Edad | → | Edad |
| Nivel educativo | → | Nivel_Educativo |
| Formación | → | Situación_Laboral |
| Cohorte | → | Fortalezas |
| Nota | → | Objetivo_Laboral |
| Activo | → | Estado (Sí→Orientación, No→Inactivo) |

---

### OPCIÓN 2: Desde **Kobo** (Encuestas)
Si usas formularios Kobo para captura de datos.

```
📊 PASO A PASO > 🔄 Sincronizar Kobo (opcional)
```

Requiere:
- API Key Kobo configurada en ⚙️ Configuración
- Acceso al formulario Kobo específico

---

### OPCIÓN 3: Importar de OTRA hoja externa
Para AYB, TECH, u otro programa.

**Crea un vinculo personalizado:**

```javascript
// En el Editor de Scripts, agrega:
const OTRA_FUENTE = {
  SPREADSHEET_ID: 'ID_DEL_SPREADSHEET_AQUI',
  HOJA: 'Nombre de la hoja',
  NCOLS: 12
};

// Luego usa el mismo patrón que sincronizarDesdeSheet()
```

---

### OPCIÓN 4: Manual (Una por una)
En el Maestro, haz clic en `➕` para agregar fila y completa manualmente.

---

## 📁 CREAR EXPEDIENTES EN DRIVE

Después de importar, crea carpetas y documentos:

```
📊 PASO A PASO > 📁 Crear Expedientes Drive
```

Esto:
- Crea carpeta: `[ID] - [Nombre completo]`
- Crea documento: `Perfil - [Nombre]`
- Genera 10 por ejecución (evita timeout)
- Puedes ejecutar varias veces hasta completar

---

## 📊 VER DATOS

### Dashboard (Estadísticas rápidas)
```
📊 PASO A PASO > 📊 Dashboard
```
Muestra:
- Total participantes
- Por perfil (A/B/C/D)
- Por estado
- Puntaje promedio

### Analytics (Análisis detallado)
```
📊 PASO A PASO > 📈 Analytics
```
Muestra:
- Dimensiones de diagnóstico (educativo, laboral, digital, etc.)
- Por zona/región
- Derivaciones (urgentes, pendientes, completadas)
- Puntajes

### Ver Ficha Individual
1. Selecciona una fila en el Maestro
2. `📊 PASO A PASO > 📋 Ver Ficha`
3. Se abre un sidebar con el perfil completo + gráficos

---

## 🔧 CAMBIAR ESTADO DE PARTICIPANTE

En la columna **"Estado"** (M), puedes cambiar entre:
- `Orientación` → `Mentoría` → `Formación` → `Cierre`
- O directamente a: `Derivación`, `Inactivo`

El sistema:
- ✅ Registra la fecha del cambio
- ✅ Actualiza el expediente en Drive
- ✅ Envía notificación por email
- ✅ Crea evento en Google Calendar

---

## 📄 GENERAR REPORTE MENSUAL

### Con IA (Recomendado)

1. Copia los números del mes desde **Analytics**
2. Abre: `REPORTE_MENSUAL_TEMPLATE.md` (en este repo)
3. Pega los números donde dice `[PEGAR AQUÍ LOS NÚMEROS]`
4. Usa un modelo de IA (Claude, ChatGPT) con el prompt incluido
5. Copia el reporte generado

### Manual

Crea un documento con estas secciones:
- **Titular de impacto**: 1 oración con logro principal
- **Resumen**: 3 oraciones (qué se hizo, resultado, qué sigue)
- **Logros**: hasta 5 puntos con datos concretos
- **Desafíos**: hasta 3 puntos honestos
- **Próximas acciones**: 3 pasos específicos

---

## 🗑️ LIMPIAR/DESINSTALAR

```
📊 PASO A PASO > 🗑️ Desinstalar & Limpiar
```

Elimina:
- ✅ Todas las hojas (Maestro, Derivaciones, Log, Dashboard, Analytics)
- ✅ Todas las carpetas en Drive
- ✅ Todos los triggers automáticos
- ✅ Toda la configuración guardada

Requiere 2 confirmaciones (es definitivo).

---

## ⚠️ TROUBLESHOOTING

### "Error: Credencial no configurada"
**Solución**: Abre `⚙️ Configuración` e ingresa Email y Folder ID

### "No se puede abrir DP_Empleabilidad"
**Solución**: 
- Verifica que tienes acceso al spreadsheet
- Comparte el spreadsheet contigo mismo si es necesario

### "Maximum call stack size exceeded"
**Solución**: 
- Si pasa en importación, intenta con menos filas (máx. 100 por vez)
- Los expedientes se crean de a 10, puedes ejecutar "📁 Crear Expedientes" varias veces

### El Dashboard no actualiza
**Solución**: Abre `📊 PASO A PASO > 📊 Dashboard` para forzar actualización

### No recibo emails de alertas
**Solución**:
- Verifica email en `⚙️ Configuración`
- Revisa carpeta Spam
- Algunos cambios de estado pueden no enviar email (solo Mentoría, Formación, Cierre)

---

## 🔐 SEGURIDAD

- ✅ API Keys se guardan en PropertiesService (NO en el código)
- ✅ Contraseñas NO se almacenan (excepto localmente)
- ✅ Logs de cambios en la hoja "Log"
- ✅ Validación de flujo de estados (no todos pueden cambiar a todos)

---

## 📞 SOPORTE

- Revisa los **Logs** en la hoja "Log" si algo falla
- En el Editor de Scripts: Ejecuciones (Ctrl+Shift+H) para ver errores
- Contacta al admin del sistema con los detalles del error

---

**Última actualización**: 2026-05-19  
**Versión**: 8.5  
**Estado**: ✅ Lista para usar
