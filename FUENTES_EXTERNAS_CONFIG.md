# 🔗 CONFIGURACIÓN DE MÚLTIPLES FUENTES DE DATOS

Este archivo te muestra cómo agregar más fuentes de importación (además de DP_Empleabilidad).

---

## 📌 FUENTES ACTUALMENTE CONFIGURADAS

### 1. **DP_Empleabilidad** (Formación/RRHH)
- **Spreadsheet ID**: `1_596FX6yr8tX93UyIks4emSeE2_vxLJMDyw9Zncsnzs`
- **Hoja**: `Paso a paso`
- **Comando**: `🏢 Importar DP_Empleabilidad`
- **Destino**: Maestro + Derivaciones

---

## ➕ AGREGAR NUEVA FUENTE

### Paso 1: Obtén el ID del Spreadsheet
1. Abre el Google Sheet de origen
2. En la URL, copia la parte en **negrita**:
   ```
   https://docs.google.com/spreadsheets/d/[AQUÍ_ESTA_EL_ID]/edit
   ```

### Paso 2: Identifica el nombre de la hoja
Ejemplo: "AYB Participantes", "TECH Cohorte 2", etc.

### Paso 3: Mapea las columnas
Tu hoja debe tener columnas como mínimo:
- Nombre completo (o similar)
- ID (o similar)
- Teléfono (opcional)
- Edad (opcional)
- Activo / Estado (opcional)

---

## 🔧 AGREGAR UNA FUENTE EN EL CÓDIGO

En el **Editor de Scripts**, agrega esto **debajo de la constante `DP_EMPLEABILIDAD`**:

```javascript
// ============================================================================
// FUENTE: [NOMBRE DEL PROGRAMA]
// ============================================================================

const [NOMBRE_FUENTE] = {
  SPREADSHEET_ID: '[ID_DEL_SPREADSHEET]',
  HOJA: '[Nombre de la hoja]',
  NCOLS: [NÚMERO_COLUMNAS],  // ejemplo: 12
  C: {
    // Mapa de columnas — los ÍNDICES (0-based) de tu hoja
    FECHA:     0,    // Columna A
    ID:        1,    // Columna B
    NOMBRE:    2,    // Columna C
    TELEFONO:  3,    // Columna D (opcional)
    GENERO:    4,    // Columna E (opcional)
    EDAD:      5,    // Columna F (opcional)
    ESTADO:    6,    // Columna G (opcional)
    // ... agrega más según tus columnas
  }
};

// Función para importar desde esta fuente
function importarDe[NOMBRE_FUENTE](silencioso) {
  sincronizarDesdeSheetGenerico([NOMBRE_FUENTE], 'Maestro', silencioso);
}
```

---

## 📋 EJEMPLOS CONCRETOS

### EJEMPLO 1: Importar de "AYB"

**URL de la hoja:** 
```
https://docs.google.com/spreadsheets/d/1xYzABC123DEF456GHI789/edit
```

**Código a agregar:**

```javascript
const FUENTE_AYB = {
  SPREADSHEET_ID: '1xYzABC123DEF456GHI789',
  HOJA: 'Maestro AYB',
  NCOLS: 10,
  C: {
    FECHA:     0,    // A: Fecha de ingreso
    ID:        1,    // B: ID Participante
    NOMBRE:    2,    // C: Nombre
    TELEFONO:  3,    // D: Teléfono
    EDAD:      4,    // E: Edad
    GENERO:    5,    // F: Género
    ESTADO:    6,    // G: Activo/Inactivo
    ZONA:      7,    // H: Zona
    EMAIL:     8,    // I: Email
    PROGRAMA:  9     // J: Programa
  }
};

function importarDesdeAYB() {
  sincronizarDesdeSheetGenerico(FUENTE_AYB, 'Maestro', false);
}
```

---

### EJEMPLO 2: Importar de "TECH Cohorte"

```javascript
const FUENTE_TECH = {
  SPREADSHEET_ID: '1aBcDEfGHiJkLmNoPqRsT789',
  HOJA: 'Cohorte 5',
  NCOLS: 8,
  C: {
    FECHA:     0,    // Fecha ingreso
    ID:        1,    // ID
    NOMBRE:    2,    // Nombre completo
    EMAIL:     3,    // Email
    TELEFONO:  4,    // Teléfono
    EDAD:      5,    // Edad
    FORMACION: 6,    // Tipo de formación
    ACTIVO:    7     // ¿Activo?
  }
};

function importarDesdeTECH() {
  sincronizarDesdeSheetGenerico(FUENTE_TECH, 'Maestro', false);
}
```

---

## 🎛️ ACTUALIZAR EL MENÚ

Después de agregar las funciones, actualiza el menú:

**Busca esta sección en `onOpen()`:**

```javascript
.addItem('🏢 Importar DP_Empleabilidad', 'sincronizarDesdeSheet')
.addItem('📁 Crear Expedientes Drive', 'crearExpedientesPendientes')
```

**Y agrega:**

```javascript
.addItem('🏢 Importar DP_Empleabilidad', 'sincronizarDesdeSheet')
.addItem('🏢 Importar AYB', 'importarDesdeAYB')
.addItem('🏢 Importar TECH', 'importarDesdeTECH')
.addItem('📁 Crear Expedientes Drive', 'crearExpedientesPendientes')
```

---

## 🧪 PROBAR LA NUEVA FUENTE

1. Pega el código en el Editor de Scripts
2. Presiona Ctrl+S para guardar
3. Recarga la página (F5)
4. En el menú debería aparecer tu nuevo item: `🏢 Importar [PROGRAMA]`
5. Haz clic y prueba

---

## 📊 FLUJOS DE DATOS

Actualmente soportamos 2 tipos de flujos:

### Flujo A: Participantes → Maestro
```
Tu hoja → sincronizarDesdeSheet → Maestro + Derivaciones
```

Ejemplo: Kobo, nuevo formulario de registro

### Flujo B: Personas en Programas Específicos → Hojas separadas
```
Tu hoja → importarDesdeFUENTE → Hoja especial [PROGRAMA]
```

Ejemplo: AYB ya tiene sus propios participantes en otra hoja

---

## ⚙️ MAPEO AUTOMÁTICO

Si tu hoja tiene columnas con EXACTAMENTE estos nombres, el sistema las detecta automáticamente:

- `Nombre`, `Nombre Completo`, `nombre_completo` → NOMBRE
- `ID`, `Creamos ID`, `id_participante` → ID
- `Teléfono`, `Telefono`, `Celular`, `Móvil` → TELEFONO
- `Edad`, `Años` → EDAD
- `Género`, `Genero`, `Sexo` → GENERO
- `Activo`, `Estado`, `Participando` → ESTADO
- `Email`, `Correo`, `E-mail` → EMAIL
- `Fecha`, `Fecha ingreso`, `Fecha de registro` → FECHA

---

## 🆘 TROUBLESHOOTING CUANDO AGREGAS FUENTE

**"No aparece el nuevo item en el menú"**
- Revisa que guardaste el código (Ctrl+S)
- Recarga la página (F5)
- Verifica que el nombre de la función sea correcto

**"Error: Spreadsheet no encontrado"**
- Verifica que el SPREADSHEET_ID sea correcto (sin espacios)
- Asegúrate de tener acceso al sheet

**"Columnas no mapean bien"**
- Revisa que los índices (0, 1, 2...) sean correctos
- Ejemplo: Columna A = índice 0, Columna B = índice 1, etc.

**"Importa pero con datos vacíos"**
- El mapeo de columnas está incorrecto
- Cuenta las columnas en tu hoja y ajusta los índices

---

## 📞 CONTACTO

Si necesitas ayuda agregando una fuente:
1. Copia el nombre exacto de tu hoja
2. Copia el SPREADSHEET_ID
3. Lista las columnas que tienes
4. Contacta al equipo técnico

---

**Ejemplo de email de soporte:**
```
Asunto: Agregar nueva fuente de datos

Quiero agregar importación desde [NOMBRE PROGRAMA]

Spreadsheet ID: [COPIAR DEL URL]
Nombre de la hoja: [NOMBRE EXACTO]
Columnas que tengo: Fecha | ID | Nombre | Email | Estado | ...

¿Puedes ayudarme?
```

---

**Última actualización:** 2026-05-19
