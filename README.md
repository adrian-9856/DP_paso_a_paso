# Sistema Integral "Paso a Paso" - Fito

Sistema centralizado de gestión de participantes con integración KoboToolbox, automatización de documentos y seguimiento inteligente de mentorías.

## 📋 Componentes

### 1. Documentación Principal
- **SISTEMA_PASO_A_PASO.md** - Especificación completa del sistema (estructura, flujos, fórmulas)

### 2. Scripts Google Apps (en carpeta `scripts/`)

| Script | Función | Trigger |
|--------|---------|---------|
| **config.gs** | Configuración centralizada, constantes, funciones base | Manual |
| **importar_kobo.gs** | Descarga y procesa datos de KoboToolbox | Diario 6 AM |
| **crear_documentos.gs** | Crea carpetas y expedientes digitales | Nuevo participante |
| **alertas_mentoria.gs** | Monitorea y genera alertas de mentoría | Diario 9 AM |
| **gestionar_estados.gs** | Maneja transiciones de estado | onEdit |
| **utilidades.gs** | Funciones helper reutilizables | Manual |

## 🚀 Instalación

### Paso 1: Crear Google Sheet Maestro

1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea un nuevo documento: **"Paso a Paso - Sistema Maestro"**
3. Renombra la hoja a: **"Maestro"**
4. Copia el ID del documento (de la URL):
   - URL: `https://docs.google.com/spreadsheets/d/{ID}/edit`

### Paso 2: Crear estructura de columnas

Copia estas columnas en la primera fila (A1:AE1):

```
A: ID_Creamos
B: Fecha_Registro
C: Nombre_Completo
D: DPI
E: Edad
F: Género
G: Teléfono
H: Zona
I: Nivel_Educativo
J: Problemas_Vitales
K: Conciencia_Participante
L: Perfil_Resultante
M: Estado_Actual
N: Fecha_Último_Cambio
O: Responsable_Actual
P: Inicio_Mentoría
Q: Sesiones_Completadas
R: Última_Sesión
S: Duración_Meses
T: Alerta_Mentoría
U: Próxima_Sesión
V: Carpeta_Drive_ID
W: Expediente_ID
X: Fuente
Y: Programa
Z: Tipo_Derivación
AA: Derivado_A
AB: Tipo_Cierre
AC: Fecha_Cierre
AD: Resultado_Final
AE: Encargado_Cierre
```

### Paso 3: Configurar permisos y Drive

1. Crea una carpeta en Google Drive: **"Participantes"**
2. Copia su ID:
   - Abre la carpeta → URL: `https://drive.google.com/drive/folders/{ID}`

### Paso 4: Instalar Scripts

1. En el Sheet maestro: `Herramientas > Editor de secuencias de comandos`
2. Copia **todo el contenido** de cada archivo `.gs` en orden:
   - `config.gs` (primero - define variables globales)
   - `importar_kobo.gs`
   - `crear_documentos.gs`
   - `alertas_mentoria.gs`
   - `gestionar_estados.gs`
   - `utilidades.gs` (último)

3. **Edita `config.gs`** con tus valores:

```javascript
const CONFIG = {
  SPREADSHEET_ID: "AQUÍ_VA_ID_DEL_SHEET",
  FOLDER_PARTICIPANTES_ID: "AQUÍ_VA_ID_CARPETA_DRIVE",
  KOBO_USERNAME: "tu_usuario_kobo",
  KOBO_PASSWORD: "tu_contraseña_kobo",
  ADMIN_EMAIL: "adrian@creamosguatemla.org",
  // ... resto de config
};
```

### Paso 5: Configurar Triggers

En el Editor de secuencias de comandos:

1. **Importación Kobo** (Diario 6 AM)
   - Función: `importarDatosKobo`
   - Tipo evento: Intervalo de tiempo
   - Frecuencia: Diaria a las 6:00 AM

2. **Alertas de Mentoría** (Diario 9 AM)
   - Función: `revisarAlertasMentoria`
   - Tipo evento: Intervalo de tiempo
   - Frecuencia: Diaria a las 9:00 AM

3. **Cambios de Estado** (onEdit)
   - Función: `onEdit`
   - Tipo evento: Desde hoja
   - Evento: En edición
   - (Se ejecuta automáticamente)

4. **Estadísticas Diarias** (Diario 8 PM - opcional)
   - Función: `enviarEstadisticasDiarias`
   - Tipo evento: Intervalo de tiempo
   - Frecuencia: Diaria a las 20:00

## 📊 Cómo Usar

### Importar Participantes (Automático)

El sistema descarga automáticamente datos de KoboToolbox cada día a las 6 AM:
- ✅ Detecta participantes nuevos
- ✅ Genera ID único (Creamos ID)
- ✅ Crea carpeta y expediente
- ✅ Envía notificación por email

### Cambiar Estado de Participante

1. Abre el Sheet maestro
2. En columna **M (Estado_Actual)**, selecciona el nuevo estado:
   - Orientación
   - Mentoría
   - Formación
   - Cierre
   - Derivación
   - Activo
   - Inactivo

3. El sistema automáticamente:
   - ✅ Registra fecha del cambio
   - ✅ Actualiza expediente
   - ✅ Envía notificación

### Registrar Sesión de Mentoría

En columna **Q (Sesiones_Completadas)**:
- Incrementa el número manualmente O
- Usa función: `registrarSesionMentoria("ID_CREAMOS")`

### Crear Derivación Interna

Usa la función desde el Editor:
```javascript
crearDerivacion(
  "IDER123", // ID del participante
  "TECH",    // Programa destino
  "Motivo de derivación",
  "responsable@email.com"
)
```

### Completar Cierre

Usa función:
```javascript
completarCierre(
  "IDER123",
  "Completado", // o "No completado", "Derivado", "Abandonado"
  "Resultado del proceso",
  "responsable@email.com"
)
```

## 🔴 Alertas y Monitoreo

### Alerta de Mentoría (Automática)

Se genera automáticamente si:
- ⏱️ Duración ≥ 6 meses O
- 📊 Sesiones ≥ 3

**Acciones automáticas:**
- Celda se pone roja (🔴 ALERTA)
- Email al responsable
- Nota agregada al expediente

**Para resolver:**
- Cambiar estado a "Cierre" O "Derivación" O
- Resetear contadores (si continúa mentoría)

## 📱 Vista Mobile-Friendly (Opcional)

Se recomienda crear filtros para acceso rápido:

1. **Mi Carga** - Participantes asignados a mí
2. **En Alerta** - Solo los que tienen 🔴 ALERTA
3. **Por Completar** - Sin cierre aún

## 🔐 Seguridad y Permisos

### Roles Recomendados

| Rol | Permisos |
|-----|----------|
| **Admin** | Acceso completo + config scripts |
| **Responsable DP** | Ver/editar participantes su programa + cambiar estados |
| **Mentor** | Solo registrar sesiones |
| **Encargado Cierre** | Registrar tipo y resultado cierre |

### Compartir Sheet

1. Click derecho en documento
2. **Compartir**
3. Añade usuarios por rol

## 🆘 Troubleshooting

### "Error de autenticación Kobo"
- Verificar usuario/contraseña en `config.gs`
- Verificar que el usuario tenga acceso a ese asset en Kobo

### "Carpeta no existe"
- Verificar que la carpeta "Participantes" exista en Drive
- Copiar ID correcto en `config.gs`

### "Expediente no se crea"
- Verificar que la carpeta tenga permiso de escritura
- Revisar logs en Editor de scripts

### Scripts no se ejecutan
- Verificar que los triggers estén activos
- Revisar zona horaria en `Configuración del Proyecto`

## 📈 Fórmulas Principales

### Generar ID Creamos (Columna A)
```
=IF(C2="","",LEFT(C2,1) & MID(D2,6,6) & TEXT(TODAY(),"DDMM"))
```

### Calcular Duración Mentoría (Columna S)
```
=IF(M2="Mentoría",INT((TODAY()-P2)/30.44),"")
```

### Alerta Automática (Columna T)
```
=IF(M2="Mentoría",
  IF(OR(S2>=6, Q2>=3),
    "🔴 ALERTA - Revisar",
    "🟢 En progreso"),
  "")
```

## 📞 Soporte

Para reportar errores o sugerencias:
1. Revisa los logs en `Ejecuciones` del Editor de scripts
2. Contacta al admin: adrian@creamosguatemla.org

## 📝 Notas Importantes

- **Backup regular**: Descarga backups del Sheet periódicamente
- **Documentación**: Actualiza `SISTEMA_PASO_A_PASO.md` si cambias la estructura
- **Capacitación**: Capacita a usuarios sobre flujos antes de usar
- **Auditoría**: Revisa regularmente logs de cambios en el Sheet

---

**Versión**: 1.0  
**Última actualización**: 2025-05-07  
**Creado para**: Organización Fito
