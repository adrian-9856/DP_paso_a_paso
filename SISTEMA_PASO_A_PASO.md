# Sistema Integral "Paso a Paso" - Fito

## Descripción General
Sistema centralizado de gestión de participantes que integra derivaciones de KoboToolbox, perfilado automático, seguimiento de estados y automatización de documentos en Google Drive/Docs.

## Arquitectura del Sistema

### 1. Capas de Entrada (Input Multicanal)
- **KoboToolbox API**: Formulario de derivación (endpoint: `/api/v2/assets/abHRWdRnPhKwzPQBajc7RZ/export-settings/esxXsxLthEAgNHbXQnV8qYV/data.csv`)
- **Hojas derivadas**: Información interna desde DPs (AYB, TECH, DP_Empleabilidad)

### 2. Cerebro Central (Google Sheets)
**Hoja Maestra**: Tabla consolidada con todas las columnas y lógicas

#### A. Información Base del Participante
- `ID_Creamos` (generado automáticamente)
- `Nombre_Completo`
- `DPI`
- `Edad`
- `Género`
- `Teléfono`
- `Zona`
- `Fecha_Registro`
- `Fuente` (KoboToolbox / Manual / Derivación interna)
- `Programa_Origen` (Educación, Empleabilidad, TECH, etc.)

#### B. Perfilado Inicial (Matriz)
- `Nivel_Educativo` (Sin educación, Primaria, Secundaria, Técnico, Universitario)
- `Problemas_Vitales` (lista de dimensiones del formulario Kobo)
- `Conciencia_Participante` (Alto, Medio, Bajo)
- `Perfil_Resultante` (Perfil 1, 2, 3 - Independiente/Quiere-puede-necesita/Quiere-no-puede)

#### C. Gestión de Estados
- `Estado_Actual` (Orientación, Mentoría, Formación, Cierre, Derivación, Activo, Inactivo)
- `Fecha_Último_Cambio_Estado`
- `Responsable_Actual`
- `Historial_Estados` (registro de transiciones)

#### D. Seguimiento de Mentoría
- `Fecha_Inicio_Mentoría`
- `Número_Sesiones_Completadas`
- `Fecha_Última_Sesión`
- `Duración_Mentoría_Meses` (calculado)
- `Alertas_Mentoría` (Rojo si >6 meses O >3 sesiones)
- `Próxima_Sesión_Programada`

#### E. Gestión de Documentos
- `Carpeta_Drive_ID` (ID único de carpeta del participante)
- `Expediente_Digital_ID` (ID del Google Doc)
- `Documentos_Compartidos` (lista de archivos asociados)
- `Última_Actualización_Expediente`

#### F. Información de Derivación (Interna)
- `Tipo_Derivación` (Interna / Externa)
- `Derivado_A` (Programa destino)
- `Motivo_Derivación`
- `Fecha_Derivación`
- `Estado_Derivación` (Pendiente, Aceptada, Rechazada)

#### G. Datos de Cierre
- `Tipo_Cierre` (Completado, No completado, Derivado, Abandonado)
- `Fecha_Cierre`
- `Resultado_Final` (texto)
- `Encargado_Cierre`

### 3. Automatización (Google Apps Script)
#### Flujos Automáticos

**3.1 Importación de KoboToolbox**
- Descarga CSV cada día a las 6 AM
- Detecta nuevos registros
- Asigna `ID_Creamos` automáticamente (formato: `PRIMERA_LETRA + NÚMEROS_DPI_ÚLTIMOS_6 + FECHA`)
- Rellena tabla maestra

**3.2 Creación de Carpeta y Expediente**
- Trigger: Nuevo participante o cambio de estado a "Mentoría"
- Crea carpeta en Google Drive con nombre: `[ID_Creamos] - [Nombre_Completo]`
- Crea Google Doc "Expediente Digital" con:
  - Datos básicos
  - Tags dinámicos (por programa, estado, perfil)
  - Secciones de notas (actualizables)
  - Historial de cambios de estado

**3.3 Alertas de Mentoría**
- Trigger: Diario a las 9 AM
- Busca participantes en "Mentoría"
- Calcula: `hoy - fecha_inicio` en meses Y `número_sesiones`
- Si `meses >= 6` O `sesiones >= 3`:
  - Marca celda en rojo (semáforo)
  - Envía email al responsable
  - Crea recordatorio en expediente

**3.4 Gestión de Estados**
- Permite transiciones NO lineales (any→any, excepto Cierre que es final)
- Registra fecha y responsable
- Actualiza expediente automáticamente
- Notifica cambios al responsable

### 4. Estructura de Columnas Excel/Sheets

| Columna | Tipo | Fuente | Fórmula/Lógica |
|---------|------|--------|-----------------|
| **A** | ID_Creamos | Auto | `=LEFT(D2,1) & MID(E2,6,6) & TEXT(TODAY(),"DDMM")` |
| **B** | Fecha_Registro | Kobo/Manual | Timestamp |
| **C** | Nombre_Completo | Kobo | Texto |
| **D** | DPI | Kobo | Texto |
| **E** | Edad | Kobo | Número |
| **F** | Género | Kobo | Dropdown |
| **G** | Teléfono | Kobo | Texto |
| **H** | Zona | Kobo | Dropdown |
| **I** | Nivel_Educativo | Kobo | Dropdown (Sin educación, Primaria, Secundaria, Técnico, Universitario) |
| **J** | Problemas_Vitales | Kobo | Texto (múltiple) |
| **K** | Conciencia_Participante | Entrevista | Dropdown (Alto, Medio, Bajo) |
| **L** | Perfil_Resultante | Auto | Fórmula de matriz |
| **M** | Estado_Actual | Dropdown | Orientación, Mentoría, Formación, Cierre, Derivación, Activo, Inactivo |
| **N** | Fecha_Último_Cambio | Auto | `=IF(M2<>M1,TODAY(),N1)` |
| **O** | Responsable_Actual | Manual | Texto |
| **P** | Inicio_Mentoría | Auto | `=IF(M2="Mentoría",TODAY(),P1)` |
| **Q** | Sesiones_Completadas | Manual | Número |
| **R** | Última_Sesión | Manual | Fecha |
| **S** | Duración_Meses | Auto | `=IF(M2="Mentoría", INT((TODAY()-P2)/30.44), "")` |
| **T** | Alerta_Mentoría | Auto | `=IF(OR(S2>=6, Q2>=3), "🔴 ALERTA", "🟢 OK")` |
| **U** | Próxima_Sesión | Manual | Fecha |
| **V** | Carpeta_Drive_ID | Auto | Script genera ID |
| **W** | Expediente_ID | Auto | Script genera ID |
| **X** | Fuente | Kobo/Manual | Dropdown |
| **Y** | Programa | Kobo/Manual | Dropdown |
| **Z** | Tipo_Derivación | Manual | Dropdown (Interna/Externa/Ninguna) |
| **AA** | Derivado_A | Manual | Texto |
| **AB** | Tipo_Cierre | Manual | Dropdown (Completado, No completado, Derivado, Abandonado) |
| **AC** | Fecha_Cierre | Manual | Fecha |
| **AD** | Resultado_Final | Manual | Texto |
| **AE** | Encargado_Cierre | Manual | Texto |

### 5. Fórmulas Críticas

#### Matriz de Perfilado (Columna L)
```
=IF(K2="Alto",
  IF(I2="Universitario", "Perfil 1: Independiente",
    IF(OR(Q2>=3, S2>=6), "Perfil 1: Independiente", "Perfil 2: Quiere-Puede-Necesita")),
  IF(K2="Medio",
    "Perfil 2: Quiere-Puede-Necesita",
    "Perfil 3: Quiere-No Puede"))
```

#### Alerta de Mentoría (Columna T)
```
=IF(M2="Mentoría",
  IF(OR(S2>=6, Q2>=3),
    "🔴 ALERTA - Revisar",
    "🟢 En progreso"),
  "")
```

### 6. Procesos de Flujo

#### Flujo 1: Nuevo Participante (Kobo)
1. Datos llegan vía Kobo
2. Script detecta nuevo registro
3. Asigna `ID_Creamos`
4. Llena información base
5. Crea carpeta en Drive
6. Envía notificación al responsable

#### Flujo 2: Cambio de Estado
1. Usuario cambia `Estado_Actual`
2. Script registra cambio (fecha + responsable)
3. Actualiza expediente
4. Envía notificación

#### Flujo 3: Mentoría (Monitoreo)
1. Script revisa diariamente participantes en "Mentoría"
2. Calcula duración y sesiones
3. Si alerta: marca celda roja + email + nota en expediente

#### Flujo 4: Cierre Normal
1. Responsable (encargado DP) marca `Tipo_Cierre`
2. Completa `Resultado_Final`
3. Script registra fecha y responsable
4. Genera resumen en expediente
5. Archiva participante

#### Flujo 5: Derivación Interna
1. Usuario marca `Tipo_Derivación = "Interna"`
2. Especifica `Derivado_A`
3. Script notifica nuevo DP
4. Estado provisional: "Derivación"
5. Nueva DP acepta/rechaza
6. Estado final: "Activo" (otro programa) o regresa

### 7. Integraciones

**KoboToolbox → Sheets**
- Trigger: Diario 6 AM
- URL CSV con autenticación
- Importa nuevos registros

**Sheets → Google Drive/Docs**
- Trigger: Nuevo participante O cambio de estado
- Crea carpeta + doc automáticamente
- Mantiene referencias (IDs)

**Alertas → Email**
- Trigger: Mentoría >6 meses O >3 sesiones
- Destinatarios: Responsable actual
- Frecuencia: Diaria a las 9 AM

### 8. Vistas Recomendadas

**Tabla Principal**: Todos los datos
**Filtro 1**: Activos en Mentoría (para seguimiento diario)
**Filtro 2**: Alertas rojas (revisión urgente)
**Filtro 3**: Pendientes de cierre
**Filtro 4**: Derivaciones pendientes

### 9. Roles y Permisos

- **Super Admin**: Acceso completo + configuración scripts
- **Responsable DP**: Ver/editar participantes su programa + cambiar estados
- **Mentores**: Registrar sesiones (columna Q, R)
- **Encargados Cierre**: Completar cierre (columnas AB-AE)

---

## Próximos Pasos

1. **Fase 1**: Crear Sheet maestro con estructura
2. **Fase 2**: Implementar Apps Scripts
3. **Fase 3**: Probar integración Kobo
4. **Fase 4**: Crear vistas y filtros
5. **Fase 5**: Capacitar usuarios
