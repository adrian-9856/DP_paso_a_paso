# 🚀 INSTALACIÓN - SISTEMA PASO A PASO v3.0

## ✅ REQUISITOS PREVIOS

- [ ] Google Sheet vacío (o existente): "Paso a Paso - Sistema Maestro"
- [ ] Carpeta en Google Drive para participantes
- [ ] Cuenta de KoboToolbox con acceso al formulario
- [ ] API Key de KoboToolbox

---

## 📋 PASO 1: OBTENER TU API KEY DE KOBO

1. Ve a: https://kf.kobotoolbox.org/admin/auth/token/
2. Login con tus credenciales de Kobo
3. Copia tu token (API Key) - Algo como: `64cc018b88067397addd36b09288be8b6539cf39`
4. Guárdalo en un lugar seguro

---

## 🗂️ PASO 2: CREAR CARPETA EN GOOGLE DRIVE

1. Abre Google Drive
2. Crea una carpeta nueva: **"Paso a Paso - Participantes"**
3. Copia el ID de la URL:
   - URL: `https://drive.google.com/drive/folders/AQUI_VA_EL_ID`
   - Ejemplo: `1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU`
4. Guárdalo

---

## 📊 PASO 3: CREAR GOOGLE SHEET

1. Ve a Google Sheets: https://sheets.google.com
2. **Crear nuevo documento**
3. Renombra a: **"Paso a Paso - Sistema Maestro"**
4. Renombra la hoja a: **"Maestro"** (sin tildes)
5. En la primera fila (A1), pega los encabezados:

```
ID_Creamos | Fecha_Registro | Nombre_Completo | DPI | Edad | Genero | Telefono | Zona | Nivel_Educativo | Problemas_Vitales | Conciencia_Participante | Perfil_Resultante | Estado_Actual | Fecha_Ultimo_Cambio | Responsable_Actual | Inicio_Mentoria | Sesiones_Completadas | Ultima_Sesion | Duracion_Meses | Alerta_Mentoria | Proxima_Sesion | Carpeta_Drive_ID | Expediente_ID | Expediente_URL | Foto_URL | Fuente | Programa | Tipo_Derivacion | Derivado_A | Tipo_Cierre | Fecha_Cierre | Resultado_Final | Encargado_Cierre
```

6. Guarda (Ctrl+S)

---

## 🔧 PASO 4: INSTALAR EL SCRIPT

1. En tu Sheet, abre: **Herramientas > Editor de secuencias de comandos**
2. Se abrirá una pestaña nueva
3. **Borra TODO el código** que encuentres (Ctrl+A, Delete)
4. Descarga el archivo `Paso_a_Paso.gs` desde:
   - https://github.com/adrian-9856/DP_paso_a_paso/blob/claude/paso-a-paso-cloud-system-FENDi/Paso_a_Paso.gs
5. **Copia TODO el código** (Ctrl+A)
6. **Pégalo** en el Editor de Secuencias (Ctrl+V)
7. Presiona **Ctrl+S** para guardar
8. **RECARGA la página** (F5)

---

## ⚙️ PASO 5: CONFIGURACIÓN INICIAL

1. En tu Sheet, aparecerá un menú nuevo: **"📊 PASO A PASO"**
2. Haz click en: **⚙️ Configuración**
3. Completa los campos:
   - **ID de Carpeta**: Pega `1pVDrNCwLRX41qiu9jJBFZ--wNm1TSWXU`
   - **API Key**: Pega tu token de Kobo
4. Haz click en: **💾 GUARDAR**
5. Haz click en: **🧪 PROBAR CONEXIÓN**

**Si ves ✅ EXITOSA**: ¡Todo está bien!
**Si ves ❌ Error**: Verifica tu API Key

---

## 🎯 PASO 6: SINCRONIZAR DATOS

1. En el menú **"📊 PASO A PASO"**, haz click en:
   - **📥 Sincronizar Kobo (Manual)**

2. El sistema:
   - ✅ Descargará datos de tu formulario Kobo
   - ✅ Detectará participantes nuevos
   - ✅ Creará una carpeta por participante
   - ✅ Generará un documento de perfil
   - ✅ Llenará el Sheet automáticamente

3. Verás un mensaje: **"✅ Sincronización completada"**

---

## 📁 ESTRUCTURA QUE SE CREA

Para cada participante se crea:

```
Paso a Paso - Participantes/
├── A1234 - Juan Pérez
│   ├── 📄 Perfil - Juan Pérez (Google Doc)
│   └── 📸 [Fotos y otros documentos]
├── M5678 - María López
│   ├── 📄 Perfil - María López
│   └── 📸 [Fotos y otros documentos]
```

---

## 🎨 DOCUMENTO DE PERFIL

Cada participante tiene un Google Doc con:

- **Información Personal**: DPI, edad, zona, educación
- **Programa**: A qué programa pertenece
- **Discusión de Caso**: Espacio para notas de sesiones
- **Plan de Acción**: Qué se va a hacer con el participante
- **Fotos**: Puedes agregar fotos manualmente

---

## 📸 AGREGAR FOTO DE PARTICIPANTE

1. Abre la carpeta del participante
2. Sube la foto (arrastra y suelta)
3. Abre el documento de perfil
4. Inserta > Imagen > Selecciona la foto que subiste

---

## 📊 USA LOS MENÚS

### 📥 Sincronizar Kobo
- Descarga datos nuevos de tu formulario
- Crea carpetas y documentos automáticamente

### 🔔 Revisar Alertas
- Busca participantes en mentoría > 6 meses
- Busca participantes con > 3 sesiones
- Marca las filas en rojo

### 📈 Ver Estadísticas
- Total de participantes
- Por estado (Orientación, Mentoría, Formación, Cierre)
- Cuántos en alerta

### 📖 Ayuda
- Guía rápida en el sistema

---

## 🆘 SOLUCIÓN DE PROBLEMAS

### "Error 401"
**Causa**: API Key incorrecta o expirada
**Solución**: 
1. Ve a https://kf.kobotoolbox.org/admin/auth/token/
2. Genera una nueva
3. Vuelve a ⚙️ Configuración y guarda

### No aparece el menú
**Causa**: No se ejecutó onOpen()
**Solución**: Recarga la página (F5)

### No se crean las carpetas
**Causa**: ID de carpeta incorrecto
**Solución**:
1. Abre la carpeta en Drive
2. Copia el ID correcto de la URL
3. Vuelve a ⚙️ Configuración y actualiza

### El Sheet da error
**Causa**: Nombre de hoja incorrecto
**Solución**: Verifica que la hoja se llame exactamente **"Maestro"**

---

## 📞 SOPORTE

Si tienes problemas:
1. Abre el Editor de Secuencias
2. Click en "Ejecuciones"
3. Revisa los logs para ver el error exacto
4. Copia el error y pregunta

---

## ✅ CHECKLIST FINAL

- [ ] API Key de Kobo obtenida
- [ ] Carpeta en Drive creada e ID copiado
- [ ] Google Sheet "Maestro" creado con encabezados
- [ ] Script `Paso_a_Paso.gs` pegado en Editor
- [ ] Página recargada (F5)
- [ ] Menú "📊 PASO A PASO" aparece
- [ ] Configuración guardada
- [ ] Conexión probada ✅
- [ ] Sincronización completada
- [ ] Participantes aparecen en el Sheet
- [ ] Carpetas creadas en Drive
- [ ] Documentos de perfil generados

**¡LISTO PARA USAR!** 🎉

---

## 🎯 PRÓXIMAS MEJORAS PLANEADAS

- [ ] Sincronización automática diaria
- [ ] Envío de reportes por email
- [ ] Integración con formularios de retroalimentación
- [ ] Gráficos de progreso
- [ ] Exportación a PDF
- [ ] Integración con WhatsApp (opcional)

---

**Versión**: 3.0
**Fecha**: 2025-05-12
**Estado**: ✅ PRODUCCIÓN
