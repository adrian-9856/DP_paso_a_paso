# 🚀 GUÍA RÁPIDA - Sistema Paso a Paso (VERSIÓN CON BOTONES)

## ⚡ 3 PASOS PARA EMPEZAR

### **Paso 1: Descarga el código**

👉 Descarga el archivo: **`Sistema_Paso_a_Paso_COMPLETO.gs`**

Desde: https://github.com/adrian-9856/DP_paso_a_paso

### **Paso 2: Pega en Google Sheets**

1. Abre tu Google Sheet "Paso a Paso - Sistema Maestro"
2. Abre: `Herramientas > Editor de secuencias de comandos`
3. **Limpia todo el código existente** (Ctrl+A, Delete)
4. **Pega el código** de `Sistema_Paso_a_Paso_COMPLETO.gs` (Ctrl+V)
5. Presiona **Ctrl+S** para guardar
6. **Recarga la página** (F5)

### **Paso 3: Verás el MENÚ**

Cuando recargues, aparecerá un menú nuevo: **"📊 PASO A PASO"**

---

## 🎯 USAR EL MENÚ

### **1️⃣ Configuración Inicial**

Click en: `📊 PASO A PASO > ⚙️ Configuración Inicial`

Se abrirá un formulario donde debes ingresar:

| Campo | Valor |
|-------|-------|
| **ID de la Carpeta** | El ID de tu carpeta en Drive |
| **Usuario Kobo** | Tu usuario de KoboToolbox |
| **Contraseña Kobo** | Tu contraseña (se guarda localmente) |

**¿Cómo obtener el ID de la carpeta?**
- Abre tu carpeta en Google Drive
- Copia de la URL: `drive.google.com/drive/folders/AQUI_VA_EL_ID`

### **2️⃣ Importar datos de KoboToolbox**

Click en: `📊 PASO A PASO > 📥 Importar de KoboToolbox`

Descargará automáticamente los datos nuevos del formulario.

### **3️⃣ Revisar Alertas**

Click en: `📊 PASO A PASO > 🔔 Revisar Alertas de Mentoría`

Muestra participantes que necesitan atención (>6 meses o >3 sesiones).

### **4️⃣ Ver Estadísticas**

Click en: `📊 PASO A PASO > 📈 Ver Estadísticas`

Resumen visual con:
- Total de participantes
- Por estado (Orientación, Mentoría, Formación, etc.)
- Alertas activas

---

## 📋 COLUMNAS QUE DEBES CREAR

En tu Sheet, crea estas columnas en la PRIMERA FILA (A1:AE1):

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

**Fácil:** Copia y pega todo en A1

---

## ✅ CHECKLIST DE INSTALACIÓN

- [ ] Descargué `Sistema_Paso_a_Paso_COMPLETO.gs`
- [ ] Pegué el código en Google Apps Script
- [ ] Recargué la página (F5)
- [ ] Veo el menú "📊 PASO A PASO"
- [ ] Abrí "⚙️ Configuración Inicial"
- [ ] Ingresé los datos requeridos
- [ ] Hice click en "💾 GUARDAR CONFIGURACIÓN"
- [ ] Probé la conexión "🧪 PROBAR CONEXIÓN"
- [ ] Creé las columnas (A1:AE1)
- [ ] ¡Listo para usar! ✅

---

## 🆘 SOLUCIÓN DE PROBLEMAS

### Error: "Invalid argument: id"
- **Causa:** Aún está usando el código viejo
- **Solución:** Usa `Sistema_Paso_a_Paso_COMPLETO.gs` en su lugar

### El menú no aparece
- **Causa:** Recarga pendiente
- **Solución:** Presiona F5 para recargar la página

### No funciona la configuración
- **Causa:** Los datos no se guardaron
- **Solución:** Abre nuevamente ⚙️ Configuración y guarda de nuevo

### "Error de autenticación Kobo"
- **Causa:** Usuario/contraseña incorrectos
- **Solución:** Verifica tus credenciales en KoboToolbox

---

## 📞 SOPORTE

Si tienes problemas, abre la consola (Ctrl+Shift+I en el script):
- Click en "Ejecuciones"
- Revisa los logs para más detalles

---

**Versión:** 2.0 (Con Menú y Botones)  
**Última actualización:** 2025-05-07  
**Estado:** ✅ LISTA PARA USAR
