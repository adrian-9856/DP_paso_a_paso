# 🔑 CÓMO OBTENER TU API KEY DE KOBO

## PASO 1: Acceder a KoboToolbox

1. Ve a: **https://kf.kobotoolbox.org/**
2. Login con tus credenciales (usuario y contraseña)

![Step 1: Login a Kobo](https://i.imgur.com/example.png)

---

## PASO 2: Ir a Configuración de Cuenta

1. Una vez dentro, haz click en tu **perfil** (arriba a la derecha)
2. Selecciona: **"Settings"** o **"Configuración"**

---

## PASO 3: Obtener el Token (API Key)

1. En el menú lateral, busca: **"API Token"** o **"Token de API"**
2. Verás una sección que dice: **"Your API Token"**
3. Copiar el token - Se verá algo así:

```
64cc018b88067397addd36b09288be8b6539cf39
```

---

## PASO 4: Guardar Seguro

⚠️ **IMPORTANTE**: Este token es como tu contraseña
- No lo compartas con nadie
- No lo publiques en GitHub
- Guárdalo en un lugar seguro

---

## Si Necesitas Generar uno Nuevo

Si perdiste tu token o quieres uno nuevo:

1. En la misma sección "API Token"
2. Busca el botón: **"Regenerate"** o **"Regenerar"**
3. Confirma la acción
4. Copia el nuevo token
5. El token viejo ya no funcionará

---

## Usar el Token en Paso a Paso

Una vez que tengas tu API Key:

1. Abre tu Sheet "Paso a Paso - Sistema Maestro"
2. Menú: **📊 PASO A PASO > ⚙️ Configuración**
3. En el campo **"Clave API de KoboToolbox"**, pega tu token
4. Haz click en **💾 GUARDAR**
5. Haz click en **🧪 PROBAR CONEXIÓN**

Si ves ✅ **"Conexión EXITOSA"**, ¡está bien!

---

## Verificar que funciona

El sistema te dirá cuántos registros tiene tu formulario:

```
✅ Conexión EXITOSA

📊 127 respuestas encontradas
```

---

## ¿Qué si sigue dando Error?

### Error 401
**Causa**: Token inválido o expirado
**Solución**:
1. Ve a https://kf.kobotoolbox.org/admin/auth/token/
2. Regenera un nuevo token
3. Vuelve a pegar en la configuración

### Error en la conexión
**Causa**: Problema de internet o servidor de Kobo
**Solución**: Intenta de nuevo en unos minutos

---

## Token Perdido o Expirado

Si no recuerdas tu token:
1. Ve a https://kf.kobotoolbox.org/admin/auth/token/
2. El token está en tu cuenta
3. Si no está, genera uno nuevo

---

**¡Listo!** Ya tienes tu API Key y puedes usar Paso a Paso. 🚀
