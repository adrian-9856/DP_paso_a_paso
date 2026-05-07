# 📤 Instrucciones para HACER PUSH a GitHub

## Opción 1: Script Automático (RECOMENDADO)

### En Terminal/PowerShell de tu máquina:

```bash
# 1. Navega a la carpeta del proyecto
cd ~/DP_paso_a_paso
# O si está en otro lugar:
cd /ruta/donde/está/DP_paso_a_paso

# 2. Ejecuta el script
chmod +x push.sh        # (solo en Mac/Linux)
./push.sh               # En Mac/Linux
bash push.sh            # En Windows (Git Bash)
```

Te pedirá:
- **Username**: Tu usuario de GitHub (`adrian-9856`)
- **Password**: Tu Personal Access Token (NO tu contraseña normal)

**¿No tienes Personal Access Token?** → Crea uno aquí:
https://github.com/settings/tokens/new
- Marca: `repo` (acceso completo)
- Copia el token y úsalo como "password"

---

## Opción 2: Comandos Manuales

```bash
# 1. Abre Terminal/PowerShell en tu máquina
cd DP_paso_a_paso

# 2. Asegúrate de estar en la rama correcta
git branch
# Debe mostrar: claude/paso-a-paso-cloud-system-FENDi

# 3. Si no existe el remoto, añádelo
git remote add origin https://github.com/adrian-9856/DP_paso_a_paso.git

# 4. O si ya existe, verifica que sea HTTPS
git remote set-url origin https://github.com/adrian-9856/DP_paso_a_paso.git

# 5. Haz push
git push -u origin claude/paso-a-paso-cloud-system-FENDi
```

---

## ✅ ¿Cómo saber que funcionó?

Después del push, deberías ver:
```
...
To github.com:adrian-9856/DP_paso_a_paso.git
 * [new branch]      claude/paso-a-paso-cloud-system-FENDi -> claude/paso-a-paso-cloud-system-FENDi
Branch 'claude/paso-a-paso-cloud-system-FENDi' set up to track remote tracking branch 'claude/paso-a-paso-cloud-system-FENDi' from 'origin'.
```

Luego abre GitHub y verás los archivos en la rama.

---

## 🆘 Solución de Problemas

### Error: "permission denied"
- Verifica que usaste el Personal Access Token, NO tu contraseña
- Genera uno nuevo aquí: https://github.com/settings/tokens

### Error: "fatal: not a git repository"
- Asegúrate de estar en la carpeta correcta con `cd DP_paso_a_paso`
- Verifica que existe la carpeta `.git` con `ls -la`

### Error: "could not read Username"
- En Windows con Git Bash, a veces requiere: `git config --global credential.helper wincred`
- Luego intenta de nuevo

### Error: "Authentication failed"
- El token expiró o es inválido
- Crea uno nuevo: https://github.com/settings/tokens

---

## 📋 Checklist Final

- [ ] Estoy en la carpeta `DP_paso_a_paso`
- [ ] Veo los archivos: `README.md`, `SISTEMA_PASO_A_PASO.md`, `scripts/`
- [ ] Ejecuto `git push` o el script
- [ ] Ingreso usuario y Personal Access Token
- [ ] El push termina exitosamente
- [ ] Veo los archivos en https://github.com/adrian-9856/DP_paso_a_paso

---

**¿Necesitas ayuda?** Comparte el error exacto que ves en terminal.
