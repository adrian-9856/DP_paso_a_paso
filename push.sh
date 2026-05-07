#!/bin/bash

# Script para hacer push del sistema "Paso a Paso" a GitHub
# Ejecuta este script en tu máquina local

echo "🚀 Iniciando push del sistema Paso a Paso..."

# 1. Cambiar a la rama correcta
echo "📌 Cambiando a rama claude/paso-a-paso-cloud-system-FENDi..."
git checkout claude/paso-a-paso-cloud-system-FENDi

# 2. Configurar remoto si no existe
echo "🔗 Verificando remoto..."
if ! git remote get-url origin &>/dev/null; then
  echo "Añadiendo remoto..."
  git remote add origin https://github.com/adrian-9856/DP_paso_a_paso.git
else
  echo "Remoto ya existe: $(git remote get-url origin)"
fi

# 3. Hacer push
echo "📤 Haciendo push a GitHub..."
git push -u origin claude/paso-a-paso-cloud-system-FENDi

if [ $? -eq 0 ]; then
  echo "✅ ¡Push completado exitosamente!"
  echo "📍 Repositorio: https://github.com/adrian-9856/DP_paso_a_paso"
  echo "🌿 Rama: claude/paso-a-paso-cloud-system-FENDi"
else
  echo "❌ Error en el push. Verifica tu conexión y credenciales."
  exit 1
fi
