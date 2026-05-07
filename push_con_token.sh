#!/bin/bash

# Script para hacer push automático con Personal Access Token
# Uso: bash push_con_token.sh <tu_token>

TOKEN="${1:-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}"
REPO="https://adrian-9856:${TOKEN}@github.com/adrian-9856/DP_paso_a_paso.git"

echo "🚀 Iniciando push del sistema Paso a Paso con token..."

# 1. Cambiar a la rama correcta
echo "📌 Rama actual:"
git branch

# 2. Configurar remoto con token
echo "🔗 Configurando remoto con autenticación..."
git remote set-url origin "$REPO"

# 3. Hacer push
echo "📤 Haciendo push a GitHub..."
git push -u origin claude/paso-a-paso-cloud-system-FENDi

if [ $? -eq 0 ]; then
  echo ""
  echo "✅ ¡¡¡PUSH COMPLETADO EXITOSAMENTE!!!"
  echo ""
  echo "📍 Repositorio: https://github.com/adrian-9856/DP_paso_a_paso"
  echo "🌿 Rama: claude/paso-a-paso-cloud-system-FENDi"
  echo ""
  echo "🎉 Ahora puedes ver todos los archivos en GitHub!"
else
  echo "❌ Error en el push. Verifica tu conexión."
  exit 1
fi
