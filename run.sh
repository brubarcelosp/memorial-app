#!/bin/bash
# Script para rodar o servidor

echo "🚀 Iniciando servidor Flask..."
echo "📝 Acesse: http://localhost:5000"
echo ""

cd "$(dirname "$0")"
python3 app.py


