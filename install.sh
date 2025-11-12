#!/bin/bash
# Script de instalação das dependências

echo "📦 Instalando dependências do projeto..."

# Tentar diferentes formas de pip
if command -v pip3 &> /dev/null; then
    PIP_CMD="pip3"
elif command -v pip &> /dev/null; then
    PIP_CMD="pip"
elif command -v python3 -m pip &> /dev/null; then
    PIP_CMD="python3 -m pip"
else
    echo "❌ pip não encontrado. Por favor, instale o pip primeiro."
    exit 1
fi

echo "Usando: $PIP_CMD"

$PIP_CMD install Flask==3.0.0 python-docx==1.1.0 beautifulsoup4==4.12.2 lxml==4.9.3 num2words==0.5.13 pandas==2.1.3 openpyxl==3.1.2 pyproj==3.6.1 Werkzeug==3.0.1

if [ $? -eq 0 ]; then
    echo "✅ Dependências instaladas com sucesso!"
    echo ""
    echo "Para rodar o servidor, execute:"
    echo "  python3 app.py"
else
    echo "❌ Erro ao instalar dependências"
    exit 1
fi


