#!/bin/bash
# filepath: c:\Users\PC\Desktop\PROMPTS ARCHICAD\PLANILHA CLIENTES\run_mac.command

# Define o diretório do script como diretório de trabalho
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

# Verifica se Python 3 está instalado
if ! command -v python3 &> /dev/null; then
    osascript -e 'tell app "System Events" to display dialog "Python 3 não está instalado. Por favor instale Python 3 para continuar." buttons {"OK"} default button 1 with icon stop with title "Erro"'
    exit 1
fi

# Verifica se o ambiente virtual existe, caso contrário, cria um
if [ ! -d "$DIR/venv" ]; then
    echo "Criando ambiente virtual..."
    python3 -m venv venv
fi

# Ativa o ambiente virtual
source venv/bin/activate

# Instala dependências necessárias
echo "Verificando dependências..."
pip install fpdf openpyxl pillow

# Cria um link simbólico para o logo se estiver faltando
if [ ! -f "logo_empresa.png" ]; then
    echo "Criando link simbólico para o logo_empresa.png..."
    # Você precisa colocar um logo padrão aqui ou criar um em branco
    touch logo_empresa.png
fi

# Modifica o código do main.py para usar caminhos relativos ao macOS
sed -i '' 's|r"c:\\Users\\PC\\Desktop\\PROMPTS ARCHICAD\\PLANILHA CLIENTES\\logo_empresa.png"|"logo_empresa.png"|g' main.py

# Executa o programa
echo "Iniciando o aplicativo..."
python3 main.py