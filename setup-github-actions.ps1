# Configurações
$RepoName = "planilha-clientes"
$GitHubUsername = "G437SG" # Substitua pelo seu nome de usuário do GitHub

# Criar estrutura de diretórios para o workflow
Write-Host "Criando estrutura para GitHub Actions..." -ForegroundColor Green
New-Item -Path ".github\workflows" -ItemType Directory -Force

# Criar arquivo de workflow
$workflowContent = @"
name: Build macOS App

on:
  push:
    branches: [ main ]
  workflow_dispatch:  # Permite executar o workflow manualmente

jobs:
  build:
    runs-on: macos-latest
    
    steps:
    - name: Checkout code
      uses: actions/checkout@v3
      
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
        
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install py2app fpdf openpyxl pillow
        
    - name: Build macOS application
      run: |
        python setup_mac.py py2app
        
    - name: Zip application
      run: |
        cd dist
        zip -r "Formulario-Projeto-Arquitetonico.zip" "Formulário de Projeto Arquitetônico.app"
        
    - name: Upload compiled application
      uses: actions/upload-artifact@v3
      with:
        name: Formulario-Projeto-Arquitetonico
        path: dist/Formulario-Projeto-Arquitetonico.zip
        retention-days: 7
"@

Set-Content -Path ".github\workflows\build-macos.yml" -Value $workflowContent -Force
Write-Host "Arquivo de workflow criado com sucesso!" -ForegroundColor Green

# Instruções para commit e push
Write-Host "`nAgora execute os seguintes comandos para enviar o código ao GitHub:" -ForegroundColor Yellow
Write-Host "--------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "1. Criar um repositório no GitHub chamado '$RepoName'" -ForegroundColor Cyan
Write-Host "   Acesse: https://github.com/new" -ForegroundColor Cyan
Write-Host "`n2. Execute os comandos git para enviar seu código:" -ForegroundColor Cyan
Write-Host "   git init" -ForegroundColor White
Write-Host "   git add ." -ForegroundColor White
Write-Host "   git commit -m 'Versão inicial'" -ForegroundColor White
Write-Host "   git branch -M main" -ForegroundColor White
Write-Host "   git remote add origin https://github.com/$GitHubUsername/$RepoName.git" -ForegroundColor White
Write-Host "   git push -u origin main" -ForegroundColor White
Write-Host "`n3. Acesse o GitHub para ver o workflow em execução:" -ForegroundColor Cyan
Write-Host "   https://github.com/$GitHubUsername/$RepoName/actions" -ForegroundColor White
Write-Host "`n4. Após a conclusão do workflow, baixe o arquivo compilado:" -ForegroundColor Cyan
Write-Host "   - Clique no workflow concluído" -ForegroundColor White
Write-Host "   - Procure por 'Artifacts' na parte inferior" -ForegroundColor White
Write-Host "   - Clique em 'Formulario-Projeto-Arquitetonico' para baixar" -ForegroundColor White
Write-Host "`nO aplicativo compilado para macOS estará dentro do arquivo ZIP baixado." -ForegroundColor Green