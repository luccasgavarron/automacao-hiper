@echo off
echo =======================================================
echo Instalando dependencias da Automacao de Faturas Hyper
echo =======================================================
echo.
echo Bibliotecas necessarias:
echo  - Selenium (Para abrir o Chrome)
echo  - Webdriver-Manager (Para gerenciar a versao do Chrome)
echo  - Pywin32 (Para comunicar direto com o Excel)
echo.

pip install selenium webdriver-manager pywin32

echo.
echo =======================================================
echo Instalacao concluida! 
echo Agora voce pode executar o arquivo automacao_faturas_hyper.py
echo =======================================================
pause
