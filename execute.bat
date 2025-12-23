@echo off
:: 1. Entra na pasta do projeto
cd /d "C:\Users\roborpa\RPA\outlookRpa"

:: 2. Ativa o ambiente virtual (Se estiver usando venv)
call venv\Scripts\activate

:: 3. Executa o robô
python main.py

:: 4. (Opcional) Desativa o venv
@REM deactivate