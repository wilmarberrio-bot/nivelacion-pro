@echo off
title Nivelacion Web App
echo Iniciando interfaz web...
echo.
echo Verificando dependencias...
echo Intentando instalar Flask y librerias necesarias...
python -m pip install flask openpyxl pandas
if %errorlevel% neq 0 (
    echo.
    echo ERROR CRITICO: No se pudo instalar Flask.
    echo Asegurese de que Python esta instalado correctamente.
    echo Intentando ejecutar de todas formas...
)

echo.
echo ===================================================
echo        Iniciando Servidor de Nivelacion
echo ===================================================
echo.
echo Por favor espere, se abrira su navegador automaticamente...
echo.
python app.py
pause
