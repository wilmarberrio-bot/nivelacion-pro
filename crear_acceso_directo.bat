@echo off
set "target=%~dp0iniciar_web.bat"
set "shortcut=%UserProfile%\Desktop\Nivelacion App.lnk"
set "icon=%SystemRoot%\System32\shell32.dll,13"

echo Creando acceso directo en el Escritorio...
echo Target: %target%
echo Shortcut: %shortcut%

powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%shortcut%');$s.TargetPath='%target%';$s.IconLocation='%icon%';$s.WorkingDirectory='%~dp0';$s.Save()"

if exist "%shortcut%" (
    echo.
    echo [EXITO] Acceso directo creado en el Escritorio.
) else (
    echo.
    echo [ERROR] No se pudo crear el acceso directo.
)
pause
