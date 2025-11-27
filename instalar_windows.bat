@echo off
REM Instalador automático Merge F42 para Windows
echo ========================================
echo   Instalando Merge F42...
echo ========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no esta instalado
    echo.
    echo Descarga Python desde: https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion
    pause
    exit /b 1
)

REM Instalar dependencias
echo Instalando pandas y openpyxl...
python -m pip install --quiet pandas openpyxl

REM Crear acceso directo en el escritorio
echo Creando acceso directo en el escritorio...

REM Obtener ruta del escritorio
set DESKTOP=%USERPROFILE%\Desktop

REM Crear archivo .bat en el escritorio
(
echo @echo off
echo cd /d "%~dp0"
echo python gui_merge.py
echo if errorlevel 1 pause
) > "%DESKTOP%\Merge F42.bat"

REM Copiar gui_merge.py al escritorio
copy /Y gui_merge.py "%DESKTOP%\gui_merge.py" >nul

echo.
echo ========================================
echo   Instalacion completada!
echo ========================================
echo.
echo El acceso directo "Merge F42.bat" esta en tu escritorio
echo Doble clic para usar la aplicacion
echo.
echo El archivo F42_MERGED.xlsx se guardara en el escritorio
echo.
pause
