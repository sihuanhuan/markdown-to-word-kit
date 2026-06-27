@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"

if "%~1"=="" (
  echo Usage:
  echo   convert.bat input.md [output.docx] [options]
  echo.
  echo Examples:
  echo   convert.bat thesis.md
  echo   convert.bat thesis.md thesis.docx
  echo   convert.bat thesis.md thesis.docx --auto-diagrams
  echo.
  echo Options:
  echo   --no-toc          Do not generate a table of contents
  echo   --auto-diagrams   Render Mermaid/PlantUML only when tools are available
  echo   --diagrams        Require Mermaid/PlantUML rendering tools
  exit /b 2
)

set "INPUT=%~1"
shift
set "OUTPUT="

if not "%~1"=="" (
  set "MAYBE_OUTPUT=%~1"
  if /I "!MAYBE_OUTPUT:~-5!"==".docx" (
    set "OUTPUT=%~1"
    shift
  )
)

py -3 --version >nul 2>nul
if errorlevel 1 (
  python --version >nul 2>nul
  if errorlevel 1 (
    echo Python 3 not found. Please install Python 3 and try again.
    exit /b 1
  )
  set "PYTHON_CMD=python"
) else (
  set "PYTHON_CMD=py -3"
)

if "%OUTPUT%"=="" (
  %PYTHON_CMD% "%SCRIPT_DIR%md2docx.py" "%INPUT%" %*
) else (
  %PYTHON_CMD% "%SCRIPT_DIR%md2docx.py" "%INPUT%" -o "%OUTPUT%" %*
)

if errorlevel 1 (
  echo.
  echo Conversion failed.
  exit /b 1
)

echo.
echo Conversion finished.
