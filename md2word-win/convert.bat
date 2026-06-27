@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
set "LOCAL_PYTHON=%SCRIPT_DIR%_md2word\tools\python\python.exe"

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

set "EXTRA_ARGS="
:collect_args
if "%~1"=="" goto args_done
set "EXTRA_ARGS=!EXTRA_ARGS! "%~1""
shift
goto collect_args
:args_done

if exist "%LOCAL_PYTHON%" (
  set "PYTHON_EXE=%LOCAL_PYTHON%"
  set "PYTHON_ARGS="
) else (
  py -3 --version >nul 2>nul
  if errorlevel 1 (
    python --version >nul 2>nul
    if errorlevel 1 (
      echo Python 3 not found. Run _md2word\install-tools.ps1 or install Python 3 and try again.
      exit /b 1
    )
    set "PYTHON_EXE=python"
    set "PYTHON_ARGS="
  ) else (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3"
  )
)

if "%OUTPUT%"=="" (
  "%PYTHON_EXE%" %PYTHON_ARGS% "%SCRIPT_DIR%_md2word\md2docx.py" "%INPUT%" !EXTRA_ARGS!
) else (
  "%PYTHON_EXE%" %PYTHON_ARGS% "%SCRIPT_DIR%_md2word\md2docx.py" "%INPUT%" -o "%OUTPUT%" !EXTRA_ARGS!
)

if errorlevel 1 (
  echo.
  echo Conversion failed.
  exit /b 1
)

echo.
echo Conversion finished.
