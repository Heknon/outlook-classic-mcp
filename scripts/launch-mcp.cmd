@echo off
rem ---------------------------------------------------------------------
rem outlook-classic-mcp launcher (used by the Claude Code plugin)
rem
rem   launch-mcp.cmd [run] [home]       start the MCP server (stdio)
rem   launch-mcp.cmd prepare [home]     install/update everything, then exit
rem
rem Installs what the server needs on first use, per user, no admin:
rem   1. uv (Astral) if it isn't installed -- uv then fetches a suitable
rem      Python and the dependencies (pywin32, mcp, pydantic) into
rem      %LOCALAPPDATA%\outlook-classic-mcp\venv;
rem   2. if uv can't be installed (blocked download / script policy), an
rem      existing Python 3.10+ ("py" or "python") is used instead.
rem
rem STDOUT IS THE MCP JSON-RPC STREAM: everything printed here goes to
rem stderr (1>&2). Don't add plain "echo" lines.
rem
rem Environment knobs:
rem   OUTLOOK_MCP_NO_AUTO_INSTALL=1  never download uv
rem   OUTLOOK_MCP_RUNTIME=python     skip uv, use an installed Python 3.10+
rem   OUTLOOK_MCP_HOME=<dir>         where the environment lives
rem ---------------------------------------------------------------------
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0.."
for %%I in ("%ROOT%") do set "ROOT=%%~fI"
rem Optional 2nd argument: the environment folder (the plugin passes its
rem persistent ${CLAUDE_PLUGIN_DATA}); otherwise OUTLOOK_MCP_HOME, else
rem %LOCALAPPDATA%\outlook-classic-mcp.
if not "%~2"=="" set "OUTLOOK_MCP_HOME=%~2"
if not defined OUTLOOK_MCP_HOME set "OUTLOOK_MCP_HOME=%LOCALAPPDATA%\outlook-classic-mcp"
if not exist "%OUTLOOK_MCP_HOME%" mkdir "%OUTLOOK_MCP_HOME%" 1>&2
set "UV_PROJECT_ENVIRONMENT=%OUTLOOK_MCP_HOME%\venv"
rem Corporate networks often intercept TLS with their own root CA, which
rem lives in the Windows certificate store: make uv trust it.
set "UV_NATIVE_TLS=1"
set "UV_SYSTEM_CERTS=1"
set "MODE=%~1"

if /i "%OUTLOOK_MCP_RUNTIME%"=="python" goto with_python
set "UV="
call :find_uv
if not defined UV if not "%OUTLOOK_MCP_NO_AUTO_INSTALL%"=="1" call :install_uv
if defined UV goto with_uv
goto with_python

rem ---------------------------------------------------------------- uv
:with_uv
if /i "%MODE%"=="prepare" goto uv_prepare
"%UV%" run --project "%ROOT%" --frozen --no-dev --quiet outlook-mcp
exit /b %errorlevel%

:uv_prepare
echo [outlook-mcp] Preparing the Outlook MCP server environment... 1>&2
"%UV%" sync --project "%ROOT%" --frozen --no-dev --quiet 1>&2
if errorlevel 1 (
  echo [outlook-mcp] uv sync failed; trying an installed Python instead. 1>&2
  goto with_python
)
echo [outlook-mcp] Ready. 1>&2
exit /b 0

rem ------------------------------------------------- fallback: Python
:with_python
set "PYEXE="
call :find_python
if not defined PYEXE goto no_runtime
set "VENV=%OUTLOOK_MCP_HOME%\pyvenv"
set "VPY=%VENV%\Scripts\python.exe"
if not exist "%VPY%" (
  echo [outlook-mcp] Creating a Python environment in !VENV! ... 1>&2
  %PYEXE% -m venv "%VENV%" 1>&2
  if errorlevel 1 goto no_runtime
)
rem Reinstall when the plugin version (pyproject.toml) changes.
set "WANT="
for /f "usebackq tokens=*" %%L in (`findstr /b /c:"version = " "%ROOT%\pyproject.toml"`) do if not defined WANT set "WANT=%%L"
set "HAVE="
if exist "%VENV%\installed.txt" set /p HAVE=<"%VENV%\installed.txt"
if not "!HAVE!"=="!WANT! !ROOT!" (
  echo [outlook-mcp] Installing dependencies with pip ^(one time^)... 1>&2
  "%VPY%" -m pip install --disable-pip-version-check --quiet --upgrade "%ROOT%" 1>&2
  if errorlevel 1 goto pip_failed
  >"%VENV%\installed.txt" echo !WANT! !ROOT!
)
if /i "%MODE%"=="prepare" (
  echo [outlook-mcp] Ready. 1>&2
  exit /b 0
)
"%VPY%" -m outlook_mcp
exit /b %errorlevel%

:pip_failed
echo [outlook-mcp] pip could not install the dependencies. If you are behind a 1>&2
echo [outlook-mcp] proxy, set HTTPS_PROXY; or ask IT to allow pypi.org. 1>&2
exit /b 1

:no_runtime
echo [outlook-mcp] Could not set up the server: uv could not be installed and no 1>&2
echo [outlook-mcp] Python 3.10+ was found. Install one of them and restart: 1>&2
echo [outlook-mcp]   uv:     powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 ^| iex" 1>&2
echo [outlook-mcp]   Python: https://www.python.org/downloads/windows/ (tick "Add to PATH") 1>&2
exit /b 1

rem ---------------------------------------------------------- helpers
:find_uv
for /f "delims=" %%U in ('where uv.exe 2^>nul') do if not defined UV set "UV=%%U"
if not defined UV if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not defined UV if exist "%LOCALAPPDATA%\Microsoft\WinGet\Links\uv.exe" set "UV=%LOCALAPPDATA%\Microsoft\WinGet\Links\uv.exe"
if not defined UV if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV=%USERPROFILE%\.cargo\bin\uv.exe"
exit /b 0

:install_uv
echo [outlook-mcp] uv not found; installing it for your user from astral.sh (one time)... 1>&2
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex" 1>&2 2>&1
call :find_uv
if not defined UV echo [outlook-mcp] uv install did not complete; falling back to Python. 1>&2
exit /b 0

:find_python
for %%C in ("py -3" "python") do (
  if not defined PYEXE (
    %%~C -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
    if not errorlevel 1 set "PYEXE=%%~C"
  )
)
exit /b 0
