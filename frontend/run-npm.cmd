@echo off
setlocal
set "NODE_DIR=%ProgramFiles%\nodejs"
if not exist "%NODE_DIR%\npm.cmd" (
  echo Node.js not found at %NODE_DIR%
  echo Install from https://nodejs.org or add Node to your PATH.
  exit /b 1
)
set "PATH=%NODE_DIR%;%PATH%"
cd /d "%~dp0"
call "%NODE_DIR%\npm.cmd" %*
