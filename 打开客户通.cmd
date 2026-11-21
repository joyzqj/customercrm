@echo off
chcp 65001 >nul
setlocal
set "DIR=%~dp0"
if exist "%DIR%dist\客户通-外贸客户管理系统.html" (
  start "" "%DIR%dist\客户通-外贸客户管理系统.html"
) else (
  start "" "%DIR%index.html"
)
endlocal
