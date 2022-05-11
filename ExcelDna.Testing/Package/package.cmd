@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg

if exist "%outputPath%\*.nupkg" del "%outputPath%\*.nupkg"

if not exist "%outputPath%" mkdir "%outputPath%"

echo on
nuget.exe pack "%basePath%\ExcelDna.Testing\ExcelDna.Testing.nuspec" -BasePath "%basePath%\ExcelDna.Testing" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive

:end
