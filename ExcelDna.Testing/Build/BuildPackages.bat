setlocal

set PackageVersion=%1
set MSBuildPath=%2

%MSBuildPath% ..\ExcelDna.Testing.sln /t:restore,build /p:Configuration=Release
@if errorlevel 1 goto end

cd ..\Package
call package.cmd %PackageVersion%
@if errorlevel 1 goto end

:end
