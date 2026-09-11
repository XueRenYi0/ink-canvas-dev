@echo off
setlocal

rem ===================================================================
rem  Inkboard - local build (uses VS 2022 Build Tools, no IDE needed)
rem
rem  Usage:   build-local.bat            build Debug
rem           build-local.bat Release    build Release
rem
rem  Output:  ..\Ink Canvas\bin\Inkboard\Inkboard.exe
rem
rem  Note: builds the .csproj directly, NOT the .sln - the solution also
rem        contains two .wapproj (MSIX packaging) projects that need the
rem        MSIX Packaging Tools component, which Build Tools usually lacks.
rem ===================================================================

set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Debug"

set "MSB=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
set "PROJ=%~dp0..\Ink Canvas\Ink Canvas.csproj"
set "OUT=%~dp0..\Ink Canvas\bin\Inkboard\Inkboard.exe"

if not exist "%MSB%"  goto :no_tools
if not exist "%PROJ%" goto :no_proj

echo [1/2] Building %CONFIG% ...
"%MSB%" "%PROJ%" -p:Configuration=%CONFIG% -v:minimal -nologo
if errorlevel 1 goto :build_failed

echo.
echo [2/2] Done. Output:
echo       %OUT%
echo.
echo   Tip: close the running Inkboard.exe before building, or the copy fails.
goto :done

:no_tools
echo.
echo [ERROR] Build Tools not found at:
echo         %MSB%
echo.
echo         Install "Visual Studio Build Tools 2022" including the
echo         ".NET desktop build tools" workload.
exit /b 1

:no_proj
echo.
echo [ERROR] Project file not found:
echo         %PROJ%
echo.
echo         Make sure this .bat stays inside the repo's Build\ folder.
exit /b 1

:build_failed
echo.
echo [FAILED] Build failed - see the messages above.
exit /b 1

:done
