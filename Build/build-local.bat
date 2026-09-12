@echo off
setlocal

rem ===================================================================
rem  InkClass - local build (uses VS 2022 Build Tools, no IDE needed)
rem
rem  Usage:   build-local.bat                       build Debug
rem           build-local.bat Release               build Release
rem           build-local.bat Debug rebuild         force full recompile
rem           build-local.bat Debug open            open output folder
rem           build-local.bat Debug rebuild open    both
rem
rem  Output:  <repo>\Ink Canvas\bin\InkClass\InkClass.exe
rem
rem  IMPORTANT: this build is INCREMENTAL. If no source file changed since
rem  the last build, MSBuild leaves the exe untouched, so its timestamp does
rem  NOT move. That is normal - the exe is still the current build.
rem  Pass "rebuild" to force a full recompile.
rem
rem  Note: builds the .csproj directly, NOT the .sln - the solution also
rem        contains two .wapproj (MSIX packaging) projects that need the
rem        MSIX Packaging Tools component, which Build Tools usually lacks.
rem
rem  Double-click friendly: when called with NO argument the window is kept
rem  open at the end so you can actually read the result. Passing an
rem  argument (e.g. "Debug") never pauses, so scripts can call this file.
rem ===================================================================

set "ERRCODE=0"
set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Debug"

rem Optional switches, order independent, given in argument 2 and 3
set "REBUILD="
set "OPEN="
if /i "%~2"=="rebuild" set "REBUILD=1"
if /i "%~3"=="rebuild" set "REBUILD=1"
if /i "%~2"=="open"    set "OPEN=1"
if /i "%~3"=="open"    set "OPEN=1"

rem No argument = most likely a double-click: keep the window open at the end.
rem With an argument (called from a script) never pause.
set "KEEP_OPEN="
if "%~1"=="" set "KEEP_OPEN=1"

set "MSB=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"

if not exist "%MSB%" goto :no_tools

rem Step up to the repo root so the printed paths are clean and absolute
rem (otherwise they show the ugly "Build\.." form).
pushd "%~dp0.." 2>nul
if errorlevel 1 goto :no_repo

set "ROOT=%CD%"
set "PROJ=%ROOT%\Ink Canvas\Ink Canvas.csproj"
set "DIR=%ROOT%\Ink Canvas\bin\InkClass"
set "OUT=%DIR%\InkClass.exe"

if not exist "%PROJ%" goto :no_proj

rem Pre-flight: a running InkClass locks its own exe. An incremental build with
rem no source change still succeeds, but as soon as something needs copying the
rem build dies in a wall of MSB3027 / MSB3021 errors that are hard to read.
rem Warn up front - do not abort, because a no-op build is still legitimate.
rem Use findstr, NOT find: if Git for Windows is installed its unix find.exe
rem may shadow the Windows one and blow up with "find: -i: No such file".
tasklist /FI "IMAGENAME eq InkClass.exe" /NH 2>nul | findstr /i /c:"InkClass.exe" >nul
if not errorlevel 1 echo [WARN] InkClass.exe is RUNNING - its exe file is locked. If this build needs to copy the exe it WILL fail. Close InkClass first.

rem Remember the exe timestamp so we can tell "really recompiled" apart from
rem "MSBuild considered it up to date". The latter is the usual reason a
rem build looks like it did nothing.
set "T0="
for %%F in ("%OUT%") do set "T0=%%~tF"

echo [1/2] Building %CONFIG% ...
echo        Project: %PROJ%
if defined REBUILD echo        Mode   : Rebuild - full recompile, will rewrite the exe
echo.
if defined REBUILD goto :do_rebuild
"%MSB%" "%PROJ%" -p:Configuration=%CONFIG% -v:minimal -nologo
set "RC=%ERRORLEVEL%"
goto :after_build

:do_rebuild
"%MSB%" "%PROJ%" -t:Rebuild -p:Configuration=%CONFIG% -v:minimal -nologo
set "RC=%ERRORLEVEL%"

:after_build
rem Capture errorlevel immediately - do not rely on it surviving a "goto".
if not "%RC%"=="0" goto :build_failed
if not exist "%OUT%" goto :no_output

set "T1="
for %%F in ("%OUT%") do set "T1=%%~tF"

echo.
if not "%T1%"=="%T0%" goto :ok_rebuilt
if not defined REBUILD goto :ok_unchanged

echo [2/2] OK. Output file:
echo        %OUT%
echo        Time  : %T1%
goto :explain

:ok_rebuilt
echo [2/2] OK - source changed, the exe was recompiled just now.
echo        Output: %OUT%
echo        Time  : %T1%
goto :explain

:ok_unchanged
echo [2/2] OK - the exe was NOT rewritten, because nothing changed since the
echo        last build. MSBuild is incremental, so its timestamp stays put.
echo        This is normal, and the exe is still the current build.
echo.
echo        Exe  : %OUT%
echo        Time : %T1%   [unchanged]
echo.
echo        If you want a guaranteed fresh artifact - for example before
echo        packaging a release - force a full recompile with:
echo            build-local.bat %CONFIG% rebuild

:explain
echo.
echo   WHERE THINGS LIVE:
echo     * bin has THREE subfolders: Debug / Release / InkClass.
echo       The runnable build is in bin\InkClass\ - that folder is named
echo       after the product, NOT after the configuration.
echo     * This is the repo build. The INSTALLED copy at
echo       %%LOCALAPPDATA%%\Programs\InkClass is a separate folder and is
echo       NOT updated by this script.
echo.
echo   Tip: close the running InkClass.exe before building, or the copy fails.
echo        Usage: build-local.bat [Debug^|Release] [rebuild] [open]

if defined OPEN start "" "%DIR%"
goto :done

:no_tools
echo.
echo [ERROR] Visual Studio Build Tools 2022 not found at:
echo         %MSB%
echo.
echo         Install "Visual Studio Build Tools 2022" including the
echo         ".NET desktop build tools" workload.
echo         (On another machine, edit the MSB= line at the top of this file.)
set "ERRCODE=1"
goto :done

:no_repo
echo.
echo [ERROR] Could not enter the repo folder from:
echo         %~dp0
set "ERRCODE=1"
goto :done

:no_proj
echo.
echo [ERROR] Project file not found:
echo         %PROJ%
echo.
echo         This .bat must stay inside the repo's Build\ folder, so that
echo         "..\Ink Canvas\Ink Canvas.csproj" exists next to it.
echo.
echo         Current folder after step-up: %ROOT%
set "ERRCODE=1"
goto :done

:build_failed
echo.
echo [FAILED] Build failed - see the messages above.
echo.
echo   MOST COMMON CAUSE: InkClass.exe is still running, so its exe file is
echo   locked and cannot be overwritten. Look for MSB3027 / MSB3021 / MSB3026
echo   in the output above - that is exactly what they mean.
echo.
echo   Close InkClass, then run this script again. To close it from here:
echo       taskkill /IM InkClass.exe /F
set "ERRCODE=1"
goto :done

:no_output
echo.
echo [WARN] Build reported success but the output exe is missing:
echo        %OUT%
set "ERRCODE=1"
goto :done

:done
popd 2>nul
echo.
if defined KEEP_OPEN (
  echo ---- press any key to close ----
  pause >nul
)
exit /b %ERRCODE%
