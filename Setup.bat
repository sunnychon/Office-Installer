@echo off

::Start
echo [%date% %time%]Starting...>>event.log

::Chk Admin
echo [%date% %time%]Cheaking Admin...>>event.log

::Get Admin If Not Admin
PUSHD "%~DP0" && CD /D "%~DP0"
IF EXIST "%Public%" >NUL 2>&1 REG QUERY "HKU\S-1-5-19\Environment"
IF NOT %errorlevel% EQU 0 (
IF EXIST "%Public%"  echo [%date% %time%]No Admin! Getting Admin Permissions...>>event.log & powershell.exe -windowstyle hidden -noprofile "Start-Process '%~dpnx0' -Verb RunAs & exit")

echo [%date% %time%]Admin Success!>>event.log

::Chk ARCH
if %PROCESSOR_ARCHITECTURE%==AMD64 set arch=64
if %PROCESSOR_ARCHITECTURE%==ARM64 set arch=64
if %PROCESSOR_ARCHITECTURE%==x86 set arch=32
echo [%date% %time%]Architecture is x%arch%

::Chk Files

::1

echo [%date% %time%]Cheaking Files...(1/3)>>event.log
IF NOT EXIST .\Runtime\setup.exe set err=nf1 goto err

::2

echo [%date% %time%]Cheaking Files...(2/3)>>event.log
IF NOT EXIST PROx%arch%.xml set err=nf2 & goto err

::3

echo [%date% %time%]Cheaking Files...(3/3)>>event.log
IF NOT EXIST .\Runtime\SOKALTE.bat set err=nf3 & goto err

::Down Files
echo [%date% %time%]Downloading Files...>>event.log
.\Runtime\setup.exe /download PROx%arch%.xml

::Inst Office
echo [%date% %time%]Installing...>>event.log
.\Runtime\setup.exe /configure PROx%arch%.xml

::Act Office
echo [%date% %time%]Activating...>>event.log
call .\Runtime\SOKALTE.bat -%arch%
wscript "%~dp0\done.vbs"

::run after install::



:::::::::::::::::::::
goto ext

:err
if %err%==nf1 set errmsg=setup.exe Not Found
if %err%==nf2 set errmsg=PROx%arch%.xml Not Found
if %err%==nf3 set errmsg=SOKALTE.bat Not Found
echo [%date% %time%]ERROR:%err%(%errmsg%)>>event.log

:ext


