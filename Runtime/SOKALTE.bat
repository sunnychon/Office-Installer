@echo off
if "%1"=="-32" cd /d "%ProgramFiles(x86)%" & goto st

if "%1"=="-64" cd /d "%ProgramFiles%" & goto st

echo %0 -^<arch^> & echo arch: 32 64 & goto exi

:st

title x%1 Mode
cd "Microsoft Office"
cd "Office16"

set retry=0
set kms_server=kms.digiboy.ir
echo Using KMS server:%kms_server%
echo Setting HOST:%kms_server%
cscript //nologo ospp.vbs /sethst:%kms_server%>nul
echo Activating Office...
cscript //nologo ospp.vbs /act>nul||goto kms_failed
::install Completed and exit
echo Activation Completed.

title Activation Completed

goto exi

:kms_failed

::kms server failed

cls
set /a retry=%retry%+1
echo Server "%kms_server%" Failed,Retrying(%retry%).
::retry kms server list
if %retry%==1 set kms_server=54.223.212.31
if %retry%==2 set kms_server=hq1.chinancce.com
if %retry%==3 set kms_server=kms.cnlic.com
if %retry%==4 set kms_server=kms.chinancce.com
if %retry%==5 set kms_server=kms.ddns.net
if %retry%==6 set kms_server=franklv.ddns.net
if %retry%==7 set kms_server=k.zpale.com
if %retry%==8 set kms_server=m.zpale.com
if %retry%==9 set kms_server=mvg.zpale.com
if %retry%==10 set kms_server=kms.shuax.com
if %retry%==11 set kms_server=kensol263.imwork.net:1688
if %retry%==12 set kms_server=xykz.f3322.org
if %retry%==13 set kms_server=kms789.com
if %retry%==14 set kms_server=dimanyakms.sytes.net:1688
if %retry%==15 set kms_server=kms.03k.org:1688
if %retry%==16 set kms_server=win.kms.pub:1688
if %retry%==17 set kms_server=kms.xingez.me:1688
if %retry%==18 set kms_server=kms.000606.xyz:1688
if %retry%==19 set kms_server=kms.cyxc.club:1688
::retry kms server list end

:: if "retry" is bigger than 20 then exit.
if %retry%==20 cls&echo KMS Key INSTALL FAILED! Please Cheak Your Network.&pause>nul&goto exi

::set the server and retry
goto retry

:exi