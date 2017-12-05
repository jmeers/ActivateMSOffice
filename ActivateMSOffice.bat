::James Meers
@echo off
::Office11 = 2003, Office12 = 2007, Office14 = 2010, Office15 = 2013
set OfficeVer="Office14"
::Check for Administrator Permissions
echo ALERT: Administrator permissions Required. 
echo Detecting permissions...
net session >nul 2>&1
if %errorLevel% == 0 (
        echo SUCCESS: Administrative permissions Accepted!
) else (
        echo FAILURE: Current permissions Insufficient.
	echo Run Batch as Administrator!
	timeout /t 5 /nobreak
	exit
)
::Check to see if Office is LICENSED
cscript "%SYSTEMDRIVE%\Program Files (x86)\Microsoft Office\%OfficeVer%\ospp.vbs" /dstatus | find /i "---LICENSED---"
::Doesn't find LICENSE
if errorlevel 1 (
echo Adding Key...
cscript "%SYSTEMDRIVE%\Program Files (x86)\Microsoft Office\%OfficeVer%\ospp.vbs" /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
::Open Word to Accept Activation of Office
"%SYSTEMDRIVE%\Program Files (x86)\Microsoft Office\%OfficeVer%\WINWORD.exe"
) else (
::Finds LICENSE
echo Working key is already been added!!
)
::Quick Time out Exit
timeout /t 5 /nobreak
