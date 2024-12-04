@echo off

Echo **********   Batch File for Running Python Test Framework from Console  ***********

Echo **********  All the files of Python and the other supported files should be under the same directory **********
cd "%~dp0"

mkdir "%~dp0TestLogs"

for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set DateTime=%%a

set Yr=%DateTime:~0,4%
set Mon=%DateTime:~4,2%
set Day=%DateTime:~6,2%
set Hr=%DateTime:~8,2%
set Min=%DateTime:~10,2%
set Sec=%DateTime:~12,2%

set TimeStamp=AOL_EDS(Phase1)-TestFramework__%Day%-%Mon%-%Yr%_(%Hr%.%Min%.%Sec%)

set LogFile="%~dp0TestLogs\%TimeStamp%.log"

Echo *****Execution of AOL_TestFramework has begun @ %Day%-%Mon%-%Yr%_(%Hr%-%Min%-%Sec%), please do not close this window****** >> "%LogFile%"

py.test -v --html=TestLogs\\AOL_EDS(Phase1)_Report.html >> "%LogFile%"

echo "The execution of AOL Test Frameowrk has been completed, explore to AOL_Reports.html file for further more details." >> "%LogFile%"





 