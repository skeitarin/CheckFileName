echo off
set n=0
 
for %%f in ( *.txt ) do call :copyFile %%f
exit /b
 
:copyFile
 
if "%n%"=="100" (goto :finish)
set /a n=n+1
set /a exp=1000+n
copy %1 %~n1%exp:~1%%~x1
 
goto :copyFile
 
:finish
set n=0
goto :EOF