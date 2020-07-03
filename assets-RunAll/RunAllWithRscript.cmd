rem --------------------------Script Variables---------------------------------
rem [TODO] read the 'R_Script_exe' value from registry, if not found read value from RunAll.ini
rem could be... set R_Script_exe=%programfiles%\R\R-3.5.1\bin\x64\Rscript.exe
rem set R_Script_exe=%userprofile%\Documents\R\R-3.5.1\bin\x64\Rscript.exe
rem [TODO] read the 'R_Script_Name' value from the RunAll.ini file

set R_Script_exe=%userprofile%\Documents\R-3.5.1\bin\x64\Rscript.exe
set R_Script_Name=RunAll.R
set /A display_error_seconds=10

rem -------------------------Script properties-----------------------------------
SETLOCAL EnableExtensions
@echo off
color b0
cls

rem ---------------------------LICENSE------------------------------------------- 
rem Authored by jeremy.gerdes@navy.mil 
rem CC0 1.0 Universal (CC0 1.0) 
rem Public Domain Dedication 
rem https://creativecommons.org/publicdomain/zero/1.0/legalcode

rem ----------------------------README------------------------------------------- 
rem This .cmd script looks for the expected instal path for Rscript.exe 
rem and calls Rscript.exe with the argument of the relative path to RunAll.R 
rem this script does not requires command extensions to be enabled to use 
rem the dynamic environment variable expansion of %__CD__% 

rem ------------------------------Script-----------------------------------------
rem In batch scripting nested branches of if statements must have syntax that 
rem resolves so we are forced to use goto statements or a function as we are 
rem going to be asking for user input with 'set \p...'
rem additionally nesting more than two if statements deep seems to fail? 
rem but I haven't tested this... could have just been syntax issues

if exist "%R_Script_exe%" (
    rem @echo Found Rscript.exe at %R_Script_exe%
) else (
    if exist "%R_Script_exe%" (
       rem @echo Found Rscript.exe at %R_Script_exe%
    ) else (
       set user_prompt=Enter the folder path of the file Rscript.exe:
       CALL :ask_user_dir R_Script_exe_Directory, %user_prompt%, %display_error_seconds%
       set R_Script_exe=%R_Script_exe_Directory%\Rscript.exe
   )
)

rem set the expected script path to current working directory's parent

echo on
call :GetDirParentN R_Script_Path %__CD__% ".."
echo %__CD__%
echo %R_Script_Path%
pause

if exist "%R_Script_Path%\%R_Script_Name%" (
    set R_Script_to_Run=%R_Script_Path%\%R_Script_Name%
    rem @echo using relative path method for expected script directory location
REM [TODO] should prompt user for path, but function is failing for some reason...
REM ) else (
REM     set user_prompt=Enter the folder path of project that contains %R_Script_Name% (without trailing backslash):
REM     CALL :ask_user_dir R_Scripts_dir, %user_prompt%, %display_error_seconds%
REM     set R_Script_to_Run=%R_Scripts_dir%\%R_Script_Name%
)
if exist "%R_Script_exe%" (
    if exist "%R_Script_to_Run%" (
        rem ---Execute R Script---
        @echo Attempting to launch 'Rscript.exe %R_Script_Name%' with:
        @echo "%R_Script_exe%" "%R_Script_to_Run%"
        "%R_Script_exe%" "%R_Script_to_Run%"
    )   
    rem let the user see whats going on
    @echo %R_Scripts_dir%
    @echo closing in %display_error_seconds% seconds
    CALL :wait %display_error_seconds% 
) else (
    @echo Failed to find Rscript.exe at: "%R_Script_exe%"
    @echo Press any key to exit:
    pause
)
EXIT /B %ERRORLEVEL%

:setVariableValue 
    set %~1=%~2
EXIT /B 0

:ask_user_dir
    rem Example function call:
    rem CALL :ask_user_dir R_Script_exe, user_prompt, is_dir, display_error_seconds
    rem example params...
    rem %~1 is R_Script_exe = ...return some path we are asking for...
    rem %~2 is user_prompt = "Enter the location of Rscript.exe"
    rem %~3 is display_error_seconds = 10
    
    @echo expected location not found if you know it;
    set /p %~1=%~2
    rem with user input we need to see if this is a directory or a file
    CALL :is_directory %~1, is_dir
    if %is_dir% neq 1 (
        if %is_dir% equ 2 (
            @echo Exiting... you entered a file when a directory was expected
            @echo closing in %~3 seconds
            CALL :wait %~3
            exit
        ) else (
             @echo Exiting... could not find directory at: %~1
             @echo closing in %~3 seconds
             CALL :wait %~3
             exit
        )
    ) 
EXIT /B 0

:is_directory
    rem fuction to check if parameter is a directory
    rem sets second parameter passed to 0 for doesn't exist, 1 for is a direcory and 2 for is a file but not a directory
    set ATTR=%~a1
    set DIRATTR=%ATTR:~0,1%
    if /I "%DIRATTR%"=="d" (
        set /A %~2=1
    ) else (
        if exist %~1 (
            set /A %~2=2
        ) else (
            set /A %~2=0
        )
    )
EXIT /B 0

:wait
    rem wait function expects an integer value and delays by that many seconds by:
    rem then pings the loopback ip, delaying one second between each ping
    rem this function can be replaced by using TIMEOUT.exe see https://ss64.com/nt/timeout.html 
    set /A wait_delay="%~1 + 1"
    ping -n %wait_delay% 127.0.0.1 > nul
EXIT /B 0



:GetDirParentN
    for %%I in ("%~2\%~3") do set "%~1=%%~fI"
EXIT /B 0
