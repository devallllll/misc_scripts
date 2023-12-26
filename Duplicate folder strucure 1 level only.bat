@echo off
SETLOCAL EnableDelayedExpansion

:: Copyright notice
echo Copyright David Lane 2023. All rights reserved.

:: Prompt the user to enter the source and destination folder paths
SET /P source=Enter the path of the source folder: 
SET /P dest=Enter the path of the destination folder: 

:: Check if the source and destination paths are not empty
IF "!source!"=="" GOTO End
IF "!dest!"=="" GOTO End

:: Check if paths ends with a backslash and add one if not (DISABLED DOESNT WORK!)
::IF NOT "%source:~-1%"=="\" SET source=%source%\
::IF NOT "%dest:~-1%"=="\" SET dest=%dest%\

:: Check if the destination folder already exists, create it if not
IF NOT EXIST "%dest%" (
    :: Create the destination folder
    mkdir "%dest%"
    echo Created destination folder: %dest%
)

:: Display the robocopy command for verification
echo Ready to execute:
echo robocopy "%source%" "%dest%" /E /LEV:2 /XF * /NJH /NJS /NDL /NC /NS /NP
pause

:: Perform the copy operation with robocopy
robocopy "%source%" "%dest%" /E /LEV:2 /XF * /NJH /NJS /NDL /NC /NS /NP

:: Explanation of robocopy options:
:: /E - Copy all subdirectories (including empty ones)
:: /LEV:1 - Only copy the top level of the source directory tree
:: /XF * - Exclude all files from being copied
:: /NJH /NJS /NDL /NC /NS /NP - Suppress various elements of the output

:End
:: Display a message upon completion
echo Operation completed.
ENDLOCAL
