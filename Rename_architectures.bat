@echo off
setlocal enabledelayedexpansion

:: Define the base directory
set BASE_DIR=.\bin\publish\Fd

:: Check if the base directory exists
if not exist "%BASE_DIR%" (
    echo Directory "%BASE_DIR%" does not exist.
    exit /b
)

:: Start recursive processing from the base directory
echo Starting recursive file renaming in %BASE_DIR%

:: Loop through all subdirectories and files recursively
for /r "%BASE_DIR%" %%d in (.) do (
    if "%%~fd" neq "%BASE_DIR%" (
        set FOLDER_NAME=%%~nxd

        :: Loop through all files in the current directory
        for %%f in (%%d\*) do (
            set FILE_NAME=%%~nxf
            set FILE_BASE=%%~nf
            set EXT=%%~xf
            set NEW_NAME=%%~dpnxf

            :: Append the folder name before the extension
            set NEW_NAME=!FILE_BASE!_!FOLDER_NAME!!EXT!

            :: Rename the file
            ren "%%f" "!NEW_NAME!"
            echo Renamed "%%f" to "!NEW_NAME!"
        )
    )
)

echo File renaming complete.
endlocal
pause
