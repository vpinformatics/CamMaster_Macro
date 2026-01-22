@echo off
setlocal ENABLEDELAYEDEXPANSION

echo =========================================
echo   VP CAMMaster License Generator
echo =========================================
echo.

set /p MACHINEID=Enter Machine ID: 

if "%MACHINEID%"=="" (
    echo Machine ID cannot be empty.
    pause
    exit /b
)

REM ===== SECRET (must match DLL exactly) =====
set SECRET=VP@2025#CamMaster$Internal

REM ===== SHA256(MachineID + SECRET) =====
for /f %%H in ('
    powershell -NoProfile -Command ^
    "$s='%MACHINEID%%SECRET%';" ^
    "$b=[Text.Encoding]::UTF8.GetBytes($s);" ^
    "$h=[Security.Cryptography.SHA256]::Create().ComputeHash($b);" ^
    "($h|ForEach-Object{ $_.ToString('x2') }) -join ''"
') do set HASH=%%H

set LICENSEDIR=C:\ProgramData\VPInformatics\CamMaster
if not exist "%LICENSEDIR%" mkdir "%LICENSEDIR%"

(
echo VP-CAMMASTER
echo %HASH%
) > "%LICENSEDIR%\license.key"

echo.
echo License generated successfully.
echo Restart CAMMaster.
pause
