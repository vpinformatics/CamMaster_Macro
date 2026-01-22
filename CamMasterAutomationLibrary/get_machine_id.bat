@echo off
echo Reading Machine ID...
echo.

reg query "HKLM\SOFTWARE\Microsoft\Cryptography" /v MachineGuid > systemid.txt

echo.
pause
