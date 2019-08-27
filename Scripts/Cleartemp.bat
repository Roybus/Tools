:: --- BATCH SCRIPT START ---
@echo off
DEL /S /Q /F "%TEMP%\*.*"
FOR /D %%d IN ("%TEMP%\*.*") DO RD /S /Q "%%d"

powershell.exe -command remove-item c:\Windows\temp\* -recurse
:: --- BATCH SCRIPT END ---