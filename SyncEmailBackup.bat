@echo off

set SOURCE=C:\Users\Loliver\Documents\emails
set DEST=C:\Users\Loliver\OneDrive\Documentation\backups\emails
set LOG=C:\Users\Loliver\OneDrive\Documentation\backups\EmailBackupSync.log

echo ================================================== >> "%LOG%"
echo Backup started %DATE% %TIME% >> "%LOG%"

robocopy "%SOURCE%" "%DEST%" /MIR /R:1 /W:1 /XJ /TEE /LOG+:"%LOG%"

echo Backup finished %DATE% %TIME% >> "%LOG%"
echo. >> "%LOG%"
