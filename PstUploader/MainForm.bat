@echo off
set "ps1File=E:\work\ps\azcopy\PstUploader\MainForm.ps1"
powershell -WindowStyle Hidden -Command "Start-Process powershell -ArgumentList '-NoExit','-ExecutionPolicy Bypass -File ""%ps1File%""' -Verb RunAs"
