@echo off
setlocal

set "drive=C:\"  REM Specify the drive to search for PST files

echo Searching for PST files in %drive%...

set "totalFiles=0"
set "totalSize=0"

set "hostname=%computername%"  REM Get the hostname of the computer

echo "PST File","File Size (bytes)","Profile Name","IP Address" > %hostname%.csv  REM Change output.csv to hostname.csv

REM Add additional paths to search
set "additionalPaths=C:\path1 C:\path2"

REM Search in the specified drive and additional paths
for %%d in (%drive% %additionalPaths%) do (
    for /r "%%d" %%f in (*.pst) do (
        echo Found PST file: %%f
        set /a "totalFiles+=1"
        for /f %%A in ('dir /-c "%%f" ^| findstr /c:"File(s)"') do (
            set "fileSize=%%A"
            set "fileSize=!fileSize:~0,-6!"
            set /a "totalSize+=fileSize"
            echo File Size: !fileSize!
            REM Get the profile name
            for /f "tokens=2 delims=\" %%P in ("%%f") do (
                set "profileName=%%P"
                echo Profile Name: !profileName!
                REM Get the IP address
                for /f "tokens=2 delims=:" %%I in ('ipconfig ^| findstr /i "IPv4 Address"') do (
                    set "ipAddress=%%I"
                    echo IP Address: !ipAddress!
                    echo "%%f","!fileSize!","!profileName!","!ipAddress!" >> %hostname%.csv  REM Change output.csv to hostname.csv
                )
            )
        )
        REM Add your desired actions here for each found PST file
    )
)

echo Total Files: %totalFiles%
echo Total Size: %totalSize% bytes

echo Search complete.

endlocal
