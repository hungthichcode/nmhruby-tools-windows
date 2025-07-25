@echo off
:: Kiểm tra xem script có đang chạy với quyền admin hay không
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    :: Nếu không có quyền admin, yêu cầu quyền admin
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %*", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /b
)

:: Chạy lệnh PowerShell với quyền admin
powershell -Command "irm https://massgrave.dev/get | iex"
pause
