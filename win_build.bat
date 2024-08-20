@echo off
SET SOURCE_SCRIPT=.\WinPin.py
SET ICON_PATH=.\image\favicon.ico
SET OUTPUT_DIR=.\dist
SET DESTINATION=%USERPROFILE%\Desktop\WinPin.exe

:: 使用 PyInstaller 打包 Python 脚本
pyinstaller --onefile --windowed --icon=%ICON_PATH% %SOURCE_SCRIPT%

:: 检查打包是否成功
IF ERRORLEVEL 1 (
    ECHO PyInstaller failed to build the application.
    PAUSE
    EXIT /B
)

:: 复制打包好的文件到桌面
echo Copying the executable to the desktop...
copy /Y "%OUTPUT_DIR%\WinPin.exe" "%DESTINATION%"

echo Done.