@echo off
chcp 65001 >nul
echo ====================================
echo  Excel 格式整理工具 - 打包脚本
echo ====================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python 3.8+
    pause
    exit /b
)

echo [1/3] 安装依赖...
pip install pandas openpyxl PySide6 pyinstaller -q
echo.

echo [2/3] 打包中...
pyinstaller --onefile --windowed --name "Excel格式整理工具" --add-data "core.py;." main.py
echo.

if exist "dist\Excel格式整理工具.exe" (
    echo [3/3] ✅ 打包成功！
    echo 输出文件：dist\Excel格式整理工具.exe
) else (
    echo [3/3] ❌ 打包失败
)

echo.
pause