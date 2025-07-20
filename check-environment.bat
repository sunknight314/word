@echo off
chcp 65001 >nul

echo ======================================
echo 环境检查工具 - Windows
echo ======================================
echo.

echo [系统信息]
echo 操作系统: %OS%
echo 计算机名: %COMPUTERNAME%
echo 用户名: %USERNAME%
echo.

echo [Python环境]
python --version 2>nul
if errorlevel 1 (
    echo Python: [未安装]
) else (
    echo Python路径: 
    where python
    echo.
    echo pip版本:
    pip --version
)
echo.

echo [Node.js环境]
node --version 2>nul
if errorlevel 1 (
    echo Node.js: [未安装]
) else (
    echo Node.js路径:
    where node
    echo.
    echo npm版本:
    npm --version
)
echo.

echo [项目依赖检查]
echo --------------------------------
if exist "backend\venv" (
    echo Python虚拟环境: [已创建]
    echo 位置: backend\venv
) else (
    echo Python虚拟环境: [未创建]
)
echo.

if exist "frontend\node_modules" (
    echo 前端依赖: [已安装]
    dir /b "frontend\node_modules" | find /c /v "" > temp.txt
    set /p count=<temp.txt
    del temp.txt
    echo 包数量: !count!
) else (
    echo 前端依赖: [未安装]
)
echo.

echo [端口占用检查]
echo --------------------------------
echo 检查端口 8000 (后端):
netstat -an | findstr :8000 | findstr LISTENING >nul
if errorlevel 1 (
    echo 端口 8000: [可用]
) else (
    echo 端口 8000: [被占用]
)

echo 检查端口 3000 (前端):
netstat -an | findstr :3000 | findstr LISTENING >nul
if errorlevel 1 (
    echo 端口 3000: [可用]
) else (
    echo 端口 3000: [被占用]
)
echo.

echo [项目文件检查]
echo --------------------------------
set files=backend\requirements.txt frontend\package.json backend\app\main.py frontend\src\App.tsx
for %%f in (%files%) do (
    if exist "%%f" (
        echo [√] %%f
    ) else (
        echo [×] %%f [缺失]
    )
)
echo.

echo ======================================
echo 检查完成
echo ======================================
pause