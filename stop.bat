@echo off
chcp 65001 >nul

echo ======================================
echo 停止服务脚本
echo ======================================
echo.

echo 正在查找并停止相关进程...
echo.

:: 停止Python进程（uvicorn）
echo [1/2] 停止后端服务...
for /f "tokens=2" %%i in ('netstat -ano ^| findstr :8000 ^| findstr LISTENING') do (
    echo 找到进程 PID: %%i
    taskkill /PID %%i /F >nul 2>&1
)

:: 停止Node.js进程（React开发服务器）
echo [2/2] 停止前端服务...
for /f "tokens=2" %%i in ('netstat -ano ^| findstr :3000 ^| findstr LISTENING') do (
    echo 找到进程 PID: %%i
    taskkill /PID %%i /F >nul 2>&1
)

:: 额外确保停止所有相关进程
taskkill /IM "node.exe" /F >nul 2>&1
taskkill /IM "python.exe" /F >nul 2>&1

echo.
echo ======================================
echo 服务已停止
echo ======================================
echo.
pause