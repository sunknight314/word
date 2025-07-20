@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ======================================
echo Word文档格式优化服务 - Windows启动脚本
echo ======================================
echo.

:: 检查Python
echo [1/6] 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.x
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [√] Python已安装

:: 检查pip
echo [2/6] 检查pip...
pip --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到pip，请确保Python安装时包含pip
    pause
    exit /b 1
)
echo [√] pip已安装

:: 检查Node.js
echo [3/6] 检查Node.js环境...
node --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Node.js，请先安装Node.js
    echo 下载地址: https://nodejs.org/
    pause
    exit /b 1
)
echo [√] Node.js已安装

:: 检查npm
echo [4/6] 检查npm...
npm --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到npm，请确保Node.js安装正确
    pause
    exit /b 1
)
echo [√] npm已安装

:: 安装后端依赖
echo.
echo [5/6] 检查并安装后端依赖...
cd backend
if not exist "venv" (
    echo 创建Python虚拟环境...
    python -m venv venv
)

:: 激活虚拟环境并安装依赖
echo 激活虚拟环境并安装依赖...
call venv\Scripts\activate.bat
pip install -r requirements.txt -q --upgrade
if errorlevel 1 (
    echo [错误] 后端依赖安装失败
    pause
    exit /b 1
)
echo [√] 后端依赖安装完成

:: 安装前端依赖
echo.
echo [6/6] 检查并安装前端依赖...
cd ..\frontend
if not exist "node_modules" (
    echo 安装前端依赖包...
    npm install
    if errorlevel 1 (
        echo [错误] 前端依赖安装失败
        pause
        exit /b 1
    )
) else (
    echo [√] 前端依赖已存在
)

:: 创建必要的目录
cd ..
if not exist "backend\uploads" mkdir backend\uploads
if not exist "backend\analysis_results" mkdir backend\analysis_results
if not exist "backend\analysis_results\paragraph_analysis" mkdir backend\analysis_results\paragraph_analysis
if not exist "backend\analysis_results\format_config" mkdir backend\analysis_results\format_config
if not exist "backend\ai_analysis_results" mkdir backend\ai_analysis_results

:: 启动服务
echo.
echo ======================================
echo 启动服务...
echo ======================================
echo.

:: 启动后端服务
echo 启动后端服务 (端口 8000)...
cd backend
start cmd /k "call venv\Scripts\activate.bat && python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000"

:: 等待后端启动
timeout /t 5 /nobreak >nul

:: 启动前端服务
echo 启动前端服务 (端口 3000)...
cd ..\frontend
start cmd /k "npm start"

:: 显示访问信息
echo.
echo ======================================
echo 服务启动成功！
echo ======================================
echo.
echo 前端访问地址: http://localhost:3000
echo 后端API地址: http://localhost:8000
echo API文档地址: http://localhost:8000/docs
echo.
echo 按任意键关闭此窗口（服务将继续运行）...
pause >nul