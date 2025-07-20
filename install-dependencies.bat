@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ======================================
echo 依赖安装脚本 - Windows
echo ======================================
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] 建议以管理员身份运行此脚本
    echo.
)

:: Python依赖检查和安装
echo [步骤1] Python环境检查和设置
echo --------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python未安装
    echo.
    echo 请从以下地址下载并安装Python 3.8或更高版本:
    echo https://www.python.org/downloads/
    echo.
    echo 安装时请勾选:
    echo - Add Python to PATH
    echo - Install pip
    echo.
    pause
    exit /b 1
) else (
    echo [√] Python已安装
    python --version
)

:: Node.js依赖检查和安装
echo.
echo [步骤2] Node.js环境检查和设置
echo --------------------------------
node --version >nul 2>&1
if errorlevel 1 (
    echo [!] Node.js未安装
    echo.
    echo 请从以下地址下载并安装Node.js LTS版本:
    echo https://nodejs.org/
    echo.
    echo 安装时会自动包含npm包管理器
    echo.
    pause
    exit /b 1
) else (
    echo [√] Node.js已安装
    node --version
    npm --version
)

:: 创建Python虚拟环境
echo.
echo [步骤3] 设置Python虚拟环境
echo --------------------------------
cd backend
if exist "venv" (
    echo [!] 虚拟环境已存在，是否重新创建？(Y/N)
    set /p recreate=
    if /i "!recreate!"=="Y" (
        echo 删除旧虚拟环境...
        rmdir /s /q venv
        echo 创建新虚拟环境...
        python -m venv venv
    )
) else (
    echo 创建虚拟环境...
    python -m venv venv
)

:: 升级pip
echo.
echo 升级pip到最新版本...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip

:: 安装Python依赖
echo.
echo [步骤4] 安装Python依赖包
echo --------------------------------
pip install -r requirements.txt
if errorlevel 1 (
    echo [!] 部分依赖安装失败，尝试逐个安装...
    for /f "tokens=*" %%i in (requirements.txt) do (
        echo 安装: %%i
        pip install %%i
    )
)

:: 安装前端依赖
echo.
echo [步骤5] 安装前端依赖包
echo --------------------------------
cd ..\frontend

:: 清理npm缓存（可选）
echo 清理npm缓存...
npm cache clean --force

:: 安装依赖
echo 安装依赖包（这可能需要几分钟）...
npm install
if errorlevel 1 (
    echo [!] npm install失败，尝试使用淘宝镜像...
    npm config set registry https://registry.npmmirror.com
    npm install
    :: 恢复默认镜像
    npm config set registry https://registry.npmjs.org
)

:: 创建必要目录
echo.
echo [步骤6] 创建必要的目录结构
echo --------------------------------
cd ..
set dirs=backend\uploads backend\analysis_results backend\analysis_results\paragraph_analysis backend\analysis_results\format_config backend\ai_analysis_results
for %%d in (%dirs%) do (
    if not exist "%%d" (
        mkdir "%%d"
        echo [√] 创建目录: %%d
    )
)

:: 完成
echo.
echo ======================================
echo 依赖安装完成！
echo ======================================
echo.
echo 现在可以运行 start.bat 启动服务
echo.
pause