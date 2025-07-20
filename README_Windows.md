# Windows环境运行指南

## 系统要求

- Windows 10/11
- Python 3.8+ (推荐3.9或3.10)
- Node.js 14+ (推荐LTS版本)

## 快速开始

### 1. 首次运行 - 安装依赖

双击运行 `install-dependencies.bat`，该脚本会：
- 检查Python和Node.js是否已安装
- 创建Python虚拟环境
- 安装所有Python依赖包
- 安装所有前端依赖包
- 创建必要的目录结构

### 2. 启动服务

双击运行 `start.bat`，该脚本会：
- 自动检查环境
- 启动后端服务（端口8000）
- 启动前端服务（端口3000）
- 在新窗口中运行服务

### 3. 访问服务

- 前端界面: http://localhost:3000
- 后端API: http://localhost:8000
- API文档: http://localhost:8000/docs

### 4. 停止服务

双击运行 `stop.bat` 停止所有服务。

## 脚本说明

| 脚本名称 | 功能说明 |
|---------|---------|
| `install-dependencies.bat` | 安装所有依赖（首次运行） |
| `start.bat` | 启动前后端服务 |
| `stop.bat` | 停止所有服务 |
| `check-environment.bat` | 检查环境配置状态 |

## 常见问题

### 1. Python未安装
- 访问 https://www.python.org/downloads/
- 下载并安装Python 3.8或更高版本
- 安装时勾选"Add Python to PATH"

### 2. Node.js未安装
- 访问 https://nodejs.org/
- 下载并安装Node.js LTS版本
- npm会随Node.js一起安装

### 3. 端口被占用
- 运行 `check-environment.bat` 查看端口状态
- 使用 `stop.bat` 停止占用端口的服务
- 或修改配置文件中的端口号

### 4. 依赖安装失败
- 确保网络连接正常
- 尝试使用管理员权限运行
- 可以尝试使用国内镜像源

### 5. 防火墙提示
- 首次运行时Windows防火墙可能会提示
- 请允许Python和Node.js通过防火墙

## 开发模式

服务默认以开发模式运行，支持热重载：
- 后端代码修改后会自动重启
- 前端代码修改后会自动刷新

## 生产部署

如需生产部署，请：
1. 构建前端：`cd frontend && npm run build`
2. 使用生产服务器运行后端
3. 配置反向代理（如nginx）

## 技术支持

如遇到问题，请检查：
1. 运行 `check-environment.bat` 查看环境状态
2. 查看控制台错误信息
3. 检查 `backend/logs` 目录下的日志文件