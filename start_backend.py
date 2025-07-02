#!/usr/bin/env python3
"""
启动后端服务
"""

import subprocess
import sys
import os

def check_dependencies():
    """检查依赖是否安装"""
    print("🔍 检查依赖...")
    
    try:
        import fastapi
        import uvicorn
        import httpx
        import docx
        print("✅ 所有依赖已安装")
        return True
    except ImportError as e:
        print(f"❌ 缺少依赖: {e}")
        print("请运行: pip3 install -r requirements.txt")
        return False

def start_server():
    """启动FastAPI服务器"""
    backend_dir = "/opt/word/backend"
    
    if not os.path.exists(backend_dir):
        print(f"❌ 后端目录不存在: {backend_dir}")
        return
    
    print("🚀 启动FastAPI服务器...")
    print(f"工作目录: {backend_dir}")
    print("API地址: http://localhost:8000")
    print("API文档: http://localhost:8000/docs")
    print()
    print("按 Ctrl+C 停止服务")
    print("-" * 50)
    
    try:
        # 切换到后端目录并启动服务
        os.chdir(backend_dir)
        subprocess.run([
            sys.executable, "-m", "uvicorn", 
            "app.main:app", 
            "--reload", 
            "--host", "0.0.0.0", 
            "--port", "8000"
        ])
    except KeyboardInterrupt:
        print("\n👋 服务已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")

if __name__ == "__main__":
    print("🔧 Word文档格式优化项目 - 后端服务")
    print("=" * 50)
    
    if check_dependencies():
        start_server()
    else:
        print("\n请先安装依赖后再启动服务")