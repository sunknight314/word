#!/bin/bash

echo "🚀 启动Word格式优化项目..."

# 检查依赖
echo "🔍 检查后端依赖..."
cd backend
if ! python3 -c "import fastapi, uvicorn, httpx, docx" 2>/dev/null; then
    echo "⚠️ 后端依赖缺失，正在安装..."
    pip3 install -r requirements.txt
fi

echo "🔍 检查前端依赖..."
cd ../frontend
if [ ! -d "node_modules" ]; then
    echo "⚠️ 前端依赖缺失，正在安装..."
    npm install
fi

# 启动后端
echo "🎯 启动后端服务..."
cd ../backend
python3 -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000 &
BACKEND_PID=$!

# 等待后端启动
echo "⏳ 等待后端启动..."
sleep 5

# 启动前端
echo "🎨 启动前端服务..."
cd ../frontend
npm start &
FRONTEND_PID=$!

echo ""
echo "✅ 服务已启动:"
echo "📡 后端API: http://localhost:8000"
echo "📱 前端界面: http://localhost:3000"
echo "📚 API文档: http://localhost:8000/docs"
echo ""
echo "💡 使用方法:"
echo "1. 打开浏览器访问 http://localhost:3000"
echo "2. 在源文件解析部分上传Word文档"
echo "3. 点击解析按钮查看AI分析结果"
echo ""
echo "🛑 按 Ctrl+C 停止所有服务"

# 清理函数
cleanup() {
    echo ""
    echo "🛑 正在停止服务..."
    kill $BACKEND_PID 2>/dev/null
    kill $FRONTEND_PID 2>/dev/null
    echo "👋 服务已停止"
    exit 0
}

# 捕获中断信号
trap cleanup INT

# 等待用户中断
wait $BACKEND_PID $FRONTEND_PID