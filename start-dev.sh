#!/bin/bash

echo "🚀 启动清关优化系统..."
echo ""

# 检查是否在正确的目录
if [ ! -f "backend/app/main.py" ]; then
    echo "❌ 错误：请在项目根目录运行此脚本"
    exit 1
fi

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误：未找到 Python3"
    exit 1
fi

# 检查后端依赖
echo "📦 检查后端依赖..."
cd backend
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
fi

source venv/bin/activate
pip3 install -r requirements.txt --quiet

# 检查 .env 文件
if [ ! -f ".env" ]; then
    echo "❌ 错误：未找到 .env 文件"
    echo "请复制 .env.example 为 .env 并填写配置"
    exit 1
fi

# 创建必要的目录
mkdir -p uploads outputs

# 启动后端
echo ""
echo "🔧 启动后端服务..."
python3 -m uvicorn app.main:app --reload --port 8000 &
BACKEND_PID=$!

# 等待后端启动
sleep 3

# 检查后端是否启动成功
if curl -s http://localhost:8000/docs > /dev/null; then
    echo "✅ 后端启动成功: http://localhost:8000"
else
    echo "❌ 后端启动失败"
    kill $BACKEND_PID
    exit 1
fi

# 启动前端
cd ../frontend
echo ""
echo "🎨 启动前端服务..."

# 检查 node_modules
if [ ! -d "node_modules" ]; then
    echo "安装前端依赖..."
    npm install
fi

npm run dev &
FRONTEND_PID=$!

echo ""
echo "✅ 系统启动完成！"
echo ""
echo "📍 访问地址："
echo "   前端: http://localhost:3000"
echo "   后端: http://localhost:8000"
echo "   API文档: http://localhost:8000/docs"
echo ""
echo "按 Ctrl+C 停止服务"
echo ""

# 等待用户中断
trap "echo ''; echo '🛑 正在停止服务...'; kill $BACKEND_PID $FRONTEND_PID; exit" INT
wait
