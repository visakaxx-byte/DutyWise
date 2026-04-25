#!/bin/bash

# 启动脚本

echo "🚀 启动清关文件处理系统后端服务..."

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "📦 创建虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "📥 安装依赖..."
pip install -r requirements.txt

# 创建必要的目录
mkdir -p uploads outputs

# 检查环境变量文件
if [ ! -f ".env" ]; then
    echo "⚠️  未找到 .env 文件，复制 .env.example..."
    cp .env.example .env
    echo "请编辑 .env 文件配置必要的环境变量"
fi

# 启动服务
echo "✅ 启动服务..."
python -m app.main
