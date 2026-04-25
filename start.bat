@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo 🚀 清关优化系统启动中...
echo.

REM 检查 Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ 错误: 未安装 Node.js
    echo 请访问 https://nodejs.org/ 下载安装
    pause
    exit /b 1
)

REM 检查 Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ 错误: 未安装 Python
    echo 请访问 https://www.python.org/ 下载安装
    pause
    exit /b 1
)

REM 检查依赖是否已安装
if not exist "node_modules" (
    echo 📦 安装根目录依赖...
    call npm install
)

if not exist "frontend\node_modules" (
    echo 📦 安装前端依赖...
    cd frontend
    call npm install
    cd ..
)

REM 启动应用
echo.
echo ✅ 依赖检查完成
echo 🎯 启动应用...
echo.
echo 前端: http://localhost:5173
echo 后端: http://localhost:8000
echo API 文档: http://localhost:8000/docs
echo.

call npm run dev
