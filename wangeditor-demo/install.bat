@echo off
echo ===================================
echo WangEditor 表格示例项目安装脚本
echo ===================================
echo.
echo 正在安装依赖包(使用淘宝源)...
call npm install --registry=https://registry.npmmirror.com
echo.
echo 依赖包安装完成!
echo.
echo 正在启动开发服务器...
call npm run serve