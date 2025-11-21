#!/bin/bash

# 启动虚拟显示服务器
Xvfb :99 -screen 0 1024x768x16 &
sleep 2

# 设置Wine环境
export WINEPREFIX=/root/.wine
export WINEARCH=win64

# 执行启动脚本
exec /app/startup.sh