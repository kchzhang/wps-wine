#!/bin/bash

# 启动FastAPI服务
cd /app
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 1