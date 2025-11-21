# 使用更轻量的基础镜像
FROM ubuntu:22.04

# 设置环境变量
ENV DEBIAN_FRONTEND=noninteractive
ENV TZ=Asia/Shanghai
ENV WINEPREFIX=/root/.wine
ENV WINEARCH=win64
ENV DISPLAY=:99

# 安装基础依赖和Wine
RUN apt-get update && apt-get install -y \
    wget \
    cabextract \
    winetricks \
    wine \
    wine64 \
    xvfb \
    python3 \
    python3-pip \
    && rm -rf /var/lib/apt/lists/*

# 安装Python COM支持库
RUN pip3 install pywin32

# 配置Wine和安装必要的运行库
RUN winecfg && \
    winetricks -q corefonts && \
    winetricks -q vb6run && \
    winetricks -q vcrun2013 && \
    winetricks -q vcrun2015

# 下载并安装WPS Office
RUN wget -O /tmp/wps-office2023.exe "https://github.com/kchzhang/wps-wine/releases/download/0.1/wps-office2023.exe" && \
    xvfb-run -a wine /tmp/wps-office2023.exe /S && \
    rm /tmp/wps-office2023.exe


# 创建工作目录
WORKDIR /app

# 复制项目文件
COPY requirements.txt .
COPY app/ ./app/
COPY entrypoint.sh .
COPY startup.sh .

# 安装Python依赖
RUN pip3 install -r requirements.txt

RUN chmod +x /app/entrypoint.sh /app/startup.sh

# 暴露服务端口
EXPOSE 8000

# # 设置数据卷
# VOLUME ["/tmp/uploads", "/tmp/outputs"]

ENTRYPOINT ["/app/entrypoint.sh"]