# 使用 Windows Server 2022 基础镜像
FROM mcr.microsoft.com/windows/server:ltsc2022

# 设置工作目录
WORKDIR C:/wps-test

# 创建测试目录
RUN mkdir test-input && mkdir test-output

# 复制最小验证脚本
COPY test_wps.py .
COPY test_doc.docx .

# 安装 Python 和必要依赖
RUN powershell -Command \
    Write-Host "安装 Python..."; \
    Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe" -OutFile "python-installer.exe"; \
    Start-Process -FilePath "python-installer.exe" -ArgumentList "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_test=0" -Wait -NoNewWindow; \
    Write-Host "Python 安装完成"; \
    & "C:\Program Files\Python39\python.exe" -m pip install pypiwin32

# 下载并安装 WPS Office
RUN powershell -Command \
    Write-Host "正在下载 WPS Office..."; \
    # 从 WPS 官网下载免费版（请替换为实际可用的下载链接）
    $wpsUrl = "https://github.com/kchzhang/wps-wine/releases/download/0.1/wps-office2023.exe"; \
    # 或者使用其他镜像源（注意：需要确认链接有效性）
    # $wpsUrl = "https://github.com/kingsoft-wps/wps-office-for-linux/releases/download/v11.1.0/wps-office_11.1.0.11691_amd64.exe"; \
    # 如果上述链接不可用，可以使用备用方法下载 Windows 版本 \
    try { \
        Invoke-WebRequest -Uri $wpsUrl -OutFile "wps-office2023.exe" -TimeoutSec 300; \
        Write-Host "WPS Office 下载完成"; \
    } catch { \
        Write-Host "WPS 下载失败，尝试备用方案..."; \
        # 备用方案：下载在线安装器 \
        $wpsOnlineUrl = "https://github.com/kchzhang/wps-wine/releases/download/0.1/wps-office2023.exe"; \
        Invoke-WebRequest -Uri $wpsOnlineUrl -OutFile "wps-office2023.exe" -TimeoutSec 300; \
        Write-Host "WPS 在线安装器下载完成"; \
    }; \
    Write-Host "正在安装 WPS Office..."; \
    # 尝试静默安装 \
    $installAttempts = @("/S", "/quiet", "/silent", "/verysilent"); \
    $installed = $false; \
    foreach ($arg in $installAttempts) { \
        try { \
            Write-Host "尝试安装参数: $arg"; \
            Start-Process -FilePath "C:\wps-test\wps-office2023.exe" -ArgumentList $arg -Wait -NoNewWindow; \
            Start-Sleep -Seconds 10; \
            # 检查是否安装成功 \
            $wpsPaths = @("C:\Program Files\Kingsoft\WPS Office", "C:\Program Files (x86)\Kingsoft\WPS Office"); \
            foreach ($path in $wpsPaths) { \
                if (Test-Path $path) { \
                    Write-Host "✓ WPS Office 安装成功"; \
                    $installed = $true; \
                    break; \
                } \
            } \
            if ($installed) { break; } \
        } catch { \
            Write-Host "安装参数 $arg 失败: $($_.Exception.Message)"; \
        } \
    }; \
    if (-not $installed) { \
        Write-Host "⚠ 静默安装失败，尝试正常安装..."; \
        # 最后尝试正常安装（需要用户交互，在容器中可能不工作） \
        try { \
            Start-Process -FilePath "C:\wps-test\wps-office2023.exe" -Wait -NoNewWindow; \
            Write-Host "正常安装完成"; \
        } catch { \
            Write-Host "安装失败: $($_.Exception.Message)"; \
            exit 1; \
        } \
    }

# 设置启动命令 - 运行验证脚本
CMD ["powershell", "-Command", "C:\\Program Files\\Python39\\python.exe C:\\wps-test\\test_wps.py"]