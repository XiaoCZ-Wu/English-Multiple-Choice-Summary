#!/bin/bash

# 使用 Docker 构建 Android APK
# 使用方法: ./docker-build.sh

set -e

echo "=========================================="
echo "使用 Docker 构建 Android APK"
echo "=========================================="

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

# 检查 Docker
if ! command -v docker &> /dev/null; then
    echo -e "${RED}错误: 未安装 Docker${NC}"
    echo "请先安装 Docker:"
    echo "  Ubuntu: sudo apt install docker.io"
    echo "  其他系统请参考 Docker 官方文档"
    exit 1
fi

# 检查 Docker 服务
if ! docker info &> /dev/null; then
    echo -e "${RED}错误: Docker 服务未运行${NC}"
    echo "请启动 Docker 服务:"
    echo "  sudo systemctl start docker"
    exit 1
fi

# 获取当前目录的绝对路径
PROJECT_DIR="$(pwd)"

echo ""
echo "项目目录: $PROJECT_DIR"
echo ""

# 检查必要的文件
if [ ! -f "main.py" ]; then
    echo -e "${RED}错误: 未找到 main.py，请在 mobile_app 目录下运行${NC}"
    exit 1
fi

if [ ! -f "buildozer.spec" ]; then
    echo -e "${RED}错误: 未找到 buildozer.spec${NC}"
    exit 1
fi

# 创建构建目录
mkdir -p "$PROJECT_DIR/bin"
mkdir -p "$HOME/.buildozer"

echo -e "${YELLOW}拉取 Buildozer 镜像...${NC}"
docker pull kivy/buildozer:latest

echo ""
echo -e "${GREEN}开始构建...${NC}"
echo "  首次构建需要下载大量依赖，可能需要 1-2 小时"
echo "  请耐心等待..."
echo ""

# 运行构建
docker run -it --rm \
    --name buildozer-build \
    -v "$PROJECT_DIR:/home/user/app" \
    -v "$HOME/.buildozer:/home/user/.buildozer" \
    -w /home/user/app \
    kivy/buildozer:latest \
    buildozer android debug

# 检查构建结果
if [ -f "$PROJECT_DIR/bin"/*.apk ]; then
    echo ""
    echo -e "${GREEN}==========================================${NC}"
    echo -e "${GREEN}构建成功!${NC}"
    echo -e "${GREEN}==========================================${NC}"
    echo ""
    
    APK_FILE=$(ls -t "$PROJECT_DIR/bin"/*.apk | head -n 1)
    echo "APK 文件: $APK_FILE"
    echo "文件大小: $(du -h "$APK_FILE" | cut -f1)"
    echo ""
    echo "下一步:"
    echo "  1. 将 APK 传输到手机"
    echo "  2. 在手机上安装并运行"
    echo ""
else
    echo -e "${RED}构建失败，未找到 APK 文件${NC}"
    exit 1
fi
