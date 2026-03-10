#!/bin/bash

# Android APK 构建脚本
# 使用方法: ./build.sh

set -e  # 遇到错误立即退出

echo "=========================================="
echo "英语单选 - Android APK 构建脚本"
echo "=========================================="

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 检查是否在正确的目录
if [ ! -f "main.py" ]; then
    echo -e "${RED}错误: 请在 mobile_app 目录下运行此脚本${NC}"
    exit 1
fi

# 检查 buildozer
if ! command -v buildozer &> /dev/null; then
    echo -e "${YELLOW}未检测到 buildozer，尝试安装...${NC}"
    
    # 检查是否在虚拟环境中
    if [ -z "$VIRTUAL_ENV" ]; then
        echo -e "${YELLOW}创建虚拟环境...${NC}"
        python3 -m venv ../buildozer-venv
        source ../buildozer-venv/bin/activate
    fi
    
    # 安装依赖
    pip install --upgrade pip
    pip install buildozer cython
fi

# 检查 Java
if ! command -v java &> /dev/null; then
    echo -e "${RED}错误: 未安装 Java，请先安装 OpenJDK 17${NC}"
    echo "sudo apt install openjdk-17-jdk"
    exit 1
fi

# 显示环境信息
echo ""
echo "环境信息:"
echo "  Python: $(python3 --version)"
echo "  Java: $(java -version 2>&1 | head -n 1)"
echo "  Buildozer: $(buildozer --version)"
echo ""

# 清理旧构建
echo -e "${YELLOW}清理旧构建...${NC}"
buildozer android clean 2>/dev/null || true

# 开始构建
echo -e "${GREEN}开始构建 APK...${NC}"
echo "  这可能需要 30 分钟到 2 小时，取决于网络速度和电脑性能"
echo "  请耐心等待..."
echo ""

# 执行构建
if buildozer android debug; then
    echo ""
    echo -e "${GREEN}==========================================${NC}"
    echo -e "${GREEN}构建成功!${NC}"
    echo -e "${GREEN}==========================================${NC}"
    echo ""
    
    # 查找生成的 APK
    APK_PATH=$(find bin -name "*.apk" -type f | head -n 1)
    
    if [ -n "$APK_PATH" ]; then
        echo "APK 文件位置: $APK_PATH"
        echo "文件大小: $(du -h "$APK_PATH" | cut -f1)"
        echo ""
        echo "安装到手机:"
        echo "  1. 使用 ADB: adb install -r $APK_PATH"
        echo "  2. 或传输到手机后手动安装"
    fi
else
    echo ""
    echo -e "${RED}==========================================${NC}"
    echo -e "${RED}构建失败!${NC}"
    echo -e "${RED}==========================================${NC}"
    echo ""
    echo "常见问题:"
    echo "  1. 检查网络连接（需要下载大量依赖）"
    echo "  2. 检查磁盘空间（需要至少 20GB）"
    echo "  3. 查看错误日志: buildozer android debug 2>&1 | tee build.log"
    echo ""
    exit 1
fi
