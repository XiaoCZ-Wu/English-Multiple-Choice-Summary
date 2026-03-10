# Android APK 快速构建指南

## 推荐方案：使用 Docker（最简单）

### 前提条件
- 安装了 Docker 的 Linux 系统（Ubuntu 推荐）
- 至少 20GB 磁盘空间
- 稳定的网络连接

### 步骤

1. **将代码复制到 Linux 系统**
   ```bash
   # 使用 scp 从 Windows 复制
   scp -r /mnt/d/Data/Python/英语单选/mobile_app user@linux-ip:~/
   
   # 或者使用 Git
   git clone https://github.com/yourusername/yourrepo.git
   ```

2. **进入目录并运行构建脚本**
   ```bash
   cd mobile_app
   chmod +x docker-build.sh
   ./docker-build.sh
   ```

3. **等待构建完成**
   - 首次构建：1-2 小时
   - 后续构建：10-30 分钟

4. **获取 APK 文件**
   - 构建完成后，APK 位于 `mobile_app/bin/` 目录
   - 文件名类似：`englishquiz-0.1-arm64-v8a_armeabi-v7a-debug.apk`

5. **安装到手机**
   ```bash
   # 使用 ADB
   adb install -r bin/englishquiz-0.1-*.apk
   
   # 或者传输到手机后手动安装
   ```

---

## 备选方案：本地安装 Buildozer

如果你不想使用 Docker，可以在 Ubuntu 上本地安装：

### 1. 一键安装脚本

创建并运行以下脚本：

```bash
cat > install-buildozer.sh << 'EOF'
#!/bin/bash
set -e

echo "安装 Buildozer 依赖..."

sudo apt update
sudo apt install -y \
    python3-pip python3-venv git zip unzip \
    openjdk-17-jdk autoconf libtool pkg-config \
    zlib1g-dev libncurses5-dev libncursesw5-dev \
    libtinfo5 cmake libffi-dev libssl-dev automake

echo "创建虚拟环境..."
python3 -m venv ~/buildozer-venv
source ~/buildozer-venv/bin/activate

echo "安装 Buildozer..."
pip install --upgrade pip
pip install buildozer cython

echo "安装完成！"
echo "使用方法:"
echo "  source ~/buildozer-venv/bin/activate"
echo "  cd mobile_app"
echo "  buildozer android debug"
EOF

chmod +x install-buildozer.sh
./install-buildozer.sh
```

### 2. 构建 APK

```bash
# 激活虚拟环境
source ~/buildozer-venv/bin/activate

# 进入项目目录
cd mobile_app

# 开始构建
buildozer android debug
```

---

## Windows 用户特别提示

### 方案 1：使用 WSL2（Windows Subsystem for Linux）

1. **安装 WSL2**
   ```powershell
   # 以管理员身份运行 PowerShell
   wsl --install -d Ubuntu-22.04
   ```

2. **重启电脑**，然后打开 Ubuntu 终端

3. **在 WSL2 中安装 Docker**
   ```bash
   sudo apt update
   sudo apt install -y docker.io
   sudo usermod -aG docker $USER
   ```

4. **访问 Windows 文件**
   ```bash
   # Windows 磁盘挂载在 /mnt/
   cd /mnt/d/Data/Python/英语单选/mobile_app
   
   # 运行构建
   ./docker-build.sh
   ```

### 方案 2：使用虚拟机

1. 安装 VirtualBox 或 VMware
2. 安装 Ubuntu 22.04
3. 设置共享文件夹
4. 在虚拟机中按照 Docker 方案构建

### 方案 3：使用 GitHub Actions（无需 Linux）

1. 将代码推送到 GitHub
2. 我已经创建了 `.github/workflows/build-android.yml`
3. 每次推送代码，GitHub 会自动构建 APK
4. 在 Actions 页面下载构建好的 APK

---

## 常见问题速查

### Q: 构建失败，提示缺少依赖
**A:** 运行以下命令安装所有依赖
```bash
sudo apt install -y python3-pip python3-venv git zip unzip \
    openjdk-17-jdk autoconf libtool pkg-config zlib1g-dev \
    libncurses5-dev libncursesw5-dev libtinfo5 cmake \
    libffi-dev libssl-dev automake
```

### Q: 构建过程中断，如何继续
**A:** 直接重新运行构建命令，会自动继续
```bash
./docker-build.sh
# 或
buildozer android debug
```

### Q: 如何清理缓存重新构建
**A:**
```bash
buildozer android clean
rm -rf ~/.buildozer/android/platform/build-*
buildozer android debug
```

### Q: 构建成功但安装失败
**A:** 确保手机开启了"允许安装未知来源应用"
- Android 8.0+: 设置 → 应用 → 特殊访问权限 → 安装未知应用
- Android 7.0 及以下: 设置 → 安全 → 未知来源

### Q: APK 文件在哪里
**A:** 构建完成后在 `mobile_app/bin/` 目录
```bash
ls -la bin/*.apk
```

---

## 文件说明

- `main.py` - APP 主程序
- `buildozer.spec` - 构建配置文件
- `requirements.txt` - Python 依赖
- `build.sh` - 本地构建脚本
- `docker-build.sh` - Docker 构建脚本
- `BUILD_GUIDE.md` - 详细构建指南
- `QUICK_START.md` - 本快速指南

---

## 获取帮助

如果构建遇到问题：

1. 查看详细日志：`buildozer android debug 2>&1 | tee build.log`
2. 查阅 [BUILD_GUIDE.md](BUILD_GUIDE.md) 的常见问题部分
3. 访问 Buildozer 官方文档：https://buildozer.readthedocs.io/
4. 访问 Kivy 官方文档：https://kivy.org/doc/stable/

---

## 下一步

构建完成后：

1. 将 APK 安装到 Android 手机
2. 确保手机和电脑在同一 WiFi 下
3. 启动电脑端的英语单选软件
4. 打开手机 APP，输入电脑的 IP 地址和端口
5. 开始使用手机录入题目！
