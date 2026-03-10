# Android APK 打包指南

## 环境要求

- Ubuntu 20.04/22.04 LTS (推荐)
- 至少 20GB 磁盘空间
- 稳定的网络连接

## 方法一：使用 Docker（推荐）

### 1. 安装 Docker

```bash
# 更新系统
sudo apt update
sudo apt upgrade -y

# 安装 Docker
sudo apt install -y docker.io

# 启动 Docker 服务
sudo systemctl start docker
sudo systemctl enable docker

# 将当前用户添加到 docker 组
sudo usermod -aG docker $USER

# 重新登录或执行
newgrp docker
```

### 2. 使用 Buildozer Docker 镜像

```bash
# 创建工作目录
mkdir -p ~/buildozer-build
cd ~/buildozer-build

# 复制手机APP代码
# 将 mobile_app 目录复制到此目录

# 运行 Buildozer 容器
docker run -it --rm \
    -v $(pwd)/mobile_app:/home/user/app \
    -v $(pwd)/.buildozer:/home/user/.buildozer \
    kivy/buildozer

# 容器内执行
cd app
buildozer android debug
```

## 方法二：本地安装 Buildozer

### 1. 安装依赖

```bash
# 更新系统
sudo apt update
sudo apt upgrade -y

# 安装基本依赖
sudo apt install -y \
    python3-pip \
    python3-venv \
    git \
    zip \
    unzip \
    openjdk-17-jdk \
    autoconf \
    libtool \
    pkg-config \
    zlib1g-dev \
    libncurses5-dev \
    libncursesw5-dev \
    libtinfo5 \
    cmake \
    libffi-dev \
    libssl-dev \
    automake
```

### 2. 安装 Android SDK/NDK

```bash
# 创建目录
mkdir -p ~/android-sdk
cd ~/android-sdk

# 下载命令行工具
wget https://dl.google.com/android/repository/commandlinetools-linux-9477386_latest.zip
unzip commandlinetools-linux-9477386_latest.zip
mkdir -p cmdline-tools/latest
mv cmdline-tools/* cmdline-tools/latest/ 2>/dev/null || true

# 设置环境变量
echo 'export ANDROID_SDK_ROOT=$HOME/android-sdk' >> ~/.bashrc
echo 'export PATH=$PATH:$ANDROID_SDK_ROOT/cmdline-tools/latest/bin' >> ~/.bashrc
source ~/.bashrc

# 安装必要组件
yes | sdkmanager --licenses
sdkmanager "platform-tools" "platforms;android-33" "build-tools;33.0.0"
```

### 3. 安装 Buildozer

```bash
# 创建虚拟环境
python3 -m venv ~/buildozer-venv
source ~/buildozer-venv/bin/activate

# 安装 buildozer
pip install --upgrade pip
pip install buildozer

# 安装 Cython
pip install cython
```

### 4. 配置 Buildozer

```bash
# 进入项目目录
cd ~/buildozer-build/mobile_app

# 初始化 buildozer（如果还没有 buildozer.spec）
# buildozer init

# 编辑 buildozer.spec 文件
# 确保以下配置正确：
# - title: 英语单选
# - package.name: englishquiz
# - requirements: python3,kivy,requests,urllib3,charset-normalizer,idna,certifi
# - android.permissions: INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,CAMERA
```

### 5. 构建 APK

```bash
# 确保在虚拟环境中
source ~/buildozer-venv/bin/activate

# 进入项目目录
cd ~/buildozer-build/mobile_app

# 清理之前的构建
buildozer android clean

# 开始构建（首次构建需要很长时间，可能需要1-2小时）
buildozer android debug

# 构建完成后，APK 文件位于:
# ./bin/englishquiz-0.1-arm64-v8a_armeabi-v7a-debug.apk
```

## 方法三：使用 GitHub Actions（自动化）

### 1. 创建 GitHub 仓库

将代码推送到 GitHub 仓库。

### 2. 创建工作流文件

```bash
mkdir -p .github/workflows
cat > .github/workflows/build-android.yml << 'EOF'
name: Build Android APK

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  build:
    runs-on: ubuntu-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
    
    - name: Install dependencies
      run: |
        sudo apt update
        sudo apt install -y \
          python3-pip \
          python3-venv \
          git \
          zip \
          unzip \
          openjdk-17-jdk \
          autoconf \
          libtool \
          pkg-config \
          zlib1g-dev \
          libncurses5-dev \
          libncursesw5-dev \
          libtinfo5 \
          cmake \
          libffi-dev \
          libssl-dev \
          automake
    
    - name: Install Buildozer
      run: |
        python3 -m venv ~/buildozer-venv
        source ~/buildozer-venv/bin/activate
        pip install --upgrade pip
        pip install buildozer cython
    
    - name: Build APK
      run: |
        source ~/buildozer-venv/bin/activate
        cd mobile_app
        buildozer android debug
    
    - name: Upload APK
      uses: actions/upload-artifact@v3
      with:
        name: android-apk
        path: mobile_app/bin/*.apk
EOF
```

### 3. 推送到 GitHub

```bash
git add .github/workflows/build-android.yml
git commit -m "Add GitHub Actions workflow for Android build"
git push
```

GitHub Actions 会自动构建 APK，你可以在 Actions 页面下载构建好的 APK。

## 常见问题

### 1. 构建失败：缺少依赖

```bash
# 安装所有可能的依赖
sudo apt install -y \
    build-essential \
    libsqlite3-dev \
    libreadline-dev \
    libbz2-dev \
    liblzma-dev \
    tk-dev \
    libgdbm-dev \
    libc6-dev
```

### 2. 内存不足

```bash
# 增加交换空间
sudo fallocate -l 8G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile

# 查看交换空间
free -h
```

### 3. 网络问题（下载慢）

```bash
# 配置镜像源
mkdir -p ~/.buildozer/android/platform

# 手动下载 android-ndk
wget https://dl.google.com/android/repository/android-ndk-r23b-linux.zip
unzip android-ndk-r23b-linux.zip -d ~/.buildozer/android/platform/

# 手动下载 android-sdk
wget https://dl.google.com/android/repository/commandlinetools-linux-9477386_latest.zip
mkdir -p ~/.buildozer/android/platform/android-sdk/cmdline-tools
unzip commandlinetools-linux-9477386_latest.zip -d ~/.buildozer/android/platform/android-sdk/cmdline-tools/
mv ~/.buildozer/android/platform/android-sdk/cmdline-tools/cmdline-tools \
   ~/.buildozer/android/platform/android-sdk/cmdline-tools/latest
```

### 4. Python 版本问题

```bash
# 确保使用 Python 3.10
python3 --version

# 如果不是 3.10，安装并切换
sudo apt install -y python3.10 python3.10-venv python3.10-dev
```

### 5. 清理构建缓存

```bash
# 如果构建失败，尝试清理后重新构建
buildozer android clean
rm -rf ~/.buildozer/android/platform/build-*
buildozer android debug
```

## 安装 APK 到手机

### 方法 1：使用 ADB

```bash
# 安装 ADB
sudo apt install -y adb

# 连接手机（需要开启USB调试）
adb devices

# 安装 APK
adb install -r bin/englishquiz-0.1-arm64-v8a_armeabi-v7a-debug.apk
```

### 方法 2：直接传输

1. 将 APK 文件传输到手机
2. 在手机上点击安装
3. 可能需要允许"安装未知来源应用"

## 调试

### 查看日志

```bash
# 连接手机后
adb logcat -s python:D
```

### 使用 Buildozer 部署并运行

```bash
# 部署到连接的设备并运行
buildozer android debug deploy run

# 查看日志
buildozer android logcat
```

## 发布版本构建

```bash
# 构建发布版本（需要签名）
buildozer android release

# 签名 APK（需要创建密钥库）
keytool -genkey -v -keystore my-release-key.keystore -alias alias_name -keyalg RSA -keysize 2048 -validity 10000

# 使用 jarsigner 签名
jarsigner -verbose -sigalg SHA1withRSA -digestalg SHA1 -keystore my-release-key.keystore bin/englishquiz-0.1-release-unsigned.apk alias_name

# 优化 APK
zipalign -v 4 bin/englishquiz-0.1-release-unsigned.apk bin/englishquiz-0.1-release.apk
```

## 文件传输到 Linux

如果你是在 Windows 上开发，需要将文件传输到 Linux：

### 方法 1：使用 SCP

```bash
# 在 Windows 上使用 PowerShell 或 Git Bash
scp -r D:/Data/Python/英语单选/mobile_app user@linux-ip:~/buildozer-build/
```

### 方法 2：使用共享文件夹

如果使用虚拟机，设置共享文件夹后直接复制。

### 方法 3：使用 Git

```bash
# 提交到 GitHub
git add mobile_app/
git commit -m "Add mobile app"
git push

# 在 Linux 上克隆
git clone https://github.com/yourusername/yourrepo.git
cd yourrepo/mobile_app
```

## 总结

最简单的方法是使用 **Docker** 或 **GitHub Actions**，不需要在本地配置复杂的构建环境。

如果是本地构建，推荐使用 **Ubuntu 22.04 LTS**，按照方法二逐步安装。

首次构建时间较长（1-2小时），请耐心等待。后续构建会快很多。
