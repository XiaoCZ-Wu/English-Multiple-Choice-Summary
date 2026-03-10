# GitHub Actions 快速开始

## 3 步构建 Android APK

### 第 1 步：创建 GitHub 仓库

访问 https://github.com/new 创建新仓库

### 第 2 步：推送代码

```bash
git init
git remote add origin https://github.com/你的用户名/仓库名.git
git add .
git commit -m "Initial commit"
git push -u origin main
```

### 第 3 步：下载 APK

1. 打开 GitHub 仓库页面
2. 点击 **Actions** 标签
3. 等待构建完成（约 15-30 分钟）
4. 点击最新的构建记录
5. 滚动到 **Artifacts** 部分
6. 点击 **android-apk** 下载

---

## 文件说明

已创建的文件：
- `.github/workflows/build-android.yml` - GitHub Actions 配置
- `GITHUB_ACTIONS_GUIDE.md` - 详细指南
- `GITHUB_QUICKSTART.md` - 本快速指南

---

## 触发构建的方式

1. **自动触发**：推送代码到 main/master 分支
2. **手动触发**：Actions 页面 → Run workflow

---

## 常见问题

**Q: 构建失败？**
A: 点击失败的构建查看日志，修复后重新推送

**Q: 构建时间多长？**
A: 首次 15-30 分钟，后续 5-15 分钟

**Q: 免费吗？**
A: 是的，GitHub Actions 对公开仓库完全免费

---

## 下一步

下载 APK 后：
1. 传输到 Android 手机
2. 安装 APK
3. 确保手机和电脑在同一 WiFi
4. 启动电脑端软件
5. 打开手机 APP，输入电脑 IP 和端口
6. 开始使用手机录入题目！
