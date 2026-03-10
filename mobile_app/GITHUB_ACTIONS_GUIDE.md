# GitHub Actions 自动构建 Android APK 指南

## 简介

使用 GitHub Actions 可以在云端自动构建 Android APK，无需本地 Linux 环境，也无需担心磁盘空间不足的问题。

## 前提条件

1. GitHub 账号（免费）
2. 将代码推送到 GitHub 仓库

## 快速开始

### 步骤 1：创建 GitHub 仓库

1. 访问 https://github.com/new
2. 输入仓库名称，例如 `english-quiz-mobile`
3. 选择"Public"（公开）或"Private"（私有）
4. 点击 "Create repository"

### 步骤 2：推送代码到 GitHub

在本地项目目录执行：

```bash
# 初始化 git（如果还没有）
git init

# 添加远程仓库（替换为你的仓库地址）
git remote add origin https://github.com/你的用户名/english-quiz-mobile.git

# 添加所有文件
git add .

# 提交
git commit -m "Initial commit"

# 推送
git push -u origin main
```

### 步骤 3：触发自动构建

推送代码后，GitHub Actions 会自动开始构建：

1. 打开 GitHub 仓库页面
2. 点击顶部的 "Actions" 标签
3. 查看构建进度

### 步骤 4：下载 APK

构建完成后：

1. 在 Actions 页面点击最新的工作流运行
2. 滚动到底部的 "Artifacts" 部分
3. 点击 "android-apk" 下载 APK 文件

## 工作流程说明

### 自动触发条件

- 推送到 `main` 或 `master` 分支
- 修改了 `mobile_app/**` 目录下的文件
- 手动触发（在 Actions 页面点击 "Run workflow"）

### 构建过程

1. **检出代码** - 从 GitHub 仓库获取代码
2. **设置 Python** - 安装 Python 3.10
3. **安装依赖** - 安装 Java、Android SDK 等依赖
4. **安装 Buildozer** - 安装构建工具
5. **缓存依赖** - 缓存构建依赖，加速后续构建
6. **构建 APK** - 使用 Buildozer 构建 Android APK
7. **上传产物** - 上传构建好的 APK 文件
8. **创建 Release** - 自动创建 GitHub Release（仅 main/master 分支）

## 查看构建状态

### 方法 1：GitHub 网站

1. 打开仓库页面
2. 点击 "Actions" 标签
3. 查看工作流运行状态

### 方法 2：邮件通知

- 构建成功或失败会发送邮件通知
- 在 GitHub 设置中配置通知偏好

## 常见问题

### Q: 构建失败怎么办？

**A:** 
1. 点击失败的构建查看日志
2. 检查错误信息
3. 修复代码后重新推送
4. 常见错误：
   - 依赖版本冲突 → 检查 `requirements.txt`
   - 构建配置错误 → 检查 `buildozer.spec`
   - 网络问题 → 重新运行工作流

### Q: 如何重新运行构建？

**A:**
1. 打开 Actions 页面
2. 点击失败的构建
3. 点击右上角的 "Re-run jobs" 按钮

### Q: 构建时间多长？

**A:**
- 首次构建：15-30 分钟（需要下载依赖）
- 后续构建：5-15 分钟（使用缓存）

### Q: 如何加快构建速度？

**A:**
- 使用缓存（已配置）
- 避免频繁推送
- 使用 `workflow_dispatch` 手动触发

### Q: APK 文件在哪里下载？

**A:**
1. 构建完成后，在 Actions 页面
2. 点击 "Artifacts" 部分的 "android-apk"
3. 或者查看 Releases 页面（自动发布）

### Q: 如何修改构建配置？

**A:**
编辑 `.github/workflows/build-android.yml` 文件：
- 修改 Python 版本
- 添加环境变量
- 修改构建命令
- 添加更多步骤

### Q: 私有仓库可以用吗？

**A:** 可以，GitHub Actions 对私有仓库也免费，但有使用限制：
- 免费账户：每月 2000 分钟
- Pro 账户：每月 3000 分钟

### Q: 如何调试构建问题？

**A:**
1. 在本地使用 Docker 测试
2. 添加更多日志输出
3. 使用 `tmate` 进行 SSH 调试

## 高级配置

### 手动触发构建

在 Actions 页面点击 "Run workflow" 按钮即可手动触发。

### 定时构建

修改 `.github/workflows/build-android.yml`：

```yaml
on:
  schedule:
    # 每天凌晨 2 点构建
    - cron: '0 2 * * *'
```

### 多版本构建

可以同时构建 debug 和 release 版本：

```yaml
- name: Build Debug APK
  run: buildozer android debug

- name: Build Release APK
  run: buildozer android release
```

### 签名 APK

添加签名步骤：

```yaml
- name: Sign APK
  uses: r0adkll/sign-android-release@v1
  with:
    releaseDirectory: mobile_app/bin
    signingKeyBase64: ${{ secrets.SIGNING_KEY }}
    alias: ${{ secrets.ALIAS }}
    keyStorePassword: ${{ secrets.KEY_STORE_PASSWORD }}
```

## 替代方案

如果 GitHub Actions 不满足需求，还可以使用：

1. **GitLab CI/CD** - 类似 GitHub Actions
2. **Bitrise** - 专门用于移动应用构建
3. **CircleCI** - 云端 CI/CD 服务
4. **Travis CI** - 开源项目免费
5. **Jenkins** - 自建 CI/CD 服务器

## 获取帮助

- GitHub Actions 文档：https://docs.github.com/cn/actions
- Buildozer 文档：https://buildozer.readthedocs.io/
- Kivy 文档：https://kivy.org/doc/stable/

## 总结

使用 GitHub Actions 的优势：
- ✅ 无需本地 Linux 环境
- ✅ 无需担心磁盘空间
- ✅ 自动构建，省时省力
- ✅ 免费使用
- ✅ 可查看构建历史
- ✅ 自动发布 Release

按照本指南操作，即可轻松构建 Android APK！
