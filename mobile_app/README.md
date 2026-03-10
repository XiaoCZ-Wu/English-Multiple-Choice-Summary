# 英语单选 - 手机端APP

使用Kivy框架开发的Android应用，用于与电脑端的英语单选软件进行局域网通信。

## 功能

1. **连接电脑端**
   - 输入电脑IP地址和端口
   - 自动测试连接

2. **获取AI列表**
   - 从电脑端拉取可用的AI配置
   - 包括AI名称、Base URL和API Key

3. **OCR识别**
   - 拍照或选择图片
   - 设置题号范围
   - 上传到电脑端进行AI识别
   - 获取识别结果

4. **编辑和导入**
   - 查看识别的题目
   - 编辑题目内容、选项、答案等
   - 添加新题目
   - 删除题目
   - 将题目发送到电脑端合并到题库

## 安装和运行

### 开发环境运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行APP
python main.py
```

### 打包为Android APK

需要使用Buildozer工具进行打包：

```bash
# 安装buildozer
pip install buildozer

# 在Linux环境下（推荐Ubuntu）
# 初始化buildozer
cd mobile_app
buildozer init

# 构建APK
buildozer android debug

# 部署到设备
buildozer android debug deploy run
```

**注意：** 打包Android APK需要在Linux环境下进行，Windows不支持直接打包。

## 使用方法

1. **启动电脑端软件**
   - 确保电脑和手机在同一WiFi网络下
   - 在设置页面查看或设置局域网端口（默认8080）
   - 服务器会自动启动

2. **打开手机APP**
   - 输入电脑的IP地址（在电脑端设置页面可以看到）
   - 输入端口号（默认8080）
   - 点击"连接电脑"

3. **使用功能**
   - 连接成功后进入主页面
   - 可以刷新AI列表查看可用AI
   - 点击"拍照/选择图片进行OCR识别"上传图片
   - 在"查看待导入题目"页面编辑和发送题目

## 文件结构

```
mobile_app/
├── main.py              # 主程序入口
├── buildozer.spec       # Buildozer打包配置
├── requirements.txt     # Python依赖
├── README.md           # 说明文档
├── src/                # 源代码目录
└── assets/             # 资源文件目录
```

## API接口

手机APP与电脑端通过HTTP API通信：

### 1. 测试连接
```
GET /api/ping
```

### 2. 获取AI列表
```
GET /api/ai-list
```

### 3. OCR识别
```
POST /api/ocr
Content-Type: multipart/form-data

参数:
- image: 图片文件
- question_range: 题号范围（可选）
```

### 4. 导入题目
```
POST /api/import
Content-Type: application/json

参数:
- questions: 题目列表
```

## 注意事项

1. **网络要求**：手机和电脑必须在同一局域网内
2. **防火墙**：确保电脑的防火墙允许该端口通信
3. **端口冲突**：如果8080端口被占用，可以在设置页面修改
4. **图片大小**：建议上传的图片大小不超过5MB

## 技术栈

- **Kivy**: 跨平台GUI框架
- **Requests**: HTTP请求库
- **Flask**: 电脑端HTTP服务器

## 许可证

MIT License
