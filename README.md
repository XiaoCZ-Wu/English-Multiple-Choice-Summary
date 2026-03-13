# 英语单选题练习系统

一个基于 PySide6 开发的英语单选题练习与错题管理桌面应用程序，支持 AI 辅助学习和 OCR 智能识别。

## 功能特性

### 📚 练习模式
- **无尽模式**：无限随机练习题目，适合日常刷题
- **套题模式**：按试卷/来源进行完整练习，模拟真实考试（未做）
- **计时功能**：记录单题用时和总用时，分析答题速度
- **答题统计**：追踪每道题的正确率和练习次数
- **智能推荐**：根据正确率优先推荐易错题

### 🤖 AI 功能
- **AI 对话**：与 AI 助手讨论题目，支持上下文连续对话
- **题目解析**：使用 AI 自动生成详细的题目解析
- **OCR 识别**：使用 AI 识别图片中的英语题目
- **多 AI 配置**：支持配置多个 AI 服务，OCR 和对话可使用不同 AI

### 📷 OCR 识别功能
- **图片识别**：支持批量导入图片进行 OCR 识别
- **截图识别**：支持截图后直接识别题目
- **局域网服务**：启动局域网服务器，手机浏览器访问即可上传图片识别
- **批量导入**：识别后的题目可批量导入题库
- **题号范围**：支持设置识别的题号范围

### 📝 错题管理
- **题目录入**：手动添加新的英语单选题
- **题目编辑**：修改已有题目内容
- **题目删除**：移除不需要的题目
- **分类管理**：9大分类（交际用语、词义辨析、时态、非谓语动词、定语从句、状语从句、情态动词、名词性从句、代词）
- **来源管理**：为题目设置来源/试卷名称，便于分类管理
- **正确率追踪**：记录每道题的练习次数和正确次数

### 📤 导出功能
- **支持格式**：Word (.docx)、PDF (.pdf)、CSV (.csv)
- **筛选导出**：可按正确率、分类、来源等条件筛选导出
- **自定义内容**：可选择是否包含答案、答题卡（未做）、出处、解析（未做）等信息
- **排版美观**：Word 导出支持自定义字体和排版

### 💾 数据管理
- **数据备份**：一键备份题目和配置为 ZIP 文件
- **数据恢复**：从备份文件恢复数据，自动保留原数据
- **自动备份**：导入备份前自动备份当前数据
- **重新加载**：重新加载 JSON 数据文件
- **正确率重置**：重置题目的正确率统计

## 开发要求

- **开发使用的操作系统**：Windows 11
- **开发使用的 Python 版本**：3.13
- **依赖库**：requirements.txt

## 安装与运行

### 方法一：直接运行源码

1. **克隆或下载项目**
```bash
git clone https://github.com/XiaoCZ-Wu/English-Multiple-Choice-Summary.git
cd 英语单选
```

2. **创建虚拟环境（推荐）**
```bash
python -m venv venv
```

3. **激活虚拟环境**
```bash
# Windows
venv\Scripts\activate
```

4. **安装依赖**
```bash
venv\Scripts\pip install -r requirements.txt
```

5. **运行程序**
```bash
venv\Scripts\python.exe main.pyw
```

### 方法二：打包为可执行文件

```bash
# 使用 PyInstaller 打包
ToEXE.bat

# 打包后的文件在 dist/英语单选错题总结/目录下
```

## 使用指南

### 首次使用

1. **启动程序**：运行 `venv\Scripts\python.exe main.pyw` 或打包后的可执行文件
2. **配置 AI**：进入"设置"页面，添加 AI 配置（用于 OCR 和对话功能）
3. **开始练习**：返回首页，选择练习模式开始练习

### AI 配置说明

本软件支持配置多个 AI 服务，OCR 识别和 AI 对话可以使用不同的 AI。

#### 通用配置参数

| 参数 | 说明 | 示例 |
|------|------|------|
| **名称** | 自定义名称，用于区分不同配置 | 智谱 AI、OpenAI、DeepSeek 等 |
| **Base URL** | AI 服务的 API 地址 | `https://api.openai.com/v1` |
| **Model** | 使用的模型 ID | `gpt-4o`、`glm-4v-flash` 等 |
| **API Key** | 从 AI 平台获取的密钥 | sk-xxxxxxxxxx |

#### 支持的 AI 服务

理论上支持所有兼容 OpenAI API 格式的 AI 服务，包括但不限于：

- **OpenAI** - GPT-4、GPT-3.5 等系列模型
- **智谱 AI** - GLM-4、GLM-3 等系列模型**（有免费内容，推荐使用）**
- **DeepSeek** - DeepSeek-V3、DeepSeek-R1 等
- **Azure OpenAI** - 微软 Azure 上的 OpenAI 服务
- **其他兼容服务** - 任何支持 OpenAI API 格式的服务

#### OCR 识别注意事项

OCR 识别需要使用**支持视觉的多模态模型**，例如：
- GPT-4o、GPT-4 Vision
- GLM-4V、GLM-4V Flash
- 其他支持图片输入的模型

普通文本模型无法识别图片内容。

### OCR 识别使用指南

#### 软件端识别
1. 点击"OCR 识别"按钮打开识别窗口
2. 点击"浏览"选择图片文件，或点击"截图"进行屏幕截图
3. 设置题号范围（可选）
4. 点击"开始识别"
5. 识别完成后，勾选需要导入的题目，点击"导入选中题目"

#### 手机端识别（局域网）
1. 在设置页面启动"局域网服务"
2. 查看控制台输出的局域网地址（如 `http://192.168.1.100:8080`）
3. 手机连接同一 WiFi，浏览器访问该地址
4. 上传图片并设置题号范围
5. 点击识别，完成后在软件端导入题目

**注意**：
- OCR 识别需要使用支持视觉的多模态模型（如 glm-4v-flash、gpt-4o）
- 图片大小建议不超过 5MB
- 软件端和手机端可同时使用，互不干扰
- 测试使用`ngrok http <lan_port(default 8080)>`从非局域网环境访问可行，但**注意个人隐私**！

### 练习模式说明

#### 无尽模式
- 从题库中随机抽取题目
- 答对自动进入下一题
- 可随时点击"生成报告"结束练习

#### 套题模式（未做）
- 选择特定来源/试卷进行练习
- 按顺序答题
- 适合模拟真实考试

### 数据备份与恢复

#### 备份数据
1. 进入"设置"页面
2. 点击"Backup"按钮
3. 备份文件保存在 `backup/backup_YYMMDD_HHMMSS.zip`

#### 恢复数据
1. 进入"设置"页面
2. 点击"导入备份"按钮
3. 选择备份文件（zip文件）
4. 程序会自动重启并加载新数据

**注意**：导入备份前会自动备份当前数据到 `backup/restore_points/` 目录

## 项目结构

```
English-Multiple-Choice-Summary/
├── main.pyw                     		# 程序入口（Windows 窗口化）
├── requirements.txt             		# 依赖列表
├── README.md                   	 	# 项目说明
├── ToEXE.bat                   	 	# 打包脚本
├── src/                         		# 源代码目录
│   ├── __init__.py
│   ├── main.py                  		# 应用程序主入口
│   ├── models/                  		# 数据模型层
│   │   ├── __init__.py
│   │   ├── question.py          		# 题目数据模型
│   │   ├── config.py            		# 配置模型
│   │   └── data_manager.py      		# 数据管理器
│   ├── core/                    		# 核心业务逻辑
│   │   ├── __init__.py
│   │   ├── practice_manager.py  		# 练习管理
│   │   ├── export_manager.py    		# 导出功能
│   │   ├── statistics.py        		# 统计功能
│   │   └── lan_server.py        		# 局域网服务器
│   ├── ui/                      		# 界面层
│   │   ├── __init__.py
│   │   ├── main_window.py       		# 主窗口
│   │   ├── ocr_window.py        		# OCR 识别窗口
│   │   ├── screenshot_tool.py   		# 截图工具
│   │   └── dialogs/
│   │       ├── __init__.py
│   │       ├── export_dialog.py 		# 导出对话框
│   │       ├── ai_chat_dialog.py		# AI 对话对话框
│   │       ├── ai_config_dialog.py 	# AI 配置对话框
│   │       └── source_dialog.py 		# 来源设置对话框
│   ├── utils/                   		# 工具函数
│   │   ├── __init__.py
│   │   ├── constants.py         		# 常量定义
│   │   ├── helpers.py           		# 辅助函数
│   │   └── logger.py            		# 日志模块
│   ├── ui_dir/                  		# UI 文件
│   │   └── main.ui              		# Qt Designer 设计的界面
│   └── ico/                     		# 图标文件
│       └── ico.ico
├── data/                        		# 数据目录
│   ├── questions.json           		# 题目数据
│   └── config.json              		# 配置文件
├── log/                         		# 日志目录
│   └── app_YYMMDD_HHMMSS.log    		# 运行日志
├── backup/                      		# 备份目录
│   └── restore_points/          		# 自动备份目录
└── output/                      		# 输出目录
```

## 数据文件说明

### questions.json
存储所有题目数据，JSON 数组格式，每个题目包含以下字段：

| 字段 | 类型 | 说明 |
|------|------|------|
| `question` | string | 题目内容 |
| `A` | string | 选项 A |
| `B` | string | 选项 B |
| `C` | string | 选项 C |
| `D` | string | 选项 D |
| `answer` | string | 正确答案，值为 "A"、"B"、"C" 或 "D" |
| `classification` | int | 分类索引，0-8 对应 9 大分类 |
| `source` | string | 来源/试卷名称 |
| `analysis` | string | 题目解析 |
| `total` | int | 练习次数，默认为 0 |
| `correct` | int | 正确次数，默认为 0 |

### config.json
存储应用程序配置，包含以下字段：

| 字段 | 类型 | 说明 |
|------|------|------|
| `version` | string | 配置文件版本 |
| `font_name` | string | 界面字体名称 |
| `font_size` | int | 字体大小 |
| `output_dir` | string | 导出文件保存目录，默认 `.\output\` |
| `ai_configs` | array | AI 配置列表 |
| `ocr_ai_name` | string | OCR 使用的 AI 名称 |
| `chat_ai_name` | string | 对话使用的 AI 名称 |
| `lan_port` | int | 局域网服务端口，默认 8080 |

#### AI 配置项结构
每个 AI 配置包含以下字段：
- `name`：AI 名称（用于显示和选择）
- `base_url`：API 基础地址
- `model`：模型名称
- `api_key`：API 密钥

## 开发规范

### 代码组织原则

1. **分层架构**：
   - `models/`：数据模型层，负责数据结构和数据访问
   - `core/`：业务逻辑层，负责核心业务功能
   - `ui/`：界面层，负责用户界面交互
   - `utils/`：工具层，提供通用工具函数

2. **单一职责原则**：
   - 每个模块只负责一个明确的功能
   - 每个类只负责一个明确的职责

3. **依赖关系**：
   - UI 层依赖 Core 层和 Models 层
   - Core 层依赖 Models 层
   - Models 层和 Utils 层不依赖其他层

### 命名规范

- **类名**：使用 PascalCase（如 `MainWindow`, `Question`）
- **函数名**：使用 snake_case（如 `start_practice`, `export_questions`）
- **私有方法**：以下划线开头（如 `_setup_ui`, `_on_confirm`）
- **常量**：使用 UPPER_SNAKE_CASE（如 `MAX_OPTIONS`, `DEFAULT_TIME_LIMIT`）

## 日志系统

程序运行时会自动记录日志，保存在 `log/app_YYMMDD_HHMMSS.log` 文件中。

日志内容包括：
- 程序启动和关闭
- 配置加载和保存
- 数据备份和恢复
- OCR 识别过程
- AI 对话记录
- 错误和异常信息

## 常见问题

### Q: 如何配置 AI 才能使用 OCR 功能？
A: 进入设置页面，添加 AI 配置。OCR 识别需要使用支持视觉的多模态模型，如智谱 AI 的 `glm-4v-flash` 或 OpenAI 的 `gpt-4o`。

### Q: 为什么 OCR 识别失败？
A: 可能原因：
1. 未配置 AI 或 API Key 错误
2. 使用的模型不支持图片识别
3. 图片过大（建议不超过 5MB）
4. 网络连接问题

### Q: 如何同时使用手机端和软件端？
A: 可以同时使用。软件端启动时会清理临时目录，但会保留网页端上传的文件（以 `web_` 开头的文件）。

### Q: 数据保存在哪里？
A: 题目数据保存在 `data/questions.json`，配置保存在 `data/config.json`。建议定期备份。

### Q: 如何重置正确率？
A: 进入题目管理页面，点击"重新加载"按钮，选择"重置正确率"。

## 更新日志

### v1.3
- 新增日志系统
- 优化配置导入导出
- 修复网页端和软件端冲突问题
- 优化套题下拉菜单更新逻辑

### v1.2
- 新增 AI 对话功能
- 优化 OCR 识别提示词
- 修复配置升级问题
- 新增局域网 OCR 服务

### v1.1
- 新增 OCR 识别功能
- 支持多 AI 配置
- 优化导出功能

### v1.0
- 初始版本发布
- 基础练习功能
- 错题管理功能

## 许可证

Apache-2.0 License

---

**作者**：AI Assistant  
**项目地址**：[GitHub 地址]  
**问题反馈**：[Issues 页面]
