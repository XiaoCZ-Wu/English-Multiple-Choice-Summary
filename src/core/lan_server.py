"""
局域网Web服务器
提供Web版OCR功能，支持手机浏览器访问
"""

import os
import json
import base64
import tempfile
import re
import logging
from typing import Dict, List, Optional
from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
from werkzeug.utils import secure_filename

try:
    from src.models import DataManager, AppConfig, Question
    from src.utils import CLASSIFICATIONS, app_logger
    from src.utils.constants import OCR_TEMP_DIR
except ImportError:
    from ..models import DataManager, AppConfig, Question
    from ..utils import CLASSIFICATIONS, app_logger
    from ..utils.constants import OCR_TEMP_DIR

# 获取logger
logger = logging.getLogger(__name__)


class LanServer:
    """局域网Web服务器"""
    
    def __init__(self, data_manager: DataManager, config: AppConfig, port: int = 8080):
        # 获取静态文件目录
        from src.utils.constants import get_resource_path
        static_dir = get_resource_path('src/core/static')
        
        self.app = Flask(__name__, static_folder=static_dir, static_url_path='/static')
        # 配置安全选项
        self.app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB
        
        # 允许跨域，但限制为局域网内访问
        CORS(self.app, resources={
            r"/api/*": {
                "origins": "*",  # 局域网内允许所有来源
                "methods": ["GET", "POST"],
                "allow_headers": ["Content-Type"]
            }
        })
        
        self.data_manager = data_manager
        self.config = config
        self.port = port
        
        # 存储上传的图片信息 {image_id: {path, filename, question_range}}
        self.uploaded_images = {}
        
        self._setup_routes()
    
    def _setup_routes(self):
        """设置路由"""
        
        @self.app.route('/')
        def index():
            """Web OCR页面"""
            return render_template_string(HTML_TEMPLATE)
        
        @self.app.route('/api/ping', methods=['GET'])
        def ping():
            """测试连接"""
            return jsonify({'status': 'ok', 'message': 'Server is running'})
        
        @self.app.route('/api/ai-info', methods=['GET'])
        def get_ai_info():
            """获取当前OCR AI信息"""
            ocr_ai = None
            if self.config.ocr_ai_name:
                ocr_ai = self.config.get_ai_config(self.config.ocr_ai_name)

            if ocr_ai:
                return jsonify({
                    'status': 'ok',
                    'ai_name': ocr_ai.name,
                    'has_key': bool(ocr_ai.api_key)
                })
            else:
                return jsonify({
                    'status': 'error',
                    'message': '未配置OCR AI'
                })

        @self.app.route('/api/classifications', methods=['GET'])
        def get_classifications():
            """获取题目分类列表"""
            return jsonify({
                'status': 'ok',
                'classifications': CLASSIFICATIONS
            })

        @self.app.route('/api/upload-image', methods=['POST'])
        def upload_image():
            """上传图片到服务器，保存到temp/ocr目录"""
            try:
                if 'image' not in request.files:
                    return jsonify({'error': '未上传图片'}), 400

                file = request.files['image']
                if file.filename == '':
                    return jsonify({'error': '未选择图片'}), 400

                # 验证文件类型
                allowed_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
                filename_lower = file.filename.lower()
                if not any(filename_lower.endswith(ext) for ext in allowed_extensions):
                    return jsonify({'error': f'不支持的文件格式: {file.filename}'}), 400

                # 保存图片到 OCR 临时目录
                ocr_temp_dir = OCR_TEMP_DIR
                os.makedirs(ocr_temp_dir, exist_ok=True)

                # 生成唯一文件名
                import uuid
                ext = os.path.splitext(file.filename)[1]
                unique_filename = f"web_{uuid.uuid4().hex}{ext}"
                image_path = os.path.join(ocr_temp_dir, unique_filename)
                file.save(image_path)

                # 生成图片ID
                image_id = uuid.uuid4().hex
                self.uploaded_images[image_id] = {
                    'path': image_path,
                    'filename': file.filename,
                    'question_range': ''
                }

                logger.info(f"[Web OCR] 图片上传成功: {file.filename} -> {unique_filename}, ID: {image_id}")

                return jsonify({
                    'status': 'ok',
                    'image_id': image_id,
                    'filename': file.filename
                })

            except Exception as e:
                logger.error(f"[Web OCR] 图片上传失败: {e}")
                return jsonify({'error': str(e)}), 500

        @self.app.route('/api/remove-image', methods=['POST'])
        def remove_image():
            """删除上传的图片"""
            try:
                data = request.get_json()
                image_id = data.get('image_id')

                if not image_id or image_id not in self.uploaded_images:
                    return jsonify({'error': '无效的图片ID'}), 400

                image_info = self.uploaded_images[image_id]
                image_path = image_info['path']

                # 删除本地文件
                try:
                    if os.path.exists(image_path):
                        os.remove(image_path)
                        logger.info(f"[Web OCR] 删除图片文件: {image_path}")
                except Exception as e:
                    logger.warning(f"[Web OCR] 删除图片文件失败: {e}")

                # 从字典中移除
                del self.uploaded_images[image_id]

                return jsonify({'status': 'ok'})

            except Exception as e:
                logger.error(f"[Web OCR] 删除图片失败: {e}")
                return jsonify({'error': str(e)}), 500

        @self.app.route('/api/update-image-range', methods=['POST'])
        def update_image_range():
            """更新图片的题号范围"""
            try:
                data = request.get_json()
                image_id = data.get('image_id')
                question_range = data.get('question_range', '')

                if not image_id or image_id not in self.uploaded_images:
                    return jsonify({'error': '无效的图片ID'}), 400

                self.uploaded_images[image_id]['question_range'] = question_range
                return jsonify({'status': 'ok'})

            except Exception as e:
                logger.error(f"[Web OCR] 更新题号范围失败: {e}")
                return jsonify({'error': str(e)}), 500

        @self.app.route('/api/recognize', methods=['POST'])
        def recognize():
            """识别上传的图片"""
            try:
                data = request.get_json()
                image_ids = data.get('image_ids', [])
                generate_analysis = data.get('generate_analysis', False)

                if not image_ids:
                    return jsonify({'error': '未选择图片'}), 400

                # 获取OCR AI配置
                ocr_ai = None
                if self.config.ocr_ai_name:
                    ocr_ai = self.config.get_ai_config(self.config.ocr_ai_name)

                if not ocr_ai or not ocr_ai.api_key:
                    error_msg = '未配置OCR AI或API Key为空'
                    app_logger.error(f"[Web OCR错误] {error_msg}")
                    return jsonify({'error': error_msg}), 400

                all_questions = []
                failed_images = []
                processed_count = 0

                for idx, image_id in enumerate(image_ids):
                    if image_id not in self.uploaded_images:
                        failed_images.append(image_id)
                        continue

                    image_info = self.uploaded_images[image_id]
                    image_path = image_info['path']
                    question_range = image_info['question_range']

                    # 检查文件是否存在
                    if not os.path.exists(image_path):
                        error_msg = f"[Web OCR] 图片文件不存在: {image_path}，请重新上传图片"
                        app_logger.error(error_msg)
                        logger.error(error_msg)
                        failed_images.append(image_info['filename'])
                        continue

                    logger.info(f"[Web OCR] 正在识别第 {idx + 1}/{len(image_ids)} 张图片: {image_info['filename']}")

                    try:
                        questions = self._perform_ocr(image_path, question_range, ocr_ai, generate_analysis)
                        all_questions.extend(questions)
                        processed_count += 1
                        logger.info(f"[Web OCR] 第 {idx + 1}/{len(image_ids)} 张图片识别成功: {image_info['filename']}, 识别出 {len(questions)} 道题目")
                    except Exception as e:
                        error_msg = f"[Web OCR] 第 {idx + 1}/{len(image_ids)} 张图片识别失败: {image_info['filename']}, 错误: {e}"
                        app_logger.error(error_msg)
                        logger.error(error_msg)
                        failed_images.append(image_info['filename'])

                return jsonify({
                    'status': 'ok',
                    'questions': all_questions,
                    'total_images': len(image_ids),
                    'processed_count': processed_count,
                    'failed_images': failed_images
                })

            except Exception as e:
                error_msg = f"[Web OCR] 识别失败: {e}"
                app_logger.error(error_msg)
                import traceback
                app_logger.error(traceback.format_exc())
                logger.error(error_msg)
                return jsonify({'error': str(e)}), 500

        @self.app.route('/api/clear-images', methods=['POST'])
        def clear_images():
            """清空所有上传的图片"""
            try:
                # 删除所有本地文件
                for image_id, image_info in list(self.uploaded_images.items()):
                    try:
                        if os.path.exists(image_info['path']):
                            os.remove(image_info['path'])
                    except:
                        pass

                # 清空字典
                self.uploaded_images.clear()
                logger.info("[Web OCR] 清空所有上传的图片")

                return jsonify({'status': 'ok'})

            except Exception as e:
                logger.error(f"[Web OCR] 清空图片失败: {e}")
                return jsonify({'error': str(e)}), 500

        @self.app.route('/api/ocr', methods=['POST'])
        def ocr():
            """OCR识别 - 支持多图"""
            try:
                # 检查是否有图片
                if 'images' not in request.files:
                    return jsonify({'error': '未上传图片'}), 400
                
                files = request.files.getlist('images')
                if not files or all(f.filename == '' for f in files):
                    return jsonify({'error': '未选择图片'}), 400
                
                # 获取参数
                question_range = request.form.get('question_range', '')
                generate_analysis = request.form.get('generate_analysis', 'false') == 'true'
                
                # 获取OCR AI配置
                ocr_ai = None
                if self.config.ocr_ai_name:
                    ocr_ai = self.config.get_ai_config(self.config.ocr_ai_name)
                
                if not ocr_ai or not ocr_ai.api_key:
                    return jsonify({'error': '未配置OCR AI或API Key为空'}), 400
                
                all_questions = []
                temp_files = []
                
                try:
                    # 处理每张图片
                    for file in files:
                        if file.filename == '':
                            continue
                        
                        # 验证文件类型
                        allowed_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
                        filename_lower = file.filename.lower()
                        if not any(filename_lower.endswith(ext) for ext in allowed_extensions):
                            return jsonify({'error': f'不支持的文件格式: {file.filename}'}), 400
                        
                        # 保存图片到 OCR 临时目录，与本地OCR保持一致
                        ocr_temp_dir = OCR_TEMP_DIR
                        os.makedirs(ocr_temp_dir, exist_ok=True)
                        
                        filename = secure_filename(file.filename)
                        image_path = os.path.join(ocr_temp_dir, f"web_ocr_{filename}")
                        file.save(image_path)
                        temp_files.append(image_path)
                        
                        # 执行OCR识别
                        questions = self._perform_ocr(image_path, question_range, ocr_ai, generate_analysis)
                        all_questions.extend(questions)
                    
                    return jsonify({
                        'status': 'ok',
                        'questions': all_questions,
                        'total_images': len(temp_files)
                    })
                    
                finally:
                    # 清理临时文件
                    for temp_file in temp_files:
                        try:
                            if os.path.exists(temp_file):
                                os.remove(temp_file)
                        except:
                            pass
                
            except Exception as e:
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/api/import', methods=['POST'])
        def import_questions():
            """导入题目到题库"""
            try:
                data = request.get_json()
                questions = data.get('questions', [])
                
                if not questions:
                    return jsonify({'error': '没有题目数据'}), 400
                
                imported_count = 0
                for q in questions:
                    # 获取分类索引
                    classification = q.get('classification', '')
                    if classification in CLASSIFICATIONS:
                        classification_idx = CLASSIFICATIONS.index(classification)
                    else:
                        classification_idx = 0
                    
                    # 创建题目对象
                    question = Question(
                        question=q.get('question', ''),
                        A=q.get('A', ''),
                        B=q.get('B', ''),
                        C=q.get('C', ''),
                        D=q.get('D', ''),
                        answer=q.get('answer', '').upper(),
                        classification=classification_idx,
                        source=q.get('source', 'Web OCR'),
                        analysis=q.get('analysis', '')
                    )
                    
                    # 添加到数据管理器
                    self.data_manager.questions.append(question)
                    
                    # 添加来源
                    if question.source and question.source not in self.data_manager.papers:
                        self.data_manager.papers.append(question.source)
                    
                    imported_count += 1
                
                # 保存
                if self.data_manager.save():
                    return jsonify({
                        'status': 'ok',
                        'imported': imported_count
                    })
                else:
                    return jsonify({'error': '保存数据失败'}), 500
                    
            except Exception as e:
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/api/ocr-single', methods=['POST'])
        def ocr_single():
            """OCR识别 - 单张图片，与ocr.ui逻辑相同"""
            try:
                # 检查是否有图片
                if 'image' not in request.files:
                    return jsonify({'error': '未上传图片'}), 400
                
                file = request.files['image']
                if file.filename == '':
                    return jsonify({'error': '未选择图片'}), 400
                
                # 获取参数
                question_range = request.form.get('question_range', '')
                generate_analysis = request.form.get('generate_analysis', 'false') == 'true'
                
                # 获取OCR AI配置
                ocr_ai = None
                if self.config.ocr_ai_name:
                    ocr_ai = self.config.get_ai_config(self.config.ocr_ai_name)
                
                if not ocr_ai or not ocr_ai.api_key:
                    return jsonify({'error': '未配置OCR AI或API Key为空'}), 400
                
                # 验证文件类型
                allowed_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
                filename_lower = file.filename.lower()
                if not any(filename_lower.endswith(ext) for ext in allowed_extensions):
                    return jsonify({'error': f'不支持的文件格式: {file.filename}'}), 400
                
                # 保存图片到 OCR 临时目录，与本地OCR保持一致
                ocr_temp_dir = OCR_TEMP_DIR
                os.makedirs(ocr_temp_dir, exist_ok=True)
                
                filename = secure_filename(file.filename)
                image_path = os.path.join(ocr_temp_dir, f"web_ocr_{filename}")
                file.save(image_path)
                
                try:
                    # 执行OCR识别
                    questions = self._perform_ocr(image_path, question_range, ocr_ai, generate_analysis)
                    
                    return jsonify({
                        'status': 'ok',
                        'questions': questions
                    })
                finally:
                    # 清理临时文件
                    try:
                        if os.path.exists(image_path):
                            os.remove(image_path)
                    except:
                        pass
                
            except Exception as e:
                return jsonify({'error': str(e)}), 500
    
    def _resize_image_for_ocr(self, image_path: str, max_size: int = 1920, quality: int = 85) -> str:
        """
        调整图片大小以适应OCR识别
        返回临时文件路径
        """
        try:
            from PIL import Image
            
            with Image.open(image_path) as img:
                # 转换为RGB模式（处理RGBA等模式）
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                
                # 检查是否需要调整大小
                width, height = img.size
                if width > max_size or height > max_size:
                    # 计算缩放比例
                    ratio = min(max_size / width, max_size / height)
                    new_width = int(width * ratio)
                    new_height = int(height * ratio)
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    logger.info(f"[Web OCR] 图片已缩放: {width}x{height} -> {new_width}x{new_height}")
                
                # 保存为临时文件
                temp_fd, temp_path = tempfile.mkstemp(suffix='.jpg')
                os.close(temp_fd)
                img.save(temp_path, 'JPEG', quality=quality, optimize=True)
                
                # 检查文件大小
                file_size = os.path.getsize(temp_path)
                logger.info(f"[Web OCR] 压缩后图片大小: {file_size / 1024:.1f} KB")
                
                return temp_path
                
        except Exception as e:
            logger.warning(f"[Web OCR] 图片压缩失败，使用原图: {e}")
            return image_path

    def _perform_ocr(self, image_path: str, question_range: str, ai_config, generate_analysis: bool = False) -> List[Dict]:
        """执行OCR识别"""
        import requests
        
        # 使用模块级别的logger
        
        # 直接读取原图，不进行压缩
        with open(image_path, 'rb') as f:
            image_data = base64.b64encode(f.read()).decode('utf-8')
        
        # 根据文件扩展名确定MIME类型
        ext = os.path.splitext(image_path)[1].lower()
        mime_types = {
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.webp': 'image/webp'
        }
        mime_type = mime_types.get(ext, 'image/jpeg')  # 默认为jpeg

        # 解析题号范围
        question_numbers = self._parse_question_range(question_range)
        if question_numbers:
            question_list_str = ', '.join(map(str, question_numbers))
            question_range_desc = f"图片中只包含以下题号的题目：{question_list_str}"
        else:
            question_list_str = "全部"
            question_range_desc = "请识别图片中的所有题目"

        # 构建分类选项字符串
        classifications_str = '、'.join(CLASSIFICATIONS)

        logger.info(f"[Web OCR] 开始识别图片: {os.path.basename(image_path)}")
        logger.info(f"[Web OCR] 题号范围: {question_range if question_range else '全部'}")
        logger.info(f"[Web OCR] 使用AI: {ai_config.name if hasattr(ai_config, 'name') else 'Unknown'}")

        # 构建提示词
        if generate_analysis:
            prompt = f"""请识别图片中的英语选择题。

【重要说明】
- {question_range_desc}
- 【关键】必须识别图片中所有可见的题目，一道都不能遗漏，请仔细检查图片的每个角落
- 严格按照图片中实际显示的题号来填写
- 【关键】题目内容必须包含原始题号，格式如："1. What is..." 或 "(1) What is..."
- 请根据题目内容判断正确答案（A/B/C/D）
- 请从以下分类中选择最符合的一个：{classifications_str}
- 【关键】题目解析请使用中文回答

【输出格式要求】
请使用HTML标签格式输出，每道题用<question>标签包裹，每个字段用对应的标签：

```
<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
<ana>解析内容（中文）</ana>
</question>

<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
<ana>解析内容（中文）</ana>
</question>
```

【重要】如果图片中有多个题目，必须为每道题都创建一个<question>标签块，不要只返回一道题。确保所有内容都在代码块内。"""
        else:
            prompt = f"""请识别图片中的英语选择题。

【重要说明】
- {question_range_desc}
- 【关键】必须识别图片中所有可见的题目，一道都不能遗漏，请仔细检查图片的每个角落
- 严格按照图片中实际显示的题号来填写
- 【关键】题目内容必须包含原始题号，格式如："1. What is..." 或 "(1) What is..."
- 请根据题目内容判断正确答案（A/B/C/D）
- 请从以下分类中选择最符合的一个：{classifications_str}

【输出格式要求】
请使用HTML标签格式输出，每道题用<question>标签包裹，每个字段用对应的标签：

```
<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
</question>

<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
</question>
```

【重要】如果图片中有多个题目，必须为每道题都创建一个<question>标签块，不要只返回一道题。确保所有内容都在代码块内。"""

        # 检查模型配置
        if not hasattr(ai_config, 'model') or not ai_config.model:
            error_msg = "[Web OCR] 错误：未配置模型名称"
            app_logger.error(error_msg)
            logger.error(error_msg)
            raise Exception("未配置模型名称，请在AI配置中设置模型")
        
        model_name = ai_config.model
        logger.info(f"[Web OCR] 使用模型: {model_name}")
        
        # 调用AI API
        headers = {
            "Authorization": f"Bearer {ai_config.api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": model_name,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{image_data}"
                            }
                        }
                    ]
                }
            ]
        }
        
        # 构建API URL，处理base_url末尾斜杠
        base_url = ai_config.base_url.rstrip('/')
        api_url = f"{base_url}/chat/completions"
        
        logger.info(f"[Web OCR] 发送请求到 AI API...")
        app_logger.info(f"[Web OCR] 模型: {model_name}")
        app_logger.info(f"[Web OCR] API URL: {api_url}")
        # 输出payload但不包含图片base64数据（太长）
        payload_copy = json.loads(json.dumps(payload))
        payload_copy['messages'][0]['content'][1]['image_url']['url'] = '[图片base64数据已省略]'
        app_logger.info(f"[Web OCR] Payload: {json.dumps(payload_copy, ensure_ascii=False, indent=2)}")
        
        logger.info(f"[Web OCR] API URL: {api_url}")

        response = requests.post(
            api_url,
            headers=headers,
            json=payload,
            timeout=120
        )

        if response.status_code != 200:
            app_logger.error(f"[Web OCR错误] API请求失败: {response.status_code}")
            app_logger.error(f"[Web OCR错误] 响应内容: {response.text}")
            logger.error(f"[Web OCR] AI API错误: {response.status_code} - {response.text}")
            raise Exception(f"AI API错误: {response.status_code}")

        result = response.json()
        content = result['choices'][0]['message']['content']

        logger.debug(f"[Web OCR] AI回复内容:\n{content}")

        # 解析结果
        questions = self._parse_ocr_result(content)

        logger.info(f"[Web OCR] 识别完成，共解析出 {len(questions)} 道题目")
        for i, q in enumerate(questions):
            logger.info(f"[Web OCR] 题目 {i+1}: {q.get('question', '')[:50]}...")

        return questions
    
    def _parse_question_range(self, range_str: str) -> List[int]:
        """解析题号范围字符串"""
        result = []
        if not range_str or not range_str.strip():
            return result
        
        range_str = range_str.replace(' ', '').replace('，', ',')
        
        parts = range_str.split(',')
        for part in parts:
            if not part:
                continue
            if '-' in part:
                try:
                    start, end = part.split('-', 1)
                    start_num = int(start)
                    end_num = int(end)
                    result.extend(range(start_num, end_num + 1))
                except ValueError:
                    continue
            else:
                try:
                    result.append(int(part))
                except ValueError:
                    continue
        
        return sorted(list(set(result)))
    
    def _parse_ocr_result(self, result) -> List[Dict]:
        """解析OCR结果 - 使用HTML解析器"""
        questions = []

        try:
            # 从结果中提取文本
            if hasattr(result, 'content'):
                text = result.content
            elif hasattr(result, 'text'):
                text = result.text
            else:
                text = str(result)

            logger.info(f"[Web OCR] 原始文本长度: {len(text)}")
            logger.info(f"[Web OCR] 原始文本内容:\n{text}")
            logger.info("=" * 80)

            # 首先尝试从代码块中提取内容
            code_pattern = r'```[\s\S]*?\n([\s\S]*?)```'
            code_matches = re.findall(code_pattern, text, re.DOTALL)

            if code_matches:
                logger.info(f"[Web OCR] 找到 {len(code_matches)} 个代码块")
                for i, match in enumerate(code_matches):
                    logger.info(f"[Web OCR] 代码块 {i+1} 内容:\n{match}")
                html_content = '\n'.join(code_matches)
            else:
                logger.info("[Web OCR] 未找到代码块，使用完整文本")
                html_content = text

            logger.info(f"[Web OCR] 用于解析的HTML内容:\n{html_content}")
            logger.info("=" * 80)

            # 使用HTML解析器解析内容
            from html.parser import HTMLParser

            class QuestionParser(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self.questions = []
                    self.current_question = {}
                    self.current_tag = None
                    self.current_data = []
                    self.tag_stack = []

                def handle_starttag(self, tag, attrs):
                    logger.debug(f"[Web OCR HTML解析] 开始标签: <{tag}>")
                    self.current_tag = tag
                    self.tag_stack.append(tag)
                    self.current_data = []

                def handle_endtag(self, tag):
                    logger.debug(f"[Web OCR HTML解析] 结束标签: </{tag}>")
                    content = ''.join(self.current_data).strip()
                    logger.info(f"[Web OCR HTML解析] 标签 <{tag}> 内容: '{content}'")

                    if tag == 'question':
                        if self.current_question:
                            logger.info(f"[Web OCR HTML解析] 完成一道题解析: {self.current_question}")
                            self.questions.append(self.current_question)
                            self.current_question = {}
                        else:
                            logger.warning("[Web OCR HTML解析] 空题目标签，无内容")
                    elif tag in ('q', 'a', 'b', 'c', 'd', 'ans', 'cat', 'ana'):
                        if not content:
                            logger.warning(f"[Web OCR HTML解析] 标签 <{tag}> 内容为空")
                        if tag == 'q':
                            self.current_question['question'] = content
                        elif tag == 'a':
                            self.current_question['A'] = content
                        elif tag == 'b':
                            self.current_question['B'] = content
                        elif tag == 'c':
                            self.current_question['C'] = content
                        elif tag == 'd':
                            self.current_question['D'] = content
                        elif tag == 'ans':
                            self.current_question['answer'] = content.upper()
                        elif tag == 'cat':
                            self.current_question['classification'] = content
                        elif tag == 'ana':
                            self.current_question['analysis'] = content
                    self.current_tag = None
                    self.current_data = []
                    if self.tag_stack and self.tag_stack[-1] == tag:
                        self.tag_stack.pop()

                def handle_data(self, data):
                    if self.current_tag:
                        self.current_data.append(data)

            parser = QuestionParser()
            parser.feed(html_content)

            logger.info(f"[Web OCR] HTML解析器找到 {len(parser.questions)} 道题目")

            # 处理解析结果
            for idx, q in enumerate(parser.questions):
                try:
                    logger.info(f"[Web OCR] 处理第 {idx+1} 道题目原始数据: {q}")

                    classification = q.get('classification', '')
                    if classification not in CLASSIFICATIONS:
                        logger.warning(f"[Web OCR] 题目 {idx+1} 分类 '{classification}' 不在有效分类列表中，置为空")
                        classification = ''

                    question_data = {
                        'question': q.get('question', ''),
                        'A': q.get('A', ''),
                        'B': q.get('B', ''),
                        'C': q.get('C', ''),
                        'D': q.get('D', ''),
                        'answer': q.get('answer', ''),
                        'classification': classification,
                        'source': 'Web OCR',
                        'analysis': q.get('analysis', '')
                    }

                    question_text = question_data['question']
                    number_match = re.match(r'^(\d+)[\.\)\s]', question_text)
                    question_number = number_match.group(1) if number_match else ''

                    logger.info(f"[Web OCR] 题目 {idx+1} 最终数据:")
                    logger.info(f"  题号: {question_number}")
                    logger.info(f"  题目: '{question_text}'")
                    logger.info(f"  A: '{question_data['A']}'")
                    logger.info(f"  B: '{question_data['B']}'")
                    logger.info(f"  C: '{question_data['C']}'")
                    logger.info(f"  D: '{question_data['D']}'")
                    logger.info(f"  答案: '{question_data['answer']}'")
                    logger.info(f"  分类: '{question_data['classification']}'")
                    logger.info(f"  解析: '{question_data['analysis'][:50]}...'" if question_data['analysis'] else "  解析: (空)")

                    if question_data['question']:
                        questions.append(question_data)
                        logger.info(f"[Web OCR] ✓ 成功添加题目 {idx+1} 到列表")
                    else:
                        logger.error(f"[Web OCR] ✗ 题目 {idx+1} 缺少题目内容，跳过")

                except Exception as e:
                    logger.error(f"[Web OCR] ✗ 处理题目 {idx+1} 时发生异常: {e}")
                    import traceback
                    logger.error(f"[Web OCR] 异常详情: {traceback.format_exc()}")
                    continue

            logger.info(f"[Web OCR] 总共成功解析并添加 {len(questions)} 道题目")

            if not questions:
                logger.error("[Web OCR] ✗ 未解析到任何有效题目")
                logger.error(f"[Web OCR] 原始文本:\n{text}")

        except Exception as e:
            logger.error(f"[Web OCR] ✗ 整体解析失败: {e}")
            import traceback
            logger.error(f"[Web OCR] 错误详情: {traceback.format_exc()}")

        return questions
    
    def start(self, host='0.0.0.0', ssl_context=None):
        """启动服务器"""
        self.app.run(host=host, port=self.port, threaded=True, ssl_context=ssl_context)
    
    def start_in_thread(self, host='0.0.0.0', ssl_context=None):
        """在后台线程启动服务器"""
        import threading
        server_thread = threading.Thread(
            target=self.start,
            kwargs={'host': host, 'ssl_context': ssl_context},
            daemon=True
        )
        server_thread.start()
        return server_thread


def start_server(data_manager, config, port=8080, threaded=True, host='0.0.0.0', use_https=False):
    """
    启动局域网服务器的便捷函数
    
    Args:
        data_manager: 数据管理器实例
        config: 应用配置实例
        port: 服务器端口
        threaded: 是否在线程中启动
        host: 监听地址
        use_https: 是否启用HTTPS（用于内网穿透）
    
    Returns:
        如果threaded=True返回线程对象，否则返回None
    """
    server = LanServer(data_manager, config, port)
    
    ssl_context = None
    if use_https:
        # 使用 Flask 内置的 adhoc 证书
        ssl_context = 'adhoc'
        logger.info(f"[Web OCR] 已启用HTTPS支持（使用临时证书）")
    
    if threaded:
        return server.start_in_thread(host, ssl_context)
    else:
        server.start(host, ssl_context)
        return None


def _create_self_signed_cert():
    """创建自签名SSL证书（用于HTTPS）"""
    import tempfile
    import ssl
    
    try:
        # 尝试使用 cryptography 生成证书
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        import datetime
        
        # 生成私钥
        key = rsa.generate_private_key(
            public_exponent=65537,
            key_size=2048,
        )
        
        # 生成证书
        subject = issuer = x509.Name([
            x509.NameAttribute(NameOID.COUNTRY_NAME, u"CN"),
            x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, u"Beijing"),
            x509.NameAttribute(NameOID.LOCALITY_NAME, u"Beijing"),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, u"OCR Server"),
            x509.NameAttribute(NameOID.COMMON_NAME, u"localhost"),
        ])
        
        cert = x509.CertificateBuilder().subject_name(
            subject
        ).issuer_name(
            issuer
        ).public_key(
            key.public_key()
        ).serial_number(
            x509.random_serial_number()
        ).not_valid_before(
            datetime.datetime.utcnow()
        ).not_valid_after(
            datetime.datetime.utcnow() + datetime.timedelta(days=365)
        ).add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName(u"localhost"),
                x509.DNSName(u"*.natfrp.cloud"),
                x509.DNSName(u"*.frp-bus.com"),
            ]),
            critical=False,
        ).sign(key, hashes.SHA256())
        
        # 创建临时文件
        cert_dir = tempfile.mkdtemp()
        key_file = os.path.join(cert_dir, 'key.pem')
        cert_file = os.path.join(cert_dir, 'cert.pem')
        
        # 保存私钥
        with open(key_file, "wb") as f:
            f.write(key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.TraditionalOpenSSL,
                encryption_algorithm=serialization.NoEncryption()
            ))
        
        # 保存证书
        with open(cert_file, "wb") as f:
            f.write(cert.public_bytes(serialization.Encoding.PEM))
        
        logger.info(f"[Web OCR] 自签名证书已生成: {cert_file}")
        return (cert_file, key_file)
        
    except ImportError:
        # 如果没有 cryptography，使用 adhoc 证书
        logger.warning("[Web OCR] 未安装 cryptography，使用 Flask 临时证书")
        return 'adhoc'
    except Exception as e:
        logger.error(f"[Web OCR] 生成证书失败: {e}")
        return None


# HTML模板 - Web OCR页面
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>OCR识别 - 英语单选</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 900px;
            margin: 0 auto;
        }
        
        h1 {
            color: #1565c0;
            text-align: center;
            margin-bottom: 20px;
            font-size: 24px;
        }
        
        h2 {
            color: #1565c0;
            margin-bottom: 15px;
            font-size: 20px;
        }
        
        .card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .form-group {
            margin-bottom: 15px;
        }
        
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            color: #333;
        }
        
        input[type="text"],
        select,
        textarea {
            width: 100%;
            padding: 12px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
            transition: border-color 0.3s;
        }
        
        input[type="text"]:focus,
        select:focus,
        textarea:focus {
            outline: none;
            border-color: #667eea;
        }
        
        textarea {
            resize: vertical;
            min-height: 80px;
        }
        
        .upload-area {
            border: 2px dashed #ccc;
            border-radius: 8px;
            padding: 30px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
            background: #fafafa;
        }
        
        .upload-area:hover {
            border-color: #667eea;
            background: #f0f4ff;
        }
        
        .upload-area.has-images {
            border-style: solid;
            border-color: #667eea;
            background: #f8f9ff;
        }
        
        .image-preview-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
            gap: 10px;
            margin-top: 15px;
        }
        
        .image-preview-item {
            position: relative;
            border-radius: 8px;
            overflow: hidden;
            border: 2px solid #e0e0e0;
            background: #fff;
        }
        
        .image-preview-item img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        
        .image-preview-item .remove-btn {
            position: absolute;
            top: 4px;
            right: 4px;
            width: 24px;
            height: 24px;
            background: rgba(255,0,0,0.8);
            color: white;
            border: none;
            border-radius: 50%;
            cursor: pointer;
            font-size: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        .image-preview-item .image-index {
            position: absolute;
            bottom: 4px;
            left: 4px;
            background: rgba(0,0,0,0.6);
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 12px;
        }
        
        .btn {
            display: inline-block;
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
            width: 100%;
            margin-bottom: 10px;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }
        
        .btn-primary:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }
        
        .btn-success {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
        }
        
        .btn-success:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(17, 153, 142, 0.4);
        }
        
        .btn-danger {
            background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%);
            color: white;
        }
        
        .btn-secondary {
            background: #f0f0f0;
            color: #333;
        }
        
        .btn-secondary:hover {
            background: #e0e0e0;
        }
        
        .btn-small {
            padding: 6px 12px;
            font-size: 14px;
            width: auto;
        }
        
        .btn-row {
            display: flex;
            gap: 10px;
            margin-top: 15px;
        }
        
        .btn-row .btn {
            flex: 1;
            margin-bottom: 0;
        }
        
        /* 题目列表样式 */
        .questions-list {
            max-height: 600px;
            overflow-y: auto;
        }
        
        .question-item {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 15px;
            border-left: 4px solid #667eea;
            transition: all 0.3s;
        }
        
        .question-item:hover {
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .question-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        .question-number {
            font-weight: bold;
            color: #667eea;
            font-size: 18px;
        }
        
        .question-checkbox {
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
        }
        
        .question-checkbox input[type="checkbox"] {
            width: 20px;
            height: 20px;
            cursor: pointer;
        }
        
        .options-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            margin-top: 10px;
        }
        
        .option-input {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .option-input label {
            font-weight: bold;
            color: #667eea;
            min-width: 24px;
            margin: 0;
            text-align: center;
        }
        
        .option-input input {
            flex: 1;
        }
        
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 30px;
        }
        
        .loading.active {
            display: block;
        }
        
        .spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .alert {
            padding: 12px;
            border-radius: 8px;
            margin-bottom: 15px;
        }
        
        .alert-error {
            background: #fee;
            color: #c33;
            border: 1px solid #fcc;
        }
        
        .alert-success {
            background: #efe;
            color: #3c3;
            border: 1px solid #cfc;
        }
        
        .alert-info {
            background: #e8f4f8;
            color: #31708f;
            border: 1px solid #bce8f1;
        }
        
        .hidden {
            display: none !important;
        }
        
        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .checkbox-group input[type="checkbox"] {
            width: 20px;
            height: 20px;
        }
        
        .stats-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px;
            background: #f0f4ff;
            border-radius: 8px;
            margin-bottom: 15px;
        }
        
        .stats-bar label {
            margin: 0;
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
        }
        
        .empty-state {
            text-align: center;
            padding: 40px;
            color: #999;
        }
        
        /* 手机端优化 */
        @media (max-width: 600px) {
            body {
                padding: 10px;
                font-size: 15px;
            }
            
            .container {
                max-width: 100%;
            }
            
            h1 {
                font-size: 20px;
                margin-bottom: 15px;
            }
            
            h2 {
                font-size: 17px;
                margin-bottom: 12px;
            }
            
            .card {
                padding: 15px;
                margin-bottom: 12px;
                border-radius: 10px;
            }
            
            .upload-area {
                padding: 20px 15px;
            }
            
            .upload-area p {
                font-size: 14px;
            }
            
            .image-preview-grid {
                grid-template-columns: repeat(2, 1fr);
                gap: 8px;
            }
            
            .image-preview-item {
                min-height: 100px;
            }
            
            .options-grid {
                grid-template-columns: 1fr;
                gap: 8px;
            }
            
            .form-row {
                grid-template-columns: 1fr;
                gap: 10px;
            }
            
            .btn-row {
                flex-direction: column;
                gap: 8px;
            }
            
            .btn {
                padding: 14px 20px;
                font-size: 15px;
            }
            
            .btn-small {
                padding: 8px 12px;
                font-size: 13px;
            }
            
            input[type="text"],
            select,
            textarea {
                padding: 10px;
                font-size: 16px; /* 防止iOS缩放 */
            }
            
            .question-item {
                padding: 15px;
                margin-bottom: 12px;
            }
            
            .question-header {
                padding: 10px 0;
            }
            
            .question-number {
                font-size: 16px;
            }
            
            .stats-bar {
                flex-direction: column;
                gap: 10px;
                align-items: flex-start;
                padding: 12px;
            }
            
            .questions-list {
                max-height: none; /* 手机端不限制高度 */
            }
            
            .form-group {
                margin-bottom: 12px;
            }
            
            label {
                font-size: 14px;
            }
            
            /* 触摸友好的按钮 */
            .btn-primary:active,
            .btn-success:active,
            .btn-danger:active,
            .btn-secondary:active {
                transform: scale(0.98);
                opacity: 0.9;
            }
            
            /* 图片预览项优化 */
            .image-preview-item .remove-btn {
                width: 28px;
                height: 28px;
                font-size: 16px;
                top: 6px;
                right: 6px;
            }
            
            /* 复选框优化 */
            .checkbox-group input[type="checkbox"] {
                width: 22px;
                height: 22px;
            }
            
            /* 选项输入优化 */
            .option-input {
                gap: 6px;
            }
            
            .option-input label {
                font-size: 15px;
                min-width: 22px;
            }
        }
        
        /* 超小屏幕优化 */
        @media (max-width: 375px) {
            body {
                padding: 8px;
            }
            
            .card {
                padding: 12px;
            }
            
            .image-preview-grid {
                grid-template-columns: repeat(2, 1fr);
                gap: 6px;
            }
            
            h1 {
                font-size: 18px;
            }
            
            h2 {
                font-size: 15px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📷 OCR识别录入</h1>
        
        <!-- AI配置信息 -->
        <div id="ai-info" class="card alert alert-info">
            正在检查AI配置...
        </div>
        
        <!-- 上传区域 -->
        <div id="upload-section" class="card">
            <h2>上传图片</h2>
            
            <div class="form-group">
                <div class="upload-area" id="uploadArea" onclick="document.getElementById('imageInput').click()">
                    <div id="uploadPlaceholder">
                        <p>📸 点击或拖拽上传图片</p>
                        <p style="color: #999; font-size: 14px; margin-top: 8px;">支持多图上传，支持 JPG、PNG 格式</p>
                    </div>
                    <div id="imagePreviewContainer" class="image-preview-grid hidden"></div>
                </div>
                <input type="file" id="imageInput" accept="image/*" multiple style="display: none;">
            </div>

            <div class="form-group checkbox-group">
                <input type="checkbox" id="generateAnalysis">
                <label for="generateAnalysis" style="margin: 0;">同时生成解析</label>
            </div>
            
            <button class="btn btn-primary" id="recognizeBtn" onclick="recognize()" disabled>
                🔍 开始识别 (<span id="imageCount">0</span>张图片)
            </button>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>正在识别，请稍候...</p>
                <p id="loadingDetail" style="color: #666; font-size: 14px; margin-top: 10px;"></p>
            </div>
        </div>

        <!-- 识别结果区域 -->
        <div id="results-section" class="hidden">
            <div class="card">
                <h2>识别结果</h2>
                
                <!-- 统计和操作栏 -->
                <div class="stats-bar">
                    <label>
                        <input type="checkbox" id="selectAll" onchange="toggleSelectAll()">
                        <span>全选</span>
                    </label>
                    <span id="selectedCount">已选择 0/0 题</span>
                </div>
                
                <!-- 题目列表 -->
                <div id="questions-list" class="questions-list">
                    <div class="empty-state">暂无识别结果</div>
                </div>
                
                <!-- 操作按钮 -->
                <div class="btn-row" style="margin-top: 20px;">
                    <button class="btn btn-success" onclick="addSelectedToDatabase()">
                        ✅ 添加选中到题库
                    </button>
                    <button class="btn btn-secondary" onclick="toggleAllQuestions()">
                        📁 折叠/展开全部
                    </button>
                    <button class="btn btn-secondary" onclick="clearResults()">
                        🗑️ 清空结果
                    </button>
                </div>
            </div>
        </div>
    </div>

    <script src="/static/ocr.js"></script>
</body>
</html>
'''