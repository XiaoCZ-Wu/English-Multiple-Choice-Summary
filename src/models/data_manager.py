"""
数据管理器
负责题目数据的加载、保存和管理
"""

import json
import os
import zipfile
from datetime import datetime
from typing import List, Optional, Callable

from .question import Question
from src.utils import QUESTIONS_FILE, CONFIG_FILE, BACKUP_DIR, get_timestamp, app_logger


class DataManager:
    """数据管理器类"""

    def __init__(self, filepath: str = QUESTIONS_FILE):
        self.filepath = filepath
        self.questions: List[Question] = []
        self.papers: List[str] = []
        self._on_data_changed: Optional[Callable] = None

    def set_on_data_changed_callback(self, callback: Callable):
        """设置数据变化回调"""
        self._on_data_changed = callback

    def load(self) -> bool:
        """加载题目数据"""
        try:
            if os.path.exists(self.filepath):
                with open(self.filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.questions = [Question.from_dict(q) for q in data]
                self._update_papers()
                return True
        except Exception as e:
            app_logger.error(f"加载数据失败: {e}")
            return False

    def save(self) -> bool:
        """保存题目数据"""
        try:
            os.makedirs(os.path.dirname(self.filepath), exist_ok=True)
            data = [q.to_dict() for q in self.questions]
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if self._on_data_changed:
                self._on_data_changed()
            return True
        except Exception as e:
            app_logger.error(f"保存数据失败: {e}")
            return False

    def backup(self) -> Optional[str]:
        """备份数据，返回备份文件路径"""
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            timestamp = get_timestamp()
            backup_path = os.path.join(BACKUP_DIR, f"backup_{timestamp}.zip")
            
            with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(QUESTIONS_FILE, arcname="questions.json")
                zf.write(CONFIG_FILE, arcname="config.json")
            
            return backup_path
        except Exception as e:
            app_logger.error(f"备份失败: {e}")
            return None

    def import_backup(self, backup_path: str) -> bool:
        """从备份文件导入数据，返回是否成功"""
        try:
            import zipfile
            import shutil

            # 验证备份文件
            if not os.path.exists(backup_path):
                app_logger.error(f"备份文件不存在: {backup_path}")
                return False

            # 创建临时目录解压
            import tempfile
            temp_dir = tempfile.mkdtemp()

            try:
                # 解压备份文件
                with zipfile.ZipFile(backup_path, 'r') as zf:
                    zf.extractall(temp_dir)

                # 检查必要的文件
                temp_questions = os.path.join(temp_dir, "questions.json")
                temp_config = os.path.join(temp_dir, "config.json")

                if not os.path.exists(temp_questions):
                    app_logger.error("备份文件损坏：缺少 questions.json")
                    return False

                # 备份当前数据（防止导入失败）
                timestamp = get_timestamp()
                restore_backup_dir = os.path.join(BACKUP_DIR, "restore_points")
                os.makedirs(restore_backup_dir, exist_ok=True)
                restore_path = os.path.join(restore_backup_dir, f"before_import_{timestamp}.zip")

                with zipfile.ZipFile(restore_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    if os.path.exists(QUESTIONS_FILE):
                        zf.write(QUESTIONS_FILE, arcname="questions.json")
                    if os.path.exists(CONFIG_FILE):
                        zf.write(CONFIG_FILE, arcname="config.json")

                # 复制新数据
                app_logger.info(f"[导入备份] 复制 questions.json: {temp_questions} -> {QUESTIONS_FILE}")
                shutil.copy2(temp_questions, QUESTIONS_FILE)
                
                if os.path.exists(temp_config):
                    app_logger.info(f"[导入备份] 复制 config.json: {temp_config} -> {CONFIG_FILE}")
                    shutil.copy2(temp_config, CONFIG_FILE)
                    # 验证复制是否成功
                    if os.path.exists(CONFIG_FILE):
                        app_logger.info(f"[导入备份] config.json 复制成功，文件大小: {os.path.getsize(CONFIG_FILE)} bytes")
                        # 读取并显示内容
                        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                            config_data = json.load(f)
                        app_logger.info(f"[导入备份] config.json 内容: {json.dumps(config_data, ensure_ascii=False, indent=2)[:200]}...")
                    else:
                        app_logger.error(f"[导入备份] 错误: config.json 复制后不存在!")
                else:
                    app_logger.info(f"[导入备份] 备份中不包含 config.json，跳过")

                app_logger.info(f"备份导入成功，原数据已保存到: {restore_path}")
                return True

            finally:
                # 清理临时目录
                shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            app_logger.error(f"导入备份失败: {e}")
            return False

    def add_question(self, question: Question) -> bool:
        """添加新题目"""
        if not question.is_valid():
            return False
        self.questions.append(question)
        if question.source and question.source not in self.papers:
            self.papers.append(question.source)
        return self.save()

    def delete_question(self, index: int) -> bool:
        """删除指定索引的题目"""
        if 0 <= index < len(self.questions):
            del self.questions[index]
            self._update_papers()
            return self.save()
        return False

    def update_question(self, index: int, question: Question) -> bool:
        """更新指定索引的题目"""
        if 0 <= index < len(self.questions) and question.is_valid():
            self.questions[index] = question
            self._update_papers()
            return self.save()
        return False

    def get_questions_by_classification(self, classification_id: int) -> List[Question]:
        """按分类获取题目"""
        return [q for q in self.questions if q.classification == classification_id]

    def get_questions_by_source(self, source: str) -> List[Question]:
        """按来源获取题目"""
        return [q for q in self.questions if q.source == source]

    def filter_questions(
        self,
        classifications: Optional[List[int]] = None,
        source: Optional[str] = None,
        max_accuracy: Optional[float] = None
    ) -> List[Question]:
        """筛选题目"""
        result = self.questions.copy()
        
        if classifications:
            result = [q for q in result if q.classification in classifications]
        
        if source and source != "Any":
            result = [q for q in result if q.source == source]
        
        if max_accuracy is not None:
            result = [q for q in result if q.get_accuracy() <= max_accuracy]
        
        return result

    def _update_papers(self):
        """更新试卷列表"""
        self.papers = []
        for q in self.questions:
            if q.source and q.source not in self.papers:
                self.papers.append(q.source)

    def get_all_questions(self) -> List[Question]:
        """获取所有题目"""
        return self.questions.copy()

    def get_papers(self) -> List[str]:
        """获取所有试卷名称"""
        return self.papers.copy()

    def get_question_count(self) -> int:
        """获取题目总数"""
        return len(self.questions)
