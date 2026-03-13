"""
重置题目统计脚本
将 questions.json 中所有题目的 total 和 correct 都设置为 0
"""

import json
import os

# 数据文件路径
QUESTIONS_FILE = os.path.join(os.path.dirname(__file__), "src", "data", "questions.json")

def reset_statistics():
    """重置所有题目的统计信息"""
    # 读取文件
    with open(QUESTIONS_FILE, "r", encoding="utf-8") as f:
        questions = json.load(f)
    
    # 重置每道题的统计
    reset_count = 0
    for question in questions:
        question["total"] = 0
        question["correct"] = 0
        reset_count += 1
    
    # 保存文件
    with open(QUESTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(questions, f, ensure_ascii=False, indent=4)
    
    print(f"成功重置 {reset_count} 道题目的统计信息！")
    print(f"所有题目的 total 和 correct 已设置为 0")

if __name__ == "__main__":
    reset_statistics()
