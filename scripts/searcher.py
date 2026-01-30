#!/usr/bin/env python3
"""
Question Bank Search Module - 按需加载
支持索引搜索 + 按需加载docx文件
"""

import os
import json
import re
from docx import Document
from typing import List, Dict, Optional, Tuple

RESOURCE_PATH = "/Users/xielk/webdata/english/lesson/resource"
INDEX_FILE = os.path.join(RESOURCE_PATH, "index.json")

class QuestionBankSearcher:
    """题库搜索器 - 索引+按需加载"""
    
    def __init__(self):
        self.index = None
        self.load_index()
    
    def load_index(self):
        """加载索引文件"""
        if not os.path.exists(INDEX_FILE):
            raise FileNotFoundError(
                f"索引文件不存在: {INDEX_FILE}\n"
                "请先运行: python scripts/indexer.py"
            )
        
        with open(INDEX_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            self.index = data.get("files", [])
            self.metadata = data.get("metadata", {})
        
        print(f"✅ 索引加载成功: {self.metadata.get('total_files', 0)} 个文件")
    
    def search(
        self,
        keyword: Optional[str] = None,
        year: Optional[str] = None,
        district: Optional[str] = None,
        exam_type: Optional[str] = None,
        question_type: Optional[str] = None,
        limit: int = 10
    ) -> List[Dict]:
        """
        搜索题库
        
        Args:
            keyword: 关键词（搜索文件名和预览内容）
            year: 年份（如: 2025, 2024）
            district: 区域（如: 徐汇, 浦东）
            exam_type: 考试类型（如: 一模, 二模）
            question_type: 题型（如: 语法, 阅读）
            limit: 返回结果数量限制
        
        Returns:
            匹配的索引项列表
        """
        results = []
        
        for item in self.index:
            # 检查各个条件
            match = True
            
            if year and item.get("year") != year:
                match = False
            
            if district and item.get("district") != district:
                match = False
            
            if exam_type and item.get("exam_type") != exam_type:
                match = False
            
            if question_type and item.get("question_type") != question_type:
                match = False
            
            if keyword:
                keyword_lower = keyword.lower()
                # 搜索文件名和预览内容
                if (keyword_lower not in item.get("filename", "").lower() and 
                    keyword_lower not in item.get("preview", "").lower()):
                    match = False
            
            if match:
                results.append(item)
        
        # 按优先级排序：年份降序、文件大小升序（优先小文件）
        results.sort(key=lambda x: (
            int(x.get("year", "0")) if x.get("year", "0").isdigit() else 0,
            x.get("size_kb", 999999)
        ), reverse=True)
        
        return results[:limit]
    
    def load_document(self, file_path: str) -> Document:
        """
        按需加载docx文档
        
        Args:
            file_path: docx文件完整路径
        
        Returns:
            Document对象
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        print(f"📄 正在加载: {os.path.basename(file_path)} ({os.path.getsize(file_path)//1024}KB)")
        return Document(file_path)
    
    def extract_questions(self, doc: Document, keyword: str = None) -> List[str]:
        """
        从文档中提取题目
        
        Args:
            doc: Document对象
            keyword: 可选的关键词过滤
        
        Returns:
            题目列表
        """
        questions = []
        current_question = []
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            
            # 检测是否是题目开始（通常包含数字、题号、问号等）
            is_question_start = bool(
                re.match(r'^\d+[\.．\s]', text) or  # 数字开头
                re.match(r'^[【\[]', text) or        # 【例题】或[例]
                '?' in text or                        # 包含问号
                '（' in text and '）' in text or     # 包含括号选项
                text.startswith(('A.', 'B.', 'C.', 'D.'))  # 选项
            )
            
            if is_question_start and current_question:
                # 保存上一题
                full_question = '\n'.join(current_question)
                if not keyword or keyword.lower() in full_question.lower():
                    questions.append(full_question)
                current_question = []
            
            current_question.append(text)
        
        # 处理最后一题
        if current_question:
            full_question = '\n'.join(current_question)
            if not keyword or keyword.lower() in full_question.lower():
                questions.append(full_question)
        
        return questions
    
    def smart_search(
        self,
        topic: str,
        district: Optional[str] = None,
        year: Optional[str] = None,
        load_docs: bool = True,
        max_docs: int = 3,
        max_questions_per_doc: int = 5
    ) -> Tuple[List[Dict], List[str]]:
        """
        智能搜索 - 搜索索引并可选加载文档
        
        Args:
            topic: 主题关键词（如: 非谓语, 定语从句）
            district: 优先区域（如学生所在区）
            year: 优先年份
            load_docs: 是否加载文档内容
            max_docs: 最多加载的文档数量
            max_questions_per_doc: 每个文档提取的最大题目数
        
        Returns:
            (索引结果列表, 题目内容列表)
        """
        print(f"\n🔍 搜索: topic='{topic}', district='{district}', year='{year}'")
        
        # 1. 搜索索引
        results = self.search(
            keyword=topic,
            district=district,
            year=year,
            limit=max_docs * 2  # 多搜一些以便筛选
        )
        
        if not results:
            print("⚠️ 未找到匹配结果")
            return [], []
        
        print(f"   索引匹配: {len(results)} 个文件")
        
        if not load_docs:
            return results, []
        
        # 2. 按需加载文档并提取内容
        all_questions = []
        loaded_count = 0
        
        for item in results:
            if loaded_count >= max_docs:
                break
            
            try:
                # 加载文档
                doc = self.load_document(item["file"])
                loaded_count += 1
                
                # 提取题目
                questions = self.extract_questions(doc, keyword=topic)
                
                # 添加来源标注
                for q in questions[:max_questions_per_doc]:
                    source = f"({item['year']} {item['district']}{item['exam_type']})"
                    all_questions.append({
                        "content": q,
                        "source": source,
                        "file": item["filename"]
                    })
                
            except Exception as e:
                print(f"   加载失败: {item['filename']} - {e}")
                continue
        
        print(f"   已加载: {loaded_count} 个文件, 提取 {len(all_questions)} 道题目")
        
        return results[:max_docs], all_questions

# 便捷函数
def search_question_bank(
    topic: str,
    district: Optional[str] = None,
    year: Optional[str] = None,
    load_content: bool = True
) -> Tuple[List[Dict], List[str]]:
    """
    快速搜索题库
    
    使用示例:
        results, questions = search_question_bank("非谓语", "嘉定", "2025")
    """
    searcher = QuestionBankSearcher()
    return searcher.smart_search(
        topic=topic,
        district=district,
        year=year,
        load_docs=load_content
    )

if __name__ == "__main__":
    # 测试代码
    print("=" * 50)
    print("题库搜索测试")
    print("=" * 50)
    
    # 初始化搜索器
    searcher = QuestionBankSearcher()
    
    # 测试1: 搜索索引
    print("\n测试1: 搜索索引 (非谓语)")
    results = searcher.search(keyword="非谓语", limit=5)
    for r in results:
        print(f"   {r['filename']} - {r['district']} {r['year']} ({r['size_kb']}KB)")
    
    # 测试2: 智能搜索并加载
    print("\n测试2: 智能搜索并加载文档")
    idx_results, questions = searcher.smart_search(
        topic="阅读",
        district="浦东",
        year="2024",
        max_docs=2,
        max_questions_per_doc=2
    )
    
    if questions:
        print("\n提取的题目示例:")
        for i, q in enumerate(questions[:2], 1):
            print(f"\n题目 {i} {q['source']}:")
            print(q['content'][:200] + "...")
