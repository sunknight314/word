#!/usr/bin/env python3
"""
测试document_parser.extract_paragraphs_info方法对论文.docx的提取功能
"""

import sys
import os
import json
from datetime import datetime

# 添加backend路径到sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), '../backend'))

from app.services.document_parser import DocumentParser


def test_paragraph_extraction():
    """测试段落提取功能"""
    
    print("📄 大论文段落提取测试")
    print("=" * 50)
    
    # 文件路径
    file_path = os.path.join(os.path.dirname(__file__), "论文.docx")
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"❌ 错误：文件不存在 - {file_path}")
        return
    
    print(f"📁 测试文件：{file_path}")
    print(f"📊 文件大小：{os.path.getsize(file_path) / 1024:.2f} KB")
    
    # 创建文档解析器
    document_parser = DocumentParser()
    
    # 测试不同的preview_length设置
    preview_lengths = [20, 50, 100]
    
    for preview_length in preview_lengths:
        print(f"\n🔍 测试预览长度：{preview_length} 字符")
        print("-" * 40)
        
        
        # 开始提取
        start_time = datetime.now()
        result = document_parser.extract_paragraphs_info(file_path, preview_length)
        end_time = datetime.now()
        
        # 计算耗时
        duration = (end_time - start_time).total_seconds()
        
        # 输出基本信息
        print(f"✅ 提取成功！耗时：{duration:.2f}秒")
        print(f"📋 文档信息：")
        doc_info = result.get("document_info", {})
        print(f"   - 总段落数：{doc_info.get('total_paragraphs', 0)}")
        print(f"   - 文档页数：{doc_info.get('total_pages', 0)}")
        print(f"   - 文档长度：{doc_info.get('total_length', 0)} 字符")
        
        # 显示段落样本
        paragraphs = result.get("paragraphs", [])
        for i, para in enumerate(paragraphs):
            para_num = para.get("paragraph_number", i+1)
            preview = para.get("preview_text", "")
            full_length = para.get("full_length", 0)
            
            # 格式化输出
            print(f"{para_num:3d}. [{full_length:3d}字] {preview}")


if __name__ == "__main__":
    print("🚀 开始论文文档提取测试")
    print(f"⏰ 测试时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 执行测试
    test_paragraph_extraction()
    
    print(f"\n🎉 测试完成！")
    print("=" * 50)