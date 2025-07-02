#!/usr/bin/env python3
"""
测试AI分析功能
"""

import sys
import json
import asyncio
sys.path.append('/opt/word/backend')

from app.services.document_parser import DocumentParser


async def test_ai_analysis():
    """测试AI分析功能"""
    test_file_path = '/opt/word/test_files/test_document.docx'
    
    # 创建解析器
    parser = DocumentParser()
    
    try:
        print("正在进行AI分析（使用模拟模式）...")
        
        # 使用模拟分析
        result = await parser.analyze_document_with_ai(test_file_path, use_mock=True)
        
        print("\n=== AI分析结果 ===")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        
        # 保存结果
        output_file = '/opt/word/test_files/ai_analysis_result.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n分析结果已保存到: {output_file}")
        
        # 显示分析摘要
        summary = result["analysis_summary"]
        print(f"\n=== 分析摘要 ===")
        print(f"总段落数: {summary['total_paragraphs']}")
        print(f"平均置信度: {summary['average_confidence']}")
        print(f"文档结构: {summary['structure_detected']}")
        
        print(f"\n=== 段落类型分布 ===")
        for para_type, count in summary["type_distribution"].items():
            print(f"{para_type}: {count}个")
        
        # 显示前几个段落的分析结果
        print(f"\n=== 前10个段落分析结果 ===")
        ai_results = result["ai_analysis"]["analysis_result"]
        for i, item in enumerate(ai_results[:10]):
            print(f"{item['paragraph_number']}. {item['preview_text']}")
            print(f"   类型: {item['type']} (置信度: {item['confidence']})")
            print(f"   理由: {item['reason']}")
            print()
        
    except Exception as e:
        print(f"测试失败: {e}")


def test_prompt_generation():
    """测试prompt生成"""
    from app.services.ai_prompts import get_analysis_prompt
    
    # 准备测试数据
    test_paragraphs = [
        {"paragraph_number": 1, "preview_text": "Word文档格式优化项目测试文档"},
        {"paragraph_number": 2, "preview_text": "第一章 项目概述"},
        {"paragraph_number": 3, "preview_text": "本项目是一个基于人工智能的Word文档格"},
        {"paragraph_number": 4, "preview_text": "1.1 项目背景"},
        {"paragraph_number": 5, "preview_text": "在现代办公环境中，文档格式的统一性对于提"}
    ]
    
    print("=== 生成的Prompt ===")
    prompt = get_analysis_prompt(test_paragraphs)
    print(prompt)
    
    # 保存prompt到文件
    with open('/opt/word/test_files/generated_prompt.txt', 'w', encoding='utf-8') as f:
        f.write(prompt)
    
    print(f"\nPrompt已保存到: /opt/word/test_files/generated_prompt.txt")


if __name__ == "__main__":
    print("=== 测试Prompt生成 ===")
    test_prompt_generation()
    
    print("\n" + "="*60)
    print("=== 测试AI分析功能 ===")
    asyncio.run(test_ai_analysis())