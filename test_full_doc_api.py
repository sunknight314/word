#!/usr/bin/env python3
"""
测试完整文档的API调用
"""

import sys
import json
import asyncio
sys.path.append('/opt/word/backend')

from app.services.document_parser import DocumentParser
from app.services.ai_analyzer import AIAnalyzer


async def test_full_document():
    """测试完整文档的API调用"""
    
    # 创建解析器
    parser = DocumentParser()
    
    # 先提取文档信息
    test_file_path = '/opt/word/test_files/test_document.docx'
    basic_result = parser.extract_paragraphs_info(test_file_path)
    
    print(f"📄 文档信息: 共 {basic_result['document_info']['total_paragraphs']} 个段落")
    print(basic_result)
    
    # 准备AI分析数据
    paragraphs_for_ai = []
    for para in basic_result["paragraphs"]:
        paragraphs_for_ai.append({
            "paragraph_number": para["paragraph_number"],
            "preview_text": para["preview_text"]
        })
    
    # 创建AI分析器
    analyzer = AIAnalyzer()
    
    print("🤖 调用AI分析器...")
    
    try:
        # 直接调用AI分析器
        ai_result = await analyzer.analyze_paragraphs(paragraphs_for_ai)
        
        print(f"📥 AI分析器返回:")
        print(f"- 成功: {ai_result.get('success', False)}")
        
        if ai_result.get("success"):
            result = ai_result["result"]
            print(f"- 分析结果包含 {len(result.get('analysis_result', []))} 个段落")
            
            # 保存结果
            output_file = '/opt/word/test_files/full_api_result.json'
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(ai_result, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 结果已保存到: {output_file}")
            
            # 显示前5个分析结果
            print(f"\n📋 前5个段落分析结果:")
            for item in result["analysis_result"][:5]:
                print(f"{item['paragraph_number']}. {item['preview_text']}")
                print(f"   类型: {item['type']} | 置信度: {item['confidence']}")
                print(f"   理由: {item['reason']}")
                print()
            
        else:
            print(f"❌ AI分析失败:")
            print(f"错误: {ai_result.get('error', '未知错误')}")
            if 'total_paragraphs' in ai_result:
                print(f"处理段落数: {ai_result['total_paragraphs']}")
            
    except Exception as e:
        print(f"💥 调用异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("🧪 测试完整文档的真实API调用...")
    asyncio.run(test_full_document())