#!/usr/bin/env python3
"""
测试真实API调用功能
"""

import sys
import json
import asyncio
sys.path.append('/opt/word/backend')

from app.services.document_parser import DocumentParser


async def test_real_api_call():
    """测试真实的API调用"""
    test_file_path = '/opt/word/test_files/test_document.docx'
    
    # 创建解析器
    parser = DocumentParser()
    
    try:
        print("=== 步骤1: 提取文档段落信息 ===")
        
        # 先进行基础解析，提取段落信息
        basic_result = parser.extract_paragraphs_info(test_file_path)
        
        print(f"文档基本信息:")
        print(f"- 总段落数: {basic_result['document_info']['total_paragraphs']}")
        print(f"- 文件路径: {basic_result['document_info']['file_path']}")
        
        # 显示前5个段落
        print(f"\n前5个段落预览:")
        for i, para in enumerate(basic_result['paragraphs'][:5]):
            print(f"{para['paragraph_number']}. {para['preview_text']}")
        
        print("\n" + "="*60)
        print("=== 步骤2: 调用真实API进行分析 ===")
        
        # 使用真实API调用进行分析
        print("正在调用硅基流动API...")
        print("模型: deepseek-ai/DeepSeek-V2.5")
        print("API: https://api.siliconflow.cn/v1")
        
        ai_result = await parser.analyze_document_with_ai(test_file_path, use_mock=False)
        
        print("\n✅ API调用成功!")
        
        # 保存完整结果
        output_file = '/opt/word/test_files/real_api_result.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(ai_result, f, ensure_ascii=False, indent=2)
        
        print(f"完整结果已保存到: {output_file}")
        
        print("\n" + "="*60)
        print("=== 步骤3: 分析结果对比 ===")
        
        # 显示分析摘要
        summary = ai_result["analysis_summary"]
        print(f"\n📊 分析摘要:")
        print(f"- 总段落数: {summary['total_paragraphs']}")
        print(f"- 平均置信度: {summary['average_confidence']}")
        print(f"- 文档结构: {summary['structure_detected']}")
        
        print(f"\n📋 段落类型分布:")
        for para_type, count in summary["type_distribution"].items():
            type_names = {
                "title": "文档标题",
                "heading1": "一级标题", 
                "heading2": "二级标题",
                "heading3": "三级标题",
                "heading4": "四级标题",
                "paragraph": "正文段落",
                "list": "列表项",
                "quote": "引用",
                "other": "其他"
            }
            print(f"- {type_names.get(para_type, para_type)}: {count}个")
        
        print(f"\n🔍 前10个段落的AI分析结果:")
        ai_results = ai_result["ai_analysis"]["analysis_result"]
        for i, item in enumerate(ai_results[:10]):
            confidence_emoji = "🟢" if item['confidence'] > 0.9 else "🟡" if item['confidence'] > 0.7 else "🔴"
            print(f"\n{item['paragraph_number']}. {item['preview_text']}")
            print(f"   {confidence_emoji} 类型: {item['type']} (置信度: {item['confidence']})")
            print(f"   💭 理由: {item['reason']}")
        
        # 对比模拟结果和真实结果
        print(f"\n" + "="*60)
        print("=== 步骤4: 模拟vs真实结果对比 ===")
        
        # 获取模拟结果
        mock_result = await parser.analyze_document_with_ai(test_file_path, use_mock=True)
        
        print("📈 结果差异分析:")
        real_types = [item['type'] for item in ai_result["ai_analysis"]["analysis_result"]]
        mock_types = [item['type'] for item in mock_result["ai_analysis"]["analysis_result"]]
        
        differences = []
        for i, (real_type, mock_type) in enumerate(zip(real_types, mock_types)):
            if real_type != mock_type:
                para = ai_result["ai_analysis"]["analysis_result"][i]
                differences.append({
                    "paragraph": para['paragraph_number'],
                    "text": para['preview_text'],
                    "real": real_type,
                    "mock": mock_type
                })
        
        if differences:
            print(f"发现 {len(differences)} 处差异:")
            for diff in differences[:5]:  # 只显示前5个差异
                print(f"- 段落{diff['paragraph']}: {diff['text']}")
                print(f"  真实API: {diff['real']} vs 模拟: {diff['mock']}")
        else:
            print("✅ 真实API结果与模拟结果完全一致!")
        
        print(f"\n🎯 测试完成! 真实API调用成功。")
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        print(f"错误类型: {type(e).__name__}")
        import traceback
        print(f"详细错误信息:")
        traceback.print_exc()


if __name__ == "__main__":
    print("🚀 开始测试真实API调用...")
    print("📄 测试文档: /opt/word/test_files/test_document.docx")
    print("🤖 API服务: 硅基流动 (SiliconCloud)")
    print("🧠 模型: deepseek-ai/DeepSeek-V2.5")
    print()
    
    asyncio.run(test_real_api_call())