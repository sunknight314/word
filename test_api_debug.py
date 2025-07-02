#!/usr/bin/env python3
"""
调试API调用问题
"""

import sys
import json
import asyncio
sys.path.append('/opt/word/backend')

from app.services.ai_analyzer import AIAnalyzer


async def debug_api_call():
    """调试API调用"""
    
    # 创建AI分析器
    analyzer = AIAnalyzer()
    
    # 准备测试数据
    test_paragraphs = [
        {"paragraph_number": 1, "preview_text": "Word文档格式优化项目测试文档"},
        {"paragraph_number": 2, "preview_text": "第一章 项目概述"},
        {"paragraph_number": 3, "preview_text": "本项目是一个基于人工智能的Word文档格"},
        {"paragraph_number": 4, "preview_text": "1.1 项目背景"},
        {"paragraph_number": 5, "preview_text": "在现代办公环境中，文档格式的统一性对于提"}
    ]
    
    print("🔧 调试API调用...")
    print(f"API Base URL: {analyzer.api_base_url}")
    print(f"Model: {analyzer.model}")
    print(f"API Key前缀: {analyzer.api_key[:20]}...")
    
    try:
        print("\n📤 发送API请求...")
        result = await analyzer.analyze_paragraphs(test_paragraphs)
        
        print(f"\n📥 API响应:")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        
        if result.get("success"):
            print("\n✅ API调用成功!")
            analysis = result["result"]
            if "analysis_result" in analysis:
                print(f"分析了 {len(analysis['analysis_result'])} 个段落")
                for item in analysis["analysis_result"][:3]:
                    print(f"- 段落{item['paragraph_number']}: {item['type']} (置信度: {item.get('confidence', 'N/A')})")
        else:
            print(f"\n❌ API调用失败: {result.get('error', '未知错误')}")
            
    except Exception as e:
        print(f"\n💥 调用异常: {e}")
        import traceback
        traceback.print_exc()


async def test_simple_request():
    """测试简单的API请求"""
    import httpx
    
    api_key = "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn"
    base_url = "https://api.siliconflow.cn/v1"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "model": "deepseek-ai/DeepSeek-V2.5",
        "messages": [
            {
                "role": "system", 
                "content": "你是一个专业的文档格式分析专家。请严格按照JSON格式输出。"
            },
            {
                "role": "user", 
                "content": "请分析这个段落的类型：'第一章 项目概述'。返回JSON格式：{\"type\": \"heading1\", \"confidence\": 0.9, \"reason\": \"一级标题\"}"
            }
        ],
        "temperature": 0.1,
        "max_tokens": 500,
        "response_format": {"type": "json_object"}
    }
    
    print("\n🧪 测试简单API请求...")
    
    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.post(
                f"{base_url}/chat/completions",
                headers=headers,
                json=payload
            )
            
            print(f"状态码: {response.status_code}")
            
            if response.status_code == 200:
                result = response.json()
                content = result["choices"][0]["message"]["content"]
                print(f"✅ 成功响应:")
                print(content)
                
                # 尝试解析JSON
                try:
                    parsed = json.loads(content)
                    print(f"✅ JSON解析成功:")
                    print(json.dumps(parsed, ensure_ascii=False, indent=2))
                except json.JSONDecodeError as e:
                    print(f"❌ JSON解析失败: {e}")
                    
            else:
                print(f"❌ 请求失败:")
                print(f"状态码: {response.status_code}")
                print(f"响应: {response.text}")
                
    except Exception as e:
        print(f"💥 请求异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("🔍 开始调试API调用问题...")
    
    # 先测试简单请求
    asyncio.run(test_simple_request())
    
    print("\n" + "="*50)
    
    # 再测试完整流程
    asyncio.run(debug_api_call())