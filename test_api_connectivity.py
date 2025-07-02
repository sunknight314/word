#!/usr/bin/env python3
"""
测试API连接性
"""

import asyncio
import httpx
import json


async def test_api_connectivity():
    """测试API连接性"""
    
    # API配置
    api_base_url = "https://api.siliconflow.cn/v1"
    api_key = "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn"
    model = "deepseek-ai/DeepSeek-V2.5"
    
    print("🔗 测试API连接性...")
    print(f"API URL: {api_base_url}")
    print(f"模型: {model}")
    print(f"API Key: {api_key[:20]}...{api_key[-10:]}")
    
    # 测试请求头
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    # 简单测试请求
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system", 
                "content": "你是一个测试助手，请返回JSON格式。"
            },
            {
                "role": "user", 
                "content": "请返回一个简单的JSON测试响应，包含status和message字段。"
            }
        ],
        "temperature": 0.1,
        "max_tokens": 100,
        "response_format": {"type": "json_object"}
    }
    
    try:
        print("\n📡 发送测试请求...")
        
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.post(
                f"{api_base_url}/chat/completions",
                headers=headers,
                json=payload
            )
            
            print(f"📊 响应状态码: {response.status_code}")
            
            if response.status_code == 200:
                result = response.json()
                print("✅ API连接成功!")
                
                # 提取响应内容
                if "choices" in result and len(result["choices"]) > 0:
                    content = result["choices"][0]["message"]["content"]
                    print(f"📝 响应内容: {content}")
                    
                    # 尝试解析JSON
                    try:
                        parsed = json.loads(content)
                        print("✅ JSON解析成功!")
                        print(f"📋 解析结果: {json.dumps(parsed, ensure_ascii=False, indent=2)}")
                    except json.JSONDecodeError as e:
                        print(f"❌ JSON解析失败: {e}")
                        
                # 检查使用情况
                if "usage" in result:
                    usage = result["usage"]
                    print(f"💰 Token使用: 输入{usage.get('prompt_tokens', 0)}, 输出{usage.get('completion_tokens', 0)}, 总计{usage.get('total_tokens', 0)}")
                    
            elif response.status_code == 401:
                print("❌ 认证失败 - API Key可能无效")
                print(f"响应: {response.text}")
                
            elif response.status_code == 429:
                print("❌ 请求过于频繁 - 触发限流")
                print(f"响应: {response.text}")
                
            elif response.status_code == 500:
                print("❌ 服务器内部错误")
                print(f"响应: {response.text}")
                
            else:
                print(f"❌ 未知错误 - 状态码: {response.status_code}")
                print(f"响应: {response.text}")
                
    except httpx.ConnectTimeout:
        print("❌ 连接超时 - 网络连接问题")
        
    except httpx.ReadTimeout:
        print("❌ 读取超时 - 服务器响应过慢")
        
    except httpx.RequestError as e:
        print(f"❌ 请求错误: {e}")
        
    except Exception as e:
        print(f"💥 未知异常: {e}")
        import traceback
        traceback.print_exc()


async def test_model_availability():
    """测试模型可用性"""
    
    api_base_url = "https://api.siliconflow.cn/v1"
    api_key = "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
    }
    
    print("\n🤖 测试模型列表...")
    
    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.get(
                f"{api_base_url}/models",
                headers=headers
            )
            
            if response.status_code == 200:
                models = response.json()
                print("✅ 模型列表获取成功!")
                
                # 查找DeepSeek模型
                if "data" in models:
                    deepseek_models = [m for m in models["data"] if "deepseek" in m.get("id", "").lower()]
                    if deepseek_models:
                        print("🧠 可用的DeepSeek模型:")
                        for model in deepseek_models:
                            print(f"  - {model.get('id', 'Unknown')}")
                    else:
                        print("⚠️ 未找到DeepSeek模型")
                        
            else:
                print(f"❌ 获取模型列表失败: {response.status_code}")
                print(f"响应: {response.text}")
                
    except Exception as e:
        print(f"💥 获取模型列表异常: {e}")


if __name__ == "__main__":
    print("🚀 开始测试API连接性...")
    
    # 测试基本连接
    asyncio.run(test_api_connectivity())
    
    # 测试模型可用性
    asyncio.run(test_model_availability())
    
    print("\n🏁 测试完成!")