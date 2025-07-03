"""
AI客户端基类 - 提供通用的AI API调用功能
"""

import json
import httpx
import re
import os
from typing import Dict, List, Any, Optional, Tuple
from abc import ABC, abstractmethod


class AIClientBase(ABC):
    """AI客户端基类"""
    
    def __init__(self, api_base_url: str = None, api_key: str = None, model: str = None, model_config: str = None):
        """
        初始化AI客户端
        
        Args:
            api_base_url: API基础URL
            api_key: API密钥
            model: 模型名称
            model_config: 模型配置名称 (如: 'deepseek', 'claude_sonnet', 'gpt4')
        """
        # 加载模型配置
        if model_config:
            config = self.load_model_config(model_config)
            self.api_base_url = api_base_url or config.get("api_base_url", "https://api.siliconflow.cn/v1")
            self.api_key = api_key or config.get("api_key", "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn")
            self.model = model or config.get("model", "deepseek-ai/DeepSeek-V2.5")
            self.model_config_name = model_config
        else:
            # 使用默认配置或直接传入的参数 (更新为DeepSeek-V3)
            self.api_base_url = api_base_url or "https://api.siliconflow.cn/v1"
            self.api_key = api_key or "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn"
            self.model = model or "deepseek-ai/DeepSeek-V3"
            self.model_config_name = "default"
        
        # HTTP客户端配置
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        self.timeout = 300.0  # 增加到5分钟
    
    def load_model_config(self, config_name: str) -> Dict[str, Any]:
        """
        加载模型配置
        
        Args:
            config_name: 配置名称
            
        Returns:
            模型配置字典
        """
        try:
            # 获取配置文件路径
            config_path = os.path.join(os.path.dirname(__file__), "..", "config", "model_configs.json")
            
            with open(config_path, 'r', encoding='utf-8') as f:
                configs = json.load(f)
            
            if config_name in configs["models"]:
                return configs["models"][config_name]
            else:
                print(f"⚠️ 未找到模型配置 '{config_name}'，使用默认配置")
                default_key = configs.get("default_model", "deepseek_v3")
                return configs["models"][default_key]
                
        except Exception as e:
            print(f"❌ 加载模型配置失败: {str(e)}")
            # 返回默认配置 (更新为DeepSeek-V3)
            return {
                "api_base_url": "https://api.siliconflow.cn/v1",
                "api_key": "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn",
                "model": "deepseek-ai/DeepSeek-V3"
            }
    
    @classmethod
    def get_available_models(cls) -> Dict[str, Any]:
        """
        获取所有可用的模型配置
        
        Returns:
            可用模型配置字典
        """
        try:
            config_path = os.path.join(os.path.dirname(__file__), "..", "config", "model_configs.json")
            
            with open(config_path, 'r', encoding='utf-8') as f:
                configs = json.load(f)
            
            return configs["models"]
            
        except Exception as e:
            print(f"❌ 获取模型配置失败: {str(e)}")
            return {}
    
    async def call_ai_api(self, system_prompt: str, user_prompt: str, max_tokens: int = 4000, temperature: float = 0.1) -> Dict[str, Any]:
        """
        调用AI API的通用方法
        
        Args:
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            max_tokens: 最大token数
            temperature: 温度参数
            
        Returns:
            AI响应结果
        """
        try:
            print("🔧 构建AI请求...")
            print(f"📋 System prompt长度: {len(system_prompt)} 字符")
            print(f"📋 User prompt长度: {len(user_prompt)} 字符")
            
            # 构建请求数据
            request_data = {
                "model": self.model,
                "messages": [
                    {
                        "role": "system",
                        "content": system_prompt
                    },
                    {
                        "role": "user",
                        "content": user_prompt
                    }
                ],
                "max_tokens": max_tokens,
                "temperature": temperature,
                "stream": False
            }
            
            print(f"🎯 使用模型: {self.model}")
            print(f"🔗 API地址: {self.api_base_url}")
            
            # 发送请求
            print("📡 发送AI请求...")
            async with httpx.AsyncClient(timeout=self.timeout) as client:
                response = await client.post(
                    f"{self.api_base_url}/chat/completions",
                    headers=self.headers,
                    json=request_data
                )
                print(f"📈 响应状态码: {response.status_code}")
                response.raise_for_status()
                
                result = response.json()
                print(f"📊 响应数据键: {list(result.keys())}")
                
                if "choices" in result and len(result["choices"]) > 0:
                    content = result["choices"][0]["message"]["content"]
                    print(f"💬 AI响应长度: {len(content)} 字符")
                    print(f"📝 AI响应前500字符: {content[:500]}...")
                    
                    return {
                        "success": True,
                        "content": content,
                        "model_info": {
                            "model": self.model,
                            "usage": result.get("usage", {})
                        }
                    }
                else:
                    print("❌ API响应格式错误 - 缺少choices字段")
                    return {
                        "success": False,
                        "error": "API响应格式错误",
                        "raw_response": result
                    }
                    
        except httpx.HTTPStatusError as e:
            print(f"❌ HTTP错误: {e.response.status_code}")
            print(f"📄 错误响应: {e.response.text}")
            return {
                "success": False,
                "error": f"HTTP错误: {e.response.status_code} - {e.response.text}",
                "raw_response": None
            }
        except httpx.TimeoutException:
            print("❌ 请求超时")
            return {
                "success": False,
                "error": "请求超时",
                "raw_response": None
            }
        except Exception as e:
            print(f"❌ 其他错误: {str(e)}")
            print(f"🔍 错误类型: {type(e).__name__}")
            return {
                "success": False,
                "error": f"AI调用失败: {str(e)}",
                "raw_response": None
            }
    
    def extract_json_from_content(self, content: str) -> str:
        """
        从AI响应内容中提取JSON部分
        
        Args:
            content: AI响应内容
            
        Returns:
            JSON字符串
        """
        content = content.strip()
        print(f"🔍 原始内容长度: {len(content)} 字符")
        
        # 方法1: 移除markdown代码块标记
        if '```json' in content or '```' in content:
            # 查找 ```json 或 ``` 开始和结束的位置
            lines = content.split('\n')
            json_lines = []
            in_json_block = False
            
            for line in lines:
                if line.strip().startswith('```json') or line.strip() == '```':
                    in_json_block = True
                    continue
                elif line.strip() == '```' and in_json_block:
                    in_json_block = False
                    break
                elif in_json_block:
                    json_lines.append(line)
            
            if json_lines:
                json_text = '\n'.join(json_lines)
                print(f"🔍 提取的JSON文本长度: {len(json_text)} 字符")
                try:
                    # 验证是否为有效JSON
                    json.loads(json_text)
                    return json_text
                except json.JSONDecodeError as e:
                    print(f"⚠️ JSON验证失败: {str(e)}")
        
        # 方法2: 寻找第一个{到最后一个}的完整JSON
        first_brace = content.find('{')
        if first_brace != -1:
            # 使用括号计数来找到匹配的结束}
            brace_count = 0
            json_end = -1
            
            for i in range(first_brace, len(content)):
                if content[i] == '{':
                    brace_count += 1
                elif content[i] == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        json_end = i + 1
                        break
            
            if json_end != -1:
                potential_json = content[first_brace:json_end]
                print(f"🔍 括号匹配提取的JSON长度: {len(potential_json)} 字符")
                try:
                    json.loads(potential_json)
                    return potential_json
                except json.JSONDecodeError as e:
                    print(f"⚠️ 括号匹配JSON验证失败: {str(e)}")
        
        # 方法3: 简单的首尾查找
        first_brace = content.find('{')
        last_brace = content.rfind('}')
        
        if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
            potential_json = content[first_brace:last_brace + 1]
            print(f"🔍 首尾查找提取的JSON长度: {len(potential_json)} 字符")
            try:
                json.loads(potential_json)
                return potential_json
            except json.JSONDecodeError as e:
                print(f"⚠️ 首尾查找JSON验证失败: {str(e)}")
        
        print("❌ 所有JSON提取方法都失败，返回原内容")
        return content
    
    @abstractmethod
    def get_prompts(self, input_data: Any) -> Tuple[str, str]:
        """
        获取系统和用户提示词（子类必须实现）
        
        Args:
            input_data: 输入数据
            
        Returns:
            (system_prompt, user_prompt) 元组
        """
        pass
    
    @abstractmethod
    async def process_ai_response(self, ai_response: str) -> Dict[str, Any]:
        """
        处理AI响应（子类必须实现）
        
        Args:
            ai_response: AI原始响应
            
        Returns:
            处理后的结果
        """
        pass
    
    async def analyze(self, input_data: Any) -> Dict[str, Any]:
        """
        分析方法 - 子类调用的通用流程
        
        Args:
            input_data: 输入数据
            
        Returns:
            分析结果
        """
        try:
            # 1. 获取提示词
            system_prompt, user_prompt = self.get_prompts(input_data)
            
            # 2. 调用AI API
            ai_result = await self.call_ai_api(system_prompt, user_prompt)
            
            if ai_result["success"]:
                # 3. 处理AI响应
                processed_result = await self.process_ai_response(ai_result["content"])
                processed_result["model_info"] = ai_result["model_info"]
                return processed_result
            else:
                return {
                    "success": False,
                    "error": ai_result["error"],
                    "raw_response": ai_result.get("raw_response")
                }
                
        except Exception as e:
            print(f"❌ 分析失败: {str(e)}")
            return {
                "success": False,
                "error": f"分析失败: {str(e)}",
                "raw_response": None
            }