"""
大模型分析服务
"""

import json
import httpx
from typing import Dict, List, Any, Optional
from app.services.ai_prompts import get_analysis_prompt


class AIAnalyzer:
    """大模型分析服务类"""
    
    def __init__(self, api_base_url: str = None, api_key: str = None, model: str = None):
        """
        初始化AI分析器
        
        Args:
            api_base_url: API基础URL
            api_key: API密钥
            model: 模型名称
        """
        # 使用硅基流动的配置
        self.api_base_url = api_base_url or "https://api.siliconflow.cn/v1"
        self.api_key = api_key or "sk-ennjwzgywmisvlgqcumcojemtajhzmcowpmoothwmlklrzcn"
        self.model = model or "deepseek-ai/DeepSeek-V2.5"
        
        # HTTP客户端配置
        self.timeout = 60.0  # 增加超时时间
        self.max_retries = 3
    
    async def analyze_paragraphs(self, paragraphs_data: List[Dict]) -> Dict[str, Any]:
        """
        分析段落类型（支持分批处理大文档）
        
        Args:
            paragraphs_data: 段落数据列表，包含paragraph_number和preview_text
            
        Returns:
            分析结果字典
        """
        try:
            # 如果段落数超过10个，进行分批处理
            if len(paragraphs_data) > 10:
                return await self._analyze_paragraphs_in_batches(paragraphs_data)
            else:
                return await self._analyze_single_batch(paragraphs_data)
            
        except Exception as e:
            import traceback
            error_detail = f"{str(e)}\n{traceback.format_exc()}"
            return {
                "success": False,
                "error": error_detail,
                "total_paragraphs": len(paragraphs_data)
            }
    
    async def _analyze_single_batch(self, paragraphs_data: List[Dict]) -> Dict[str, Any]:
        """
        分析单批段落
        """
        # 生成prompt
        prompt = get_analysis_prompt(paragraphs_data)
        
        # 调用大模型API
        response = await self._call_ai_api(prompt)
        
        # 解析响应
        analysis_result = self._parse_ai_response(response)
        
        return {
            "success": True,
            "result": analysis_result,
            "total_paragraphs": len(paragraphs_data)
        }
    
    async def _analyze_paragraphs_in_batches(self, paragraphs_data: List[Dict]) -> Dict[str, Any]:
        """
        分批处理大文档
        """
        batch_size = 8  # 每批处理8个段落
        all_results = []
        
        print(f"📦 文档较大，分批处理: {len(paragraphs_data)}个段落，每批{batch_size}个")
        
        for i in range(0, len(paragraphs_data), batch_size):
            batch = paragraphs_data[i:i + batch_size]
            batch_num = i // batch_size + 1
            total_batches = (len(paragraphs_data) + batch_size - 1) // batch_size
            
            print(f"  📄 处理第{batch_num}批 (共{total_batches}批): 段落{batch[0]['paragraph_number']}-{batch[-1]['paragraph_number']}")
            
            try:
                batch_result = await self._analyze_single_batch(batch)
                
                if batch_result["success"]:
                    batch_analysis = batch_result["result"]["analysis_result"]
                    all_results.extend(batch_analysis)
                    print(f"  ✅ 第{batch_num}批完成")
                else:
                    print(f"  ❌ 第{batch_num}批失败: {batch_result['error'][:100]}...")
                    # 如果某批失败，可以继续处理其他批次
                    
            except Exception as e:
                print(f"  💥 第{batch_num}批异常: {str(e)[:100]}...")
                continue
        
        return {
            "success": True,
            "result": {
                "analysis_result": all_results
            },
            "total_paragraphs": len(paragraphs_data),
            "processed_paragraphs": len(all_results)
        }
    
    async def _call_ai_api(self, prompt: str) -> str:
        """
        调用大模型API
        
        Args:
            prompt: 分析prompt
            
        Returns:
            API响应文本
        """
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system", 
                    "content": "你是一个专业的文档格式分析专家，专门负责分析Word文档段落类型。请严格按照JSON格式输出分析结果。"
                },
                {
                    "role": "user", 
                    "content": prompt
                }
            ],
            "temperature": 0.1,  # 低温度确保稳定输出
            "max_tokens": 3000,
            "response_format": {"type": "json_object"}  # 启用JSON模式
        }
        
        async with httpx.AsyncClient(timeout=self.timeout) as client:
            response = await client.post(
                f"{self.api_base_url}/chat/completions",
                headers=headers,
                json=payload
            )
            
            if response.status_code != 200:
                raise Exception(f"API调用失败: {response.status_code} - {response.text}")
            
            result = response.json()
            return result["choices"][0]["message"]["content"]
    
    def _parse_ai_response(self, response_text: str) -> Dict[str, Any]:
        """
        解析大模型响应
        
        Args:
            response_text: API响应文本
            
        Returns:
            解析后的结果字典
        """
        try:
            # 尝试直接解析JSON
            if response_text.strip().startswith("{"):
                return json.loads(response_text)
            
            # 尝试从markdown代码块中提取JSON
            if "```json" in response_text:
                start = response_text.find("```json") + 7
                end = response_text.find("```", start)
                json_text = response_text[start:end].strip()
                return json.loads(json_text)
            
            # 尝试从```代码块中提取JSON
            if "```" in response_text:
                start = response_text.find("```") + 3
                end = response_text.find("```", start)
                json_text = response_text[start:end].strip()
                return json.loads(json_text)
            
            # 如果都失败，返回原始文本
            return {
                "error": "无法解析响应",
                "raw_response": response_text
            }
            
        except json.JSONDecodeError as e:
            return {
                "error": f"JSON解析失败: {str(e)}",
                "raw_response": response_text
            }
    
    def create_mock_analysis(self, paragraphs_data: List[Dict]) -> Dict[str, Any]:
        """
        创建模拟分析结果（用于测试）
        
        Args:
            paragraphs_data: 段落数据列表
            
        Returns:
            模拟的分析结果
        """
        analysis_result = []
        
        for para in paragraphs_data:
            preview = para["preview_text"]
            para_num = para["paragraph_number"]
            
            # 简单的规则判断（模拟AI分析）
            if "文档" in preview and para_num == 1:
                para_type = "title"
                confidence = 0.95
                reason = "文档标题，位于第一段"
            elif preview.startswith("第") and "章" in preview:
                para_type = "heading1"
                confidence = 0.92
                reason = "使用'第X章'格式的一级标题"
            elif preview.count(".") == 1 and preview[0].isdigit():
                para_type = "heading2"
                confidence = 0.88
                reason = "使用'X.X'格式的二级标题"
            elif preview.count(".") == 2 and preview[0].isdigit():
                para_type = "heading3"
                confidence = 0.85
                reason = "使用'X.X.X'格式的三级标题"
            elif len(preview) <= 15 and not preview.endswith("，"):
                para_type = "heading4"
                confidence = 0.75
                reason = "短文本，可能是低级标题"
            else:
                para_type = "paragraph"
                confidence = 0.80
                reason = "正文段落特征"
            
            analysis_result.append({
                "paragraph_number": para_num,
                "preview_text": preview,
                "type": para_type,
                "confidence": confidence,
                "reason": reason
            })
        
        return {
            "analysis_result": analysis_result
        }