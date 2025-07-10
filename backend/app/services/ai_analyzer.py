"""
AI段落分析服务 - 重构后的统一架构
"""

import os
import json
from datetime import datetime
from typing import Dict, List, Any, Tuple
from .ai_client_base import AIClientBase
from .document_processor import DocumentProcessor


class AIAnalyzer(AIClientBase):
    """AI段落分析服务类"""
    
    def __init__(self, api_base_url: str = None, api_key: str = None, model: str = None, model_config: str = None):
        """
        初始化AI分析器
        
        Args:
            api_base_url: API基础URL
            api_key: API密钥
            model: 模型名称
            model_config: 模型配置名称
        """
        # 初始化基类
        super().__init__(api_base_url, api_key, model, model_config)
        
        # 初始化文档处理器
        self.document_processor = DocumentProcessor()
        
        # 结果保存配置
        self.save_results = True
        self.results_dir = "ai_analysis_results"
        self._ensure_results_dir()
    
    def get_prompts(self, input_data: List[Dict]) -> Tuple[str, str]:
        """
        获取段落分析的系统和用户提示词
        
        Args:
            input_data: 段落数据列表
            
        Returns:
            (system_prompt, user_prompt) 元组
        """
        from .ai_prompts import get_paragraph_analysis_prompt
        return get_paragraph_analysis_prompt(input_data)
    
    async def process_ai_response(self, ai_response: str) -> Dict[str, Any]:
        """
        处理AI响应，解析段落分析结果
        
        Args:
            ai_response: AI原始响应
            
        Returns:
            处理后的分析结果
        """
        try:
            print("🔍 开始提取段落分析JSON...")
            json_content = self.extract_json_from_content(ai_response)
            print(f"📋 提取的JSON长度: {len(json_content)} 字符")
            
            analysis_result = json.loads(json_content)
            print("✅ 段落分析JSON解析成功")
            
            return {
                "success": True,
                "analysis_result": analysis_result.get("analysis_result", []),
                "raw_response": ai_response
            }
            
        except json.JSONDecodeError as e:
            print(f"❌ 段落分析JSON解析失败: {str(e)}")
            if 'json_content' in locals():
                print(f"🔍 尝试解析的内容: {json_content}")
            else:
                print("🔍 json_content变量未定义")
            return {
                "success": False,
                "error": f"JSON解析失败: {str(e)}",
                "raw_response": ai_response
            }
    
    async def analyze_paragraphs(self, paragraphs_data: List[Dict]) -> Dict[str, Any]:
        """
        分析段落类型（支持分批处理）
        
        Args:
            paragraphs_data: 段落数据列表，包含paragraph_number和preview_text
            
        Returns:
            分析结果字典
        """
        try:
            print(f"📄 开始分析 {len(paragraphs_data)} 个段落...")
            
            # 判断是否需要分批处理
            if len(paragraphs_data) > 200:
                print(f"📦 段落数量较多，将进行分批处理...")
                return await self._analyze_paragraphs_in_batches(paragraphs_data)
            
            # 直接使用AI分析，一次性处理
            result = await self.analyze(paragraphs_data)
            
            # 保存分析结果
            if self.save_results and result.get("success"):
                self._save_analysis_result(result, paragraphs_data)
            
            return result
            
        except Exception as e:
            import traceback
            error_detail = f"{str(e)}\n{traceback.format_exc()}"
            error_result = {
                "success": False,
                "error": error_detail,
                "total_paragraphs": len(paragraphs_data)
            }
            
            # 即使失败也保存错误信息
            if self.save_results:
                self._save_analysis_result(error_result, paragraphs_data, is_error=True)
            
            return error_result
    
    async def _analyze_paragraphs_in_batches(self, paragraphs_data: List[Dict]) -> Dict[str, Any]:
        """
        分批处理段落，带重叠上下文
        
        Args:
            paragraphs_data: 段落数据列表
            
        Returns:
            合并后的分析结果
        """
        batch_size = 100  # 每批大小
        overlap_size = 10  # 重叠段落数
        
        batches = []
        i = 0
        
        # 创建批次
        while i < len(paragraphs_data):
            # 计算批次的结束位置
            end = min(i + batch_size, len(paragraphs_data))
            
            # 如果不是最后一批，包含重叠
            if end < len(paragraphs_data):
                batch = paragraphs_data[i:end + overlap_size]
            else:
                batch = paragraphs_data[i:end]
            
            batches.append({
                'data': batch,
                'start_index': i,
                'end_index': end,
                'has_overlap': end < len(paragraphs_data)
            })
            
            i = end
        
        print(f"📦 分成 {len(batches)} 批进行处理")
        
        # 处理每个批次
        all_results = []
        batch_errors = []
        
        for idx, batch_info in enumerate(batches):
            print(f"🔄 处理第 {idx + 1}/{len(batches)} 批...")
            try:
                batch_result = await self.analyze(batch_info['data'])
                
                if batch_result.get("success"):
                    analysis_results = batch_result.get("analysis_result", [])
                    
                    # 如果有重叠，去除重叠部分的分析结果
                    if batch_info['has_overlap'] and idx < len(batches) - 1:
                        # 只保留到原始批次大小的结果
                        analysis_results = analysis_results[:batch_size]
                    
                    all_results.extend(analysis_results)
                else:
                    batch_errors.append(f"批次 {idx + 1} 失败: {batch_result.get('error', '未知错误')}")
                    
            except Exception as e:
                batch_errors.append(f"批次 {idx + 1} 异常: {str(e)}")
        
        # 构建最终结果
        if all_results:
            final_result = {
                "success": True,
                "analysis_result": all_results,
                "total_paragraphs": len(paragraphs_data),
                "batch_info": {
                    "total_batches": len(batches),
                    "batch_size": batch_size,
                    "overlap_size": overlap_size
                }
            }
            
            if batch_errors:
                final_result["partial_errors"] = batch_errors
            
            # 保存分析结果
            if self.save_results:
                self._save_analysis_result(final_result, paragraphs_data)
            
            return final_result
        else:
            return {
                "success": False,
                "error": "所有批次处理失败",
                "batch_errors": batch_errors,
                "total_paragraphs": len(paragraphs_data)
            }
    
    def _ensure_results_dir(self):
        """确保结果保存目录存在"""
        if not os.path.exists(self.results_dir):
            os.makedirs(self.results_dir)
            print(f"📁 创建AI分析结果目录: {self.results_dir}")
    
    def _save_analysis_result(self, result: Dict[str, Any], paragraphs_data: List[Dict], is_error: bool = False):
        """
        保存AI分析结果到JSON文件
        
        Args:
            result: 分析结果
            paragraphs_data: 原始段落数据
            is_error: 是否为错误结果
        """
        try:
            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            status = "error" if is_error else "success"
            filename = f"ai_analysis_{status}_{timestamp}.json"
            filepath = os.path.join(self.results_dir, filename)
            
            # 准备保存的数据
            save_data = {
                "timestamp": datetime.now().isoformat(),
                "model": self.model,
                "status": status,
                "input_data": {
                    "total_paragraphs": len(paragraphs_data),
                    "paragraphs": paragraphs_data
                },
                "analysis_result": result
            }
            
            # 保存到文件
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
            
            print(f"💾 AI分析结果已保存: {filepath}")
            
        except Exception as e:
            print(f"⚠️ 保存AI分析结果失败: {str(e)}")