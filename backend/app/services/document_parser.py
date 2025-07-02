from docx import Document
import json
from typing import List, Dict, Any
import os
from app.services.ai_analyzer import AIAnalyzer


class DocumentParser:
    def __init__(self):
        self.ai_analyzer = AIAnalyzer()
    
    def extract_paragraphs_info(self, file_path: str) -> Dict[str, Any]:
        """
        提取Word文档中每段的前20个字符
        
        Args:
            file_path: Word文档文件路径
            
        Returns:
            包含段落信息的字典
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        try:
            # 打开Word文档
            doc = Document(file_path)
            
            paragraphs_info = []
            paragraph_count = 0
            
            for paragraph in doc.paragraphs:
                # 获取段落文本
                paragraph_text = paragraph.text.strip()
                
                # 跳过空段落
                if not paragraph_text:
                    continue
                    
                paragraph_count += 1
                
                # 提取前20个字符
                preview_text = paragraph_text[:20]
                
                # 构建段落信息
                paragraph_info = {
                    "paragraph_number": paragraph_count,
                    "preview_text": preview_text,
                    "full_text_length": len(paragraph_text),
                    "is_complete": len(paragraph_text) <= 20
                }
                
                paragraphs_info.append(paragraph_info)
            
            # 构建最终结果
            result = {
                "document_info": {
                    "total_paragraphs": paragraph_count,
                    "file_path": file_path
                },
                "paragraphs": paragraphs_info
            }
            
            return result
            
        except Exception as e:
            raise Exception(f"解析文档失败: {str(e)}")
    
    def save_paragraphs_to_json(self, file_path: str, output_path: str = None) -> str:
        """
        提取段落信息并保存为JSON文件
        
        Args:
            file_path: Word文档文件路径
            output_path: JSON输出文件路径，如果为None则生成默认路径
            
        Returns:
            JSON文件路径
        """
        result = self.extract_paragraphs_info(file_path)
        
        if output_path is None:
            # 生成默认输出路径
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = f"/tmp/{base_name}_paragraphs.json"
        
        # 保存为JSON文件
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        return output_path
    
    async def analyze_document_with_ai(self, file_path: str, use_mock: bool = True) -> Dict[str, Any]:
        """
        使用AI分析文档段落类型
        
        Args:
            file_path: Word文档文件路径
            use_mock: 是否使用模拟分析（True=模拟，False=真实API调用）
            
        Returns:
            包含AI分析结果的字典
        """
        # 先提取段落信息
        paragraphs_info = self.extract_paragraphs_info(file_path)
        
        # 准备AI分析数据
        paragraphs_for_ai = []
        for para in paragraphs_info["paragraphs"]:
            paragraphs_for_ai.append({
                "paragraph_number": para["paragraph_number"],
                "preview_text": para["preview_text"]
            })
        
        # 进行AI分析
        if use_mock:
            # 使用模拟分析
            ai_analysis = self.ai_analyzer.create_mock_analysis(paragraphs_for_ai)
        else:
            # 使用真实API调用
            ai_result = await self.ai_analyzer.analyze_paragraphs(paragraphs_for_ai)
            if ai_result["success"]:
                ai_analysis = ai_result["result"]
            else:
                raise Exception(f"AI分析失败: {ai_result['error']}")
        
        # 合并结果
        result = {
            "file_info": paragraphs_info["document_info"],
            "paragraphs": paragraphs_info["paragraphs"],
            "ai_analysis": ai_analysis,
            "analysis_summary": self._create_analysis_summary(ai_analysis)
        }
        
        return result
    
    def _create_analysis_summary(self, ai_analysis: Dict[str, Any]) -> Dict[str, Any]:
        """
        创建分析结果摘要
        
        Args:
            ai_analysis: AI分析结果
            
        Returns:
            分析摘要
        """
        if "analysis_result" not in ai_analysis:
            return {"error": "无效的分析结果"}
        
        # 统计各类型段落数量
        type_counts = {}
        total_confidence = 0
        
        for item in ai_analysis["analysis_result"]:
            para_type = item.get("type", "unknown")
            confidence = item.get("confidence", 0)
            
            # 处理置信度可能是字符串的情况
            if isinstance(confidence, str):
                confidence_map = {"high": 0.9, "medium": 0.7, "low": 0.5}
                confidence = confidence_map.get(confidence.lower(), 0.8)
            
            type_counts[para_type] = type_counts.get(para_type, 0) + 1
            total_confidence += confidence
        
        avg_confidence = total_confidence / len(ai_analysis["analysis_result"]) if ai_analysis["analysis_result"] else 0
        
        return {
            "total_paragraphs": len(ai_analysis["analysis_result"]),
            "type_distribution": type_counts,
            "average_confidence": round(avg_confidence, 3),
            "structure_detected": self._detect_document_structure(type_counts)
        }
    
    def _detect_document_structure(self, type_counts: Dict[str, int]) -> str:
        """
        检测文档结构类型
        
        Args:
            type_counts: 各类型段落数量统计
            
        Returns:
            文档结构描述
        """
        has_title = type_counts.get("title", 0) > 0
        has_h1 = type_counts.get("heading1", 0) > 0
        has_h2 = type_counts.get("heading2", 0) > 0
        has_h3 = type_counts.get("heading3", 0) > 0
        
        if has_title and has_h1 and has_h2:
            if has_h3:
                return "完整层级结构（标题-一级-二级-三级）"
            else:
                return "标准层级结构（标题-一级-二级）"
        elif has_h1 and has_h2:
            return "基本层级结构（一级-二级）"
        elif has_h1:
            return "简单结构（仅一级标题）"
        else:
            return "平铺结构（无明显层级）"