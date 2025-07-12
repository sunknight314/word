"""
文档处理器 - 统一的文档上传和预处理模块
"""

import os
from datetime import datetime
from typing import Dict, List, Any, Optional
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


class DocumentProcessor:
    """统一的文档处理器类"""
    
    def __init__(self):
        """初始化文档处理器"""
        pass
    
    def extract_paragraphs_preview(self, file_path: str, preview_length: int = 20) -> Dict[str, Any]:
        """
        提取文档段落预览（用于段落分析）
        
        Args:
            file_path: Word文档文件路径
            preview_length: 每段预览字符长度
            
        Returns:
            包含段落预览信息的字典
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        try:
            doc = Document(file_path)
            
            paragraphs_data = []
            paragraph_count = 0
            
            # 获取文档body元素，用于识别表格
            body = doc._element.body
            
            # 遍历body中的所有元素（段落和表格）
            for element in body:
                # 处理表格
                if element.tag.endswith('tbl'):
                    paragraph_count += 1
                    paragraph_info = {
                        "paragraph_number": paragraph_count,
                        "preview_text": "[表格]",
                        "full_length": 0,
                        "_is_empty": True,  # 内部使用的标记
                        "_local_type": "表格"  # 内部使用的类型
                    }
                    paragraphs_data.append(paragraph_info)
                
                # 处理段落
                elif element.tag.endswith('p'):
                    # 创建段落对象（用于兼容python-docx的API）
                    from docx.text.paragraph import Paragraph
                    paragraph = Paragraph(element, doc)
                    
                    # 使用更准确的方法检查是否包含文本
                    text_nodes = element.xpath('.//w:t')
                    has_text = any(t.text.strip() for t in text_nodes if t.text)
                    
                    if has_text:
                        print("文本文本文本文本")
                        # 这是文本段落
                        paragraph_count += 1
                        paragraph_text = paragraph.text.strip()
                        
                        # 提取前N个字符作为预览
                        preview_text = paragraph_text[:preview_length]
                        if len(paragraph_text) > preview_length:
                            preview_text += "..."
                        
                        paragraph_info = {
                            "paragraph_number": paragraph_count,
                            "preview_text": preview_text,
                            "full_length": len(paragraph_text)
                        }
                        paragraphs_data.append(paragraph_info)
                    else:
                        # 不是文本段落，检查是否包含特殊内容
                        special_type = self._identify_empty_paragraph_type(paragraph)
                        
                        # 所有非文本段落都保留（包括空段落）
                        paragraph_count += 1
                        paragraph_info = {
                            "paragraph_number": paragraph_count,
                            "preview_text": f"[{special_type}]",
                            "full_length": 0,
                            "_is_empty": True,  # 内部使用的标记
                            "_local_type": special_type  # 内部使用的类型
                        }
                        paragraphs_data.append(paragraph_info)
            
            # 构建最终结果
            result = {
                "document_info": {
                    "total_paragraphs": paragraph_count,
                    "file_path": file_path,
                    "extracted_at": datetime.now().isoformat(),
                    "preview_length": preview_length
                },
                "paragraphs": paragraphs_data
            }
            
            return result
            
        except Exception as e:
            raise Exception(f"提取段落预览失败: {str(e)}")
    
    def extract_full_document_text(self, file_path: str) -> Dict[str, Any]:
        """
        提取文档完整文本内容（用于格式分析）
        
        Args:
            file_path: Word文档文件路径
            
        Returns:
            包含完整文档内容的字典
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        try:
            doc = Document(file_path)
            
            all_paragraphs = []
            full_text = ""
            paragraph_count = 0
            
            # 获取文档body元素，用于识别表格
            body = doc._element.body
            
            # 遍历body中的所有元素（段落和表格）
            for element in body:
                # 处理表格
                if element.tag.endswith('tbl'):
                    paragraph_count += 1
                    paragraph_info = {
                        "paragraph_number": paragraph_count,
                        "text": "[表格]",
                        "length": 0,
                        "is_empty": True,
                        "local_type": "表格"
                    }
                    
                    all_paragraphs.append(paragraph_info)
                    full_text += "[表格]\n"
                
                # 处理段落
                elif element.tag.endswith('p'):
                    # 创建段落对象（用于兼容python-docx的API）
                    from docx.text.paragraph import Paragraph
                    paragraph = Paragraph(element, doc)
                    
                    # 使用更准确的方法检查是否包含文本
                    text_nodes = element.xpath('.//w:t')
                    has_text = any(t.text and t.text.strip() for t in text_nodes if t.text)
                    
                    if has_text:
                        # 这是文本段落
                        paragraph_count += 1
                        paragraph_text = paragraph.text.strip()
                        
                        # 保存完整段落信息
                        paragraph_info = {
                            "paragraph_number": paragraph_count,
                            "text": paragraph_text,
                            "length": len(paragraph_text),
                            "is_empty": False,
                            "local_type": "text"
                        }
                        
                        all_paragraphs.append(paragraph_info)
                        full_text += paragraph_text + "\n"
                    else:
                        # 不是文本段落，检查是否包含特殊内容
                        special_type = self._identify_empty_paragraph_type(paragraph)
                        
                        # 所有非文本段落都保留（包括空段落）
                        paragraph_count += 1
                        paragraph_info = {
                            "paragraph_number": paragraph_count,
                            "text": f"[{special_type}]",
                            "length": 0,
                            "is_empty": True,
                            "local_type": special_type
                        }
                        
                        all_paragraphs.append(paragraph_info)
                        full_text += f"[{special_type}]\n"
            
            # 构建最终结果
            result = {
                "document_info": {
                    "total_paragraphs": paragraph_count,
                    "total_length": len(full_text),
                    "file_path": file_path,
                    "extracted_at": datetime.now().isoformat()
                },
                "paragraphs": all_paragraphs,
                "full_text": full_text.strip()
            }
            
            return result
            
        except Exception as e:
            raise Exception(f"提取文档内容失败: {str(e)}")
    
    def estimate_token_count(self, text: str, chars_per_token: int = 4) -> int:
        """
        估算文本的token数量
        
        Args:
            text: 文本内容
            chars_per_token: 每个token平均字符数（中文约2-3，英文约4-5）
            
        Returns:
            估算的token数量
        """
        # 简单估算：中英文混合取平均值
        return len(text) // chars_per_token
    
    def should_batch_process(self, data: Any, max_tokens: int = 3000) -> bool:
        """
        判断是否需要分批处理
        
        Args:
            data: 要处理的数据
            max_tokens: 最大token限制
            
        Returns:
            是否需要分批处理
        """
        if isinstance(data, str):
            # 字符串类型：估算token数量
            estimated_tokens = self.estimate_token_count(data)
            return estimated_tokens > max_tokens
        elif isinstance(data, list):
            # 列表类型：按数量判断
            return len(data) > 10
        else:
            return False
    
    def split_paragraphs_for_batch(self, paragraphs_data: List[Dict], batch_size: int = 10) -> List[List[Dict]]:
        """
        将段落数据分批
        
        Args:
            paragraphs_data: 段落数据列表
            batch_size: 每批数量
            
        Returns:
            分批后的段落数据
        """
        batches = []
        for i in range(0, len(paragraphs_data), batch_size):
            batch = paragraphs_data[i:i + batch_size]
            batches.append(batch)
        return batches
    
    def split_text_for_batch(self, text: str, max_chars: int = 12000) -> List[str]:
        """
        将长文本按字符数分批
        
        Args:
            text: 要分割的文本
            max_chars: 每批最大字符数
            
        Returns:
            分批后的文本列表
        """
        if len(text) <= max_chars:
            return [text]
        
        # 尝试按段落分割
        paragraphs = text.split('\n')
        batches = []
        current_batch = ""
        
        for paragraph in paragraphs:
            # 如果加上这个段落会超出限制
            if len(current_batch + paragraph + '\n') > max_chars:
                if current_batch:
                    batches.append(current_batch.strip())
                    current_batch = paragraph + '\n'
                else:
                    # 单个段落就超出限制，强制分割
                    for i in range(0, len(paragraph), max_chars):
                        batches.append(paragraph[i:i + max_chars])
            else:
                current_batch += paragraph + '\n'
        
        # 添加最后一批
        if current_batch:
            batches.append(current_batch.strip())
        
        return batches
    
    def _identify_empty_paragraph_type(self, paragraph) -> str:
        """
        识别空段落的类型（图片、表格、公式等）
        
        Args:
            paragraph: Word段落对象
            
        Returns:
            段落类型字符串
        """
        try:
            p_element = paragraph._element
            # 匹配公式（跨命名空间）
            if p_element.xpath('.//*[local-name()="oMath" or local-name()="oMathPara"]'):
                return "Word公式"

            # 2. 检测OLE对象（Visio/MathType/其他）
            if ole_objects := p_element.xpath('.//*[local-name()="OLEObject"]'):
                for ole in ole_objects:
                    prog_id = ole.get('ProgID', '')
                    # 识别Visio对象（关键！）
                    if prog_id.startswith('Visio.Drawing.'):  # 匹配所有Visio版本
                        return "Visio图形"
                    # 识别MathType公式
                    elif 'MathType' in prog_id or 'Equation.DSMT' in prog_id:
                        return "MathType公式"
                return "嵌入对象"  # 其他未识别的OLE对象

            # 匹配图片
            if p_element.xpath('.//*[local-name()="drawing" or local-name()="pict"]'):
                return "图片"

            return "空段落"
    
        except Exception as e:
            return f"识别错误: {str(e)}"
            
    
    def merge_analysis_results(self, local_paragraphs: List[Dict], ai_analysis: List[Dict]) -> List[Dict]:
        """
        合并本地段落检测结果和AI分析结果
        
        Args:
            local_paragraphs: 本地提取的所有段落信息（包含空段落）
            ai_analysis: AI分析的段落类型结果（仅包含文本段落）
            
        Returns:
            合并后的段落分析结果
        """
        merged_results = []
        
        # 创建AI分析结果的映射（段落号 -> AI分析）
        ai_analysis_map = {}
        for ai_item in ai_analysis:
            para_num = ai_item.get("paragraph_number")
            if para_num:
                ai_analysis_map[para_num] = ai_item
        
        # 合并本地和AI分析结果
        for local_para in local_paragraphs:
            para_num = local_para.get("paragraph_number")
            
            # 检查是否为空段落
            if local_para.get("_is_empty", False):
                # 空段落：使用本地检测的类型
                merged_item = {
                    "paragraph_number": para_num,
                    "type": local_para.get("_local_type", "空段落"),
                }
            else:
                # 文本段落：使用AI分析结果
                ai_result = ai_analysis_map.get(para_num, {})
                merged_item = {
                    "paragraph_number": para_num,
                    "type": ai_result.get("type", "paragraph"),
                }
            
            merged_results.append(merged_item)
        
        return merged_results