"""
文档处理器 - 统一的文档上传和预处理模块
"""

import os
from datetime import datetime
from typing import Dict, List, Any, Optional
from docx import Document


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
            
            for paragraph in doc.paragraphs:
                paragraph_text = paragraph.text.strip()
                
                # 跳过空段落
                if not paragraph_text:
                    continue
                    
                paragraph_count += 1
                
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
            
            for paragraph in doc.paragraphs:
                paragraph_text = paragraph.text.strip()
                
                # 跳过空段落
                if not paragraph_text:
                    continue
                    
                paragraph_count += 1
                
                # 保存完整段落信息
                paragraph_info = {
                    "paragraph_number": paragraph_count,
                    "text": paragraph_text,
                    "length": len(paragraph_text)
                }
                
                all_paragraphs.append(paragraph_info)
                full_text += paragraph_text + "\n"
            
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