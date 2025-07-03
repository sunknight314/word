"""
格式配置生成器 - 重构后的统一架构
"""

import os
import json
import re
from datetime import datetime
from typing import Dict, Any, Tuple
from .ai_client_base import AIClientBase
from .document_processor import DocumentProcessor


class FormatConfigGenerator(AIClientBase):
    """格式配置生成器类"""
    
    def __init__(self, api_base_url: str = None, api_key: str = None, model: str = None, model_config: str = None):
        """
        初始化格式配置生成器
        
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
    
    def get_prompts(self, input_data: str) -> Tuple[str, str]:
        """
        获取格式配置生成的系统和用户提示词
        
        Args:
            input_data: 文档完整内容
            
        Returns:
            (system_prompt, user_prompt) 元组
        """
        from .ai_prompts import get_format_config_generation_prompt
        return get_format_config_generation_prompt(input_data)
    
    async def process_ai_response(self, ai_response: str) -> Dict[str, Any]:
        """
        处理AI响应，解析格式配置
        
        Args:
            ai_response: AI原始响应
            
        Returns:
            处理后的配置结果
        """
        try:
            print("🔍 开始提取格式配置JSON...")
            json_content = self.extract_json_from_content(ai_response)
            print(f"📋 提取的JSON长度: {len(json_content)} 字符")
            
            import json
            format_config = json.loads(json_content)
            print("✅ 格式配置JSON解析成功")
            
            return {
                "success": True,
                "format_config": format_config,
                "raw_response": ai_response
            }
            
        except json.JSONDecodeError as e:
            print(f"❌ 格式配置JSON解析失败: {str(e)}")
            if 'json_content' in locals():
                print(f"🔍 尝试解析的内容: {json_content}")
            return {
                "success": False,
                "error": f"JSON解析失败: {str(e)}",
                "raw_response": ai_response
            }
    
    async def generate_format_config_with_ai(self, document_content: str) -> Dict[str, Any]:
        """
        使用AI生成格式配置（一次性处理）
        
        Args:
            document_content: 文档完整内容
            
        Returns:
            AI生成的结果
        """
        print(f"📄 开始格式配置生成，文档长度: {len(document_content)} 字符")
        
        # 直接使用AI分析，一次性处理
        return await self.analyze(document_content)
    
    def _old_extract_json_from_content(self, content: str) -> str:
        """
        从AI响应内容中提取JSON部分
        
        Args:
            content: AI响应内容
            
        Returns:
            JSON字符串
        """
        # 去除可能的markdown代码块标记
        content = content.strip()
        
        # 尝试多种方法提取JSON
        # 方法1: 查找完整的JSON对象
        import re
        
        # 使用正则表达式查找JSON对象
        json_pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
        matches = re.findall(json_pattern, content, re.DOTALL)
        
        if matches:
            # 找到最长的匹配项（通常是完整的JSON）
            longest_match = max(matches, key=len)
            try:
                # 验证是否为有效JSON
                json.loads(longest_match)
                return longest_match
            except json.JSONDecodeError:
                pass
        
        # 方法2: 查找JSON开始和结束位置
        start_markers = ['{', '```json\n{', '```\n{']
        end_markers = ['}', '}\n```', '}\n```']
        
        json_start = -1
        json_end = -1
        
        # 找到JSON开始位置
        for marker in start_markers:
            idx = content.find(marker)
            if idx != -1:
                if marker.startswith('```'):
                    json_start = content.find('{', idx)
                else:
                    json_start = idx
                break
        
        # 从后往前找JSON结束位置
        for marker in end_markers:
            idx = content.rfind(marker)
            if idx != -1:
                if marker.endswith('```'):
                    json_end = content.rfind('}', 0, idx) + 1
                else:
                    json_end = idx + 1
                break
        
        if json_start != -1 and json_end != -1 and json_end > json_start:
            extracted = content[json_start:json_end]
            try:
                # 验证提取的内容是否为有效JSON
                json.loads(extracted)
                return extracted
            except json.JSONDecodeError:
                pass
        
        # 方法3: 尝试修复常见的JSON格式问题
        # 移除可能的前后缀
        cleaned_content = content
        
        # 移除markdown代码块标记
        if '```' in cleaned_content:
            cleaned_content = re.sub(r'```json\s*', '', cleaned_content)
            cleaned_content = re.sub(r'```\s*$', '', cleaned_content)
        
        # 查找第一个{和最后一个}
        first_brace = cleaned_content.find('{')
        last_brace = cleaned_content.rfind('}')
        
        if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
            potential_json = cleaned_content[first_brace:last_brace + 1]
            try:
                # 尝试修复常见的JSON问题
                # 移除可能的转义字符问题
                fixed_json = potential_json.replace('\\"', '"').replace('\\n', '\n')
                json.loads(fixed_json)
                return fixed_json
            except json.JSONDecodeError:
                pass
        
        # 如果所有方法都失败，返回原内容
        return content
    
    def validate_and_fix_config(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """
        验证和修正格式配置
        """
        return self._old_validate_and_fix_config(config)
    
    def _old_validate_and_fix_config(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """
        验证和修正格式配置
        
        Args:
            config: AI生成的配置
            
        Returns:
            验证和修正后的配置
        """
        # 默认的format_config.json结构
        default_config = {
            "document_info": {
                "title": "AI生成的文档",
                "author": "系统自动生成",
                "description": "基于AI分析的文档格式配置"
            },
            "page_settings": {
                "margins": {
                    "top": "2.54cm",
                    "bottom": "2.54cm",
                    "left": "3.17cm",
                    "right": "3.17cm"
                },
                "orientation": "portrait",
                "size": "A4"
            },
            "styles": {
                "title": {
                    "name": "CustomTitle",
                    "font": {
                        "chinese": "黑体",
                        "english": "Times New Roman",
                        "size": "22pt",
                        "bold": True,
                        "italic": False
                    },
                    "paragraph": {
                        "alignment": "center",
                        "line_spacing": "20pt",
                        "space_before": "0pt",
                        "space_after": "24pt",
                        "first_line_indent": "0pt",
                        "left_indent": "0pt",
                        "right_indent": "0pt",
                        "hanging_indent": "0pt"
                    },
                    "outline_level": None
                },
                "heading1": {
                    "name": "CustomHeading1",
                    "font": {
                        "chinese": "黑体",
                        "english": "Times New Roman",
                        "size": "16pt",
                        "bold": True,
                        "italic": False
                    },
                    "paragraph": {
                        "alignment": "center",
                        "line_spacing": "20pt",
                        "space_before": "24pt",
                        "space_after": "18pt",
                        "first_line_indent": "0pt",
                        "left_indent": "0pt",
                        "right_indent": "0pt",
                        "hanging_indent": "0pt"
                    },
                    "outline_level": 0
                },
                "heading2": {
                    "name": "CustomHeading2",
                    "font": {
                        "chinese": "宋体",
                        "english": "Times New Roman",
                        "size": "15pt",
                        "bold": True,
                        "italic": False
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": "20pt",
                        "space_before": "18pt",
                        "space_after": "12pt",
                        "first_line_indent": "0pt",
                        "left_indent": "0pt",
                        "right_indent": "0pt",
                        "hanging_indent": "0pt"
                    },
                    "outline_level": 1
                },
                "heading3": {
                    "name": "CustomHeading3",
                    "font": {
                        "chinese": "宋体",
                        "english": "Times New Roman",
                        "size": "14pt",
                        "bold": True,
                        "italic": False
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": "20pt",
                        "space_before": "12pt",
                        "space_after": "6pt",
                        "first_line_indent": "28pt",
                        "left_indent": "0pt",
                        "right_indent": "0pt",
                        "hanging_indent": "0pt"
                    },
                    "outline_level": 2
                },
                "paragraph": {
                    "name": "CustomBody",
                    "font": {
                        "chinese": "宋体",
                        "english": "Times New Roman",
                        "size": "12pt",
                        "bold": False,
                        "italic": False
                    },
                    "paragraph": {
                        "alignment": "left",
                        "line_spacing": "20pt",
                        "space_before": "0pt",
                        "space_after": "0pt",
                        "first_line_indent": "24pt",
                        "left_indent": "0pt",
                        "right_indent": "0pt",
                        "hanging_indent": "0pt"
                    },
                    "outline_level": None
                }
            },
            "toc_settings": {
                "title": "目录",
                "levels": "1-3",
                "hyperlinks": True,
                "tab_leader": "dots",
                "page_format": "roman",
                "page_start": 1,
                "headers": {
                    "odd": "目录",
                    "even": "目录"
                }
            },
            "page_numbering": {
                "toc_section": {
                    "format": "upperRoman",
                    "start": 1,
                    "restart": True,
                    "template": "{page}"
                },
                "content_sections": {
                    "format": "decimal",
                    "start": 1,
                    "restart_first_chapter": True,
                    "continue_others": True,
                    "template": "第 {page} 页"
                }
            },
            "headers_footers": {
                "odd_even_different": True,
                "first_page_different": False,
                "toc_section": {
                    "odd_header": "目录",
                    "even_header": "目录",
                    "odd_footer": "",
                    "even_footer": ""
                },
                "content_sections": {
                    "odd_header_template": "{chapter_title}",
                    "even_header_template": "Word文档格式优化项目",
                    "odd_footer": "",
                    "even_footer": ""
                }
            },
            "section_breaks": {
                "between_toc_and_content": "odd_page",
                "between_chapters": "odd_page",
                "between_sections": "continuous"
            },
            "document_structure": {
                "auto_generate_toc": True,
                "chapter_start_page": "odd",
                "chapter_numbering": "arabic",
                "section_numbering": "decimal"
            }
        }
        
        # 递归合并配置 - AI配置优先，默认配置填补空缺
        def merge_dict(default: dict, ai_config: dict) -> dict:
            result = ai_config.copy()  # 以AI配置为基础
            
            # 遍历默认配置，只添加AI配置中缺失的部分
            for key, default_value in default.items():
                if key not in result:
                    # AI配置中没有这个键，使用默认值
                    result[key] = default_value
                elif isinstance(default_value, dict) and isinstance(result[key], dict):
                    # 两边都是字典，递归合并
                    result[key] = merge_dict(default_value, result[key])
            
            return result
        
        # 合并AI生成的配置和默认配置 - AI配置优先
        merged_config = merge_dict(default_config, config)
        
        return merged_config
    
    async def process_format_document(self, file_path: str) -> Dict[str, Any]:
        """
        处理格式要求文档的完整流程
        
        Args:
            file_path: 格式要求文档路径
            
        Returns:
            处理结果
        """
        try:
            # 1. 提取文档内容
            print(f"📄 开始提取文档内容: {file_path}")
            content_result = self.document_processor.extract_full_document_text(file_path)
            
            # 2. 使用AI生成格式配置
            print("🤖 开始AI分析生成格式配置...")
            ai_result = await self.generate_format_config_with_ai(content_result["full_text"])
            
            # 3. 验证和修正配置
            if ai_result.get("success") and ai_result.get("format_config"):
                print("✅ AI生成成功，开始验证和修正配置...")
                final_config = self.validate_and_fix_config(ai_result["format_config"])
                ai_result["format_config"] = final_config
            else:
                print("❌ AI生成失败，使用默认配置...")
                default_config = self.validate_and_fix_config({})
                ai_result["format_config"] = default_config
                ai_result["fallback"] = True
            
            # 4. 整合文档信息和AI结果
            result = {
                "document_info": content_result["document_info"],
                **ai_result
            }
            
            return result
                
        except Exception as e:
            print(f"❌ 处理失败: {str(e)}")
            # 返回默认配置作为降级方案
            default_config = self.validate_and_fix_config({})
            
            return {
                "success": False,
                "format_config": default_config,
                "error": str(e),
                "fallback": True
            }