"""
格式配置生成器 - 重构后的统一架构
"""

import os
import json
import re
import logging
from datetime import datetime
from typing import Dict, Any, Tuple
from .ai_client_base import AIClientBase
from .document_processor import DocumentProcessor
from app.utils import UnitConverter

logger = logging.getLogger(__name__)


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
        验证和修正格式配置，包括单位转换
        
        Args:
            config: AI生成的配置（包含原始单位）
            
        Returns:
            验证、修正并转换单位后的配置
        """
        # 保存AI原始配置（如果不为空）
        if config:
            self._save_ai_original_config(config)
        
        # 使用默认配置作为基础
        default_config = self._get_default_config()
        
        # 智能合并配置：只添加AI未识别的样式
        final_config = self._smart_merge_config(default_config, config)
        
        # 保存智能合并后的配置
        if config:  # 只在有AI配置时保存合并结果
            self._save_merged_config(final_config)
        
        # 转换所有单位为pt
        final_config = self._convert_all_units(final_config)
        
        return final_config
    
    def _save_ai_original_config(self, config: Dict[str, Any]) -> None:
        """保存AI返回的原始配置JSON"""
        try:
            # 生成唯一文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_id = self._generate_config_id(config)
            filename = f"ai_original_{file_id}_{timestamp}.json"
            
            # 创建保存目录
            save_dir = "backend/analysis_results/ai_original_configs"
            os.makedirs(save_dir, exist_ok=True)
            
            # 保存文件
            file_path = os.path.join(save_dir, filename)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            logger.info(f"AI原始配置已保存: {file_path}")
            print(f"💾 AI原始配置已保存: {filename}")
            
        except Exception as e:
            logger.error(f"保存AI原始配置失败: {str(e)}")
            print(f"❌ 保存AI原始配置失败: {str(e)}")
    
    def _save_merged_config(self, config: Dict[str, Any]) -> None:
        """保存智能合并后的配置JSON"""
        try:
            # 生成唯一文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_id = self._generate_config_id(config)
            filename = f"merged_config_{file_id}_{timestamp}.json"
            
            # 创建保存目录
            save_dir = "backend/analysis_results/merged_configs"
            os.makedirs(save_dir, exist_ok=True)
            
            # 保存文件
            file_path = os.path.join(save_dir, filename)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            logger.info(f"智能合并配置已保存: {file_path}")
            print(f"💾 智能合并配置已保存: {filename}")
            
        except Exception as e:
            logger.error(f"保存智能合并配置失败: {str(e)}")
            print(f"❌ 保存智能合并配置失败: {str(e)}")
    
    def _generate_config_id(self, config: Dict[str, Any]) -> str:
        """生成配置的唯一标识符"""
        try:
            import hashlib
            # 使用配置内容生成hash作为ID
            config_str = json.dumps(config, sort_keys=True, ensure_ascii=False)
            return hashlib.md5(config_str.encode()).hexdigest()[:8]
        except:
            # 如果hash生成失败，使用时间戳
            return datetime.now().strftime("%H%M%S")
    
    def _get_default_config(self) -> Dict[str, Any]:
        """从JSON文件获取默认配置"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), '..', '..', 'config', 'default_format_config.json')
            config_path = os.path.abspath(config_path)
            
            if not os.path.exists(config_path):
                logger.error(f"默认配置文件不存在: {config_path}")
                return self._get_fallback_config()
            
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            logger.info(f"已加载默认配置: {config_path}")
            return config
            
        except Exception as e:
            logger.error(f"加载默认配置失败: {str(e)}")
            print(f"❌ 加载默认配置失败: {str(e)}")
            return self._get_fallback_config()
    
    def _get_fallback_config(self) -> Dict[str, Any]:
        """获取降级默认配置（当JSON文件加载失败时使用）"""
        return {
            "document_info": {
                "title": "系统自动生成",
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
                "Title": self._get_default_style("Title", "二号", True, "center"),
                "Heading1": self._get_default_style("Heading1", "三号", True, "center"),
                "Heading2": self._get_default_style("Heading2", "小三", True, "left"),
                "Heading3": self._get_default_style("Heading3", "四号", True, "left"),
                "Heading4": self._get_default_style("Heading4", "小四", True, "left"),
                "Normal": self._get_default_style("Normal", "小四", False, "justify"),
                "AbstractTitleCN": self._get_default_style("AbstractTitleCN", "三号", True, "center"),
                "AbstractTitleEN": self._get_default_style("AbstractTitleEN", "三号", True, "center")
            }
        }
    
    def _get_default_style(self, style_name: str, font_size: str, bold: bool, alignment: str) -> Dict[str, Any]:
        """获取默认样式配置"""
        return {
            "name": style_name,
            "font": {
                "chinese": "宋体",
                "english": "Times New Roman",
                "size": font_size,
                "bold": bold,
                "italic": False
            },
            "paragraph": {
                "alignment": alignment,
                "line_spacing": "20磅",
                "space_before": "0磅",
                "space_after": "0磅",
                "first_line_indent": "2字符" if alignment == "justify" else "0磅",
                "left_indent": "0磅",
                "right_indent": "0磅",
                "hanging_indent": "0磅"
            },
            "outline_level": None
        }
    
    def _deep_merge_config(self, base: Dict, override: Dict) -> Dict:
        """深度合并配置"""
        result = base.copy()
        
        for key, value in override.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._deep_merge_config(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def _smart_merge_config(self, default_config: Dict, ai_config: Dict) -> Dict:
        """
        智能合并配置：优先使用AI配置，只在AI未识别某种样式时才添加默认样式
        处理AI返回的__NOT_SPECIFIED__标志，用默认值替换
        """
        result = ai_config.copy()
        
        # 处理AI配置中的__NOT_SPECIFIED__标志
        result = self._replace_not_specified_values(result, default_config)
        
        # 确保基础结构存在
        if "document_info" not in result:
            result["document_info"] = default_config.get("document_info", {})
        
        if "page_settings" not in result:
            result["page_settings"] = default_config.get("page_settings", {})
        else:
            # 深度合并页面设置
            result["page_settings"] = self._deep_merge_config(
                default_config.get("page_settings", {}), 
                result["page_settings"]
            )
        
        # 智能合并样式：只添加AI未识别的样式
        if "styles" not in result:
            result["styles"] = {}
        
        ai_styles = set(result["styles"].keys())
        default_styles = set(default_config.get("styles", {}).keys())
        
        # 只添加AI未识别的默认样式
        missing_styles = default_styles - ai_styles
        for style_name in missing_styles:
            result["styles"][style_name] = default_config["styles"][style_name]
            logger.info(f"添加默认样式 {style_name}，因为AI未识别此样式")
        
        return result
    
    def _replace_not_specified_values(self, ai_config: Dict, default_config: Dict) -> Dict:
        """
        递归替换AI配置中的__NOT_SPECIFIED__标志为默认值
        
        Args:
            ai_config: AI返回的配置（可能包含__NOT_SPECIFIED__标志）
            default_config: 默认配置
            
        Returns:
            替换后的配置
        """
        def replace_recursive(obj, default_obj=None):
            if isinstance(obj, dict):
                result = {}
                for key, value in obj.items():
                    default_value = default_obj.get(key) if isinstance(default_obj, dict) else None
                    result[key] = replace_recursive(value, default_value)
                return result
            elif isinstance(obj, list):
                return [replace_recursive(item) for item in obj]
            elif obj == "__NOT_SPECIFIED__":
                if default_obj is not None:
                    logger.info(f"将__NOT_SPECIFIED__替换为默认值: {default_obj}")
                    print(f"🔄 将__NOT_SPECIFIED__替换为默认值: {default_obj}")
                    return default_obj
                else:
                    # 如果没有默认值，移除该字段（返回None，稍后过滤）
                    logger.info(f"移除__NOT_SPECIFIED__字段，无对应默认值")
                    print(f"🗑️ 移除__NOT_SPECIFIED__字段，无对应默认值")
                    return None
            else:
                return obj
        
        result = replace_recursive(ai_config, default_config)
        
        # 清理None值（移除__NOT_SPECIFIED__且无默认值的字段）
        result = self._clean_none_values(result)
        
        return result
    
    def _clean_none_values(self, obj):
        """递归清理配置中的None值"""
        if isinstance(obj, dict):
            return {k: self._clean_none_values(v) for k, v in obj.items() if v is not None}
        elif isinstance(obj, list):
            return [self._clean_none_values(item) for item in obj if item is not None]
        else:
            return obj
    
    def _convert_all_units(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """转换配置中的所有单位为pt"""
        # 转换页边距
        if "page_settings" in config and "margins" in config["page_settings"]:
            margins = config["page_settings"]["margins"]
            for key in ["top", "bottom", "left", "right"]:
                if key in margins:
                    margins[key] = UnitConverter.convert_margin(margins[key])
        
        # 转换样式中的单位
        if "styles" in config:
            for style_name, style in config["styles"].items():
                if isinstance(style, dict):
                    # 转换字号
                    if "font" in style and "size" in style["font"]:
                        style["font"]["size"] = UnitConverter.convert_font_size(style["font"]["size"])
                    
                    # 转换段落格式
                    if "paragraph" in style:
                        para = style["paragraph"]
                        # 获取基准字号用于计算相对单位
                        base_size = 12  # 默认
                        if "font" in style and "size" in style["font"]:
                            size_str = style["font"]["size"]
                            if size_str.endswith("pt"):
                                try:
                                    base_size = float(size_str[:-2])
                                except:
                                    pass
                        
                        # 转换各种间距和缩进
                        for spacing_key in ["line_spacing", "space_before", "space_after"]:
                            if spacing_key in para:
                                para[spacing_key] = UnitConverter.convert_spacing(para[spacing_key], base_size)
                        
                        for indent_key in ["first_line_indent", "left_indent", "right_indent", "hanging_indent"]:
                            if indent_key in para:
                                para[indent_key] = UnitConverter.convert_indent(para[indent_key], base_size)
        
        return config
    
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