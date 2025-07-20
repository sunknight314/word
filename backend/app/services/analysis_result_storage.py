"""
分析结果存储服务
保存段落分析和格式配置分析的结果
"""
import json
import os
from datetime import datetime
from typing import Dict, Any, Optional
from pathlib import Path
import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)



class AnalysisResultStorage:
    """分析结果存储类"""
    
    def __init__(self, storage_dir: str = "analysis_results"):
        """
        初始化存储服务
        
        Args:
            storage_dir: 存储目录路径
        """
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        
        # 创建子目录
        self.paragraph_dir = self.storage_dir / "paragraph_analysis"
        self.format_dir = self.storage_dir / "format_config"
        self.paragraph_dir.mkdir(exist_ok=True)
        self.format_dir.mkdir(exist_ok=True)
    
    def save_paragraph_analysis(self, file_id: str, analysis_result: Dict[str, Any]) -> str:
        """
        保存段落分析结果
        
        Args:
            file_id: 文件ID
            analysis_result: 分析结果
            
        Returns:
            保存的文件路径
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"paragraph_{file_id}_{timestamp}.json"
        filepath = self.paragraph_dir / filename
        
        # 添加元数据
        data = {
            "file_id": file_id,
            "timestamp": datetime.now().isoformat(),
            "analysis_type": "paragraph_analysis",
            "result": analysis_result
        }
        
        # 保存到文件
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return str(filepath)
    
    def save_format_config(self, file_id: str, format_config: Dict[str, Any], 
                          ai_info: Optional[Dict[str, Any]] = None) -> str:
        """
        保存格式配置分析结果
        
        Args:
            file_id: 文件ID
            format_config: 格式配置
            ai_info: AI模型信息
            
        Returns:
            保存的文件路径
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"format_{file_id}_{timestamp}.json"
        filepath = self.format_dir / filename
        
        # 添加元数据
        data = {
            "file_id": file_id,
            "timestamp": datetime.now().isoformat(),
            "analysis_type": "format_config",
            "ai_info": ai_info,
            "format_config": format_config
        }
        
        # 保存到文件
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return str(filepath)
    
    def get_latest_paragraph_analysis(self, file_id: str) -> Optional[Dict[str, Any]]:
        """
        获取最新的段落分析结果
        
        Args:
            file_id: 文件ID
            
        Returns:
            分析结果，如果不存在则返回None
        """
        # 查找匹配的文件
        pattern = f"paragraph_{file_id}_*.json"
        files = list(self.paragraph_dir.glob(pattern))
        
        if not files:
            return None
        
        # 按修改时间排序，获取最新的
        latest_file = max(files, key=lambda f: f.stat().st_mtime)
        
        logger.info(f"段落分析json路径: {latest_file}")

        with open(latest_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        return data
    
    def get_latest_format_config(self, file_id: str) -> Optional[Dict[str, Any]]:
        """
        获取最新的格式配置
        
        Args:
            file_id: 文件ID
            
        Returns:
            格式配置，如果不存在则返回None
        """
        # 查找匹配的文件
        pattern = f"format_{file_id}_*.json"
        files = list(self.format_dir.glob(pattern))
        
        if not files:
            return None
        
        # 按修改时间排序，获取最新的
        latest_file = max(files, key=lambda f: f.stat().st_mtime)

        logger.info(f"格式配置json路径: {latest_file}")
        
        with open(latest_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        return data
    
    def get_all_analysis_for_file(self, file_id: str) -> Dict[str, Any]:
        """
        获取文件的所有分析结果
        
        Args:
            file_id: 文件ID
            
        Returns:
            包含段落分析和格式配置的字典
        """
        return {
            "paragraph_analysis": self.get_latest_paragraph_analysis(file_id),
            "format_config": self.get_latest_format_config(file_id)
        }
    
    def list_all_analyses(self) -> Dict[str, Any]:
        """
        列出所有保存的分析结果
        
        Returns:
            分析结果列表
        """
        paragraph_files = [
            {
                "filename": f.name,
                "filepath": str(f),
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            }
            for f in self.paragraph_dir.glob("*.json")
        ]
        
        format_files = [
            {
                "filename": f.name,
                "filepath": str(f),
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            }
            for f in self.format_dir.glob("*.json")
        ]
        
        return {
            "paragraph_analyses": sorted(paragraph_files, key=lambda x: x["modified"], reverse=True),
            "format_configs": sorted(format_files, key=lambda x: x["modified"], reverse=True)
        }
    
    def cleanup_old_files(self, days: int = 30):
        """
        清理超过指定天数的旧文件
        
        Args:
            days: 保留天数
        """
        from datetime import timedelta
        
        cutoff_time = datetime.now() - timedelta(days=days)
        
        # 清理段落分析文件
        for file in self.paragraph_dir.glob("*.json"):
            if datetime.fromtimestamp(file.stat().st_mtime) < cutoff_time:
                file.unlink()
        
        # 清理格式配置文件
        for file in self.format_dir.glob("*.json"):
            if datetime.fromtimestamp(file.stat().st_mtime) < cutoff_time:
                file.unlink()