import os
import uuid
from typing import Optional
from fastapi import UploadFile
import aiofiles


class FileStorage:
    def __init__(self, upload_dir: str = "uploads"):
        # 使用项目目录下的uploads文件夹
        self.upload_dir = os.path.join(os.getcwd(), upload_dir)
        os.makedirs(self.upload_dir, exist_ok=True)
    
    async def save_uploaded_file(self, file: UploadFile, file_type: str = "source") -> dict:
        """
        保存上传的文件
        
        Args:
            file: 上传的文件对象
            file_type: 文件类型 ("source" 或 "format")
            
        Returns:
            包含文件信息的字典
        """
        # 生成唯一的文件ID
        file_id = f"{file_type}_{uuid.uuid4().hex[:8]}"
        
        # 获取文件扩展名
        file_extension = os.path.splitext(file.filename)[1]
        if not file_extension:
            file_extension = ".docx"
        
        # 生成保存路径
        saved_filename = f"{file_id}{file_extension}"
        file_path = os.path.join(self.upload_dir, saved_filename)
        
        # 异步保存文件
        async with aiofiles.open(file_path, 'wb') as f:
            content = await file.read()
            await f.write(content)
        
        return {
            "file_id": file_id,
            "original_filename": file.filename,
            "saved_filename": saved_filename,
            "file_path": file_path,
            "file_size": len(content)
        }
    
    def get_file_path(self, file_id: str) -> Optional[str]:
        """
        根据文件ID获取文件路径
        
        Args:
            file_id: 文件ID
            
        Returns:
            文件路径，如果文件不存在返回None
        """
        # 在上传目录中查找以file_id开头的文件
        for filename in os.listdir(self.upload_dir):
            if filename.startswith(file_id):
                file_path = os.path.join(self.upload_dir, filename)
                if os.path.exists(file_path):
                    return file_path
        return None