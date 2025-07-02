from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List, Dict, Any
from app.services.file_storage import FileStorage
from app.services.document_parser import DocumentParser

router = APIRouter()

# 初始化服务
file_storage = FileStorage()
document_parser = DocumentParser()

# 步骤1：上传源文件
@router.post("/upload/source")
async def upload_source_file(file: UploadFile = File(..., description="源Word文档")):
    try:
        # 检查文件类型
        if not file.filename.lower().endswith(('.docx', '.doc')):
            raise HTTPException(status_code=400, detail="只支持.docx和.doc格式的文件")
        
        # 保存文件
        file_info = await file_storage.save_uploaded_file(file, "source")
        
        return {
            "file_id": file_info["file_id"],
            "filename": file_info["original_filename"],
            "message": "源文件上传成功"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"上传失败: {str(e)}")

# 步骤2：解析源文件 (包含AI分析)
@router.get("/parse/source/{file_id}")
async def parse_source_file(file_id: str):
    try:
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="文件不存在")
        
        # 使用AI分析文档 (use_mock=False 使用真实API)
        result = await document_parser.analyze_document_with_ai(file_path, use_mock=False)
        result["file_id"] = file_id
        
        return result
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="文件不存在")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"AI分析失败: {str(e)}")

# 步骤3：上传格式要求文件
@router.post("/upload/format")
async def upload_format_file(file: UploadFile = File(..., description="格式要求文档")):
    return {
        "file_id": "format_456",
        "filename": file.filename,
        "message": "格式要求文件上传成功"
    }

# 步骤4：解析格式要求文件
@router.get("/parse/format/{file_id}")
async def parse_format_file(file_id: str):
    # 模拟返回解析结果
    mock_result = {
        "file_id": file_id,
        "format_requirements": {
            "font_family": "Times New Roman",
            "font_size": 14,
            "line_spacing": 2.0,
            "margins": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
            "header_footer": True
        }
    }
    return mock_result

# 步骤5：执行格式转换
@router.post("/process")
async def process_documents(source_id: str, format_id: str):
    return {
        "task_id": "task_789",
        "status": "processing",
        "message": "格式转换已开始"
    }

# 查询处理状态
@router.get("/status/{task_id}")
async def get_processing_status(task_id: str):
    return {
        "task_id": task_id,
        "status": "completed",
        "progress": 100,
        "download_url": f"/api/download/{task_id}"
    }

# 下载处理结果
@router.get("/download/{task_id}")
async def download_result(task_id: str):
    return {"message": f"文件下载功能待实现 - {task_id}"}