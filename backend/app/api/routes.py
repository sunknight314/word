from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List, Dict, Any
from app.services.file_storage import FileStorage
from app.services.document_parser import DocumentParser
from app.services.format_config_generator import FormatConfigGenerator
from app.services.ai_client_base import AIClientBase
router = APIRouter()

# 初始化服务
file_storage = FileStorage()
document_parser = DocumentParser()
format_config_generator = FormatConfigGenerator()

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

# 步骤2：解析源文件 (包含AI分析) - 重构后的架构
@router.get("/parse/source/{file_id}")
async def parse_source_file(file_id: str, preview_length: int = 20, model_config: str = "deepseek_v3"):
    try:
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="文件不存在")
        
        # 1. 提取段落预览信息
        paragraphs_info = document_parser.extract_paragraphs_info(file_path, preview_length)
        
        # 2. 使用指定的模型配置创建AI分析器
        from app.services.ai_analyzer import AIAnalyzer
        ai_analyzer = AIAnalyzer(model_config=model_config)
        
        # 3. 使用AI分析段落
        ai_result = await ai_analyzer.analyze_paragraphs(paragraphs_info["paragraphs"])
        
        # 4. 整合结果
        result = {
            "file_id": file_id,
            "document_info": paragraphs_info["document_info"],
            "analysis_result": ai_result.get("analysis_result", []),
            "success": ai_result.get("success", False),
            "processing_info": {
                "preview_length": preview_length,
                "total_paragraphs": len(paragraphs_info["paragraphs"]),
                "model_config": model_config
            },
            # 添加完整的段落分析JSON（前端显示用）
            "paragraph_analysis_json": {
                "analysis_result": ai_result.get("analysis_result", []),
                "success": ai_result.get("success", False),
                "total_paragraphs": len(paragraphs_info["paragraphs"]),
                "model_used": model_config
            }
        }
        
        # 添加错误信息和批处理信息
        if not ai_result.get("success"):
            result["error"] = ai_result.get("error", "分析失败")
        if "batch_info" in ai_result:
            result["batch_info"] = ai_result["batch_info"]
        if "model_info" in ai_result:
            result["model_info"] = ai_result["model_info"]
        
        return result
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="文件不存在")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"段落分析失败: {str(e)}")

# 步骤2.1：解析源文件 (兼容旧版本)
@router.get("/parse/source/legacy/{file_id}")
async def parse_source_file_legacy(file_id: str):
    try:
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="文件不存在")
        
        # 使用旧版AI分析文档
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
    try:
        # 检查文件类型
        if not file.filename.lower().endswith(('.docx', '.doc')):
            raise HTTPException(status_code=400, detail="只支持.docx和.doc格式的文件")
        
        # 保存文件
        file_info = await file_storage.save_uploaded_file(file, "format")
        
        return {
            "file_id": file_info["file_id"],
            "filename": file_info["original_filename"],
            "message": "格式要求文件上传成功"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"上传失败: {str(e)}")

# 步骤4：AI生成格式配置 - 重构后的架构
@router.post("/generate/format-config/{file_id}")
async def generate_format_config(file_id: str, model_config: str = "deepseek_v3"):
    """
    使用AI从格式要求文档生成format_config.json配置
    """
    try:
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="格式文件不存在")
        
        # 使用指定的模型配置创建格式配置生成器
        from app.services.format_config_generator import FormatConfigGenerator
        config_generator = FormatConfigGenerator(model_config=model_config)
        
        # 使用重构后的格式配置生成器
        result = await config_generator.process_format_document(file_path)
        result["file_id"] = file_id
        
        # 添加模型信息供前端显示
        result["ai_info"] = {
            "model": {
                "model": config_generator.model,
                "config_name": config_generator.model_config_name
            }
        }
        
        return result
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="格式文件不存在")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"格式配置生成失败: {str(e)}")

# 步骤4.2：AI生成格式配置 (兼容旧版本)
@router.post("/generate/format-config/legacy/{file_id}")
async def generate_format_config_legacy(file_id: str):
    """
    使用AI从格式要求文档生成format_config.json配置 (旧版本)
    """
    try:
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="格式文件不存在")
        
        # 使用旧版AI生成格式配置
        result = await format_config_generator.process_format_document(file_path)
        result["file_id"] = file_id
        
        return result
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="格式文件不存在")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"AI生成格式配置失败: {str(e)}")

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

# 模型管理接口
@router.get("/models/available")
async def get_available_models():
    """获取所有可用的模型配置"""
    try:
        models = AIClientBase.get_available_models()
        return {
            "success": True,
            "models": models,
            "count": len(models)
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"获取模型配置失败: {str(e)}")

