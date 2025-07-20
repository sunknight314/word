from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List, Dict, Any
from pydantic import BaseModel
from app.services.file_storage import FileStorage
from app.services.format_config_generator import FormatConfigGenerator
from app.services.ai_client_base import AIClientBase
from app.services.document_processor import DocumentProcessor
from app.services.analysis_result_storage import AnalysisResultStorage
from app.utils.html_generator import generate_format_display_html

router = APIRouter()

# 初始化服务
file_storage = FileStorage()
format_config_generator = FormatConfigGenerator()
analysis_storage = AnalysisResultStorage()

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
        
        document_processor = DocumentProcessor()

        # 1. 提取段落预览信息
        paragraphs_info = document_processor.extract_paragraphs_preview(file_path, preview_length)
        
        # 2. 使用指定的模型配置创建AI分析器
        from app.services.ai_analyzer import AIAnalyzer
        ai_analyzer = AIAnalyzer(model_config=model_config)
        
        # 3. 过滤出非空段落发送给AI分析
        # 只发送文本段落给AI，不发送图片、公式等空段落
        text_paragraphs = [p for p in paragraphs_info["paragraphs"] if not p.get("_is_empty", False)]
        
        # 记录完整的段落信息到日志
        import json
        print(f"[DEBUG] 文档解析 - 文件ID: {file_id}")
        print(f"[DEBUG] 总段落数: {len(paragraphs_info['paragraphs'])}")
        print(f"[DEBUG] 文本段落数: {len(text_paragraphs)}")
        print(f"[DEBUG] 段落预览信息:")
        print(json.dumps(paragraphs_info, ensure_ascii=False, indent=2))
        print("=" * 80)
        
        # 使用AI分析文本段落
        ai_result = await ai_analyzer.analyze_paragraphs(text_paragraphs)
        
        # 4. 合并本地检测和AI分析结果
        merged_results = document_processor.merge_analysis_results(
            paragraphs_info["paragraphs"], 
            ai_result.get("analysis_result", []) if ai_result.get("success") else []
        )
        
        # 5. 整合结果
        result = {
            "file_id": file_id,
            "document_info": paragraphs_info["document_info"],
            "analysis_result": merged_results,
            "paragraphs": paragraphs_info["paragraphs"],  # 添加段落信息供前端显示
            "success": ai_result.get("success", False),
            "processing_info": {
                "preview_length": preview_length,
                "total_paragraphs": len(paragraphs_info["paragraphs"]),
                "model_config": model_config
            },
            # 添加完整的段落分析JSON（前端显示用）
            "paragraph_analysis_json": {
                "analysis_result": merged_results,
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
        
        # 记录最终分析结果到日志
        print(f"[DEBUG] AI分析结果:")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        print("=" * 80)
        
        # 保存段落分析结果
        if result.get("success"):
            saved_path = analysis_storage.save_paragraph_analysis(file_id, result)
            result["saved_path"] = saved_path
            print(f"[INFO] 段落分析结果已保存到: {saved_path}")
        
        return result
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="文件不存在")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"段落分析失败: {str(e)}")

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
        
        # 添加格式化的展示数据
        if result.get("success") and result.get("format_config"):
            from app.utils.format_display_helper import FormatDisplayHelper
            display_helper = FormatDisplayHelper()
            result["display_data"] = display_helper.format_for_display(result["format_config"])
            
            # 保存格式配置分析结果
            saved_path = analysis_storage.save_format_config(
                file_id, 
                result["format_config"],
                result.get("ai_info")
            )
            result["saved_path"] = saved_path
            print(f"[INFO] 格式配置分析结果已保存到: {saved_path}")
            
            # 添加查看链接
            result["view_url"] = f"/api/format-config/view/{file_id}"
            result["message"] = "格式识别成功！请点击查看链接查看友好的格式展示。"
        
        return result
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="格式文件不存在")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"格式配置生成失败: {str(e)}")

# 获取格式配置的HTML展示页面
@router.get("/format-config/view/{file_id}")
async def view_format_config(file_id: str):
    """
    返回格式配置的HTML展示页面
    """
    from fastapi.responses import HTMLResponse
    
    try:
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="格式文件不存在")
        
        # 使用默认模型配置创建格式配置生成器
        from app.services.format_config_generator import FormatConfigGenerator
        config_generator = FormatConfigGenerator(model_config="deepseek_v3")
        
        # 获取已处理的结果（实际应用中应该从缓存或数据库获取）
        result = await config_generator.process_format_document(file_path)
        
        # 生成展示数据
        if result.get("success") and result.get("format_config"):
            from app.utils.format_display_helper import FormatDisplayHelper
            display_helper = FormatDisplayHelper()
            display_data = display_helper.format_for_display(result["format_config"])
            
            # 生成HTML页面
            html_content = generate_format_display_html(display_data)
            return HTMLResponse(content=html_content, status_code=200)
        else:
            return HTMLResponse(
                content="<h1>格式识别失败</h1><p>无法生成格式展示</p>", 
                status_code=400
            )
            
    except Exception as e:
        return HTMLResponse(
            content=f"<h1>错误</h1><p>{str(e)}</p>", 
            status_code=500
        )


# 步骤5：执行格式转换
class ProcessRequest(BaseModel):
    source_id: str
    format_id: str

@router.post("/process")
async def process_documents(request: ProcessRequest):
    """
    执行文档格式转换
    
    Args:
        source_id: 源文档ID
        format_id: 格式文档ID
    """
    import logging
    import uuid
    
    # 设置日志
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    
    try:
        logger.info("🚀 开始执行文档格式转换")
        
        # 获取源文件路径
        source_path = file_storage.get_file_path(request.source_id)
        if not source_path:
            raise HTTPException(status_code=404, detail="源文件不存在")
        
        format_path = file_storage.get_file_path(request.format_id)
        if not format_path:
            raise HTTPException(status_code=404, detail="格式文件不存在")
        
        logger.info(f"📄 源文件: {request.source_id} -> {source_path.split('/')[-1]}")
        logger.info(f"📋 格式文件: {request.format_id} -> {format_path.split('/')[-1]}")
        
        # 初始化格式转换服务
        from app.services.document_formatter_v2 import DocumentFormatterV2
        formatter = DocumentFormatterV2()
        
        # 执行格式转换
        result = formatter.format_document(source_path, request.source_id, request.format_id)
        
        if result["success"]:
            logger.info("✅ 格式转换成功完成!")
            
            # 生成任务ID
            task_id = str(uuid.uuid4())
            
            # 保存结果信息
            result["task_id"] = task_id
            result["source_id"] = request.source_id
            result["format_id"] = request.format_id
            
            # 展示处理结果
            logger.info(f"📊 输出文件: {result.get('output_path', 'N/A')}")
            if result.get("report"):
                report = result["report"]
                for step, status in report.items():
                    if isinstance(status, dict):
                        success = status.get("success", False)
                        status_icon = "✅" if success else "❌"
                        logger.info(f"   {status_icon} {step}")
            
            return {
                "task_id": task_id,
                "status": "completed",
                "message": "格式转换已完成",
                "output_path": result["output_path"],
                "report": result["report"]
            }
        else:
            logger.error(f"❌ 格式转换失败: {result.get('error', '未知错误')}")
            raise HTTPException(status_code=500, detail=result.get("error", "格式转换失败"))
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ 格式转换异常: {str(e)}")
        raise HTTPException(status_code=500, detail=f"格式转换失败: {str(e)}")

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
@router.get("/download/{file_path:path}")
async def download_result(file_path: str):
    """
    下载处理后的文档
    
    Args:
        file_path: 文件路径
    """
    from fastapi.responses import FileResponse
    import os
    
    try:
        # 安全检查：确保文件路径是绝对路径且存在
        if not os.path.isabs(file_path):
            raise HTTPException(status_code=400, detail="文件路径必须是绝对路径")
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="文件不存在")
        
        # 获取文件名
        filename = os.path.basename(file_path)
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"下载失败: {str(e)}")

# 更新格式配置
@router.put("/format-config/{file_id}")
async def update_format_config(file_id: str, format_config: Dict[str, Any]):
    """
    更新格式配置
    
    Args:
        file_id: 文件ID
        format_config: 新的格式配置
    """
    try:
        # 验证格式配置
        from app.utils.format_config_validator import FormatConfigValidator
        is_valid, errors, warnings = FormatConfigValidator.validate_format_config(format_config)
        
        if not is_valid:
            raise HTTPException(
                status_code=400, 
                detail={
                    "message": "格式配置验证失败",
                    "errors": errors,
                    "warnings": warnings
                }
            )
        
        # 获取文件路径
        file_path = file_storage.get_file_path(file_id)
        if not file_path:
            raise HTTPException(status_code=404, detail="文件不存在")
        
        # 保存格式配置（这里简单地保存到同目录下的json文件）
        import json
        import os
        config_path = os.path.splitext(file_path)[0] + "_format_config.json"
        
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(format_config, f, ensure_ascii=False, indent=2)
        
        return {
            "success": True,
            "message": "格式配置已更新",
            "file_id": file_id,
            "config_path": config_path,
            "warnings": warnings if warnings else None
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"更新格式配置失败: {str(e)}")

# 分析结果查询接口
@router.get("/analysis/{file_id}")
async def get_analysis_results(file_id: str):
    """获取文件的所有分析结果"""
    try:
        results = analysis_storage.get_all_analysis_for_file(file_id)
        return {
            "success": True,
            "file_id": file_id,
            "paragraph_analysis": results["paragraph_analysis"],
            "format_config": results["format_config"]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"获取分析结果失败: {str(e)}")

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

