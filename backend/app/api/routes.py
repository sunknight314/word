from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List, Dict, Any
from app.services.file_storage import FileStorage
from app.services.format_config_generator import FormatConfigGenerator
from app.services.ai_client_base import AIClientBase
from app.services.document_processor import DocumentProcessor
router = APIRouter()

# 初始化服务
file_storage = FileStorage()
format_config_generator = FormatConfigGenerator()

def generate_format_display_html(display_data: Dict[str, Any]) -> str:
    """生成格式展示的HTML页面"""
    
    # 生成页面设置HTML
    page_settings_html = ""
    if display_data.get("page_settings"):
        settings = display_data["page_settings"]
        items_html = ""
        for item in settings["items"]:
            items_html += f"""
                <div class="setting-item">
                    <span class="setting-label">{item["label"]}:</span>
                    <span class="setting-value">{item["value"]}</span>
                </div>
            """
        
        page_settings_html = f"""
            <div class="section">
                <h2 class="section-title">{settings["title"]}</h2>
                <div class="settings-grid">
                    {items_html}
                </div>
            </div>
        """
    
    # 生成样式设置HTML
    styles_html = ""
    if display_data.get("styles"):
        style_cards = ""
        for style in display_data["styles"]:
            font_items = ""
            for item in style["font"]:
                font_items += f"""
                    <div class="detail-item">
                        <span class="detail-label">{item["label"]}</span>
                        <span class="detail-value">{item["value"]}</span>
                    </div>
                """
            
            para_items = ""
            for item in style["paragraph"]:
                para_items += f"""
                    <div class="detail-item">
                        <span class="detail-label">{item["label"]}</span>
                        <span class="detail-value">{item["value"]}</span>
                    </div>
                """
            
            style_cards += f"""
                <div class="style-card">
                    <div class="style-name">{style["name"]}</div>
                    <div class="style-details">
                        <div class="detail-group">
                            <div class="detail-group-title">字体设置</div>
                            {font_items}
                        </div>
                        <div class="detail-group">
                            <div class="detail-group-title">段落设置</div>
                            {para_items}
                        </div>
                    </div>
                </div>
            """
        
        styles_html = f"""
            <div class="section">
                <h2 class="section-title">样式设置</h2>
                {style_cards}
            </div>
        """
    
    # 生成完整HTML
    return f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>格式要求展示</title>
        <style>
            * {{
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }}
            body {{
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif;
                background-color: #f5f7fa;
                padding: 20px;
                line-height: 1.6;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
            }}
            .header {{
                background: white;
                padding: 25px;
                border-radius: 12px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                margin-bottom: 25px;
                text-align: center;
            }}
            .header h1 {{
                color: #2c3e50;
                font-size: 28px;
                margin-bottom: 10px;
            }}
            .section {{
                background: white;
                padding: 25px;
                border-radius: 12px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                margin-bottom: 20px;
            }}
            .section-title {{
                font-size: 20px;
                color: #34495e;
                margin-bottom: 20px;
                padding-bottom: 10px;
                border-bottom: 2px solid #ecf0f1;
            }}
            .settings-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
                gap: 15px;
            }}
            .setting-item {{
                display: flex;
                justify-content: space-between;
                padding: 12px 16px;
                background: #f8f9fa;
                border-radius: 8px;
            }}
            .setting-label {{
                font-weight: 500;
                color: #495057;
            }}
            .setting-value {{
                color: #2c3e50;
                font-weight: 600;
            }}
            .style-card {{
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 20px;
                margin-bottom: 15px;
            }}
            .style-name {{
                font-size: 18px;
                color: #2c3e50;
                font-weight: 600;
                margin-bottom: 15px;
            }}
            .style-details {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 15px;
            }}
            .detail-group {{
                background: #f8f9fa;
                padding: 15px;
                border-radius: 6px;
            }}
            .detail-group-title {{
                font-weight: 600;
                color: #495057;
                margin-bottom: 10px;
                font-size: 14px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            .detail-item {{
                display: flex;
                justify-content: space-between;
                padding: 5px 0;
                font-size: 14px;
            }}
            .detail-label {{
                color: #6c757d;
            }}
            .detail-value {{
                color: #2c3e50;
                font-weight: 500;
            }}
            .success-message {{
                background: #d4edda;
                color: #155724;
                padding: 15px;
                border-radius: 8px;
                margin-bottom: 20px;
                text-align: center;
            }}
            @media (max-width: 768px) {{
                .style-details {{
                    grid-template-columns: 1fr;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>文档格式要求</h1>
                <p style="color: #7f8c8d;">AI智能识别的格式配置</p>
            </div>
            
            <div class="success-message">
                ✅ 格式识别成功！以下是识别到的格式要求：
            </div>
            
            {page_settings_html}
            {styles_html}
        </div>
    </body>
    </html>
    """

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

