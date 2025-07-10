"""
HTML生成工具类
"""
from typing import Dict, Any


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