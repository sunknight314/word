#!/usr/bin/env python3
"""
示例：如何使用基于配置的文档修改系统
"""

import os
import sys
from document_modifier import DocumentModifier

def modify_test_document():
    """修改测试文档的示例"""
    
    # 源文档路径（使用document_modify_test目录中的文档）
    source_doc = "../document_modify_test/test_document.docx"
    
    # 配置文件路径
    config_file = "modify_config.json"
    
    # 输出文档路径
    output_doc = "../document_modify_test/test_document_modified.docx"
    
    print("📚 基于配置的Word文档修改示例")
    print("="*60)
    
    # 检查文件是否存在
    if not os.path.exists(source_doc):
        print(f"❌ 源文档不存在: {source_doc}")
        return
    
    if not os.path.exists(config_file):
        print(f"❌ 配置文件不存在: {config_file}")
        return
    
    try:
        # 创建文档修改器
        modifier = DocumentModifier(source_doc, config_file)
        
        # 执行修改
        results = modifier.modify_document(output_doc)
        
        # 检查结果
        if results['success']:
            print(f"\n✅ 文档修改成功！")
            print(f"📄 输出文档: {output_doc}")
        else:
            print(f"\n❌ 文档修改失败！")
            for error in results.get('errors', []):
                print(f"  - {error}")
    
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()

def create_custom_config():
    """创建自定义配置的示例"""
    
    import json
    
    custom_config = {
        "modify_mode": "merge",
        
        "document_info": {
            "backup_original": True
        },
        
        "styles": {
            "Heading 1": {
                "font": {
                    "name": "黑体",
                    "size": "18pt",
                    "bold": True,
                    "color": "#FF0000"
                }
            },
            "Normal": {
                "font": {
                    "name": "仿宋",
                    "size": "14pt"
                },
                "paragraph": {
                    "line_spacing": 2.0
                }
            }
        },
        
        "content_modifications": {
            "replacements": [
                {
                    "find": "项目",
                    "replace": "工程",
                    "options": {
                        "case_sensitive": False
                    }
                }
            ]
        }
    }
    
    # 保存自定义配置
    custom_config_path = "custom_modify_config.json"
    with open(custom_config_path, 'w', encoding='utf-8') as f:
        json.dump(custom_config, f, ensure_ascii=False, indent=2)
    
    print(f"✅ 已创建自定义配置文件: {custom_config_path}")
    return custom_config_path

def batch_modify_documents():
    """批量修改文档的示例"""
    
    print("\n📚 批量修改文档示例")
    print("="*60)
    
    # 要修改的文档列表
    documents = [
        "../document_modify_test/test_document.docx",
        # 可以添加更多文档
    ]
    
    config_file = "modify_config.json"
    
    for i, doc_path in enumerate(documents, 1):
        if not os.path.exists(doc_path):
            print(f"❌ 文档不存在: {doc_path}")
            continue
        
        print(f"\n处理第 {i}/{len(documents)} 个文档: {doc_path}")
        
        try:
            # 生成输出文件名
            base_name = os.path.splitext(os.path.basename(doc_path))[0]
            output_path = os.path.join(
                os.path.dirname(doc_path),
                f"{base_name}_batch_modified.docx"
            )
            
            # 创建修改器并执行
            modifier = DocumentModifier(doc_path, config_file)
            results = modifier.modify_document(output_path)
            
            if results['success']:
                print(f"  ✅ 成功 -> {output_path}")
            else:
                print(f"  ❌ 失败")
        
        except Exception as e:
            print(f"  ❌ 错误: {e}")

def main():
    """主函数"""
    
    print("🎯 Word文档配置化修改系统使用示例")
    print("="*60)
    print("1. 使用默认配置修改单个文档")
    print("2. 创建并使用自定义配置")
    print("3. 批量修改多个文档")
    print("4. 退出")
    
    choice = input("\n请选择操作 (1-4): ").strip()
    
    if choice == "1":
        modify_test_document()
    
    elif choice == "2":
        # 创建自定义配置
        custom_config_path = create_custom_config()
        
        # 使用自定义配置修改文档
        source_doc = "../document_modify_test/test_document.docx"
        output_doc = "../document_modify_test/test_document_custom_modified.docx"
        
        if os.path.exists(source_doc):
            modifier = DocumentModifier(source_doc, custom_config_path)
            results = modifier.modify_document(output_doc)
            
            if results['success']:
                print(f"\n✅ 使用自定义配置修改成功！")
                print(f"📄 输出文档: {output_doc}")
        
        # 清理自定义配置文件
        os.remove(custom_config_path)
    
    elif choice == "3":
        batch_modify_documents()
    
    else:
        print("退出程序")

if __name__ == "__main__":
    main()