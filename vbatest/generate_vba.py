import win32com.client
import os

def create_docm_with_macro(output_path, macro_code=None):
    try:
        # 启动Word应用
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 隐藏界面
        
        # 创建新文档
        doc = word.Documents.Add()
        
        # 添加内容（可选）
        doc.Content.Text = "这是由程序生成的启用宏的文档。\n"
        
        # 添加VBA宏（若需嵌入）
        if macro_code:
            # 导入宏模块到文档
            vba_project = doc.VBProject
            vba_module = vba_project.VBComponents.Add(1)  # 1表示标准模块
            vba_module.CodeModule.AddFromString(macro_code)  # 写入宏代码
        
        # 保存为.docm格式
        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)  # 17对应.docm
        doc.Close()
        word.Quit()
        print(f"文档已保存：{output_path}")
    
    except Exception as e:
        print(f"错误：{e}")
        if 'doc' in locals(): doc.Close(SaveChanges=False)
        if 'word' in locals(): word.Quit()

# 示例调用
macro_code = """
Sub AutoOpen()
    MsgBox "文档已打开！", vbInformation
End Sub
"""
create_docm_with_macro("output.docm", macro_code)