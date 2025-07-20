Attribute VB_Name = "Semi_Auto_Typesetting"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'本模块用于提供各种半自动排版功能

Sub Text_Style(control As IRibbonControl)
'
' 正文样式 宏
' 将所选段落设置为“正文_PXL”样式
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.normal_PXL)

    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title1(control As IRibbonControl)
Attribute Title1.VB_Description = "将所选段落设置为样式“标题1”。"
Attribute Title1.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.标题1"
'
' 标题1 宏
' 将所选段落设置为样式“标题1_PXL”
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level1_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title2(control As IRibbonControl)
Attribute Title2.VB_Description = "将所选段落设置为样式“标题2”。"
Attribute Title2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.宏1"
'
' 标题2 宏
' 将所选段落设置为样式“标题2_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level2_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title3(control As IRibbonControl)
Attribute Title3.VB_Description = "将所选段落设置为样式“标题3”。"
Attribute Title3.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.标题3"
'
' 标题3 宏
' 将所选段落设置为样式“标题3_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level3_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title4(control As IRibbonControl)
Attribute Title4.VB_Description = "将所选段落设置为样式“标题4”。"
Attribute Title4.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.标题4"
'
' 标题4 宏
' 将所选段落设置为样式“标题4_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level4_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Table_Title(control As IRibbonControl)
Attribute Table_Title.VB_Description = "将所选段落设置为样式“表题”。"
Attribute Table_Title.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.表题"
'
' 表题 宏
' 将所选段落设置为样式“表题_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.tab_title_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Picture_Title(control As IRibbonControl)
Attribute Picture_Title.VB_Description = "将所选段落设置为样式“图题”。"
Attribute Picture_Title.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.图题"
'
' 图题 宏
' 将所选段落设置为样式“图题_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.fig_title_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Table_Text(control As IRibbonControl)
Attribute Table_Text.VB_Description = "将所选段落设置为样式“表文字”。"
Attribute Table_Text.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.表文字"
'
' 表文字 宏
' 将所选段落设置为样式“表文字_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.tab_text_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Table_Style(control As IRibbonControl)
Attribute Table_Style.VB_Description = "将所选表格设置为样式“PXL”。"
Attribute Table_Style.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.表"
'
' 表 宏
' 将所选表格设置为样式“tab_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.tab_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新表格样式'''''

    '''''检查所选范围是否存在表格'''''
    Dim tb As Table
    On Error Resume Next '启用错误处理
    Set tb = Selection.Tables(1)
    On Error GoTo 0 ' 关闭当前错误处理设置
    
    If Not tb Is Nothing Then
        For Each tb In Selection.Tables
            With tb
                .style = style_name
            End With
        Next
    End If

End Sub
Sub Appendix1(control As IRibbonControl)
Attribute Appendix1.VB_Description = "将所选段落设置为样式“附标题1”。"
Attribute Appendix1.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.附·标题1"
'
' 附录1 宏
' 将所选段落设置为样式“附录1_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level1_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)

    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix2(control As IRibbonControl)
Attribute Appendix2.VB_Description = "将所选段落设置为样式“附录2”。"
Attribute Appendix2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.附录2"
'
' 附录2 宏
' 将所选段落设置为样式“附录2_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level2_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix3(control As IRibbonControl)
Attribute Appendix3.VB_Description = "将所选段落设置为样式“附录3”。"
Attribute Appendix3.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.附录"
'
' 附录3 宏
' 将所选段落设置为样式“附录3_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level3_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix4(control As IRibbonControl)
'
' 附录4 宏
' 将所选段落设置为样式“附录4_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level4_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)

    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix_Table(control As IRibbonControl)
Attribute Appendix_Table.VB_Description = "将所选段落设置为样式“附录表”。"
Attribute Appendix_Table.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.附录表"
'
' 附录表 宏
' 将所选段落设置为样式“附录表_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_tab_title_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix_Picture(control As IRibbonControl)
Attribute Appendix_Picture.VB_Description = "将所选段落设置为样式“附录图”。"
Attribute Appendix_Picture.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.附录图"
'
' 附录图 宏
' 将所选段落设置为样式“附录图_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_fig_title_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub References_Style(control As IRibbonControl)
Attribute References_Style.VB_Description = "将所选段落设置为样式“参考文献”。"
Attribute References_Style.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.参考文献"
'
' 参考文献 宏
' 将所选段落设置为样式“参考文献_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.ref_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Formula_Style(control As IRibbonControl)
Attribute Formula_Style.VB_Description = "将所选段落设置为样式“公式样式”。"
Attribute Formula_Style.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.公式样式"
'
' 公式样式 宏
' 将所选段落设置为样式“公式_PXL”。
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.equ_PXL)
    
    '''''样式检查'''''
    Exit_Flag = False '为True则退出
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    '''''更新段落样式'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Formula_Number(control As IRibbonControl)
Attribute Formula_Number.VB_Description = "插入公示编号"
Attribute Formula_Number.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.公式编号"
'
' 公式编号 宏
' 插入公示编号
'
    On Error Resume Next '启用错误处理
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "LISTNUM  ListN_PXL \l 7 ", PreserveFormatting:=False
    On Error GoTo 0 ' 关闭当前错误处理设置
End Sub
Sub Style_Check(str As String)
'''''样式检查'''''
    Dim style_test As style
    Dim arr_style_check() As Boolean
    ReDim arr_style_check(Config.nLength - 1) '当前总共有17种样式
    Dim style_check_result As Boolean
    On Error Resume Next
    Set style_test = ActiveDocument.Styles(str)
    If Err <> 0 Then '有样式不存在
        style_check_result = Style_Check_and_Repair.Style_Check_All(arr_style_check)  '检查所有样式
        Call Style_Check_and_Repair.Style_Create(arr_style_check) '恢复缺失的样式
        Err.Clear
    End If
    On Error GoTo 0
End Sub


