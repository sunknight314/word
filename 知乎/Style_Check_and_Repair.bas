Attribute VB_Name = "Style_Check_and_Repair"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'本模块用于检查和恢复样式

Function Style_Check_All(arr_style_check() As Boolean) As Boolean
'''''模板全样式筛查'''''
    Dim style_test As style
    Style_Check_All = True
    For i = 1 To Config.nLength
        On Error Resume Next
        
        Set style_test = ActiveDocument.Styles(Config.dict.Item(i))
        If Err <> 0 Then
            arr_style_check(i - 1) = False
            Style_Check_All = False
        Else
            arr_style_check(i - 1) = True
        End If
        
        Err.Clear
        
        On Error GoTo 0
    Next
End Function

Sub Style_Create(arr_style_check() As Boolean)
'重置缺失的样式
'
    Dim template_temp As Template
    Dim templatePath As String
    Dim templateDoc As Document
    Dim currentDoc As Document
    Dim savedRange As Range
    Dim style_name As String

'    Application.ScreenUpdating = False '关闭屏幕刷新

    Dim newListTemplate As ListTemplate

'临时保存Selection.Range指定的段落范围
    Set savedRange = Selection.Range

'设置模板文件的路径
    templatePath = "Error" '初始化
    For Each template_temp In Templates
        If template_temp.Name = "《Word自动排版工具》V3.0.dotm" Then
            templatePath = template_temp.FullName
            Exit For
        End If
    Next

    If templatePath = "Error" Then
        Exit_Flag = True
        MsgBox "未能找到自动排版工具!" & vbCrLf & "请检查模板文件名称等！"
        Exit Sub
    End If
'    templatePath = "C:\Program Files\Microsoft Office\root\Office16\STARTUP\《Word自动排版工具》V3.0.dotm"

'获取当前文档对象
    Set currentDoc = ActiveDocument

'打开模板文件
    Set templateDoc = Documents.Open(templatePath, ReadOnly:=True, Visible:=False)

'修复缺失的样式
    For i = 1 To Config.nLength
        If arr_style_check(i - 1) = False Then
            Select Case i
                Case Config.normal_PXL
                    style_name = Config.dict.Item(Config.normal_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.normal_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font '仅靠Application.OrganizerCopy无法稳定加载模板的文字格式
                Case Config.level1_PXL
                    style_name = Config.dict.Item(Config.level1_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.level1_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.level2_PXL
                    style_name = Config.dict.Item(Config.level2_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.level2_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.level3_PXL
                    style_name = Config.dict.Item(Config.level3_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.level3_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.level4_PXL
                    style_name = Config.dict.Item(Config.level4_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.level4_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.tab_title_PXL
                    style_name = Config.dict.Item(Config.tab_title_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.tab_title_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.fig_title_PXL
                    style_name = Config.dict.Item(Config.fig_title_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.fig_title_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.tab_text_PXL
                    style_name = Config.dict.Item(Config.tab_text_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.tab_text_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.equ_PXL
                    style_name = Config.dict.Item(Config.equ_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
                    Call Style_Library.equ_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.ref_PXL
                    style_name = Config.dict.Item(Config.ref_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.ref_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.appx_level1_PXL
                    style_name = Config.dict.Item(Config.appx_level1_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.appx_level1_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.appx_level2_PXL
                    style_name = Config.dict.Item(Config.appx_level2_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.appx_level2_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.appx_level3_PXL
                    style_name = Config.dict.Item(Config.appx_level3_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.appx_level3_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.appx_level4_PXL
                    style_name = Config.dict.Item(Config.appx_level4_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.appx_level4_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.appx_tab_title_PXL
                    style_name = Config.dict.Item(Config.appx_tab_title_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.appx_tab_title_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.appx_fig_title_PXL
                    style_name = Config.dict.Item(Config.appx_fig_title_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.appx_fig_title_PXL
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font
                Case Config.tab_PXL
                    style_name = Config.dict.Item(Config.tab_PXL)
                    Application.OrganizerCopy Source:=templatePath, _
                        Destination:=currentDoc, _
                        Name:=style_name, _
                        Object:=wdOrganizerObjectStyles
'                    Call Style_Library.tab_PXL
                Case Else
                '无效
            End Select
        End If
    Next

' 关闭模板文件，不保存更改
    templateDoc.Close SaveChanges:=wdDoNotSaveChanges
    
'检测正文多级列表“ListN_PXL”是否存在，不存在则创建
    On Error Resume Next
        Set newListTemplate = ActiveDocument.ListTemplates("ListN_PXL")

        If Err <> 0 Then
            Call Style_Library.Multi_level_list '创建多级列表关联正文的各种样式
        End If
    On Error GoTo 0

'检测附录多级列表“ListA_PXL”是否存在，不存在则创建
    On Error Resume Next
        Set newListTemplate = ActiveDocument.ListTemplates("ListA_PXL")

        If Err <> 0 Then
            Call Style_Library.Multi_level_list_APPX '创建多级列表关联附录的各种样式
        End If
    On Error GoTo 0

'检测参考文献列表“ListR_PXL”是否存在，不存在则创建
    On Error Resume Next
        Set newListTemplate = ActiveDocument.ListTemplates("ListR_PXL")
        If Err <> 0 Then
            Call Style_Library.ref_list '创建单级列表关联参考文献样式
        End If
    On Error GoTo 0

    '关联样式和列表
    Set newListTemplate = currentDoc.ListTemplates("ListN_PXL")
    With newListTemplate
        .ListLevels(1).LinkedStyle = Config.dict.Item(Config.level1_PXL) ' "标题1_PXL"
        .ListLevels(2).LinkedStyle = Config.dict.Item(Config.level2_PXL) ' "标题2_PXL"
        .ListLevels(3).LinkedStyle = Config.dict.Item(Config.level3_PXL) ' "标题3_PXL"
        .ListLevels(4).LinkedStyle = Config.dict.Item(Config.level4_PXL) ' "标题4_PXL"
        .ListLevels(8).LinkedStyle = Config.dict.Item(Config.fig_title_PXL) ' "图题_PXL"
        .ListLevels(9).LinkedStyle = Config.dict.Item(Config.tab_title_PXL) ' "表题_PXL"
    End With

    '关联样式和列表
    Set newListTemplate = currentDoc.ListTemplates("ListA_PXL")
    With newListTemplate
        .ListLevels(1).LinkedStyle = Config.dict.Item(Config.appx_level1_PXL) ' "附录1_PXL"
        .ListLevels(2).LinkedStyle = Config.dict.Item(Config.appx_level2_PXL) ' "附录2_PXL"
        .ListLevels(3).LinkedStyle = Config.dict.Item(Config.appx_level3_PXL) ' "附录3_PXL"
        .ListLevels(4).LinkedStyle = Config.dict.Item(Config.appx_level4_PXL) ' "附录4_PXL"
        .ListLevels(7).LinkedStyle = Config.dict.Item(Config.appx_fig_title_PXL) ' "附录图_PXL"
        .ListLevels(8).LinkedStyle = Config.dict.Item(Config.appx_tab_title_PXL) ' "附录表_PXL"
    End With

    '关联样式和列表
    Set newListTemplate = currentDoc.ListTemplates("ListR_PXL")
    newListTemplate.ListLevels(1).LinkedStyle = Config.dict.Item(Config.ref_PXL) ' "参考文献_PXL"

'恢复原本选择的范围
    savedRange.Select
    
'打开屏幕刷新
'    Application.ScreenUpdating = True
End Sub



