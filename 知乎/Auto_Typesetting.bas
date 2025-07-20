Attribute VB_Name = "Auto_Typesetting"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'本模块用于提供一键自动排版功能
Public Config As New Configuration '全局参数设置
Public Exit_Flag As Boolean

Sub Auto_Typesetting_by_Selections_Normal(control As IRibbonControl)
'按选择排版
    Dim oDoc As Document
    Dim oRange As Range
    Dim savedRange As Range '临时保存Selection.Range指定的段落范围
    
'    Dim sub_exit As Boolean
'    sub_exit = False
    Set oDoc = ActiveDocument
    Set oRange = Selection.Range
    Set savedRange = Selection.Range
    Exit_Flag = False '为True则退出
    
    Application.ScreenUpdating = False '关闭屏幕刷新

    Call Auto_Typesetting(oRange, oDoc, 0)

    Application.ScreenUpdating = True '打开屏幕刷新
    savedRange.Select
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    
    MsgBox ("完成！")
End Sub

Sub Auto_Typesetting_by_Selections_Appendix(control As IRibbonControl)
'按选择排版
    Dim oDoc As Document
    Dim oRange As Range
    Dim savedRange As Range '临时保存Selection.Range指定的段落范围
    
'    Dim sub_exit As Boolean
'    sub_exit = False
    Set oDoc = ActiveDocument
    Set oRange = Selection.Range
    Set savedRange = Selection.Range
    Exit_Flag = False '为True则退出
    
    Application.ScreenUpdating = False '关闭屏幕刷新
   
    Call Auto_Typesetting(oRange, oDoc, 1)
    
    Application.ScreenUpdating = True '打开屏幕刷新
    
    savedRange.Select
    
    If Exit_Flag = True Then '为True则退出
        Exit Sub
    End If
    
    MsgBox ("完成！")
End Sub

'Sub Auto_Typesetting_by_Pages(control As IRibbonControl)
''按页数排版
'    Dim pg_fr As Integer
'    Dim pg_to As Integer
'    Dim oDoc As Document
'    Dim oRange As Range
'    Dim savedRange As Range '临时保存Selection.Range指定的段落范围
'    Dim sub_exit As Boolean
'    sub_exit = False
'    Set oDoc = ActiveDocument
'
'    Application.ScreenUpdating = False '关闭屏幕刷新
'
'    '输入页码检测
'    Call input_management(pg_fr, pg_to, oDoc, sub_exit)
'    If sub_exit = True Then
'        Exit Sub
'    End If
'
'    '捕获段落范围
'    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pg_fr
'    Set oRange = Selection.Range
'    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pg_to
'    oRange.End = Selection.Bookmarks("\Page").Range.End
'    Set savedRange = oRange
'    '自动排版
'    Call Auto_Typesetting(oRange, oDoc)
'
'    Application.ScreenUpdating = True '打开屏幕刷新
'
'    savedRange.Select
'    MsgBox ("完成！")
'End Sub

'Sub input_management(pg_fr As Integer, pg_to As Integer, oDoc As Document, sub_exit As Boolean)
'    Dim input_str As String
'    Dim pg_count As Integer
'    pg_count = oDoc.ComputeStatistics(wdStatisticPages)
'
'    Do
'        input_str = InputBox("请输入排版的  起始  页码：", Title:="请输入排版的  起始  页码")
'
'        If input_str = "" Then '取消或关闭窗口
'            sub_exit = True
'            Exit Sub
'        End If
'
'        If Len(input_str) = 0 Then
'            answer = MsgBox("未输入任何值！", vbOKOnly, "请输入排版的  起始  页码")
'        ElseIf IsNumeric(input_str) = False Then
'            answer = MsgBox("输入不为数字！请输入整数！", vbOKOnly, "请输入排版的  起始  页码")
'        ElseIf IsNumeric(input_str) = True Then
'            If CInt(input_str) = input_str Then
'                pg_fr = CInt(input_str)
'                If pg_fr > pg_count Or pg_fr <= 0 Then
'                    answer = MsgBox("输入的页码超过文档范围！", vbOKOnly, "请输入排版的  起始  页码")
'                Else
'                    Exit Do
'                End If
'            Else
'                answer = MsgBox("输入不为整数！请输入整数！", vbOKOnly, "请输入排版的  起始  页码")
'            End If
'        End If
'    Loop Until False
'
'    Do
'        input_str = InputBox("请输入排版的  终止  页码：", Title:="请输入排版的  终止  页码")
'
'        If input_str = "" Then '取消或关闭窗口
'            sub_exit = True
'            Exit Sub
'        End If
'
'        If Len(input_str) = 0 Then
'            answer = MsgBox("未输入任何值！", vbOKOnly, "请输入排版的  终止  页码")
'        ElseIf IsNumeric(input_str) = False Then
'            answer = MsgBox("输入不为数字！请输入整数！", vbOKOnly, "请输入排版的  终止  页码")
'        ElseIf IsNumeric(input_str) = True Then
'            If CInt(input_str) = input_str Then
'                pg_to = CInt(input_str)
'                If pg_to > pg_count Then
'                    answer = MsgBox("输入的页码超过文档范围！", vbOKOnly, "请输入排版的  终止  页码")
'                ElseIf pg_fr > pg_to Then
'                    answer = MsgBox("输入的页码小于起始页码！", vbOKOnly, "请输入排版的  终止  页码")
'                Else
'                    Exit Do
'                End If
'            Else
'                answer = MsgBox("输入不为整数！请输入整数！", vbOKOnly, "请输入排版的  终止  页码")
'            End If
'        End If
'    Loop Until False
'End Sub

Sub Auto_Typesetting(oRange As Range, oDoc As Document, model As Integer)
    Dim pg As Paragraph
    Dim pic_reg As Object
    Dim content_reg As Object
    Dim equation_reg As Object
    Dim Next_Exit As Boolean
    Next_Exit = False

    Set pic_reg = CreateObject("vbscript.regexp")
    Set content_reg = CreateObject("vbscript.regexp")
    Set equation_reg = CreateObject("vbscript.regexp")

    With pic_reg
        .Global = False
        .Pattern = "[^\s/]"
    End With
    
    With content_reg
        .Global = False
        .Pattern = "目[\s]*录"
    End With
    
    With equation_reg
        .Global = False
        .Pattern = "[a-zA-Z\u4e00-\u9fa5]"
    End With
    
    Dim level1 As String '大纲1级段落样式
    Dim level2 As String '大纲2级段落样式
    Dim level3 As String '大纲3级段落样式
    Dim level4 As String '大纲4级段落样式
    Dim tab_title As String '表标题
    Dim fig_title As String '图表题
    
    If model = 0 Then
        level1 = Config.dict.Item(Config.level1_PXL) '"标题1_PXL"
        level2 = Config.dict.Item(Config.level2_PXL) '"标题2_PXL"
        level3 = Config.dict.Item(Config.level3_PXL) '"标题3_PXL"
        level4 = Config.dict.Item(Config.level4_PXL) '"标题4_PXL"
        tab_title = Config.dict.Item(Config.tab_title_PXL) '"表题_PXL"
        fig_title = Config.dict.Item(Config.fig_title_PXL) '"图题_PXL"
    ElseIf model = 1 Then
        level1 = Config.dict.Item(Config.appx_level1_PXL) '"附录1_PXL"
        level2 = Config.dict.Item(Config.appx_level2_PXL) '"附录2_PXL"
        level3 = Config.dict.Item(Config.appx_level3_PXL) '"附录3_PXL"
        level4 = Config.dict.Item(Config.appx_level4_PXL) '"附录4_PXL"
        tab_title = Config.dict.Item(Config.appx_tab_title_PXL) '"附录表_PXL"
        fig_title = Config.dict.Item(Config.appx_fig_title_PXL) '"附录图_PXL"
    Else
        Exit_Flag = True
        MsgBox ("自动排版模式错误！")
        Exit Sub
    End If
'''''自定义样式检查'''''
    Dim style_check_result As Boolean
    Dim arr_style_check() As Boolean
    ReDim arr_style_check(Config.nLength - 1) '当前总共有17种样式
    style_check_result = True
    style_check_result = Style_Check_and_Repair.Style_Check_All(arr_style_check)  '检查所有样式
    If style_check_result <> True Then
        Call Style_Check_and_Repair.Style_Create(arr_style_check) '仅恢复缺失的样式
    End If
    
    If Exit_Flag = True Then
        Exit Sub
    End If
           
'''''设置1~4级标题、图标题、表标题、图片'''''
    '初始化，避免复制内容时，选择的位置已有样式（比如 “正文_PXL”），后续跳过处理的段落会变成已有样式（比如“正文_PXL”）
'    With oRange.Paragraphs
'        .CharacterUnitFirstLineIndent = 0 '首行缩进，字符
'        .CharacterUnitLeftIndent = 0 '左缩进，字符
'        .CharacterUnitRightIndent = 0 '右缩进，字符
'        .FirstLineIndent = 0 '首行缩进，磅值
'        .LeftIndent = 0 '左缩进，磅值
'        .RightIndent = 0 '右缩进，磅值
'        .Alignment = wdAlignParagraphCenter '两端对齐
'    End With
    
    For Each pg In oRange.Paragraphs
        With pg
            
            If Next_Exit = True Then '由于图片后续内容作为标题，连续处理了2段，故跳过第2段
                Next_Exit = False
                GoTo NextIter1
            End If
            
            '表格部分段落先不处理
            If .Range.Information(wdWithInTable) Then
                GoTo NextIter1
            End If
            
'            跳过“公式_PXL”样式段落
            If .style.NameLocal = Config.dict.Item(Config.equ_PXL) Then
                GoTo NextIter1
            End If
            
            'word内置公式对象且无中英文，不处理
            If .Range.OMaths.Count <> 0 And equation_reg.Execute(.Range.Text).Count = 0 Then
                GoTo NextIter1
            End If
            
            ' 跳过目录部分
            If .Range.Fields.Count > 0 Then
'                If .Range.Fields(1).Type = wdFieldTOC Then
'                    GoTo NextIter1
'                ElseIf equation_reg.Execute(.Range.Text).Count = 0 Then
'                    GoTo NextIter1
'                End If
                If .Range.Fields(1).Type = wdFieldTOC Then
                    GoTo NextIter1
                ElseIf .Range.Fields(1).Type = wdFieldEmbed Then
                    If .Range.Fields(1).OLEFormat.ClassType Like "Visio.Drawing.15" Then
                        GoTo Picture_layout
                    ElseIf equation_reg.Execute(.Range.Text).Count = 0 _
                        And .Range.Fields(1).OLEFormat.ClassType Like "Equation.DSMT4" Then 'mathtype公式
                        GoTo NextIter1
                    End If
                
                
'                ElseIf equation_reg.Execute(.Range.Text).Count = 0 Then
'                    If .Range.Fields(1).OLEFormat.ClassType = "Equation.DSMT4" Then 'mathtype公式
'
'
'
'                    End If
'                    GoTo NextIter1
                End If
            End If
            
            '目录标题不处理
            If content_reg.Execute(.Range.Text).Count <> 0 Then
                GoTo NextIter1
            End If
                
            '将非嵌入式图片转换为嵌入式图片
            For Each pg_shape In .Range.ShapeRange
                If pg_shape.Type = msoPicture Then
                    pg_shape.ConvertToInlineShape
                End If
            Next

            '段落含有嵌入式图片且不存在非空白（\f换页符，\n换行符，\r回车符，\t Tab制表符，\v垂直制表符，/图片）的字符
            '整段直接居中，单倍行距，取消缩进
            If .Range.InlineShapes.Count > 0 And pic_reg.Execute(.Range.Text).Count = 0 Then
                If .Range.InlineShapes(1).Type = wdInlineShapePicture Or _
                    .Range.InlineShapes(1).Type = wdInlineShapeChart Then  '对象类型为嵌入式图片或内嵌图表
'                    .Range.InlineShapes(1).Type = wdInlineShapeEmbeddedOLEObject
                    '设置图片（若存在）后续段落为图标题
Picture_layout:
                    On Error Resume Next
                    .Range.Next(Unit:=wdParagraph, Count:=1).style = oDoc.Styles(fig_title)
                    If Err = 0 Then '存在后续段落
                        Next_Exit = True
                    End If
                    On Error GoTo 0
                    
                    '设置图片段落格式
                    .CharacterUnitFirstLineIndent = 0 '首行缩进，字符
                    .CharacterUnitLeftIndent = 0 '左缩进，字符
                    .CharacterUnitRightIndent = 0 '右缩进，字符
                    .FirstLineIndent = 0 '首行缩进，磅值
                    .LeftIndent = 0 '左缩进，磅值
                    .RightIndent = 0 '右缩进，磅值
                    .Alignment = wdAlignParagraphCenter '居中
                    .LineSpacingRule = wdLineSpaceSingle '单倍行距
                    .SpaceBefore = 6
                    GoTo NextIter1
                End If
            End If
             
            '设置大纲段落1~4级
            If .OutlineLevel = wdOutlineLevel1 Then
                .Range.Select
'                Selection.ClearFormatting
                .Range.style = oDoc.Styles(level1)
            ElseIf .OutlineLevel = wdOutlineLevel2 Then
                .Range.Select
'                Selection.ClearFormatting
                .Range.style = oDoc.Styles(level2)
            ElseIf .OutlineLevel = wdOutlineLevel3 Then
                .Range.Select
'                Selection.ClearFormatting
                .Range.style = oDoc.Styles(level3)
            ElseIf .OutlineLevel = wdOutlineLevel4 Then
                .Range.Select
'                Selection.ClearFormatting
                .Range.style = oDoc.Styles(level4)
'            ElseIf .OutlineLevel = wdOutlineLevel5 Or _
'                .OutlineLevel = wdOutlineLevel6 Or _
'                .OutlineLevel = wdOutlineLevel7 Or _
'                .OutlineLevel = wdOutlineLevel8 Or _
'                .OutlineLevel = wdOutlineLevel9 Then'
            Else
                
                '对于已使用“PXL”列表样式的段落，存在被用户修改格式的可能, _
                '需要重置段落样式，尤其是"图题_PXL"、"表题_PXL"、 _
                '"附录表_PXL"、"附录图_PXL"这4个样式本身需要判断才应用的样式。
                Select Case .Range.style.NameLocal
                    Case fig_title
                        .Range.style = oDoc.Styles(fig_title)
                        GoTo NextIter1
                    Case tab_title
                        .Range.style = oDoc.Styles(tab_title)
                        GoTo NextIter1
                    Case Config.dict.Item(Config.ref_PXL)
                        .Range.style = oDoc.Styles(Config.dict.Item(Config.ref_PXL))
                        GoTo NextIter1
                End Select
            
                '其余列表样式处理
                If .Range.ListParagraphs.Count <> 0 Then
                        .CharacterUnitFirstLineIndent = 0 '首行缩进，字符
                        .CharacterUnitLeftIndent = 0 '左缩进，字符
                        .CharacterUnitRightIndent = 0 '右缩进，字符
                        .FirstLineIndent = CentimetersToPoints(1) '悬挂缩进1厘米
                        .LeftIndent = CentimetersToPoints(1) '左缩进1厘米
                        .RightIndent = 0 '右缩进，磅值
                        .Alignment = wdAlignParagraphJustify '两端对齐
                    GoTo NextIter1
                End If
                
                '其它为“正文_PXL”样式。
                .Range.Select
'                Selection.ClearFormatting
                .Range.style = oDoc.Styles(Config.dict.Item(Config.normal_PXL))
            End If
NextIter1:
        End With
    Next
    
'''''设置表标题、表格样式和内部文字样式'''''

    Dim tb As Table
    For Each tb In oRange.Tables
        With tb
            '以表格形式进行多图片排版的特殊情况添加图表题，仅针对无内外边框的表格
            If .Borders.InsideLineStyle = wdLineStyleNone _
                And .Borders.OutsideLineStyle = wdLineStyleNone Then
'                On Error Resume Next
                .Range.Next(Unit:=wdParagraph, Count:=1).style = oDoc.Styles(fig_title)
                GoTo NextIter2 '还需设置表格中的文字
'                On Error GoTo 0
            End If
            
            '设置表格标题
'            On Error Resume Next
            .Range.Previous(Unit:=wdParagraph, Count:=1).style = oDoc.Styles(tab_title)
'            On Error GoTo 0
    
            '设置表格为“tab_pxl”样式
            .style = oDoc.Styles(Config.dict.Item(Config.tab_PXL))
NextIter2:
            '设置表格中文字为“表文字_PXL”样式
            .Range.Select
'            Selection.ClearFormatting
            .Range.style = oDoc.Styles(Config.dict.Item(Config.tab_text_PXL))
        End With
    Next

   
'''''更新目录'''''
'修改目录1~4级样式的基准为列表段落（wdStyleListParagraph），以避免目录样式受基准样式影响
    With oDoc
        .Styles(wdStyleTOC1).BaseStyle = wdStyleListParagraph
        .Styles(wdStyleTOC2).BaseStyle = wdStyleListParagraph
        .Styles(wdStyleTOC3).BaseStyle = wdStyleListParagraph
        .Styles(wdStyleTOC4).BaseStyle = wdStyleListParagraph
    End With
    
    For Each cont In oDoc.TablesOfContents
        cont.Update
    Next
    
'''''页眉、页脚'''''
'    Dim sec As Section
'
'    '修改“页眉”、“页脚”样式的基准为列表段落（wdStyleListParagraph）
'    With oDoc
'        .Styles(wdStyleHeader).BaseStyle = wdStyleListParagraph
'        .Styles(wdStyleFooter).BaseStyle = wdStyleListParagraph
'    End With
'
'    '删除页眉内容
'    For Each sec In oRange.Sections
''        sec.Headers(wdHeaderFooterPrimary).Range.Delete
'        sec.Headers(wdHeaderFooterPrimary).Range.Text = ""
'        sec.Headers(wdHeaderFooterPrimary).Range.Borders.InsideLineStyle = wdLineStyleNone
'        sec.Headers(wdHeaderFooterPrimary).Range.Borders.OutsideLineStyle = wdLineStyleNone
'    Next
'
'    '“页脚”样式居中，取消缩进
'    With oDoc.Styles(wdStyleFooter).ParagraphFormat
'        .Alignment = wdAlignParagraphCenter
'        .CharacterUnitFirstLineIndent = 0 '首行缩进，字符
'        .CharacterUnitLeftIndent = 0 '左缩进，字符
'        .CharacterUnitRightIndent = 0 '右缩进，字符
'        .FirstLineIndent = 0 '首行缩进，磅值
'        .LeftIndent = 0 '左缩进，磅值
'        .RightIndent = 0 '右缩进，磅值
'        .Alignment = wdAlignParagraphCenter '居中
'    End With
    
'''''更新域'''''
    
    oRange.Fields.Update
End Sub


