Attribute VB_Name = "Auto_Typesetting"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'��ģ�������ṩһ���Զ��Ű湦��
Public Config As New Configuration 'ȫ�ֲ�������
Public Exit_Flag As Boolean

Sub Auto_Typesetting_by_Selections_Normal(control As IRibbonControl)
'��ѡ���Ű�
    Dim oDoc As Document
    Dim oRange As Range
    Dim savedRange As Range '��ʱ����Selection.Rangeָ���Ķ��䷶Χ
    
'    Dim sub_exit As Boolean
'    sub_exit = False
    Set oDoc = ActiveDocument
    Set oRange = Selection.Range
    Set savedRange = Selection.Range
    Exit_Flag = False 'ΪTrue���˳�
    
    Application.ScreenUpdating = False '�ر���Ļˢ��

    Call Auto_Typesetting(oRange, oDoc, 0)

    Application.ScreenUpdating = True '����Ļˢ��
    savedRange.Select
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    
    MsgBox ("��ɣ�")
End Sub

Sub Auto_Typesetting_by_Selections_Appendix(control As IRibbonControl)
'��ѡ���Ű�
    Dim oDoc As Document
    Dim oRange As Range
    Dim savedRange As Range '��ʱ����Selection.Rangeָ���Ķ��䷶Χ
    
'    Dim sub_exit As Boolean
'    sub_exit = False
    Set oDoc = ActiveDocument
    Set oRange = Selection.Range
    Set savedRange = Selection.Range
    Exit_Flag = False 'ΪTrue���˳�
    
    Application.ScreenUpdating = False '�ر���Ļˢ��
   
    Call Auto_Typesetting(oRange, oDoc, 1)
    
    Application.ScreenUpdating = True '����Ļˢ��
    
    savedRange.Select
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    
    MsgBox ("��ɣ�")
End Sub

'Sub Auto_Typesetting_by_Pages(control As IRibbonControl)
''��ҳ���Ű�
'    Dim pg_fr As Integer
'    Dim pg_to As Integer
'    Dim oDoc As Document
'    Dim oRange As Range
'    Dim savedRange As Range '��ʱ����Selection.Rangeָ���Ķ��䷶Χ
'    Dim sub_exit As Boolean
'    sub_exit = False
'    Set oDoc = ActiveDocument
'
'    Application.ScreenUpdating = False '�ر���Ļˢ��
'
'    '����ҳ����
'    Call input_management(pg_fr, pg_to, oDoc, sub_exit)
'    If sub_exit = True Then
'        Exit Sub
'    End If
'
'    '������䷶Χ
'    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pg_fr
'    Set oRange = Selection.Range
'    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pg_to
'    oRange.End = Selection.Bookmarks("\Page").Range.End
'    Set savedRange = oRange
'    '�Զ��Ű�
'    Call Auto_Typesetting(oRange, oDoc)
'
'    Application.ScreenUpdating = True '����Ļˢ��
'
'    savedRange.Select
'    MsgBox ("��ɣ�")
'End Sub

'Sub input_management(pg_fr As Integer, pg_to As Integer, oDoc As Document, sub_exit As Boolean)
'    Dim input_str As String
'    Dim pg_count As Integer
'    pg_count = oDoc.ComputeStatistics(wdStatisticPages)
'
'    Do
'        input_str = InputBox("�������Ű��  ��ʼ  ҳ�룺", Title:="�������Ű��  ��ʼ  ҳ��")
'
'        If input_str = "" Then 'ȡ����رմ���
'            sub_exit = True
'            Exit Sub
'        End If
'
'        If Len(input_str) = 0 Then
'            answer = MsgBox("δ�����κ�ֵ��", vbOKOnly, "�������Ű��  ��ʼ  ҳ��")
'        ElseIf IsNumeric(input_str) = False Then
'            answer = MsgBox("���벻Ϊ���֣�������������", vbOKOnly, "�������Ű��  ��ʼ  ҳ��")
'        ElseIf IsNumeric(input_str) = True Then
'            If CInt(input_str) = input_str Then
'                pg_fr = CInt(input_str)
'                If pg_fr > pg_count Or pg_fr <= 0 Then
'                    answer = MsgBox("�����ҳ�볬���ĵ���Χ��", vbOKOnly, "�������Ű��  ��ʼ  ҳ��")
'                Else
'                    Exit Do
'                End If
'            Else
'                answer = MsgBox("���벻Ϊ������������������", vbOKOnly, "�������Ű��  ��ʼ  ҳ��")
'            End If
'        End If
'    Loop Until False
'
'    Do
'        input_str = InputBox("�������Ű��  ��ֹ  ҳ�룺", Title:="�������Ű��  ��ֹ  ҳ��")
'
'        If input_str = "" Then 'ȡ����رմ���
'            sub_exit = True
'            Exit Sub
'        End If
'
'        If Len(input_str) = 0 Then
'            answer = MsgBox("δ�����κ�ֵ��", vbOKOnly, "�������Ű��  ��ֹ  ҳ��")
'        ElseIf IsNumeric(input_str) = False Then
'            answer = MsgBox("���벻Ϊ���֣�������������", vbOKOnly, "�������Ű��  ��ֹ  ҳ��")
'        ElseIf IsNumeric(input_str) = True Then
'            If CInt(input_str) = input_str Then
'                pg_to = CInt(input_str)
'                If pg_to > pg_count Then
'                    answer = MsgBox("�����ҳ�볬���ĵ���Χ��", vbOKOnly, "�������Ű��  ��ֹ  ҳ��")
'                ElseIf pg_fr > pg_to Then
'                    answer = MsgBox("�����ҳ��С����ʼҳ�룡", vbOKOnly, "�������Ű��  ��ֹ  ҳ��")
'                Else
'                    Exit Do
'                End If
'            Else
'                answer = MsgBox("���벻Ϊ������������������", vbOKOnly, "�������Ű��  ��ֹ  ҳ��")
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
        .Pattern = "Ŀ[\s]*¼"
    End With
    
    With equation_reg
        .Global = False
        .Pattern = "[a-zA-Z\u4e00-\u9fa5]"
    End With
    
    Dim level1 As String '���1��������ʽ
    Dim level2 As String '���2��������ʽ
    Dim level3 As String '���3��������ʽ
    Dim level4 As String '���4��������ʽ
    Dim tab_title As String '�����
    Dim fig_title As String 'ͼ����
    
    If model = 0 Then
        level1 = Config.dict.Item(Config.level1_PXL) '"����1_PXL"
        level2 = Config.dict.Item(Config.level2_PXL) '"����2_PXL"
        level3 = Config.dict.Item(Config.level3_PXL) '"����3_PXL"
        level4 = Config.dict.Item(Config.level4_PXL) '"����4_PXL"
        tab_title = Config.dict.Item(Config.tab_title_PXL) '"����_PXL"
        fig_title = Config.dict.Item(Config.fig_title_PXL) '"ͼ��_PXL"
    ElseIf model = 1 Then
        level1 = Config.dict.Item(Config.appx_level1_PXL) '"��¼1_PXL"
        level2 = Config.dict.Item(Config.appx_level2_PXL) '"��¼2_PXL"
        level3 = Config.dict.Item(Config.appx_level3_PXL) '"��¼3_PXL"
        level4 = Config.dict.Item(Config.appx_level4_PXL) '"��¼4_PXL"
        tab_title = Config.dict.Item(Config.appx_tab_title_PXL) '"��¼��_PXL"
        fig_title = Config.dict.Item(Config.appx_fig_title_PXL) '"��¼ͼ_PXL"
    Else
        Exit_Flag = True
        MsgBox ("�Զ��Ű�ģʽ����")
        Exit Sub
    End If
'''''�Զ�����ʽ���'''''
    Dim style_check_result As Boolean
    Dim arr_style_check() As Boolean
    ReDim arr_style_check(Config.nLength - 1) '��ǰ�ܹ���17����ʽ
    style_check_result = True
    style_check_result = Style_Check_and_Repair.Style_Check_All(arr_style_check)  '���������ʽ
    If style_check_result <> True Then
        Call Style_Check_and_Repair.Style_Create(arr_style_check) '���ָ�ȱʧ����ʽ
    End If
    
    If Exit_Flag = True Then
        Exit Sub
    End If
           
'''''����1~4�����⡢ͼ���⡢����⡢ͼƬ'''''
    '��ʼ�������⸴������ʱ��ѡ���λ��������ʽ������ ������_PXL������������������Ķ������������ʽ�����硰����_PXL����
'    With oRange.Paragraphs
'        .CharacterUnitFirstLineIndent = 0 '�����������ַ�
'        .CharacterUnitLeftIndent = 0 '���������ַ�
'        .CharacterUnitRightIndent = 0 '���������ַ�
'        .FirstLineIndent = 0 '������������ֵ
'        .LeftIndent = 0 '����������ֵ
'        .RightIndent = 0 '����������ֵ
'        .Alignment = wdAlignParagraphCenter '���˶���
'    End With
    
    For Each pg In oRange.Paragraphs
        With pg
            
            If Next_Exit = True Then '����ͼƬ����������Ϊ���⣬����������2�Σ���������2��
                Next_Exit = False
                GoTo NextIter1
            End If
            
            '��񲿷ֶ����Ȳ�����
            If .Range.Information(wdWithInTable) Then
                GoTo NextIter1
            End If
            
'            ��������ʽ_PXL����ʽ����
            If .style.NameLocal = Config.dict.Item(Config.equ_PXL) Then
                GoTo NextIter1
            End If
            
            'word���ù�ʽ����������Ӣ�ģ�������
            If .Range.OMaths.Count <> 0 And equation_reg.Execute(.Range.Text).Count = 0 Then
                GoTo NextIter1
            End If
            
            ' ����Ŀ¼����
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
                        And .Range.Fields(1).OLEFormat.ClassType Like "Equation.DSMT4" Then 'mathtype��ʽ
                        GoTo NextIter1
                    End If
                
                
'                ElseIf equation_reg.Execute(.Range.Text).Count = 0 Then
'                    If .Range.Fields(1).OLEFormat.ClassType = "Equation.DSMT4" Then 'mathtype��ʽ
'
'
'
'                    End If
'                    GoTo NextIter1
                End If
            End If
            
            'Ŀ¼���ⲻ����
            If content_reg.Execute(.Range.Text).Count <> 0 Then
                GoTo NextIter1
            End If
                
            '����Ƕ��ʽͼƬת��ΪǶ��ʽͼƬ
            For Each pg_shape In .Range.ShapeRange
                If pg_shape.Type = msoPicture Then
                    pg_shape.ConvertToInlineShape
                End If
            Next

            '���京��Ƕ��ʽͼƬ�Ҳ����ڷǿհף�\f��ҳ����\n���з���\r�س�����\t Tab�Ʊ����\v��ֱ�Ʊ����/ͼƬ�����ַ�
            '����ֱ�Ӿ��У������о࣬ȡ������
            If .Range.InlineShapes.Count > 0 And pic_reg.Execute(.Range.Text).Count = 0 Then
                If .Range.InlineShapes(1).Type = wdInlineShapePicture Or _
                    .Range.InlineShapes(1).Type = wdInlineShapeChart Then  '��������ΪǶ��ʽͼƬ����Ƕͼ��
'                    .Range.InlineShapes(1).Type = wdInlineShapeEmbeddedOLEObject
                    '����ͼƬ�������ڣ���������Ϊͼ����
Picture_layout:
                    On Error Resume Next
                    .Range.Next(Unit:=wdParagraph, Count:=1).style = oDoc.Styles(fig_title)
                    If Err = 0 Then '���ں�������
                        Next_Exit = True
                    End If
                    On Error GoTo 0
                    
                    '����ͼƬ�����ʽ
                    .CharacterUnitFirstLineIndent = 0 '�����������ַ�
                    .CharacterUnitLeftIndent = 0 '���������ַ�
                    .CharacterUnitRightIndent = 0 '���������ַ�
                    .FirstLineIndent = 0 '������������ֵ
                    .LeftIndent = 0 '����������ֵ
                    .RightIndent = 0 '����������ֵ
                    .Alignment = wdAlignParagraphCenter '����
                    .LineSpacingRule = wdLineSpaceSingle '�����о�
                    .SpaceBefore = 6
                    GoTo NextIter1
                End If
            End If
             
            '���ô�ٶ���1~4��
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
                
                '������ʹ�á�PXL���б���ʽ�Ķ��䣬���ڱ��û��޸ĸ�ʽ�Ŀ���, _
                '��Ҫ���ö�����ʽ��������"ͼ��_PXL"��"����_PXL"�� _
                '"��¼��_PXL"��"��¼ͼ_PXL"��4����ʽ������Ҫ�жϲ�Ӧ�õ���ʽ��
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
            
                '�����б���ʽ����
                If .Range.ListParagraphs.Count <> 0 Then
                        .CharacterUnitFirstLineIndent = 0 '�����������ַ�
                        .CharacterUnitLeftIndent = 0 '���������ַ�
                        .CharacterUnitRightIndent = 0 '���������ַ�
                        .FirstLineIndent = CentimetersToPoints(1) '��������1����
                        .LeftIndent = CentimetersToPoints(1) '������1����
                        .RightIndent = 0 '����������ֵ
                        .Alignment = wdAlignParagraphJustify '���˶���
                    GoTo NextIter1
                End If
                
                '����Ϊ������_PXL����ʽ��
                .Range.Select
'                Selection.ClearFormatting
                .Range.style = oDoc.Styles(Config.dict.Item(Config.normal_PXL))
            End If
NextIter1:
        End With
    Next
    
'''''���ñ���⡢�����ʽ���ڲ�������ʽ'''''

    Dim tb As Table
    For Each tb In oRange.Tables
        With tb
            '�Ա����ʽ���ж�ͼƬ�Ű������������ͼ���⣬�����������߿�ı��
            If .Borders.InsideLineStyle = wdLineStyleNone _
                And .Borders.OutsideLineStyle = wdLineStyleNone Then
'                On Error Resume Next
                .Range.Next(Unit:=wdParagraph, Count:=1).style = oDoc.Styles(fig_title)
                GoTo NextIter2 '�������ñ���е�����
'                On Error GoTo 0
            End If
            
            '���ñ�����
'            On Error Resume Next
            .Range.Previous(Unit:=wdParagraph, Count:=1).style = oDoc.Styles(tab_title)
'            On Error GoTo 0
    
            '���ñ��Ϊ��tab_pxl����ʽ
            .style = oDoc.Styles(Config.dict.Item(Config.tab_PXL))
NextIter2:
            '���ñ��������Ϊ��������_PXL����ʽ
            .Range.Select
'            Selection.ClearFormatting
            .Range.style = oDoc.Styles(Config.dict.Item(Config.tab_text_PXL))
        End With
    Next

   
'''''����Ŀ¼'''''
'�޸�Ŀ¼1~4����ʽ�Ļ�׼Ϊ�б���䣨wdStyleListParagraph�����Ա���Ŀ¼��ʽ�ܻ�׼��ʽӰ��
    With oDoc
        .Styles(wdStyleTOC1).BaseStyle = wdStyleListParagraph
        .Styles(wdStyleTOC2).BaseStyle = wdStyleListParagraph
        .Styles(wdStyleTOC3).BaseStyle = wdStyleListParagraph
        .Styles(wdStyleTOC4).BaseStyle = wdStyleListParagraph
    End With
    
    For Each cont In oDoc.TablesOfContents
        cont.Update
    Next
    
'''''ҳü��ҳ��'''''
'    Dim sec As Section
'
'    '�޸ġ�ҳü������ҳ�š���ʽ�Ļ�׼Ϊ�б���䣨wdStyleListParagraph��
'    With oDoc
'        .Styles(wdStyleHeader).BaseStyle = wdStyleListParagraph
'        .Styles(wdStyleFooter).BaseStyle = wdStyleListParagraph
'    End With
'
'    'ɾ��ҳü����
'    For Each sec In oRange.Sections
''        sec.Headers(wdHeaderFooterPrimary).Range.Delete
'        sec.Headers(wdHeaderFooterPrimary).Range.Text = ""
'        sec.Headers(wdHeaderFooterPrimary).Range.Borders.InsideLineStyle = wdLineStyleNone
'        sec.Headers(wdHeaderFooterPrimary).Range.Borders.OutsideLineStyle = wdLineStyleNone
'    Next
'
'    '��ҳ�š���ʽ���У�ȡ������
'    With oDoc.Styles(wdStyleFooter).ParagraphFormat
'        .Alignment = wdAlignParagraphCenter
'        .CharacterUnitFirstLineIndent = 0 '�����������ַ�
'        .CharacterUnitLeftIndent = 0 '���������ַ�
'        .CharacterUnitRightIndent = 0 '���������ַ�
'        .FirstLineIndent = 0 '������������ֵ
'        .LeftIndent = 0 '����������ֵ
'        .RightIndent = 0 '����������ֵ
'        .Alignment = wdAlignParagraphCenter '����
'    End With
    
'''''������'''''
    
    oRange.Fields.Update
End Sub


