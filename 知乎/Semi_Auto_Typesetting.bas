Attribute VB_Name = "Semi_Auto_Typesetting"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'��ģ�������ṩ���ְ��Զ��Ű湦��

Sub Text_Style(control As IRibbonControl)
'
' ������ʽ ��
' ����ѡ��������Ϊ������_PXL����ʽ
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.normal_PXL)

    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title1(control As IRibbonControl)
Attribute Title1.VB_Description = "����ѡ��������Ϊ��ʽ������1����"
Attribute Title1.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.����1"
'
' ����1 ��
' ����ѡ��������Ϊ��ʽ������1_PXL��
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level1_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title2(control As IRibbonControl)
Attribute Title2.VB_Description = "����ѡ��������Ϊ��ʽ������2����"
Attribute Title2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��1"
'
' ����2 ��
' ����ѡ��������Ϊ��ʽ������2_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level2_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title3(control As IRibbonControl)
Attribute Title3.VB_Description = "����ѡ��������Ϊ��ʽ������3����"
Attribute Title3.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.����3"
'
' ����3 ��
' ����ѡ��������Ϊ��ʽ������3_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level3_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Title4(control As IRibbonControl)
Attribute Title4.VB_Description = "����ѡ��������Ϊ��ʽ������4����"
Attribute Title4.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.����4"
'
' ����4 ��
' ����ѡ��������Ϊ��ʽ������4_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.level4_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Table_Title(control As IRibbonControl)
Attribute Table_Title.VB_Description = "����ѡ��������Ϊ��ʽ�����⡱��"
Attribute Table_Title.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.����"
'
' ���� ��
' ����ѡ��������Ϊ��ʽ������_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.tab_title_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Picture_Title(control As IRibbonControl)
Attribute Picture_Title.VB_Description = "����ѡ��������Ϊ��ʽ��ͼ�⡱��"
Attribute Picture_Title.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.ͼ��"
'
' ͼ�� ��
' ����ѡ��������Ϊ��ʽ��ͼ��_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.fig_title_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Table_Text(control As IRibbonControl)
Attribute Table_Text.VB_Description = "����ѡ��������Ϊ��ʽ�������֡���"
Attribute Table_Text.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.������"
'
' ������ ��
' ����ѡ��������Ϊ��ʽ��������_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.tab_text_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Table_Style(control As IRibbonControl)
Attribute Table_Style.VB_Description = "����ѡ�������Ϊ��ʽ��PXL����"
Attribute Table_Style.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��"
'
' �� ��
' ����ѡ�������Ϊ��ʽ��tab_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.tab_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���±����ʽ'''''

    '''''�����ѡ��Χ�Ƿ���ڱ��'''''
    Dim tb As Table
    On Error Resume Next '���ô�����
    Set tb = Selection.Tables(1)
    On Error GoTo 0 ' �رյ�ǰ����������
    
    If Not tb Is Nothing Then
        For Each tb In Selection.Tables
            With tb
                .style = style_name
            End With
        Next
    End If

End Sub
Sub Appendix1(control As IRibbonControl)
Attribute Appendix1.VB_Description = "����ѡ��������Ϊ��ʽ��������1����"
Attribute Appendix1.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��������1"
'
' ��¼1 ��
' ����ѡ��������Ϊ��ʽ����¼1_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level1_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)

    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix2(control As IRibbonControl)
Attribute Appendix2.VB_Description = "����ѡ��������Ϊ��ʽ����¼2����"
Attribute Appendix2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��¼2"
'
' ��¼2 ��
' ����ѡ��������Ϊ��ʽ����¼2_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level2_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix3(control As IRibbonControl)
Attribute Appendix3.VB_Description = "����ѡ��������Ϊ��ʽ����¼3����"
Attribute Appendix3.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��¼"
'
' ��¼3 ��
' ����ѡ��������Ϊ��ʽ����¼3_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level3_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix4(control As IRibbonControl)
'
' ��¼4 ��
' ����ѡ��������Ϊ��ʽ����¼4_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_level4_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)

    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix_Table(control As IRibbonControl)
Attribute Appendix_Table.VB_Description = "����ѡ��������Ϊ��ʽ����¼����"
Attribute Appendix_Table.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��¼��"
'
' ��¼�� ��
' ����ѡ��������Ϊ��ʽ����¼��_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_tab_title_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Appendix_Picture(control As IRibbonControl)
Attribute Appendix_Picture.VB_Description = "����ѡ��������Ϊ��ʽ����¼ͼ����"
Attribute Appendix_Picture.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��¼ͼ"
'
' ��¼ͼ ��
' ����ѡ��������Ϊ��ʽ����¼ͼ_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.appx_fig_title_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub References_Style(control As IRibbonControl)
Attribute References_Style.VB_Description = "����ѡ��������Ϊ��ʽ���ο����ס���"
Attribute References_Style.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.�ο�����"
'
' �ο����� ��
' ����ѡ��������Ϊ��ʽ���ο�����_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.ref_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Formula_Style(control As IRibbonControl)
Attribute Formula_Style.VB_Description = "����ѡ��������Ϊ��ʽ����ʽ��ʽ����"
Attribute Formula_Style.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��ʽ��ʽ"
'
' ��ʽ��ʽ ��
' ����ѡ��������Ϊ��ʽ����ʽ_PXL����
'
    Dim style_name As String
    style_name = Config.dict.Item(Config.equ_PXL)
    
    '''''��ʽ���'''''
    Exit_Flag = False 'ΪTrue���˳�
    
    Call Style_Check(style_name)
    
    If Exit_Flag = True Then 'ΪTrue���˳�
        Exit Sub
    End If
    '''''���¶�����ʽ'''''
    Selection.style = ActiveDocument.Styles(style_name)
End Sub
Sub Formula_Number(control As IRibbonControl)
Attribute Formula_Number.VB_Description = "���빫ʾ���"
Attribute Formula_Number.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��ʽ���"
'
' ��ʽ��� ��
' ���빫ʾ���
'
    On Error Resume Next '���ô�����
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "LISTNUM  ListN_PXL \l 7 ", PreserveFormatting:=False
    On Error GoTo 0 ' �رյ�ǰ����������
End Sub
Sub Style_Check(str As String)
'''''��ʽ���'''''
    Dim style_test As style
    Dim arr_style_check() As Boolean
    ReDim arr_style_check(Config.nLength - 1) '��ǰ�ܹ���17����ʽ
    Dim style_check_result As Boolean
    On Error Resume Next
    Set style_test = ActiveDocument.Styles(str)
    If Err <> 0 Then '����ʽ������
        style_check_result = Style_Check_and_Repair.Style_Check_All(arr_style_check)  '���������ʽ
        Call Style_Check_and_Repair.Style_Create(arr_style_check) '�ָ�ȱʧ����ʽ
        Err.Clear
    End If
    On Error GoTo 0
End Sub


