Attribute VB_Name = "Style_Check_and_Repair"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'��ģ�����ڼ��ͻָ���ʽ

Function Style_Check_All(arr_style_check() As Boolean) As Boolean
'''''ģ��ȫ��ʽɸ��'''''
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
'����ȱʧ����ʽ
'
    Dim template_temp As Template
    Dim templatePath As String
    Dim templateDoc As Document
    Dim currentDoc As Document
    Dim savedRange As Range
    Dim style_name As String

'    Application.ScreenUpdating = False '�ر���Ļˢ��

    Dim newListTemplate As ListTemplate

'��ʱ����Selection.Rangeָ���Ķ��䷶Χ
    Set savedRange = Selection.Range

'����ģ���ļ���·��
    templatePath = "Error" '��ʼ��
    For Each template_temp In Templates
        If template_temp.Name = "��Word�Զ��Ű湤�ߡ�V3.0.dotm" Then
            templatePath = template_temp.FullName
            Exit For
        End If
    Next

    If templatePath = "Error" Then
        Exit_Flag = True
        MsgBox "δ���ҵ��Զ��Ű湤��!" & vbCrLf & "����ģ���ļ����Ƶȣ�"
        Exit Sub
    End If
'    templatePath = "C:\Program Files\Microsoft Office\root\Office16\STARTUP\��Word�Զ��Ű湤�ߡ�V3.0.dotm"

'��ȡ��ǰ�ĵ�����
    Set currentDoc = ActiveDocument

'��ģ���ļ�
    Set templateDoc = Documents.Open(templatePath, ReadOnly:=True, Visible:=False)

'�޸�ȱʧ����ʽ
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
                    currentDoc.Styles(style_name).Font = templateDoc.Styles(style_name).Font '����Application.OrganizerCopy�޷��ȶ�����ģ������ָ�ʽ
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
                '��Ч
            End Select
        End If
    Next

' �ر�ģ���ļ������������
    templateDoc.Close SaveChanges:=wdDoNotSaveChanges
    
'������Ķ༶�б�ListN_PXL���Ƿ���ڣ��������򴴽�
    On Error Resume Next
        Set newListTemplate = ActiveDocument.ListTemplates("ListN_PXL")

        If Err <> 0 Then
            Call Style_Library.Multi_level_list '�����༶�б�������ĵĸ�����ʽ
        End If
    On Error GoTo 0

'��⸽¼�༶�б�ListA_PXL���Ƿ���ڣ��������򴴽�
    On Error Resume Next
        Set newListTemplate = ActiveDocument.ListTemplates("ListA_PXL")

        If Err <> 0 Then
            Call Style_Library.Multi_level_list_APPX '�����༶�б������¼�ĸ�����ʽ
        End If
    On Error GoTo 0

'���ο������б�ListR_PXL���Ƿ���ڣ��������򴴽�
    On Error Resume Next
        Set newListTemplate = ActiveDocument.ListTemplates("ListR_PXL")
        If Err <> 0 Then
            Call Style_Library.ref_list '���������б�����ο�������ʽ
        End If
    On Error GoTo 0

    '������ʽ���б�
    Set newListTemplate = currentDoc.ListTemplates("ListN_PXL")
    With newListTemplate
        .ListLevels(1).LinkedStyle = Config.dict.Item(Config.level1_PXL) ' "����1_PXL"
        .ListLevels(2).LinkedStyle = Config.dict.Item(Config.level2_PXL) ' "����2_PXL"
        .ListLevels(3).LinkedStyle = Config.dict.Item(Config.level3_PXL) ' "����3_PXL"
        .ListLevels(4).LinkedStyle = Config.dict.Item(Config.level4_PXL) ' "����4_PXL"
        .ListLevels(8).LinkedStyle = Config.dict.Item(Config.fig_title_PXL) ' "ͼ��_PXL"
        .ListLevels(9).LinkedStyle = Config.dict.Item(Config.tab_title_PXL) ' "����_PXL"
    End With

    '������ʽ���б�
    Set newListTemplate = currentDoc.ListTemplates("ListA_PXL")
    With newListTemplate
        .ListLevels(1).LinkedStyle = Config.dict.Item(Config.appx_level1_PXL) ' "��¼1_PXL"
        .ListLevels(2).LinkedStyle = Config.dict.Item(Config.appx_level2_PXL) ' "��¼2_PXL"
        .ListLevels(3).LinkedStyle = Config.dict.Item(Config.appx_level3_PXL) ' "��¼3_PXL"
        .ListLevels(4).LinkedStyle = Config.dict.Item(Config.appx_level4_PXL) ' "��¼4_PXL"
        .ListLevels(7).LinkedStyle = Config.dict.Item(Config.appx_fig_title_PXL) ' "��¼ͼ_PXL"
        .ListLevels(8).LinkedStyle = Config.dict.Item(Config.appx_tab_title_PXL) ' "��¼��_PXL"
    End With

    '������ʽ���б�
    Set newListTemplate = currentDoc.ListTemplates("ListR_PXL")
    newListTemplate.ListLevels(1).LinkedStyle = Config.dict.Item(Config.ref_PXL) ' "�ο�����_PXL"

'�ָ�ԭ��ѡ��ķ�Χ
    savedRange.Select
    
'����Ļˢ��
'    Application.ScreenUpdating = True
End Sub



