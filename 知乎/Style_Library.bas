Attribute VB_Name = "Style_Library"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'��ģ�鶨��ģ��ĸ��ָ�ʽ

''''''���ָ�ʽ����˵������''''''
'    With ActiveDocument.Styles("xxxxx").Font '������������
'        .NameFarEast = "����" '���ö����ַ���Ϊ�����塱
'        .NameAscii = "Times New Roman" '���������ı��ַ���Ϊ��Times New Roman��
'        .NameOther = "Times New Roman" '���������ַ���Ϊ��Times New Roman��
'        .Name = "Times New Roman" '���������ַ����µ�����Ϊ��Times New Roman��
'        .Size = 22 '�趨�����СΪ22��
'        .Bold = False '�رմ���
'        .Italic = False '�ر�б��
'        .Underline = wdUnderlineNone '�ر��»���
'        .UnderlineColor = wdColorAutomatic '���»�����ɫ��Ϊ�Զ����������ı���ɫ��
'        .StrikeThrough = False '�ر�ɾ����Ч��
'        .DoubleStrikeThrough = False '�ر�˫ɾ����Ч��
'        .Outline = False '�ر�����Ч��
'        .Emboss = False '�رո���Ч��
'        .Shadow = False '�ر���ӰЧ��
'        .Hidden = False '�ر������ı�Ч��
'        .SmallCaps = False '�ر�С�ʹ�д��ĸЧ��
'        .AllCaps = False '�ر�ȫ����д��ĸЧ��
'        .Color = wdColorAutomatic '����������ɫΪ�Զ����������ĵ�����򸸼���ʽ����ɫ��
'        .Engrave = False '�رյ��Ч��
''        .Superscript = False '�ر��ϱ�Ч��
''        .Subscript = False '�ر��±�Ч��
'        .Scaling = 100 '�����������ű���Ϊ100%��Ĭ��ֵ��
'        .Kerning = 1 '�����ַ����Ϊ1pt
'        .Animation = wdAnimationNone '�ر����ֶ���Ч��
'        .DisableCharacterSpaceGrid = False '�����ַ����񲼾�
'        .EmphasisMark = wdEmphasisMarkNone '�ر�ǿ�����
'        .Ligatures = wdLigaturesNone '�ر�����
'        .NumberSpacing = wdNumberSpacingDefault '�趨���ּ��ΪĬ��
'        .NumberForm = wdNumberFormDefault '�趨������ʽΪĬ��
'        .StylisticSet = wdStylisticSetDefault '�趨���ΪĬ��
'        .ContextualAlternates = 0 '���ر������Ľ�������
'    End With
'    With ActiveDocument.Styles("xxxxx").ParagraphFormat '���ö����ʽ����
'        .LeftIndent = 0 '���ö����������Ϊ0����
'        .RightIndent = 0 '���ö����������Ϊ0����
'        .SpaceBefore = 6 '���ö�ǰ���Ϊ6��
'        .SpaceBeforeAuto = False '�ر��Զ�������ǰ���Ĺ���
'        .SpaceAfter = 6 '���öκ���Ϊ6��
'        .SpaceAfterAuto = False '�ر��Զ������κ���Ĺ���
'        .LineSpacingRule = wdLineSpace1pt5 '�����о����Ϊ1.5���о�
'        .Alignment = wdAlignParagraphCenter '�������ı����뷽ʽ����Ϊ���ж���
'        .WidowControl = False '�رչ��п��ƹ��ܣ��������βֻ��һ�г�������һҳ����һ����
'        .KeepWithNext = False '��ǿ�Ƶ�ǰ���������Ķ���ͬҳ��ʾ
'        .KeepTogether = False '��������ҳ�Ͽ�����ǿ���������䱣����ͬһҳ��
'        .PageBreakBefore = True '�ڸö���ǰ�����ҳ����ʹ�ö���ʼ�մ��µ�һҳ��ʼ��
'        .NoLineNumber = False '�������к���ʾ�����ö���������ʾ�кţ�����ĵ��ѿ����к���ʾ��
'        .Hyphenation = True '�����Զ��ϴʹ��ܣ�����Word�ڵ��ʻ���ʱ�������ַ�
'        .FirstLineIndent = 0 '������������Ϊ0���ס�
'        .OutlineLevel = wdOutlineLevel1 '����������Ϊ��ټ���1
'        .CharacterUnitLeftIndent = 0 '�������ַ�Ϊ��λ��������Ϊ0���ַ�
'        .CharacterUnitRightIndent = 0 '�������ַ�Ϊ��λ��������Ϊ0���ַ�
'        .CharacterUnitFirstLineIndent = 0 '�������ַ�Ϊ��λ����������Ϊ0���ַ�
'        .LineUnitBefore = 0 '��������Ϊ��λ�Ķ�ǰ���Ϊ0��
'        .LineUnitAfter = 0 '��������Ϊ��λ�Ķκ���Ϊ0��
'        .MirrorIndents = False '�رվ�������������ʹ���б�������븸�б�Գ�
'        .TextboxTightWrap = wdTightNone '�����ı�����ܻ�����ʽΪ�޽��ܻ��ƣ����ı�������Χ�ı��䱣����׼���
'        .CollapsedByDefault = False '��ʹ�ö�����ʽ�ڴ����ͼ��Ĭ���۵�
'        .AutoAdjustRightIndent = True '����Word���ݶ�����Ʊ�λ�������������Զ�����������
'        .DisableLineHeightGrid = False '�������и����񣬼������о����ĵ���Ĭ���������
'        .FarEastLineBreakControl = True ' �Զ�����������������ж��ֿ���
'        .WordWrap = True '���õ��ʻ��й��ܣ����������ʿ�Խ���С�
'        .HangingPunctuation = True '�������ұ����ţ�ʹ�䳬�������ı��߽磬����Ű����۶�
'        .HalfWidthPunctuationOnTopOfLine = False '��ֹ��Ǳ����ų���������
'        .AddSpaceBetweenFarEastAndAlpha = True '�ڶ����ַ��������ַ�֮����Ӷ���ո�
'        .AddSpaceBetweenFarEastAndDigit = True '�ڶ����ַ�������֮����Ӷ���ո�
'        .BaseLineAlignment = wdBaselineAlignAuto '���û��߶��뷽ʽΪ�Զ�����Ĭ�ϵĻ��߶��뷽ʽ
'    End With



Sub Multi_level_list()

' ���Ķ༶�б�


    Dim newListTemplate As ListTemplate
    Set newListTemplate = ActiveDocument.ListTemplates.Add(Name:="ListN_PXL") '�����б�Ĭ��Ϊ����
    newListTemplate.Convert (1) 'תΪ�༶�б�

    With newListTemplate.ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.level1_PXL) '����1_PXL
    End With
    With newListTemplate.ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.86) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.level2_PXL) '����2_PXL
    End With
    With newListTemplate.ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.2) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.level3_PXL) '����3_PXL
    End With
    With newListTemplate.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.49) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.level4_PXL) '����4_PXL
    End With
'    With newListTemplate.ListLevels(5)
'        .NumberFormat = "%1.%2.%3.%4.%5"
'        .TrailingCharacter = wdTrailingSpace
'        .NumberStyle = wdListNumberStyleArabic
'        .NumberPosition = 0
'        .Alignment = wdListLevelAlignLeft
'        .TextPosition = 0
'        .TabPosition = wdUndefined
'        .ResetOnHigher = 4
'        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
'    End With
'    With newListTemplate.ListLevels(6)
'        .NumberFormat = "%1.%2.%3.%4.%5.%6"
'        .TrailingCharacter = wdTrailingTab
'        .NumberStyle = wdListNumberStyleArabic
'        .NumberPosition = 0
'        .Alignment = wdListLevelAlignLeft
'        .TextPosition = 0
'        .TabPosition = wdUndefined
'        .ResetOnHigher = 5
'        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
'    End With
    With newListTemplate.ListLevels(7)
        .NumberFormat = "��%1.%7��"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
    End With
    With newListTemplate.ListLevels(8)
        .NumberFormat = "ͼ%1.%8"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.fig_title_PXL) 'ͼ��_PXL
    End With
    With newListTemplate.ListLevels(9)
        .NumberFormat = "��%1.%9"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.tab_title_PXL) '����_PXL
    End With

End Sub

Sub Multi_level_list_APPX()

' ��¼�༶����



    Dim newListTemplate As ListTemplate
    Set newListTemplate = ActiveDocument.ListTemplates.Add(Name:="ListA_PXL") '�����б�Ĭ��Ϊ����
    newListTemplate.Convert (1) 'תΪ�༶�б�

    With newListTemplate.ListLevels(1)
        .NumberFormat = "��¼%1"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleUppercaseLetter
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.appx_level1_PXL) '��¼1_PXL
    End With
    With newListTemplate.ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.99) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.appx_level2_PXL) '��¼2_PXL
    End With
    With newListTemplate.ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.33) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.appx_level3_PXL) '��¼3_PXL
    End With
    With newListTemplate.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.6) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.appx_level4_PXL) '��¼4_PXL
    End With
'    With newListTemplate.ListLevels(5)
'        .NumberFormat = "%1.%2.%3.%4.%5"
'        .TrailingCharacter = wdTrailingTab
'        .NumberStyle = wdListNumberStyleArabic
'        .NumberPosition = 0
'        .Alignment = wdListLevelAlignLeft
'        .TextPosition = 0
'        .TabPosition = wdUndefined
'        .ResetOnHigher = 4
'        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
'    End With
'    With newListTemplate.ListLevels(6)
'        .NumberFormat = "%1.%2.%3.%4.%5.%6"
'        .TrailingCharacter = wdTrailingTab
'        .NumberStyle = wdListNumberStyleArabic
'        .NumberPosition = 0
'        .Alignment = wdListLevelAlignLeft
'        .TextPosition = 0
'        .TabPosition = wdUndefined
'        .ResetOnHigher = 5
'        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
'    End With
    With newListTemplate.ListLevels(7)
        .NumberFormat = "ͼ%1.%7"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.appx_fig_title_PXL) '��¼ͼ_PXL
    End With
    With newListTemplate.ListLevels(8)
        .NumberFormat = "��%1.%8"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.appx_tab_title_PXL) '��¼��_PXL
    End With
'    With newListTemplate.ListLevels(9)
'        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9"
'        .TrailingCharacter = wdTrailingTab
'        .NumberStyle = wdListNumberStyleArabic
'        .NumberPosition = 0
'        .Alignment = wdListLevelAlignLeft
'        .TextPosition = 0
'        .TabPosition = wdUndefined
'        .ResetOnHigher = 8
'        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = ""
'    End With

End Sub

Sub ref_list()

' �ο������б�


    Dim newListTemplate As ListTemplate
    Set newListTemplate = ActiveDocument.ListTemplates.Add(Name:="ListR_PXL") '�����б�Ĭ��Ϊ����

    With newListTemplate.ListLevels(1)
        .NumberFormat = "[%1]"
        .TrailingCharacter = wdTrailingTab '��ź� wdTrailingTab=�Ʊ�λ,wdTrailingSpace=�ո�
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1) '��������
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
'        With .Font
'            .Bold = wdUndefined
'            .Italic = wdUndefined
'            .StrikeThrough = wdUndefined
'            .Subscript = wdUndefined
'            .Superscript = wdUndefined
'            .Shadow = wdUndefined
'            .Outline = wdUndefined
'            .Emboss = wdUndefined
'            .Engrave = wdUndefined
'            .AllCaps = wdUndefined
'            .Hidden = wdUndefined
'            .Underline = wdUndefined
'            .Color = wdUndefined
'            .Size = wdUndefined
'            .Animation = wdUndefined
'            .DoubleStrikeThrough = wdUndefined
'            .Name = ""
'        End With
'        .LinkedStyle = Config.dict.Item(Config.ref_PXL) '�ο�����_PXL
    End With
End Sub

'Sub normal_PXL()
'
' ������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.normal_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 14
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphJustify
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0.99)
'        .OutlineLevel = wdOutlineLevelBodyText
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 2
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub level1_PXL()
'
' ����1��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.level1_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph '��Ӷ�����ʽ
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False 'ȡ���Զ�����
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font '������������
'        .NameFarEast = "����" '���ö����ַ���Ϊ�����塱
'        .NameAscii = "Times New Roman" '���������ı��ַ���Ϊ��Times New Roman��
'        .NameOther = "Times New Roman" '���������ַ���Ϊ��Times New Roman��
'        .Name = "Times New Roman" '���������ַ����µ�����Ϊ��Times New Roman��
'        .Size = 22 '�趨�����СΪ22��
'        .Bold = False '�رմ���
'        .Italic = False '�ر�б��
'        .Underline = wdUnderlineNone '�ر��»���
'        .UnderlineColor = wdColorAutomatic '���»�����ɫ��Ϊ�Զ����������ı���ɫ��
'        .StrikeThrough = False '�ر�ɾ����Ч��
'        .DoubleStrikeThrough = False '�ر�˫ɾ����Ч��
'        .Outline = False '�ر�����Ч��
'        .Emboss = False '�رո���Ч��
'        .Shadow = False '�ر���ӰЧ��
'        .Hidden = False '�ر������ı�Ч��
'        .SmallCaps = False '�ر�С�ʹ�д��ĸЧ��
'        .AllCaps = False '�ر�ȫ����д��ĸЧ��
'        .Color = wdColorAutomatic '����������ɫΪ�Զ����������ĵ�����򸸼���ʽ����ɫ��
'        .Engrave = False '�رյ��Ч��
''        .Superscript = False '�ر��ϱ�Ч��
''        .Subscript = False '�ر��±�Ч��
'        .Scaling = 100 '�����������ű���Ϊ100%��Ĭ��ֵ��
'        .Kerning = 1 '�����ַ����Ϊ1pt
'        .Animation = wdAnimationNone '�ر����ֶ���Ч��
'        .DisableCharacterSpaceGrid = False '�����ַ����񲼾�
'        .EmphasisMark = wdEmphasisMarkNone '�ر�ǿ�����
'        .Ligatures = wdLigaturesNone '�ر�����
'        .NumberSpacing = wdNumberSpacingDefault '�趨���ּ��ΪĬ��
'        .NumberForm = wdNumberFormDefault '�趨������ʽΪĬ��
'        .StylisticSet = wdStylisticSetDefault '�趨���ΪĬ��
'        .ContextualAlternates = 0 '���ر������Ľ�������
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat '���ö����ʽ����
'        .LeftIndent = 0 '���ö����������Ϊ0����
'        .RightIndent = 0 '���ö����������Ϊ0����
'        .SpaceBefore = 6 '���ö�ǰ���Ϊ6��
'        .SpaceBeforeAuto = False '�ر��Զ�������ǰ���Ĺ���
'        .SpaceAfter = 6 '���öκ���Ϊ6��
'        .SpaceAfterAuto = False '�ر��Զ������κ���Ĺ���
'        .LineSpacingRule = wdLineSpace1pt5 '�����о����Ϊ1.5���о�
'        .Alignment = wdAlignParagraphCenter '�������ı����뷽ʽ����Ϊ���ж���
'        .WidowControl = False '�رչ��п��ƹ��ܣ��������βֻ��һ�г�������һҳ����һ����
'        .KeepWithNext = False '��ǿ�Ƶ�ǰ���������Ķ���ͬҳ��ʾ
'        .KeepTogether = False '��������ҳ�Ͽ�����ǿ���������䱣����ͬһҳ��
'        .PageBreakBefore = True '�ڸö���ǰ�����ҳ����ʹ�ö���ʼ�մ��µ�һҳ��ʼ��
'        .NoLineNumber = False '�������к���ʾ�����ö���������ʾ�кţ�����ĵ��ѿ����к���ʾ��
'        .Hyphenation = True '�����Զ��ϴʹ��ܣ�����Word�ڵ��ʻ���ʱ�������ַ�
'        .FirstLineIndent = 0 '������������Ϊ0���ס�
'        .OutlineLevel = wdOutlineLevel1 '����������Ϊ��ټ���1
'        .CharacterUnitLeftIndent = 0 '�������ַ�Ϊ��λ��������Ϊ0���ַ�
'        .CharacterUnitRightIndent = 0 '�������ַ�Ϊ��λ��������Ϊ0���ַ�
'        .CharacterUnitFirstLineIndent = 0 '�������ַ�Ϊ��λ����������Ϊ0���ַ�
'        .LineUnitBefore = 0 '��������Ϊ��λ�Ķ�ǰ���Ϊ0��
'        .LineUnitAfter = 0 '��������Ϊ��λ�Ķκ���Ϊ0��
'        .MirrorIndents = False '�رվ�������������ʹ���б�������븸�б�Գ�
'        .TextboxTightWrap = wdTightNone '�����ı�����ܻ�����ʽΪ�޽��ܻ��ƣ����ı�������Χ�ı��䱣����׼���
'        .CollapsedByDefault = False '��ʹ�ö�����ʽ�ڴ����ͼ��Ĭ���۵�
'        .AutoAdjustRightIndent = True '����Word���ݶ�����Ʊ�λ�������������Զ�����������
'        .DisableLineHeightGrid = False '�������и����񣬼������о����ĵ���Ĭ���������
'        .FarEastLineBreakControl = True ' �Զ�����������������ж��ֿ���
'        .WordWrap = True '���õ��ʻ��й��ܣ����������ʿ�Խ���С�
'        .HangingPunctuation = True '�������ұ����ţ�ʹ�䳬�������ı��߽磬����Ű����۶�
'        .HalfWidthPunctuationOnTopOfLine = False '��ֹ��Ǳ����ų���������
'        .AddSpaceBetweenFarEastAndAlpha = True '�ڶ����ַ��������ַ�֮����Ӷ���ո�
'        .AddSpaceBetweenFarEastAndDigit = True '�ڶ����ַ�������֮����Ӷ���ո�
'        .BaseLineAlignment = wdBaselineAlignAuto '���û��߶��뷽ʽΪ�Զ�����Ĭ�ϵĻ��߶��뷽ʽ
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub level2_PXL()
'
' ����2��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.level2_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 16
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 6
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 6
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevel2
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub level3_PXL()
'
' ����3��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.level3_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 15
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 3
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 3
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevel3
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub level4_PXL()
'
' ����4��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.level4_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 14
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(3.5)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 3
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(-3.5)
'        .OutlineLevel = wdOutlineLevel4
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub tab_title_PXL()
'
' ���ı�����ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.tab_title_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:= _
'        wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 10.5
'        .Bold = True
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 2.5
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphCenter
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevel5
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0.5
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle _
'         = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub fig_title_PXL()
'
' ����ͼ����ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.fig_title_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:= _
'        wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 10.5
'        .Bold = True
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 2.5
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphCenter
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevel5
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0.5
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle _
'         = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub tab_text_PXL()
'
' ��������ʽ��ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.tab_text_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 12
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpaceSingle
'        .Alignment = wdAlignParagraphCenter
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevelBodyText
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle _
'        = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

Sub equ_PXL()
'
' ��ʽ��ʽ
'
'
    Dim style_name As String
'    Dim style_name_next As String
    style_name = Config.dict.Item(Config.equ_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 14
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevelBodyText
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
    Dim FirstTapPosition As Double
    Dim SecondTapPosition As Double
    Dim PageWidth  As Double 'ҳ��ʵ�ʿ��
    PageWidth = ActiveDocument.Sections(1).PageSetup.PageWidth _
                - ActiveDocument.Sections(1).PageSetup.LeftMargin _
                - ActiveDocument.Sections(1).PageSetup.RightMargin '��ͨ����һ���ж�
    FirstTapPosition = PageWidth / 2
    SecondTapPosition = PageWidth
    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.Add Position:= _
        FirstTapPosition, Alignment:=wdAlignTabCenter, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.Add Position:= _
        SecondTapPosition, Alignment:=wdAlignTabRight, Leader:= _
        wdTabLeaderSpaces
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
End Sub

'Sub ref_PXL()
'
' �ο�������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.ref_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 14
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0.99)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(-0.99)
'        .OutlineLevel = wdOutlineLevelBodyText
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = -2
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name).NoSpaceBetweenParagraphsOfSameStyle = _
'        False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.Add Position:= _
'        CentimetersToPoints(1), Alignment:=wdAlignTabLeft, Leader:= _
'        wdTabLeaderSpaces
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub appx_level1_PXL()
'
' ��¼1��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.appx_level1_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 22
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 6
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 6
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphCenter
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = True
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = 0
'        .OutlineLevel = wdOutlineLevel1
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name). _
'        NoSpaceBetweenParagraphsOfSameStyle = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub appx_level2_PXL()
'
' ��¼2��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.appx_level2_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 16
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = 0
'        .RightIndent = 0
'        .SpaceBefore = 6
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 6
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = 0
'        .OutlineLevel = wdOutlineLevel2
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name). _
'        NoSpaceBetweenParagraphsOfSameStyle = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub appx_level3_PXL()
'
' ��¼3��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.appx_level3_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 15
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = 0
'        .RightIndent = 0
'        .SpaceBefore = 3
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 3
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = 0
'        .OutlineLevel = wdOutlineLevel3
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name). _
'        NoSpaceBetweenParagraphsOfSameStyle = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub
'Sub appx_level4_PXL()
'
' ��¼4��������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.appx_level4_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 14
'        .Bold = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(3.5)
'        .RightIndent = 0
'        .SpaceBefore = 3
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphLeft
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(-3.5)
'        .OutlineLevel = wdOutlineLevel4
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name). _
'        NoSpaceBetweenParagraphsOfSameStyle = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops.ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub appx_tab_title_PXL()
'
' ��¼������ʽ
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.appx_tab_title_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 10.5
'        .Bold = True
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 2.5
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphCenter
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevel5
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0.5
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name). _
'        NoSpaceBetweenParagraphsOfSameStyle = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops. _
'        ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub appx_fig_title_PXL()
'
' ��¼ͼ����ʽ
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.appx_fig_title_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font
'        .NameFarEast = "����"
'        .NameAscii = "Times New Roman"
'        .NameOther = "Times New Roman"
'        .Name = "Times New Roman"
'        .Size = 10.5
'        .Bold = True
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
''        .Superscript = False
''        .Subscript = False
'        .Scaling = 100
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        .LeftIndent = CentimetersToPoints(0)
'        .RightIndent = CentimetersToPoints(0)
'        .SpaceBefore = 0
'        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
'        .SpaceAfterAuto = False
'        .LineSpacingRule = wdLineSpace1pt5
'        .Alignment = wdAlignParagraphCenter
'        .WidowControl = False
'        .KeepWithNext = False
'        .KeepTogether = False
'        .PageBreakBefore = False
'        .NoLineNumber = False
'        .Hyphenation = True
'        .FirstLineIndent = CentimetersToPoints(0)
'        .OutlineLevel = wdOutlineLevel5
'        .CharacterUnitLeftIndent = 0
'        .CharacterUnitRightIndent = 0
'        .CharacterUnitFirstLineIndent = 0
'        .LineUnitBefore = 0
'        .LineUnitAfter = 0
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
'        .AutoAdjustRightIndent = True
'        .DisableLineHeightGrid = False
'        .FarEastLineBreakControl = True
'        .WordWrap = True
'        .HangingPunctuation = True
'        .HalfWidthPunctuationOnTopOfLine = False
'        .AddSpaceBetweenFarEastAndAlpha = True
'        .AddSpaceBetweenFarEastAndDigit = True
'        .BaseLineAlignment = wdBaselineAlignAuto
'    End With
'    ActiveDocument.Styles(style_name). _
'        NoSpaceBetweenParagraphsOfSameStyle = False
'    ActiveDocument.Styles(style_name).ParagraphFormat.TabStops. _
'        ClearAll
'    With ActiveDocument.Styles(style_name).ParagraphFormat
'        With .Shading
'            .Texture = wdTextureNone
'            .ForegroundPatternColor = wdColorAutomatic
'            .BackgroundPatternColor = wdColorAutomatic
'        End With
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
'        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'        With .Borders
'            .DistanceFromTop = 1
'            .DistanceFromLeft = 4
'            .DistanceFromBottom = 1
'            .DistanceFromRight = 4
'            .Shadow = False
'        End With
'    End With
'    ActiveDocument.Styles(style_name).Frame.Delete
'End Sub

'Sub tab_PXL()
''
'' �����ʽ
''
''
'    ActiveDocument.Styles.Add Name:="tab_PXL", Type:=wdStyleTypeTable '��ӱ����ʽ
'    With ActiveDocument.Styles("tab_PXL").Font '���ñ������������
'        .NameFarEast = "����" '���������֣������ģ������µ�������Ϊ�����塱��
'        .NameAscii = "Times New Roman" '��ASCII�ַ���������Ӣ�ġ����ֵȣ��µ�������Ϊ��Times New Roman����
'        .NameOther = "Times New Roman" '���Ƕ��ǡ���ASCII�ַ������µ�����������Ϊ��Times New Roman����
'        .Name = "Times New Roman" 'ͳһ���������ַ������µ�����Ϊ��Times New Roman�������п�����ǰ�������ظ�����ȷ�����������Ƶ�һ���ԡ�
'        .Size = 10.5 '���������СΪ10.5����
''        .Bold = False '���ô���Ч����
''        .Italic = False '����б��Ч����
''        .Underline = wdUnderlineNone 'ȡ���»���Ч��������wdUnderlineNone��Word���õĳ�������ʾ���»��ߡ�
''        .UnderlineColor = wdColorAutomatic '�����»�����ɫΪ�Զ���ͨ��Ϊ�ı���ɫ������������һ���ѽ������»��ߣ��˴����ò�Ӱ��ʵ��Ч����
''        .StrikeThrough = False '����ɾ����Ч����
''        .DoubleStrikeThrough = False '����˫ɾ����Ч����
''        .Outline = False '�����������Ч����
''        .Emboss = False '���ø���Ч����
''        .Shadow = False '������ӰЧ����
''        .Hidden = False 'ȷ���ı��������ء�
''        .SmallCaps = False '����С�ʹ�д��ĸЧ����
''        .AllCaps = False '����ȫ��д��ĸЧ����
''        .Color = wdColorAutomatic '�����ı���ɫΪ�Զ���ͨ��Ϊ��ɫ��������ָ���ض���ɫ��
''        .Engrave = False '���õ��Ч����
''        .Superscript = False '�����ϱ�Ч����
''        .Subscript = False '�����±�Ч����
''        .Scaling = 100 '�����������ű���Ϊ100%���������ж������š�
''        .Kerning = 1 '�����ּ���������ż�ࣩΪ1 pt����ֵԽ�������ַ�֮��ľ���Խ��
''        .Animation = wdAnimationNone '���ö���Ч����
''        .DisableCharacterSpaceGrid = False '�����ַ����������ã�������Word����������ֺ��Զ������ַ���ࡣ
''        .EmphasisMark = wdEmphasisMarkNone '������κ�ǿ����ǣ������غŵȣ���
''        .Ligatures = wdLigaturesNone '�������֣���ff��fi��fl����ĸ��ϵ����������ʽ����
''        .NumberSpacing = wdNumberSpacingDefault 'ʹ��Ĭ�ϵ����ּ�����á�
''        .NumberForm = wdNumberFormDefault 'ʹ��Ĭ�ϵ�������ʽ���á�
''        .StylisticSet = wdStylisticSetDefault 'ʹ��Ĭ�ϵķ�񼯣�ĳЩOpenType�����ṩ�Ķ�����ʽѡ���
''        .ContextualAlternates = 0 '�ر������Ľ������Σ�ĳЩOpenType���������Χ�ַ��Զ�ѡ������ʵ����Σ���ֵΪ0��ʾ�����á�
'    End With
'    With ActiveDocument.Styles("tab_PXL").ParagraphFormat '���ñ���ж�������
''        .LeftIndent = CentimetersToPoints(0) '������������Ϊ0���ף�ת��Ϊ������
''        .RightIndent = CentimetersToPoints(0) '������������Ϊ0���ף�ת��Ϊ��������
''        .SpaceBefore = 0 '��ǰ�����Ϊ0����
''        .SpaceBeforeAuto = False '�ر��Զ���ǰ��ࡣ
''        .SpaceAfter = 0 '�κ�����Ϊ0����
''        .SpaceAfterAuto = False '�ر��Զ��κ��ࡣ
'        .LineSpacingRule = wdLineSpaceSingle '�����о����Ϊ�����оࡣ
'        .Alignment = wdAlignParagraphCenter '������뷽ʽ��Ϊ���ж��롣
''        .WidowControl = True '���ù��п��ƣ������ĩֻʣһ���г�������һҳ��������
''        .KeepWithNext = False '�رա����¶�ͬҳ��ѡ���������ҳ��
''        .KeepTogether = False '�رա����䲻��֡�ѡ���������ҳ��
''        .PageBreakBefore = False 'ȡ����ǰ��ҳ��
''        .NoLineNumber = False '�����к���ʾ������ĵ��������кŹ��ܣ���
''        .Hyphenation = True '�����Զ����ֹ��ܡ�
''        .FirstLineIndent = CentimetersToPoints(0) '����������Ϊ0���ף�ת��Ϊ��������
''        .OutlineLevel = wdOutlineLevelBodyText '�������趨Ϊ���ļ��𣨷Ǵ�ټ��𣩡�
''        .CharacterUnitLeftIndent = 0 '���ַ�Ϊ��λ����������Ϊ0��
''        .CharacterUnitRightIndent = 0 '���ַ�Ϊ��λ����������Ϊ0��
''        .CharacterUnitFirstLineIndent = 0 '���ַ�Ϊ��λ������������Ϊ0��
''        .LineUnitBefore = 0 '���и�Ϊ��λ�Ķ�ǰ�����Ϊ0��
''        .LineUnitAfter = 0 '���и�Ϊ��λ�Ķκ�����Ϊ0��
''        .MirrorIndents = False '�ر������������񣨼������������Ӱ���Ҳ���������
'        .TextboxTightWrap = wdTightNone '�ı�����ܻ�������Ϊ���޽��ܻ��ơ���
''        .CollapsedByDefault = False '��Ĭ���۵�����ʽ���䣨�ڴ����ͼ�У���
''        .AutoAdjustRightIndent = True '�����Զ������������������ı�����Ͷ��뷽ʽ����
''        .DisableLineHeightGrid = False '���رջ���������о������
''        .FarEastLineBreakControl = True '���ö������ԣ������ġ����ģ��Ļ��п��ơ�
''        .WordWrap = True '�����Զ����С�
''        .HangingPunctuation = True '�������ұ�㣨ĳЩ�����ų�������߽磩��
''        .HalfWidthPunctuationOnTopOfLine = False '��ֹ��Ǳ���������ס�
''        .AddSpaceBetweenFarEastAndAlpha = True '�ڶ�����������ĸ����Ӷ���ո�
''        .AddSpaceBetweenFarEastAndDigit = True '�ڶ������������ּ���Ӷ���ո�
''        .BaseLineAlignment = wdBaselineAlignAuto '���߶��뷽ʽ��Ϊ�Զ���
'    End With
'    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = False '����Ĭ�ϵĶμ�ࣨ����ȡ���μ�ࣩ
'    ActiveDocument.Styles("tab_PXL").ParagraphFormat.TabStops.ClearAll 'ɾ�����ж����Ѷ����ˮƽ����λ�ã�ʹ��ʹ�ô���ʽ�Ķ��䲻����֮ǰ���õ��Ʊ��Ӱ��
'    ActiveDocument.Styles("tab_PXL").Frame.Delete 'ʹ�ø���ʽ���ı��������ܵ���ܵ�Ӱ�졣
'    With ActiveDocument.Styles("tab_PXL").Table
'        .TableDirection = 1 '�����������Ϊ�����ң�ˮƽ���򣩡���Word�У�1����ˮƽ����0����ֱ���򣨴��ϵ��£���
'        .TopPadding = 0 '���ñ���ڲ���Ԫ������߿�֮����ϱ߾�Ϊ0���ס�CentimetersToPoints�������ڽ�����ֵת��ΪWordʹ�õĵ�����λ��
'        .BottomPadding = 0 '���ñ���ڲ���Ԫ������߿�֮����±߾�Ϊ0���ס�
'        .LeftPadding = CentimetersToPoints(0.1) '���ñ���ڲ���Ԫ������߿�֮�����߾�Ϊ0.1���ס�������Ԫ��� TopPadding �������ý������������ TopPadding �������á�
'        .RightPadding = CentimetersToPoints(0.1) '���ñ���ڲ���Ԫ������߿�֮����ұ߾�Ϊ0.1���ס�������Ԫ��� RightPadding �������ý������������ RightPadding �������á�
'        .Alignment = wdAlignRowCenter '������е����������ݾ��ж��롣
''        .Spacing = 0 '���ñ����֮��ļ��Ϊ0����������֮��û�ж���Ĵ�ֱ��ࡣ
'        .AllowPageBreaks = True '�������ҳ���С��������ڱ�Ҫʱ��ҳ��ʾ��
'        .AllowBreakAcrossPage = True '�������ҳ���С���������Ԫ������ݿ�Խҳ��߽硣
''        .LeftIndent = CentimetersToPoints(0) '���ñ������������ڶ����������Ϊ0���ף����������������롣
''        .RowStripe = 0 '���ñ����м����ƣ����汳��ɫ������Ϊ�ޣ�ֵΪ0������ȡ�����Ľ�����ɫЧ����
''        .ColumnStripe = 0 '���ñ����м����ƣ����汳��ɫ������Ϊ�ޣ�ֵΪ0������ȡ�����Ľ�����ɫЧ����
'    End With
'    With ActiveDocument.Styles("tab_PXL").Table
'        With .Shading
'            .Texture = wdTextureNone '����񱳾���������Ϊ����������ɫ��䣩��
'            .ForegroundPatternColor = wdColorAutomatic '�����ǰ��ͼ����ɫ����Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'            .BackgroundPatternColor = wdColorAutomatic '����񱳾�ͼ����ɫ����Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        With .Borders(wdBorderHorizontal) '����ˮƽ�߿�
'            .LineStyle = wdLineStyleSingle '���߿�������Ϊ��ʵ�ߡ�
'            .LineWidth = wdLineWidth050pt '���߿��߿���Ϊ0.5��������ϸ����
'            .Color = wdColorAutomatic '���߿���ɫ��Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        With .Borders(wdBorderVertical) '���ô�ֱ�߿�
'            .LineStyle = wdLineStyleSingle '���߿�������Ϊ��ʵ�ߡ�
'            .LineWidth = wdLineWidth050pt '���߿��߿���Ϊ0.5��������ϸ����
'            .Color = wdColorAutomatic '���߿���ɫ��Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        With .Borders(wdBorderLeft) '������߿�
'            .LineStyle = wdLineStyleSingle '���߿�������Ϊ��ʵ�ߡ�
'            .LineWidth = wdLineWidth150pt '���߿��߿���Ϊ1.5�����ϴ֣���
'            .Color = wdColorAutomatic '���߿���ɫ��Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        With .Borders(wdBorderRight) '�����ұ߿�
'            .LineStyle = wdLineStyleSingle '���߿�������Ϊ��ʵ�ߡ�
'            .LineWidth = wdLineWidth150pt '���߿��߿���Ϊ1.5�����ϴ֣���
'            .Color = wdColorAutomatic '���߿���ɫ��Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        With .Borders(wdBorderTop) '�����ϱ߿�
'            .LineStyle = wdLineStyleSingle '���߿�������Ϊ��ʵ�ߡ�
'            .LineWidth = wdLineWidth150pt '���߿��߿���Ϊ1.5�����ϴ֣���
'            .Color = wdColorAutomatic '���߿���ɫ��Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        With .Borders(wdBorderBottom) '�����±߿�
'            .LineStyle = wdLineStyleSingle '���߿�������Ϊ��ʵ�ߡ�
'            .LineWidth = wdLineWidth150pt '���߿��߿���Ϊ1.5�����ϴ֣���
'            .Color = wdColorAutomatic '���߿���ɫ��Ϊ�Զ���Ĭ��ϵͳ��ɫ����
'        End With
'        .Borders.Shadow = False 'ȡ�����߿���ӰЧ����
'    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstRow)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstRow)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstRow). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstRow).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstRow). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastRow)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastRow)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastRow). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastRow).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastRow). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstColumn)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstColumn)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstColumn). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstColumn).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdFirstColumn). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastColumn)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastColumn)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastColumn). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastColumn).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdLastColumn). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddColumnBanding)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddColumnBanding)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddColumnBanding). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddColumnBanding). _
''        Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddColumnBanding). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenColumnBanding)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenColumnBanding)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenColumnBanding) _
''        .ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenColumnBanding) _
''        .Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenColumnBanding). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddRowBanding)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddRowBanding)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddRowBanding). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddRowBanding). _
''        Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdOddRowBanding). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenRowBanding)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenRowBanding)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenRowBanding). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenRowBanding). _
''        Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdEvenRowBanding). _
''        ParagraphFormat.TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNECell)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNECell)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNECell). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNECell).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdNECell).ParagraphFormat _
''        .TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNWCell)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNWCell)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNWCell). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdNWCell).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdNWCell).ParagraphFormat _
''        .TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSECell)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSECell)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSECell). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSECell).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdSECell).ParagraphFormat _
''        .TabStops.ClearAll
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSWCell)
''        .TopPadding = CentimetersToPoints(0)
''        .BottomPadding = CentimetersToPoints(0)
''        .LeftPadding = CentimetersToPoints(0.1)
''        .RightPadding = CentimetersToPoints(0.1)
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSWCell)
''        With .Shading
''            .Texture = wdTextureNone
''            .ForegroundPatternColor = wdColorAutomatic
''            .BackgroundPatternColor = wdColorAutomatic
''        End With
''        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
''        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
''        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
''        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
''        .Borders.Shadow = False
''    End With
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSWCell). _
''        ParagraphFormat
''        .LeftIndent = CentimetersToPoints(0)
''        .RightIndent = CentimetersToPoints(0)
''        .SpaceBefore = 0
''        .SpaceBeforeAuto = False
''        .SpaceAfter = 0
''        .SpaceAfterAuto = False
''        .LineSpacingRule = wdLineSpaceSingle
''        .Alignment = wdAlignParagraphLeft
''        .WidowControl = True
''        .KeepWithNext = False
''        .KeepTogether = False
''        .PageBreakBefore = False
''        .NoLineNumber = False
''        .Hyphenation = True
''        .FirstLineIndent = CentimetersToPoints(0)
''        .OutlineLevel = wdOutlineLevelBodyText
''        .CharacterUnitLeftIndent = 0
''        .CharacterUnitRightIndent = 0
''        .CharacterUnitFirstLineIndent = 0
''        .LineUnitBefore = 0
''        .LineUnitAfter = 0
''        .MirrorIndents = False
''        .TextboxTightWrap = wdTightNone
''        .CollapsedByDefault = False
''        .AutoAdjustRightIndent = True
''        .DisableLineHeightGrid = False
''        .FarEastLineBreakControl = True
''        .WordWrap = True
''        .HangingPunctuation = True
''        .HalfWidthPunctuationOnTopOfLine = False
''        .AddSpaceBetweenFarEastAndAlpha = True
''        .AddSpaceBetweenFarEastAndDigit = True
''        .BaseLineAlignment = wdBaselineAlignAuto
''    End With
''    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = _
''        False
''    With ActiveDocument.Styles("tab_PXL").Table.Condition(wdSWCell).Font
''        .NameFarEast = ""
''        .NameAscii = ""
''        .NameOther = ""
''        .Name = ""
''    End With
''    ActiveDocument.Styles("tab_PXL").Table.Condition(wdSWCell).ParagraphFormat _
''        .TabStops.ClearAll
'End Sub




