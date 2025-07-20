Attribute VB_Name = "Style_Library"
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0
'
'本模块定义模板的各种格式

''''''部分格式属性说明如下''''''
'    With ActiveDocument.Styles("xxxxx").Font '字体属性设置
'        .NameFarEast = "黑体" '设置东亚字符集为“黑体”
'        .NameAscii = "Times New Roman" '设置西文文本字符集为“Times New Roman”
'        .NameOther = "Times New Roman" '设置其他字符集为“Times New Roman”
'        .Name = "Times New Roman" '设置所有字符集下的字体为“Times New Roman”
'        .Size = 22 '设定字体大小为22磅
'        .Bold = False '关闭粗体
'        .Italic = False '关闭斜体
'        .Underline = wdUnderlineNone '关闭下划线
'        .UnderlineColor = wdColorAutomatic '将下划线颜色设为自动（即跟随文本颜色）
'        .StrikeThrough = False '关闭删除线效果
'        .DoubleStrikeThrough = False '关闭双删除线效果
'        .Outline = False '关闭轮廓效果
'        .Emboss = False '关闭浮雕效果
'        .Shadow = False '关闭阴影效果
'        .Hidden = False '关闭隐藏文本效果
'        .SmallCaps = False '关闭小型大写字母效果
'        .AllCaps = False '关闭全部大写字母效果
'        .Color = wdColorAutomatic '设置字体颜色为自动（即跟随文档主题或父级样式的颜色）
'        .Engrave = False '关闭雕刻效果
''        .Superscript = False '关闭上标效果
''        .Subscript = False '关闭下标效果
'        .Scaling = 100 '设置字体缩放比例为100%（默认值）
'        .Kerning = 1 '设置字符间距为1pt
'        .Animation = wdAnimationNone '关闭文字动画效果
'        .DisableCharacterSpaceGrid = False '启用字符网格布局
'        .EmphasisMark = wdEmphasisMarkNone '关闭强调标记
'        .Ligatures = wdLigaturesNone '关闭连字
'        .NumberSpacing = wdNumberSpacingDefault '设定数字间距为默认
'        .NumberForm = wdNumberFormDefault '设定数字形式为默认
'        .StylisticSet = wdStylisticSetDefault '设定风格集为默认
'        .ContextualAlternates = 0 '并关闭上下文交替字体
'    End With
'    With ActiveDocument.Styles("xxxxx").ParagraphFormat '设置段落格式属性
'        .LeftIndent = 0 '设置段落的左缩进为0厘米
'        .RightIndent = 0 '设置段落的右缩进为0厘米
'        .SpaceBefore = 6 '设置段前间距为6磅
'        .SpaceBeforeAuto = False '关闭自动调整段前间距的功能
'        .SpaceAfter = 6 '设置段后间距为6磅
'        .SpaceAfterAuto = False '关闭自动调整段后间距的功能
'        .LineSpacingRule = wdLineSpace1pt5 '设置行距规则为1.5倍行距
'        .Alignment = wdAlignParagraphCenter '将段落文本对齐方式设置为居中对齐
'        .WidowControl = False '关闭孤行控制功能，即允许段尾只有一行出现在下一页或下一栏。
'        .KeepWithNext = False '不强制当前段落与其后的段落同页显示
'        .KeepTogether = False '允许段落跨页断开，不强制整个段落保持在同一页面
'        .PageBreakBefore = True '在该段落前插入分页符，使得段落始终从新的一页开始。
'        .NoLineNumber = False '不禁用行号显示，即该段落正常显示行号（如果文档已开启行号显示）
'        .Hyphenation = True '启用自动断词功能，允许Word在单词换行时插入连字符
'        .FirstLineIndent = 0 '设置首行缩进为0厘米。
'        .OutlineLevel = wdOutlineLevel1 '将段落设置为大纲级别1
'        .CharacterUnitLeftIndent = 0 '设置以字符为单位的左缩进为0个字符
'        .CharacterUnitRightIndent = 0 '设置以字符为单位的右缩进为0个字符
'        .CharacterUnitFirstLineIndent = 0 '设置以字符为单位的首行缩进为0个字符
'        .LineUnitBefore = 0 '设置以行为单位的段前间距为0行
'        .LineUnitAfter = 0 '设置以行为单位的段后间距为0行
'        .MirrorIndents = False '关闭镜像缩进，即不使子列表的缩进与父列表对称
'        .TextboxTightWrap = wdTightNone '设置文本框紧密环绕样式为无紧密环绕，即文本框与周围文本间保留标准间距
'        .CollapsedByDefault = False '不使该段落样式在大纲视图中默认折叠
'        .AutoAdjustRightIndent = True '允许Word根据段落的制表位和悬挂缩进来自动调整右缩进
'        .DisableLineHeightGrid = False '不禁用行高网格，即保持行距与文档的默认网格对齐
'        .FarEastLineBreakControl = True ' 对东亚语言启用特殊的行断字控制
'        .WordWrap = True '启用单词换行功能，即允许长单词跨越多行。
'        .HangingPunctuation = True '启用悬挂标点符号，使其超出常规文本边界，提高排版美观度
'        .HalfWidthPunctuationOnTopOfLine = False '禁止半角标点符号出现在行首
'        .AddSpaceBetweenFarEastAndAlpha = True '在东亚字符与拉丁字符之间添加额外空格
'        .AddSpaceBetweenFarEastAndDigit = True '在东亚字符与数字之间添加额外空格
'        .BaseLineAlignment = wdBaselineAlignAuto '设置基线对齐方式为自动，即默认的基线对齐方式
'    End With



Sub Multi_level_list()

' 正文多级列表


    Dim newListTemplate As ListTemplate
    Set newListTemplate = ActiveDocument.ListTemplates.Add(Name:="ListN_PXL") '创建列表，默认为单级
    newListTemplate.Convert (1) '转为多级列表

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
'        .LinkedStyle = Config.dict.Item(Config.level1_PXL) '标题1_PXL
    End With
    With newListTemplate.ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.86) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.level2_PXL) '标题2_PXL
    End With
    With newListTemplate.ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.2) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.level3_PXL) '标题3_PXL
    End With
    With newListTemplate.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.49) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.level4_PXL) '标题4_PXL
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
        .NumberFormat = "（%1.%7）"
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
        .NumberFormat = "图%1.%8"
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
'        .LinkedStyle = Config.dict.Item(Config.fig_title_PXL) '图题_PXL
    End With
    With newListTemplate.ListLevels(9)
        .NumberFormat = "表%1.%9"
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
'        .LinkedStyle = Config.dict.Item(Config.tab_title_PXL) '表题_PXL
    End With

End Sub

Sub Multi_level_list_APPX()

' 附录多级标题



    Dim newListTemplate As ListTemplate
    Set newListTemplate = ActiveDocument.ListTemplates.Add(Name:="ListA_PXL") '创建列表，默认为单级
    newListTemplate.Convert (1) '转为多级列表

    With newListTemplate.ListLevels(1)
        .NumberFormat = "附录%1"
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
'        .LinkedStyle = Config.dict.Item(Config.appx_level1_PXL) '附录1_PXL
    End With
    With newListTemplate.ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.99) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.appx_level2_PXL) '附录2_PXL
    End With
    With newListTemplate.ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.33) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.appx_level3_PXL) '附录3_PXL
    End With
    With newListTemplate.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.6) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.appx_level4_PXL) '附录4_PXL
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
        .NumberFormat = "图%1.%7"
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
'        .LinkedStyle = Config.dict.Item(Config.appx_fig_title_PXL) '附录图_PXL
    End With
    With newListTemplate.ListLevels(8)
        .NumberFormat = "表%1.%8"
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
'        .LinkedStyle = Config.dict.Item(Config.appx_tab_title_PXL) '附录表_PXL
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

' 参考文献列表


    Dim newListTemplate As ListTemplate
    Set newListTemplate = ActiveDocument.ListTemplates.Add(Name:="ListR_PXL") '创建列表，默认为单级

    With newListTemplate.ListLevels(1)
        .NumberFormat = "[%1]"
        .TrailingCharacter = wdTrailingTab '编号后 wdTrailingTab=制表位,wdTrailingSpace=空格
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1) '悬挂缩进
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
'        .LinkedStyle = Config.dict.Item(Config.ref_PXL) '参考文献_PXL
    End With
End Sub

'Sub normal_PXL()
'
' 正文样式
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
'        .NameFarEast = "仿宋"
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
' 正文1级标题样式
'
'
'    Dim style_name As String
'    Dim style_name_next As String
'    style_name = Config.dict.Item(Config.level1_PXL)
'    style_name_next = Config.dict.Item(Config.normal_PXL)
'    ActiveDocument.Styles.Add Name:=style_name, Type:=wdStyleTypeParagraph '添加段落样式
'    ActiveDocument.Styles(style_name).AutomaticallyUpdate = False '取消自动更新
'    ActiveDocument.Styles(style_name).BaseStyle = ""
'    ActiveDocument.Styles(style_name).NextParagraphStyle = ActiveDocument.Styles(style_name_next)
'    With ActiveDocument.Styles(style_name).Font '字体属性设置
'        .NameFarEast = "黑体" '设置东亚字符集为“黑体”
'        .NameAscii = "Times New Roman" '设置西文文本字符集为“Times New Roman”
'        .NameOther = "Times New Roman" '设置其他字符集为“Times New Roman”
'        .Name = "Times New Roman" '设置所有字符集下的字体为“Times New Roman”
'        .Size = 22 '设定字体大小为22磅
'        .Bold = False '关闭粗体
'        .Italic = False '关闭斜体
'        .Underline = wdUnderlineNone '关闭下划线
'        .UnderlineColor = wdColorAutomatic '将下划线颜色设为自动（即跟随文本颜色）
'        .StrikeThrough = False '关闭删除线效果
'        .DoubleStrikeThrough = False '关闭双删除线效果
'        .Outline = False '关闭轮廓效果
'        .Emboss = False '关闭浮雕效果
'        .Shadow = False '关闭阴影效果
'        .Hidden = False '关闭隐藏文本效果
'        .SmallCaps = False '关闭小型大写字母效果
'        .AllCaps = False '关闭全部大写字母效果
'        .Color = wdColorAutomatic '设置字体颜色为自动（即跟随文档主题或父级样式的颜色）
'        .Engrave = False '关闭雕刻效果
''        .Superscript = False '关闭上标效果
''        .Subscript = False '关闭下标效果
'        .Scaling = 100 '设置字体缩放比例为100%（默认值）
'        .Kerning = 1 '设置字符间距为1pt
'        .Animation = wdAnimationNone '关闭文字动画效果
'        .DisableCharacterSpaceGrid = False '启用字符网格布局
'        .EmphasisMark = wdEmphasisMarkNone '关闭强调标记
'        .Ligatures = wdLigaturesNone '关闭连字
'        .NumberSpacing = wdNumberSpacingDefault '设定数字间距为默认
'        .NumberForm = wdNumberFormDefault '设定数字形式为默认
'        .StylisticSet = wdStylisticSetDefault '设定风格集为默认
'        .ContextualAlternates = 0 '并关闭上下文交替字体
'    End With
'    With ActiveDocument.Styles(style_name).ParagraphFormat '设置段落格式属性
'        .LeftIndent = 0 '设置段落的左缩进为0厘米
'        .RightIndent = 0 '设置段落的右缩进为0厘米
'        .SpaceBefore = 6 '设置段前间距为6磅
'        .SpaceBeforeAuto = False '关闭自动调整段前间距的功能
'        .SpaceAfter = 6 '设置段后间距为6磅
'        .SpaceAfterAuto = False '关闭自动调整段后间距的功能
'        .LineSpacingRule = wdLineSpace1pt5 '设置行距规则为1.5倍行距
'        .Alignment = wdAlignParagraphCenter '将段落文本对齐方式设置为居中对齐
'        .WidowControl = False '关闭孤行控制功能，即允许段尾只有一行出现在下一页或下一栏。
'        .KeepWithNext = False '不强制当前段落与其后的段落同页显示
'        .KeepTogether = False '允许段落跨页断开，不强制整个段落保持在同一页面
'        .PageBreakBefore = True '在该段落前插入分页符，使得段落始终从新的一页开始。
'        .NoLineNumber = False '不禁用行号显示，即该段落正常显示行号（如果文档已开启行号显示）
'        .Hyphenation = True '启用自动断词功能，允许Word在单词换行时插入连字符
'        .FirstLineIndent = 0 '设置首行缩进为0厘米。
'        .OutlineLevel = wdOutlineLevel1 '将段落设置为大纲级别1
'        .CharacterUnitLeftIndent = 0 '设置以字符为单位的左缩进为0个字符
'        .CharacterUnitRightIndent = 0 '设置以字符为单位的右缩进为0个字符
'        .CharacterUnitFirstLineIndent = 0 '设置以字符为单位的首行缩进为0个字符
'        .LineUnitBefore = 0 '设置以行为单位的段前间距为0行
'        .LineUnitAfter = 0 '设置以行为单位的段后间距为0行
'        .MirrorIndents = False '关闭镜像缩进，即不使子列表的缩进与父列表对称
'        .TextboxTightWrap = wdTightNone '设置文本框紧密环绕样式为无紧密环绕，即文本框与周围文本间保留标准间距
'        .CollapsedByDefault = False '不使该段落样式在大纲视图中默认折叠
'        .AutoAdjustRightIndent = True '允许Word根据段落的制表位和悬挂缩进来自动调整右缩进
'        .DisableLineHeightGrid = False '不禁用行高网格，即保持行距与文档的默认网格对齐
'        .FarEastLineBreakControl = True ' 对东亚语言启用特殊的行断字控制
'        .WordWrap = True '启用单词换行功能，即允许长单词跨越多行。
'        .HangingPunctuation = True '启用悬挂标点符号，使其超出常规文本边界，提高排版美观度
'        .HalfWidthPunctuationOnTopOfLine = False '禁止半角标点符号出现在行首
'        .AddSpaceBetweenFarEastAndAlpha = True '在东亚字符与拉丁字符之间添加额外空格
'        .AddSpaceBetweenFarEastAndDigit = True '在东亚字符与数字之间添加额外空格
'        .BaseLineAlignment = wdBaselineAlignAuto '设置基线对齐方式为自动，即默认的基线对齐方式
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
' 正文2级标题样式
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
'        .NameFarEast = "宋体"
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
' 正文3级标题样式
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
'        .NameFarEast = "宋体"
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
' 正文4级标题样式
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
'        .NameFarEast = "宋体"
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
' 正文表题样式
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
'        .NameFarEast = "宋体"
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
' 正文图题样式
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
'        .NameFarEast = "宋体"
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
' 表文字样式样式
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
'        .NameFarEast = "宋体"
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
' 公式样式
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
'        .NameFarEast = "仿宋"
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
    Dim PageWidth  As Double '页面实际宽度
    PageWidth = ActiveDocument.Sections(1).PageSetup.PageWidth _
                - ActiveDocument.Sections(1).PageSetup.LeftMargin _
                - ActiveDocument.Sections(1).PageSetup.RightMargin '仅通过第一节判断
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
' 参考文献样式
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
'        .NameFarEast = "仿宋"
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
' 附录1级标题样式
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
'        .NameFarEast = "黑体"
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
' 附录2级标题样式
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
'        .NameFarEast = "宋体"
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
' 附录3级标题样式
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
'        .NameFarEast = "宋体"
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
' 附录4级标题样式
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
'        .NameFarEast = "仿宋"
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
' 附录表题样式
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
'        .NameFarEast = "宋体"
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
' 附录图题样式
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
'        .NameFarEast = "宋体"
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
'' 表格样式
''
''
'    ActiveDocument.Styles.Add Name:="tab_PXL", Type:=wdStyleTypeTable '添加表格样式
'    With ActiveDocument.Styles("tab_PXL").Font '设置表格中字体属性
'        .NameFarEast = "宋体" '将东亚文字（如中文）环境下的字体设为“宋体”。
'        .NameAscii = "Times New Roman" '将ASCII字符环境（如英文、数字等）下的字体设为“Times New Roman”。
'        .NameOther = "Times New Roman" '将非东亚、非ASCII字符环境下的字体名称设为“Times New Roman”。
'        .Name = "Times New Roman" '统一设置所有字符环境下的字体为“Times New Roman”。此行可能与前面三行重复，但确保了字体名称的一致性。
'        .Size = 10.5 '设置字体大小为10.5磅。
''        .Bold = False '禁用粗体效果。
''        .Italic = False '禁用斜体效果。
''        .Underline = wdUnderlineNone '取消下划线效果，其中wdUnderlineNone是Word内置的常量，表示无下划线。
''        .UnderlineColor = wdColorAutomatic '设置下划线颜色为自动（通常为文本颜色），但由于上一行已禁用了下划线，此处设置不影响实际效果。
''        .StrikeThrough = False '禁用删除线效果。
''        .DoubleStrikeThrough = False '禁用双删除线效果。
''        .Outline = False '禁用外框轮廓效果。
''        .Emboss = False '禁用浮雕效果。
''        .Shadow = False '禁用阴影效果。
''        .Hidden = False '确保文本不被隐藏。
''        .SmallCaps = False '禁用小型大写字母效果。
''        .AllCaps = False '禁用全大写字母效果。
''        .Color = wdColorAutomatic '设置文本颜色为自动（通常为黑色），即不指定特定颜色。
''        .Engrave = False '禁用雕刻效果。
''        .Superscript = False '禁用上标效果。
''        .Subscript = False '禁用下标效果。
''        .Scaling = 100 '设置字体缩放比例为100%，即不进行额外缩放。
''        .Kerning = 1 '设置字间距调整（字偶距）为1 pt。数值越大，相邻字符之间的距离越宽。
''        .Animation = wdAnimationNone '禁用动画效果。
''        .DisableCharacterSpaceGrid = False '保持字符网格功能启用，即允许Word根据字体和字号自动调整字符间距。
''        .EmphasisMark = wdEmphasisMarkNone '不添加任何强调标记（如着重号等）。
''        .Ligatures = wdLigaturesNone '禁用连字（如ff、fi、fl等字母组合的特殊呈现形式）。
''        .NumberSpacing = wdNumberSpacingDefault '使用默认的数字间距设置。
''        .NumberForm = wdNumberFormDefault '使用默认的数字形式设置。
''        .StylisticSet = wdStylisticSetDefault '使用默认的风格集（某些OpenType字体提供的额外样式选项）。
''        .ContextualAlternates = 0 '关闭上下文交替字形（某些OpenType字体根据周围字符自动选择更合适的字形）。值为0表示不启用。
'    End With
'    With ActiveDocument.Styles("tab_PXL").ParagraphFormat '设置表格中段落属性
''        .LeftIndent = CentimetersToPoints(0) '将左缩进设置为0厘米（转换为点数）
''        .RightIndent = CentimetersToPoints(0) '将右缩进设置为0厘米（转换为点数）。
''        .SpaceBefore = 0 '段前间距设为0磅。
''        .SpaceBeforeAuto = False '关闭自动段前间距。
''        .SpaceAfter = 0 '段后间距设为0磅。
''        .SpaceAfterAuto = False '关闭自动段后间距。
'        .LineSpacingRule = wdLineSpaceSingle '设置行距规则为单倍行距。
'        .Alignment = wdAlignParagraphCenter '段落对齐方式设为居中对齐。
''        .WidowControl = True '启用孤行控制（避免段末只剩一两行出现在下一页顶部）。
''        .KeepWithNext = False '关闭“与下段同页”选项，允许段落分页。
''        .KeepTogether = False '关闭“段落不拆分”选项，允许段落跨页。
''        .PageBreakBefore = False '取消段前分页。
''        .NoLineNumber = False '开启行号显示（如果文档已启用行号功能）。
''        .Hyphenation = True '启用自动断字功能。
''        .FirstLineIndent = CentimetersToPoints(0) '首行缩进设为0厘米（转换为点数）。
''        .OutlineLevel = wdOutlineLevelBodyText '将段落设定为正文级别（非大纲级别）。
''        .CharacterUnitLeftIndent = 0 '以字符为单位的左缩进设为0。
''        .CharacterUnitRightIndent = 0 '以字符为单位的右缩进设为0。
''        .CharacterUnitFirstLineIndent = 0 '以字符为单位的首行缩进设为0。
''        .LineUnitBefore = 0 '以行高为单位的段前间距设为0。
''        .LineUnitAfter = 0 '以行高为单位的段后间距设为0。
''        .MirrorIndents = False '关闭左右缩进镜像（即左侧缩进不会影响右侧缩进）。
'        .TextboxTightWrap = wdTightNone '文本框紧密环绕设置为“无紧密环绕”。
''        .CollapsedByDefault = False '不默认折叠该样式段落（在大纲视图中）。
''        .AutoAdjustRightIndent = True '启用自动调整右缩进（根据文本方向和对齐方式）。
''        .DisableLineHeightGrid = False '不关闭基于网格的行距调整。
''        .FarEastLineBreakControl = True '启用东亚语言（如中文、日文）的换行控制。
''        .WordWrap = True '启用自动换行。
''        .HangingPunctuation = True '启用悬挂标点（某些标点符号超出常规边界）。
''        .HalfWidthPunctuationOnTopOfLine = False '禁止半角标点置于行首。
''        .AddSpaceBetweenFarEastAndAlpha = True '在东亚文字与字母间添加额外空格。
''        .AddSpaceBetweenFarEastAndDigit = True '在东亚文字与数字间添加额外空格。
''        .BaseLineAlignment = wdBaselineAlignAuto '基线对齐方式设为自动。
'    End With
'    ActiveDocument.Styles("tab_PXL").NoSpaceBetweenParagraphsOfSameStyle = False '保留默认的段间距（即不取消段间距）
'    ActiveDocument.Styles("tab_PXL").ParagraphFormat.TabStops.ClearAll '删除所有段落已定义的水平对齐位置，使得使用此样式的段落不再受之前设置的制表符影响
'    ActiveDocument.Styles("tab_PXL").Frame.Delete '使用该样式的文本将不再受到框架的影响。
'    With ActiveDocument.Styles("tab_PXL").Table
'        .TableDirection = 1 '将表格方向设置为从左到右（水平方向）。在Word中，1代表水平方向，0代表垂直方向（从上到下）。
'        .TopPadding = 0 '设置表格内部单元格与表格边框之间的上边距为0厘米。CentimetersToPoints函数用于将厘米值转换为Word使用的点数单位。
'        .BottomPadding = 0 '设置表格内部单元格与表格边框之间的下边距为0厘米。
'        .LeftPadding = CentimetersToPoints(0.1) '设置表格内部单元格与表格边框之间的左边距为0.1厘米。单个单元格的 TopPadding 属性设置将覆盖整个表的 TopPadding 属性设置。
'        .RightPadding = CentimetersToPoints(0.1) '设置表格内部单元格与表格边框之间的右边距为0.1厘米。单个单元格的 RightPadding 属性设置将覆盖整个表的 RightPadding 属性设置。
'        .Alignment = wdAlignRowCenter '将表格中的所有行内容居中对齐。
''        .Spacing = 0 '设置表格行之间的间距为0，即相邻行之间没有额外的垂直间距。
'        .AllowPageBreaks = True '允许表格跨页断行。允许表格在必要时分页显示。
'        .AllowBreakAcrossPage = True '允许表格跨页断行。允许单个单元格的内容跨越页面边界。
''        .LeftIndent = CentimetersToPoints(0) '设置表格相对于其所在段落的左缩进为0厘米，即表格与段落左侧对齐。
''        .RowStripe = 0 '设置表格的行间条纹（交替背景色）设置为无（值为0），即取消表格的交替颜色效果。
''        .ColumnStripe = 0 '设置表格的列间条纹（交替背景色）设置为无（值为0），即取消表格的交替颜色效果。
'    End With
'    With ActiveDocument.Styles("tab_PXL").Table
'        With .Shading
'            .Texture = wdTextureNone '将表格背景纹理设置为无纹理（即纯色填充）。
'            .ForegroundPatternColor = wdColorAutomatic '将表格前景图案颜色设置为自动（默认系统颜色）。
'            .BackgroundPatternColor = wdColorAutomatic '将表格背景图案颜色设置为自动（默认系统颜色）。
'        End With
'        With .Borders(wdBorderHorizontal) '设置水平边框。
'            .LineStyle = wdLineStyleSingle '将边框线型设为单实线。
'            .LineWidth = wdLineWidth050pt '将边框线宽设为0.5磅。（较细）。
'            .Color = wdColorAutomatic '将边框颜色设为自动（默认系统颜色）。
'        End With
'        With .Borders(wdBorderVertical) '设置垂直边框。
'            .LineStyle = wdLineStyleSingle '将边框线型设为单实线。
'            .LineWidth = wdLineWidth050pt '将边框线宽设为0.5磅。（较细）。
'            .Color = wdColorAutomatic '将边框颜色设为自动（默认系统颜色）。
'        End With
'        With .Borders(wdBorderLeft) '设置左边框。
'            .LineStyle = wdLineStyleSingle '将边框线型设为单实线。
'            .LineWidth = wdLineWidth150pt '将边框线宽设为1.5磅（较粗）。
'            .Color = wdColorAutomatic '将边框颜色设为自动（默认系统颜色）。
'        End With
'        With .Borders(wdBorderRight) '设置右边框。
'            .LineStyle = wdLineStyleSingle '将边框线型设为单实线。
'            .LineWidth = wdLineWidth150pt '将边框线宽设为1.5磅（较粗）。
'            .Color = wdColorAutomatic '将边框颜色设为自动（默认系统颜色）。
'        End With
'        With .Borders(wdBorderTop) '设置上边框。
'            .LineStyle = wdLineStyleSingle '将边框线型设为单实线。
'            .LineWidth = wdLineWidth150pt '将边框线宽设为1.5磅（较粗）。
'            .Color = wdColorAutomatic '将边框颜色设为自动（默认系统颜色）。
'        End With
'        With .Borders(wdBorderBottom) '设置下边框。
'            .LineStyle = wdLineStyleSingle '将边框线型设为单实线。
'            .LineWidth = wdLineWidth150pt '将边框线宽设为1.5磅（较粗）。
'            .Color = wdColorAutomatic '将边框颜色设为自动（默认系统颜色）。
'        End With
'        .Borders.Shadow = False '取消表格边框阴影效果。
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




