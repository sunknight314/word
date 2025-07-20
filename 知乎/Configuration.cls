VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This program is used for MS-Word automatic layout.
' This file is copyrighted by Xianglu Pang.
' Email:460386414@qq.com
' This program is free software.
' It is distributed under the GNU General Public License (GPL) version 2.0
' Version: Beta version 3.0

Private m_normal_PXL As Integer
Private m_level1_PXL As Integer
Private m_level2_PXL As Integer
Private m_level3_PXL As Integer
Private m_level4_PXL As Integer
Private m_fig_title_PXL As Integer
Private m_tab_title_PXL As Integer
Private m_tab_text_PXL As Integer
Private m_equ_PXL As Integer
Private m_ref_PXL As Integer
Private m_appx_level1_PXL As Integer
Private m_appx_level2_PXL As Integer
Private m_appx_level3_PXL As Integer
Private m_appx_level4_PXL As Integer
Private m_appx_fig_title_PXL As Integer
Private m_appx_tab_title_PXL As Integer
Private m_tab_PXL As Integer

Private m_nLength As Integer

Private m_dict As Object

Private Sub Class_Initialize()
    m_normal_PXL = 1
    m_level1_PXL = 2
    m_level2_PXL = 3
    m_level3_PXL = 4
    m_level4_PXL = 5
    m_fig_title_PXL = 6
    m_tab_title_PXL = 7
    m_tab_text_PXL = 8
    m_equ_PXL = 9
    m_ref_PXL = 10
    m_appx_level1_PXL = 11
    m_appx_level2_PXL = 12
    m_appx_level3_PXL = 13
    m_appx_level4_PXL = 14
    m_appx_fig_title_PXL = 15
    m_appx_tab_title_PXL = 16
    m_tab_PXL = 17
    
    m_nLength = 17

    Set m_dict = CreateObject("Scripting.Dictionary")
    m_dict.Add 1, "正文_PXL"
    m_dict.Add 2, "标题1_PXL"
    m_dict.Add 3, "标题2_PXL"
    m_dict.Add 4, "标题3_PXL"
    m_dict.Add 5, "标题4_PXL"
    m_dict.Add 6, "图题_PXL"
    m_dict.Add 7, "表题_PXL"
    m_dict.Add 8, "表文字_PXL"
    m_dict.Add 9, "公式_PXL"
    m_dict.Add 10, "参考文献_PXL"
    m_dict.Add 11, "附录1_PXL"
    m_dict.Add 12, "附录2_PXL"
    m_dict.Add 13, "附录3_PXL"
    m_dict.Add 14, "附录4_PXL"
    m_dict.Add 15, "附录图_PXL"
    m_dict.Add 16, "附录表_PXL"
    m_dict.Add 17, "tab_PXL"
End Sub

Public Property Get normal_PXL() As Integer
    normal_PXL = m_normal_PXL
End Property
Public Property Get level1_PXL() As Integer
    level1_PXL = m_level1_PXL
End Property
Public Property Get level2_PXL() As Integer
    level2_PXL = m_level2_PXL
End Property
Public Property Get level3_PXL() As Integer
    level3_PXL = m_level3_PXL
End Property
Public Property Get level4_PXL() As Integer
    level4_PXL = m_level4_PXL
End Property
Public Property Get fig_title_PXL() As Integer
    fig_title_PXL = m_fig_title_PXL
End Property
Public Property Get tab_title_PXL() As Integer
    tab_title_PXL = m_tab_title_PXL
End Property
Public Property Get tab_text_PXL() As Integer
    tab_text_PXL = m_tab_text_PXL
End Property
Public Property Get equ_PXL() As Integer
    equ_PXL = m_equ_PXL
End Property
Public Property Get ref_PXL() As Integer
    ref_PXL = m_ref_PXL
End Property
Public Property Get appx_level1_PXL() As Integer
    appx_level1_PXL = m_appx_level1_PXL
End Property
Public Property Get appx_level2_PXL() As Integer
    appx_level2_PXL = m_appx_level2_PXL
End Property
Public Property Get appx_level3_PXL() As Integer
    appx_level3_PXL = m_appx_level3_PXL
End Property
Public Property Get appx_level4_PXL() As Integer
    appx_level4_PXL = m_appx_level4_PXL
End Property
Public Property Get appx_fig_title_PXL() As Integer
    appx_fig_title_PXL = m_appx_fig_title_PXL
End Property
Public Property Get appx_tab_title_PXL() As Integer
    appx_tab_title_PXL = m_appx_tab_title_PXL
End Property
Public Property Get tab_PXL() As Integer
    tab_PXL = m_tab_PXL
End Property
Public Property Get nLength() As Integer
    nLength = m_nLength
End Property
Public Property Get dict() As Object
    Set dict = m_dict
End Property
'

'Public Property Let normal_PXL(ByVal Value As Integer)
'    m_normal_PXL = Value
'End Property
'Public Property Let level1_PXL(ByVal Value As Integer)
'    m_level1_PXL = Value
'End Property
'Public Property Let level2_PXL(ByVal Value As Integer)
'    m_level2_PXL = Value
'End Property
'Public Property Let level3_PXL(ByVal Value As Integer)
'    m_level3_PXL = Value
'End Property
'Public Property Let level4_PXL(ByVal Value As Integer)
'    m_level4_PXL = Value
'End Property
'Public Property Let fig_title_PXL(ByVal Value As Integer)
'    m_fig_title_PXL = Value
'End Property
'Public Property Let tab_title_PXL(ByVal Value As Integer)
'    m_tab_title_PXL = Value
'End Property
'Public Property Let tab_text_PXL(ByVal Value As Integer)
'    m_tab_text_PXL = Value
'End Property
'Public Property Let equ_PXL(ByVal Value As Integer)
'    m_equ_PXL = Value
'End Property
'Public Property Let ref_PXL(ByVal Value As Integer)
'    m_ref_PXL = Value
'End Property
'Public Property Let appx_level1_PXL(ByVal Value As Integer)
'    m_appx_level1_PXL = Value
'End Property
'Public Property Let appx_level2_PXL(ByVal Value As Integer)
'    m_appx_level2_PXL = Value
'End Property
'Public Property Let appx_level3_PXL(ByVal Value As Integer)
'    m_appx_level3_PXL = Value
'End Property
'Public Property Let appx_level4_PXL(ByVal Value As Integer)
'    m_appx_level4_PXL = Value
'End Property
'Public Property Let appx_fig_title_PXL(ByVal Value As Integer)
'    m_appx_fig_title_PXL = Value
'End Property
'Public Property Let appx_tab_title_PXL(ByVal Value As Integer)
'    m_appx_tab_title_PXL = Value
'End Property
'Public Property Let tab_PXL(ByVal Value As Integer)
'    m_tab_PXL = Value
'End Property
'Public Property Let nLength(ByVal Value As Integer)
'    m_nLength = Value
'End Property
'Public Property Let dict(ByVal Value As Object)
'    Set m_dict = Value
'End Property


