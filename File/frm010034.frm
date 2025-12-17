VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010034 
   BorderStyle     =   1  '單線固定
   Caption         =   "圖書基本資料維護"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7536
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   7536
   Begin VB.ComboBox cboClass 
      Height          =   300
      ItemData        =   "frm010034.frx":0000
      Left            =   1080
      List            =   "frm010034.frx":0016
      TabIndex        =   2
      Text            =   "cboClass"
      Top             =   1020
      Width           =   1680
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "借閱記錄"
      Height          =   400
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   720
      Width           =   950
   End
   Begin VB.ComboBox cboStatus 
      Height          =   300
      ItemData        =   "frm010034.frx":0044
      Left            =   1080
      List            =   "frm010034.frx":0046
      TabIndex        =   11
      Text            =   "cboStatus"
      Top             =   3300
      Width           =   1680
   End
   Begin VB.CommandButton CmdPaper 
      Caption         =   "封面"
      Height          =   450
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   700
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   600
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":0048
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":0364
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":0680
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":085C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":0B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":0E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":11B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":14CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":17E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":1B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010034.frx":1E20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   7536
      _ExtentX        =   13293
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   690
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   4
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "(流水號)"
      Height          =   255
      Left            =   1965
      TabIndex        =   17
      Top             =   720
      Width           =   1020
   End
   Begin MSForms.ComboBox cboBK10 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   3000
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   270
      Index           =   13
      Left            =   4320
      TabIndex        =   12
      Top             =   3300
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   270
      Index           =   11
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Width           =   1200
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   9
      Left            =   4350
      TabIndex        =   3
      Top             =   1020
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   2670
      Width           =   6195
      VariousPropertyBits=   679493659
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   2340
      Width           =   6195
      VariousPropertyBits=   679493659
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   2010
      Width           =   6195
      VariousPropertyBits=   679493659
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   6195
      VariousPropertyBits=   679493659
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   800
      Index           =   14
      Left            =   1080
      TabIndex        =   22
      Top             =   3960
      Width           =   6200
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "10936;1411"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   2
      Left            =   4350
      TabIndex        =   1
      Top             =   690
      Width           =   1995
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "3519;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TxtBK 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   1350
      Width           =   6195
      VariousPropertyBits=   679493659
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblLR 
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   37
      Top             =   3630
      Width           =   855
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   960
   End
   Begin MSForms.Label LblLR 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   35
      Top             =   3630
      Width           =   1200
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "借閱人員："
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   34
      Top             =   3630
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "借閱日期："
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   33
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "出刊日期："
      Height          =   255
      Index           =   14
      Left            =   3360
      TabIndex        =   31
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "狀　　態："
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   30
      Top             =   3300
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "保管單位："
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   29
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "保管人員："
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "上架日期："
      Height          =   255
      Index           =   7
      Left            =   3300
      TabIndex        =   27
      Top             =   1050
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "譯　　者："
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "作者 (外)："
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   25
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "作者 (中)："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "書名 (外)："
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label Label23 
      Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4800
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "備　　註："
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "ＩＳＢＮ："
      Height          =   255
      Index           =   1
      Left            =   3285
      TabIndex        =   16
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "書名 (中)："
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "類　　別："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "圖書編號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frm010034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 TxtBK()全部/cboBK10/LblLR()全部
'2016/10/03 Create by Amy
Option Explicit

'宣告欄位內容結構
Private Type FieldData
   fiName As String
   fiOldData As String
   fiType As Integer
End Type

Dim m_FieldList() As FieldData
Dim i As Integer, BD_F As Integer
Dim m_PrevForm As Form '前畫面
Dim EditMode As Integer '0:Add 1:Update 2:Del 3:Cancel 5:Query
Dim m_AttachPath As String
'執行各項功能的權限
Dim m_bInsert As Boolean, m_bUpdate As Boolean, m_bDelete As Boolean
Dim m_FirstKEY As String, m_LastKEY As String, m_CurrKEY As String '第一筆/最後一筆/目前顯示 資料
'Dim bolBK10Click As Boolean '保管人是否是選的 '2025/01/10 不使用

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cboBK10_Click()
    If EditMode <> 0 And EditMode <> 1 Then Exit Sub
    'bolBK10Click = True 'Mark by Amy 2025/01/10 不使用
End Sub

'Modify by Amy 2021/12/01 原:Integer
Private Sub cboBK10_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If EditMode <> 0 And EditMode <> 1 Then Exit Sub
    KeyAscii = UpperCase(KeyAscii)
    'bolBK10Click = False 'Mark by Amy 2025/01/10 不使用
End Sub

Private Sub cboBK10_Validate(Cancel As Boolean)
    Dim stST01 As String, stST02 As String
    
    If EditMode <> 0 And EditMode <> 1 Then Exit Sub
    
    'Memo by Amy 保管人 可輸員編 or 姓名
    'Modify by Amy 2025/01/10 改Form2.0 元件後,輸完員編會預帶下拉選單中有的資料,導致員編+姓名查資料會錯
    If Trim(cboBK10) = MsgText(601) Then Exit Sub
    If ByInputGetST01or02(cboBK10, stST01, stST02) = False Then
         Cancel = True
         cboBK10.SetFocus
         Exit Sub
    End If
    If cboBK10 <> stST01 & " " & stST02 Then
         cboBK10 = stST01 & " " & stST02
    End If
    'end 2025/01/10
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    If EditMode <> 0 Or EditMode <> 1 Then Exit Sub
    KeyAscii = 0 '設只能選
End Sub

Private Sub cboStatus_Validate(Cancel As Boolean)
    If EditMode <> 0 And EditMode <> 1 Then Exit Sub
    
    Select Case cboStatus
        Case "遺失", "銷毀"
        Case Else
    End Select
End Sub

Private Sub cmdHistory_Click()
    If frm010035_3.QueryRecord(TxtBK(1)) = True Then
        Me.Hide
        frm010035_3.strPreFormName = Me.Name
        frm010035_3.Show
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        Exit Sub
    End If
End Sub

Private Sub CmdPaper_Click()
    Dim hLocalFile As Long
    Dim stFileName As String
    
    If TxtBK(1) = MsgText(601) Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If GetAttachFile(m_AttachPath) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    stFileName = m_AttachPath & "\" & TxtBK(1) & ".pdf"
    ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
    Dim oText
    
    strExc(0) = "Select * From BooksData Where RowNum<1"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    BD_F = RsTemp.Fields.Count
    ReDim m_FieldList(1 To BD_F) As FieldData
    
    For i = 1 To BD_F
        m_FieldList(i).fiName = "BK" & Format(i, "00")

        'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
        'If RsTemp.Fields(m_FieldList(i).fiName).Type = 200 Then
            m_FieldList(i).fiType = 0
        'Else
        '    m_FieldList(i).fiType = 1
        'End If
        'end 2017/06/29
    Next i
End Sub

'Mark by Amy 2021/12/08 原程式搬至Form_KeyU
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim SeqNo As String, bolUpdStatus As Boolean

    'Add by Amy 2021/12/01 從Form_KeyDown搬來
    If KeyCode = vbKeyF2 And m_bInsert = False Then Exit Sub
    If KeyCode = vbKeyF3 And m_bUpdate = False Then Exit Sub
    If KeyCode = vbKeyF5 And m_bDelete = False Then Exit Sub

    Select Case KeyCode
        Case vbKeyF2 'Add
            TxtBK(2).SetFocus
            TxtBK(1).TabStop = False
            RbEdit 0
        Case vbKeyF3 'Upd
            TxtBK(2).SetFocus
            TxtBK(1).TabStop = False
            RbEdit 1
        Case vbKeyF5 'Del
            '已有借閱記錄不可刪(非系統產生的一筆)
            If ChkNotDelete(TxtBK(1)) = True Then
                MsgBox "有借閱記錄不可刪除！", vbInformation
                Exit Sub
            End If
            If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                EditMode = 2: SeqNo = TxtBK(1)
                If ActionRecord = True Then
                    '只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
                    If SeqNo = m_LastKEY Or SeqNo = m_FirstKEY Then
                       RefreshRange
                    End If
                    QueryRecord m_CurrKEY
                    ToolBarSet 1
                    EditMode = 3
                Else
                    Exit Sub
                End If
            End If
        Case vbKeyF4 'Query
            RbEdit 5
        Case vbKeyHome 'MoveFirst
             If Not (EditMode = 0 Or EditMode = 1) Then
                TxtLock 0
                GetRecord 0
             End If
        Case vbKeyPageUp 'MovePre
             If Not (EditMode = 0 Or EditMode = 1) Then
                TxtLock 0
                If m_CurrKEY = m_FirstKEY Then
                    ShowMsg MsgText(9008)
                    Exit Sub
                End If
                GetRecord 1
             End If
        Case vbKeyPageDown 'MoveNext
             If Not (EditMode = 0 Or EditMode = 1) Then
                TxtLock 0
                If m_CurrKEY = m_LastKEY Then
                    ShowMsg MsgText(9009)
                    Exit Sub
                End If
                GetRecord 2
             End If
        Case vbKeyEnd 'MoveLast
             If Not (EditMode = 0 Or EditMode = 1) Then
                  TxtLock 0
                  GetRecord 3
             End If
        Case vbKeyF9 'OK
            '檢查欄位有效性
            If TxtValidate = False Then Exit Sub

            '在 新增/修改 狀態按Enter鍵
            If EditMode = 0 Or EditMode = 1 Then
                If EditMode = 1 Then
                    If m_FieldList(10).fiOldData <> Left(Trim(cboBK10), 5) Then
                        '修改保管人判斷原保管人是否有歸還記錄
                        If ChkBK10GiveBack(m_FieldList(10).fiOldData, SeqNo) = True Then
                            If MsgBox("原保管人有借閱中的記錄是否一同更新？", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbNo Then
                                SeqNo = ""
                            End If
                        End If
                    Else
                        '修改狀態成 遺失或銷毀 則新增一筆借閱記錄
                        If cboStatus.Tag <> cboStatus And Not (cboStatus = "遺失" Or cboStatus = "銷毀") Then
                            MsgBox "修改狀態有誤,請確認 ！"
                            Exit Sub
                        'Modify by Amy 2017/02/07 避免未修改時新增一筆銷毀記錄
                        ElseIf cboStatus = "遺失" Or cboStatus = "銷毀" Then
                            bolUpdStatus = True
                        End If
                    End If
                End If
                If ActionRecord(bolUpdStatus, SeqNo) = True Then
                    RefreshRange
                End If
            '在查詢狀態按Enter鍵
            ElseIf EditMode = 5 Then
                If QueryRecord(TxtBK(1), TxtBK(2)) = False Then
                    MsgBox "查無資料！", vbInformation
                End If
            End If
            QueryRecord m_CurrKEY
            ToolBarSet 1
            EditMode = 3

        Case vbKeyF10 'Cancel
            If EditMode <> 5 Then
                If MsgBox("並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
                    If EditMode = 0 Then
                        TxtLock 0
                        GetRecord 3
                    ElseIf EditMode = 1 Then
                        QueryRecord m_CurrKEY
                    End If
                    ToolBarSet 1
                    EditMode = 3 'Cancel
                Else
                    Exit Sub
                End If
            Else
                ToolBarSet 1
                EditMode = 3 'Cancel
                QueryRecord m_CurrKEY
            End If
            TxtBK(1).SetFocus
        Case vbKeyEscape 'Exit
            If m_PrevForm Is Nothing Then
                Unload Me
            Else
                tmpBol = fnCancelNowFormAndShowParentForm(Me)
            End If
    End Select
End Sub

Private Sub Form_Load()
    If m_PrevForm Is Nothing Then
        '取得使用者執行各項功能的權限
        m_bInsert = IsUserHasRightOfFunction("frm010034", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm010034", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm010034", strDel, False)
        ToolBarSet 1     '設定ToolBar按鈕顯示
        EditMode = 3
    End If
    MoveFormToCenter Me
    
    ClearField
    
    m_AttachPath = App.path & "\BooksAttach"
    If Dir(m_AttachPath, vbDirectory) = MsgText(601) Then MkDir m_AttachPath
    setCboStatus
    setCboBK10
    RefreshRange
    GetRecord 0     '設定第一筆key值
    
End Sub

Private Sub setCboStatus()
    cboStatus.AddItem ""
    cboStatus.AddItem "遺失"
    cboStatus.AddItem "銷毀"
End Sub

Private Sub setCboBK10()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, stDeptN As String
    
    'Modify By Sindy 2023/12/22
    If strSrvDate(1) >= 新部門啟用日 Then
      strQ = "Select a0922,st01,st02 From staff,acc090NEW Where st04='1' And st01>'63' and st01<'F' And st93=a0921(+) And substr(st01,4,1)<>'9' " & _
              "Order by st93,st01 asc"
    Else
    '2023/12/22 END
      strQ = "Select a0902,st01,st02 From staff,acc090 Where st04='1' And st01>'63' and st01<'F' And st03=a0901(+) And substr(st01,4,1)<>'9' " & _
              "Order by st03,st01 asc"
    End If
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    cboBK10.AddItem ""
    If RsQ.RecordCount > 0 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            cboBK10.AddItem "" & RsQ.Fields("st01") & " " & RsQ.Fields("st02")
            'Modify By Sindy 2023/12/22
            If strSrvDate(1) >= 新部門啟用日 Then
               stDeptN = "" & RsQ.Fields("a0922")
            Else
            '2023/12/22 END
               stDeptN = "" & RsQ.Fields("a0902")
            End If
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
    Dim txt, chk As CheckBox, i As Integer
    
    Select Case Lt
        'Cancel/OK
        Case 0
            For Each txt In TxtBK
                txt.Locked = True
            Next
            cboClass.Locked = True
            cboStatus.Locked = True
            TxtBK(1).Locked = False
            'Mark by Amy 2021/12/01 Form2.0 不支援Appearance屬性
'            TxtBK(1).Appearance = 1
'            TxtBK(1).BorderStyle = 1
        'Add/Upd
        Case 1
            For Each txt In TxtBK
               txt.Locked = False
            Next
            cboClass.Locked = False
            cboStatus.Locked = False
            TxtBK(1).Locked = True
            'Mark by Amy 2021/12/01 Form2.0 不支援Appearance屬性
'            TxtBK(1).Appearance = 0
'            TxtBK(1).BorderStyle = 0
        'Query
        Case 5
            For Each txt In TxtBK
                txt.Locked = True
                If txt.Index = 1 Then
                    txt.Locked = False
                    'Mark by Amy 2021/12/01 Form2.0 不支援Appearance屬性
'                    txt.Appearance = 1
'                    txt.BorderStyle = 1
                    txt.SetFocus
                ElseIf txt.Index = 2 Then
                    txt.Locked = False
                End If
            Next
            cboClass.Locked = True
            cboStatus.Locked = True
    End Select
End Sub

Private Sub RbEdit(intStatus As Integer)
    Dim strQ As String, strBK01 As String
    Dim SeqNo As String

    Select Case intStatus
        'Add
        Case 0
            ClearField
            ToolBarSet 0
            CmdPaper.Enabled = False
            EditMode = 0
            TxtBK(1).TabStop = True
            TxtBK(2).SetFocus
            TextInverse TxtBK(2)
            TxtBK(9) = strSrvDate(2) '預設上架日
            cboStatus = "借閱中"
            CmdPaper.Enabled = False
        'Update
        Case 1
            ToolBarSet 0
            EditMode = 1
            TxtBK(1).TabStop = True
            TxtBK(2).SetFocus
            TextInverse TxtBK(2)
       'Query
        Case 5
            ClearField
            ToolBarSet 0
            TxtLock 5
            EditMode = 5
   End Select
End Sub

Private Function ActionRecord(Optional ByVal bolUpdStatus As Boolean = False, Optional ByVal stLR01 As String = "") As Boolean
    Dim stExe1 As String, stExe2 As String, stExe3 As String
    Dim SeqNo As String
    
 On Error GoTo ErrHand
 
    ActionRecord = False
    Select Case EditMode
        'Add
        Case 0
            SeqNo = GetSerialNo_Lib(0)
            If SeqNo = MsgText(601) Then Exit Function
            If GetSerialNo_Lib(1) = MsgText(601) Then Exit Function
            stExe1 = GetSql(0, SeqNo)
            '寫一筆借閱記錄
            stExe2 = GetSql(2, SeqNo)
            m_CurrKEY = SeqNo
        'Upd
        Case 1
            stExe1 = GetSql(1, TxtBK(1))
            '修改狀態成 遺失或銷毀 則新增一筆借閱記錄
            If bolUpdStatus = True Then
                SeqNo = Format(Val(GetSerialNo_Lib(2)) + 1, "0000000")
                stExe2 = GetSql(IIf(cboStatus = "遺失", 3, 4), TxtBK(1), SeqNo)
            End If
            '修改保管人時,原保管人尚未歸還則需更新借閱記錄之保管人
            If stLR01 <> MsgText(601) Then
                stExe3 = "Update LoanRecord Set LR08='" & Left(Trim(cboBK10), 5) & "' Where  LR01='" & stLR01 & "' And LR03='" & TxtBK(1) & "' "
            End If
        'Del
        Case 2
            stExe1 = "Delete BooksData Where BK01='" & TxtBK(1) & "' "
            '刪除借閱記錄
            stExe2 = "Delete LoanRecord Where LR03='" & TxtBK(1) & "' "
            '刪除ImgByteFile (圖書封面)
            stExe3 = "Delete ImgByteFile Where IBF01='BOK' And IBF02='00" & TxtBK(1) & "' And IBF03='0' And IBF04='00' And IBF05='6' "
    End Select
    cnnConnection.BeginTrans
    If EditMode = 2 Then Pub_SeekTbLog stExe1
    cnnConnection.Execute stExe1
    If EditMode = 2 Then Pub_SeekTbLog stExe2
    If stExe2 <> MsgText(601) Then cnnConnection.Execute stExe2
    If stExe3 <> MsgText(601) Then cnnConnection.Execute stExe3
    cnnConnection.CommitTrans
    
    ActionRecord = True
    Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "新增作業失敗, 請洽電腦中心人員!!!"
End Function

'組Sql語法
'stEdit:0-新增 圖書基本資料/1-修改 圖書基本資料
'         2-新增 借閱記錄(借閱)/3-新增 借閱記錄(遺失)/4-新增 借閱記錄(銷毀)
Private Function GetSql(ByVal stEdit As Integer, ByVal stBK01 As String, Optional ByVal stLR01 As String = "") As String
    Dim bDifference As Boolean, bFirst As Boolean
    Dim NowTime As String
    Dim strSql As String, strTmp As String, strNowData As String
    
    bFirst = True: bDifference = False
    NowTime = ServerTime
    NowTime = IIf(Len(NowTime) = 6, Left(NowTime, 4), Left(NowTime, 3))
                 
    'Add
    If stEdit = 0 Then
        strSql = "Insert Into BooksData ("
        For i = 1 To BD_F
            strTmp = Empty: strNowData = Empty
            Select Case i
                Case 1
                    strNowData = stBK01
                Case 3
                    strNowData = GetBK03(True, cboClass)
                Case 9, 13
                    If TxtBK(i) <> MsgText(601) Then
                        strNowData = Val(TxtBK(i)) + 19110000
                    End If
                Case 10
                    strNowData = Left(Trim(cboBK10), 5)
                '目前未使用
                Case 12
                    'strNowData = cboStatus
                Case 15
                    strNowData = strUserNum
                Case 16
                    strNowData = strSrvDate(1)
                Case 17
                    strNowData = NowTime
                Case 18 To 20
                Case Else
                    strNowData = TxtBK(i)
            End Select
            If m_FieldList(i).fiOldData <> strNowData Then
                strTmp = m_FieldList(i).fiName
            End If
            If strTmp <> Empty Then
                bDifference = True
                If bFirst = True Then
                    strSql = strSql & strTmp
                    bFirst = False
                Else
                    strSql = strSql & "," & strTmp
               End If
            End If
        Next i
        strSql = strSql & ") "
        strSql = strSql & "Values ("
        
        bFirst = True
        For i = 1 To BD_F
            strTmp = Empty: strNowData = Empty
            Select Case i
                Case 1
                    strNowData = stBK01
                Case 3
                    strNowData = GetBK03(True, cboClass)
                Case 9, 13
                    If TxtBK(i) <> MsgText(601) Then
                        strNowData = Val(TxtBK(i)) + 19110000
                    End If
                Case 10
                    strNowData = Left(Trim(cboBK10), 5)
                '目前未使用
                Case 12
                    'strNowData = cboStatus
                Case 15
                    strNowData = strUserNum
                Case 16
                    strNowData = strSrvDate(1)
                Case 17
                    strNowData = NowTime
                Case 18 To 20
                Case Else
                    strNowData = TxtBK(i)
            End Select
            If m_FieldList(i).fiOldData <> strNowData Then
                If m_FieldList(i).fiType = 0 Then
                    strTmp = "'" & ChgSQL(strNowData) & "'"
                Else
                   strTmp = strNowData
                End If
            End If
            If strTmp <> Empty Then
                If bFirst = True Then
                    strSql = strSql & strTmp
                    bFirst = False
                Else
                    strSql = strSql & "," & strTmp
                End If
            End If
        Next i
        strSql = strSql & ")"
    'Update
    ElseIf stEdit = 1 Then
        strSql = "Update BooksData Set "
        For i = 2 To BD_F
            strTmp = Empty: strNowData = Empty
            Select Case i
                Case 3
                    strNowData = GetBK03(True, cboClass)
                Case 9, 13
                    If TxtBK(i) <> MsgText(601) Then
                        strNowData = Val(TxtBK(i)) + 19110000
                    End If
                Case 10
                    strNowData = Left(Trim(cboBK10), 5)
                '目前未使用
                Case 12
                    'strNowData = cboStatus
                Case 15 To 17
                Case 18
                    strNowData = strUserNum
                Case 19
                    strNowData = strSrvDate(1)
                Case 20
                    strNowData = NowTime
                Case Else
                    strNowData = TxtBK(i)
            End Select
            If i < 15 Or i > 17 Then
                If m_FieldList(i).fiOldData <> strNowData Then
                    If m_FieldList(i).fiType = 0 Then
                        If strNowData = Empty Then
                            strTmp = m_FieldList(i).fiName & " = NULL "
                        Else
                            strTmp = m_FieldList(i).fiName & " = '" & ChgSQL(strNowData) & "'"
                        End If
                    Else
                        If strNowData = Empty Then
                            strTmp = m_FieldList(i).fiName & " = NULL "
                        Else
                            strTmp = m_FieldList(i).fiName & " = " & strNowData
                        End If
                    End If
                End If
            End If
            If strTmp <> Empty Then
                bDifference = True
                If bFirst = True Then
                    strSql = strSql & strTmp
                    bFirst = False
                Else
                    strSql = strSql & "," & strTmp
                End If
            End If
        Next i
        strSql = strSql & " Where BK01 = '" & stBK01 & "' "
    '新增 借閱記錄
    Else
        strTmp = ""
        Select Case stEdit
            Case 2 '借閱
                strTmp = "1"
                If cboStatus = "遺失" Then
                    strTmp = "Y"
                ElseIf cboStatus = "銷毀" Then
                    strTmp = "Z"
                End If
                stLR01 = GetSerialNo_Lib(1)
            Case 3 '遺失
                strTmp = "Y"
            Case 4 '銷毀
                strTmp = "Z"
        End Select
        
        If strTmp = MsgText(601) Then
            strSql = ""
        ElseIf strTmp = "1" Then
            '應還日 預設系統後一個月
            strSql = "Insert Into LoanRecord (LR01,LR02,LR03,LR04,LR05,LR06,LR08,LR09,LR10) Values(" & _
                        CNULL(stLR01) & ",'" & strTmp & "'," & CNULL(stBK01) & "," & strSrvDate(1) & "," & DBDATE(DateAdd("m", 1, Format(strSrvDate(1), "####/##/##"))) & "," & _
                        strSrvDate(1) & "," & CNULL(Left(Trim(cboBK10), 5)) & "," & strSrvDate(1) & "," & ChgSQL(NowTime) & ")"
        Else
            '當狀態改為 遺失/銷毀 時同時新增一筆借閱記錄
            strSql = "Insert Into LoanRecord (LR01,LR02,LR03,LR04,LR06,LR08,LR09,LR10) Values(" & _
                        CNULL(stLR01) & ",'" & strTmp & "'," & CNULL(stBK01) & "," & strSrvDate(1) & "," & _
                        strSrvDate(1) & ",'" & strUserNum & "'," & strSrvDate(1) & "," & ChgSQL(NowTime) & ")"
        End If
    End If
    GetSql = strSql
End Function

Private Function TxtValidate() As Boolean
    Dim bolCancel As Boolean, strMsg As String
    
    TxtValidate = False
    
    If EditMode = 0 Or EditMode = 1 Then
        'Add by Amy 2021/12/01檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
        If PUB_ChkUniText(Me, True, True) = False Then
            Exit Function
        End If

        If cboClass = MsgText(601) Then MsgBox "書籍類別不可為空值", vbInformation: cboClass.SetFocus: Exit Function
        If TxtBK(4) = MsgText(601) And TxtBK(5) = MsgText(601) Then MsgBox "書名(中)/(外文)需擇一輸入", vbInformation: TxtBK(4).SetFocus: Exit Function
        If TxtBK(6) = MsgText(601) And TxtBK(7) = MsgText(601) Then MsgBox "作者(中)/(外文)需擇一輸入", vbInformation: TxtBK(6).SetFocus: Exit Function
        
        If TxtBK(9) = MsgText(601) Then
            MsgBox "上架日期不可為空值", vbInformation: TxtBK(9).SetFocus: Exit Function
        Else
            Call TxtBK_Validate(9, bolCancel)
            If bolCancel = True Then Exit Function
        End If
        
        If TxtBK(11) <> MsgText(601) And Trim(cboBK10) = MsgText(601) Then
            MsgBox "保管單位有輸,請輸入保管單位聯絡人", vbInformation: cboBK10.SetFocus: Exit Function
        ElseIf Trim(cboBK10) = MsgText(601) Then
            MsgBox "保管人不可為空", vbInformation: cboBK10.SetFocus: Exit Function
        Else
            strMsg = "Y": strExc(0) = ""
            strExc(0) = GetStaffName(Left(Trim(cboBK10), 5), strMsg)
            If strMsg <> MsgText(601) Then MsgBox strMsg, vbInformation: cboBK10.SetFocus: Exit Function
        End If
        
        Call TxtBK_Validate(13, bolCancel)
        If bolCancel = True Then Exit Function
        Call TxtBK_Validate(14, bolCancel)
        If bolCancel = True Then Exit Function
        If cboStatus = MsgText(601) Then MsgBox "狀態選取有誤,請確認！", vbInformation: cboStatus.SetFocus: Exit Function
    ElseIf EditMode = 5 Then
        If TxtBK(1) = MsgText(601) And TxtBK(2) = MsgText(601) Then
            MsgBox "請輸入查詢條件", vbInformation: TxtBK(1).SetFocus: Exit Function
        End If
    End If
    
    TxtValidate = True
End Function

Private Sub RefreshRange()
    Dim RsQ As New ADODB.Recordset
    Dim strSql As String

    strSql = "Select Nvl(MIN(BK01),0) as BK01 From BooksData "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        If IsNull(RsQ.Fields("BK01")) = False Then: m_FirstKEY = RsQ.Fields("BK01")
    End If
    RsQ.Close

    strSql = "Select Nvl(MAX(BK01),0) as BK01 From BooksData "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        If IsNull(RsQ.Fields("BK01")) = False Then: m_LastKEY = RsQ.Fields("BK01")
    End If
    RsQ.Close
    
    Set RsQ = Nothing
End Sub

Private Sub GetRecord(intChoose As Integer)
    Dim RsQ As New ADODB.Recordset
    Dim strSql As String
    
    Select Case intChoose
        'FirstRec
        Case 0
            m_CurrKEY = m_FirstKEY
        'PreRec
        Case 1
            strSql = "Select BK01  From BooksData Where BK01=(" & _
                         "Select MAX(BK01)  From BooksData Where BK01 < '" & m_CurrKEY & "') "
            RsQ.CursorLocation = adUseClient
            RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If RsQ.RecordCount > 0 Then
                If IsNull(RsQ.Fields("BK01")) = False Then: m_CurrKEY = RsQ.Fields("BK01")
                RsQ.Close
            End If
        'NextRec
        Case 2
            strSql = "Select BK01  From BooksData Where BK01=(" & _
                         "Select MIN(BK01) From BooksData Where BK01 > '" & m_CurrKEY & "') "
            RsQ.CursorLocation = adUseClient
            RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If RsQ.RecordCount > 0 Then
                If IsNull(RsQ.Fields("BK01")) = False Then: m_CurrKEY = RsQ.Fields("BK01")
                RsQ.Close
            End If
        'LastRec
        Case 3
            m_CurrKEY = m_LastKEY
    End Select
    QueryRecord m_CurrKEY
End Sub

Public Function QueryRecord(ByVal stBK01 As String, Optional ByVal stBK02 As String = "") As Boolean
    Dim oText, idx As Integer
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    QueryRecord = False
    
    If stBK01 <> MsgText(601) Then strQ = " And BK01='" & stBK01 & "' "
    If stBK02 <> MsgText(601) Then strQ = strQ & " And BK02='" & stBK02 & "' "
    
    strQ = "Select * From BooksData Where " & Mid(strQ, 6)
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ClearField
        With RsQ
            For Each oText In TxtBK
                idx = oText.Index
                m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
                If (idx = 9 Or idx = 13) And m_FieldList(idx).fiOldData <> MsgText(601) Then
                    oText.Text = Val(m_FieldList(idx).fiOldData) - 19110000
                Else
                    oText.Text = m_FieldList(idx).fiOldData
                End If
            Next
            '書籍類別
            m_FieldList(3).fiOldData = "" & .Fields(m_FieldList(3).fiName)
            cboClass = GetBK03(False, m_FieldList(3).fiOldData)
            '保管人/聯絡人
            m_FieldList(10).fiOldData = "" & .Fields(m_FieldList(10).fiName)
            cboBK10 = .Fields("BK10") & " " & StaffQuery(.Fields("BK10"))
            '狀態
'            m_FieldList(12).fiOldData = "" & .Fields(m_FieldList(12).fiName)-BK12不使用
'            cboStatus = "" & RsQ.Fields("LR02")
        End With
        UpdateCUID RsQ
        m_CurrKEY = stBK01
        If Not m_PrevForm Is Nothing Then EditMode = 5
        QueryRecord = True
        
        '抓取借閱記錄
        strQ = "Select ST02,LR02,LR04,LR06," & _
                "Decode(LR06,null,Decode(LR02, '1','借閱申請中', 'X','可借閱', 'Y','遺失', 'Z','銷毀', '延期申請中'), Decode(LR02, 'X','可借閱', 'Y','遺失', 'Z','銷毀', '借閱中')) as LR02N " & _
                "From LoanRecord a,Staff " & _
                "Where LR03='" & TxtBK(1) & "' And LR08=ST01(+) " & _
                "And LR01||LR02=(Select Max(LR01||LR02) as LR01 From LoanRecord Where a.LR03=LR03) "

        If RsQ.State = adStateOpen Then RsQ.Close
        RsQ.CursorLocation = adUseClient
        RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
        If RsQ.RecordCount > 0 Then
            '狀態
            cboStatus = "" & RsQ.Fields("LR02N")
            cboStatus.Tag = cboStatus
            If "" & RsQ.Fields("LR02") >= "1" And "" & RsQ.Fields("LR02") <= "W" And Not IsNull(RsQ.Fields("LR06")) Then
                LblLR(0) = "" & RsQ.Fields("ST02")
                LblLR(1) = Val("" & RsQ.Fields("LR04")) - 19110000
            End If
        End If
    End If
    If RsQ.State = adStateOpen Then RsQ.Close
    
    '判斷是否取圖書封面檔
    strQ = "Select  * From ImgByteFile Where IBF01='BOK' And IBF02='00" & TxtBK(1) & "' And IBF03='0' And IBF04='00' And IBF05='6' "
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
       CmdPaper.Enabled = True
    Else
       CmdPaper.Enabled = False
    End If
    If RsQ.State = adStateOpen Then RsQ.Close
End Function

Public Sub ToolBarSet(ByVal intStatus As Integer)
    Dim i As Integer
    Dim bolAccept As Boolean
         
    Select Case intStatus
        Case -1 'ReadyOnly
            TxtLock 0
            For i = 1 To 4
               TBar1.Buttons(i).Enabled = False
               TBar1.Buttons(i + 5).Enabled = False
            Next
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
        Case 0 'Add/Upd
            TxtLock 1
            For i = 1 To 4
               TBar1.Buttons(i).Enabled = False
               TBar1.Buttons(i + 5).Enabled = False
            Next
            TBar1.Buttons(11).Enabled = True
            TBar1.Buttons(12).Enabled = True
            TBar1.Buttons(14).Enabled = True
        Case 1
            TxtLock 0
            For i = 1 To 4
                bolAccept = False
                TBar1.Buttons(i).Enabled = False
                Select Case i
                    Case 1
                        If m_bInsert = True Then bolAccept = True
                    Case 2
                        If m_bUpdate = True Then bolAccept = True
                    Case 3
                        If m_bDelete = True Then bolAccept = True
                    Case 4 '查詢
                        bolAccept = True
                        TBar1.Buttons(i).Enabled = True
                End Select
                TBar1.Buttons(i).Enabled = bolAccept
                TBar1.Buttons(i + 5).Enabled = True
            Next
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
    End Select
End Sub

Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
    Dim strCName As String, strCDate As String, strCTime As String
    Dim strUName As String, strUDate As String, strUTime As String

    If IsNull(rsSrcTmp.Fields("BK15")) = False Then
       If IsEmptyText(rsSrcTmp.Fields("BK15")) = False Then
          strCName = StaffQuery(rsSrcTmp.Fields("BK15"))
          m_FieldList(15).fiOldData = "" & rsSrcTmp.Fields("BK15")
       End If
    End If
    If IsNull(rsSrcTmp.Fields("BK16")) = False Then
       If IsEmptyText(rsSrcTmp.Fields("BK16")) = False Then
          strCDate = Format(TAIWANDATE(rsSrcTmp.Fields("BK16")), "###/##/##")
          m_FieldList(16).fiOldData = "" & rsSrcTmp.Fields("BK16")
       End If
    End If
    If IsNull(rsSrcTmp.Fields("BK17")) = False Then
       If IsEmptyText(rsSrcTmp.Fields("BK17")) = False Then
          strCTime = Format(rsSrcTmp.Fields("BK17"), "##:##")
          m_FieldList(17).fiOldData = "" & rsSrcTmp.Fields("BK17")
       End If
    End If
    If IsNull(rsSrcTmp.Fields("BK18")) = False Then
        If IsEmptyText(rsSrcTmp.Fields("BK18")) = False Then
            strUName = StaffQuery(rsSrcTmp.Fields("BK18"))
            m_FieldList(18).fiOldData = "" & rsSrcTmp.Fields("BK18")
       End If
    End If
    If IsNull(rsSrcTmp.Fields("BK19")) = False Then
       If IsEmptyText(rsSrcTmp.Fields("BK19")) = False Then
          strUDate = Format(TAIWANDATE(rsSrcTmp.Fields("BK19")), "###/##/##")
          m_FieldList(19).fiOldData = "" & rsSrcTmp.Fields("BK19")
       End If
    End If
    If IsNull(rsSrcTmp.Fields("BK20")) = False Then
       If IsEmptyText(rsSrcTmp.Fields("BK20")) = False Then
          strUTime = Format(rsSrcTmp.Fields("BK20"), "##:##")
          m_FieldList(20).fiOldData = "" & rsSrcTmp.Fields("BK20")
       End If
    End If
    
    ' 設定CUID中的文字
    Label23.Caption = "CREATE : " & strCName & " " & _
               " " & strCDate & " " & _
               " " & strCTime & String(10, " ") & _
               "UPDATE : " & strUName & " " & _
               " " & strUDate & " " & _
               " " & strUTime
               
End Sub

Private Function GetBK03(ByVal bolTxt As Boolean, ByVal stVal As String) As String
    GetBK03 = ""
    If bolTxt = True Then
        GetBK03 = Mid(stVal, 1, InStr(stVal, ".") - 1)
    Else
        Select Case stVal
            Case 1
                GetBK03 = stVal & ".專利"
            Case 2
                GetBK03 = stVal & ".商標"
            Case 3
                GetBK03 = stVal & ".法律"
            Case 4
                GetBK03 = stVal & ".電腦"
            Case 5
                GetBK03 = stVal & ".其他"
        End Select
    End If
End Function

Private Sub ClearField()
    Dim oText, Lbl
    
    For Each oText In TxtBK
        oText.Text = Empty
    Next
    For Each Lbl In LblLR
       Lbl = ""
    Next
    cboBK10 = ""
    cboClass = ""
    cboStatus = "": cboStatus.Tag = ""
    Label23 = ""
    For i = 1 To BD_F
       m_FieldList(i).fiOldData = Empty
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm010034 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Modify by Amy 2021/12/01 原程式搬到Action,加記錄鍵盤傳入順序
    Dim KeyCode As Integer
    
    Select Case Button.Index
        Case 1 'Add
            KeyCode = vbKeyF2
        Case 2 'Update
            KeyCode = vbKeyF3
        Case 3 'Del
            KeyCode = vbKeyF5
        Case 4 'Query
            KeyCode = vbKeyF4
        Case 6 'MoveFirst
            KeyCode = vbKeyHome
        Case 7 'MovePrv
            KeyCode = vbKeyPageUp
        Case 8 'MoveNext
            KeyCode = vbKeyPageDown
        Case 9 'MoveLast
            KeyCode = vbKeyEnd
        Case 11 'OK
            KeyCode = vbKeyF9
        Case 12 'Cancel
            KeyCode = vbKeyF10
        Case 14 'Exit
            KeyCode = vbKeyEscape
    End Select
    
    Screen.MousePointer = vbHourglass
    Action Button.Index
    Screen.MousePointer = vbDefault
    'end 2021/12/01
End Sub

Private Sub TxtBK_GotFocus(Index As Integer)
   TextInverse TxtBK(Index)
    Select Case Index
        Case 1, 2, 5, 7, 9, 10, 13
            CloseIme
        Case 4, 6, 8, 11, 14
            OpenIme
    End Select
End Sub

'Modify by Amy 2021/12/01 原:Integer
Private Sub TxtBK_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
        Case 1, 2
            'Mark by Amy 2021/12/01取消以ENTER控制為換行的功能
            'If KeyAscii = 13 And EditMode = 5 Then Call Form_KeyDown(vbKeyF9, 0)
        Case 10
            KeyAscii = UpperCase(KeyAscii)
    End Select
End Sub

Private Sub TxtBK_Validate(Index As Integer, Cancel As Boolean)
    Dim stTmp As String 'Add by Amy 2018/01/31
    
    If EditMode > 1 Then Exit Sub
    If TxtBK(Index) = MsgText(601) Then Exit Sub
    
    Select Case Index
        '上架日/出刊日
        Case 9, 13
            'Modify by Amy 2018/01/31 出刊日期可能只有年月
            stTmp = TxtBK(Index)
            If Index = 13 And (Len(stTmp) = 7 Or Len(stTmp) = 6) And Right(stTmp, 2) = "00" Then
                stTmp = Left(stTmp, IIf(Len(stTmp) = 7, 5, 4)) & "01"
            End If
            If CheckIsTaiwanDate(stTmp) = False Then
                TxtBK(Index).SetFocus: Cancel = True: Exit Sub
            ElseIf Not ChkWorkDay(DBDATE(stTmp)) And Index = 9 Then
                MsgBox "上架日期必須是工作天 !", vbInformation: TxtBK(Index).SetFocus: Cancel = True: Exit Sub
            ElseIf Val(stTmp) > Val(strSrvDate(2)) And Index = 13 Then
                MsgBox "出刊日期不應該為未來日期!", vbInformation: TxtBK(Index).SetFocus: Cancel = True: Exit Sub
            End If
            'end 2018/01/31
        '備註
        Case 14
            If CheckLengthIsOK(TxtBK(Index), 500) = False Then TxtBK(Index).SetFocus: Cancel = True: Exit Sub
    End Select
End Sub

'檢查是否可以刪除-借閱記錄是否只有一筆(系統新增圖書館基本資料所寫的借閱記錄)
Private Function ChkNotDelete(ByVal strKEY01 As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
   
    ChkNotDelete = False
     strSql = "Select Sum(CRec) as Crec From(" & _
                    "Select Count(*) as CRec From LoanRecord a " & _
                    "Where LR03 = '" & strKEY01 & "' " & _
                    "And LR01<>(Select Min(LR01) as LR01 From LoanRecord Where a.LR03=LR03) " & _
                    "Union Select Count(*) as CRec From LoanRecord a " & _
                    "Where LR03 = '" & strKEY01 & "' " & _
                    "And LR01=(Select Min(LR01) as LR01 From LoanRecord Where a.LR03=LR03) And LR02<>'1' " & _
                  ")"

    '讀取資料庫
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    '檢查讀取的資料筆數
    If rsTmp.RecordCount > 0 Then
        If Val(rsTmp.Fields("CRec")) > 0 Then ChkNotDelete = True
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

'取得員工姓名(空白-顯示用/1-新增修改用/2-檢查用)
Public Function GetStaffName(ByVal stNo As String, Optional ByRef stShowMsg As String = "", Optional ByRef strDeptName) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String, strMsg As String
   
   GetStaffName = Empty

   strDeptName = Empty
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      'Modify by Amy 2025/01/10 bug-有+a0902但未串Acc090 會錯
      'strSql = "Select ST02,ST04,nvl(A0922,'(舊)'||A0902) A0902 From Staff,ACC090NEW Where ST01='" & stNo & "' And ST93=A0921(+) "
      strSql = "Select ST02,ST04,nvl(A0922,'(舊)'||A0902) A0902 From Staff,ACC090NEW,ACC090 " & _
                     "Where ST01='" & stNo & "' And ST93=A0921(+) And st03=a0901(+) "
   Else
      strSql = "Select ST02,ST04,A0902 From Staff,ACC090 Where ST01='" & stNo & "' And ST03=A0901(+) "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffName = rsTmp.Fields("ST02")
         If "" & rsTmp.Fields("A0902") <> MsgText(601) Then strDeptName = rsTmp.Fields("A0902")
      End If
      If stShowMsg <> MsgText(601) Then
          If rsTmp.Fields("ST04") = "2" Then
              stShowMsg = stNo & " [" & "" & rsTmp.Fields("ST02") & "] 已離職！"
              GetStaffName = Empty
              strDeptName = Empty
          Else
              stShowMsg = MsgText(601)
          End If
      End If
   Else
        stShowMsg = MsgText(601)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function GetAttachFile(ByVal m_AttachPath As String) As Boolean
    Dim lngSize As Long
    Dim iFileNo As Integer
    Dim bytes() As Byte
   
On Error GoTo ErrHnd

   GetAttachFile = False
    strExc(0) = "Select  * From ImgByteFile Where IBF01='BOK' And IBF02='00" & TxtBK(1) & "' And IBF03='0' And IBF04='00' And IBF05='6' "
    If RsTemp.State <> adStateClosed Then RsTemp.Close
    RsTemp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        If Dir(m_AttachPath) <> "" Then Kill m_AttachPath
        
         'Add By Sindy 2017/8/10
'         If "" & RsTemp.Fields("IBF15") <> "" Then
            GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("IBF15"), m_AttachPath & "\" & TxtBK(1) & ".pdf", UCase("ImgByteFile"))
'         Else
'         '2017/8/10 END
'            With RsTemp
'                lngSize = Val(.Fields("IBF13").Value)
'                ReDim bytes(lngSize)
'                If lngSize > 0 Then bytes() = .Fields("IBF14").GetChunk(lngSize)
'            End With
'            iFileNo = FreeFile
'            Open m_AttachPath & "\" & TxtBK(1) & ".pdf" For Binary Access Write As #iFileNo
'            If lngSize > 0 Then Put #iFileNo, , bytes()
'            Close #iFileNo
'            GetAttachFile = True
'         End If
    Else
        Close #iFileNo
    End If
   Exit Function
   
ErrHnd:
    MsgBox Err.Description, vbCritical
    If iFileNo > 0 Then Close #iFileNo
End Function

'確認保管人是否有歸回記錄
Private Function ChkBK10GiveBack(ByVal stBK10 As String, ByRef stLR01 As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    ChkBK10GiveBack = False: stLR01 = ""
    
    strQ = "Select LR01 From BooksData, LoanRecord a Where BK10='" & stBK10 & "' " & _
              "And LR01=(Select Max(LR01) as LR01 From LoanRecord Where a.LR03=LR03) " & _
              "And Not Exists(Select * From LoanRecord b Where a.LR03=b.LR03(+) " & _
              "And (LR02='X' or LR02='Y' or LR02='Z') And LR01=(Select Max(LR01) as LR01 From LoanRecord Where b.LR03=LR03)) " & _
              "And BK01=LR03(+) "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        stLR01 = "" & RsQ.Fields("LR01")
        ChkBK10GiveBack = True
    End If
    
    RsQ.Close
    Set RsQ = Nothing
End Function

'Add by Amy 2021/12/01 從Form_KeyDown搬過來改為Form_KeyUp
Private Sub Action(Index As Integer)
    If TBar1.Buttons(Index).Enabled = False Then Exit Sub
    
    Select Case Index
        Case 1 'Add
            Call Form_KeyUp(vbKeyF2, 0)
        Case 2 'Update
            Call Form_KeyUp(vbKeyF3, 0)
        Case 3 'Del
            Call Form_KeyUp(vbKeyF5, 0)
        Case 4 'Query
            Call Form_KeyUp(vbKeyF4, 0)
        Case 6 'MoveFirst
            Call Form_KeyUp(vbKeyHome, 0)
        Case 7 'MovePrv
            Call Form_KeyUp(vbKeyPageUp, 0)
        Case 8 'MoveNext
            Call Form_KeyUp(vbKeyPageDown, 0)
        Case 9 'MoveLast
            Call Form_KeyUp(vbKeyEnd, 0)
        Case 11 'OK
            Call Form_KeyUp(vbKeyF9, 0)
        Case 12 'Cancel
            Call Form_KeyUp(vbKeyF10, 0)
        Case 14 'Exit
            Call Form_KeyUp(vbKeyEscape, 0)
    End Select
End Sub

