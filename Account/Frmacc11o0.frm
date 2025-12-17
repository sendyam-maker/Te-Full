VERSION 5.00
Begin VB.Form Frmacc11o0 
   AutoRedraw      =   -1  'True
   Caption         =   "發票號碼維護"
   ClientHeight    =   5424
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5424
   ScaleWidth      =   5100
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   23
      Top             =   2010
      Width           =   1572
   End
   Begin VB.CommandButton CmdSearch 
      Height          =   300
      Left            =   3600
      Picture         =   "Frmacc11o0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   30
      Width           =   350
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   3240
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1590
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1590
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1230
      Width           =   500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3240
      MaxLength       =   8
      TabIndex        =   4
      Top             =   750
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   3
      Top             =   750
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   2
      Top             =   420
      Width           =   500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   30
      Width           =   800
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "Label9(5)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2500
      Index           =   5
      Left            =   180
      TabIndex        =   31
      Top             =   3960
      Width           =   4605
   End
   Begin VB.Label Lbl_Display1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl_Display1(5)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   30
      Top             =   3660
      Width           =   2100
   End
   Begin VB.Label Lbl_Display1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl_Display1(4)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2940
      TabIndex        =   29
      Top             =   3420
      Width           =   2100
   End
   Begin VB.Label Lbl_Display1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl_Display1(3)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2940
      TabIndex        =   28
      Top             =   3180
      Width           =   2100
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "最後修改發票日期欄人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   3660
      Width           =   2700
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "最後修改發票日期欄人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   26
      Top             =   3420
      Width           =   2700
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "最後修改發票日期欄人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   3180
      Width           =   2700
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "申    報    日    期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Lbl_Display1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl_Display1(2)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   21
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "目前使用發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2010
      Width           =   2100
   End
   Begin VB.Label Lbl_Display1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl_Display1(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   19
      Top             =   2640
      Width           =   2100
   End
   Begin VB.Label Labe13 
      BackStyle       =   0  '透明
      Caption         =   "目前使用最大號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   2100
   End
   Begin VB.Label Lbl_Display1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl_Display1(0)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   17
      Top             =   2400
      Width           =   2100
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "目前使用號碼範圍："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   2100
   End
   Begin VB.Label Labe8 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   1605
      Width           =   195
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "字軌２："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1230
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "號碼２："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   750
      Width           =   195
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "迄："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   30
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "號碼１："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   10
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "字軌１："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "年月起："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   30
      Width           =   1005
   End
End
Attribute VB_Name = "Frmacc11o0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已檢查 (無需修改的物件)
'Create By Amy 2013/11/29
Option Explicit

Dim RsAcc410 As New ADODB.Recordset
Dim intR As Integer
Dim OldA4112 As String '記錄發票日期
Dim ii As Integer

' 第一筆資料
Dim m_FirstKEY(2) As String
' 最後一筆資料
Dim m_LastKEY(2) As String
' 目前正在顯示
Dim m_CurrKEY(2) As String

Private Sub Form_Activate()
   Dim strKey() As String
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'Modify by Amy 2024/06/04 strItemNo為共用變數,避免有未清資料,無","而導致陣列索引錯誤當掉
   '一直測不出瑞婷說的錯,先加if 判斷觀察
   If InStr(strItemNo, ",") > 0 Then
      strKey = Split(strItemNo, ",")
      QueryRecord strKey(0), strKey(1)
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        If CmdSearch.Enabled = True Then
            cmdSearch_Click
        End If
    Else
        KeyEnter KeyCode
    End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5225, 6060 'Modify by Amy 2023/10/05 原5290
   FormClear
   RefreshRange
   GetFirstRecordVal     '設定第一筆key值
   GetNowRecordVal Left(strSrvDate(2), 5) 'Add by Amy 2014/03/13 '顯示當月那筆
   FormEnabled
   'Modify by Amy 2023/05/23 同frmacc1127
   Label9(5).Caption = "．收據抬頭為「可扣繳」資料則發票地址：" & vbCrLf & _
   "　先抓[客戶檔]中文地址(即營業登記地址)，" & vbCrLf & _
   "　無資料時再抓[收據抬頭檔]營業地址" & vbCrLf & _
                                    "．收據抬頭為「不可扣繳」資料則發票地址：" & vbCrLf & _
   "　先抓[客戶檔]聯絡地址，" & vbCrLf & _
   "　無資料時[客戶檔]中文地址(即營業登記地址)" & vbCrLf & _
   "　客戶檔[無]資料時再抓，" & vbCrLf & _
   "　[收據抬頭檔]郵寄地址，無資料時再抓營業地址" & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        Cancel = 1
        Exit Sub
    End If
    StatusClear
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Set Frmacc11o0 = Nothing
End Sub

Private Sub cmdSearch_Click()
    If Text1(0) = "" Then
        MsgBox "請輸入欲查詢起始年月！"
    ElseIf Text1(1) = "" Then
        MsgBox "請輸入欲查詢截止年月！"
    Else
        Screen.MousePointer = vbHourglass
        Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
        QueryRecord Text1(0), Text1(1)
    End If
    Screen.MousePointer = vbDefault
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Dim strTp As String
    
    If Trim(Text1(Index)) = MsgText(601) Then
         MsgBox "年月" & IIf(Index = 0, "起始", "截止") & "不可為空！"
         Cancel = True
    Else
         If Not IsNumeric(Text1(Index)) Then
            MsgBox "年月" & IIf(Index = 0, "起始", "截止") & "只能輸入數字！"
            Cancel = True
        Else
            If Index = 0 And Val(Right(Text1(Index), 2)) = 12 Then
                MsgBox "年月起始有誤，請確認！"
                Cancel = True
            ElseIf Len(Text1(Index)) < 4 Or Len(Text1(Index)) > 5 Or Val(Right(Text1(Index), 2)) > 12 Then
                MsgBox "年月" & IIf(Index = 0, "起始", "截止") & "格式錯誤！"
                Cancel = True
            ElseIf Index = 0 Then
                strTp = Right(Text1(Index), 2) + 1
                Text1(1) = Left(Text1(0), Len(Text1(0)) - 2) & "" & String(2 - Len(strTp), "0") & strTp
            End If
        End If
    End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    TextInverse Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If (Index = 1 Or Index = 2) And Val(Text2(Index)) > 0 Then
        Text2(Index) = "" & String(8 - Len(Text2(Index)), "0") & Text2(Index)
    End If
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
    If (Index = 1 Or Index = 2) And Trim(Text2(Index)) <> MsgText(601) Then
        If Not IsNumeric(Text2(Index)) Then
           MsgBox "字軌1" & IIf(Index = 1, "起始", "截止") & "號碼只能輸入數字！"
           Cancel = True
        Else
            If Val(Text2(1)) > 0 And Val(Text2(2)) > 0 Then
                If Val(Text2(1)) > Val(Text2(2)) Then
                    MsgBox "字軌1起始號碼不可大於截止號碼！"
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    TextInverse Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    If (Index = 1 Or Index = 2) And Val(Text3(Index)) > 0 Then
        Text3(Index) = "" & String(8 - Len(Text3(Index)), "0") & Text3(Index)
    End If
End Sub

Private Sub Text3_Validate(Index As Integer, Cancel As Boolean)
    If (Index = 1 Or Index = 2) And Trim(Text3(Index)) <> MsgText(601) Then
        If Not IsNumeric(Text3(Index)) Then
            MsgBox "字軌2" & IIf(Index = 1, "起始", "截止") & "號碼只能輸入數字！"
            Cancel = True
        Else
            If Val(Text3(1)) > 0 And Val(Text3(2)) > 0 Then
                If Val(Text3(1)) > Val(Text3(2)) Then
                    MsgBox "字軌2起始號碼不可大於截止號碼！"
                    Cancel = True
                End If
            End If
        End If
    End If

End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    If Trim(Text4) <> MsgText(601) Then
        If CheckIsTaiwanDate(Text4) = False Then
            Cancel = True
        Else
            If ChkWorkDay(Val(Text4) + 19110000) = True Then
                If Val(Text4) < Val(OldA4112) And OldA4112 <> MsgText(601) Then
                    MsgBox "目前使用發票日期不可小於原日期！"
                    Cancel = True
                End If
            Else
                MsgBox "目前使用發票日期需為工作日！"
                Cancel = True
            End If
        End If
    End If
End Sub

Public Function FormSave() As Boolean
   Dim strSave As String, strMsg As String
   Dim txt As TextBox
   Dim bCancel As Boolean, DMLlog As Boolean
   Dim CountTp As Integer
On Error GoTo Checking
   
   FormSave = False
   strSave = "": strMsg = ""
   
   For Each txt In Text1
        Text1_Validate txt.Index, bCancel
        If bCancel = True Then
            Text1(txt.Index).SetFocus
            Exit Function
        Else
            If strSaveConfirm = MsgText(3) Then '新增
                strSave = strSave & "," & CNULL(ChgSQL(txt.Text))
            End If
        End If
    Next
      
    For Each txt In Text2
        If txt = MsgText(601) Then
            If txt.Index <> 0 Then
                strMsg = "號碼" & IIf(txt.Index = 1, "起始", "截止")
            End If
            MsgBox "字軌1" & strMsg & "不可為空！"
            Text2(txt.Index).SetFocus
            Exit Function
        Else
            If txt.Index <> 0 Then
                Text2_Validate txt.Index, bCancel
                If bCancel = True Then
                    Text2(txt.Index).SetFocus
                    Exit Function
                End If
                If Len(txt) <> 8 Then
                    Text2(txt.Index) = "" & String(8 - Len(Text2(txt.Index)), "0") & Text2(txt.Index)
                End If
            End If
            
            If strSaveConfirm = MsgText(3) Then '新增
                strSave = strSave & "," & CNULL(ChgSQL(txt.Text))
            ElseIf strSaveConfirm = MsgText(4) Then
                strSave = strSave & ",A410" & txt.Index + 3 & "=" & CNULL(ChgSQL(IIf(txt.Index = 0, txt.Text, Val(txt.Text))))
            End If
        End If
    Next
   
   CountTp = 0
    For Each txt In Text3
        If txt <> MsgText(601) Then
            CountTp = CountTp + 1
            If txt.Index <> 0 Then
                Text3_Validate txt.Index, bCancel
                If bCancel = True Then
                    Text3(txt.Index).SetFocus
                    Exit Function
                End If
                If Len(txt) <> 8 And Trim(txt) <> MsgText(601) Then
                    Text3(txt.Index) = "" & String(8 - Len(Text3(txt.Index)), "0") & Text3(txt.Index)
                End If
            End If
        End If
    Next
    If CountTp > 0 And CountTp <> 3 Then
        MsgBox "字軌2輸入不完整，請確認！"
        Exit Function
    End If
    '判斷字軌範圍不可重疊
    If Trim(Text2(0)) <> MsgText(601) And Trim(Text3(0)) <> MsgText(601) And Text2(0) = Text3(0) Then
        If (Text2(1) <= Text3(1) And Text2(2) >= Text3(1)) Or (Text2(1) <= Text3(2) And Text2(2) >= Text3(2)) Then
            MsgBox "字軌範圍重疊，請確認！"
            Exit Function
        End If
    End If
    If strSaveConfirm = MsgText(3) Then '新增
        strSave = strSave & "," & CNULL(ChgSQL(Text3(0))) & "," & CNULL(ChgSQL(Text3(1))) & "," & CNULL(ChgSQL(Text3(2)))
    ElseIf strSaveConfirm = MsgText(4) Then
        strSave = strSave & ",A4106= " & CNULL(ChgSQL(Text3(0))) & ",A4107= " & CNULL(ChgSQL(IIf(Val(Text3(1)) > 0, Val(Text3(1)), Text3(1)))) & ",A4108= " & CNULL(ChgSQL(IIf(Val(Text3(2)) > 0, Val(Text3(2)), Text3(2))))
    End If
        
    Text4_Validate bCancel
    If bCancel = True Then
        Text4.SetFocus
        Exit Function
    Else
        If strSaveConfirm = MsgText(3) Then '新增
            strSave = strSave & "," & CNULL(ChgSQL(Text4))
        ElseIf strSaveConfirm = MsgText(4) And OldA4112 <> Text4 Then
            strSave = strSave & ",A4112= " & CNULL(ChgSQL(Text4)) & ",A4113='" & strUserNum & "',A4114=" & Val(strSrvDate(2)) & ",A4115=" & ServerTime
            DMLlog = True
        End If
    End If
    
   '判斷資料是否重疊
   If strSaveConfirm = MsgText(3) Then '新增
        If Len(Text1(0)) = 5 Then
            CountTp = 3
        Else
            CountTp = 2
        End If
        strExc(0) = "Select a4101,a4102 From acc410 Where substr(A4101,1," & CountTp & ")=" & Left(Text1(0), CountTp) & " Order by A4101"
        intI = 1: strExc(1) = ""
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            Do While Not RsTemp.EOF
                strExc(1) = strExc(1) & RsTemp.Fields("A4101") & "," & RsTemp.Fields("A4102") & ","
                RsTemp.MoveNext
            Loop
            If InStr(1, strExc(1), Text1(0)) > 0 Then
                MsgBox "年月資料重覆！"
                Text1(0).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Len(strSave) > 0 Then
        If strSaveConfirm = MsgText(3) Then
            '新增時 a4109 目前使用號碼範圍 預設為1
            strSave = "Insert Into acc410 (a4101,a4102,a4103,a4104,a4105,a4106,a4107,a4108,a4112,a4109) Values(" & Mid(strSave, 2) & ",1)"
        ElseIf strSaveConfirm = MsgText(4) Then
            strSave = "Update acc410 Set " & Mid(strSave, 2) & " Where A4101='" & Val(Text1(0)) & "' And A4102='" & Val(Text1(1)) & "' "
        End If
        adoTaie.BeginTrans
        If DMLlog = True Then
            Pub_SeekTbLog strSave
        End If
        adoTaie.Execute strSave
        adoTaie.CommitTrans
        FormSave = True
        If strSaveConfirm = MsgText(3) Then
            GetCurrRecordVal Text1(0), Text1(1)
        Else
            RefreshRange
        End If
    End If
    FormShow
    
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
    adoTaie.RollbackTrans
    MsgBox "更新失敗！"
    MsgBox Err.Description, , MsgText(5)
End Function

Public Function FormDel() As Boolean
    If DelRecord = True Then
        RefreshRange
    Else
        Exit Function
    End If
End Function

Private Function DelRecord() As Boolean
    Dim strDel As String, strNo1 As String, strNo2 As String
    DelRecord = False
On Error GoTo ErrHand
   
    strNo1 = Val(Text1(0))
    strNo2 = Val(Text1(1))
    strDel = "Delete Acc410 Where A4101=" & strNo1 & " And A4102=" & strNo2 & " "
    cnnConnection.Execute strDel
    
         
    ' 只有刪除的是最後一筆或第一筆須重新取的第一筆及最後一筆
   If (strNo1 = m_LastKEY(0) And strNo2 = m_LastKEY(1)) Or (strNo1 = m_FirstKEY(0) And strNo2 = m_FirstKEY(1)) Then
      RefreshRange
   End If
    GetCurrRecordVal strNo1, strNo2
    DelRecord = True
   
ErrHand:
   If Err.Number = 0 Then
      Exit Function
   End If
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

Public Sub FormShow()
  If m_CurrKEY(0) <> MsgText(601) And m_CurrKEY(1) <> MsgText(601) Then
    strSql = "Select * From Acc410 " & _
                "Where A4101=" & m_CurrKEY(0) & " And A4102=" & m_CurrKEY(1) & " ORDER BY A4101,A4102"
                
    intR = 1
    Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
    If intR = 1 Then
         FormClear
         With RsAcc410
             OldA4112 = "" & .Fields("A4112") '記錄目前使用發票日期
             Text1(0) = "" & .Fields("A4101")
             Text1(1) = "" & .Fields("A4102")
      
             '字軌1
             Text2(0) = "" & .Fields("A4103")
             Text2(1) = String(8 - Len("" & .Fields("A4104")), "0") & "" & .Fields("A4104")
             Text2(2) = String(8 - Len("" & .Fields("A4105")), "0") & "" & .Fields("A4105")
             '字軌2
             Text3(0) = "" & .Fields("A4106")
             If Not IsNull(.Fields("A4107")) Then
                Text3(1) = String(8 - Len("" & .Fields("A4107")), "0") & "" & .Fields("A4107")
             End If
             If Not IsNull(.Fields("A4108")) Then
                Text3(2) = String(8 - Len("" & .Fields("A4108")), "0") & "" & .Fields("A4108")
             End If
             '目前使用發票日期
             Text4 = "" & .Fields("A4112")
             '目前使用號碼範圍
             Lbl_Display1(0).Caption = "" & .Fields("A4109")
             '目前使用最大號碼
             If Not IsNull(.Fields("A4110")) Then
                Lbl_Display1(1).Caption = IIf(Val(.Fields("A4109")) = 1, .Fields("A4103"), .Fields("A4106")) & String(8 - Len("" & .Fields("A4110")), "0") & "" & .Fields("A4110")
             End If
             '申報日期
             If Not IsNull(.Fields("A4111")) Then
                Lbl_Display1(2).Caption = Format(.Fields("A4111"), "###/##/##")
             End If
             '最後修改發票日期人員
             If Not IsNull(.Fields("A4113")) Then
                Lbl_Display1(3).Caption = GetStaffName(.Fields("A4113"), True)
             End If
             '最後修改發票日期日期
             If Not IsNull(.Fields("A4114")) Then
                Lbl_Display1(4).Caption = Format(.Fields("A4114"), "###/##/##")
             End If
             '最後修改發票日期時間
             If Not IsNull(.Fields("A4115")) Then
                Lbl_Display1(5).Caption = Format(IIf(Len(.Fields("A4115")) = 6, Left(.Fields("A4115"), 4), Left(.Fields("A4115"), 3)), "##:##")
             End If
            
        End With
   End If
   End If
   Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
End Sub

Public Sub FormEnabled()
   If strSaveConfirm = MsgText(4) Then '修改狀態時
       Text1(0).Enabled = False
       Text1(1).Enabled = False
       CmdSearch.Enabled = False
       TextLocked (False)
   Else
        Text1(0).Enabled = True
        Text1(1).Enabled = True
        If strSaveConfirm = MsgText(3) Then '新增狀態時
            CmdSearch.Enabled = False
            TextLocked (False)
        Else
            CmdSearch.Enabled = True
            TextLocked (True)
        End If
   End If
End Sub

Private Sub TextLocked(ByVal IsLock As Boolean)
    Dim oText As TextBox
    For Each oText In Text2
      oText.Locked = IsLock
    Next
    For Each oText In Text3
      oText.Locked = IsLock
    Next
    Text4.Locked = IsLock
End Sub

Public Sub FormClear()
   Dim oText As TextBox
   Dim oLbl As LABEL
   For Each oText In Text1
      oText = ""
   Next
   For Each oText In Text2
      oText = ""
   Next
   For Each oText In Text3
      oText = ""
   Next
   Text4 = ""
   For Each oLbl In Lbl_Display1
       oLbl.Caption = ""
   Next
   Frmacc0000.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub RefreshRange()

Dim strSql As String

   strSql = "Select A4101,A4102 From Acc410 " & _
               "Where A4101 = (Select MIN(A4101) From Acc410) AND " & _
                          "A4102 = (Select MIN(A4102) From Acc410 " & _
                                         "Where A4101 = (Select MIN(A4101) FROM Acc410)) "
   
    intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      If IsNull(RsAcc410.Fields("A4101")) = False Then: m_FirstKEY(0) = RsAcc410.Fields("A4101")
      If IsNull(RsAcc410.Fields("A4102")) = False Then: m_FirstKEY(1) = RsAcc410.Fields("A4102")
   End If
   RsAcc410.Close

    strSql = "Select A4101,A4102 From Acc410 " & _
               "Where A4101 = (Select MAX(A4101) From Acc410) AND " & _
                          "A4102 = (Select MAX(A4102) From Acc410 " & _
                                         "Where A4101 = (Select MAX(A4101) FROM Acc410)) "
 
    intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      If IsNull(RsAcc410.Fields("A4101")) = False Then: m_LastKEY(0) = RsAcc410.Fields("A4101")
      If IsNull(RsAcc410.Fields("A4102")) = False Then: m_LastKEY(1) = RsAcc410.Fields("A4102")
   End If
   RsAcc410.Close
   
   Set RsAcc410 = Nothing
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim strSql As String
   
   IsRecordExist = False
   strSql = "Select * From Acc410 " & _
                "Where A4101 = '" & strKEY01 & "' AND " & _
                          "A4102 = '" & strKEY02 & "' "
                  
   ' 讀取資料庫
    intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   RsAcc410.Close
   Set RsAcc410 = Nothing
End Function
' 顯示資料
Private Sub GetCurrRecordVal(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String

   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "Select A4101,A4102 From Acc410 " & _
                  "Where A4101 = '" & m_CurrKEY(0) & "' AND " & _
                             "A4102 = (Select MIN(A4102) From Acc410 " & _
                                          "Where A4101 = '" & m_CurrKEY(0) & "' AND " & _
                                                    "A4102 > '" & m_CurrKEY(1) & "' )"
      intR = 1
      Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
      If intR = 1 Then
        If IsNull(RsAcc410.Fields("A4101")) = False Then: m_CurrKEY(0) = RsAcc410.Fields("A4101")
        If IsNull(RsAcc410.Fields("A4102")) = False Then: m_CurrKEY(1) = RsAcc410.Fields("A4102")
         RsAcc410.Close
         RefreshRange
         FormShow
         GoTo EXITSUB
      End If
      RsAcc410.Close
      
      strSql = "Select A4101,A4102 From Acc410 " & _
                  "Where A4101 = (Select MIN(A4101) From Acc410 " & _
                                         "Where A4101 > '" & m_CurrKEY(0) & "') AND " & _
                                                    "A4102 = (Select MIN(A4102) From Acc410 " & _
                                                                 "Where A4101 = (Select MIN(A4101) From Acc410 " & _
                                                                                        "Where A4101 > '" & m_CurrKEY(0) & "')) "
       intR = 1
       Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
       If intR = 1 Then
     
         If IsNull(RsAcc410.Fields("A4101")) = False Then: m_CurrKEY(0) = RsAcc410.Fields("A4101")
         If IsNull(RsAcc410.Fields("A4102")) = False Then: m_CurrKEY(1) = RsAcc410.Fields("A4102")
       Else
         GetLastRecordVal
         GoTo EXITSUB
       End If
       RsAcc410.Close
   End If
   RefreshRange
   FormShow
EXITSUB:
End Sub
' 第一筆資料
Public Sub GetFirstRecordVal()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   FormShow
End Sub
'上一筆資料
Public Sub GetPreRecordVal()
    Dim strSql As String
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "Select A4101,A4102 From Acc410 " & _
               "Where A4101 = '" & m_CurrKEY(0) & "' And " & _
                         "A4102 = (Select MAX(A4102) From Acc410 " & _
                                        "Where A4101 = '" & m_CurrKEY(0) & "' AND " & _
                                                   "A4102 < '" & m_CurrKEY(1) & "' )"
    intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      If IsNull(RsAcc410.Fields("A4101")) = False Then: m_CurrKEY(0) = RsAcc410.Fields("A4101")
      If IsNull(RsAcc410.Fields("A4102")) = False Then: m_CurrKEY(1) = RsAcc410.Fields("A4102")
      RsAcc410.Close
      FormShow
      GoTo EXITSUB
   End If
   RsAcc410.Close
   
   strSql = "Select A4101,A4102 From Acc410 " & _
               "Where A4101 = (Select MAX(A4101) From Acc410 " & _
                                       "Where A4101 < '" & m_CurrKEY(0) & "') AND " & _
                                       "A4102 = (Select MAX(A4102) From Acc410 " & _
                                       "Where A4101 = (Select MAX(A4101) From Acc410 " & _
                                       "Where A4101 < '" & m_CurrKEY(0) & "')) "

   intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      If IsNull(RsAcc410.Fields("A4101")) = False Then: m_CurrKEY(0) = RsAcc410.Fields("A4101")
      If IsNull(RsAcc410.Fields("A4102")) = False Then: m_CurrKEY(1) = RsAcc410.Fields("A4102")
   End If
   RsAcc410.Close
   FormShow
   
EXITSUB:
   Set RsAcc410 = Nothing
End Sub
'下一筆資料
Public Sub GetNextRecordVal()
    Dim strSql As String
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "Select A4101,A4102 From Acc410 " & _
               "Where A4101 = '" & m_CurrKEY(0) & "' AND " & _
                         "A4102 = (Select MIN(A4102) From Acc410 " & _
                                       "Where A4101 = '" & m_CurrKEY(0) & "' AND " & _
                                                  "A4102 > '" & m_CurrKEY(1) & "' )"
    intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      If IsNull(RsAcc410.Fields("A4101")) = False Then: m_CurrKEY(0) = RsAcc410.Fields("A4101")
      If IsNull(RsAcc410.Fields("A4102")) = False Then: m_CurrKEY(1) = RsAcc410.Fields("A4102")
      RsAcc410.Close
      FormShow
      GoTo EXITSUB
   End If
   RsAcc410.Close
   
   strSql = "Select A4101,A4102 From Acc410 " & _
            "Where A4101 = (Select MIN(A4101) From Acc410 " & _
                                    "Where A4101 > '" & m_CurrKEY(0) & "') AND " & _
                                              "A4102 = (Select MIN(A4102) FROM customer " & _
                                                           "Where A4101 = (Select MIN(A4101) From Acc410 " & _
                                                                                     "Where A4101 > '" & m_CurrKEY(0) & "')) "

    intR = 1
   Set RsAcc410 = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
      If IsNull(RsAcc410.Fields("A4101")) = False Then: m_CurrKEY(0) = RsAcc410.Fields("A4101")
      If IsNull(RsAcc410.Fields("A4102")) = False Then: m_CurrKEY(1) = RsAcc410.Fields("A4102")
   End If
   RsAcc410.Close
   FormShow
   
EXITSUB:
   Set RsAcc410 = Nothing
End Sub
' 最後一筆資料
Public Sub GetLastRecordVal()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   FormShow
End Sub

' 查詢記錄
Private Function QueryRecord(ByVal strA4101 As String, ByVal strA4102 As String) As Boolean

   QueryRecord = False
   strA4101 = Val(strA4101)
   strA4102 = Val(strA4102)
  
   If IsRecordExist(strA4101, strA4102) = True Then
      m_CurrKEY(0) = strA4101
      m_CurrKEY(1) = strA4102
      QueryRecord = True
   Else
      MsgBox ("無此資料")
   End If
   FormShow
 
End Function

Private Sub CheckOC()
    If RsAcc410.State = adStateOpen Then
        RsAcc410.Close
    End If
End Sub

Private Sub CheckOCrs()
    If RsAcc410.State = adStateOpen Then
        RsAcc410.Close
    End If
End Sub

'Add by Amy 2014/03/13 抓取當月那筆
Private Sub GetNowRecordVal(ByVal strSrvYM As String)
    Dim strSql As String 'Add by Amy 2019/12/19 共用變數可能造成錯
    
    strSql = "Select * From acc410 Where A4101<='" & strSrvYM & "' and  A4102>='" & strSrvYM & "' "
    intR = 1
    Set RsTemp = ClsLawReadRstMsg(intR, strSql)
    If intR = 1 Then
        m_CurrKEY(0) = RsTemp.Fields("A4101")
        m_CurrKEY(1) = RsTemp.Fields("A4102")
         FormShow
    Else
        m_CurrKEY(0) = m_LastKEY(0)
        m_CurrKEY(1) = m_LastKEY(1)
        FormShow
    End If
End Sub
'end 2014/03/03/13
