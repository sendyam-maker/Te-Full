VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140112_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "預約內容"
   ClientHeight    =   5700
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6924
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6924
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   510
      Width           =   750
   End
   Begin VB.ComboBox cboTime 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2475
      TabIndex        =   2
      Text            =   "cboTime"
      Top             =   1260
      Width           =   1185
   End
   Begin VB.ComboBox cboTime 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1125
      TabIndex        =   1
      Text            =   "cboTime"
      Top             =   1260
      Width           =   1185
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Top             =   1650
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1785
      Left            =   180
      TabIndex        =   13
      Top             =   3750
      Width           =   6540
      Begin VB.CheckBox Check2 
         Caption         =   "每日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   32
         Top             =   270
         Width           =   870
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "要預約日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3510
         TabIndex        =   21
         Top             =   600
         Width           =   1410
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "全選"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3510
         TabIndex        =   9
         Top             =   1260
         Width           =   1410
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "不預約日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3510
         TabIndex        =   8
         Top             =   930
         Width           =   1410
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "所有日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3510
         TabIndex        =   7
         Top             =   270
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "每週"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   870
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1056
         ItemData        =   "frm140112_1.frx":0000
         Left            =   4950
         List            =   "frm140112_1.frx":0002
         Style           =   1  '項目包含核取方塊
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   300
         Left            =   1125
         TabIndex        =   24
         Top             =   570
         Width           =   1125
         _ExtentX        =   1990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(  請輸入預約日期起算 3 個月以內的日期 )"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   180
         TabIndex        =   20
         Top             =   930
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "結束日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   630
         Width           =   900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5355
      TabIndex        =   12
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4050
      TabIndex        =   11
      Top             =   90
      Width           =   1230
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1125
      TabIndex        =   23
      Top             =   900
      Width           =   1125
      _ExtentX        =   1990
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CheckBox Check3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   3000
      Value           =   1  '核取
      Width           =   330
   End
   Begin MSForms.ComboBox cboRoom 
      Height          =   345
      Left            =   1125
      TabIndex        =   0
      Top             =   510
      Width           =   3210
      VariousPropertyBits=   545343515
      DisplayStyle    =   7
      Size            =   "6773;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "教育訓練編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4560
      TabIndex        =   30
      Top             =   570
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(主題：請務必填寫)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1125
      TabIndex        =   29
      Top             =   2040
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "預約"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   28
      Top             =   570
      Width           =   450
   End
   Begin MSForms.Label lblOldUser 
      Height          =   924
      Left            =   4860
      TabIndex        =   27
      Top             =   1296
      Width           =   1752
      ForeColor       =   255
      Size            =   "3090;1630"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   276
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTimes 
      AutoSize        =   -1  'True
      Caption         =   "第1次"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3915
      TabIndex        =   26
      Top             =   1650
      Width           =   750
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "EMail 通知( 於預約日期凌晨寄發 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   540
      TabIndex        =   25
      Top             =   3000
      Width           =   3345
   End
   Begin MSForms.Label lblCreateData 
      Height          =   300
      Left            =   180
      TabIndex        =   22
      Top             =   3336
      Width           =   6492
      VariousPropertyBits=   27
      Size            =   "11451;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   216
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "預約日期                     ( 請輸入 4 個月以內的日期 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   19
      Top             =   945
      Width           =   4905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "時間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   18
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "預約人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   17
      Top             =   1710
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   2250
      X2              =   2475
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "使用單位"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   16
      Top             =   2040
      Width           =   900
   End
   Begin MSForms.Label lblUserName 
      Height          =   264
      Left            =   2472
      TabIndex        =   15
      Top             =   1716
      Width           =   1308
      VariousPropertyBits=   27
      Size            =   "2307;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   216
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtContent 
      Height          =   645
      Left            =   180
      TabIndex        =   4
      Top             =   2280
      Width           =   6450
      VariousPropertyBits=   -1467987941
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "11377;1138"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm140112_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/4/22 改成Form2.0 (txtContent)
'Memo by Morgan 2024/4/21 改成Form2.0 (lblOldUser,lblUserName,lblCreateData)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Morgan 2011/5/16
Option Explicit

Public m_State As String 'A=新增,E=修改,D=刪除,S=檢視
Public m_Users As String 'Added by Morgan 2019/7/12

Private Type DateList
   Date As String
   Selected As Boolean
End Type

Dim arrDateList() As DateList '週期性預約日期清單
Dim bolRefresh As Boolean '是否重新抓日期資料
'Add by Amy 2019/01/24
Public bolIs140113 As Boolean '由教育訓練進入
Public strRoomNo As String, strRR20 As String '會議室編號/教育訓練編號
Dim oRR(16) As String '由FormSave搬過來
Public cboRoomItemData As String  'Added by Morgan 2021/4/22

Private Sub cboRoom_Change()
   SetlblEmail
End Sub

Private Sub SetlblEmail()
   'Add by Amy 2019/01/24 +if 因教育訓練進入cboRoom.ItemData(cboRoom.ListIndex) = 9會 Error
   If bolIs140113 = True Or strRR20 <> MsgText(601) Then
      lblEmail = "EMail 通知( 於預約日期前一日凌晨寄發 )"
   ElseIf PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex) = 9 Then
      lblEmail = "EMail 通知( 於預約日期前一日凌晨寄發 )"
   Else
      lblEmail = "EMail 通知( 於預約日期凌晨寄發 )"
   End If
End Sub

Private Sub cboTime_Change(Index As Integer)
   'Add by Amy 2019/11/12
   If bolIs140113 = True Or strRR20 <> MsgText(601) Then Exit Sub
   SetlblOldUser
End Sub

Private Sub cboTime_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      MaskEdBox2.Enabled = True
      cmdFunc(0).Enabled = True
      cmdFunc(1).Enabled = True
      cmdFunc(2).Enabled = True
      cmdFunc(3).Enabled = True
      Check3.Value = 0
      Check2.Value = 0 'Added by Morgan 2019/7/3
   Else
      MaskEdBox2.Mask = ""
      MaskEdBox2.Text = ""
      MaskEdBox2.Mask = DFormat
      MaskEdBox2.Enabled = False
      cmdFunc(0).Enabled = False
      cmdFunc(1).Enabled = False
      cmdFunc(2).Enabled = False
      cmdFunc(3).Enabled = False
      Check3.Value = 1
      ResetList
   End If
End Sub

'Added by Morgan 2019/7/3
Private Sub Check2_Click()
   If Check2.Value = 1 Then
      MaskEdBox2.Enabled = True
      cmdFunc(0).Enabled = True
      cmdFunc(1).Enabled = True
      cmdFunc(2).Enabled = True
      cmdFunc(3).Enabled = True
      Check3.Value = 0
      Check1.Value = 0
   Else
      MaskEdBox2.Mask = ""
      MaskEdBox2.Text = ""
      MaskEdBox2.Mask = DFormat
      MaskEdBox2.Enabled = False
      cmdFunc(0).Enabled = False
      cmdFunc(1).Enabled = False
      cmdFunc(2).Enabled = False
      cmdFunc(3).Enabled = False
      Check3.Value = 1
      ResetList
   End If
End Sub

Private Sub ResetList()
   List1.Clear
   Erase arrDateList
   bolRefresh = True
End Sub

Public Sub ReadDetail(pDate As String)
   Dim rr(3) As String, jj As Integer, Index As Integer
   'Modified by Morgan 2019/7/3 +Check2
   If Check1.Value = 1 Or Check2.Value = 1 Then
      rr(1) = PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex)
      rr(2) = DBDATE(MaskEdBox1)
      rr(3) = Replace(cboTime(0), ":", "")
      strExc(0) = "select * from RoomResDetail where rd01=" & rr(1) & " and rd02=" & rr(2) & " and rd03=" & rr(3)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ResetList
         jj = 0
         List1.Visible = False
         With RsTemp
         Do While Not .EOF
            ReDim Preserve arrDateList(jj)
            arrDateList(jj).Date = Format(TransDate(.Fields("RD04"), 1), "###/##/##")
            List1.AddItem arrDateList(jj).Date, jj
            If IsNull(.Fields("RD05")) Then
               arrDateList(jj).Selected = True
               List1.Selected(jj) = True
            End If
            If .Fields("RD04") = pDate Then Index = jj
            jj = jj + 1
            .MoveNext
         Loop
         End With
         List1.ListIndex = Index
         List1.Visible = True
         bolRefresh = False
      End If
   End If
End Sub

Private Function RefreshList() As Boolean
   Dim ii As Single, jj As Single
   Dim strStartDate As String, strEndDate As String
   
   strStartDate = Replace(MaskEdBox1, "/", "")
   strEndDate = Replace(MaskEdBox2, "/", "")
   
   If MaskEdBox2 = "___/__/__" Then
      MsgBox "請輸入結束日期！"
      MaskEdBox2.SetFocus
      Exit Function
   ElseIf ChkDate(strEndDate) = False Then
      MaskEdBox2.SetFocus
      Exit Function
   ElseIf Val(strEndDate) < Val(strStartDate) Then
      MsgBox "結束日期不可早於預約日期！"
      MaskEdBox2.SetFocus
      Exit Function
      
   ElseIf Pub_StrUserSt03 <> "M51" Then
      strExc(1) = CompDate(1, 2, strStartDate)
      If DBDATE(strEndDate) > strExc(1) Then
         MsgBox "結束日期請輸入預約日期起算 3 個月以內的日期！"
         MaskEdBox2.SetFocus
         Exit Function
      End If
   End If
   
   ResetList
   jj = 0
   ii = Val(strStartDate)
   List1.Visible = False
   Do While ii <= Val(strEndDate)
      ReDim Preserve arrDateList(jj)
      arrDateList(jj).Date = Format(ii, "###/##/##")
      arrDateList(jj).Selected = True
      List1.AddItem arrDateList(jj).Date, jj
      List1.Selected(jj) = True
      
      'Modified by Morgan 2019/7/3 +Check2
      If Check2.Value = vbChecked Then
         ii = TransDate(CompDate(2, 1, ii), 1)
      Else
         ii = TransDate(CompDate(2, 7, ii), 1)
      End If
      jj = jj + 1
   Loop
   List1.ListIndex = 0
   List1.Visible = True
  
   bolRefresh = False
   RefreshList = True
End Function
Private Sub ListAllDate()
   Dim ii As Single, jj As Single
   If bolRefresh = True Then
      RefreshList
   Else
      List1.Visible = False
      List1.Clear
      jj = 0
      For ii = LBound(arrDateList) To UBound(arrDateList)
         List1.AddItem arrDateList(ii).Date, jj
         List1.Selected(jj) = arrDateList(ii).Selected
         jj = jj + 1
      Next
      List1.ListIndex = 0
      List1.Visible = True
   End If
End Sub

Private Sub SelectAll()
   Dim ii As Single, Index As Integer
   
   If List1.ListCount > 0 Then
      Index = List1.ListIndex
      For ii = 0 To List1.ListCount - 1
         If List1.Selected(ii) = False Then
            List1.Selected(ii) = True
         End If
      Next
      List1.ListIndex = Index
   End If
End Sub

Private Sub cmdFunc_Click(Index As Integer)
   Select Case Index
   Case 0: ListAllDate
   Case 1: ListCancelDate
   Case 2: SelectAll
   Case 3: ListSelectedDate
   End Select
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 0 '確定
      If TxtValidate Then
         If FormSave Then
            frm140112.RefreshGridData
            frm140112.RefreshGridData2 'Added by Lydia 2025/01/09
            Unload Me
         End If
      End If
      
   Case 1 '取消
      'Add by Amy 2019/01/24 +教育訓練
      If bolIs140113 = False Then
         frm140112.RefreshGridData  '都更新
         frm140112.RefreshGridData2 'Added by Lydia 2025/01/09
      End If
      Unload Me
   End Select
End Sub

'Memo by Amy 2019/01/24 此若修改請確認 frm140113.TxtValidate是否也需修改
Private Function TxtValidate() As Boolean
   Dim ii As Integer, strStartDate As String
   
   strStartDate = Replace(MaskEdBox1, "/", "")
   If strStartDate = "" Then
      MsgBox "請輸入預約日期！", vbExclamation
      If MaskEdBox1.Enabled Then MaskEdBox1.SetFocus
      Exit Function
      
   ElseIf ChkDate(strStartDate) = False Then
      If MaskEdBox1.Enabled Then MaskEdBox1.SetFocus
      Exit Function
      
   'Modified by Morgan 2019/7/3 +Check2
   ElseIf Check1.Value = 0 And Check2.Value = 0 And Pub_StrUserSt03 <> "M51" Then
      
      If DBDATE(strStartDate) < strSrvDate(1) Then
         MsgBox "預約日期已經過了！", vbCritical
         If MaskEdBox1.Enabled Then MaskEdBox1.SetFocus
         Exit Function
         
      Else
         'Modified by Morgan 2019/7/10 改4個月(原2個月)--經理
         strExc(1) = CompDate(1, 4, strSrvDate(1))
         If DBDATE(strStartDate) > strExc(1) Then
            MsgBox "預約日期請輸入 4 個月以內的日期！"
            If MaskEdBox1.Enabled Then MaskEdBox1.SetFocus
            Exit Function
         End If
      End If
   End If
   
   For intI = 0 To 1
      If cboTime(intI) = "" Then
         MsgBox "請點選時間！", vbExclamation
         cboTime(intI).SetFocus
         Exit Function
      End If
   Next
   If cboTime(0) >= cboTime(1) Then
      MsgBox "結束時間必須晚於開始時間！", vbCritical
      cboTime(1).SetFocus
      Exit Function
   End If
   
   'Added by Morgan 2015/8/13
   If strStartDate = strSrvDate(2) Then
      If cboTime(0).Enabled = True Then
         If Val(Replace(cboTime(0), ":", "") & "00") < ServerTime Then
            MsgBox "開始時間已過請重新點選！", vbCritical
            cboTime(0).SetFocus
            Exit Function
         End If
      End If
      
      If cboTime(1).Enabled = True Then
         If Val(Replace(cboTime(1), ":", "") & "00") < ServerTime Then
            MsgBox "結束時間已過請重新點選！", vbCritical
            cboTime(1).SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2015/8/13
      
   If Pub_StrUserSt03 <> "M51" Then
      If txtUser = "" Then
         MsgBox "請輸入預約人員工編號！", vbExclamation
         If txtUser.Enabled Then txtUser.SetFocus
         Exit Function
      End If
   End If
   
   If ChkStaffID(strUserNum) = True Then
      If txtUser.Enabled Then txtUser.SetFocus
      Exit Function
   End If
   
   If txtUser <> "" Then
      If lblUserName = "" Then
         MsgBox "預約人員工編號輸入錯誤！", vbCritical
         If txtUser.Enabled Then txtUser.SetFocus
         Exit Function
         
      'Added by Morgan 2019/5/10
      ElseIf PUB_GetST14(txtUser) = "99997" Then
         MsgBox "【" & lblUserName & "】設定為不寄信，不可為預約人！", vbCritical
         If txtUser.Enabled Then txtUser.SetFocus
         Exit Function
      'end 2019/5/10
      End If
   End If
   
   'Added by Morgan 2015/8/24
   If PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex) = 9 Then
      If txtContent = "" Or txtContent = GetDepartmentName(GetStaffDepartment(txtUser)) Then
         MsgBox "主題欄位中，請加註預定地點及事由！", vbInformation
         txtContent.SetFocus
         Exit Function
      End If
   End If
   'end 2015/8/24
   
   If txtContent = "" Then
      MsgBox "請輸入使用單位(主題)！"
      txtContent.SetFocus
      Exit Function
   ElseIf GetTextLength(txtContent) > txtContent.MaxLength Then
      MsgBox "使用單位(主題)超過 100 字元限制(1個中文算2個字元)！"
      txtContent.SetFocus
      Exit Function
   End If
   
   If m_State = "D" Then
      'Modified by Morgan 2019/7/3 +Check2
      If Check1.Value = 1 Or Check2.Value = 1 Then
         If MsgBox("本預約為週期性，是否確定要取消所有預約？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      Else
         If MsgBox("確定要取消預約？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
   Else
      'Modified by Morgan 2019/7/3 +Check2
      If Check1.Value = 1 Or Check2.Value = 1 Then
         If bolRefresh = True Then
            If RefreshList = False Then
               Exit Function
            End If
         End If
         For ii = LBound(arrDateList) To UBound(arrDateList)
            If arrDateList(ii).Selected = True Then
               Exit For
            End If
         Next
         If ii > UBound(arrDateList) Then
            If m_State = "A" Then
               MsgBox "週期性預約至少要保留一天不取消！", vbInformation
            ElseIf m_State = "E" Then
               MsgBox "週期性預約若要全部取消請改用刪除功能！", vbInformation
            End If
            Exit Function
         End If
      End If
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim rr(16) As String ', oRR(16) As String 'Modify by Amy  2019/01/24 改成全域
   Dim ii As Integer, RD04 As String, RD05 As String, RR19 As String
   Dim strErrMsg As String
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   'Modify by Amy 2019/11/12
   If bolIs140113 = True Then
        rr(1) = strRoomNo
   Else
        rr(1) = PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex)
   End If
   'end 2019/11/12
   
   rr(2) = DBDATE(MaskEdBox1)
   rr(3) = Replace(cboTime(0), ":", "")
   rr(4) = Replace(cboTime(1), ":", "")
   If Check1.Value = 1 Then
      rr(5) = "1"
      rr(6) = DBDATE(MaskEdBox2)
      
   'Added by Morgan 2019/7/3
   ElseIf Check2.Value = 1 Then
      rr(5) = "2"
      rr(6) = DBDATE(MaskEdBox2)
   'end 2019/7/3
   Else
      rr(5) = "N"
      rr(6) = "NULL"
   End If
   rr(7) = txtUser
   rr(8) = txtContent
   If Check3.Value = 0 Then
      rr(9) = "N"
   End If
      
   Select Case m_State
   Case "A" '新增
      rr(10) = strUserNum
      rr(11) = strSrvDate(1)
      rr(12) = "to_char(sysdate,'hh24miss')"
      '週期性預約明細
      If rr(5) <> "N" Then
         '檢查主檔是否重複
         'Modified by Morgan 2015/8/14 +未取消判斷
         strExc(0) = "select 1 from RoomReservation where rr01=" & rr(1) & " and rr02=" & rr(2) & " and rr03=" & rr(3) & " and rr05='" & rr(5) & "' and rr18=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strErrMsg = "該時段已有其他週期性預約,請重新確認！"
            GoTo ErrHnd
         End If
         
         strSql = "insert into RoomReservation(rr01,rr02,rr03,rr04,rr05,rr06,rr07,rr08,rr09,rr10,rr11,rr12)" & _
            " values(" & rr(1) & "," & rr(2) & "," & rr(3) & "," & rr(4) & ",'" & rr(5) & "'," & rr(6) & _
            ",'" & rr(7) & "','" & ChgSQL(rr(8)) & "','" & rr(9) & "','" & rr(10) & "'," & rr(11) & "," & rr(12) & ")"
         cnnConnection.Execute strSql, intI
         
         
         For ii = LBound(arrDateList) To UBound(arrDateList)
            RD04 = DBDATE(arrDateList(ii).Date)
            If arrDateList(ii).Selected = True Then
               If ChkValidate(rr(1), RD04, rr(3), rr(4)) = False Then
                  strErrMsg = "週期性預約建立失敗，[" & RD04 & "] 該時段已有其他預約,請重新確認！"
                  GoTo ErrHnd
               End If
               RD05 = "NULL"
            Else
               RD05 = strSrvDate(1)
            End If
            strSql = "insert into RoomResDetail(rd01,rd02,rd03,rd04,rd05)" & _
               " values(" & rr(1) & "," & rr(2) & "," & rr(3) & "," & RD04 & "," & RD05 & ")"
            cnnConnection.Execute strSql, intI
         Next
      '單次預約
      Else
         
         'Added by Morgan 2015/8/14
         '智權人員借車要記錄次數
         rr(16) = "null"
         '*** 公務車 ***
         If PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex) = 9 Then
            'Modified by Morgan 2016/10/21 +總務(M11)也比照智權人員規則
            'Modified by Morgan 2019/7/12 部門也抓設定(+客戶服務組 W10--文雄)
            'strExc(1) = PUB_GetST03(txtUser)
            'If Left(strExc(1), 1) = "S" Or strExc(1) = "M11" Or strExc(1) = "M10" Then
            strExc(1) = PUB_GetStaffST15(txtUser, 1)
            If Left(strExc(1), 1) = "S" Or InStr(m_Users, strExc(1)) > 0 Then
            'end 2019/7/12
               rr(16) = frm140112.GetTimes(DBDATE(MaskEdBox1), txtUser)
               '申請日晚於申請日時,第1次借車可取代預約期間的其他非第1次預約
               If rr(2) > strSrvDate(1) And rr(16) = "1" Then
                  RR19 = ServerTime
                  strSql = "update RoomReservation set rr17='" & strUserNum & "',rr18=" & strSrvDate(1) & ",RR19=" & RR19 & " where RR01=" & rr(1) & " and RR02=" & rr(2) & _
                     " and ( (rr03<=" & rr(3) & " and rr04>" & rr(3) & ") or (rr03<" & rr(4) & " and rr04>=" & rr(4) & ") or (rr03>" & rr(3) & " and rr04<=" & rr(4) & ") ) and RR05='N' and rr16>1 AND RR18=0"
                  cnnConnection.Execute strSql, intI
                  '預約被取消要發EMail通知
                  If intI > 0 Then
                     strExc(1) = ""
                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                        " select '" & strUserNum & "',rr07,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'通知借車預約已被取消!!('||st02||')'" & _
                        ",'" & lblUserName & "第1次借車優先，您 '||sqldatet(rr02)||' '||to_char(to_date(lpad(rr03,4,'0'),'HH24mi'),'hh24:mi')||'-'||to_char(to_date(lpad(rr04,4,'0'),'HH24mi'),'hh24:mi')||' 的借車預約(第'||rr16||'次)已被取消!!'" & _
                        " from RoomReservation,staff where rr17='" & strUserNum & "' and rr18=" & strSrvDate(1) & " and rr19=" & RR19 & " and st01(+)=rr07"
                     cnnConnection.Execute strSql, intI
                  End If
               End If
            End If
         End If
         '*** end 公務車 ***
         'end 2015/8/14
         
         '檢查預約是否有重疊
         '單次預約-檢查預約是否有重疊(含公務車)
         'Modify by Amy 2019/01/24
         'If ChkValidate(rr(1), rr(2), rr(3), rr(4)) = False Then
         If ChkReservation(rr(1), rr(2), rr(3), rr(4), , , , , , strRR20) = False Then
            strErrMsg = "該時段已有其他預約,請重新確認！"
            GoTo ErrHnd
         End If
         
         'Modified by Morgan 2015/8/14 +rr16
         'Modify by Amy 2019/01/24 +rr20
         strSql = "insert into RoomReservation(rr01,rr02,rr03,rr04,rr05,rr06,rr07,rr08,rr09,rr10,rr11,rr12,rr16" & IIf(bolIs140113 = True Or strRR20 <> "", ",rr20", "") & ")" & _
            " values(" & rr(1) & "," & rr(2) & "," & rr(3) & "," & rr(4) & ",'" & rr(5) & "'," & rr(6) & _
            ",'" & rr(7) & "','" & ChgSQL(rr(8)) & "','" & rr(9) & "','" & rr(10) & "'," & rr(11) & "," & rr(12) & "," & rr(16) & IIf(bolIs140113 = True Or strRR20 <> "", "," & strRR20, "") & ")"
         cnnConnection.Execute strSql, intI
         'end 2015/8/14
      
      End If 'rr(5) <> "N"
      
      
   Case "E" '修改
      rr(13) = strUserNum
      rr(14) = strSrvDate(1)
      rr(15) = "to_char(sysdate,'hh24miss')"
         
      '週期性預約明細 (2019/01/24 Memo by Amy 教育訓練週期預約複雜,故未開放)
      If rr(5) <> "N" Then
         strSql = "Update RoomReservation" & _
            " set RR07='" & rr(7) & "',RR08='" & ChgSQL(rr(8)) & "',RR09='" & rr(9) & "'" & _
            ",RR13='" & rr(13) & "',RR14=" & rr(14) & ",RR15=" & rr(15) & _
            " where RR01=" & rr(1) & " and RR02=" & rr(2) & " and RR03=" & rr(3) & " and RR05='" & rr(5) & "'"
            
         cnnConnection.Execute strSql, intI
         
         '明細都刪除重新建立
         strSql = "delete RoomResDetail where rd01=" & rr(1) & " and rd02=" & rr(2) & " and rd03=" & rr(3)
         cnnConnection.Execute strSql, intI
         
         For ii = LBound(arrDateList) To UBound(arrDateList)
            RD04 = DBDATE(arrDateList(ii).Date)
            If arrDateList(ii).Selected = True Then
               If ChkValidate(rr(1), RD04, rr(3), rr(4)) = False Then
                  strErrMsg = "週期性預約建立失敗，[" & RD04 & "] 該時段已有其他預約,請重新確認！"
                  GoTo ErrHnd
               End If
               RD05 = "NULL"
            Else
               RD05 = strSrvDate(1)
            End If
            strSql = "insert into RoomResDetail(rd01,rd02,rd03,rd04,rd05)" & _
               " values(" & rr(1) & "," & rr(2) & "," & rr(3) & "," & RD04 & "," & RD05 & ")"
            cnnConnection.Execute strSql, intI
         Next
      '單次預約
      Else
        If bolIs140113 = False Then
            oRR(1) = cboRoom.Tag
            oRR(2) = DBDATE(MaskEdBox1.Tag)
            oRR(3) = Replace(cboTime(0).Tag, ":", "")
            oRR(4) = Replace(cboTime(1).Tag, ":", "") 'Modify by Amy 2019/01/24 由下搬上來
        End If
     
         '單次預約-檢查預約是否有重疊
         'Modify by Amy 2019/01/24
         'If ChkValidate(rr(1), rr(2), rr(3), rr(4), True, oRR(1), oRR(2), oRR(3)) = False Then
         If ChkReservation(rr(1), rr(2), rr(3), rr(4), True, oRR(1), oRR(2), oRR(3), IIf(bolIs140113 = True Or strRR20 <> "", oRR(4), ""), strRR20) = False Then
            strErrMsg = "該時段已有其他預約,請重新確認！"
            GoTo ErrHnd
         End If
         
         'oRR(4) = Replace(cboTime(1).Tag, ":", "") 'Mark by Amy 2019/01/24 往搬上
         
         oRR(7) = txtUser.Tag
         oRR(8) = txtContent.Tag
         If Val(Check3.Tag) = 0 Then
            oRR(9) = "N"
         End If
         
         If rr(1) <> oRR(1) Or rr(2) <> oRR(2) Or rr(3) <> oRR(3) Or rr(4) <> oRR(4) Or rr(7) <> oRR(7) Or rr(8) <> oRR(8) Or rr(9) <> oRR(9) Then
            rr(13) = strUserNum
            rr(14) = strSrvDate(1)
            rr(15) = "to_char(sysdate,'hh24miss')"
            'Modify by Amy 2019/01/24 +RR20
            strSql = "And RR01=" & oRR(1) & " and RR02=" & oRR(2) & " and RR03=" & oRR(3) & " "
            If bolIs140113 = True Or strRR20 <> MsgText(601) Then
                strSql = " And RR20='" & strRR20 & "' "
            End If
            strSql = "Update RoomReservation" & _
               " set RR01=" & rr(1) & ",RR02=" & rr(2) & ",RR03=" & rr(3) & ",RR04=" & rr(4) & _
               ",RR07='" & rr(7) & "',RR08='" & ChgSQL(rr(8)) & "',RR09='" & rr(9) & "'" & _
               ",RR13='" & rr(13) & "',RR14=" & rr(14) & ",RR15=" & rr(15) & _
               " Where RR05='N' " & strSql
            'end 2019/01/24
            cnnConnection.Execute strSql, intI
         End If
      End If 'rr(5) <> "N"
      
   Case "D" '刪除
      '週期性預約刪除明細
      If rr(5) <> "N" Then
         strSql = "delete RoomResDetail" & _
            " where RD01=" & rr(1) & " and RD02=" & rr(2) & " and RD03=" & rr(3)
         cnnConnection.Execute strSql, intI
      End If
      
      'Added by Morgan 2015/8/14
      '當日取消借車要保留記錄
      'Modified by Morgan 2017/10/13 +M11,M10
      'Modified by Morgan 2019/7/12 部門也抓設定(+客戶服務組 W10--文雄)
      'strExc(1) = PUB_GetST03(txtUser)
      'If cboRoom.ItemData(cboRoom.ListIndex) = 9 And (Left(strExc(1), 1) = "S" Or strExc(1) = "M11" Or strExc(1) = "M10") Then
      strExc(1) = PUB_GetStaffST15(txtUser, 1)
      If PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex) = 9 And (Left(strExc(1), 1) = "S" Or InStr(m_Users, strExc(1)) > 0) Then
      'end 2019/7/12
         If rr(2) = strSrvDate(1) Then
            strSql = "update RoomReservation set RR17='" & strUserNum & "',rr18=" & strSrvDate(1) & ",RR19=TO_CHAR(SYSDATE,'HH24MISS')" & _
               " where RR01=" & rr(1) & " and RR02=" & rr(2) & " and RR03=" & rr(3) & " and RR05='" & rr(5) & "' and rr18=0"
            cnnConnection.Execute strSql, intI
         Else
            strSql = "delete RoomReservation" & _
               " where RR01=" & rr(1) & " and RR02=" & rr(2) & " and RR03=" & rr(3) & " and RR05='" & rr(5) & "' and rr18=0"
            cnnConnection.Execute strSql, intI
            '更新借車次數
            UpdateTimes rr(2), txtUser
         End If
         
         'Added by Morgan 2015/8/18
         '原時段的預約取消記錄也一併刪除(當日取消的除外)
         strSql = "delete RoomReservation a" & _
            " where RR01=" & rr(1) & " and RR02=" & rr(2) & " and rr18<rr02" & _
            " and ( (rr03<=" & rr(3) & " and rr04>" & rr(3) & ") or (rr03<" & rr(4) & " and rr04>=" & rr(4) & ") or (rr03>" & rr(3) & " and rr04<=" & rr(4) & ") ) " & _
            " and not exists(select * from RoomReservation b where b.rr01=a.rr01 and b.rr02=a.rr02 and b.rr18=0 " & _
            " and ( (b.rr03<=a.rr03 and b.rr04>a.rr03) or (b.rr03<a.rr04 and b.rr04>=a.rr04) or (b.rr03>a.rr03 and b.rr04<=a.rr04) ))"
         cnnConnection.Execute strSql, intI
      Else
      'end 2015/8/14
         strSql = "delete RoomReservation" & _
            " where RR01=" & rr(1) & " and RR02=" & rr(2) & " and RR03=" & rr(3) & " and RR05='" & rr(5) & "'"
         cnnConnection.Execute strSql, intI
      End If
   End Select
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If strErrMsg <> "" Then
      MsgBox strErrMsg, vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Function UpdateTimes(pDate As String, Optional pUserNo As String)
   Dim stSQL As String, intR As Integer, iRR16 As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stStartDate As String, stEndDate As String
   
   If pUserNo = "" Then pUserNo = strUserNum
   
   intR = Weekday(ChangeWStringToWDateString(pDate))
   If intR = 1 Then
      stStartDate = pDate
   Else
      stStartDate = CompDate(2, -1 * (intR - 1), pDate)
   End If
   If intR = 7 Then
      stEndDate = pDate
   Else
      stEndDate = CompDate(2, 7 - intR, pDate)
   End If
   
   '當日取消也要算
   stSQL = "select rr01,rr02,rr03,rr05,rr16,rr18,rr19 from RoomReservation" & _
      " where rr07='" & pUserNo & "' and rr01=9 and rr02>=" & stStartDate & " and rr02<=" & stEndDate & " and (rr18=0 or rr02=rr18) order by rr16 asc"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
      iRR16 = 0
      Do While Not .EOF
         iRR16 = iRR16 + 1
         If iRR16 <> Val("" & .Fields("rr16")) Then
            stSQL = "update RoomReservation set rr16=" & iRR16 & " where rr01=" & .Fields("rr01") & " and rr02=" & .Fields("rr02") & " and rr03=" & .Fields("rr03") & " and rr05='" & .Fields("rr05") & "' and rr18=" & .Fields("rr18") & " and rr19=" & .Fields("rr19")
            cnnConnection.Execute stSQL, intR
         End If
         .MoveNext
      Loop
      End With
   End If
   Set rsQuery = Nothing
End Function

Public Sub SetEnable()
   Select Case m_State
      Case "A"
         Me.Caption = Me.Caption & " - 新增"
         'Modify by Amy 2019/01/24 +bolIs140113判斷
         If bolIs140113 = False And (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M12") Then
            txtUser.Enabled = True
         Else
            txtUser.Enabled = False
         End If
         
         If bolIs140113 = False And Pub_StrUserSt03 = "M51" Then
            Frame1.Enabled = True
            Check1.Enabled = True
            Check2.Enabled = True 'Added by Morgan 2019/7/3
            MaskEdBox2.Enabled = True
         Else
            Frame1.Enabled = False
         End If
         Check3.Enabled = False
      Case "E"
         Me.Caption = Me.Caption & " - 修改"
         
         If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M12" Then
            txtUser.Enabled = True
         Else
            txtUser.Enabled = False
         End If
         
         txtContent.Enabled = True
         Check3.Enabled = False
         
         'Modified by Morgan 2019/7/3 +Check2
         If Check1.Value = 0 And Check2.Value = 0 Then
            cboRoom.Enabled = True
            MaskEdBox1.Enabled = True
            cboTime(0).Enabled = True
            cboTime(1).Enabled = True
            'Add by Amy 2020/02/06 +if 教育訓練進入不判斷
            If bolIs140113 = False And strRR20 = MsgText(601) Then
                'Added by Morgan 2015/8/13 使用完不可改,使用中只能改結束時間
                If DBDATE(MaskEdBox1) <= strSrvDate(1) Then
                   cboRoom.Enabled = False
                   MaskEdBox1.Enabled = False
                   If DBDATE(MaskEdBox1) < strSrvDate(1) Then
                      cboTime(0).Enabled = False
                      cboTime(1).Enabled = False
                   Else
                      If Val(Replace(cboTime(1), ":", "") & "00") < ServerTime Then
                         cboTime(0).Enabled = False
                         cboTime(1).Enabled = False
                      ElseIf Val(Replace(cboTime(0), ":", "") & "00") < ServerTime Then
                         cboTime(0).Enabled = False
                      End If
                   End If
                End If
            End If
            'end 2015/8/13
            Frame1.Enabled = False
         Else
            cboRoom.Enabled = False
            MaskEdBox1.Enabled = False
            cboTime(0).Enabled = False
            cboTime(1).Enabled = False
            cboTime(1).Enabled = False
            Frame1.Enabled = True
            Check1.Enabled = False
            Check2.Enabled = False 'Added by Morgan 2019/7/3
            MaskEdBox2.Enabled = False
         End If

      Case "D"
         Me.Caption = Me.Caption & " - 刪除"
         cboRoom.Enabled = False
         MaskEdBox1.Enabled = False
         cboTime(0).Enabled = False
         cboTime(1).Enabled = False
         txtUser.Enabled = False
         txtContent.Enabled = False
         Check3.Enabled = False
         Frame1.Enabled = False
         Command1(0).Caption = "刪除"
         
      Case "S"
         Me.Caption = Me.Caption & " - 檢視"
         cboRoom.Enabled = False
         MaskEdBox1.Enabled = False
         cboTime(0).Enabled = False
         cboTime(1).Enabled = False
         txtUser.Enabled = False
         txtContent.Enabled = False
         Check3.Enabled = False
         Frame1.Enabled = False
         Command1(0).Visible = False
         
   End Select
   
   cboRoom.Enabled = False 'Added by Morgan 2015/8/12
 
End Sub

Private Sub ListCancelDate()
   Dim ii As Single
   List1.Visible = False
   List1.Clear
   For ii = LBound(arrDateList) To UBound(arrDateList)
      If arrDateList(ii).Selected = False Then
         List1.AddItem arrDateList(ii).Date
      End If
   Next
   If List1.ListCount > 0 Then
      List1.ListIndex = 0
   End If
   List1.Visible = True
End Sub

Private Sub ListSelectedDate()
   Dim ii As Single, jj As Single
   List1.Visible = False
   List1.Clear
   jj = 0
   For ii = LBound(arrDateList) To UBound(arrDateList)
      If arrDateList(ii).Selected = True Then
         List1.AddItem arrDateList(ii).Date, jj
         List1.Selected(jj) = True
         jj = jj + 1
      End If
   Next
   If List1.ListCount > 0 Then
      List1.ListIndex = 0
   End If
   List1.Visible = True
End Sub

Private Sub Form_Activate()
   If txtUser.Enabled Then txtUser.SetFocus
   'Added by Morgan 2015/8/14
   SetlblEmail
   If m_State = "A" Then
      'Modify by Amy 2019/11/12 +if 否則會error
      If Not (bolIs140113 = True Or strRR20 = MsgText(601)) Then
        SetlblTimes
        SetlblOldUser
      End If
   End If
   'end 2015/8/14
End Sub

Private Sub Form_Load()
    Dim strTp(1 To 4) As String 'Add by Amy 2019/01/24
    
    If Pub_StrUserSt03 <> "M51" Then
       Me.Height = 4050
    End If
   'Modify by Amy 2019/11/12
   If (bolIs140113 = True Or strRR20 <> MsgText(601)) And m_State <> "S" And m_State <> "D" Then
     m_State = "A"
     If ChkHasRR20(Val(strRR20), strTp(1), strTp(2), strTp(3), strTp(4), True) = True Then
        m_State = "E"
        oRR(1) = strTp(1)
        oRR(2) = strTp(2)
        oRR(3) = strTp(3)
        oRR(4) = strTp(4)
     End If
   End If
   MoveFormToCenter Me
   SetCombo
   lblTimes = "" 'Added by Morgan 2015/8/14
   Text1 = strRR20 'Add by Amy 2019/01/24
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim stTP(1 To 4) As String 'Add by Amy 2019/12/09
   
   PUB_SendMailCache 'Added by Morgan 2015/8/14
   'Add by Amy 2019/01/24
   bolIs140113 = False
   strRR20 = ""
   'end 2019/01/24
   Set frm140112_1 = Nothing
End Sub

Private Sub List1_Click()
   Dim ii As Single, jj As Single
   If List1.Visible = True Then
      For ii = 0 To List1.ListCount - 1
         For jj = ii To UBound(arrDateList)
            If arrDateList(jj).Date = List1.List(ii) Then
               arrDateList(jj).Selected = List1.Selected(ii)
               Exit For
            End If
         Next
      Next
   End If
End Sub

Private Sub MaskEdBoxInverse(pBox As MaskEdBox)
   pBox.SelStart = 0
   pBox.SelLength = Len(pBox.Text)
End Sub

Private Sub MaskEdBox1_GotFocus()
   CloseIme
   MaskEdBoxInverse MaskEdBox1
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If m_State = "A" Then SetlblTimes
End Sub

Private Sub SetlblTimes()
   lblTimes = ""
   If PUB_GetItemData(cboRoomItemData, Me.cboRoom.ListIndex) = 9 Then
      'Modified by Morgan 2016/10/21 +總務(M11)也比照智權人員規則
      'Modified by Morgan 2019/7/12 部門也抓設定(+客戶服務組 W10--文雄)
      'strExc(1) = PUB_GetST03(txtUser)
      'If Left(strExc(1), 1) = "S" Or strExc(1) = "M11" Or strExc(1) = "M10" Then
      strExc(1) = PUB_GetStaffST15(txtUser, 1)
      If Left(strExc(1), 1) = "S" Or InStr(m_Users, strExc(1)) > 0 Then
      'end 2019/7/12
         lblTimes = "第" & frm140112.GetTimes(DBDATE(MaskEdBox1), txtUser) & "次"
      End If
   End If
End Sub

Public Sub SetlblOldUser()
   Dim rr02 As Long, rr03 As Integer, rr04 As Integer, stCon As String
   lblOldUser = ""
   If PUB_GetItemData(cboRoomItemData, cboRoom.ListIndex) = 9 And txtUser <> "" Then
      rr02 = DBDATE(MaskEdBox1)
      rr03 = Val(Replace(cboTime(0), ":", ""))
      rr04 = Val(Replace(cboTime(1), ":", ""))
      If rr02 > 0 And rr03 > 0 And rr04 > 0 Then
         If m_State = "A" Then
            stCon = " and rr18=0"
         Else
            stCon = " and rr18>0 and rr17='" & txtUser & "'"
         End If
         strExc(0) = "select distinct st02 from roomreservation,staff where rr01=9 and rr02=" & rr02 & _
            " and ( (rr03<=" & rr03 & " and rr04>" & rr03 & ") or (rr03<" & rr04 & " and rr04>=" & rr04 & ") or (rr03>" & rr03 & " and rr04<=" & rr04 & ") ) and rr16>1 and rr07<>'" & txtUser & "'" & stCon & " and st01(+)=rr07"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            lblOldUser = "原時段借車人 " & .Fields("st02")
            .MoveNext
            Do While Not .EOF
               lblOldUser = lblOldUser & "、" & .Fields("st02")
               .MoveNext
            Loop
            End With
         End If
      End If
   End If
End Sub

Private Sub MaskEdBox2_Change()
   ResetList
End Sub

Private Sub MaskEdBox2_GotFocus()
   CloseIme
   MaskEdBoxInverse MaskEdBox2
End Sub

Private Sub txtContent_GotFocus()
   TextInverse txtContent
End Sub

Private Sub txtContent_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtContent
End Sub

Private Sub txtUser_Change()
   Dim stDeptNo As String, stDeptN As String 'Add by Amy 2020/11/27
   
   If Len(txtUser) >= 5 Then
      'Modify by Amy 2020/11/27 部門重新編制前先抓SC03沒才抓GetStaffDepartment 原:txtContent = GetDepartmentName(GetStaffDepartment(txtUser))
      'Modify by Amy 2020/12/22 杜經理操作智權部教育訓練,部門需以st15抓
      stDeptNo = GetStaffDepartment(txtUser, IIf(strRR20 <> "", True, False))
      lblUserName = GetStaffName(txtUser, True)
      If m_State <> "S" And lblUserName <> "" Then
         stDeptN = GetSC03(stDeptNo)
         If stDeptN = MsgText(601) Then
            stDeptN = GetDepartmentName(stDeptNo)
         End If
          txtContent = stDeptN
          'end 2020/11/27
      End If
      If m_State = "A" Then SetlblTimes 'Added by Morgan 2015/8/14
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtUser_GotFocus()
   CloseIme
   TextInverse txtUser
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub SetCombo()
   Dim ii As Integer
   For ii = 0 To 23
      cboTime(0).AddItem Format(ii, "00") & ":" & "00"
      cboTime(1).AddItem Format(ii, "00") & ":" & "30"
      cboTime(0).AddItem Format(ii, "00") & ":" & "30"
      cboTime(1).AddItem Format(ii + 1, "00") & ":" & "00"
   Next
End Sub

Private Function ChkValidate(pRoom As String, pDate As String, pFromTime As String, pToTime As String, Optional bolEdit As Boolean _
   , Optional pOldRoom As String, Optional pOldDate As String, Optional pOldFromTime As String) As Boolean
   Dim strCon As String
   If bolEdit Then
      strCon = " and not (rr01=" & pOldRoom & " and rr02=" & pOldDate & " and rr03=" & pOldFromTime & ")"
   End If
   
   'Added by Morgan 2015/8/14 +排除已取消者
   strCon = strCon & " and rr18=0"
   'end 2015/8/14
   'Add by Amy 2019/11/12 由教育訓練進入排除教育訓練預約
   If bolIs140113 = True Or strRR20 <> MsgText(601) Then
        strCon = strCon & " And rr20<>" & Val(strRR20)
   End If
   '檢查預約是否有重疊
   strExc(0) = "select 1 from RoomReservation" & _
      " where rr01=" & pRoom & " and rr02=" & pDate & " and " & pFromTime & ">=rr03 and " & pFromTime & "<rr04 and rr05='N'" & strCon & _
      " union select 1 from RoomReservation" & _
      " where rr01=" & pRoom & " and rr02=" & pDate & " and " & pToTime & ">rr03 and " & pToTime & "<=rr04 and rr05='N'" & strCon & _
      " union select 1 from RoomReservation" & _
      " where rr01=" & pRoom & " and rr02=" & pDate & " and " & pFromTime & "<rr03 and " & pToTime & ">=rr04 and rr05='N'" & strCon & _
      " union select 1 from RoomResDetail,RoomReservation" & _
      " where rd01=" & pRoom & " and rd04=" & pDate & " and rd05 is null" & _
      " and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03 and " & pFromTime & ">=rr03 and " & pFromTime & "<rr04" & strCon & _
      " union select 1 from RoomResDetail,RoomReservation" & _
      " where rd01=" & pRoom & " and rd04=" & pDate & " and rd05 is null" & _
      " and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03 and " & pToTime & ">rr03 and " & pToTime & "<=rr04" & strCon & _
      " union select 1 from RoomResDetail,RoomReservation" & _
      " where rd01=" & pRoom & " and rd04=" & pDate & " and rd05 is null" & _
      " and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03 and " & pFromTime & "<rr03 and " & pToTime & ">=rr04" & strCon

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      ChkValidate = True
   End If
End Function

'Add by Amy 2019/01/24 依教育訓練編號抓取資料
Public Sub ReadData()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select * From RoomReservation Where RR20=" & Val(strRR20)
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If RsQ.RecordCount > 0 Then
            cboTime(0) = Format("" & RsQ.Fields("RR03"), "00:00")
            cboTime(0).Tag = cboTime(0)
            cboTime(1) = Format("" & RsQ.Fields("RR04"), "00:00")
            cboTime(1).Tag = cboTime(1)
            MaskEdBox1.Mask = ""
            MaskEdBox1.Text = CFDate(Val("" & RsQ.Fields("RR02")) - 19110000)
            MaskEdBox1.Tag = MaskEdBox1.Text
            MaskEdBox1.Mask = DFormat
            MaskEdBox2.Mask = DFormat
            txtUser = strUserNum
            txtUser.Tag = txtUser
            txtContent = "" & RsQ.Fields("RR08")
            txtContent.Tag = txtContent
            If IsNull(RsQ.Fields("RR09")) Then Check3.Value = 1
            Check3.Tag = Check3.Value
        End If
    End If
    RsQ.Close
    
End Sub

'Add by Amy 2020/11/27
Private Function GetSC03(ByVal stDeptNo As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    Dim bolOpen As Boolean
   
    GetSC03 = "": bolOpen = True
    If Left(stDeptNo, 2) = "F2" Then bolOpen = False '外專公開只有一筆資料共用,故抓「不公開」
    strQ = GetSeminarContactSql(2, stDeptNo, bolOpen, , txtUser)
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetSC03 = "" & RsQ.Fields("SC03")
    End If
    
    Set RsQ = Nothing
End Function

