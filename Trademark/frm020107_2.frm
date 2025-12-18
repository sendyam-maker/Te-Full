VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020107_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "T案大陸指示信"
   ClientHeight    =   5745
   ClientLeft      =   30
   ClientTop       =   945
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox txtLetterHead 
      Height          =   300
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2805
      Width           =   435
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2685
      Left            =   30
      TabIndex        =   55
      Top             =   3060
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4736
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "申請"
      TabPicture(0)   =   "frm020107_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstText"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.ListBox lstText 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         ItemData        =   "frm020107_2.frx":001C
         Left            =   3810
         List            =   "frm020107_2.frx":001E
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   360
         Width           =   1725
      End
      Begin VB.Frame Frame2 
         Height          =   825
         Left            =   5550
         TabIndex        =   57
         Top             =   480
         Width           =   3285
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 加入"
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   180
            Width           =   735
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Left            =   90
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   0
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "T"
            Top             =   180
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   1
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   16
            Top             =   180
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   2
            Left            =   2370
            MaxLength       =   1
            TabIndex        =   17
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   3
            Left            =   2760
            MaxLength       =   2
            TabIndex        =   18
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Height          =   180
            Left            =   870
            TabIndex        =   60
            Top             =   540
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Line Line2 
            X1              =   1170
            X2              =   3030
            Y1              =   300
            Y2              =   300
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   675
         Left            =   180
         TabIndex        =   56
         Top             =   390
         Width           =   2535
         Begin VB.CheckBox Check1 
            Caption         =   "多件申請，同圖，不同類別"
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   14
            Top             =   330
            Width           =   2475
         End
         Begin VB.CheckBox Check1 
            Caption         =   "多件申請，同類別，不同圖"
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   13
            Top             =   60
            Width           =   2475
         End
      End
      Begin VB.Label Label3 
         Caption         =   "其他案件："
         Height          =   165
         Left            =   2880
         TabIndex        =   59
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務(&I)"
      Height          =   315
      Left            =   4050
      TabIndex        =   2
      Top             =   30
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "已設定代表圖(&I)"
      Height          =   315
      Left            =   5490
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   30
      Width           =   1395
   End
   Begin VB.TextBox textTM21 
      Height          =   300
      Left            =   900
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1590
      Width           =   1092
   End
   Begin VB.TextBox textTM22 
      Height          =   300
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1590
      Width           =   1092
   End
   Begin VB.TextBox textCP44 
      Height          =   300
      Left            =   5370
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2490
      Width           =   975
   End
   Begin VB.TextBox textTM23 
      Height          =   300
      Left            =   900
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1890
      Width           =   975
   End
   Begin VB.TextBox textTM79 
      Height          =   300
      Left            =   900
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2190
      Width           =   975
   End
   Begin VB.TextBox textTM81 
      Height          =   300
      Left            =   900
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2490
      Width           =   975
   End
   Begin VB.TextBox textTM78 
      Height          =   300
      Left            =   5370
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1890
      Width           =   975
   End
   Begin VB.TextBox textTM80 
      Height          =   300
      Left            =   5370
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   9
      Top             =   2190
      Width           =   975
   End
   Begin VB.TextBox textCP10 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6150
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   38
      Top             =   397
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   315
      Index           =   1
      Left            =   7770
      TabIndex        =   3
      Top             =   30
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   6930
      TabIndex        =   0
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   21
      Top             =   6240
      Width           =   972
   End
   Begin VB.TextBox textTM04 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   28
      Top             =   676
      Width           =   375
   End
   Begin VB.TextBox textTM03 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2220
      MaxLength       =   1
      TabIndex        =   27
      Top             =   676
      Width           =   255
   End
   Begin VB.TextBox textTM02 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   26
      Top             =   676
      Width           =   855
   End
   Begin VB.TextBox textTM01 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   915
      MaxLength       =   3
      TabIndex        =   25
      Top             =   676
      Width           =   495
   End
   Begin MSForms.TextBox textTM05 
      Height          =   300
      Left            =   900
      TabIndex        =   65
      Top             =   977
      Width           =   7905
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "13944;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP44Nm 
      Height          =   255
      Left            =   6405
      TabIndex        =   64
      Top             =   2513
      Width           =   2295
      VariousPropertyBits=   27
      Caption         =   "lblCP44Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP13Nm 
      Height          =   255
      Left            =   930
      TabIndex        =   62
      Top             =   1335
      Width           =   2625
      VariousPropertyBits=   27
      Caption         =   "lblCP13Nm"
      Size            =   "4630;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP14Nm 
      Height          =   255
      Left            =   5370
      TabIndex        =   63
      Top             =   1335
      Width           =   2295
      VariousPropertyBits=   27
      Caption         =   "lblCP14Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否印信頭及簽章:             (N:不印)"
      Height          =   180
      Left            =   120
      TabIndex        =   61
      Top             =   2850
      Width           =   2715
   End
   Begin VB.Label Label5 
      Caption         =   "專用期限:"
      Height          =   225
      Left            =   90
      TabIndex        =   54
      Top             =   1635
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2280
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "商品類別:"
      Height          =   180
      Left            =   4560
      TabIndex        =   53
      Top             =   1650
      Width           =   765
   End
   Begin MSForms.Label lblTM09 
      Height          =   255
      Left            =   5370
      TabIndex        =   52
      Top             =   1613
      Width           =   2295
      VariousPropertyBits=   27
      Caption         =   "lblTM09"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1"
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   51
      Top             =   1950
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3"
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   50
      Top             =   2250
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5"
      Height          =   195
      Index           =   4
      Left            =   210
      TabIndex        =   49
      Top             =   2550
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2"
      Height          =   180
      Index           =   5
      Left            =   4695
      TabIndex        =   48
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4"
      Height          =   180
      Index           =   6
      Left            =   4695
      TabIndex        =   47
      Top             =   2250
      Width           =   630
   End
   Begin MSForms.Label lblTM23Nm 
      Height          =   255
      Left            =   1965
      TabIndex        =   46
      Top             =   1913
      Width           =   2295
      VariousPropertyBits=   27
      Caption         =   "lblTM23Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblTM78Nm 
      Height          =   255
      Left            =   6405
      TabIndex        =   45
      Top             =   1913
      Width           =   2295
      VariousPropertyBits=   27
      Caption         =   "lblTM78Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblTM79Nm 
      Height          =   255
      Left            =   1965
      TabIndex        =   44
      Top             =   2213
      Width           =   2295
      VariousPropertyBits=   27
      Caption         =   "lblTM79Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblTM80Nm 
      Height          =   255
      Left            =   6405
      TabIndex        =   43
      Top             =   2213
      Width           =   2295
      Caption         =   "lblTM80Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblTM81Nm 
      Height          =   255
      Left            =   1965
      TabIndex        =   42
      Top             =   2513
      Width           =   2295
      Caption         =   "lblTM81Nm"
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   195
      Left            =   4695
      TabIndex        =   41
      Top             =   2550
      Width           =   630
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   5340
      TabIndex        =   40
      Top             =   457
      Width           =   765
   End
   Begin VB.Label lblCP10Nm 
      Caption         =   "lblCP10Nm"
      Height          =   255
      Left            =   6720
      TabIndex        =   39
      Top             =   420
      Width           =   2055
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱:"
      Height          =   180
      Left            =   90
      TabIndex        =   37
      Top             =   1037
      Width           =   765
   End
   Begin VB.Label lblNation 
      Caption         =   "lblNation"
      Height          =   255
      Left            =   3780
      TabIndex        =   36
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   1
      Left            =   2970
      TabIndex        =   35
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   5340
      TabIndex        =   34
      Top             =   736
      Width           =   765
   End
   Begin VB.Label lblCP07 
      Caption         =   "lblCP07"
      Height          =   255
      Left            =   6180
      TabIndex        =   33
      Top             =   750
      Width           =   1290
   End
   Begin VB.Label lblCP06 
      Caption         =   "lblCP06"
      Height          =   255
      Left            =   3780
      TabIndex        =   32
      Top             =   750
      Width           =   1410
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "承 辦 人:"
      Height          =   180
      Left            =   4650
      TabIndex        =   31
      Top             =   1372
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   2970
      TabIndex        =   30
      Top             =   736
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   90
      TabIndex        =   29
      Top             =   736
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   1380
      Width           =   765
   End
   Begin VB.Label lblCP09 
      Caption         =   "lblCP09"
      Height          =   255
      Left            =   915
      TabIndex        =   23
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   22
      Top             =   457
      Width           =   585
   End
End
Attribute VB_Name = "frm020107_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/06 改成Form2.0 ; lblCP13Nm、lblCP14Nm、lblCP13Nm、lblTM09、lblTM23Nm、lblTM78Nm、lblTM79Nm、lblTM80Nm、lblTM81Nm、lblCP44Nm
'Memo by Lydia 2021/10/06 注意：因為本程式當時是以Word選項->儲存->以Word97-2003文件(*.doc)為預設Word格式來設計，
                                      '若預設為*.docx檔當有插圖時會自動插入位置不準；目前尚未調整程式寫法，先以改變使用者的預設Word格式來解決。
'Create By Sindy 2014/4/29
Option Explicit

Dim strReceiveNo As String
Dim tm() As String, cp() As String
Dim intWhere As Integer
Public ChkTG As Boolean '檢查是否已經有商品及服務

'加入代表圖用
'Const msoBringInFrontOfText = 4
'Const msoFalse = 0
'Const msoLineSolid = 1
'Const msoLineSingle = 1
Const msoTrue = -1
'Const msoPictureAutomatic = 1

Dim strFilePathN As String


Private Sub Check1_Click(Index As Integer)
   If Check1(Index).Value = 1 Then
      If Index = 0 Then
         Check1(1).Value = 0
      Else
         Check1(0).Value = 0
      End If
      Frame2.Enabled = True
   Else
      Frame2.Enabled = False
   End If
End Sub

'加入本所案號
Private Sub cmdAdd_Click()
Dim rsA As New ADODB.Recordset
   
   '檢查是否有此案號存在,並且為大陸案
   Text1(1) = Right("000000" & Text1(1), 6)
   Text1(2) = Right("0" & Text1(2), 1)
   Text1(3) = Right("00" & Text1(3), 2)
   lblText = Text1(0) & "-" & Text1(1) & "-" & Text1(2) & "-" & Text1(3)
   If textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 = Text1(0) & "-" & Text1(1) & "-" & Text1(2) & "-" & Text1(3) Then
      MsgBox "本所案號重覆！"
      Text1(1).SetFocus
      Exit Sub
   End If
   strExc(0) = "select tm10 from trademark where tm01='" & Text1(0) & "' and tm02='" & Text1(1) & "' and tm03='" & Text1(2) & "' and tm04='" & Text1(3) & "'"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If rsA.Fields(0) <> "020" Then
         MsgBox "此案" & lblText & "非大陸案！"
         Text1(1).SetFocus
         Exit Sub
      End If
   Else
      MsgBox "無此案號" & lblText & "資料！"
      Text1(1).SetFocus
      Exit Sub
   End If
   AddlstText
   Text1(1).SetFocus
   Set rsA = Nothing
End Sub

'移除本所案號
Private Sub cmdRemove_Click()
   RemovelstText
   Text1(1).SetFocus
End Sub

Private Sub AddlstText()
   Dim idx As Integer, bFound As Boolean
   If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" And Text1(3) <> "" Then
      For idx = 0 To lstText.ListCount - 1
         If lstText.List(idx) = lblText Then
            MsgBox "本所案號已存在於清單中！"
            Text1(1).SetFocus
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         lstText.AddItem lblText
         Text1(1) = "": Text1(2) = "": Text1(3) = ""
         lblText = ""
      End If
   End If
End Sub

Private Sub RemovelstText()
   Dim idx As Integer, ii As Integer
   If lstText.ListCount > 0 Then
      ii = 0
      For idx = 0 To lstText.ListCount - 1
         If lstText.Selected(ii) = True Then
            lstText.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub

Private Sub Command2_Click()
'   frm03010303_04.Hide
'   Set frm03010303_04.UpForm = Me
'   frm03010303_04.TGKey = textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04
'   frm03010303_04.AllClass = Trim(lblTM09.Caption)
'   frm03010303_04.cmdok(2).Visible = True
'
''   If m_EditMode <> 1 And m_EditMode <> 2 Then
'       frm03010303_04.Label2.Visible = False
'       frm03010303_04.cmdok(0).Visible = False
'       frm03010303_04.cmdok(2).Visible = False
'       frm03010303_04.cmd.Visible = False
'       frm03010303_04.cmd2.Visible = False
'       frm03010303_04.txt2(0).Visible = False
'       frm03010303_04.txt2(1).Visible = False
'       frm03010303_04.txt2(2).Visible = False
'       frm03010303_04.txt2(3).Visible = False
'       frm03010303_04.Line1.Visible = False
''   End If
'   If Trim(lblTM09.Caption) <> "" Then
'      Me.Hide
'      frm03010303_04.QueryData
'      frm03010303_04.Show vbModal
'   Else
'      MsgBox ("無商品類別，不可使用此按鈕 !")
'   End If
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04
   frm03010303_04.AllClass = Trim(lblTM09.Caption)
   frm03010303_04.Caption = "商品及服務資料"
   frm03010303_04.Label2.Visible = True
   Me.Hide
   frm03010303_04.QueryData
   frm03010303_04.Show vbModal
End Sub

Private Sub Command3_Click()
   frmPic001.oCP01 = textTM01
   frmPic001.oCP02 = textTM02
   frmPic001.oCP03 = textTM03
   frmPic001.oCP04 = textTM04
   frmPic001.StrMenu
   frmPic001.CanScan
   frmPic001.SetSeekCmdok 'Add by Amy 2018/07/20
   frmPic001.Show vbModal
   '檢查有無代表圖
   'Modify by Amy 2018/07/20  改寫至function
'   strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & textTM01 & "' and ibf02='" & textTM02 & "' and ibf03='" & textTM03 & "' and ibf04='" & textTM04 & "' and ibf05='1'"
'   CheckOC2
'   adoRecordset1.CursorLocation = adUseClient
'   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
  If ChkImgByteFile(textTM01, textTM02, textTM03, textTM04) = True Then
       Command3.Caption = "已設定代表圖(&I)"
       Command3.BackColor = &HC0FFC0
   Else
       Command3.Caption = "未設定代表圖(&I)"
       Command3.BackColor = &HC0C0FF
   End If
'   CheckOC2
   'end 2018/07/20
End Sub

Private Sub Form_Initialize()
ReDim tm(1 To TF_TM) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   With frm020107_1
      textTM01 = .Text1
      textTM02 = .Text2
      textTM03 = .Text3
      textTM04 = .Text4
      strReceiveNo = .Tag
   End With
   '顯示收文號
   lblCP09 = strReceiveNo
   ClearAll
   ReadAllData
   Frame2.Enabled = False
End Sub

'清除欄位值
Private Sub ClearAll()
   lblNation.Caption = Empty
   textCP10.Text = Empty
   lblCP10Nm.Caption = Empty
   lblCP06.Caption = Empty
   lblCP07.Caption = Empty
   textTM05.Text = Empty
   lblCP13Nm.Caption = Empty
   lblCP14Nm.Caption = Empty
   textTM21.Text = Empty
   textTM22.Text = Empty
   lblTM09.Caption = Empty
   textTM23.Text = Empty
   lblTM23Nm.Caption = Empty
   textTM78.Text = Empty
   lblTM78Nm.Caption = Empty
   textTM79.Text = Empty
   lblTM79Nm.Caption = Empty
   textTM80.Text = Empty
   lblTM80Nm.Caption = Empty
   textTM81.Text = Empty
   lblTM81Nm.Caption = Empty
   textCP44.Text = Empty
   lblCP44Nm.Caption = Empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020107_2 = Nothing
End Sub

Private Sub ReadAllData()
Dim bolTmp As Boolean
   
   tm(1) = textTM01
   tm(2) = textTM02
   tm(3) = textTM03
   tm(4) = textTM04
   Select Case tm(1)
      Case "T"
         '讀取商標基本檔
         If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
            '案件名稱
            textTM05 = tm(5)
            '商品類別
            lblTM09 = tm(9)
            '申請國家
            If tm(10) <> "" Then
               If ClsPDGetNation(tm(10), strExc(0)) Then
                  lblNation = strExc(0)
               End If
            End If
            '專用期限
            If Val(tm(21)) > 0 Then
               textTM21 = tm(21)
            End If
            If Val(tm(22)) > 0 Then
               textTM22 = tm(22)
            End If
            '申請人1,2,3,4,5
            If tm(23) <> "" Then
               textTM23.Text = ChangeCustomerL(tm(23))
               If ClsPDGetCustomer(tm(23), strExc(0)) Then
                  lblTM23Nm = strExc(0)
               End If
            End If
            If tm(78) <> "" Then
               textTM78.Text = ChangeCustomerL(tm(78))
               If ClsPDGetCustomer(tm(78), strExc(0)) Then
                  lblTM78Nm = strExc(0)
               End If
            End If
            If tm(79) <> "" Then
               textTM79.Text = ChangeCustomerL(tm(79))
               If ClsPDGetCustomer(tm(79), strExc(0)) Then
                  lblTM79Nm = strExc(0)
               End If
            End If
            If tm(80) <> "" Then
               textTM80.Text = ChangeCustomerL(tm(80))
               If ClsPDGetCustomer(tm(80), strExc(0)) Then
                  lblTM80Nm = strExc(0)
               End If
            End If
            If tm(81) <> "" Then
               textTM81.Text = ChangeCustomerL(tm(81))
               If ClsPDGetCustomer(tm(81), strExc(0)) Then
                  lblTM81Nm = strExc(0)
               End If
            End If
            '檢查有無代表圖
            'Modify by Amy 2018/07/20  改寫至function
'            strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & textTM01 & "' and ibf02='" & textTM02 & "' and ibf03='" & textTM03 & "' and ibf04='" & textTM04 & "' and ibf05='1'"
'            CheckOC2
'            adoRecordset1.CursorLocation = adUseClient
'            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            If ChkImgByteFile(textTM01, textTM02, textTM03, textTM04) = True Then
                Command3.Caption = "已設定代表圖(&I)"
                Command3.BackColor = &HC0FFC0
            Else
                Command3.Caption = "未設定代表圖(&I)"
                Command3.BackColor = &HC0C0FF
            End If
'            CheckOC2
            'end 2018/07/20
         End If
         
      Case "TT"
         '讀取服務業務基本檔
         If ClsPDReadServicePracticeDatabase(tm(), intWhere) Then
            '案件名稱
            textTM05 = tm(5) & tm(6) & tm(7)
            '申請國家
            If tm(9) <> "" Then
               If ClsPDGetNation(tm(9), strExc(0)) Then
                  lblNation = strExc(0)
               End If
            End If
            '申請人1,2,3,4,5
            If tm(8) <> "" Then
               textTM23.Text = ChangeCustomerL(tm(8))
               If ClsPDGetCustomer(tm(8), strExc(0)) Then
                  lblTM23Nm = strExc(0)
               End If
            End If
            If tm(58) <> "" Then
               textTM78.Text = ChangeCustomerL(tm(58))
               If ClsPDGetCustomer(tm(58), strExc(0)) Then
                  lblTM78Nm = strExc(0)
               End If
            End If
            If tm(59) <> "" Then
               textTM79.Text = ChangeCustomerL(tm(59))
               If ClsPDGetCustomer(tm(59), strExc(0)) Then
                  lblTM79Nm = strExc(0)
               End If
            End If
            If tm(65) <> "" Then
               textTM80.Text = ChangeCustomerL(tm(65))
               If ClsPDGetCustomer(tm(65), strExc(0)) Then
                  lblTM80Nm = strExc(0)
               End If
            End If
            If tm(66) <> "" Then
               textTM81.Text = ChangeCustomerL(tm(66))
               If ClsPDGetCustomer(tm(66), strExc(0)) Then
                  lblTM81Nm = strExc(0)
               End If
            End If
         End If
   End Select
   cp(9) = strReceiveNo
   '讀取案件進度檔
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      '智權人員
      If cp(13) <> "" Then
         If ClsPDGetStaff(cp(13), strExc(0)) Then lblCP13Nm = cp(13) & " " & strExc(0)
      End If
      '承辦人
      If cp(14) <> "" Then
         If ClsPDGetStaff(cp(14), strExc(0)) Then lblCP14Nm = cp(14) & " " & strExc(0)
      End If
      '本所期限
      If Val(cp(6)) > 0 Then
         lblCP06 = cp(6)
      End If
      '法定期限
      If Val(cp(7)) > 0 Then
         lblCP07 = cp(7)
      End If
      '案件性質
      If cp(10) <> "" Then
         textCP10 = cp(10)
         If tm(10) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         If ClsPDGetCaseProperty(tm(1), textCP10, strExc(0), bolTmp) Then
            lblCP10Nm = strExc(0)
         End If
      End If
      '代理人
      If cp(44) <> "" Then
         textCP44 = ChangeCustomerL(cp(44))
         If PUB_GetAgentName(tm(1), cp(44), strExc(0)) Then
            lblCP44Nm = strExc(0)
         End If
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim strTmp As String  '指示信
'Dim strTmp2 As String '給客戶函
Dim ii As Integer
Dim BolHaveImg As Boolean
Dim strText As String
Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String
Dim bolReturn As Boolean
'Added by Lydia 2018/10/26
Dim iPicNo As Integer, iPicNo2 As Integer '信頭、信尾圖檔代碼
'Dim stFileName As String '暫存圖檔檔名 'Remove by Lydia 2018/11/02
'Dim oShape   'Remove by Lydia 2018/11/02
'end 2018/10/26

   Select Case Index
      Case 0 '確定
         strTmp = ""
         Select Case textCP10
            Case "101"
               If (Check1(0).Value = 1 Or Check1(1).Value = 1) And lstText.ListCount = 0 Then
                  MsgBox "多件申請時，請輸入相關的本所案號！"
                  Text1(1).SetFocus
                  Exit Sub
               End If
               '多件申請，同圖，不同類別
               If Check1(1).Value = 1 Then
                  strTmp = "03"
               '多件申請，同類別，不同圖
               ElseIf Check1(0).Value = 1 Then
                  strTmp = "04"
               Else
                  strTmp = "01"
               End If
            Case "102"
               strTmp = "01"
            Case "301"
               strTmp = "01"
            Case "501"
               strTmp = "01"
            Case "502"
               strTmp = "01"
         End Select
         If strTmp <> "" Then
            'Added by Lydia 2018/10/26 信頭、信尾圖檔代碼
            If txtLetterHead <> "N" Then
               'Added by Morgan 2020/3/30
               If strSrvDate(1) >= 智慧所更名日 Then
                  PUB_GetLetterPicID tm(130), textTM01, iPicNo, iPicNo2, 1, True
               Else
               'end 2020/3/30
   
                 If tm(130) = "J" Then
                     iPicNo = 21
                     iPicNo2 = 55
                 Else
                     iPicNo = 7
                 End If
               End If 'Added by Morgan 2020/3/30
            End If
            'end 2018/10/26
            
            ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
            Call InsExpField(strTmp)
            'Modify By Sindy 2019/12/25 不存定稿資料檔
            'NowPrint strReceiveNo, "18", strTmp, True, strUserNum, 0
            NowPrint strReceiveNo, "18", strTmp, True, strUserNum, 0, , , , , , , , False
            '2019/12/25 END
          '  If strTmp2 <> "" Then
          '     NowPrint strReceiveNo, "18", strTmp2, True, strUserNum, 0
          '  End If

            'Added by Lydia 2018/10/26
            '切換為整頁模式,信頭才會正常顯示
            If g_WordAp.ActiveWindow.View.SplitSpecial = wdPaneNone Then
               g_WordAp.ActiveWindow.ActivePane.View.Type = wdPageView
            Else
               g_WordAp.ActiveWindow.View.Type = wdPageView
            End If
            'end 2018/10/26
            g_WordAp.Selection.PageSetup.TopMargin = g_WordAp.CentimetersToPoints(4) '上邊界4公分

            Select Case textCP10
               Case "101" '申請
                  If Not (g_WordAp Is Nothing) Then
                     With g_WordAp
                        Call WordFindText("***圖表***")  '尋找Word檔中文字
                        'Modify by Amy 2018/07/27 原:GetImgByteFile
                        strFilePathN = "": BolHaveImg = GetImgByteFile_Case(textTM01, textTM02, textTM03, textTM04, strFilePathN, 9)
                        '非多件申請
                        If Check1(0).Value = 0 And Check1(1).Value = 0 Then
                           'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                           'If BolHaveImg = True Then Call AddInPicToWordR(False)  '插入圖片檔案
                           If BolHaveImg = True Then Call AddInPicToWordR2(False)
                           Call WordFindText("***代表圖***")
                           'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                           'If BolHaveImg = True Then Call AddInPicToWordR(False)
                           If BolHaveImg = True Then Call AddInPicToWordR2(False)
                           'Added by Lydia 2018/11/02 改成模組(信頭不只第一頁要有,後面給客戶的都要有)
                           'Modify By Sindy 2021/1/21 mark
'                           Call GetLetterHead("2", iPicNo, iPicNo2)
                           '2021/1/21 END
                        '多件申請
                        Else
                           '插入表格
                           '同類別，不同圖
                           If Check1(0).Value = 1 Then
                              .Selection.Tables.add Range:=.Selection.Range, NumRows:=(lstText.ListCount + 1), NumColumns:=1
                           '同圖，不同類別
                           Else
                              '.Selection.Tables.Add Range:=.Selection.Range, NumRows:=(lstText.ListCount + 1) + 1, NumColumns:=2
                              .Selection.Tables.add Range:=.Selection.Range, NumRows:=(lstText.ListCount + 1), NumColumns:=2
                           End If
                           '設定框線
                           .Selection.Tables(1).Select
                           With .Selection.Borders(wdBorderTop)
                               .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
                               .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
                           End With
                           With .Selection.Borders(wdBorderLeft)
                               .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
                               .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
                           End With
                           With .Selection.Borders(wdBorderBottom)
                               .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
                               .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
                           End With
                           With .Selection.Borders(wdBorderRight)
                               .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
                               .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
                           End With
                           With .Selection.Borders(wdBorderHorizontal)
                               .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
                           End With
                           With .Selection.Borders(wdBorderVertical)
                               .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
                           End With
                           '同類別，不同圖
                           If Check1(0).Value = 1 Then
                              '設定表格高度
                              .Selection.Cells.SetHeight RowHeight:=200, HeightRule:=wdRowHeightExactly
                              .Selection.MoveRight Unit:=wdCell
                              strText = "本所案號：" & textTM01 & "-" & textTM02 & IIf(textTM03 & textTM04 = "000", "", "-" & textTM03 & "-" & textTM04) & vbCrLf
                              strText = strText & "圖樣：***" & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 & "***" & vbCrLf
                              'Modify By Sindy 2015/3/10 +IIf(Left(cp(44), 6) = "Y52269", "　　（英文涵義：　　　　　）", "")
                              strText = strText & "商標說明：" & Trim(textTM05) & IIf(Left(cp(44), 6) = "Y52269", "　　（英文涵義：　　　　　）", "")
                              .Selection.TypeText Text:=strText
                              For ii = 0 To lstText.ListCount - 1
                                 strTM01 = SystemNumber(lstText.List(ii), 1)
                                 strTM02 = SystemNumber(lstText.List(ii), 2)
                                 strTM03 = SystemNumber(lstText.List(ii), 3)
                                 strTM04 = SystemNumber(lstText.List(ii), 4)
                                 bolReturn = PUB_ReadTradeMarkData(tm(), strTM01, strTM02, strTM03, strTM04)
                                 .Selection.MoveRight Unit:=wdCell
                                 strText = "本所案號：" & strTM01 & "-" & strTM02 & IIf(strTM03 & strTM04 = "000", "", "-" & strTM03 & "-" & strTM04) & vbCrLf
                                 strText = strText & "圖樣：***" & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "***" & vbCrLf
                                 'Modify By Sindy 2015/3/10 +IIf(Left(cp(44), 6) = "Y52269", "　　（英文涵義：　　　　　）", "")
                                 strText = strText & "商標說明：" & tm(5) & IIf(Left(cp(44), 6) = "Y52269", "　　（英文涵義：　　　　　）", "")
                                 .Selection.TypeText Text:=strText
                              Next ii
                              Call WordFindText("***" & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 & "***")
                              'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                              'If BolHaveImg = True Then Call AddInPicToWordR(True)
                              If BolHaveImg = True Then Call AddInPicToWordR2(False)
                              Call WordFindText("***代表圖***")
                              'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                              'If BolHaveImg = True Then Call AddInPicToWordR(False)
                              If BolHaveImg = True Then Call AddInPicToWordR2(False)
                              'Added by Lydia 2018/11/02 改成模組(信頭不只第一頁要有,後面給客戶的都要有)
                              'Modify By Sindy 2021/1/21 mark
'                              Call GetLetterHead("2", iPicNo, iPicNo2)
                              For ii = 0 To lstText.ListCount - 1
                                 strTM01 = SystemNumber(lstText.List(ii), 1)
                                 strTM02 = SystemNumber(lstText.List(ii), 2)
                                 strTM03 = SystemNumber(lstText.List(ii), 3)
                                 strTM04 = SystemNumber(lstText.List(ii), 4)
                                 Call WordFindText("***" & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "***")
                                 'Modify by Amy 2018/07/27 原:GetImgByteFile
                                 strFilePathN = "": BolHaveImg = GetImgByteFile_Case(strTM01, strTM02, strTM03, strTM04, strFilePathN, 9)
                                 'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                                 'If BolHaveImg = True Then Call AddInPicToWordR(True)
                                 If BolHaveImg = True Then Call AddInPicToWordR2(False)
                                 .Selection.EndKey Unit:=wdStory '游標移到文章最後
                                 .Selection.InsertBreak Type:=wdPageBreak '新增下一頁
                                 .Selection.TypeParagraph
                                 .Selection.TypeParagraph 'Added by Lydia 2018/11/02 多空一行(配合信頭)
                                 .Selection.TypeText Text:=GetCaseCustSheet(strTM01, strTM02, strTM03, strTM04) '案件回覆單
                                 Call WordFindText("***代表圖***")
                                 'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                                 'If BolHaveImg = True Then Call AddInPicToWordR(True)
                                 If BolHaveImg = True Then Call AddInPicToWordR2(False)
                                 'Added by Lydia 2018/11/02 改成模組(信頭不只第一頁要有,後面給客戶的都要有)
                                 'Modify By Sindy 2021/1/21 mark
'                                 Call GetLetterHead("2", iPicNo, iPicNo2)
                              Next ii
                           '同圖，不同類別
                           Else
'                              '設定表格高度
'                              .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                              .Selection.SelectRow
'                              .Selection.Cells.SetHeight RowHeight:=200, HeightRule:=wdRowHeightExactly
                              '設定表格欄寬
                              .Selection.MoveLeft Unit:=wdCharacter, Count:=1
                              .Selection.SelectColumn
                              .Selection.Cells.SetWidth ColumnWidth:=.CentimetersToPoints(0.5), RulerStyle:=wdAdjustProportional
                              .Selection.SelectRow
            '                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            '                  .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                              .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                              .Selection.MoveRight Unit:=wdCell
'                              .Selection.TypeText Text:="商標名稱：" & Trim(textTM05) & vbCrLf & "商標圖樣：" & vbCrLf & "***" & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 & "***"
'                              .Selection.MoveRight Unit:=wdCell
                              .Selection.TypeText Text:="1"
                              .Selection.MoveRight Unit:=wdCell
                              strText = "Our Ref：" & textTM01 & "-" & textTM02 & IIf(textTM03 & textTM04 = "000", "", "-" & textTM03 & "-" & textTM04)
                              If lblCP14Nm = "" Then
                                 strText = strText & vbCrLf & "商品類別：" & vbCrLf & "指定商品："
                              Else
                                 strText = strText & BeforePrintGetDBData("TMGoods:" & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 & "-中文含標題")
                              End If
                              .Selection.TypeText Text:=strText
                              For ii = 0 To lstText.ListCount - 1
                                 strTM01 = SystemNumber(lstText.List(ii), 1)
                                 strTM02 = SystemNumber(lstText.List(ii), 2)
                                 strTM03 = SystemNumber(lstText.List(ii), 3)
                                 strTM04 = SystemNumber(lstText.List(ii), 4)
                                 bolReturn = PUB_ReadTradeMarkData(tm(), strTM01, strTM02, strTM03, strTM04)
                                 .Selection.MoveRight Unit:=wdCell
                                 .Selection.TypeText Text:=ii + 2
                                 .Selection.MoveRight Unit:=wdCell
                                 strText = "Our Ref：" & strTM01 & "-" & strTM02 & IIf(strTM03 & strTM04 = "000", "", "-" & strTM03 & "-" & strTM04)
                                 If tm(9) = "" Then
                                    strText = strText & vbCrLf & "商品類別：" & vbCrLf & "指定商品："
                                 Else
                                    strText = strText & BeforePrintGetDBData("TMGoods:" & lstText.List(ii) & "-中文含標題")
                                 End If
                                 .Selection.TypeText Text:=strText
                              Next ii
'                              Call WordFindText("***" & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 & "***")
'                              If BolHaveImg = True Then Call AddInPicToWordR(True)
                              Call WordFindText("***代表圖***")
                              'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                              'If BolHaveImg = True Then Call AddInPicToWordR(False)
                              If BolHaveImg = True Then Call AddInPicToWordR2(False)
                              Call WordFindText("***代表圖***")
                              'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                              'If BolHaveImg = True Then Call AddInPicToWordR(False)
                              If BolHaveImg = True Then Call AddInPicToWordR2(False)
                              'Added by Lydia 2018/11/02 改成模組(信頭不只第一頁要有,後面給客戶的都要有)
                              'Modify By Sindy 2021/1/21 mark
'                              Call GetLetterHead("2", iPicNo, iPicNo2)
                              '其他案件的案件回覆單
                              For ii = 0 To lstText.ListCount - 1
                                 strTM01 = SystemNumber(lstText.List(ii), 1)
                                 strTM02 = SystemNumber(lstText.List(ii), 2)
                                 strTM03 = SystemNumber(lstText.List(ii), 3)
                                 strTM04 = SystemNumber(lstText.List(ii), 4)
                                 'Modify by Amy 2018/07/27 原:GetImgByteFile
                                 strFilePathN = "": BolHaveImg = GetImgByteFile_Case(strTM01, strTM02, strTM03, strTM04, strFilePathN, 9)
                                 .Selection.EndKey Unit:=wdStory '游標移到文章最後
                                 .Selection.InsertBreak Type:=wdPageBreak '新增下一頁
                                 .Selection.TypeParagraph '空一行
                                 .Selection.TypeParagraph 'Added by Lydia 2018/11/02 多空一行(配合信頭)
                                 .Selection.TypeText Text:=GetCaseCustSheet(strTM01, strTM02, strTM03, strTM04) '案件回覆單
                                 Call WordFindText("***代表圖***")
                                 'Modified by Lydia 2018/10/26 插入圖片(改用物件)
                                 'If BolHaveImg = True Then Call AddInPicToWordR(True)
                                 If BolHaveImg = True Then Call AddInPicToWordR2(False)
                                 'Added by Lydia 2018/11/02 改成模組(信頭不只第一頁要有,後面給客戶的都要有)
                                 'Modify By Sindy 2021/1/21 mark
'                                 Call GetLetterHead("2", iPicNo, iPicNo2)
                              Next ii
                           End If
                           '多件申請,因此要Rename 第一頁的Our Ref:
                           If textTM03 <> "0" Or textTM04 <> "00" Then
                              strText = "Our Ref: " & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04
                           Else
                              strText = "Our Ref: " & textTM01 & "-" & textTM02
                           End If
                           Call WordFindText(strText)
                           '串其他案號
                           For ii = 0 To lstText.ListCount - 1
                              strTM01 = SystemNumber(lstText.List(ii), 1)
                              strTM02 = SystemNumber(lstText.List(ii), 2)
                              strTM03 = SystemNumber(lstText.List(ii), 3)
                              strTM04 = SystemNumber(lstText.List(ii), 4)
                              strText = strText & "," & strTM01 & "-" & strTM02 & IIf(strTM03 & strTM04 = "000", "", "-" & strTM03 & "-" & strTM04)
                           Next ii
                           .Selection.TypeText strText
                        End If
                     End With
                  End If
            End Select

            'Added by Lydia 2018/10/26 第 2 頁以後不要有信頭,故放在本文;
            'Modified by Lydia 2018/11/02 改成模組(信頭不只第一頁要有,後面給客戶的都要有)
            'Modify By Sindy 2021/1/21 mark
'            Call GetLetterHead("1", iPicNo, iPicNo2)
            
            '簽章
            Call WordFindText("***簽章***")  '尋找Word檔中文字
            If txtLetterHead <> "N" Then
                'Modified by Lydia 2018/11/02 要求與MCTF的定稿簽章一致
                'If PUB_ReadDB2File(strFilePathN, 56) = True Then '林藍
                If PUB_ReadDB2File(strFilePathN, 47) = True Then '林黑
                    Call AddInPicToWordR2(True)
                End If
            End If
            'Added by Lydia 2018/11/02 申請以外,移到下一頁(案件回覆單)
            If textCP10 <> "101" Then
                g_WordAp.Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                'Modify By Sindy 2021/1/21 mark
'                Call GetLetterHead("2", iPicNo, iPicNo2)
            End If
            'end 2018/11/02

            Set g_WordAp = Nothing
            'end 2018/10/26
             
            'Add By Sindy 2018/8/13
            'Modified by Lydia 2018/11/14 區分是否前畫面或工作進度維護而來
            'If frm020107_1.Tag <> "" Then
            'Modify By Sindy 2021/1/21
            'If frm020107_1.Tag <> "" And Left(UCase(frm020107_1.Tag), 1) = "F" Then
            If frm020107_1.cmdOK(2).Tag <> "" And Left(UCase(frm020107_1.cmdOK(2).Tag), 1) = "F" Then
            '2021/1/21 END
               Unload frm020107_1
               Unload Me
            Else
            '2018/8/13 END
               frm020107_1.Show
               '回第一個畫面清除
               frm020107_1.Clear
               Unload Me
            End If
         Else
            MsgBox "無指示信定稿！"
         End If
      Case 1 '回前畫面
         'Add By Sindy 2018/8/13
         'Modified by Lydia 2018/11/14 區分是否前畫面或工作進度維護而來
         'If frm020107_1.Tag <> "" Then
         'Modify By Sindy 2021/1/21
         'If frm020107_1.Tag <> "" And Left(UCase(frm020107_1.Tag), 1) = "F" Then
         If frm020107_1.cmdOK(2).Tag <> "" And Left(UCase(frm020107_1.cmdOK(2).Tag), 1) = "F" Then
         '2021/1/21 END
            Unload frm020107_1
            Unload Me
         Else
         '2018/8/13 END
            frm020107_1.Show
            Unload Me
         End If
   End Select
End Sub

'尋找Word檔中文字
Private Sub WordFindText(strFindText As String, Optional strReplaceText As String = "")
   If Trim(strFindText) = "" Then Exit Sub
   With g_WordAp
'      .Selection.WholeStory
'      .Selection.Copy
      .Selection.GoTo what:=wdGoToPage, which:=wdGoToPrevious, Count:=3
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strFindText
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      .Selection.Find.Execute
      .Selection.Delete
      .Selection.Font.ColorIndex = wdBlack
      .Selection.TypeText strReplaceText
   End With
End Sub

'Word檔插入圖片
Private Sub AddInPicToWordR(bolImgCoverText As Boolean)
Dim dblHeight As Double
Dim dblWidth As Double

On Error GoTo ErrEnd
   
   With g_WordAp
      '插入1列1欄的表格
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=1
      '設定欄框為無線
      .Selection.Tables(1).AutoFormat Format:=wdTableFormatSimple1, ApplyBorders:=False, ApplyShading:=True, ApplyFont:=True, ApplyColor:=True, _
            ApplyHeadingRows:=True, ApplyLastRow:=False, ApplyFirstColumn:=True, ApplyLastColumn:=False, AutoFit:=False
NotAddTable:
      '指定檔名
      .Selection.InlineShapes.AddPicture FileName:=strFilePathN, LinkToFile:=False, SaveWithDocument:=True
      .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
      '定義大小
      '鎖定最高 圖區
      '圖大小
      .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).LockAspectRatio = msoTrue
      dblHeight = .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height
      dblWidth = .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width
      If dblHeight > 150 Then
         .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height = 150
         .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width = Int(dblWidth / Round((dblHeight / 150), 5))
      Else
         If dblWidth > 150 Then
            .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width = 150
            .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height = Int(dblHeight / Round((dblWidth / 150), 5))
         End If
      End If
      
'      '指定檔名
'      .ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=strFilePathN, LinkToFile:=False, SaveWithDocument:=True
'      .ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
'      '定義大小
'      '鎖定最高 圖區
'      '圖大小
'      .Selection.ShapeRange.LockAspectRatio = msoTrue
'      .Selection.ShapeRange.Height = 230
'      If .Selection.ShapeRange.Width > 150 Then
'         .Selection.ShapeRange.Width = 150
'      End If
'      '移到指定位置
'      '.Selection.ShapeRange.Left = .CentimetersToPoints(12) '11.2
'      '.Selection.ShapeRange.Top = .CentimetersToPoints(1)
'      .Selection.ShapeRange.LockAnchor = False
'      If bolImgCoverText = True Then
'         .Selection.ShapeRange.WrapFormat.Type = wdWrapNone '圖蓋文
'      Else
'         '.Selection.ShapeRange.WrapFormat.Type = wdWrapSquare '文字繞圖
'         .Selection.ShapeRange.WrapFormat.Type = wdWrapTopBottom '上下
'      End If
'      '游標移到文章最後
'      '.Selection.EndKey Unit:=wdStory
   End With
   Exit Sub
   
ErrEnd:
   If Err.Number = 5962 Then
      GoTo NotAddTable
   End If
End Sub

'客戶函
Private Function GetCaseCustSheet(strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String)
Dim bolReturn As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strAppl1CU15 As String
   
   GetCaseCustSheet = ""
   bolReturn = PUB_ReadTradeMarkData(tm(), strTM01, strTM02, strTM03, strTM04)
   If bolReturn = True Then
      strSql = "select DECODE(CU15,'0','台端','1','貴公司','貴單位') From customer" & _
               " where cu01='" & Left(tm(23), 8) & "' and cu02='" & Right(tm(23), 1) & "'"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strAppl1CU15 = rsTmp.Fields(0)
      End If
      GetCaseCustSheet = strAppl1CU15 & "委託本所代為申請之大陸商標，資料如下：" & vbCrLf & vbCrLf
      GetCaseCustSheet = GetCaseCustSheet & "一、本件案號：" & strTM01 & "-" & strTM02 & IIf(strTM03 & strTM04 = "000", "", "-" & strTM03 & "-" & strTM04) & vbCrLf & vbCrLf
      GetCaseCustSheet = GetCaseCustSheet & "二、" & GetApplData(tm(23), tm(78), tm(79), tm(80), tm(81)) & vbCrLf & vbCrLf
      'GetCaseCustSheet = GetCaseCustSheet & "三、商標說明：" & tm(5) & vbCrLf & vbCrLf
      GetCaseCustSheet = GetCaseCustSheet & "三、商品類別及名稱：" & BeforePrintGetDBData("TMGoods:" & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "-中文含第類") & vbCrLf & vbCrLf
      GetCaseCustSheet = GetCaseCustSheet & "四、商標圖樣：" & vbCrLf
      GetCaseCustSheet = GetCaseCustSheet & "***代表圖***" & vbCrLf
      'Add By Sindy 2021/1/15
      GetCaseCustSheet = GetCaseCustSheet & "　　商標說明：" & tm(5) & IIf(Left(cp(44), 6) = "Y52269", "　　（英文涵義：　　　　　）", "")
      '2021/1/15 END
   End If
   
   Set rsTmp = Nothing
End Function

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField(strET03 As String)
Dim strText As String
Dim ii As Integer
Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String
Dim tmpArr 'Add By Sindy 2015/5/27
   
   '清除定稿例外欄位檔原有資料
   EndLetter "18", lblCP09.Caption, strET03, strUserNum
   Select Case textCP10
      Case "101"
'         '清除定稿例外欄位檔原有資料
'         EndLetter "18", lblCP09.Caption, strET03, strUserNum
         '申請人資料
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','申請人資料','" & ChgSQL(GetApplData(textTM23, textTM78, textTM79, textTM80, textTM81)) & "')"
         cnnConnection.Execute strSql
         
         '英文涵義
         If Left(cp(44), 6) = "Y52269" Then
            strText = "（英文涵義：　　　　　）"
         Else
            strText = ""
         End If
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','英文涵義','" & strText & "')"
         cnnConnection.Execute strSql
         '同圖
         If strET03 = "03" Then
'            strText = Replace(BeforePrintGetDBData("TMGoods:" & textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04 & "-中文含第類"), vbCrLf, "") & vbCrLf
'            For ii = 0 To lstText.ListCount - 1
'               strTM01 = SystemNumber(lstText.List(ii), 1)
'               strTM02 = SystemNumber(lstText.List(ii), 2)
'               strTM03 = SystemNumber(lstText.List(ii), 3)
'               strTM04 = SystemNumber(lstText.List(ii), 4)
'               strText = strText & Replace(BeforePrintGetDBData("TMGoods:" & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "-中文含第類"), vbCrLf, "") & vbCrLf
'            Next ii
'            '商品類別及名稱
'            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
'                     "','商品類別及名稱','" & strText & "')"
'            cnnConnection.Execute strSql
         '同類別
         ElseIf strET03 = "04" Then
'            strText = "「" & GetPrjName(textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04) & "」"
'            For ii = 0 To lstText.ListCount - 1
'               strTM01 = SystemNumber(lstText.List(ii), 1)
'               strTM02 = SystemNumber(lstText.List(ii), 2)
'               strTM03 = SystemNumber(lstText.List(ii), 3)
'               strTM04 = SystemNumber(lstText.List(ii), 4)
'               strText = strText & "、「" & GetPrjName(strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04) & "」"
'            Next ii
'            '商標名稱
'            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
'                     "','商標名稱','" & ChgSQL(strText) & "')"
'            cnnConnection.Execute strSql
            '委託書幾份
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                     "','委託書幾份','" & lstText.ListCount + 1 & "')"
            cnnConnection.Execute strSql
         'Add By Sindy 2015/5/27
         Else
            tmpArr = Split(lblTM09, ",")
            '是否有跨類
            If (Val(UBound(tmpArr)) + 1) > 1 Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                        "','一案多類','一案多類')"
               cnnConnection.Execute strSql
            Else
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                        "','一案多類','')"
               cnnConnection.Execute strSql
            End If
         '2015/5/27 END
         End If
      Case Else
         'Add By Sindy 2015/3/18
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','公司營業編號1','" & GetApplSingleData(textTM23) & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','公司營業編號2','" & GetApplSingleData(textTM78) & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','公司營業編號3','" & GetApplSingleData(textTM79) & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','公司營業編號4','" & GetApplSingleData(textTM80) & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
                  "','公司營業編號5','" & GetApplSingleData(textTM81) & "')"
         cnnConnection.Execute strSql
         '2015/3/18 END
   End Select
   'Add By Sindy 2014/12/12
   '開台一智權
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & "18" & "','" & lblCP09.Caption & "','" & strET03 & "','" & strUserNum & _
            "','開台一智權','" & IIf(tm(130) = "J", "※ 帳單抬頭請開「台一智權股份有限公司」", "") & "')"
   cnnConnection.Execute strSql
   '2014/12/12 END
End Sub

'取得申請人資料
Private Function GetApplData(strAppl1 As String, strAppl2 As String, strAppl3 As String, strAppl4 As String, strAppl5 As String) As String
Dim rsTmp As New ADODB.Recordset
Dim i As Integer
Dim strCustID As String
Dim strTempAddr As String 'Add By Sindy 2022/4/14
   
   GetApplData = ""
   For i = 1 To 5
      If i = 1 Then strCustID = Trim(strAppl1)
      If i = 2 Then strCustID = Trim(strAppl2)
      If i = 3 Then strCustID = Trim(strAppl3)
      If i = 4 Then strCustID = Trim(strAppl4)
      If i = 5 Then strCustID = Trim(strAppl5)
      If strCustID = "" Then Exit For
      If Len(strCustID) < 9 Then: strCustID = strCustID & String(9 - Len(strCustID), "0")
      strSql = "select CU104,CU04,CU05,CU88,CU89,CU90,CU11,decode(substr(CU10,1,2),'00','000',CU10) CU10,na03,CU23,CU24,CU25,CU26,CU27,CU28,CU102,CU15" & _
               " From customer,nation" & _
               " where cu01='" & Left(strCustID, 8) & "' and cu02='" & Right(strCustID, 1) & "'" & _
               " and decode(substr(CU10,1,2),'00','000',CU10)=na01(+)"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If GetApplData <> "" Then GetApplData = GetApplData & vbCrLf & "　　"
         GetApplData = GetApplData & "申請人名稱" & IIf(i <> 1, i, "") & "：" & IIf(rsTmp.Fields("CU104") <> "", rsTmp.Fields("CU104"), IIf(rsTmp.Fields("CU04") <> "", rsTmp.Fields("CU04"), rsTmp.Fields("CU05") & " " & rsTmp.Fields("CU88") & " " & rsTmp.Fields("CU89") & " " & rsTmp.Fields("CU90")))
         If rsTmp.Fields("cu10") = "000" Then
            GetApplData = GetApplData & vbCrLf
            'Modify By Sindy 2014/12/31
            If "" & rsTmp.Fields("cu15") = "0" Then '0.個人
               GetApplData = GetApplData & "　　身分證字號：" & rsTmp.Fields("cu11") & vbCrLf
            Else
            '2014/12/31 END
               GetApplData = GetApplData & "　　公司營業編號：" & rsTmp.Fields("cu11") & vbCrLf
            End If
         Else
            '非台灣者,申請人名稱後面+英文名稱
            GetApplData = GetApplData & IIf(rsTmp.Fields("CU104") & rsTmp.Fields("CU04") <> "", rsTmp.Fields("CU05") & " " & rsTmp.Fields("CU88") & " " & rsTmp.Fields("CU89") & " " & rsTmp.Fields("CU90"), "") & vbCrLf
            GetApplData = GetApplData & "　　國籍：" & rsTmp.Fields("na03") & vbCrLf
         End If
         'Modify by Sindy 2022/4/14 除去地址前面的郵遞區號
         strTempAddr = Trim(PUB_ChgNumeralStyle("" & rsTmp.Fields("cu23")))
         If strTempAddr <> "" Then
            Do While (Left(strTempAddr, 1) >= "１" And Left(strTempAddr, 1) <= "９") Or (Left(strTempAddr, 1) >= "1" And Left(strTempAddr, 1) <= "9")
               strTempAddr = Mid(strTempAddr, 2)
            Loop
         End If
         '2022/4/14 END
         GetApplData = GetApplData & "　　地址：" & rsTmp.Fields("na03") & "　" & strTempAddr 'PUB_ChgNumeralStyle("" & rsTmp.Fields("cu23"))
         '非台灣者,地址後面+英文地址
         If rsTmp.Fields("cu10") <> "000" Then
            GetApplData = GetApplData & IIf(rsTmp.Fields("cu24") <> "", rsTmp.Fields("cu24") & " " & rsTmp.Fields("cu25") & " " & rsTmp.Fields("cu26") & " " & rsTmp.Fields("cu27") & " " & rsTmp.Fields("cu28") & " " & rsTmp.Fields("cu102"), "")
         End If
      End If
   Next i
   
   Set rsTmp = Nothing
End Function

'取得申請人ID資料
Private Function GetApplSingleData(strAppl As String) As String
Dim rsTmp As New ADODB.Recordset
   
   GetApplSingleData = ""
   If Len(strAppl) < 9 Then: strAppl = strAppl & String(9 - Len(strAppl), "0")
   strSql = "select CU104,CU04,CU05,CU88,CU89,CU90,CU11,decode(substr(CU10,1,2),'00','000',CU10) CU10,na03,CU23,CU24,CU25,CU26,CU27,CU28,CU102,CU15" & _
            " From customer,nation" & _
            " where cu01='" & Left(strAppl, 8) & "' and cu02='" & Right(strAppl, 1) & "'" & _
            " and decode(substr(CU10,1,2),'00','000',CU10)=na01(+)"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If rsTmp.Fields("cu10") = "000" Then
         If "" & rsTmp.Fields("cu15") = "0" Then '0.個人
            GetApplSingleData = "身分證字號：" & rsTmp.Fields("cu11")
         Else
            GetApplSingleData = "公司營業編號：" & rsTmp.Fields("cu11")
         End If
      End If
   End If
   
   Set rsTmp = Nothing
End Function

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2018/10/26
Private Sub txtLetterHead_Change()
   TextInverse txtLetterHead
End Sub

Private Sub txtLetterHead_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
       KeyAscii = 0
       Beep
    End If
End Sub

Private Sub AddInPicToWordR2(ByVal bRight As Boolean)
Dim dblHeight As Double
Dim dblWidth As Double
Dim oShape

On Error GoTo ErrEnd

   With g_WordAp

      '插入1列1欄的表格
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=1
      '設定欄框為無線
      .Selection.Tables(1).AutoFormat Format:=wdTableFormatSimple1, ApplyBorders:=False, ApplyShading:=True, ApplyFont:=True, ApplyColor:=True, _
            ApplyHeadingRows:=True, ApplyLastRow:=False, ApplyFirstColumn:=True, ApplyLastColumn:=False, AutoFit:=False
NotAddTable:
      '指定檔名
'Modified by Lydia 2018/11/14 改成舊寫法(因為竹平的代表圖無法在正確位置
      'Modified by Lydia 2018/11/14
      'Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=strFilePathN, LinkToFile:=False, SaveWithDocument:=True)
      Set oShape = .Selection.InlineShapes.AddPicture(FileName:=strFilePathN, LinkToFile:=False, SaveWithDocument:=True).ConvertToShape
      oShape.ZOrder 4
      oShape.LockAnchor = True
      oShape.LockAspectRatio = msoTrue
      oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
      oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
      dblHeight = oShape.Height
      dblWidth = oShape.Width
      If bRight = True Then '簽章(目前圖片左邊有空白)
          If dblWidth > 320 Then
                 oShape.Width = 320
                 oShape.Height = Int(dblHeight / Round((dblWidth / 320), 5))
          End If
          oShape.Left = .CentimetersToPoints(3.5)
      Else '代表圖靠左
            If dblHeight > 150 Then
                   oShape.Height = 150
                   oShape.Width = Int(dblWidth / Round((dblHeight / 150), 5))
            ElseIf dblWidth > 150 Then
                    oShape.Width = 150
                     oShape.Height = Int(dblHeight / Round((dblWidth / 150), 5))
            End If
          oShape.Left = .CentimetersToPoints(0)
      End If
      oShape.WrapFormat.Type = wdWrapTopBottom 'Added by Lydia 2018/11/14 文字上下
'--------------------------------
'      .Selection.InlineShapes.AddPicture FileName:=strFilePathN, LinkToFile:=False, SaveWithDocument:=True
'      .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'      '定義大小
'      '鎖定最高 圖區
'      '圖大小
'      .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).LockAspectRatio = msoTrue
'      dblHeight = .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height
'      dblWidth = .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width
'      If bRight = True Then '簽章(目前圖片左邊有空白)
'          If dblWidth > 320 Then
'                 .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width = 320
'                 .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height = Int(dblHeight / Round((dblWidth / 320), 5))
'          End If
'          .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Left = .CentimetersToPoints(3.5)
'      Else
'            If dblHeight > 150 Then
'               .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height = 150
'               .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width = Int(dblWidth / Round((dblHeight / 150), 5))
'            Else
'               If dblWidth > 150 Then
'                  .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Width = 150
'                  .Selection.InlineShapes(Trim(.Selection.InlineShapes.Count)).Height = Int(dblHeight / Round((dblWidth / 150), 5))
'               End If
'            End If
'      End If
      'end 2018/11/14
   End With
   Exit Sub
   
ErrEnd:
   If Err.Number = 5962 Then
      GoTo NotAddTable
   End If
End Sub
'end 2018/10/26

'Added by Lydia 2018/11/02 設定頁面的信頭和信尾
Private Sub GetLetterHead(ByVal iType As String, ByVal iPicNo As Integer, iPicNo2 As Integer)
Dim stFileName As String
Dim oShape

    If iPicNo > 0 And Not (g_WordAp Is Nothing) Then
        If iType = "1" Then '第一頁的信頭和信尾
            g_WordAp.Selection.HomeKey Unit:=wdStory
        Else   '目前頁面
            g_WordAp.Selection.GoTo what:=wdGoToPage, which:=wdGoToPrevious, Count:=1
            g_WordAp.Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
        End If
        
        '信頭
        'Added by Lydia 2018/11/08 比照撰寫信函WordChinese1
        If iPicNo2 > 0 Then '有信尾
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               Set oShape = g_WordAp.ActiveDocument.Shapes.AddPicture(Anchor:=g_WordAp.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = msoTrue
               oShape.Width = g_WordAp.CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = g_WordAp.CentimetersToPoints(0)
               'Modified by Lydia 2018/11/08
               'oShape.Top = g_WordAp.CentimetersToPoints(0)
               oShape.Top = g_WordAp.CentimetersToPoints(0.5)
               If iPicNo2 > 0 Then
                    '信尾
                    If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                       Set oShape = g_WordAp.ActiveDocument.Shapes.AddPicture(Anchor:=g_WordAp.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                       oShape.ZOrder 4
                       oShape.LockAnchor = True
                       oShape.LockAspectRatio = msoTrue
                       oShape.Width = g_WordAp.CentimetersToPoints(21)
                       oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                       oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                       oShape.Left = g_WordAp.CentimetersToPoints(0)
                       'Modified by Lydia 2018/11/08
                       'oShape.Top = g_WordAp.CentimetersToPoints(27.6)
                       oShape.Top = g_WordAp.CentimetersToPoints(27)
                    End If
               End If '信尾
               g_WordAp.Selection.EndKey Unit:=wdStory
            End If '信頭
        'Added by Lydia 2018/11/08
        Else  '無信尾
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               Set oShape = g_WordAp.ActiveDocument.Shapes.AddPicture(Anchor:=g_WordAp.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = msoTrue
               oShape.Width = 546.5
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = g_WordAp.CentimetersToPoints(1)
               oShape.Top = g_WordAp.CentimetersToPoints(0.8)
               oShape.WrapFormat.Type = wdWrapNone
               g_WordAp.Selection.EndKey Unit:=wdStory
            End If
        End If
        'end 2018/11/08
    End If
End Sub
