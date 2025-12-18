VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090130 
   BorderStyle     =   1  '單線固定
   Caption         =   "委查結果"
   ClientHeight    =   5472
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7872
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5472
   ScaleWidth      =   7872
   Begin VB.CheckBox ChkPass 
      Caption         =   "無須查名/自行查名"
      ForeColor       =   &H00FF00FF&
      Height          =   465
      Left            =   6540
      TabIndex        =   34
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   4920
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   10
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   20
         Top             =   122
         Width           =   450
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   9
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   18
         Top             =   122
         Width           =   320
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   8
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   17
         Top             =   122
         Width           =   700
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   7
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   16
         Top             =   122
         Width           =   480
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "複製前案"
         Height          =   315
         Left            =   0
         TabIndex        =   22
         Top             =   100
         Width           =   1000
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   3120
         Y1              =   240
         Y2              =   240
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   1080
      TabIndex        =   25
      Top             =   60
      Width           =   5295
      _ExtentX        =   9335
      _ExtentY        =   3620
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "查名代號"
      TabPicture(0)   =   "frm090130.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(4)"
      Tab(0).Control(1)=   "cmdRefresh(0)"
      Tab(0).Control(2)=   "txtNo(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtNo(1)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "新增委查單"
      TabPicture(1)   =   "frm090130.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Line1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtUnicode"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Line2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblSname"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtField(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtField(5)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtField(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtField(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtField(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtField(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtField(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Check1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdRefresh(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "查詢"
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   420
         Width           =   960
      End
      Begin VB.CheckBox Check1 
         Caption         =   "含已收文"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtNo 
         Height          =   280
         Index           =   1
         Left            =   -73320
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtNo 
         Height          =   280
         Index           =   0
         Left            =   -73800
         MaxLength       =   3
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "查詢"
         Height          =   315
         Index           =   0
         Left            =   -72300
         TabIndex        =   2
         Top             =   480
         Width           =   960
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   420
         Width           =   650
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   1
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   10
         Top             =   760
         Width           =   1200
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   2
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   11
         Top             =   760
         Width           =   1200
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   3
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   2640
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   4
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   5
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   14
         Top             =   1700
         Width           =   1200
      End
      Begin VB.TextBox txtField 
         Height          =   270
         Index           =   6
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1700
         Width           =   1200
      End
      Begin MSForms.Label lblSname 
         Height          =   255
         Left            =   1890
         TabIndex        =   35
         Top             =   450
         Width           =   855
         Caption         =   "lblSname"
         Size            =   "1508;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line2 
         X1              =   2460
         X2              =   2600
         Y1              =   1833
         Y2              =   1833
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委查單號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   1750
         Width           =   900
      End
      Begin MSForms.TextBox txtUnicode 
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Top             =   1365
         Width           =   2640
         VariousPropertyBits=   671105051
         Size            =   "4657;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   2460
         X2              =   2600
         Y1              =   893
         Y2              =   893
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "查名代號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   -74760
         TabIndex        =   30
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委查人員："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   470
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委查期間："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   1130
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "文字商標："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1425
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   525
      Left            =   3720
      TabIndex        =   19
      Top             =   4920
      Width           =   4140
      Begin VB.CommandButton cmdExit 
         Caption         =   "結束"
         Height          =   315
         Left            =   2805
         TabIndex        =   6
         Top             =   90
         Width           =   1000
      End
      Begin VB.CommandButton CmdRead 
         Caption         =   "查覆明細"
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   90
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "確定"
         Height          =   315
         Index           =   1
         Left            =   600
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   90
         Width           =   1200
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   140
         Width           =   315
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGRD1 
      Height          =   2445
      Left            =   60
      TabIndex        =   3
      Top             =   2400
      Width           =   7755
      _ExtentX        =   13674
      _ExtentY        =   4318
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|總收文號|申請編號|委查單號|客戶名稱|檔案名稱"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label3 
      Caption         =   "註:勾選 V 表示收文，空白表示不收文"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "雙擊開啟該筆查覆明細"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   2160
      Width           =   2895
   End
End
Attribute VB_Name = "frm090130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/01 改成Form2.0 ; MGRD1改字型=新細明體-ExtB、lblSname、txtUnicode改字型=新細明體-ExtB
'Created by Lydia 2015/08/28
Option Explicit
Public mbolCall As Integer '櫃台收文存檔呼叫
Public m_CP09 As String '總收文號
Public iStiu As Integer '0:新增收文, 1:修改,  2:查詢
Public cmdState As Integer '紀錄作用按鍵
Public m_CP13 As String '智權人員

Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_CP27 As String '是否已發文
Dim m_PrevForm As Form '前一畫面
Dim reSearch As Integer '記錄重整的按鍵
Dim iPrevRow1 As Integer '前次點選列
Dim lTotRows1 As Long, lSelRows1 As Long
Dim m_blnColOrderAsc1 As Boolean '欄位資料由小到大排序

Dim m_AttachPath As String

Dim colMTQC02 As Integer '是否已收文
Dim colCp09 As Integer '總收文號(行位置)
Dim colTQF02 As Integer '委查單號
Dim R_type As String 'W:櫃台收文, T:商標承辦
Dim colTQC01 As Integer 'Added by Lydia 2016/05/04
'Added by Lydia 2019/04/19
Dim m_TM10 As String '申請國家
Dim m_CP10 As String '案件性質
Dim m_CP06 As String, m_CP07 As String '所限、法限
Dim m_EP06 As String '文件齊備日
Dim m_CP48 As String '承辦期限

'Modified by Lydia 2018/03/20 先傳變數iCp13
Public Sub SetParent(ByRef fm As Form, Optional ByVal iCp13 As String)
   Set m_PrevForm = fm
   
   'Added by Lydia 2018/03/20
   If iCp13 <> "" Then m_CP13 = iCp13
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim iRow As Integer, bContinue As Boolean
   Dim strFirst As String
   Dim iMode As String, iStr As String
   Dim cnt2 As Integer, cnt3 As Integer

   If mbolCall = True Then GoTo JumpSave
   
   bContinue = False
   If Index = 1 Then

      If TMQ_CtrRead = True Then cnt2 = PUB_MGridGetId("讀", MGRD1)
      If TMQ_ReApp = False Then cnt3 = PUB_MGridGetId("TMQ21", MGRD1)
      
      With MGRD1
      For iRow = 1 To .Rows - 1
         If Trim(.TextMatrix(iRow, 0) & .TextMatrix(iRow, colMTQC02)) <> "" Then
            If strFirst = "" Then
               strFirst = .TextMatrix(iRow, colTQC01)
            ElseIf strFirst <> .TextMatrix(iRow, colTQC01) And .TextMatrix(iRow, colTQC01) <> "" Then
                   MsgBox "請選取同一查名代號的委查單!", vbCritical
                   Exit Sub
            End If
            'Modified by Lydia 2016/03/28 是否控制已讀
            If TMQ_CtrRead And .TextMatrix(iRow, cnt2) <> "Y" And cnt2 > 0 Then
                MsgBox "委查單 " & MGRD1.TextMatrix(iRow, colTQF02) & " 尚未讀取查覆附件,不能收文,請洽委查人!", vbCritical
                Exit Sub
            End If
            'Modified by Lydia 2016/04/06 控制是否可重複申請
            If TMQ_ReApp = False And cnt3 > 0 Then
                If .TextMatrix(iRow, 0) = "V" And .TextMatrix(iRow, cnt3) <> "" And Trim(.TextMatrix(iRow, cnt3)) <> m_CP09 Then
                    MsgBox "委查單 " & MGRD1.TextMatrix(iRow, colTQF02) & " 已做其他收文,請洽委查人!", vbCritical
                    Exit Sub
                End If
            End If
            
            bContinue = True
            '取消已收文
            If .TextMatrix(iRow, 0) = "" And .TextMatrix(iRow, colMTQC02) <> "" Then
               If MsgBox("請確認 " & MGRD1.TextMatrix(iRow, colTQF02) & " 不列入收文範圍?", vbInformation + vbYesNo) = vbNo Then
                  .TextMatrix(iRow, 0) = "V"
                  iMode = ""
                  Exit Sub
               Else
                  iMode = "-"
               End If
            ElseIf .TextMatrix(iRow, 0) = "V" And .TextMatrix(iRow, colMTQC02) = "" Then
               iMode = "+"
            Else
               iMode = ""
            End If
            If iMode <> "" Then iStr = iStr & iMode & .TextMatrix(iRow, colTQF02) & ","
            
         End If
      Next
      End With
   End If
   
   'Added by Lydia 2019/04/19 先處理「無須查名/自行查名」
   If ChkPass.Value = vbChecked Or ChkPass.Tag <> "" Then
        'Added by Lydia 2024/01/09
        If iMode <> "" Then
            MsgBox "請先確認勾選委查單的結果！", vbCritical
            Exit Sub
        End If
        'end 2024/01/09
        If Process1(iStr) = False Then
            Exit Sub
        End If
        '「無須查名/自行查名」，後面不用繼續
        If ChkPass.Value = vbChecked Then
            Unload Me
            Exit Sub
        End If
   End If
   'end 2019/04/19
   
   If bContinue = False Then
        If ChkPass.Value = vbUnchecked Then 'Added by Lydia 2019/04/19
            If MsgBox("確定不要勾選任何委查單嗎?", vbInformation + vbYesNo) = vbYes Then
               cmdExit_Click
            End If
        End If
   ElseIf Index = 1 Then
JumpSave:
       If m_CP09 <> "" Then
          'Added by Lydia 2016/05/04
          If iStr = "" Then
             If MsgBox("委查結果無變更,確定繼續?", vbInformation + vbYesNo) = vbNo Then
                Unload Me
             End If
             Exit Sub
          End If
          'end 2016/05/04

          'Modified by Lydia 2018/12/10 判斷查名是否齊備
          'If PUB_TMQtoCP(m_AttachPath, m_CP09, iStr, IIf(R_type = "W", CompTxTno, "")) = True Then
          'Modified by Lydia 2024/03/14 +True
          'If PUB_TMQtoCP(m_AttachPath, m_CP09, iStr, IIf(R_type = "W", CompTxTno, ""), True) = True Then
          If PUB_TMQtoCP(True, m_AttachPath, m_CP09, iStr, IIf(R_type = "W", CompTxTno, ""), True) = True Then
             MsgBox "收文確定完畢!", vbInformation
             If TypeName(m_PrevForm) <> "Nothing" Then
                  If R_type = "W" Then m_PrevForm.TMQList = ""
                  cmdExit_Click
             Else
                If reSearch < 0 Then
                   OpenTable 0
                Else
                   OpenTable reSearch
                End If
             End If
          End If
       Else
          'Memo by Lydia 2016/04/27 櫃台收文改成直接在畫面輸入查名代號
          If TypeName(m_PrevForm) <> "Nothing" Then
             If CompTxTno <> Trim(txtNo(0).Tag) & Trim(txtNo(1).Tag) Then
                 MsgBox "查詢結果與查名代號不符,請確定!", vbCritical
                 Exit Sub
             End If
             m_PrevForm.TMQList = CompTxTno & "|" & iStr
          End If
          cmdExit_Click
       End If
   End If

End Sub

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MGRD1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MGRD1.MousePointer = vbDefault
End Sub

Private Sub Read_Click(Optional ByRef bolU As Boolean = False)
Dim tmplist As String
Dim iR As Integer
   
    If iPrevRow1 > 0 And bolU = True Then
       tmplist = tmplist & Trim(MGRD1.TextMatrix(iPrevRow1, colTQF02)) & ","
    Else
        For iR = 1 To MGRD1.Rows - 1
           If MGRD1.TextMatrix(iR, 0) = "V" Then
              tmplist = tmplist & Trim(MGRD1.TextMatrix(iR, colTQF02)) & ","
           End If
        Next iR
    End If
    
    If tmplist <> "" Then
       frm090128.m_NoList = tmplist
       frm090128.R_type = "Q"
       frm090128.iStiu = 0
       frm090128.SetParent Me
       frm090128.m_NoIdx = 0
       frm090128.mbolCall = True
       frm090128.Show
       If frm090128.QueryData = True Then
          Me.Hide
       Else
          Unload frm090128
       End If
    End If
End Sub
Private Sub CmdRead_Click()
    Call Read_Click(False)
End Sub

Private Sub cmdRefresh_Click(Index As Integer)
Dim Cancel As Boolean
   For intI = 0 To 6
       txtField_Validate intI, Cancel
       If Cancel = True Then
          Exit Sub
       End If
   Next intI
   
   SetMouseBusy
   reSearch = Index
   OpenTable Index
   SetMouseReady
End Sub

'Remove by Lydia 2018/03/20 改在form_load
'Private Sub Form_Activate()
'
'  If TypeName(m_PrevForm) <> "Nothing" Then
'    'Modified by Lydia 2016/04/25 +TS案收文
'    If m_PrevForm.Name = "frm010004" Or m_PrevForm.Name = "frm010007" Then
'       SSTab1.TabVisible(1) = False
'       cmdRefresh(0).Default = True
'       R_type = "W"
'       If Me.txtNo(1).Enabled = True Then Me.txtNo(1).SetFocus
'    Else
'       SSTab1.TabVisible(0) = False
'       cmdRefresh(1).Default = True
'       R_type = "T"
'    End If
'  Else
'    Me.Caption = "委查結果"
'    SSTab1.TabVisible(0) = False
'    R_type = ""
'  End If
'
'End Sub

Private Sub Form_Load()
Dim oText As TextBox
Dim strTmp As String 'Added by Lydia 2019/04/19

   MoveFormToCenter Me

    m_AttachPath = App.path & "\" & strUserNum
    If Dir(m_AttachPath, vbDirectory) = "" Then
       MkDir m_AttachPath
    End If
    
   Call PUB_GetTMQans("1", True) 'Added by Lydia 2016/06/02 求近似本所案
   
   '清除資料
   lblSname.Caption = ""
   For Each oText In txtField
      oText.Text = "": oText.Tag = ""
   Next
   txtUnicode = ""
   txtField(1) = ChangeWStringToTString(PUB_GetWorkDay1(CompDate(1, -3, strSrvDate(1)), True))
   txtField(2) = strSrvDate(2)
   
   SSTab1.Tab = 1 'Added by Lydia 2021/10/01
   
   'Added by Lydia 2018/03/20 預設資料
   If m_CP13 <> "" Then
       txtField(0).Text = m_CP13
       lblSname.Caption = GetStaffName(m_CP13, True)
   End If
   'end 2018/03/20
   
   If R_type = "" Then
      If TypeName(m_PrevForm) <> "Nothing" Then
         'Modified by Lydia 2016/04/25 +TS案收文
         If m_PrevForm.Name = "frm010004" Or m_PrevForm.Name = "frm010007" Then
            R_type = "W"
         Else
            R_type = "T"
         End If
      End If
   End If
       
   If m_CP09 <> "" Then
      'Modified by Lydia 2016/04/06 改變查名代號規則
      'strSql = "SELECT TMQ20 FROM trademarkquery WHERE TMQ21='" & m_CP09 & "' "
      strSql = "select tqc01 from tmqcasemap where tqc02='" & m_CP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '傳查名代號
         If "" & RsTemp(0) <> "" Then
             SetTxtNo (RsTemp(0))
             cmdRefresh(0).Visible = False
             txtNo(0).Enabled = False
             txtNo(1).Enabled = False
             Call TmpTxtNo
         End If

         Call OpenTable(0)
         If iStiu = 2 Then cmdOK(1).Visible = False

         GoTo JumpReset
      Else
         'Modified by Lydia 2019/04/19
         'MsgBox "無委查結果!"
         strTmp = "無委查結果!"
         
         If R_type = "T" And TMQ_ReApp = True Then
            Frame2.Visible = True
         End If
      End If
      '已發文/取消收文,不可變更
      'Modified by Lydia 2019/01/30 +CP01,CP02,CP03,CP04
      'Modified by Lydia 2019/04/19
      'strSql = "select CP27||CP57,CP01,CP02,CP03,CP04 from caseprogress where cp09='" & m_CP09 & "' "
      strSql = "select cp27||cp57,cp01,cp02,cp03,cp04,cp10,cp06,cp07,cp48,cp64,ep06,tm10 " & _
                  "from caseprogress, engineerprogress,trademark where cp09='" & m_CP09 & "' and cp09=ep02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) "

      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(RsTemp(0)) Then
           MsgBox "已發文或取消收文,不可變更查名結果", vbInformation
           iStiu = 2
         End If
         'Added by Lydia 2019/01/30
         m_CP01 = "" & RsTemp.Fields("CP01")
         m_CP02 = "" & RsTemp.Fields("CP02")
         m_CP03 = "" & RsTemp.Fields("CP03")
         m_CP04 = "" & RsTemp.Fields("CP04")
         'Added by Lydia 2019/04/19 預設無須查名/自行查名
         m_CP06 = "" & RsTemp.Fields("CP06")
         m_CP07 = "" & RsTemp.Fields("CP07")
         m_CP10 = "" & RsTemp.Fields("CP10")
         m_TM10 = "" & RsTemp.Fields("TM10")
         m_EP06 = "" & RsTemp.Fields("EP06")
         m_CP48 = "" & RsTemp.Fields("CP48")
         If "" & RsTemp.Fields("CP64") <> "" And InStr("" & RsTemp.Fields("CP64"), "查名備註:無須查名/自行查名") > 0 _
                  And InStr("" & RsTemp.Fields("CP64"), "查名備註:取消無須查名/自行查名") = 0 Then
             ChkPass.Value = 1
             ChkPass.Tag = "Y"
             If strTmp <> "" Then
                 MsgBox "本案無須查名/自行查名 !", vbInformation
                 strTmp = ""
             End If
         End If
         'end 2019/04/19
      End If
      If strTmp <> "" Then MsgBox strTmp 'Added by Lydia 2019/04/19
      
      If iStiu = 2 Then
         cmdOK(1).Visible = False
         Frame2.Visible = False
      End If
      
   ElseIf R_type <> "" Then
      If m_PrevForm.TMQList <> "" And Len(CompTxTno) <= 3 Then
          SetTxtNo (Mid(m_PrevForm.TMQList, 1, InStr(m_PrevForm.TMQList, "|") - 1))
         Call OpenTable(0)
         GoTo JumpReset
      End If
   End If
   If Len(CompTxTno) <= 3 Then
      '改成民國年
      txtNo(0) = Mid(strSrvDate(2), 1, 3)
   End If
   Call TmpTxtNo
   
   reSearch = -1
   
   KillTemp

   SetGrid1 True

JumpReset:
End Sub

Private Sub txtField_GotFocus(Index As Integer)
txtField(Index).SelStart = 0
txtField(Index).SelLength = Len(txtField(Index))
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1, 2 '申請日期
          KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
          KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
Dim tmpArr As Variant
Dim i As Integer
Dim tmpGrp As String
Dim rsMe As New ADODB.Recordset
Dim bolSpec As Boolean

   Select Case Index
      Case 0
         If txtField(0) <> Empty Then
            If ClsPDGetStaff(txtField(0).Text, strExc(1), strExc(2)) Then
               lblSname.Caption = strExc(1)
            Else
               txtField(0).SetFocus
               Cancel = True
            End If
         Else
            lblSname.Caption = ""
         End If
      Case 1, 2
         If txtField(Index) <> "" Then
            If CheckIsTaiwanDate(txtField(Index)) Then
               If Index = 2 And txtField(Index) < txtField(Index - 1) Then
                  MsgBox "申請期間終止日不可小於起始日!", vbCritical
                  txtField(Index).SetFocus
                  Cancel = True
               End If
            Else
               txtField(Index).SetFocus
               Cancel = True
            End If
         End If
      Case 5, 6
         If txtField(Index) <> "" Then
            If Left(txtField(Index), 1) = "H" And IsNumeric(Mid(txtField(Index), 3, 1)) Then
                If Index = 6 And txtField(Index) < txtField(Index - 1) Then
                   MsgBox "委查單號止不可小於委查單號起!", vbCritical
                   txtField(Index).SetFocus
                   Cancel = True
                End If
            Else
               MsgBox "委查單號的編碼為H!", vbCritical
               txtField(Index).SetFocus
               Cancel = True
            End If
         End If
      Case 7
         'Modified by Lydia 2016/06/08
         'If txtField(Index) <> "" And txtField(Index) <> "T" Then
         '   MsgBox "請輸入T案!", vbCritical
         If txtField(Index) <> "" And txtField(Index) <> "T" And txtField(Index) <> "TS" Then
            MsgBox "請輸入T、TS案!", vbCritical
            txtField(Index).SetFocus
            Cancel = True
         End If
      Case Else
   End Select
End Sub

Private Sub SetGrid1(Optional pReset As Boolean = False, Optional ByVal iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   'Modified by Lydia 2016/04/21 + TQA21
   arrGridHeadText = Array("V", "讀", "文字1", "文字2", "客戶名稱", "結果", "組群", "總收文號", "檔案名稱", "申請編號", "委查單號", "TQC01", "TMQ02", "MTQC02", "TMQ21", "TQA02")
   If R_type = "W" Or R_type = "" Then
      arrGridHeadWidth = Array(200, 300, 860, 860, 800, 860, 0, 0, 0, 0, 1000, 0, 0, 0, 0, 0)
   Else
      arrGridHeadWidth = Array(200, 300, 860, 860, 800, 860, 860, 0, 1500, 0, 1000, 0, 0, 0, 0, 0)
   End If

   MGRD1.Visible = False
   MGRD1.Cols = UBound(arrGridHeadText) + 1
   MGRD1.Rows = iR
   With MGRD1
        If pReset = True Then
           .Clear
           .Rows = 2
           iPrevRow1 = 0
           lTotRows1 = 0
           lSelRows1 = 0
           lblCount(1) = lSelRows1 & " / " & lTotRows1
        End If
        For iRow = 0 To .Cols - 1
           .row = 0
           .col = iRow
           .Text = arrGridHeadText(iRow)
           .ColWidth(iRow) = arrGridHeadWidth(iRow)
           .CellAlignment = flexAlignCenterCenter
        Next
        For intI = 1 To iR - 1
          .row = intI
          For iRow = 0 To 5
            .col = iRow
            .CellBackColor = QBColor(15)
          Next iRow
        Next intI
   End With
   
   MGRD1.Visible = True
   colMTQC02 = PUB_MGridGetId("MTQC02", MGRD1)
   colCp09 = PUB_MGridGetId("總收文號", MGRD1)
   colTQF02 = PUB_MGridGetId("委查單號", MGRD1)
   colTQC01 = PUB_MGridGetId("TQC01", MGRD1) 'Added by Lydia 2016/05/04
End Sub

Private Sub OpenTable(ByVal aKind As Integer)
   Dim iRow As Integer
   Dim stCon As String  '共同查詢(含查名代號,收文號)
   Dim st2Con As String '條件查詢
   Dim caseStr As String
   Dim Cancel As Boolean
   Dim btTemp() As Byte
   'iStiu '0:新增收文, 1:修改,  2:查詢
   
   stCon = ""
   st2Con = ""
   
   
   If iStiu = 0 And TMQ_CtrRead Then
      '控制是否查覆完畢
      st2Con = st2Con & " AND TMQ11>0"
   End If
   'Modified by Lydia 2016/07/06 改TQC07
   'caseStr = "DECODE(TQC07,NULL,'V','') 已收"
   caseStr = "'V' 已收"
   '查名代號
   If aKind = 0 Then
      If Len(CompTxTno) = 9 Then
         stCon = stCon & " AND TQC01='" & CompTxTno & "'"
         If m_CP09 = "" Or Trim(txtNo(0).Text) & Trim(txtNo(1).Text) <> Trim(txtNo(0).Tag) & Trim(txtNo(1).Tag) Then
            stCon = stCon & " AND TQC02 IS NULL "
         End If
      ElseIf m_CP13 <> "" Then
            If m_CP09 = "" Then
               MsgBox "請輸入查名代號!", vbCritical
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            st2Con = st2Con & " AND TMQ02=" & CNULL(m_CP13)
      End If
   Else
        '委查人
        If txtField(0) <> "" Then st2Con = st2Con & " AND TMQ02=" & CNULL(Trim(txtField(0)))
        '委查日期
        If txtField(1) <> "" Or txtField(2) <> "" Then
           If txtField(1) = "" Or txtField(2) = "" Then
              st2Con = st2Con & " AND TMQ04=" & IIf(txtField(1) <> "", DBDATE(txtField(1)), DBDATE(txtField(2)))
           Else
              st2Con = st2Con & " AND TMQ04>=" & DBDATE(txtField(1)) & " And TMQ04<=" & DBDATE(txtField(2))
           End If
        End If
        '客戶名稱
        If txtField(3) <> "" Then st2Con = st2Con & " AND TQA04 LIKE '%" & Trim(txtField(3)) & "%'"
        '混合輸入
        txtField(4) = txtUnicode
        'Added by Lydia 2021/10/01 用日期控制不用經過二進位處理存入TQA07-TQA08，直接存入TQA13-TQA14
        If strSrvDate(1) >= Form20上線日 Then
            If Trim(txtUnicode) <> "" Then
                st2Con = st2Con & " AND (TQA13 LIKE '%" & Trim(txtUnicode) & "%' OR TQA14 LIKE '%" & Trim(txtUnicode) & "%') "
            End If
        Else
        'end 2021/10/01
            If txtField(4).Text <> txtUnicode.Text Then
               Call UnicodeWR(strExc(1))
               st2Con = st2Con & " AND (TQA07 LIKE '%" & strExc(1) & "' OR TQA08 LIKE '%" & strExc(1) & "') "
            ElseIf txtField(4).Text <> "" Or txtUnicode.Text <> "" Then
               st2Con = st2Con & " AND (TQA13 LIKE '%" & Trim(txtField(4)) & "%' OR TQA14 LIKE '%" & Trim(txtField(4)) & "%') "
            End If
        End If 'Added by Lydia 2021/10/01
        
        '委查單號
        If txtField(5) <> "" Or txtField(6) <> "" Then
           If txtField(5) = "" Or txtField(6) = "" Then
              st2Con = st2Con & " AND TMQ01=" & CNULL(IIf(txtField(5) <> "", txtField(5), txtField(6)))
           Else
              st2Con = st2Con & " AND TMQ01>=" & CNULL(txtField(5)) & " And TMQ01<=" & CNULL(txtField(6))
           End If
        End If
        
        'Added by Lydia 2016/04/18
        '不含已收文
        If Check1.Value = 0 Then
           st2Con = st2Con & " AND TMQ21 IS NULL"
        End If
   End If
   If m_CP09 <> "" And stCon = "" Then  '追加查名結果,無查名代號
       stCon = stCon & " AND TQC02='" & m_CP09 & "'"
   End If
   
   'Added by Lydia 2018/03/20 用TMQ20判斷是否已刪除明細
   'Remove by Lydia 2018/03/21 影響速度
   'stCon = stCon & " And nvl(tmq20,'N') = 'N' "
   'st2Con = st2Con & " And nvl(tmq20,'N') = 'N' "
   'end 2018/03/20
   
   SetGrid1 True
'   strExc(0) = "SELECT " & caseStr & ",TMQ19 讀,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1,DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,'','(非正常字)'),TQA14),'2','') 文字2," & _
'               "TQA04 客戶名稱,decode(v1c2," & TMQ_結果查詢 & ") 結果, TMQ03 組群,TMQ21 總收文號,CPP02 檔案名稱,TMQ18 申請編號,TMQ01 委查單號,TMQ20,TMQ02 " & _
'               "FROM TMQAPP,trademarkquery,CASEPAPERPDF,STAFF,(select tqd02 v1c1, min(tqd06) v1c2 from tmqdetail group by tqd02) VT1 " & _
'               "WHERE TQA01=TMQ18(+) AND TMQ21=CPP01(+) and tmq01=v1c1(+) AND TMQ02=ST01(+) " & IIf(TMQ_CtrRead = True, "AND TQA09>=" & CNULL(TMQ電子化啟用日, True), "") & _
'               "AND (TMQ21 IS NULL OR (TMQ21>'A' AND INSTR(upper(CPP02),upper(TMQ01||'.'||'" & UCase(TMQ_查名作業 & ".menu") & "')) > 0 )) " & _
'               "AND TQA01 IN (SELECT TMQ18 FROM trademarkquery WHERE 1=1" & stCon & ") " & st2Con & _
'               "AND (V1C2 IS NULL OR V1C2<" & CNULL(TMQ_不查, True) & ")"
'   '若新增櫃台收文,預設帶未收文的查名單
'   If m_CP09 = "" Then 'Or R_type = "T" Then 或承辦人進度維護(條件查詢排除)
'      strExc(0) = strExc(0) & " AND TMQ21 IS NULL "
'   End If
'   '有收文號,預設帶入收文號的委查單
'   If m_CP09 <> "" Then
'       strExc(2) = "Union SELECT " & caseStr & ",TMQ19 讀,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1,DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,'','(非正常字)'),TQA14),'2','') 文字2," & _
'               "TQA04 客戶名稱,decode(v1c2," & TMQ_結果查詢 & ") 結果, TMQ03 組群,TMQ21 總收文號,CPP02 檔案名稱,TMQ18 申請編號,TMQ01 委查單號,TMQ20,TMQ02 " & _
'               "FROM TMQAPP,trademarkquery a1,CASEPAPERPDF,STAFF,(select tqd02 v1c1, min(tqd06) v1c2 from tmqdetail group by tqd02) VT1 " & _
'               "WHERE TQA01=TMQ18(+) AND TMQ21=CPP01(+) and tmq01=v1c1(+) AND TMQ02=ST01(+) " & IIf(TMQ_CtrRead = True, "AND TQA09>=" & CNULL(TMQ電子化啟用日, True), "") & _
'               "AND (TMQ21 IS NULL OR (TMQ21>'A' AND INSTR(upper(CPP02),upper(TMQ01||'.'||'" & UCase(TMQ_查名作業 & ".menu") & "')) > 0 )) " & _
'               "and exists (select * from trademarkquery a2 where a1.tmq01=a2.tmq01(+) AND TMQ21='" & m_CP09 & "')"
'      strExc(0) = strExc(0) & strExc(2)
'   End If

   'Modified by Lydia 2016/04/06
   '查名代號查詢,有收文對照檔
   'Modified by Lydia 2016/04/21 + TQA02
   'Modified by Lydia 2016/06/01 覆核結果取代查名結果MIN(TQD06)=>MIN(NVL(TQD09,TQD06))
   'Modified by Lydia 2016/06/02 TMQ_結果查詢改成模組PUB_GetTMQans
   'Modified by Lydia 2016/07/06 DECODE(TQC07,'N','',TQC02) MTQC02 =>直接抓TQC02
   strExc(2) = "SELECT " & caseStr & ",TMQ19 讀,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1,DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,'','(非正常字)'),TQA14),'2','') 文字2," & _
            "TQA04 客戶名稱,DECODE(V1C2," & PUB_GetTMQans("3", True) & ") 結果, TMQ03 組群,TQC02 總收文號,TMQ01||'." & TMQ_查名作業 & ".menu" & "' 檔案名稱,TMQ18 申請編號,TMQ01 委查單號,TQC01,TMQ02,TQC02 MTQC02,TMQ21,TQA02 " & _
            "FROM TMQAPP,TRADEMARKQUERY ,TMQCASEMAP,(SELECT TQD02 V1C1, MIN(NVL(TQD09,TQD06)) V1C2 FROM TMQDETAIL GROUP BY TQD02) VT1 " & _
            "WHERE TQA01=TMQ18(+) AND TQA20 IS NULL AND TMQ01=V1C1(+) AND TMQ01=TQC03(+) " & IIf(TMQ_CtrRead = True, "AND TQA09>=" & CNULL(TMQ電子化啟用日, True), "")
   '若新增櫃台收文,預設帶未收文的查名單
   strExc(0) = strExc(2) & stCon & IIf(m_CP09 = "", " AND TQC02 IS NULL ", "")
   
   'Remove by Lydia 2016/07/06 取消收文不保留記錄
   ''承辦人取消收文
   ' strExc(2) = "SELECT " & caseStr & ",TMQ19 讀,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1,DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,'','(非正常字)'),TQA14),'2','') 文字2," & _
            "TQA04 客戶名稱,DECODE(V1C2," & PUB_GetTMQans("3", True) & ") 結果, TMQ03 組群,TQC02 總收文號,TMQ01||'." & TMQ_查名作業 & ".menu" & "' 檔案名稱,TMQ18 申請1編號,TMQ01 委查單號,TQC01,TMQ02,TQC02 MTQC02,TMQ21,TQA02 " & _
            "FROM TMQAPP,TRADEMARKQUERY ,TMQCASEMAP,(SELECT TQD02 V1C1, MIN(NVL(TQD09,TQD06)) V1C2 FROM TMQDETAIL GROUP BY TQD02) VT1 " & _
            "WHERE TQA01=TMQ18(+) AND TQA20 IS NULL AND TMQ01=V1C1(+) AND TMQ01=TQC03(+) " & IIf(TMQ_CtrRead = True, "AND TQA09>=" & CNULL(TMQ電子化啟用日, True), "")
   ' strExc(0) = strExc(0) & " Union " & strExc(2) & stCon & " AND TQC07 IS NOT NULL"
    
   'Modified by Lydia 2018/03/20 按查詢鈕才加入
   'If m_CP09 <> "" And st2Con <> "" Then
   If aKind <> 0 And m_CP09 <> "" And st2Con <> "" Then
      '無收文對照檔
      strExc(2) = " UNION SELECT '' 已收,TMQ19 讀,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1,DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,'','(非正常字)'),TQA14),'2','') 文字2," & _
                  "TQA04 客戶名稱,DECODE(V1C2," & PUB_GetTMQans("3", True) & ") 結果, TMQ03 組群," & CNULL(m_CP09) & " 總收文號,TMQ01||'." & TMQ_查名作業 & ".menu" & "' 檔案名稱,TMQ18 申請編號,TMQ01 委查單號,'' TQC01,TMQ02,'' MTQC02,TMQ21,TQA02 " & _
                  "FROM TMQAPP,TRADEMARKQUERY,(SELECT TQD02 V1C1, MIN(NVL(TQD09,TQD06)) V1C2 FROM TMQDETAIL GROUP BY TQD02) VT1 " & _
                  "WHERE TQA01=TMQ18(+) AND TQA20 IS NULL AND TMQ01=V1C1(+) " & IIf(TMQ_CtrRead = True, "AND TQA09>=" & CNULL(TMQ電子化啟用日, True), "") & _
                  "AND TMQ01 NOT IN (SELECT TQC03 FROM TMQCASEMAP WHERE TQC02='" & m_CP09 & "') "
      strExc(0) = strExc(0) & strExc(2) & st2Con
   End If
   
   strExc(0) = strExc(0) & " ORDER BY 1 DESC,申請編號,委查單號"

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MGRD1
        .Visible = False
        .FixedCols = 0
        Set .Recordset = RsTemp
        SetGrid1 False, RsTemp.RecordCount + 1
        .FixedCols = 6
        RsTemp.MoveFirst
        strExc(1) = "" & RsTemp.Fields("TQA02") 'Added by Lydia 2016/04/21
        Do While Not RsTemp.EOF
           If RsTemp.Fields("已收") = "V" Then lSelRows1 = lSelRows1 + 1
           RsTemp.MoveNext
        Loop
        lTotRows1 = RsTemp.RecordCount
        lblCount(1) = lSelRows1 & " / " & lTotRows1
        .Visible = True
      End With
      TmpTxtNo
      'Added by Lydia 2016/04/21 提示智權人員不是委查人
      If m_CP13 <> "" And m_CP13 <> strExc(1) And aKind = 0 And m_CP09 = "" Then
         strExc(2) = GetStaffName(strExc(1))
         MsgBox "查名代號的申請人(" & strExc(2) & ")與接洽單的智權人員不同，請確認資料!", vbCritical
      End If
      'end 2016/04/21
   Else
      If aKind = 0 And (m_CP09 = "" Or Trim(txtNo(0).Text) & Trim(txtNo(1).Text) <> Trim(txtNo(0).Tag) & Trim(txtNo(1).Tag)) Then
         MsgBox "資料庫無資料或查名代號已收文 !", vbInformation
      Else
         MsgBox "資料庫無資料 !", vbInformation
      End If
   End If
End Sub

Private Sub SelectRow(ByRef pRow As Integer, ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With FlexGrid
   nCol = .col
   If pPrevRow > 0 Then
      If pPrevRow <> pRow Then
         .row = pPrevRow
      For iCol = 0 To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
         Next
      End If
   End If

   If pRow > 0 Then
      .row = pRow
      For iCol = 0 To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   pPrevRow = pRow
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TypeName(m_PrevForm) <> "Nothing" Then
         Set m_PrevForm.Tmpfrm090130 = Me
         m_PrevForm.Show
         'Added by Lydia 2016/03/28 提醒承辦人追蹤覆核結果
         If m_PrevForm.Name = "frm090201_b" Then
            'Modified by Lydia 2019/01/30 增加判斷查名是否齊備
            'If m_PrevForm.GetCheckTMQ23(m_CP09) = False Then
            '   m_PrevForm.cmd(2).Enabled = False
            If m_PrevForm.GetCheckTMQ23(m_CP09, True) = False Or m_PrevForm.cmd(5).Tag = "N" Then
               m_PrevForm.cmd(2).Enabled = False
               m_PrevForm.cmd(5).Enabled = False
            Else
               m_PrevForm.cmd(2).Enabled = True
            'end 2019/01/30
            End If
         End If
         'end 2016/03/28
    End If
    
    m_CP09 = ""
    mbolCall = False
    Set m_PrevForm = Nothing
    Set frm090130 = Nothing
End Sub

Private Sub KillTemp()
On Error GoTo ErrHnd
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
   Exit Sub
   
ErrHnd:
   Resume Next
End Sub

Private Sub MGRD1_DblClick()

   If MGRD1.MouseRow > 0 Then
     Call Read_Click(True)
   End If
End Sub

Private Sub MGRD1_Click()
   Dim nCol As Integer, nRow As Integer, iRow As Integer, iCol As Integer
   Dim stValue As String
   Dim stCP09 As String
      
   With MGRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      '紀錄前次點選的收文號
      If iPrevRow1 > 0 Then
         stCP09 = MGRD1.TextMatrix(iPrevRow1, colCp09)
      End If
      
      .col = nCol
      If m_blnColOrderAsc1 = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc1 = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc1 = False
      End If
               
      '重設排序後前次點選的位置
      If iPrevRow1 > 0 Then
         For iRow = 1 To .Rows - 1
            If stCP09 = MGRD1.TextMatrix(iRow, colCp09) Then
               iPrevRow1 = iRow
               Exit For
            End If
         Next
      End If
   ElseIf nRow > 0 And .TextMatrix(nRow, 4) <> "" Then
      SelectRow nRow, MGRD1, iPrevRow1
      
      .row = nRow
      .col = nCol
      If nCol = 0 Then
         ClickGrid MGRD1, 1
      End If

   End If
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(ByRef FlexGrid As MSHFlexGrid, Index As Integer)
   Dim iCol As Integer
     '新增或修改收文的查名結果才可變更
     If iStiu = 0 Or iStiu = 1 Then
        With FlexGrid
           If Index = 1 Then
              If .TextMatrix(.row, 0) = "V" Then
                 .TextMatrix(.row, 0) = ""
                 lSelRows1 = lSelRows1 - 1
              ElseIf .TextMatrix(.row, 0) = "" Then
                 .TextMatrix(.row, 0) = "V"
                 lSelRows1 = lSelRows1 + 1
              End If
              lblCount(1) = lSelRows1 & " / " & lTotRows1
           End If
        End With
     End If
End Sub


'Oracle 8i 用RAW存UNICODE, 之後用UNICODE查詢
'1. Oracle升級到 9,使用內建函數UTL_RAW.CAST_TO_RAW ('字串') (未測試)
'2. 用like '%unicode雙位元的內碼%'; (只能做模糊比對,有可能誤判)
Private Sub UnicodeWR(ByRef UniStrCode As String)
Dim btHead(1) As Byte
Dim btTemp() As Byte
Dim btR() As Byte
Dim p As String, p2 As String, p3 As String
Dim idx As Integer
    '先刪檔
    If Dir(m_AttachPath & "\readunicode.txt") <> "" Then Kill m_AttachPath & "\readunicode.txt"
    UniStrCode = ""
    btHead(0) = 255
    btHead(1) = 254
    '寫入二進位元檔
    If txtUnicode.Text <> "" Then
        btTemp = txtUnicode.Text
        Open m_AttachPath & "\readunicode.txt" For Binary As #1
        Put #1, , btHead
        Put #1, , btTemp
        Close #1
    End If
    '讀取二進位元檔
     If txtUnicode.Text <> "" Then
        Open m_AttachPath & "\readunicode.txt" For Binary As #1
        Get #1, , btHead
        Get #1, , btTemp
        btR = btTemp
        Close #1
     End If
     For idx = 1 To UBound(btR)
        '因為實際寫入到DB的碼有經過DB處理,可能有部份差異,所以每一位元加%做模糊比對
        UniStrCode = UniStrCode & Hex$(btR(idx)) & "%"
     Next idx

End Sub
'Modified by Lydia 2016/04/06 改變查名代號(收文組群)規則
Private Function CompTxTno() As String
'原本記錄收文的最小委查單號,改成民國年+流水號6碼
    'CompTxTno = Trim(txtNo(0)) & Trim(txtNo(1)) & Trim(txtNo(2))
    CompTxTno = Trim(txtNo(0)) & Trim(txtNo(1))
End Function
Private Sub TmpTxtNo()
  txtNo(0).Tag = txtNo(0).Text
  txtNo(1).Tag = txtNo(1).Text
End Sub

Private Sub SetTxtNo(ByVal RCode As String)
    txtNo(0) = Mid(RCode, 1, txtNo(0).MaxLength)
    txtNo(1) = Mid(RCode, txtNo(0).MaxLength + 1, txtNo(1).MaxLength)
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
    TextInverse txtNo(Index)
End Sub

Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If txtNo(Index).Text > txtNo(Index).Tag Or Val(txtNo(Index).Text) - Val(txtNo(Index).Tag) < -1 Or Val(txtNo(Index).Text) - Val(txtNo(Index).Tag) > 1 Then
           MsgBox "年份輸入錯誤!!"
           Cancel = True
           txtNo(Index).SetFocus
        End If
    End If
End Sub
Private Sub cmdCopy_Click()
Dim strBatch As String

   If txtField(9) = "" Then txtField(9) = "0"
   If txtField(10) = "" Then txtField(10) = "00"
   
   If lTotRows1 > 0 Then
      If MsgBox("複製前案不可和勾選結果同時做收文,確定繼續複製前案?", vbCritical + vbYesNo) = vbNo Then
         Exit Sub
      End If
   End If
   'Modified by Lydia 2016/04/25 +TS案
   'strExc(0) = "select cp09,tqc01,tqc02,tqc03 from caseprogress,tmqcasemap where cp01='" & txtField(7) & "' and cp02='" & txtField(8) & "' and cp03='" & txtField(9) & "' and cp04='" & txtField(10) & "'" & _
               " and cp10='101' and cp57 is null and cp09=tqc02(+) and tqc07 is null order by tqc03"
   'Modified by Lydia 2016/07/06 改TQC07
   'strExc(0) = "select cp09,tqc01,tqc02,tqc03 from caseprogress,tmqcasemap where cp01='" & txtField(7) & "' and cp02='" & txtField(8) & "' and cp03='" & txtField(9) & "' and cp04='" & txtField(10) & "'" & _
               " and cp10='" & IIf(txtField(7) = "T", TMQ_T案, IIf(txtField(7) = "TS", TMQ_TS案, "")) & "' and cp57 is null and cp09=tqc02(+) and tqc07 is null order by tqc03"
   'Modified by Lydia 2021/11/19 增加737智財協作之T案
   'strExc(0) = "select cp09,tqc01,tqc02,tqc03 from caseprogress,tmqcasemap where cp01='" & txtField(7) & "' and cp02='" & txtField(8) & "' and cp03='" & txtField(9) & "' and cp04='" & txtField(10) & "'" & _
               " and cp10='" & IIf(txtField(7) = "T", TMQ_T案, IIf(txtField(7) = "TS", TMQ_TS案, "")) & "' and cp57 is null and cp09=tqc02(+) order by tqc03"
   strExc(0) = "select cp09,tqc01,tqc02,tqc03 from caseprogress,tmqcasemap where cp01='" & txtField(7) & "' and cp02='" & txtField(8) & "' and cp03='" & txtField(9) & "' and cp04='" & txtField(10) & "'" & _
               " and instr('" & IIf(txtField(7) = "T", TMQ_T案, IIf(txtField(7) = "TS", TMQ_TS案, "")) & "', cp10) > 0  and cp57 is null and cp09=tqc02(+) order by tqc03"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If IsNull(RsTemp.Fields("tqc02")) Or IsNull(RsTemp.Fields("tqc03")) Then
            Exit Do
         Else
            strBatch = strBatch & "+" & Trim(RsTemp.Fields("tqc03")) & ","
         End If
         RsTemp.MoveNext
      Loop
      If m_CP09 <> "" And strBatch <> "" Then
         'Modified by Lydia 2018/12/10 判斷查名是否齊備
         'If PUB_TMQtoCP(m_AttachPath, m_CP09, strBatch, "") = True Then
         'Modified by Lydia 2024/03/14 +True
         'If PUB_TMQtoCP(m_AttachPath, m_CP09, strBatch, "", True) = True Then
         If PUB_TMQtoCP(True, m_AttachPath, m_CP09, strBatch, "", True) = True Then
             MsgBox "複製完畢!", vbInformation
             Call OpenTable(0) '重整
             Frame2.Enabled = False
         End If
      Else
         If strBatch = "" Then
             MsgBox "該案無查名結果可複製!", vbCritical
         End If
      End If
   End If
End Sub

'Added by Lydia 2019/04/19 無須查名/自行查名: MCTF案件不用查名或智權人員自行查名將結果PDF寄到查名中心
Private Function Process1(Optional ByVal m_strNo As String = "") As Boolean
Dim strB1 As String
Dim intB As Integer
Dim strM As String
Dim strDate As String

       strM = PUB_TMQchkCP143(m_CP09)
       If ChkPass.Value = vbChecked Then
          strB1 = "Y"
       Else
          strB1 = ""
       End If
       
On Error GoTo ErrorHandle:

       If strB1 <> ChkPass.Tag Then
           If strM <> "" And m_strNo <> "" Then
                MsgBox "已有勾選委查單，不可設為無須查名/自行查名 ！", vbCritical
                Exit Function
           End If
           cnnConnection.BeginTrans
           If strB1 = "Y" Then
                If Val(m_EP06) > 0 Then   '文件齊備＋查名齊備
                    strDate = Pub_GetHandleDay(m_CP01, m_TM10, m_CP10, strSrvDate(1), m_CP06, m_CP09)
                    If strDate <> "" Then
                        strDate = ", cp48=" & strDate
                    End If
                End If
                'Modified by Lydia 2019/10/01 T-223954先設無須查名又「取消無須查名」,最後送件時判斷為無須查名
                'strSql = "update caseprogress set cp143=" & strSrvDate(1) & strDate & " , cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & " 查名備註:無須查名/自行查名(" & Val(Format(ServerTime, "000000")) & " " & strUserNum & ");'||cp64 " & _
                           " where cp09='" & m_CP09 & "' and instr(nvl(cp64,'N'),'無須查名/自行查名') = 0 "
                strSql = "update caseprogress set cp143=" & strSrvDate(1) & strDate & " , cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & " 查名備註:無須查名/自行查名(" & Val(Format(ServerTime, "000000")) & " " & strUserNum & ");'||cp64 " & _
                           " where cp09='" & m_CP09 & "' "
                cnnConnection.Execute strSql, intB
           Else
                strDate = ", cp48=null "
                strSql = "update caseprogress set cp143=null " & strDate & " , cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & " 查名備註:取消無須查名/自行查名(" & Val(Format(ServerTime, "000000")) & " " & strUserNum & ");'||cp64 " & _
                           " where cp09='" & m_CP09 & "' and instr(nvl(cp64,'N'),'無須查名/自行查名') > 0 "
                cnnConnection.Execute strSql, intB
           End If
           cnnConnection.CommitTrans
           ChkPass.Tag = strB1
       End If
       Process1 = True
       Exit Function
       
ErrorHandle:
       If Err.Number <> 0 Then
           cnnConnection.RollbackTrans
           MsgBox Err.Description
       End If
End Function

