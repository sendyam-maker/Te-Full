VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010519 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利處系統收件區"
   ClientHeight    =   6630
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8960
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      Height          =   555
      Left            =   4320
      TabIndex        =   51
      Top             =   630
      Width           =   945
      Begin VB.ComboBox cboReason 
         Height          =   300
         ItemData        =   "frm04010519.frx":0000
         Left            =   495
         List            =   "frm04010519.frx":0010
         Style           =   2  '單純下拉式
         TabIndex        =   52
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "備註/原因："
         Height          =   180
         Left            =   0
         TabIndex        =   53
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "不處理/2次確認信件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   750
      TabIndex        =   46
      Top             =   90
      Width           =   2475
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm04010519.frx":003C
      Left            =   2520
      List            =   "frm04010519.frx":0046
      Style           =   2  '單純下拉式
      TabIndex        =   44
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   330
      Index           =   1
      Left            =   3570
      Style           =   1  '圖片外觀
      TabIndex        =   43
      Top             =   2130
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   345
      Left            =   7320
      TabIndex        =   37
      Top             =   1500
      Visible         =   0   'False
      Width           =   1635
      Begin VB.CommandButton cmdAgree 
         Caption         =   "同意"
         Height          =   330
         Left            =   45
         TabIndex        =   23
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdback 
         Caption         =   "退回"
         Height          =   330
         Left            =   840
         TabIndex        =   24
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdReInput 
         Caption         =   "來函期限2次確認"
         Height          =   330
         Left            =   30
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   345
      Left            =   4710
      TabIndex        =   36
      Top             =   2130
      Width           =   4095
      Begin VB.CommandButton cmdRecall 
         Caption         =   "回覆確收"
         Height          =   330
         Left            =   3060
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdProDel 
         Caption         =   "已處理"
         Height          =   330
         Left            =   2301
         TabIndex        =   20
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdPDF 
         Caption         =   "歸卷"
         Height          =   330
         Left            =   1544
         TabIndex        =   19
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdInput 
         Caption         =   "輸入"
         Height          =   330
         Left            =   30
         TabIndex        =   17
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdNotProDel 
         Caption         =   "不處理"
         Height          =   330
         Left            =   787
         TabIndex        =   18
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      Height          =   555
      Left            =   4740
      TabIndex        =   39
      Top             =   1500
      Width           =   4185
      Begin VB.CommandButton cmdCACK 
         Caption         =   "客戶確收"
         Height          =   300
         Left            =   3270
         TabIndex        =   50
         Top             =   250
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選總收文號"
         Height          =   300
         Left            =   2160
         TabIndex        =   16
         Top             =   250
         Width           =   1095
      End
      Begin VB.TextBox txtPI21 
         Height          =   270
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   14
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtPI19 
         Height          =   270
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtPI18 
         Height          =   270
         Left            =   810
         MaxLength       =   3
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtPI20 
         Height          =   270
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtRecvNo 
         Height          =   270
         Left            =   810
         MaxLength       =   9
         TabIndex        =   15
         Top             =   250
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "本所案號:"
         Height          =   225
         Left            =   0
         TabIndex        =   41
         Top             =   30
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "總收文號:"
         Height          =   225
         Left            =   0
         TabIndex        =   40
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "刪除"
      Height          =   330
      Left            =   4590
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   4980
      TabIndex        =   33
      Top             =   420
      Width           =   3975
      Begin VB.CommandButton cmdSendMail 
         Caption         =   "立即寄發通知信"
         Height          =   300
         Left            =   2490
         TabIndex        =   26
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         Caption         =   "程式結束後”寄發通知信”或"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   60
         TabIndex        =   34
         Top             =   30
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdUpdRow 
      Caption         =   "轉寄"
      Height          =   330
      Left            =   5364
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "信件狀況"
      Height          =   330
      Left            =   7098
      TabIndex        =   5
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   330
      Left            =   3390
      TabIndex        =   1
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "記錄查詢"
      Height          =   330
      Left            =   6156
      TabIndex        =   4
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8040
      TabIndex        =   6
      Top             =   0
      Width           =   765
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm04010519.frx":004F
      Height          =   3885
      Left            =   60
      TabIndex        =   25
      Top             =   2490
      Width           =   8865
      _ExtentX        =   15646
      _ExtentY        =   6862
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|符號|確|收信日期時間|本所案號|主旨|收受者|轉寄者|轉寄日期時間|讀取日期時間|檔名|總收文號|處理原因"
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
      _Band(0).Cols   =   13
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-->"
      Height          =   285
      Left            =   1950
      TabIndex        =   30
      Top             =   1350
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<--"
      Height          =   285
      Left            =   1950
      TabIndex        =   31
      Top             =   1170
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CheckBox Check1 
      Caption         =   "轉寄後不顯示"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1170
      TabIndex        =   10
      Top             =   915
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSForms.TextBox TxtIR20 
      Height          =   735
      Left            =   5250
      TabIndex        =   22
      Top             =   720
      Width           =   3645
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "6429;1296"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox List1 
      Height          =   315
      Left            =   2730
      TabIndex        =   8
      Top             =   660
      Width           =   1545
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2725;556"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextContext 
      Height          =   825
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   4335
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "7646;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   0
      Top             =   360
      Width           =   1605
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "2831;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboII06 
      Height          =   300
      Left            =   1140
      TabIndex        =   7
      Top             =   660
      Width           =   1590
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2805;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   30
      TabIndex        =   48
      Top             =   30
      Visible         =   0   'False
      Width           =   1725
      VariousPropertyBits=   746604571
      Size            =   "3043;503"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "符號："
      Height          =   180
      Left            =   1950
      TabIndex        =   45
      Top             =   2220
      Width           =   540
   End
   Begin VB.Label Label8 
      Caption         =   "轉寄信件內容："
      Height          =   255
      Left            =   150
      TabIndex        =   42
      Top             =   1020
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "下方收受者點二下即可移除"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   2760
      TabIndex        =   38
      Top             =   420
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "* 代表有轉寄給他人"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   2205
      Width           =   1635
   End
   Begin VB.Label Label5 
      Caption         =   "轉寄收受者:"
      Height          =   255
      Left            =   150
      TabIndex        =   32
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   150
      TabIndex        =   29
      Top             =   405
      Width           =   900
   End
   Begin VB.Label LblTotCnt 
      Caption         =   "總筆數:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7320
      TabIndex        =   28
      Top             =   6420
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "備註：雙擊”主旨”開啟信件"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   27
      Top             =   6420
      Width           =   2535
   End
   Begin VB.Label LblSec2Query 
      BackColor       =   &H0080FFFF&
      Height          =   330
      Left            =   660
      TabIndex        =   47
      Top             =   30
      Visible         =   0   'False
      Width           =   2685
   End
End
Attribute VB_Name = "frm04010519"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/19 Form2.0已修改
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Dim m_iRow As Integer
Public m_strIR01 As String, m_strIR02 As String, m_strIR03 As String, m_strIR04 As String
Public m_strPi12 As String
Public cmdState As Integer, bolQuery As Boolean '紀錄作用按鍵
Dim m_strPi05 As String
Dim strIR22Emp As String
Dim m_TxtIR20 As String
Dim m_strUserList As String
Public m_AppNo As String
Public m_RegNo As String
Dim stIR16 As String 'Added by Morgan 2020/8/13


'Add By Sindy 2022/2/8
Private Sub cboReason_Click()
   If cboReason.List(cboReason.ListIndex) = "已處理：" Or cboReason.List(cboReason.ListIndex) = "不處理：" Then
   Else
      If TxtIR20.Text = "" Then
         If Trim(cboReason.List(cboReason.ListIndex)) = "其他" Then
            TxtIR20.Text = "其他，"
         Else
            TxtIR20.Text = Trim(cboReason.List(cboReason.ListIndex))
         End If
      End If
   End If
   txtPI20_GotFocus
   TxtIR20.SetFocus
   cboReason.ListIndex = -1
   cboReason.Width = 315
End Sub
Private Sub cboReason_DropDown()
   cboReason.Width = 3645
End Sub
Private Sub cboReason_LostFocus()
   cboReason.Width = 315
End Sub

'Add By Sindy 2022/6/2
Private Sub Check2_Click()
   Call cmdQuery_Click
End Sub

'同意
Private Sub cmdAgree_Click()
Dim bolSelectRow As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
Dim strCPP02 As String 'Add By Sindy 2020/2/14
Dim strNation As String 'Add By Sindy 2022/11/9
   
   bolSelectRow = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         bolSelectRow = True
         Exit For
      End If
   Next i
   If bolSelectRow = False Then
      MsgBox "請至少勾選一筆要同意的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定同意嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         '歸卷
         If GRD1.TextMatrix(i, 23) = "4" Then
            If Trim(GRD1.TextMatrix(i, 4)) = "" Then
               MsgBox "無本所案號，不可歸檔！", vbExclamation, "警告！"
               Exit Sub
            End If
            If Trim(GRD1.TextMatrix(i, 24)) = "" Then
               MsgBox "無指定總收文號，不可歸檔！", vbExclamation, "警告！"
               Exit Sub
            End If
         End If
         
         'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         Screen.MousePointer = vbHourglass
         strUpdTime = Right("000000" & ServerTime, 6)
         cnnConnection.BeginTrans: bolConn = True
         '歸卷
         If GRD1.TextMatrix(i, 23) = "4" Then
            '下載信件檔,上傳卷宗區
            'Modify By Sindy 2019/3/8 更改P案系統收件區處理歸卷的整封郵件副檔名命名為:PAT.陸代郵件,原為(外來郵件)
            If SystemNumber(Trim(GRD1.TextMatrix(i, 4)), 1) = "P" Or SystemNumber(Trim(GRD1.TextMatrix(i, 4)), 1) = "PS" Then
               'Modify By Sindy 2022/11/9 + IIf(strNation <> 台灣國家代號, "PAT", "RX")
               strNation = GetPrjNation1(Trim(GRD1.TextMatrix(i, 4)))
               If PUB_UploadPatentLetterFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 16), GRD1.TextMatrix(i, 24), IIf(strNation <> 台灣國家代號, "PAT", "RX"), , , strCPP02) = False Then
                  cnnConnection.RollbackTrans
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            Else
            '2019/3/8 END
               If PUB_UploadPatentLetterFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 16), GRD1.TextMatrix(i, 24), , , , strCPP02) = False Then
                  cnnConnection.RollbackTrans
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            'Add By Sindy 2020/2/14 第2次核對,但歸卷的新增人員還是應該維持原處理者
            If "" & GRD1.TextMatrix(i, 27) <> "" Then
               strExc(0) = "update CasePaperPDF set cpp05='" & GRD1.TextMatrix(i, 27) & "'" & _
                           " where cpp01='" & GRD1.TextMatrix(i, 24) & "'" & _
                           " and upper(cpp02)=upper('" & ChgSQL(strCPP02) & "')"
               cnnConnection.Execute strExc(0)
            End If
            '2020/2/14 END
         End If
         
         '同意
         strExc(0) = "update InputRecord set " & _
                     " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
         'Add By Sindy 2018/6/5
         If TxtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20='" & ChgSQL(TxtIR20) & "'"
         End If
         '2018/6/5 END
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         Call SaveInputRecord(i, False)
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 同意失敗！" & vbCrLf & Err.Description
End Sub

'退回
Private Sub cmdBack_Click()
Dim bolHavdBack As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
   
   bolHavdBack = False
   '先檢查是否有資料要退回
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         bolHavdBack = True
         Exit For
      End If
   Next i
   If bolHavdBack = False Then
      MsgBox "請至少勾選一筆要退回的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定要退回信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         '退回
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         
         strExc(0) = "update InputRecord set " & _
                     " ir16='3',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'"
         'Add By Sindy 2018/6/5
         If TxtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20='" & ChgSQL(TxtIR20) & "'"
         End If
         '2018/6/5 END
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         
         strExc(0) = "select cum02 from CaseUseMemo" & _
                     " where cum05='02'" & _
                       " and cum06=" & CNULL(strUserNum) & _
                       " and cum02=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                        " values('0',upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "'),'0','0','02')"
            cnnConnection.Execute strExc(0)
            Frame1.Visible = True '*****
         End If
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 退回失敗！" & vbCrLf & Err.Description
End Sub

'Added by Morgan 2021/3/31
Private Sub cmdCACK_Click()
Dim strUpdTime As String
Dim bolConn As Boolean
   
On Error GoTo ErrHand
   
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            If dblPrevRow <> .row Then
               MsgBox "資料列選取有誤，請重新確認！"
               Exit Sub
            End If
            '檢查資料
            If txtPI18 = "" Or txtPI19 = "" Then
               MsgBox "請輸入本所案號！", vbExclamation, "警告！"
               If txtPI18 = "" Then
                  Me.txtPI18.SetFocus
               ElseIf txtPI19 = "" Then
                  Me.txtPI19.SetFocus
               End If
               Exit Sub
               
            ElseIf Left(GRD1.TextMatrix(.row, 16), 1) = "P" Then '專利處
               If txtPI18 <> "P" And txtPI18 <> "PS" And _
                  txtPI18 <> "CFP" And txtPI18 <> "CPS" Then
                  MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
                  Me.txtPI18.SetFocus
                  Exit Sub
               End If
            
            Else
               MsgBox "有問題請洽電腦中心!!"
               Exit Sub
               
            End If
            
            If txtPI20 = "" Then txtPI20 = "0"
            If txtPI21 = "" Then txtPI21 = "00"
            If txtRecvNo = "" Then
               MsgBox "請選擇確收的總收文號！", vbExclamation, "警告！"
               'Me.txtRecvNo.SetFocus
               cmdSelCp09.Value = True
               If txtRecvNo = "" Then Exit Sub
            End If
            
            strExc(0) = "select cp09,cp01,lp01 from caseprogress,letterprogress" & _
                        " where cp01='" & txtPI18 & "'" & _
                          " and cp02='" & txtPI19 & "'" & _
                          " and cp03='" & txtPI20 & "'" & _
                          " and cp04='" & txtPI21 & "'" & _
                          " and cp09='" & txtRecvNo & "' and lp01(+)=cp09 and lp10(+)='Y' and lp26(+)='E' and lp39(+)>0 and lp47(+)=0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               MsgBox "查無進度資料！", vbExclamation, "警告！"
               Me.txtRecvNo.SetFocus
               Exit Sub
            Else
               If IsNull(RsTemp("lp01")) Then
                  MsgBox "該收文號無待確收信函！", vbExclamation, "警告！"
                  Me.txtRecvNo.SetFocus
                  Exit Sub
               Else
                  
               End If
            End If
            
            '確收
            cnnConnection.BeginTrans: bolConn = True
            
            If PUB_RecpConfirm(txtRecvNo, Me) = False Then
               cnnConnection.RollbackTrans
               Exit Sub
            End If
                  
            strUpdTime = Right("000000" & ServerTime, 6)
            
            '更新信函確收日期
            strExc(0) = "update letterprogress set lp46='" & strUserNum & "'" & _
                        ",lp47='" & strSrvDate(1) & "',lp48='" & strUpdTime & "'" & _
                        " where lp01='" & txtRecvNo & "' and lp47=0"
            cnnConnection.Execute strExc(0)
            
            
            strExc(0) = "update patentinput set " & _
                        "pi18='" & txtPI18 & "',pi19='" & txtPI19 & "'," & _
                        "pi20='" & txtPI20 & "',pi21='" & txtPI21 & "'" & _
                        " where pi01=" & GRD1.TextMatrix(.row, 10) & _
                        " and pi02=" & GRD1.TextMatrix(.row, 11) & _
                        " and pi03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'"
            cnnConnection.Execute strExc(0)

            .TextMatrix(.row, 4) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
            .TextMatrix(.row, 24) = txtRecvNo
                        
            Screen.MousePointer = vbHourglass
            '下載信件檔,上傳卷宗區
            If PUB_UploadPatentLetterFile(GRD1.TextMatrix(.row, 10), GRD1.TextMatrix(.row, 16), txtRecvNo, "CACK") = False Then
               If bolConn = True Then cnnConnection.RollbackTrans
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            
            strExc(0) = "update InputRecord set " & _
                        " ir16='5',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                        ",ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
            'Add By Sindy 2021/4/8
            If TxtIR20 <> m_TxtIR20 Then
               strExc(0) = strExc(0) & _
                           ",IR20='" & ChgSQL(TxtIR20) & "'"
            End If
            '2021/4/8 END
            strExc(0) = strExc(0) & _
                        " where ir01=" & GRD1.TextMatrix(.row, 10) & _
                          " and ir02=" & GRD1.TextMatrix(.row, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(.row, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
            
            Call SaveInputRecord(.row, False)
            
            cnnConnection.CommitTrans: bolConn = False
            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
            Call CancelRowColor(.row) '清除反白
            GRD1.RowHeight(.row) = 0
            Exit For
         End If
      Next m_iRow
   End With
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox " 信件歸卷註記失敗！" & vbCrLf & Err.Description
End Sub

'刪除鍵
Private Sub cmdDelRow_Click()
Dim bolSelectRow As Boolean
Dim strUpdTime As String
Dim bolConn As String
   
   bolSelectRow = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         bolSelectRow = True
         Exit For
      End If
   Next i
   If bolSelectRow = False Then
      MsgBox "請至少勾選一筆要刪除的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定要刪除信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         'Add By Sindy 2022/2/22 針對分類有*號時，要增加提醒訊息。
         Dim strBox As String
         If InStr(GRD1.TextMatrix(i, 29), "*") > 0 Then
            strBox = Mid(GRD1.TextMatrix(i, 29), InStr(GRD1.TextMatrix(i, 29), "*") - 1, 1)
            If MsgBox(GRD1.TextMatrix(i, 5) & vbCrLf & vbCrLf & _
                   "提醒：信件狀態「" & IIf(strBox = "F", "國外部", IIf(strBox = "P", "專利處", "商標處")) & "」是直接刪除的，若人員處理信件時發現非該單位信件，請轉寄回該信箱。" & vbCrLf & vbCrLf & _
                   "確定要刪除信件嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
               'Call CancelRowColor(i) '清除反白
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
         '2022/2/22 END
         
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
                  
         '刪除
         strExc(0) = "update InputRecord set " & _
                     " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
         'Add By Sindy 2021/1/22
         If TxtIR20 <> m_TxtIR20 Then
            strExc(0) = strExc(0) & _
                        ",IR20='" & ChgSQL(TxtIR20) & "'"
         End If
         '2021/1/22 END
         strExc(0) = strExc(0) & _
                     " where ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                       " and ir08=0"
         cnnConnection.Execute strExc(0)
         Call SaveInputRecord(i, False)
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 刪除失敗！" & vbCrLf & Err.Description
End Sub

Private Function CheckDataValid(Optional strPI06 As String = "") As Boolean
Dim intTotList As Integer
Dim strChkEmp As String, strChkName As String
   
   CheckDataValid = False
   TextContext.Enabled = False
   '檢查收受者是否重覆
   If strPI06 <> "" Or List1.ListCount > 0 Then
      '欲檢查幾個收受者
      If strPI06 <> "" Then
         intTotList = 0
      Else
         intTotList = List1.ListCount - 1
      End If
      Screen.MousePointer = vbHourglass
      For i = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(i, 0) = "V" Then
            For j = 0 To intTotList
               If strPI06 <> "" Then
                  strChkEmp = Left(strPI06, 5)
                  strChkName = Trim(Mid(strPI06, 6)) 'Add By Sindy 2021/11/19
               Else
                  strChkEmp = Left(List1.List(j), 5)
                  strChkName = Trim(Mid(List1.List(j), 6)) 'Add By Sindy 2021/11/19
               End If
               
               '非專利處程序人員
               If PUB_GetST03(strChkEmp) <> "P12" Then
                  TextContext.Enabled = True
               End If
               
               strExc(0) = "select ir04 from inputrecord" & _
                           " where ir01=" & GRD1.TextMatrix(i, 10) & _
                             " and ir02=" & GRD1.TextMatrix(i, 11) & _
                             " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                             " and ir04='" & strChkEmp & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Screen.MousePointer = vbDefault
                  MsgBox "收受者（" & strChkName & "）已收過此郵件！" & vbCrLf & _
                         GRD1.TextMatrix(i, 3) & vbCrLf & _
                         GRD1.TextMatrix(i, 5), vbExclamation
                  Me.List1.SetFocus
                  Exit Function
               End If
            Next j
         End If
      Next i
      Screen.MousePointer = vbDefault
   End If
   
   CheckDataValid = True
End Function

Private Sub CancelRowColor(intRow As Integer)
   '清除反白
   GRD1.TextMatrix(intRow, 0) = ""
   GRD1.col = 0
   GRD1.row = intRow
   For j = 0 To GRD1.Cols - 1
      GRD1.col = j
      GRD1.CellBackColor = QBColor(15)
   Next j
   Call SetColor(CDbl(intRow))
   Call ClearText
End Sub

Private Sub cmdDetail_Click()
   cmdState = 99
   Call PubShowNextData
End Sub

Public Function PubShowNextData() As Boolean
Dim rsA As New ADODB.Recordset
Dim stFileName As String
Dim hLocalFile As Long

Select Case cmdState
'Case 0 '基本資料
'   If bolQuery = True Then
'      Me.Enabled = False
''      For i = 1 To GrdDataList.Rows - 1
''         GrdDataList.col = 0
''         GrdDataList.row = i
''         If Trim(GrdDataList.Text) = "V" Then
'           Dim Str01 As String
''           GrdDataList.col = 0
''           GrdDataList.Text = ""
''           For j = 0 To GrdDataList.Cols - 1
''               GrdDataList.col = j
''               GrdDataList.CellBackColor = QBColor(15)
''           Next j
''           GrdDataList.col = 1
'           Str01 = SystemNumber(lblCaseNo, 1)
'           If Mid(UCase(Str01), 1, 1) = "N" Then
'               Str01 = Mid(Str01, 2, 3)
'           End If
''           If Not IsNull(GrdDataList.Text) Then
'               'Modified by Morgan 2016/3/24 排除母層是共同查詢
'               If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
'                  fnCloseAllFrm100 'Added by Morgan 2016/2/22
'               End If
'               'end 2016/3/24
'               If fnSaveParentForm(Me) = False Then
'                   Me.Enabled = True
'                   Exit Sub
'               End If
'               bolQuery = False
'               Select Case Pub_RplStr(Str01)
'                   Case "CFP", "FCP", "P"   '專利
'                         Screen.MousePointer = vbHourglass
'                         frm100101_3.Show
'                         frm100101_3.Tag = Pub_RplStr(lblCaseNo)
'                         frm100101_3.StrMenu
'                         Screen.MousePointer = vbDefault
'                   Case "CFT", "FCT", "T", "TF"   '商標
'                         Screen.MousePointer = vbHourglass
'                         frm100101_4.Show
'                         frm100101_4.Tag = Pub_RplStr(lblCaseNo)
'                         frm100101_4.StrMenu
'                         Screen.MousePointer = vbDefault
'                   'Modify By Sindy 2009/07/24 增加LIN系統類別
'                   Case "CFL", "FCL", "L", "LIN"          '法務
'                         Screen.MousePointer = vbHourglass
'                         frm100101_5.Show
'                         frm100101_5.Tag = Pub_RplStr(lblCaseNo)
'                         frm100101_5.StrMenu
'                         Screen.MousePointer = vbDefault
'                   Case "LA"            '顧問
'                         Screen.MousePointer = vbHourglass
'                         frm100101_6.Show
'                         frm100101_6.Tag = Pub_RplStr(lblCaseNo)
'                         frm100101_6.StrMenu
'                         Screen.MousePointer = vbDefault
'                   Case Else                  '服務
'                        Select Case Pub_RplStr(Str01)
'                            Case "TB"    '條碼
'                               Screen.MousePointer = vbHourglass
'                               frm100101_7.Show
'                               frm100101_7.Tag = Pub_RplStr(lblCaseNo)
'                               frm100101_7.StrMenu
'                               Screen.MousePointer = vbDefault
'                            Case "TM"
'                               Screen.MousePointer = vbHourglass
'                               frm100101_8.Show
'                               frm100101_8.Tag = Pub_RplStr(lblCaseNo)
'                               frm100101_8.StrMenu
'                               Screen.MousePointer = vbDefault
'                            Case "TD"
'                               Screen.MousePointer = vbHourglass
'                               frm100101_9.Show
'                               frm100101_9.Tag = Pub_RplStr(lblCaseNo)
'                               frm100101_9.StrMenu
'                               Screen.MousePointer = vbDefault
'                            Case "TC", "CFC"
'                               Screen.MousePointer = vbHourglass
'                               frm100101_A.Show
'                               frm100101_A.Tag = Pub_RplStr(lblCaseNo)
'                               frm100101_A.StrMenu
'                               Screen.MousePointer = vbDefault
'                            Case Else
'                               Screen.MousePointer = vbHourglass
'                               frm100101_B.Show
'                               frm100101_B.Tag = Pub_RplStr(lblCaseNo)
'                               frm100101_B.StrMenu
'                               Screen.MousePointer = vbDefault
'                         End Select
'               End Select
''           End If
'           Me.Enabled = True
'           Exit Sub
''         End If
''      Next i
'      Me.Enabled = True
'   End If
Case 1 '進度
   If bolQuery = True Then
      Me.Enabled = False
      For i = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = i
         If Trim(GRD1.Text) = "V" Then
'            GRD1.col = 0
'            GRD1.Text = ""
'            For j = 0 To GRD1.Cols - 1
'                GRD1.col = j
'                GRD1.CellBackColor = QBColor(15)
'            Next j
             GRD1.col = 4
             If GRD1.Text <> "" Then
'                'Modified by Morgan 2016/3/24 排除母層是共同查詢
'                If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
'                   fnCloseAllFrm100 'Added by Morgan 2016/2/22
'                End If
'                'end 2016/3/24

'                If fnSaveParentForm(Me) = False Then
'                    Me.Enabled = True
'                    Exit Function
'                End If
                Screen.MousePointer = vbHourglass
                bolQuery = False
                frm100101_2.Show
                frm100101_2.Tag = Pub_RplStr(GRD1.TextMatrix(dblPrevRow, 4))
                'frm100101_2.cmdOK(6).Visible = False
                frm100101_2.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Function
             Else
                MsgBox "無本所案號！", vbExclamation, "警告！"
             End If
         End If
      Next i
      Me.Enabled = True
   End If
'Case 4 '完整卷宗
'   Screen.MousePointer = vbHourglass
'   frm100101_L.m_strKey = lblCaseNo.Caption
'   frm100101_L.SetParent Me
'   If frm100101_L.QueryData = True Then
'      frm100101_L.Show
'      Me.Hide
'   Else
'      Unload frm100101_L
'   End If
'   Screen.MousePointer = vbDefault
Case Else
   PubShowNextData = False
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" And GRD1.TextMatrix(i, 16) <> "" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Function
         End If
         PubShowNextData = True
         '明細資料
         frm06010613_1.m_II01 = GRD1.TextMatrix(i, 10)
         frm06010613_1.m_II02 = GRD1.TextMatrix(i, 11)
         frm06010613_1.m_II03 = GRD1.TextMatrix(i, 16)
         frm06010613_1.m_II19 = GRD1.TextMatrix(i, 12)
         'Modify By Sindy 2017/12/25 Mark
'         Call CancelRowColor(i)
         frm06010613_1.CmdNext.Enabled = False
'         For j = i To grd1.Rows - 1
'            If grd1.TextMatrix(j, 0) = "V" And grd1.TextMatrix(j, 16) <> "" Then
'               frm06010613_1.cmdNext.Enabled = True
'               Exit For
'            End If
'         Next j
         Call frm06010613_1.SetParent(Me)
         frm06010613_1.Show
         frm06010613_1.QueryData
         Me.Hide
         Exit Function
      End If
   Next i
End Select
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Function QueryData(Optional bolShowMsg As Boolean = True) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strUser As String
   
On Error GoTo ErrHand
   
   m_blnColOrderAsc = True
   QueryData = False
   Screen.MousePointer = vbHourglass
   
   Call ClearText
   SetButton True 'Added by Morgan 2020/8/13
   Frame2.Visible = False: cmdSelCp09.Visible = False: Frame4.Enabled = False
   Label9.Visible = False
   Combo2.Visible = False
   Label6.Visible = False
   Frame3.Visible = False '同意/退回
   cmdDelRow.Enabled = False '刪除
   cmdDelRow.Tag = cmdDelRow.Enabled 'Add By Sindy 2021/1/22 記錄原狀況
   GRD1.Clear
   Call SetGrd
   
   LblSec2Query.Visible = False
   cmdRecall.Visible = False '回覆確收
   
   '[不處理/2次確認信件]
   '專利處程序須第二次確認
   'Modified by Morgan 2020/7/17 +IR16=1 也要列出(職代2次確認來函期限)
   'Modified by Morgan 2020/8/20 + Srt (2次確認OK或退回的排前面)
   'Modified by Sindy 2022/3/1 + getmailbox(pi01,pi03) 信箱來源
   strSql = "select '' V,IR23 符號,GetInputRecordReply(ir01,ir02,ir03) 確,sqldatet(PI12)||' '||sqltime6(PI13) 收信日期時間,decode(pi18,null,'',PI18||'-'||PI19||'-'||PI20||'-'||PI21) 本所案號,PI17 主旨" & _
            ",'' 收受者,s2.st02||'-'||decode(ir16," & 信件處理狀態 & ",ir16) 處理人員" & _
            ",sqldatet(ir17)||' '||sqltime6(ir18) 處理日期時間" & _
            ",sqldatet(IR05)||' '||sqltime6(IR06) 讀取日期時間" & _
            ",IR01,IR02,PI15,IR04,PI06,PI14,PI03 檔名,Pi08,Pi09,ir11,ir12,pi12,pi05,IR16,IR21 總收文號,IR20 處理原因,ir24,ir19,'' Srt,getmailbox(pi01,pi03) 信箱來源" & _
            " From inputrecord,PatentInput,staff s1,staff s2" & _
            " where IR08=0 and IR16 in('1','2','4')" & _
            " and IR01=PI01(+) and IR02=PI02(+) and IR03=PI03(+)" & _
            " and ir13=s1.st01(+)" & _
            " and ir19=s2.st01(+)" & _
            " and ir22='" & Trim(Left(Combo1, 6)) & "'" & _
            " order by IR23 asc,ir11 asc,ir12 asc"
   If Check2.Value = 1 Then
      cmdCACK.Visible = False 'Added by Morgan 2021/3/31
      Check2.BackColor = &H80FFFF
      LblSec2Query.Visible = True
      Frame3.Visible = True '專利處程序須第二次確認:同意/退回
   Else
      'Removed by Morgan 2022/2/16 取消(要改為智權確收)
      'If strSrvDate(1) >= e化客戶啟用日 Then cmdCACK.Visible = True 'Added by Morgan 2021/3/31
      'end 2022/2/16
      '顯示[不處理/2次確認信件]的筆數
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If InStr(Check2.Caption, "(") > 0 Then Check2.Caption = Left(Check2.Caption, InStr(Check2.Caption, "(") - 1)
      If intI = 1 Then
         Check2.Caption = Check2.Caption & "(" & RsTemp.RecordCount & ")"
      Else
         Check2.Caption = Check2.Caption & "(0)"
      End If
      Check2.BackColor = &H8000000F
      
      Label9.Visible = True
      Combo2.Visible = True
      Label6.Visible = True
      
      Frame2.Visible = True: cmdSelCp09.Visible = True: Frame4.Enabled = True
      'P12.專利處程序人員
      If PUB_GetST03(Trim(Left(Combo1, 6))) = "P12" Then
         cmdInput.Visible = True '輸入
      Else
         cmdInput.Visible = False '輸入
      End If
      If Not (PUB_GetST03(Trim(Left(Combo1, 6))) = "P12" Or _
              InStr(m_strUserList, Trim(Left(Combo1, 6))) > 0) Then
         cmdDelRow.Enabled = True '刪除
         cmdDelRow.Tag = cmdDelRow.Enabled 'Add By Sindy 2021/1/22 記錄原狀況
         cmdPDF.Enabled = False: Frame4.Enabled = False: cmdSelCp09.Enabled = False: cmdOK(1).Enabled = False '歸卷
      End If
      'Modified by Morgan 2020/8/13 + or IR16='7' or IR16='8'
      'Modified by Morgan 2020/8/20 + Srt (2次確認OK或退回的排前面)
      'Modified by Sindy 2022/3/1 + getmailbox(pi01,pi03) 信箱來源
      strSql = "select '' V,IR23 符號,GetInputRecordReply(ir01,ir02,ir03) 確,sqldatet(PI12)||' '||sqltime6(PI13) 收信日期時間,decode(pi18,null,'',PI18||'-'||PI19||'-'||PI20||'-'||PI21) 本所案號,PI17 主旨" & _
               ",'' 收受者,decode(ir16,null,decode(ir15,'Y',decode(length(ir03),5,decode(substr(ir03,1,1),'P','Patent'),'IPDept'),s1.st02),s2.st02||decode(ir16," & 信件處理狀態 & ",ir16)) 轉寄者" & _
               ",decode(ir16,null,sqldatet(Pi08)||' '||sqltime6(Pi09),sqldatet(ir17)||' '||sqltime6(ir18)) 轉寄日期時間" & _
               ",sqldatet(IR05)||' '||sqltime6(IR06) 讀取日期時間" & _
               ",IR01,IR02,PI15,IR04,PI06,PI14,PI03 檔名,Pi08,Pi09,ir11,ir12,pi12,pi05,IR16,IR21 總收文號,IR20 處理原因,ir24,ir19,decode(IR16,'7',1,'8',1,2) Srt,getmailbox(pi01,pi03) 信箱來源" & _
               " From inputrecord,PatentInput,staff s1,staff s2" & _
               " where IR08=0 and (IR16 is null or IR16='3' or IR16='7' or IR16='8')" & _
               " and IR01=PI01 and IR02=PI02 and IR03=PI03" & _
               " and ir13=s1.st01(+)" & _
               " and ir19=s2.st01(+)" & _
               " and IR04='" & Trim(Left(Combo1, 6)) & "'"
      strSql = strSql & " order by Srt asc, ir11 desc,ir12 desc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblTotCnt.Caption = "總筆數: "
   'Add By Sindy 2019/6/10
   If Check2.Value = 1 Then
      If InStr(Check2.Caption, "(") > 0 Then Check2.Caption = Left(Check2.Caption, InStr(Check2.Caption, "(") - 1)
   End If
   '2019/6/10 END
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      LblTotCnt.Caption = "總筆數: " & rsTmp.RecordCount
      QueryData = True
      For i = 1 To GRD1.Rows - 1
         '有無轉寄給他人:*.有
         strSql = "SELECT ir04 FROM inputrecord" & _
                  " WHERE ir01=" & GRD1.TextMatrix(i, 10) & _
                    " and ir02=" & GRD1.TextMatrix(i, 11) & _
                    " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                    " and (ir11<>" & GRD1.TextMatrix(i, 19) & " or ir12<>" & GRD1.TextMatrix(i, 20) & ")" & _
                    " and (ir13='" & Trim(Left(Combo1, 6)) & "' or ir14='" & Trim(Left(Combo1, 6)) & "')" & _
                    " and instr(ir04,'確收')=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            GRD1.TextMatrix(i, 3) = "*" & GRD1.TextMatrix(i, 3)
         End If
         '解析收受者:
         '轉寄日期和轉寄時間相同
         If GRD1.TextMatrix(i, 17) = GRD1.TextMatrix(i, 19) And GRD1.TextMatrix(i, 18) = GRD1.TextMatrix(i, 20) Then
            'Modify By Sindy 2019/7/22 + IIf(Trim(GRD1.TextMatrix(i, 26)) = "Y", "[副]", "") &
            GRD1.TextMatrix(i, 6) = IIf(Trim(GRD1.TextMatrix(i, 26)) = "Y", "[副]", "") & PUB_ReadUserData(GRD1.TextMatrix(i, 14))
         Else
            strSql = "SELECT ir04,ir24 FROM inputrecord" & _
                     " WHERE ir01=" & GRD1.TextMatrix(i, 10) & _
                       " and ir02=" & GRD1.TextMatrix(i, 11) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                       " and ir11=" & GRD1.TextMatrix(i, 19) & _
                       " and ir12=" & GRD1.TextMatrix(i, 20)
            intI = 1
            strUser = ""
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               With RsTemp
                  RsTemp.MoveFirst
                  Do While RsTemp.EOF = False
                     'Modify By Sindy 2019/7/22 + IIf(Trim("" & RsTemp.Fields("ir24")) = "Y", "[副]", "") &
                     strUser = strUser & ";" & IIf(Trim("" & RsTemp.Fields("ir24")) = "Y", "[副]", "") & PUB_ReadUserData(RsTemp.Fields("ir04"))
                     RsTemp.MoveNext
                  Loop
               End With
            End If
            GRD1.TextMatrix(i, 6) = Mid(strUser, 2)
         End If
      Next i
      Call SetColor
   Else
      If bolShowMsg = True Then
         Screen.MousePointer = vbDefault
         ShowNoData
      End If
   End If
   rsTmp.Close
   
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   GRD1.Visible = True
   dblPrevRow = 0
   
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   MsgBox " 查詢失敗！" & vbCrLf & Err.Description & vbCrLf & vbCrLf & strSql
End Function

Private Sub cmdHistory_Click()
Dim nFrm As Form
   
   '檢查表單是否已開啟，若是，則關閉
   For Each nFrm In Forms
      If StrComp(nFrm.Name, "frm06010613", vbTextCompare) = 0 Then
         Unload frm06010613
      End If
   Next
   
   Call frm06010613.SetParent(Me)
   frm06010613.m_WorkType = 1 '轉寄 Add By Sindy 2017/12/12
   frm06010613.m_MailUsernum = Me.Combo1.Text
   frm06010613.Show
   Me.Hide
End Sub

Private Function CheckStatus(pIR01 As String, pIR02 As String, pIR03 As String, pIR04 As String) As Boolean
   'Modify By Sindy 2017/12/26
   'Modified by Morgan 2020/7/30 +ir08,ir19
   'strExc(0) = "select ir16 from inputRecord where ir01=" & pIR01 & " and ir02=" & pIR02 & " and ir03='" & pIR03 & "' and ir04='" & pIR04 & "'"
   strExc(0) = "select ir16,ir08,ir19 from inputRecord where ir01=" & pIR01 & " and ir03='" & pIR03 & "' and ir04='" & pIR04 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2020/7/30 考慮2次確認或退回等改判斷處理人員
      'If RsTemp("ir16") = "1" Then '1. 輸入
      If RsTemp("ir08") > 0 Or (RsTemp("ir16") <> "" And RsTemp("ir19") = strUserNum) Then
      'end 2020/7/30
         CheckStatus = True
      Else
         CheckStatus = False
      End If
   End If
End Function

Public Sub GoNext()
   With GRD1
      If Val(m_iRow) = Val(txtPI18.Tag) Then
'         '上刪除標記,高度設零
'         .row = m_iRow
         If CheckStatus(m_strIR01, m_strIR02, m_strIR03, m_strIR04) = True Then
            If m_strIR01 = .TextMatrix(m_iRow, 10) And _
               m_strIR03 = .TextMatrix(m_iRow, 16) And _
               m_strIR04 = .TextMatrix(m_iRow, 13) Then
               LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
               Call SaveInputRecord(m_iRow)
               Call CancelRowColor(m_iRow) '清除反白
               GRD1.RowHeight(m_iRow) = 0
               If Val(Replace(LblTotCnt.Caption, "總筆數:", "")) = 0 Then
                  Call QueryData(False)
               End If
            Else
               Call QueryData(False)
            End If
'            If .TextMatrix(.row, 0) = "V" Then
'            End If
'            .TextMatrix(.row, 0) = "X"
'            .RowHeight(.row) = 0
'            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
'            Call ReadFirstGrd1Text '查詢勾選的第一筆資料
         End If
      Else
         Call QueryData(False)
      End If
   End With
End Sub

'不處理
Private Sub cmdNotProDel_Click()
Dim bolHavdData As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
   
   bolHavdData = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         bolHavdData = True
         Exit For
      End If
   Next i
   If bolHavdData = False Then
      MsgBox "請至少勾選一筆不處理的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
'      'Add By Sindy 2019/6/11
'      If PUB_GetST03(strUserNum) <> "P12" Then '非專利處,要輸入原因
'         If Trim(TxtIR20) = "" Then
'            MsgBox "原因不可空白！", vbExclamation, "警告！"
'            TxtIR20.SetFocus
'            Exit Sub
'         End If
'      End If
'      '2019/6/11 END
      If MsgBox("確定信件不處理嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   '專利處程序須第二次確認
   'Modify By Sindy 2017/12/26
   'strIR22Emp = PUB_GetWorkDeputyEmp(Trim(Left(Combo1.Text, 6)))
   strIR22Emp = PUB_GetWorkDeputyEmp(strUserNum, False)
   '2017/12/26 END
   If strIR22Emp = "" Then
      MsgBox "無職代可重覆確認資料！", vbExclamation, "警告！"
      Exit Sub
   End If
   
   'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         
         '不處理
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         
         '專利處程序須第二次確認
'         If PUB_GetST03(strUserNum) = "P12" Then
            strExc(0) = "update InputRecord set " & _
                        " ir16='2',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                        ",ir22='" & strIR22Emp & "'"
            'Add By Sindy 2021/4/8
            If TxtIR20 <> m_TxtIR20 Then
               strExc(0) = strExc(0) & _
                           ",IR20='" & ChgSQL(TxtIR20) & "'"
            End If
            '2021/4/8 END
            strExc(0) = strExc(0) & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
            
            strExc(0) = "select cum02 from CaseUseMemo" & _
                        " where cum05='02'" & _
                          " and cum06=" & CNULL(strUserNum) & _
                          " and cum02='" & strIR22Emp & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                           " values('0','" & strIR22Emp & "','0','0','02')"
               cnnConnection.Execute strExc(0)
               Frame1.Visible = True '*****
            End If
            
'         '其他人員
'         Else
'            strExc(0) = "update InputRecord set " & _
'                        " ir16='2',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
'                        ",ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
'            'Add By Sindy 2021/4/8
'            If TxtIR20 <> m_TxtIR20 Then
'               strExc(0) = strExc(0) & _
'                           ",IR20='" & TxtIR20 & "'"
'            End If
'            '2021/4/8 END
'            strExc(0) = strExc(0) & _
'                        " where ir01=" & grd1.TextMatrix(i, 10) & _
'                          " and ir02=" & grd1.TextMatrix(i, 11) & _
'                          " and ir03='" & ChgSQL(grd1.TextMatrix(i, 16)) & "'" & _
'                          " and upper(ir04)=upper('" & ChgSQL(grd1.TextMatrix(i, 13)) & "')" & _
'                          " and ir08=0"
'            cnnConnection.Execute strExc(0)
'
'            Call SaveInputRecord(i, True)
'         End If
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 信件不處理註記失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2013/9/5
Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   bolQuery = True
   PubShowNextData
   Exit Sub
End Sub

'歸卷
Private Sub cmdPDF_Click()
Dim strUpdTime As String
Dim bolConn As Boolean
   
On Error GoTo ErrHand
   
   '專利處程序須第二次確認
   'Modify By Sindy 2017/12/26
   'strIR22Emp = PUB_GetWorkDeputyEmp(Trim(Left(Combo1.Text, 6)))
   strIR22Emp = PUB_GetWorkDeputyEmp(strUserNum, False)
   '2017/12/26 END
   If strIR22Emp = "" Then
      MsgBox "無職代可重覆確認資料！", vbExclamation, "警告！"
      Exit Sub
   End If
   
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            If dblPrevRow <> .row Then
               MsgBox "資料列選取有誤，請重新確認！"
               Exit Sub
            End If
            '檢查資料
            If txtPI18 = "" Or txtPI19 = "" Then
               MsgBox "請輸入本所案號！", vbExclamation, "警告！"
               If txtPI18 = "" Then
                  Me.txtPI18.SetFocus
               ElseIf txtPI19 = "" Then
                  Me.txtPI19.SetFocus
               End If
               Exit Sub
            Else
               If txtPI18 <> "P" And txtPI18 <> "PS" And _
                  txtPI18 <> "CFP" And txtPI18 <> "CPS" Then
                  MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
                  Me.txtPI18.SetFocus
                  Exit Sub
               End If
            End If
            If txtPI20 = "" Then txtPI20 = "0"
            If txtPI21 = "" Then txtPI21 = "00"
            If txtRecvNo = "" Then
               MsgBox "請選擇歸卷的總收文號！", vbExclamation, "警告！"
               Me.txtRecvNo.SetFocus
               Exit Sub
            End If
            
            strExc(0) = "select cp09,cp01 from caseprogress" & _
                        " where cp01='" & txtPI18 & "'" & _
                          " and cp02='" & txtPI19 & "'" & _
                          " and cp03='" & txtPI20 & "'" & _
                          " and cp04='" & txtPI21 & "'" & _
                          " and cp09='" & txtRecvNo & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               MsgBox "查無進度資料！", vbExclamation, "警告！"
               Me.txtRecvNo.SetFocus
               Exit Sub
            End If
            
            '歸卷
            cnnConnection.BeginTrans: bolConn = True
            strUpdTime = Right("000000" & ServerTime, 6)
            
            If Left(GRD1.TextMatrix(.row, 16), 1) = "P" Then '專利處
               strExc(0) = "update patentinput set " & _
                           "pi18='" & txtPI18 & "',pi19='" & txtPI19 & "'," & _
                           "pi20='" & txtPI20 & "',pi21='" & txtPI21 & "'" & _
                           " where pi01=" & GRD1.TextMatrix(.row, 10) & _
                           " and pi02=" & GRD1.TextMatrix(.row, 11) & _
                           " and pi03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'"
               cnnConnection.Execute strExc(0)
            Else
               MsgBox "有問題請洽電腦中心!!"
               Exit Sub
            End If
            .TextMatrix(.row, 4) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
            .TextMatrix(.row, 24) = txtRecvNo
            
            '專利處程序須第二次確認
'            If PUB_GetST03(strUserNum) = "P12" Then
               strExc(0) = "update InputRecord set ir21='" & txtRecvNo & "'" & _
                           ",ir16='4',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                           ",ir22='" & strIR22Emp & "'"
               'Add By Sindy 2018/1/11
               If TxtIR20 <> m_TxtIR20 Then
                  strExc(0) = strExc(0) & _
                              ",IR20='" & ChgSQL(TxtIR20) & "'"
               End If
               '2018/1/11 END
               strExc(0) = strExc(0) & _
                           " where ir01=" & GRD1.TextMatrix(.row, 10) & _
                             " and ir02=" & GRD1.TextMatrix(.row, 11) & _
                             " and ir03='" & ChgSQL(GRD1.TextMatrix(.row, 16)) & "'" & _
                             " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(.row, 13)) & "')" & _
                             " and ir08=0"
               cnnConnection.Execute strExc(0)
               
               strExc(0) = "select cum02 from CaseUseMemo" & _
                           " where cum05='02'" & _
                             " and cum06=" & CNULL(strUserNum) & _
                             " and cum02='" & strIR22Emp & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                              " values('0','" & strIR22Emp & "','0','0','02')"
                  cnnConnection.Execute strExc(0)
                  Frame1.Visible = True '*****
               End If
               
'            '其他人員
'            Else
'               Screen.MousePointer = vbHourglass
'               '下載信件檔,上傳卷宗區
'               If PUB_UploadPatentLetterFile(grd1.TextMatrix(.row, 10), grd1.TextMatrix(.row, 16), grd1.TextMatrix(.row, 24)) = False Then
'                  Screen.MousePointer = vbDefault
'                  Exit Sub
'               End If
'
'               strExc(0) = "update InputRecord set " & _
'                           " ir16='4',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
'                           ",ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
'               'Add By Sindy 2018/1/11
'               If TxtIR20 <> m_TxtIR20 Then
'                  strExc(0) = strExc(0) & _
'                              ",IR20='" & TxtIR20 & "'"
'               End If
'               '2018/1/11 END
'               strExc(0) = strExc(0) & _
'                           " where ir01=" & grd1.TextMatrix(.row, 10) & _
'                             " and ir02=" & grd1.TextMatrix(.row, 11) & _
'                             " and ir03='" & ChgSQL(grd1.TextMatrix(.row, 16)) & "'" & _
'                             " and upper(ir04)=upper('" & ChgSQL(grd1.TextMatrix(.row, 13)) & "')" & _
'                             " and ir08=0"
'               cnnConnection.Execute strExc(0)
'
'               Call SaveInputRecord(.row, True)
'            End If
            
            cnnConnection.CommitTrans: bolConn = False
            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
            Call CancelRowColor(.row) '清除反白
            GRD1.RowHeight(.row) = 0
            'Call ReadFirstGrd1Text '查詢勾選的第一筆資料
            Exit For
         End If
      Next m_iRow
   End With
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox " 信件歸卷註記失敗！" & vbCrLf & Err.Description
End Sub

'已處理
Private Sub cmdProDel_Click()
Dim bolHavdData As Boolean
Dim strUpdTime As String
Dim bolConn As Boolean
   
   bolHavdData = False
   If GRD1.Rows - 1 < 1 Then Exit Sub
'   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         bolHavdData = True
         Exit For
      End If
   Next i
   If bolHavdData = False Then
      MsgBox "請至少勾選一筆已處理的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      'Add By Sindy 2022/2/8 控管「已處理」備註要有才可執行
      If Trim(TxtIR20) = "" Then
         MsgBox "原因不可空白！", vbExclamation, "警告！"
         TxtIR20.SetFocus
         Exit Sub
      End If
      '2022/2/8 END
      
      If MsgBox("確定信件已處理？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         
         '5.已處理
         'Added by Morgan 2020/8/13
         '來函期限確認/退回只要上確認人員日期時間
         If stIR16 = "7" Or stIR16 = "8" Then
            'Added by Morgan 2020/9/16
            '更新報價日期，清除列印日期註記
            strSql = "update lettercache set lc13=null,lc11=to_char(sysdate,'yyyymmdd'),lc12=to_char(sysdate,'hh24miss') where lc03='" & Trim(GRD1.TextMatrix(i, 24)) & "' and lc13=19221111"
            cnnConnection.Execute strSql, intI
            '來函無報價且承辦人為程序(輸入人員)的要補上發文日
            If intI = 0 Then
               strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & Trim(GRD1.TextMatrix(i, 24)) & "' and cp14=cp65 and cp27 is null"
               cnnConnection.Execute strSql, intI
            End If
            'end 2020/9/16
            
            'Added by Morgan 2023/4/10 FMP案通知來函的EMail於確認後寄出
            strSql = "update mailcache set mc01='" & strUserNum & "',mc12=0 where mc13='" & Trim(GRD1.TextMatrix(i, 24)) & "' and mc12=99999999"
            cnnConnection.Execute strSql, intI
            'end 2023/4/10
            
            'Added by Morgan 2021/2/26 清除來函法定期限(暫存)
            strSql = "update caseprogress set cp142='' where cp09='" & Trim(GRD1.TextMatrix(i, 24)) & "' and cp07 is null"
            cnnConnection.Execute strSql, intI
            'end 2021/2/26
               
            strExc(0) = "update InputRecord set ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                        " and ir02=" & GRD1.TextMatrix(i, 11) & _
                        " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                        " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                        " and ir08=0"
         Else
         'end 2020/8/13
            strExc(0) = "update InputRecord set " & _
                        " ir16='5',ir17=" & strSrvDate(1) & ",ir18=" & strUpdTime & ",ir19='" & strUserNum & "'" & _
                        ",ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'"
            If TxtIR20 <> m_TxtIR20 Then
               strExc(0) = strExc(0) & _
                           ",IR20='" & ChgSQL(TxtIR20) & "'"
            End If
            strExc(0) = strExc(0) & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                        " and ir02=" & GRD1.TextMatrix(i, 11) & _
                        " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                        " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                        " and ir08=0"
         End If 'Added by Morgan 2020/8/13
         
         cnnConnection.Execute strExc(0)
         
         Call SaveInputRecord(i, False)
         
         cnnConnection.CommitTrans: bolConn = False
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         Call CancelRowColor(i) '清除反白
         GRD1.RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   PUB_SendMailCache 'Added by Morgan 2023/4/10
   'Call QueryData(False)
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 信件已處理註記失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdQuery_Click()
   If Me.Combo1.Text = "" Then
      MsgBox "員工編號不可空白！", vbExclamation, "警告！"
      Combo1.SetFocus
      Exit Sub
   End If
   
   Call QueryData
End Sub

'回覆確收
Private Sub cmdRecall_Click()
Dim strFileName As String, strFullFileName As String
Dim strUpdTime As String
'Dim objOutLook As Object
'Dim objMail As Object
'Dim myForward As Object
'Dim objNS As Object
'Dim strSocSubject As String
'Dim jj As Integer
Dim bolConn As Boolean
Dim strIR20 As String
Dim strCnt As String
   
On Error GoTo ErrHand

'   Screen.MousePointer = vbHourglass
'   Set objOutLook = CreateObject("Outlook.Application")
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If GRD1.TextMatrix(i, 2) = "Y" Then
            If MsgBox("要【重覆確收】嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Exit Sub
            End If
         End If
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = Mid(GRD1.TextMatrix(i, 15), InStrRev(GRD1.TextMatrix(i, 15), "/") + 1)
         If GetAttachFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), GRD1.TextMatrix(i, 16), strFullFileName, m_AttachPath & "\" & strFileName) = True Then
            ShellExecute 0, "open", strFullFileName, vbNullString, vbNullString, 1
'            Set objMail = objOutLook.CreateItemFromTemplate(strFullFileName)
'
'''Dim myItem As Outlook.MailItem
''
'' 'Dim myAction As Outlook.Action
''
'' Set myItem = objOutLook.CreateItem(olMailItem)
''
'' Set myAction = myItem.Actions.add
''
''
''
'' myAction.Name = "Link Original"
''
'' myAction.ShowOn = olMenuAndToolbar
''
'' myAction.ReplyStyle = olLinkOriginalItem
''
'' myItem.To = "Dan Wilson"
''
'' myItem.Display
'' myItem.Send
''
''
''' Dim myItem As Outlook.MailItem
''
''' Dim myAction As Outlook.Action
''
''
''
'' 'Set myItem = objOutLook.CreateItem(strFullFileName)
''
'' Set myAction = objMail.Actions.add
''
'' myAction.Name = "Agree"
''
'' objMail.To = objOutLook.GetNamespace("MAPI").CurrentUser
'' objMail.Display
'' objMail.Send
''
''
''            Set objMail = objOutLook.CreateItem(strFullFileName)
'            strSocSubject = objMail.Subject
'            Text2.Text = objMail.Subject
'            'Reply-To
''            PUB_SendMail strUserNum, "97038", "", "回覆: " & Text2, objMail.SenderEmailAddress & vbCrLf & "已收悉~" & vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , , , , , , "TM@taie.com.tw", , , True, False
''            PUB_SendMail strUserNum, "sindygirllu@gmail.com", "", "回覆: " & Text2, objMail.SenderEmailAddress & vbCrLf & "已收悉~" & vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , , , , , , "TM@taie.com.tw", , , True, False
'
'            'Set objOL = New Outlook.Application
''Set objNS = objOutLook.GetNamespace("MAPI")
''Set objFolder = objNS.GetDefaultFolder(olFolderContacts)
'            If objMail.Class = 43 Then '43.olMail
'                  '*** 轉寄 *** 會用inbound名義寄出
'                  Set myForward = objMail.Forward
'                  'Set myForward = objMail.ReplyAll
'                  'objMail.Reply
''                  strCC = "" '副本
''                  If strII05 = "6" Then
''                     '新知不轉職代
''                  Else
''                     '檢查收件者是否有休假,若有,則加發職代
''                     ArrStr = Split(strTo, ";")
''                     For jj = 0 To UBound(ArrStr)
''                        strTempCC = GetCaseDutyAgent(ArrStr(jj), "", False)
''                        If strTempCC <> "" Then
''                           ArrStrkk = Split(strTempCC, ";")
''                           For kk = 0 To UBound(ArrStrkk)
''                              If InStr(strTo, ArrStrkk(kk)) = 0 Then '收件者
''                                 If InStr(strCC, ArrStrkk(kk)) = 0 Then '副本
''                                    If strCC = "" Then
''                                       strCC = ArrStrkk(kk)
''                                    Else
''                                       strCC = strCC & ";" & ArrStrkk(kk)
''                                    End If
''                                 End If
''                              End If
''                           Next kk
''                        End If
''                     Next jj
''                  End If
'                  '移除原信的收件人及副本;密件副本不會留在msg中
'                  For jj = myForward.Recipients.Count To 1 Step -1
'                     myForward.Recipients.Remove jj
'                  Next jj
''                  If InStr(UCase(PUB_GetDbTerminal), 正式資料庫電腦名稱) = 0 Then '測試資料庫
''                     strTo = Pub_GetSpecMan("電腦中心郵件檢核人員")
''                     strCC = ""
''                  End If
'                  '收件者
'                  myForward.Recipients.add objMail.SenderEmailAddress
''                  ArrStr = Split(strTo, ";")
''                  For kk = 0 To UBound(ArrStr)
''                     strExc(0) = "select st01,st04 from staff" & _
''                                 " where st01='" & ArrStr(kk) & "'"
''                     intI = 1
''                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''                     '台一人員
''                     If intI = 1 Then
''                        If RsTemp.Fields("st04") = "1" Then '在職的才能寄出(outlook無此聯絡人)時會出現錯誤訊息-2147467259:Outlook 無法識別一或多個名稱。
''                           myForward.Recipients.add ArrStr(kk)
''                        End If
''                     '特殊信箱
''                     Else
''                        If UCase(ArrStr(kk)) <> UCase("patent") Then '外專不寄走系統
''                           If UCase(ArrStr(kk)) = UCase("account") Then ArrStr(kk) = "account@taie.com.tw"
''                           myForward.Recipients.add ArrStr(kk)
''                        End If
''                     End If
''                  Next kk
''                  '副本
''                  myForward.cc = strCC
''                  '主旨增加,當個案且有案號時,顯示歸入那一個案號
''                  If strII05 = "1" And strCaseNo <> "" Then
''                     myForward.Subject = myForward.Subject & "【" & strCaseNo & " Saved】"
''                  End If
''                  'myForward.senderemailaddress = "ipdept@taie.com.tw"
''                  'myForward.sentonbehalfofname = "ipdept"
'
'                  myForward.Subject = "【已確收】" & myForward.Subject
'                  'myForward.htmlbody = "增加內文" & vbCrLf & "增加內文" & myForward.htmlbody
'                  'myForward.Body = "增加內文2" & vbCrLf & vbCrLf & "增加內文2" & vbCrLf & myForward.Body
'
'                  myForward.Display
'                  'myForward.Send
'                  'DoEvents
'                  Set myForward = Nothing
'                  '*** END
'               '2017/6/26 END
'               End If
'
'            Set objMail = Nothing
'         Else
'            MsgBox "無此郵件！", vbInformation
'         End If
         
            cnnConnection.BeginTrans: bolConn = True
            strUpdTime = Right("000000" & ServerTime, 6)
            
            If TxtIR20 <> m_TxtIR20 Then
               strIR20 = TxtIR20
            End If
            
            '檢查目前此封郵件已確收幾次
            strExc(0) = "select count(ir01) from inputrecord" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and instr(ir04,'確收')>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            strCnt = ""
            If intI = 1 Then
               strCnt = CStr(RsTemp.Fields(0) + 1)
            End If
            '確收
            strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR14" & _
                        ",IR08,IR09,IR10,IR20)" & _
                        " values(" & GRD1.TextMatrix(i, 10) & _
                                 "," & GRD1.TextMatrix(i, 11) & _
                                 ",'" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                                 ",'確收" & strCnt & "'," & strSrvDate(1) & "," & _
                                 strUpdTime & ",'" & strUserNum & "'," & CNULL(IIf(Trim(Left(Combo1, 6)) <> strUserNum, Trim(Left(Combo1, 6)), "")) & _
                                 "," & strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "','" & ChgSQL(strIR20) & "')"
            cnnConnection.Execute strExc(0)
            
            cnnConnection.CommitTrans: bolConn = False
            GRD1.TextMatrix(i, 2) = "Y" '確收註記
            'Call CancelRowColor(i) '清除反白
         End If
      End If
   Next i
   Screen.MousePointer = vbDefault
   
'   Set objMail = Nothing
'   Set objOutLook = Nothing
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox " 確收失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdReInput_Click()
   With GRD1
   For m_iRow = 1 To .Rows - 1
      .row = m_iRow
      txtPI18.Tag = m_iRow
      If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
         If dblPrevRow <> .row Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         m_strIR01 = .TextMatrix(.row, 10)
         m_strIR02 = .TextMatrix(.row, 11)
         m_strIR03 = .TextMatrix(.row, 16)
         m_strIR04 = .TextMatrix(.row, 13)
         Call Forms(0).SetTmpfrm04010519(Me) 'Add By Sindy 2022/5/23
         frm02010605.m_CP09 = Trim(.TextMatrix(.row, 24))
         'Added by Morgan 2023/4/12
         frm02010605.m_strIR01 = m_strIR01
         frm02010605.m_strIR02 = m_strIR02
         frm02010605.m_strIR03 = m_strIR03
         frm02010605.m_strIR04 = m_strIR04
         'end 2023/4/12
         frm02010605.Show
         Exit For
      End If
   Next
   End With
End Sub

Private Sub cmdSelCp09_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
    
   If Trim(txtPI18) <> "" And Trim(txtPI19) <> "" Then
      Me.Tag = ""
      txtPI20.Text = IIf(txtPI20 = "", "0", txtPI20)
      txtPI21.Text = IIf(txtPI21 = "", "00", txtPI21)
      'cp159=0
      sqlB = "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(pa09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,patent " & _
             "where cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "' " & _
             "and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 "
      'Add By Sindy 2019/4/30 + 服務
      sqlB = sqlB & " union " & _
             "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(sp09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,servicepractice " & _
             "where cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "' " & _
             "and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 "
      sqlB = sqlB & " ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      intI = 0
      Set rsRead = ClsLawReadRstMsg(intI, sqlB)
      If intI = 1 Then
         Set frm880012.grdDataList.Recordset = rsRead
         Set frm880012.fmParent = Me
         frm880012.iTyp = "1"
         frm880012.Show vbModal
         If Me.Tag <> "" Then
            txtRecvNo.Text = Me.Tag
            txtRecvNo.SetFocus
         Else
            txtRecvNo.Text = ""
         End If
      End If
   Else
      MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
      If Me.txtPI18.Enabled = True Then Me.txtPI18.SetFocus
   End If
End Sub

'立即寄送通知信
Private Sub cmdSendMail_Click()
Dim strContent As String
Dim bolCaseDutyAgentMsg As Boolean, strRestKind As String
Dim strTempCC As String
Dim strRestEmp As String, strNormalEmp As String
Dim ArrStr As Variant, jj As Integer
   
   strExc(0) = "select cum02 from CaseUseMemo" & _
               " where cum05='02'" & _
                 " and cum06=" & CNULL(strUserNum)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         '因為有休假問題,所以有休假人員各自發信,其他人則一封
         strTempCC = GetCaseDutyAgent(RsTemp.Fields("cum02"), "", bolCaseDutyAgentMsg, strRestKind)
         If strTempCC <> "" Then
            strRestEmp = strRestEmp & ";" & RsTemp.Fields("cum02")
         Else
            strNormalEmp = strNormalEmp & ";" & RsTemp.Fields("cum02")
         End If
         RsTemp.MoveNext
      Loop
      strContent = "請至案件管理系統的一般作業\系統收件區，進行看查。"
      '薛經理:一起通知,減少操作人員等發信的時間
      If strNormalEmp <> "" Then
         strNormalEmp = Mid(strNormalEmp, 2)
         PUB_SendMail strUserNum, strNormalEmp, "", "通知已有信件轉入系統收件區", strContent, , , , , , , , , , , False, , , , , , , , , , , , , "1"
      End If
      '有休假人員各自發信
      If strRestEmp <> "" Then
         strRestEmp = Mid(strRestEmp, 2)
         ArrStr = Split(strRestEmp, ";")
         For jj = 0 To UBound(ArrStr)
            '暫不用轉寄職代,因實際做事的人不同
            PUB_SendMail strUserNum, ArrStr(jj), "", "通知已有信件轉入系統收件區", strContent, , , , , , , , , , True, False, , , , , , , , , , , , , "1"
         Next jj
      End If
      '刪除記錄
      strExc(0) = "delete from CaseUseMemo" & _
                  " where cum05='02'" & _
                  " and cum06=" & CNULL(strUserNum)
      cnnConnection.Execute strExc(0)
   End If
   Frame1.Visible = False '*****
End Sub

'專利處程序轉寄
Private Sub PatentTransMail()
Dim strUpdTime As String
Dim strFileName As String, strFullFileName As String
Dim bolConn As Boolean
Dim strTo As String 'Add By Sindy 2022/1/27
   
On Error GoTo ErrHand
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         For j = 0 To List1.ListCount - 1
            '非專利處程序人員,須轉寄Outlook先下載信件檔
            If PUB_GetST03(Left(Trim(List1.List(j)), 5)) <> "P12" Then
               '讀取檔案
               strFileName = Mid(GRD1.TextMatrix(i, 15), InStrRev(GRD1.TextMatrix(i, 15), "/") + 1)
               If GetAttachFile(GRD1.TextMatrix(i, 10), GRD1.TextMatrix(i, 11), GRD1.TextMatrix(i, 16), strFullFileName, m_AttachPath & "\" & strFileName) = False Then
                  MsgBox "下載檔案失敗，無法轉寄！", vbExclamation, "警告！"
                  Exit Sub
               Else
                  Exit For
               End If
            End If
         Next j
         
         'Add by Sindy 2021/11/19 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         Screen.MousePointer = vbHourglass
         cnnConnection.BeginTrans: bolConn = True
         strUpdTime = Right("000000" & ServerTime, 6)
         '清除主檔msg檔可刪除日期-轉寄就要恢復主檔的控管
         strExc(0) = "update PatentInput set " & _
                     " pi16=0" & _
                     " where pi01=" & GRD1.TextMatrix(i, 10) & _
                       " and pi02=" & GRD1.TextMatrix(i, 11) & _
                       " and pi03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "' and nvl(pi16,0)>0"
         cnnConnection.Execute strExc(0)
         
         strTo = "" 'Add By Sindy 2022/1/27
         '新增收受者
         For j = 0 To List1.ListCount - 1
            strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR14)" & _
                        " values(" & GRD1.TextMatrix(i, 10) & _
                                 "," & GRD1.TextMatrix(i, 11) & _
                                 ",'" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                                 ",'" & Left(Trim(List1.List(j)), 5) & "'," & strSrvDate(1) & "," & _
                                 strUpdTime & ",'" & strUserNum & "'," & CNULL(IIf(Trim(Left(Combo1, 6)) <> strUserNum, Trim(Left(Combo1, 6)), "")) & ")"
            cnnConnection.Execute strExc(0)
            '非專利處程序人員
            If PUB_GetST03(Left(Trim(List1.List(j)), 5)) <> "P12" Then
               'Add By Sindy 2022/1/27
               If strTo <> "" Then strTo = strTo & ";"
               strTo = strTo & Left(Trim(List1.List(j)), 5)
               '2022/1/27 END
'               '轉寄Outlook
'               PUB_SendMail strUserNum, Left(Trim(List1.List(j)), 5), "", GRD1.TextMatrix(i, 5), _
'                     IIf(TextContext.Enabled = True And Trim(TextContext) <> "", TextContext, vbCrLf & "信件內容參附件！"), , strFullFileName
'               If bolMailSendOk = False Then GoTo ErrHand
'               '並且該收受者上刪除日期時間人員
'               strExc(0) = "update InputRecord set " & _
'                           " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
'                           " where ir01=" & GRD1.TextMatrix(i, 10) & _
'                             " and ir02=" & GRD1.TextMatrix(i, 11) & _
'                             " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
'                             " and ir04='" & Left(Trim(List1.List(j)), 5) & "'" & _
'                             " and ir08=0"
'               cnnConnection.Execute strExc(0)
               'Add By Sindy 2018/1/17 轉寄Outlook後,
               '非專利處人員該筆結束
               If Left(PUB_GetST03(Left(Trim(List1.List(j)), 5)), 2) <> "P1" Then
                  Check1.Value = 1
               '專利處程序人員要針對此信件做結果處理
               Else
                  Check1.Value = 0
               End If
               '2018/1/17 END
            '進系統收件區人員
            Else
               '寫入要發通知信的人員
               'CaseUseMemo:
               'cum01 = 0
               'cum02 = 收受者
               'cum03 = 0
               'cum04 = 0
               'cum05 = 02.信件轉寄通知信
               'cum06 = 操作人員
               strExc(0) = "select cum02 from CaseUseMemo" & _
                           " where cum05='02'" & _
                             " and cum06=" & CNULL(strUserNum) & _
                             " and cum02=" & CNULL(Left(Trim(List1.List(j)), 5))
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                              " values('0','" & Left(Trim(List1.List(j)), 5) & "','0','0','02')"
                  cnnConnection.Execute strExc(0)
                  Frame1.Visible = True '*****
               End If
               '轉寄專利處程序人員後該筆結束,由另一位程序接手處理信件
               If PUB_GetST03(Left(Trim(List1.List(j)), 5)) = "P12" Then
                  Check1.Value = 1 'Add By Sindy 2017/12/21 轉寄程序要上刪除日期
               Else
                  Check1.Value = 0
               End If
            End If
         Next j
         'Add By Sindy 2022/1/27
         '秀玲寄-
         '洪副理 您好：原信是同時寄到IPDEPT及PATENT信箱，PATENT的部分是由林慧汶轉寄您及Monica、May三人，
         '程式原是分3 封信寄發，已請SINDY調整程式，改為一封信同時寄發三人，
         '這樣您就不會再有重覆收到正副本不同的2封信，也可以看出同時發給May。
         If strTo <> "" Then
            '轉寄Outlook
            PUB_SendMail strUserNum, strTo, "", GRD1.TextMatrix(i, 5), _
                  IIf(TextContext.Enabled = True And Trim(TextContext) <> "", TextContext, vbCrLf & "信件內容參附件！"), , strFullFileName
            If bolMailSendOk = False Then GoTo ErrHand
            '並且該收受者上刪除日期時間人員
            strExc(0) = "update InputRecord set " & _
                        " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and ir04 in('" & Replace(strTo, ";", "','") & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
         End If
         '2022/1/27 END
         
         If Check1.Value = 1 Then '上刪除日期
            strExc(0) = "update InputRecord set " & _
                        " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                        " where ir01=" & GRD1.TextMatrix(i, 10) & _
                          " and ir02=" & GRD1.TextMatrix(i, 11) & _
                          " and ir03='" & ChgSQL(GRD1.TextMatrix(i, 16)) & "'" & _
                          " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(i, 13)) & "')" & _
                          " and ir08=0"
            cnnConnection.Execute strExc(0)
         End If
         Call SaveInputRecord(i, False)
         cnnConnection.CommitTrans: bolConn = False
         If Left(GRD1.TextMatrix(i, 3), 1) <> "*" Then
            GRD1.TextMatrix(i, 3) = "*" & GRD1.TextMatrix(i, 3)
         End If
         If Check1.Value = 1 Then '上刪除日期
            LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
            Call CancelRowColor(i) '清除反白
            GRD1.RowHeight(i) = 0
         End If
      End If
   Next i
   Screen.MousePointer = vbDefault
   
   '清除收受者
   cboII06.Text = ""
   List1.Clear
   List1.Tag = ""
   
   Exit Sub
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox " 轉寄失敗！" & vbCrLf & Err.Description & vbCrLf & vbCrLf & strExc(0)
End Sub

Private Sub cmdUpdRow_Click()
Dim bolHavdSel As Boolean
   
   '收受者不可空白
   If List1.ListCount <= 0 Then
      MsgBox "收受者不可空白！", vbExclamation, "警告！"
      cboII06.SetFocus
      Exit Sub
   End If
   
   bolHavdSel = False
   '先檢查是否有資料要刪除
   If GRD1.Rows - 1 < 1 Then Exit Sub
   If GRD1.Rows - 1 >= 1 And GRD1.TextMatrix(1, 16) = "" Then Exit Sub
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If dblPrevRow <> i Then
            MsgBox "資料列選取有誤，請重新確認！"
            Exit Sub
         End If
         bolHavdSel = True
         Exit For
      End If
   Next i
   If bolHavdSel = False Then
      MsgBox "請至少勾選一筆要轉寄的資料！", vbExclamation, "警告！"
      Exit Sub
   Else
      If MsgBox("確定要轉寄信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         For i = 1 To grd1.Rows - 1
'            If grd1.TextMatrix(i, 0) = "V" Then
'               Call CancelRowColor(i) '清除反白
'            End If
'         Next i
         Exit Sub
      End If
   End If
   
   '專利處程序轉寄
   Call PatentTransMail
End Sub

Private Sub Combo1_Click()
   If Combo1.Text <> "" Then
      Call QueryData(False)
   End If
End Sub

'加註符號
Private Sub Combo2_Click()
   If dblPrevRow > 0 Then
      If GRD1.TextMatrix(dblPrevRow, 0) = "V" Then
         strExc(0) = "update InputRecord set ir23=" & CNULL(Left(Combo2.Text, 1)) & _
                     " where ir01=" & GRD1.TextMatrix(dblPrevRow, 10) & _
                       " and ir03='" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 16)) & "'" & _
                       " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 13)) & "')"
         cnnConnection.Execute strExc(0)
         GRD1.TextMatrix(dblPrevRow, 1) = Left(Combo2.Text, 1)
      End If
   End If
End Sub

Private Sub Command1_Click()
   If cboII06.Text <> "" Then
      If InStr(List1.Tag, cboII06.Text) = 0 Then
         If List1.Tag = "" Then List1.Clear
         If CheckDataValid(cboII06.Text) = False Then GRD1.Visible = True: Exit Sub
         List1.AddItem cboII06.Text
         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.Text
      End If
      cboII06.Text = ""
   End If
End Sub

'Private Sub Command2_Click()
'   Call List1_DblClick
'End Sub

Private Sub Form_Activate()
   'Add By Sindy 2017/12/20 內專程序人員請使用「專利管理系統（Patpro）」操作系統收件區
   If Pub_StrUserSt03 = "P12" And (UCase(App.EXEName) <> UCase("Tepatpro") And UCase(App.EXEName) <> UCase("patpro")) Then
      MsgBox "內專程序人員請使用「專利管理系統（Patpro）」" & vbCrLf & vbCrLf & "操作系統收件區！", vbExclamation
      Unload Me
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ClearText
   
   'If Me.Combo1.Text = "" Then Me.Combo1.Text = strUserNum & " " & strUserName
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   '檢查是否有未寄送通知信
   strExc(0) = "select cum02 from CaseUseMemo" & _
               " where cum05='02'" & _
                 " and cum06=" & CNULL(strUserNum) & _
                 " and rownum<=1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Frame1.Visible = True '*****
   Else
      Frame1.Visible = False
   End If
   '收受者
   cboII06.Clear
   cboII06.AddItem "": m_strUserList = ""
   
   '轉信收受者預設部門前二碼相同者,
   '但與操作人員相同部門者排在前面,
   '其他人以部門 所別 + 員工編號排序
   strSql = "SELECT a0902,st01,st02,st03,st06,1 sort FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st01 not in('96029','96030') and st03='" & Pub_StrUserSt03 & "'" & _
            " Union" & _
            " SELECT a0902,st01,st02,st03,st06,2 sort FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st01 not in('96029','96030') and substr(st03,1,2)='" & Left(Pub_StrUserSt03, 2) & "' and st03<>'" & Pub_StrUserSt03 & "'" & _
            " order by sort,st03,st06,st01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            cboII06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   
'   'CFP程序人員不可執行歸卷
'   If PUB_GetST05(strUserNum) = "83" Or PUB_GetST05(strUserNum) = "85" Then
'      cmdPDF.Enabled = False
'   End If
   
   '設定同意/退回的位置
   Frame3.Left = 7140
   Frame3.Top = 2070
   Check2.Caption = "不處理/2次確認信件"
   
   Call SetCombo1
   Call QueryData(False)
   
   'Added by Sindy 2021/11/19 如果一開始將ListBox拉到需要的大小，字型會自動放大；
   '所以畫面預設為一列高度(315)，Form_Load才放大到需要的大小
   List1.Clear
   List1.Height = 600
   List1.Width = 1500
End Sub

Private Sub SetCombo1()
   Combo1.Clear
   Combo1.AddItem strUserNum & " " & strUserName
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
   Combo1.Text = Combo1.List(0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modify By Sindy 2018/6/20 Run執行檔.exe才發E-Mail; 或測試時要執行
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Or _
      UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
   '2018/6/20 END
      '立即寄送通知信
      If Frame1.Visible = True Then Call cmdSendMail_Click
   End If
   
   DestroyToolTip '清除物件
   Set frm04010519 = Nothing
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   'Modify By Sindy 2019/7/22 + ,ir24
   '                        0    1     2     3               4           5       6         7         8               9               10      11      12      13      14      15      16      17      18      19      20      21      22      23      24          25          26      27      28     29
   arrGridHeadText = Array("V", "符", "確", "收信日期時間", "本所案號", "主旨", "收受者", "轉寄者", "轉寄日期時間", "讀取日期時間", "IR01", "IR02", "PI15", "IR04", "PI06", "PI14", "檔名", "Pi08", "Pi09", "ir11", "ir12", "pi12", "pi05", "IR16", "總收文號", "處理原因", "ir24", "ir19", "Srt", "信箱來源")
   arrGridHeadWidth = Array(200, 250, 0, 1300, 1200, 2800, 900, 1200, 1200, 1200, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1000, 2000, 0, 0, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next iRow
   If GRD1.RowHeight(1) = 0 Then GRD1.RowHeight(1) = 255 '***
   GRD1.Visible = True
End Sub

Private Sub Grd1_Click()
GRD1.Visible = False
GRD1.row = GRD1.MouseRow
GRD1.col = GRD1.MouseCol
nRow = GRD1.row
nCol = GRD1.col
If nRow = 0 Then
   If GRD1.Text <> "V" Then
      If GRD1.Text = "無" Then
         If m_blnColOrderAsc = True Then
            GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
Else
   'Modify By Sindy 2017/12/22
   If dblPrevRow <= 0 Then
      dblPrevRow = 0
   Else
      'Modify By Sindy 2017/12/29 記錄的目前資料列是未選取狀況,尋找目前反白的資料列,清除反白
      GRD1.col = 3
      GRD1.row = dblPrevRow
      If GRD1.CellBackColor <> &HFFC0C0 Then
         For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            If GRD1.CellBackColor = &HFFC0C0 Then
               Call CancelRowColor(GRD1.row) '清除反白
               dblPrevRow = 0
               Exit For
            End If
         Next i
      '2017/12/29 END
      ElseIf dblPrevRow <> nRow Then
         GRD1.TextMatrix(dblPrevRow, 0) = ""
         Call CancelRowColor(CInt(dblPrevRow)) '清除反白
      End If
   End If
   '2017/12/22 END
   GRD1.row = nRow 'GRD1.MouseRow
   dblPrevRow = GRD1.row '記錄目前筆數
   GRD1.col = 0
   'Add By Sindy 2021/1/22 收受者前面加[副],已無電子檔時,才要亮刪除按鍵
   If GRD1.TextMatrix(GRD1.row, 6) <> "" And Trim(GRD1.TextMatrix(GRD1.row, 15)) = "" Then
      If Left(GRD1.TextMatrix(GRD1.row, 6), 3) = "[副]" Then '副本目前均為主管,查看的信件,是從別的信箱轉入的信件
         cmdDelRow.Enabled = True
      Else
         cmdDelRow.Enabled = cmdDelRow.Tag
      End If
   End If
   '2021/1/22 END
   If GRD1.TextMatrix(GRD1.row, 16) <> "" Then
'      '清除反白
'      'If GRD1.TextMatrix(GRD1.row, 0) = "V" Then
'      If grd1.CellBackColor = &HFFC0C0 Then
'         Call CancelRowColor(grd1.row) '清除反白
'         If txtPI18.Tag <> "" And (Val(grd1.row) <= Val(txtPI18.Tag)) Then Call ReadFirstGrd1Text '查詢勾選的第一筆資料
'      Else
''         If Trim(GRD1.TextMatrix(GRD1.row, 9)) = "" And nCol = 0 Then '無讀取日期不可以操作其功能
''            GRD1.Visible = True
''            MsgBox "此郵件尚未讀取(開啟)不可操作其功能，因此不可勾選!!", vbExclamation, "警告！"
''         Else
            '將點選資料列反白
            GRD1.TextMatrix(GRD1.row, 0) = "V"
            GRD1.col = 0
            GRD1.row = nRow
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
            If List1.ListCount > 0 And CheckDataValid() = False Then
               Call CancelRowColor(GRD1.row) '清除反白
            End If
            SetColor2 'Added by Morgan 2020/7/29
            
            'If txtPI18.Tag = "" Or (Val(grd1.row) <> Val(txtPI18.Tag)) Then Call ReadFirstGrd1Text '查詢勾選的第一筆資料
            Call ReadFirstGrd1Text '查詢勾選的資料
''         End If
'      End If
   End If
End If
GRD1.Visible = True
End Sub

'Modify By Sindy 2017/12/27
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And (GRD1.MouseCol = 3 Or GRD1.MouseCol = 5) Then
      If iRow <> GRD1.MouseRow Or iCol <> GRD1.MouseCol Then
         If GRD1.MouseCol = 5 Then
            If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
               'GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
               CreateToolTip GetHWndForToolTip(GRD1), GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
               iRow = GRD1.MouseRow
               iCol = GRD1.MouseCol
            End If
         ElseIf GRD1.MouseCol = 3 Then
            'GRD1.ToolTipText = "信件編號:" & GRD1.TextMatrix(GRD1.MouseRow, 10) & "-" & GRD1.TextMatrix(GRD1.MouseRow, 16)
            CreateToolTip GetHWndForToolTip(GRD1), "信件編號:" & GRD1.TextMatrix(GRD1.MouseRow, 10) & "-" & GRD1.TextMatrix(GRD1.MouseRow, 16)
            iRow = GRD1.MouseRow
            iCol = GRD1.MouseCol
         End If
      End If
   End If
End Sub

'開啟附件
Private Sub GRD1_DblClick()
Dim strFileName As String, strFullFileName As String
Dim strUpdTime As String
Dim bolConn As Boolean
   
On Error GoTo ErrHand
   
   GRD1.row = GRD1.MouseRow
   GRD1.col = GRD1.MouseCol
   nRow = GRD1.row
   nCol = GRD1.col
   If GRD1.col = 5 And nRow > 0 Then
      If GRD1.TextMatrix(dblPrevRow, 16) <> "" Then
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = Mid(GRD1.TextMatrix(dblPrevRow, 15), InStrRev(GRD1.TextMatrix(dblPrevRow, 15), "/") + 1)
         Call PUB_ChkFileTypeOpenExE(strFileName) 'Add By Sindy 2017/9/13
         If GetAttachFile(GRD1.TextMatrix(dblPrevRow, 10), GRD1.TextMatrix(dblPrevRow, 11), GRD1.TextMatrix(dblPrevRow, 16), strFullFileName, m_AttachPath & "\" & strFileName) = True Then
            ShellExecute 0, "open", strFullFileName, vbNullString, vbNullString, 1
            
            '非電腦中心 (或電腦中心人員操作時,若收受者是自己信件時,可以更新讀取日期時間), 才需更新資料
            If Pub_StrUserSt03 <> "M51" Or _
               (Pub_StrUserSt03 = "M51" And Trim(Left(Combo1, 6)) = strUserNum) Then
               If Trim(GRD1.TextMatrix(dblPrevRow, 9)) = "" Then '無讀取日期時間才需更新資料
                  cnnConnection.BeginTrans: bolConn = True
                  strUpdTime = Right("000000" & ServerTime, 6)
                  strExc(0) = "update InputRecord set " & _
                              " ir05=" & strSrvDate(1) & ",ir06=" & strUpdTime & ",ir07='" & strUserNum & "'" & _
                              " where ir01=" & GRD1.TextMatrix(dblPrevRow, 10) & _
                                " and ir02=" & GRD1.TextMatrix(dblPrevRow, 11) & _
                                " and ir03='" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 16)) & "'" & _
                                " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 13)) & "')"
                  cnnConnection.Execute strExc(0)
                  'Add By Sindy 2019/7/17 副本人員只要有讀取信件就上核銷(刪除資訊)
                  strExc(0) = "update InputRecord set " & _
                              " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                              " where ir01=" & GRD1.TextMatrix(dblPrevRow, 10) & _
                                " and ir02=" & GRD1.TextMatrix(dblPrevRow, 11) & _
                                " and ir03='" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 16)) & "'" & _
                                " and upper(ir04)=upper('" & ChgSQL(GRD1.TextMatrix(dblPrevRow, 13)) & "')" & _
                                " and ir24='Y'"
                  cnnConnection.Execute strExc(0)
                  '2019/7/17 END
                  Call SaveInputRecord(CInt(dblPrevRow), False)
                  cnnConnection.CommitTrans: bolConn = False
                  
                  GRD1.TextMatrix(dblPrevRow, 9) = ChangeWStringToTDateString(strSrvDate(1)) & " " & Format(strUpdTime, "00:00:00")
               End If
            End If
            'Modify By Sindy 2017/12/22 Mark
'            If Trim(GRD1.TextMatrix(dblPrevRow, 9)) = "" Then '無讀取日期時間代表開啟沒成功
'               Call CancelRowColor(CInt(dblPrevRow)) '清除反白
'            End If
'         Else
'            MsgBox "無此郵件！", vbInformation
         End If
         Screen.MousePointer = vbDefault
      End If
   End If
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox " 讀取寫入失敗！" & vbCrLf & Err.Description
End Sub

Private Function GetAttachFile(ByVal strPkey1 As String, ByVal strPkey2 As String, _
                               ByVal strPkey3 As String, ByRef pFileName As String, _
                               Optional pSavePath As String) As Boolean
Dim stAttPath As String
   
On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      '改傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      stAttPath = pSavePath
   End If
   
   GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
   
   Exit Function
   
ErrHnd:
   If Err.NUMBER = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub SaveInputRecord(intRow As Integer, Optional bolSendMail As Boolean = True)
Dim strIR22 As String, strIR16 As String 'Modify By Sindy 2022/8/19

   'Modify By Sindy 2019/7/17
   'If (Frame2.Visible = False And Frame3.Visible = False) Or bolRunDel = True Then
   If Trim(GRD1.TextMatrix(intRow, 8)) <> "" Then '已有轉寄資料才須執行下列核銷
   '2019/7/17 END
      '若信件收受者全部已處理或已刪除,主檔就可以掛上msg檔刪除日期,等待AutoBatchDay一個月後刪除實體檔
      strExc(0) = "select ir01 from InputRecord" & _
                  " where ir01=" & GRD1.TextMatrix(intRow, 10) & _
                    " and ir02=" & GRD1.TextMatrix(intRow, 11) & _
                    " and ir03='" & GRD1.TextMatrix(intRow, 16) & "'" & _
                    " and ir08=0" 'and ir05=0 and ir08=0 : 若信件收受者全部已讀取或已刪除
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         '更新"無"Msg檔刪除日期
         strExc(0) = "update PatentInput set" & _
                     " pi16=" & strSrvDate(1) & _
                     " where pi01=" & GRD1.TextMatrix(intRow, 10) & _
                       " and pi02=" & GRD1.TextMatrix(intRow, 11) & _
                       " and pi03='" & GRD1.TextMatrix(intRow, 16) & "'" & _
                       " and pi16=0"
         cnnConnection.Execute strExc(0)
      End If
      
      'Modify By Sindy 2022/8/19
      If bolSendMail = True Then
         '檢查是否有需主管核准的通知信要發
         strExc(0) = "select * from inputrecord" & _
                     " where ir01=" & GRD1.TextMatrix(intRow, 10) & _
                       " and ir03='" & GRD1.TextMatrix(intRow, 16) & "'" & _
                       " and ir04='" & GRD1.TextMatrix(intRow, 13) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strIR16 = "" & RsTemp.Fields("IR16")
            strIR22 = "" & RsTemp.Fields("IR22")
         End If
         If strIR22 <> "" And strIR16 = "1" Then '1.輸入
            strExc(0) = "select cum02 from CaseUseMemo" & _
                        " where cum05='02'" & _
                          " and cum06=" & CNULL(strUserNum) & _
                          " and cum02='" & strIR22 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                           " values('0','" & strIR22 & "','0','0','02')"
               cnnConnection.Execute strExc(0)
               Frame1.Visible = True '*****
            End If
         End If
      End If
      '2022/8/19 END
   End If
End Sub

Private Sub SetColor(Optional intSetRow As Double = 0)
   Dim ii As Integer, jj As Integer
   
   With GRD1
   If .Rows > 1 Then
      .Visible = False
      For ii = IIf(intSetRow = 0, 1, intSetRow) To IIf(intSetRow = 0, .Rows - 1, intSetRow)
         '若有讀取日期時間時,則變灰色
         .row = ii
         If Trim(.TextMatrix(ii, 9)) <> "" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '灰
               .CellBackColor = &HE0E0E0
            Next jj
         End If
         SetColor2 'Added by Morgan 2020/7/28
         
'         If Trim(.TextMatrix(ii, 17)) = "" Then
'            .col = 3
'            '淺黃色 '灰
'            .CellBackColor = &HC0FFFF   '&HE0E0E0
'         End If
      Next ii
      If intSetRow = 0 Then .TopRow = 1
      .Visible = True
   End If
   End With
End Sub

'點二下可刪除List1資料列
Private Sub List1_DblClick(Cancel As MSForms.ReturnBoolean)
Dim strText As String
   
   If List1.ListIndex >= 0 Then
      strText = List1.List(List1.ListIndex)
      List1.RemoveItem List1.ListIndex
      If List1.Tag <> "" Then
         List1.Tag = Replace(List1.Tag, ";" & strText, "")
         List1.Tag = Replace(List1.Tag, strText, "")
      End If
   End If
End Sub

'收受者
'Private Sub cboII06_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   cboII06.ToolTipText = "按Tab或點二下加入資料至右邊資料區"
'End Sub
'Private Sub cboII06_DblClick(Cancel As MSForms.ReturnBoolean)
'   If cboII06.ListIndex >= 0 Then
'      If InStr(List1.Tag, cboII06.List(cboII06.ListIndex)) = 0 Then
'         If List1.Tag = "" Then List1.Clear
'         If CheckDataValid(cboII06.List(cboII06.ListIndex)) = False Then GRD1.Visible = True: Exit Sub
'         List1.AddItem cboII06.List(cboII06.ListIndex)
'         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.List(cboII06.ListIndex)
'         List1.SetFocus
'      End If
'      cboII06.ListIndex = 0
'   End If
'End Sub
Private Sub cboII06_Click()
   If cboII06.ListIndex >= 0 Then
      If InStr(List1.Tag, cboII06.List(cboII06.ListIndex)) = 0 Then
         If List1.Tag = "" Then List1.Clear
         If CheckDataValid(cboII06.List(cboII06.ListIndex)) = False Then GRD1.Visible = True: Exit Sub
         List1.AddItem cboII06.List(cboII06.ListIndex)
         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.List(cboII06.ListIndex)
         List1.SetFocus
      End If
      cboII06.ListIndex = 0
   End If
End Sub
Private Sub cboII06_Validate(Cancel As Boolean)
   Call cboII06_LostFocus '收受者輸至第五碼時自動帶出姓名
   If cboII06.Text <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(cboII06, 5)) = True Then
         'cboII06.Text = ""
         cboII06.SetFocus
         Call cboII06_GotFocus
         Exit Sub
      Else
         Call Command1_Click
      End If
   End If
End Sub

Private Sub cboII06_GotFocus()
   cboII06.SelStart = 0
   cboII06.SelLength = Len(cboII06.Text)
End Sub
Private Sub cboII06_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboII06_LostFocus()
Dim strText As String
   
   cboII06.Text = Trim(cboII06.Text) 'Add By Sindy 2021/11/19
   If cboII06.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboII06.Text)
      If strText <> "" Then
         cboII06.Text = strText & " " & cboII06.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboII06.Text, 5))
         If strText <> "" Then
            cboII06.Text = Left(cboII06.Text, 5) & " " & strText
         End If
      End If
   End If
End Sub

Private Sub ClearText()
   'm_TxtIR20 = "可輸入處理原因"
   m_TxtIR20 = ""
   TxtIR20 = m_TxtIR20
   txtPI18 = "": txtPI18.Tag = ""
   txtPI19 = ""
   txtPI20 = ""
   txtPI21 = ""
   txtRecvNo = "": Me.Tag = ""
   TextContext = vbCrLf & "信件內容參附件！"
End Sub

'查詢勾選的資料
Private Sub ReadFirstGrd1Text()
   Call ClearText
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            '將本所案號顯示在畫面上
            'If Trim(.TextMatrix(.row, 4)) <> "" Then
               'txtPI18.Tag = m_iRow
               txtPI18 = SystemNumber(Trim(.TextMatrix(.row, 4)), 1)
               txtPI19 = SystemNumber(Trim(.TextMatrix(.row, 4)), 2)
               txtPI20 = SystemNumber(Trim(.TextMatrix(.row, 4)), 3)
               txtPI21 = SystemNumber(Trim(.TextMatrix(.row, 4)), 4)
               txtRecvNo = Trim(.TextMatrix(.row, 24))
               TxtIR20 = Trim(.TextMatrix(.row, 25))
               If TxtIR20 = "" And Check2.Value = 0 Then
                  TxtIR20 = m_TxtIR20
               ElseIf Check2.Value = 1 Then
                  'Added by Morgan 2020/7/21 來函期限2次確認
                  If stIR16 = "1" Then
                     cmdReInput.Visible = True
                     'Modify By Sindy 2021/4/6
                     'Frame3.Visible = False
                     cmdAgree.Visible = False
                     cmdback.Visible = False
                     '2021/4/6 END
                     m_TxtIR20 = TxtIR20
                  Else
                     cmdReInput.Visible = False
                     'Modify By Sindy 2021/4/6
                     'Frame3.Visible = True
                     cmdAgree.Visible = True
                     cmdback.Visible = True
                     '2021/4/6 END
                  'end 2020/7/21
                     'Add By Sindy 2018/6/5 專利處程序須第二次確認
                     m_TxtIR20 = IIf(TxtIR20 <> "", TxtIR20 & "; ", "") & "第2次確認原因:"
                     
                  End If 'Added by Morgan 2020/7/21
                  TxtIR20 = m_TxtIR20
               End If
               
               'Added by Morgan 2020/8/13
               If Check2.Value = 0 Then
                  If stIR16 = "7" Or stIR16 = "8" Then
                     SetButton False, stIR16
                  Else
                     SetButton True
                  End If
               End If
               'end 2020/8/13
               
               If txtRecvNo.Visible = True And txtRecvNo.Enabled = True Then
                  txtRecvNo.SetFocus
               End If
               Exit Sub
            'End If
         End If
      Next m_iRow
   End With
End Sub

Private Function PatentInputForm() As Boolean
   PatentInputForm = False
   'Added by Morgan 2020/7/29
   If PUB_MGridGetValue(GRD1.row, "IR16", GRD1) = "8" Then
      strExc(1) = PUB_MGridGetValue(GRD1.row, "總收文號", GRD1)
      If strExc(1) <> "" Then
         strExc(0) = "select * from caseprogress where cp09='" & strExc(1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "本信函為來函期限2次確認後退回，若原輸入期限確實有誤，請先刪除該來函後再重新輸入！", vbCritical
            Exit Function
         End If
      End If
   End If
   'end 2020/7/29
   
   strExc(0) = "select pa11 from patent" & _
               " where pa01='" & txtPI18 & "'" & _
                 " and pa02='" & txtPI19 & "'" & _
                 " and pa03='" & txtPI20 & "'" & _
                 " and pa04='" & txtPI21 & "'" & _
               " union select sp11 from servicepractice" & _
               " where sp01='" & txtPI18 & "'" & _
                 " and sp02='" & txtPI19 & "'" & _
                 " and sp03='" & txtPI20 & "'" & _
                 " and sp04='" & txtPI21 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'If ChkSysName(txtPI18) = True Then
         If txtPI18 <> "P" And txtPI18 <> "PS" And _
            txtPI18 <> "CFP" And txtPI18 <> "CPS" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
            Me.txtPI18.SetFocus
            Exit Function
         End If
      'End If
      m_AppNo = "" & RsTemp.Fields("pa11")
'               If m_AppNo <> "" Then
'                  If txtAppNo = "" Then
'                     If txtPI18 = "P" Or txtPI18 = "PS" Then
'                        MsgBox "請輸入申請案號！", vbExclamation
'                        Me.txtAppNo.SetFocus
'                        Exit Sub
'                     End If
'                  ElseIf m_AppNo <> txtAppNo Then
'                     MsgBox "申請案號輸入錯誤！", vbExclamation
'                     Me.txtAppNo.SetFocus
'                     Exit Sub
'                  Else
'                  End If
'               End If
      If Trim(GRD1.TextMatrix(GRD1.row, 4)) <> txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21 Then
         strExc(0) = "update patentinput set " & _
                     "pi18='" & txtPI18 & "',pi19='" & txtPI19 & "'," & _
                     "pi20='" & txtPI20 & "',pi21='" & txtPI21 & "'" & _
                     " where pi01=" & m_strIR01 & _
                     " and pi02=" & m_strIR02 & _
                     " and pi03='" & m_strIR03 & "'"
         cnnConnection.Execute strExc(0)
         GRD1.TextMatrix(GRD1.row, 4) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
      End If
   Else
      MsgBox "本所案號輸入錯誤！", vbExclamation, "警告！"
      Me.txtPI18.SetFocus
      Exit Function
   End If
   '暫時:
   'Pub_SeekTbLog "Update caseprogress set cp64=cp64||'信件編號:" & m_strIR01 & "-" & m_strIR03 & "-" & m_strIR04 & "' where cp01='" & txtPI18 & "' and cp02='" & txtPI19 & "' and cp03='" & txtPI20 & "' and cp04='" & txtPI21 & "'"
   'END
   If txtPI18 = "P" Or txtPI18 = "PS" Then
      Call Forms(0).SetTmpfrm04010519(Me) 'Add By Sindy 2022/5/20
      PopupMenu mdiMain.mnuPopEMail1
   Else
      Call Forms(0).SetTmpfrm04010519(Me) 'Add By Sindy 2022/5/20
      PopupMenu mdiMain.mnuPopEMail2
   End If
   PatentInputForm = True
End Function

'輸入
Private Sub cmdInput_Click()
   With GRD1
      For m_iRow = 1 To .Rows - 1
         .row = m_iRow
         txtPI18.Tag = m_iRow
         If .TextMatrix(.row, 0) = "V" And .RowHeight(.row) > 0 Then
            If dblPrevRow <> .row Then
               MsgBox "資料列選取有誤，請重新確認！"
               Exit Sub
            End If
            m_strIR01 = .TextMatrix(.row, 10)
            m_strIR02 = .TextMatrix(.row, 11)
            m_strIR03 = .TextMatrix(.row, 16)
            m_strIR04 = .TextMatrix(.row, 13)
            'm_strPi12 = .TextMatrix(.row, 21) '收信日期
            m_strPi12 = .TextMatrix(.row, 17) '轉寄日期
            If Val(m_strPi12) > 0 Then
               m_strPi12 = Val(m_strPi12) - 19110000
            End If
            m_strPi05 = .TextMatrix(.row, 22)
            
            m_AppNo = "": m_RegNo = ""
            '檢查是否有輸入本所案號
'            If Trim(.TextMatrix(.row, 4)) = "" Then
'               If txtPI18 = "" Or txtPI19 = "" Then
'                  MsgBox "請輸入本所案號！", vbExclamation, "警告！"
'                  Me.txtPI18.SetFocus
'                  Exit Sub
'               End If
'            Else
'               txtPI18 = SystemNumber(Trim(.TextMatrix(.row, 4)), 1)
'               txtPI19 = SystemNumber(Trim(.TextMatrix(.row, 4)), 2)
'               txtPI20 = SystemNumber(Trim(.TextMatrix(.row, 4)), 3)
'               txtPI21 = SystemNumber(Trim(.TextMatrix(.row, 4)), 4)
'            End If
            If txtPI18 = "" Or txtPI19 = "" Then
               MsgBox "請輸入本所案號！", vbExclamation, "警告！"
               Me.txtPI18.SetFocus
               Exit Sub
            End If
            If txtPI20 = "" Then txtPI20 = "0"
            If txtPI21 = "" Then txtPI21 = "00"
            
            '專利處程序輸入
            If PatentInputForm = False Then
               Exit Sub
            End If
            Exit For
         End If
      Next m_iRow
   End With
End Sub

Private Sub TextContext_Change()
   PUB_RefreshText TextContext
End Sub

Private Sub txtIR20_Change()
   PUB_RefreshText TxtIR20
End Sub

Private Sub txtPI18_GotFocus()
   TextInverse txtPI18
   CloseIme
End Sub

Private Sub txtPI18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPI18_Validate(Cancel As Boolean)
   If txtPI18 <> "" Then
      txtPI18 = UCase(txtPI18)
      If txtPI18 <> "P" And txtPI18 <> "PS" And _
         txtPI18 <> "CFP" And txtPI18 <> "CPS" Then
         MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtPI18
End Sub

Private Sub txtPI19_GotFocus()
   TextInverse txtPI19
End Sub

Private Sub txtPI20_GotFocus()
   TextInverse txtPI20
End Sub

Private Sub txtPI20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPI20_LostFocus()
   If txtPI18 <> "" And txtPI19 <> "" And txtPI20 = "" Then txtPI20 = "0"
End Sub

Private Sub txtPI21_GotFocus()
   TextInverse txtPI21
End Sub

Private Sub txtPI21_LostFocus()
   If txtPI18 <> "" And txtPI19 <> "" And txtPI21 = "" Then txtPI21 = "00"
End Sub

Private Sub txtPI21_Validate(Cancel As Boolean)
Dim strPI05 As String
Dim strPI06 As String
   
   If txtPI18 <> "" And txtPI19 <> "" Then
      If txtPI20 = "" Then txtPI20 = "0"
      If txtPI21 = "" Then txtPI21 = "00"
   End If
End Sub

Private Sub txtRecvNo_GotFocus()
   TextInverse txtRecvNo
End Sub

'Added by Morgan 2020/7/29
'來函期限2次確認/退回變色
Private Sub SetColor2()
   With GRD1
   stIR16 = PUB_MGridGetValue(.row, "IR16", GRD1)
   If Pub_StrUserSt03 = "P12" Then
      If Check2.Value = 1 Then
         '來函期限2次確認
         If stIR16 = "1" Then
            .col = 7
            .CellBackColor = &H80FFFF '黃
         End If
      Else
         '來函期限確認OK
         If stIR16 = "7" Then
            .col = 7
            .CellBackColor = &H80FFFF '黃
         '來函期限確認退回
         ElseIf stIR16 = "8" Then
            .col = 7
            .CellBackColor = &HFF '紅
         End If
      End If
   End If
   End With
End Sub

'Added by Morgan 2020/8/13
Private Sub SetButton(pEnabled As Boolean, Optional pStatus As String)
   cmdInput.Enabled = pEnabled
   cmdNotProDel.Enabled = pEnabled
   cmdPDF.Enabled = pEnabled
   cmdProDel.Enabled = pEnabled
   cmdRecall.Enabled = pEnabled
   TxtIR20.Enabled = pEnabled
   If pEnabled = False Then
      If pStatus = "8" Then
         cmdInput.Enabled = True
      Else
         cmdProDel.Enabled = True
      End If
   End If
End Sub
