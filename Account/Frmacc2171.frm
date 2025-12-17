VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc2171 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子結匯作業"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8892
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8892
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "點我展開"
      Height          =   345
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   2940
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   2970
      ExtentX         =   5239
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2880
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   5865
      _ExtentX        =   10351
      _ExtentY        =   5080
      _Version        =   393216
      Cols            =   10
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|單據編號|急件付款日|幣別|金額|單據日期|代理人名稱|代理人編號|智|註記"
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
      _Band(0).Cols   =   10
   End
   Begin VB.Frame Frame1 
      Height          =   2955
      Left            =   3015
      TabIndex        =   3
      Top             =   -90
      Width           =   5850
      Begin VB.CommandButton cmdOK 
         Caption         =   "註記"
         Height          =   315
         Index           =   3
         Left            =   2550
         Style           =   1  '圖片外觀
         TabIndex        =   40
         Top             =   150
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "智權公司"
         Height          =   255
         Left            =   3510
         TabIndex        =   39
         Top             =   1230
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "其他幣(非大陸)"
         Height          =   195
         Index           =   3
         Left            =   3465
         TabIndex        =   38
         Top             =   1560
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "USD(非大陸)"
         Height          =   195
         Index           =   2
         Left            =   2025
         TabIndex        =   37
         Top             =   1560
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "大陸"
         Height          =   195
         Index           =   1
         Left            =   1125
         TabIndex        =   36
         Top             =   1560
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Caption         =   "全部"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   35
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "符合寬度"
         Height          =   315
         Left            =   3996
         TabIndex        =   34
         Top             =   150
         Width           =   984
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   2412
         TabIndex        =   9
         Top             =   840
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   1164
         TabIndex        =   8
         Top             =   840
         Width           =   1005
      End
      Begin VB.CommandButton Command3 
         Default         =   -1  'True
         Height          =   300
         Left            =   2250
         Picture         =   "Frmacc2171.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   516
         Width           =   350
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "查詢"
         Height          =   315
         Index           =   2
         Left            =   3492
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1176
         TabIndex        =   6
         Text            =   "U12345678"
         Top             =   516
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "轉入待結匯"
         Height          =   315
         Index           =   0
         Left            =   1335
         Style           =   1  '圖片外觀
         TabIndex        =   5
         Top             =   150
         Width           =   1140
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "暫不結匯"
         Height          =   315
         Index           =   1
         Left            =   180
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   150
         Width           =   1092
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1155
         Left            =   75
         TabIndex        =   15
         Top             =   1770
         Width           =   4950
         _ExtentX        =   8721
         _ExtentY        =   2032
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "智|本所案號|案件性質|單據金額|案件名稱"
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
         _Band(0).Cols   =   5
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   1155
         TabIndex        =   10
         Top             =   1170
         Width           =   1005
         _ExtentX        =   1778
         _ExtentY        =   550
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   315
         Left            =   2415
         TabIndex        =   11
         Top             =   1170
         Width           =   1005
         _ExtentX        =   1778
         _ExtentY        =   550
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '透明
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   13.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2235
         TabIndex        =   33
         Top             =   1170
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "單據日期"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   8
         Left            =   180
         TabIndex        =   32
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   13.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2235
         TabIndex        =   31
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人編號"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   30
         Top             =   900
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "單據編號"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   192
         TabIndex        =   14
         Top             =   576
         Width           =   720
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   4560
         TabIndex        =   13
         Top             =   930
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   1848
      Left            =   3852
      TabIndex        =   16
      Top             =   5868
      Visible         =   0   'False
      Width           =   4944
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "106/1/1"
         Top             =   36
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   912
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Y12345678"
         Top             =   396
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1896
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "AAA"
         Top             =   396
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1272
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Y12345678"
         Top             =   732
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   912
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "USD"
         Top             =   1056
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   315
         Index           =   6
         Left            =   3432
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "999,999"
         Top             =   1056
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Index           =   7
         Left            =   912
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1392
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "單據日期"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   2892
         TabIndex        =   29
         Top             =   96
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   144
         TabIndex        =   28
         Top             =   456
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人D/N No."
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   144
         TabIndex        =   27
         Top             =   792
         Width           =   1128
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "幣別"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   144
         TabIndex        =   26
         Top             =   1116
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "金額"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   2940
         TabIndex        =   25
         Top             =   1116
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   144
         TabIndex        =   24
         Top             =   1392
         Width           =   360
      End
   End
End
Attribute VB_Name = "Frmacc2171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/01 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、MSHFlexGrid2改字型=新細明體-ExtB
'Created by Morgan 2017/1/10
Option Explicit
Dim iPrevRow As Integer '前次點選列
Dim lTotRows As Long, lSelRows As Long
Dim m_blnColOrderAsc As Boolean
Dim m_AttachPath As String

Private Sub Check1_Click()
   cmdOK(2).Value = True
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0 '轉入待結匯
      FormSave
   Case 1 '暫不結匯
      HideRow
   Case 2 '查詢
      SetMouseBusy
      FindData
      SetMouseReady
   'Added by Morgan 2019/5/9
   Case 3 '註記
      If iPrevRow > 0 Then
         strExc(0) = GetValue(iPrevRow, "註記", MSHFlexGrid1)
         strExc(1) = InputBox("請輸入註記內容！(空白或取消都會清除註記)", "設定/取消註記", strExc(0))
         If SetValue(iPrevRow, "註記", strExc(1), MSHFlexGrid1) = True Then
            strExc(2) = GetValue(iPrevRow, "a1501", MSHFlexGrid1)
            If Left(strExc(2), 1) = "U" Then
               cnnConnection.Execute "update acc150 set a1511='" & ChgSQL(strExc(1)) & "' where a1501='" & strExc(2) & "'", intI
            Else
               cnnConnection.Execute "update acc160 set a1609='" & ChgSQL(strExc(1)) & "' where a1601='" & strExc(2) & "'", intI
            End If
            SetTagColor iPrevRow
         End If
      End If
   End Select
End Sub

Private Sub Command1_Click()
            WebBrowser1.SetFocus
            DoEvents
            SendKeys "^2"
End Sub

Private Sub Command3_Click()
   If Text1(0) <> "" Then
      MSHFlexGrid1.Recordset.MoveFirst
      MSHFlexGrid1.Recordset.Find " a1501='" & Text1(0) & "'"
      If MSHFlexGrid1.Recordset.EOF Then
         MsgBox "無此帳單編號！", vbExclamation
         
      Else
         MSHFlexGrid1.row = MSHFlexGrid1.Recordset.AbsolutePosition
         MSHFlexGrid1.TopRow = MSHFlexGrid1.row
         SelectRow MSHFlexGrid1.row
      End If
   End If
End Sub

Private Sub Command4_Click()
   If Command4.Caption = "點我展開" Then
      RePosForm True
   Else
      RePosForm False
   End If
End Sub

Private Sub Form_Activate()
   strFormName = Me.Name
   tool3_enabled
   MenuDisabled
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   m_AttachPath = App.path & "\" & strUserNum
   KillTemp
   OpenTable
   Me.WindowState = 2
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   FormReset
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

Private Sub FindData()
   Dim iRow As Integer, ii As Integer
   Dim bolHidden As Boolean
   Dim iA1502 As Integer, iA1503 As Integer, iRecs As Integer
   Dim iJComp As Integer, iCur As Integer, iFA10 As Integer
   
   With MSHFlexGrid1
   iA1502 = GetFieldId("a1502", MSHFlexGrid1)
   iA1503 = GetFieldId("a1503", MSHFlexGrid1)
   
   iJComp = GetFieldId("智", MSHFlexGrid1)
   iCur = GetFieldId("幣別", MSHFlexGrid1)
   iFA10 = GetFieldId("fa10", MSHFlexGrid1)
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) <> "X" Then
         bolHidden = False
         For ii = 1 To 1
            If Text2(0) <> "" Then
               If .TextMatrix(iRow, iA1503) < Text2(0) Then
                  bolHidden = True
                  Exit For
               End If
            End If
            If Text2(1) <> "" Then
               If .TextMatrix(iRow, iA1503) > Text2(1) Then
                  bolHidden = True
                  Exit For
               End If
            End If
            If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "" Then
               If .TextMatrix(iRow, iA1502) < Val(FCDate(MaskEdBox1.Text)) Then
                  bolHidden = True
                  Exit For
               End If
            End If
            
            If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "" Then
               If .TextMatrix(iRow, iA1502) > Val(FCDate(MaskEdBox2.Text)) Then
                  bolHidden = True
                  Exit For
               End If
            End If
            
            'Added by Morgan 2018/11/16
            '智權
            If Check1.Value = vbChecked Then
               If .TextMatrix(iRow, iJComp) <> "智" Then
                  bolHidden = True
                  Exit For
               End If
            '非智權
            Else
               If .TextMatrix(iRow, iJComp) = "智" Then
                  bolHidden = True
                  Exit For
               End If
            End If
            '大陸
            If Option1(1).Value = True Then
               If Not (.TextMatrix(iRow, iFA10) = "020") Then
                  bolHidden = True
               End If
            'USD(非大陸)
            ElseIf Option1(2).Value = True Then
               If Not (.TextMatrix(iRow, iFA10) <> "020" And .TextMatrix(iRow, iCur) = "USD") Then
                  bolHidden = True
               End If
            '其他幣(非大陸)
            ElseIf Option1(3).Value = True Then
               If .TextMatrix(iRow, iFA10) = "020" Or .TextMatrix(iRow, iCur) = "USD" Then
                  bolHidden = True
               End If
            End If
            'end 2018/11/16
         Next ii
         
         If bolHidden Then
            .TextMatrix(iRow, 0) = ""
            .RowHeight(iRow) = 0
         Else
            .RowHeight(iRow) = .RowHeight(0)
            iRecs = iRecs + 1
         End If
      End If
   Next
   lTotRows = iRecs
   lSelRows = 0
   lblCount = lSelRows & " / " & lTotRows
   DoEvents
   End With
   FormReset
End Sub

Private Sub OpenTable()
   Dim iRow As Integer
   Dim stVTB1 As String, stVTB2 As String
   
   SetGrid True
   'Modified by Morgan 2019/9/3 +輸入人員改控制P12部門(財務處會輸翻譯社的帳單,也不要電子結匯)--婉莘
   'Modified by Morgan 2017/2/20 調整欄位--婉莘
   'Modified by Morgan 2017/2/21 +抵帳單
   'Modified by Morgan 2018/1/25 改帳單都要已審核才可結匯
   'Modified by Morgan 2018/11/30 +CFP,CPS
   'Modified by Morgan 2019/5/9 +財務註記 a1511,a1609
   'Modified by Morgan 2020/10/26 +L公司--婉莘(目前帳單未電子化,若要電子結匯需手動上傳pdf檔)
   'Modify By Sindy 2021/1/19 + P22.商標處程序, 取消 and cp01 in ('P','PS','CFP','CPS','L')
   'Modified by Morgan 2023/3/30 +a1527
   'Modified by Morgan 2023/4/19 +F11外商承辦
   'Modified by Morgan 2025/8/22 +F22外專程序(不用再限制部門)，取消 and st03 in ('P12','L02','P22','F11')
   '帳單
   stVTB1 = "select a1501,a1502,a1503,a1504,a1505,a1506,a1509,a1511,a1527,ayf02,cp01,cp02,cp03,cp04" & _
      " from acc150,acc152,acc170,acc151,caseprogress" & _
      " where a1507||a1512 is null and a1502>920000 and length(a1501)=9" & _
      " and NVL(a1521,'Y')='Y' and ayf01(+)=a1501" & _
      " and a1702(+)=a1501 and a1702 is null" & _
      " and axf01(+)=a1501 and cp09(+)=axf02 "
   '抵帳單
   'Modified by Morgan 2018/1/17 +剔除有結匯日期者(有可能沒有結匯資料,直接上結匯日期)--秀玲
   'Modify By Sindy 2021/1/19 + P22.商標處程序, 取消 and cp01 in ('P','PS','CFP','CPS')
   'and st03='P12' => and st03 in ('P12','L02','P22')
   'Modified by Morgan 2023/4/19 +F11外商承辦
   'Modified by Morgan 2025/8/22 +F22外專程序(不用再限制部門)，取消 and st03 in ('P12','L02','P22','F11')
   stVTB2 = "select a1601,a1602,a1603,a1604,a1605,-1*a1606,a1608,a1609,null,ayf02,cp01,cp02,cp03,cp04" & _
      " from acc160,acc152,acc170,acc161,caseprogress" & _
      " where a1602>920000 and length(a1601)=9 and nvl(a1607,0)=0" & _
      " and ayf01(+)=a1601" & _
      " and a1702(+)=a1601 and a1702 is null" & _
      " and axg01(+)=a1601 and cp09(+)=axg02 "
   'Modified by Morgan 2019/1/25 +distinct(一張帳單不只一個案號時會變多筆 Ex:U10711065)
   'Modified by Morgan 2019/5/30 Y49572 USPTO是固定信用卡刷卡請款. 所以雖然有建帳單, 但其實已經付款完成. 所以不需要列示在電子結匯區--婉莘
   'Modified by Morgan 2020/10/23 +L公司顯示"法"--婉莘(只有L案,其他還是用智慧所)
   'Modified by Morgan 2021/9/24 +SrtCol 排序用的暫存欄位
   'Modified by Morgan 2021/11/5 刷卡方式支付的代理人新增:Y55627 EUROP, Y55645 EUIPO--婉莘
   'Modified by Morgan 2022/7/7 德國專利局 Y5576600 帳單也不顯示--婉莘
   'Modified by Morgan 2023/3/30 +急件付款日，調整公司別欄位到後面
   'Modified by Morgan 2024/5/31 金額加,號並顯示到小數2位--斯閔
   'Modified by Morgan 2024/12/4 德國專利局 Y5576600 帳單改要顯示但可不用帳單結匯--斯閔
   strExc(0) = "select distinct '' Chk,a1501 DocNo,sqldatet(a1527) a1527T" & _
      ",a1505 Cur,to_char(a1506,'9,999,990.00') Amt,sqldatet(a1502) a1502T" & _
      ",nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),nvl(fa04,fa06)) FName" & _
      ",a1503,decode(nvl(pa161,nvl(tm130,nvl(lc48,sp85))),'J','智',decode(lc01,'L','法')) JFlg" & _
      ",a1511,decode(a1503,'Y55766000','Y','') Read,ayf02,a1501,a1502,a1503,a1504,a1509,fa10,'' SrtCol" & _
      " from (" & stVTB1 & " union " & stVTB2 & "),fagent,patent,trademark,lawcase,servicepractice" & _
      " where fa01(+)=substr(a1503,1,8) and fa02(+)=substr(a1503,9)" & _
      " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
      " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
      " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" & _
      " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
      " and a1503 not in ('Y49572000','Y55627000','Y55645000')" & _
      " order by a1527T asc,a1502 asc,a1501 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      lSelRows = 0
      lTotRows = RsTemp.RecordCount
      lblCount = lSelRows & " / " & lTotRows
      SetGrid
      SetTagColor 'Added by Morgan 2019/5/9
      .col = 1: .row = 1
       'SelectRow 1
       Option1(0).Value = True 'Added by Morgan 2018/11/16 預設全部,非智權
      .Visible = True
      End With
   Else
      MsgBox "無待結匯帳單！", vbExclamation
   End If
   FormReset
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 950, 1000, 500, 880, 920, 2000, 950, 260, 2000)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
      
      iPrevRow = 0
      lTotRows = 0
      lSelRows = 0
      lblCount = lSelRows & " / " & lTotRows
   End If
   .FixedCols = 2
   .FormatString = .FormatString
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol = 4 Then
            .ColAlignment(iCol) = flexAlignRightCenter
         Else
            .ColAlignment(iCol) = flexAlignLeftCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub SetGrid2(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(260, 1000, 950, 900, 1540)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid2
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 0
   .FormatString = "智|本所案號|案件性質|單據金額|案件名稱"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol = 3 Then
            .ColAlignment(iCol) = flexAlignRightCenter
         Else
            .ColAlignment(iCol) = flexAlignLeftCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      If Command4.Caption = "點我展開" Then
         RePosForm False
      Else
         RePosForm True
      End If
   End If
End Sub

Private Sub RePosForm(pFull As Boolean)
   If Forms(0).WindowState <> 1 Then
      If pFull = True Then
         WebBrowser1.Width = Me.Width - 90
         Command4.Caption = "點我還原"
      Else
         WebBrowser1.Width = Me.Width - 90 - MSHFlexGrid1.Width
         Command4.Caption = "點我展開"
      End If
      Command4.Width = WebBrowser1.Width
      WebBrowser1.Height = Me.Height - Command4.Height - 390

      MSHFlexGrid1.Top = Frame1.Top + Frame1.Height
      MSHFlexGrid1.Height = Me.Height - Frame1.Top - Frame1.Height - 400
      
      MSHFlexGrid1.Left = Me.Width - 90 - MSHFlexGrid1.Width
      Frame1.Left = MSHFlexGrid1.Left
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc2171 = Nothing
End Sub

Private Sub MSHFlexGrid1_DblClick()
   Dim nCol As Integer
   
   If MSHFlexGrid1.MouseRow > 0 Then
      nCol = MSHFlexGrid1.MouseCol
      If OpenPdf = True Then
         If nCol <> 0 Then
            MSHFlexGrid1.col = 0
            If MSHFlexGrid1.Text = "" Then
               ClickGrid MSHFlexGrid1
            End If
         End If
         'ReadRow 'Removed by Morgan 2023/4/19
      End If
      ReadRow 'Added by Morgan 2023/4/19
   End If
End Sub

Private Function OpenPdf() As Boolean
   Dim stFileName As String
   Dim stA1501 As String, stAyf02 As String
            
   If iPrevRow = 0 Then
      MsgBox "請先點要預覽的帳單！", vbInformation
   Else
      SetMouseBusy
      With MSHFlexGrid1
      If .TextMatrix(iPrevRow, 1) <> "" Then
         FormReset
         DoEvents
         
         stA1501 = GetValue(iPrevRow, "a1501", MSHFlexGrid1)
         stAyf02 = GetValue(iPrevRow, "ayf02", MSHFlexGrid1)
         If stAyf02 = "" Then
            MsgBox "【" & stA1501 & "】電子檔尚未上傳無法預覽！", vbCritical
         ElseIf PUB_GetAttachFile_Invoice(stA1501, stAyf02, m_AttachPath, stFileName) = True Then
            WebBrowser1.Navigate m_AttachPath & "\" & stFileName
            SetValue iPrevRow, "Read", "Y", MSHFlexGrid1
            OpenPdf = True
         End If
      End If
      End With
      SetMouseReady
   End If
End Function

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
   Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim stValue As String
   Dim stA1501 As String
   Dim iSort As Integer 'Added by Morgan 2023/5/19
   
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      iSort = GetFieldId("SrtCol", MSHFlexGrid1) 'Added by Morgan 2023/5/19
      
      '紀錄前次點選的收文號
      If iPrevRow > 0 Then
         stA1501 = GetValue(iPrevRow, "A1501", MSHFlexGrid1)
      End If
      
      'Modified by Morgan 2021/9/24 無法指定第2個排序欄位,改用固定排序欄位排
      '.col = nCol
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, iSort) = .TextMatrix(iRow, nCol) & .TextMatrix(iRow, 1)
      Next
      .col = iSort
      'end 2021/9/24
      
      If m_blnColOrderAsc = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc = False
      End If
               
      '重設排序後前次點選的位置
      If iPrevRow > 0 Then
         For iRow = 1 To .Rows - 1
            If stA1501 = GetValue(iRow, "A1501", MSHFlexGrid1) Then
               iPrevRow = iRow
               Exit For
            End If
         Next
      End If
      
   ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      .row = nRow
      .col = nCol
      If nCol = 0 Then
         ClickGrid MSHFlexGrid1
      End If
      SelectRow nRow
      ReadRow 'Added by Morgan 2023/4/19
   End If
   
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(FlexGrid As MSHFlexGrid)
   With FlexGrid
   If .Text = "V" Then
      lSelRows = lSelRows - 1
      .Text = ""
      
   '已刪除資料標示為 X
   ElseIf .Text = "" Then
      lSelRows = lSelRows + 1
      .Text = "V"
      
   End If
   lblCount = lSelRows & " / " & lTotRows
   End With
End Sub

Private Sub SelectRow(pRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With MSHFlexGrid1
   nCol = .col
   If iPrevRow > 0 Then
      If iPrevRow <> pRow Then
         .row = iPrevRow
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         'Modified by Morgan 2019/5/9 有註記要變色
         'SetRowBK
         SetTagColor iPrevRow
         'end 2019/5/9
      End If
   End If
   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      
      For iCol = .FixedCols + 1 To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   iPrevRow = pRow
   End With
End Sub

Private Sub ReadRow()
   Text1(0) = GetValue(iPrevRow, "單據編號", MSHFlexGrid1)
   
   'Memo by Morgan 2018/1/17 婉莘說不需要明細,測試完後可刪除(表單上的物件也要)
   'Text1(1) = GetValue(iPrevRow, "單據日期", MSHFlexGrid1)
   'Text1(2) = GetValue(iPrevRow, "代理人編號", MSHFlexGrid1)
   'Text1(3) = GetValue(iPrevRow, "代理人名稱", MSHFlexGrid1)
   'Text1(4) = GetValue(iPrevRow, "a1504", MSHFlexGrid1)
   'Text1(5) = GetValue(iPrevRow, "幣別", MSHFlexGrid1)
   'Text1(6) = GetValue(iPrevRow, "金額", MSHFlexGrid1)
   'Text1(7) = GetValue(iPrevRow, "a1509", MSHFlexGrid1)
   'END 2018/1/17
   
   SetGrid2 True
   If Left(Text1(0), 1) = "U" Then
      'Modified by Morgan 2020/10/23 +L公司顯示"法"--婉莘(只有L案,其他還是用智慧所)
      strExc(0) = "select decode(nvl(pa161,nvl(tm130,nvl(lc48,sp85))),'J','智',decode(lc01,'L','法')) ChkJ, axf03, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as cpm03, to_char(axf04,'9,999,990.00') axf04, axf12" & _
      " from acc151, caseprogress, casepropertymap,patent,trademark,lawcase,servicepractice" & _
      " where axf01 = '" & Text1(0) & "' and axf02 = cp09(+) and cp01 = cpm01(+) and cp10 = cpm02(+)" & _
      " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
      " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
      " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" & _
      " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
      " order by axf02 asc"
   Else
      'Modified by Morgan 2020/10/23 +L公司顯示"法"--婉莘(只有L案,其他還是用智慧所)
      strExc(0) = "select decode(nvl(pa161,nvl(tm130,nvl(lc48,sp85))),'J','智',decode(lc01,'L','法')) ChkJ, axg03, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as cpm03, to_char(-1*axg04,'9,999,990.00'), axg12" & _
      " from acc161, caseprogress, casepropertymap,patent,trademark,lawcase,servicepractice" & _
      " where axg01 = '" & Text1(0) & "' and axg02 = cp09(+) and cp01 = cpm01(+) and cp10 = cpm02(+)" & _
      " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
      " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
      " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" & _
      " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
      " order by axg02 asc"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid2
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      SetGrid2
      .col = 1: .row = 1
      .Visible = True
      End With
   End If
            
End Sub

Private Sub FormReset()
   Dim oText As TextBox
   For Each oText In Text1
      oText.Text = ""
   Next
   WebBrowser1.Navigate "about:blank"
   SetGrid2 True
End Sub

Private Sub SetRowBK(Optional pCellBackColor As Long = 0)
   Dim ii As Integer
   With MSHFlexGrid1
      If pCellBackColor = 0 Then pCellBackColor = .BackColor
      For ii = .FixedCols To .Cols - 1
         .col = ii
         .CellBackColor = pCellBackColor
      Next
   End With
End Sub

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function

Private Function GetValue(pRow As Integer, pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As String
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String, ByRef FlexGrid As MSHFlexGrid) As Boolean
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Sub HideRow()
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         .TextMatrix(iRow, 0) = "X"
         .RowHeight(iRow) = 0
         lSelRows = lSelRows - 1
         lTotRows = lTotRows - 1
         lblCount = lSelRows & " / " & lTotRows
         DoEvents
      End If
   Next
   End With
   FormReset
End Sub

Private Sub FormSave()
   Dim iRow As Integer, strDocNo As String, bolResult As Boolean
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         strDocNo = GetValue(iRow, "a1501", MSHFlexGrid1)
         
         'Added by Morgan 2019/5/10
         If GetValue(iRow, "註記", MSHFlexGrid1) <> "" Then
            MsgBox "單據號碼【" & strDocNo & "】有註記，不可轉入待結匯！", vbCritical
         'end 2019/5/10
         ElseIf GetValue(iRow, "Read", MSHFlexGrid1) = "Y" Then
            If Left(strDocNo, 1) = "U" Then
               bolResult = Frmacc2170.Acc170SaveNew(strDocNo, , True)
            Else
               bolResult = Frmacc2180.Acc170SaveNew(strDocNo, , True)
            End If
            If bolResult = True Then
               .TextMatrix(iRow, 0) = "X"
               .RowHeight(iRow) = 0
               lSelRows = lSelRows - 1
               lTotRows = lTotRows - 1
               lblCount = lSelRows & " / " & lTotRows
               DoEvents
            End If
         Else
            MsgBox "單據號碼【" & strDocNo & "】尚未讀取Pdf檔，不可轉入待結匯！", vbCritical
         End If
      End If
   Next
   End With
   FormReset
End Sub

Private Sub Option1_Click(Index As Integer)
   cmdOK(2).Value = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    TextInverse Text2(Index)
    CloseIme
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_LostFocus(Index As Integer)
   If Index = 0 Then
      If Text2(0) <> "" Then
         Text2(1) = Left(Text2(0), 6) & "ZZZ"
      Else
         Text2(1) = Text2(0) 'Added by Morgan 2018/3/26 --婉莘
      End If
   End If
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
    If Text2(Index).Text <> "" Then
        Text2(Index) = Left(Text2(Index) & "00000000", 9)
    End If
End Sub
'Added by Morgan 2019/5/9
'有註記的變色
Private Sub SetTagColor(Optional pRow As Integer = 0)
   Dim iRow As Integer, ii As Integer, iTag As Integer
   With MSHFlexGrid1
   iTag = GetFieldId("註記", MSHFlexGrid1)
   If pRow > 0 Then
      .row = pRow
      If .TextMatrix(pRow, iTag) <> "" Then
         SetRowBK &HFF00&
      Else
         SetRowBK
      End If
   Else
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, iTag) <> "" Then
            .row = iRow
            SetRowBK &HFF00&
         End If
      Next
   End If
   End With
End Sub
