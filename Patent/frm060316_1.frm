VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm060316_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准函"
   ClientHeight    =   4092
   ClientLeft      =   132
   ClientTop       =   996
   ClientWidth     =   9108
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4092
   ScaleWidth      =   9108
   Begin VB.CommandButton cmdCount 
      Caption         =   "件數計算"
      Height          =   345
      Left            =   4440
      TabIndex        =   22
      Top             =   390
      Width           =   915
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只列印承辦單"
      Height          =   225
      Left            =   6090
      TabIndex        =   37
      Top             =   720
      Width           =   1545
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   4
      Left            =   7530
      MaxLength       =   2
      TabIndex        =   18
      Top             =   1365
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   3
      Left            =   7290
      MaxLength       =   1
      TabIndex        =   17
      Top             =   1365
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   2
      Left            =   6450
      MaxLength       =   6
      TabIndex        =   16
      Top             =   1365
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   5970
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "FCP"
      Top             =   1365
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新增"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7950
      TabIndex        =   19
      Top             =   1350
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "刪除"
      Height          =   400
      Index           =   1
      Left            =   7950
      TabIndex        =   20
      Top             =   1770
      Width           =   600
   End
   Begin VB.ListBox List1 
      Height          =   1308
      Index           =   1
      ItemData        =   "frm060316_1.frx":0000
      Left            =   5970
      List            =   "frm060316_1.frx":0002
      TabIndex        =   35
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "設定請款單及定稿"
      Height          =   570
      Left            =   4620
      TabIndex        =   33
      Top             =   3270
      Width           =   4335
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   750
         Style           =   2  '單純下拉式
         TabIndex        =   26
         Top             =   210
         Width           =   3465
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   34
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.TextBox txtLetterDate 
      Height          =   264
      Left            =   5970
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1050
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1365
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1365
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1365
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "FCP"
      Top             =   1365
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   400
      Index           =   0
      Left            =   3765
      TabIndex        =   13
      Top             =   1350
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刪除"
      Height          =   400
      Index           =   1
      Left            =   3765
      TabIndex        =   14
      Top             =   1770
      Width           =   600
   End
   Begin VB.ListBox List1 
      Height          =   1308
      Index           =   0
      ItemData        =   "frm060316_1.frx":0004
      Left            =   1815
      List            =   "frm060316_1.frx":0006
      TabIndex        =   30
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Height          =   400
      Left            =   5460
      TabIndex        =   21
      Top             =   4350
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7920
      TabIndex        =   24
      Top             =   180
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   6900
      TabIndex        =   23
      Top             =   180
      Width           =   972
   End
   Begin MSFlexGridLib.MSFlexGrid grdList 
      Height          =   2955
      Left            =   150
      TabIndex        =   27
      Top             =   4470
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   5207
      _Version        =   393216
   End
   Begin VB.TextBox textPA04 
      Height          =   264
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1050
      Width           =   375
   End
   Begin VB.TextBox textPA03 
      Height          =   264
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1050
      Width           =   255
   End
   Begin VB.TextBox textPA02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1050
      Width           =   855
   End
   Begin VB.TextBox textPA01 
      Height          =   264
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1050
      Width           =   495
   End
   Begin VB.TextBox textCP05_2 
      Height          =   264
      Left            =   3300
      MaxLength       =   7
      TabIndex        =   1
      Top             =   750
      Width           =   1035
   End
   Begin VB.TextBox textCP05_1 
      Height          =   264
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   750
      Width           =   1035
   End
   Begin VB.OptionButton optSel 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton optSel 
      Caption         =   "核准發文日："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   570
      Left            =   180
      TabIndex        =   28
      Top             =   3270
      Width           =   4365
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   735
         Style           =   2  '單純下拉式
         TabIndex        =   25
         Top             =   210
         Width           =   3510
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   29
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3000
      X2              =   3120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   3300
      TabIndex        =   41
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   2
      Left            =   1800
      TabIndex        =   40
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "上次列印發文日："
      Height          =   180
      Index           =   3
      Left            =   255
      TabIndex        =   39
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label LblCount 
      Caption         =   " 0 筆"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4440
      TabIndex        =   38
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本次不適用案號："
      Height          =   180
      Index           =   1
      Left            =   4530
      TabIndex        =   36
      Top             =   1380
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "定稿日期："
      Height          =   180
      Index           =   0
      Left            =   5070
      TabIndex        =   32
      Top             =   1095
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "有檢索案號："
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   31
      Top             =   1380
      Width           =   1080
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3000
      X2              =   3120
      Y1              =   855
      Y2              =   855
   End
End
Attribute VB_Name = "frm060316_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/3/2 Form2.0畫面無物件需修改 (Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim m_PA01 As String
Dim m_PA02 As String
Dim m_PA03 As String
Dim m_PA04 As String
Dim m_PA08 As String 'Add by Morgan 2011/7/1 報價要用
Dim m_PA09 As String

Dim m_CP09 As String
Dim m_CP10 As String 'Add by Sindy 2018/10/18
'Add by Morgan 2004/6/23
Dim m_CP05 As String '來函收文日
Dim m_CurrSel As Integer
'Add By Cheng 2003/01/02
Dim m_blnPriData As Boolean '是否有優先權資料
Dim m_bln3PriData As Boolean '是否有三個以上優先權資料
'Add By Cheng 2003/01/28
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Add by Morgan 2004/7/27
Dim m_LetterLanguage As String
Dim m_PA75 As String '代理人
Dim m_PA26 As String '申請人 Add by Morgan 2011/6/27
'Added by Morgan 2014/7/24
Dim m_PA27 As String
Dim m_PA28 As String
Dim m_PA29 As String
Dim m_PA30 As String
'end 2014/7/24
Dim m_WithReportList As String '有檢索案號
Public m_bPrintBill As Boolean '是否列印請款單 Add by Morgan 2011/6/23
Public m_iBillPageCount As Integer '請款單頁數 Add by Morgan 2011/6/27
Dim m_bolDivSug As Boolean 'Added by Morgan 2012/12/12 有分割建議文字
Dim m_strDivState As String 'Added by Morgan 2012/12/26 N:不可提分割,Y:可提分割,"":不確定
Dim strPrinter2 As String 'Add By Sindy 2015/6/26
Const m_ModifyDateUp = 20151101 'Add By Sindy 2015/10/22 原來函收文日改為核准發文日
Dim PrinterIndex As Integer, m_AttachPath As String 'Add By Sindy 2017/6/8
Dim strErrText As String 'Add By Sindy 2019/1/28


'Add By Sindy 2015/7/14
Private Sub cmdCount_Click()
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strText As String
Dim strHadCnt As String
   
   If Val(textCP05_1) > 0 And Val(textCP05_2) > 0 Then
      lblCount = " 0 / 0 筆"
      If optSel(0).Value = True Then
         Call QueryLetterData(True, strHadCnt)
         lblCount = Val(strHadCnt) - Val(List1(1).ListCount) & " / " & Val(List1(1).ListCount) & " 筆"
         '檢查是否有不適用案號,但單筆未下過的案號
         If List1(1).ListCount > 0 Then
            strText = ""
            For i = 0 To List1(1).ListCount - 1
               strSql = "SELECT * FROM CaseUseMemo" & _
                        " WHERE CUM01||CUM02||CUM03||CUM04='" & List1(1).List(i) & "' AND CUM05='01'"
               Set rsTmp = New ADODB.Recordset
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount = 0 Then
                  strText = strText & "," & List1(1).List(i)
               End If
               rsTmp.Close
            Next i
            If strText <> "" Then
               strText = Mid(strText, 2)
               MsgBox "尚有 " & strText & " 未下過單筆列印，" & vbCrLf & _
                      "請先行處理定稿。", vbInformation
            End If
         End If
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Me.Enabled = False
   
   'Move to Unload by Morgan 2004/10/26
'    'Add By Cheng 2003/09/10
'    '列印定稿整批列印清單
'    PUB_PrintLetterList strUserNum
'    '刪除定稿整批列印資料
'    PUB_DeleteLetterList strUserNum
'    'Add By Cheng 2003/01/29
'    '列印地址條
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    'Add By Cheng 2003/02/05
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
   '2004/10/26 end
   
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim strTmp As String
   If Index = 0 And Text1(2).Text <> "" Then
      strTmp = Text1(1) & Text1(2)
      If Text1(3).Text = "" Then
         strTmp = strTmp & "0"
      Else
         strTmp = strTmp & Text1(3).Text
      End If
      If Text1(4).Text = "" Then
         strTmp = strTmp & "00"
      Else
         strTmp = strTmp & Text1(4).Text
      End If
      intI = 1
      strExc(0) = "SELECT PA57 FROM PATENT WHERE " & ChgPatent(strTmp)
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) = "Y" Then
            MsgBox "必須為未閉卷之案號，請重新輸入 !", vbCritical
         Else
            List1(0).AddItem strTmp
            Text1(2).Text = ""
         End If
      Else
         MsgBox "案號不存在，請重新輸入 !", vbCritical
      End If
      Text1(2).SetFocus
   Else
      If List1(0).ListIndex > -1 Then List1(0).RemoveItem List1(0).ListIndex
   End If
End Sub

'Add By Sindy 2014/4/18
Private Sub Command2_Click(Index As Integer)
   Dim strTmp As String
   If Index = 0 And Text2(2).Text <> "" Then
      strTmp = Text2(1) & Text2(2)
      If Text2(3).Text = "" Then
         strTmp = strTmp & "0"
      Else
         strTmp = strTmp & Text2(3).Text
      End If
      If Text2(4).Text = "" Then
         strTmp = strTmp & "00"
      Else
         strTmp = strTmp & Text2(4).Text
      End If
      intI = 1
      strExc(0) = "SELECT PA57 FROM PATENT WHERE " & ChgPatent(strTmp)
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) = "Y" Then
            MsgBox "必須為未閉卷之案號，請重新輸入 !", vbCritical
         Else
            List1(1).AddItem strTmp
            Text2(2).Text = ""
         End If
      Else
         MsgBox "案號不存在，請重新輸入 !", vbCritical
      End If
      Text2(2).SetFocus
   Else
      If List1(1).ListIndex > -1 Then List1(1).RemoveItem List1(1).ListIndex
   End If
End Sub

Private Sub Form_Load()
'Add By Cheng 2003/02/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
   MoveFormToCenter Me
   
   'Add By Sindy 2017/6/8
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   '2017/6/8 END
   
'   'Add By Sindy 2015/10/22 原來函收文日改為核准發文日
'   If strSrvDate(1) >= m_ModifyDateUp Then
'      optSel(0).Caption = "核准發文日："
'      Label1(3).Caption = "上次列印發文日："
'   End If
'   '2015/10/22 END
   
   'InitialGridList
   If optSel(0).Value = True Then
      cmdQuery.Enabled = False
      EnableTextBox textCP05_1, True
      EnableTextBox textCP05_2, True
      EnableTextBox textPA01, False
      EnableTextBox textPA02, False
      EnableTextBox textPA03, False
      EnableTextBox textPA04, False
   Else
      cmdQuery.Enabled = True
      EnableTextBox textCP05_1, False
      EnableTextBox textCP05_2, False
      EnableTextBox textPA01, True
      EnableTextBox textPA02, True
      EnableTextBox textPA03, True
      EnableTextBox textPA04, True
   End If
   
   'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1
   'end 2011/3/15
   
   txtLetterDate = strSrvDate(2) 'Add by Morgan 2009/8/20 定稿日期
   
   'Add by Morgan 2011/6/23
   'Modify By Sindy 2015/6/26 +strPrinter2
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
   
   'Add By Sindy 2015/6/18 刪除3個月前的核准函單筆”本所案號用途記錄”資料
   strExc(0) = "delete from CaseUseMemo where CUM05='01' and CUM07<" & DBDATE(DateAdd("m", -3, Format(strSrvDate(1), "####/##/##")))
   cnnConnection.Execute strExc(0)
   
   'Add By Sindy 2015/7/16 紀錄在資料庫,否則換電腦或使用者會讀不到
   Label2(2).Caption = PUB_GetLastDate(Me.Name, "DATE1")
   Label2(3).Caption = PUB_GetLastDate(Me.Name, "DATE2")
   '2015/7/16 END
   
   'MsgBox "本程式已改為直接列印定稿，請先選定印表機並放好定稿紙！", vbExclamation
   MsgBox "本程式已改為直接列印定稿，請先選定印表機並放好空白紙！", vbExclamation
End Sub

Public Sub Clear()
   textCP05_1.Text = Empty
   textCP05_2.Text = Empty
   textPA01.Text = Empty
   textPA02.Text = Empty
   textPA03.Text = Empty
   textPA04.Text = Empty
'   InitialGridList
End Sub

Private Sub doProcess()

   Dim bFind As Boolean
   
   'Add by Morgan 2009/8/19
   g_LetterDate = DBDATE(txtLetterDate)
   m_WithReportList = ""
   For intI = 0 To List1(0).ListCount - 1
      m_WithReportList = m_WithReportList & List1(0).List(intI) & ";"
   Next
'   '來函收文日
'   If optSel(0).Value = True Then
      If CheckDataValid() = True Then
         Screen.MousePointer = vbHourglass
         If QueryLetterData() = False Then
            Screen.MousePointer = vbDefault
            MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "查詢資料"
            Exit Sub
         Else
            Screen.MousePointer = vbDefault
            MsgBox "作業完成", vbOKOnly + vbInformation, "執行作業"
         End If
         Clear
         SetInputFocus
      End If
'   '本所案號
'   Else
'      If grdList.row > 0 And grdList.row <= grdList.Rows Then
'         '92.2.13 modify by sonia
'         'If grdList.TextMatrix(grdList.Row, 3) <> "核准" Then
'         If grdList.TextMatrix(grdList.row, 3) <> "核准" And grdList.TextMatrix(grdList.row, 3) <> "改變原處分" Then
'         '92.2.13 end
'            MsgBox "請先選取一筆案件性質為核准或改變原處分的記錄", vbOKOnly + vbCritical, "檢核資料"
'            Exit Sub
'         End If
'
'         'Added by Morgan 2012/12/5
'         '配合發明初審核准控制
'         strExc(0) = "select cp27 from caseprogress where cp09='" & m_CP09 & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If IsNull(RsTemp(0)) Then
'               MsgBox "尚未發文不可列印！"
'               Exit Sub
'            End If
'         End If
'         'end 2012/12/5
'
'         frm060316_2.SetData 0, m_PA01, True
'         frm060316_2.SetData 1, m_PA02, False
'         frm060316_2.SetData 2, m_PA03, False
'         frm060316_2.SetData 3, m_PA04, False
'         frm060316_2.SetData 4, m_CP09, False
'         frm060316_2.SetData 5, "frm060316_1", False
'         'Add by Morgan 2004/7/27
'         frm060316_2.SetData 6, m_CP05, False
'         'Add by Morgan 2011/1/14
'         strExc(1) = GetPS(m_PA01, m_PA02, m_PA03, m_PA04)
'         If strExc(1) <> "" Then
'            frm060316_2.Combo1.AddItem strExc(1), 0
'            frm060316_2.Combo1.ListIndex = 0
'         End If
'         'end 2011/1/4
'         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
'         frm060316_2.Show
'         frm060316_2.QueryData
'         Me.Hide
'      End If
'   End If
End Sub

'Modified by Morgan 2011/12/2
'因為有發生執行了兩次的情形,所改寫成呼叫sub方便鎖住按鈕
Private Sub cmdOK_Click()
   cmdOK.Enabled = False
   CmdExit.Enabled = False 'Add By Sindy 2015/10/27
   doProcess
   cmdOK.Enabled = True
   CmdExit.Enabled = True 'Add By Sindy 2015/10/27
End Sub

'Removed by Morgan 2016/11/29 併入 StartLetter
'Private Function GetPS(PA01 As String, PA02 As String, PA03 As String, PA04 As String) As String

'   Dim stSQL As String, iR As Integer
'   stSQL = "select pa75,pa26||pa27||pa28||pa29||pa30 from patent where pa01='" & PA01 & "'" & _
'      " and pa02='" & PA02 & "' and pa03='" & PA03 & "' and pa04='" & PA04 & "'"
'   iR = 1
'   Set AdoRecordSet3 = ClsLawReadRstMsg(iR, stSQL)
'   If iR = 1 Then
'      With AdoRecordSet3
'      'Modified by Morgan 2014/2/10 設定重複,留範圍大的 --毓芳;判斷所有申請人
'      'If .Fields("pa75") = "Y34232000" And .Fields("pa26") = "X30299000" Then
'      If InStr("" & .Fields(1), "X3029900") > 0 Then
'         GetPS = "Per the general instructions of Sekisui Chemical, attached please find the allowed Taiwanese claims for your review."
'      'Added by Morgan 2014/2/10--江如玉
'      ElseIf Left("" & .Fields("pa75"), 8) = "Y1892300" And (InStr("" & .Fields(1), "X2798400") > 0 Or InStr("" & .Fields(1), "X4771900") > 0) Then
'         GetPS = "The English translation of the allowed claims will follow shortly."
'      'end 2014/2/10
'      'Added by Morgan 2016/7/29 Y20990010--鄭詠心, Y52527010--葉敏莉
'      'Modified by Morgan 2016/11/21 +Y45161000--陳亭妙
'      ElseIf .Fields("pa75") = "Y20990010" Or .Fields("pa75") = "Y52527010" Or .Fields("pa75") = "Y45616000" Then
'         GetPS = "Enclosed please find the allowed claims for this application for your reference and records."
'
'      'Added by Morgan 2016/10/20 --陳怡蓉
'      ElseIf .Fields("pa75") = "Y45801B20" And (InStr("" & .Fields(1), "X47805000") > 0 Or InStr("" & .Fields(1), "X70269010") > 0) Then
'         GetPS = "Enclosed please find the allowed claims in English for your reference and records."
'
'      'Added by Morgan 2016/11/21 --陳怡蓉
'      ElseIf .Fields("pa75") = "Y54047000" Then
'         GetPS = "Enclosed please find a copy of the most recently approved claims for your records and prompt reference."
'
'      'Added by Morgan 2016/11/24 --葉子寧
'      'Removed by Morgan 2016/11/29 與Elisa 2016/8/22 的需求重複
'      'ElseIf .Fields("pa75") = "Y52242000" Then
'      '   GetPS = "Enclosed please find an English version of the allowed claims for your records and prompt reference."
'      'end 2016/11/29
'      End If
'      End With
'   End If
'End Function
'end 2016/11/29

'Modify By Sindy 2015/7/15 +Optional ByVal bolQueryCnt As Boolean, Optional ByRef strHadCnt As String
Private Function QueryLetterData(Optional ByVal bolQueryCnt As Boolean, Optional ByRef strHadCnt As String) As Boolean
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim strBillNo As String '待印請款單號 Add by Morgan 2011/6/23
Dim strTmp As String 'Add By Sindy 2014/4/18
Dim bFind As Boolean
'Added by Morgan 2014/6/3
Dim bolDNEmail As Boolean, bolDNPlusPaper As Boolean
Dim strStarTime As String 'Add By Sindy 2015/6/26
Dim bolSingleRun As Boolean
Dim strCP48 As String 'Add By Sindy 2017/1/12
   
On Error GoTo ErrHnd
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & optSel(0).Caption & textCP05_1 & "-" & textCP05_2 'Add By Sindy 2010/12/7
   
   '本所案號
   m_PA01 = textPA01
   m_PA02 = textPA02
   m_PA03 = textPA03
   m_PA04 = textPA04
   
   'Add By Sindy 2014/4/18
   strTmp = ""
   For i = 0 To List1(1).ListCount - 1
      strTmp = strTmp & List1(1).List(i) & ","
   Next
   '2014/4/18 END
   
   bFind = False
    'Modify By Cheng 2003/01/28
    '限定FCP的案件
'   strSQL = "SELECT CP43 FROM CASEPROGRESS " & _
'            "WHERE CP05 >= " & DBDATE(textCP05_1) & " AND " & _
'                  "CP05 <= " & DBDATE(textCP05_2) & " AND " & _
'                  "CP10 = '1001' "
   'Modify by Morgan 2004/7/13   '加閉卷不印，只要申請案
   'Modify by Morgan 2004/7/14   '加排序 PA02,PA03
   '2005/6/14 modify by sonia 加核准的改變原處分
   'strSQL = "SELECT CP43,CP01,CP02,CP03,CP04, CP05 FROM CASEPROGRESS A, PATENT " & _
   '         "WHERE CP01='FCP' AND CP05 >= " & DBDATE(textCP05_1) & " AND " & _
   '               "CP05 <= " & DBDATE(textCP05_2) & " AND " & _
   '               "CP10 = '1001' AND EXISTS(SELECT * FROM CASEPROGRESS B WHERE B.CP09=A.CP43 AND B.CP10 IN ('101','102','103','104','105','107','301','302','303','304','305','306','307'))" & _
   '               " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL order by CP01,CP02,CP03,CP04"
   'Modified by Morgan 2012/12/5 +判斷要有發文日(配合發明初審核准定稿控制)
   'Modified by Morgan 2013/11/21 +衍生設計(125)
   strSql = "SELECT CP43,CP01,CP02,CP03,CP04,CP05,PA75,PA26,PA08,CP09,pa27,pa28,pa29,pa30,GetEmailFlag(CP09) eMail,CP27,CP10" & _
             " FROM CASEPROGRESS A,PATENT" & _
            " WHERE CP01='FCP'"
   If optSel(0).Value = True Then '整批
      'Add By Sindy 2015/10/22
'      If Replace(Trim(optSel(0).Caption), "：", "") = "核准發文日" Then
         strSql = strSql & " AND CP27>=" & DBDATE(textCP05_1) & " AND CP27<=" & DBDATE(textCP05_2)
'      Else
'      '2015/10/22 END
'         strSql = strSql & " AND CP05>=" & DBDATE(textCP05_1) & " AND CP05<=" & DBDATE(textCP05_2)
'      End If
   End If
   strSql = strSql & " and (CP10 = '1001' OR (CP10 = '1503' AND CP24='1')) AND EXISTS(SELECT * FROM CASEPROGRESS B WHERE B.CP09=A.CP43 AND B.CP10 IN ('101','102','103','104','105','107','125','301','302','303','304','305','306','307','308','309'))" & _
                     " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND cp27>0"
   If optSel(1).Value = True Then '本所案號
      strSql = strSql & " AND CP01='" & m_PA01 & "' AND CP02='" & m_PA02 & "' AND CP03='" & m_PA03 & "' AND CP04='" & m_PA04 & "'"
   End If
   strSql = strSql & " order by eMail,CP01,CP02,CP03,CP04"
   '2005/6/14 END
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'Add By Sindy 2015/7/15
   If bolQueryCnt = True Then
      strHadCnt = rsTmp.RecordCount
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   '2015/7/15 END
   
   'Add By Sindy 2017/6/9
   '檢查是否有安裝PDFCreator
   PrinterIndex = -1
   For i = 0 To Printers.Count - 1
    If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
     PrinterIndex = i
     Exit For
    End If
   Next i
   If PrinterIndex < 0 Then
      MsgBox "請通知電腦中心安裝PDFCreator !!!"
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   '2017/6/9 END
   strErrText = "檢查是否有安裝PDFCreator" & vbCrLf
   
   If rsTmp.RecordCount > 0 Then
      strErrText = strErrText & "rsTmp.RecordCount=" & rsTmp.RecordCount & vbCrLf
      'Add by Morgan 2011/7/8
      pub_OsPrinter = PUB_GetOsDefaultPrinter
      PUB_SetOsDefaultPrinter Combo2.Text
      PUB_SetWordActivePrinter
      'end 2011/7/8
      PUB_RestorePrinter Combo2.Text 'Add By Sindy 2015/6/26
      strErrText = strErrText & "開始" & vbCrLf
      InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/12/7
      bFind = True
      rsTmp.MoveFirst
      strStarTime = Format(ServerTime, "##:##:##") 'Add By Sindy 2015/6/26
      m_PA75 = ""
      Do While rsTmp.EOF = False
         m_CP09 = rsTmp.Fields("CP09") 'Added by Morgan 2012/11/8
         m_CP10 = rsTmp.Fields("CP10") 'Add by Sindy 2018/10/18
         ' 以相關總收文號列印定稿
         '92.2.27 ADD BY SONIA 未傳值定稿全部印中文
         m_PA01 = rsTmp.Fields("CP01")
         m_PA02 = rsTmp.Fields("CP02")
         m_PA03 = rsTmp.Fields("CP03")
         m_PA04 = rsTmp.Fields("CP04")
         'Add by Morgan 2004/6/23
         'Add By Sindy 2015/10/22
'         If Replace(Trim(optSel(0).Caption), "：", "") = "核准發文日" Then
            'Modified by Morgan 2019/7/30 目前已無用，改回紀錄來函收文日,新法核准函要用
            'm_CP05 = Val("" & rsTmp.Fields("CP27")) - 19110000
            m_CP05 = rsTmp.Fields("CP05")
            'end 2019/7/30
'         Else
'         '2015/10/22 END
'            m_CP05 = Val("" & rsTmp.Fields("CP05")) - 19110000
'         End If
         '92.2.27 END
         strErrText = strErrText & "m_CP09=" & m_CP09 & vbCrLf
         
         'Add By Sindy 2017/1/12 讀取核准案的"通知告准" D類進度之CP48承辦期限
         strCP48 = ""
         strExc(0) = "select cp48 from Caseprogress" & _
                     " where cp43='" & m_CP09 & "' and cp10='1917' and cp57 is null" & _
                     " order by cp09 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCP48 = Trim("" & RsTemp.Fields(0))
         End If
         '2017/1/12 END
         
         'Add By Sindy 2015/12/2 檢查有無單筆Run過
         bolSingleRun = False
         strExc(0) = "select * from CaseUseMemo" & _
                     " where CUM01='" & rsTmp.Fields("CP01").Value & "'" & _
                     " and CUM02='" & rsTmp.Fields("CP02").Value & "'" & _
                     " and CUM03='" & rsTmp.Fields("CP03").Value & "'" & _
                     " and CUM04='" & rsTmp.Fields("CP04").Value & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            bolSingleRun = True
         End If
         '2015/12/2 END
         'Modify By Sindy 2015/12/2 + Or bolSingleRun = False
         If InStr(strTmp, m_PA01 & m_PA02 & m_PA03 & m_PA04) = 0 Or bolSingleRun = False Then 'Add By Sindy 2014/4/18 +if
            m_PA75 = "" & rsTmp.Fields("PA75") 'Add by Morgan 2007/5/8
            m_PA26 = "" & rsTmp.Fields("PA26") 'Add by Morgan 2011/6/27
            'Added by Morgan 2014/7/24
            m_PA27 = "" & rsTmp.Fields("PA27")
            m_PA28 = "" & rsTmp.Fields("PA28")
            m_PA29 = "" & rsTmp.Fields("PA29")
            m_PA30 = "" & rsTmp.Fields("PA30")
            'end 2014/7/24
            m_PA08 = "" & rsTmp.Fields("PA08") 'Add by Morgan 2011/7/4
            'Add By Sindy 2015/6/30 只列印承辦單
            If Check1.Value = 1 Then
               strErrText = strErrText & "PUB_PrintFCPEmpBill : 1" & vbCrLf
               'Modified by Lydia 2019/03/04 更換類別代號;04=>01
               Call PUB_PrintFCPEmpBill(m_PA01, m_PA02, m_PA03, m_PA04, "01", m_CP09)
            Else
               If IsNull(rsTmp.Fields("CP43")) = False Then
                  If IsEmptyText(rsTmp.Fields("CP43")) = False Then
                     'Add By Sindy 2015/6/17 列印FCP承辦單
                     strErrText = strErrText & "PUB_PrintFCPEmpBill : 2" & vbCrLf
                     'Modified by Lydia 2019/03/04 更換類別代號;04=>01
                     Call PUB_PrintFCPEmpBill(m_PA01, m_PA02, m_PA03, m_PA04, "01", m_CP09)

                     'Add by Morgan 2004/7/27
                     'Modify by Morgan 2006/6/2
                     'm_LetterLanguage = GetLetterLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
                     m_LetterLanguage = PUB_GetLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
                     'Add by Morgan 2008/3/24 判斷是否產生電子檔
                     bolEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , bolPlusPaper)
                     'Added by Morgan 2014/6/3
                     If bolEmail = False Then
                        bolDNEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , bolDNPlusPaper, , True)
                     Else
                        bolDNEmail = bolEmail
                        bolDNPlusPaper = bolPlusPaper
                     End If
                     'end 2014/6/3

                     'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
                     If bolPlusPaper Then
                        iCopy = 0
                     Else
                        iCopy = 1
                     End If
                     'end 2009/10/20

                     'Added by Morgan 2018/6/12 -- Lina
                     '代理人 Y45814000 BASF SE Global Intellectual Property 告准函不論是否有前款未清 , 皆正常告准, 不增加欠款未付段落
                     If m_PA75 = "Y45814000" Then
                        m_bPrintBill = False
                        strBillNo = ""
                     Else
                     'end 2018/6/12
                     
                        'Add by Morgan 2011/6/23
                        m_bPrintBill = PUB_GetUnPaidBill(m_PA01, m_PA02, m_PA03, m_PA04, strBillNo)
                        strErrText = strErrText & "m_bPrintBill : " & m_bPrintBill & " strBillNo=" & strBillNo & vbCrLf
                        '列印請款單
                        If strBillNo <> "" Then
                           'Modified by Morgan 2014/6/3
                           'PUB_PrintBill strBillNo, Combo2.Text, bolEmail, bolPlusPaper, Me.Name, m_iBillPageCount, 2
                           PUB_PrintBill strBillNo, Combo2.Text, bolDNEmail, bolDNPlusPaper, Me.Name, m_iBillPageCount, 2
                        End If
                        'end 2011/6/23
                        
                     End If 'Added by Morgan 2018/6/12
                     
                     PrintLetter rsTmp.Fields("CP43"), m_LetterLanguage, bolEmail, iCopy
                     PUB_PrintLetter rsTmp.Fields("CP43") '列印通知函 Add by Morgan 2011/6/23
                     
                     'Add By Sindy 2017/1/12 定稿產生時，同時將畫面上之定稿日期更新至該案"通知告准" D類進度之CP85(FCP定稿日期)，以便下一功能可整批上發文日
'                     strSql = "Update Caseprogress set cp85=" & DBDATE(txtLetterDate) & _
'                              " where cp43='" & m_CP09 & "' and cp10='1917' and cp85 is null and cp27||cp57 is null"
                     'Modify By Sindy 2017/3/2 Bobbie說有時會做1次以上的產生定稿,要記錄最後一次的定稿日期
                     strSql = "Update Caseprogress set cp85=" & DBDATE(txtLetterDate) & _
                              " where cp09=(select max(cp09) from Caseprogress where cp43='" & m_CP09 & "' and cp10='1917' and cp27||cp57 is null)"
                     cnnConnection.Execute strSql
                     '2017/1/12 END
                     
                     If Not bolEmail Or bolPlusPaper Then
'                        'Add By Sindy 2015/9/21 日文定稿才要印地址條
'                        If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'                        '2015/9/21 END
                           '新增地址條列表資料
                           pub_AddressListSN = pub_AddressListSN + 1
                           PUB_AddNewAddressList strUserNum, "" & rsTmp.Fields("CP01").Value, "" & rsTmp.Fields("CP02").Value, "" & rsTmp.Fields("CP03").Value, "" & rsTmp.Fields("CP04").Value, "" & pub_AddressListSN, "0"
'                        End If
                     End If

                     'Add By Sindy 2015/6/18
                     If optSel(1).Value = True Then '執行單筆本所案號時,記錄下來
                        strExc(0) = "select * from CaseUseMemo" & _
                           " where CUM01='" & rsTmp.Fields("CP01").Value & "'" & _
                           " and CUM02='" & rsTmp.Fields("CP02").Value & "'" & _
                           " and CUM03='" & rsTmp.Fields("CP03").Value & "'" & _
                           " and CUM04='" & rsTmp.Fields("CP04").Value & "'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 0 Then
                           strSql = "INSERT INTO CaseUseMemo(CUM01,CUM02,CUM03,CUM04,CUM05)" & _
                                    " VALUES('" & rsTmp.Fields("CP01").Value & "','" & rsTmp.Fields("CP02").Value & "','" & rsTmp.Fields("CP03").Value & "','" & rsTmp.Fields("CP04").Value & "','01')"
                           cnnConnection.Execute strSql
                        End If
                     End If

                     '新增整批定稿列印清單資料
                     'Modify by Morgan 2004/12/27 因列印要依收文日，條件前8字元放 收文日&":"
                     'Modified by Morgan 2012/12/5 +日文定稿標記
                     'Modified by Morgan 2014/9/26 +e化註記
                     'Modify By Sindy 2017/1/12 Right("0" & m_CP05 & ":", 8) ==> Trim(m_CP09) & ":"
                     'PUB_AddNewLetterList "核准函", Right("0" & m_CP05 & ":", 8) & Me.textCP05_1.Text & "-" & Me.textCP05_2.Text, "" & rsTmp.Fields("CP01").Value, "" & rsTmp.Fields("CP02").Value, "" & rsTmp.Fields("CP03").Value, "" & rsTmp.Fields("CP04").Value, IIf(InStr(strTmp, m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0, "■", "　") & IIf(InStr(m_WithReportList, m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0, "○", "　") & IIf(m_bolDivSug = True, "★", "") & IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "") & IIf(m_LetterLanguage = "3", "日", "")
                     PUB_AddNewLetterList "核准函", Trim(m_CP09) & ":" & IIf(strCP48 = "", "00000000", strCP48) & ":" & Me.textCP05_1.Text & "-" & Me.textCP05_2.Text, "" & rsTmp.Fields("CP01").Value, "" & rsTmp.Fields("CP02").Value, "" & rsTmp.Fields("CP03").Value, "" & rsTmp.Fields("CP04").Value, IIf(InStr(strTmp, m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0, "■", "　") & IIf(InStr(m_WithReportList, m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0, "○", "　") & IIf(m_bolDivSug = True, "★", "") & IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "") & IIf(m_LetterLanguage = "3", "日", "")
                  End If
               End If
            End If
         Else
            'Modify By Sindy 2015/7/14
            'Modify By Sindy 2017/1/12 Right("0" & m_CP05 & ":", 8) ==> Trim(m_CP09) & ":"
            'PUB_AddNewLetterList "核准函", Right("0" & m_CP05 & ":", 8) & Me.textCP05_1.Text & "-" & Me.textCP05_2.Text, "" & rsTmp.Fields("CP01").Value, "" & rsTmp.Fields("CP02").Value, "" & rsTmp.Fields("CP03").Value, "" & rsTmp.Fields("CP04").Value, "■"
            PUB_AddNewLetterList "核准函", Trim(m_CP09) & ":" & IIf(strCP48 = "", "00000000", strCP48) & ":" & Me.textCP05_1.Text & "-" & Me.textCP05_2.Text, "" & rsTmp.Fields("CP01").Value, "" & rsTmp.Fields("CP02").Value, "" & rsTmp.Fields("CP03").Value, "" & rsTmp.Fields("CP04").Value, "■"
            '2015/7/14 END
         End If 'Add By Sindy 2014/4/18 +end
         rsTmp.MoveNext
      Loop
      PUB_SetOsDefaultPrinter pub_OsPrinter 'Add by Morgan 2011/6/23
      PUB_RestorePrinter strPrinter2 'Add By Sindy 2015/6/26
      'Add By Sindy 2015/6/26
      If optSel(0).Value = True Then '整批
         'Add By Sindy 2015/7/16 紀錄在資料庫,否則換電腦或使用者會讀不到
         PUB_SaveLastDate Me.Name, "DATE1", textCP05_1.Text
         PUB_SaveLastDate Me.Name, "DATE2", textCP05_2.Text
         Label2(2).Caption = PUB_GetLastDate(Me.Name, "DATE1")
         Label2(3).Caption = PUB_GetLastDate(Me.Name, "DATE2")
         '2015/7/16 END
'         PUB_SendMail strUserNum, "97038", "", "外專執行＜核准函＞整批的執行時間: " & strStarTime & " ~ " & Format(ServerTime, "##:##:##"), "如主旨"
      End If
      '2015/6/26 END
      'Add By Sindy 2014/4/18
      For i = 0 To List1(1).ListCount - 1
         List1(1).RemoveItem 0
      Next
      '2014/4/18 END
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/7
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
   QueryLetterData = bFind
   
   Exit Function
   
ErrHnd:
'   'Add By Sindy 2013/1/28
'   If Err.Number = -2147217900 Then 'ORA-00917: 遺漏逗點
'      '寫Log
'      Call ReadTxt3(strSql)
'      '接著發生錯誤陳述式的下個陳述式開始執行
'      Resume Next
'   End If
'   '2013/1/28 End
'   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      PUB_SendMail strUserNum, "97038", "", "frm060316_1 核准函", strErrText & vbCrLf & vbCrLf & _
         Err.Number & vbCrLf & _
         Err.Description, , , , , , , , , , True, False, , , False
      MsgBox Err.Number & vbCrLf & _
             Err.Description & vbCrLf & vbCrLf & _
             IIf(strExc(0) <> "", "strExc(0)=" & strExc(0) & vbCrLf & vbCrLf, "") & _
             IIf(strSql <> "", "strSql=" & strSql & vbCrLf & vbCrLf, ""), vbCritical
   End If
End Function

Private Sub QueryOtherData()
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim bFind As Boolean
Dim strData As String
   
   Screen.MousePointer = vbHourglass
   
   '本所案號
   m_PA01 = textPA01
   m_PA02 = textPA02
   m_PA03 = textPA03
   m_PA04 = textPA04
   
   '本次不適用案號
   If optSel(0).Value = True Then
      strSql = "SELECT CP43,CP01,CP02,CP03,CP04,CP05,PA75,PA26,PA08,CP09,pa27,pa28,pa29,pa30,GetEmailFlag(CP09) eMail,cp148" & _
                " FROM CASEPROGRESS A,PATENT,CaseUseMemo" & _
               " WHERE CP01='FCP'"
      'Add By Sindy 2015/10/22
'      If Replace(Trim(optSel(0).Caption), "：", "") = "核准發文日" Then
         strSql = strSql & " AND CP27>=" & DBDATE(textCP05_1) & " AND CP27<=" & DBDATE(textCP05_2)
'      Else
'      '2015/10/22 END
'         strSql = strSql & " AND CP05>=" & DBDATE(textCP05_1) & " AND CP05<=" & DBDATE(textCP05_2)
'      End If
      strSql = strSql & " and (CP10 = '1001' OR (CP10 = '1503' AND CP24='1')) AND EXISTS(SELECT * FROM CASEPROGRESS B WHERE B.CP09=A.CP43 AND B.CP10 IN ('101','102','103','104','105','107','125','301','302','303','304','305','306','307','308','309'))" & _
                        " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND cp27>0"
      strSql = strSql & " AND CP01=CUM01 AND CP02=CUM02 AND CP03=CUM03 AND CP04=CUM04 AND CUM05='01'"
      strSql = strSql & " order by eMail,CP01,CP02,CP03,CP04"
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            bFind = False
            strData = rsTmp.Fields("CP01") & rsTmp.Fields("CP02") & rsTmp.Fields("CP03") & rsTmp.Fields("CP04")
            For i = 0 To List1(1).ListCount - 1
               If List1(1).List(i) = strData Then
                  bFind = True: Exit For
               End If
            Next i
            If bFind = False Then
               List1(1).AddItem strData
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      'Add By Sindy 2015/7/14
      strSql = "SELECT CP43,CP01,CP02,CP03,CP04,CP05,PA75,PA26,PA08,CP09,pa27,pa28,pa29,pa30,GetEmailFlag(CP09) eMail,cp148" & _
                " FROM CASEPROGRESS A,PATENT" & _
               " WHERE CP01='FCP'"
      'Add By Sindy 2015/10/22
'      If Replace(Trim(optSel(0).Caption), "：", "") = "核准發文日" Then
         strSql = strSql & " AND CP27>=" & DBDATE(textCP05_1) & " AND CP27<=" & DBDATE(textCP05_2)
'      Else
'      '2015/10/22 END
'         strSql = strSql & " AND CP05>=" & DBDATE(textCP05_1) & " AND CP05<=" & DBDATE(textCP05_2)
'      End If
      strSql = strSql & " and (CP10 = '1001' OR (CP10 = '1503' AND CP24='1')) AND EXISTS(SELECT * FROM CASEPROGRESS B WHERE B.CP09=A.CP43 AND B.CP10 IN ('101','102','103','104','105','107','125','301','302','303','304','305','306','307','308','309'))" & _
                        " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND cp27>0"
      strSql = strSql & " AND (instr(cp64,'優先告准')>0 or instr(cp64,'分割期限')>0)"
      strSql = strSql & " order by eMail,CP01,CP02,CP03,CP04"
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            bFind = False
            strData = rsTmp.Fields("CP01") & rsTmp.Fields("CP02") & rsTmp.Fields("CP03") & rsTmp.Fields("CP04")
            For i = 0 To List1(1).ListCount - 1
               If List1(1).List(i) = strData Then
                  bFind = True: Exit For
               End If
            Next i
            If bFind = False Then
               List1(1).AddItem strData
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      Call cmdCount_Click
      '2015/7/14 END
   End If
   
   '有檢索案號
   strSql = "SELECT CP43,CP01,CP02,CP03,CP04,CP05,PA75,PA26,PA08,CP09,pa27,pa28,pa29,pa30,GetEmailFlag(CP09) eMail,cp148" & _
             " FROM CASEPROGRESS A,PATENT" & _
            " WHERE CP01='FCP'"
   If optSel(0).Value = True Then '整批
      'Add By Sindy 2015/10/22
'      If Replace(Trim(optSel(0).Caption), "：", "") = "核准發文日" Then
         strSql = strSql & " AND CP27>=" & DBDATE(textCP05_1) & " AND CP27<=" & DBDATE(textCP05_2)
'      Else
'      '2015/10/22 END
'         strSql = strSql & " AND CP05>=" & DBDATE(textCP05_1) & " AND CP05<=" & DBDATE(textCP05_2)
'      End If
   End If
   strSql = strSql & " and (CP10 = '1001' OR (CP10 = '1503' AND CP24='1')) AND EXISTS(SELECT * FROM CASEPROGRESS B WHERE B.CP09=A.CP43 AND B.CP10 IN ('101','102','103','104','105','107','125','301','302','303','304','305','306','307','308','309'))" & _
                     " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND cp27>0"
   If optSel(1).Value = True Then '本所案號
      strSql = strSql & " AND CP01='" & m_PA01 & "' AND CP02='" & m_PA02 & "' AND CP03='" & m_PA03 & "' AND CP04='" & m_PA04 & "'"
   End If
   strSql = strSql & " and cp148='Y'" '有檢索
   strSql = strSql & " order by eMail,CP01,CP02,CP03,CP04"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         bFind = False
         strData = rsTmp.Fields("CP01") & rsTmp.Fields("CP02") & rsTmp.Fields("CP03") & rsTmp.Fields("CP04")
         For i = 0 To List1(0).ListCount - 1
            If List1(0).List(i) = strData Then
               bFind = True: Exit For
            End If
         Next i
         If bFind = False Then
            List1(0).AddItem strData
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Sub

'' 初始化列表
'Public Sub InitialGridList()
'   grdList.Clear
'   grdList.Rows = 1
'   grdList.Cols = 7
'
'   grdList.ColWidth(0) = 300
'   grdList.row = 0
'
'   grdList.col = 0
'   grdList.ColAlignment(0) = flexAlignCenterCenter
'   grdList.col = 1
'   grdList.Text = "收文日"
'   grdList.ColWidth(1) = 1000
'   grdList.ColAlignment(1) = flexAlignCenterCenter
'   grdList.col = 2
'   grdList.Text = "總收文號"
'   grdList.ColWidth(2) = 1200
'   grdList.ColAlignment(2) = flexAlignCenterCenter
'   grdList.col = 3
'   grdList.Text = "案件性質"
'   grdList.ColWidth(3) = 1400
'   grdList.ColAlignment(3) = flexAlignLeftCenter
'   grdList.col = 4
'   grdList.Text = "結果"
'   grdList.ColWidth(4) = 1000
'   grdList.ColAlignment(4) = flexAlignCenterCenter
'   grdList.col = 5
'   grdList.Text = "相關總收文號"
'   grdList.ColWidth(5) = 1400
'   grdList.ColAlignment(5) = flexAlignCenterCenter
'   grdList.col = 6
'   grdList.Text = "相關人"
'   grdList.ColWidth(6) = 2000
'   grdList.ColAlignment(6) = flexAlignLeftCenter
'End Sub
'
'' 讀取專利基本檔
'Private Function ReadPatentData() As Boolean
'   Dim strSql As String
'   Dim rsTmp As ADODB.Recordset
'   Dim bFind As Boolean
'   bFind = False
'   strSql = "SELECT * FROM PATENT " & _
'            "WHERE PA01 = '" & m_PA01 & "' AND " & _
'                  "PA02 = '" & m_PA02 & "' AND " & _
'                  "PA03 = '" & m_PA03 & "' AND " & _
'                  "PA04 = '" & m_PA04 & "' "
'   Set rsTmp = New ADODB.Recordset
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      bFind = True
'      m_PA09 = "" & rsTmp.Fields("PA09")
'      m_PA75 = "" & rsTmp.Fields("PA75")
'      m_PA26 = "" & rsTmp.Fields("PA26") 'Add by Morgan 2011/6/27
'      'Added by Morgan 2014/7/24
'      m_PA27 = "" & rsTmp.Fields("PA27")
'      m_PA28 = "" & rsTmp.Fields("PA28")
'      m_PA29 = "" & rsTmp.Fields("PA29")
'      m_PA30 = "" & rsTmp.Fields("PA30")
'      'end 2014/7/24
'      m_PA08 = "" & rsTmp.Fields("PA08") 'Add by Morgan 20111/7/1
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'   ReadPatentData = bFind
'End Function
'
'Private Function ListData() As Boolean
'   Dim strSql As String
'   Dim rsTmp As ADODB.Recordset
'   Dim bDeal As Boolean
'   Dim bFind As Boolean
'
'   bFind = False
'   InitialGridList
'
'   If m_PA09 < "010" Then
'      strSql = "SELECT CP05,CP09,CPM03 AS CP10,CP24,CP37,CP38,CP39,CP43,CP50,CP51,CP52,CP56 FROM CASEPROGRESS, CASEPROPERTYMAP " & _
'               "WHERE CP01 = '" & m_PA01 & "' AND " & _
'                     "CP02 = '" & m_PA02 & "' AND " & _
'                     "CP03 = '" & m_PA03 & "' AND " & _
'                     "CP04 = '" & m_PA04 & "' AND " & _
'                     "CP01 = CPM01(+) AND " & _
'                     "CP10 = CPM02(+) "
'   Else
'      strSql = "SELECT CP05,CP09,CPM04 AS CP10,CP24,CP37,CP38,CP39,CP43,CP50,CP51,CP52,CP56 FROM CASEPROGRESS, CASEPROPERTYMAP " & _
'               "WHERE CP01 = '" & m_PA01 & "' AND " & _
'                     "CP02 = '" & m_PA02 & "' AND " & _
'                     "CP03 = '" & m_PA03 & "' AND " & _
'                     "CP04 = '" & m_PA04 & "' AND " & _
'                     "CP01 = CPM01(+) AND " & _
'                     "CP10 = CPM02(+) "
'   End If
'   Set rsTmp = New ADODB.Recordset
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/12/7
'      Do While rsTmp.EOF = False
'         ' 列入資料
'         grdList.Rows = grdList.Rows + 1
'         grdList.row = grdList.Rows - 1
'
'         bFind = True
'         ' 收文日
'         If IsNull(rsTmp.Fields("CP05")) = False Then
'            grdList.TextMatrix(grdList.row, 1) = TAIWANDATE(rsTmp.Fields("CP05"))
'         End If
'         ' 收文號
'         If IsNull(rsTmp.Fields("CP09")) = False Then
'            grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("CP09")
'         End If
'         ' 案件性質
'         If IsNull(rsTmp.Fields("CP10")) = False Then
'            grdList.TextMatrix(grdList.row, 3) = rsTmp.Fields("CP10")
'         End If
'         ' 結果
'         If IsNull(rsTmp.Fields("CP24")) = False Then
'            Select Case rsTmp.Fields("CP24")
'               Case "1":
'                  grdList.TextMatrix(grdList.row, 4) = "准勝"
'               Case "2":
'                  grdList.TextMatrix(grdList.row, 4) = "駁敗"
'               Case Else:
'            End Select
'         End If
'         ' 相關總收文號
'         If IsNull(rsTmp.Fields("CP43")) = False Then
'            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("CP43")
'         End If
'         ' 相關人
'         bDeal = False
'         If bDeal = False And IsNull(rsTmp.Fields("CP37")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP37")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP37")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP38")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP38")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP38")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP39")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP39")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP39")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP50")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP51")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP52")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
'               bDeal = True
'            End If
'         End If
'         If bDeal = False Then: grdList.TextMatrix(grdList.row, 6) = Empty
'NextRecord:
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   ListData = bFind
'   Set rsTmp = Nothing
'End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim Cancel As Boolean
   Dim nResponse
   CheckDataValid = False
   
   ' 選項
   If optSel(0).Value = True Then
      ' 來函收文日不可空白
      If IsEmptyText(textCP05_1) = True Or IsEmptyText(textCP05_2) = True Then
         strTit = "檢核資料"
         strMsg = Replace(Trim(optSel(0).Caption), "：", "") & "不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         If IsEmptyText(textCP05_1) = True Then
            textCP05_1.SetFocus
         Else
            textCP05_2.SetFocus
         End If
         GoTo EXITSUB
      End If
      'Add By Cheng 2002/03/20
      If PUB_CheckKeyInDate(Me.textCP05_1) = -1 Then
         Me.textCP05_1.SetFocus
         textCP05_1_GotFocus
         GoTo EXITSUB
      End If
      If PUB_CheckKeyInDate(Me.textCP05_2) = -1 Then
         Me.textCP05_2.SetFocus
         textCP05_2_GotFocus
         GoTo EXITSUB
      End If
      
      ' 範圍
      If Val(textCP05_1) > Val(textCP05_2) Then
         strTit = "檢核資料"
         strMsg = Replace(Trim(optSel(0).Caption), "：", "") & "範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textCP05_1.SetFocus
         GoTo EXITSUB
      End If
   Else
      ' 本所案號
      If IsEmptyText(textPA01) = True Then
         strTit = "檢核資料"
         strMsg = "本所案號系統類別不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA01.SetFocus
         GoTo EXITSUB
      End If
      ' 本所案號
      If IsEmptyText(textPA02) = True Then
         strTit = "檢核資料"
         strMsg = "本所案號流水號不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA02.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add by Morgan 2009/8/20
   If txtLetterDate <> "" Then
      txtLetterDate_Validate Cancel
      If Cancel = True Then GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   g_LetterDate = "" 'Add by Morgan 2009/8/20
   
   'Copy from cmdExit_Click by Morgan 2004/10/26
   '列印定稿整批列印清單
   PUB_PrintLetterList strUserNum, "2", Combo2, strPrinter2
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, "and LL02='核准函' "
   
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2004/10/26 end
   'Add by Morgan 2011/6/23
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   Set frm060316_1 = Nothing
End Sub

Private Sub optSel_Click(Index As Integer)
   If optSel(0).Value = True Then
      textCP05_1.SetFocus
      cmdQuery.Enabled = False
      EnableTextBox textCP05_1, True
      EnableTextBox textCP05_2, True
      EnableTextBox textPA01, False
      EnableTextBox textPA02, False
      EnableTextBox textPA03, False
      EnableTextBox textPA04, False
   Else
      textPA01.SetFocus
      cmdQuery.Enabled = True
      EnableTextBox textCP05_1, False
      EnableTextBox textCP05_2, False
      EnableTextBox textPA01, True
      EnableTextBox textPA02, True
      EnableTextBox textPA03, True
      EnableTextBox textPA04, True
   End If
End Sub

'Private Sub grdList_SelChange()
'   If grdList.row > 0 Then
'      'Add by Morgan 2004/6/24
'      m_CP05 = grdList.TextMatrix(grdList.row, 1)
'      m_CP09 = grdList.TextMatrix(grdList.row, 2)
'   End If
'   grdList_ShowSelection
'End Sub
'
'' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
'Private Sub grdList_ShowSelection()
'   Dim nCurrSel As Integer
'   Dim nCol As Integer
'
'   nCurrSel = grdList.row
'
'   ' 與前一選擇的列位置相同則不處理
'   If m_CurrSel = grdList.row Then
'      Dim nOldCol As Integer
'      nOldCol = grdList.col
'      grdList.col = 1
'      If grdList.CellBackColor <> &H8000000D Then
'         For nCol = 1 To grdList.Cols - 1
'            grdList.col = nCol
'            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
'            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
'         Next nCol
'      End If
'      grdList.col = nOldCol
'      GoTo EXITSUB
'   End If
'
'   ' 將原先選取的列回復到正常的顏色
'   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
'      grdList.row = m_CurrSel
'      grdList.col = 1
'      If grdList.CellBackColor <> &H80000005 Then
'         For nCol = 1 To grdList.Cols - 1
'            grdList.col = nCol
'            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
'            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
'         Next nCol
'      End If
'      grdList.col = 0
'   End If
'   ' 設定成所選取的列
'   m_CurrSel = nCurrSel
'   ' 將所選取的列反白
'   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
'      grdList.row = m_CurrSel
'      grdList.col = 1
'      For nCol = 1 To grdList.Cols - 1
'         grdList.col = nCol
'         grdList.CellBackColor = &H8000000D
'         grdList.CellForeColor = &H80000005
'      Next nCol
'      grdList.col = 0
'   End If
'EXITSUB:
'End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2014/4/18
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 來函收文日(起)
Private Sub textCP05_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05_1) = False Then
      If CheckIsTaiwanDate(textCP05_1, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = Replace(Trim(optSel(0).Caption), "：", "") & "(起)日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textCP05_1_GotFocus
      End If
   End If
End Sub

' 來函收文日(迄)
Private Sub textCP05_2_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textCP05_2) = False Then
      If CheckIsTaiwanDate(textCP05_2, False) = False Then
         strTit = "檢核資料"
         strMsg = Replace(Trim(optSel(0).Caption), "：", "") & "(迄)日期格式不正確"
         textCP05_2.SetFocus
         InverseTextBox textCP05_2
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Exit Sub 'Add By Sindy 2015/6/18
      Else
         If Not ChkRange(textCP05_1, textCP05_2, Replace(Trim(optSel(0).Caption), "：", "")) Then
            textCP05_1.SetFocus
            InverseTextBox textCP05_1
            Exit Sub 'Add By Sindy 2015/6/18
         End If
      End If
      Call QueryOtherData
   End If
End Sub

Private Sub textPA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA01) = False Then
      Select Case textPA01
         Case "FCP":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPA01_GotFocus
      End Select
   End If
End Sub

Private Sub textCP05_1_GotFocus()
   InverseTextBox textCP05_1
End Sub

Private Sub textCP05_2_GotFocus()
   InverseTextBox textCP05_2
End Sub

Private Sub textPA01_GotFocus()
   InverseTextBox textPA01
End Sub

Private Sub textPA02_GotFocus()
   InverseTextBox textPA02
End Sub

Private Sub textPA02_LostFocus()
   If textPA02.Text <> "" Then
      If textPA03 = "" Then textPA03 = "0"
      If textPA04 = "" Then textPA04 = "00"
   End If
End Sub

Private Sub textPA03_GotFocus()
   InverseTextBox textPA03
End Sub

Private Sub textPA03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPA04_GotFocus()
   InverseTextBox textPA04
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
'Modify by Morgan 2004/7/27
'加定稿語文參數
Private Sub InsExpField(ByVal strCP09 As String, Optional ByVal stLetterLanguage As String = "2")
Dim strSql As String
Dim strTemp As String
'Add By Cheng 2003/01/24
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
       
    'Add By Cheng 2003/01/24
    '若為英文定稿
    'Modify by Morgan 2004/7/27
    'If GetLetterLanguage(m_PA01, m_PA02, m_PA03, m_PA04) = "2" Then
    If stLetterLanguage = "2" Then
        StrSQLa = "Select COUNT(*) From PriDate Where PD01='" & m_PA01 & "' AND PD02='" & m_PA02 & "' AND PD03='" & m_PA03 & "' AND PD04='" & m_PA04 & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If rsA.Fields(0).Value > 0 Then
                m_blnPriData = True
                '若有三個以上優先權資料
                If rsA.Fields(0).Value >= 3 Then
                    m_bln3PriData = True
                Else
                    m_bln3PriData = False
                End If
            Else
                m_blnPriData = False
                m_bln3PriData = False
            End If
        Else
            m_blnPriData = False
            m_bln3PriData = False
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
   ' 定稿語文
   'Modified by Lydia 2022/10/05 統一使用模組
   'Select Case GetLetterLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
   Select Case PUB_GetLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
      ' 中文
      Case "1":
         ' 清除定稿例外欄位檔原有資料
         EndLetter "04", strCP09, "01", strUserNum
         
'Removed by Morgan 2012/8/30 舊定稿已刪除
'      ' 英文
'      Case "2":
'         'Add by Morgan 2004/6/23
'         '93.7.1以後用二合一定稿
'         If Val(m_CP05) < 930701 Then
'           'Modify By Cheng 2003/01/24
'           '依優先權個數出不同定稿
'   '         EndLetter "04", strCP09, "02", strUserNum
'           '若有三個優先權資料
'           If m_bln3PriData = True Then
'                EndLetter "04", strCP09, "05", strUserNum
'               'Add By Cheng 2003/02/14
'               '附件
'                EndLetter "04", strCP09, "08", strUserNum
'            '若有優先權資料
'            ElseIf m_blnPriData = True Then
'                EndLetter "04", strCP09, "02", strUserNum
'               'Add By Cheng 2003/02/14
'               '附件
'                EndLetter "04", strCP09, "06", strUserNum
'            '若無優先權資料
'            Else
'                EndLetter "04", strCP09, "04", strUserNum
'               'Add By Cheng 2003/02/14
'               '附件
'                EndLetter "04", strCP09, "07", strUserNum
'            End If
'         End If

      ' 日文
      Case "3":
         EndLetter "04", strCP09, "03", strUserNum
      Case Else:
   End Select
End Sub
'Modify by Morgan 2004/7/27 加定稿語文參數
Private Sub PrintLetter(ByVal strCP09 As String, Optional ByVal stLetterLanguage As String = "2", Optional bolByEmail As Boolean, Optional iCopys As Integer)
   'Add by Morgan 2004/7/27
   Dim stET03 As String
   Dim stContent As String
   Dim stPS As String 'Add by Morgan 2011/1/14
   'Add By Sindy 2017/6/8
   Dim strFolder As String, strFileName As String
   Dim strTmpFileName As String, strTmpFileName2 As String
   Dim strMergeFN As String, strCmd As String, strMergeName As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim fs As Object
   '2017/6/8 END
   Dim stLtKind As String '特殊定稿控制 Added by Morgan 2019/7/30
   
On Error GoTo ErrHnd2 'Add By Sindy 2019/8/26

   'stPS = GetPS(m_PA01, m_PA02, m_PA03, m_PA04) 'Add by Morgan 2011/1/14 'Removed by Morgan 2016/11/29 併入 StartLetter
               
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField strCP09
   
   ' 定稿語文
   'Modify by Morgan 2004/7/27
   'Select Case GetLetterLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
   Select Case stLetterLanguage
      ' 中文
      Case "1":
         stET03 = "01"
         'Added by Morgan 2016/3/31
         StartLetter "04", strCP09, stET03, m_PA01 & m_PA02 & m_PA03 & m_PA04, stPS, "98"
         ' 列印定稿
         'Modified by Morgan 2016/3/31
         'NowPrint strCP09, "04", "01", False, strUserNum, 0
         If bolByEmail = True Then
            NowPrint strCP09, "04", stET03, False, strUserNum, , , , , , , True, True
         Else
            NowPrint strCP09, "04", "98", False, strUserNum, , , , , 1
            NowPrint strCP09, "04", stET03, False, strUserNum
         End If
         
      ' 英文
      Case "2":
         'Add by Morgan 2004/6/23
         '93.7.1以後用二合一定稿
         '新法
'         If Val(m_CP05) >= 930701 Then
            '指示信 09~15
            stET03 = GetET03(m_PA01 & m_PA02 & m_PA03 & m_PA04)
            
            'Added by Morgan 2019/7/30
            '來函收文日108.8.1以後的英文定稿 一般(09)/自動代繳(10)/收款後辦案(12)/Y4526800(04) 定稿改用合併定稿(02)
            'If m_CP05 >= "20190801" Then 'Removed by Morgan 2022/11/23 舊定稿已刪除
               If stET03 = "09" Or stET03 = "10" Or stET03 = "12" Or stET03 = "04" Then
                  stLtKind = stET03
                  stET03 = "02"
               End If
            'End If
            'end 2019/7/30
   
            StartLetter "04", strCP09, stET03, m_PA01 & m_PA02 & m_PA03 & m_PA04, stPS, "98", stLtKind
            
            'Add by Morgan 2008/3/24 判斷是否產生電子檔
            If bolByEmail = True Then
               NowPrint strCP09, "04", stET03, False, strUserNum, , , , , iCopys
               '因為要印EMail資料,所以Save2File參數也要設
               NowPrint strCP09, "04", stET03, False, strUserNum, , , True, stContent, , , , True
            Else
            'End 2008/3/24
               'Add by Morgan 2006/2/13
               '英文定稿加傳真封面
               NowPrint strCP09, "04", "98", False, strUserNum, , , , , 1
               NowPrint strCP09, "04", stET03, False, strUserNum
            End If
            
            'Added by Morgan 2022/8/23 Murata以日譯文取代原本的英譯文--Bobbie
            'Modified by Morgan 2022/8/29 +Y53618--Arashi
            If InStr("Y27766000,Y45002000,Y52519000,Y33898000,Y28043000,Y53618000", m_PA75) > 0 And Left(m_PA26, 8) = "X2776600" Then
               '日譯文
               If m_blnPriData = True Then
                  stET03 = "18"
               Else
                  stET03 = "17"
               End If
            Else
            'end 2022/8/23
            
               '附件
               '若有三個優先權資料
               If m_bln3PriData = True Then
                  stET03 = "15"
               '若有優先權資料
               ElseIf m_blnPriData = True Then
                  stET03 = "13"
               '若無優先權資料
               Else
                  stET03 = "14"
               End If
               
            End If 'Added by Morgan 2022/8/24
            
            'Add by Morgan 2008/3/24 判斷是否產生電子檔
            If bolByEmail = True Then
               NowPrint strCP09, "04", stET03, False, strUserNum, , , , , iCopys
               NowPrint strCP09, "04", stET03, False, strUserNum, , stContent, , , , , True, True
            Else
            'end 2008/3/24
               NowPrint strCP09, "04", stET03, False, strUserNum
            End If
            
'Removed by Morgan 2012/8/30 舊定稿已刪除
'         '舊法
'         Else
'            '若有三個優先權資料
'            If m_bln3PriData = True Then
'               'Add by Morgan 2008/3/24 判斷是否產生電子檔
'               If bolByEmail = True Then
'                  NowPrint strCP09, "04", "05", False, strUserNum, , , , , iCopys
'                  NowPrint strCP09, "04", "05", False, strUserNum, , , True, stContent, , , , True
'                  '附件
'                  NowPrint strCP09, "04", "08", False, strUserNum, , , , , iCopys
'                  NowPrint strCP09, "04", "08", False, strUserNum, , stContent, , , , , True, True
'
'               Else
'               'end 2008/3/24
'                  NowPrint strCP09, "04", "05", False, strUserNum
'                  '附件
'                  NowPrint strCP09, "04", "08", False, strUserNum
'               End If
'            '若有優先權資料
'            ElseIf m_blnPriData = True Then
'               'Add by Morgan 2008/3/24 判斷是否產生電子檔
'               If bolByEmail = True Then
'                  NowPrint strCP09, "04", "02", False, strUserNum, , , , , iCopys
'                  NowPrint strCP09, "04", "02", False, strUserNum, , , True, stContent, , , , True
'                  '附件
'                  NowPrint strCP09, "04", "06", False, strUserNum, , , , , iCopys
'                  NowPrint strCP09, "04", "06", False, strUserNum, , stContent, , , , , True, True
'               Else
'               'end 2008/3/24
'                  NowPrint strCP09, "04", "02", False, strUserNum
'                  '附件
'                  NowPrint strCP09, "04", "06", False, strUserNum
'               End If
'            '若無優先權資料
'            Else
'               'Add by Morgan 2008/3/24 判斷是否產生電子檔
'               If bolByEmail = True Then
'                  NowPrint strCP09, "04", "04", False, strUserNum, , , , , iCopys
'                  NowPrint strCP09, "04", "04", False, strUserNum, , , True, stContent, , , , True
'                  '附件
'                  NowPrint strCP09, "04", "07", False, strUserNum, , , , , iCopys
'                  NowPrint strCP09, "04", "07", False, strUserNum, , stContent, , , , , True, True
'               Else
'               'end 2008/3/24
'                  NowPrint strCP09, "04", "04", False, strUserNum
'                  '附件
'                  NowPrint strCP09, "04", "07", False, strUserNum
'               End If
'            End If

'         End If
      ' 日文
      Case "3":
         stET03 = "03"
         'Modify by Morgan 2004/12/22
         'NowPrint strCP09, "04", stET03, False, strUserNum, 0
         '照英文抓法判斷是否自動代繳
         stET03 = GetET03(m_PA01 & m_PA02 & m_PA03 & m_PA04, "3")
         StartLetter "04", strCP09, stET03, m_PA01 & m_PA02 & m_PA03 & m_PA04, stPS, "98"
         'Add by Morgan 2008/3/24 判斷是否產生電子檔
         If bolByEmail = True Then
            NowPrint strCP09, "04", stET03, False, strUserNum, , , , , iCopys
            NowPrint strCP09, "04", stET03, False, strUserNum, , , True, stContent, , , , True
            '譯文
            If m_blnPriData = True Then
               stET03 = "18"
            Else
               stET03 = "17"
            End If
            NowPrint strCP09, "04", stET03, False, strUserNum, , , , , , iCopys
            NowPrint strCP09, "04", stET03, False, strUserNum, , stContent, , , , , True, True
         Else
         'end 2008/3/24
            'Add by Morgan 2006/3/15
            '加英文傳真封面
            NowPrint strCP09, "04", "98", False, strUserNum, , , , , 1
            '2006/3/15 end
            NowPrint strCP09, "04", stET03, False, strUserNum
            '譯文
            If m_blnPriData = True Then
               stET03 = "18"
            Else
               stET03 = "17"
            End If
            NowPrint strCP09, "04", stET03, False, strUserNum
            '2004/12/22
         End If
      Case Else:
   End Select
   
   'Add By Sindy 2017/6/8
   '英,日文E化案件要產生"譯文+電子公文"的pdf檔(\\typing2\fcp_workflow\Notice of Allowance with translation)
   If (stLetterLanguage = 2 Or stLetterLanguage = 3) And bolByEmail = True Then
      strErrText = strErrText & "Dir(m_AttachPath..." & vbCrLf
      If Dir(m_AttachPath & "\.") <> "" Then
         Kill m_AttachPath & "\*.*"
      End If
      strErrText = strErrText & "Kill m_AttachPath" & vbCrLf
      
      PUB_SetOsDefaultPrinter Printers(PrinterIndex).DeviceName 'Printer.DeviceName '作業系統預設印表機指到PDFCreator
      PUB_SetWordActivePrinter
      strErrText = strErrText & "呼叫定稿..." & vbCrLf
      '呼叫定稿
      strUserLevel = "發FC郵件"
      NowPrint strCP09, "04", stET03, True, strUserNum, , , , , , , True, , False, , , , , True
      strUserLevel = ""
      '轉PDF
      If Pub_StrUserSt03 = "M51" Then
         strFolder = PUB_Getdesktop
      Else
         'Modified by Lydia 2024/07/22 改用變數
         'strFolder = "\\typing2\fcp_workflow\Notice of Allowance with translation"
         strFolder = "\\" & strTyping2Path & "\fcp_workflow\Notice of Allowance with translation"
      End If
      strFileName = m_PA01 & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", m_PA03 & m_PA04, "") & EfileNameFCP_04
      strTmpFileName = m_PA01 & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", m_PA03 & m_PA04, "") & "_Letter"
      strErrText = strErrText & "strFileName:" & strFileName & vbCrLf
      strErrText = strErrText & "strTmpFileName:" & strTmpFileName & vbCrLf
      frmPDF.Show
      frmPDF.StartProcess m_AttachPath, strTmpFileName
      '切換印表機
      'g_WordAp.Visible = True
      If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord 'Added by Morgan 2019/2/13
      g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
      g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
      DoEvents
      frmPDF.EndtProcess
      Unload frmPDF
      g_WordAp.Quit wdDoNotSaveChanges
      Set g_WordAp = Nothing
      
      '切回欲列印的印表機
      PUB_SetOsDefaultPrinter Combo2.Text
      PUB_RestorePrinter Combo2.Text
      strErrText = strErrText & "切回欲列印的印表機" & vbCrLf
      
      '切換至來源目錄
      If m_AttachPath <> "." Then ChDir m_AttachPath
      'Modify By Sindy 2018/10/18
      strTmpFileName2 = Replace(m_PA01 & m_PA02 & "Notice of Allowance.pdf", " ", "")
      'If PUB_GetAttachFile_CPP(m_CP09, m_PA01 & m_PA02 & "Notice of Allowance.pdf", m_AttachPath & "\" & strTmpFileName2, True) = True Then
      'Modify By Sindy 2020/2/5
      'If PUB_GetAttachFile_CPP(m_CP09, m_PA01 & Val(m_PA02) & "." & m_CP10 & ".pdf", m_AttachPath & "\" & strTmpFileName2, True) = True Then
      If PUB_GetAttachFile_CPP(m_CP09, m_PA01 & m_PA02 & "." & m_CP10 & ".pdf", m_AttachPath & "\" & strTmpFileName2, True) = True Then
      '2020/2/5 END
      '2018/10/18 END
         DoEvents
         '合併
         strMergeName = "merge" & ServerTime & ".pdf"
         strMergeFN = ".\" & strTmpFileName & ".PDF " & ".\" & strTmpFileName2
         strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & strMergeName
         process_id = SHELL(strCmd, vbHide)
         process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
         If process_handle <> 0 Then
            DoEvents
            For intI = 1 To 10
               If PUB_CheckIsRunning(pub_PdftkName) = True Then
                  Sleep 1000
               Else
                  Exit For
               End If
            Next
            If intI > 10 Then
               TerminateProcess process_handle, 0&
               CloseHandle process_handle
               MsgBox m_PA01 & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", m_PA03 & m_PA04, "") & "合併PDF失敗！"
            Else
               CloseHandle process_handle
            End If
            Set fs = CreateObject("Scripting.FileSystemObject")
            fs.CopyFile m_AttachPath & "\" & strMergeName, strFolder & "\" & strFileName
            DoEvents
            Set fs = Nothing
            strErrText = strErrText & "合併成功" & vbCrLf
         Else
            MsgBox m_PA01 & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", m_PA03 & m_PA04, "") & "合併PDF失敗！"
         End If
      Else
         MsgBox "無法儲存檔案[ " & strTmpFileName2 & " ]！"
      End If
'      If Me.optSel(1).Value = True Then
'         MsgBox "PDF檔已存於 " & strFolder & "！"
'      End If
   End If
   '2017/6/8 END
   
ErrHnd2:
   'Add By Sindy 2019/8/26
   If Err.Number = 70 Then '70:沒有使用權限
      strErrText = ""
      '接著發生錯誤陳述式的下個陳述式開始執行
      Resume Next
   End If
   '2019/8/26 END
End Sub

Public Sub SetInputFocus()
   If optSel(0).Value = True Then
      textCP05_1.SetFocus
   Else
      textPA01.SetFocus
   End If
End Sub
'Add by Morgan 2004/6/23
'取得英日文定稿處理方式
'stMode:2=英,3=日
Public Function GetET03(strPA0104 As String, Optional ByVal stMode As String = "2") As String

   GetET03 = "09" '預設
   
On Error GoTo ErrHnd

   strSql = "Select PA71, PA75, PA11, FA42, FA39, PA26,CU75,CU72 From PATENT, FAGENT, CUSTOMER" & " WHERE " & ChgPatent(strPA0104) & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)"
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   '若有資料
   If adoRecordset.RecordCount > 0 Then
      '若基本檔有設定FCP領證自動代繳欄
      If "" & adoRecordset("PA71").Value = "Y" Then
         GetET03 = "10" '自動代繳
      '申請案號的第九碼非NULL
      'Modify by Morgan 2010/12/27 申請案號改碼數
      'ElseIf "" & Mid("" & adoRecordset("PA11").Value, 9, 1) <> "" Then
      'Modified by Morgan 2012/12/24 +衍生設計也有單獨的證書
      'ElseIf "" & Mid("" & adoRecordset("PA11").Value, 10, 1) <> "" Then
      ElseIf "" & Mid("" & adoRecordset("PA11").Value, 10, 1) <> "" And "" & Mid("" & adoRecordset("PA11").Value, 10, 1) <> "D" Then
         GetET03 = "11" '聯合
      '若基本檔有代理人
'Modify by Morgan 2010/2/2 代理人沒設時要抓申請人
'      ElseIf "" & adoRecordset("PA75").Value <> "" Then
'         If "" & adoRecordset("FA42").Value = "Y" Then
'            GetET03 = "10" '自動代繳欄
'         ElseIf "" & adoRecordset("FA39").Value <> "" Then
'            GetET03 = "12" '收款後辦案有值
'         End If
      ElseIf "" & adoRecordset("FA42").Value = "Y" Then
         GetET03 = "10" '自動代繳欄
      ElseIf "" & adoRecordset("FA39").Value <> "" Then
         GetET03 = "12" '收款後辦案有值
'end 2010/2/2
      '若基本檔有申請人(一定)
      ElseIf "" & adoRecordset("PA26").Value <> "" Then
         If "" & adoRecordset("CU75").Value = "Y" Then
            GetET03 = "10" '自動代繳
         ElseIf "" & adoRecordset("CU72").Value <> "" Then
            GetET03 = "12" '收款後辦案有值
         End If
       End If
       
'Removed by Morgan 2012/12/11 與一般統一
'       'Add by Morgan 2007/4/14 Nikon(Y45148)定稿特別 20~23
'       If stMode = "2" Then
'         If Left("" & adoRecordset("PA75").Value, 6) = "Y45148" Then
'            GetET03 = "20"
'         End If
'      End If
'      'end 2007/4/14
'end 2012/12/11

      'Added by Morgan 2014/1/28
      'Modified by Morgan 2021/3/3 +Y45268B1
      If InStr("Y4526800,Y45268B1", Left("" & adoRecordset("PA75"), 8)) > 0 Then
         GetET03 = "04"
      'Added by Morgan 2016/8/24 --郭怡瑩
      'Modified by Morgan 2016/8/29 Y45002+X27766 --邱子瑜
      'Modified by Morgan 2016/9/26 +Y48804 --郭怡瑩
      'Remove by Lydia 2017/08/01 因為內文只差一句,所以合併為一般定稿09
      'ElseIf (Left("" & adoRecordset("PA75"), 8) = "Y4945600") Or (Left("" & adoRecordset("PA75"), 8) = "Y4880400") Or (Left("" & adoRecordset("PA75"), 8) = "Y4500200" And Left("" & adoRecordset("PA26"), 8) = "X2776600") Then
      '   GetET03 = "20"
      'end 2017/08/01
      End If
      'end 2014/1/28
      
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
   
   'Add by Morgan 2004/12/22 加日文
   If stMode = "3" Then
      If GetET03 = "10" Then
         GetET03 = "16"
      ElseIf GetET03 = "11" Then
         GetET03 = "19"
      Else
         GetET03 = "03"
      End If
   End If
End Function


'Add by Morgan 2004/6/23
'加列印備註參數
Public Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String, ByVal PA0104 As String, Optional ByVal stPS As String, Optional ByVal ET03_1 As String, Optional ByVal stKind As String)
   'Modified by Lydia 2017/07/20 strTxt(1 To 30) => strTxt(1 To 40)
   Dim strTxt(1 To 50) As String
   Dim ii As Integer
   Dim stNP08 As String, stNP09 As String, dblFee As Double, dblFeeA As Double, dblFeeB As Double
   Dim iPage As Integer
   Dim bDisc As Boolean '是否可減免
   Dim stContent As String
   Dim strComb As String '特殊文面組合 A:設計可減免,B:核對已准專利,C:有欠款;
   Dim bolCanDiv As Boolean 'Added by Morgan 2020/5/19 是否有分割期限
   Dim strMemo As String, strClaims As String 'Added by Lydia 2020/12/30
   Dim stNP23 As String 'Add By Sindy 2021/4/26
   
   ii = 0
   EndLetter ET01, ET02, ET03, strUserNum
   If ET03_1 <> "" Then
     EndLetter ET01, ET02, ET03_1, strUserNum
   End If
   
   '不論是否續辦與否
   strSql = "SELECT NP08, NP09, NP23 FROM NEXTPROGRESS WHERE " & ChgNextProgress(PA0104) & " AND NP07 IN ('601','602','603') ORDER BY NP08 Desc "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      stNP08 = "" & adoRecordset.Fields("NP08")
      stNP09 = "" & adoRecordset.Fields("NP09")
      stNP23 = "" & adoRecordset.Fields("NP23") 'Add By Sindy 2021/4/26
   End If
   If Val(stNP08) > 0 Then
      '例外欄位--本所期限
      ii = ii + 1
      'Modify By Sindy 2021/4/23
      If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','約定期限','" & stNP23 & "')"
      Else
      '2021/4/23 END
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','本所期限','" & stNP08 & "')"
      End If
   End If
   'Add by Morgan 2007/4/14
   If Val(stNP09) > 0 Then
      '例外欄位--法定期限
       ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','法定期限','" & stNP09 & "')"
   End If
   'end 2007/4/14
      
   bDisc = PUB_GetFCPCaseDiscState(PA0104) 'Add by Morgan 2008/4/17
   'Modify by Morgan 2011/7/1 要傳專利種類,因為前3年的費用修法後已不相同
   'dblFee = Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "605", 1, 1))
   dblFee = Val(PUB_GetYF0607(台灣國家代號, m_PA08, "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(台灣國家代號, m_PA08, "Y00000000", "605", 1, 1))
   
   If bDisc = True Then dblFee = dblFee - 800 'Add by Morgan 2008/4/17
   'Add by Morgan 2005/11/3
   '規費(領證+第一年年費)
   'Modify by Morgan 2011/7/1 要傳專利種類,因為前3年的費用修法後已不相同
   'dblFeeA = Val(PUB_GetYF07(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF07(台灣國家代號, "1", "Y00000000", "605", 1, 1))
   dblFeeA = Val(PUB_GetYF07(台灣國家代號, m_PA08, "Y00000000", "601", 1, 1)) + Val(PUB_GetYF07(台灣國家代號, m_PA08, "Y00000000", "605", 1, 1))
   If bDisc = True Then dblFeeA = dblFeeA - 800 'Add by Morgan 2008/4/17
   '服務費(領證+第一年年費)
   dblFeeB = dblFee - dblFeeA
   '2005/11/3 end
   
    ii = ii + 1
    '領證費
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','領證費','" & Format(dblFee) & "')"
    ii = ii + 1
    dblFee = dblFee / PUB_GetUSXRate
    '費用
   '美金取至整數位(無條件捨去)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用','" & Format(Fix(dblFee), "0.00") & "')"
   
   'Add by Morgan 2005/11/3
   ii = ii + 1
   '規費
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','規費','" & Format(dblFeeA) & "')"
   ii = ii + 1
   '服務費
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','服務費','" & Format(dblFeeB) & "')"
      
   '規費2
   'Modify by Morgan 2011/7/1 要傳專利種類,因為前3年的費用修法後已不相同
   'dblFeeA = dblFeeA + Val(PUB_GetYF07(台灣國家代號, "1", "Y00000000", "605", 2, 2))
   dblFeeA = dblFeeA + Val(PUB_GetYF07(台灣國家代號, m_PA08, "Y00000000", "605", 2, 2))
   If bDisc = True Then dblFeeA = dblFeeA - 800 'Add by Morgan 2008/4/17
   '服務費2
   'Modify by Morgan 2011/7/1 要傳專利種類,因為前3年的費用修法後已不相同
   'dblFeeB = dblFeeB + 0.5 * Val(PUB_GetYF06(台灣國家代號, "1", "Y00000000", "605", 2, 2))
   dblFeeB = dblFeeB + 0.5 * Val(PUB_GetYF06(台灣國家代號, m_PA08, "Y00000000", "605", 2, 2))
   dblFee = dblFeeA + dblFeeB
   
    ii = ii + 1
    '領證費
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','領證費2','" & Format(dblFee) & "')"
    ii = ii + 1
    dblFee = dblFee / PUB_GetUSXRate
    '費用
   '美金取至整數位(無條件捨去)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用2','" & Format(Fix(dblFee), "0.00") & "')"
   ii = ii + 1
   '規費2
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','規費2','" & Format(dblFeeA) & "')"
   ii = ii + 1
   '服務費2
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','服務費2','" & Format(dblFeeB) & "')"
   
   '規費3
   'Modify by Morgan 2011/7/1 要傳專利種類,因為前3年的費用修法後已不相同
   'dblFeeA = dblFeeA + Val(PUB_GetYF07(台灣國家代號, "1", "Y00000000", "605", 3, 3))
   dblFeeA = dblFeeA + Val(PUB_GetYF07(台灣國家代號, m_PA08, "Y00000000", "605", 3, 3))
   If bDisc = True Then dblFeeA = dblFeeA - 800 'Add by Morgan 2008/4/17
   '服務費3
   'Modify by Morgan 2011/7/1 要傳專利種類,因為前3年的費用修法後已不相同
   'dblFeeB = dblFeeB + 0.5 * Val(PUB_GetYF06(台灣國家代號, "1", "Y00000000", "605", 3, 3))
   dblFeeB = dblFeeB + 0.5 * Val(PUB_GetYF06(台灣國家代號, m_PA08, "Y00000000", "605", 3, 3))
   '領證費3
   dblFee = dblFeeA + dblFeeB
   
   ii = ii + 1
    '領證費3
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','領證費3','" & Format(dblFee) & "')"
    ii = ii + 1
    dblFee = dblFee / PUB_GetUSXRate
    '費用3
   '美金取至整數位(無條件捨去)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用3','" & Format(Fix(dblFee), "0.00") & "')"
   ii = ii + 1
   '規費3
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','規費3','" & Format(dblFeeA) & "')"
   ii = ii + 1
   '服務費3
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','服務費3','" & Format(dblFeeB) & "')"
      
   '2005/11/3 end
   
   'Add by Morgan 2009/8/19
   If InStr(m_WithReportList, PA0104) = 0 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','不印檢索','♀')"
   End If
   
   'Added by Morgan 2012/11/28
   If InStr(m_WithReportList, PA0104) > 0 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有檢索要印','♀')"
   End If
   
   'Added by Morgan 2017/3/9
   '特定Y+X不寄紙本
   If m_PA75 = "Y34232000" And (m_PA26 = "X64636000" Or m_PA26 = "X30299000" Or m_PA26 = "X54783000" Or m_PA26 = "X45816000" Or m_PA26 = "X45207000" Or m_PA26 = "X53359000" Or m_PA26 = "X30151000" Or m_PA26 = "X48269000" Or m_PA26 = "X47023000" Or m_PA26 = "X73190000") Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','寄紙本才印','')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','不寄紙本才印','♀')"
   End If
   'end 2017/3/9
   
   'Move by Lydia 2020/12/30 為了先取得特殊設定，從下面移到此處
   'Modified by Lydia 2020/12/30 增加”日文定稿請求項”，並且另用變數
   'If GetApprovalPS(m_PA01 & m_PA02 & m_PA03 & m_PA04, m_PA75, m_PA26 & "," & m_PA27 & "," & m_PA28 & "," & m_PA29 & "," & m_PA30, strExc(1)) = True Then 'Modified by Lyda 2019/3/11
   If m_PA08 <> "3" Then 'Added by Lydia 2021/09/03 排除設計案
      'Modified by Lydia 2022/10/05 傳入定稿語文
      'If GetApprovalPS(m_PA01 & m_PA02 & m_PA03 & m_PA04, m_PA75, m_PA26 & "," & m_PA27 & "," & m_PA28 & "," & m_PA29 & "," & m_PA30, strMemo, strClaims) = True Then
      'Modified by Lydia 2023/03/22 +pKind=1
      If PUB_GetApprovalPS("1", m_PA01 & m_PA02 & m_PA03 & m_PA04, m_PA75, m_PA26 & "," & m_PA27 & "," & m_PA28 & "," & m_PA29 & "," & m_PA30, strMemo, strClaims, m_LetterLanguage) = True Then
         If strMemo <> "" Then
      'end 2022/10/05
             ii = ii + 1
             strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                 "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','額外段落','" & strMemo & "')"
         End If 'Add ed by Lydia 2022/10/05
      End If
   End If 'Added by Lydia 2021/09/03
   
   'Added by Morgan 2017/9/7
   'Added by Sindy 2020/6/29 + Y34232(Yasutomi)+X30299(積水化學) 優先告准，需附中文Claims請修改定稿內容 (FCP-04-000-03)
   'Modified by Lydia 2020/12/30 改用模組判斷”日文定稿請求項”
   'If (m_PA75 = "Y54732000" And m_PA26 = "X30299000") Or _
   '   (m_PA75 = "Y34232000" And m_PA26 = "X30299000") Then
   '   ii = ii + 1
   '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
   '      " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有附中文claims才印','♀')"
   'End If
   ''end 2017/9/7
   If strClaims = "1" Then  '1.  需附中文請求項
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','需附中文claims才印','♀')"
   ElseIf strClaims = "2" Then '2.  需附原文(含英文及日文)請求項
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','需附原文claims才印','♀')"
   'Added by Lydia 2024/03/18
   ElseIf strClaims = "3" Then '3.  需附中文及原文請求項
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','需附中文及原文claims才印','♀')"
   'end 2024/03/18
   End If
   'end 2020/12/30
   
   'Added by Lydia 2017/07/20 Y34210+ X47717的 The notice of allowance and its translation<寄紙本> 或<不寄紙本才印> 取代為"Attached please find the notice of allowance with translation and allowed claims"
   'Modified by Lydia 2017/08/01 因為與特殊定稿相同,所以合併在一起
   'If m_PA75 = "Y34210000" And m_PA26 = "X47717000" Then
   'Modified by Morgan 2019/1/9 +Y27766
   'Modified by Morgan 2020/3/6 +Y48804010
   'Modified by Lydia 2022/02/08 先將Y34210000+X47717000寫在文中的設定拿掉(因已另行設定) -- (m_PA75 = "Y34210000" And m_PA26 = "X47717000") Or
   If (Left(m_PA75, 8) = "Y4945600") Or (Left(m_PA75, 8) = "Y4880400") Or (Left(m_PA75, 8) = "Y4500200" And Left(m_PA26, 8) = "X2776600") Or (Left(m_PA75, 8) = "Y2776600") Or (Left(m_PA75, 8) = "Y4880401") Then
   'end 2017/07/31
      '直接將判斷欄位設為空白
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','寄紙本才印','')"
      '直接將判斷欄位設為空白
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','不寄紙本才印','')"
         
      'Added by Morgan 2022/8/23 Y27766,Y4500200+X2776600不印例外備註(有其他附件改人工維護) --Bobbie
      If (Left(m_PA75, 8) = "Y4500200" And Left(m_PA26, 8) = "X2776600") Or (Left(m_PA75, 8) = "Y2776600") Then
      
      Else
      'end 2022/8/23
         '取代句子
         ii = ii + 1
         'Modified by Lydia 2017/08/01 內文改寫在定稿內
         'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','例外備註','Attached please find the notice of allowance with translation and allowed claims')"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','例外備註','♀')"
      End If 'Added by Morgan
   End If
   'end 2017/07/20
   
   'Added by Morgan 2012/11/8
   '發明初審核准加分割期限
   m_bolDivSug = False
   m_strDivState = "N"
   If m_PA08 = "1" Then
         strExc(1) = ""
         strExc(2) = ""
         strExc(3) = "" 'Add By Sindy 2021/4/26
         
         'Modified by Morgan 2012/12/26 +考慮分割案核准
         strExc(0) = "SELECT a.cp05,pa162,b.cp10,b.cp09,pa163 FROM caseprogress a,caseprogress b,patent" & _
            " WHERE a.cp09='" & m_CP09 & "' and a.cp10='1001' and a.cp05>=20121202 and b.cp09(+)=a.cp43 and b.cp10 in ('101','307')" & _
            " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Added by Morgan 2019/7/30 108.11.1 新法分割期限同領證期限
            If m_CP05 >= "20191001" Then
               strExc(1) = stNP09
               strExc(2) = stNP08
               strExc(3) = stNP23 'Add By Sindy 2021/4/26
            Else
            'end 2019/7/30
            
               strExc(1) = CompDate(2, 30, RsTemp.Fields(0))
               strExc(2) = CompDate(2, -2, strExc(1))
               
            End If 'Added by Morgan 2019/7/30
            
            '發明申請
            If RsTemp.Fields("cp10") = "101" Then
               m_strDivState = "Y"
            '分割
            ElseIf RsTemp.Fields("cp10") = "307" And RsTemp.Fields("pa163") = "Y" Then
               m_strDivState = "Y"
            End If
         End If 'Added by Morgan 2016/10/20
            
            If m_strDivState = "Y" Then
               If Val(strExc(1)) > 0 Then bolCanDiv = True 'Added by Morgan 2020/5/19
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                  " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明初審核准要印','♀')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                  " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割法定期限','" & strExc(1) & "')"
               ii = ii + 1
               'Modify By Sindy 2021/4/23
               If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                     " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割約定期限','" & strExc(3) & "')"
               Else
               '2021/4/23 END
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                     " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割本所期限','" & strExc(2) & "')"
               End If
               strComb = strComb & "D"
               
               'Removed by Morgan 2019/10/2 108.11.1 新法發明、新型都可分割，分割建議改到下面
               'If RsTemp("pa162") = "Y" Then
               '   strComb = strComb & "E"
               '   strExc(0) = "select dst05 from divsugtext where dst01='" & m_PA01 & "'" & _
               '      " and dst02='" & m_PA02 & "' and dst03='" & m_PA03 & "'" & _
               '      " and dst04='" & m_PA04 & "' and dst05 is not null"
               '   intI = 1
               '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               '   If intI = 1 Then
               '      m_bolDivSug = True
               '      ii = ii + 1
               '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               '         " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割建議','" & ChgSQL(Trim(RsTemp(0))) & " ')"
               '   End If
               'End If
               'end 2019/10/2
               
            'Added by Morgan 2016/1/19
            Else
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                  " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明非初審核准要印','♀')"
            'end 2016/1/19
            
            End If
            
         'End If 'Removed by Morgan 2016/10/20
   End If
   'end 2012/11/8
   
   'Add by Morgan 2011/7/27
   '設計可減免
   If bDisc = True And m_PA08 = "3" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','新式樣可減免才印','♀')"
         
      strComb = strComb & "A" 'Add by Morgan 2011/8/17
   End If
   
   'Add by Morgan 2011/8/17
   strExc(0) = "SELECT 1 FROM caseprogress WHERE " & ChgCaseprogress(PA0104) & " AND cp10='926' and cp10='926' and cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strComb = strComb & "B"
   End If
   
   'Added by Morgan 2013/7/19
   '一案兩請提醒
   If m_PA08 = "2" Then
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
         " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and " & ChgCaseMap(PA0104, 0, 0) & _
         " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and " & ChgCaseMap(PA0104, , 1) & ") X" & _
         ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','一案兩請新型案要印','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案申請號','" & adoRecordset("pa11") & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案彼所案號','" & IIf(IsNull(adoRecordset("pa77")), "", "" & adoRecordset("pa77")) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案本所案號','" & adoRecordset("CNo") & "')"
      End If
   End If
   'end 2013/7/19
                           
   'Add by Morgan 2011/6/23
   If m_bPrintBill = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有欠款才印','♀')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有欠款不印','♀')"
      
      strComb = strComb & "C" 'Add by Morgan 2011/8/17
   End If
   
   strExc(1) = ""
   'Added by Morgan 2012/11/28
   If InStr(strComb, "D") > 0 Then
      If InStr(strComb, "E") = 0 Then '有分割建議時不控制跳頁
         strExc(1) = "有D段"
      End If
   Else
   'end 2012/11/28
   
      'Add by Morgan 2011/8/17
      If Len(strComb) = 1 Then
         strExc(1) = "只有" & strComb & "段"
         
      'A+(B or C or BC)
      ElseIf InStr(strComb, "A") > 0 Then
         strExc(1) = "有AX段"
         
      'B+C
      ElseIf strComb = "BC" Then
         strExc(1) = "有BC段"
      End If
   End If
   
   If strExc(1) <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','" & strExc(1) & "','♀')"
   End If
   'end 2011/8/17
   
   strExc(1) = ""
   'Added by Mogan 2013/9/24
   'Added by Lydia 2014/10/28 針對代理人 Y4835301 且申請人為NIKE(X55265,X72195) 的案件,於專利核准函之已准收文單的備註列印提示
   'Modified by Morgan 2016/5/13 +Y53793 --陳亭妙
   'Modified by Morgan 2016/7/29 +Y20990010--鄭詠心, +Y52527010--葉敏莉
   'Modified by Morgan 2016/11/21 +Y45616000--陳亭妙
   'Modified by Morgan 2017/2/21 +Y45659000,Y45659010--Bobbie
   'Modified by Morgan 2017/2/24 +Y52859000,Y51799010--Bobbie
   'Modified by Morgan 2017/3/8 +Y4830904 --鄭詠心
   'Modified by Morgan 2017/8/1 +Y53565000,X46435010--葉敏莉 (代理人統一抓前8碼並將相同內容合併)
   'Modified by Morgan 2017/8/8 +Y52259且X45149 --吳若芬
   'Modified by Morgan 2017/12/11 +Y51799040--Bobbie
   'Modified by Morgan 2018/5/23 + Y20372000 --Lina
   'Modified by Morgan 2018/6/8 + Y45814 BASF SE Global Intellectual Property --葉子寧
   'Modified by Morgan 2018/7/19 +FCP59268,FCP59269 --Jessica
   'Modified by Morgan 2018/11/30 +(Y52283+X49019)--Landy
   'Modified by Lydia 2019/03/11 額外段落改成開Table維護(ApprovalPS)
'   If Left(m_PA75, 8) = "Y4827900" _
'      Or (Left(m_PA75, 8) = "Y4835301" And (Left(m_PA26, 8) = "X5526500" Or Left(m_PA26, 8) = "X7219500")) _
'      Or Left(m_PA75, 8) = "Y5379300" _
'      Or Left(m_PA75, 8) = "Y2099001" Or Left(m_PA75, 8) = "Y5252701" _
'      Or Left(m_PA75, 8) = "Y4561600" _
'      Or Left(m_PA75, 8) = "Y4565900" Or Left(m_PA75, 8) = "Y4565901" _
'      Or Left(m_PA75, 8) = "Y5285900" Or Left(m_PA75, 8) = "Y5179901" _
'      Or Left(m_PA75, 8) = "Y4830904" _
'      Or Left(m_PA75, 8) = "Y5356500" Or InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X4643501") > 0 _
'      Or (Left(m_PA75, 8) = "Y5225900" And InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X4514900") > 0) _
'      Or Left(m_PA75, 8) = "Y5179904" _
'      Or Left(m_PA75, 8) = "Y2037200" _
'      Or Left(m_PA75, 8) = "Y4581400" _
'      Or PA0104 = "FCP059268000" Or PA0104 = "FCP059269000" _
'      Or (Left(m_PA75, 8) = "Y5228300" And Left(m_PA26, 8) = "X4901900") Then
'
'      strExc(1) = "Enclosed please find the allowed claims for this application for your reference and records."
'
'   'Added by Morgan 2018/2/23 +Y27856B70--Phoebe
'   ElseIf Left(m_PA75, 8) = "Y27856B7" Then
'      strExc(1) = "For your prompt reference and file, enclosed please find the English allowed claims for the referenced application."
'
'   'Added by Morgan 2017/12/27 --敏莉
'   ElseIf Left(m_PA75, 8) = "Y5458900" Then
'      strExc(1) = "Attached please find the granted claims in word format for your prompt attention and records."
'
'   'Added by Morgan 2014/7/24 --Sharon
'   ElseIf InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X5863100") > 0 Then
'      strExc(1) = "Enclosed please find the English translation of the allowed claims for this application for your reference and records."
'
'   'Modified by Morgan 2016/8/22 +Y52242 --Elisa(確認不要與X5863100用同一句)
'   ElseIf Left(m_PA75, 8) = "Y5224200" Then
'      strExc(1) = "Enclosed please find an English version of the allowed claims for your records and prompt reference."
'
'   'Modified by Morgan 2016/11/29 不會再進單筆畫面,從GetPS移過來(避免檢查時遺漏)
'   'Modified by Morgan 2014/2/10 設定重複,留範圍大的 --毓芳;判斷所有申請人
'   'ElseIf m_PA75 = "Y34232000" And m_PA26 = "X30299000" Then
'   ElseIf InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X3029900") > 0 Then
'      strExc(1) = "Per the general instructions of Sekisui Chemical, attached please find the allowed Taiwanese claims for your review."
'
'   'Added by Morgan 2014/2/10--江如玉
'   ElseIf Left(m_PA75, 8) = "Y1892300" And (InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X2798400") > 0 Or InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X4771900") > 0) Then
'      strExc(1) = "The English translation of the allowed claims will follow shortly."
'
'   'Added by Morgan 2016/10/20 --陳怡蓉
'   ElseIf Left(m_PA75, 8) = "Y45801B2" And (InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X47805000") > 0 Or InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X70269010") > 0) Then
'      strExc(1) = "Enclosed please find the allowed claims in English for your reference and records."
'
'   'Added by Morgan 2016/11/21 --陳怡蓉
'   ElseIf Left(m_PA75, 8) = "Y5404700" Then
'      strExc(1) = "Enclosed please find a copy of the most recently approved claims for your records and prompt reference."
'
'   'Added by Morgan 2017/2/20 --Joyce
'   ElseIf Left(m_PA75, 8) = "Y2049500" Or Left(m_PA75, 8) = "Y2049503" Or Left(m_PA75, 8) = "Y2049504" Or Left(m_PA75, 8) = "Y2049501" Or Left(m_PA75, 8) = "Y4625500" Or Left(m_PA75, 8) = "Y5327500" Or Left(m_PA75, 8) = "Y5327502" Then
'      strExc(1) = "Pursuant to your instructions, attached please find a copy of the granted claims in English for your reference."
'
'   'add by sonia 2017/7/13  --洪培堯
'   'Modified by Morgan 2017/11/8 +Y33801010  --洪培堯
'   ElseIf Left(m_PA75, 8) = "Y2782000" Or Left(m_PA75, 8) = "Y3380101" Then
'      strExc(1) = "Enclosed please find a copy of the allowed claims for your records."
'
'   'Added by Morgan 2017/8/16 --潘子微
'   ElseIf Left(m_PA75, 8) = "Y2068300" Then
'      strExc(1) = "Enclosed please find a scanned version of patent specification with the allowed claims for your prompt reference and records."
'
'   'end 2016/11/29
'   End If
'
'   If strExc(1) <> "" Then
   'Move by Lydia 2020/12/30 為了先取得特殊設定，從下面移到上方
   'end 2013/9/24
   
   'Modified by Morgan 2014/2/10 不續辦准通知應該都要判斷
   'If stPS = "" Then
   'Modified by Lydia 2021/03/09 改用Table的設定
   'If InStr(stPS, "This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately.") = 0 Then
   If InStr(strMemo, "This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately.") = 0 Then
      'Modify by Morgan 2009/12/8 +X30299備註
      '判斷是否不續辦但准通知
      strSql = "Select PA89,PA26 From Patent Where " & ChgPatent(PA0104) & " and pa57 is not null"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Removed by Morgan 2014/2/10 移到 GetPS
         'If "" & adoRecordset.Fields(1).Value = "X30299000" Then
         '   strExc(1) = "Per the general instructions of Sekisui Chemical, enclosed please find the allowed Taiwanese claims for your review."
         'End If
         'end 2014/2/10
         
         If "" & adoRecordset.Fields(0).Value = "Y" Then
            If stPS <> "" Then
               stPS = stPS & vbCrLf & "     "
            End If
            stPS = stPS & "This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately."
         End If
      End If
   End If
   
   If stPS <> "" Then
      ii = ii + 1
      '列印備註
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','列印備註','P.S. : " & stPS & "')"
   End If
      
   'Added by Morgan 2019/7/30
   '英文
   If ET03 = "02" Then
      '自動代繳(10)
      If stKind = "10" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','自動代繳不印','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','自動代繳才印','♀')"
      End If
      
      '收款後辦案(12)
      If stKind = "12" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','收款後辦案才印','♀')"
      End If
      
      'Y4526800(04)
      If stKind = "04" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','Y45268不印','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','Y45268才印','♀')"

      End If
   End If
   
   '英文、日文
   If ET03 = "02" Or ET03 = "03" Or ET03 = "16" Then
      '2019/10/1以後
      If m_CP05 >= "20191001" Then
         '發明新型
         If m_PA08 = "1" Or m_PA08 = "2" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/10/1起發明新型准才印','♀')"
         End If
      '2019/8/1-9/30
      ElseIf m_CP05 >= "20190801" Then
         '發明初審
         If m_strDivState = "Y" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/8/1-9/30發明初審准才印','♀')"
         End If
         '發明新型
         If m_PA08 = "1" Or m_PA08 = "2" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/8/1-9/30發明新型准才印','♀')"
         End If
      End If
      
      If m_CP05 >= "20190801" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/8/1起不印','♀')"
            
         '發明新型
         If m_PA08 = "1" Or m_PA08 = "2" Then
            '分割期限=領證期限(發明初審核准期限在上面設定)
            If m_strDivState <> "Y" Then
               If Val(stNP09) > 0 Then bolCanDiv = True 'Added by Morgan 2020/5/19
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                  " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割法定期限','" & stNP09 & "')"
               ii = ii + 1
               'Modify By Sindy 2021/4/23
               If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                     " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割約定期限','" & stNP23 & "')"
               Else
               '2021/4/23 END
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                     " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割本所期限','" & stNP08 & "')"
               End If
            End If
            
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','核准函收文日','" & m_CP05 & "')"
         End If
      End If
   End If
   
   '中文
   If ET03 = "01" Then
      '2019/10/1以後
      If m_CP05 >= "20191001" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/10/1起不印','♀')"
      
      'Added by Morgan 2019/9/11 即日起都帶新法規定--David
      End If
      If m_CP05 >= "20190801" Then
      'end 2019/9/11
      
         '發明新型
         If m_PA08 = "1" Or m_PA08 = "2" Then
            ii = ii + 1
            'Modified by Morgan 2019/9/11 即日起都帶新法規定--David
            'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/10/1起發明新型准才印','♀')"
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/8/1起發明新型准才印','♀')"
            'end 2019/9/11
            
            '分割期限=領證期限(發明初審核准期限在上面設定)
            If m_strDivState <> "Y" Then
               If Val(stNP09) > 0 Then bolCanDiv = True 'Added by Morgan 2020/5/19
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                  " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割法定期限','" & stNP09 & "')"
               ii = ii + 1
               'Modify By Sindy 2021/4/23
               If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                     " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割約定期限','" & stNP23 & "')"
               Else
               '2021/4/23 END
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                     " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割本所期限','" & stNP08 & "')"
               End If
            End If
         End If
      
      'Removed by Morgan 2019/9/11 即日起都帶新法規定--David
      'End If
      'If m_CP05 >= "20190801" Then
      'end 2019/9/11
      
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','108/8/1起不印','♀')"
      End If
   End If
   'end 2019/7/30
   
   'Added by Morgan 2019/10/2
   If m_PA08 = "1" Or m_PA08 = "2" Then
      strExc(0) = "select dst05 from patent,divsugtext where pa01='" & m_PA01 & "'" & _
         " and pa02='" & m_PA02 & "' and pa03='" & m_PA03 & "' and pa04='" & m_PA04 & "'" & _
         " and pa162='Y' and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04 and dst05 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strComb = strComb & "E"
         m_bolDivSug = True
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','分割建議','" & ChgSQL(Trim(RsTemp(0))) & " ')"
      End If
   End If
   'end 2019/10/2
   
   'Added by Morgan 2020/5/19
   If bolCanDiv = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有分割期限才印','♀')"
   End If
   'end 2020/5/19
   
   'Added by Morgan 2021/1/28
   'FACEBOOK Group之告准定稿增加核對已准之固定報價 --Anny
   If InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X8066800") > 0 _
      Or InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X8066900") > 0 _
      Or InStr(m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, "X8067000") > 0 Then
      ii = ii + 1
      If m_PA08 = "1" Or m_PA08 = "2" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','核對已准固定報價',', in the amount of USD150,')"
      ElseIf m_PA08 = "3" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','核對已准固定報價',', in the amount of USD100,')"
      End If
   'Added by Morgan 2023/10/6
   'Y55948000 SATELLOGIC 核對已准之固定報價 --Franny
   ElseIf m_PA75 = "Y55948000" Then
      ii = ii + 1
      If m_PA08 = "1" Or m_PA08 = "2" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','核對已准固定報價',', in the amount of NTD3,825 (=NTD4,500 x 85%),')"
      ElseIf m_PA08 = "3" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','核對已准固定報價',', in the amount of NTD2,550 (=NTD3,000 x 85%),')"
      End If
   End If
   'end 2021/1/28
   
   CheckOC
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(ii, strTxt) Then
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
On Error GoTo ErrHnd

   '傳真頁數(通知函例外欄位寫入後才可讀出正確內容)
   'Add by Morgan 2006/2/13
   If ET03_1 <> "" Then
      'Add by Morgan 2011/6/27
      If (m_PA75 = "Y49456000" And m_PA26 = "X47325000") Or (m_PA75 = "Y34232000" And m_PA26 = "X30299000") Then
         '因為要加日文請求項頁數地方空白--何淑華 2011/6/24
         
      Else
      'end 2011/6/27
         '傳真頁數
         'Modified by Morgan 2014/8/5 改與核准函輸入一樣抓前8碼(已檢查資料OK)
         Select Case Left(m_PA75, 8)
            Case "Y4908300" 'DRAGON
               iPage = 3
            Case "Y4514800" 'Nikon
               iPage = 5
            Case "Y3423200" 'YASUTOMI
               iPage = 0 'Modified by Morgan 2012/12/26 不固定不印人工填--- 何淑華
            'Add by Morgan 2011/6/27 加傳真來函公文-- 何淑華 2011/6/24
            'Modified by Morgan 2013/12/9 +Y20049 -- 何淑華
            'Modified by Morgan 2014/8/5 +Y51306,Y28043 -- 何淑華
            'Modified by Morgan 2014/8/6 +Y52061 -- 何淑華(吳彩菱)
            'Modified by Morgan 2015/6/11 +Y34271 -- 何淑華
            'Modified by Morgan 2015/9/8 +Y51622 -- 何淑華
            'Modified by Morgan 2015/12/22 +Y20065 -- 何淑華
            Case "Y4745300", "Y2204600", "Y5292200", "Y2004900", "Y5130600", "Y2804300", "Y5206100", "Y3427100", "Y5162200", "Y2006500"
               iPage = 5
            'Modified by Sindy 2016/8/30 +Y52117 -- 鄭詠心
            'Modified by Sindy 2017/6/7 +Y20050 -- 鄭詠心
            Case "Y5211700", "Y2005000"
               iPage = 6
            Case Else
               iPage = 2
         End Select
         
         'Added by Morgan 2014/7/2 日文定稿內的跳頁符號已拿掉,改固定加1頁
         If ET03 = "03" Or ET03 = "06" Then
            iPage = iPage + 1
         End If
         'end 2014/7/2
         
         If iPage > 0 Then
            'Add by Morgan 2007/4/20
            'Modify by Morgan 2011/8/17 抓解析後的定稿內容判斷才準確
            'If PUB_IsMultiPage(ET01, ET02, ET03) = True Then
            NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent
            If InStr(stContent, Chr(12)) > 0 Then
            'end 2011/8/17
               '因為信函要多印一份給客戶所以要多加一頁
               If Left(m_PA75, 6) = "Y49083" Then
                  iPage = iPage + 2
               Else
                  iPage = iPage + 1
               End If
            End If
            'end 2007/4/20
            
            'Add by Morgan 2011/6/27 若有列印請款單時頁數也要加入
            If m_bPrintBill = True Then
               iPage = iPage + m_iBillPageCount
            End If
            'end 2011/6/27
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & ET02 & "','" & ET03_1 & "','" & strUserNum & "','傳真頁數','" & iPage & "')"
            cnnConnection.Execute strSql, intI
         End If
      End If
      
   End If 'Add by Morgan 2011/6/27
   
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2015/6/18
Private Sub textPA04_LostFocus()
   m_PA03 = textPA03
   If IsEmptyText(m_PA03) Then m_PA03 = "0": textPA03 = "0"
   m_PA04 = textPA04
   If IsEmptyText(m_PA04) Then m_PA04 = "00": textPA04 = "00"
   Call QueryOtherData
End Sub

Private Sub txtLetterDate_GotFocus()
   CloseIme
   TextInverse txtLetterDate
End Sub

Private Sub txtLetterDate_Validate(Cancel As Boolean)
   If ChkDate(txtLetterDate) = False Then
      Cancel = True
   End If
End Sub

'Removed by Morgan 2022/10/18 P案也要用,改到 PUB_GetApprovalPS
''Added by Lydia 2019/03/11 通知告准加註(ApprvoalPS)
''dbCaseNo:本所案號,dbFA:代理人編號,dbCu:申請人編號
''Modified by Lydia 2020/12/30 pJnClaims 日文定稿請求項
''Memo by Lydia 2021/02/02 增加”通知工程師Email設定”，模組寫在frm06010602_3，若有變更程式兩邊都要檢查一下
''Modified by Lydia 2022/10/05 增加"定稿語文"pStLang
'Private Function GetApprovalPS(dbCaseNo As String, dbFA As String, dbCu As String, Optional ByRef pMemo As String = "", Optional ByRef pJnClaims As String = "", Optional ByVal pStLang As String = "2") As Boolean
'Dim stSQL As String, iR As Integer
'Dim stCon As String
'Dim rsQuery As ADODB.Recordset
''逐筆判斷Y代理人+X申請人1~5;若有一筆以上,只使用第一筆符合
'Dim m_Memo As String
'Dim iCall As Integer, iRound As Integer
'Dim tmpArr As Variant
'Dim m_Claims As String 'Added by Lydia 2020/12/30
'
'   '判斷有幾個申請人
'   tmpArr = Split(dbCu, ",")
'   For iR = 0 To UBound(tmpArr)
'       If Trim(tmpArr(iR)) <> "" Then
'           iCall = iCall + 1
'       End If
'   Next iR
'
'   For iRound = 1 To iCall
'        '順序 1.本所案號 2.代理人+申請人 3.代理人 4.申請人
'        'Modified by Lydia 2020/12/30 + APS12 日文定稿請求項
'        'Modified by Lydia 2022/10/05 +APS15 日文定稿加註
'        stSQL = "select 0 Od1, APS02, APS12, APS15 from ApprovalPS where APS03='" & dbCaseNo & "' " & stCon & _
'           " union select 1 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
'           " union select 2 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
'           " union select 3 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05 is null" & stCon & _
'           " union select 4 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
'           " union select 5 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
'           " union select 6 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05 is null" & stCon & _
'           " union select 7 Od1, APS02, APS12, APS15 from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
'           " union select 8 Od1, APS02, APS12, APS15 from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
'           " order by Od1, APS02"
'            iR = 1
'            Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
'            If iR = 1 Then
'               'Modified by Lydia 2021/03/09 重新整理,逐筆判斷只使用第一筆符合; FCP-58141曾經另外設工程師Email通知並且優先權更大
'               rsQuery.MoveFirst
'               Do While Not rsQuery.EOF
'                    'Modified by Lydia 2022/10/05
'                    'If "" & rsQuery.Fields("APS02") & rsQuery.Fields("APS12") <> "" Then
'                    '     m_Memo = "" & rsQuery.Fields("APS02")
'                    If (pStLang <> "3" And "" & rsQuery.Fields("APS02") <> "") Or _
'                        (pStLang = "3" And "" & rsQuery.Fields("APS15") & rsQuery.Fields("APS12") <> "") Then
'                         If pStLang <> "3" Then
'                            m_Memo = "" & rsQuery.Fields("APS02")
'                         Else
'                            m_Memo = "" & rsQuery.Fields("APS15")
'                         End If
'                    'end 2022/10/05
'                         m_Claims = "" & rsQuery.Fields("APS12")
'                         GoTo JumpToEnd
'                    End If
'                    rsQuery.MoveNext
'               Loop
'               'end 2021/03/09
'            End If
'   Next iRound
'
'JumpToEnd:
'   pMemo = m_Memo
'   pJnClaims = m_Claims 'Added by Lydia 2020/12/30
'   'Added by Lydia 2021/02/02 改判斷是否有備註
'   If pMemo <> "" Or pJnClaims <> "" Then
'       GetApprovalPS = True
'   End If
'   'end 2021/02/02
'   Set rsQuery = Nothing
'End Function


