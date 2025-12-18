VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010401_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標分割子案視窗"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   9180
   Begin VB.TextBox txtDate 
      Height          =   270
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   0
      Top             =   4890
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7065
      TabIndex        =   3
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6240
      TabIndex        =   2
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8280
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4092
      Left            =   48
      TabIndex        =   6
      Top             =   744
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   7218
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "定稿公報日期(民國)"
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   4920
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "母案有申請意見書的期限要管制，請點選一筆子案繼承管制期限！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   7395
   End
End
Attribute VB_Name = "frm02010401_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/28 Form2.0已修改 grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
'Create by nickc 2006/07/21
Option Explicit

Public IsHaveTM15 As Boolean
Public oKey As String    '母案收文號
Public UpForm As Form
Public oStrCDate As String  '來函收文日及結果日
Dim m_TM01 As String   '母案
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_CP01 As String    '要設定期限的子案
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Public IsHaveNp202 As Boolean
Public IsHaveCp202 As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Public SeekGrdIndex As Integer
Dim i  As Integer
'由下層畫面傳入資料用           內商申請案號    內商註冊號輸入  FCT申請案號輸入  FCT註冊號輸入
Public PutSeekData01 As String '  textCP47              textTM15                textTM09                  textTM14
Public PutSeekData02 As String '  textCP30              textTM14                textTM32                  textTM15
Public PutSeekData03 As String '  textTM27             text1                        textTM11                 textTM21
Public PutSeekData04 As String '  textCP05              textTM21                textTM12                 textTM22
Public PutSeekData05 As String '  textPrint               textTM22                textTM27                 text1
Public PutSeekData06 As String '  textCF09             textCP47                 textCP05                  textCreFee
Public PutSeekData07 As String '  textCP45             textDate                   textPrint                   textPrint
Public PutSeekData08 As String '  textPs                  textMoney               textPriorityDoc        textPrtTrans
Public PutSeekData09 As String '  textNP08             textTC1                   textAddDate            text2
Public PutSeekData10 As String '  textNP09             textTC2                   textAdd                   textNP08
Public PutSeekData11 As String '                             textPrint                  textDN                     textNP09
Public PutSeekData12 As String '                             textPS                    textToEng
Public PutSeekData13 As String '                             textNP08                textPrtTrans
Public PutSeekData14 As String '                             textNP09                textPS
Public PutSeekData15 As String '                                                           textTM67
Public PutSeekData16 As String '                                                           textNP08
Public PutSeekData17 As String '                                                           textNP09
Public PutSeekData18 As String '                                                           combo2(0)
Public PutSeekData19 As String '                                                           combo2(1)
Public PutSeekData20 As String '                                                           combo2(2)
Public PutSeekData21 As String '                                                           combo2(3)
Public PutSeekData22 As String '                                                           combo2(4)
Public PutSeekData23 As String '                                                           combo2(5)
Public PutSeekData24 As String '                                                           combo2(6)
Public PutSeekData25 As String '                                                           combo2(7)
Public PutSeekData26 As String '                                                           combo2(8)
Public PutSeekData27 As String '                                                           combo2(9)
Public PutSeekData28 As String '                                                           textTM47
Public PutSeekData29 As String '                                                           textTM48
Public PutSeekData30 As String '                                                           textTM49
Public PutSeekData31 As String '                                                           textTM50
Public PutSeekData32 As String '                                                           textTM51
Public PutSeekData33 As String '                                                           textTM52
Public PutSeekData34 As String '                                                           textTM94
Public PutSeekData35 As String '                                                           textTM95
Public PutSeekData36 As String '                                                           textTM96
Public PutSeekData37 As String '                                                           textTM97
Public PutSeekData38 As String '                                                           textTM98
Public PutSeekData39 As String '                                                           textTM99
Public PutSeekData40 As String '                                                           textTM100
Public PutSeekData41 As String '                                                           textTM101
Public PutSeekData42 As String '                                                           textTM102
Public PutSeekData43 As String '                                                           textTM103
Public PutSeekData44 As String '                                                           textTM104
Public PutSeekData45 As String '                                                           textTM105
Public PutSeekData46 As String '                                                           textTM106
Public PutSeekData47 As String '                                                           textTM107
Public PutSeekData48 As String '                                                           textTM108
Public PutSeekData49 As String '                                                           textTM109
Public PutSeekData50 As String '                                                           textTM110
Public PutSeekData51 As String '                                                           textTM111
Public PutSeekData52 As String '                                                           textTM112
Public PutSeekData53 As String '                                                           textTM113
Public PutSeekData54 As String '                                                           textTM114
Public PutSeekData55 As String '                                                           textTM115
Public PutSeekData56 As String '                                                           textTM116
Public PutSeekData57 As String '                                                           textTM117
Public PutSeekData58 As String '                                                           textLaw
Dim strCP05 As String
Dim strCP09 As String
Dim strCP27 As String
Dim m_CP12 As String
Dim m_CP13 As String
Dim m_TM67 As String
Dim m_TM09 As String
Dim m_TM15 As String
Dim m_NP08 As String
Dim m_NP09 As String   '2014/12/9 ADD BY SONIA
Dim m_TM12 As String
'add by nickc 2007/08/10
Dim m_Law As String
'add by nickc 2008/01/23  加入可以取消
Public m_IsCancal As Boolean
'Added by Morgan 2017/5/3 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
Dim m_DocPdf As String
Dim m_DocPdfDate As String
Dim m_DocPdfTime As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
'end 2017/5/3
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub
Private Sub cmdCancel_Click()
   Unload Me
   UpForm.Show
End Sub

Private Sub cmdExit_Click()
   If UCase(UpForm.Name) = "FRM02010401_3" Then   '從內商來
       Unload frm02010401_2
       Unload frm02010401_1
   ElseIf UCase(UpForm.Name) = "FRM03020401_03" Then  '從外商來
       Unload frm03020401_02
       Unload frm03020401_01
   End If
   Unload UpForm
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nRow As Integer
   Dim nCol As Integer
   Dim IsSel As Boolean
   Dim Cancel As Boolean
   
   If grdList.Rows > 0 Then
      If (IsEmptyText(m_CP01) = True Or IsEmptyText(m_CP02) = True Or IsEmptyText(m_CP03) = True Or IsEmptyText(m_CP04) = True) And (IsHaveNp202 = True Or IsHaveCp202 = True) Then
         strMsg = "請先選取一筆記錄承接期限"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   Else
      strMsg = "無符合的資料"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   IsSel = False
   If IsHaveNp202 Or IsHaveCp202 Then
        For nRow = 1 To grdList.Rows - 1
           grdList.row = nRow
           grdList.col = 1
           If grdList.CellBackColor = &H8000000D Then
               IsSel = True
               Exit For
           End If
        Next nRow
        If IsSel = False Then
            strMsg = "母案有期限，請選擇一筆承接期限"
            strTit = "資料檢核"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
        End If
   End If
   
   If txtDate.Visible = True Then
      If txtDate.Text = "" Then
        strMsg = "定稿上的公報日期不可空白"
        strTit = "資料檢核"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        txtDate.SetFocus
        GoTo EXITSUB
      End If
      Cancel = False
      txtDate_Validate Cancel
      If Cancel = True Then GoTo EXITSUB
   End If
   

   
   strCP05 = DBDATE(oStrCDate)
   SeekGrdIndex = 0
   Me.Hide
   On Error GoTo oErr
   InputFun
EXITSUB:
Exit Sub
oErr:
    cnnConnection.RollbackTrans
    MsgBox "存檔失敗！", vbInformation
    Me.Show
    'Resume Next
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010401_3.m_strIR01
   m_strIR02 = frm02010401_3.m_strIR02
   m_strIR03 = frm02010401_3.m_strIR03
   m_strIR04 = frm02010401_3.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm02010401_6 = Nothing
End Sub

Sub StrMenu()
Dim strNationCode As String
With rsTmp
    If rsTmp.State = 1 Then rsTmp.Close
    Set rsTmp = New ADODB.Recordset
    strSql = "select c1.cp01 as cp01,c1.cp02 as cp02,c1.cp03 as cp03,c1.cp04 as cp04,c1.cp12 as cp12,c1.cp13 as cp13,tm01,tm02,tm03,tm04,tm05,tm10,tm08,tm09,tm23,c2.cp09 as cp09 from caseprogress c1,divisioncase,trademark,caseprogress c2 where c1.cp09='" & oKey & "' and c1.cp01=dc05(+) and c1.cp02=dc06(+) and c1.cp03=dc07(+) and c1.cp04=dc08(+) and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+) and dc01=c2.cp01(+) and dc02=c2.cp02(+) and dc03=c2.cp03(+) and dc04=c2.cp04(+) and c2.cp10='308' "
    strSql = strSql & " ORDER BY TM01,TM02,TM03,TM04"  'ADD BY SONIA 2016/8/4
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        m_TM01 = CheckStr(.Fields("cp01"))
        m_TM02 = CheckStr(.Fields("cp02"))
        m_TM03 = CheckStr(.Fields("cp03"))
        m_TM04 = CheckStr(.Fields("cp04"))
        m_CP12 = CheckStr(.Fields("cp12"))
        m_CP13 = CheckStr(.Fields("cp13"))
        InitialGrdList
        .MoveFirst
        Do While .EOF = False
            strNationCode = Empty
            grdList.Rows = grdList.Rows + 1
            grdList.row = grdList.Rows - 1
            ' 本所案號欄位
            grdList.TextMatrix(grdList.row, 1) = .Fields("TM01") & .Fields("TM02") & .Fields("TM03") & .Fields("TM04")
            ' 商標名稱欄位
            If IsNull(.Fields("TM05")) = False Then
               grdList.TextMatrix(grdList.row, 2) = .Fields("TM05")
            End If
            ' 申請國家
            If IsNull(.Fields("TM10")) = False Then
               strNationCode = .Fields("TM10")
               grdList.TextMatrix(grdList.row, 6) = GetNationName(.Fields("TM10"), 0)
            End If
            ' 商標種類欄位
            If IsNull(.Fields("TM08")) = False Then
               If strNationCode < "010" Then
                  grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(.Fields("TM08"), 0)
               Else
                  grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(.Fields("TM08"), 1)
               End If
            End If
            ' 商品類別
            If IsNull(.Fields("TM09")) = False Then
               grdList.TextMatrix(grdList.row, 4) = .Fields("TM09")
            End If
            ' 申請人
            If IsNull(.Fields("TM23")) = False Then
               grdList.TextMatrix(grdList.row, 5) = GetCustomerName(.Fields("TM23"), 0)
            End If
            ' 隱藏起來的本所案號
            grdList.TextMatrix(grdList.row, 7) = .Fields("TM01")
            grdList.TextMatrix(grdList.row, 8) = .Fields("TM02")
            grdList.TextMatrix(grdList.row, 9) = .Fields("TM03")
            grdList.TextMatrix(grdList.row, 10) = .Fields("TM04")
            grdList.TextMatrix(grdList.row, 11) = .Fields("cp09")
            .MoveNext
        Loop
        '檢查母案有無未收文或未發文的申請意見書
        IsHaveNp202 = False
        strSql = "select np01 from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np06 is null and np07=202 "
        Set rsTmp = New ADODB.Recordset
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 Then
            IsHaveNp202 = True
            Label1.Visible = True
        End If
        If IsHaveNp202 = False Then
            IsHaveCp202 = False
            strSql = " select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='202' and cp27 is null "
            Set rsTmp = New ADODB.Recordset
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 Then
                IsHaveCp202 = True
                Label1.Visible = True
            End If
        End If
        '控制畫面上輸入項目的出現
        If m_TM01 = "FCT" Then
            If IsHaveTM15 = True Then
                'edit by nickc 2007/08/10
                'txtLaw.Visible = False
                'Label3.Visible = False
                txtDate.Visible = True
                Label2.Visible = True
            Else
                'edit by nickc 2007/08/10
                'txtLaw.Visible = True
                'Label3.Visible = True
                txtDate.Visible = False
                Label2.Visible = False
            End If
        Else
            'edit by nickc 2007/08/10
            'txtLaw.Visible = False
            'Label3.Visible = False
            txtDate.Visible = False
            Label2.Visible = False
        End If
        'Added by Lydia 2023/10/18
        If grdList.Rows >= 2 Then
           grdList.FixedRows = 1
        End If
        'end 2023/10/18
    End If
End With
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   'edit by nickc 2007/08/10
   'grdList.Cols = 69
   grdList.Cols = 70
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "本所案號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "商標名稱"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "商標種類"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "商品類別"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "申請人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "申請國家"
   grdList.ColWidth(6) = 1200
   ' 本所案號 欄位一
   grdList.col = 7
   grdList.Text = Empty
   grdList.ColWidth(7) = 0
   ' 本所案號 欄位二
   grdList.col = 8
   grdList.Text = Empty
   grdList.ColWidth(8) = 0
   ' 本所案號 欄位三
   grdList.col = 9
   grdList.Text = Empty
   grdList.ColWidth(9) = 0
   ' 本所案號 欄位四
   grdList.col = 10
   grdList.Text = Empty
   grdList.ColWidth(10) = 0
   ' 收文號 欄位五
   grdList.col = 11
   grdList.Text = Empty
   grdList.ColWidth(11) = 0
   '儲存傳回的資料
   For i = 12 To grdList.Cols - 1
        grdList.col = i
        grdList.Text = Empty
        grdList.ColWidth(i) = 0
   Next i
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 7
      m_CP01 = grdList.Text
      grdList.col = 8
      m_CP02 = grdList.Text
      grdList.col = 9
      m_CP03 = grdList.Text
      grdList.col = 10
      m_CP04 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub
' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = 1
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

Public Sub InputFun()
    Dim IsSel As Boolean
    SeekGrdIndex = SeekGrdIndex + 1
    If SeekGrdIndex >= grdList.Rows Then
        SaveData
        'Modified by Morgan 2017/5/3 電子公文
        'Unload Me
        If Me.m_DocNo <> "" Then
            cmdExit_Click
            frm02010412.GoNext
        Else
            Unload Me
        End If
        'end 2017/5/3
        Exit Sub
    End If
    IsSel = False
    If IsHaveNp202 = True Or IsHaveCp202 = True Then
        grdList.row = SeekGrdIndex
        grdList.col = 1
        If grdList.CellBackColor = &H8000000D Then
            IsSel = True
        End If
    End If
    Select Case m_TM01
    Case "T"
        If IsHaveTM15 = True Then    '有審定號代表註冊後分割，要輸註冊號
            Set frm02010404_3.UpForm = Me
            
            'Added by Morgan 2024/7/11 電子公文
            frm02010404_3.m_DocWord = m_DocWord
            frm02010404_3.m_DocNo = m_DocNo
            frm02010404_3.m_DocPdf = m_DocPdf
            frm02010404_3.m_DocPdfDate = m_DocPdfDate
            frm02010404_3.m_DocPdfTime = m_DocPdfTime
            'end 2024/7/11
                            
            frm02010404_3.SetData 0, grdList.TextMatrix(SeekGrdIndex, 7), True
            frm02010404_3.SetData 1, grdList.TextMatrix(SeekGrdIndex, 8)
            frm02010404_3.SetData 2, grdList.TextMatrix(SeekGrdIndex, 9)
            frm02010404_3.SetData 3, grdList.TextMatrix(SeekGrdIndex, 10)
            frm02010404_3.SetData 4, oStrCDate
            frm02010404_3.m_MonCP09 = oKey
            If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                frm02010404_3.textNP08.Visible = True
                frm02010404_3.textNP09.Visible = True
                frm02010404_3.Label1(18).Visible = True
                frm02010404_3.Label1(17).Visible = True
                frm02010404_3.textNP08.Enabled = True
                frm02010404_3.textNP09.Enabled = True
                'Add By Sindy 2012/5/18
                frm02010404_3.LabNP07 = "202"
                frm02010404_3.Label32.Visible = True
                frm02010404_3.Frame1.Visible = True
                frm02010404_3.Frame2.Visible = True
                '2012/5/18 End
            Else
                frm02010404_3.textNP08.Visible = False
                frm02010404_3.textNP09.Visible = False
                frm02010404_3.Label1(18).Visible = False
                frm02010404_3.Label1(17).Visible = False
                frm02010404_3.textNP08.Enabled = False
                frm02010404_3.textNP09.Enabled = False
                'Add By Sindy 2012/5/18
                frm02010404_3.LabNP07 = ""
                frm02010404_3.Label32.Visible = False
                frm02010404_3.Frame1.Visible = False
                frm02010404_3.Frame2.Visible = False
                '2012/5/18 End
            End If
            'edit by nickc 2008/01/23 加入可以取消
            'frm02010404_3.cmdCancel.Visible = False
            frm02010404_3.cmdExit.Visible = False
            frm02010404_3.textCP47.Enabled = False
            frm02010404_3.textEditPrint.Visible = False
            frm02010404_3.Label11.Visible = False
            frm02010404_3.Label12.Visible = False
            frm02010404_3.textTM21.Enabled = False
            frm02010404_3.textTM22.Enabled = False
            frm02010404_3.textTM14.Enabled = False
            frm02010404_3.QueryData
            'add by nickc 2008/01/23 加入可以取消
            m_IsCancal = False
            frm02010404_3.Show vbModal
            'add by nickc 2008/01/23 加入可以取消
            If m_IsCancal = True Then
                Me.Show
                Exit Sub
            End If
            '將回傳的資料記錄下來
            grdList.TextMatrix(SeekGrdIndex, 12) = PutSeekData01
            grdList.TextMatrix(SeekGrdIndex, 13) = PutSeekData02
            grdList.TextMatrix(SeekGrdIndex, 14) = PutSeekData03
            grdList.TextMatrix(SeekGrdIndex, 15) = PutSeekData04
            grdList.TextMatrix(SeekGrdIndex, 16) = PutSeekData05
            grdList.TextMatrix(SeekGrdIndex, 17) = PutSeekData06
            grdList.TextMatrix(SeekGrdIndex, 18) = PutSeekData07
            grdList.TextMatrix(SeekGrdIndex, 19) = PutSeekData08
            grdList.TextMatrix(SeekGrdIndex, 20) = PutSeekData09
            grdList.TextMatrix(SeekGrdIndex, 21) = PutSeekData10
            grdList.TextMatrix(SeekGrdIndex, 22) = PutSeekData11
            grdList.TextMatrix(SeekGrdIndex, 23) = PutSeekData12
            grdList.TextMatrix(SeekGrdIndex, 24) = PutSeekData13
            grdList.TextMatrix(SeekGrdIndex, 25) = PutSeekData14
        Else   '無審定號要輸申請案號
            Set frm02010301_2.UpForm = Me
            
            'Added by Morgan 2024/7/11 電子公文
            frm02010301_2.m_DocWord = m_DocWord
            frm02010301_2.m_DocNo = m_DocNo
            frm02010301_2.m_DocPdf = m_DocPdf
            frm02010301_2.m_DocPdfDate = m_DocPdfDate
            frm02010301_2.m_DocPdfTime = m_DocPdfTime
            'end 2024/7/11
            
            frm02010301_2.SetData grdList.TextMatrix(SeekGrdIndex, 7), grdList.TextMatrix(SeekGrdIndex, 8), grdList.TextMatrix(SeekGrdIndex, 9), grdList.TextMatrix(SeekGrdIndex, 10), grdList.TextMatrix(SeekGrdIndex, 11)
            frm02010301_2.m_MonCP09 = oKey
            If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                frm02010301_2.textNP08.Visible = True
                frm02010301_2.textNP09.Visible = True
                frm02010301_2.Label1(18).Visible = True
                frm02010301_2.Label1(17).Visible = True
                frm02010301_2.textNP08.Enabled = True
                frm02010301_2.textNP09.Enabled = True
                'Add By Sindy 2012/5/18
                frm02010301_2.LabNP07 = "202"
                frm02010301_2.Label32.Visible = True
                frm02010301_2.Frame1.Visible = True
                frm02010301_2.Frame2.Visible = True
                '2012/5/18 End
            Else
                frm02010301_2.textNP08.Visible = False
                frm02010301_2.textNP09.Visible = False
                frm02010301_2.Label1(18).Visible = False
                frm02010301_2.Label1(17).Visible = False
                frm02010301_2.textNP08.Enabled = False
                frm02010301_2.textNP09.Enabled = False
                'Add By Sindy 2012/5/18
                frm02010301_2.LabNP07 = ""
                frm02010301_2.Label32.Visible = False
                frm02010301_2.Frame1.Visible = False
                frm02010301_2.Frame2.Visible = False
                '2012/5/18 End
            End If
            'edit by nickc 2008/01/23 加入可以取消
            'frm02010301_2.cmdCancel.Visible = False
            frm02010301_2.cmdExit.Visible = False
            frm02010301_2.textCP47.Enabled = False
            frm02010301_2.UpdateCtrl
            'add by nickc 2008/01/23 加入可以取消
            m_IsCancal = False
            frm02010301_2.Show vbModal
            'add by nickc 2008/01/23 加入可以取消
            If m_IsCancal = True Then
                Me.Show
                Exit Sub
            End If
            '將回傳的資料記錄下來
            grdList.TextMatrix(SeekGrdIndex, 12) = PutSeekData01
            grdList.TextMatrix(SeekGrdIndex, 13) = PutSeekData02
            grdList.TextMatrix(SeekGrdIndex, 14) = PutSeekData03
            grdList.TextMatrix(SeekGrdIndex, 15) = PutSeekData04
            grdList.TextMatrix(SeekGrdIndex, 16) = PutSeekData05
            grdList.TextMatrix(SeekGrdIndex, 17) = PutSeekData06
            grdList.TextMatrix(SeekGrdIndex, 18) = PutSeekData07
            grdList.TextMatrix(SeekGrdIndex, 19) = PutSeekData08
            grdList.TextMatrix(SeekGrdIndex, 20) = PutSeekData09
            grdList.TextMatrix(SeekGrdIndex, 21) = PutSeekData10
        End If
    Case "FCT"
        If IsHaveTM15 = True Then    '有審定號代表註冊後分割，要輸註冊號
            Set frm03020404_03.UpForm = Me
            
            'Added by Morgan 2024/7/11 電子公文
            frm03020404_03.m_DocWord = m_DocWord
            frm03020404_03.m_DocNo = m_DocNo
            frm03020404_03.m_DocPdf = m_DocPdf
            frm03020404_03.m_DocPdfDate = m_DocPdfDate
            frm03020404_03.m_DocPdfTime = m_DocPdfTime
            'end 2024/7/11
                
            frm03020404_03.SetData 0, grdList.TextMatrix(SeekGrdIndex, 7), True
            frm03020404_03.SetData 1, grdList.TextMatrix(SeekGrdIndex, 8)
            frm03020404_03.SetData 2, grdList.TextMatrix(SeekGrdIndex, 9)
            frm03020404_03.SetData 3, grdList.TextMatrix(SeekGrdIndex, 10)
            frm03020404_03.SetData 4, oStrCDate
            frm03020404_03.m_MonCP09 = oKey
            If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                frm03020404_03.textNP08.Visible = True
                frm03020404_03.textNP09.Visible = True
                frm03020404_03.Label1(18).Visible = True
                frm03020404_03.Label1(17).Visible = True
                frm03020404_03.textNP08.Enabled = True
                frm03020404_03.textNP09.Enabled = True
                'Add By Sindy 2012/5/18
                frm03020404_03.LabNP07 = "202"
                frm03020404_03.Label32.Visible = True
                frm03020404_03.Frame1.Visible = True
                frm03020404_03.Frame2.Visible = True
                '2012/5/18 End
            Else
                frm03020404_03.textNP08.Visible = False
                frm03020404_03.textNP09.Visible = False
                frm03020404_03.Label1(18).Visible = False
                frm03020404_03.Label1(17).Visible = False
                frm03020404_03.textNP08.Enabled = False
                frm03020404_03.textNP09.Enabled = False
                'Add By Sindy 2012/5/18
                frm03020404_03.LabNP07 = ""
                frm03020404_03.Label32.Visible = False
                frm03020404_03.Frame1.Visible = False
                frm03020404_03.Frame2.Visible = False
                '2012/5/18 End
            End If
            'edit by nickc 2008/01/23 加入可以取消
            'frm03020404_03.cmdCancel.Visible = False
            frm03020404_03.cmdExit.Visible = False
            frm03020404_03.textTM14.Enabled = False
            frm03020404_03.textTM21.Enabled = False
            frm03020404_03.textTM22.Enabled = False
            frm03020404_03.textCreFee.Enabled = False
            frm03020404_03.Combo2.Enabled = False
            frm03020404_03.QueryData
            'add by nickc 2008/01/23 加入可以取消
            m_IsCancal = False
            frm03020404_03.Show vbModal
            'add by nickc 2008/01/23 加入可以取消
            If m_IsCancal = True Then
                Me.Show
                Exit Sub
            End If
            '將回傳的資料記錄下來
            grdList.TextMatrix(SeekGrdIndex, 12) = PutSeekData01
            grdList.TextMatrix(SeekGrdIndex, 13) = PutSeekData02
            grdList.TextMatrix(SeekGrdIndex, 14) = PutSeekData03
            grdList.TextMatrix(SeekGrdIndex, 15) = PutSeekData04
            grdList.TextMatrix(SeekGrdIndex, 16) = PutSeekData05
            grdList.TextMatrix(SeekGrdIndex, 17) = PutSeekData06
            grdList.TextMatrix(SeekGrdIndex, 18) = PutSeekData07
            grdList.TextMatrix(SeekGrdIndex, 19) = PutSeekData08
            grdList.TextMatrix(SeekGrdIndex, 20) = PutSeekData09
            grdList.TextMatrix(SeekGrdIndex, 21) = PutSeekData10
            grdList.TextMatrix(SeekGrdIndex, 22) = PutSeekData11
        Else   '無審定號要輸申請案號
            Set frm030203_02.UpForm = Me
            
            'Added by Morgan 2024/7/11 電子公文
            frm030203_02.m_DocWord = m_DocWord
            frm030203_02.m_DocNo = m_DocNo
            frm030203_02.m_DocPdf = m_DocPdf
            frm030203_02.m_DocPdfDate = m_DocPdfDate
            frm030203_02.m_DocPdfTime = m_DocPdfTime
            'end 2024/7/11
            
            frm030203_02.SetData 0, grdList.TextMatrix(SeekGrdIndex, 7), True
            frm030203_02.SetData 1, grdList.TextMatrix(SeekGrdIndex, 8)
            frm030203_02.SetData 2, grdList.TextMatrix(SeekGrdIndex, 9)
            frm030203_02.SetData 3, grdList.TextMatrix(SeekGrdIndex, 10)
            frm030203_02.SetData 4, grdList.TextMatrix(SeekGrdIndex, 11)
            frm030203_02.m_MonCP09 = oKey
            If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                frm030203_02.textNP08.Visible = True
                frm030203_02.textNP09.Visible = True
                frm030203_02.Label1(18).Visible = True
                frm030203_02.Label1(17).Visible = True
                frm030203_02.textNP08.Enabled = True
                frm030203_02.textNP09.Enabled = True
                'add by nickc 2007/08/10
                frm030203_02.txtLaw.Visible = True
                frm030203_02.Label16.Visible = True
                'Add By Sindy 2012/5/18
                frm030203_02.LabNP07 = "202"
                frm030203_02.Label32.Visible = True
                frm030203_02.Frame1.Visible = True
                frm030203_02.Frame2.Visible = True
                '2012/5/18 End
            Else
                frm030203_02.textNP08.Visible = False
                frm030203_02.textNP09.Visible = False
                frm030203_02.Label1(18).Visible = False
                frm030203_02.Label1(17).Visible = False
                frm030203_02.textNP08.Enabled = False
                frm030203_02.textNP09.Enabled = False
                'add by nickc 2007/08/10
                frm030203_02.txtLaw.Visible = False
                frm030203_02.Label16.Visible = False
                'Add By Sindy 2012/5/18
                frm030203_02.LabNP07 = ""
                frm030203_02.Label32.Visible = False
                frm030203_02.Frame1.Visible = False
                frm030203_02.Frame2.Visible = False
                '2012/5/18 End
            End If
            'edit by nickc 2008/01/23 加入可以取消
            'frm030203_02.cmdCancel.Visible = False
            frm030203_02.cmdExit.Visible = False
            frm030203_02.textTM11.Enabled = False
            frm030203_02.cmdPriority.Enabled = False
            frm030203_02.textPriorityDoc.Enabled = False
            frm030203_02.SSTab1.TabEnabled(1) = False
            frm030203_02.QueryData
            'add by nickc 2008/01/23 加入可以取消
            m_IsCancal = False
            frm030203_02.Show vbModal
            'add by nickc 2008/01/23 加入可以取消
            If m_IsCancal = True Then
                Me.Show
                Exit Sub
            End If
            '將回傳的資料記錄下來
            grdList.TextMatrix(SeekGrdIndex, 12) = PutSeekData01
            grdList.TextMatrix(SeekGrdIndex, 13) = PutSeekData02
            grdList.TextMatrix(SeekGrdIndex, 14) = PutSeekData03
            grdList.TextMatrix(SeekGrdIndex, 15) = PutSeekData04
            grdList.TextMatrix(SeekGrdIndex, 16) = PutSeekData05
            grdList.TextMatrix(SeekGrdIndex, 17) = PutSeekData06
            grdList.TextMatrix(SeekGrdIndex, 18) = PutSeekData07
            grdList.TextMatrix(SeekGrdIndex, 19) = PutSeekData08
            grdList.TextMatrix(SeekGrdIndex, 20) = PutSeekData09
            grdList.TextMatrix(SeekGrdIndex, 21) = PutSeekData10
            grdList.TextMatrix(SeekGrdIndex, 22) = PutSeekData11
            grdList.TextMatrix(SeekGrdIndex, 23) = PutSeekData12
            grdList.TextMatrix(SeekGrdIndex, 24) = PutSeekData13
            grdList.TextMatrix(SeekGrdIndex, 25) = PutSeekData14
            grdList.TextMatrix(SeekGrdIndex, 26) = PutSeekData15
            grdList.TextMatrix(SeekGrdIndex, 27) = PutSeekData16
            grdList.TextMatrix(SeekGrdIndex, 28) = PutSeekData17
            'add by nickc 2007/05/01  加入代表人
            grdList.TextMatrix(SeekGrdIndex, 29) = PutSeekData18
            grdList.TextMatrix(SeekGrdIndex, 30) = PutSeekData19
            grdList.TextMatrix(SeekGrdIndex, 31) = PutSeekData20
            grdList.TextMatrix(SeekGrdIndex, 32) = PutSeekData21
            grdList.TextMatrix(SeekGrdIndex, 33) = PutSeekData22
            grdList.TextMatrix(SeekGrdIndex, 34) = PutSeekData23
            grdList.TextMatrix(SeekGrdIndex, 35) = PutSeekData24
            grdList.TextMatrix(SeekGrdIndex, 36) = PutSeekData25
            grdList.TextMatrix(SeekGrdIndex, 37) = PutSeekData26
            grdList.TextMatrix(SeekGrdIndex, 38) = PutSeekData27
            grdList.TextMatrix(SeekGrdIndex, 39) = PutSeekData28
            grdList.TextMatrix(SeekGrdIndex, 40) = PutSeekData29
            grdList.TextMatrix(SeekGrdIndex, 41) = PutSeekData30
            grdList.TextMatrix(SeekGrdIndex, 42) = PutSeekData31
            grdList.TextMatrix(SeekGrdIndex, 43) = PutSeekData32
            grdList.TextMatrix(SeekGrdIndex, 44) = PutSeekData33
            grdList.TextMatrix(SeekGrdIndex, 45) = PutSeekData34
            grdList.TextMatrix(SeekGrdIndex, 46) = PutSeekData35
            grdList.TextMatrix(SeekGrdIndex, 47) = PutSeekData36
            grdList.TextMatrix(SeekGrdIndex, 48) = PutSeekData37
            grdList.TextMatrix(SeekGrdIndex, 49) = PutSeekData38
            grdList.TextMatrix(SeekGrdIndex, 50) = PutSeekData39
            grdList.TextMatrix(SeekGrdIndex, 51) = PutSeekData40
            grdList.TextMatrix(SeekGrdIndex, 52) = PutSeekData41
            grdList.TextMatrix(SeekGrdIndex, 53) = PutSeekData42
            grdList.TextMatrix(SeekGrdIndex, 54) = PutSeekData43
            grdList.TextMatrix(SeekGrdIndex, 55) = PutSeekData44
            grdList.TextMatrix(SeekGrdIndex, 56) = PutSeekData45
            grdList.TextMatrix(SeekGrdIndex, 57) = PutSeekData46
            grdList.TextMatrix(SeekGrdIndex, 58) = PutSeekData47
            grdList.TextMatrix(SeekGrdIndex, 59) = PutSeekData48
            grdList.TextMatrix(SeekGrdIndex, 60) = PutSeekData49
            grdList.TextMatrix(SeekGrdIndex, 61) = PutSeekData50
            grdList.TextMatrix(SeekGrdIndex, 62) = PutSeekData51
            grdList.TextMatrix(SeekGrdIndex, 63) = PutSeekData52
            grdList.TextMatrix(SeekGrdIndex, 64) = PutSeekData53
            grdList.TextMatrix(SeekGrdIndex, 65) = PutSeekData54
            grdList.TextMatrix(SeekGrdIndex, 66) = PutSeekData55
            grdList.TextMatrix(SeekGrdIndex, 67) = PutSeekData56
            grdList.TextMatrix(SeekGrdIndex, 68) = PutSeekData57
            grdList.TextMatrix(SeekGrdIndex, 69) = PutSeekData58
        End If
    Case Else
    End Select
    InputFun
End Sub

'存檔兼產生定稿
Function SaveData() As Boolean
Dim ijk As Integer
Dim IsSel As Boolean
Dim StrExtString As String
Dim rsMe As New ADODB.Recordset
Dim StrSQLMe As String
Dim tmpChild_Tm15 As String
Dim tmpChild_Tm67 As String
Dim tmpChild_Tm09 As String
Dim strSubData As String 'Add By Sindy 2010/6/2
Dim arrTM09 As Variant, strGoodsKind As String 'Add By Sindy 2010/11/12
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean, ET03_1 As String
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End

On Error GoTo oErr

cnnConnection.BeginTrans

   'Add By Sindy 2012/1/13
   ET01 = "03"
   ET02 = oKey
   bolEdit = False
   '2012/1/13 End
    
    Select Case m_TM01
    Case "T"
        If IsHaveTM15 = True Then    '有審定號代表註冊後分割，要輸註冊號
             '更新母案核准及結果日
             strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & oKey & "' "
             cnnConnection.Execute strSql
            '新增母案 C 來文
             strCP09 = AutoNo("C", 6)
             strCP05 = DBDATE(oStrCDate)
             strCP27 = "null"
             ' 組成SQL語法
             strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
                      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "','" & "N" & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & oKey & "')"
             ' 新增資料到資料庫
             cnnConnection.Execute strSql
             
            'Added by Morgan 2017/6/14 電子公文
            If m_DocNo <> "" Then
               strSql = "update caseprogress set cp08='" & m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號' where cp09='" & strCP09 & "'"
               cnnConnection.Execute strSql, intI
               '下載公文給子案用
               m_DocPdf = "$" & m_DocNo & ".pdf"
               If PUB_GetAttachFile_CPP(m_DocNo, m_DocPdf, App.path & "\" & strUserNum) = True Then
                  Set oFile = oFileSys.GetFile(m_DocPdf)
                  m_DocPdfDate = Format(oFile.DateLastModified, "YYYYMMDD")
                  m_DocPdfTime = Format(oFile.DateLastModified, "HHMMSS")
                  PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, "1001"
               End If
            End If
            'end 2017/6/14
            
             For ijk = 1 To grdList.Rows - 1
                SeekGrdIndex = ijk
                IsSel = False
                If IsHaveNp202 = True Or IsHaveCp202 = True Then
                    grdList.row = SeekGrdIndex
                    grdList.col = 1
                    If grdList.CellBackColor = &H8000000D Then
                        IsSel = True
                    End If
                End If
                 Load frm02010404_3
                Set frm02010404_3.UpForm = Me
                frm02010404_3.SetData 0, grdList.TextMatrix(SeekGrdIndex, 7), True
                frm02010404_3.SetData 1, grdList.TextMatrix(SeekGrdIndex, 8)
                frm02010404_3.SetData 2, grdList.TextMatrix(SeekGrdIndex, 9)
                frm02010404_3.SetData 3, grdList.TextMatrix(SeekGrdIndex, 10)
                frm02010404_3.SetData 4, oStrCDate
                frm02010404_3.m_MonCP09 = oKey
                If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                    frm02010404_3.textNP08.Visible = True
                    frm02010404_3.textNP09.Visible = True
                    frm02010404_3.textNP08.Enabled = True
                    frm02010404_3.textNP09.Enabled = True
                Else
                    frm02010404_3.textNP08.Visible = False
                    frm02010404_3.textNP09.Visible = False
                    frm02010404_3.textNP08.Enabled = False
                    frm02010404_3.textNP09.Enabled = False
                End If
                frm02010404_3.cmdCancel.Visible = False
                frm02010404_3.cmdExit.Visible = False
                frm02010404_3.textCP47.Enabled = False
                frm02010404_3.QueryData
                frm02010404_3.textTM15 = grdList.TextMatrix(SeekGrdIndex, 12)
                frm02010404_3.textTM14 = grdList.TextMatrix(SeekGrdIndex, 13)
                'frm02010404_3.Text1 = grdList.TextMatrix(SeekGrdIndex, 14)
                frm02010404_3.textTM21 = grdList.TextMatrix(SeekGrdIndex, 15)
                frm02010404_3.textTM22 = grdList.TextMatrix(SeekGrdIndex, 16)
                frm02010404_3.textCP47 = grdList.TextMatrix(SeekGrdIndex, 17)
                frm02010404_3.textDate = grdList.TextMatrix(SeekGrdIndex, 18)
                frm02010404_3.textMoney = grdList.TextMatrix(SeekGrdIndex, 19)
                frm02010404_3.textTC1 = grdList.TextMatrix(SeekGrdIndex, 20)
                frm02010404_3.textTC2 = grdList.TextMatrix(SeekGrdIndex, 21)
                frm02010404_3.textPrint = grdList.TextMatrix(SeekGrdIndex, 22)
                frm02010404_3.textPS = grdList.TextMatrix(SeekGrdIndex, 23)
                frm02010404_3.textNP08 = grdList.TextMatrix(SeekGrdIndex, 24)
                frm02010404_3.textNP09 = grdList.TextMatrix(SeekGrdIndex, 25)
                
                'Added by Morgan 2017/6/14 電子公文
                frm02010404_3.m_DocWord = m_DocWord
                frm02010404_3.m_DocNo = m_DocNo
                frm02010404_3.m_DocPdf = m_DocPdf
                frm02010404_3.m_DocPdfDate = m_DocPdfDate
                frm02010404_3.m_DocPdfTime = m_DocPdfTime
                'end 2017/6/14
                
                'frm02010404_3.cmdok_Click
                Call frm02010404_3.cmdOK_Click(0)  'Modify By Sindy 2009/05/14
                
                'Add By Sindy 2010/11/12
                m_TM09 = ""
                StrExtString = "select * from trademark where tm01='" & grdList.TextMatrix(SeekGrdIndex, 7) & "' and tm02='" & grdList.TextMatrix(SeekGrdIndex, 8) & "' and tm03='" & grdList.TextMatrix(SeekGrdIndex, 9) & "' and tm04='" & grdList.TextMatrix(SeekGrdIndex, 10) & "' "
                Set rsMe = New ADODB.Recordset
                If rsMe.State = 1 Then rsMe.Close
                rsMe.CursorLocation = adUseClient
                rsMe.Open StrExtString, cnnConnection, adOpenStatic, adLockReadOnly
                If rsMe.RecordCount <> 0 Then
                    m_TM09 = CheckStr(rsMe.Fields("TM09"))
                End If
                '1-34商品 35-45服務
                strGoodsKind = "商品"
                If Trim(m_TM09) > "" Then
                  arrTM09 = Split(m_TM09, ",")
                  If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
                     strGoodsKind = "服務"
                  End If
                End If
                '2010/11/12 End
                
                If grdList.TextMatrix(SeekGrdIndex, 22) = "1" Then   '<===此句與下段不同喔
                    EndLetter "03", grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "12", strUserNum
                    'Add By Sindy 2010/11/12
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308" & "','" & "12" & "','" & strUserNum & "'," & _
                                 "'商品或服務','" & strGoodsKind & "')"
                    cnnConnection.Execute strSql
                    '2010/11/12 End
'                    NowPrint grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "03", "12", False, strUserNum, 0
                     'Modify By Sindy 2012/1/13
                     ET02 = grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308"
                     ET03 = "12"
                     '2012/1/13 End
                     Call SaveNowPrint(ET01, ET02, ET03, bolEdit, ET03_1, grdList.TextMatrix(SeekGrdIndex, 7), grdList.TextMatrix(SeekGrdIndex, 8), grdList.TextMatrix(SeekGrdIndex, 9), grdList.TextMatrix(SeekGrdIndex, 10)) 'Add By Sindy 2012/5/3
                'Add By Sindy 2013/11/5
                ElseIf grdList.TextMatrix(SeekGrdIndex, 22) = "2" Then
                     EndLetter "03", grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "20", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308" & "','" & "20" & "','" & strUserNum & "'," & _
                                 "'商品或服務','" & strGoodsKind & "')"
                    cnnConnection.Execute strSql
                     ET02 = grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308"
                     ET03 = "20"
                     Call SaveNowPrint(ET01, ET02, ET03, bolEdit, ET03_1, grdList.TextMatrix(SeekGrdIndex, 7), grdList.TextMatrix(SeekGrdIndex, 8), grdList.TextMatrix(SeekGrdIndex, 9), grdList.TextMatrix(SeekGrdIndex, 10))
                End If
                '2013/11/5 END
             Next ijk
             '印結案單，改最後一起印
            PUB_PrintCaseCloseSheet strUserNum, "0", False, False
            '刪除暫存資料
            PUB_DeleteCaseCloseSheet strUserNum
        Else   '無審定號要輸申請案號

            '更新母案核准及結果日
             strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & oKey & "' "
             cnnConnection.Execute strSql
            '新增母案 C 來文
             strCP09 = AutoNo("C", 6)
             strCP05 = DBDATE(oStrCDate)
             strCP27 = "null"
             ' 組成SQL語法
             strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
                      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "','" & "N" & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & oKey & "')"
             ' 新增資料到資料庫
             cnnConnection.Execute strSql
             
            'Added by Morgan 2017/6/14 電子公文
            If m_DocNo <> "" Then
               strSql = "update caseprogress set cp08='" & m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號' where cp09='" & strCP09 & "'"
               cnnConnection.Execute strSql, intI
               '下載公文給子案用
               m_DocPdf = "$" & m_DocNo & ".pdf"
               If PUB_GetAttachFile_CPP(m_DocNo, m_DocPdf, App.path & "\" & strUserNum) = True Then
                  Set oFile = oFileSys.GetFile(m_DocPdf)
                  m_DocPdfDate = Format(oFile.DateLastModified, "YYYYMMDD")
                  m_DocPdfTime = Format(oFile.DateLastModified, "HHMMSS")
                  PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, "1001"
               End If
            End If
            'end 2017/6/14
   
             'add by nickc  2006/11/09 將母案下一程序的催審，皆上不續辦
             strSql = "update nextprogress set np06='N' where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=305 "
             cnnConnection.Execute strSql
             For ijk = 1 To grdList.Rows - 1
                SeekGrdIndex = ijk
                IsSel = False
                If IsHaveNp202 = True Or IsHaveCp202 = True Then
                    grdList.row = SeekGrdIndex
                    grdList.col = 1
                    If grdList.CellBackColor = &H8000000D Then
                        IsSel = True
                    End If
                End If
                 Load frm02010301_2
                Set frm02010301_2.UpForm = Me
                frm02010301_2.SetData grdList.TextMatrix(SeekGrdIndex, 7), grdList.TextMatrix(SeekGrdIndex, 8), grdList.TextMatrix(SeekGrdIndex, 9), grdList.TextMatrix(SeekGrdIndex, 10), grdList.TextMatrix(SeekGrdIndex, 11)
                frm02010301_2.m_MonCP09 = oKey
                If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                    frm02010301_2.textNP08.Visible = True
                    frm02010301_2.textNP09.Visible = True
                    frm02010301_2.textNP08.Enabled = True
                    frm02010301_2.textNP09.Enabled = True
                Else
                    frm02010301_2.textNP08.Visible = False
                    frm02010301_2.textNP09.Visible = False
                    frm02010301_2.textNP08.Enabled = False
                    frm02010301_2.textNP09.Enabled = False
                End If
                frm02010301_2.cmdCancel.Visible = False
                frm02010301_2.cmdExit.Visible = False
                frm02010301_2.UpdateCtrl
                frm02010301_2.textCP47 = grdList.TextMatrix(SeekGrdIndex, 12)
                frm02010301_2.textCP30 = grdList.TextMatrix(SeekGrdIndex, 13)
                frm02010301_2.textTM27 = grdList.TextMatrix(SeekGrdIndex, 14)
                frm02010301_2.textCP05 = grdList.TextMatrix(SeekGrdIndex, 15)
                frm02010301_2.textPrint = grdList.TextMatrix(SeekGrdIndex, 16)
                frm02010301_2.textCF09 = grdList.TextMatrix(SeekGrdIndex, 17)
                frm02010301_2.textCP45 = grdList.TextMatrix(SeekGrdIndex, 18)
                frm02010301_2.textPS = grdList.TextMatrix(SeekGrdIndex, 19)
                frm02010301_2.textNP08 = grdList.TextMatrix(SeekGrdIndex, 20)
                frm02010301_2.textNP09 = grdList.TextMatrix(SeekGrdIndex, 21)
                
                'Added by Morgan 2017/6/14 電子公文
                frm02010301_2.m_DocWord = m_DocWord
                frm02010301_2.m_DocNo = m_DocNo
                frm02010301_2.m_DocPdf = m_DocPdf
                frm02010301_2.m_DocPdfDate = m_DocPdfDate
                frm02010301_2.m_DocPdfTime = m_DocPdfTime
                'end 2017/6/14
                
                'frm02010301_2.cmdok_Click
                Call frm02010301_2.cmdOK_Click(0)  'Modify By Sindy 2009/05/14
                
                'Add By Sindy 2010/11/12
                m_TM09 = ""
                StrExtString = "select * from trademark where tm01='" & grdList.TextMatrix(SeekGrdIndex, 7) & "' and tm02='" & grdList.TextMatrix(SeekGrdIndex, 8) & "' and tm03='" & grdList.TextMatrix(SeekGrdIndex, 9) & "' and tm04='" & grdList.TextMatrix(SeekGrdIndex, 10) & "' "
                Set rsMe = New ADODB.Recordset
                If rsMe.State = 1 Then rsMe.Close
                rsMe.CursorLocation = adUseClient
                rsMe.Open StrExtString, cnnConnection, adOpenStatic, adLockReadOnly
                If rsMe.RecordCount <> 0 Then
                    m_TM09 = CheckStr(rsMe.Fields("TM09"))
                End If
                '1-34商品 35-45服務
                strGoodsKind = "商品"
                If Trim(m_TM09) > "" Then
                  arrTM09 = Split(m_TM09, ",")
                  If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
                     strGoodsKind = "服務"
                  End If
                End If
                '2010/11/12 End
                
                If grdList.TextMatrix(SeekGrdIndex, 16) = "1" Then  '<===此句與上段不同喔
                    EndLetter "03", grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "11", strUserNum
                    'Add By Sindy 2010/11/12
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308" & "','" & "11" & "','" & strUserNum & "'," & _
                                 "'商品或服務','" & strGoodsKind & "')"
                    cnnConnection.Execute strSql
                    '2010/11/12 End
'                    NowPrint grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "03", "11", False, strUserNum, 0
                     'Modify By Sindy 2012/1/13
                     ET02 = grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308"
                     ET03 = "11"
                     '2012/1/13 End
                     Call SaveNowPrint(ET01, ET02, ET03, bolEdit, ET03_1, grdList.TextMatrix(SeekGrdIndex, 7), grdList.TextMatrix(SeekGrdIndex, 8), grdList.TextMatrix(SeekGrdIndex, 9), grdList.TextMatrix(SeekGrdIndex, 10)) 'Add By Sindy 2012/5/3
                '2010/11/19 add by sonia
                ElseIf grdList.TextMatrix(SeekGrdIndex, 16) = "2" Then
                    EndLetter "03", grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "09", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308" & "','" & "09" & "','" & strUserNum & "'," & _
                                 "'商品或服務','" & strGoodsKind & "')"
                    cnnConnection.Execute strSql
'                    NowPrint grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308", "03", "09", False, strUserNum, 0
                     'Modify By Sindy 2012/1/13
                     ET02 = grdList.TextMatrix(SeekGrdIndex, 7) & grdList.TextMatrix(SeekGrdIndex, 8) & grdList.TextMatrix(SeekGrdIndex, 9) & grdList.TextMatrix(SeekGrdIndex, 10) & "&308"
                     ET03 = "09"
                     '2012/1/13 End
                     Call SaveNowPrint(ET01, ET02, ET03, bolEdit, ET03_1, grdList.TextMatrix(SeekGrdIndex, 7), grdList.TextMatrix(SeekGrdIndex, 8), grdList.TextMatrix(SeekGrdIndex, 9), grdList.TextMatrix(SeekGrdIndex, 10)) 'Add By Sindy 2012/5/3
                '2010/11/19 end
                End If
            Next ijk
        End If
        
         'Add by Sindy 2019/5/10
         Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
         If m_strIR01 <> "" Then
            PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010401_1", strCP09
         End If
         '2019/5/10 END
        
         cnnConnection.CommitTrans
         
    Case "FCT"
        If IsHaveTM15 = True Then    '有審定號代表註冊後分割，要輸註冊號
            '更新母案核准及結果日
             strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & oKey & "' "
             cnnConnection.Execute strSql
             '新增母案 C 來文
             strCP09 = AutoNo("C", 6)
             strCP05 = DBDATE(oStrCDate)
             strCP27 = "null"
             ' 組成SQL語法
             strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
                      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "','" & "N" & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & oKey & "')"
             ' 新增資料到資料庫
             cnnConnection.Execute strSql
             
            'Added by Morgan 2017/6/14 電子公文
            If m_DocNo <> "" Then
               strSql = "update caseprogress set cp08='" & m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號' where cp09='" & strCP09 & "'"
               cnnConnection.Execute strSql, intI
               '下載公文給子案用
               m_DocPdf = "$" & m_DocNo & ".pdf"
               If PUB_GetAttachFile_CPP(m_DocNo, m_DocPdf, App.path & "\" & strUserNum) = True Then
                  Set oFile = oFileSys.GetFile(m_DocPdf)
                  m_DocPdfDate = Format(oFile.DateLastModified, "YYYYMMDD")
                  m_DocPdfTime = Format(oFile.DateLastModified, "HHMMSS")
                  PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, "1001"
               End If
            End If
            'end 2017/6/14
               
             strSubData = "" '子案資料
             For ijk = 1 To grdList.Rows - 1
                SeekGrdIndex = ijk
                IsSel = False
                If IsHaveNp202 = True Or IsHaveCp202 = True Then
                    grdList.row = SeekGrdIndex
                    grdList.col = 1
                    If grdList.CellBackColor = &H8000000D Then
                        IsSel = True
                    End If
                End If
                 Load frm03020404_03
                Set frm03020404_03.UpForm = Me
                frm03020404_03.SetData 0, grdList.TextMatrix(SeekGrdIndex, 7), True
                frm03020404_03.SetData 1, grdList.TextMatrix(SeekGrdIndex, 8)
                frm03020404_03.SetData 2, grdList.TextMatrix(SeekGrdIndex, 9)
                frm03020404_03.SetData 3, grdList.TextMatrix(SeekGrdIndex, 10)
                frm03020404_03.SetData 4, oStrCDate
                frm03020404_03.m_MonCP09 = oKey
                If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                    frm03020404_03.textNP08.Visible = True
                    frm03020404_03.textNP09.Visible = True
                    frm03020404_03.Label1(18).Visible = True
                    frm03020404_03.Label1(17).Visible = True
                    frm03020404_03.textNP08.Enabled = True
                    frm03020404_03.textNP09.Enabled = True
                Else
                    frm03020404_03.textNP08.Visible = False
                    frm03020404_03.textNP09.Visible = False
                    frm03020404_03.Label1(18).Visible = False
                    frm03020404_03.Label1(17).Visible = False
                    frm03020404_03.textNP08.Enabled = False
                    frm03020404_03.textNP09.Enabled = False
                    
                End If
                '將回傳的資料記錄下來
                frm03020404_03.cmdCancel.Visible = False
                frm03020404_03.cmdExit.Visible = False
                frm03020404_03.QueryData
                frm03020404_03.textTM14 = grdList.TextMatrix(SeekGrdIndex, 12)
                frm03020404_03.textTM15 = grdList.TextMatrix(SeekGrdIndex, 13)
                frm03020404_03.textTM21 = grdList.TextMatrix(SeekGrdIndex, 14)
                frm03020404_03.textTM22 = grdList.TextMatrix(SeekGrdIndex, 15)
                frm03020404_03.Text1 = grdList.TextMatrix(SeekGrdIndex, 16)
                frm03020404_03.textCreFee = grdList.TextMatrix(SeekGrdIndex, 17)
                frm03020404_03.textPrint = grdList.TextMatrix(SeekGrdIndex, 18)
                frm03020404_03.textPrtTrans = grdList.TextMatrix(SeekGrdIndex, 19)
                frm03020404_03.Text2 = grdList.TextMatrix(SeekGrdIndex, 20)
                frm03020404_03.textNP08 = grdList.TextMatrix(SeekGrdIndex, 21)
                frm03020404_03.textNP09 = grdList.TextMatrix(SeekGrdIndex, 22)
                
                'Added by Morgan 2017/6/14 電子公文
                frm03020404_03.m_DocWord = m_DocWord
                frm03020404_03.m_DocNo = m_DocNo
                frm03020404_03.m_DocPdf = m_DocPdf
                frm03020404_03.m_DocPdfDate = m_DocPdfDate
                frm03020404_03.m_DocPdfTime = m_DocPdfTime
                'end 2017/6/14
            
                'frm03020404_03.cmdok_Click
                Call frm03020404_03.cmdOK_Click(0)  'Modify By Sindy 2009/05/14
                m_TM67 = ""
                m_TM09 = ""
                m_TM15 = ""
                StrExtString = "select * from trademark where tm01='" & grdList.TextMatrix(SeekGrdIndex, 7) & "' and tm02='" & grdList.TextMatrix(SeekGrdIndex, 8) & "' and tm03='" & grdList.TextMatrix(SeekGrdIndex, 9) & "' and tm04='" & grdList.TextMatrix(SeekGrdIndex, 10) & "' "
                Set rsMe = New ADODB.Recordset
                If rsMe.State = 1 Then rsMe.Close
                rsMe.CursorLocation = adUseClient
                rsMe.Open StrExtString, cnnConnection, adOpenStatic, adLockReadOnly
                If rsMe.RecordCount <> 0 Then
                    m_TM67 = CheckStr(rsMe.Fields("TM67"))
                    m_TM09 = CheckStr(rsMe.Fields("TM09"))
                    m_TM15 = CheckStr(rsMe.Fields("TM15"))
                End If
                'Add By Sindy 2010/6/2
                '子案資料
                strSubData = strSubData & _
                                     "Registration No.: " & m_TM15 & "　　　(Our Ref: " & grdList.TextMatrix(SeekGrdIndex, 7) & "-" & grdList.TextMatrix(SeekGrdIndex, 8) & IIf(Trim(grdList.TextMatrix(SeekGrdIndex, 9)) = "0" And Trim(grdList.TextMatrix(SeekGrdIndex, 10)) = "00", "", Trim(grdList.TextMatrix(SeekGrdIndex, 9)) & Trim(grdList.TextMatrix(SeekGrdIndex, 10))) & ")" & vbCrLf & _
                                     "Class : " & m_TM09 & vbCrLf & _
                                     "Goods/services designated : " & vbCrLf & _
                                     "|?TMGoods:" & grdList.TextMatrix(SeekGrdIndex, 7) & "-" & grdList.TextMatrix(SeekGrdIndex, 8) & "-" & grdList.TextMatrix(SeekGrdIndex, 9) & "-" & grdList.TextMatrix(SeekGrdIndex, 10) & "-英文?|" & vbCrLf
                '2010/6/2 End
             Next ijk
             
             'Modify By Sindy 2019/9/11
            'FCT定稿語文為日文之案件若無定稿，請不要帶英文定稿，例FCT-41430分割核准
            ' 定稿語文
            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
               ' 英文
               Case "2":
            '2019/9/11 END
                '【函】
                '英文-註冊後分割核准通知函
                EndLetter "03", oKey, "03", strUserNum
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('03','" & oKey & "','03','" & strUserNum & "'," & _
                    "'" & "公報日期" & "','" & ChangeTStringToWString(txtDate) & "')"
                cnnConnection.Execute strSql
   '             NowPrint oKey, "03", "03", False, strUserNum, 0
               ET03 = "03" 'Modify By Sindy 2012/1/13
               
                'Modify By Sindy 2010/6/2
                '英文-註冊後分割核准譯文
                EndLetter "03", oKey, "04", strUserNum
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('03','" & oKey & "','04','" & strUserNum & "'," & _
                        "'" & "子案案件數" & "','" & CStr(grdList.Rows - 1) & "')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('03','" & oKey & "','04','" & strUserNum & "'," & _
                        "'" & "公報日期" & "','" & ChangeTStringToWString(txtDate) & "')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('03','" & oKey & "','04','" & strUserNum & "'," & _
                        "'" & "子案資料" & "','" & strSubData & "')"
                cnnConnection.Execute strSql
   '             NowPrint oKey, "03", "04", False, strUserNum, 0
               ET03_1 = "04" 'Modify By Sindy 2012/1/13
                '2010/6/2 End
             End Select
             
        Else   '無審定號要輸申請案號
            '更新母案核准及結果日
             strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & oKey & "' "
             cnnConnection.Execute strSql
            '新增母案 C 來文
             strCP09 = AutoNo("C", 6)
             strCP05 = DBDATE(oStrCDate)
             strCP27 = "null"
             ' 組成SQL語法
             strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
                      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "','" & "N" & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & oKey & "')"
             ' 新增資料到資料庫
             cnnConnection.Execute strSql
             
            'Added by Morgan 2017/6/14 電子公文
            If m_DocNo <> "" Then
               strSql = "update caseprogress set cp08='" & m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號' where cp09='" & strCP09 & "'"
               cnnConnection.Execute strSql, intI
               '下載公文給子案用
               m_DocPdf = "$" & m_DocNo & ".pdf"
               If PUB_GetAttachFile_CPP(m_DocNo, m_DocPdf, App.path & "\" & strUserNum) = True Then
                  Set oFile = oFileSys.GetFile(m_DocPdf)
                  m_DocPdfDate = Format(oFile.DateLastModified, "YYYYMMDD")
                  m_DocPdfTime = Format(oFile.DateLastModified, "HHMMSS")
                  PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, "1001"
               End If
            End If
            'end 2017/6/14
            
             'add by nickc  2006/11/09 將母案下一程序的催審，皆上不續辦
             strSql = "update nextprogress set np06='N' where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=305 "
             cnnConnection.Execute strSql
             m_NP08 = "": m_NP09 = ""
             m_TM12 = ""
             'add by nickc 2007/08/10
             m_Law = ""
             
             Dim ExString As String
             ExString = ""
             For ijk = 1 To grdList.Rows - 1
                SeekGrdIndex = ijk
                IsSel = False
                If IsHaveNp202 = True Or IsHaveCp202 = True Then
                    grdList.row = SeekGrdIndex
                    grdList.col = 1
                    If grdList.CellBackColor = &H8000000D Then
                        IsSel = True
                    End If
                End If
                 Load frm030203_02
                Set frm030203_02.UpForm = Me
                frm030203_02.SetData 0, grdList.TextMatrix(SeekGrdIndex, 7), True
                frm030203_02.SetData 1, grdList.TextMatrix(SeekGrdIndex, 8)
                frm030203_02.SetData 2, grdList.TextMatrix(SeekGrdIndex, 9)
                frm030203_02.SetData 3, grdList.TextMatrix(SeekGrdIndex, 10)
                frm030203_02.SetData 4, grdList.TextMatrix(SeekGrdIndex, 11)
                frm030203_02.m_MonCP09 = oKey
                If (IsHaveNp202 Or IsHaveCp202) And IsSel Then    '秀出新本所期限和法定期限
                    frm030203_02.textNP08.Visible = True
                    frm030203_02.textNP09.Visible = True
                    frm030203_02.Label1(18).Visible = True
                    frm030203_02.Label1(17).Visible = True
                    frm030203_02.textNP08.Enabled = True
                    frm030203_02.textNP09.Enabled = True
                Else
                    frm030203_02.textNP08.Visible = False
                    frm030203_02.textNP09.Visible = False
                    frm030203_02.Label1(18).Visible = False
                    frm030203_02.Label1(17).Visible = False
                    frm030203_02.textNP08.Enabled = False
                    frm030203_02.textNP09.Enabled = False
                End If
                frm030203_02.cmdCancel.Visible = False
                frm030203_02.cmdExit.Visible = False
                frm030203_02.textTM11.Enabled = False
                frm030203_02.cmdPriority.Enabled = False
                frm030203_02.textPriorityDoc.Enabled = False
                frm030203_02.SSTab1.TabEnabled(1) = False
                '將回傳的資料記錄下來
                frm030203_02.cmdCancel.Visible = False
                frm030203_02.cmdExit.Visible = False
                frm030203_02.QueryData
                frm030203_02.textTM09 = grdList.TextMatrix(SeekGrdIndex, 12)
                frm030203_02.textTM32 = grdList.TextMatrix(SeekGrdIndex, 13)
                frm030203_02.textTM11 = grdList.TextMatrix(SeekGrdIndex, 14)
                frm030203_02.textTM12 = grdList.TextMatrix(SeekGrdIndex, 15)
                frm030203_02.textTM27 = grdList.TextMatrix(SeekGrdIndex, 16)
                frm030203_02.textCP05 = grdList.TextMatrix(SeekGrdIndex, 17)
                frm030203_02.textPrint = grdList.TextMatrix(SeekGrdIndex, 18)
                frm030203_02.textPriorityDoc = grdList.TextMatrix(SeekGrdIndex, 19)
                frm030203_02.textAddDate = grdList.TextMatrix(SeekGrdIndex, 20)
                frm030203_02.textAdd = grdList.TextMatrix(SeekGrdIndex, 21)
                frm030203_02.textDN = grdList.TextMatrix(SeekGrdIndex, 22)
                frm030203_02.txtToEng = grdList.TextMatrix(SeekGrdIndex, 23)
                frm030203_02.textPrtTrans = grdList.TextMatrix(SeekGrdIndex, 24)
                frm030203_02.textPS = grdList.TextMatrix(SeekGrdIndex, 25)
                frm030203_02.textTM67 = grdList.TextMatrix(SeekGrdIndex, 26)
                frm030203_02.textNP08 = grdList.TextMatrix(SeekGrdIndex, 27)
                frm030203_02.textNP09 = grdList.TextMatrix(SeekGrdIndex, 28)
                'add by nickc 2007/05/01 加入代理人
                If frm030203_02.SSTab1.TabEnabled(1) = True Then
                    frm030203_02.Combo2(0).Text = grdList.TextMatrix(SeekGrdIndex, 29)
                    frm030203_02.Combo2(1).Text = grdList.TextMatrix(SeekGrdIndex, 30)
                    frm030203_02.Combo2(2).Text = grdList.TextMatrix(SeekGrdIndex, 31)
                    frm030203_02.Combo2(3).Text = grdList.TextMatrix(SeekGrdIndex, 32)
                    frm030203_02.Combo2(4).Text = grdList.TextMatrix(SeekGrdIndex, 33)
                    frm030203_02.Combo2(5).Text = grdList.TextMatrix(SeekGrdIndex, 34)
                    frm030203_02.Combo2(6).Text = grdList.TextMatrix(SeekGrdIndex, 35)
                    frm030203_02.Combo2(7).Text = grdList.TextMatrix(SeekGrdIndex, 36)
                    frm030203_02.Combo2(8).Text = grdList.TextMatrix(SeekGrdIndex, 37)
                    frm030203_02.Combo2(9).Text = grdList.TextMatrix(SeekGrdIndex, 38)
                End If
                frm030203_02.textTM47 = grdList.TextMatrix(SeekGrdIndex, 39)
                frm030203_02.textTM48 = grdList.TextMatrix(SeekGrdIndex, 40)
                frm030203_02.textTM49 = grdList.TextMatrix(SeekGrdIndex, 41)
                frm030203_02.textTM50 = grdList.TextMatrix(SeekGrdIndex, 42)
                frm030203_02.textTM51 = grdList.TextMatrix(SeekGrdIndex, 43)
                frm030203_02.textTM52 = grdList.TextMatrix(SeekGrdIndex, 44)
                frm030203_02.textTM94 = grdList.TextMatrix(SeekGrdIndex, 45)
                frm030203_02.textTM95 = grdList.TextMatrix(SeekGrdIndex, 46)
                frm030203_02.textTM96 = grdList.TextMatrix(SeekGrdIndex, 47)
                frm030203_02.textTM97 = grdList.TextMatrix(SeekGrdIndex, 48)
                frm030203_02.textTM98 = grdList.TextMatrix(SeekGrdIndex, 49)
                frm030203_02.textTM99 = grdList.TextMatrix(SeekGrdIndex, 50)
                frm030203_02.textTM100 = grdList.TextMatrix(SeekGrdIndex, 51)
                frm030203_02.textTM101 = grdList.TextMatrix(SeekGrdIndex, 52)
                frm030203_02.textTM102 = grdList.TextMatrix(SeekGrdIndex, 53)
                frm030203_02.textTM103 = grdList.TextMatrix(SeekGrdIndex, 54)
                frm030203_02.textTM104 = grdList.TextMatrix(SeekGrdIndex, 55)
                frm030203_02.textTM105 = grdList.TextMatrix(SeekGrdIndex, 56)
                frm030203_02.TextTM106 = grdList.TextMatrix(SeekGrdIndex, 57)
                frm030203_02.TextTM107 = grdList.TextMatrix(SeekGrdIndex, 58)
                frm030203_02.textTM108 = grdList.TextMatrix(SeekGrdIndex, 59)
                frm030203_02.TextTM109 = grdList.TextMatrix(SeekGrdIndex, 60)
                frm030203_02.TextTM110 = grdList.TextMatrix(SeekGrdIndex, 61)
                frm030203_02.textTM111 = grdList.TextMatrix(SeekGrdIndex, 62)
                frm030203_02.TextTM112 = grdList.TextMatrix(SeekGrdIndex, 63)
                frm030203_02.TextTM113 = grdList.TextMatrix(SeekGrdIndex, 64)
                frm030203_02.textTM114 = grdList.TextMatrix(SeekGrdIndex, 65)
                frm030203_02.TextTM115 = grdList.TextMatrix(SeekGrdIndex, 66)
                frm030203_02.TextTM116 = grdList.TextMatrix(SeekGrdIndex, 67)
                frm030203_02.textTM117 = grdList.TextMatrix(SeekGrdIndex, 68)
                'add by nickc 2007/08/10
                frm030203_02.txtLaw = grdList.TextMatrix(SeekGrdIndex, 69)
                If frm030203_02.textNP08.Enabled = True Then
                    '2014/12/9 MODIFY BY SONIA 改通知法定期限
                    'm_NP08 = grdList.TextMatrix(SeekGrdIndex, 27)
                    m_NP09 = grdList.TextMatrix(SeekGrdIndex, 28)
                    '2014/12/9 END
                    m_TM12 = grdList.TextMatrix(SeekGrdIndex, 15)
                    'add by nickc 2007/08/10
                    m_Law = grdList.TextMatrix(SeekGrdIndex, 69)
                End If
                
                'Added by Morgan 2017/6/14 電子公文
                frm030203_02.m_DocWord = m_DocWord
                frm030203_02.m_DocNo = m_DocNo
                frm030203_02.m_DocPdf = m_DocPdf
                frm030203_02.m_DocPdfDate = m_DocPdfDate
                frm030203_02.m_DocPdfTime = m_DocPdfTime
                'end 2017/6/14
                
                'frm030203_02.cmdok_Click
                Call frm030203_02.cmdOK_Click(0)  'Modify By Sindy 2009/05/14
                'edit by nickc 2008/01/25 琬姿說，第二個以後往前一格
                If ExString = "" Then ExString = ExString & " "
                'ExString = ExString & "   Application No.: " & grdList.TextMatrix(SeekGrdIndex, 15) & vbCrLf
                ExString = ExString & "  Application No.: " & grdList.TextMatrix(SeekGrdIndex, 15) & vbCrLf
                ExString = ExString & "  Our Ref: " & grdList.TextMatrix(SeekGrdIndex, 7) & "-" & grdList.TextMatrix(SeekGrdIndex, 8) & "-" & grdList.TextMatrix(SeekGrdIndex, 9) & "-" & grdList.TextMatrix(SeekGrdIndex, 10) & vbCrLf
                'edit by nickc 2008/01/25 琬姿說要判斷
                'ExString = ExString & "  Class: " & grdList.TextMatrix(SeekGrdIndex, 12) & vbCrLf
                ExString = ExString & "  Class" & IIf(InStr(1, grdList.TextMatrix(SeekGrdIndex, 12), ",") > 0, "es", "") & ": " & grdList.TextMatrix(SeekGrdIndex, 12) & vbCrLf
                'add by nickc 2008/01/25 琬姿說要空一行
                ExString = ExString & vbCrLf
                
                'Modify By Sindy 2009/07/06
                'ExString = ExString & "  Goods/Service designated: |?TMGoods:" & grdList.TextMatrix(SeekGrdIndex, 7) & "-" & grdList.TextMatrix(SeekGrdIndex, 8) & "-" & grdList.TextMatrix(SeekGrdIndex, 9) & "-" & grdList.TextMatrix(SeekGrdIndex, 10) & "?|" & vbCrLf & vbCrLf
                ExString = ExString & "  Goods/Service designated: |?TMGoods:" & grdList.TextMatrix(SeekGrdIndex, 7) & "-" & grdList.TextMatrix(SeekGrdIndex, 8) & "-" & grdList.TextMatrix(SeekGrdIndex, 9) & "-" & grdList.TextMatrix(SeekGrdIndex, 10) & "-英文?|" & vbCrLf & vbCrLf
            Next ijk
            ExString = "2. " & Mid(ExString, 4)
            
            'Modify By Sindy 2019/9/11
            'FCT定稿語文為日文之案件若無定稿，請不要帶英文定稿，例FCT-41430分割核准
            ' 定稿語文
            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
               ' 英文
               Case "2":
            '2019/9/11 END
               '函
               EndLetter "03", oKey, "02", strUserNum
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','02','" & strUserNum & "'," & _
                   "'" & "例申請案號" & "','" & m_TM12 & "')"
               cnnConnection.Execute strSql
               'edit by nickc  2007/08/10
               'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','02','" & strUserNum & "'," & _
                   "'" & "例法條" & "','" & txtLaw & "')"
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','02','" & strUserNum & "'," & _
                   "'" & "例法條" & "','" & m_Law & "')"
               cnnConnection.Execute strSql
               '2014/12/9 MODIFY BY SONIA 改通知法定期限
               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','02','" & strUserNum & "'," & _
                   "'" & "例本所期限" & "','" & DBDATE(m_NP08) & "')"
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','02','" & strUserNum & "'," & _
                   "'" & "例法定期限" & "','" & DBDATE(m_NP09) & "')"
               '2014/12/9 END
               cnnConnection.Execute strSql
   '            NowPrint oKey, "03", "02", False, strUserNum, 0
               ET03 = "02" 'Modify By Sindy 2012/1/13
               
               '譯文
               EndLetter "03", oKey, "07", strUserNum
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','07','" & strUserNum & "'," & _
                   "'" & "子案案件數" & "','" & ShowNumber(grdList.Rows - 1) & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','07','" & strUserNum & "'," & _
                   "'" & "例子案" & "','" & ExString & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','07','" & strUserNum & "'," & _
                   "'" & "例申請案號" & "','" & m_TM12 & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','07','" & strUserNum & "'," & _
                   "'" & "例法條" & "','" & m_Law & "')"
               cnnConnection.Execute strSql
               '2014/12/9 MODIFY BY SONIA 改通知法定期限
               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','07','" & strUserNum & "'," & _
                   "'" & "例本所期限" & "','" & DBDATE(m_NP08) & "')"
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('03','" & oKey & "','07','" & strUserNum & "'," & _
                   "'" & "例法定期限" & "','" & DBDATE(m_NP09) & "')"
               '2014/12/9 END
               cnnConnection.Execute strSql
   '            NowPrint oKey, "03", "07", False, strUserNum, 0
               ET03_1 = "07" 'Modify By Sindy 2012/1/13
            End Select
        End If
        '2010/6/25 Add By Sindy 核准分割要印地址條
        pub_AddressListSN = pub_AddressListSN + 1
        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
         
         'Add by Sindy 2019/5/10
         Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
         If m_strIR01 <> "" Then
            PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010401_1", strCP09
         End If
         '2019/5/10 END
   
         cnnConnection.CommitTrans
         
         Call SaveNowPrint(ET01, ET02, ET03, bolEdit, ET03_1, m_TM01, m_TM02, m_TM03, m_TM04) 'Modify By Sindy 2012/5/3
'         'Add By Sindy 2012/1/13
'         If ET03 <> "" Then
'            bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
'            If bolEmail Then
'               '判斷是否EMail同時寄紙本
'               If Not bolPlusPaper Then
'                  iCopy = 1
'               End If
'               NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
'               If ET03_1 <> "" Then
'                  NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , True, True
'               End If
'               MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
'            Else
'               NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0
'               If ET03_1 <> "" Then
'                  NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0
'               End If
'            End If
'         End If
'         '2012/1/13 End
    Case Else
    End Select
    
Me.Show
MsgBox "存檔成功！", vbInformation
If UCase(UpForm.Name) = "FRM02010401_3" Then   '從內商來
    Unload frm02010401_2
   'Add By Sindy 2019/5/10
   If Me.m_strIR01 <> "" Then
     Unload frm02010401_1
     If Not m_PrevForm Is Nothing Then
        Call m_PrevForm.GoNext
     End If
   Else
   '2019/5/10 END
      frm02010401_1.m_txtTMBM07_1 = ""
      frm02010401_1.m_txtTMBM07_2 = ""
      frm02010401_1.m_txtTM14 = ""
      frm02010401_1.m_blnNotFirst = True
      frm02010401_1.Show
   End If
ElseIf UCase(UpForm.Name) = "FRM03020401_03" Then  '從外商來
   Unload frm03020401_02
   'Add By Sindy 2019/5/10
   If Me.m_strIR01 <> "" Then
     Unload frm03020401_01
     If Not m_PrevForm Is Nothing Then
        Call m_PrevForm.GoNext
     End If
   Else
   '2019/5/10 END
      frm03020401_01.Show
   End If
End If
Unload UpForm
Exit Function

oErr:
    cnnConnection.RollbackTrans
    MsgBox "存檔失敗！", vbCritical
    Me.Show
    'Resume Next
End Function

Private Sub txtDate_GotFocus()
    InverseTextBox txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(txtDate) = False Then
      ' 檢查是否為民國年
      If CheckIsTaiwanDate(txtDate, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的公報日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtDate_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2012/5/3
Private Sub SaveNowPrint(ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean, ET03_1 As String, strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String)
Dim bolPlusPaper As Boolean, bolEmail As Boolean, iCopy As Integer
Dim strLD18 As String 'Add By Sindy 2021/1/25
   
   If ET03 <> "" Then
      'Add By Sindy 2021/1/25 商標電子化
      If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(strTM01, 1) = "T" Then
         strSql = " select cp09 from caseprogress" & _
            " where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "'" & _
            " and substr(cp09,1,1)='C' and cp27=" & strSrvDate(1)
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strLD18 = rsTmp.Fields("CP09")
         End If
      End If
      '2021/1/25 END
      
      bolEmail = PUB_GetEMailFlag(strTM01 & strTM02 & strTM03 & strTM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2021/1/25 + 信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True, , , , , strLD18
         If ET03_1 <> "" Then
            'Add By Sindy 2021/1/25 + 信函總收文號
            NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , True, True, , , , , strLD18
         End If
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(strTM01) & " ]！"
      Else
         'Add By Sindy 2021/1/25 + 信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         If ET03_1 <> "" Then
            'Add By Sindy 2021/1/25 + 信函總收文號
            NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         End If
      End If
   End If
End Sub
