VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm030208 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子檔整批匯入"
   ClientHeight    =   4824
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   9084
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4824
   ScaleWidth      =   9084
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frm030208.frx":0000
      Left            =   1740
      List            =   "frm030208.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   870
      Width           =   2445
   End
   Begin VB.Frame Frame3 
      Caption         =   "匯入錯誤訊息："
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
      Height          =   2895
      Left            =   90
      TabIndex        =   11
      Top             =   1260
      Width           =   8865
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2352
         ItemData        =   "frm030208.frx":0004
         Left            =   90
         List            =   "frm030208.frx":0006
         TabIndex        =   12
         Top             =   270
         Width           =   8685
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   8520
      TabIndex        =   1
      Top             =   480
      Width           =   345
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6390
      TabIndex        =   3
      Top             =   60
      Width           =   940
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5370
      TabIndex        =   2
      Top             =   60
      Width           =   940
   End
   Begin VB.TextBox txtPath1 
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
      Left            =   1740
      TabIndex        =   0
      Text            =   "C:\temp"
      Top             =   480
      Width           =   6765
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7410
      TabIndex        =   4
      Top             =   60
      Width           =   940
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   90
      TabIndex        =   8
      Top             =   4260
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   8820
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   840
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   720
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frm030208.frx":0008
   End
   Begin VB.FileListBox File1 
      Height          =   432
      Left            =   1380
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1155
      Left            =   4590
      TabIndex        =   10
      Top             =   450
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6054
      _ExtentY        =   2032
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱                                                             "
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "匯入的檔案類型："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   930
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "電子檔存放路徑："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1560
   End
End
Attribute VB_Name = "frm030208"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/10/15 Form2.0已檢查 (無需修改的物件)
'Create By Lydia 2020/07/22 電子檔整批匯入(FCT專用)
Option Explicit

Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine1 As Integer
Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
Dim intUpdStarRow As Integer, intUpdEndRow As Integer
Dim strUpdCP01 As String, strUpdCP02 As String, strUpdCP03 As String, strUpdCP04 As String
Dim strUpdCP09 As String, strUpdCP10 As String
Dim m_strReFileN As String '副檔名

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194


Private Sub cmdExit_Click()
   Unload Me
End Sub

'匯入
Private Sub cmdImPort_Click()
Dim fs
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim strTotRow As String
Dim strCaseNo As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strFileName As String
Dim strErr As String
Dim strTemp As String
Dim jj As Integer
Dim varTmp As Variant
   
On Error GoTo ErrHand
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   m_strReFileN = ""

    varTmp = Split(Combo1.Text, " ")
    If UCase(varTmp(0)) = "0" Then '發文
       m_strReFileN = "DATA"
    ElseIf UCase(varTmp(0)) = "1" Then '收據
       m_strReFileN = "RECEIPT"
    End If
    If m_strReFileN = "" Then
       MsgBox "匯入的檔案類型不可空白", vbExclamation
       Exit Sub
    End If
   
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   
   '檢查資料夾
   Set fs = CreateObject("Scripting.FileSystemObject")
   File1.path = txtPath1.Text
   File1.Refresh
   If File1.ListCount = 0 Then
      MsgBox txtPath1.Text & " 此資料夾中，尚無電子檔！"
      txtPath1.SetFocus
      Exit Sub
   End If
   Set fs = Nothing
   
   dblMaxWidth = 8820
   Text2.Width = 0
   List1.Clear
   Grid2.Clear
   Grid2.Cols = 1
   Grid2.Rows = 1

   For dblFCnt = 0 To File1.ListCount - 1
      '檢查檔案是否正在使用中
      If PUB_ChkFileOpening(txtPath1.Text & "\" & Trim(File1.List(dblFCnt))) = True Then
         MsgBox Trim(File1.List(dblFCnt)) & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！", vbExclamation
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      '檔名後4碼為.PDF者才須匯入
      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
         Grid2.AddItem Trim(File1.List(dblFCnt))
      End If
   Next dblFCnt

   Grid2.col = 0
   Grid2.row = 0
   Me.Grid2.Sort = 5 '字串昇冪
   
   strTotRow = Grid2.Rows - 1
   
   If Val(strTotRow) = 0 Then
      MsgBox "無資料！", vbInformation
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   '清空變數值
   intUpdStarRow = 0
   intUpdEndRow = 0
   strUpdCP01 = ""
   strUpdCP02 = ""
   strUpdCP03 = ""
   strUpdCP04 = ""
   strUpdCP09 = ""
   strUpdCP10 = ""
   For dblFCnt = 1 To strTotRow
      Text2.Width = dblMaxWidth / Val(strTotRow) * dblFCnt: DoEvents
      strErr = "": strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
      strFileName = UCase(Grid2.TextMatrix(dblFCnt, 0))
      strFileName = Mid(strFileName, 1, InStrRev(strFileName, ".") - 1) '去掉.PDF
      
      '取得案號
      If InStr(strFileName, ".") > 0 Then
         strCaseNo = Trim(Left(strFileName, InStr(strFileName, ".") - 1))
         'Modify By Sindy 2020/7/27 Mark,移至上面
'         '發文
'         If Trim(Left(Combo1.Text, 2)) = "0" Then
'            strFileName = Mid(strFileName, 1, InStrRev(strFileName, ".") - 1) '去掉.PDF
'         End If
      Else
         strCaseNo = strFileName
      End If
'      For jj = 1 To 3
'         If Asc(Mid(strCaseNo, jj, 1)) >= 65 And Asc(Mid(strCaseNo, jj, 1)) <= 90 And Len(strCP01) < 3 Then '系統別
'            strCP01 = strCP01 & Mid(strCaseNo, jj, 1)
'         Else
'            Exit For
'         End If
'      Next jj
      'Modify By Sindy 2020/7/27 檢查檔名前面是否為本所案號的系統別
      strSql = "select sk01 from Systemkind where sk01=substr('" & strFileName & "',1,length(sk01))"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         For jj = 1 To 3
            If Asc(Mid(strCaseNo, jj, 1)) >= 65 And Asc(Mid(strCaseNo, jj, 1)) <= 90 And Len(strCP01) < 3 Then '系統別
               strCP01 = strCP01 & Mid(strCaseNo, jj, 1)
            Else
               Exit For
            End If
         Next jj
      End If
      If strCP01 = "" Then
         '申請案號
         'Modify By Sindy 2021/3/16 + and tm29 is null and tm57 is null
         strSql = "select tm01,tm02,tm03,tm04 From trademark" & _
                  " where tm12='" & strCaseNo & "' and tm29 is null and tm57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.RecordCount = 1 Then
               strCP01 = RsTemp.Fields("tm01")
               strCP02 = RsTemp.Fields("tm02")
               strCP03 = RsTemp.Fields("tm03")
               strCP04 = RsTemp.Fields("tm04")
            End If
         End If
      End If
      If strCP01 = "" Then
         '註冊號
         'Modify By Sindy 2021/3/16 + and tm29 is null and tm57 is null
         'Modified by Lydia 2021/10/22 +限制FCT案 and tm01='FCT' ; ex.FCT-047982的審定號00075728和T-088124相同
         strSql = "select tm01,tm02,tm03,tm04 From trademark" & _
                  " where tm15='" & strCaseNo & "' and tm01='FCT' and tm29 is null and tm57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.RecordCount = 1 Then
               strCP01 = RsTemp.Fields("tm01")
               strCP02 = RsTemp.Fields("tm02")
               strCP03 = RsTemp.Fields("tm03")
               strCP04 = RsTemp.Fields("tm04")
            End If
         End If
      End If
      If strCP01 = "" Then
         For jj = 1 To 3
            If Asc(Mid(strCaseNo, jj, 1)) >= 65 And Asc(Mid(strCaseNo, jj, 1)) <= 90 And Len(strCP01) < 3 Then '系統別
               strCP01 = strCP01 & Mid(strCaseNo, jj, 1)
            Else
               Exit For
            End If
         Next jj
      End If
      '2020/7/27 END
      
      If CheckSys(strCP01) = "" Then
         strErr = convForm(CheckStr(strFileName), 30) & "，找不到本所案號"
         GoTo RunSave
      Else
         If strCP01 <> "FCT" Then
            strErr = convForm(CheckStr(strFileName), 30) & "，系統別有誤"
            GoTo RunSave
         End If
      End If
      If strCP02 = "" Then
         If InStr(strCaseNo, "-") = 0 Then
            strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1), "000000")
            strCP03 = "0"
            strCP04 = "00"
         Else
            strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1, InStr(strCaseNo, "-") - 1 - Len(strCP01)), "000000")
            strCP03 = Mid(strCaseNo, InStr(strCaseNo, "-") + 1, 1)
            If InStr(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), "-") > 0 Then
               strCP04 = Format(Mid(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), InStr(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), "-") + 1), "00")
            Else
               strCP04 = "00"
            End If
         End If
         '檢查strCP02的長度是否為6碼且為數字
         If Len(strCP02) <> 6 Then
            strErr = convForm(CheckStr(strFileName), 30) & "，案號CP02長度非6碼有誤"
            GoTo RunSave
         ElseIf IsNumeric(strCP02) = False Then
            strErr = convForm(CheckStr(strFileName), 30) & "，案號CP02非數字型態有誤"
            GoTo RunSave
         End If
         '檢查strCP03的長度是否為1碼且為數字
         If Len(strCP03) <> 1 Then
            strErr = convForm(CheckStr(strFileName), 30) & "，案號CP03長度非1碼有誤"
            GoTo RunSave
         'Mark by Lydia 2024/03/21 可以為英文/數字; Ex. FCT026472-T
         'ElseIf IsNumeric(strCP03) = False Then
         '   strErr = convForm(CheckStr(strFileName), 30) & "，案號CP03非數字型態有誤"
         '   GoTo RunSave
         'end 2024/03/21
         End If
         '檢查strCP04的長度是否為2碼且為數字
         If Len(strCP04) <> 2 Then
            strErr = convForm(CheckStr(strFileName), 30) & "，案號CP04長度非2碼有誤"
            GoTo RunSave
         ElseIf IsNumeric(strCP04) = False Then
            strErr = convForm(CheckStr(strFileName), 30) & "，案號CP04非數字型態有誤"
            GoTo RunSave
         End If
      End If
      
      '中文檔名 ex.申請書
      If strFileName <> PUB_GetSimpleName(strFileName) Then
         strErr = convForm(CheckStr(strFileName), 30) & "，不符檔案命名原則：含非英數字。"
         GoTo RunSave
      End If
      
      '取得收文號
      If strCP03 & strCP04 = "000" Then
         strCaseNo = strCP01 & strCP02
      Else
         strCaseNo = strCP01 & strCP02 & "-" & strCP03 & "-" & strCP04
      End If
      Call GetUpdCP09(strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt, strErr)
      
RunSave:
      If strUpdCP09 <> "" Or strErr <> "" Then
         If intUpdStarRow > 0 Then
            If intUpdStarRow > 0 And intUpdEndRow = 0 Then
               intUpdEndRow = intUpdStarRow
            End If
            If strUpdCP09 = "" Then
               If strErr = "" Then Call GetErrText(strFileName)
            Else
               Call SaveFilePDF '存檔
            End If
         End If
         '清空變數值
         intUpdStarRow = 0
         intUpdEndRow = 0
         strUpdCP09 = ""
         strUpdCP10 = ""
         If strErr <> "" Then
            If InStr(UCase(strErr), UCase(strFileName)) = 0 Then
               strErr = convForm(CheckStr(strFileName), 30) & "，" & Trim(strErr)
            End If
            List1.AddItem UCase(strErr), 0: SetListScroll List1
            strErr = ""
         End If
      End If
   Next dblFCnt
   
   Text2.Width = dblMaxWidth: DoEvents
   
   Screen.MousePointer = vbDefault
   
   MsgBox "匯入完畢！" & IIf(List1.ListCount > 0, vbCrLf & vbCrLf & "有匯入失敗的電子檔，請查看畫面上的【匯入錯誤訊息】", "")
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

Private Sub GetUpdCP09(strFileName As String, _
                       strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
                       dblFCnt As Double, ByRef strErrDesc As String)
Dim strConSql As String
Dim strTemp As String
   
   If intUpdStarRow = 0 Then
      strUpdCP01 = strCP01
      strUpdCP02 = strCP02
      strUpdCP03 = strCP03
      strUpdCP04 = strCP04
      
      intUpdStarRow = dblFCnt
   Else
      intUpdEndRow = dblFCnt
   End If
   strSql = "": strUpdCP09 = "": strConSql = ""
   
   '判斷傳入檔案的案件性質
   If InStr(strFileName, ".") > 0 Then
      'FCT可傳入多筆PDF,所以可能會例外+副檔名(ex.案號.案件性質.POA.PDF)
      If InStr(strFileName, ".") <> InStrRev(strFileName, ".") Then
          '案號後第一組 .. ; ex. 判斷data.1.pdf
          strExc(2) = InStr(strFileName, ".")
          strExc(1) = Mid(strFileName, Val(strExc(2)) + 1, InStr(Mid(strFileName, Val(strExc(2)) + 1), ".") - 1)
      Else
          strExc(1) = Mid(strFileName, InStrRev(strFileName, ".") + 1)
      End If
      '判斷是數值,長度是3碼或4碼
      'If Val(strExc(1)) > 100 Then
      If Val(strExc(1)) > 0 And _
         (Len(strExc(1)) >= 3 Or Len(strExc(1)) <= 4) Then
         strConSql = strConSql & " and cp10=" & CNULL(strExc(1))
      End If
   End If
   
   If Trim(Left(Combo1.Text, 2)) = "0" Then '發文
      '有案件性質匯入時移入最近一道相同案件性質，沒有性質直接歸入最新一道收文。
      strSql = "select cp09,cp10 From caseprogress" & _
               " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
               strConSql & " order by cp66 desc,cp67 desc,cp09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         strUpdCP09 = RsTemp.Fields("cp09")
         strUpdCP10 = RsTemp.Fields("cp10")
      Else
         strErrDesc = strErrDesc & "找不到歸卷的文號"
      End If
      
   'Add By Sindy 2020/7/27
   ElseIf Trim(Left(Combo1.Text, 2)) = "1" Then '收據
      'AB類,有發文規費
      '發文日期最大
      strSql = "select cp09,cp10 From caseprogress" & _
               " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
               " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B')" & _
               " and cp84>0" & _
               " and cp158>0 and cp159=0" & strConSql & _
               " order by cp27 desc,cp82 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         strUpdCP09 = RsTemp.Fields("cp09")
         strUpdCP10 = RsTemp.Fields("cp10")
         '檢查是否已有收據電子檔,若有,不可以歸檔
         strSql = "select cpp02" & _
                  " From casepaperpdf" & _
                  " where cpp01='" & strUpdCP09 & "'" & _
                  " and instr(upper(cpp02),'." & UCase(m_strReFileN) & ".')>0" & _
                  " and substr(upper(cpp02),-4)<>'.DEL'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strErrDesc = strErrDesc & "收據電子檔已存在(" & strUpdCP09 & IIf(strUpdCP10 <> "", "-" & strUpdCP10, "") & ")"
            strUpdCP09 = "": strUpdCP10 = ""
         End If
      Else
         strErrDesc = strErrDesc & "找不到歸卷的文號"
      End If
      '2020/7/27 END
   End If
   
   Exit Sub
End Sub

Private Sub GetErrText(strFName As String)
Dim i As Integer
Dim strText As String

   If intUpdStarRow > 0 Then
      For i = intUpdStarRow To intUpdEndRow
         If UCase(Trim(strFName)) <> UCase(Trim(Grid2.TextMatrix(i, 0))) Then
            strText = convForm(CheckStr(Grid2.TextMatrix(i, 0)), 30) & IIf(strUpdCP09 = "", "找不到歸卷的文號，", "")
            strText = Left(strText, Len(strText) - 1)
            List1.AddItem UCase(strText), 0: SetListScroll List1
         End If
      Next
   End If
End Sub

'存檔
Private Function SaveFilePDF() As Boolean
Dim dblFCnt As Double
Dim strFileName As String
Dim strFullFileName As String
Dim stReName As String
Dim fs, f
Dim strErr As String
Dim bolSave As Boolean
Dim bolCnn As Boolean
Dim strTcp01 As String, strTcp02 As String, strTcp03 As String, strTcp04 As String
Dim strFilePath As String, strFN As String 'Added by Lydia 2023/05/09

On Error GoTo ErrHand
   
   For dblFCnt = intUpdStarRow To intUpdEndRow
      strErr = ""
      strFileName = Grid2.TextMatrix(dblFCnt, 0)
      strFullFileName = txtPath1.Text & "\" & strFileName
      bolSave = False
      
      '更名
      If Trim(Left(Combo1.Text, 2)) = "0" Then '發文
         '只有案號.PDF或案號.案件性質.PDF才加副檔名DATA
         strExc(1) = ""
         If InStr(strUpdCP01 & strUpdCP02 & IIf(strUpdCP03 <> "0", strUpdCP03, "") & IIf(strUpdCP04 <> "00", strUpdCP04, "") & ".PDF," & _
                  strUpdCP01 & Val(strUpdCP02) & IIf(strUpdCP03 <> "0", strUpdCP03, "") & IIf(strUpdCP04 <> "00", strUpdCP04, "") & ".PDF," & _
                  strUpdCP01 & strUpdCP02 & IIf(strUpdCP03 <> "0", strUpdCP03, "") & IIf(strUpdCP04 <> "00", strUpdCP04, "") & "." & strUpdCP10 & ".PDF," & _
                  strUpdCP01 & Val(strUpdCP02) & IIf(strUpdCP03 <> "0", strUpdCP03, "") & IIf(strUpdCP04 <> "00", strUpdCP04, "") & "." & strUpdCP10 & ".PDF", UCase(strFileName)) > 0 Then
            strExc(1) = "DATA"
         End If
         If PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, strFileName, stReName, True, 1, False, strErr, , strExc(1)) = False Then
            GoTo ReadNext
         End If
         
      'Add By Sindy 2020/7/27
      ElseIf Trim(Left(Combo1.Text, 2)) = "1" Then '收據
         stReName = PUB_FCPCaseNo2FileName(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04) & _
                    "." & strUpdCP10 & "." & m_strReFileN & _
                    Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") + 1)
         '2020/7/27 END
      End If
      
      '檢查檔案是否已存在
      strSql = "select cpp02" & _
               " From casepaperpdf" & _
               " where cpp01='" & strUpdCP09 & "'" & _
               " and upper(cpp02)='" & UCase(stReName) & "'" & _
               " and substr(upper(cpp02),-4)<>'.DEL'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strErr = strFileName & " => " & stReName & "　　，檔案已存在"
         GoTo ReadNext
      End If
      '檢查此文號的案號是否與系統抓到的案號一致
      strSql = "select cp01,cp02,cp03,cp04" & _
               " From caseprogress" & _
               " where cp09='" & strUpdCP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strTcp01 = RsTemp.Fields("cp01")
         strTcp02 = RsTemp.Fields("cp02")
         strTcp03 = RsTemp.Fields("cp03")
         strTcp04 = RsTemp.Fields("cp04")
         If strTcp01 <> strUpdCP01 Or _
            strTcp02 <> strUpdCP02 Or _
            strTcp03 <> strUpdCP03 Or _
            strTcp04 <> strUpdCP04 Then
            strErr = convForm(CheckStr(strFileName), 30) & "文號" & strUpdCP09 & _
                     "本所案號" & strTcp01 & strTcp02 & strTcp03 & strTcp04 & _
                     "與系統抓到的案號" & strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 & _
                     "不一致!"
            GoTo ReadNext
         End If
      End If
      
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(strFullFileName)
      '檔案大小為 0 KB 有誤
      If f.Size = 0 Then
         strErr = convForm(CheckStr(strFileName), 30) & MsgText(9221)
         GoTo ReadNext
      End If
      
      cnnConnection.BeginTrans
      bolCnn = True
      If SaveAttFile_PDF(strUpdCP09, strFullFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & " => " & stReName & "　　存檔失敗！" & vbCrLf & Err.Description
         
         cnnConnection.RollbackTrans
         bolCnn = False
         
         GoTo ReadNext
      Else
         bolSave = True
         cnnConnection.CommitTrans
         bolCnn = False
         'Added by Lydia 2023/05/09 報告客戶之資料統一存檔FCT_WORKFLOW：增加送件後之收據及紙本申請書存檔FCT_WORKFLOW
         strFilePath = Pub_GetEFilePath_All(strTcp01, strTcp02, strTcp03, strTcp04)
         If UCase(strFileName) <> UCase(stReName) Then
            f.Name = stReName
         End If
         strFN = Pub_GetEFileName(strFilePath, stReName)
         If Len(strFN) <> Len(stReName) Then
            f.Name = strFN
         End If
         fs.CopyFile txtPath1.Text & "\" & strFN, strFilePath & "\" & strFN
         Sleep 1000
         strFullFileName = txtPath1.Text & "\" & strFN
         'end 2023/05/09
         fs.DeleteFile strFullFileName, True '刪檔
      End If
      
ReadNext:
      If bolSave = False Then
         strErr = Replace(strErr, vbCrLf, "")
         List1.AddItem UCase(strErr), 0: SetListScroll List1
      End If
   Next dblFCnt
   
   Exit Function
   
ErrHand:
   If bolCnn = True Then
      cnnConnection.RollbackTrans
   End If
   MsgBox Err.Description
End Function

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1.Text Then
      strExc(0) = ""
      txtPath1.Tag = ""
      Command2.Enabled = True
      txtPath1.Enabled = True 'Add By Sindy 2020/7/27
      Select Case Combo1.ListIndex
         Case 0 '發文
            'Command2.Enabled = False '保留,可變更
            strExc(0) = GetSetting("TAIE", "FCT", UCase("frm030202_01") & "Dir", "")  '在發文第一畫面設定
         'Modify By Sindy 2020/7/27
         Case 1 '收據
            'Modified by Lydia 2024/07/22 改成變數
            'strExc(0) = "\\SALE1\FCT_RECEIPT_SCAN"
            strExc(0) = "\\" & strSale1Path & "\FCT_RECEIPT_SCAN"
            Command2.Enabled = False
            txtPath1.Enabled = False
         '2020/7/27 END
         Case Else
            strExc(0) = GetSetting("TAIE", "FCT", UCase(Me.Name) & "Dir-" & Trim(Left(UCase(Combo1.Text), 2)), "")
      End Select
      If strExc(0) = "" Then strExc(0) = PUB_Getdesktop  '預設個人桌面
      
      If strExc(0) <> "" Then txtPath1.Text = strExc(0)
      
      txtPath1.Tag = txtPath1.Text
   End If
   Combo1.Tag = Combo1.Text
   
   'Add By Sindy 2020/7/27
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
      Command2.Enabled = True
      txtPath1.Enabled = True
   End If
   '2020/7/27 END
End Sub

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.pdf"
      .Filter = "PDF檔案 (*.pdf)|*.pdf"
      .InitDir = txtPath1.Tag '抓預設資料夾
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            If Trim(Left(Combo1, 2)) = "0" Then '發文
                SaveSetting "TAIE", "FCT", UCase("frm030202_01") & "Dir", sFile(0)
            Else
                SaveSetting "TAIE", "FCT", UCase(Me.Name) & "Dir-" & Trim(Left(UCase(Combo1.Text), 2)), sFile(0)
            End If
            txtPath1.Text = sFile(0)
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
                If Trim(Left(Combo1, 2)) = "0" Then '發文
                    SaveSetting "TAIE", "FCT", UCase("frm030202_01") & "Dir", Left(.FileName, InStrRev(.FileName, "\") - 1)
                Else
                    SaveSetting "TAIE", "FCT", UCase(Me.Name) & "Dir-" & Trim(Left(UCase(Combo1.Text), 2)), Left(.FileName, InStrRev(.FileName, "\") - 1)
                End If
            End If
            txtPath1.Text = Left(.FileName, InStrRev(.FileName, "\") - 1)
         End If
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrintL As Integer
Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   
   m_DefaultPrinter = Printer.DeviceName
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      j = j + 1
      If Printer.DeviceName = m_DefaultPrinter Then
         SeekPrint = i
      End If
   Next i
   Set Printer = Printers(SeekPrint)

   '預設下拉選單
   Combo1.Clear
   Combo1.AddItem "0 發文", 0
   Combo1.AddItem "1 收據", 1
   Combo1.ListIndex = 0
   Call Combo1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030208 = Nothing
End Sub

Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub

Private Function TxtValidate() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim Cancel As Boolean

TxtValidate = False

If IsEmptyText(txtPath1) = True Then
   strTit = "檢核資料"
   strMsg = "請輸入電子檔存放路徑！"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtPath1.SetFocus
   txtPath1_GotFocus
   Exit Function
Else
   'Modify By Sindy 2020/7/30
'   If Dir(txtPath1, vbDirectory) = "" Then
'        strTit = "檢核資料"
'        strMsg = "電子檔存放路徑不存在！"
'        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'        txtPath1.SetFocus
'        txtPath1_GotFocus
'        Exit Function
'   End If
   If PUB_ChkDir(txtPath1) = False Then
      MsgBox "檔案存放路徑不存在，請重新選擇！"
      If txtPath1.Enabled = True Then txtPath1.SetFocus
      Exit Function
   End If
End If

TxtValidate = True
End Function

'列印
Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer
   
   iLine1 = 0
   For j = List1.ListCount - 1 To 0 Step -1
      For i = 1 To 1
         strTemp(i) = ""
      Next i
      strTemp(1) = List1.List(j)
      If iLine1 > 52 Or iLine1 = 0 Then
         If iLine1 > 0 Then Printer.NewPage: iLine1 = 0
         PrintTitle '列印表頭
      End If
      PrintDetail '列印明細
   Next j
   Printer.EndDoc
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 5500
End Sub

Sub PrintTitle()
GetPleft
iLine1 = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("匯入電子檔錯誤訊息清單") / 2)
Printer.CurrentY = iLine1 * 300
Printer.Print "匯入電子檔錯誤訊息清單"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine1 = iLine1 + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine1 = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print "錯誤訊息"

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print String(148, "-")
iLine1 = iLine1 + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
   For m_j = 1 To 1
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine1 * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine1 = iLine1 + 1
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub
