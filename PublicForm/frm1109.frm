VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1109 
   BorderStyle     =   1  '單線固定
   Caption         =   "部門別電子送件清單列印"
   ClientHeight    =   2916
   ClientLeft      =   456
   ClientTop       =   996
   ClientWidth     =   4776
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2916
   ScaleWidth      =   4776
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   285
      Left            =   1215
      TabIndex        =   4
      Top             =   2280
      Width           =   1005
      _ExtentX        =   1778
      _ExtentY        =   508
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   510
      Width           =   3405
   End
   Begin VB.Timer Timer1 
      Left            =   4140
      Top             =   960
   End
   Begin VB.ComboBox cboListType 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frm1109.frx":0000
      Left            =   1215
      List            =   "frm1109.frx":000A
      TabIndex        =   2
      Top             =   1530
      Width           =   2625
   End
   Begin VB.ComboBox cboListTime 
      Height          =   300
      ItemData        =   "frm1109.frx":001A
      Left            =   1215
      List            =   "frm1109.frx":001C
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1890
      Width           =   1005
   End
   Begin VB.TextBox txtPrinter 
      Enabled         =   0   'False
      Height          =   264
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1200
      Width           =   1005
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   3735
      TabIndex        =   6
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2745
      TabIndex        =   5
      Top             =   30
      Width           =   912
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   870
      Width           =   1005
      _ExtentX        =   1778
      _ExtentY        =   487
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "內外商送件只分時段不管分段時間，一律列出所有未繳費案件。如上午清單列印後還有送件且要一起繳費，則時段仍選上午，不可選下午。"
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   2370
      TabIndex        =   17
      Top             =   1950
      Width           =   2355
   End
   Begin VB.Label lblServerTime 
      Height          =   180
      Left            =   1215
      TabIndex        =   16
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "系統時間："
      Height          =   180
      Index           =   7
      Left            =   210
      TabIndex        =   15
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "分段時間："
      Height          =   180
      Index           =   6
      Left            =   210
      TabIndex        =   14
      Top             =   2325
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Index           =   5
      Left            =   390
      TabIndex        =   13
      Top             =   570
      Width           =   720
   End
   Begin MSForms.Label lblPrinter 
      Height          =   255
      Left            =   2265
      TabIndex        =   12
      Top             =   1215
      Width           =   1245
      VariousPropertyBits=   27
      Size            =   "2196;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "部門別："
      Height          =   180
      Index           =   2
      Left            =   390
      TabIndex        =   11
      Top             =   1590
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "送件時段："
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   10
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "列印人員："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   9
      Top             =   1242
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "清單日期："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   915
      Width           =   900
   End
End
Attribute VB_Name = "frm1109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (lblPrinter,Printer列印未改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create by Morgan 2011/6/2 參考 frm1108
Option Explicit

Dim m_OriPrinterName As String '原印表機名稱
Dim m_Dept As String '部門別
Dim m_intCnt As Integer '查詢出幾筆資料
Dim strPrinter As String
Dim prnPrint As Printer

Private Sub cboListTime_Click()
   GetTime
End Sub

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
   Set frm1108 = Nothing
End Sub

'檢查當日清單是否有無支票號的
Private Function isChecked(Optional p_bAll As Boolean = True) As Boolean
   
   isChecked = True
   
   strExc(0) = "SELECT AL06 FROM APPLISTe" & _
      " WHERE AL01=" & DBDATE(MaskEdBox2) & " AND AL02='" & m_Dept & "'"
   
   If p_bAll = False Then
      strExc(0) = strExc(0) & " and AL03='" & cboListTime.ListIndex & "'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         If IsNull(.Fields(0)) Then
            isChecked = False
            Exit Function
         End If
         .MoveNext
      Loop
      End With
   Else
      isChecked = False
   End If
End Function

Private Function Process() As Boolean
   Dim arrList
   Dim strConCP As String, strInsSQL As String, iEffect As Integer
   Dim strDate As String, strTime As String
  
   strDate = DBDATE(MaskEdBox2)
   strTime = Replace(MaskEdBox1, ":", "")
   
   strConCP = " and cp27=" & strDate
   '上午
   If cboListTime.ListIndex = 0 Then
      strConCP = strConCP & " and cp82<" & strTime
   '下午
   Else
      strConCP = strConCP & " and cp82>=" & strTime
   End If
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   '刪除舊資料
   'Modified by Morgan 2011/11/18 因為財務處會忘記傳送,所以改只要刪除當天的資料,倘若發生前日清單已產生但不要繳費時才由人工通知刪除資料(目前尚未發生)
   '內商
   'Modified by Morgan 2014/8/8
   'If m_Dept = "P2" Then
   If m_Dept = "P2" Or (m_Dept = "F1" And strDate > "20140808") Then
      'strSql = " DELETE FROM APPLISTDETAILe" & _
         " WHERE (ALD01,ALD02,ALD03) in (SELECT AL01,AL02,AL03 FROM APPLISTe WHERE AL02='" & m_Dept & "' AND AL06 IS NULL AND ((AL01=" & strDate & " AND AL03='" & cboListTime.ListIndex & "') or AL01<" & strDate & "))"
      strSql = " DELETE FROM APPLISTDETAILe" & _
         " WHERE (ALD01,ALD02,ALD03) in (SELECT AL01,AL02,AL03 FROM APPLISTe WHERE AL02='" & m_Dept & "' AND AL06 IS NULL AND AL01=" & strDate & " AND AL03='" & cboListTime.ListIndex & "')"
         
   '其他
   Else
      'strSql = " DELETE FROM APPLISTeDETAIL" & _
         " WHERE (ALD01,ALD02,ALD03) in (SELECT AL01,AL02,AL03 FROM APPLISTe WHERE AL02='" & m_Dept & "' AND AL06 IS NULL AND ((AL01=" & strDate & " AND AL03='" & cboListTime.ListIndex & "') or AL01<" & strDate & "))"
      strSql = " DELETE FROM APPLISTeDETAIL" & _
         " WHERE (ALD01,ALD02,ALD03) in (SELECT AL01,AL02,AL03 FROM APPLISTe WHERE AL02='" & m_Dept & "' AND AL06 IS NULL AND AL01=" & strDate & " AND AL03='" & cboListTime.ListIndex & "')"
   End If
   cnnConnection.Execute strSql, iEffect
   
   'strSql = " DELETE FROM APPLISTe" & _
      " WHERE AL02='" & m_Dept & "' AND AL06 IS NULL AND ((AL01=" & strDate & " AND AL03='" & cboListTime.ListIndex & "') or AL01<" & strDate & ")"
      
   strSql = " DELETE FROM APPLISTe" & _
      " WHERE AL02='" & m_Dept & "' AND AL06 IS NULL AND AL01=" & strDate & " AND AL03='" & cboListTime.ListIndex & "'"
   cnnConnection.Execute strSql, iEffect
   
   '新增清單明細
   strSql = GetSql
   cnnConnection.Execute strSql, iEffect
   
   '新增清單主檔
   '內商
   'Modified by Morgan 2014/8/8 +外商
   'If m_Dept = "P2" Then
   If m_Dept = "P2" Or m_Dept = "F1" Then
      strSql = " INSERT INTO APPLISTe(AL01,AL02,AL03,AL04,AL05,AL07,AL08,AL09)" & _
         " SELECT ALD01,ALD02,ALD03,ALD04," & strTime & ",'" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MI')" & _
         " FROM APPLISTDETAILe,APPLISTe WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
         " AND AL01(+)=ALD01 AND AL02(+)=ALD02 AND AL03(+)=ALD03 AND AL06 IS NULL GROUP BY ALD01,ALD02,ALD03,ALD04"
      cnnConnection.Execute strSql, iEffect
   '其他
   Else
      strSql = " INSERT INTO APPLISTe(AL01,AL02,AL03,AL04,AL05,AL07,AL08,AL09)" & _
         " SELECT ALD01,ALD02,ALD03,ALD04," & strTime & ",'" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MI')" & _
         " FROM APPLISTeDETAIL,APPLISTe WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
         " AND AL01(+)=ALD01 AND AL02(+)=ALD02 AND AL03(+)=ALD03 AND AL04(+)=ALD04 AND AL06 IS NULL GROUP BY ALD01,ALD02,ALD03,ALD04"
      cnnConnection.Execute strSql, iEffect
   End If
   
   '若無明細時新增一筆
   strExc(0) = "select * from APPLISTe where al01=" & strDate & " and al02='" & m_Dept & "' and al03='" & cboListTime.ListIndex & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      strSql = " INSERT INTO APPLISTe(AL01,AL02,AL03,AL04,AL05,AL07,AL08,AL09)" & _
         " VALUES(" & strDate & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','8'," & strTime & ",'" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MI'))"
      cnnConnection.Execute strSql, iEffect
   End If
   
   cnnConnection.CommitTrans
   Process = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function


Private Sub cmdPrint_Click()
   Dim strDesc As String
   
   If Trim(txtPrinter) = "" Then
      MsgBox "請輸入列印人員!!"
      txtPrinter.SetFocus
      Exit Sub
   End If
   
   If MaskEdBox2 = "___/__/__" Then
      MsgBox "請輸入發文日期!!"
      If MaskEdBox2.Enabled Then MaskEdBox2.SetFocus
      Exit Sub
   End If
   
   If MaskEdBox1 = "__:__:__" Then
      MsgBox "請輸入分段時間!!"
      If MaskEdBox1.Enabled Then MaskEdBox1.SetFocus
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   If cboListType.ListIndex >= 0 Then
      '該時段有未開支票就要建立清單資料
      If isChecked(False) = False Then
         If Process() = False Then
            MsgBox "無法產生清單資料！", vbCritical
            Exit Sub
         End If
      Else
         If MsgBox("財務處支票皆已開立，將不會重新產生清單明細！是否確定要列印？" & vbCrLf & vbCrLf & "( ※若案件有更動，請與財務處確認　貴部門該時段支票已作廢後再行列印！ )", vbOKCancel + vbQuestion + vbDefaultButton2, "重印確認") = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      
      ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
      
      pub_QL05 = pub_QL05 & ";" & Label1(0) & MaskEdBox2
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txtPrinter & lblPrinter
      pub_QL05 = pub_QL05 & ";" & Label1(2) & cboListType.Text
      pub_QL05 = pub_QL05 & ";" & Label1(3) & cboListTime.Text
      pub_QL05 = pub_QL05 & ";" & Label1(6) & MaskEdBox1
      
      m_intCnt = 0

      PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
      
      If DoPrint = True Then
         strDesc = "電子送件清單列印完畢!!"
      Else
         strDesc = "無電子送件案件可列印!!"
      End If
      
      PUB_RestorePrinter strPrinter 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
      
      InsertQueryLog (m_intCnt)
      MsgBox strDesc, vbInformation
   Else
      MsgBox "該部門無清單可列印！", vbInformation
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintHead(Optional pRptType As String = "8")

      Dim strTitle As String
      Dim lngY As Long
      
      strTitle = cboListType.Text & " " & MaskEdBox2 & " " & cboListTime & "電子送件清單" & IIf(pRptType <> "8", "(非智慧局)", "")
            
      Printer.Print
      Printer.FontSize = 16
      Printer.CurrentX = 5000
      Printer.Print strTitle
      Printer.FontSize = 12
      Printer.CurrentX = 13000
      Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
      Printer.CurrentX = 0
      lngY = Printer.CurrentY
      Printer.Print "區分時段時間：" & MaskEdBox1
      Printer.CurrentY = lngY
      Printer.CurrentX = 13000
      Printer.Print "列印時間：" & lblServerTime
      Printer.CurrentX = 13000
      Printer.Print "列印人員：" & lblPrinter
      
      Printer.Print
      If m_Dept = "P2" Or m_Dept = "F1" Then
         'Modified by Morgan 2025/3/24 +本所案號長度+4,案件名稱長度-4
         Printer.Print "本所案號            規費      規費小計 申請案號   案件性質　　 申請人               案件名稱                             商品類別      "
         Printer.Print "------------------- --------- -------- ---------- ------------ -------------------- ------------------------------------ --------------"
      Else
         Printer.Print "本所案號        規費      規費小計 申請案號   案件性質　　 申請人               案件名稱                                "
         Printer.Print "--------------- --------- -------- ---------- ------------ -------------------- ----------------------------------------"
      End If
End Sub

Private Sub PrintTail(iPage As Integer, Optional p_lngTot As Long, Optional p_iRecs As Integer, Optional p_iCaseCnt As Integer)
   Dim stData As String
   
   If p_lngTot > 0 Then
      '內商不印案件數
      If m_Dept = "P2" Then
         stData = "總計" & Space(4) & Right(Space(8) & "規費 " & Format(p_lngTot, "###,###,###") & " 元", 16) & Space(4) & Right(Space(2) & "明細 " & Format(p_iRecs) & " 筆", 9) & Space(4)
      Else
         stData = "總計" & Space(4) & Right(Space(8) & "規費 " & Format(p_lngTot, "###,###,###") & " 元", 16) & Space(4) & Right(Space(2) & "明細 " & Format(p_iRecs) & " 筆", 9) & Space(4) & Right(Space(2) & "案號 " & Format(p_iCaseCnt) & " 筆", 9)
      End If
   End If
               
   Printer.FontSize = 12
   If m_Dept = "P2" Or m_Dept = "F1" Then
      Printer.Print String(135, "-")
   Else
      Printer.Print String(120, "-")
   End If
   Printer.Print stData
   Printer.CurrentY = 10700
   Printer.CurrentX = 7000
   Printer.Print "第 " & Format(iPage) & " 頁"
   
End Sub

Private Sub PrintTail1(iPage As Integer, Optional p_iRecs As Integer)
   Dim stData As String
   
   Printer.FontSize = 12
   Printer.Print String(135, "-")
   If p_iRecs > 0 Then
      Printer.Print "共 " & p_iRecs & " 筆"
   End If
   
   Printer.CurrentY = 10700
   Printer.CurrentX = 7000
   Printer.Print "第 " & Format(iPage) & " 頁"
   
End Sub

Private Function GetSql() As String

   Dim strDate As String, strTime As String, strDate1 As String
   Dim strInsSQL As String, strCon As String, stVTable As String
   
  
   strDate = DBDATE(MaskEdBox2)
   strTime = Replace(MaskEdBox1, ":", "")
   
   strCon = " and A.cp22 is null and A.cp27=" & strDate
   '上午
   If cboListTime.ListIndex = 0 Then
      strCon = strCon & " and A.cp82<" & strTime
   '下午
   Else
      strCon = strCon & " and A.cp82>=" & strTime
   End If
   
   Select Case m_Dept
      '專利
      Case "P1", "F2"
         'Modified by Morgan 2019/6/18 +非智慧局清單(7)(Ex:P-087216)
         'stVTable = stVTable & " SELECT CP01, CP02, CP03, CP04,'8' ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         stVTable = stVTable & " SELECT CP01, CP02, CP03, CP04,decode(cp130,null,'8','經濟部智慧財產局','8','7') ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         'end 2019/6/18
         stVTable = stVTable & " FROM CASEPROGRESS A,PATENT,STAFF"
         stVTable = stVTable & " where cp01='" & IIf(m_Dept = "P1", "P", "FCP") & "'" & strCon
         stVTable = stVTable & " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"
         stVTable = stVTable & " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'"
         stVTable = stVTable & " and cp118='Y'"
         'Modified by Morgan 2019/6/18
         'stVTable = stVTable & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         stVTable = stVTable & " GROUP BY decode(cp130,null,'8','經濟部智慧財產局','8','7'),CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         'end 2019/6/18
         
         strInsSQL = "INSERT INTO APPLISTeDETAIL(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10)" & _
            " SELECT " & strDate & ",'" & m_Dept & "','" & cboListTime.ListIndex & "',ALD04, CP01,CP02,CP03,CP04,ALD09,ALD10" & _
            " FROM  (" & stVTable & ") X"
   
      '外商
      Case "F1"
         'Modified by Morgan 2014/8/8 一天可能發文3次
         'strInsSQL = "INSERT INTO APPLISTeDETAIL(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10)" & _
            " SELECT " & strDate & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','8', CP01,CP02,CP03,CP04,SUM(CP84),MAX(CP22)" & _
            " FROM CASEPROGRESS A, trademark, staff" & _
            " where cp01='FCT' and cp84>0" & strCon & _
            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 and TM10='000'" & _
            " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "' and cp118='Y'"
         'strInsSQL = strInsSQL & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         'Modified by Morgan 2015/7/13 改5天(颱風假,部分放假)
         strDate1 = PUB_GetWorkDay1(CompDate(2, -5, strDate), True)
         
         strCon = " and cp27>=" & strDate1
         'Modified by Morgan 2016/1/5 +類別:7=非智慧局 8=智慧局
         'Modified by Morgan 2016/7/26 非智慧局的電子送件不需要印(目前外商沒有此類案件)--阿蓮
         'Modified by Morgan 2020/5/27 部門改抓發文人員判斷 Ex:FCT-045833發文時未分案
         'Modified by Morgan 2022/7/8 +CFT的英文證明(拿掉TM10='000'條件) --阿蓮
         strInsSQL = "INSERT INTO APPLISTDETAILe(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10,ALD11)" & _
            " SELECT " & strDate & ",'" & m_Dept & "','" & cboListTime.ListIndex & "',decode(cp130,null,'8','經濟部智慧財產局','8','7'), CP01,CP02,CP03,CP04,CP84,CP22,CP09" & _
            " FROM CASEPROGRESS, trademark, staff" & _
            " where (cp01='FCT' or (cp01='CFT' and cp10='304' and cp30 is not null)) and cp84>0 and nvl(cp130,'經濟部智慧財產局')='經濟部智慧財產局'" & strCon & _
            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04" & _
            " and st01(+)=cp83 and substrb(st03,1,2)='F1' and cp118='Y'" & _
            " and not exists(select * from applistdetaile where ald01>=" & strDate1 & " and ald11=cp09)"
          'end 2014/8/8
      '內商
      Case "P2"
         '前一工作天
         'Modified by Morgan 2012/8/3 改3天
         'Modified by Morgan 2015/7/13 改5天(颱風假,部分放假)
         strDate1 = PUB_GetWorkDay1(CompDate(2, -5, strDate), True)
         
         'Modified by Morgan 2015/6/30 排除不出名
         'Modified by Morgan 2016/1/5 +類別:7=非智慧局 8=智慧局
         strCon = " and cp22 is null and cp85<=" & strDate & " and cp85>=" & strDate1
         'Modified by Morgan 2016/7/26 非智慧局的電子送件不需要印(繳費單給財務處去便利商店繳)--桂英
         'Modified by Morgan 2020/5/27 部門改抓發文人員判斷 Ex:FCT-045833發文時未分案
         'Modified by Morgan 2020/5/28 改回抓承辦人,因為內商是用承辦人發文日判斷
         strInsSQL = "INSERT INTO APPLISTDETAILe(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10,ALD11)" & _
            " SELECT " & strDate & ",'" & m_Dept & "','" & cboListTime.ListIndex & "',decode(cp130,null,'8','經濟部智慧財產局','8','7'), CP01,CP02,CP03,CP04,CP84,CP22,CP09" & _
            " FROM CASEPROGRESS, trademark, staff" & _
            " where cp01 IN ('T','FCT') and cp84>0 and nvl(cp130,'經濟部智慧財產局')='經濟部智慧財產局'" & strCon & _
            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 and TM10='000'" & _
            " and st01(+)=cp14 and substrb(st03,1,2)='P2' and cp118='Y'" & _
            " and not exists(select * from applistdetaile where ald01<=" & strDate & " and ald01>=" & strDate1 & " and ald11=cp09)"
   End Select
   
   GetSql = strInsSQL
      
End Function
'Modified by Morgan 2012/8/6 規費扣除已銷帳
Private Function DoPrint() As Boolean

   Dim strTmp As String, iRec As Integer, strTitle As String, iRecs As Integer, i As Integer
   Dim lngTot As Long, strCon As String, stVTable As String, stLastNo As String
   Dim iCaseNo As Integer '案件筆數
   Dim iPage As Integer '頁次
   Dim nCopys As Integer '份數
   Dim iCopys As Integer
   Dim strDate As String, strTime As String
   Dim stVTable2 As String, stVTable3 As String, stVTable4 As String
   Dim stRptType As String 'Added by Morgan 2016/1/5
   
   strDate = DBDATE(MaskEdBox2)
   strTime = Replace(MaskEdBox1, ":", "")
   
   strCon = " and A.cp27=" & strDate
   '上午
   If cboListTime.ListIndex = 0 Then
      strCon = strCon & " and A.cp82<" & strTime
   '下午
   Else
      strCon = strCon & " and A.cp82>=" & strTime
   End If
   
   strSql = ""
   '上下午,出名否不同清單
   
   '專利
   '排序:是否出名(C09),清單種類(C10),主管機關(C12),發文規費加總(C03),本所案號(C01),發文規費(C02),案件性質(C05)
   If m_Dept = "P1" Or m_Dept = "F2" Then
         nCopys = 2
         'Modified by Morgan 2019/6/18 +非智慧局清單(Ex:P-087216)
         stVTable = " SELECT ALD05 X01, ALD06 X02, ALD07 X03, ALD08 X04, PA05, PA11, PA26, ALD09 X05, ALD04 X06, ALD10,ALD01,ALD02,ALD03,ALD04 " & _
            " FROM APPLISTeDETAIL A,PATENT" & _
            " WHERE ALD01=" & strDate & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04 in ('7','8')" & _
            " AND PA01(+)=ALD05 AND PA02(+)=ALD06 AND PA03(+)=ALD07 AND PA04(+)=ALD08"
         
         stVTable2 = "select cp09 T2C1,sum(a1u09) T2C2" & _
            " from APPLISTeDETAIL,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & strDate & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04 in ('7','8')" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0" & strCon & _
            " AND A1U03(+)=CP09 GROUP BY CP09"
         
         stVTable3 = "select CP43 T3C1,sum(a1u09) T3C2" & _
            " from APPLISTeDETAIL,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & strDate & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04 in ('7','8')" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0" & strCon & _
            " AND CP10='404' AND CP09>'B' AND CP43 IS NOT NULL AND A1U03(+)=CP43 GROUP BY CP43"


         strSql = "SELECT LPAD(A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,15,' ') C01, A.CP84 C02, X05 C03" & _
            ",LPAD(NVL(PA11,' '),10,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(NVL(CU04,' '),20,' ') C06" & _
            ",RPAD(PA05,40,' ') C07, RPAD(' ',10,' ') C08, DECODE(ALD10,'N',1,0) C09, X06 C10" & _
            ",DECODE( NVL(DECODE(A.CP10,'404',DECODE(SUBSTR(A.CP09,1,1),'A',A.CP17-NVL(T2C2,0),B.CP17-NVL(T3C2,0)),A.CP17-NVL(T2C2,0)),0)-A.CP84,0,' ','*') C11" & _
            ",decode(A.cp130,null,'經濟部智慧財產局',decode(instr(A.cp130,','),0,A.cp130,substr(A.cp130,1,instr(A.cp130,',')-1))) C12,ALD04 C13" & _
            " FROM  (" & stVTable & ") X, CASEPROGRESS A, CUSTOMER" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,(" & stVTable2 & ") T2,(" & stVTable3 & ") T3" & _
            " where A.CP01=X01 AND A.CP02=X02 AND A.CP03=X03 AND A.CP04=X04" & " and A.cp84>0" & strCon & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(A.CP10,'404',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'404',A.CP43,NULL)" & _
            " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'404',NVL(NP07,B.CP10),A.CP10)" & _
            " AND T2C1(+)=A.CP09 AND T3C1(+)=B.CP09"
            
         strSql = strSql & " order by 1,2 desc"
         
   '外商
   'Modified by Morgan 2014/8/8 外商 2014/8/8 以後改用內商模式
   'ElseIf m_Dept = "F1" Then
   ElseIf m_Dept = "F1" And strDate <= "20140808" Then
   'end 2014/8/8
         nCopys = 2
         
         stVTable = " SELECT ALD05 X01, ALD06 X02, ALD07 X03, ALD08 X04, TM05, TM09, TM12, TM15, TM23, ALD09 X05, ALD04 X06,ALD10 " & _
            " FROM APPLISTeDETAIL,TRADEMARK A" & _
            " WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
            " AND A.TM01(+)=ALD05 AND A.TM02(+)=ALD06 AND A.TM03(+)=ALD07 AND A.TM04(+)=ALD08"
         
         stVTable2 = "select cp09 T2C1,sum(a1u09) T2C2" & _
            " from APPLISTeDETAIL,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & strDate & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04='8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0" & strCon & _
            " AND A1U03(+)=CP09 GROUP BY CP09"
         
         stVTable3 = "select CP43 T3C1,sum(a1u09) T3C2" & _
            " from APPLISTeDETAIL,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & strDate & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04='8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0" & strCon & _
            " AND CP10='303' AND CP09>'B' AND CP43 IS NOT NULL AND A1U03(+)=CP43 GROUP BY CP43"
         
         '若性質為延期303時:A類收文用CP43抓NP07，規費用延期的CP17比較；B類收文用CP43抓相關收文號的CP10，規費用相關收文號的CP17比較。
         strSql = "SELECT LPAD(A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,15,' ') C01, A.CP84 C02, X05 C03" & _
            ", LPAD(NVL(DECODE(A.CP01,'CFT',A.CP30,NVL(TM15,TM12)),' '),10,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(CU04,20,' ') C06" & _
            ", RPAD(TM05,40,' ') C07, RPAD(TM09,10,' ') C08, DECODE(ALD10,'N',1,0) C09,X06 C10" & _
            ", DECODE( NVL(DECODE(A.CP10,'303',DECODE(SUBSTR(A.CP09,1,1),'A',A.CP17-NVL(T2C2,0),B.CP17-NVL(T3C2,0)),A.CP17-NVL(T2C2,0)),0)-A.CP84,0,' ','*') C11" & _
            ",decode(A.cp130,null,'經濟部智慧財產局',decode(instr(A.cp130,','),0,A.cp130,substr(A.cp130,1,instr(A.cp130,',')-1))) C12,'8' C13" & _
            " FROM  (" & stVTable & ") X, CASEPROGRESS A,  CUSTOMER" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,(" & stVTable2 & ") T2,(" & stVTable3 & ") T3" & _
            " where A.CP01=X01 AND A.CP02=X02 AND A.CP03=X03 AND A.CP04=X04" & " and A.cp84>0" & strCon & _
            " and cu01(+)=substr(TM23,1,8) and cu02(+)=substr(TM23,9,1)" & _
            " AND B.CP09(+)=DECODE(A.CP10,'303',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'303',A.CP43,NULL)" & _
            " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'303',NVL(NP07,B.CP10),A.CP10)" & _
            " AND T2C1(+)=A.CP09 AND T3C1(+)=B.CP09"

         strSql = strSql & " order by 4,1,2 desc"
         
   '內商
   'Modified by Morgan 2014/8/8 外商 2014/8/8 以後改用內商模式
   'ElseIf m_Dept = "P2" Then
   ElseIf m_Dept = "P2" Or m_Dept = "F1" Then
         nCopys = 2
         
         strCon = " AND ALD01=" & strDate & " AND ALD02='" & m_Dept & "'" & _
            " AND ALD03='" & cboListTime.ListIndex & "'"
         
         'Modified by Morgan 2016/1/5 +非智慧局清單
         stVTable = " SELECT ALD05 X01, ALD06 X02, ALD07 X03, ALD08 X04, SUM(ALD09) X05,ALD04 X06" & _
            " FROM APPLISTDETAILe" & _
            " WHERE 1=1" & strCon & _
            " GROUP BY ALD04,ALD05,ALD06,ALD07,ALD08"
         
         stVTable2 = "select cp09 T2C1,sum(a1u09) T2C2" & _
            " from APPLISTDETAILe,caseprogress A,acc1u0" & _
            " WHERE  CP09(+)=ALD11" & strCon & _
            " AND A1U03(+)=CP09 GROUP BY CP09"
         
         stVTable3 = "select cp43 T3C1,sum(a1u09) T3C2" & _
            " from APPLISTDETAILe,caseprogress A,acc1u0" & _
            " WHERE CP09(+)=ALD11" & strCon & _
            " AND CP10='303' AND CP09>'B' AND CP43 IS NOT NULL" & _
            " AND A1U03(+)=CP43 GROUP BY CP43"
            
         'Added by Morgan 2025/3/24 CFT 的智權
         stVTable4 = "select cp09 T4C1,'智權' T4C2" & _
            " from APPLISTDETAILe,caseprogress A,acc0j0,acc0k0" & _
            " WHERE  CP09(+)=ALD11" & strCon & _
            " AND ALD05='CFT'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0" & _
            " and a0j01(+)=cp09 and a0k01(+)=a0j13 and a0k11='J'"
                     

         '若性質為延期303時:A類收文用CP43抓NP07，規費用延期的CP17比較；B類收文用CP43抓相關收文號的CP10，規費用相關收文號的CP17比較。
         strSql = "SELECT LPAD(T4C2||A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,19,' ') C01, A.CP84 C02, X05 C03" & _
            ", LPAD(NVL(DECODE(A.CP01,'CFT',A.CP30,NVL(TM15,TM12)),' '),10,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(CU04,20,' ') C06" & _
            ", RPAD(TM05,36,' ') C07, RPAD(TM09,10,' ') C08, DECODE(ALD10,'N',1,0) C09,X06 C10" & _
            ", DECODE( NVL(DECODE(A.CP10,'303',DECODE(SUBSTR(A.CP09,1,1),'A',A.CP17-NVL(T2C2,0),B.CP17-NVL(T3C2,0)),A.CP17-NVL(T2C2,0)),0)-A.CP84,0,' ','*') C11" & _
            ",decode(A.cp130,null,'經濟部智慧財產局',decode(instr(A.cp130,','),0,A.cp130,substr(A.cp130,1,instr(A.cp130,',')-1))) C12, X06 C13" & _
            " FROM  APPLISTDETAILe,(" & stVTable & ") X, TRADEMARK, CASEPROGRESS A, CUSTOMER" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,(" & stVTable2 & ") T2,(" & stVTable3 & ") T3,(" & stVTable4 & ") T4" & _
            " WHERE X01(+)=ALD05 AND X02(+)=ALD06" & strCon & _
            " AND X03(+)=ALD07 AND X04(+)=ALD08" & _
            " AND TM01(+)=ALD05 AND TM02(+)=ALD06 AND TM03(+)=ALD07 AND TM04(+)=ALD08" & _
            " AND A.CP09(+)=ALD11 and cu01(+)=substr(TM23,1,8) and cu02(+)=substr(TM23,9,1)" & _
            " AND B.CP09(+)=DECODE(A.CP10,'303',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'303',A.CP43,NULL)" & _
            " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'303',NVL(NP07,B.CP10),A.CP10)" & _
            " AND T2C1(+)=A.CP09 AND T3C1(+)=B.CP09 AND T4C1(+)=A.CP09"
            
         strSql = strSql & " order by 13 desc,4,1,2 desc"
   End If

On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         m_intCnt = m_intCnt + .RecordCount
         Printer.Orientation = 2
         Printer.Font = "細明體"
         
         For iCopys = 1 To nCopys
            .MoveFirst
            stRptType = "" & .Fields("C13") 'Added by Morgan 2016/1/5
            If iCopys > 1 Then Printer.NewPage
         
            iPage = 1: iRec = 0: lngTot = 0: iCaseNo = 0: iRecs = 0: stLastNo = ""
            
            PrintHead stRptType
            
            Do While Not .EOF
               'Added by Morgan 2016/1/5
               If stRptType <> .Fields("C13") Then
                  PrintTail iPage, lngTot, iRecs, iCaseNo
                  Printer.NewPage
                  iPage = 1: iRec = 0: lngTot = 0: iCaseNo = 0: iRecs = 0: stLastNo = ""
                  
                  stRptType = .Fields("C13")
                  PrintHead stRptType
               End If
                  
               iRec = iRec + 1: iRecs = iRecs + 1
               If iRec > 26 Then
                  PrintTail iPage
                  Printer.NewPage
                  iPage = iPage + 1
                  PrintHead stRptType
                  iRec = 0
               End If
               strTmp = ""
               
               For i = 0 To 7
                  '規費
                  If i = 1 Then
                     strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(i)), "###,###"), 8) & .Fields("C11") & Space(1)
                     lngTot = lngTot + Val("" & .Fields(i))
                     
                  '規費小計
                  ElseIf i = 2 Then
                     '內商不印小計
                     'If .Fields(0) <> stLastNo Then
                     If m_Dept <> "P2" And .Fields(0) <> stLastNo Then
                        strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(i)), "###,###"), 8) & Space(1)
                     Else
                        strTmp = strTmp & Space(9)
                     End If
                  '案件性質
                  ElseIf i = 4 Then
                     strTmp = strTmp & .Fields(i) & Space(1)
                  Else
                     '內商同案號也要印
                     If m_Dept = "P2" Or .Fields(0) <> stLastNo Then
                        strTmp = strTmp & .Fields(i) & Space(1)
                     Else
                        strTmp = strTmp & Space(Len(.Fields(i)) + 1)
                     End If
                  End If
               Next
               Printer.CurrentY = Printer.CurrentY + 60
               Printer.Print strTmp
               If .Fields(0) <> stLastNo Then iCaseNo = iCaseNo + 1
               stLastNo = .Fields(0)
               .MoveNext
            Loop
            PrintTail iPage, lngTot, iRecs, iCaseNo
         Next
         
         Printer.EndDoc
         DoPrint = True
      End If
   End With


ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   CheckOC

End Function

Private Sub Form_Load()
 
   MoveFormToCenter Me
   
   '更新系統時間
   'Modified by Morgan 2017/1/4
   'Date = Format(strSrvDate(1), "####/##/##") '校正日期與DB同步
   'time = Format(ServerTime, "##:##:##")   '校正時間與DB同步
   PUB_SyncClientDateTime
   'end 2017/1/4
   lblServerTime.Caption = Format(Now, "HH:MM:SS")
   Timer1.Interval = 1000
   
   '分段時間
   MaskEdBox1.Mask = ""
   MaskEdBox1 = lblServerTime
   MaskEdBox1.Mask = "##:##:##"
   
   '發文日期
   MaskEdBox2.Mask = ""
   MaskEdBox2 = Format(strSrvDate(2), "###/##/##")
   MaskEdBox2.Mask = "###/##/##"
   
   '印表機
   PUB_SetPrinter Me.Name, Combo1, strPrinter      'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
      
   '列印人員
   txtPrinter.Text = strUserNum
   lblPrinter.Caption = GetStaffName(strUserNum)
   
   '清單種類
   cboListType.Clear
   cboListType.AddItem "內專"
   cboListType.AddItem "內商"
   cboListType.AddItem "外商"
   cboListType.AddItem "外專"
   
   '送件時段
   cboListTime.Clear
   cboListTime.AddItem "上午"
   cboListTime.AddItem "下午"
   '12點前預設上午
   If lblServerTime < "12:00:00" Then
      cboListTime.ListIndex = 0
   Else
      cboListTime.ListIndex = 1
   End If
   
   SetRef
   
   If Pub_StrUserSt03 = "M51" Then
      txtPrinter.Enabled = True
   End If
   
End Sub

Private Sub SetRef()
   m_Dept = Left(GetStaffDepartment(txtPrinter), 2)
   Select Case m_Dept
      Case "P1": cboListType.ListIndex = 0
      Case "P2": cboListType.ListIndex = 1
      Case "F1": cboListType.ListIndex = 2
      Case "F2": cboListType.ListIndex = 3
      Case Else: cboListType.ListIndex = -1
   End Select
   GetTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
       '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    MenuEnabled
   Set frm1109 = Nothing
End Sub

Private Sub MaskEdBox1_GotFocus()
   MaskEdBoxInverse MaskEdBox1
End Sub

Private Sub MaskEdBox2_GotFocus()
   MaskEdBoxInverse MaskEdBox2
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   strExc(1) = Replace(MaskEdBox2, "/", "")
   If ChkDate(strExc(1)) = False Then
      Cancel = True
      MaskEdBox2.SetFocus
      MaskEdBox2_GotFocus
   End If
End Sub

Private Sub Timer1_Timer()
   lblServerTime.Caption = Format(Now, "HH:MM:SS")
End Sub

Private Function GetTime() As Boolean
   Dim stTime As String
   
On Error GoTo ErrHnd
   
   If cboListType.ListIndex >= 0 Then
      MaskEdBox1.Enabled = True
      strSql = "select AL05,AL06 from APPLISTe" & _
         " where AL01=" & DBDATE(MaskEdBox2) & " and AL02='" & m_Dept & "' ORDER BY AL06"
      
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
         '已列印
         If .RecordCount > 0 Then
            stTime = Format("" & .Fields("AL05"), "00:00:00")
            ' 若財務處當日已產生傳票分錄則分段時間不可再改
            If Not IsNull(.Fields(1)) Then
               MaskEdBox1.Enabled = False
            End If
            '下午都不可改分隔時間以免發生上下午不同
            If cboListTime.ListIndex = 1 Then
               MaskEdBox1.Enabled = False
            End If
         '未列印
         Else
            MaskEdBox1.Enabled = True
            If cboListTime.ListIndex = 1 Then
               stTime = "00:00:00"
            Else
               stTime = lblServerTime
            End If
         End If
      End With
      
      If stTime <> "" Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = stTime
         MaskEdBox1.Mask = "##:##:##"
      End If
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
   
End Function

Private Sub txtPrinter_GotFocus()
   TextInverse txtPrinter
End Sub

Private Sub txtPrinter_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPrinter_Validate(Cancel As Boolean)
   Dim strTempName As String
   If txtPrinter <> "" Then
      If ClsPDGetStaff(txtPrinter, strTempName) = True Then
         lblPrinter = strTempName
         SetRef
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub MaskEdBoxInverse(pBox As MaskEdBox)
   pBox.SelStart = 0
   pBox.SelLength = Len(pBox.Text)
End Sub
