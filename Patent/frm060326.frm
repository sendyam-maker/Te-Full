VERSION 5.00
Begin VB.Form frm060326 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯點數/核稿件數統計表"
   ClientHeight    =   1080
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3210
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1485
      MaxLength       =   5
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   2205
      TabIndex        =   2
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1170
      TabIndex        =   1
      Top             =   30
      Width           =   912
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "統計年月："
      Height          =   180
      Left            =   630
      TabIndex        =   3
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frm060326"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit
Const cntX As Long = 500 '左邊界
Const cntY As Long = 500 '上邊界
Const cntL As Long = 400 '列高
Dim iX As Long, iY As Long '現在列印位置
Dim iPLeft(1 To 10) As Long '各欄位起始X座標
Dim iSum(1 To 4) As Long, iOutSum(1 To 2) As Long

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
   Set frm060326 = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   If Text1 = "" Then
      MsgBox Label1 & "條件不可空白！", vbExclamation
      Cancel = True
      Text1.SetFocus
      Text1_GotFocus
   Else
      Cancel = False
      Text1_Validate Cancel
   End If
   TxtValidate = Not Cancel
End Function

Private Sub cmdPrint_Click(Index As Integer)
   
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      If DoPrint = True Then
         MsgBox "列印完成！"
      End If
   End If
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub SetPLeft()
   iPLeft(1) = cntX
   iPLeft(2) = iPLeft(1) + 1200
   iPLeft(3) = iPLeft(2) + 1200
   iPLeft(4) = iPLeft(3) + 2300
   iPLeft(5) = iPLeft(4) + 4800
End Sub

Private Sub PrintHead(stTitle As String, iPage As Integer, iPageTot As Integer)

   iY = cntY
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5500 - (Printer.TextWidth(stTitle) / 2)
   Printer.CurrentY = iY
   Printer.Print stTitle
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   iY = Printer.CurrentY
   Printer.CurrentX = cntX
   Printer.CurrentY = iY
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iY
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   
   iY = iY + cntL
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iY
   Printer.Print "列印時間：" & Format(ServerTime, Tformat)
   
   iY = iY + cntL
   Printer.CurrentX = cntX
   Printer.CurrentY = iY
   Printer.Print "統計年月：" & Text1
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iY
   Printer.Print "頁　　次：" & Format(iPage) & "/" & Format(iPageTot)
   

   PrintLine
   
   iY = iY + cntL
   Printer.CurrentX = iPLeft(1)
   Printer.CurrentY = iY
   Printer.Print "工程師"
   Printer.CurrentX = iPLeft(2)
   Printer.CurrentY = iY
   Printer.Print "組別"
   Printer.CurrentX = iPLeft(3)
   Printer.CurrentY = iY
   Printer.Print "翻譯點數/件數"
   Printer.CurrentX = iPLeft(4)
   Printer.CurrentY = iY
   Printer.Print "I.U.核稿或D.承辦件數/日本申請人件數"
   Printer.CurrentX = iPLeft(5)
   Printer.CurrentY = iY
   Printer.Print "備註"
   
   iY = iY + cntL
   Printer.CurrentX = iPLeft(4)
   Printer.CurrentY = iY
   Printer.Print "(翻譯201或檢視209或製作中說210發文)"
      
   PrintLine
   
End Sub

Private Sub PrintHead1(stTitle As String, iPage As Integer, iPageTot As Integer)

   iY = cntY
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5500 - (Printer.TextWidth(stTitle) / 2)
   Printer.CurrentY = iY
   Printer.Print stTitle
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   iY = Printer.CurrentY
   Printer.CurrentX = cntX
   Printer.CurrentY = iY
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iY
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   
   iY = iY + cntL
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iY
   Printer.Print "列印時間：" & Format(ServerTime, Tformat)
   
   iY = iY + cntL
   Printer.CurrentX = cntX
   Printer.CurrentY = iY
   Printer.Print "統計年月：" & Text1
   Printer.CurrentX = cntX + 8000
   Printer.CurrentY = iY
   Printer.Print "頁　　次：" & Format(iPage) & "/" & Format(iPageTot)
   

   PrintLine
   
   iY = iY + cntL
   Printer.CurrentX = iPLeft(1)
   Printer.CurrentY = iY
   Printer.Print "請款單號"
   Printer.CurrentX = iPLeft(2)
   Printer.CurrentY = iY
   Printer.Print "本所案號"
      
   PrintLine
   
End Sub
Private Sub PrintLine()
   iY = iY + cntL
   Printer.CurrentX = cntX
   Printer.Print String(180, "-")
   iY = iY - cntL / 2
End Sub

Private Sub PrintTail()
   Dim strTmp As String
   iY = iY + cntL
   iX = iPLeft(1)
   Printer.CurrentX = iX: Printer.CurrentY = iY
   Printer.Print "FCP" & Val(Text1) Mod 100 & "月總數"
   '翻譯點數/件數
   strTmp = iSum(1)
   iX = iPLeft(3) + Printer.TextWidth("翻譯點數") - Printer.TextWidth(strTmp)
   Printer.CurrentX = iX: Printer.CurrentY = iY
   Printer.Print strTmp & "/" & iSum(2)
   'I.U.核稿或D.承辦件數/日本申請人件數
   strTmp = iSum(3)
   iX = iPLeft(4) + Printer.TextWidth("I.U.核稿或D.承辦件數") - Printer.TextWidth(strTmp)
   Printer.CurrentX = iX: Printer.CurrentY = iY
   Printer.Print strTmp & "/" & iSum(4)
   
   '外翻統計
   iY = iY + cntL
   strTmp = "( " & iOutSum(1)
   iX = iPLeft(3) + Printer.TextWidth("翻譯點數") - Printer.TextWidth(strTmp)
   Printer.CurrentX = iX: Printer.CurrentY = iY
   Printer.Print strTmp & "/" & iOutSum(2) & " )"
   
End Sub

Private Function DoPrint() As Boolean
   Dim stVTB1 As String, stVTB2 As String, stVTB3 As String
   Dim stGroup As String
   Dim strTitle As String, strTmp As String
   Dim iPage As Integer, iPageTot As Integer
   Dim iRec As Integer, iRecs As Integer, lngTot As Long
   Const PageRec As Integer = 30
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Label1 & Text1 'Add By Sindy 2010/12/13
   
   '虛擬表格1:請款統計資料(翻譯)
   stVTB1 = "SELECT CP14,SUM(FEAT) TFEAT,COUNT(*) CNT1" & _
      " FROM ( SELECT A1L01,NVL(A1L05,0)-NVL(A1L07,0) FEAT" & _
         " FROM acc1k0, acc1l0" & _
         " where a1k02>=" & Text1 & "01 and a1k02<=" & Text1 & "31 and a1k13||''='FCP'" & _
         " and a1l01(+)=a1k01 AND a1l04='201'" & _
      " ) X,CASEPROGRESS WHERE CP60(+)=A1L01 AND CP10='201' GROUP BY CP14"
       
   '虛擬表格2:發文統計資料(翻譯,檢視中說,製作中說)
   'Modified by Morgan 2013/11/6 +235核對中說格式
   stVTB2 = "SELECT NVL(EP04,CP14) NO2,COUNT(*) CNT2,SUM(DECODE(SUBSTR(CU10,1,3),'011',1,0)) CNT3" & _
      " From CASEPROGRESS, ENGINEERPROGRESS, PATENT, CUSTOMER" & _
      " WHERE CP27>=" & (Text1 + 191100) & "01 AND CP27<=" & (Text1 + 191100) & "31 AND CP01||''='FCP' AND CP10||'' IN ('201','209','210','235')" & _
      " AND EP02(+)=CP09" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
      " GROUP BY NVL(EP04,CP14)"
   
   '2008/4/8 MODIF BY SONIA 加ST03=F81
   'strSQL = "SELECT ST02,nvl(TFEAT,0) TFEAT,NVL(CNT1,0) CNT1,NVL(CNT2,0) CNT2,NVL(CNT3,0) CNT3,ST16,ST05,ST01" & _
   '   " FROM STAFF, (" & stVTB1 & ") VTB1, (" & stVTB2 & ") VTB2" & _
   '   " WHERE ST03='F21' AND ST04='1' AND CP14(+)=ST01 AND NO2(+)=ST01"
   strSql = "SELECT ST02,nvl(TFEAT,0) TFEAT,NVL(CNT1,0) CNT1,NVL(CNT2,0) CNT2,NVL(CNT3,0) CNT3,ST16,ST05,ST01" & _
      " FROM STAFF, (" & stVTB1 & ") VTB1, (" & stVTB2 & ") VTB2" & _
      " WHERE (ST15='F21' OR ST15='F81') AND ST04='1' AND CP14(+)=ST01 AND NO2(+)=ST01"
      
   '所有點數/件數統計
   strSql = strSql & " UNION ALL" & _
      " SELECT NULL,nvl(SUM(TFEAT),0) TFEAT,NVL(SUM(CNT1),0) CNT1,0,0,'0',NULL,NULL" & _
      " FROM (" & stVTB1 & ") VTB1" & _
      ""
      
   strSql = strSql & " ORDER BY ST16,ST05,ST01"

On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
         Erase iSum
         Printer.Orientation = 1
         Printer.Font = "細明體"
         .MoveFirst
         iPage = 1: lngTot = 0: iRec = 0: iRecs = 0
         iPageTot = .RecordCount \ PageRec + IIf((.RecordCount Mod PageRec) = 0, 0, 1)
         strTitle = "FCP" & Me.Caption
         PrintHead strTitle, iPage, iPageTot
         '所有點數/件數統計
         If "" & .Fields("ST16") = "0" Then
            iOutSum(1) = Format(.Fields("TFEAT") / 1000, "0")
            iOutSum(2) = Val(.Fields("CNT1"))
            .MoveNext
         End If
         Do While Not .EOF
            iRec = iRec + 1: iRecs = iRecs + 1: lngTot = lngTot + .Fields(1)
            If iRec > PageRec Then
               Printer.NewPage
               iPage = iPage + 1
               PrintHead strTitle, iPage, iPageTot
               iRec = 0
            End If
            iY = iY + cntL
            '工程師
            Printer.CurrentX = iPLeft(1)
            Printer.CurrentY = iY
            Printer.Print "" & .Fields("ST02")
            If "" & .Fields("ST16") <> stGroup Then
               stGroup = "" & .Fields("ST16")
               Printer.CurrentX = iPLeft(2)
               Printer.CurrentY = iY
               '2010/1/8 modify by sonia 改有function
               'Select Case stGroup
               '   Case "1"
               '      Printer.Print "電機組"
               '   Case "2"
               '      Printer.Print "化學組"
               '   Case "3"
               '      Printer.Print "日文組"
               '   '2008/2/22 add by sonia
               '   Case "4"
               '      Printer.Print "德文組"
               'End Select
               Printer.Print PUB_GetFCPGrpName(stGroup, True)
               '2010/1/8 end
            End If
            '翻譯點數/件數
            strTmp = Format(.Fields("TFEAT") / 1000, "0")
            iSum(1) = iSum(1) + Val(strTmp)
            iSum(2) = iSum(2) + Val(.Fields("CNT1"))
            iSum(3) = iSum(3) + Val(.Fields("CNT2"))
            iSum(4) = iSum(4) + Val(.Fields("CNT3"))
            iX = iPLeft(3) + Printer.TextWidth("翻譯點數") - Printer.TextWidth(strTmp)
            Printer.CurrentX = iX: Printer.CurrentY = iY
            Printer.Print strTmp & "/" & .Fields("CNT1")
            'I.U.核稿或D.承辦件數/日本申請人件數
            strTmp = Format("" & .Fields("CNT2"))
            iX = iPLeft(4) + Printer.TextWidth("I.U.核稿或D.承辦件數") - Printer.TextWidth(strTmp)
            Printer.CurrentX = iX: Printer.CurrentY = iY
            Printer.Print strTmp & "/" & .Fields("CNT3")
            .MoveNext
         Loop
         PrintLine
         PrintTail
         PrintExtra
         Printer.EndDoc
         DoPrint = True
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/13
         MsgBox "無可列印資料！", vbInformation
      End If
   End With

ErrHnd:

   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'列印無承辦人翻譯案件
Private Sub PrintExtra()

   Dim strTitle As String, strTmp As String
   Dim iPage As Integer, iPageTot As Integer
   Dim iRec As Integer, iRecs As Integer, lngTot As Long
   Const PageRec As Integer = 30
   
   strSql = "SELECT CP60,CP01||'-'||CP02||'-'||CP03||'-'||CP04 CN" & _
      " FROM ( SELECT A1L01,NVL(A1L05,0)-NVL(A1L07,0) FEAT" & _
         " FROM acc1k0, acc1l0" & _
         " where a1k02>=" & Text1 & "01 and a1k02<=" & Text1 & "31 and a1k13||''='FCP'" & _
         " and a1l01(+)=a1k01 AND a1l04='201'" & _
      " ) X,CASEPROGRESS WHERE CP60(+)=A1L01 AND CP10='201' AND CP14 IS NULL"
      
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         iY = iY + cntL
         Printer.CurrentX = iPLeft(1)
         Printer.CurrentY = iY
         Printer.Print "待續..."
         
         iPage = 1: lngTot = 0: iRec = 0: iRecs = 0
         iPageTot = .RecordCount \ PageRec + IIf((.RecordCount Mod PageRec) = 0, 0, 1)
         strTitle = "FCP已請款無承辦翻譯案件清單"
         Printer.NewPage
         PrintHead1 strTitle, iPage, iPageTot
         Do While Not .EOF
            iRec = iRec + 1: iRecs = iRecs + 1
            If iRec > PageRec Then
               Printer.NewPage
               iPage = iPage + 1
               PrintHead1 strTitle, iPage, iPageTot
               iRec = 0
            End If
            iY = iY + cntL
            Printer.CurrentX = iPLeft(1)
            Printer.CurrentY = iY
            Printer.Print "" & .Fields(0)
            Printer.CurrentX = iPLeft(2)
            Printer.CurrentY = iY
            Printer.Print "" & .Fields(1)
            .MoveNext
         Loop
         PrintLine
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetPLeft
   'Modified by Lydia 2019/10/28
   'Text1 = Left(TransDate(CompDate(1, -1, strSrvDate(1)), 1), 4)
   Text1 = Mid(strSrvDate(2), 1, Len(strSrvDate(2)) - 2)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "" Then
      If CheckIsTaiwanDate(Text1 & "01") = False Then
         Cancel = True
      End If
   End If
End Sub
