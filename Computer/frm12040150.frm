VERSION 5.00
Begin VB.Form frm12040150 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員客戶名冊 (依最後收文日期區間)"
   ClientHeight    =   3630
   ClientLeft      =   2960
   ClientTop       =   1620
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5340
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   7
      Top             =   2250
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   8
      Left            =   3180
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2250
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1845
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   3180
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1845
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   3180
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   3180
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   10
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   10
      Text            =   "N"
      Top             =   3000
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   9
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "N"
      Top             =   2715
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   0
      Top             =   630
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4230
      TabIndex        =   12
      Top             =   48
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3285
      TabIndex        =   11
      Top             =   48
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "讀檔進度："
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3270
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第四段收文日期"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   2295
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2940
      X2              =   3060
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第三段收文日期"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   19
      Top             =   1890
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2940
      X2              =   3060
      Y1              =   1965
      Y2              =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第二段收文日期"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   1485
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2940
      X2              =   3060
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第一段收文日期"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2940
      X2              =   3060
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "是否含有客戶狀態者:             (N: 不含)"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   3075
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "是否含不寄雜誌對象:             (N: 不含)"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   2310
      TabIndex        =   14
      Top             =   690
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   13
      Top             =   675
      Width           =   720
   End
End
Attribute VB_Name = "frm12040150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2011/3/8 改寫 by sonia
'Modify By Sindy 2012/6/25 接著後續完成
Option Explicit

Dim PLeft(0 To 8) As Integer
' 預設印表機
Dim m_DefaultPrinter As String

Private Sub cmdok_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   
   Select Case Index
      Case 0 '確定
         '檢查輸入的資料是否齊全完整
         If CheckDataValid() = False Then
            GoTo EXITSUB
         End If
         
         '名冊
         PrintCase
            
      Case 1 '結束
         Unload Me
   End Select
EXITSUB:
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Prn As Printer
   
   ' 將印表機設為原先的預設印表機
   For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next
   
   Set frm12040150 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 9, 10
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp As String
   
   Select Case Index
      Case 0
         lblName.Caption = ""
         If Text1(Index) <> "" Then
            lblName.Caption = GetPrjSalesNM(Text1(Index))
            If Len(Text1(Index)) <> 0 Then
               If Len(lblName.Caption) = 0 Then
                  Cancel = True
                  MsgBox "智權人員輸入錯誤！", vbCritical
                  Text1(Index).SetFocus
                  Text1_GotFocus (Index)
                  Exit Sub
               End If
            End If
         End If
      Case 1, 2, 3, 4, 5, 6, 7, 8
         If Text1(Index) <> "" Then
            Cancel = Not ChkDate(Text1(Index).Text)
         End If
         If Index = 1 Or Index = 3 Or Index = 5 Or Index = 7 Then
            If Text1(Index) <> "" And Text1(Index + 1) = "" Then
               Text1(Index + 1) = Text1(Index)
            End If
         ElseIf Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
            If RunNick2(Text1(Index - 1), Text1(Index)) Then
               Call Text1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub PrintCase()
Dim Page As Integer, iPrint As Integer, i As Integer, j As Integer
Dim strNo As String '員工代號
Dim ReportID As Integer
Dim rsTmp As New ADODB.Recordset
Dim bolConnection  As Boolean
Dim intRow As Integer
Dim strCon As String
   
On Error GoTo ErrHand
   
   strCon = ""
   '是否含不寄雜誌對象
   If Text1(9) = "N" Then
      strCon = strCon & " AND CU32 IS NULL "
   End If
   '是否含有客戶狀態者
   If Text1(10) = "N" Then
      'modify by sonia 2023/2/23 再加其他,不得代理專利,不得代理商標,解除對造,國內同業
      'Modify By Sindy 2025/6/27 +OR CU80='業務自行處理' or cu80='其他' or cu80='不得代理專利' or cu80='不得代理商標' or cu80='解除對造' or cu80='國內同業'
      '                          改抓常變數
      strCon = strCon & " AND (CU80 IS NULL or instr('" & 客戶及代理人可讀取的狀態 & "',cu80)>0) "
   End If
   
   '產生資料
   bolConnection = True
   cnnConnection.BeginTrans
   '刪除暫存檔中該智權人員資料
   strSql = "delete from R12040150 where sales='" & Text1(0) & "'"
   cnnConnection.Execute strSql
   '先讀出該智權人員所有客戶
   strExc(0) = "select cu01,cu02 from customer where cu13='" & Text1(0) & "' and cu02='0'" & strCon & " ORDER BY CU12,CU13,CU01,CU02"
   intI = 1: intRow = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Label2.Visible = True
      Do While Not RsTemp.EOF
         '逐筆抓出最後A類收文日期寫入R12040150(8碼客戶編號+8碼最後A類收文日)
         'Modify By Sindy 2012/7/3 使用substr(tm23,1,8)='" & RsTemp.Fields("cu01") & "'抓資料慢,改用>=及<=
         strSql = "select max(cp05) from caseprogress,( " & _
                  "select tm01 as t1,tm02 as t2,tm03 as t3,tm04 as t4 from trademark where (tm23>='" & RsTemp.Fields("cu01") & "0' and tm23<='" & RsTemp.Fields("cu01") & "9') or (tm78>='" & RsTemp.Fields("cu01") & "0' and tm78<='" & RsTemp.Fields("cu01") & "9') or (tm79>='" & RsTemp.Fields("cu01") & "0' and tm79<='" & RsTemp.Fields("cu01") & "9') or (tm80>='" & RsTemp.Fields("cu01") & "0' and tm80<='" & RsTemp.Fields("cu01") & "9') or (tm81>='" & RsTemp.Fields("cu01") & "0' and tm81<='" & RsTemp.Fields("cu01") & "9') " & _
                  "Union " & _
                  "select pa01 as t1,pa02 as t2,pa03 as t3,pa04 as t4 from patent where (pa26>='" & RsTemp.Fields("cu01") & "0' and pa26<='" & RsTemp.Fields("cu01") & "9') or (pa27>='" & RsTemp.Fields("cu01") & "0' and pa27<='" & RsTemp.Fields("cu01") & "9') or (pa28>='" & RsTemp.Fields("cu01") & "0' and pa28<='" & RsTemp.Fields("cu01") & "9') or (pa29>='" & RsTemp.Fields("cu01") & "0' and pa29<='" & RsTemp.Fields("cu01") & "9') or (pa30>='" & RsTemp.Fields("cu01") & "0' and pa30<='" & RsTemp.Fields("cu01") & "9') " & _
                  "Union " & _
                  "select sp01 as t1,sp02 as t2,sp03 as t3,sp04 as t4 from servicepractice where (sp08>='" & RsTemp.Fields("cu01") & "0' and sp08<='" & RsTemp.Fields("cu01") & "9') or (sp58>='" & RsTemp.Fields("cu01") & "0' and sp58<='" & RsTemp.Fields("cu01") & "9') or (sp59>='" & RsTemp.Fields("cu01") & "0' and sp59<='" & RsTemp.Fields("cu01") & "9') or (sp65>='" & RsTemp.Fields("cu01") & "0' and sp65<='" & RsTemp.Fields("cu01") & "9') or (sp66>='" & RsTemp.Fields("cu01") & "0' and sp66<='" & RsTemp.Fields("cu01") & "9') " & _
                  "Union " & _
                  "select hc01 as t1,hc02 as t2,hc03 as t3,hc04 as t4 from hirecase where (hc05>='" & RsTemp.Fields("cu01") & "0' and hc05<='" & RsTemp.Fields("cu01") & "9') or (hc24>='" & RsTemp.Fields("cu01") & "0' and hc24<='" & RsTemp.Fields("cu01") & "9') or (hc25>='" & RsTemp.Fields("cu01") & "0' and hc25<='" & RsTemp.Fields("cu01") & "9') or (hc26>='" & RsTemp.Fields("cu01") & "0' and hc26<='" & RsTemp.Fields("cu01") & "9') or (hc27>='" & RsTemp.Fields("cu01") & "0' and hc27<='" & RsTemp.Fields("cu01") & "9') " & _
                  "Union " & _
                  "select lc01 as t1,lc02 as t2,lc03 as t3,lc04 as t4 from lawcase where (lc11>='" & RsTemp.Fields("cu01") & "0' and lc11<='" & RsTemp.Fields("cu01") & "9') or (lc43>='" & RsTemp.Fields("cu01") & "0' and lc43<='" & RsTemp.Fields("cu01") & "9') or (lc44>='" & RsTemp.Fields("cu01") & "0' and lc44<='" & RsTemp.Fields("cu01") & "9') or (lc45>='" & RsTemp.Fields("cu01") & "0' and lc45<='" & RsTemp.Fields("cu01") & "9') or (lc46>='" & RsTemp.Fields("cu01") & "0' and lc46<='" & RsTemp.Fields("cu01") & "9') " & _
                  ") a " & _
                  "where a.t1=cp01(+) and a.t2=cp02(+) and a.t3=cp03(+) and a.t4=cp04(+) " & _
                  "and cp09<'B' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection
         If rsTmp.RecordCount > 0 Then
            strDate = "" & rsTmp.Fields(0)
         End If
         rsTmp.Close

         strSql = "insert into R12040150 values('" & RsTemp.Fields("cu01") & "'," & CNULL(strDate) & "," & CNULL(Text1(0)) & ")"
         cnnConnection.Execute strSql
         intRow = intRow + 1
         Label2.Caption = "讀檔進度：" & intRow & " / " & RsTemp.RecordCount: DoEvents
         RsTemp.MoveNext
      Loop
      Label2.Visible = False
      cnnConnection.CommitTrans
      bolConnection = False
   Else
      cnnConnection.CommitTrans
      bolConnection = False
      MsgBox "該智權人員無客戶資料！"
      Exit Sub
   End If
   
   '再以R12040150及畫面條件抓基本資料列印名冊
   strSql = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,PCC05,CU22 " & _
                "FROM CUSTOMER,POTCUSTCONT,NATION,STAFF,R12040150 " & _
                "WHERE SALES='" & Me.Text1(0).Text & "' AND CUNO=CU01(+) AND cu02='0' AND CU10=NA01(+) AND CU13=ST01(+) AND PCC01(+)=CU01 AND PCC02(+)=CU127 " & strCon
   For ReportID = 1 To 5
      If ReportID = 1 Then
         If Val(Text1(1)) > 0 And Val(Text1(2)) > 0 Then
            strExc(0) = strSql + " AND RECDATE Between " & DBDATE(Text1(1)) & " AND " & DBDATE(Text1(2))
         Else
            GoTo ReadNext
         End If
      End If
      If ReportID = 2 Then
         If Val(Text1(3)) > 0 And Val(Text1(4)) > 0 Then
            strExc(0) = strSql + " AND RECDATE Between " & DBDATE(Text1(3)) & " AND " & DBDATE(Text1(4))
         Else
            GoTo ReadNext
         End If
      End If
      If ReportID = 3 Then
         If Val(Text1(5)) > 0 And Val(Text1(6)) > 0 Then
            strExc(0) = strSql + " AND RECDATE Between " & DBDATE(Text1(5)) & " AND " & DBDATE(Text1(6))
         Else
            GoTo ReadNext
         End If
      End If
      If ReportID = 4 Then
         If Val(Text1(7)) > 0 And Val(Text1(8)) > 0 Then
            strExc(0) = strSql + " AND RECDATE Between " & DBDATE(Text1(7)) & " AND " & DBDATE(Text1(8))
         Else
            GoTo ReadNext
         End If
      End If
      If ReportID = 5 Then
         strExc(0) = strSql & " and CUNO not in(select CUNO from R12040150 where SALES='" & Me.Text1(0).Text & "' AND (" & _
                     "(RECDATE Between " & DBDATE(Text1(1)) & " AND " & DBDATE(Text1(2)) & ")"
         If Val(Text1(3)) > 0 And Val(Text1(4)) > 0 Then
            strExc(0) = strExc(0) + " or (RECDATE Between " & DBDATE(Text1(3)) & " AND " & DBDATE(Text1(4)) & ")"
         End If
         If Val(Text1(5)) > 0 And Val(Text1(6)) > 0 Then
            strExc(0) = strExc(0) + " or (RECDATE Between " & DBDATE(Text1(5)) & " AND " & DBDATE(Text1(6)) & ")"
         End If
         If Val(Text1(7)) > 0 And Val(Text1(8)) > 0 Then
            strExc(0) = strExc(0) + " or (RECDATE Between " & DBDATE(Text1(7)) & " AND " & DBDATE(Text1(8)) & ")"
         End If
         strExc(0) = strExc(0) & "))"
      End If
      strExc(0) = strExc(0) + " ORDER BY CU12,CU13,CU01,CU02 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Screen.MousePointer = vbHourglass
         GetPrintLeft
         Page = 1
         CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value, ReportID
         iPrint = 2700 + 300 + 300
         If Not IsNull(RsTemp.Fields(10).Value) Then strNo = RsTemp.Fields(10).Value
         
         i = 0
         With RsTemp
            i = 0
            Do While Not .EOF
               For j = 0 To 4
                  Printer.CurrentX = PLeft(j)
                  Printer.CurrentY = iPrint
                  If j = 0 Then
                     If Not IsNull(RsTemp.Fields(9).Value) Then
                        Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                        If Not IsNull(RsTemp.Fields(8).Value) Then
                           Printer.Print "＊" & Left(.Fields(j) & "000", 9) & " N"
                        Else
                           Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                        End If
                     Else
                        If Not IsNull(RsTemp.Fields(8).Value) Then
                           Printer.Print " " & Left(.Fields(j) & "000", 9) & " N"
                        Else
                           Printer.Print "" & Left(.Fields(j) & "000", 9)
                        End If
                     End If
                  Else
                     Printer.Print "" & .Fields(j)
                  End If
               Next j
               
               iPrint = iPrint + 300
               
               For j = 5 To 7
                  If j = 7 Then
                     Printer.CurrentX = PLeft(8)
                     Printer.CurrentY = iPrint
                     Printer.Print "" & .Fields(15)
                  End If
                  Printer.CurrentX = PLeft(j)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields(j)
               Next j
               
               iPrint = iPrint + 300
                           
               Printer.CurrentX = PLeft(5)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(9)
               Printer.CurrentX = PLeft(6)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(13)
               Printer.CurrentX = PLeft(8)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(16)
               Printer.CurrentX = PLeft(7)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(14)
               iPrint = iPrint + 300
                           
               Printer.CurrentX = PLeft(0)
               Printer.CurrentY = iPrint
               Printer.Print String(250, "-")
               
               iPrint = iPrint + 300
               
               i = i + 1
               .MoveNext
               If RsTemp.EOF Then Exit Do
               If i > 9 Or "" & RsTemp.Fields(10).Value <> strNo Then
                  strNo = "" & RsTemp.Fields(10).Value
                  Printer.CurrentX = PLeft(0)
                  Printer.CurrentY = iPrint
                  Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
                  Printer.NewPage
                  Page = Page + 1
                  CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value, ReportID
                  iPrint = 2700 + 300 + 300
                  i = 0
               End If
            Loop
         End With
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
         
         Printer.EndDoc
      End If
ReadNext:
   Next ReportID
   
   ShowPrintOk
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   MsgBox Err.Description
   If bolConnection = True Then
      cnnConnection.RollbackTrans
   End If
End Sub

Private Sub GetPrintLeft()
   '第一列
   PLeft(0) = 200
   PLeft(1) = 1500
   PLeft(2) = 4200 + 3000 - 500 - 500
   PLeft(3) = 5200 + 3000 - 300
   PLeft(4) = 6200 + 3000 + 500 - 300
   '第二列
   PLeft(5) = 200
   PLeft(6) = 1500
   PLeft(7) = 6200 + 3000 + 500 - 300
   PLeft(8) = 5200 + 3000 - 300          '接洽人及手機
End Sub

Private Sub CaseTitle(ByVal Page As String, ByVal strSNo As String, ByVal strSName As String, Optional iReportID As Integer, Optional strDept As String)
Dim i As Integer
'Page : 頁數
'strSNo : 員工編號
'strSName : 員工姓名
   
   i = 500
   
   If Page = 1 Then Printer.Orientation = vbPRORPortrait
   Printer.FontName = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("智權人員客戶名冊 ") / 2)
   Printer.CurrentY = i
   Printer.Print "智權人員客戶名冊 "
   Printer.Font.Underline = False
   
   strExc(2) = ""
   Select Case iReportID
      Case 1
         strExc(2) = "(收文日期：" & ChangeTStringToTDateString(Text1(1)) & " 至 " & ChangeTStringToTDateString(Text1(2)) & ")"
      Case 2
         strExc(2) = "(收文日期：" & ChangeTStringToTDateString(Text1(3)) & " 至 " & ChangeTStringToTDateString(Text1(4)) & ")"
      Case 3
         strExc(2) = "(收文日期：" & ChangeTStringToTDateString(Text1(5)) & " 至 " & ChangeTStringToTDateString(Text1(6)) & ")"
      Case 4
         strExc(2) = "(收文日期：" & ChangeTStringToTDateString(Text1(7)) & " 至 " & ChangeTStringToTDateString(Text1(8)) & ")"
      Case 5
         If ChangeTStringToTDateString(Text1(2)) <> "" Then strDate = ChangeTStringToTDateString(Text1(1))
         If ChangeTStringToTDateString(Text1(4)) <> "" Then strDate = ChangeTStringToTDateString(Text1(3))
         If ChangeTStringToTDateString(Text1(6)) <> "" Then strDate = ChangeTStringToTDateString(Text1(5))
         If ChangeTStringToTDateString(Text1(8)) <> "" Then strDate = ChangeTStringToTDateString(Text1(7))
         strExc(2) = "(" & strDate & "之後未收文)"
   End Select
   If strExc(2) <> "" Then
      Printer.Font.Size = 14
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(2)) / 2)
      Printer.CurrentY = i + 500
      Printer.Print strExc(2)
   End If
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 800 - 300
   Printer.Print "列印人　 : " & strUserName
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 800
   Printer.Print "智權人員 : " & strSNo & " " & strSName & " " & strDept
   
   strExc(2) = ""
   If Text1(9) = "N" Then
      strExc(2) = "不含不寄雜誌對象"
   End If
   If Text1(10) = "N" Then
      If strExc(2) <> "" Then strExc(2) = strExc(2) & "；"
      strExc(2) = strExc(2) & "不含有客戶狀態者"
   End If
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(2)) / 2)
   Printer.CurrentY = i + 800
   Printer.Print strExc(2)
   
   Printer.CurrentX = 7000 + 1500
   Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))
   Printer.CurrentX = 7000 + 1500
   Printer.CurrentY = i + 1100
   Printer.Print "頁　　次 : " & Page
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1400
   Printer.Print String(250, "-")
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1700
   Printer.Print "編號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = i + 1700
   Printer.Print "公司名稱"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = i + 1700
   Printer.Print "負責人"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = i + 1700
   Printer.Print "國籍"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = i + 1700
   Printer.Print "電話"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "郵遞區號"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "客戶聯絡地址"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "接洽人"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "傳真"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = i + 1700 + 300 + 300
   Printer.Print "客戶狀態"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = i + 1700 + 300 + 300
   Printer.Print "中文地址"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = i + 1700 + 300 + 300
   Printer.Print "手機"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = i + 1700 + 300 + 300
   Printer.Print "統一編號"
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 2000 + 300 + 300
   Printer.Print String(250, "-")
   Printer.Font.Size = 10
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   lblName.Caption = ""
   If Text1(0) <> "" Then
      lblName.Caption = GetPrjSalesNM(Text1(0))
      If Len(Text1(0)) <> 0 Then
         If Len(lblName.Caption) = 0 Then
            MsgBox "智權人員輸入錯誤！", vbCritical
            Text1(0).SetFocus
            Call Text1_GotFocus(0)
            GoTo EXITSUB
         End If
      End If
   End If
   
   If Text1(1) = "" Then
      MsgBox "第一段收文起始日期不可空白！", vbCritical
      Text1(1).SetFocus
      Call Text1_GotFocus(1)
      GoTo EXITSUB
   End If
   If Text1(2) = "" Then
      MsgBox "第一段收文截止日期不可空白！", vbCritical
      Text1(2).SetFocus
      Call Text1_GotFocus(2)
      GoTo EXITSUB
   End If
      
   CheckDataValid = True

EXITSUB:
End Function
