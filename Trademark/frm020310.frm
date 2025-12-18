VERSION 5.00
Begin VB.Form frm020310 
   BorderStyle     =   1  '單線固定
   Caption         =   "對外案件延展未提申明細表"
   ClientHeight    =   1320
   ClientLeft      =   3645
   ClientTop       =   1965
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5370
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1395
      MaxLength       =   8
      TabIndex        =   0
      Top             =   735
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2595
      MaxLength       =   8
      TabIndex        =   1
      Top             =   735
      Width           =   1110
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3255
      TabIndex        =   2
      Top             =   105
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4050
      TabIndex        =   3
      Top             =   105
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "（西元年日期）"
      Height          =   180
      Index           =   0
      Left            =   3735
      TabIndex        =   5
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "專用期止日："
      Height          =   180
      Index           =   3
      Left            =   255
      TabIndex        =   4
      Top             =   780
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2610
      Y1              =   885
      Y2              =   885
   End
End
Attribute VB_Name = "frm020310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer
Dim strPerson As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
   
      Case 0
      
            Printer.Orientation = 2
            DoEvents
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            Call PrintData
            Me.Enabled = True
            Screen.MousePointer = vbDefault
           
      Case 1
           Unload Me
      Case Else
      
   End Select
   
End Sub

Sub PrintData()
   Dim strSql As String, strLastPerson As String
   
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/9/30
   End If
   
   '此報表是為了在可辦前就先辦理,才要出此報表抓其資料,防止代理人那裡漏辦理
   'Modify By Sindy 2020/2/20
   '由於大陸法令規定有改,之前於六個月可提交續展申請,現在可於十二個月提交續展申請,煩請協助修改本所抓的期限規則
   'ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),6) => ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),12)
   strSql = " SELECT ST02 AS C1, SUBSTR(NVL(NVL(FA04,FA05),FA06),1,15) AS C2, CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C3" & _
      " , SUBSTR(NVL(NVL(TM05,TM06),TM07),1,17) AS C4, (SUBSTR(CP27,1,4)-1911) || '/' || SUBSTR(CP27,5,2) || '/' || SUBSTR(CP27,7,2) AS C5" & _
      " , TM15 AS C6, SUBSTR(NVL(NVL(CU04,CU05),CU06),1,15) AS C7" & _
      " , SUBSTR(TM21,1,4) || '/' || SUBSTR(TM21,5,2) || '/' || SUBSTR(TM21,7,2) || ' - '" & _
      " || SUBSTR(TM22,1,4) || '/' || SUBSTR(TM22,5,2) || '/' || SUBSTR(TM22,7,2) AS C8" & _
      " From CASEPROGRESS, Trademark, CUSTOMER, STAFF, FAGENT" & _
      " WHERE CP47 IS NULL AND CP07>'" & txt1(5).Text & "' AND CP07<'" & txt1(6).Text & "'" & _
      " AND TO_DATE(CP07,'YYYYMMDD')>ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),12)" & _
      " AND CP01 IN ('T','TF') AND CP10='102'" & _
      " AND TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04 AND TM10<>'000'" & _
      " AND CU01=SUBSTR(TM23,1,8) AND CU02=SUBSTR(TM23,9,1)" & _
      " AND ST01(+)=CP14" & _
      " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9,1)" & _
      " ORDER BY CP14,CP44,CP01,CP02,CP03,CP04"
   
   CheckOC
   Page = 1
   strPerson = ""
   strLastPerson = ""
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
           .MoveFirst
           strPerson = .Fields(0)
           PrintTitle
           Do While .EOF = False
               strPerson = .Fields(0)
               If (strLastPerson <> "" And strPerson <> strLastPerson) Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
               End If
               
               For i = 1 To 7
                   strTemp(i + 1) = CheckStr(.Fields(i))
               Next i
               PrintDatil
               
               If (iPrint >= 10000) Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
               End If
               strLastPerson = strPerson
               .MoveNext
           Loop
       'add by nick 2005/01/03
       Else
            InsertQueryLog (0) 'Add By Sindy 2010/9/30
            ShowNoData
            Exit Sub
       End If
   End With
   PrintEnd
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()

   Call GetPleft
   
   iPrint = 500
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "對外案件延展未提申明細表"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   
   Printer.CurrentX = 6200
   Printer.CurrentY = iPrint
   Printer.Print "專用期止日：" & Format(ChangeWStringToWDateString(txt1(5)) & " ", "@@@@@@@@") & "－" & ChangeWStringToWDateString(txt1(6))
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "承辦人：" & strPerson
   
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   
   Printer.Font.Size = 10
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "代理人"
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "發文日期"
   
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "審定號"
   
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "客戶名稱"
   
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "專用期間"
   iPrint = iPrint + 300
   
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10

End Sub

Sub PrintDatil()

   For i = 2 To 8
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   iPrint = iPrint + 300
   
End Sub

Sub GetPleft()

   Erase PLeft
   
   PLeft(0) = 500
   PLeft(1) = 1500
   
   PLeft(2) = 500
   PLeft(3) = 3500
   PLeft(4) = 5000
   PLeft(5) = 8400
   PLeft(6) = 9400
   PLeft(7) = 10400
   PLeft(8) = 13400
   
End Sub

Sub PrintEnd()

End Sub

Private Sub Form_Load()

   MoveFormToCenter Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frm020310 = Nothing
   
End Sub

Private Sub txt1_GotFocus(Index As Integer)

   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If (KeyCode = 13) Then cmdok(0).SetFocus
   
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)

   If (KeyAscii <> 8) Then
      If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
   End If

End Sub
