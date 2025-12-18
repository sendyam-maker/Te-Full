VERSION 5.00
Begin VB.Form frm084003 
   BorderStyle     =   1  '單線固定
   Caption         =   "法務案件年度統計表"
   ClientHeight    =   2160
   ClientLeft      =   2610
   ClientTop       =   3165
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4815
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   3420
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1500
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3792
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2964
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.Label lblNote 
      Caption         =   "此報表尚需法律所重新調整！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   60
      TabIndex        =   9
      Top             =   150
      Width           =   2925
   End
   Begin VB.Label lblSysKind 
      Caption         =   "(1.法務 2.顧問 3.全部)"
      Height          =   180
      Left            =   2220
      TabIndex        =   8
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別︰"
      Height          =   180
      Index           =   2
      Left            =   420
      TabIndex        =   7
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第一年度："
      Height          =   180
      Index           =   1
      Left            =   420
      TabIndex        =   6
      Top             =   864
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第二年度："
      Height          =   180
      Index           =   0
      Left            =   2328
      TabIndex        =   5
      Top             =   864
      Width           =   900
   End
End
Attribute VB_Name = "frm084003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 5) As Integer

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   If ChkRange(Text1(0), Text1(1), "年度") = False Then Exit Sub
   If Text1(2) = "" Then
      MsgBox "系統別不得為空值 !", vbCritical
      Text1(2).SetFocus
      Exit Sub
   End If
   Screen.MousePointer = 11
   GetPrintLeft
   PrintCase
   Screen.MousePointer = 0
   MsgBox "列印結束!", vbInformation
End Sub

Private Sub InsertDb(ByVal SituS As String)
 Dim StS(1 To 11) As String, StRng(0 To 9) As String
 Dim stTmp(1 To 5) As String, StN(1 To 4) As String
 Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset
 Dim Qty As String, i As Integer, j As Integer
 Dim strSql As String
 Dim strCP01 As String
 Dim strCP02 As String
 Dim strCP03 As String
 Dim strCP04 As String
 Dim nCount As Integer
 Dim rs As New ADODB.Recordset
 Dim m_Type As Integer
 
   Set Wo = DBEngine.Workspaces(0)
   Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   Qty = "DELETE FROM TEMP"
   Db.Execute Qty
   If RsTemp.State = adStateOpen Then RsTemp.Close
   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES " & _
      "('案件','專利','商標','著作權','其他','總計')"
   Db.Execute Qty
   StS(1) = "刑事":  StS(2) = "民事": StS(3) = "強制執行"
   StS(4) = "行政訴訟": StS(5) = "雜文": StS(6) = "總件數"
   StRng(1) = " CP10 IN ('2101','2102','2103','2104','2105','2106','2108'," & _
      "'2121','2122','2123','2124','2201','2202','2203','2204','2205','2206'," & _
      "'2207','2109','2137','2139','2140','2221','2222')"
   
   StRng(2) = " CP10 IN ('1101','1102','1103','1104','1105','1111','1112','1113','1114','1115','1121','1122','1123','1124','1131','1132','1133','1134')"
   
   StRng(3) = " CP10 IN ('1301','1311','1312','1313','1314')"
   
   StRng(4) = " CP10 LIKE '54%'"
   
   StRng(5) = " ((CP10 NOT IN ('2101','2102','2103','2104','2105','2106','2108'," & _
      "'2121','2122','2123','2124','2201','2202','2203','2204','2205','2206'," & _
      "'2207','2109','2137','2139','2140','2221','2222'," & _
      "'1101','1102','1103','1104','1105','1111','1112','1113','1114','1115','1121','1122','1123','1124','1131','1132','1133','1134'," & _
      "'1301','1311','1312','1313','1314')) AND (CP10 NOT LIKE '54%'))"
      
   StN(1) = " AND CR05 IN ('P','FCP','CFP'))"
   StN(2) = " AND CR05 IN ('T','FCT','CFT','TF'))"
   StN(3) = " AND CR05 IN ('TC','CFC'))"
   StN(4) = " AND NOT (CR05 IN ('P','FCP','CFP','T','FCT','CFT','TF','TC','CFC')))"

   For i = 1 To 5
       strCP01 = ""
       strCP02 = ""
       strCP03 = ""
       strCP04 = ""
      For j = 1 To 5
         stTmp(j) = "0"
      Next
      'Modify By Cheng 2002/03/26
      '多加CP09<'C'控制
'      strSQL = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE" & StRng(i) & SituS & GetSQL
      strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP09<'C' AND " & StRng(i) & SituS & GetSql
      rs.Open strSql, cnnConnection
      If rs.EOF = False Then
         Do While rs.EOF = False
            m_Type = 0
            If Not IsNull(rs.Fields("CP01")) Then
               strCP01 = rs.Fields("CP01")
            Else
               strCP01 = ""
            End If
            If Not IsNull(rs.Fields("CP02")) Then
               strCP02 = rs.Fields("CP02")
            Else
               strCP02 = ""
            End If
            If Not IsNull(rs.Fields("CP03")) Then
               strCP03 = rs.Fields("CP03")
            Else
               strCP03 = ""
            End If
            If Not IsNull(rs.Fields("CP04")) Then
               strCP04 = rs.Fields("CP04")
            Else
               strCP04 = ""
            End If
         
            For j = 1 To 3
               'Modified by Lydia 2020/04/10 CASERELATION=>CASERELATION1
               strExc(2) = " select count(*) from (SELECT DISTINCt cr01,cr02,cr03,cr04 FROM CASERELATION1 WHERE " & _
                            " CR01 ='" & strCP01 & "'" & _
                            " AND CR02 ='" & strCP02 & "'" & _
                            " AND CR03 ='" & strCP03 & "'" & _
                            " AND CR04 ='" & strCP04 & "'" & StN(j)
               RsTemp.Open strExc(2), cnnConnection
               If RsTemp.EOF = False Then
                  If Val(RsTemp.Fields(0)) <> 0 Then
                     m_Type = 1
                     stTmp(j) = Format(Val(stTmp(j)) + Val(RsTemp.Fields(0)))
                     stTmp(5) = Format(Val(stTmp(5)) + Val(RsTemp.Fields(0)))
                  End If
               End If
               RsTemp.Close
            Next
            If m_Type <> 1 Then
               stTmp(4) = Val(stTmp(4)) + 1
               stTmp(5) = Val(stTmp(5)) + 1
            End If
            rs.MoveNext
            nCount = nCount + 1
          Loop
          rs.Close
       Else
           For j = 1 To 4
               stTmp(j) = Format(Val(stTmp(j)) + 0)
               stTmp(5) = Format(Val(stTmp(5)) + 0)
           Next
           rs.Close
       End If
'       If nCount <> StTmp(5) Then
'          StTmp(4) = StTmp(4) + CInt(nCount) - CInt(StTmp(5))
'          StTmp(5) = StTmp(5) + CInt(nCount) - CInt(StTmp(5))
'       End If
      Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) VALUES ('" & _
         StS(i) & "','" & stTmp(1) & "','" & stTmp(2) & "','" & stTmp(3) & "','" & _
         stTmp(4) & "','" & stTmp(5) & "')"
      Db.Execute Qty
   Next
   Qty = "INSERT INTO TEMP (TMP01,TMP02,TMP03,TMP04,TMP05,TMP06) SELECT '" & _
      StS(6) & "',SUM(VAL(TMP02)),SUM(VAL(TMP03)),SUM(VAL(TMP04)),SUM(VAL(TMP05))," & _
      "SUM(VAL(TMP06)) FROM TEMP WHERE TMP01<>'案件'"
   Db.Execute Qty
End Sub

Private Sub PrintCase()
 Dim i As Integer, iPrint As Integer
 Dim Wo As DAO.Workspace, Db As DAO.Database, Rc As DAO.Recordset
 Dim Qty As String
 Dim nYear As Integer
 Dim nYear1 As Integer
 
 nYear = CInt(Text1(0)) + 1911
 nYear1 = CInt(Text1(1)) + 1911
 
On Error GoTo ErrHand

   If CreateDatabase = False Then
      MsgBox "無法建立暫存區，列印失敗 !", vbInformation
      Exit Sub
   End If
   'InsertDb " AND SUBSTR(CP27,1,4) ='" & CStr(CInt(Text1(0)) + 1911) + "'"
   InsertDb " AND CP05 BETWEEN " & nYear & "0101 and " & nYear & "1231"
   i = 500
   Printer.Orientation = vbPRORPortrait
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "法務案件年度統計表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 500
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 9000:             Printer.CurrentY = i + 500
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate))
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print Text1(0) & " 年"
   Printer.CurrentX = 9000:             Printer.CurrentY = i + 800
   Printer.Print "頁次 : 1"
   iPrint = i + 1100
   Printer.CurrentX = 500:              Printer.CurrentY = iPrint
   Printer.Print String(180, "-")
   iPrint = iPrint + 300
   Set Wo = DBEngine.Workspaces(0)
   Set Db = Wo.OpenDatabase(App.path & "\Case.mdb", False, False, ";PWD=taie")
   Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06 FROM TEMP"
   Set Rc = Db.OpenRecordset(Qty)
   With Rc
      Do While Not .EOF
         For i = 0 To 5
             If i = 0 Then
                Printer.CurrentX = PLeft(i)
                Printer.CurrentY = iPrint
                Printer.Print CheckStr(.Fields(i))
             Else
                Printer.CurrentX = CInt(PLeft(i)) + 1000 - (Printer.TextWidth(CheckStr(.Fields(i))))
                Printer.CurrentY = iPrint
                Printer.Print CheckStr(.Fields(i))
             End If

         Next
         iPrint = iPrint + 300
         .MoveNext
      Loop
   End With
   Printer.CurrentX = 500:          Printer.CurrentY = iPrint
   Printer.Print String(180, "-")
   InsertDb " AND CP05 BETWEEN " & nYear1 & "0101 and " & nYear1 & "1231"
   'InsertDb " AND SUBSTR(CP27,1,4) ='" & CStr(CInt(Text1(1)) + 1911) + "'"
   iPrint = iPrint + 500
   Printer.CurrentX = 500:             Printer.CurrentY = iPrint
   Printer.Print Text1(1) & " 年"
   iPrint = iPrint + 300
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   Printer.Print String(180, "-")
   iPrint = iPrint + 300
   Qty = "SELECT TMP01,TMP02,TMP03,TMP04,TMP05,TMP06 FROM TEMP"
   Set Rc = Db.OpenRecordset(Qty)
   With Rc
      Do While Not .EOF
         For i = 0 To 5
             If i = 0 Then
                 Printer.CurrentX = PLeft(i)
                 Printer.CurrentY = iPrint
                 Printer.Print .Fields(i)
             Else
                 Printer.CurrentX = CInt(PLeft(i)) + 1000 - (Printer.TextWidth(CheckStr(.Fields(i))))
                 Printer.CurrentY = iPrint
                 Printer.Print .Fields(i)
             End If
         Next
         iPrint = iPrint + 300
         .MoveNext
      Loop
   End With
   Printer.CurrentX = 500:          Printer.CurrentY = iPrint
   Printer.Print String(180, "-")
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 500:     PLeft(1) = 1900
   PLeft(2) = 3100:    PLeft(3) = 4300
   PLeft(4) = 5800:    PLeft(5) = 7100
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub
Private Sub Form_Paint()
  If Me.Tag = 1 Then
     lblSysKind.Caption = "(1.FCL 2.CFL 3.全部)"
  ElseIf Me.Tag = 0 Then
     lblSysKind.Caption = "(1.法務 2.顧問 3.全部)"
  End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 2
         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Function GetSql() As String
   If Me.Tag = 0 Then
      Select Case Text1(2)
         Case "1"
            strExc(1) = " AND CP01='L'"
         Case "2"
            strExc(1) = " AND CP01='LA'"
         Case "3"
            strExc(1) = " AND CP01 IN ('L','LA')"
      End Select
   Else
      Select Case Text1(2)
         Case "1"
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            strExc(1) = " AND CP01 in ('FCL','LIN')"
         Case "2"
            strExc(1) = " AND CP01='CFL'"
         Case "3"
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            strExc(1) = " AND CP01 IN ('FCL','CFL','LIN')"
      End Select
   End If
   GetSql = strExc(1) & " AND CP26 IS NULL AND CP57 IS NULL"
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm084003 = Nothing
End Sub
