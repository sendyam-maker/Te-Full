VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm12040123 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件自動核准作業"
   ClientHeight    =   2400
   ClientLeft      =   2490
   ClientTop       =   2505
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4635
   Begin VB.TextBox textPercent 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   3048
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   1068
   End
   Begin MSComctlLib.ProgressBar proBar 
      Height          =   252
      Left            =   384
      TabIndex        =   4
      Top             =   1320
      Width           =   3732
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3588
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox textSys 
      Height          =   264
      Left            =   1584
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "PS：T大陸授權案18個月後, TF申請一年後"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別 ："
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frm12040123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

' 系統別
Dim m_Sys As String

Private Sub Form_Load()
   textPercent = Empty
   textPercent.BackColor = &H8000000F
   MoveFormToCenter Me
   
   textSys = "T,TF"
   textSys.Locked = True
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bTData As Boolean
   Dim bTFData As Boolean
   
   bTData = False
   bTFData = False
   
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      
      ' 執行作業
      bTData = OnSaveData("T")
      bTFData = OnSaveData("TF")
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      If bTData = False And bTFData = False Then
         strTit = "搜尋資料"
         strMsg = "資料庫中沒有符合的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         strTit = "自動核准"
         strMsg = "資料處理完畢"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Function OnSaveData(ByVal strSys As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   OnSaveData = False
   
   m_Sys = strSys
   Select Case m_Sys
      Case "T":
         '2007/10/24 MODIFY BY SONIA 加判斷申請國家為大陸,原程式未判斷,發文日改判斷新系統上線日
         strSql = "SELECT * FROM CASEPROGRESS, TRADEMARK " & _
                  "WHERE CP01 = '" & m_Sys & "' AND " & _
                        "CP10 = '502' AND " & _
                        "CP47 IS NOT NULL AND " & _
                        "CP01 = TM01(+) AND " & _
                        "CP02 = TM02(+) AND " & _
                        "CP03 = TM03(+) AND " & _
                        "CP04 = TM04(+) AND '020'=TM10(+) AND " & _
                        "CP27>=20030201 AND " & _
                        "(CP24 IS NULL OR CP24 = '' OR CP24 = ' ') AND " & _
                        "(TO_DATE('" & DBDATE(SystemDate()) & "','YYYYMMDD') - TO_DATE(CP47,'YYYYMMDD')) / 30 > 18 "
                        'SystemDate() & " - CP47 > 1800 "
      Case "TF":
         '2007/10/24 MODIFY BY SONIA 改判斷母案基本檔之目前准駁及申請日
'         strSQL = "SELECT * FROM CASEPROGRESS, TRADEMARK " & _
'                  "WHERE CP01 = '" & m_Sys & "' AND " & _
'                        "CP10 = '101' AND " & _
'                        "CP03 <> '0' AND " & _
'                        "CP04 <> '00' AND " & _
'                        "CP01 = TM01(+) AND " & _
'                        "CP02 = TM02(+) AND " & _
'                        "CP03 = TM03(+) AND " & _
'                        "CP04 = TM04(+) AND " & _
'                        "CP27 IS NOT NULL AND " & _
'                        "(CP24 IS NULL OR CP24 = '' OR CP24 = ' ') AND " & _
'                        "(TO_DATE('" & DBDATE(SystemDate()) & "','YYYYMMDD') - TO_DATE(TM11,'YYYYMMDD')) / 30 > 24 "
'                        'SystemDate() & " - TM11 > 2400 "
         strSql = "SELECT * FROM CASEPROGRESS, TRADEMARK " & _
                  "WHERE CP01 = '" & m_Sys & "' AND " & _
                        "CP10 = '101' AND " & _
                        "CP03 <> '0' AND " & _
                        "CP04 <> '00' AND " & _
                        "CP01 = TM01(+) AND " & _
                        "CP02 = TM02(+) AND " & _
                        "CP03 = TM03(+) AND " & _
                        "CP04 = '00' AND " & _
                        "CP27 IS NOT NULL AND " & _
                        "(TM16 IS NULL OR TM16 = '' OR TM16 = ' ') AND " & _
                        "(TO_DATE('" & DBDATE(SystemDate()) & "','YYYYMMDD') - TO_DATE(TM11,'YYYYMMDD')) / 30 > 24 "
         '2007/10/24 END
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      textPercent = Empty
      proBar.max = rsTmp.RecordCount
      proBar.Min = 0
      proBar.Value = 0
      textPercent = "0/" & CStr(rsTmp.RecordCount)
      UpdateDB rsTmp
      textPercent = Empty
      proBar.Min = 0
      proBar.Value = 0
      
      OnSaveData = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub UpdateDB(ByRef rsTmp As ADODB.Recordset)
   Dim strSql As String
   Dim strCP05 As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP13 As String
   Dim strCP27 As String
   Dim strCP43 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP10 As String
   Dim strNP14 As String
   Dim strNP15 As String
   Dim strNP22 As String
   Dim nRow As Long
      
   rsTmp.MoveFirst
   nRow = 0
   Do While rsTmp.EOF <> True
      '''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新案件進度檔的實際結果為准, 准駁日為系統日
      strSql = "UPDATE CASEPROGRESS SET CP24 = '1', " & _
                                       "CP25 = " & DBDATE(SystemDate()) & " " & _
               "WHERE CP01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                     "CP02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                     "CP03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                     "CP04 = '" & rsTmp.Fields("CP04") & "' AND " & _
                     "CP09 = '" & rsTmp.Fields("CP09") & "' "
      cnnConnection.Execute strSql
      '''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新商標基本檔的目前准駁為准, 審定來函日為系統日
      If IsNull(rsTmp.Fields("CP01")) = False Then
         If rsTmp.Fields("CP01") = "TF" Then
            strSql = "UPDATE TRADEMARK SET TM16 = '1', " & _
                                          "TM13 = " & DBDATE(SystemDate()) & " " & _
                     "WHERE TM01 = '" & rsTmp.Fields("CP01") & "' AND " & _
                           "TM02 = '" & rsTmp.Fields("CP02") & "' AND " & _
                           "TM03 = '" & rsTmp.Fields("CP03") & "' AND " & _
                           "TM04 = '" & rsTmp.Fields("CP04") & "' "
            cnnConnection.Execute strSql
         End If
      End If
      '''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 新增資料到案件進度檔
      ' 收文號
      strCP09 = Empty
      strCP09 = AutoNo("C", 6)
      ' 收文日為系統日
      strCP05 = DBDATE(SystemDate())
      ' 案件性質為核准
      strCP10 = "1001"
      ' 業務區別
      strCP12 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         '91.9.10 MODIFY BY SONIA
         'strCP12 = GetStaffDepartment(rsTmp.Fields("CP13"))
         strCP12 = GetST15(rsTmp.Fields("CP13"))
         '91.9.10 END
      Else
         If IsNull(rsTmp.Fields("CP12")) = False Then
            strCP12 = rsTmp.Fields("CP12")
         End If
      End If
      ' 智權人員
      strCP13 = "NULL"
      If IsNull(rsTmp.Fields("CP13")) = False Then
         If IsEmptyText(rsTmp.Fields("CP13")) = False Then: strCP13 = rsTmp.Fields("CP13")
      End If
      ' 發文日
      strCP27 = DBDATE(SystemDate())
      ' 相關總收文號
      strCP43 = Empty
      If IsNull(rsTmp.Fields("CP09")) = False Then
         strCP43 = rsTmp.Fields("CP09")
      End If
      '2007/10/24 MODIFY BY SONIA 加CP64註明自動核准
      strSql = "INSERT INTO CASEPROGRESS " & _
                     "(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
               "VALUES " & _
                     "('" & rsTmp.Fields("CP01") & "','" & rsTmp.Fields("CP02") & "','" & rsTmp.Fields("CP03") & "','" & rsTmp.Fields("CP04") & "'," & _
                     strCP05 & ",'" & strCP09 & "','" & strCP10 & "'," & _
                     "'" & strCP12 & "','" & strCP13 & "','" & strUserNum & "'," & _
                     "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & strCP43 & "','自動核准')"
      cnnConnection.Execute strSql
      '''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新下一程序檔中案件性質為催審的資料, 將是否續辦欄位改為Y
      strSql = "UPDATE NEXTPROGRESS SET NP06 = 'Y' " & _
               "WHERE NP01 = '" & rsTmp.Fields("CP09") & "' AND " & _
                     "NP02 = '" & rsTmp.Fields("CP01") & "' AND " & _
                     "NP03 = '" & rsTmp.Fields("CP02") & "' AND " & _
                     "NP04 = '" & rsTmp.Fields("CP03") & "' AND " & _
                     "NP05 = '" & rsTmp.Fields("CP04") & "' AND " & _
                     "NP07 = 305 AND " & _
                     "(NP06 IS NULL OR NP06 = '' OR NP06 = ' ')"
      cnnConnection.Execute strSql
      '''''''''''''''''''''''''''''''''''''''''''''''''''
      Select Case m_Sys
         '2007/10/24 與葉經理確認仍掛終止授權期限
         Case "T":
            ' 本所期限
            strNP08 = "NULL"
            If IsNull(rsTmp.Fields("CP54")) = False Then
               If IsEmptyText(rsTmp.Fields("CP54")) = False Then: strNP08 = rsTmp.Fields("CP53")
            End If
            ' 法定期限
            strNP09 = "NULL"
            If IsNull(rsTmp.Fields("CP54")) = False Then
               If IsEmptyText(rsTmp.Fields("CP54")) = False Then: strNP09 = rsTmp.Fields("CP54")
            End If
            ' 智權人員代號
            strNP10 = "NULL"
            If IsNull(rsTmp.Fields("CP13")) = False Then
               If IsEmptyText(rsTmp.Fields("CP13")) = False Then: strNP10 = "'" & rsTmp.Fields("CP13") & "'"
            End If
            ' 相關人
            strNP14 = "NULL"
            If IsNull(rsTmp.Fields("CP50")) = False Then
               If IsEmptyText(rsTmp.Fields("CP50")) = False Then: strNP14 = rsTmp.Fields("CP50")
            End If
            If strNP14 = "NULL" And IsNull(rsTmp.Fields("CP51")) = False Then
               If IsEmptyText(rsTmp.Fields("CP51")) = False Then: strNP14 = rsTmp.Fields("CP51")
            End If
            If strNP14 = "NULL" And IsNull(rsTmp.Fields("CP52")) = False Then
               If IsEmptyText(rsTmp.Fields("CP52")) = False Then: strNP14 = rsTmp.Fields("CP52")
            End If
            If strNP14 <> "NULL" Then
               strNP14 = "'" & strNP14 & "'"
            End If
            ' 備註
            strNP15 = "'授權終止日"
            If IsNull(rsTmp.Fields("CP54")) = False Then
               If IsEmptyText(rsTmp.Fields("CP54")) = False Then: strNP15 = strNP15 & ":" & rsTmp.Fields("CP54")
            End If
            strNP15 = strNP15 & "'"
            ' 序號
            strNP22 = GetNextProgressNo()
            'Modify By Cheng 2002/09/25
'            strSQL = "INSERT INTO NEXTPROGRESS " & _
'                           "(NP01, NP02, NP03, NP04, NP05, NP07, NP08, NP09, NP10, NP14, NP15, NP22) " & _
'                     "VALUES " & _
'                           "('" & strCP09 & "','" & rsTmp.Fields("CP01") & "','" & rsTmp.Fields("CP02") & "','" & rsTmp.Fields("CP03") & "','" & rsTmp.Fields("CP04") & "'," & _
'                           "503" & "," & strNP08 & "," & strNP09 & "," & strNP10 & "," & strNP14 & "," & strNP15 & "," & strNP22 & ")"
            strSql = "INSERT INTO NEXTPROGRESS " & _
                           "(NP01, NP02, NP03, NP04, NP05, NP07, NP08, NP09, NP10, NP14, NP15, NP22) " & _
                     "VALUES " & _
                           "('" & strCP09 & "','" & rsTmp.Fields("CP01") & "','" & rsTmp.Fields("CP02") & "','" & rsTmp.Fields("CP03") & "','" & rsTmp.Fields("CP04") & "'," & _
                           "503" & "," & strNP08 & "," & strNP09 & ",'" & strNP10 & "','" & ChgSQL(strNP14) & "','" & strNP15 & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
         Case Else:
      End Select

      ' 更新ProgressBar的進度及顯示的文字
      proBar.Value = proBar.Value + 1
      nRow = nRow + 1
      textPercent = CStr(nRow) & "/" & CStr(rsTmp.RecordCount)
      
      ' 下一筆
      rsTmp.MoveNext
   Loop

End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False

   ' 系統類別
   If IsEmptyText(textSys) = True Then
      strTit = "查詢資料"
      strMsg = "請輸入系統類別"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSys.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040123 = Nothing
End Sub

Private Sub textSys_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
'Private Sub textSys_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(textSys) = False Then
'      Select Case textSys
'         Case "T", "TF":
'         Case Else:
'            Cancel = True
'            strTit = "搜尋資料"
'            strMsg = "系統類別只可為T或TF"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textSys_GotFocus
'      End Select
'   End If
'End Sub

Private Sub textSys_GotFocus()
   InverseTextBox textSys
End Sub

