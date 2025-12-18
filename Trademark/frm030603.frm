VERSION 5.00
Begin VB.Form frm030603 
   BorderStyle     =   1  '單線固定
   Caption         =   "更新審定號作業"
   ClientHeight    =   1425
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4785
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2820
      TabIndex        =   1
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3780
      TabIndex        =   2
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTMBM07 
      Height          =   264
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期 :"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
End
Attribute VB_Name = "frm030603"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/10 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

'Add By Cheng 2003/05/16
Dim PLeft(0 To 5) As Integer
Public bolNotShowMsg As Boolean


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Public Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nAffect As Long
   Dim nResponse
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 執行作業
      'Modify By Sindy 2018/12/17 改成共用函數
      'nAffect = Process
      nAffect = frm030603_Process(textTMBM07)
      '2018/12/17 END
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      If nAffect <= 0 Then
         strTit = "檢核資料"
         strMsg = "沒有符合條件的資料可更新!"
         If bolNotShowMsg = False Then 'Add By Sindy 2011/11/28
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
      Else
         strTit = "檢核資料"
         strMsg = "此公報資料已更新完畢!"
         If bolNotShowMsg = False Then 'Add By Sindy 2011/11/28
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
        'Add By Cheng 2003/05/16
        '列印商標已公告缺公告日案件清單
        '93.8.23 取消列印此清單, 因92/11/28修法後核准繳第一期註冊費後才公告,故一定沒有公告日
        'PrintData
        '93.8.23 END
         textTMBM07 = Empty
      End If
      'textTMBM07.SetFocus
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030603 = Nothing
End Sub

' 公報卷期
Private Sub textTMBM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTMBM07) = False Then
      If IsNumeric(textTMBM07) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公報卷期只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_GotFocus
      End If
   End If
End Sub

'Modify By Sindy 2018/12/17 Mark : 改成共用函數(frm030603_Process)
'Private Function Process() As Long
'Dim strSql As String
'Dim strTemp As String
'Dim rsTmp As New ADODB.Recordset
'Dim nAffect As Long
'Dim nCount As Long
'Dim nTotal As Long
''Add By Cheng 2003/05/16
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
''Add By Cheng 2003/05/16
''刪除商標已公告缺公告日暫存檔資料
'StrSQLa = "Delete From R030603 Where ID='" & strUserNum & "' "
'cnnConnection.Execute StrSQLa
' '911107 nick transation
'On Error GoTo CheckingErr
'cnnConnection.BeginTrans
'
'   nAffect = 0
'   Process = 0
'    'Add By Cheng 2003/05/16
'    '申請國家為台灣者
'    'Modify By Cheng 2003/06/24
''    strSQLA = "Select TM01, TM02, TM03, TM04, TM12, TM15,'" & strUserNum & "' From Trademark " & _
''                    "WHERE TM12 IN (SELECT TMBM04 AS TM12 FROM TMBULLETIN " & _
''                    "WHERE TMBM07 = '" & textTMBM07 & "') AND " & _
''                    "(TM01 = 'T' OR TM01 = 'FCT') And TM14 Is Null And TM10 < '010' "
'    StrSQLa = "Select TM01, TM02, TM03, TM04, TM12, TMBM01,'" & strUserNum & "' From Trademark, TMBULLETIN " & _
'                    "WHERE TM12=TMBM04 And " & _
'                    " TMBM07 = '" & textTMBM07 & "' AND " & _
'                    "(TM01 = 'T' OR TM01 = 'FCT') And TM14 Is Null And TM10 < '010' "
'    rsA.CursorLocation = adUseClient
'    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        While Not rsA.EOF
'            StrSQLa = "Insert Into R030603 Values ('" & rsA.Fields(0).Value & "','" & rsA.Fields(1).Value & "','" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "','" & rsA.Fields(4).Value & "','" & rsA.Fields(5).Value & "','" & rsA.Fields(6).Value & "' )"
'            cnnConnection.Execute StrSQLa
'            rsA.MoveNext
'        Wend
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
'   ' 以公報卷期找商標公報檔的公報卷期, 在所找到的記錄中以申請案號找尋商標基本檔的申請案號欄且申請國家必須為台灣, 系統別必須為T或FCT, 若有找到則更新商標基本檔的審定號為商標公報檔的審定號
'   strSql = "UPDATE TRADEMARK SET TM15 = (SELECT TMBM01 FROM TMBULLETIN " & _
'                                         "WHERE TMBM07 = '" & textTMBM07 & "' AND " & _
'                                               "TMBM04 = TM12) " & _
'            "WHERE TM12 IN (SELECT TMBM04 AS TM12 FROM TMBULLETIN " & _
'                           "WHERE TMBM07 = '" & textTMBM07 & "') AND " & _
'                  "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' "
'   cnnConnection.Execute strSql, nAffect
'
'   Process = nAffect
'   nTotal = 0
'   nCount = 0
'   ' 以公報卷期找商標公報檔的公報卷期, 在所找到的記錄中以申請案號找尋商標基本檔的正商標號數且申請國家必須為台灣, 系統別必須為T或FCT, 若有找到則更新商標基本檔的正商標號數為商標公報檔的審定號
'   strSql = "SELECT TMBM04 FROM TMBULLETIN " & _
'            "WHERE TMBM07 = '" & textTMBM07 & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   strTemp = Empty
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Dim nIndex As Integer
'      nIndex = 0
'      Do While rsTmp.EOF = False
'         If IsNull(rsTmp.Fields("TMBM04")) = False Then
'            If IsEmptyText(rsTmp.Fields("TMBM04")) = False Then
'               If IsEmptyText(strTemp) = False Then: strTemp = strTemp & ","
'               strTemp = strTemp & "'" & rsTmp.Fields("TMBM04") & "'"
'               nIndex = nIndex + 1
'               ' 90.10.09 modify by louis (串列數目不可超過254)
'               If nIndex > 250 Then
'                  strSql = "UPDATE TRADEMARK SET TM27 = (SELECT TMBM01 FROM TMBULLETIN " & _
'                                            "WHERE TMBM07 = '" & textTMBM07 & "' AND " & _
'                                                  "TMBM04 = TM27) " & _
'                           "WHERE TM27 IN (" & strTemp & ") AND " & _
'                                 "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' "
'                  cnnConnection.Execute strSql, nCount
'
'                  nTotal = nTotal + nCount
'                  strTemp = Empty
'                  nIndex = 0
'               End If
'            End If
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'
'   If IsEmptyText(strTemp) = False Then
'      strSql = "UPDATE TRADEMARK SET TM27 = (SELECT TMBM01 FROM TMBULLETIN " & _
'                                            "WHERE TMBM07 = '" & textTMBM07 & "' AND " & _
'                                                  "TMBM04 = TM27) " & _
'               "WHERE TM27 IN (" & strTemp & ") AND " & _
'                     "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' "
'      cnnConnection.Execute strSql, nCount
'      nTotal = nTotal + nCount
'      If Process = 0 Then
'         'Process = nAffect
'         Process = nTotal
'      End If
'   End If
'   Set rsTmp = Nothing
'
'   'strSQL = "UPDATE TRADEMARK SET TM27 = (SELECT TMBM01 FROM TMBULLETIN " & _
'   '                                      "WHERE TMBM07 = '" & textTMBM07 & "' AND " & _
'   '                                            "TMBM04 = TM27) " & _
'   '         "WHERE TM27 IN (SELECT TMBM04 AS TM27 FROM TMBULLETIN " & _
'   '                        "WHERE TMBM07 = '" & textTMBM07 & "') AND " & _
'   '               "(TM01 = 'T' OR TM01 = 'FCT') AND " & _
'   '               "TM10 < '010' "
'   'cnnConnection.Execute strSQL
'
' '911107 nick transation
'  cnnConnection.CommitTrans
'     Exit Function
'CheckingErr:
'    MsgBox (Err.Description)
'     cnnConnection.RollbackTrans
'End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   ' 審定號不可空白
   If IsEmptyText(textTMBM07) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入審定號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM07.SetFocus
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTMBM07_GotFocus()
   InverseTextBox textTMBM07
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textTMBM07.Enabled = True Then
   Cancel = False
   textTMBM07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

'Add By Cheng 2003/05/16
Private Sub PrintTitle(Page As Integer)
'Page : 頁數
Dim i As Integer
  
i = 500
If Page = 1 Then Printer.Orientation = vbPRORPortrait
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 2800
Printer.CurrentY = i
Printer.Print "商標已公告缺公告日案件清單"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 800
Printer.Print "列印人 : " & strUserName
Printer.CurrentX = PLeft(3) - 1200
Printer.CurrentY = i + 800
Printer.Print "公報卷期 : " & Me.textTMBM07.Text
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
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = i + 1700
Printer.Print "申請案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = i + 1700
Printer.Print "審定號"
Printer.CurrentX = PLeft(3) - 300
Printer.CurrentY = i + 1700
Printer.Print "｜"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = i + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = i + 1700
Printer.Print "申請案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = i + 1700
Printer.Print "審定號"

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 2000
Printer.Print String(250, "-")

End Sub

'Add By Cheng 2003/05/16
Private Sub GetPrintLeft()
PLeft(0) = 200
PLeft(1) = 2000
PLeft(2) = 4000 - 300

PLeft(3) = 6200 - 300
PLeft(4) = 8000 - 300
PLeft(5) = 10000 - 600
End Sub

'Add By Cheng 2003/05/16
Private Sub PrintData()
Dim rs As New ADODB.Recordset
Dim intPage As Integer
Dim strDate As String
Dim strNation As String
Dim ii As Integer
Dim jj As Integer
Dim arrJJ
Dim intMaxJJ As Integer
Dim kk As Integer
Dim arrKK
Dim intMaxKK As Integer
Dim Prn As Printer
Dim iPrint As Integer
Dim iPrint1 As Integer
Dim strDeadLineCon As String
Dim strDLCon As String

strSql = "Select * From R030603 Where ID='" & strUserNum & "' Order By 1,2,3,4 "
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
rs.CursorLocation = adUseClient
rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
   intPage = 1
   GetPrintLeft
   PrintTitle intPage
   ii = 0
   iPrint = 2700
   iPrint1 = 2700
   rs.MoveFirst
   While Not rs.EOF
      If ii >= 40 Then
         intPage = intPage + 1
         Printer.NewPage
         PrintTitle intPage
         ii = 0
         iPrint = 2700
         iPrint1 = 2700
      End If
      '列印左半邊
      If ii < 20 Then
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "" & rs.Fields(0).Value & "-" & rs.Fields(1).Value & "-" & rs.Fields(2).Value & "-" & rs.Fields(3).Value
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "" & rs.Fields(4).Value
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "" & rs.Fields(5).Value
         Printer.CurrentX = PLeft(3) - 300
         Printer.CurrentY = iPrint
         Printer.Print "｜"
         iPrint = iPrint + 300
         
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print String(250, "-")
         iPrint = iPrint + 300
      '列印右半邊
      Else
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "" & rs.Fields(0).Value & "-" & rs.Fields(1).Value & "-" & rs.Fields(2).Value & "-" & rs.Fields(3).Value
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "" & rs.Fields(4).Value
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "" & rs.Fields(5).Value
         iPrint1 = iPrint1 + 300
         iPrint1 = iPrint1 + 300
      End If
      rs.MoveNext
      ii = ii + 1
   Wend
   Printer.EndDoc
    ShowPrintOk
End If
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing

End Sub

