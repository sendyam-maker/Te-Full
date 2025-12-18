VERSION 5.00
Begin VB.Form frm030615 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報資料檢核表"
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
      Caption         =   "公報卷期："
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
End
Attribute VB_Name = "frm030615"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

'Add By Cheng 2003/05/16
Dim PLeft(0 To 5) As Integer
Dim m_strKind As String '報表種類

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub cmdok_Click()
    If CheckDataValid() = False Then Exit Sub
    '重新檢查欄位有效性
    If TxtValidate = False Then Exit Sub
    '列印國內公報資料檢核表
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/21 清除查詢印表記錄檔欄位
    PrintData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030615 = Nothing
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
        ElseIf Val(Right(Me.textTMBM07.Text, 2)) < 1 Or Val(Right(Me.textTMBM07.Text, 2)) > 24 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "公報期數輸入錯誤!!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTMBM07_GotFocus
        End If
    End If
End Sub

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
Printer.CurrentX = PLeft(2)
Printer.CurrentY = i
Printer.Print "商標公報資料檢核表"
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
Printer.Print "列印日期 : " & ChangeTStringToTDateString(strSrvDate(2))

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1100
Select Case m_strKind
Case "1"
    Printer.Print "※未發證（商標基本檔無專用期限者）※"
Case "2"
    Printer.Print "※商標基本檔之註冊公告日為當期但該期公報中無該筆資料※"
Case "3"
    Printer.Print "※商標基本檔之註冊號與公報資料不同※"
Case Else
    Printer.Print "※商標基本檔之商品類別與公報資料不同※"
End Select
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
Dim strTMBM07 As String

strTMBM07 = ChgTMBM07ToDate(Me.textTMBM07.Text)
If Len(textTMBM07) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07 'Add By Sindy 2010/10/21
End If
'未發證資料
strSql = "Select TM01, TM02, TM03, TM04, TMBM01, TMBM04, '1' From Trademark, Tmbulletin Where TM10='000' And TM12=TMBM04 And (TM21 Is Null Or TM22 Is Null) And TMBM07='" & Me.textTMBM07.Text & "' "
'註冊公告日為當期但該期公報中無資料
strSql = strSql & " Union Select TM01, TM02, TM03, TM04, TM15, TM12, '2' From Trademark, Tmbulletin Where TM10='000' And TM12=TMBM04(+) And TM14=" & Val(strTMBM07) & " And TMBM04 Is Null "
'商標基本檔的註冊號與公報資料不同
strSql = strSql & " Union Select TM01, TM02, TM03, TM04, TM15, TM12, '3' From Trademark, Tmbulletin Where TM10='000' And TM12=TMBM04(+) And TM14=" & Val(strTMBM07) & " And TMBM01 Is Null "
'商標基本檔的商品類別與公報資料不同
strSql = strSql & " Union Select TM01, TM02, TM03, TM04, TM15, TM12, '4' From Trademark, Tmbulletin Where TM10='000' And TM12=TMBM04 And TMBM07='" & Me.textTMBM07.Text & "' And TM09<>TMBM08 And (TM09<>'7' And TM09<>'8')  "
strSql = strSql & " Order By 7, 1, 2, 3, 4 "
rs.CursorLocation = adUseClient
rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
   InsertQueryLog (rs.RecordCount) 'Add By Sindy 2010/10/21
   m_strKind = "" & rs.Fields(6).Value
   intPage = 1
   GetPrintLeft
   PrintTitle intPage
   ii = 0
   iPrint = 2700
   iPrint1 = 2700
   rs.MoveFirst
   While Not rs.EOF
        If m_strKind <> "" & rs.Fields(6).Value Then
            m_strKind = "" & rs.Fields(6).Value
            intPage = intPage + 1
            Printer.NewPage
            PrintTitle intPage
            ii = 0
            iPrint = 2700
            iPrint1 = 2700
        End If
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
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint1
         Printer.Print "" & rs.Fields(0).Value & "-" & rs.Fields(1).Value & "-" & rs.Fields(2).Value & "-" & rs.Fields(3).Value
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint1
         Printer.Print "" & rs.Fields(4).Value
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint1
         Printer.Print "" & rs.Fields(5).Value
         iPrint1 = iPrint1 + 300
         iPrint1 = iPrint1 + 300
      End If
      rs.MoveNext
      ii = ii + 1
   Wend
   Printer.EndDoc
    ShowPrintOk
Else
    InsertQueryLog (0) 'Add By Sindy 2010/10/21
    ShowNoData
End If
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
End Sub

'將公報卷期轉換為日期
Private Function ChgTMBM07ToDate(strTMBM07 As String)
Dim strYY As String
Dim strMM As String
Dim strDD As String
'920101 : 3001, 920116 : 3002 ...(每年會有24期)

strYY = (Val(Mid(strTMBM07, 1, Len(strTMBM07) - 2)) + 62)
strMM = Format(Right(strTMBM07, 2) / 2, "00")
If Right(strTMBM07, 2) Mod 2 <> 0 Then
    strDD = "01"
Else
    strDD = "16"
End If
ChgTMBM07ToDate = DBDATE(strYY & strMM & strDD)
End Function
