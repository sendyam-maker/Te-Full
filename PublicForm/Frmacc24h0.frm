VERSION 5.00
Begin VB.Form Frmacc24h0 
   AutoRedraw      =   -1  'True
   Caption         =   "折讓單列印"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   5160
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1668
      TabIndex        =   1
      Top             =   336
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1668
      TabIndex        =   3
      Top             =   744
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1200
      Width           =   4692
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   252
      TabIndex        =   0
      Top             =   348
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   744
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   480
      TabIndex        =   5
      Top             =   756
      Width           =   1272
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   1404
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc24h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSql As String
Dim strNo As String
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Private Const intDefault As Integer = 500
Private Const intTop As Integer = 1000
Dim strNewPage As String
Dim strCurr As String
Public m_iCopy As Integer '列印份數 Add by Morgan 2009/10/15
Dim strPrintCurr As String 'Add By Sindy 2013/1/15


Public Sub Command2_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   PrintData
   If strCon10 <> MsgText(602) Then
      FormClear
   End If
   Screen.MousePointer = vbDefault
   StatusView MsgText(100)
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   If IsObject(mdiMain) Then
      ToolShow
   End If
   '93.3.16 END

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(100)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 2200
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text3.Enabled = False
   StatusView MsgText(100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24h0 = Nothing
End Sub


Private Sub Option1_Click()
   If Option1.Value Then
      Text1.Enabled = True
      Text3.Enabled = False
   End If
End Sub

Private Sub Option2_Click()
   If Option2.Value Then
      Text1.Enabled = False
      Text3.Enabled = True
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text3 = ""
   If Option1.Value Then
      Text1.SetFocus
   Else
      Text3.SetFocus
   End If
End Sub

'*************************************************
' 列印明細資料
'
'*************************************************
Private Sub PrintData()
Dim strDescription As String
'add by nickc 2007/02/08
Dim lngAmount As Long
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
   lngAmount = 0
   intLength = 0
   douAmount = 0
   douUSDollar = 0
   strNewPage = ""
   strSql = MsgText(601)
   If Option1.Value Then
      If Text1 <> MsgText(601) Then
         strSql = strSql & " and a1k01 = '" & Text1 & "'"
         pub_QL05 = pub_QL05 & ";" & Label1 & Text1 'Add By Sindy 2010/12/23
      End If
   Else
      If Text3 <> MsgText(601) Then
         strSql = strSql & " and a1201 = '" & Text3 & "'"
         pub_QL05 = pub_QL05 & ";" & Label3 & Text3 'Add By Sindy 2010/12/23
      End If
   End If
   strNo = ""
   strCon10 = ""
   
   If Me.m_iCopy > 0 Then Printer.Copies = Me.m_iCopy 'Add by Morgan 2009/10/15
   
   Printer.FontSize = 12
   adoacc1k0.CursorLocation = adUseClient
   If Option1.Value Then
      'Modify by Morgan 2007/1/24 加 FA70
      'Modify By Sindy 2011/3/7 +FA108
      'Modify By Sindy 2013/1/15 +, a1k31
'      adoacc1k0.Open "select a1k27, fa05, fa63, fa64, fa65, fa32, fa18, a1k02, fa33, fa19, fa34, fa20, a1k13, a1k14, a1k15, a1k16, fa21, fa22, fa35, a1k03, fa36, a1k01, fa06, fa23, a1k04, a1k10, fa43, a1k18 as Curr, fa70 as cu102, a1k06, a1k07, fa108 from acc1k0, fagent where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02" & strSql & " union " & _
'                     "select a1k27, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, a1k02, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, a1k13, a1k14, a1k15, a1k16, cu27 as fa21, cu28 as fa22, cu68 as fa35, a1k03, cu69 as fa36, a1k01, cu06 as fa06, cu29 as fa23, a1k04, a1k10, cu76 as fa43, a1k18 as Curr, cu102, a1k06, a1k07,cu148 as fa108 from acc1k0, customer where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02" & strSql & "", adoTaie, adOpenStatic, adLockReadOnly
      adoacc1k0.Open "select a1k27, fa05, fa63, fa64, fa65, fa32, fa18, a1k02, fa33, fa19, fa34, fa20, a1k13, a1k14, a1k15, a1k16, fa21, fa22, fa35, a1k03, fa36, a1k01, fa06, fa23, a1k04, a1k10, fa43, a1k18 as Curr, fa70 as cu102, a1k06, a1k07, fa108, a1k31 from acc1k0, fagent where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02" & strSql & " union " & _
                     "select a1k27, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, a1k02, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, a1k13, a1k14, a1k15, a1k16, cu27 as fa21, cu28 as fa22, cu68 as fa35, a1k03, cu69 as fa36, a1k01, cu06 as fa06, cu29 as fa23, a1k04, a1k10, cu76 as fa43, a1k18 as Curr, cu102, a1k06, a1k07,cu148 as fa108, a1k31 from acc1k0, customer where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02" & strSql & "", adoTaie, adOpenStatic, adLockReadOnly
      '2013/1/15 End
   Else
      'Modify by Morgan 2007/1/24 加 FA70
      'Modify By Sindy 2011/3/7 +FA108
      adoacc1k0.Open "select a1203 as a1k27, fa05, fa63, fa64, fa65, fa32, fa18, a1202 as a1k02, fa33, fa19, fa34, fa20, substr(a1208, 1, length(a1208) - 9) as a1k13, substr(a1208, length(a1208) - 8, 6) as a1k14, substr(a1208, length(a1208) - 2, 1) as a1k15, substr(a1208, length(a1208) - 1, 2) as a1k16, fa21, fa22, fa35, a1203 as a1k03, fa36, a1201 as a1k01, fa06, fa23, '' as a1k04, a1205 as a1k10, fa43, a1204 as Curr, fa70 as cu102, a1207 as a1k06, a1202 as a1k07, fa108 from acc120, fagent where substr(a1203, 1, 8) = fa01 and substr(a1203, 9, 1) =  fa02" & strSql & " union " & _
                     "select a1203 as a1k27, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, a1202 as a1k02, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, substr(a1208, 1, length(a1208) - 9) as a1k13, substr(a1208, length(a1208) - 8, 6) as a1k14, substr(a1208, length(a1208) - 2, 1) as a1k15, substr(a1208, length(a1208) - 1, 2) as a1k16, cu27 as fa21, cu28 as fa22, cu68 as fa35, a1203 as a1k03, cu69 as fa36, a1201 as a1k01, cu06 as fa06, cu29 as fa23, '' as a1k04, a1205 as a1k10, cu76 as fa43, a1204 as Curr, cu102, a1207 as a1k06, a1202 as a1k07,cu148 as fa108 from acc120, customer where substr(a1203, 1, 8) = cu01 and substr(a1203, 9, 1) = cu02" & strSql & "", adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/23
      strCon10 = MsgText(602)
      MsgBox MsgText(28), , MsgText(5)
      adoacc1k0.Close
      Exit Sub
   Else
      InsertQueryLog (adoacc1k0.RecordCount) 'Add By Sindy 2010/12/23
   End If
   Do While adoacc1k0.EOF = False
      If strNo <> adoacc1k0.Fields("a1k27").Value Then
         If douAmount <> 0 Then
            Printer.Line (500 + intDefault, 6300 + intCounter * 300 - 200 + intTop)-(10500 + intDefault, 6300 + intCounter * 300 - 200 + intTop)
'            PrintSum
            douAmount = 0
            douUSDollar = 0
            strNewPage = ""
            Printer.NewPage
         End If
         intCounter = 0
         PrintHead "N"
         strNo = adoacc1k0.Fields("a1k27").Value
      End If
      'Modify By Sindy 2011/3/7
      If CheckSys("" & adoacc1k0.Fields("a1k13").Value) = "2" Or _
         CheckSys("" & adoacc1k0.Fields("a1k13").Value) = "6" Then
         If IsNull(adoacc1k0.Fields("fa108").Value) = False Then
            strCurr = adoacc1k0.Fields("fa108").Value
         Else
            strCurr = MsgText(601)
         End If
      '2011/3/7 End
      Else
         If IsNull(adoacc1k0.Fields("fa43").Value) = False Then
            strCurr = adoacc1k0.Fields("fa43").Value
         Else
            strCurr = MsgText(601)
         End If
      End If
      Printer.CurrentX = 1000 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(109)
      intCounter = intCounter + 1
      Printer.CurrentX = 500 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
'      Printer.Print ReportSum(110) & adoacc1k0.Fields("a1k01").Value
'      intCounter = intCounter + 1
'      Printer.CurrentX = 3000 + intDefault
'      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      'Add By Sindy 2013/1/15
      If Option1.Value Then
         Select Case strCurr
            'Modify By Sindy 2013/1/17
            Case "NTD"
               Printer.Print ReportSum(110) & adoacc1k0.Fields("a1k01").Value & "(NTD" & Format(adoacc1k0.Fields("a1k06").Value, FDollar) & ")" & ReportSum(111)
            Case Else
               Printer.Print ReportSum(110) & adoacc1k0.Fields("a1k01").Value & "(" & adoacc1k0.Fields("Curr").Value & Format(adoacc1k0.Fields("a1k31").Value, FDollar) & ")" & ReportSum(111)
         End Select
      Else
      '2013/1/15 End
         Select Case strCurr
            'Modify By Sindy 2013/1/17
            Case "NTD"
               Printer.Print ReportSum(110) & adoacc1k0.Fields("a1k01").Value & "(NTD" & Format(adoacc1k0.Fields("a1k06").Value * adoacc1k0.Fields("a1k10").Value, FDollar) & ")" & ReportSum(111)
            Case Else
               Printer.Print ReportSum(110) & adoacc1k0.Fields("a1k01").Value & "(USD" & Format(adoacc1k0.Fields("a1k06").Value, FDollar) & ")" & ReportSum(111)
         End Select
      End If
      intCounter = intCounter + 2
      Printer.CurrentX = 1000 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(112)
      intCounter = intCounter + 1
      Printer.CurrentX = 500 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(113)
      intCounter = intCounter + 3
      Printer.CurrentX = 6000 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(94)
      intCounter = intCounter + 2
      Printer.CurrentX = 6000 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(114)
      intCounter = intCounter + 1
      Printer.CurrentX = 6000 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(115)
      intCounter = intCounter + 1
      Printer.CurrentX = 6000 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(116)
      intCounter = intCounter + 1
      Printer.CurrentX = 500 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(117)
      intCounter = intCounter + 2
      Printer.CurrentX = 500 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(118)
      Printer.NewPage
      intCounter = 0
      PrintHead "Y"
      intCounter = intCounter + 2
      Printer.Line (500 + intDefault, 6300 + intCounter * 300 - 200 + intTop)-(10500 + intDefault, 6300 + intCounter * 300 - 200 + intTop)
      intCounter = intCounter + 2
      Printer.CurrentX = 500 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print ReportSum(119)
      Printer.CurrentX = 8500 + intDefault
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      'Add By Sindy 2013/1/15
      If Option1.Value Then
         Select Case strCurr
            'Modify By Sindy 2013/1/17
            Case "NTD"
               Printer.Print "NTD"
               strPrintCurr = "NTD"
               strAmount = Format(adoacc1k0.Fields("a1k06").Value, FDollar)
            Case Else
               Printer.Print adoacc1k0.Fields("Curr").Value
               strPrintCurr = adoacc1k0.Fields("Curr").Value
               strAmount = Format(adoacc1k0.Fields("a1k31").Value, FDollar)
         End Select
      Else
      '2013/1/15 End
         Select Case strCurr
            'Modify By Sindy 2013/1/17
            Case "NTD"
               Printer.Print "NTD"
               strPrintCurr = "NTD" 'Add By Sindy 2013/1/15
               strAmount = Format(adoacc1k0.Fields("a1k06").Value * adoacc1k0.Fields("a1k10").Value, FDollar)
            Case Else
               Printer.Print "USD"
               strPrintCurr = "USD" 'Add By Sindy 2013/1/15
               strAmount = Format(adoacc1k0.Fields("a1k06").Value, FDollar)
         End Select
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 10500 + intDefault - intLength
      Printer.CurrentY = 6300 + intCounter * 300 + intTop
      Printer.Print strAmount
      intCounter = intCounter + 2
      adoacc1k0.MoveNext
   Loop
   Printer.Line (500 + intDefault, 6300 + intCounter * 300 - 200 + intTop)-(10500 + intDefault, 6300 + intCounter * 300 - 200 + intTop)
   PrintSum
   adoacc1k0.Close
   Printer.EndDoc
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead(strYes As String)
Dim intRow As Integer
Dim strCustName As String
Dim strSystemName As String
Dim strCaseName As String
Dim strAppNo As String
Dim strCustNo As String
Dim strProperty As String
   
'   If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select cu05, cu88, cu89, cu90, pa85 as Lang from patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
                    "union select cu05, cu88, cu89, cu90, tm53 as Lang from trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
                    "union select cu05, cu88, cu89, cu90, '' as Lang from lawcase, customer where substr(lc11, 1, 8) = cu01 and substr(lc11, 9, 1) = cu02 and lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
                    "union select cu05, cu88, cu89, cu90, '' as Lang from hirecase, customer where substr(hc05, 1, 8) = cu01 and substr(hc05, 9, 1) = cu02 and hc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and hc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and hc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and hc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
                    "union select cu05, cu88, cu89, cu90, sp34 as Lang from servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields("Lang").Value) = False Then
            strLanguage = adoquery.Fields("Lang").Value
         Else
            strLanguage = "2"
         End If
         strCustName = ""
         If IsNull(adoquery.Fields("cu05").Value) = False Then
            If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
                  Printer.Print adoquery.Fields("cu05").Value
               End If
            End If
            strCustName = adoquery.Fields("cu05").Value
         End If
         If IsNull(adoquery.Fields("cu88").Value) = False Then
            If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
               intRow = intRow + 1
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
                  Printer.Print adoquery.Fields("cu88").Value
               End If
            End If
            strCustName = strCustName & adoquery.Fields("cu88").Value
         End If
         If IsNull(adoquery.Fields("cu89").Value) = False Then
            If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
               intRow = intRow + 1
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
                  Printer.Print adoquery.Fields("cu89").Value
               End If
            End If
            strCustName = strCustName & adoquery.Fields("cu89").Value
         End If
         If IsNull(adoquery.Fields("cu90").Value) = False Then
            If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
               intRow = intRow + 1
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
                  Printer.Print adoquery.Fields("cu90").Value
               End If
            End If
            strCustName = strCustName & adoquery.Fields("cu90").Value
         End If
      Else
         strLanguage = "2"
         strCustName = ""
      End If
      adoquery.Close
   intRow = intRow + 2
   If adoacc1k0.Fields("a1k04").Value = MsgText(602) Then
      Printer.CurrentX = 0 + intDefault
      Printer.CurrentY = 1500 + intRow * 300 + intTop
      Printer.Print "C/O"
   End If
   Select Case strLanguage
      Case "2"
         If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa05").Value
         End If
         If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
            intRow = intRow + 1
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa63").Value
         End If
         If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
            intRow = intRow + 1
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa64").Value
         End If
         If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
            intRow = intRow + 1
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa65").Value
         End If
         intRow = intRow + 1
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               Printer.Print adoacc1k0.Fields("fa18").Value
            End If
         Else
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa32").Value
         End If
         Printer.CurrentX = 5500 + intDefault
         Printer.CurrentY = 2100 + intTop
         Printer.Print "Date:"
         Printer.CurrentX = 6500 + intDefault
         Printer.CurrentY = 2100 + intTop
         Printer.Print Format(AFDate(CADate(adoacc1k0.Fields("a1k07").Value)), "mmm. d, yyyy")
         intRow = intRow + 1
         'Modify by Morgan 2007/1/24 應該都要判斷FA32才對
         'If IsNull(adoacc1k0.Fields("fa33").Value) Then
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               Printer.Print adoacc1k0.Fields("fa19").Value
            End If
         Else
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "" & adoacc1k0.Fields("fa33").Value
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select pa77 as Yno, pa48 as Cno, ptm05, pa06 as Cname, pa11 as Ano, pa26 as Custno from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm01 (+) and sk02 = ptm02 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select tm45 as Yno, tm35 as Cno, ptm05, tm06 as Cname, tm12 as Ano, tm23 as Custno from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm01 (+) and sk02 = ptm02 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select lc23 as Yno, lc17 as Cno, '' as ptm05, lc06 as Cname, '' as Ano, lc11 as Custno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select sp27 as Yno, sp29 as Cno, '' as ptm05, sp06 as Cname, sp11 as Ano, sp08 as Custno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            Printer.CurrentX = 5500 + intDefault
            Printer.CurrentY = 2400 + intTop
            Printer.Print "Your Ref:"
            Printer.CurrentX = 6500 + intDefault
            Printer.CurrentY = 2400 + intTop
            If IsNull(adoquery.Fields("Yno").Value) Then
               Printer.Print ""
            Else
               Printer.Print adoquery.Fields("Yno").Value
            End If
            intRow = intRow + 1
            'Modify by Morgan 2007/1/24 應該都要判斷FA32才對
            'If IsNull(adoacc1k0.Fields("fa34").Value) Then
            If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
                  Printer.CurrentX = 500 + intDefault
                  Printer.CurrentY = 1500 + intRow * 300 + intTop
                  Printer.Print adoacc1k0.Fields("fa20").Value
               End If
            Else
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               Printer.Print "" & adoacc1k0.Fields("fa34").Value
            End If
            Printer.CurrentX = 5500 + intDefault
            Printer.CurrentY = 2700 + intTop
            Printer.Print "Our Ref:"
            Printer.CurrentX = 6500 + intDefault
            Printer.CurrentY = 2700 + intTop
            Printer.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
            intRow = intRow + 1
            'Modify by Morgan 2007/1/24 應該都要判斷FA32才對
            'If IsNull(adoacc1k0.Fields("fa35").Value) Then
            If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
                  Printer.CurrentX = 500 + intDefault
                  Printer.CurrentY = 1500 + intRow * 300 + intTop
                  Printer.Print adoacc1k0.Fields("fa21").Value
               End If
            Else
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               Printer.Print "" & adoacc1k0.Fields("fa35").Value
            End If
            If IsNull(adoquery.Fields("Cno").Value) = False Then
               Printer.CurrentX = 5500 + intDefault
               Printer.CurrentY = 3000 + intTop
               Printer.Print "Case No:"
               Printer.CurrentX = 6500 + intDefault
               Printer.CurrentY = 3000 + intTop
               If IsNull(adoquery.Fields("Cno").Value) Then
                  Printer.Print ""
               Else
                  Printer.Print adoquery.Fields("Cno").Value
               End If
            Else
               If adoacc1k0.Fields("a1k03").Value = "Y20438020" Then
                  Printer.CurrentX = 5500 + intDefault
                  Printer.CurrentY = 3000 + intTop
                  Printer.Print "Vendor Code # 88125-0"
               End If
            End If
            If IsNull(adoquery.Fields("ptm05").Value) Then
               strSystemName = ""
            Else
               strSystemName = adoquery.Fields("ptm05").Value
            End If
            If IsNull(adoquery.Fields("Cname").Value) Then
               strCaseName = ""
            Else
               strCaseName = adoquery.Fields("Cname").Value
            End If
            If IsNull(adoquery.Fields("Ano").Value) Then
               strAppNo = ""
            Else
               strAppNo = adoquery.Fields("Ano").Value
            End If
            If IsNull(adoquery.Fields("Custno").Value) Then
               strCustNo = ""
            Else
               strCustNo = adoquery.Fields("Custno").Value
            End If
         End If
         adoquery.Close
         intRow = intRow + 1
         'Modify by Morgan 2007/1/24 應該都要判斷FA32才對
         'If IsNull(adoacc1k0.Fields("fa36").Value) Then
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               Printer.Print adoacc1k0.Fields("fa22").Value
            End If
         Else
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "" & adoacc1k0.Fields("fa36").Value
         End If
         intRow = intRow + 1
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("cu102").Value) = False Then
               Printer.CurrentX = 500 + intDefault
               Printer.CurrentY = 1500 + intRow * 300 + intTop
               Printer.Print adoacc1k0.Fields("cu102").Value
            End If
         End If
         If strYes = MsgText(602) Then
            intRow = intRow + 1
            Printer.CurrentX = 4000 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "CREDIT NOTE"
            Printer.CurrentX = 7000 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "No."
            Printer.CurrentX = 7500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("a1k01").Value
            Printer.Line (500 + intDefault, 1500 + intRow * 300 + 350 + intTop)-(10500 + intDefault, 1500 + intRow * 300 + 350 + intTop)
         End If
         intRow = intRow + 1
      Case "3"
         If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa06").Value
         End If
         intRow = intRow + 1
         If IsNull(adoacc1k0.Fields("fa23").Value) = False Then
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa23").Value
         End If
         Printer.CurrentX = 5500 + intDefault
         Printer.CurrentY = 2100 + intTop
         Printer.Print "Date:"
         Printer.CurrentX = 6500 + intDefault
         Printer.CurrentY = 2100 + intTop
         Printer.Print Format(AFDate(CADate(adoacc1k0.Fields("a1k07").Value)), "mmm. d, yyyy")
         intRow = intRow + 1
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select pa77 as Yno, pa48 as Cno, ptm06, pa07 as Cname, pa11 as Ano, pa26 as Custno from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm01 (+) and sk02 = ptm02 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select tm45 as Yno, tm35 as Cno, ptm06, tm07 as Cname, tm12 as Ano, tm23 as Custno from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm01 (+) and sk02 = ptm02 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select lc23 as Yno, lc17 as Cno, '' as ptm06, lc07 as Cname, '' as Ano, lc11 as Custno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select sp27 as Yno, sp29 as Cno, '' as ptm06, sp07 as Cname, sp11 as Ano, sp08 as Custno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            Printer.CurrentX = 5500 + intDefault
            Printer.CurrentY = 2400 + intTop
            Printer.Print "Your Ref:"
            Printer.CurrentX = 6500 + intDefault
            Printer.CurrentY = 2400 + intTop
            If IsNull(adoquery.Fields("Yno").Value) Then
               Printer.Print ""
            Else
               Printer.Print adoquery.Fields("Yno").Value
            End If
            intRow = intRow + 1
            Printer.CurrentX = 5500 + intDefault
            Printer.CurrentY = 2700 + intTop
            Printer.Print "Our Ref:"
            Printer.CurrentX = 6500 + intDefault
            Printer.CurrentY = 2700 + intTop
            Printer.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
            intRow = intRow + 1
            If IsNull(adoquery.Fields("Cno").Value) = False Then
               Printer.CurrentX = 5500 + intDefault
               Printer.CurrentY = 3000 + intTop
               Printer.Print "Case No:"
               Printer.CurrentX = 6500 + intDefault
               Printer.CurrentY = 3000 + intTop
               If IsNull(adoquery.Fields("Cno").Value) Then
                  Printer.Print ""
               Else
                  Printer.Print adoquery.Fields("Cno").Value
               End If
            Else
               If adoacc1k0.Fields("a1k03").Value = "Y20438020" Then
                  Printer.CurrentX = 5500 + intDefault
                  Printer.CurrentY = 3000 + intTop
                  Printer.Print "Vendor Code # 88125-0"
               End If
            End If
            If IsNull(adoquery.Fields("ptm06").Value) Then
               strSystemName = ""
            Else
               strSystemName = adoquery.Fields("ptm06").Value
            End If
            If IsNull(adoquery.Fields("Cname").Value) Then
               strCaseName = ""
            Else
               strCaseName = adoquery.Fields("Cname").Value
            End If
            If IsNull(adoquery.Fields("Ano").Value) Then
               strAppNo = ""
            Else
               strAppNo = adoquery.Fields("Ano").Value
            End If
            If IsNull(adoquery.Fields("Custno").Value) Then
               strCustNo = ""
            Else
               strCustNo = adoquery.Fields("Custno").Value
            End If
         End If
         adoquery.Close
         If strYes = MsgText(602) Then
            intRow = intRow + 1
            Printer.CurrentX = 4000 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "CREDIT NOTE"
            Printer.CurrentX = 7000 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "No."
            Printer.CurrentX = 7500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("a1k01").Value
            Printer.Line (500 + intDefault, 1500 + intRow * 300 + 350 + intTop)-(10500 + intDefault, 1500 + intRow * 300 + 350 + intTop)
         End If
         intRow = intRow + 1
      Case "1"
         If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa04").Value
         End If
         intRow = intRow + 1
         If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
            Printer.CurrentX = 500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("fa18").Value
         End If
         Printer.CurrentX = 5500 + intDefault
         Printer.CurrentY = 2100 + intTop
         Printer.Print "Date:"
         Printer.CurrentX = 6500 + intDefault
         Printer.CurrentY = 2100 + intTop
         Printer.Print Format(AFDate(CADate(adoacc1k0.Fields("a1k07").Value)), "mmm. d, yyyy")
         intRow = intRow + 1
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select pa77 as Yno, pa48 as Cno, ptm05, pa07 as Cname, pa11 as Ano, pa26 as Custno from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm01 (+) and sk02 = ptm02 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select tm45 as Yno, tm35 as Cno, ptm05, tm07 as Cname, tm12 as Ano, tm23 as Custno from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm01 (+) and sk02 = ptm02 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select lc23 as Yno, lc17 as Cno, '' as ptm05, lc07 as Cname, '' as Ano, lc11 as Custno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select sp27 as Yno, sp29 as Cno, '' as ptm05, sp07 as Cname, sp11 as Ano, sp08 as Custno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            Printer.CurrentX = 5500 + intDefault
            Printer.CurrentY = 2400 + intTop
            Printer.Print "Your Ref:"
            Printer.CurrentX = 6500 + intDefault
            Printer.CurrentY = 2400 + intTop
            If IsNull(adoquery.Fields("Yno").Value) Then
               Printer.Print ""
            Else
               Printer.Print adoquery.Fields("Yno").Value
            End If
            intRow = intRow + 1
            Printer.CurrentX = 5500 + intDefault
            Printer.CurrentY = 2700 + intTop
            Printer.Print "Our Ref:"
            Printer.CurrentX = 6500 + intDefault
            Printer.CurrentY = 2700 + intTop
            Printer.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
            intRow = intRow + 1
            If IsNull(adoquery.Fields("Cno").Value) = False Then
               Printer.CurrentX = 5500 + intDefault
               Printer.CurrentY = 3000 + intTop
               Printer.Print "Case No:"
               Printer.CurrentX = 6500 + intDefault
               Printer.CurrentY = 3000 + intTop
               If IsNull(adoquery.Fields("Cno").Value) Then
                  Printer.Print ""
               Else
                  Printer.Print adoquery.Fields("Cno").Value
               End If
            Else
               If adoacc1k0.Fields("a1k03").Value = "Y20438020" Then
                  Printer.CurrentX = 5500 + intDefault
                  Printer.CurrentY = 3000 + intTop
                  Printer.Print "Vendor Code # 88125-0"
               End If
            End If
            If IsNull(adoquery.Fields("ptm05").Value) Then
               strSystemName = ""
            Else
               strSystemName = adoquery.Fields("ptm05").Value
            End If
            If IsNull(adoquery.Fields("Cname").Value) Then
               strCaseName = ""
            Else
               strCaseName = adoquery.Fields("Cname").Value
            End If
            If IsNull(adoquery.Fields("Ano").Value) Then
               strAppNo = ""
            Else
               strAppNo = adoquery.Fields("Ano").Value
            End If
            If IsNull(adoquery.Fields("Custno").Value) Then
               strCustNo = ""
            Else
               strCustNo = adoquery.Fields("Custno").Value
            End If
         End If
         adoquery.Close
         If strYes = MsgText(602) Then
            intRow = intRow + 1
            Printer.CurrentX = 4000 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "CREDIT NOTE"
            Printer.CurrentX = 7000 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print "No."
            Printer.CurrentX = 7500 + intDefault
            Printer.CurrentY = 1500 + intRow * 300 + intTop
            Printer.Print adoacc1k0.Fields("a1k01").Value
            Printer.Line (500 + intDefault, 1500 + intRow * 300 + 350 + intTop)-(10500 + intDefault, 1500 + intRow * 300 + 350 + intTop)
         End If
         intRow = intRow + 1
   End Select
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select cp10 from caseprogress where cp60 = '" & adoacc1k0.Fields("a1k01").Value & "' and cp10 >= '101' and cp10 <= '105'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      strProperty = "new  "
   Else
      strProperty = " "
   End If
   adoquery.Close
   intRow = intRow + 1
   Printer.CurrentX = 500 + intDefault
   Printer.CurrentY = 1500 + intRow * 300 + intTop
   Printer.Print ReportSum(84) & strProperty & strSystemName & " No. " & strAppNo
   If strCustNo = "X22232010" Then
      intRow = intRow + 1
      Printer.CurrentX = 900 + intDefault
      Printer.CurrentY = 1500 + intRow * 300 + intTop
      Printer.Print strCaseName
   End If
   intRow = intRow + 1
   Printer.CurrentX = 900 + intDefault
   Printer.CurrentY = 1500 + intRow * 300 + intTop
   Printer.Print "Applicant: " & strCustName
End Sub

'*************************************************
' 合計位置
'
'*************************************************
Private Sub PrintSum()
   intCounter = intCounter - 1
   Printer.CurrentX = 7000 + intDefault
   Printer.CurrentY = 6600 + intCounter * 300 + intTop
   Printer.Print ReportSum(120)
   Printer.CurrentX = 8500 + intDefault
   Printer.CurrentY = 6600 + intCounter * 300 + intTop
   'Modify By Sindy 2013/1/15
'   Select Case strCurr
'      Case "N"
'         Printer.Print "NTD"
'      Case Else
'         Printer.Print "USD"
'   End Select
   Printer.Print strPrintCurr
   '2013/1/15 End
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 10500 + intDefault - intLength
   Printer.CurrentY = 6600 + intCounter * 300 + intTop
   Printer.Print strAmount
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Option1.Value Then
      If Text1 <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   Else
      If Text3 <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   FormCheck = False
End Function

Private Sub Text3_GotFocus()
   TextInverse Text3
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
