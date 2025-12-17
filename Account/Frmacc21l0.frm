VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc21l0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外收款分析表"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1320
   ScaleWidth      =   5160
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel表"
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
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   720
      Width           =   4452
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   252
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc21l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/03 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
'Remove by Lydia 2018/07/12
'Public adoacc1p0 As New ADODB.Recordset
'Dim strSql As String

Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2108/07/12 改版
   'ExcelSave
   If MaskEdBox1.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(29) Then
       MsgBox "請輸入日期 !"
       If MaskEdBox1.Text = MsgText(29) Then
           MaskEdBox1.SetFocus
       Else
           MaskEdBox2.SetFocus
       End If
   Else
        ExcelSave2
   End If
   'end 2018/07/12
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 1700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21l0 = Nothing
End Sub

'*************************************************
'  轉成Excel檔案
'
'*************************************************
'Remove by Lydia 2018/07/12
'Private Sub ExcelSave()
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt225 As New Worksheet
'Dim lngCounter As Long, strTotalAmt As String
'Dim strName As String
'Dim strTotal(7) As String
'Dim intCounter As Integer
'
'On Error GoTo Checking
'   strSql = ""
'   If MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a1p18 >= " & FCDate(MaskEdBox1.Text) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1p18 <= " & FCDate(MaskEdBox2.Text) & ""
'   End If
''   If Dir(strExcelPath & ReportTitle(225) & ACDate(ServerDate) & MsgText(43)) = MsgText(601) Then
''      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
''         MkDir strExcelPath
''      End If
''   Else
''      Kill strExcelPath & ReportTitle(225) & ACDate(ServerDate) & MsgText(43)
''   End If
'   If Dir(strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43)
'   End If
'   xlsSalesPoint.Workbooks.add
'   Set wksaccrpt225 = xlsSalesPoint.Worksheets(1)
'   wksaccrpt225.Columns("a:a").ColumnWidth = 15
'   wksaccrpt225.Columns("b:b").ColumnWidth = 15
'   wksaccrpt225.Columns("c:c").ColumnWidth = 15
'   wksaccrpt225.Columns("d:d").ColumnWidth = 15
'   wksaccrpt225.Columns("e:e").ColumnWidth = 15
'   wksaccrpt225.Columns("f:f").ColumnWidth = 15
'   wksaccrpt225.Columns("g:g").ColumnWidth = 15
'   wksaccrpt225.Range("a1").Value = ReportTitle(225)
'   wksaccrpt225.Range("a2").Value = "收款日期"
'   wksaccrpt225.Range("b2").Value = MaskEdBox1.Text
'   wksaccrpt225.Range("c2").Value = "~"
'   wksaccrpt225.Range("d2").Value = MaskEdBox2.Text
'   wksaccrpt225.Range("a3").Value = "IR"
'   wksaccrpt225.Range("a3").HorizontalAlignment = xlCenter
'   wksaccrpt225.Range("b3").Value = "CB"
'   wksaccrpt225.Range("b3").HorizontalAlignment = xlCenter
'   wksaccrpt225.Range("c3").Value = "託收"
'   wksaccrpt225.Range("c3").HorizontalAlignment = xlCenter
'   wksaccrpt225.Range("d3").Value = "其他"
'   wksaccrpt225.Range("d3").HorizontalAlignment = xlCenter
'   wksaccrpt225.Range("e3").Value = "匯入台幣"
'   wksaccrpt225.Range("e3").HorizontalAlignment = xlCenter
'   wksaccrpt225.Range("f3").Value = "託收到期"
'   wksaccrpt225.Range("f3").HorizontalAlignment = xlCenter
'   wksaccrpt225.Range("g3").Value = "託收原幣"
'   wksaccrpt225.Range("g3").HorizontalAlignment = xlCenter
'   lngCounter = 4
'   adoacc1p0.CursorLocation = adUseClient
'   adoacc1p0.Open "select * from (select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p24 = '1'" & strSql & " union " & _
'                  "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p24 = '1' and a1p07 > 0" & strSql & ") new order by a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc1p0.EOF = False
'      If IsNull(adoacc1p0.Fields("a1p21").Value) Then
'         wksaccrpt225.Range("a" & lngCounter).Value = MsgText(601)
'      Else
'         If Mid(adoacc1p0.Fields("a1p19").Value, 1, 2) = "US" Then
'            wksaccrpt225.Range("a" & lngCounter).Value = Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'            strTotal(0) = Val(strTotal(0)) + Val(adoacc1p0.Fields("a1p21").Value)
'         Else
'            wksaccrpt225.Range("a" & lngCounter).Value = adoacc1p0.Fields("a1p19").Value & "  " & Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'         End If
'      End If
'      lngCounter = lngCounter + 1
'      adoacc1p0.MoveNext
'   Loop
'   adoacc1p0.Close
'   intCounter = lngCounter
'   lngCounter = 4
'   adoacc1p0.CursorLocation = adUseClient
'   adoacc1p0.Open "select * from (select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p24 = '2'" & strSql & " union " & _
'                  "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p24 = '2' and a1p07 > 0" & strSql & ") new order by a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc1p0.EOF = False
'      If IsNull(adoacc1p0.Fields("a1p21").Value) Then
'         wksaccrpt225.Range("b" & lngCounter).Value = MsgText(601)
'      Else
'         If Mid(adoacc1p0.Fields("a1p19").Value, 1, 2) = "US" Then
'            wksaccrpt225.Range("b" & lngCounter).Value = Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'            strTotal(1) = Val(strTotal(1)) + Val(adoacc1p0.Fields("a1p21").Value)
'         Else
'            wksaccrpt225.Range("b" & lngCounter).Value = adoacc1p0.Fields("a1p19").Value & "  " & Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'         End If
'      End If
'      lngCounter = lngCounter + 1
'      adoacc1p0.MoveNext
'   Loop
'   adoacc1p0.Close
'   If intCounter < lngCounter Then
'      intCounter = lngCounter
'   End If
'   lngCounter = 4
'   adoacc1p0.CursorLocation = adUseClient
'   adoacc1p0.Open "select * from (select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p24 = '3'" & strSql & " union " & _
'                  "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p24 = '3' and a1p07 > 0" & strSql & ") new order by a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc1p0.EOF = False
'      If IsNull(adoacc1p0.Fields("a1p21").Value) Then
'         wksaccrpt225.Range("c" & lngCounter).Value = MsgText(601)
'      Else
'         If Mid(adoacc1p0.Fields("a1p19").Value, 1, 2) = "US" Then
'            wksaccrpt225.Range("c" & lngCounter).Value = Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'            strTotal(2) = Val(strTotal(2)) + Val(adoacc1p0.Fields("a1p21").Value)
'         Else
'            wksaccrpt225.Range("c" & lngCounter).Value = adoacc1p0.Fields("a1p19").Value & "  " & Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'         End If
'      End If
'      lngCounter = lngCounter + 1
'      adoacc1p0.MoveNext
'   Loop
'   adoacc1p0.Close
'   If intCounter < lngCounter Then
'      intCounter = lngCounter
'   End If
'   lngCounter = 4
'   adoacc1p0.CursorLocation = adUseClient
'   adoacc1p0.Open "select * from (select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p24 in ('4', '5')" & strSql & " union " & _
'                  "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p24 in ('4', '5') and a1p07 > 0" & strSql & ") new order by a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc1p0.EOF = False
'      If IsNull(adoacc1p0.Fields("a1p21").Value) Then
'         wksaccrpt225.Range("d" & lngCounter).Value = MsgText(601)
'      Else
'         If Mid(adoacc1p0.Fields("a1p19").Value, 1, 2) = "US" Then
'            wksaccrpt225.Range("d" & lngCounter).Value = Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'            strTotal(3) = Val(strTotal(3)) + Val(adoacc1p0.Fields("a1p21").Value)
'         Else
'            wksaccrpt225.Range("d" & lngCounter).Value = adoacc1p0.Fields("a1p19").Value & "  " & Format(Val(adoacc1p0.Fields("a1p21").Value), FDollar)
'         End If
'      End If
'      lngCounter = lngCounter + 1
'      adoacc1p0.MoveNext
'   Loop
'   adoacc1p0.Close
'   If intCounter < lngCounter Then
'      intCounter = lngCounter
'   End If
'   intCounter = intCounter + 1
'   wksaccrpt225.Range("a" & intCounter).Value = Format(strTotal(0), FDollar)
'   wksaccrpt225.Range("b" & intCounter).Value = Format(strTotal(1), FDollar)
'   wksaccrpt225.Range("c" & intCounter).Value = Format(strTotal(2), FDollar)
'   wksaccrpt225.Range("d" & intCounter).Value = Format(strTotal(3), FDollar)
'    'Modify By Cheng 2003/06/09
''   xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & MsgText(43)
'   'Modify by Amy 2016/06/23 +判斷版本
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   MsgBox "檔案已產生"
'   FormClear
'   Exit Sub
'
'Checking:
'    MsgBox Err.Description, , MsgText(5)
'    If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'    xlsSalesPoint.Workbooks.Close
'    xlsSalesPoint.Quit
'    Set xlsSalesPoint = Nothing
'    Set wksaccrpt225 = Nothing
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   MaskEdBox1.SetFocus
End Sub

'Added by Lydia 2018/07/12 改版
Private Sub ExcelSave2()
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt225 As New Worksheet
Dim rsRead As New ADODB.Recordset
Dim strTempName As String, strFileName As String
Dim xRows As Integer, intJ As Integer
Dim intCounter As Integer
Dim strCurr As String, strA1p05 As String, strA1P24 As String
Dim strX1 As String, strX2 As String
Dim strDate1 As String, strDate2 As String
Dim strCon1 As String, strCon2 As String

On Error GoTo Checking
   
   strTempName = "國外收款各幣別分析"
   strFileName = strTempName
   
   If MaskEdBox1.Text <> MsgText(29) Then
      strDate1 = FCDate(MaskEdBox1.Text)
      strFileName = strFileName & strDate1
      strCon1 = strCon1 & " and a1p18 >= " & strDate1
      strCon2 = strCon2 & " and a1p18 >= " & Val(Mid(strDate1, 1, 3)) - 1 & Mid(strDate1, 4)
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      strDate2 = FCDate(MaskEdBox2.Text)
      strFileName = strFileName & IIf(MaskEdBox1.Text <> MsgText(29), "-", "") & strDate2
      strCon1 = strCon1 & " and a1p18 <= " & strDate2
      strCon2 = strCon2 & " and a1p18 <= " & Val(Mid(strDate2, 1, 3)) - 1 & Mid(strDate2, 4)
   End If
   If Dir(strExcelPath & strFileName & MsgText(43)) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & strFileName & MsgText(43)
   End If
   
   '當期
   '因為國外匯款進來,多半是入台銀,所以抓指定科目;科目名稱簡化
   'Modified by Lydia 2020/09/10 +北京招商(1917)
   'Modified by Lydia 2020/09/11 已不用1公司(商標)美金帳戶110228
   'strExc(1) = "decode(a1p05,'110205','台銀綜專','110228','台銀綜商','113002','外票','1917','北京招商',a1p05)"
   'strSql = " select a1p05," & strExc(1) & " as acc_name ,decode(a1p19,'USD','1','2') ord1,substr(a1p18,1,3) yy, a1p24,a1p19,sum(a1p21) uamt,sum(nvl(a1p07,0)+nvl(a1p08,0)) tamt" & _
                     " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon1 & _
                     " and a1p05 in ('110205','110228','113002','1917') and a1p24 >='1' and a1p24<='5' " & _
                     " group by a1p05," & strExc(1) & ",decode(a1p19,'USD','1','2'),substr(a1p18,1,3),a1p24,a1p19"
   strExc(1) = "decode(a1p05,'110205','台銀綜專','113002','外票','1917','北京招商',a1p05)"
   strSql = " select a1p05," & strExc(1) & " as acc_name ,decode(a1p19,'USD','1','2') ord1,substr(a1p18,1,3) yy, a1p24,a1p19,sum(a1p21) uamt,sum(nvl(a1p07,0)+nvl(a1p08,0)) tamt" & _
                     " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon1 & _
                     " and a1p05 in ('110205','113002','1917') and a1p24 >='1' and a1p24<='5' " & _
                     " group by a1p05," & strExc(1) & ",decode(a1p19,'USD','1','2'),substr(a1p18,1,3),a1p24,a1p19"
   '前一期
   'Modified by Lydia 2020/09/10 +北京招商(1917)
   'Modified by Lydia 2020/09/11 已不用1公司(商標)美金帳戶110228
   'strSql = strSql & " Union all select a1p05," & strExc(1) & " as acc_name,decode(a1p19,'USD','1','2') ord1,substr(a1p18,1,3) yy, a1p24,a1p19,sum(a1p21) uamt,sum(nvl(a1p07,0)+nvl(a1p08,0)) tamt" & _
                     " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon2 & _
                     " and a1p05 in ('110205','113002','1917') and a1p24 >='1' and a1p24<='5' " & _
                     " group by a1p05," & strExc(1) & ",decode(a1p19,'USD','1','2'),substr(a1p18,1,3),a1p24,a1p19"
   strSql = strSql & " Union all select a1p05," & strExc(1) & " as acc_name,decode(a1p19,'USD','1','2') ord1,substr(a1p18,1,3) yy, a1p24,a1p19,sum(a1p21) uamt,sum(nvl(a1p07,0)+nvl(a1p08,0)) tamt" & _
                     " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon2 & _
                     " and a1p05 in ('110205','110228','113002','1917') and a1p24 >='1' and a1p24<='5' " & _
                     " group by a1p05," & strExc(1) & ",decode(a1p19,'USD','1','2'),substr(a1p18,1,3),a1p24,a1p19"
   strSql = strSql & " order by a1p05,ord1,a1p19,a1p24,yy desc "
   intJ = 0
   Set rsRead = ClsLawReadRstMsg(intJ, strSql)
   If intJ = 0 Then
       Set rsRead = Nothing
       Exit Sub
   End If

   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt225 = xlsSalesPoint.Worksheets(1)
   '設欄寬
   For intJ = Asc("A") To Asc("G")
        If intJ <= Asc("C") Then
            wksaccrpt225.Columns(Chr(intJ) & ":" & Chr(intJ)).ColumnWidth = 9
        Else
            wksaccrpt225.Columns(Chr(intJ) & ":" & Chr(intJ)).ColumnWidth = 13.5
        End If
   Next intJ

   wksaccrpt225.Range("A2").Value = strTempName
   wksaccrpt225.Range("A2").Font.Size = 16
   wksaccrpt225.Range("A3").Value = "查詢資料區間: " & ChangeTStringToTDateString(strDate1) & IIf(strDate1 <> "", "~", "") & ChangeTStringToTDateString(strDate2)
   wksaccrpt225.Range("A3").Font.Size = 16
   xRows = 5
   strX1 = "D"
   wksaccrpt225.Range(strX1 & xRows).Value = ChangeTStringToTDateString(strDate1) & IIf(strDate1 <> "", "~", "") & ChangeTStringToTDateString(strDate2)
   wksaccrpt225.Range(strX1 & xRows & ":" & Chr(Asc(strX1) + 1) & xRows).Merge
   strX2 = "F": strExc(0) = ""
   If strDate1 <> "" Then
        strExc(0) = strExc(0) & "~" & ChangeTStringToTDateString(Val(Mid(strDate1, 1, 3)) - 1 & Mid(strDate1, 4))
   End If
   If strDate2 <> "" Then
        strExc(0) = strExc(0) & "~" & ChangeTStringToTDateString(Val(Mid(strDate2, 1, 3)) - 1 & Mid(strDate2, 4))
   End If
   wksaccrpt225.Range(strX2 & xRows).Value = IIf(strExc(0) <> "", Mid(strExc(0), 2), "")
   wksaccrpt225.Range(strX2 & xRows & ":" & Chr(Asc(strX2) + 1) & xRows).Merge
   xRows = 6
   wksaccrpt225.Range("A" & xRows).Value = "幣別"
   wksaccrpt225.Range("B" & xRows).Value = "科目"
   wksaccrpt225.Range("C" & xRows).Value = "來源"
   wksaccrpt225.Range("D" & xRows).Value = "原幣金額"
   wksaccrpt225.Range("E" & xRows).Value = "當時收款台幣"
   wksaccrpt225.Range("F" & xRows).Value = "原幣金額"
   wksaccrpt225.Range("G" & xRows).Value = "當時收款台幣"
   wksaccrpt225.Range("A5:G6").HorizontalAlignment = xlCenter '置中
   xRows = 7
   intCounter = xRows
   
   '寫入資料
   rsRead.MoveFirst
   With rsRead
      Do While Not .EOF
          '不同科目別
          If strA1p05 <> "" & .Fields("a1p05") Then
              '加總-科目別
              If strA1p05 <> "" Then
                   xRows = xRows + 1
                   wksaccrpt225.Range("A" & xRows).Value = "小計"
                   wksaccrpt225.Range(Chr(Asc(strX1) + 1) & xRows).Formula = "=SUM(" & Chr(Asc(strX1) + 1) & intCounter & ":" & Chr(Asc(strX1) + 1) & xRows - 1 & ")"
                   wksaccrpt225.Range(Chr(Asc(strX2) + 1) & xRows).Formula = "=SUM(" & Chr(Asc(strX2) + 1) & intCounter & ":" & Chr(Asc(strX2) + 1) & xRows - 1 & ")"
                   xRows = xRows + 2
              End If
              intCounter = xRows
           '不同幣別或收款類別
          ElseIf strCurr <> "" & .Fields("a1p19") Or strA1P24 <> "" & .Fields("a1p24") Then
              xRows = xRows + 1
          End If
          
          '顯示幣別和科目
          wksaccrpt225.Range("A" & xRows).Value = "" & .Fields("a1p19")
          wksaccrpt225.Range("B" & xRows).Value = "" & .Fields("acc_name")

          '顯示收款類別
           If wksaccrpt225.Range(Chr(Asc(strX1) - 1) & xRows).Value = "" Then
              strExc(1) = ""
              Select Case "" & .Fields("a1p24")
                  Case "1": strExc(1) = ComboItem(51)
                  Case "2": strExc(1) = ComboItem(52)
                  Case "3": strExc(1) = ComboItem(53)
                  Case "4": strExc(1) = ComboItem(54)
                  Case "5": strExc(1) = ComboItem(55)
              End Select
              wksaccrpt225.Range(Chr(Asc(strX1) - 1) & xRows).Value = Mid(strExc(1), 4)
          End If
          
          '當期
          If "" & .Fields("yy") = Mid(strDate1, 1, 3) Then
               wksaccrpt225.Range(strX1 & xRows).Value = "" & .Fields("uamt")
               wksaccrpt225.Range(Chr(Asc(strX1) + 1) & xRows).Value = "" & .Fields("tamt")
          Else '前一期
               wksaccrpt225.Range(strX2 & xRows).Value = "" & .Fields("uamt")
               wksaccrpt225.Range(Chr(Asc(strX2) + 1) & xRows).Value = "" & .Fields("tamt")
          End If
          strA1p05 = "" & .Fields("a1p05")
          strCurr = "" & .Fields("a1p19")
          strA1P24 = "" & .Fields("a1p24")
          .MoveNext
      Loop
   End With
    '加總-科目別
    If strA1p05 <> "" Then
         xRows = xRows + 1
         wksaccrpt225.Range("A" & xRows).Value = "小計"
         wksaccrpt225.Range(Chr(Asc(strX1) + 1) & xRows).Formula = "=SUM(" & Chr(Asc(strX1) + 1) & intCounter & ":" & Chr(Asc(strX1) + 1) & xRows - 1 & ")"
         wksaccrpt225.Range(Chr(Asc(strX2) + 1) & xRows).Formula = "=SUM(" & Chr(Asc(strX2) + 1) & intCounter & ":" & Chr(Asc(strX2) + 1) & xRows - 1 & ")"
         xRows = xRows + 1
    End If
              
   '抓區間的美金匯率
   xRows = xRows + 1
   'Modified by Lydia 2020/09/10 +北京招商(1917)
   'Modified by Lydia 2020/09/11 已不用1公司(商標)美金帳戶110228
   'strSql = " select substr(a1p18,1,3) yy, min(a1p20) min_rate,max(a1p20) max_rate" & _
               " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon1 & _
               " and a1p05 in ('110205','110228','113002','1917') and a1p24 >='1' and a1p24<='5' and a1p19='USD' group by substr(a1p18,1,3)"
   strSql = " select substr(a1p18,1,3) yy, min(a1p20) min_rate,max(a1p20) max_rate" & _
               " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon1 & _
               " and a1p05 in ('110205','113002','1917') and a1p24 >='1' and a1p24<='5' and a1p19='USD' group by substr(a1p18,1,3)"
   'Modified by Lydia 2020/09/10 +北京招商(1917)
   'Modified by Lydia 2020/09/11 已不用1公司(商標)美金帳戶110228
   'strSql = strSql & "Union all select substr(a1p18,1,3) yy, min(a1p20) min_rate,max(a1p20) max_rate" & _
               " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon2 & _
               " and a1p05 in ('110205','110228','113002','1917') and a1p24 >='1' and a1p24<='5' and a1p19='USD' group by substr(a1p18,1,3)"
   strSql = strSql & "Union all select substr(a1p18,1,3) yy, min(a1p20) min_rate,max(a1p20) max_rate" & _
               " from acc1p0 where a1p01 = '1' and a1p02 in ('F','G')" & strCon2 & _
               " and a1p05 in ('110205','113002','1917') and a1p24 >='1' and a1p24<='5' and a1p19='USD' group by substr(a1p18,1,3)"
   strSql = strSql & " order by yy desc "
   intJ = 1
   Set rsRead = ClsLawReadRstMsg(intJ, strSql)
   If intJ = 1 Then
        rsRead.MoveFirst
        Do While Not rsRead.EOF
           '當期
           If "" & rsRead.Fields("yy") = Mid(strDate1, 1, 3) Then
                wksaccrpt225.Range(strX1 & xRows).Value = "美金匯率"
                wksaccrpt225.Range(Chr(Asc(strX1) + 1) & xRows).Value = "" & rsRead.Fields("min_rate") & "-" & rsRead.Fields("max_rate")
                wksaccrpt225.Range(Chr(Asc(strX1) + 1) & xRows).HorizontalAlignment = xlLeft
           Else '前一期
                wksaccrpt225.Range(strX2 & xRows).Value = "美金匯率"
                wksaccrpt225.Range(Chr(Asc(strX2) + 1) & xRows).Value = "" & rsRead.Fields("min_rate") & "-" & rsRead.Fields("max_rate")
                wksaccrpt225.Range(Chr(Asc(strX2) + 1) & xRows).HorizontalAlignment = xlLeft
           End If
           rsRead.MoveNext
        Loop
   End If
   
   '格式
   wksaccrpt225.Range(strX1 & "7:" & strX1 & xRows).NumberFormatLocal = "#,##0.00" '外幣
   wksaccrpt225.Range(strX2 & "7:" & strX2 & xRows).NumberFormatLocal = "#,##0.00"
   wksaccrpt225.Range(Chr(Asc(strX1) + 1) & "7:" & Chr(Asc(strX1) + 1) & xRows).NumberFormatLocal = "#,##0" '台幣
   wksaccrpt225.Range(Chr(Asc(strX2) + 1) & "7:" & Chr(Asc(strX2) + 1) & xRows).NumberFormatLocal = "#,##0"
   
   '全部表格-畫虛線
   wksaccrpt225.Range("A5:" & Chr(Asc(strX2) + 1) & xRows).Borders.LineStyle = xlContinuous
   wksaccrpt225.Range("A5:" & Chr(Asc(strX2) + 1) & xRows).Borders.Weight = xlHairline
   '畫實線
   wksaccrpt225.Range(strX1 & "5:" & Chr(Asc(strX2) + 1) & xRows & "5").Borders(xlEdgeTop).Weight = xlThin
   wksaccrpt225.Range(strX1 & xRows & ":" & Chr(Asc(strX2) + 1) & xRows).Borders(xlEdgeBottom).Weight = xlThin
   wksaccrpt225.Range(strX1 & "5:" & strX1 & xRows).Borders(xlEdgeLeft).Weight = xlThin
   wksaccrpt225.Range(strX2 & "5:" & strX2 & xRows).Borders(xlEdgeLeft).Weight = xlThin
   wksaccrpt225.Range(Chr(Asc(strX2) + 1) & "5:" & Chr(Asc(strX2) + 1) & xRows).Borders(xlEdgeRight).Weight = xlThin
'-------------------------------------------
   '判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & MsgText(43), FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   MsgBox "檔案已產生"
   FormClear
   Set rsRead = Nothing
   
   Exit Sub
   
Checking:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ReportTitle(225) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set rsRead = Nothing
    Set xlsSalesPoint = Nothing
    Set wksaccrpt225 = Nothing
End Sub

