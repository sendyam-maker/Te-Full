VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14t0 
   AutoRedraw      =   -1  'True
   Caption         =   "同仁介紹案源獎金明細表"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5505
   Begin VB.OptionButton Option1 
      Caption         =   "發放日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   870
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "未收款明細"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   780
      TabIndex        =   3
      Top             =   1410
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1260
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   2130
      Width           =   3105
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   870
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   870
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   330
      Left            =   2160
      TabIndex        =   5
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   330
      Left            =   3600
      TabIndex        =   6
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "上次發放日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   780
      TabIndex        =   7
      Top             =   420
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   3330
      X2              =   3540
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      X1              =   3330
      X2              =   3540
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc14t0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2014/1/7
Option Explicit

Dim adoquery As New ADODB.Recordset
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim lngPageNo As Long '頁數
Dim dblSumAmt As Double '獎金合計


Private Sub Command1_Click()
Dim strCon As String
   
On Error GoTo ErrHnd
   
   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   lngPageNo = 0
   Screen.MousePointer = vbHourglass

   If Option1(0).Value = True Then '可發放
      strCon = " and a0k36 between " & ACDate(DBDATE(MaskEdBox1)) & " and " & ACDate(DBDATE(MaskEdBox2))
   Else '未收款 '2014/1/22 modify by sonia 加a0k37 is null
      strCon = " and nvl(a0k36,0)=0 and a0k37 is null "
   End If
   'modify by sonia 2016/10/3 剔除F5639 北京寰華介紹案源,2016/9/1另有管制辦法
   strExc(0) = "select s1.st02 a0k34NM,a0k34,s2.st02 a0k20NM,a0k20,a0k01,a0k04,a0k17,sqldatet(cp05) cp05T,a0j01,a0j02,a0j09,a0j10,cpm03,cp01,cp02,cp03,cp04,cp10,a0j04" & _
               " from acc0k0,staff s1,staff s2,acc0j0,caseprogress,casepropertymap" & _
               " where a0k34=s1.st01(+) and a0k20=s2.st01(+)" & _
               " and a0k01=a0j13(+)" & _
               " and a0j01=cp09(+)" & _
               " and cp01=cpm01(+) and cp10=cpm02(+)" & _
               " and a0k34 is not null and a0k34<>'F5639'" & strCon & _
               " order by cp05,a0k20,a0k01"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoquery
         .MoveFirst
         Set xlsAnnuity = New Excel.Application
         Call SetExcelWorksheets
         PrintHead_Excel intCounter '頁首
         dblSumAmt = 0 '獎金合計
         Do While Not .EOF
            'Modify By Sindy 2015/9/30 Option1(0).Value = True : 第2頁切頁有誤 +  And intCounter <> 48 判斷
            If (Option1(0).Value = True And lngPageNo = 1 And intCounter Mod 43 = 0) Or _
               (Option1(0).Value = True And lngPageNo <> 1 And intCounter Mod 48 = 0 And intCounter <> 48) Or _
               (Option1(1).Value = True And intCounter Mod 48 = 0) Then
               '換頁
               intCounter = intCounter + 1
               wksAnnuity.Range("A" & intCounter).Select
               wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
               PrintHead_Excel intCounter '頁首
            End If
            '明細資料
            PrintData_Excel adoquery, intCounter
            .MoveNext
         Loop
         If Option1(0).Value = True Then
            '合計
            intCounter = intCounter + 1
            wksAnnuity.Range("G" & intCounter).Value = "合計："
            wksAnnuity.Range("H" & intCounter).Value = Format(dblSumAmt, DDollar)
         End If
      End With
   Else
      Screen.MousePointer = vbDefault
      MsgBox "無資料可供列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Sub
   End If
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   
   Screen.MousePointer = vbDefault
   
   If Option1(0).Value = True Then '可發放
      PUB_SaveLastDate Me.Name, "MaskEdBox3", ChangeTDateStringToTString(MaskEdBox1)
      PUB_SaveLastDate Me.Name, "MaskEdBox4", ChangeTDateStringToTString(MaskEdBox2)
      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox3"))
      MaskEdBox4.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
      MaskEdBox1.Mask = ""
      MaskEdBox2.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox2.Text = ""
      MaskEdBox1.Mask = DFormat
      MaskEdBox2.Mask = DFormat
   End If
   
   adoquery.Close
   Set adoquery = Nothing
   Exit Sub
   
ErrHnd:
   Screen.MousePointer = vbDefault
   adoquery.Close
   Set adoquery = Nothing
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub SetExcelWorksheets()
   xlsAnnuity.Visible = True
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 9
   wksAnnuity.Columns("B:B").ColumnWidth = 9
   wksAnnuity.Columns("C:C").ColumnWidth = 9
   wksAnnuity.Columns("D:D").ColumnWidth = 10
   wksAnnuity.Columns("E:E").ColumnWidth = 16
   wksAnnuity.Columns("F:F").ColumnWidth = 8
   wksAnnuity.Columns("G:G").ColumnWidth = 6
   wksAnnuity.Columns("H:H").ColumnWidth = 8
   wksAnnuity.Columns("I:I").ColumnWidth = 9
   intCounter = 1
End Sub

'表頭
Private Sub PrintHead_Excel(ByRef iRow As Integer)
Dim i As Integer, strTemp As String
   
   lngPageNo = lngPageNo + 1
   With wksAnnuity
      If Option1(0).Value = True Then
         .Range("E" & iRow).Value = "同仁介紹案源獎金明細表"
      Else
         .Range("E" & iRow).Value = "同仁介紹案源未收款明細表"
      End If
      '選取,儲存格合併,置中,粗體字
      strTemp = "A" & iRow & ":I" & iRow
      .Range(strTemp).Select
      With .Application.Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With .Application.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      .Application.Selection.Font.Bold = True
      
      If Option1(0).Value = True Then
         iRow = iRow + 1
         .Range("A" & iRow).Value = "列印人：" & strUserName
         .Range("D" & iRow).Value = "發放期間：" & MaskEdBox1 & " ~ " & MaskEdBox2
         .Range("H" & iRow).Value = "頁數：" & lngPageNo
         
         If lngPageNo = 1 Then
            iRow = iRow + 2
            '插入2列
            .Rows("2:2").Select
            .Application.Selection.Insert Shift:=xlDown
            .Application.Selection.Insert Shift:=xlDown
            .Range("A2:I3").Select
            '加框線
            With .Application.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Application.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Application.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Application.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Application.Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Application.Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            '合併儲存格
            .Range("A2:B2").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("C2:D2").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("E2:F2").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("G2:I2").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("A2:B2").Value = "受文者"
            .Range("C2:D2").Value = "發文者"
            .Range("E2:F2").Value = "發文單位"
            .Range("G2:I2").Value = "列印日期"
            '合併儲存格
            .Range("A3:B3").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("C3:D3").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("E3:F3").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            .Range("G3:I3").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            '將欄位拉高
            .Rows("3:3").RowHeight = 100 '66
            .Range("A3:I3").Select
            With .Application.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
            End With
            .Range("A3:B3").Value = "全體同仁"
            .Range("E3:F3").Value = "財務處"
            .Range("G3:I3").Value = "中華民國" & Left(strSrvDate(2), 3) & "年" & Mid(strSrvDate(2), 4, 2) & "月" & Right(strSrvDate(2), 2) & "日"
         End If
         
      Else
         iRow = iRow + 1
         .Range("A" & iRow).Value = "列印人：" & strUserName
         .Range("H" & iRow).Value = "列印日期："
         .Range("I" & iRow).Value = Format(strSrvDate(2), "###/##/##")
         iRow = iRow + 1
         .Range("H" & iRow).Value = "頁數："
         .Range("I" & iRow).Value = lngPageNo
         strTemp = "H" & iRow - 1 & ":H" & iRow
         .Range(strTemp).Select
         With .Application.Selection
            .HorizontalAlignment = xlRight '靠右
         End With
         strTemp = "I" & iRow & ":I" & iRow
         .Range(strTemp).Select
         With .Application.Selection
            .HorizontalAlignment = xlLeft '靠左
         End With
      End If
      
      iRow = iRow + 1
      .Range("A" & iRow).Value = "收文日期"
      .Range("B" & iRow).Value = "介紹人員"
      .Range("C" & iRow).Value = "智權人員"
      .Range("D" & iRow).Value = "本所案號"
      .Range("E" & iRow).Value = "收據抬頭"
      .Range("F" & iRow).Value = "金額"
      .Range("G" & iRow).Value = "點數"
      .Range("H" & iRow).Value = "介紹獎金"
      .Range("I" & iRow).Value = "案件性質"
      strTemp = "A" & iRow & ":I" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
      With .Application.Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlInsideVertical)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
   End With
End Sub

Private Sub PrintData_Excel(p_Rst As ADODB.Recordset, ByRef iRow As Integer)
Dim dblRate As Double
Dim strTemp As String
Dim adoTmp As ADODB.Recordset
Dim dbla1u07 As Double, dbla1u09 As Double
Dim dblAmt As Double
   
   iRow = iRow + 1
   With wksAnnuity
      .Range("A" & iRow).Value = "" & p_Rst.Fields("cp05T")
      .Range("B" & iRow).Value = "" & p_Rst.Fields("a0k34NM") '介紹人員
      .Range("C" & iRow).Value = "" & p_Rst.Fields("a0k20NM") '智權人員
      .Range("D" & iRow).Value = p_Rst.Fields("cp01") & "-" & p_Rst.Fields("cp02") & IIf(p_Rst.Fields("cp03") & p_Rst.Fields("cp04") = "000", "", "-" & p_Rst.Fields("cp03") & IIf(p_Rst.Fields("cp04") = "00", "", "-" & p_Rst.Fields("cp04")))
      .Range("E" & iRow).Value = "" & p_Rst.Fields("a0k04")
      '銷帳金額
      dbla1u07 = 0 '銷帳服務費
      dbla1u09 = 0 '銷帳規費
      strSql = "select nvl(sum(a1u07),0) a1u07,nvl(sum(a1u09),0) a1u09 from acc1u0 where a1u02='" & p_Rst.Fields("a0k01") & "' and a1u03='" & p_Rst.Fields("a0j01") & "'"
      intI = 1
      Set adoTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         dbla1u07 = adoTmp.Fields("a1u07")
         dbla1u09 = adoTmp.Fields("a1u09")
      End If
      'a0j09 = 服務費
      'a0j10 = 規費
      .Range("F" & iRow).Value = Format(Val("" & p_Rst.Fields("a0j09")) + Val("" & p_Rst.Fields("a0j10")) - dbla1u07 - dbla1u09, DDollar) '金額
      .Range("G" & iRow).Value = Round((Val("" & p_Rst.Fields("a0j09")) - dbla1u07) / 1000, 3) '點數
      .Range("I" & iRow).Value = "" & p_Rst.Fields("cpm03")
      '取得獎金的比率:
      '1.000台灣T商標案件性質4XX或6XX者 * 3%, 其它 * 5%
      '2.非T均為 * 3%
      If p_Rst.Fields("cp01") = "T" Then
         If "" & p_Rst.Fields("a0j04") = "000" And (Left(p_Rst.Fields("cp10"), 1) = "4" Or Left(p_Rst.Fields("cp10"), 1) = "6") Then
            dblRate = 3
         Else
            dblRate = 5
         End If
      Else
         dblRate = 3
      End If
      '介紹獎金
      dblAmt = 0
      If (Val("" & p_Rst.Fields("a0j09")) - dbla1u07) = 0 Or Option1(1).Value = True Then
         .Range("H" & iRow).Value = ""
      Else
         dblAmt = Round((Val("" & p_Rst.Fields("a0j09")) - dbla1u07) * dblRate / 100, 0)
         .Range("H" & iRow).Value = Format(dblAmt, DDollar)
         dblSumAmt = dblSumAmt + dblAmt '獎金合計
      End If
      strTemp = "A" & iRow & ":I" & iRow
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlInsideVertical)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
   End With
   Set adoTmp = Nothing
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5625
   Me.Height = 3240
   '改單線固定(調整大小不用再設定)
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next
   
'   '起日預設為系統日期一個月的30日,若無30日則為下一個日曆天
'   strDate = CompDate(1, -1, strSrvDate(1))
'   strDate = Left(strDate, 6) & "30"
'   If IsDate(ChangeWStringToWDateString(strDate)) = False Then
'      strDate = Left(strSrvDate(1), 6) & "01"
'   End If
'   MaskEdBox1.Text = CFDate(ACDate(strDate))
   MaskEdBox1.Mask = DFormat
'   '止日預設為系統日期當月的29日,若無29日則為前一個日曆天
'   strDate = Left(strSrvDate(1), 6) & "29"
'   If IsDate(ChangeWStringToWDateString(strDate)) = False Then
'      strDate = Left(strSrvDate(1), 6) & "28"
'   End If
'   MaskEdBox2.Text = CFDate(ACDate(strDate))
   MaskEdBox2.Mask = DFormat
   
   '上次發放日期
   If PUB_GetLastDate(Me.Name, "MaskEdBox3") <> "" Then
      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox3"))
   End If
   If PUB_GetLastDate(Me.Name, "MaskEdBox4") <> "" Then
      MaskEdBox4.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
   End If
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc14t0 = Nothing
End Sub

''*************************************************
'' 清除畫面
''
''*************************************************
'Private Sub FormClear()
'   MaskEdBox1.Mask = ""
'   MaskEdBox1.Text = ""
'   MaskEdBox1.Mask = DFormat
'   MaskEdBox2.Mask = ""
'   MaskEdBox2.Text = CFDate(ACDate(ServerDate))
'   MaskEdBox2.Mask = DFormat
'End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Option1(0).Value = True Then
      '日期檢查
      If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
         MsgBox "發放起始日期格式錯誤！", vbExclamation
         FormCheck = False
         MaskEdBox1.SetFocus
         Exit Function
      End If
   
      If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
         MsgBox "發放迄止日期格式錯誤！", vbExclamation
         FormCheck = False
         MaskEdBox2.SetFocus
         Exit Function
      End If
   End If
   FormCheck = True
End Function
