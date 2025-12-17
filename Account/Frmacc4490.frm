VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc4490 
   AutoRedraw      =   -1  'True
   Caption         =   "預算實績比較表"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      Top             =   150
      Width           =   3500
   End
   Begin VB.CommandButton cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
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
      Left            =   1380
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1410
      Width           =   2300
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   2040
      Width           =   3450
   End
   Begin VB.TextBox Text1 
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
      Left            =   2280
      TabIndex        =   2
      Text            =   "Y"
      Top             =   960
      Width           =   495
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   570
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   150
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   2460
      Width           =   2300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   10
      Top             =   210
      Width           =   675
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   360
      TabIndex        =   8
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "是否含未超支科目"
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
      TabIndex        =   6
      Top             =   960
      Width           =   1932
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年月"
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
      Left            =   360
      TabIndex        =   5
      Top             =   570
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc4490"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc040 As New ADODB.Recordset
Public adoaccrpt410 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim intAutoNo As Integer
Dim dllaccrpt410 As Object
Dim douMonthAmt As Double
Dim douYearAmt As Double
Dim stra0403 As String  '2014/1/24 add by sonia
Dim strPrinter As String 'Added by Lydia 2016/02/16 加印表機選項
Dim strFieldN(), intWidth() 'Add by Amy 2016/08/11
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/4/27


'Add by Sindy 2020/4/27
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/4/27

'Add by Amy 2016/08/11
Private Sub Cmd_Excel_Click()
    
    If FormCheck = False Then
        Exit Sub
    End If
    
    Call SetCompN 'Add by Sindy 2020/4/27
    
    'add by sonia 2017/12/8 檢查是否有未過帳傳票
    'Modify by Sindy 2020/4/27
    'If CheckAX210(Text6, Replace(MaskEdBox1.Text, "/", ""), Replace(MaskEdBox1.Text, "/", "")) = True Then
    If CheckAX210(strCmp, Replace(MaskEdBox1.Text, "/", ""), Replace(MaskEdBox1.Text, "/", "")) = True Then
        Exit Sub
    End If
    'end 2017/12/8
    
    Screen.MousePointer = vbHourglass

    Accrpt410Delete
    ProduceData
    ExcelSave
    
    Screen.MousePointer = vbDefault
    FormClear
    'Mark by Amy 2020/02/17 不顯示請更換大表-瑞婷 2018/11/05 mail to 秀玲
    'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub ExcelSave()
    Dim xlsAnnuity As New Excel.Application
    Dim wksAnnuity As New Worksheet
    Dim strFileName As String, strQ As String, strTemp As String
    Dim strCol As String, strFind As String, strReplace As String
    Dim bolFormula As Boolean
    Dim ii As Integer, intField As Integer, intCounter As Integer, intTitleRow As Integer
  
    ReDim strFieldN(9)
    ReDim intWidth(9)

    strFieldN = Array("科目代碼", "會計科目", "當月預算", "當月實績", "當月差額", "當月佔經費(%)", "累計預算", "累計實績", "累計差額", "累計佔經費(%)")
    'Modified by Lydia 2017/06/06 改變科目代碼、會計科目的欄寬
    intWidth = Array(7.5, 21, 12, 12, 12, 9.5, 12, 12, 12, 9.5)
                                
On Error GoTo ErrHnd
    
    intField = 65:  intCounter = 1
    'Modify By Sindy 2020/5/22
    'strFileName = Val(Replace(MaskEdBox1.Text, "/", "")) & "預算實績比較表" & ServerDate & MsgText(43)
    strFileName = Val(Replace(MaskEdBox1.Text, "/", "")) & "預算實績比較表" & ServerDate & "-" & Replace(strCmpN, "/", "") & MsgText(43)
    '2020/5/22 END
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
             MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAnnuity.Workbooks.add
    Set wksAnnuity = xlsAnnuity.Worksheets(1)
    
    With wksAnnuity
        '***表頭設定***
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = "預算實績比較表"
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).HorizontalAlignment = xlCenter
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).VerticalAlignment = xlCenter
        intCounter = intCounter + 1
        
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField - 1) & intCounter).Value = "公司別："
        'Modify by Sindy 2020/4/27
        '.Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = Trim(Text6) & IIf(Text7 = "", "台一　專利商標/智權", Text7)
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = Trim(strCmp) & strCmpN
        intCounter = intCounter + 1
        
        .Range(Chr(intField) & intCounter).Value = "列印人員：" & strUserName
        .Range(Chr(UBound(strFieldN) + intField - 1) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField - 1) & intCounter).Value = "年　月："
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = MaskEdBox1.Text
        intCounter = intCounter + 2
         
         '公式顯示
        .Range(Chr(intField + GetValue("當月預算")) & intCounter).Value = "A"
        .Range(Chr(intField + GetValue("當月實績")) & intCounter).Value = "B"
        .Range(Chr(intField + GetValue("當月差額")) & intCounter).Value = "A-B"
        .Range(Chr(intField + GetValue("當月佔經費(%)")) & intCounter).Value = "B/E"
        .Range(Chr(intField + GetValue("累計預算")) & intCounter).Value = "C"
        .Range(Chr(intField + GetValue("累計實績")) & intCounter).Value = "D"
        .Range(Chr(intField + GetValue("累計差額")) & intCounter).Value = "C-D"
        .Range(Chr(intField + GetValue("累計佔經費(%)")) & intCounter).Value = "D/F"
        .Range(Chr(intField + GetValue("當月預算")) & intCounter & ":" & Chr(intField + GetValue("累計佔經費(%)")) & intCounter).HorizontalAlignment = xlCenter
        
        intCounter = intCounter + 1
        For ii = 0 To UBound(strFieldN)
            .Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = intWidth(ii)
            .Range(Chr(intField + ii) & intCounter).Value = strFieldN(ii)
            .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
        Next ii
        Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(GetValue("會計科目") + intField) & intCounter)
        Call SetExcelLine(0, wksAnnuity, Chr(GetValue("當月預算") + intField) & intCounter & ":" & Chr(GetValue("當月佔經費(%)") + intField) & intCounter)
        Call SetExcelLine(0, wksAnnuity, Chr(GetValue("累計預算") + intField) & intCounter & ":" & Chr(GetValue("累計佔經費(%)") + intField) & intCounter)
       
        intTitleRow = intCounter: intCounter = intCounter + 1
        '列印資料
        If adoaccrpt410.State = adStateOpen Then adoaccrpt410.Close
        strQ = "Select * From accrpt410 Order by R41012"
        adoaccrpt410.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
        Do While adoaccrpt410.EOF = False
            For ii = LBound(strFieldN) To UBound(strFieldN)
                bolFormula = True
                Select Case ii
                    Case GetValue("當月差額")
                        strTemp = "=" & Chr(GetValue("當月預算") + intField) & intCounter & "-" & Chr(GetValue("當月實績") + intField) & intCounter
                    Case GetValue("當月佔經費(%)")
                        strTemp = "=Round(" & Chr(GetValue("當月實績") + intField) & intCounter & "/$" & Chr(UBound(strFieldN) + intField + 1) & "$1" & "*100,2)"
                    Case GetValue("累計差額")
                        strTemp = "=" & Chr(GetValue("累計預算") + intField) & intCounter & "-" & Chr(GetValue("累計實績") + intField) & intCounter
                    Case GetValue("累計佔經費(%)")
                        strTemp = "=Round(" & Chr(GetValue("累計實績") + intField) & intCounter & "/$" & Chr(UBound(strFieldN) + intField + 1) & "$2" & "*100,2)"
                    Case Else
                        bolFormula = False
                        strTemp = "" & adoaccrpt410.Fields(ii + 1)
                End Select
                If ii <> GetValue("科目代號") And ii <> GetValue("會計科目") Then
                    wksAnnuity.Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
                ElseIf ii = GetValue("科目代號") Then
                    wksAnnuity.Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "@"
                End If
                If bolFormula = True Then
                    wksAnnuity.Range(Chr(intField + ii) & intCounter).Formula = strTemp
                Else
                    wksAnnuity.Range(Chr(intField + ii) & intCounter).Value = strTemp
                End If
            Next ii
            adoaccrpt410.MoveNext
            intCounter = intCounter + 1
        Loop
        adoaccrpt410.Close
        Call SetExcelLine(2, wksAnnuity, Chr(intField) & intTitleRow + 1 & ":" & Chr(GetValue("會計科目") + intField) & intCounter - 1)
        Call SetExcelLine(2, wksAnnuity, Chr(GetValue("當月預算") + intField) & intTitleRow + 1 & ":" & Chr(GetValue("當月佔經費(%)") + intField) & intCounter - 1)
        Call SetExcelLine(2, wksAnnuity, Chr(GetValue("累計預算") + intField) & intTitleRow + 1 & ":" & Chr(GetValue("累計佔經費(%)") + intField) & intCounter - 1)
        
        '加總
        For ii = GetValue("會計科目") To GetValue("累計佔經費(%)")
            If ii = GetValue("會計科目") Then
                wksAnnuity.Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlLeft
                wksAnnuity.Range(Chr(intField + ii) & intCounter).Value = "合　　計"
            Else
                wksAnnuity.Range(Chr(intField + ii) & intCounter).Formula = "=Sum(" & Chr(intField + ii) & intTitleRow + 1 & ":" & Chr(intField + ii) & intCounter - 1 & ")"
            End If
        Next ii
        Call SetExcelLine(3, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(GetValue("會計科目") + intField) & intCounter)
        Call SetExcelLine(1, wksAnnuity, Chr(GetValue("當月預算") + intField) & intCounter & ":" & Chr(GetValue("當月佔經費(%)") + intField) & intCounter)
        Call SetExcelLine(1, wksAnnuity, Chr(GetValue("累計預算") + intField) & intCounter & ":" & Chr(GetValue("累計佔經費(%)") + intField) & intCounter)
        
        intCounter = intCounter + 2
        '備註
        wksAnnuity.Range(Chr(intField + 1) & intCounter).Value = "當月營業支出(E)"
        wksAnnuity.Range(Chr(intField + 2) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
        wksAnnuity.Range(Chr(intField + 2) & intCounter).HorizontalAlignment = xlRight
        wksAnnuity.Range(Chr(intField + 2) & intCounter).Value = douMonthAmt
        strFind = "$" & Chr(UBound(strFieldN) + intField + 1) & "$1"
        strReplace = "$" & Chr(intField + 2) & "$" & intCounter
        wksAnnuity.Columns(Chr(GetValue("當月佔經費(%)") + intField)).Replace what:=strFind, Replacement:=strReplace, LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
        intCounter = intCounter + 1
        
        strFind = "$" & Chr(UBound(strFieldN) + intField + 1) & "$2"
        strReplace = "$" & Chr(intField + 2) & "$" & intCounter
        wksAnnuity.Range(Chr(intField + 1) & intCounter).Value = "累計營業支出(F)"
        wksAnnuity.Range(Chr(intField + 2) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
        wksAnnuity.Range(Chr(intField + 2) & intCounter).HorizontalAlignment = xlRight
        wksAnnuity.Range(Chr(intField + 2) & intCounter).Value = douYearAmt
        wksAnnuity.Columns(Chr(GetValue("累計佔經費(%)") + intField)).Replace what:=strFind, Replacement:=strReplace, LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
        
        For ii = GetValue("當月差額") To GetValue("當月佔經費(%)")
            wksAnnuity.Columns(Chr(intField + ii)).Replace what:="當月", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
        Next ii
        For ii = GetValue("累計差額") To GetValue("累計佔經費(%)")
            wksAnnuity.Columns(Chr(intField + ii)).Replace what:="累計", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
        Next ii
  
    End With
    'Excel字型大小設定
    With wksAnnuity.Range(Chr(intField) & "1:" & Chr(UBound(strFieldN) + intField) & intCounter)
        .Font.Name = "新細明體"
        .Font.Size = 11
    End With
    With wksAnnuity
        .PageSetup.PaperSize = 9 '設A4
        .PageSetup.PrintTitleRows = "$1:$" & intTitleRow
        .PageSetup.Orientation = xlLandscape '橫印
        'Modified by Lydia 2017/06/06 橫印改直印 xlLandscape => xlPortrait
        '.PageSetup.Orientation = xlLandscape '橫印
        .PageSetup.Orientation = xlPortrait '直印
        'Added by Lydia 2017/06/06 縮放比例為78%,列印頁面水平置中
        .PageSetup.Zoom = 78
        .PageSetup.CenterHorizontally = True
        'end 2017/06/06
        
        .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.3) '上
        .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.3) '下
        'Modified by Lydia 2017/06/06 為了印整張A4,左右邊界改為0
        '.PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.5) '左邊界
        '.PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.5) '右邊界
        .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0) '左邊界
        .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0) '右邊界
    End With
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set wksAnnuity = Nothing
    Set xlsAnnuity = Nothing
    MsgBox "檔案已產生~"
    Exit Sub
   
ErrHnd:
    If adoaccrpt410.State = adStateOpen Then adoaccrpt410.Close
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set wksAnnuity = Nothing
    Set xlsAnnuity = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub SetExcelLine(intChoose As Integer, ByRef m_Xls As Worksheet, strField As String)

    With m_Xls.Range(strField)
        Select Case intChoose
            Case 0 '抬頭
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
            Case 1 '最後合計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeBottom).Weight = xlThick '粗線
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
            Case 2 '資料內容
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlHairline
            Case 3 '最後合計-合計字樣欄
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeBottom).Weight = xlThick '粗線
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
        End Select
    End With
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'end 2016/08/11

'Private Sub Command1_Click()
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'
'   'add by sonia 2017/12/8 檢查是否有未過帳傳票
'   If CheckAX210(Text6, Replace(MaskEdBox1.Text, "/", ""), Replace(MaskEdBox1.Text, "/", "")) = True Then
'       Exit Sub
'   End If
'   'end 2017/12/8
'
'   Screen.MousePointer = vbHourglass
'  'Added by Lydia 2016/02/16 加印表機選項
'   PUB_SetOsDefaultPrinter Combo1
'
'   Accrpt410Delete
'   ProduceData
'   '2014/1/24 modify by sonia
'   'dllaccrpt410.Acc4490 ReportTitle(410), Text6, Text7, MaskEdBox1.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'   dllaccrpt410.Acc4490 ReportTitle(410), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), MaskEdBox1.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'
'  'Added by Lydia 2016/02/16 加印表機選項
'   PUB_SetOsDefaultPrinter strPrinter
'
'   Screen.MousePointer = vbDefault
'   FormClear
'   'Mark by Amy 2020/02/17 不顯示請更換大表-瑞婷 2018/11/05 mail to 秀玲
'   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   'Mark by Amy 2020/02/17 不顯示請更換大表-瑞婷 2018/11/05 mail to 秀玲
'   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
'   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   'Modified by Lydia 2016/02/16
   'Me.Height = 2400
   Me.Height = 2310 '3000
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/4/27 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/4/27
   
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   'Mark by Amy 2020/02/17 不顯示請更換大表-瑞婷 2018/11/05 mail to 秀玲
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt410 = CreateObject("AccReport.ReportSelect")
   'Added by Lydia 2016/02/16 加印表機選項
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Added by Lydia 2016/02/16 程式結束後要還原為預設印表機
   PUB_SetOsDefaultPrinter strPrinter
   
   Set dllaccrpt410 = Nothing
   Set Frmacc4490 = Nothing
End Sub


Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Sindy 2020/4/27 公司別改下拉
'Private Sub Text6_Change()
'   '2014/1/24 modify by sonia
'   'If Text6 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   'Text7 = A0802Query(Text6)
'   Select Case Text6
'      Case "1"
'         Text7 = A0802Query(Text6)
'      Case "2"
'         Text7 = A0802Query("J")
'      Case ""
'         Text7 = "台一　專利商標/智權"
'   End Select
'   '2014/1/24 end
'End Sub
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
''2014/1/24 add by sonia
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/1/24 end

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim douPreAmount, douDebit, douCredit As Double
Dim strSql As String

On Error GoTo Checking
   SumShow
   intAutoNo = 0
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt410.CursorLocation = adUseClient
   adoaccrpt410.Open "select * from accrpt410 ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/1/24 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Sindy 2020/4/27
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If strCmp <> "" Then
      If InStr(strCmp, "+") > 0 Then
         adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "')) order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & strCmp & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   '2020/4/27 END
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/1/24 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      adoaccrpt410.AddNew
      adoaccrpt410.Fields("r41001").Value = strUserNum
      adoaccrpt410.Fields("r41002").Value = adoacc010.Fields("a0101").Value
      If IsNull(adoacc010.Fields("a0102").Value) Then
         adoaccrpt410.Fields("r41003").Value = Null
      Else
         adoaccrpt410.Fields("r41003").Value = adoacc010.Fields("a0102").Value
      End If
'------------------------------------------------
' 當月金額
'------------------------------------------------
      adoacc040.CursorLocation = adUseClient
      strSql = MsgText(601)
      stra0403 = MsgText(601) '2014/1/24 add by sonia
      
      'Modify by Sindy 2020/4/27
'      If Text6 <> MsgText(601) Then
'         '2014/1/24 modify by sonia
'         'strSql = " and a0403 = '" & Text6 & "'"
'         strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'         stra0403 = " a0403 = '" & IIf(Text6 = "2", "J", "1") & "' and"
'         '2014/1/24 end
'      End If
      If strCmp <> MsgText(601) Then
         '2014/1/24 modify by sonia
         'strSql = " and a0403 = '" & Text6 & "'"
         If InStr(strCmp, "+") > 0 Then
            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
            stra0403 = " a0403 In ('" & Replace(strCmp, "+", "','") & "') and"
         Else
            strSql = " and a0403 = '" & strCmp & "'"
            stra0403 = " a0403 = '" & strCmp & "' and"
         End If
         '2014/1/24 end
      End If
      '2020/4/27 END
      
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
         strSql = strSql & " and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & ""
      End If
      adoacc040.Open "select sum(a0408), sum(a0409) from acc040 where a0405 = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(1).Value) Then
            adoaccrpt410.Fields("r41004").Value = 0
         Else
            adoaccrpt410.Fields("r41004").Value = Val(adoacc040.Fields(1).Value)
         End If
         If IsNull(adoacc040.Fields(0).Value) Then
            adoaccrpt410.Fields("r41005").Value = 0
         Else
            adoaccrpt410.Fields("r41005").Value = Val(adoacc040.Fields(0).Value)
         End If
         adoaccrpt410.Fields("r41006").Value = Val(adoaccrpt410.Fields("r41004").Value) - Val(adoaccrpt410.Fields("r41005").Value)
         If douMonthAmt <> 0 Then
            adoaccrpt410.Fields("r41007").Value = Val(Format(Val(adoaccrpt410.Fields("r41005").Value) / douMonthAmt * 100, FAmount))
         Else
            adoaccrpt410.Fields("r41007").Value = 0
         End If
      Else
         adoaccrpt410.Fields("r41004").Value = 0
         adoaccrpt410.Fields("r41005").Value = 0
         adoaccrpt410.Fields("r41006").Value = 0
         adoaccrpt410.Fields("r41007").Value = 0
      End If
      adoacc040.Close
'------------------------------------------------
' 累計金額
'------------------------------------------------
      adoaccsum.CursorLocation = adUseClient
      strSql = MsgText(601)
      
      'Modify by Sindy 2020/4/27
'      If Text6 <> MsgText(601) Then
'         '2014/1/24 modify by sonia
'         'strSql = " and a0403 = '" & Text6 & "'"
'         strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'         '2014/1/24 end
'      End If
      If strCmp <> MsgText(601) Then
         '2014/1/24 modify by sonia
         'strSql = " and a0403 = '" & Text6 & "'"
         If InStr(strCmp, "+") > 0 Then
            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
         Else
            strSql = " and a0403 = '" & strCmp & "'"
         End If
         '2014/1/24 end
      End If
      '2020/4/27 END
      
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
         strSql = strSql & " and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 <= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & ""
      End If
      adoaccsum.Open "select sum(a0408), sum(a0409) from acc040 where a0405 = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(1).Value) Then
            adoaccrpt410.Fields("r41008").Value = 0
         Else
            adoaccrpt410.Fields("r41008").Value = Val(adoaccsum.Fields(1).Value)
         End If
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt410.Fields("r41009").Value = 0
         Else
            adoaccrpt410.Fields("r41009").Value = Val(adoaccsum.Fields(0).Value)
         End If
         adoaccrpt410.Fields("r41010").Value = Val(adoaccrpt410.Fields("r41008").Value) - Val(adoaccrpt410.Fields("r41009").Value)
         If douYearAmt <> 0 Then
            adoaccrpt410.Fields("r41011").Value = Val(Format(Val(adoaccrpt410.Fields("r41009").Value) / douYearAmt * 100, FAmount))
         Else
            adoaccrpt410.Fields("r41011").Value = 0
         End If
      Else
         adoaccrpt410.Fields("r41008").Value = 0
         adoaccrpt410.Fields("r41009").Value = 0
         adoaccrpt410.Fields("r41010").Value = 0
         adoaccrpt410.Fields("r41011").Value = 0
      End If
      adoaccsum.Close
      AutoNoSave
      adoaccrpt410.UpdateBatch
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt410.Close
   If Text1 = MsgText(603) Then
      adoTaie.Execute "delete from accrpt410 where r41006 >= 0"
   End If
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt410Delete()
    adoTaie.Execute "delete from accrpt410 "
End Sub

'*************************************************
' 自動編號存入 r41012 欄位
'
'*************************************************
Private Sub AutoNoSave()
   intAutoNo = intAutoNo + 1
   adoaccrpt410.Fields("r41012").Value = intAutoNo
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
'   Text6 = ""
'   Text7 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   Text1 = ""
'   Text6.SetFocus
   
   'Add By Sindy 2020/4/27
   CboCmp.ListIndex = -1
   CboCmp.SetFocus
   '2020/4/27 END
End Sub

'*************************************************
'  顯示欄位資料(傳票合計)
'
'*************************************************
Public Sub SumShow()
   douMonthAmt = 0
   douYearAmt = 0
   adoaccsum.CursorLocation = adUseClient
   '2014/1/24 modify by sonia 加公司別條件
   'adoaccsum.Open "select sum(a0408) from acc040 where substr(a0405, 1, 1) in ('6', '8') and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a0408) from acc040 where " & stra0403 & " substr(a0405, 1, 1) in ('6', '8') and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         douMonthAmt = adoaccsum.Fields(0).Value
      End If
   End If
   adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   '2014/1/24 modify by sonia 加公司別條件
   'adoaccsum.Open "select sum(a0408) from acc040 where substr(a0405, 1, 1) in ('6', '8') and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 <= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a0408) from acc040 where " & stra0403 & " substr(a0405, 1, 1) in ('6', '8') and a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 <= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         douYearAmt = adoaccsum.Fields(0).Value
      End If
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
Dim bolCancel As Boolean

   'Add by Sindy 2020/4/27 +公司別判斷
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bolCancel)
      If bolCancel = True Then
          Exit Function
      End If
   End If
   'end 2020/4/27
   
   If MaskEdBox1.Text = Mid(MsgText(29), 1, 6) Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Function
   End If
   
   FormCheck = True
End Function

'Added by Lydia 2016/02/17
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox1.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox1.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox1.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox1.SetFocus
         End If
      End If
   End If
End Sub
