VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc4470 
   AutoRedraw      =   -1  'True
   Caption         =   "綜合損益比較表"
   ClientHeight    =   2280
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   210
      Width           =   3500
   End
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   1680
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2346
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   10
      Top             =   1200
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "本期年月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   8
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "上期年月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   6
      Top             =   210
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc4470"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc021 As New ADODB.Recordset
Public adoaccrpt408 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc040L As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim dllaccrpt408 As Object
Dim intAutoNo As Integer
'Add by Amy 2018/03/01
Dim strF1(), strF2(), intWidth()
Dim i As Integer
Dim intField As Integer, intCounter As Integer, intTitle As Integer
'Add by Amy 2020/04/16
Dim strCmp As String, strCmpN As String

'Add by Amy 2020/04/16
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
'end 2020/04/16

Private Sub Command1_Click()
   'Add by Amy 2020/04/16
   Dim bolShowMsg As Boolean
   
   'Modify by Amy 2020/04/16
   If FormCheck(bolShowMsg) = False Then
      If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
   'end 2020/04/16
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Call SetCompN 'Add by Amy 2020/04/16
   Accrpt408Delete
   ProduceData
   '2014/2/20 modify by sonia
   'dllaccrpt408.Acc4470 ReportTitle(408), Text6, Text7, MaskEdBox3.Text, MaskEdBox4.Text, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   'Modify by Amy 2020/04/16 公司別改抓變數
   'dllaccrpt408.Acc4470 ReportTitle(408), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), MaskEdBox3.Text, MaskEdBox4.Text, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   dllaccrpt408.Acc4470 ReportTitle(408), strCmp, strCmpN, MaskEdBox3.Text, MaskEdBox4.Text, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
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
   Me.Height = 2650
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/16 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/16
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt408 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt408 = Nothing
   Set Frmacc4470 = Nothing
End Sub

'Modify by Amy 2020/04/16 公司別改下拉
'Private Sub Text6_Change()
'   '2014/2/20 modify by sonia
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
'   '2014/2/20 end
'End Sub
'
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'
''2014/2/20 add by sonia
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/2/20 end
'end 2020/04/16

'Add by Amy 2018/03/01
Private Sub Cmd_Excel_Click()
    Dim strQ As String
    Dim bolShowMsg As Boolean 'Add by Amy 2020/04/16
    
    'Modify by Amy 2020/04/16
    If FormCheck(bolShowMsg) = False Then
        If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
    'end 2020/04/16
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Call SetCompN 'Add by Amy 2020/04/16
    Accrpt408Delete
    ProduceData
    strQ = "Select R40803 as AccN,Nvl(R40804,0),Nvl(R40805,0),Nvl(R40806,0),Nvl(R40807,0),Nvl(R40808,0),R40802 as AccNo " & _
                "From Accrpt408 Where R40801='" & strUserNum & "' And R40809=1 Order by R40802"
    If adoaccrpt408.State <> adStateClosed Then adoaccrpt408.Close
    adoaccrpt408.CursorLocation = adUseClient
    adoaccrpt408.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If adoaccrpt408.RecordCount = 0 Then
        MsgBox "無資料產生！"
    Else
        If SaveExcel = True Then FormClear
    End If
    Screen.MousePointer = vbDefault
    StatusClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Function SaveExcel() As Boolean
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim j As Integer
    Dim xlsFileName As String, stTemp(1) As String
    Dim stStartC As String, stPAndL As String '會計科目年度合計起始欄/營業損益列
    Dim intStartR As Integer, stStyle As String '合計起始列/格式
    
On Error GoTo ErrHand
    'Modify by Amy 2020/04/16 公司名稱改抓變數 原:IIf(Text7 = "", "台一智權", Text7)
    'Modify by Amy 2024/06/20 檔名同表單名,因太多損益表名稱一樣
    xlsFileName = Replace(strCmpN, "/", " ") & Trim(Replace(ReportTitle(408), "*", "")) & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & xlsFileName
    End If
    
    ReDim strF1(4)
    ReDim strF2(5)
    ReDim intWidth(5)
    strF1 = Array("", "", "本期", "", "上期")
    strF2 = Array("會計科目", "實際數", "預算數", "差異數", "上期實際數", "上期差異數")
    intWidth = Array(15, 13, 13, 13, 13, 13, 13)
   
    intField = 65:  intCounter = 1
    'Modify by Amy 2019/12/06 原:1
    xlsAgentPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAgentPoint.Workbooks.add
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    'xlsAgentPoint.Visible = True
    Call SetField(wksrpt)
    
    intStartR = intCounter
    adoaccrpt408.MoveFirst
    Do While adoaccrpt408.EOF = False
        '合計
        If InStr("" & adoaccrpt408.Fields("AccNo"), "T") > 0 Then
            '記錄營業收入 or 營業支出 or 營業外收入
            If "" & adoaccrpt408.Fields("AccNo") = "4T" Or "" & adoaccrpt408.Fields("AccNo") = "6T" Or _
               "" & adoaccrpt408.Fields("AccNo") = "71T" Or "" & adoaccrpt408.Fields("AccNo") = "72T" Then
                If "" & adoaccrpt408.Fields("AccNo") = "71T" Then
                    stPAndL = stPAndL & "+" & Chr(intField + LBound(strF2) + 1) & intCounter
                Else
                    stPAndL = stPAndL & "-" & Chr(intField + LBound(strF2) + 1) & intCounter
                End If
                '若沒資料合計也需顯示-婧瑄
                If intStartR < intCounter Then
                    stTemp(1) = "Sum(" & Chr(intField + LBound(strF2) + 1) & intStartR & ":" & Chr(intField + LBound(strF2) + 1) & intCounter - 1 & ")"
                Else
                    stTemp(1) = "0"
                End If
                stStyle = "xlContinuous" '單線
            '記錄營業損益/營業外支出/本期損益
            ElseIf "" & adoaccrpt408.Fields("AccNo") = "6ZT" Or "" & adoaccrpt408.Fields("AccNo") = "ZZT" Then
                stTemp(1) = Mid(stPAndL, 2)
                stPAndL = "+" & Chr(intField + LBound(strF2) + 1) & intCounter
                stStyle = "xlDouble" '雙線
            End If
            wksrpt.Range(Chr(intField + LBound(strF2)) & intCounter).Value = "" & adoaccrpt408.Fields("AccN")
            wksrpt.Range(Chr(intField + LBound(strF2)) & intCounter).Font.Bold = True
            wksrpt.Range(Chr(intField + LBound(strF2)) & intCounter).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
            wksrpt.Range(Chr(intField + LBound(strF2)) & intCounter).Interior.tintandshade = 0.5 '設深淺
            For i = LBound(strF2) + 1 To UBound(strF2)
                If InStr("" & adoaccrpt408.Fields("AccNo"), "T") > 0 Then
                    stTemp(0) = Replace(stTemp(1), Chr(intField + LBound(strF2) + 1), Chr(intField + i))
                End If
                wksrpt.Range(Chr(intField + i) & intCounter).Value = IIf(stTemp(0) <> "0", "=", "") & stTemp(0)
                wksrpt.Range(Chr(intField + i) & intCounter).NumberFormatLocal = "#,##0.00_ ;[紅色]-#,##0.00 "
                wksrpt.Range(Chr(intField + i) & intCounter).Font.Bold = True
                wksrpt.Range(Chr(intField + i) & intCounter).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
                wksrpt.Range(Chr(intField + i) & intCounter).Interior.tintandshade = 0.5 '設深淺
            Next i
            If stStyle <> MsgText(601) Then
                If stStyle = "xlDouble" Then
                    wksrpt.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strF2)) & intCounter).Borders(xlEdgeBottom).LineStyle = xlDouble
                Else
                    wksrpt.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strF2)) & intCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous
                End If
                stStyle = ""
            End If
            intStartR = intCounter + 1
        '內容
        Else
            For i = LBound(strF2) To UBound(strF2)
                If i = GetValue("會計科目") Then
                    stTemp(0) = "" & adoaccrpt408.Fields("AccN")
                'Add by Amy 2019/12/06 改用公式
                ElseIf InStr(strF2(i), "差異數") > 0 Then
                    '本期差異數(第一個差異數)
                    If strF2(i) = "差異數" Then
                        stTemp(0) = "=" & Chr(intField + GetValue("實際數")) & intCounter & "-" & Chr(intField + GetValue("預算數")) & intCounter
                    '上期差異數
                    Else
                        stTemp(0) = "=" & Chr(intField + GetValue("實際數")) & intCounter & "-" & Chr(intField + GetValue("上期實際數")) & intCounter
                    End If
                Else
                    stTemp(0) = Val("" & adoaccrpt408.Fields(i))
                End If
                If i <> GetValue("會計科目") Then
                    wksrpt.Range(Chr(intField + i) & intCounter).NumberFormatLocal = "#,##0.00_ ;[紅色]-#,##0.00 "
                End If
                'end 2019/12/06
                wksrpt.Range(Chr(intField + i) & intCounter).Value = stTemp(0)
            Next i
        End If
        '合計且非支出項目中間多空一行
        If InStr("" & adoaccrpt408.Fields("AccNo"), "T") > 0 And "" & adoaccrpt408.Fields("AccNo") <> "6T" And "" & adoaccrpt408.Fields("AccNo") <> "72T" Then
            intCounter = intCounter + 2
        Else
            intCounter = intCounter + 1
        End If
        adoaccrpt408.MoveNext
    Loop
    
    Call SetField(wksrpt, True) '修改顯示欄名
    wksrpt.Range(Chr(intField) & intTitle + 1 & ":" & Chr(intField + UBound(strF2)) & intCounter).Font.Size = 10
    wksrpt.PageSetup.PaperSize = 9 '設定紙張 A4
    wksrpt.PageSetup.Orientation = xlPortrait '直印
    wksrpt.PageSetup.PrintTitleRows = "$1:$" & intTitle '表頭保留列
    wksrpt.PageSetup.PrintGridlines = True '列印格線
    wksrpt.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    '邊界
    wksrpt.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0.5)
    wksrpt.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0.5)
    wksrpt.PageSetup.TopMargin = xlsAgentPoint.InchesToPoints(0.3)
    wksrpt.PageSetup.BottomMargin = xlsAgentPoint.InchesToPoints(0.3)
    wksrpt.PageSetup.HeaderMargin = xlsAgentPoint.InchesToPoints(0.5)
    wksrpt.PageSetup.FooterMargin = xlsAgentPoint.InchesToPoints(0.5)
    wksrpt.PageSetup.Zoom = 100 '縮放比例
    
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    SaveExcel = True
    MsgBox "Excel已產生！"
Exit Function
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
End Function

Private Sub SetField(ByRef Wks As Worksheet, Optional ByVal bolIsLast As Boolean = False)
    If bolIsLast = False Then
        Wks.Range(Chr(intField) & intCounter).Value = Replace(ReportTitle(408), "*", "")
        Wks.Range(Chr(intField) & intCounter).Font.Size = 12
        Wks.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strF2)) & intCounter).HorizontalAlignment = xlCenter
        Wks.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strF2)) & intCounter).MergeCells = True
        intCounter = intCounter + 1
        Wks.Range(Chr(intField + 1) & intCounter).Value = "公司別："
        'Modify by Amy 2020/04/16 公司別改抓變數
        'Wks.Range(Chr(intField + 2) & intCounter).Value = IIf(Text6 = "2", "J", Text6) & "　" & IIf(Text7 = "", "台一　專利商標/智權", Text7)
        Wks.Range(Chr(intField + 2) & intCounter).Value = strCmp & "　" & strCmpN
        intCounter = intCounter + 1
        Wks.Range(Chr(intField + 1) & intCounter).Value = "上期月份："
        Wks.Range(Chr(intField + 2) & intCounter).Value = MaskEdBox3.Text & "~" & MaskEdBox4.Text
        intCounter = intCounter + 1
        Wks.Range(Chr(intField + 1) & intCounter).Value = "本期月份："
        Wks.Range(Chr(intField + 2) & intCounter).Value = MaskEdBox1.Text & "~" & MaskEdBox2.Text
        intCounter = intCounter + 1
        Wks.Range(Chr(intField) & intCounter).Value = "列印人員："
        Wks.Range(Chr(intField + 1) & intCounter).Value = StaffQuery(strUserNum)
        Wks.Range(Chr(intField + UBound(strF2) - 2) & intCounter).Value = "列印日期："
        Wks.Range(Chr(intField + UBound(strF2) - 1) & intCounter).Value = CFDate(ACDate(ServerDate))
        intCounter = intCounter + 2
        For i = LBound(strF1) To UBound(strF1)
            If strF1(i) <> MsgText(601) Then
                Wks.Range(Chr(intField + i) & intCounter).Value = strF1(i)
                Wks.Range(Chr(intField + i) & intCounter).Font.Bold = True
                If strF1(i) = "本期" Then
                    Wks.Range(Chr(intField + i - 1) & intCounter & ":" & Chr(intField + i + 1) & intCounter).HorizontalAlignment = xlCenter
                    Wks.Range(Chr(intField + i - 1) & intCounter & ":" & Chr(intField + i + 1) & intCounter).MergeCells = True
                Else
                    Wks.Range(Chr(intField + UBound(strF1)) & intCounter & ":" & Chr(intField + UBound(strF2)) & intCounter).HorizontalAlignment = xlCenter
                    Wks.Range(Chr(intField + UBound(strF1)) & intCounter & ":" & Chr(intField + UBound(strF2)) & intCounter).MergeCells = True
                End If
            End If
        Next i
        intCounter = intCounter + 1
    End If
    For i = LBound(strF2) To UBound(strF2)
        If bolIsLast = True Then
            Wks.Range(Chr(intField + i) & intTitle).Font.Size = 11
            If InStr(strF2(i), "上期") > 0 Then
                Wks.Range(Chr(intField + i) & intTitle).Value = Replace(strF2(i), "上期", "")
            End If
        Else
            Wks.Range(Chr(intField + i) & intCounter).Value = strF2(i)
            Wks.Range(Chr(intField + i) & intCounter).Font.Bold = True
            Wks.Columns(Chr(i + intField)).ColumnWidth = intWidth(i)
            Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
        End If
    Next i
    If bolIsLast = False Then intTitle = intCounter
    intCounter = intCounter + 1
End Sub

'  Add by Amy 2018/03/01 改寫法原寫法相減後會產生不止兩位小數,可能報表會與Excel加總數不相同
'*************************************************
'  產生報表資料(內外帳共用暫存檔)
'*************************************************
Private Sub ProduceData()
    Dim strQ As String, strWhere As String
    
On Error GoTo Checking
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
    
    'Modify by Amy 2020/04/16 公司別改變數 原:IIf(Text6 = "2", "J", Text6)
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
           strWhere = "And (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "')) "
        Else
            strWhere = "And (a0109 is null or a0109='" & strCmp & "') "
        End If
    End If
'------------------------------------------------
' 營業收入明細
'------------------------------------------------
    strQ = "Select * From  Acc010 Where a0101 >= '4' and a0101 < '5' and a0104 = '3' and InStr(a0102,'不用')=0 " & strWhere & "Order by a0101 asc "
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
       SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
       adoacc010.MoveNext
    Loop
    adoacc010.Close
 
'------------------------------------------------
' 營業支出明細
'------------------------------------------------
    strQ = "Select * From  Acc010 Where a0101 >= '6' and a0101 < '7' and a0104 = '3' and InStr(a0102,'不用')=0 " & strWhere & "Order by a0101 asc "
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
       SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
       adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'------------------------------------------------
' 非營業收入
'------------------------------------------------
    strQ = "Select * From  Acc010 Where a0101 >= '71' and a0101 < '72' and a0104 = '3' and InStr(a0102,'不用')=0 " & strWhere & "Order by a0101 asc "
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
       SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
       adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'------------------------------------------------
' 非營業支出明細
'------------------------------------------------
    strQ = "Select * From  Acc010 Where a0101 >= '72' and a0101 < '8' and a0104 = '3' and InStr(a0102,'不用')=0 " & strWhere & "Order by a0101 asc "
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
       SaveAcc040 adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value
       adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'更新上期差異數
    strSql = "Update Accrpt408 Set R40808=R40804-R40807 " & _
                "Where R40801= '" & strUserNum & "' And R40809=1  "
    adoTaie.Execute strSql
    
'-------------------------------------------------
' 合計
'-------------------------------------------------
    '營業收入
    SaveSum "4", "499999", "4T"

    '營業支出
    SaveSum "6", "699999", "6T"
    
    '營業損益
    SaveSum "4T", "6T", "6ZT"

    '營業外收入
    SaveSum "71", "719999", "71T"
    
    '營業外支出
    SaveSum "72", "729999", "72T"
  
'-------------------------------------------------
' 本期損益
'-------------------------------------------------
    SaveSum "ZZT", "ZZT", "ZZT"
    
Checking:
    If Err.Number = 0 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub SaveSum(strAccNo1 As String, strAccNo2 As String, strSumNo As String)
    Dim stSQL As String, GetSumName As String
    Dim intIns As Integer
    
    If InStr(strSumNo, "T") > 0 Then
        Select Case strSumNo
            Case "4T"
                GetSumName = "營業收入"
            '營業支出
            Case "6T"
                GetSumName = ReportSum(2)
            '營業損益
            Case "6ZT"
                GetSumName = ReportSum(3)
            '營業外收入
            Case "71T"
                GetSumName = ReportSum(5)
            '營業外支出
            Case "72T"
                GetSumName = ReportSum(6)
            '稅前淨損益
            Case "ZZT"
                GetSumName = ReportSum(7)
        End Select
        GetSumName = Replace(GetSumName, ":", "")
    End If
    If strSumNo = "6ZT" Then
        '營業損益(6ZT)=營業收入(4T)-營業支出(6T)
        stSQL = "Select '" & strUserNum & "','" & strSumNo & "','" & GetSumName & "',S1-E1,S2-E2,SDif1-EDif1,S3-E3,SDif2-EDif2,1 From " & _
                  "(Select '6ZT' as K1,Sum(R40804) as S1,Sum(R40805) as S2,Sum(R40804)-Sum(R40805) as SDif1,Sum(R40807) as S3,Sum(R40804)-Sum(R40807) as SDif2 From Accrpt408 Where R40801='" & strUserNum & "' And R40809=1 And R40802='4T' )," & _
                  "(Select '6ZT' as K2,Sum(R40804) as E1,Sum(R40805) as E2,Sum(R40804)-Sum(R40805) as EDif1,Sum(R40807) as E3,Sum(R40804)-Sum(R40807) as EDif2 From Accrpt408 Where R40801='" & strUserNum & "' And R40809=1 And R40802='6T' ) " & _
                    "Where K1=K2(+) "
    ElseIf strSumNo = "ZZT" Then
        '稅前淨損益(ZZT)=營業損益(6ZT)+營業外收入(71T)-營業外支出(72T)
        stSQL = "Select '" & strUserNum & "','" & strSumNo & "','" & GetSumName & "',S1-E1,S2-E2,SDif1-EDif1,S3-E3,SDif2-EDif2,1 From " & _
                  "(Select 'ZZT' as K1,Sum(R40804) as S1,Sum(R40805) as S2,Sum(R40804)-Sum(R40805) as SDif1,Sum(R40807) as S3,Sum(R40804)-Sum(R40807) as SDif2 From Accrpt408 Where R40801='" & strUserNum & "' And R40809=1 And (R40802='6ZT' or R40802='71T') )," & _
                  "(Select 'ZZT' as K2,Sum(R40804) as E1,Sum(R40805) as E2,Sum(R40804)-Sum(R40805) as EDif1,Sum(R40807) as E3,Sum(R40804)-Sum(R40807) as EDif2 From Accrpt408 Where R40801='" & strUserNum & "' And R40809=1 And R40802='72T' ) " & _
                "Where K1=K2(+) "
    Else
        stSQL = "Select '" & strUserNum & "','" & strSumNo & "','" & GetSumName & "',Sum(Nvl(R40804,0)),Sum(Nvl(R40805,0)),Sum(Nvl(R40804,0))-Sum(Nvl(R40805,0)),Sum(Nvl(R40807,0)),Sum(Nvl(R40804,0))-Sum(Nvl(R40807,0)),1 " & _
                    "From Accrpt408 Where R40801='" & strUserNum & "' And R40809=1 And R40802>='" & strAccNo1 & "' And R40802<='" & strAccNo2 & "' "
    End If
    stSQL = "Insert Into Accrpt408 (R40801,R40802,R40803,R40804,R40805,R40806,R40807,R40808,R40809) " & stSQL
    adoTaie.Execute stSQL, intIns
    '若沒資料可合計也需要新增,要照損益表格式顯示-婧瑄
    If intIns = 0 Then
        stSQL = "Insert Into Accrpt408 (R40801,R40802,R40803,R40804,R40805,R40806,R40807,R40808,R40809) " & _
                    "Values( '" & strUserNum & "','" & strSumNo & "','" & GetSumName & "',0,0,0,0,0,1) "
        adoTaie.Execute stSQL
    End If
End Sub

Private Sub SaveAcc040(strAccNo1 As String, strA0102 As String)
    Dim strSql As String, strWhere1 As String, strWhere2 As String
    
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
        strWhere1 = strWhere1 & " And a0401 || decode(length(a0402),1, '0'||a0402, a0402) >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & " "
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
        strWhere1 = strWhere1 & " And a0401 || decode(length(a0402),1, '0'||a0402, a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & " "
    End If
           
    '會計編號
    If strAccNo1 <> MsgText(601) Then
        strWhere2 = strWhere2 & " And SubStr(a0405, 1, 4) = '" & strAccNo1 & "' "
    End If
    '公司別
    'Modify by Amy 2020/04/16 公司別改抓變數 原:Text6
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strWhere2 = strWhere2 & " And a0403 in ('" & Replace(strCmp, "+", "','") & "' ) "
        Else
            strWhere2 = strWhere2 & " And a0403 = '" & strCmp & "' "
        End If
    End If
    
    strWhere2 = strWhere2 & " And a0404 = '" & MsgText(55) & "' "
    
    '本期實際數/預算數
    strSql = "Insert Into Accrpt408 (R40801,R40802,R40803,R40804,R40805,R40806,R40809) " & _
                "Select '" & strUserNum & "','" & strAccNo1 & "','" & strA0102 & "',Sum(Nvl(a0408,0)),Sum(Nvl(a0409,0)),Sum(Nvl(a0408,0))-Sum(Nvl(a0409,0)),1 From Acc040 " & _
                "Where " & Mid(strWhere1, 5) & strWhere2
    adoTaie.Execute strSql
    
    '上期實際/差異數
    strWhere1 = ""
    If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> Mid(MsgText(29), 1, 6) Then
        strWhere1 = strWhere1 & " And a0401 || decode(length(a0402),1, '0'||a0402, a0402) >= " & Val(Mid(MaskEdBox3.Text, 1, 3) & Mid(MaskEdBox3.Text, 5, 2)) & " "
    End If
    If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> Mid(MsgText(29), 1, 6) Then
        strWhere1 = strWhere1 & " And a0401 || decode(length(a0402),1, '0'||a0402, a0402) <= " & Val(Mid(MaskEdBox4.Text, 1, 3) & Mid(MaskEdBox4.Text, 5, 2)) & " "
    End If
    
    strSql = "Update Accrpt408 Set R40807=(Select Sum(Nvl(a0408,0)) From acc040 Where " & Mid(strWhere1, 5) & strWhere2 & ") " & _
                    "Where R40801= '" & strUserNum & "' And R40809=1 And R40802='" & strAccNo1 & "' "
    adoTaie.Execute strSql
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = LBound(strF2) To UBound(strF2)
       If UCase(strF2(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'end 2018/03/01

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData_Old()
'Dim douManaIn As Double
'Dim douManaOut As Double
'Dim douExManaIn As Double
'Dim douExManaOut As Double
'Dim douEsManaIn As Double
'Dim douEsManaOut As Double
'Dim douEsExManaIn As Double
'Dim douEsExManaOut As Double
'Dim douLastManaIn As Double
'Dim douLastManaOut As Double
'Dim douLastExManaIn As Double
'Dim douLastExManaOut As Double
'Dim douAmount As Double
'
'On Error GoTo Checking
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   adoaccrpt408.CursorLocation = adUseClient
'   adoaccrpt408.Open "select * from accrpt408", adoTaie, adOpenDynamic, adLockBatchOptimistic
''------------------------------------------------
'' 營業收入明細
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt408.AddNew
'      Accrpt408Save
'      adoaccrpt408.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 營業收入小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40804), sum(r40805), sum(r40807) from accrpt408 where r40801 = '" & strUserNum & "' and r40802 >= '4' and r40802 < '5'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(4)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40803").Value = ReportSum(1)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt408.Fields("r40804").Value = 0
'         douManaIn = 0
'      Else
'         adoaccrpt408.Fields("r40804").Value = Val(adoaccsum.Fields(0).Value)
'         douManaIn = Val(adoaccsum.Fields(0).Value)
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt408.Fields("r40805").Value = 0
'         douEsManaIn = 0
'      Else
'         adoaccrpt408.Fields("r40805").Value = Val(adoaccsum.Fields(1).Value)
'         douEsManaIn = Val(adoaccsum.Fields(1).Value)
'      End If
'      adoaccrpt408.Fields("r40806").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value))
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt408.Fields("r40807").Value = 0
'         douLastManaIn = 0
'      Else
'         adoaccrpt408.Fields("r40807").Value = CStr(Val(adoaccsum.Fields(2).Value))
'         douLastManaIn = Val(adoaccsum.Fields(2).Value)
'      End If
'      adoaccrpt408.Fields("r40808").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value))
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'      'Add By Cheng 2002/01/18
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(8)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 營業支出明細
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt408.AddNew
'      Accrpt408Save
'      adoaccrpt408.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 營業支出小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40804), sum(r40805), sum(r40807) from accrpt408 where r40801 = '" & strUserNum & "' and r40802 >= '6' and r40802 < '7'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(4)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40803").Value = ReportSum(2)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt408.Fields("r40804").Value = 0
'         douManaOut = 0
'      Else
'         adoaccrpt408.Fields("r40804").Value = Val(adoaccsum.Fields(0).Value)
'         douManaOut = Val(adoaccsum.Fields(0).Value)
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt408.Fields("r40805").Value = 0
'         douEsManaOut = 0
'      Else
'         adoaccrpt408.Fields("r40805").Value = Val(adoaccsum.Fields(1).Value)
'         douEsManaOut = Val(adoaccsum.Fields(1).Value)
'      End If
'      adoaccrpt408.Fields("r40806").Value = Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value)
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt408.Fields("r40807").Value = 0
'         douLastManaOut = 0
'      Else
'         adoaccrpt408.Fields("r40807").Value = Val(adoaccsum.Fields(2).Value)
'         douLastManaOut = Val(adoaccsum.Fields(2).Value)
'      End If
'      adoaccrpt408.Fields("r40808").Value = Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'      'Add By Cheng 2002/01/18
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(8)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 營業損益
''------------------------------------------------
'   adoaccrpt408.AddNew
'   adoaccrpt408.Fields("r40801").Value = strUserNum
'   adoaccrpt408.Fields("r40803").Value = ReportSum(3)
'   adoaccrpt408.Fields("r40804").Value = CStr(douManaIn - douManaOut)
'   adoaccrpt408.Fields("r40805").Value = CStr(douEsManaIn - douEsManaOut)
'   adoaccrpt408.Fields("r40806").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value))
'   adoaccrpt408.Fields("r40807").Value = CStr(douLastManaIn - douLastManaOut)
'   adoaccrpt408.Fields("r40808").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value))
'   AutoNoSave
'   adoaccrpt408.UpdateBatch
'
'   'Add By Cheng 2002/01/18
'   adoaccrpt408.AddNew
'   adoaccrpt408.Fields("r40801").Value = strUserNum
'   adoaccrpt408.Fields("r40804").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40805").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40806").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40807").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40808").Value = ReportSum(8)
'   AutoNoSave
'   adoaccrpt408.UpdateBatch
'
''------------------------------------------------
'' 非營業收入
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt408.AddNew
'      Accrpt408Save
'      adoaccrpt408.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 非營業收入小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40804), sum(r40805), sum(r40807) from accrpt408 where r40801 = '" & strUserNum & "' and r40802 >= '71' and r40802 < '72'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(4)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40803").Value = ReportSum(5)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt408.Fields("r40804").Value = 0
'         douExManaIn = 0
'      Else
'         adoaccrpt408.Fields("r40804").Value = CStr(Val(adoaccsum.Fields(0).Value))
'         douExManaIn = Val(adoaccsum.Fields(0).Value)
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt408.Fields("r40805").Value = 0
'         douEsExManaIn = 0
'      Else
'         adoaccrpt408.Fields("r40805").Value = CStr(Val(adoaccsum.Fields(1).Value))
'         douEsExManaIn = Val(adoaccsum.Fields(1).Value)
'      End If
'      adoaccrpt408.Fields("r40806").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value))
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt408.Fields("r40807").Value = 0
'         douLastExManaIn = 0
'      Else
'         adoaccrpt408.Fields("r40807").Value = CStr(Val(adoaccsum.Fields(2).Value))
'         douLastExManaIn = Val(adoaccsum.Fields(2).Value)
'      End If
'      adoaccrpt408.Fields("r40808").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value))
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'      'Add By Cheng 2002/01/18
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(8)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 非營業支出明細
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/2/20 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/2/20 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt408.AddNew
'      Accrpt408Save
'      adoaccrpt408.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 非營業支出小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40804), sum(r40805), sum(r40807) from accrpt408 where r40801 = '" & strUserNum & "' and r40802 >= '72' and r40802 < '8'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(4)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(4)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40803").Value = ReportSum(6)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt408.Fields("r40804").Value = 0
'         douExManaOut = 0
'      Else
'         adoaccrpt408.Fields("r40804").Value = CStr(Val(adoaccsum.Fields(0).Value))
'         douExManaOut = Val(adoaccsum.Fields(0).Value)
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt408.Fields("r40805").Value = 0
'         douEsExManaOut = 0
'      Else
'         adoaccrpt408.Fields("r40805").Value = CStr(Val(adoaccsum.Fields(1).Value))
'         douEsExManaOut = Val(adoaccsum.Fields(1).Value)
'      End If
'      adoaccrpt408.Fields("r40806").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value))
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt408.Fields("r40807").Value = 0
'         douLastExManaOut = 0
'      Else
'         adoaccrpt408.Fields("r40807").Value = CStr(Val(adoaccsum.Fields(2).Value))
'         douLastExManaOut = Val(adoaccsum.Fields(2).Value)
'      End If
'      adoaccrpt408.Fields("r40808").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value))
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'      'Add By Cheng 2002/01/18
'      adoaccrpt408.AddNew
'      adoaccrpt408.Fields("r40801").Value = strUserNum
'      adoaccrpt408.Fields("r40804").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40805").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40806").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40807").Value = ReportSum(8)
'      adoaccrpt408.Fields("r40808").Value = ReportSum(8)
'      AutoNoSave
'      adoaccrpt408.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 稅前損益
''------------------------------------------------
'   adoaccrpt408.AddNew
'   adoaccrpt408.Fields("r40801").Value = strUserNum
'   adoaccrpt408.Fields("r40803").Value = ReportSum(7)
'   adoaccrpt408.Fields("r40804").Value = CStr(douManaIn - douManaOut + douExManaIn - douExManaOut)
'   adoaccrpt408.Fields("r40805").Value = CStr(douEsManaIn - douEsManaOut + douEsExManaIn - douEsExManaOut)
'   adoaccrpt408.Fields("r40806").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value))
'   adoaccrpt408.Fields("r40807").Value = CStr(douLastManaIn - douLastManaOut + douLastExManaIn - douLastExManaOut)
'   adoaccrpt408.Fields("r40808").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value))
'   AutoNoSave
'   adoaccrpt408.UpdateBatch
'   adoaccrpt408.AddNew
'   adoaccrpt408.Fields("r40801").Value = strUserNum
'   adoaccrpt408.Fields("r40804").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40805").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40806").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40807").Value = ReportSum(8)
'   adoaccrpt408.Fields("r40808").Value = ReportSum(8)
'   AutoNoSave
'   adoaccrpt408.UpdateBatch
'   adoaccrpt408.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt408Delete()
    'Modify by Amy 2018/03/01 避免同時進入+R40801, 避免內外帳同時操作+R40809='1'
    adoTaie.Execute "Delete From Accrpt408 Where R40801='" & strUserNum & "' And R40809='1' "
    'intAutoNo = 0
End Sub

'*************************************************
' 計算會計科目餘額資料並儲存至損益表資料暫存檔中
'
'*************************************************
'Private Sub Accrpt408Save2()
    'Mark by Amy 2018/03/01
'Dim strSql As String
'
'   adoacc040.CursorLocation = adUseClient
'   If Text6 <> MsgText(601) Then
'      '2014/2/20 modify by sonia
'      'strSql = " and a0403 = '" & Text6 & "'"
'      strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      '2014/2/20 end
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
'      strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
'      strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
'   End If
'   '2012/1/11 MODIFY BY SONIA 科目有"不用"字眼的不抓
'   'adoacc040.Open "select sum(a0408), sum(a0409) from acc040 where substr(a0405, 1, 4) = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   adoacc040.Open "select sum(a0408), sum(a0409) from acc040,acc010 where a0405=a0101(+) and instr(a0102,'/不用')=0 and substr(a0405, 1, 4) = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc040.RecordCount <> 0 Then
'      If IsNull(adoacc040.Fields(0).Value) Then
'         adoaccrpt408.Fields("r40804").Value = 0
'      Else
'         adoaccrpt408.Fields("r40804").Value = adoacc040.Fields(0).Value
'      End If
'      If IsNull(adoacc040.Fields(1).Value) Then
'         adoaccrpt408.Fields("r40805").Value = 0
'      Else
'         adoaccrpt408.Fields("r40805").Value = adoacc040.Fields(1).Value
'      End If
'      adoaccrpt408.Fields("r40806").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40805").Value))
'   Else
'      adoaccrpt408.Fields("r40804").Value = 0
'      adoaccrpt408.Fields("r40805").Value = 0
'      adoaccrpt408.Fields("r40806").Value = 0
'   End If
'   adoacc040.Close
'   adoacc040L.CursorLocation = adUseClient
'   strSql = MsgText(601)
'   If Text6 <> MsgText(601) Then
'      strSql = " and a0403 = '" & Text6 & "'"
'   End If
'   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> Mid(MsgText(29), 1, 6) Then
'      strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) >= " & Val(Mid(MaskEdBox3.Text, 1, 3) & Mid(MaskEdBox3.Text, 5, 2)) & ""
'   End If
'   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> Mid(MsgText(29), 1, 6) Then
'      strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) <= " & Val(Mid(MaskEdBox4.Text, 1, 3) & Mid(MaskEdBox4.Text, 5, 2)) & ""
'   End If
'   adoacc040L.Open "select sum(a0408) from acc040 where substr(a0405, 1, 4) = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc040L.RecordCount <> 0 Then
'      If IsNull(adoacc040L.Fields(0).Value) Then
'         adoaccrpt408.Fields("r40807").Value = 0
'      Else
'         adoaccrpt408.Fields("r40807").Value = adoacc040L.Fields(0).Value
'      End If
'   Else
'      adoaccrpt408.Fields("r40807").Value = 0
'   End If
'   adoacc040L.Close
'   adoaccrpt408.Fields("r40808").Value = CStr(Val(adoaccrpt408.Fields("r40804").Value) - Val(adoaccrpt408.Fields("r40807").Value))
'   AutoNoSave
'End Sub

'*************************************************
' 設定初始及結束年月
'
'*************************************************
'Private Sub Accrpt408Save()
    'Mark by Amy 2018/03/01
'   adoaccrpt408.Fields("r40801").Value = strUserNum
'   adoaccrpt408.Fields("r40802").Value = adoacc010.Fields("a0101").Value
'   If IsNull(adoacc010.Fields("a0102").Value) Then
'      adoaccrpt408.Fields("r40803").Value = Null
'   Else
'      adoaccrpt408.Fields("r40803").Value = adoacc010.Fields("a0102").Value
'   End If
'   Accrpt408Save2
'End Sub

'*************************************************
' 自動編號存入 r40809 欄位
'
'*************************************************
'Private Sub AutoNoSave()
    'Mark by Amy 2018/03/01
'   intAutoNo = intAutoNo + 1
'   adoaccrpt408.Fields("r40809").Value = intAutoNo
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/16 公司別改下拉
'   Text6 = ""
'   Text7 = ""
   CboCmp = ""
   'end 2020/04/16
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   CboCmp.SetFocus 'Modify by Amy 2020/04/16 原:Text6
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Moidfy by Amy 2020/04/16 +bolShowMsg
Public Function FormCheck(bolShowMsg As Boolean) As Boolean
   'Add by Amy 2020/04/16 +公司別判斷
   Dim bCancel As Boolean
   
   If Trim(CboCmp) <> MsgText(601) Then
        Call CboCmp_Validate(bCancel)
        If bCancel = True Then
            bolShowMsg = True
            Exit Function
        End If
   End If
   'end 2020/04/16
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox3.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox4.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
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
'Added by Lydia 2016/02/17
Private Sub MaskEdBox2_LostFocus()
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox2.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox2.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox2.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox2.SetFocus
         End If
      End If
   End If
End Sub
'Added by Lydia 2016/02/17
Private Sub MaskEdBox3_LostFocus()
   If MaskEdBox3.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox3.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox3.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox3.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox3.SetFocus
         End If
      End If
   End If
End Sub
'Added by Lydia 2016/02/17
Private Sub MaskEdBox4_LostFocus()
   If MaskEdBox4.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox4.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox4.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox4.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox4.SetFocus
         End If
      End If
   End If
End Sub

'Add by Amy 2020/04/16
Private Sub SetCompN()
    strCmpN = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

