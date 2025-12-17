VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmacc14j0 
   AutoRedraw      =   -1  'True
   Caption         =   "翻譯費總表"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   5010
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1230
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   900
      Width           =   2820
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
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1440
      Width           =   4500
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   2
      Top             =   495
      Width           =   300
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   1575
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
      Left            =   3165
      TabIndex        =   1
      Top             =   120
      Width           =   1575
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
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   225
      TabIndex        =   7
      Top             =   930
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Left            =   225
      TabIndex        =   6
      Top             =   165
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   3110
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "身分              (1:內翻  2:外翻  空白:全部)"
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
      Left            =   225
      TabIndex        =   5
      Top             =   525
      Width           =   4215
   End
End
Attribute VB_Name = "frmacc14j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/04/07 Form2.0已修改 (Printer 改以Excel印)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2007/6/4
Option Explicit

Dim PLeft(0 To 6) As Integer
Dim m_intPage As Integer
Dim m_iPrint As Integer
'預設印表機
Dim m_DefaultPrinter As String, m_Prn As Printer
Dim m_Grp As String
'Add by Amy 2022/04/07
Dim strField, intWidth
Dim i As Integer, intField As Integer, intR As Integer, intTitleRow As Integer
Dim strPrinter As String 'Add By Amy 2022/05/02

Private Sub Command1_Click()
   FormPrint
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5130, 2460
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   '畫面初值設定
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text4 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   '若印表機變動, 則更新列印設定
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   PUB_RestorePrinter m_DefaultPrinter
   
   Set frmacc14j0 = Nothing
End Sub


Private Function FormCheck() As Boolean
   If MaskEdBox1.Text = MsgText(29) Then
      MsgBox "入帳日期不可空白！"
      MaskEdBox1.SetFocus
      Exit Function
   ElseIf MaskEdBox2.Text = MsgText(29) Then
      MsgBox "入帳日期不可空白！"
      MaskEdBox2.SetFocus
      Exit Function
   Else
      FormCheck = True
   End If
End Function

Private Sub MaskEdBox1_GotFocus()
   If MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox1.SelStart = 0
      MaskEdBox1.SelLength = MaskEdBox1.MaxLength
   End If
End Sub

Private Sub MaskEdBox2_GotFocus()
   If MaskEdBox2.Text = MsgText(29) And MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox2 = MaskEdBox1
      MaskEdBox2.SelStart = 0
      MaskEdBox2.SelLength = MaskEdBox2.MaxLength
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub FormPrint()
   Dim strCon As String, strCon1 As String
   strCon = ""
   '入帳日期
   If MaskEdBox1.Text <> MsgText(29) Then
      strCon = strCon & " and a1p18>=" & Val(FCDate(MaskEdBox1.Text))
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a1p18<=" & Val(FCDate(MaskEdBox2.Text))
   End If
   
   'Modify by Morgan 2011/3/3 改判斷部門
   '內翻
   If Text4 = "1" Then
      'strCon1 = strCon1 & " and s2.st04='1'"
      strCon1 = strCon1 & " and s1.st03='F52'"
   '外翻
   ElseIf Text4 = "2" Then
      'strCon1 = strCon1 & " and nvl(s2.st04,'2')='2'"
      strCon1 = strCon1 & " and s1.st03='F51'"
   End If
   
   'strExc(0) = "select nvl(s2.st04,'2') C01,a1p15||' '||s1.st02 C02, pay C03,a1p18 C04 from (" & _
      " select a1p15,a1p18,sum(a1p07) pay from acc1p0 where a1p05='6130'" & strCon & _
      " group by a1p15,a1p18),staff s1,staff_idmap,staff s2" & _
      " where s1.st01(+)=a1p15 and s1.st01 is not null and sim02(+)=a1p15 and s2.st01(+)=sim01" & strCon1 & _
      " order by 1,2"
   'Modified by Morgan 2018/4/2 剔除沒有 Transfee 資料者--婧瑄
   strExc(0) = "select decode(s1.st03,'F52','1','2') C01,a1p15||' '||s1.st02 C02, pay C03,a1p18 C04 from (" & _
      " select a1p15,a1p18,sum(a1p07) pay from acc1p0,transfee where a1p05='6130'" & strCon & _
      " AND TF07(+)=A1P04 AND TF14(+)=A1P17 and tf01 is not null group by a1p15,a1p18),staff s1,staff_idmap,staff s2" & _
      " where s1.st01(+)=a1p15 and s1.st01 is not null and sim02(+)=a1p15 and s2.st01(+)=sim01" & strCon1 & _
      " order by 1,2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify by Amy 2022/04/07 改Excel印
      'DoPrint RsTemp.Clone
      Call SetPrinter(False) 'Add by Amy 2022/05/02
      If ExcelSave(RsTemp.Clone) = True Then
        MsgBox "列印完成！"
      End If
   Else
      MsgBox "查無資料！"
   End If
   Call SetPrinter(True) 'Add by Amy 2022/05/02
   
End Sub

'Add by Amy 2022/04/07 以Excel印
Private Function ExcelSave(p_Rst As ADODB.Recordset) As Boolean
    Dim Xls As New Excel.Application, Wks As New Worksheet
    Dim strWkName As String, strFileName As String, strTitleN As String, strFormat As String '工作表名稱為中文/檔案名稱/表名/儲存格格式
    Dim strOldState As String, strTmp As String
    
On Error GoTo ErrHnd
    
    If Trim(Text4) = MsgText(601) Or Trim(Text4) = "1" Then
        strTitleN = "所內工程師外譯"
    Else
        strTitleN = "外譯人員"
    End If
    
    intField = 65:  intR = 1
    strFileName = strTitleN & "翻譯費總表" & ServerDate & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    Xls.SheetsInNewWorkbook = 3
    Xls.Workbooks.add
    Xls.Visible = False
    '工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
    Set Wks = Xls.Worksheets(strWkName & "1")
    
    ReDim stField(2): ReDim intWidth(2)
    strField = Array("姓名", "　", "金額")
    intWidth = Array(15, 10, 20)
    
    With Wks
        .Range(Chr(intField) & intR).Font.Size = 20
        .Range(Chr(intField) & intR).Font.Name = "新細明體"
        .Range(Chr(intField) & intR).Font.Bold = True
        .Range(Chr(intField) & intR).Value = strTitleN & "翻譯費總表"
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).MergeCells = True
        intR = intR + 1
        .Range(Chr(intField) & intR).Font.Size = 11
        .Range(Chr(intField) & intR).Value = "入帳日期：" & MaskEdBox1 & " － " & MaskEdBox2
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).MergeCells = True
        intR = intR + 1
        .Range(Chr(intField) & intR).Font.Size = 11
        .Range(Chr(intField) & intR).Value = "列印人：" & StaffQuery(strUserNum)
        .Range(Chr(intField + UBound(stField) - 1) & intR).Value = "列印日期：" & CFDate(strSrvDate(2))
        .Range(Chr(intField + 1) & intR & ":" & Chr(UBound(stField) + intField) & intR).MergeCells = True
        .Range(Chr(intField + UBound(stField)) & intR).HorizontalAlignment = xlRight
        
        intR = intR + 2
        For i = LBound(strField) To UBound(strField)
            .Range(Chr(intField + i) & intR).Value = strField(i)
            .Columns(Chr(intField + i) & ":" & Chr(intField + i)).ColumnWidth = intWidth(i)
        Next i
        '設定格式
        .Range(Chr(intField) & "2:" & Chr(UBound(stField) + intField) & intR).Font.Size = 12
        .Range(Chr(intField) & "2:" & Chr(UBound(stField) + intField) & intR).Font.Name = "新細明體"
        .Range(Chr(intField) & "2:" & Chr(UBound(stField) + intField) & intR).Font.Bold = True
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeTop).Weight = xlThin
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeBottom).Weight = xlThin
        intTitleRow = intR
        intR = intR + 1
         
         '版面設定
        .PageSetup.PaperSize = 9 'A4
        .PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
        .PageSetup.Orientation = xlPortrait '直印
        .PageSetup.PrintTitleRows = "$1:$" & intTitleRow '標題列
        .PageSetup.LeftMargin = Xls.InchesToPoints(0.5) '邊界
        .PageSetup.RightMargin = Xls.InchesToPoints(0.5)
        .PageSetup.TopMargin = Xls.InchesToPoints(0.5)
        .PageSetup.BottomMargin = Xls.InchesToPoints(1)
        .PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
       
        Do While RsTemp.EOF = False
            '外翻 / 內翻　分開印
            If strOldState <> MsgText(601) And strOldState <> RsTemp.Fields("C01") Then
                .Range(Chr(intField + LBound(strField)) & intR).Value = "總　　　計"
                .Range(Chr(intField + LBound(strField)) & intR).Font.Bold = True
                
                .Range(Chr(intField + UBound(strField)) & intR).Value = "=Sum(" & Chr(intField + UBound(strField)) & intTitleRow + 1 & ":" & Chr(intField + UBound(strField)) & intR - 1 & ")"
                .Range(Chr(intField + UBound(strField)) & intR).Font.Bold = True
                .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeTop).Weight = xlThin  '細線
                .PrintOut Copies:=1, Collate:=True
                
                '刪除資料
                .Range(Chr(intField) & intTitleRow + 1 & ":" & Chr(intField + UBound(stField)) & intR).Delete
                .Range(Chr(intField) & "1").Value = "外譯人員翻譯費總表"
                intR = intTitleRow + 1
            End If
            
            For i = LBound(strField) To UBound(strField)
                strFormat = "": strTmp = ""
                Select Case strField(i)
                    Case "姓名"
                        strTmp = "" & RsTemp.Fields("C02")
                    Case "金額"
                        strTmp = "" & RsTemp.Fields("C03")
                        If strTmp <> MsgText(601) Then
                            strFormat = "#,##0"
                        End If
                End Select
                If strFormat <> MsgText(601) Then
                    .Range(Chr(intField + i) & intR).NumberFormatLocal = strFormat
                    .Range(Chr(intField + i) & intR).HorizontalAlignment = xlRight
                End If
                .Range(Chr(intField + i) & intR).Value = strTmp
            Next i
            strOldState = RsTemp.Fields("C01")
            intR = intR + 1
            RsTemp.MoveNext
        Loop
        .Range(Chr(intField + LBound(strField)) & intR).Value = "總　　　計"
        .Range(Chr(intField + LBound(strField)) & intR).Font.Bold = True
        
        .Range(Chr(intField + UBound(strField)) & intR).Value = "=Sum(" & Chr(intField + UBound(strField)) & intTitleRow + 1 & ":" & Chr(intField + UBound(strField)) & intR - 1 & ")"
        .Range(Chr(intField + UBound(strField)) & intR).Font.Bold = True
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeTop).Weight = xlThin  '細線
        .PrintOut Copies:=1, Collate:=True
    End With

    '判斷版本
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
    Set Wks = Nothing
    Set Xls = Nothing
    ExcelSave = True
    Exit Function
  
ErrHnd:
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
    Set Wks = Nothing
    Set Xls = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function
'end 2022/04/07

'Mark by Amy 2022/04/07 不使用
Private Sub DoPrint(p_Rst As ADODB.Recordset)
'
'   Dim strTemp As String, lngTotal As Long
'
'On Error GoTo flgErr
'
'   '設定使用者所選擇的印表機成預設印表機
'   For Each m_Prn In Printers
'      If m_Prn.DeviceName = cmbPrinter.Text Then
'         Set Printer = m_Prn
'         Exit For
'      End If
'   Next
'   GetPleft
'   Printer.Orientation = 1 '直印
'   m_Grp = ""
'   With p_Rst
'      .MoveFirst
'      Do While Not .EOF
'         If m_Grp <> .Fields("C01") Then
'            If m_Grp <> "" Then
'               If m_iPrint + 900 > Printer.ScaleHeight Then
'                  Printer.NewPage
'                  PrintHead
'               Else
'                  m_iPrint = m_iPrint + 300
'               End If
'               PrintTotal lngTotal
'               Printer.NewPage
'            End If
'            lngTotal = 0
'            m_intPage = 0
'            m_Grp = "" & .Fields("C01")
'            PrintHead
'         Else
'            NewLine
'         End If
'         Printer.CurrentX = PLeft(0)
'         Printer.CurrentY = m_iPrint
'         Printer.Print "" & .Fields("C02")
'         strTemp = Format("" & .Fields("C03"), "#,##0")
'         Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp) - 50
'         Printer.CurrentY = m_iPrint
'         Printer.Print strTemp
'         lngTotal = lngTotal + Val("" & .Fields("C03"))
'         .MoveNext
'      Loop
'      NewLine
'      PrintTotal lngTotal
'      Printer.EndDoc
'      MsgBox "列印完成！"
'   End With
'
'flgErr:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description, vbCritical
'   End If
   
End Sub

'Mark by Amy 2022/04/07 不使用
Private Sub NewLine()
'   m_iPrint = m_iPrint + 300
'   If m_iPrint + 900 > Printer.ScaleHeight Then
'      Printer.CurrentX = PLeft(0)
'      Printer.CurrentY = m_iPrint
'      Printer.DrawStyle = vbSolid
'      Printer.Line (PLeft(0), m_iPrint)-(PLeft(2), m_iPrint)
'      Printer.NewPage
'      PrintHead
'   End If
End Sub

Private Sub GetPleft()
'   PLeft(0) = 3000 '姓名
'   PLeft(1) = PLeft(0) + 2500 '金額
'   PLeft(2) = PLeft(1) + 2500 '右邊
End Sub

Private Sub PrintHead()
'
'   Dim strTemp As String
'
'   m_intPage = m_intPage + 1
'   m_iPrint = 500:
'
'   With Printer
'
'      '表頭
'      .Font.Size = 20
'      .Font.Bold = True
'      .Font.Underline = True
'
'      If m_Grp = "1" Then
'         strTemp = "所內工程師外譯翻譯費總表"
'      Else
'         strTemp = "外譯人員翻譯費總表"
'      End If
'      .CurrentX = PLeft(0) + (PLeft(2) - PLeft(0) - .TextWidth(strTemp)) / 2
'      .CurrentY = m_iPrint
'      Printer.Print strTemp
'
'      .Font.Size = 10
'      .Font.Bold = False
'      .Font.Underline = False
'
'      '跳列
'      m_iPrint = m_iPrint + 600
'
'      '條件
'      strTemp = "入帳日期: " & MaskEdBox1 & " － " & MaskEdBox2
'      .CurrentX = PLeft(0) + (PLeft(2) - PLeft(0) - .TextWidth(strTemp)) / 2
'      .CurrentY = m_iPrint
'      Printer.Print strTemp
'
'      '跳列
'      m_iPrint = m_iPrint + 500
'
'      Printer.CurrentX = PLeft(0)
'      Printer.CurrentY = m_iPrint
'      Printer.Print "列印人：" & strUserName
'
'
'      strTemp = "列印日期：" & CFDate(strSrvDate(2))
'      .CurrentX = PLeft(2) - .TextWidth(strTemp)
'      .CurrentY = m_iPrint
'      Printer.Print strTemp
'
'      '跳列
'      m_iPrint = m_iPrint + 300
'
'      strTemp = "列印日期：" & CFDate(strSrvDate(2))
'      .CurrentX = PLeft(2) - .TextWidth(strTemp)
'      .CurrentY = m_iPrint
'      Printer.Print "頁　　次：　" & str(m_intPage)
'
'      m_iPrint = m_iPrint + 300
'      DrawLine
'
'      .Font.Size = 12
'      .Font.Bold = True
'      .CurrentX = PLeft(0)
'      .CurrentY = m_iPrint
'      Printer.Print "姓名"
'
'      .CurrentX = PLeft(1)
'      .CurrentY = m_iPrint
'      Printer.Print "金額"
'
'      .Font.Bold = False
'
'      m_iPrint = m_iPrint + 300
'      DrawLine
'   End With
End Sub

Private Sub PrintTotal(p_lngTot As Long)
'   Dim strTemp As String
'
'   DrawLine
'   Printer.Font.Bold = True
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = m_iPrint
'   Printer.Print "總　　　計"
'   strTemp = Format(p_lngTot, "$#,##0")
'   Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp) - 50
'   Printer.CurrentY = m_iPrint
'   Printer.Print strTemp
'   Printer.Font.Bold = False
End Sub

Private Sub DrawLine()
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = m_iPrint
'   Printer.DrawStyle = vbSolid
'   Printer.DrawWidth = 4
'   Printer.Line (PLeft(0), m_iPrint)-(PLeft(2), m_iPrint)
'   m_iPrint = m_iPrint + 100
End Sub
'end 2022/04/07

'Add by Amy 2022/05/02
Private Sub SetPrinter(ByVal bolReCovery As Boolean)
    If bolReCovery = False Then
        '切換印表機
        PUB_SetOsDefaultPrinter cmbPrinter
        PUB_RestorePrinter cmbPrinter
    Else
        '還原印表機
        PUB_SetOsDefaultPrinter strPrinter
        PUB_RestorePrinter strPrinter
    End If
End Sub

