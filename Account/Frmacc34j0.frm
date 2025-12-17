VERSION 5.00
Begin VB.Form Frmacc34j0 
   AutoRedraw      =   -1  'True
   Caption         =   "支票列印"
   ClientHeight    =   3456
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5136
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3456
   ScaleWidth      =   5136
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Top             =   1860
      Width           =   705
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   1560
      Width           =   705
   End
   Begin VB.CheckBox chk 
      Caption         =   "不刪表格資料（測）"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   90
      Width           =   1950
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Word 套印"
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
      TabIndex        =   13
      Top             =   1140
      Width           =   4692
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1320
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   780
      Width           =   3540
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2250
      TabIndex        =   8
      Top             =   2640
      Width           =   705
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2250
      TabIndex        =   7
      Top             =   2940
      Width           =   705
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   2
      Top             =   420
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   60
      Width           =   1572
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
      Left            =   228
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2220
      Width           =   4692
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   420
      Width           =   1572
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "左邊界：　　　　　　　(單位公分)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   390
      TabIndex        =   18
      Top             =   1890
      Width           =   3780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "上邊界：　　　　　　　(單位公分)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   390
      TabIndex        =   17
      Top             =   1590
      Width           =   3780
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　(單位公分)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   2670
      Width           =   3900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　(單位公分)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   2970
      Width           =   3900
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   420
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc34j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/02/24 Form2.0已修改 (改為Word套印-以信封方式進紙)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strOrder1, strOrder2 As String
Dim strSql As String, intCounter As Integer
Dim strAmount As String, intLength As Integer
Dim strName As String
'Add By Cheng 2003/03/27
Dim m_dblPLeft As Double 'X軸
Dim m_dblPTop As Double 'Y軸
'預設印表機
Dim m_DefaultPrinter As String
Dim strFileName As String, strPrinter As String 'Add by Amy 2022/02/24

'Add by Amy 2022/05/12
Private Sub Combo1_Click()
    SetPrinterStartPos Combo1
End Sub

Private Sub Command1_Click()
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    'Modify by Amy 2022/05/12
    If InStr(Combo1, "IBM 5577-KC2") = 0 Then
        MsgBox "請選擇IBM 5577-KC2 印表機"
        Exit Sub
    End If
    'end 2022/05/12
    Screen.MousePointer = vbHourglass
    Command1.Enabled = False
    'Add by Amy 2022/02/24 從ProcessData1搬過來
    strSql = ""
    If Text2 <> MsgText(601) Then
        strSql = " and a0e07 = '" & Text2 & "'"
    End If
    If Text1 <> MsgText(601) Then
        strSql = strSql & " and a0e02 >= '" & Text1 & "'"
    End If
    If Text4 <> MsgText(601) Then
        strSql = strSql & " and a0e02 <= '" & Text4 & "'"
    End If
    PUB_RestorePrinter Combo1 'Modify by Amy 2022/05/12
    'ProcessData1 'Mark by Amy 2022/02/24 未使用 第一信用合作社--本票
    ProcessData2 '瑞興銀行-支票
    'ProcessData3 'Mark by Amy 2022/02/24 未使用 台北國際商業銀行--支票
    PUB_RestorePrinter strPrinter 'Modify by Amy 2022/05/12
    Command1.Enabled = True
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

'Add by Amy 2022/02/24 ProcessData2改Word 套印
Private Sub Command2_Click()
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    'Add by Amy 2022/05/12 +Word 微調邊界
    If Trim(Text5.Text) <> MsgText(601) Or Trim(Text6.Text) <> MsgText(601) Then
        If Text5.Text <> MsgText(601) Then
            Text5.Text = Trim(Text5.Text)
            If Not IsNumeric(Text5.Text) Then
                MsgBox "上邊界請輸入數字", , MsgText(5)
                Exit Sub
            '避免下邊界扣掉後變負數
            ElseIf Val(Text5.Text) > 0.43 Then
                MsgBox "上邊界輸入值過大", , MsgText(5)
                Exit Sub
            End If
        End If
        If Text6.Text <> MsgText(601) Then
            Text6.Text = Trim(Text6.Text)
            If Not IsNumeric(Text6.Text) Then
                MsgBox "左邊界請輸入數字", , MsgText(5)
                Exit Sub
            '避免右邊界扣掉後變負數
            ElseIf Val(Text6.Text) > 0.5 Then
                MsgBox "左邊界輸入值過大", , MsgText(5)
                Exit Sub
            End If
        End If
        
    End If
    Screen.MousePointer = vbHourglass
    Command2.Enabled = False
    Call SetPrinter(False)
    
    strSql = ""
    If Text2 <> MsgText(601) Then
        strSql = " and a0e07 = '" & Text2 & "'"
    End If
    If Text1 <> MsgText(601) Then
        strSql = strSql & " and a0e02 >= '" & Text1 & "'"
    End If
    If Text4 <> MsgText(601) Then
        strSql = strSql & " and a0e02 <= '" & Text4 & "'"
    End If
    If ProcessData4 = True Then
        FormClear
    End If
    Call SetPrinter(True)
    Command2.Enabled = True
    Screen.MousePointer = vbDefault
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

'Add by Amy 2022/02/24 瑞興銀行-支票 (Word套印)
'Memo 避免套印不準,請以財務 Word2010 且 用財務印表機列測式列印 (因雖電腦中心印表機與財務印表機同型號,但同機器印仍會有誤差）
Private Function ProcessData4() As Boolean
    Dim ii As Integer, bVisible As Boolean, m_WordLeft As Long, m_WordTop As Long
    Dim strA0E10(2) As String, strA0E13(2) As String, strName(1) As String, strNTWord As String, strNT As String, strMemo As String
    Dim strQ As String, strTp As String
   
On Error GoTo Checking
    Me.MousePointer = vbHourglass
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
    ProcessData4 = False
   
    strQ = "Select * From Acc0e0 Where a0e04 = '" & MsgText(19) & "' and a0e07 in ('1756650') and a0e08 = '1'" & strSql & _
            " Order by a0e02 asc, a0e26 asc, a0e27 asc"
    adoacc0e0.CursorLocation = adUseClient
    adoacc0e0.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoacc0e0.RecordCount = 0 Then
        ProcessData4 = True
        adoacc0e0.Close
        MsgBox "無資料可列印！"
        StatusClear
        Me.MousePointer = vbDefault
        Exit Function
    End If
    
    'Modify by Amy 2023/03/21 開過Word 切換其他模式,再跑此程式會Error
    'If g_WordAp Is Nothing Then
    If TypeName(g_WordAp) <> "Application" Then
        Set g_WordAp = New Word.Application
    End If
    If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Function
  
    With g_WordAp.Application
        '切換為整頁模式
        If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
        Else
            .ActiveWindow.View.Type = wdPageView
        End If
    End With
    'end 2023/03/21
    g_WordAp.Documents.Open App.path & "\" & strUserNum & "\" & strFileName
    g_WordAp.ActiveDocument.SaveAs App.path & "\" & strUserNum & "\" & strFileName
        
    strMemo = "禁止背書轉讓"
    Do While adoacc0e0.EOF = False
        If adoacc0e0.Fields("a0e46").Value >= 1 Then
            strDelConfirm = MsgBox(MsgText(159) & adoacc0e0.Fields("a0e02").Value & MsgText(160), vbOKCancel + vbDefaultButton1, MsgText(5))
            If strDelConfirm = vbCancel Then
               GoTo NextSkip
            End If
        End If
        
        '票根發票日
        strTp = "" & adoacc0e0.Fields("a0e13")
        strA0E13(0) = Mid(strTp, 1, 3)
        strA0E13(1) = Mid(strTp, 4, 2)
        strA0E13(2) = Mid(strTp, 6, 2)
        
        '票根到期日/日期
        strTp = "" & adoacc0e0.Fields("a0e10")
        strA0E10(0) = Mid(strTp, 1, 3)
        strA0E10(1) = Mid(strTp, 4, 2)
        strA0E10(2) = Mid(strTp, 6, 2)
        
        '票根受款人
        strTp = "" & adoacc0e0.Fields("a0e12").Value
        '取10個字換行
        If Len(strTp) > 10 Then
            strName(0) = Mid(strTp, 1, 10)
            strName(0) = strName(0) & vbCrLf & Replace(strTp, strName(0), "")
        Else
            strName(0) = strTp
        End If
        
        '憑票支付
        strTp = "" & adoacc0e0.Fields("a0e12").Value
       'Modify by Amy 2022/03/09 取20個字換行
        If Len(strTp) > 20 Then
            strName(1) = Mid(strTp, 1, 20)
            strName(1) = strName(1) & vbCrLf & Replace(strTp, strName(1), "")
        Else
            strName(1) = strTp
        End If
        
        '票根金額/NT
        strNT = Format(Val("" & adoacc0e0.Fields("a0e11")), FDollar)
        '新臺幣(大寫)
        strNTWord = ChangeNumber(Val("" & adoacc0e0.Fields("a0e11")))
        
        '填入位置
        With g_WordAp
            .Selection.PageSetup.TopMargin = .CentimetersToPoints(1 + Val(Text5.Text))
            .Selection.PageSetup.BottomMargin = .CentimetersToPoints(0.43 - Val(Text5.Text))
            .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.2 + Val(Text6.Text))
            .Selection.PageSetup.RightMargin = .CentimetersToPoints(0.5 - Val(Text6.Text))
            .ActiveDocument.Tables(1).Select
            .Selection.Font.Name = "細明體"
            .Selection.Font.Size = 10
            .Selection.Font.Bold = False
            .Selection.MoveLeft Unit:=wdCell, Count:=1
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.MoveRight Unit:=wdCell, Count:=1
            '*** 票根 ***
            'Modify by Amy 2022/03/02 設定靠左(避免印不準)
            '發票日
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:=strA0E13(0)
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:=strA0E13(1)
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:=strA0E13(2)
            
            '到期日(倒著印)
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:=strA0E10(2)
            .Selection.MoveLeft Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:=strA0E10(1)
            .Selection.MoveLeft Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.TypeText Text:=strA0E10(0)
            'end 2022/03/02
            
            '授款人
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.TypeText Text:=strName(0)
           
            '金額
            .Selection.MoveDown Unit:=wdLine, Count:=1
            'Modify by Amy 2022/05/03 0002652 因授款人跳行導致跳錯欄,將範本合併儲存格加空白印
            .Selection.TypeText Text:=String(3, "　　") & strNT
            
            '*** 支票內容 ***
            .Selection.MoveRight Unit:=wdCell, Count:=3
            .Selection.MoveUp Unit:=wdLine, Count:=3
            .Selection.MoveRight Unit:=wdCell, Count:=3
            
            '日期
            'Modify by Amy 2022/03/02 設定靠右(避免印不準)
            .Selection.Font.Size = 12
            .Selection.Font.Bold = True
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Selection.TypeText Text:=strA0E10(0)
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.Font.Size = 12
            .Selection.Font.Bold = True
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Selection.TypeText Text:=strA0E10(1)
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.Font.Size = 12
            .Selection.Font.Bold = True
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Selection.TypeText Text:=strA0E10(2)
            'end 2022/03/02
            
            '憑票支付
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.Font.Size = 12
            .Selection.Font.Bold = True
            .Selection.TypeText Text:=strName(1)
            
            '新臺幣
            .Selection.MoveDown Unit:=wdLine, Count:=1
             .Selection.Font.Size = 14
            .Selection.Font.Bold = True
            .Selection.TypeText Text:=strNTWord
            
            'NT$
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.Font.Size = 12
            .Selection.Font.Bold = True
            .Selection.TypeText Text:=strNT
            
            '禁止背書轉讓
            'Mark by Amy 2022/03/14 調整範本(合併NT及禁止背書轉讓儲存格),拿掉右移,因發現 壹拾貳萬柒仟伍佰元整 字太多若右移會跳錯格,strMemo前加全型空白對位置
            '.Selection.MoveRight Unit:=wdCharacter, Count:=1 '最後一格最好別用wdCell,避免跳錯
            .Selection.MoveDown Unit:=wdLine, Count:=1 '先跳右再跳下,避免跳錯 ex:Word2010與Word21013 會跳不同位置
            .Selection.Font.Size = 12
            .Selection.Font.Bold = True
            .Selection.TypeText Text:=String(10, "　") & strMemo
            'end 2022/03/14
            
            '列印
            .PrintOut Background:=False, Copies:=1, Collate:=True
            
            '刪除表格內容
            If chk.Value = vbChecked Then
                '電腦中心測式內容用
            Else
                .ActiveDocument.Tables(1).Select
                .Selection.Delete Unit:=wdCharacter, Count:=1
            End If
        End With
        '更新列印次數
        strQ = "Update Acc0e0 Set a0e46 = 1 where a0e01 = '" & adoacc0e0.Fields("a0e01") & "' and a0e02 = '" & adoacc0e0.Fields("a0e02") & "' "
        adoTaie.Execute strQ
       
NextSkip:
        adoacc0e0.MoveNext
    Loop
    adoacc0e0.Close
    g_WordAp.ActiveDocument.SaveAs App.path & "\" & strUserNum & "\" & strFileName
    '關閉
    g_WordAp.ActiveDocument.Close
    g_WordAp.Quit
    Set g_WordAp = Nothing
    ProcessData4 = True
    Exit Function
   
Checking:
    g_WordAp.ActiveDocument.SaveAs App.path & "\" & strUserNum & "\" & strFileName
    MsgBox Err.Description, , MsgText(5)
    '關閉
    g_WordAp.ActiveDocument.Close
    g_WordAp.Quit
    Set g_WordAp = Nothing
    
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

'設定列印起點
'Modify by Amy 2022/05/12 +stPSP06參數
Private Sub SetPrinterStartPos(ByVal stPSP06 As String)
   'Modified by Morgan 2017/11/8
   'strExc(0) = "Select * From PrintStartPoint Where PSP01='" & strUserNum & "' And PSP02='" & Me.Name & "' And PSP03='" & Me.Name & "' and psp06='" & Printer.DeviceName & "'"
   strExc(0) = "Select PSP04,PSP05,PSP06,'1' Srt From PrintStartPoint Where PSP01='" & strUserNum & "@" & pub_HostName & "' And PSP02='" & Me.Name & "' And PSP03='" & Me.Name & "' and psp06='" & stPSP06 & "'"
   strExc(0) = strExc(0) & " union Select PSP04,PSP05,PSP06,'2' Srt From PrintStartPoint Where PSP01='" & strUserNum & "' And PSP02='" & Me.Name & "' And PSP03='" & Me.Name & "' and psp06='" & stPSP06 & "' order by Srt Asc"
   'end 2017//18
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '若有資料
   If intI = 1 Then
      If InStr(stPSP06, "HP LaserJet P3005 PCL6") > 0 Then
        Me.Text5.Text = "" & CDbl("" & RsTemp("PSP04").Value)
        Me.Text6.Text = "" & CDbl("" & RsTemp("PSP05").Value)
        Me.Text5.Tag = Me.Text5.Text
        Me.Text6.Tag = Me.Text6.Text
      Else
        Me.Text7.Text = "" & CDbl("" & RsTemp("PSP04").Value)
        Me.Text8.Text = "" & CDbl("" & RsTemp("PSP05").Value)
        Me.Text7.Tag = Me.Text7.Text
        Me.Text8.Tag = Me.Text8.Text
      End If
   '若無資料
   Else
      If InStr(stPSP06, "HP LaserJet P3005 PCL6") > 0 Then
        Me.Text5.Text = "0"
        Me.Text6.Text = "0"
        Me.Text5.Tag = ""
        Me.Text6.Tag = ""
      Else
        Me.Text7.Text = "0"
        Me.Text8.Text = "0"
        Me.Text7.Tag = ""
        Me.Text8.Tag = ""
      End If
   End If
End Sub
'end 2022/05/012

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5250, 3900  'Modify by Amy 2023/08/18 原: 3690
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   'SetPrinterStartPos 'Mark by Amy 2022/05/12 往下搬
   'Modify by Amy 2022/02/24 改 Form2.0
   chk.Visible = False
   If Pub_StrUserSt03 = "M51" Then chk.Visible = True
   If Dir(App.path & "\" & strUserNum, vbDirectory) = MsgText(601) Then
        MkDir App.path & "\" & strUserNum
   End If
   '清除暫存檔
   If Dir(App.path & "\" & strUserNum & "\") <> "" Then
      PUB_KillTempFile strUserNum & "\$$*.docx"
   End If
   strFileName = "$$瑞興銀行支票.docx"
   If Dir(App.path & "\" & strUserNum & "\" & strFileName) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & strFileName
   End If
   Call PUB_GetSampleFile(strFileName, "M31-000014-0-00", , App.path & "\" & strUserNum & "\")
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   'end 2022/02/24
   SetPrinterStartPos Combo1  'Add by Amy 2022/05/12 從上面搬下來 +Combo1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Add by Amy 2022/02/24 +印表機
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '若偏移值有變動, 則更新列印設定
   'Add by Amy 2022/05/12 +雷射印表機設定
   If InStr(Combo1, "HP LaserJet P3005 PCL6") > 0 And (Me.Text5.Text <> Me.Text5.Tag Or Me.Text6.Text <> Me.Text6.Tag) Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Name, Me.Text5.Text, Me.Text6.Text, Combo1
   End If
   'Modify by Amy 2022/05/12 +IBM 5577-KC2,原Printer.DeviceName改Combo1
    If InStr(Combo1, "IBM 5577-KC2") > 0 And (Me.Text7.Text <> Me.Text7.Tag Or Me.Text8.Text <> Me.Text8.Tag) Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Name, Me.Text7.Text, Me.Text8.Text, Combo1
    End If
   Set Frmacc34j0 = Nothing
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   'edit by nickc 2007/02/08
   'Text3 = ""
   Text4 = ""
   Text2.SetFocus
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by sonia 2020/2/1
Private Sub Text1_LostFocus()
   If Text4 = "" Then Text4 = Text1
End Sub
'end 2020/2/1

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生支票資料並列印 (第一信用合作社--本票)
'  Memo by Amy 2022/02/24 目前未使用-與瑞婷確認
'*************************************************
Private Sub ProcessData1()
   
'On Error GoTo Checking
'   strSql = ""
'   Me.MousePointer = vbHourglass
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   If Text2 <> MsgText(601) Then
'      strSql = " and a0e07 = '" & Text2 & "'"
'   End If
'   If Text1 <> MsgText(601) Then
'      strSql = strSql & " and a0e02 >= '" & Text1 & "'"
'   End If
'   If Text4 <> MsgText(601) Then
'      strSql = strSql & " and a0e02 <= '" & Text4 & "'"
'   End If
'
'   adoacc0e0.CursorLocation = adUseClient
'   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149980') and a0e08 = '2' and (a0e46 is null or a0e46 = 0)" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'   '2010/6/21 MODIFY BY SONIA 加0149951
'   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149980') and a0e08 = '2'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'modify by sonia 2020/6/19
'   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149951', '0149980') and a0e08 = '2'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'modify by sonia 2020/6/20 加回0149951
'   'Modify by Amy 2020/07/24 取消0149951
'   adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('1756650') and a0e08 = '2'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0e0.RecordCount = 0 Then
'      adoacc0e0.Close
'      StatusClear
'      Me.MousePointer = vbDefault
'      Exit Sub
'   End If
'
'   'Modify by Morgan 2008/3/25 控制 9x 才自訂
'   If pub_OS = "1" Then
'      Printer.Height = 5800
'      Printer.Width = 14000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(8)
'   End If
'   'end 2008/3/25
'   Printer.FontSize = 10
'   'Add by Morgan 2004/7/29
'   Printer.Font.Name = "細明體"
'
'   Do While adoacc0e0.EOF = False
'      If adoacc0e0.Fields("a0e46").Value >= 1 Then
'         strDelConfirm = MsgBox(MsgText(159) & adoacc0e0.Fields("a0e02").Value & MsgText(160), vbOKCancel + vbDefaultButton1, MsgText(5))
'         If strDelConfirm = vbCancel Then
'            GoTo NextSkip
'         End If
'      End If
'      Printer.CurrentX = 8800
'      Printer.CurrentY = 300
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e13").Value), 1, 3))
'      Printer.CurrentX = 9500
'      Printer.CurrentY = 300
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e13").Value), 5, 2))
'      Printer.CurrentX = 10200
'      Printer.CurrentY = 300
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e13").Value), 8, 2))
'      Printer.CurrentX = 6000
'      Printer.CurrentY = 700
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
'      Printer.CurrentX = 6700
'      Printer.CurrentY = 700
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 5, 2))
'      Printer.CurrentX = 7400
'      Printer.CurrentY = 700
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 8, 2))
'      Printer.CurrentX = 5300
'      Printer.CurrentY = 1000
'      'adoquery.CursorLocation = adUseClient
'      'adoquery.Open "select * from acc0q0 where a0q01 = " & Val(IIf(IsNull(adoacc0e0.Fields("a0e03").Value), "", adoacc0e0.Fields("a0e03").Value)) & " and a0q03 = '" & IIf(IsNull(adoacc0e0.Fields("a0e06").Value), "", adoacc0e0.Fields("a0e06").Value) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      'If adoquery.RecordCount <> 0 Then
'      '   If IsNull(adoquery.Fields("a0q05").Value) Then
'      '      Select Case adoacc0e0.Fields("a0e05").Value
'      '         Case "1"
'      '            strName = CustomerQuery(adoacc0e0.Fields("a0e06").Value, 1)
'      '         Case "2"
'      '            strName = A0i02Query(adoacc0e0.Fields("a0e06").Value)
'      '         Case "3"
'      '            strName = StaffQuery(adoacc0e0.Fields("a0e06").Value)
'      '         Case Else
'      '            strName = ""
'      '      End Select
'      '   Else
'      '      strName = adoquery.Fields("a0q05").Value
'      '   End If
'      'Else
'      '   Select Case adoacc0e0.Fields("a0e05").Value
'      '      Case "1"
'      '         strName = CustomerQuery(adoacc0e0.Fields("a0e06").Value, 1)
'      '      Case "2"
'      '         strName = A0i02Query(adoacc0e0.Fields("a0e06").Value)
'      '      Case "3"
'      '         strName = StaffQuery(adoacc0e0.Fields("a0e06").Value)
'      '      Case Else
'      '         strName = ""
'      '   End Select
'      'End If
'      'adoquery.Close
'      If IsNull(adoacc0e0.Fields("a0e12").Value) Then
'         strName = ""
'      Else
'         strName = adoacc0e0.Fields("a0e12").Value
'      End If
'      Printer.Print strName
'      '92.10.24 MODIFY BY SONIA
'      'strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), DDollar)
'      strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), FDollar)
'      '92.10.24 END
'      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = 10200 - intLength
'      Printer.CurrentY = 1000
'      Printer.Print strAmount
'      Printer.FontSize = 12
'      Printer.CurrentX = 5300
'      Printer.CurrentY = 1300
'      Printer.Print ChangeNumber(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value))
'      Printer.FontSize = 10
'      Printer.CurrentX = 800
'      Printer.CurrentY = 1500
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
'      Printer.CurrentX = 1500
'      Printer.CurrentY = 1500
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 5, 2))
'      Printer.CurrentX = 2200
'      Printer.CurrentY = 1500
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 8, 2))
'      Printer.CurrentX = 0
'      Printer.CurrentY = 1800
'      Printer.Print strName
'      '92.10.24 MODIFY BY SONIA
'      'strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), DDollar)
'      strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), FDollar)
'      '92.10.24 END
'      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = 2000 - intLength
'      Printer.CurrentY = 3200
'      Printer.Print strAmount
''      Printer.CurrentX = 0
''      Printer.CurrentY = 3800
''      Printer.Print IIf(IsNull(adoacc0e0.Fields("a0e12").Value), MsgText(601), adoacc0e0.Fields("a0e12").Value)
'      adoTaie.Execute "update acc0e0 set a0e46 = 1 where a0e01 = '" & adoacc0e0.Fields("a0e01").Value & "' and a0e02 = '" & adoacc0e0.Fields("a0e02").Value & "'"
'      Printer.NewPage
'NextSkip:
'      adoacc0e0.MoveNext
'   Loop
'   adoacc0e0.Close
'   Printer.EndDoc
'   Me.MousePointer = vbDefault
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生支票資料並列印 (第一信用合作社--支票)
'
'*************************************************
Private Sub ProcessData2()
'Add By Cheng 2003/04/01
Dim ii As Integer
Dim jj As Integer
   
On Error GoTo Checking
   Me.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   
   adoacc0e0.CursorLocation = adUseClient
   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149980') and a0e08 = '1' and (a0e46 is null or a0e46 = 0)" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
   '2010/6/21 MODIFY BY SONIA 加0149951
   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149980') and a0e08 = '1'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2020/6/19
   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149951', '0149980') and a0e08 = '1'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2020/6/20 加回0149951
   'Modify by Amy 2020/07/24 取消0149951
   adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('1756650') and a0e08 = '1'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      StatusClear
      Me.MousePointer = vbDefault
      Exit Sub
   End If
   
   'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
   '9x
   If pub_OS = "1" Then
      Printer.Height = 4820
      Printer.Width = 14000
   Else
      Printer.PaperSize = PUB_GetPaperSize(6)
   End If
   'end 2008/3/25
   
   Printer.FontSize = 10
   Printer.Font.Name = "細明體"
   'Add By Cheng 2003/03/27
   m_dblPLeft = CDbl(Me.Text7.Text) * 567
   m_dblPTop = CDbl(Me.Text8.Text) * 567
   
   Do While adoacc0e0.EOF = False
      If adoacc0e0.Fields("a0e46").Value >= 1 Then
         strDelConfirm = MsgBox(MsgText(159) & adoacc0e0.Fields("a0e02").Value & MsgText(160), vbOKCancel + vbDefaultButton1, MsgText(5))
         If strDelConfirm = vbCancel Then
            GoTo NextSkip
         End If
      End If
      '日期
       Printer.FontSize = 12
        Printer.Font.Bold = True
'      Printer.CurrentX = 7000
      Printer.CurrentX = 7000 + 1620 + m_dblPLeft
'      Printer.CurrentY = 300
      Printer.CurrentY = 300 + 495 - 170 - 170 - 60 + m_dblPTop + 150
      'Modify by Morgan 2011/2/11 100年後會壓到年字，左移一格
      'Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
      If adoacc0e0.Fields("a0e10").Value > 1000000 Then
         Printer.Print Format(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
      Else
         Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
      End If
'      Printer.CurrentX = 7500
      Printer.CurrentX = 7500 + 1620 + 115 + m_dblPLeft
'      Printer.CurrentY = 300
      Printer.CurrentY = 300 + 495 - 170 - 170 - 60 + m_dblPTop + 150
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 5, 2))
'      Printer.CurrentX = 8000
      Printer.CurrentX = 8000 + 1620 + 115 + 60 + m_dblPLeft
'      Printer.CurrentY = 300
      Printer.CurrentY = 300 + 495 - 170 - 170 - 60 + m_dblPTop + 150
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 8, 2))
      '憑票支付
      Printer.FontSize = 12
      Printer.Font.Bold = True
      Printer.CurrentX = 5300 - 360 + m_dblPLeft
      Printer.CurrentY = 1000 - 190 + m_dblPTop + 150
      If IsNull(adoacc0e0.Fields("a0e12").Value) Then
         strName = ""
      Else
         strName = adoacc0e0.Fields("a0e12").Value
      End If
      If Len(strName) Mod 13 = 0 Then
          ii = Len(strName) / 13
      Else
          ii = Fix(Len(strName) / 13) + 1
      End If
      For jj = 1 To ii
          Printer.CurrentX = 5300 - 280 - 110 + m_dblPLeft
          Printer.CurrentY = 1000 - 115 + m_dblPTop - 55 - 230 * (ii - jj) + 150
          Printer.Print Mid(strName, (jj - 1) * 13 + 1, 13)
      Next jj
      'NT金額
       Printer.FontSize = 12
        Printer.Font.Bold = True
      'Modify by Morgan 2011/4/26 瑞婷
      'strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), DDollar)
      strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), FDollar)
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 4700 + m_dblPLeft
      Printer.CurrentY = 1940 + m_dblPTop + 150
      Printer.Print strAmount
      '新台幣
      Printer.FontSize = 14
      Printer.Font.Bold = True
      Printer.CurrentX = 5300 - 650 + m_dblPLeft
      Printer.CurrentY = 1800 - 410 + m_dblPTop + 150
      Printer.Print ChangeNumber(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value))
      Printer.FontSize = 10
      '票根發票日
       Printer.FontSize = 10
        Printer.Font.Bold = False
      Printer.CurrentX = 800 - 230 - 90 + m_dblPLeft
      Printer.CurrentY = 1500 - 510 - 230 - 20 + m_dblPTop
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e13").Value), 1, 3))
      Printer.CurrentX = 1500 - 400 - 90 + m_dblPLeft
      Printer.CurrentY = 1500 - 510 - 230 - 20 + m_dblPTop
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e13").Value), 5, 2))
      Printer.CurrentX = 2200 - 680 + m_dblPLeft
      Printer.CurrentY = 1500 - 510 - 230 - 20 + m_dblPTop
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e13").Value), 8, 2))
      '票根到期日
       Printer.FontSize = 10
        Printer.Font.Bold = False
'      Printer.CurrentX = 800
'      Printer.CurrentY = 1500
      Printer.CurrentX = 800 - 230 - 90 + m_dblPLeft
      Printer.CurrentY = 1500 - 510 + m_dblPTop
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
'      Printer.CurrentX = 1500
'      Printer.CurrentY = 1500
      Printer.CurrentX = 1500 - 400 - 90 + m_dblPLeft
      Printer.CurrentY = 1500 - 510 + m_dblPTop
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 5, 2))
'      Printer.CurrentX = 2200
'      Printer.CurrentY = 1500
      Printer.CurrentX = 2200 - 680 + m_dblPLeft
      Printer.CurrentY = 1500 - 510 + m_dblPTop
      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 8, 2))
      '票根受款人
       Printer.FontSize = 10
        Printer.Font.Bold = False
      Printer.CurrentX = 0 + m_dblPLeft
'      Printer.CurrentY = 1800
      Printer.CurrentY = 1800 - 170 + m_dblPTop - 220
        'Modify By Cheng 2003/04/01
        '取10個字
'      Printer.Print strName
        If Len(strName) Mod 10 = 0 Then
            ii = Len(strName) / 10
        Else
            ii = Fix(Len(strName) / 10) + 1
        End If
        For jj = 1 To ii
            Printer.CurrentX = 0 + m_dblPLeft
            Printer.CurrentY = 1800 - 170 + m_dblPTop - 220 + 190 * (jj - 1)
            Printer.Print Mid(strName, (jj - 1) * 10 + 1, 10)
        Next jj
        '票根金額
       Printer.FontSize = 10
        Printer.Font.Bold = False
      'Modify by Morgan 2011/4/26 瑞婷
      'strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), DDollar)
      strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), FDollar)
      
      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = 2000 - intLength + m_dblPLeft
      Printer.CurrentX = 570 + m_dblPLeft
'      Printer.CurrentY = 3200
      Printer.CurrentY = 3200 - 1250 + m_dblPTop
      Printer.Print strAmount
        'Add By Cheng 2003/04/02
        '禁止背書轉讓
       Printer.FontSize = 12
        Printer.Font.Bold = True
        Printer.CurrentX = 6080 + m_dblPLeft
        Printer.CurrentY = 2700 + m_dblPTop
        Printer.Print "禁止背書轉讓"
'      Printer.CurrentX = 0
'      Printer.CurrentY = 3800
'      Printer.Print IIf(IsNull(adoacc0e0.Fields("a0e12").Value), MsgText(601), adoacc0e0.Fields("a0e12").Value)
      adoTaie.Execute "update acc0e0 set a0e46 = 1 where a0e01 = '" & adoacc0e0.Fields("a0e01").Value & "' and a0e02 = '" & adoacc0e0.Fields("a0e02").Value & "'"
      Printer.NewPage
NextSkip:
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   Printer.EndDoc
   Me.MousePointer = vbDefault
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生支票資料並列印 (台北國際商業銀行--支票)
'  Memo by Amy 2022/02/24 目前未使用-與瑞婷確認
'*************************************************
Private Sub ProcessData3()
   
'On Error GoTo Checking
'   Me.MousePointer = vbHourglass
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'
'   adoacc0e0.CursorLocation = adUseClient
'   'adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('0149950', '0149980') and a0e08 = '1' and (a0e46 is null or a0e46 = 0)" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'   adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e07 in ('02369300', '02369900') and a0e08 = '1'" & strSql & " order by a0e02 asc, a0e26 asc, a0e27 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0e0.RecordCount = 0 Then
'      adoacc0e0.Close
'      StatusClear
'      Me.MousePointer = vbDefault
'      Exit Sub
'   End If
'
'   'Modify by Morgan 2008/3/25 控制 9x 才自訂
'   If pub_OS = "1" Then
'      Printer.Height = 5800
'      Printer.Width = 14000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(8)
'   End If
'   'end 2008/3/25
'   Printer.FontSize = 10
'   'Add by Morgan 2004/7/29
'   Printer.Font.Name = "細明體"
'
'   Do While adoacc0e0.EOF = False
'      If adoacc0e0.Fields("a0e46").Value >= 1 Then
'         strDelConfirm = MsgBox(MsgText(159) & adoacc0e0.Fields("a0e02").Value & MsgText(160), vbOKCancel + vbDefaultButton1, MsgText(5))
'         If strDelConfirm = vbCancel Then
'            GoTo NextSkip
'         End If
'      End If
'      Printer.CurrentX = 7000
'      Printer.CurrentY = 300
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
'      Printer.CurrentX = 7500
'      Printer.CurrentY = 300
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 5, 2))
'      Printer.CurrentX = 8000
'      Printer.CurrentY = 300
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 8, 2))
'      Printer.CurrentX = 5300
'      Printer.CurrentY = 1000
'      'adoquery.CursorLocation = adUseClient
'      'adoquery.Open "select * from acc0q0 where a0q01 = " & Val(IIf(IsNull(adoacc0e0.Fields("a0e03").Value), "", adoacc0e0.Fields("a0e03").Value)) & " and a0q03 = '" & IIf(IsNull(adoacc0e0.Fields("a0e06").Value), "", adoacc0e0.Fields("a0e06").Value) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      'If adoquery.RecordCount <> 0 Then
'      '   If IsNull(adoquery.Fields("a0q05").Value) Then
'      '      Select Case adoacc0e0.Fields("a0e05").Value
'      '         Case "1"
'      '            strName = CustomerQuery(adoacc0e0.Fields("a0e06").Value, 1)
'      '         Case "2"
'      '            strName = A0i02Query(adoacc0e0.Fields("a0e06").Value)
'      '         Case "3"
'      '            strName = StaffQuery(adoacc0e0.Fields("a0e06").Value)
'      '         Case Else
'      '            strName = ""
'      '      End Select
'      '   Else
'      '      strName = adoquery.Fields("a0q05").Value
'      '   End If
'      'Else
'      '   Select Case adoacc0e0.Fields("a0e05").Value
'      '      Case "1"
'      '         strName = CustomerQuery(adoacc0e0.Fields("a0e06").Value, 1)
'      '      Case "2"
'      '         strName = A0i02Query(adoacc0e0.Fields("a0e06").Value)
'      '      Case "3"
'      '         strName = StaffQuery(adoacc0e0.Fields("a0e06").Value)
'      '      Case Else
'      '         strName = ""
'      '   End Select
'      'End If
'      'adoquery.Close
'      If IsNull(adoacc0e0.Fields("a0e12").Value) Then
'         strName = ""
'      Else
'         strName = adoacc0e0.Fields("a0e12").Value
'      End If
'      Printer.Print strName
'      strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), DDollar)
'      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = 10200 - intLength
'      Printer.CurrentY = 1000
'      Printer.Print strAmount
'      Printer.FontSize = 12
'      Printer.CurrentX = 5300
'      Printer.CurrentY = 1800
'      Printer.Print ChangeNumber(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value))
'      Printer.FontSize = 10
'      Printer.CurrentX = 800
'      Printer.CurrentY = 1500
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 1, 3))
'      Printer.CurrentX = 1500
'      Printer.CurrentY = 1500
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 5, 2))
'      Printer.CurrentX = 2200
'      Printer.CurrentY = 1500
'      Printer.Print Val(Mid(CFDate(adoacc0e0.Fields("a0e10").Value), 8, 2))
'      Printer.CurrentX = 0
'      Printer.CurrentY = 1800
'      Printer.Print strName
'      Printer.CurrentX = 0
'      Printer.CurrentY = 2100
'      Printer.Print IIf(IsNull(adoacc0e0.Fields("a0e12").Value), MsgText(601), adoacc0e0.Fields("a0e12").Value)
'      strAmount = Format(IIf(IsNull(adoacc0e0.Fields("a0e11").Value), 0, adoacc0e0.Fields("a0e11").Value), DDollar)
'      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = 2000 - intLength
'      Printer.CurrentY = 3500
'      Printer.Print strAmount
'      adoTaie.Execute "update acc0e0 set a0e46 = 1 where a0e01 = '" & adoacc0e0.Fields("a0e01").Value & "' and a0e02 = '" & adoacc0e0.Fields("a0e02").Value & "'"
'      Printer.NewPage
'NextSkip:
'      adoacc0e0.MoveNext
'   Loop
'   adoacc0e0.Close
'   Printer.EndDoc
'   Me.MousePointer = vbDefault
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add by Amy 2022/02/24
Private Sub SetPrinter(ByVal bolReCovery As Boolean)
    If bolReCovery = False Then
        '切換印表機
        PUB_SetOsDefaultPrinter Combo1
        PUB_RestorePrinter Combo1
    Else
        '還原印表機
        PUB_SetOsDefaultPrinter strPrinter
        PUB_RestorePrinter strPrinter
    End If
End Sub
