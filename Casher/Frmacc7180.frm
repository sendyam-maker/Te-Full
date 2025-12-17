VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc7180 
   AutoRedraw      =   -1  'True
   Caption         =   "分所智權人員收款明細表-智權人員繳款"
   ClientHeight    =   3615
   ClientLeft      =   3660
   ClientTop       =   1875
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   5640
   Begin VB.TextBox Text3 
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
      Left            =   2490
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "Y"
      Top             =   2010
      Width           =   525
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1485
      TabIndex        =   12
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox Text2 
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1590
      Width           =   945
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   195
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   330
      Value           =   -1  'True
      Width           =   195
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
      Left            =   450
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2610
      Width           =   4692
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
      Height          =   315
      Left            =   1950
      MaxLength       =   5
      TabIndex        =   1
      Top             =   360
      Width           =   945
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1950
      TabIndex        =   3
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
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
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
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
   Begin VB.Label Label7 
      Caption         =   "資料內容：當月的第二個工作日~至隔月的第一個工作日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "是否列印明細資料：         (Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2040
      Width           =   3795
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   " 印表機"
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
      Left            =   450
      TabIndex        =   13
      Top             =   3150
      Width           =   975
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
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
      Left            =   2280
      TabIndex        =   11
      Top             =   1620
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   10
      Top             =   1620
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   1110
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "出納確認日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   8
      Top             =   1110
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "出納確認年月                  (Ex : 10301)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   7
      Top             =   420
      Width           =   4635
   End
End
Attribute VB_Name = "Frmacc7180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add by Lydia 2014/10/3 分所智權人員收款明細表-智權人員繳款
'Memo by Lydia 2015/07/16 因為報表欄位時有修改,將程式改成變數控制
Option Explicit
Public adoacc310 As New ADODB.Recordset
Dim M7180 As New ADODB.Recordset
Dim strSql, strNo, strSalesNo, strSalesName As String
Dim intLength As Integer
Dim intCounter As Integer
Dim intPage As Integer
'Modified by Lydia 2015/07/16
'Dim PLeft(0 To 19) As Integer
'Dim SPLeft(0 To 13) As Integer
'Dim strTemp(0 To 20) As String
'Dim m_sColumn(0 To 17) As String '明細欄位名稱
'Dim Sm_sColumn(1 To 12) As String '總計欄位名稱
'Dim lngSubTot() As String, lngTot(1 To 12) As Double  '小計,合計
    'Modified by Lydia 2015/07/16 設定欄位數
    Private Const LR1 = 13 '13 '第1行的欄位數(從0開始)
    Private Const LR2 = 4  '第2行的欄位數(從0開始)
    Private Const LRt = 13 '尾頁總計的欄位數(從0-不顯示的業務代號開始)
    Dim PLeft(0 To LR1 + LR2 + 3) As Integer
    Dim SPLeft(0 To LRt + 1) As Integer
    Dim strTemp() As String
    Dim m_sColumn(0 To LR1 + LR2 + 1) As String '明細欄位名稱(第1行+第2行)
    Dim Sm_sColumn(1 To LRt) As String '總計欄位名稱(從1開始,這樣才與明細的欄位一致)
    Dim lngSubTot() As String, lngTot(1 To LRt) As Double  '小計,合計
    Dim sChk4401 As String  '員工代號
    Dim sChk4402 As Long  '繳款時間
    Dim sChk4403 As Long '繳款日期
    Dim mDiff As Integer '溢收款的位置
    Dim mDot As Double '點數
    'end 2015/07/16
'列印用
Dim iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 14, ciFontSize = 10
'Modified by Lydia 2015/07/16
'Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Private Const ciStartX = 500
Private Const ciStartY = 500
Private Const ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim prnPrint As Printer
Dim strPrint As String
Dim mTitle As Integer  '判斷頁面第一筆資料位置

Dim mChk4401 As String, mChk4402 As Long, mChk4403 As Long '單據區間判斷
Dim mSName As String, mNo01 As Integer, detailChk As Boolean   '群組換頁,小計陣列位置,是否列印明細

Private Sub Command1_Click()
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    For Each prnPrint In Printers
       If prnPrint.DeviceName = Combo1 Then
          Set Printer = prnPrint
       End If
    Next

    DoPrintRun
    For Each prnPrint In Printers
       If prnPrint.DeviceName = strPrint Then
          Set Printer = prnPrint
       End If
    Next
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = "列印A4報表"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
    If KeyCode <> vbKeyEscape Then
        Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
    End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 5760
    Me.Height = 4020
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath4)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    Option1_Click 0
    Text1 = Left(strSrvDate(1), 6) - 191100
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = DFormat
    Frmacc0000.StatusBar1.Panels(1).Text = "列印A4報表"
    SendKeys "{Tab}"
   strPrint = Printer.DeviceName
   For Each prnPrint In Printers
         Combo1.AddItem prnPrint.DeviceName
      If Combo1 = "" Then
         Combo1 = Printer.DeviceName
      End If
   Next
    '設定列印印表機
    StrSQLa = "Select * From PrintStartPoint Where PSP01='" & strUserNum & "' And PSP02='" & Me.Name & "' And PSP03='" & Me.Combo1.Name & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若有資料
    If rsA.RecordCount > 0 Then
        If Me.Combo1.ListCount > 0 Then
            For ii = 0 To Me.Combo1.ListCount - 1
                If Me.Combo1.List(ii) = "" & rsA("PSP06").Value Then
                    Me.Combo1.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
        '記錄原設定值
        Me.Combo1.Tag = Me.Combo1.Text
    '若無資料
    Else
        '記錄原設定值
        Me.Combo1.Tag = ""
    End If
    
SetColumnName
SSetColumnName

    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    StatusClear
    Set Frmacc7180 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If MaskEdBox1.Text = MsgText(29) Then
        Exit Sub
    End If
    MaskEdBox2.Mask = ""
    MaskEdBox2.Mask = DFormat
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
    Case 0
        Me.Text1.Enabled = True
        Me.Option1(1).Value = False
        Me.MaskEdBox1.Enabled = False
        Me.MaskEdBox2.Enabled = False
    Case 1
        Me.MaskEdBox1.Enabled = True
        Me.MaskEdBox2.Enabled = True
        Me.Option1(0).Value = False
        Me.Text1.Enabled = False
    End Select
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Me.Text1.Text <> "" Then
        If IsDate(Format(Val(Me.Text1.Text & "01") + 19110000, "####/##/##")) = False Then
            MsgBox "收款年月輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
        End If
    End If
    If Cancel = True Then Text1_GotFocus
End Sub


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
    If Me.Option1(0).Value = True Then
        Text1.SetFocus
    Else
        Me.MaskEdBox1.SetFocus
    End If
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
    If Me.Option1(0).Value = True Then
        If Text1 <> MsgText(601) Then
           FormCheck = True
           Exit Function
        End If
    End If
    If Me.Option1(1).Value = True Then
        If MaskEdBox1.Text <> MsgText(29) Then
           FormCheck = True
           Exit Function
        End If
        If MaskEdBox2.Text <> MsgText(29) Then
           FormCheck = True
           Exit Function
        End If
    End If
    FormCheck = False
End Function

Private Sub Text2_GotFocus()
    TextInverse Me.Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Me.Text2.Text = "" Then Me.lblSalesName.Caption = "": Exit Sub
    Me.lblSalesName.Caption = GetStaffName(Me.Text2.Text)
    If Me.lblSalesName.Caption = "" Then
        MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
        Cancel = True
    End If
    If Cancel = True Then Text2_GotFocus
End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 78 And KeyAscii <> 89 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Text3.Text = "" Then
        MsgBox "是否列印明細，請輸入 Y 或 N！", , "錯誤！"
        Text3.SetFocus
        Cancel = True
    End If
End Sub


Private Sub GetPleft() '明細表邊界

Printer.Font.Name = "新細明體"
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

Erase PLeft

   PLeft(0) = ciStartX
   'Modified by Lydia 2015/07/16
'   PLeft(19) = lngPageWidth - ciStartX * 2  '出納備註的邊界=列印右邊界
'For intI = 1 To 18
'   Select Case intI
'
'    Case 1 To 12, 17, 18
'        PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(4, "　")) + ciColGap
'    Case 13 '收據號碼起
'        PLeft(intI) = PLeft(1)
'
'    Case 14  '收據號碼止,客戶名稱起
'        PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(5, "　")) + ciColGap
'
'    Case 15
'        PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(12, "　")) + ciColGap
'
'    Case 16
'        PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(14, "　")) + ciColGap
'   End Select
'Next intI

   PLeft(LR1 + LR2 + 3) = lngPageWidth - ciStartX
For intI = 1 To LR1 + LR2 + 2
   If intI > LR1 And intI <= LR1 + 4 Then
      If intI = LR1 + 1 Then
         PLeft(intI) = PLeft(1) '收據號碼起-換行
      Else
        Select Case intI - LR1 - 1
         Case 1  '收據號碼止,客戶名稱起
             PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(5, "　")) + ciColGap
         Case 2
             PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(12, "　")) + ciColGap
         Case 3
             PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(14, "　")) + ciColGap
        End Select
      End If
   Else
        PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(4, "　")) + ciColGap
   End If
Next intI

End Sub

Private Sub GetSPleft() '總計表邊界

Erase SPLeft

   SPLeft(0) = ciStartX
'Modified by Lydia 2015/07/16
'For intI = 1 To 13
For intI = 1 To LRt
   '縮小-姓名、其他、溢收款、手續費、補扣繳、外幣
    If intI = 1 Or (intI >= mDiff - 1 And intI < mDiff + 4) Then
       SPLeft(intI) = SPLeft(intI - 1) + Printer.TextWidth(String(4, "　")) + ciColGap
    Else
       SPLeft(intI) = SPLeft(intI - 1) + Printer.TextWidth(String(5, "　")) + ciColGap
    End If
Next intI

End Sub


Private Sub SetColumnName()
'Modified by Lydia 2015/07/16
'   m_sColumn(0) = "確認日期"
'   m_sColumn(1) = "繳款人"
'   m_sColumn(2) = "票據金額"
'   m_sColumn(3) = "北所電匯"
'   m_sColumn(4) = "分所電匯"
'   m_sColumn(5) = "現金"
'   m_sColumn(6) = "抵暫收款"
'   m_sColumn(7) = "溢收款"
'   m_sColumn(8) = "手續費"
'   m_sColumn(9) = "補扣繳"
'   m_sColumn(10) = "外幣"
'   m_sColumn(11) = "會計簽名"
'   m_sColumn(12) = "出納備註"
'   m_sColumn(13) = "收據號碼"
'   m_sColumn(14) = "客戶名稱"
'   m_sColumn(15) = "案件性質"
'   m_sColumn(16) = "點數"
'   m_sColumn(17) = "扣繳金額"
Dim TmpArr As Variant
Dim TmpStr As String

   TmpStr = "確認日期|繳款人|票據金額|北所電匯|分所電匯|現金|抵暫收款|其他|溢收款|手續費|補扣繳|外幣|會計簽名|出納備註"
   TmpStr = TmpStr & "|收據號碼|客戶名稱|案件性質|點數|扣繳金額"
   TmpArr = Split(TmpStr, "|")
   
   For intI = 0 To UBound(TmpArr)
       m_sColumn(intI) = TmpArr(intI)
   Next intI
End Sub

Private Sub SSetColumnName()
'Modified by Lydia 2015/07/16
'   Sm_sColumn(1) = "繳款人"
'   Sm_sColumn(2) = "票據金額"
'   Sm_sColumn(3) = "北所電匯"
'   Sm_sColumn(4) = "分所電匯"
'   Sm_sColumn(5) = "現金"
'   Sm_sColumn(6) = "抵暫收款"
'   Sm_sColumn(7) = "溢收款"
'   Sm_sColumn(8) = "手續費"
'   Sm_sColumn(9) = "補扣繳"
'   Sm_sColumn(10) = "外幣"
'   Sm_sColumn(11) = "合計"
'   Sm_sColumn(12) = "點數"
Dim TmpArr As Variant
Dim TmpStr As String

   TmpStr = "繳款人|票據金額|北所電匯|分所電匯|現金|抵暫收款|其他|溢收款|手續費|補扣繳|外幣|合計|點數"
   TmpArr = Split(TmpStr, "|")
   
   For intI = 0 To UBound(TmpArr)
       '從1開始
       Sm_sColumn(intI + 1) = TmpArr(intI)
   Next intI
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader

   End If
    
End Sub

Private Sub PrintNewSub(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
Dim xPrt As Long
xPrt = iPrint

   xPrt = xPrt + lngLineHeight * 2
   If xPrt >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = xPrt

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader

   End If
    
End Sub

Private Sub PrintLine(mKind As Integer)
   Dim iNo As Integer

   Printer.CurrentX = ciStartX
   If mKind = 2 Then
     'Modified by Lydia 2015/07/16
     'iNo = (PLeft(19) - ciStartX) \ Printer.TextWidth("=")
     iNo = (PLeft(LR1 + LR2 + 3) - ciStartX) \ Printer.TextWidth("=")
     Printer.Print String(iNo, "=")
   Else
    'Modified by Lydia 2015/07/16
    'iNo = (PLeft(19) - ciStartX) \ Printer.TextWidth("-")
    iNo = (PLeft(LR1 + LR2 + 3) - ciStartX) \ Printer.TextWidth("-")
     Printer.Print String(iNo - 2, "-")
   End If

   iPrint = iPrint + 150

End Sub

Private Sub WriteSSum(m_SState As Integer, m_Ix As Integer) '寫入小計2維陣列
Dim m_In As Integer, m_Si As Integer, m_Stot As Long
Dim m_Rnum As Integer, m_Rwrite As Boolean

m_Stot = 0
                      
'Modified by Lydia 2015/07/16
'If mChk4401 = strTemp(18) And mChk4402 = strTemp(19) And mChk4403 = strTemp(20) Then
If mChk4401 = sChk4401 And mChk4402 = sChk4402 And mChk4403 = sChk4403 Then
'主檔只計算一次金額,明細點數需要累計
       m_Rwrite = False
       m_Rnum = 0
       For m_In = 0 To m_Ix - 1
           'Modified by Lydia 2015/07/16
'          If LTrim(RTrim(lngSubTot(m_In, 0))) = LTrim(RTrim(strTemp(18))) Then  '已有小計記錄
'            lngSubTot(m_In, 12) = Val(lngSubTot(m_In, 12)) + Val(strTemp(16)) '個人小計-點數
'            lngTot(12) = lngTot(12) + Val(strTemp(16)) '總計-點數
          If LTrim(RTrim(lngSubTot(m_In, 0))) = LTrim(RTrim(sChk4401)) Then
            lngSubTot(m_In, LRt) = Val(lngSubTot(m_In, LRt)) + mDot  '個人小計-點數
            lngTot(LRt) = lngTot(LRt) + mDot '總計-點數
            m_Rwrite = True
          Else
            If Len(lngSubTot(m_In, 1)) = 0 Then '查無記錄->寫入新陣列
               Exit For
            End If
          End If
       Next m_In
          
           
Else
    
    If m_SState = 1 Then '第一個智權人小計
      Erase lngSubTot
      'Modified by Lydia 2015/07/16
'      ReDim lngSubTot(0 To m_Ix - 1, 0 To 12)
'        lngSubTot(0, 0) = LTrim(RTrim(strTemp(18))) '小計-代號
'        lngSubTot(0, 1) = LTrim(RTrim(strTemp(1))) '小計-繳款人
'        For m_Si = 2 To 10
      ReDim lngSubTot(0 To m_Ix - 1, 0 To LRt)
        lngSubTot(0, 0) = LTrim(RTrim(sChk4401)) '小計-代號
        lngSubTot(0, 1) = LTrim(RTrim(strTemp(1))) '小計-繳款人
        For m_Si = 2 To LRt - 2
           lngSubTot(0, m_Si) = Val(strTemp(m_Si))
           m_Stot = m_Stot + Val(strTemp(m_Si))
        Next m_Si
        'Modified by Lydia 2015/07/16
'        lngSubTot(0, 11) = m_Stot
'        lngSubTot(0, 12) = strTemp(16)
        lngSubTot(0, LRt - 1) = m_Stot
        lngSubTot(0, LRt) = mDot
    Else
       m_Rwrite = False
       m_Rnum = 0
       For m_In = 0 To m_Ix - 1
          'Modified by Lydia 2015/07/16
'          If LTrim(RTrim(lngSubTot(m_In, 0))) = LTrim(RTrim(strTemp(18))) Then '已有小計記錄
'            For m_Si = 2 To 10
          If LTrim(RTrim(lngSubTot(m_In, 0))) = LTrim(RTrim(sChk4401)) Then '已有小計記錄
            For m_Si = 2 To LRt - 2
             lngSubTot(m_In, m_Si) = Val(lngSubTot(m_In, m_Si)) + Val(strTemp(m_Si))
             'Modified by Lydia 2015/07/16
             'If m_Si = 7 Then '減溢收款
             If m_Si = mDiff Then '減溢收款
               m_Stot = m_Stot - Val(strTemp(m_Si))
             Else
               m_Stot = m_Stot + Val(strTemp(m_Si))
             End If
            Next m_Si
             'Modified by Lydia 2015/07/16
'            lngSubTot(m_In, 11) = Val(lngSubTot(m_In, 11)) + m_Stot
'            lngSubTot(m_In, 12) = Val(lngSubTot(m_In, 12)) + Val(strTemp(16))
            lngSubTot(m_In, LRt - 1) = Val(lngSubTot(m_In, LRt - 1)) + m_Stot
            lngSubTot(m_In, LRt) = Val(lngSubTot(m_In, LRt)) + mDot
            m_Rwrite = True
          Else
            If Len(lngSubTot(m_In, 1)) = 0 Then '查無記錄->寫入新陣列
               m_Rnum = m_In
               Exit For
            End If
          End If
       Next m_In
          
       If m_Rnum > 0 And m_Rwrite = False Then '寫入新陣列
        'Modified by Lydia 2015/07/16
'        lngSubTot(m_Rnum, 0) = LTrim(RTrim(strTemp(18))) '小計-代號
        lngSubTot(m_Rnum, 0) = LTrim(RTrim(sChk4401)) '小計-代號
        lngSubTot(m_Rnum, 1) = LTrim(RTrim(strTemp(1))) '小計-繳款人
        'Modified by Lydia 2015/07/16
'         For m_Si = 2 To 10
'          lngSubTot(m_Rnum, m_Si) = Val(strTemp(m_Si))
'            If m_Si = 7 Then '減溢收款
         For m_Si = 2 To LRt - 2
          lngSubTot(m_Rnum, m_Si) = Val(strTemp(m_Si))
            If m_Si = mDiff Then '減溢收款
             m_Stot = m_Stot - Val(strTemp(m_Si))
            Else
             m_Stot = m_Stot + Val(strTemp(m_Si))
            End If
           
         Next m_Si
         'Modified by Lydia 2015/07/16
'         lngSubTot(m_Rnum, 11) = m_Stot
'         lngSubTot(m_Rnum, 12) = strTemp(16)
         lngSubTot(m_Rnum, LRt - 1) = m_Stot
         lngSubTot(m_Rnum, LRt) = mDot
       End If
       
    End If
    If detailChk = False Then '不印明細，判斷值提前變動
        'Modified by Lydia 2015/07/16
'        mChk4401 = strTemp(18)
'        mChk4402 = strTemp(19)
'        mChk4403 = strTemp(20)
        mChk4401 = sChk4401
        mChk4402 = sChk4402
        mChk4403 = sChk4403
    End If
    WriteTSum
End If
End Sub

Private Sub WriteTSum() '寫入總計陣列
Dim m_TIn As Integer, m_Ttot As Long

m_Ttot = 0
lngTot(1) = "0000" '總計
'Modified by Lydia 2015/07/16
'   For m_TIn = 2 To 12
'    If m_TIn < 11 Then
'      lngTot(m_TIn) = lngTot(m_TIn) + Val(strTemp(m_TIn))
'      If m_TIn = 7 Then '減溢收款
   For m_TIn = 2 To LRt
    If m_TIn <= LRt - 2 Then
      lngTot(m_TIn) = lngTot(m_TIn) + Val(strTemp(m_TIn))
      If m_TIn = mDiff Then '減溢收款
         m_Ttot = m_Ttot - Val(strTemp(m_TIn))
      Else
         m_Ttot = m_Ttot + Val(strTemp(m_TIn))
      End If
     
    Else
      'Modified by Lydia 2015/07/16
'      If m_TIn = 11 Then lngTot(m_TIn) = lngTot(m_TIn) + m_Ttot '合計
'      If m_TIn = 12 Then lngTot(m_TIn) = lngTot(m_TIn) + Val(strTemp(16)) '點數
       Select Case LRt - m_TIn
          Case 1 '合計
                 lngTot(m_TIn) = lngTot(m_TIn) + m_Ttot
          Case 0 '點數
                 lngTot(m_TIn) = lngTot(m_TIn) + mDot
       End Select
    End If
   Next m_TIn


End Sub

Private Sub DoPrintRun()
Dim tmpRec As New ADODB.Recordset
Dim strA As String, sRec As Integer
Dim idR As Integer 'Added by Lydia 2015/07/16

Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
Printer.PaperSize = 9  'A4
   
lngPageHeight = Printer.ScaleHeight
lngPageWidth = Printer.ScaleWidth
lngLineHeight = 300
strSql = ""

    If Me.Option1(0).Value = True Then
       'Modified by Lydia 2015/01/07 出納確認年月改當月的第二個工作日~至隔月的第一個工作日
        If Text1 <> MsgText(601) Then
           ' strSql = strSql & " And A4413>=" & Val(Me.Text1.Text & "01") + 19110000 & " And A4413<=" & Val(Me.Text1.Text & "31") + 19110000 & " "
            strExc(0) = Val(Me.Text1.Text & "01") + 19110000
            strExc(0) = PUB_GetWorkDay1(CompDate(2, 1, PUB_GetWorkDay1(strExc(0), False)), False)
            strExc(1) = Val(Me.Text1.Text & "01") + 19110000
            strExc(1) = PUB_GetWorkDay1(CompDate(1, 1, strExc(1)), False)
            strSql = strSql & " And A4413>=" & strExc(0) & " and A4413 <= " & strExc(1) & " "
        End If
    Else
        If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " And A4413>=" & Val(FCDate(MaskEdBox1.Text)) + 19110000 & " "
        End If
        If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " And A4413<=" & Val(FCDate(MaskEdBox2.Text)) + 19110000 & " "
        End If
    End If

    If Me.Text2.Text <> "" Then
        strSql = strSql & " And A4401='" & Me.Text2.Text & "' "
    End If

    If Text3.Text = "Y" Then
        detailChk = True
    Else
        detailChk = False
    End If

    strSql = strSql + " and st06='" & pub_strUserOffice & "' "

'---智權人員人數(陣列定義)
strA = " select count(distinct(a4401)) as RR1 From acc441, acc440, staff Where a4401=axd01(+) and a4402=axd02(+) and a4403=axd03(+) and a4401=st01(+) " & strSql

If tmpRec.State = 1 Then tmpRec.Close
tmpRec.CursorLocation = adUseClient
tmpRec.Open strA, cnnConnection, adOpenStatic, adLockReadOnly
sRec = tmpRec!rr1

 '注意排序A4401+A4402+A4403(單據號碼)
 'Added by Lydia 2015/07/16 +A4430
strSql = strSql & " order by a4401,a4413,a4402,a4403, st06,axd04 " '--智權人,確認日期
strA = "select st06,A4413,A4401,ST02,A4402,A4403,A4404,A4415, " & _
       "nvl(A4405,0) as A4405,nvl(A4406,0) as A4406,nvl(A4407,0) as A4407,nvl(A4408,0) as A4408,nvl(A4409,0) as A4409,nvl(A4410,0) as A4410, " & _
       "nvl(A4411,0) as A4411,nvl(A4422,0) as A4422,NVL(A4426,0) as A4426,NVL(A4430,0) as A4430,AXD04,A0K04,nvl(a0k09, 0) as A0K09,a0j02, " & _
       "nvl(A0j09,0) as A0j09,nvl(A0j10,0) as A0j10,getcp10desc(cp01,cp10,a0j04) As 案件性質 , " & _
       "(AXD06/1000) as addot, AXD08 From acc440, acc441, staff, acc0k0, ACC0J0, caseprogress " & _
       "Where a4401=axd01(+) and a4402=axd02(+) and a4403=axd03(+) and a4401=st01(+)  " & _
       "and AXD04=A0K01(+) and axd04=a0j13(+) AND AXD05=A0J01(+) and cp09(+)=AXD05  " & strSql
If M7180.State = 1 Then M7180.Close
M7180.CursorLocation = adUseClient
M7180.Open strA, cnnConnection, adOpenStatic, adLockReadOnly
iPage = 0
'Added by Lydia 2015/07/16
mDiff = 8
ReDim strTemp(0 To 18) As String '列印明細

GetPleft
GetSPleft
Erase lngTot
mChk4401 = ""
mChk4402 = 0
mChk4403 = 0

If Not M7180.EOF And Not M7180.BOF Then
    With M7180
       M7180.MoveFirst
       mSName = LTrim(RTrim(M7180!st02)) '智權人員-群組換頁
       mNo01 = 0
       If detailChk = True Then
         PrintHeader '列印表頭
       End If
       Do While Not M7180.EOF
           iPage = iPage + 1
           '智權人員-群組換頁
           If mSName <> LTrim(RTrim(M7180!st02)) And detailChk = True Then
              PrintSalesSum mNo01 '列印個人小計

              iPage = iPage + 1
              Printer.NewPage
              mSName = LTrim(RTrim(M7180!st02)) '智權人員-群組換頁
              mNo01 = mNo01 + 1
              PrintHeader
           End If
           '明細
           'Modified by Lydia 2015/07/16
'           strTemp(0) = ChangeTStringToTDateString(M7180!a4413 - 19110000)  '確認日期
'           strTemp(1) = Trim(M7180!st02) '繳款人(智權人員)
'           strTemp(18) = RTrim(LTrim(M7180!A4401)) '員工代號
'           strTemp(19) = Val(M7180!A4402) '繳款時間
'           strTemp(20) = Val(M7180!A4403) '繳款日期
'           strTemp(2) = Val(M7180!A4405) '--  票據金額
'           strTemp(3) = Val(M7180!A4406) '--  北所電匯
'           strTemp(4) = Val(M7180!A4407) '--  分所電匯
'           strTemp(5) = Val(M7180!A4408) '--  現金
'           strTemp(6) = Val(M7180!A4409) '--  抵暫收款
'           strTemp(7) = Val(M7180!A4410) '--  溢收款
'           strTemp(8) = Val(M7180!A4411) '--  手續費
'           strTemp(9) = Val(M7180!A4422) '--  補扣繳
'           strTemp(10) = Val(M7180!A4426) '--  外幣
'
'           strTemp(11) = "" '--  會計簽名(留白)
'           strTemp(12) = "" & StrConv(LeftB(StrConv(LTrim(RTrim(M7180!A4415)), vbFromUnicode), 24), vbUnicode) '--  出納備註
'           strTemp(13) = "" & M7180!axd04 '--  收據號碼
'            If M7180!a0k09 > 0 Then '繳款後, 出納(acc0k0)作廢
'               strTemp(13) = strTemp(13) & "(廢)"
'            End If
'           strTemp(14) = "" & StrConv(LeftB(StrConv(LTrim(RTrim(M7180!A0K04)), vbFromUnicode), 24), vbUnicode) '取中文混雜指定長度(中文2,英文1 byte)
'           strTemp(15) = "" & Trim(M7180!A0J02) & StrConv(LeftB(StrConv(LTrim(RTrim(M7180!案件性質)), vbFromUnicode), 26), vbUnicode) & Trim(str(M7180!A0j09 + M7180!A0J10))
'
'           strTemp(16) = Val(M7180!addot) '--  點數
'           strTemp(17) = Val(M7180!AXD08) '--  扣繳金額

           sChk4401 = RTrim(LTrim(M7180!A4401)) '員工代號
           sChk4402 = Val(M7180!A4402) '繳款時間
           sChk4403 = Val(M7180!A4403) '繳款日期
           idR = 0
           strTemp(idR) = ChangeTStringToTDateString(M7180!a4413 - 19110000)  '確認日期
           idR = idR + 1: strTemp(idR) = Trim(M7180!st02) '繳款人(智權人員)
           idR = idR + 1: strTemp(idR) = Val(M7180!A4405) '--  票據金額
           idR = idR + 1: strTemp(idR) = Val(M7180!A4406) '--  北所電匯
           idR = idR + 1: strTemp(idR) = Val(M7180!A4407) '--  分所電匯
           idR = idR + 1: strTemp(idR) = Val(M7180!A4408) '--  現金
           idR = idR + 1: strTemp(idR) = Val(M7180!A4409) '--  抵暫收款
           idR = idR + 1: strTemp(idR) = Val(M7180!A4430) '--  其他
           idR = idR + 1: strTemp(idR) = Val(M7180!A4410) '--  溢收款
           'mDiff = idR '記錄溢收款位置
           idR = idR + 1: strTemp(idR) = Val(M7180!A4411) '--  手續費
           idR = idR + 1: strTemp(idR) = Val(M7180!A4422) '--  補扣繳
           idR = idR + 1: strTemp(idR) = Val(M7180!A4426) '--  外幣
           idR = idR + 1: strTemp(idR) = "" '--  會計簽名(留白)
           idR = idR + 1: strTemp(idR) = "" & StrConv(LeftB(StrConv(LTrim(RTrim(M7180!A4415)), vbFromUnicode), 24), vbUnicode) '--  出納備註
           idR = idR + 1: strTemp(idR) = "" & M7180!axd04 '--  收據號碼
            If M7180!a0k09 > 0 Then '繳款後, 出納(acc0k0)作廢
               strTemp(idR) = strTemp(idR) & "(廢)"
            End If
           idR = idR + 1: strTemp(idR) = "" & StrConv(LeftB(StrConv(LTrim(RTrim(M7180!A0K04)), vbFromUnicode), 24), vbUnicode)  '取中文混雜指定長度(中文2,英文1 byte)
           idR = idR + 1: strTemp(idR) = "" & Trim(M7180!A0J02) & StrConv(LeftB(StrConv(LTrim(RTrim(M7180!案件性質)), vbFromUnicode), 26), vbUnicode) & Trim(str(M7180!A0j09 + M7180!A0J10))

           idR = idR + 1: strTemp(idR) = Val(M7180!addot)  '--  點數
           mDot = Val(M7180!addot)
           idR = idR + 1: strTemp(idR) = Val(M7180!AXD08)  '--  扣繳金額

           If .AbsolutePosition = 1 Then
              Call WriteSSum(1, sRec) '首筆重設陣列
           Else
              Call WriteSSum(2, sRec)
           End If
            
           If detailChk = True Then
            PrintDetailRun '列印明細
           End If
            M7180.MoveNext
            
        Loop

    End With


'列印總計和表尾
If detailChk = True Then
PrintSalesSum mNo01 '末筆列印個人小計
PrintLine 1
Printer.NewPage
End If
PrintHeaderT
'----逐筆列印個人小計
For intI = 0 To sRec - 1
   PrintDetailT (intI)
   PrintLine 1
Next intI

PrintSum

PrintLine 2
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If


Printer.EndDoc
ShowPrintOk
End Sub

Private Sub PrintHeader()
Dim strPTmp As String
Dim pa1 As Integer
iPrint = ciStartY

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = False
'title line=1
'Modified by Lydia 2015/07/16
'Printer.CurrentX = PLeft(3) + 500
strPTmp = IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & _
              IIf(Me.Option1(0).Value = True, " " & Left(Me.Text1.Text, Len(Me.Text1.Text) - 2) & " 年 " & Right(Me.Text1.Text, 2) & " 月 份", " " & Me.MaskEdBox1.Text & " ~ " & Me.MaskEdBox2.Text & " ") & "每日收款明細表-智權人員繳款"
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
'Modified by Lydia 2015/07/16
'Printer.Print IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & _
              IIf(Me.Option1(0).Value = True, " " & Left(Me.Text1.Text, Len(Me.Text1.Text) - 2) & " 年 " & Right(Me.Text1.Text, 2) & " 月 份", " " & Me.MaskEdBox1.Text & " ~ " & Me.MaskEdBox2.Text & " ") & "每日收款明細表-智權人員繳款"
Printer.Print strPTmp
'title line = 2
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False
PrintNewLine
'Modified by Lydia 2015/07/16
'Printer.CurrentX = PLeft(11)
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & CFDate(strSrvDate(2))
'title line = 3
PrintNewLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & mSName
'Modified by Lydia 2015/07/16
'Printer.CurrentX = PLeft(11)
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page
    
'title line = 4
PrintNewLine
'Modified by Lydia 2015/07/16
'For pa1 = 0 To 12
'  If pa1 >= 2 And pa1 <= 10 Then '金額類-置右
For pa1 = 0 To LR1
  If pa1 >= 2 And pa1 <= LR1 - 2 Then '金額類-置右
    Printer.CurrentX = PLeft(pa1 + 1) - Printer.TextWidth(m_sColumn(pa1)) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print m_sColumn(pa1)
  Else
    Printer.CurrentX = PLeft(pa1)
    Printer.CurrentY = iPrint
    Printer.Print m_sColumn(pa1)
  End If
Next pa1


'title line = 5 第二行資料
PrintNewLine
'Modified by Lydia 2015/07/16
'For pa1 = 13 To 17
'  If pa1 >= 16 Then '金額類-置右
For pa1 = LR1 + 1 To LR1 + LR2 + 1
  If pa1 >= LR1 + LR2 Then  '金額類-置右
    Printer.CurrentX = PLeft(pa1 + 1) - Printer.TextWidth(m_sColumn(pa1)) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print m_sColumn(pa1)
  Else
    Printer.CurrentX = PLeft(pa1)
    Printer.CurrentY = iPrint
    Printer.Print m_sColumn(pa1)
  End If
Next pa1

PrintNewLine

PrintLine 2

mTitle = iPrint
End Sub

Private Sub PrintDetailRun()
Dim aP1 As Integer, pB As String

'Modified by Lydia 2015/07/16
'If mChk4401 = strTemp(18) And mChk4402 = strTemp(19) And mChk4403 = strTemp(20) Then
If mChk4401 = sChk4401 And mChk4402 = sChk4402 And mChk4403 = sChk4403 Then

'主檔只印一筆
           
Else
    If Len(mChk4401) > 0 And mChk4402 > 0 And mChk4403 > 0 Then
      If iPrint <> mTitle Then '判斷頁面第一筆資料位置
       PrintLine 1 '區隔不同區間
      End If
    End If
        'Modified by Lydia 2015/07/16
'        For aP1 = 0 To 12
'        If aP1 >= 2 And aP1 <= 10 Then '金額類-置右
        For aP1 = 0 To LR1
        If aP1 >= 2 And aP1 <= LR1 - 2 Then '金額類-置右
          pB = Format(strTemp(aP1), DDollar2)
          Printer.CurrentX = PLeft(aP1 + 1) - Printer.TextWidth(pB) - ciColGap
          Printer.CurrentY = iPrint
          Printer.Print pB
        Else
          Printer.CurrentX = PLeft(aP1)
          Printer.CurrentY = iPrint
          Printer.Print strTemp(aP1)
        End If
        Next aP1
        
        PrintNewLine
        'Modified by Lydia 2015/07/16
'        mChk4401 = strTemp(18)
'        mChk4402 = strTemp(19)
'        mChk4403 = strTemp(20)
        mChk4401 = sChk4401
        mChk4402 = sChk4402
        mChk4403 = sChk4403
End If

'第二行資料
'Modified by Lydia 2015/07/16
'For aP1 = 13 To 17
'  If aP1 >= 16 Then '金額類-置右
'    If aP1 = 16 Then
For aP1 = LR1 + 1 To LR1 + LR2 + 1
  If aP1 >= LR1 + LR2 Then  '金額類-置右
    If aP1 = LR1 + LR2 Then
      pB = Format(strTemp(aP1), "##,###,##0.000")
    Else
      pB = Format(strTemp(aP1), DDollar2)
    End If
    Printer.CurrentX = PLeft(aP1 + 1) - Printer.TextWidth(pB) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print pB
  Else
    Printer.CurrentX = PLeft(aP1)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(aP1)
  End If
Next aP1

PrintNewLine


End Sub


Private Sub PrintHeaderT()
Dim sa1 As Integer
Dim strPTmp As String 'Added by Lydia 2015/07/16

iPrint = ciStartY
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = False
'title line=1
'Modified by Lydia 2015/07/16
'Printer.CurrentX = PLeft(3) + 500
strPTmp = IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & _
              IIf(Me.Option1(0).Value = True, " " & Left(Me.Text1.Text, Len(Me.Text1.Text) - 2) & " 年 " & Right(Me.Text1.Text, 2) & " 月 份", " " & Me.MaskEdBox1.Text & " ~ " & Me.MaskEdBox2.Text & " ") & "每日收款明細表-智權人員繳款"
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
'Printer.Print IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & _
              IIf(Me.Option1(0).Value = True, " " & Left(Me.Text1.Text, Len(Me.Text1.Text) - 2) & " 年 " & Right(Me.Text1.Text, 2) & " 月 份", " " & Me.MaskEdBox1.Text & " ~ " & Me.MaskEdBox2.Text & " ") & "每日收款明細表-智權人員繳款"
Printer.Print strPTmp

'title line = 2
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False
PrintNewLine
'Modified by Lydia 2015/07/16
'Printer.CurrentX = PLeft(11)
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & CFDate(strSrvDate(2))
'title line = 3
PrintNewLine
'Modified by Lydia 2015/07/16
'Printer.CurrentX = PLeft(11)
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page
    
'title line = 4
PrintNewLine

'Modified by Lydia 2015/07/16
'For sa1 = 1 To 12
'  If sa1 > 1 And sa1 <= 12 Then '金額類-置右
For sa1 = 1 To LRt
  If sa1 > 1 And sa1 <= LRt Then '金額類-置右
    Printer.CurrentX = SPLeft(sa1) - Printer.TextWidth(Sm_sColumn(sa1)) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print Sm_sColumn(sa1)
  Else
    Printer.CurrentX = SPLeft(sa1 - 1)
    Printer.CurrentY = iPrint
    Printer.Print Sm_sColumn(sa1)
  End If
Next sa1

PrintNewLine

PrintLine 2

End Sub

Private Sub PrintDetailT(tR As Integer)
Dim aP3 As Integer, pSA As String
'Modified by Lydia 2015/07/16
'For aP3 = 1 To 12
'  If aP3 > 1 Then '金額類-置右
'    If aP3 = 12 Then
For aP3 = 1 To LRt
  If aP3 > 1 Then '金額類-置右
    If aP3 = LRt Then
      pSA = Format(Val(lngSubTot(tR, aP3)), "##,###,##0.000") '點數
    Else
      pSA = Format(Val(lngSubTot(tR, aP3)), DDollar2)
    End If
    Printer.CurrentX = SPLeft(aP3) - Printer.TextWidth(pSA) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print pSA
  Else
    Printer.CurrentX = SPLeft(aP3 - 1)
    Printer.CurrentY = iPrint
    Printer.Print lngSubTot(tR, aP3)
  End If
Next aP3

PrintNewLine

End Sub

Private Sub PrintSum()
Dim aP2 As Integer, pSB As String ', dSpace As Integer

Printer.CurrentX = SPLeft(0)
Printer.CurrentY = iPrint
Printer.Print "總計："
'Modified by Lydia 2015/07/16
'For aP2 = 2 To 12
''金額類-置右
'    If aP2 = 12 Then
For aP2 = 2 To LRt
'金額類-置右
    If aP2 = LRt Then
      pSB = Format(lngTot(aP2), "##,###,##0.000")
    Else
      pSB = Format(lngTot(aP2), DDollar2)
    End If
    Printer.CurrentX = SPLeft(aP2) - Printer.TextWidth(pSB) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print pSB

Next aP2

PrintNewLine

End Sub


Private Sub PrintSalesSum(Sno1 As Integer)
Dim sP2 As Integer, prB As String

PrintLine 1

PrintNewSub '檢查頁面空間是否足夠印２行

Printer.CurrentX = SPLeft(0)
Printer.CurrentY = iPrint
Printer.Print "小計："
'Modified by Lydia 2015/07/16
'For sP2 = 2 To 12
'  If sP2 > 1 And sP2 <= 12 Then '金額類-置右
For sP2 = 2 To LRt
  
    Printer.CurrentX = SPLeft(sP2) - Printer.TextWidth(Sm_sColumn(sP2)) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print Sm_sColumn(sP2)

Next sP2
    
PrintNewLine
'Modified by Lydia 2015/07/16
'For sP2 = 2 To 12
''金額類-置右
'    If sP2 = 12 Then
For sP2 = 2 To LRt
'金額類-置右
    If sP2 = LRt Then
      prB = Format(lngSubTot(Sno1, sP2), "##,###,##0.000")
    Else
      prB = Format(lngSubTot(Sno1, sP2), DDollar2)
    End If

    Printer.CurrentX = SPLeft(sP2) - Printer.TextWidth(prB) - ciColGap
    Printer.CurrentY = iPrint
    Printer.Print prB
Next sP2

End Sub


