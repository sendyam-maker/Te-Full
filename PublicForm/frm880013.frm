VERSION 5.00
Begin VB.Form frm880013 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "報表紙張格式設定"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox Check1 
      Caption         =   "Word Visible"
      Height          =   375
      Left            =   2205
      TabIndex        =   6
      Top             =   3540
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "顯示目前紙張代碼"
      Height          =   405
      Index           =   2
      Left            =   135
      TabIndex        =   5
      Top             =   3510
      Width           =   1770
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5445
      TabIndex        =   2
      Top             =   3510
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      ItemData        =   "frm880013.frx":0000
      Left            =   135
      List            =   "frm880013.frx":0002
      TabIndex        =   4
      Top             =   450
      Width           =   6390
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   855
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   90
      Width           =   5655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "開始(&S)"
      Height          =   400
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   90
      TabIndex        =   3
      Top             =   150
      Width           =   675
   End
End
Attribute VB_Name = "frm880013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'2010/8/18 sonia 日期欄已修改
'Create by Morgan 2008/3/27
Option Explicit
Dim m_DefaultPrinter As String

Private Sub Check1_Click()
  g_LetterDebug = IIf(Check1.Value = vbChecked, True, False) 'Added by Morgan 2017/9/5
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         SetPaperSize
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      Case 1
         Unload Me
      
      Case 2
         ShowPaperSize
         
   End Select
End Sub

Private Sub ShowPaperSize()
On Error Resume Next
   
   Dim PrnOld As String
   
   PrnOld = Printer.DeviceName
   PUB_RestorePrinter cmbPrinter.Text
   Printer.Orientation = 1
   Printer.EndDoc
   'Modified by Morgan 2018/7/17 +可印範圍等訊息
   MsgBox "印表機：" & Printer.DeviceName & vbCrLf & vbCrLf & "目前紙張代碼：" & Printer.PaperSize & vbCrLf & vbCrLf & "紙張大小(WxH)：" & Round(Printer.Width / 567, 1) & "cm x " & Round(Printer.Height / 567, 1) & "cm" & vbCrLf & vbCrLf & "可印大小(WxH)：" & Round(Printer.ScaleWidth / 567, 1) & "cm x " & Round(Printer.ScaleHeight / 567, 1) & "cm" & vbCrLf & vbCrLf & "ScaleMode：" & Printer.ScaleMode, vbInformation
   
   PUB_RestorePrinter PrnOld
   
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   Check1.Value = IIf(g_LetterDebug, vbChecked, vbUnchecked) 'Added by Morgan 2017/9/5
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
   End If
   Set frm880013 = Nothing
End Sub

'Modify by Morgan 2009/1/16 紙張定義改抓資料庫
Private Sub SetPaperSize()
   Dim PrnOld As String, PrnNew As String, iNo As Integer
   Dim arrNo(1 To 11, 2) As String, i As Integer
   Dim arrPaperName(1 To 11) As String
   Dim arrPaper() As String, lngWidth As Long, lngHeight As Long
   Dim iRec As Integer, iRecFound As Integer
   Dim bFound As Boolean
   
   strExc(0) = "select * from papersizedefine order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ReDim arrPaper(RsTemp.RecordCount, 3)
      iRecFound = 0
      PrnOld = Printer.DeviceName
      PrnNew = cmbPrinter.Text
      PUB_RestorePrinter PrnNew
      List1.Clear
      List1.AddItem "開始設定紙張格式..."
      DoEvents
      
On Error GoTo ErrHnd
   
      'Modify by Morgan 2010/1/18 改先抓自訂格式以便所有印表機都能共用
      'For iNo = 1 To 300
      'Modified by Morgan 2014/12/11 代碼改抓到 400(原300) 誤差 +-18twips(原50)
      For iNo = 400 To 1 Step -1
         Printer.PaperSize = iNo
         '判斷紙張代碼是否有設定成功
         If Printer.PaperSize = iNo Then
            With RsTemp
            .MoveFirst
            Do While Not .EOF
               iRec = .AbsolutePosition
               '格式未設定
               If arrPaper(iRec, 0) = "" Then
                  lngWidth = Val("" & .Fields("pd03")) * 567
                  lngHeight = Val("" & .Fields("pd04")) * 567
                  '誤差值 18twips
                  'Modified by Morgan 2015/8/3 寬度檢查大於就好
                  If Printer.Width >= lngWidth - 18 And (Printer.Height >= lngHeight - 18 And Printer.Height <= lngHeight + 18) Then
                     arrPaper(iRec, 0) = "" & .Fields("pd01") '格式代碼
                     arrPaper(iRec, 1) = Printer.PaperSize '紙張代碼
                     arrPaper(iRec, 2) = Round(Printer.Width / 567, 2) '紙張寬度
                     arrPaper(iRec, 3) = Round(Printer.Height / 567, 2) '紙張高度
                     strExc(0) = .Fields("pd01") & ".【" & .Fields("pd02") & "】"
                     List1.AddItem strExc(0) & String(30 - GetTextLength(strExc(0)), " ") & "紙張代碼 --> " & arrPaper(iRec, 1), 0
                     DoEvents
                     iRecFound = iRecFound + 1
                     Exit Do
                  End If
               End If
               .MoveNext
            Loop
            End With
         End If
         '全部都設定好就結束
         If iRecFound = UBound(arrPaper, 1) Then
            Exit For
         End If
      Next
      
      PUB_RestorePrinter PrnOld
      List1.AddItem "存檔中...", 0
      DoEvents
      If SaveData(arrPaper) = True Then
         List1.AddItem "存檔成功!", 0
      Else
         List1.AddItem "存檔失敗!", 0
      End If
   Else
      MsgBox "無法讀取紙張定義檔!!"
   End If
   Exit Sub
   
ErrHnd:
   If Err.Number = 380 Then
      Resume Next
   Else
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Function SaveData(arrNum() As String) As Boolean
   Dim idx As Integer, stSQL As String
   adoTaie.BeginTrans
   
On Error GoTo ErrHnd
   stSQL = "delete from papersizemap where pm01='" & pub_HostName & "'"
   adoTaie.Execute stSQL, intI
   For idx = LBound(arrNum) To UBound(arrNum)
      If arrNum(idx, 0) <> "" Then
         stSQL = "insert into papersizemap(PM01,PM02,PM03,PM04,PM05,PM06)" & _
            " values('" & ChgSQL(pub_HostName) & "','" & arrNum(idx, 0) & "'," & arrNum(idx, 1) & ",'" & ChgSQL(cmbPrinter.Text) & "','" & arrNum(idx, 2) & "'," & arrNum(idx, 3) & ")"
         adoTaie.Execute stSQL, intI
      End If
   Next
   adoTaie.CommitTrans
   SaveData = True
   Exit Function
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

