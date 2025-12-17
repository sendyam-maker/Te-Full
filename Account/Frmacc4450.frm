VERSION 5.00
Begin VB.Form Frmacc4450 
   AutoRedraw      =   -1  'True
   Caption         =   "科目分類帳"
   ClientHeight    =   4284
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4284
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdExcel 
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
      Left            =   480
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   1308
      Width           =   4215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1320
      TabIndex        =   0
      Top             =   96
      Width           =   3500
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   960
      Style           =   2  '單純下拉式
      TabIndex        =   20
      Top             =   2316
      Width           =   4050
   End
   Begin VB.TextBox Text3 
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
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   612
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
      Left            =   480
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1872
      Width           =   4215
   End
   Begin VB.ComboBox Combo13 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   6
      Top             =   3420
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   7
      Top             =   3420
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   8
      Top             =   3780
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   9
      Top             =   3780
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   4
      Top             =   840
      Width           =   1095
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
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   3
      Top             =   840
      Width           =   1095
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
      Left            =   3240
      TabIndex        =   2
      Top             =   480
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
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   120
      TabIndex        =   21
      Top             =   2340
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "是否列印傳票號碼"
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
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   2016
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      TabIndex        =   18
      Top             =   3060
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "1."
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
      Left            =   720
      TabIndex        =   17
      Top             =   3420
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image2 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc4450.frx":0000
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "2."
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
      Left            =   720
      TabIndex        =   16
      Top             =   3780
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image3 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc4450.frx":0442
      Stretch         =   -1  'True
      Top             =   3780
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1332
      Left            =   240
      Top             =   2940
      Visible         =   0   'False
      Width           =   4692
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
      Left            =   2520
      TabIndex        =   15
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "日期"
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
      TabIndex        =   14
      Top             =   840
      Width           =   612
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
      Left            =   3000
      TabIndex        =   13
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   12
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Left            =   360
      TabIndex        =   11
      Top             =   96
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc4450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc021 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt406 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Dim strSort1 As String
'edit by nickc 2007/02/08
'Dim trSort2 As String
Dim strSort2 As String, intCounter As Integer, intRow As Integer, intLength As Integer, strAmount As String
Dim douLastMonth As Double, douBalance As Double, douDebit As Double, douCredit As Double
'Dim dllaccrpt406 As Object
Private Const lngLeft As Long = 1400
'Add By Cheng 2003/06/10
'Modified by Lydia 2015/05/04
'Const m_dblPLeft As Double = 500
'Const m_dblPTop As Double = 540
Const m_dblPLeft As Double = 300
Const m_dblPTop As Double = 400
'預設印表機
Dim m_DefaultPrinter As String
Dim strPrinter As String 'Add By Sindy 2013/6/4
'Modified by Lydia 2015/05/04
Dim startY As Integer
Dim PLeft(0 To 6) As Integer
'Add by Amy 2024/06/05
Dim bolData As Boolean, strCmp As String, strReportN As String, strYear As String, strMonth As String  'strCmp從ProduceData搬出來
Dim strF() As String, arrWidth, intField As Integer, intTitleR As Integer, intPreR As Integer, bolShowPre As Boolean

'Added by Lydia 2015/05/04 設欄位Ｘ軸座標
Private Sub GetPleft()
    PLeft(0) = 0 '傳票日期
    PLeft(1) = PLeft(0) + Printer.TextWidth(String(5, "　"))  '傳票號碼，原1248
    PLeft(2) = PLeft(1) + Printer.TextWidth(String(6, "　"))  '摘要　　，原2676
    PLeft(3) = PLeft(2) + Printer.TextWidth(String(18, "　"))  '借方金額，原6891
    PLeft(4) = PLeft(3) + Printer.TextWidth(String(5, "　")) + 100  '貸方金額，原8391
    PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + 100 '餘額　　，原9891
    PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + 100
End Sub

'Add by Amy 2024/06/05
Private Sub CmdExcel_Click()
   Dim hLocalFile As Long
   
   If FormCheck = False Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   strYear = "": strMonth = ""
   Accrpt406Delete
   ProduceData
   If bolData = True Then
      strReportN = "科目分類帳"
      If SaveExcel = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      Else
         If MsgBox("EXCEL檔案已產生！" & vbCrLf & vbCrLf & strExcelPath & vbCrLf & vbCrLf & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
            ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
   FormClear
End Sub

Private Sub Combo13_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo4.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'Add by Amy 2020/04/08
Private Sub Combo2_GotFocus()
    TextInverse Combo2
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(Combo2) = MsgText(601) Then Exit Sub
    
    strCmp = Combo2
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo2.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo2)) = 1 Then
        Combo2 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/08

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo4_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(Combo2) = MsgText(601) Then Exit Sub
    
    strCmp = Combo2
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo2.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo2)) = 1 Then
        Combo2 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         'Modify by Amy 2020/04/08 公司別改下拉
         'Text6.SetFocus
         Combo2.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(195), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt406Delete
   ProduceData
   PUB_RestorePrinter Combo1 'Add By Sindy 2013/6/4
   PrintTitle 'Added by Lydia 2016/03/28 紙張,字型設定
   GetPleft 'Added by Lydia 2015/05/04
   PrintData
   PUB_RestorePrinter strPrinter 'Add By Sindy 2013/6/4
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
'Added by Lydia 2015/07/03
Dim iPrt As Integer, inX As Integer
Dim tmpPrt As String
Dim tmpArr As Variant

   '表單初始化
   'Modify by Amy 2023/07/19 修改寬高
   'PUB_InitForm Me, 5250, 3180 '2640
   'Modify by Amy 2024/06/05 原3300
   PUB_InitForm Me, 5400, 2290
   'Add by Amy 2020/04/08
   Combo2.AddItem "", 0
   Call Pub_SetCboCmp(Combo2, False, False, False, , 1)
   'end 2020/04/08
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   Text3 = MsgText(602)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'   Set dllaccrpt406 = CreateObject("AccReport.ReportSelect")
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
   
   'Added by Lydia 2015/07/03 剔除財務室的點陣印表機
   For iPrt = 0 To Combo1.ListCount - 1
       If InStr(Combo1.List(iPrt), "5577") = 0 Then
          tmpPrt = tmpPrt & "," & Trim(Combo1.List(iPrt))
       End If
   Next iPrt
   tmpPrt = Mid(tmpPrt, 2, Len(tmpPrt) - 1)
   Combo1.Clear
   tmpArr = Split(tmpPrt, ",")
   For iPrt = 0 To UBound(tmpArr)
      Combo1.AddItem tmpArr(iPrt)
      If Combo1.Tag = tmpArr(iPrt) Then inX = iPrt
   Next iPrt
   If inX > 0 Then
      Combo1.ListIndex = inX
   Else
      Combo1.ListIndex = 0
   End If
   'end 2015/07/3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Add By Sindy 2013/6/4
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2013/6/4 END

'   Set dllaccrpt406 = Nothing
   Set Frmacc4450 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

'Mark by Amy 2020/04/08 公司別改下拉
'Private Sub Text6_Change()
'   If Text6 = MsgText(601) Then
'      Exit Sub
'   End If
'
'   '20140123START Add By eric
'   If Text6 <> "1" And Text6 <> "J" Then
'      MsgBox "公司別僅能為 1 或 J ! (1:台一/J:智權)"
'      Text6.Text = ""
'      Text6.SetFocus
'   End If
'   '20140123END
'
'   Text7 = A0802Query(Text6)
'End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()

   strSort1 = "日期"
   strSort2 = "傳票號碼"
   Combo13.AddItem strSort1
   Combo13.AddItem strSort2
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
End Sub

'Add By Sindy 2009/06/01
'抓上期年月
Private Sub SetLastMonth(ByRef strYear As String, strMonth As String)
   If Text4 <> MsgText(601) Then
'      If Len(Text4) = 4 Then
'         If Val(Mid(Text4, 3, 2)) = 1 Then
'            strYear = Val(Mid(Text4, 1, 2)) - 1
'            strMonth = "12"
'         Else
'            strYear = Mid(Text4, 1, 2)
'            strMonth = Val(Mid(Text4, 3, 2)) - 1
'         End If
'      Else
'         If Val(Mid(Text4, 4, 2)) = 1 Then
'            strYear = Val(Mid(Text4, 1, 3)) - 1
'            strMonth = "12"
'         Else
'            strYear = Mid(Text4, 1, 3)
'            strMonth = Val(Mid(Text4, 4, 2)) - 1
'         End If
'      End If
      'Modify By Sindy 2013/6/4
      If Len(Text4) = 6 Then
         If Val(Mid(Text4, 3, 2)) = 1 Then
            strYear = Val(Mid(Text4, 1, 2)) - 1
            strMonth = "12"
         Else
            strYear = Mid(Text4, 1, 2)
            strMonth = Val(Mid(Text4, 3, 2)) - 1
         End If
      Else
         If Val(Mid(Text4, 4, 2)) = 1 Then
            strYear = Val(Mid(Text4, 1, 3)) - 1
            strMonth = "12"
         Else
            strYear = Mid(Text4, 1, 3)
            strMonth = Val(Mid(Text4, 4, 2)) - 1
         End If
      End If
      '2013/6/4 End
   End If
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1 As String
Dim strOrder2 As String
Dim strSql As String
Dim lngStartDate As Long
Dim lngEndDate As Long
Dim lngCounter As Long
'Add By Sindy 2009/06/01
Dim strYear As String
Dim strMonth As String
'2009/06/01 End

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   strCmp = "": bolData = False 'Modify by Amy 2024/06/05
   Select Case Combo13
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = ", a0205 asc"
         Else
            strOrder1 = ", a0205 desc"
         End If
         Select Case Combo5
            Case strSort2
               If Combo6 = MsgText(1) Then
                  strOrder2 = ", ax202 asc"
               Else
                  strOrder2 = ", ax202 desc"
               End If
            Case Else
               strOrder2 = MsgText(601)
         End Select
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = ", ax202 asc"
         Else
            strOrder1 = ", ax202 desc"
         End If
         Select Case Combo5
            Case strSort1
               If Combo6 = MsgText(1) Then
                  strOrder2 = ", a0205 asc"
               Else
                  strOrder2 = ", a0205 desc"
               End If
            Case Else
               strOrder2 = MsgText(601)
         End Select
      Case Else
         strOrder1 = MsgText(601)
         strOrder2 = MsgText(601)
   End Select
   lngCounter = 1
   If adoaccrpt406.State <> adStateClosed Then adoaccrpt406.Close 'Add by Amy 2024/06/07
   adoaccrpt406.CursorLocation = adUseClient
   adoaccrpt406.Open "select * from accrpt406", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   If adoacc021.State <> adStateClosed Then adoacc021.Close 'Add by Amy 2024/06/07
   adoacc021.CursorLocation = adUseClient
   If Text4 <> MsgText(601) Then
      lngStartDate = Text4 'Val(Text4 & MsgText(12)) Modify By Sindy 2013/6/4
   Else
      lngStartDate = 0
   End If
   If Text5 <> MsgText(601) Then
      lngEndDate = Text5 'Val(Text5 & MsgText(13)) Modify By Sindy 2013/6/4
   Else
      lngEndDate = 0
   End If
   'Remove by Lydia 2015/05/04 移除不用的格式
'   Select Case strAccount
'      Case "2"
'         If Text2 <> MsgText(601) Then
'            strSql = " and ax305 >= '" & Text2 & "'"
'         End If
'         If Text1 <> MsgText(601) Then
'            strSql = strSql & " and ax305 <= '" & Text1 & "'"
'         End If
'         If Text6 <> MsgText(601) Then
'            strSql = strSql & " and ax301 = '" & Text6 & "'"
'         End If
'         If lngStartDate <> 0 Then
'            strSql = strSql & " and a0305 >= " & lngStartDate & ""
'         End If
'         If lngEndDate <> 0 Then
'            strSql = strSql & " and a0305 <= " & lngEndDate & ""
'         End If
'      Case Else
         If Text2 <> MsgText(601) Then
            'Modify By Sindy 2009/06/01
            'strSQL = " and ax205 >= '" & Text2 & "'"
         End If
         If Text1 <> MsgText(601) Then
            'Modify By Sindy 2009/06/01
            'strSQL = strSQL & " and ax205 <= '" & Text1 & "'"
         End If
         'Modify by Amy 2020/04/08 公司別改下拉 原:Text6
         strCmp = Combo2
         If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
         End If
         If strCmp <> MsgText(601) Then
            strSql = strSql & " and ax201 = '" & strCmp & "'"
         End If
         'end 2020/04/08
         If lngStartDate <> 0 Then
            strSql = strSql & " and a0205 >= " & lngStartDate & ""
         End If
         If lngEndDate <> 0 Then
            strSql = strSql & " and a0205 <= " & lngEndDate & ""
         End If
'   End Select
'2015/05/04

'   If strSQL <> MsgText(601) Then
'      strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 1)
'   End If
   
   SetLastMonth strYear, strMonth 'Add by Sindy 2009/06/01
   
   'Remove by Lydia 2015/05/04 移除不用的格式
'   Select Case strAccount
'      Case "2"
'         adoacc021.Open "select ax305 as ax205, a0305 as a0205, ax302 as ax202, ax308 as ax208, ax309 as ax209, ax314 as ax214, ax312 as ax212, ax306 as ax206, ax307 as ax207 from (select * from acc031, acc030, acc010 where acc031.ax301 = acc030.a0301 and acc031.ax302 = acc030.a0302 and ax305 = a0101) new " & strSql & " order by ax305 asc, ax302 asc, ax303 asc" & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
'      Case Else
         'Modify By Sindy 2009/06/01
         'adoacc021.Open "select * from (select * from acc021, acc020, acc010 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax205 = a0101) new " & strSQL & " order by ax205 asc, ax202 asc, ax203 asc" & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
         'Moidfy by Amy 2020/04/08 原:Text6
         adoacc021.Open "select * from (select * from acc021, acc020 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 " & strSql & " ) a, " & _
         "(select * from acc010 where a0101 >= '" & Text2 & "' and a0101 <= '" & Text1 & "') b, " & _
         "(select * from acc040 where a0401 = " & Val(strYear) & " and a0402 = " & Val(strMonth) & " and a0403 = '" & strCmp & "' and a0404 = '" & MsgText(55) & "') c " & _
         "where a0101=ax205(+) and a0101=a0405(+) " & _
         "and ((substr(a0101,1,1) in ('1','2','3') and (ax206 <> 0 or ax207 <> 0 or a0408 <> 0)) or (substr(a0101,1,1) not in ('1','2','3') and (ax206 <> 0 or ax207 <> 0))) " & _
         "order by ax205 asc, ax202 asc, ax203 asc " & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
'   End Select

   If adoacc021.RecordCount = 0 Then
      adoacc021.Close
      adoaccrpt406.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   bolData = True 'Add by Amy 2024/06/05 有資料
   Do While adoacc021.EOF = False
      adoaccrpt406.AddNew
      adoaccrpt406.Fields("r40601").Value = strUserNum
      'Modify By Sindy 2009/06/01
      'adoaccrpt406.Fields("r40611").Value = adoacc021.Fields("ax205").Value
      adoaccrpt406.Fields("r40611").Value = adoacc021.Fields("a0101").Value
      If IsNull(adoacc021.Fields("a0205").Value) Then
         adoaccrpt406.Fields("r40602").Value = Null
      Else
         adoaccrpt406.Fields("r40602").Value = adoacc021.Fields("a0205").Value
      End If
      If IsNull(adoacc021.Fields("ax202").Value) Then
         adoaccrpt406.Fields("r40603").Value = Null
      Else
         adoaccrpt406.Fields("r40603").Value = adoacc021.Fields("ax202").Value
      End If
      If IsNull(adoacc021.Fields("ax208").Value) Then
         adoaccrpt406.Fields("r40604").Value = Null
      Else
         adoaccrpt406.Fields("r40604").Value = adoacc021.Fields("ax208").Value
      End If
      If IsNull(adoacc021.Fields("ax209").Value) Then
         adoaccrpt406.Fields("r40605").Value = Null
      Else
         adoaccrpt406.Fields("r40605").Value = adoacc021.Fields("ax209").Value
      End If
      If IsNull(adoacc021.Fields("ax214").Value) Then
         adoaccrpt406.Fields("r40606").Value = Null
      Else
         adoaccrpt406.Fields("r40606").Value = adoacc021.Fields("ax214").Value
      End If
      If adoacc021.Fields("ax202").Value = "D110011995" Then
        strExc(0) = ""
      End If
      If IsNull(adoacc021.Fields("ax212").Value) Then
         adoaccrpt406.Fields("r40607").Value = Null
      Else
        'Modify by Amy 2021/02/01 摘要過長會錯
         adoaccrpt406.Fields("r40607").Value = StrToStr(adoacc021.Fields("ax212").Value, 75)
      End If
      If IsNull(adoacc021.Fields("ax206").Value) Then
         adoaccrpt406.Fields("r40608").Value = 0
      Else
         adoaccrpt406.Fields("r40608").Value = Val(adoacc021.Fields("ax206").Value)
      End If
      If IsNull(adoacc021.Fields("ax207").Value) Then
         adoaccrpt406.Fields("r40609").Value = 0
      Else
         adoaccrpt406.Fields("r40609").Value = Val(adoacc021.Fields("ax207").Value)
      End If
      'Add by Amy 2024/06/05 +對沖其他
      If IsNull(adoacc021.Fields("ax213").Value) Then
         adoaccrpt406.Fields("r40613").Value = Null
      Else
         adoaccrpt406.Fields("r40613").Value = adoacc021.Fields("ax213").Value
      End If
      
      If adoacc010.State <> adStateClosed Then adoacc010.Close 'Add by Amy 2024/06/07
      adoacc010.CursorLocation = adUseClient
      'Modify By Sindy 2009/06/01
      'adoacc010.Open "select a0103 from acc010 where a0101 = '" & adoacc021.Fields("ax205").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoacc010.Open "select a0103 from acc010 where a0101 = '" & adoacc021.Fields("a0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc010.Fields(0).Value = "1" Then
         adoaccrpt406.Fields("r40610").Value = Val(adoaccrpt406.Fields("r40608").Value) - Val(adoaccrpt406.Fields("r40609").Value)
      Else
         adoaccrpt406.Fields("r40610").Value = Val(adoaccrpt406.Fields("r40609").Value) - Val(adoaccrpt406.Fields("r40608").Value)
      End If
      adoacc010.Close
      adoaccrpt406.Fields("r40612").Value = lngCounter
      adoaccrpt406.UpdateBatch
      lngCounter = lngCounter + 1
      adoacc021.MoveNext
   Loop
   adoacc021.Close
   adoaccrpt406.Close
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
Private Sub Accrpt406Delete()
   'Modify by Amy 2024/06/05 避免財務同時使用,造成資料有問題 +Where
   adoTaie.Execute "delete from accrpt406 Where R40601='" & strUserNum & "' "
End Sub

''*************************************************
''  執行報表之 Dll
''
''*************************************************
'Private Sub RunReportDll()
'   dllaccrpt406.Acc4450 ReportTitle(406), Text6, Text7, Text4, Text5, adoacc010.Fields("a0101").Value, adoacc010.Fields("a0102").Value, strUserNum, CFDate(ACDate(ServerDate))
'End Sub

'*************************************************
'  列印報表
'
'*************************************************
Private Sub PrintData()
Dim strYear As String
Dim strMonth As String
Dim strCmp As String 'Add by Amy 2020/04/08
   
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select a0b02 from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields("a0b02").Value) = False Then
'         strYear = Mid(CFDate(adoaccsum.Fields("a0b02").Value), 1, 3)
'         strMonth = Mid(CFDate(adoaccsum.Fields("a0b02").Value), 5, 2)
'      End If
'   End If
'   adoaccsum.Close
   
   'Modify By Sindy 2009/06/01
'   If Text4 <> MsgText(601) Then
'      If Len(Text4) = 4 Then
'         If Val(Mid(Text4, 3, 2)) = 1 Then
'            strYear = Val(Mid(Text4, 1, 2)) - 1
'            strMonth = "12"
'         Else
'            strYear = Mid(Text4, 1, 2)
'            strMonth = Val(Mid(Text4, 3, 2)) - 1
'         End If
'      Else
'         If Val(Mid(Text4, 4, 2)) = 1 Then
'            strYear = Val(Mid(Text4, 1, 3)) - 1
'            strMonth = "12"
'         Else
'            strYear = Mid(Text4, 1, 3)
'            strMonth = Val(Mid(Text4, 4, 2)) - 1
'         End If
'      End If
'   End If
   SetLastMonth strYear, strMonth
   
   PrintTitle '紙張設定
  intCounter = 1
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select * from acc010 where a0101 >= '" & Text2 & "' and a0101 <= '" & Text1 & "' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc010.EOF = False
      intRow = 0
      douBalance = 0
      douDebit = 0
      douCredit = 0
      adoaccsum.CursorLocation = adUseClient
      '20140123START Modify By eric
      'Modify by Amy 2020/04/08 原:Text6
      strCmp = Combo2
      If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      adoaccsum.Open "select a0408 from acc040 where a0401 = " & Val(strYear) & " and a0402 = " & Val(strMonth) & " and a0403 ='" & strCmp & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc010.Fields("a0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      'adoaccsum.Open "select a0408 from acc040 where a0401 = " & Val(strYear) & " and a0402 = " & Val(strMonth) & " and a0403 = '1' and a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc010.Fields("a0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      '20140123END
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields("a0408").Value) = False Then
            douLastMonth = adoaccsum.Fields("a0408").Value
         Else
            douLastMonth = 0
         End If
      Else
         douLastMonth = 0
      End If
      adoaccsum.Close
      Select Case Mid(adoacc010.Fields("a0101").Value, 1, 1)
         Case "1", "2", "3"
            douBalance = douLastMonth
         Case Else
            douBalance = 0
      End Select
      '93.3.5 add by sonia
      If adoacc010.Fields("a0103").Value = "2" And Mid(adoacc010.Fields("a0101").Value, 1, 1) And douBalance < 0 Then
         douBalance = douBalance * -1
      End If
      '93.3.5 emd
      adoaccrpt406.CursorLocation = adUseClient
      'Modify by Amy 2024/06/05 +strUserNum
      adoaccrpt406.Open "select * from accrpt406 where r40611 = '" & adoacc010.Fields("a0101").Value & "' And R40601='" & strUserNum & "' order by r40602 asc, r40603 asc, r40612 asc" _
            , adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccrpt406.RecordCount <> 0 Then
         PrintHead
         Select Case Mid(adoacc010.Fields("a0101").Value, 1, 1)
            Case "1", "2", "3"
            '抓上期餘額
            'Remove by Lydia 2015/05/04 移除不用的格式
'               Select Case strAccount
'                  Case "2"
'                     Printer.CurrentX = 2676 - lngLeft + m_dblPLeft
'                     Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'                     Printer.Print MsgText(204)
'                     strAmount = Format(douBalance, FDollar)
'                     intLength = Printer.TextWidth(strAmount)
'                     Printer.CurrentX = 11091 - intLength - lngLeft + m_dblPLeft
'                     Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'                     Printer.Print strAmount
'                     intRow = intRow + 1
'                  Case Else
                     'Modified by Lydia 2016/03/28 調整位置
                     'Printer.CurrentX = 2676 + m_dblPLeft
                     Printer.CurrentX = PLeft(2) + m_dblPLeft
                     Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
                     Printer.Print MsgText(204)
                     strAmount = Format(douBalance, FDollar)
                     intLength = Printer.TextWidth(strAmount)
                     'Modified by Lydia 2015/05/04 改列印位置(A4) , 靠右
                     'Printer.CurrentX = 11091 - intLength + m_dblPLeft
                     Printer.CurrentX = PLeft(6) - intLength + m_dblPLeft
                     Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
                     Printer.Print strAmount
                     intRow = intRow + 1
'               End Select

         End Select
      End If
      Do While adoaccrpt406.EOF = False
         douDebit = douDebit + Val(adoaccrpt406.Fields("r40608").Value)
         douCredit = douCredit + Val(adoaccrpt406.Fields("r40609").Value)
         PrintDetail
         intRow = intRow + 1
         RowCheck
         adoaccrpt406.MoveNext
         If adoaccrpt406.EOF Then
            PrintSum
         End If
      Loop
'      If adoaccrpt406.RecordCount <> 0 Then
'         RunReportDll
'         Sleep intSleep
'      End If
      If adoaccrpt406.RecordCount <> 0 Then
         intCounter = intCounter + 1
         Printer.EndDoc
         PrintTitle
      End If
      adoaccrpt406.Close
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   Printer.EndDoc
End Sub

'*************************************************
'  列印表頭
'
'*************************************************
Private Sub PrintHead()
   'Add by Amy 2020/04/08
   Dim strCmp As String
   strCmp = Combo2
   If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
   End If
   'end 2020/04/08
   Printer.FontSize = 14
   Printer.CurrentX = 4000 + m_dblPLeft
   Printer.CurrentY = 60 + m_dblPTop
   Printer.Print ReportTitle(406)
   Printer.FontSize = 10
   Printer.CurrentX = 3100 + m_dblPLeft
   Printer.CurrentY = 600 + m_dblPTop
   Printer.Print "公司別:"
   Printer.CurrentX = 4300 + m_dblPLeft
   Printer.CurrentY = 600 + m_dblPTop
   Printer.Print strCmp 'Moidfy by Amy 2020/04/08 原:Text6
   Printer.CurrentX = 4630 + m_dblPLeft
   Printer.CurrentY = 600 + m_dblPTop
   Printer.Print A0802Query(strCmp) 'Moidfy by Amy 2020/04/08 原:Text7
   Printer.CurrentX = 3100 + m_dblPLeft
   Printer.CurrentY = 900 + m_dblPTop
   Printer.Print "會計科目:"
   Printer.CurrentX = 4300 + m_dblPLeft
   Printer.CurrentY = 900 + m_dblPTop
   Printer.Print adoacc010.Fields("a0101").Value
   Printer.CurrentX = 5250 + m_dblPLeft
   Printer.CurrentY = 900 + m_dblPTop
   Printer.Print IIf(IsNull(adoacc010.Fields("a0102").Value), MsgText(601), adoacc010.Fields("a0102").Value)
   Printer.CurrentX = 0 + m_dblPLeft
   Printer.CurrentY = 1300 + m_dblPTop
   Printer.Print "列印人員:"
   Printer.CurrentX = 1185 + m_dblPLeft
   Printer.CurrentY = 1300 + m_dblPTop
   Printer.Print StaffQuery(strUserNum)
   Printer.CurrentX = 8172 + m_dblPLeft + 400
   Printer.CurrentY = 1300 + m_dblPTop
   Printer.Print "列印日期:"
   'Modified by Lydia 2015/05/04 改列印位置(A4)
   'Printer.CurrentX = 9372 + m_dblPLeft
   Printer.CurrentX = 9420 + m_dblPLeft
   Printer.CurrentY = 1300 + m_dblPTop
   Printer.Print CFDate(ACDate(ServerDate))
   Printer.CurrentX = 8172 + m_dblPLeft + 400
   Printer.CurrentY = 1600 + m_dblPTop
   Printer.Print "頁次:"
   'Modified by Lydia 2015/05/04 改列印位置(A4)
   'Printer.CurrentX = 9372 + m_dblPLeft
   Printer.CurrentX = 9420 + m_dblPLeft
   Printer.CurrentY = 1600 + m_dblPTop
   Printer.Print intCounter
   'Modified by Lydia 2015/05/04 改列印位置(A4)
   'Printer.CurrentX = 0 + m_dblPLeft
   Printer.CurrentX = PLeft(0) + m_dblPLeft
   Printer.CurrentY = 2000 + m_dblPTop
   Printer.Print "傳票日期"
   'Remove by Lydia 2015/05/04 移除不用的格式
'   Select Case strAccount
'      Case "2"
'         If Text3 = MsgText(602) Then
'            Printer.CurrentX = 10000 + m_dblPLeft
'            Printer.CurrentY = 2000 + m_dblPTop
'            Printer.Print "傳票號碼"
'         End If
'         Printer.CurrentX = 2676 - lngLeft + m_dblPLeft
'         Printer.CurrentY = 2000 + m_dblPTop
'         Printer.Print "摘要"
'         Printer.CurrentX = 6891 - lngLeft + m_dblPLeft + 400
'         Printer.CurrentY = 2000 + m_dblPTop
'         Printer.Print "借方金額"
'         Printer.CurrentX = 8391 - lngLeft + m_dblPLeft + 400
'         Printer.CurrentY = 2000 + m_dblPTop
'         Printer.Print "貸方金額"
'         Printer.CurrentX = 10591 - lngLeft + m_dblPLeft + 400
'         Printer.CurrentY = 2000 + m_dblPTop
'         Printer.Print "餘額"
'         Printer.CurrentX = 0 + m_dblPLeft
'         Printer.CurrentY = 2300 + m_dblPTop
'         Printer.Print String(50, "─")
'      Case Else
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 1248 + m_dblPLeft
         Printer.CurrentX = PLeft(1) + m_dblPLeft
         Printer.CurrentY = 2000 + m_dblPTop
         Printer.Print "傳票號碼"
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 2676 + m_dblPLeft
         Printer.CurrentX = PLeft(2) + m_dblPLeft
         Printer.CurrentY = 2000 + m_dblPTop
         Printer.Print "摘要"
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 6891 + m_dblPLeft + 400
         Printer.CurrentX = PLeft(4) + m_dblPLeft - Printer.TextWidth(String(4, "　"))
         Printer.CurrentY = 2000 + m_dblPTop
         Printer.Print "借方金額"
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 8391 + m_dblPLeft + 400
         Printer.CurrentX = PLeft(5) + m_dblPLeft - Printer.TextWidth(String(4, "　"))
         Printer.CurrentY = 2000 + m_dblPTop
         Printer.Print "貸方金額"
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 9891 + m_dblPLeft + 400
         Printer.CurrentX = PLeft(6) + m_dblPLeft - Printer.TextWidth(String(2, "　"))
         Printer.CurrentY = 2000 + m_dblPTop
         Printer.Print "餘額"
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 0 + m_dblPLeft
         Printer.CurrentX = PLeft(0) + m_dblPLeft
         Printer.CurrentY = 2300 + m_dblPTop
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.Print String(57, "─")
         Printer.Print String(55, "─")
'   End Select
End Sub

'*************************************************
'  列印報表明細
'
'*************************************************
Private Sub PrintDetail()
   '傳票日期
   'Modified by Lydia 2015/05/04 改列印位置(A4)
   'Printer.CurrentX = 0 + m_dblPLeft
   Printer.CurrentX = PLeft(0) + m_dblPLeft
   Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
   Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40602").Value), MsgText(601), Format(adoaccrpt406.Fields("r40602").Value, DFormat))
   'Remove by Lydia 2015/05/04 移除不用的格式
'   Select Case strAccount
'      Case "2"
'         If Text3 = MsgText(602) Then
'            Printer.CurrentX = 10000 + m_dblPLeft
'            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'            Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40603").Value), MsgText(601), adoaccrpt406.Fields("r40603").Value)
'         End If
'         Printer.CurrentX = 2676 - lngLeft + m_dblPLeft
'         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'        'Modify By Cheng 2003/06/10
''         Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40607").Value), MsgText(601), StrConv(MidB(StrConv(adoaccrpt406.Fields("r40607").Value, vbFromUnicode), 1, 46), vbUnicode))
'         Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40607").Value), MsgText(601), StrConv(MidB(StrConv(adoaccrpt406.Fields("r40607").Value, vbFromUnicode), 1, 42), vbUnicode))
'         strAmount = IIf(IsNull(adoaccrpt406.Fields("r40608").Value), 0, Format(adoaccrpt406.Fields("r40608").Value, FDollar))
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 8091 - intLength - lngLeft + m_dblPLeft + 400
'         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'         Printer.Print strAmount
'         strAmount = IIf(IsNull(adoaccrpt406.Fields("r40609").Value), 0, Format(adoaccrpt406.Fields("r40609").Value, FDollar))
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 9591 - intLength - lngLeft + m_dblPLeft + 400
'         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'         Printer.Print strAmount
'         If adoacc010.Fields("a0103").Value = "1" Then
'            douBalance = douBalance + Val(adoaccrpt406.Fields("r40608").Value) - Val(adoaccrpt406.Fields("r40609").Value)
'         Else
'            douBalance = douBalance + (Val(adoaccrpt406.Fields("r40609").Value) - Val(adoaccrpt406.Fields("r40608").Value))
'         End If
'         strAmount = Format(douBalance, FDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 11091 - intLength - lngLeft + m_dblPLeft + 400
'         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'         Printer.Print strAmount
'      Case Else
         '傳票號碼
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 1248 + m_dblPLeft
         Printer.CurrentX = PLeft(1) + m_dblPLeft
         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
         Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40603").Value), MsgText(601), adoaccrpt406.Fields("r40603").Value)
         '摘要
         'Modified by Lydia 2015/05/04 改列印位置(A4)
         'Printer.CurrentX = 2676 + m_dblPLeft
         Printer.CurrentX = PLeft(2) + m_dblPLeft
         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
        'Modify By Cheng 2003/06/10
'         Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40607").Value), MsgText(601), StrConv(MidB(StrConv(adoaccrpt406.Fields("r40607").Value, vbFromUnicode), 1, 46), vbUnicode))
         'Modified by Lydia 2015/05/04
         'Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40607").Value), MsgText(601), StrConv(MidB(StrConv(adoaccrpt406.Fields("r40607").Value, vbFromUnicode), 1, 42), vbUnicode))
         strExc(0) = "" & adoaccrpt406.Fields("r40607").Value ': strExc(0) = convForm(strExc(0), 20)
         Printer.Print IIf(IsNull(adoaccrpt406.Fields("r40607").Value), MsgText(601), StrConv(MidB(StrConv(adoaccrpt406.Fields("r40607").Value, vbFromUnicode), 1, 42), vbUnicode))
         '借方金額
         strAmount = IIf(IsNull(adoaccrpt406.Fields("r40608").Value), 0, Format(adoaccrpt406.Fields("r40608").Value, FDollar))
         intLength = Printer.TextWidth(strAmount)
         'Modified by Lydia 2015/05/04 改列印位置(A4), 靠右
         'Printer.CurrentX = 8091 - intLength + m_dblPLeft + 400
         Printer.CurrentX = PLeft(4) - intLength + m_dblPLeft
         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
         Printer.Print strAmount
         '貸方金額
         strAmount = IIf(IsNull(adoaccrpt406.Fields("r40609").Value), 0, Format(adoaccrpt406.Fields("r40609").Value, FDollar))
         intLength = Printer.TextWidth(strAmount)
         'Modified by Lydia 2015/05/04 改列印位置(A4), 靠右
         'Printer.CurrentX = 9591 - intLength + m_dblPLeft + 400
         Printer.CurrentX = PLeft(5) - intLength + m_dblPLeft
         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
         Printer.Print strAmount
         '餘額
         If adoacc010.Fields("a0103").Value = "1" Then
            douBalance = douBalance + Val(adoaccrpt406.Fields("r40608").Value) - Val(adoaccrpt406.Fields("r40609").Value)
         Else
            douBalance = douBalance + (Val(adoaccrpt406.Fields("r40609").Value) - Val(adoaccrpt406.Fields("r40608").Value))
         End If
         strAmount = Format(douBalance, FDollar)
         intLength = Printer.TextWidth(strAmount)
         'Modified by Lydia 2015/05/04 改列印位置(A4), 靠右
         'Printer.CurrentX = 11091 - intLength + m_dblPLeft + 400
         Printer.CurrentX = PLeft(6) - intLength + m_dblPLeft
         Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
         Printer.Print strAmount
'   End Select
End Sub

'*************************************************
'  紙張設定
'
'*************************************************
Private Sub PrintTitle()
'Modified by Lydia 2015/05/04 因為目前已改用A4紙張故取消紙張設定直接用所選印表機的預設紙張這樣才不會受限
   'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
'   '9x
'   If pub_OS = "1" Then
'      Printer.Height = 16000
'      Printer.Width = 19000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(7)
'   End If
'   'end 2008/3/25

   'Added by Lydia 2015/07/03 預設紙張A4
   Printer.PaperSize = 9
   Printer.Font = "新細明體" 'Added by Lydia 2016/03/28
   Printer.FontSize = intFontSize
   
End Sub

'*************************************************
'  列數設定
'
'*************************************************
Private Sub RowCheck()
  'Modified by Lydia 2015/05/04 改列印位置(A4)
   'If intRow > 40 Then
   If intRow > 42 Then
      intRow = 0
      intCounter = intCounter + 1
      Printer.NewPage
      PrintHead
   End If
End Sub

'*************************************************
'  合計列印
'
'*************************************************
Private Sub PrintSum()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(r40608), sum(r40609), sum(r40610) from accrpt406 where r40611 = '" & adoacc010.Fields("a0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
   'Remove by Lydia 2015/05/04 移除不用的格式
'      Select Case strAccount
'         Case "2"
'            Printer.CurrentX = 6091 - lngLeft + m_dblPLeft
'            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'            Printer.Print String(27, "─")
'            intRow = intRow + 1
'            RowCheck
'            Printer.CurrentX = 5391 - lngLeft + m_dblPLeft
'            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'            Printer.Print ReportSum(25)
'            strAmount = Format(douDebit, FDollar)
'            intLength = Printer.TextWidth(strAmount)
'            Printer.CurrentX = 8091 - intLength - lngLeft + m_dblPLeft + 400
'            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'            Printer.Print strAmount
'            strAmount = Format(douCredit, FDollar)
'            intLength = Printer.TextWidth(strAmount)
'            Printer.CurrentX = 9591 - intLength - lngLeft + m_dblPLeft + 400
'            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'            Printer.Print strAmount
'            strAmount = Format(douBalance, FDollar)
'            intLength = Printer.TextWidth(strAmount)
'            Printer.CurrentX = 11091 - intLength - lngLeft + m_dblPLeft + 400
'            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
'            Printer.Print strAmount
'         Case Else
            'Modified by Lydia 2015/05/04 改列印位置(A4)
            'Printer.CurrentX = 6091 + m_dblPLeft
            Printer.CurrentX = 5300 + m_dblPLeft
            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
            'Modified by Lydia 2015/05/04
            'Printer.Print String(27, "─")
            Printer.Print String(29, "─")
            intRow = intRow + 1
            RowCheck
            Printer.CurrentX = 5391 + m_dblPLeft
            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
            Printer.Print ReportSum(25)
            strAmount = Format(douDebit, FDollar)
            intLength = Printer.TextWidth(strAmount)
            'Modified by Lydia 2015/05/04 改列印位置(A4)
            'Printer.CurrentX = 8091 - intLength + m_dblPLeft + 400
            Printer.CurrentX = PLeft(4) - intLength + m_dblPLeft
            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
            Printer.Print strAmount
            strAmount = Format(douCredit, FDollar)
            intLength = Printer.TextWidth(strAmount)
            'Modified by Lydia 2015/05/04 改列印位置(A4)
            'Printer.CurrentX = 9591 - intLength + m_dblPLeft + 400
            Printer.CurrentX = PLeft(5) - intLength + m_dblPLeft
            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
            Printer.Print strAmount
            strAmount = Format(douBalance, FDollar)
            intLength = Printer.TextWidth(strAmount)
            'Modified by Lydia 2015/05/04 改列印位置(A4)
            'Printer.CurrentX = 11091 - intLength + m_dblPLeft + 400
            Printer.CurrentX = PLeft(6) - intLength + m_dblPLeft
            Printer.CurrentY = 2500 + intRow * 300 + m_dblPTop
            Printer.Print strAmount
'      End Select
   End If
   adoaccsum.Close
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy  2020/04/08
'   Text6 = ""
'   Text7 = ""
   Combo2 = ""
   'end 2020/04/08
   'edit by nickc 2007/02/08
   'Text8 = ""
   'Text9 = ""
   Text2 = ""
   Text1 = ""
   Text4 = ""
   Text5 = ""
   Text3 = MsgText(602)
   Combo13 = ""
   Combo5 = ""
   Combo2.SetFocus 'Moidfy by Amy 2020/04/08 原:Text6.SetFocus
End Sub

'Mark by Amy 2020/04/08 公司別改下拉
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'   '20140123START Add By eric
'   CloseIme
'   '20140123END
'End Sub

''20140123START Add By eric
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''20140123START By eric
'Private Sub Text6_LostFocus()
'   If Text6.Text = "" Then
'      MsgBox "公司別僅可為 1 或 J   ! (1:台一/J:智權)"
'      Text6.Text = ""
'      Text6.SetFocus
'      Exit Sub
'   End If
'End Sub
'end 2020/04/08
'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   Dim bCancel As Boolean 'Add by Amy 2020/04/08
      
   '2013/8/29 add by sonia 瑞婷說日期一定要輸,故從下面移上來
   FormCheck = False
   'Modify by Amy 2024/06/05 公司別一定要輸,避免餘額抓錯及傳票資料混在一起(原於日期後判斷)
   'Add by Amy 2020/04/08
   If Combo2 = MsgText(601) Then
      MsgBox "公司別必填", , MsgText(5)
      Exit Function
   Else
        Call Combo2_Validate(bCancel)
        If bCancel = True Then
            Exit Function
        End If
   End If
   'end 2020/04/08
   If Text4 = MsgText(601) Then
      MsgBox "日期欄位必填", , MsgText(5)
      Exit Function
   End If
   If Text5 = MsgText(601) Then
      MsgBox "日期欄位必填", , MsgText(5)
      Exit Function
   End If
   '2013/8/29 end
   
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
'2013/8/29 cancel by sonia
'   If Text4 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text5 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
   FormCheck = True
'2013/8/19 end
End Function

'Add by Amy 2024/06/05
Private Function SaveExcel() As Boolean
   Dim xlsAp As New Excel.Application, wksrpt As New Worksheet, i As Integer, intWkPage As Integer, strWkName As String, bolOpenXls As Boolean
   Dim strAllF As String, strAllW As String, strFileN As String, strOldAccNo As String, strOldA0103 As String, strFormat As String, strTmp(2) As String, strMsg(2) As String
   Dim strSql As String, intQ1 As Integer, intQ2 As Integer, strLastAmt As String, strBalance As String

On Error GoTo ErrHnd
   
   '依 目分類帳查詢的欄位 輸出-斯閔
   strAllF = "傳票日,傳票號碼,借方金額,貸方金額,餘額,摘要,對沖(客),對沖(業),對沖(案號),對沖(其他)"
   strAllW = "6.8, 9.5, 11, 11,11,12.5, 8.5, 5, 11.5, 5.5"
   strF = Split(strAllF, ",")
   arrWidth = Split(strAllW, ",")
   
   SetLastMonth strYear, strMonth
   
   '抓餘額資料
   intRow = 1: intField = 65: intWkPage = 1: bolShowPre = False
   'Memo by Amy 有餘額或傳票有資料就出現
   strSql = "Select R40611,a0102,a0103,a0408 From Acc010,Acc040 " & _
                  ",(Select Distinct R40611 From accrpt406 Where R40601='" & strUserNum & "' ) " & _
                  "Where a0405(+)=R40611 And a0404(+)='" & MsgText(55) & "'  And a0403(+)='" & strCmp & "' " & _
                  " And a0401(+)= " & Val(strYear) & " And a0402(+)= " & Val(strMonth) & _
                  " And a0405=a0101(+) Order by R40611 asc"
   intQ1 = 1
   Set adoacc010 = ClsLawReadRstMsg(intQ1, strSql)
   If intQ1 = 1 Then
      adoacc010.MoveFirst
      xlsAp.SheetsInNewWorkbook = 3
      xlsAp.Workbooks.add
      strWkName = Left(xlsAp.Worksheets(1).Name, Len(xlsAp.Worksheets(1).Name) - 1)
      bolOpenXls = True
      Do While adoacc010.EOF = False
         If strOldAccNo <> MsgText(601) And strOldAccNo <> "" & adoacc010.Fields("R40611").Value Then
            '*** 不同科目換工作表 ***
            wksrpt.Name = strOldAccNo
            strMsg(1) = ""
            Call SetExcelEnd(1, strOldAccNo, xlsAp, wksrpt, intTitleR, strOldA0103, strMsg(1))
            If strMsg(1) <> MsgText(601) Then
               strMsg(2) = strMsg(2) & ";" & strMsg(1)
            End If
            intWkPage = intWkPage + 1
            '*** End 不同科目換工作表 ***
            '不同類別,先存檔,再開新檔
            If Left(strOldAccNo, 1) <> Left(adoacc010.Fields("R40611").Value, 1) Then
               strFileN = Text2 & "-" & Text1 & " " & strReportN & "(" & Left(strOldAccNo, 1) & "字頭) " & ACDate(ServerDate) & ServerTime
               '畫面可允許輸跨類輸入,避免工作表無法全數列示,故不同類別換檔案
               If SetExcelEnd(0, strOldAccNo, xlsAp, wksrpt, intTitleR, strOldA0103, strMsg(0), strFileN) = False Then
                  MsgBox strFileN & "-" & strMsg(0) & vbCrLf & _
                                 "儲存失敗,請洽電腦中心"
                  Exit Do
               End If
               bolOpenXls = False
               intWkPage = 1
               xlsAp.SheetsInNewWorkbook = 3
               xlsAp.Workbooks.add
               bolOpenXls = True
            End If
         End If
         If intWkPage > 3 Then
            xlsAp.Worksheets.add After:=wksrpt '插入sheet
         End If
         Set wksrpt = xlsAp.Worksheets(strWkName & intWkPage)
         intRow = 1
         wksrpt.Activate
         Call SetTitle(True, xlsAp, wksrpt, intRow, strReportN)
         intTitleR = intRow
         intRow = intRow + 1
         bolShowPre = False
         strLastAmt = "" & adoacc010.Fields("a0408")
         '1/2/3 開頭科目,顯示 上期餘額
         Select Case Mid(adoacc010.Fields("R40611").Value, 1, 1)
            Case "1", "2", "3"
               strBalance = strLastAmt
               bolShowPre = True
            Case Else
               strBalance = 0
         End Select
         '貸方科目 且為 1字頭科目,若為負數需顯示正數
         If adoacc010.Fields("a0103").Value = "2" And Mid(adoacc010.Fields("R40611").Value, 1, 1) = "1" And Val(strBalance) < 0 Then
            strBalance = Val(strBalance) * -1
         End If
         If bolShowPre = True Then
            wksrpt.Range(Chr(GetColVal(strF, "摘要", LBound(strF)) + intField) & intRow).Value = "上期餘額"
            wksrpt.Range(Chr(GetColVal(strF, "餘額", LBound(strF)) + intField) & intRow).Value = strBalance
            wksrpt.Range(Chr(GetColVal(strF, "餘額", LBound(strF)) + intField) & intRow).NumberFormatLocal = "#,##0.00"
            intPreR = intRow
            intRow = intRow + 1
         End If
         
         '抓暫存檔會計科目資料
         strSql = "Select R40602,R40603,R40608,R40609,R40610,R40607,R40604,R40605,R40606,R40613 " & _
                        "From Accrpt406 Where R40601='" & strUserNum & "' And R40611 = '" & adoacc010.Fields("R40611").Value & "' " & _
                        "Order by r40602 asc, r40603 asc, r40612 asc"
         intQ2 = 1
         Set adoaccrpt406 = ClsLawReadRstMsg(intQ2, strSql)
         If intQ2 = 1 Then
            adoaccrpt406.MoveFirst
            Do While adoaccrpt406.EOF = False
               For i = LBound(strF) To UBound(strF)
                  strFormat = ""
                  strTmp(0) = "" & adoaccrpt406.Fields(i)
                  Select Case strF(i)
                     Case "傳票日期"
                        strTmp(0) = Format(strTmp(0), "###/##/##")
                     Case "借方金額", "貸方金額", "餘額"
                        strFormat = "#,##0.00"
                        If strF(i) = "餘額" Then
                           If bolShowPre = True Then
                              strTmp(0) = Chr(i + intField) & intRow - 1 & "+"
                           Else
                              strTmp(0) = ""
                           End If
                           strTmp(1) = Chr(GetColVal(strF, "借方金額", LBound(strF)) + intField)
                           strTmp(2) = Chr(GetColVal(strF, "貸方金額", LBound(strF)) + intField)
                           If adoacc010.Fields("a0103").Value = "1" Then
                              '原餘額+借-貸
                              strTmp(0) = strTmp(0) & strTmp(1) & intRow & "-" & strTmp(2) & intRow
                           Else
                              '原餘額+貸-借
                              strTmp(0) = strTmp(0) & strTmp(2) & intRow & "-" & strTmp(1) & intRow
                           End If
                           strTmp(0) = "=" & strTmp(0)
                        End If
                     Case "對沖代號(業)", "對沖代號(其他)"
                        strFormat = "@"
                  End Select
                  If strFormat <> MsgText(601) Then
                     wksrpt.Range(Chr(i + intField) & intRow).NumberFormatLocal = strFormat
                  End If
                  If strTmp(0) = MsgText(601) Then strTmp(0) = " "
                  wksrpt.Range(Chr(i + intField) & intRow).Value = strTmp(0)
               Next i
               intRow = intRow + 1
               adoaccrpt406.MoveNext
            Loop
         End If
         strOldA0103 = "" & adoacc010.Fields("a0103").Value
         strOldAccNo = "" & adoacc010.Fields("R40611").Value
         adoacc010.MoveNext
      Loop
      wksrpt.Name = strOldAccNo
      strMsg(1) = ""
      Call SetExcelEnd(1, strOldAccNo, xlsAp, wksrpt, intTitleR, strOldA0103, strMsg(1))
      If strMsg(1) <> MsgText(601) Then
         strMsg(2) = strMsg(2) & ";" & strMsg(1)
      End If
      strFileN = Text2 & "-" & Text1 & " " & strReportN & "(" & Left(strOldAccNo, 1) & "字頭) " & ACDate(ServerDate) & ServerTime
      If SetExcelEnd(0, strOldAccNo, xlsAp, wksrpt, intTitleR, strOldA0103, strMsg(0), strFileN) = False Then
         Exit Function
      End If
      
      If strMsg(2) <> MsgText(601) Then
         MsgBox "會計科目" & vbCrLf & Replace(Mid(strMsg(2), 2), ";", vbCrLf) & vbCrLf & "餘額有誤,請洽電腦中心"
      End If
   End If
   
   SaveExcel = True
   
   Exit Function
    
ErrHnd:
   If bolOpenXls = True Then
      If strFileN = MsgText(601) Then
         strFileN = Text2 & "-" & Text1 & " " & strReportN
      End If
      If Val(xlsAp.Version) < 12 Then
         xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & MsgText(43), FileFormat:=-4143
      Else
         xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ".xlsx", FileFormat:=51
      End If
      xlsAp.Workbooks.Close
      xlsAp.Quit
      Set xlsAp = Nothing
   End If
   MsgBox Err.Description, vbCritical
End Function

Private Sub SetTitle(IsFirst As Boolean, XlsApp As Excel.Application, Wks As Worksheet, ByRef intRow As Integer, Optional ByVal stTitleN As String)
   Dim ii As Integer, intTpR As Integer, stTxt As String

   If IsFirst = True Then
       Wks.Range(Chr(intField) & intRow).Value = stTitleN
       Wks.Range(Chr(intField) & intRow).Font.Size = 18
       Wks.Range(Chr(intField) & intRow).Font.Bold = True
       Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Select
      
       With XlsApp.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
       End With
       intRow = intRow + 1
       '條件
       stTxt = "公司別：　" & A0802Query(strCmp)
       Wks.Range(Chr(intField) & intRow).Value = stTxt
       Wks.Range(Chr(intField) & intRow & ":" & Chr(GetColVal(strF, "案件性質", LBound(strF)) + intField) & intRow).MergeCells = True
       intRow = intRow + 1
       stTxt = "會計科目：" & adoacc010.Fields("R40611").Value
       If "" & adoacc010.Fields("a010２").Value <> MsgText(601) Then
          stTxt = stTxt & " " & adoacc010.Fields("a0102").Value
       End If
       Wks.Range(Chr(intField) & intRow).Value = stTxt
       Wks.Range(Chr(intField) & intRow & ":" & Chr(GetColVal(strF, "案件性質", LBound(strF)) + intField) & intRow).MergeCells = True
       intRow = intRow + 1
       '列印人員/日期
       stTxt = "列印人員：" & StaffQuery(strUserNum)
       Wks.Range(Chr(intField) & intRow).Value = stTxt
       stTxt = "列印日期：" & CFDate(ACDate(ServerDate))
       Wks.Range(Chr((intField + UBound(strF) - 2)) & intRow).Value = stTxt
       intRow = intRow + 1
       intTpR = intRow
   Else
      '畫線
      Wks.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strF)) & intTitleR).Borders(xlEdgeBottom).LineStyle = xlContinuous
      intTpR = intTitleR
   End If
   
   For ii = LBound(strF) To UBound(strF)
      stTxt = strF(ii)
      If IsFirst = False Then
         stTxt = Replace(stTxt, "(", vbCrLf & "(")
         If ii = UBound(strF) Then
            Wks.Range(Chr(intField + ii) & intTpR).RowHeight = 27
         End If
      End If
      Wks.Range(Chr(intField + ii) & intTpR).Value = stTxt
      Wks.Range(Chr(intField + ii) & intTpR).Font.Bold = True
      Wks.Range(Chr(intField + ii) & intTpR).ColumnWidth = Val(arrWidth(ii))
      Wks.Range(Chr(intField + ii) & intTpR).HorizontalAlignment = xlCenter
   Next ii
   
End Sub

'intChosoe:0-存檔/1-換工作表
Private Function SetExcelEnd(intChoose As Integer, stAccNo As String, XlsApp As Excel.Application, Wks As Worksheet, intTitleR As Integer, stA0103 As String, ByRef stMsg As String, Optional ByVal stFieldN As String = "") As Boolean
   Dim RsQ As New ADODB.Recordset, stQ As String, intQ As Integer, ii As Integer, stTP(3) As String, stAmt(2) As String
   
   SetExcelEnd = False
   stQ = "Select Sum(R40608),Sum(R40609),Sum(R40610) From Accrpt406 " & _
            "Where R40601='" & strUserNum & "' And R40611 = '" & stAccNo & "' "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stQ)
   If intQ = 1 Then
      For ii = LBound(stAmt) To UBound(stAmt)
         stAmt(ii) = "" & RsQ.Fields(ii)
      Next ii
   End If
   '加總
   Wks.Range(Chr(GetColVal(strF, "傳票號碼", LBound(strF)) + intField) & intRow).Value = "合　計"
   Wks.Range(Chr(GetColVal(strF, "傳票號碼", LBound(strF)) + intField) & intRow).Font.Bold = True
   Wks.Range(Chr(GetColVal(strF, "傳票號碼", LBound(strF)) + intField) & intRow).HorizontalAlignment = xlCenter
   For ii = GetColVal(strF, "借方金額", LBound(strF)) To GetColVal(strF, "餘額", LBound(strF))
      stTP(1) = stAmt(ii - GetColVal(strF, "借方金額", LBound(strF)))
      If ii = GetColVal(strF, "餘額", LBound(strF)) Then
         If bolShowPre = True Then
            stTP(0) = Chr(ii + intField) & intPreR
            stTP(1) = Val(Wks.Range(stTP(0)).Value) + Val(stTP(1)) '上期餘額+Sum(R40610)
            stTP(0) = stTP(0) & "+"
         '[無]上期餘額
         Else
            stTP(0) = ""
            stTP(1) = stTP(1) 'Sum(R40610)
         End If
         stTP(2) = Chr(GetColVal(strF, "借方金額", LBound(strF)) + intField)
         stTP(3) = Chr(GetColVal(strF, "貸方金額", LBound(strF)) + intField)
         If stA0103 = "1" Then
            '原餘額+借-貸
            stTP(0) = stTP(0) & "Sum(" & stTP(2) & intTitleR + 1 & ":" & stTP(2) & intRow - 1 & ")" & _
                                             "-Sum(" & stTP(3) & intTitleR + 1 & ":" & stTP(3) & intRow - 1 & ")"
         Else
            '原餘額+貸-借
            stTP(0) = stTP(0) & "Sum(" & stTP(3) & intTitleR + 1 & ":" & stTP(3) & intRow - 1 & ")" & _
                                             "-Sum(" & stTP(2) & intTitleR + 1 & ":" & stTP(2) & intRow - 1 & ")"
         End If
      Else
         stTP(0) = "Sum(" & Chr(ii + intField) & intTitleR + 1 & ":" & Chr(ii + intField) & intRow - 1 & ")"
      End If
      Wks.Range(Chr(ii + intField) & intRow).Value = "=" & stTP(0)
      Wks.Range(Chr(ii + intField) & intRow).NumberFormatLocal = "#,##0.00"
      stTP(0) = Wks.Range(Chr(ii + intField) & intRow).Value
      If Val(stTP(1)) <> Val(stTP(0)) Then
         stMsg = stMsg & ";" & stAccNo
      End If
   Next ii
   '內容字大小
   Wks.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strF)) & intRow).Font.Size = 10
   '畫線
   Wks.Range(Chr(intField) & intRow - 1 & ":" & Chr(intField + UBound(strF)) & intRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
   '更新欄名
   Call SetTitle(False, XlsApp, Wks, intRow)
   '版面設定
   Wks.PageSetup.Orientation = xlPortrait '直印
   Wks.PageSetup.Zoom = 100 '縮放比例為100%,列印頁面水平置中
   Wks.PageSetup.HeaderMargin = XlsApp.Application.InchesToPoints(0) '頁首
   Wks.PageSetup.FooterMargin = XlsApp.Application.InchesToPoints(0) '頁尾
   Wks.PageSetup.TopMargin = XlsApp.InchesToPoints(0.3) '上
   Wks.PageSetup.BottomMargin = XlsApp.InchesToPoints(0.3) '下
   Wks.PageSetup.LeftMargin = XlsApp.InchesToPoints(0.1) '左邊界
   Wks.PageSetup.RightMargin = XlsApp.InchesToPoints(0.1) '右邊界
   Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
   Wks.PageSetup.CenterHorizontally = True '水平置中(版面設定->邊界->水平置中)
   
   If intChoose = 0 Then
      If stFieldN = MsgText(601) Then
         stMsg = stMsg & ";檔案名稱為空"
      Else
         If Dir(strExcelPath & stFieldN & ".xlsx") = MsgText(601) Then
            If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
               MkDir strExcelPath
            End If
         Else
            Kill strExcelPath & stFieldN & ".xlsx"
         End If
         
         If Val(XlsApp.Version) < 12 Then
            XlsApp.Workbooks(1).SaveAs FileName:=strExcelPath & stFieldN & MsgText(43), FileFormat:=-4143
         Else
            XlsApp.Workbooks(1).SaveAs FileName:=strExcelPath & stFieldN & ".xlsx", FileFormat:=51
         End If
         XlsApp.Workbooks.Close
         XlsApp.Quit
      End If
   End If
   If stMsg <> MsgText(601) Then
      stMsg = Mid(stMsg, 2)
   End If
   If InStr(stMsg, "檔案名稱為空") > 0 Then Exit Function
   SetExcelEnd = True
End Function



   
