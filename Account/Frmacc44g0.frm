VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44g0 
   AutoRedraw      =   -1  'True
   Caption         =   "會計傳票列印"
   ClientHeight    =   2760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5004
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   5004
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      Style           =   2  '單純下拉式
      TabIndex        =   20
      Top             =   1380
      Width           =   3540
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "A4中一刀紙(A4空白紙開啟Excel印)-事務機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   2130
      Width           =   4800
   End
   Begin VB.ComboBox CboComp 
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
      Left            =   1170
      TabIndex        =   0
      Text            =   "CboComp"
      Top             =   60
      Width           =   3520
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
      Height          =   315
      Left            =   3090
      MaxLength       =   10
      TabIndex        =   4
      Top             =   990
      Width           =   1572
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   930
      TabIndex        =   6
      Top             =   3540
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
      Height          =   330
      Left            =   3570
      TabIndex        =   7
      Top             =   3540
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   930
      TabIndex        =   8
      Top             =   3900
      Visible         =   0   'False
      Width           =   1812
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
      Height          =   330
      Left            =   3570
      TabIndex        =   9
      Top             =   3900
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "連續報表紙(紙有表格套印)-點陣印表機-(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
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
      Top             =   2796
      Width           =   4800
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
      Height          =   315
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   3
      Top             =   990
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1170
      TabIndex        =   1
      Top             =   540
      Width           =   1575
      _ExtentX        =   2773
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3090
      TabIndex        =   2
      Top             =   540
      Width           =   1575
      _ExtentX        =   2773
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
   Begin VB.CheckBox Check1 
      Caption         =   "新格式A4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   19
      Top             =   1770
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "印A4空白中一刀，請勿開啟Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   810
      TabIndex        =   22
      Top             =   1740
      Width           =   4005
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
      Left            =   180
      TabIndex        =   21
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   330
      TabIndex        =   17
      Top             =   3180
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   690
      TabIndex        =   16
      Top             =   3540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2970
      Picture         =   "Frmacc44g0.frx":0000
      Stretch         =   -1  'True
      Top             =   3540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   690
      TabIndex        =   15
      Top             =   3900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   240
      Top             =   3060
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2970
      Picture         =   "Frmacc44g0.frx":0442
      Stretch         =   -1  'True
      Top             =   3900
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
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
      Left            =   2850
      TabIndex        =   14
      Top             =   990
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "傳票號碼"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   2850
      TabIndex        =   12
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Left            =   240
      TabIndex        =   11
      Top             =   540
      Width           =   975
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
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc44g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/01/18 Form2.0已修改 (改為Excel A4中一刀套印,分上下表格)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit

Dim adoacc020 As New ADODB.Recordset
Dim adoacc021 As New ADODB.Recordset
Dim adoaccsum As New ADODB.Recordset
Dim strSort1 As String
Dim strSort2 As String
Private Const intDefault As Integer = 300
Dim strDefCmp As String 'Add by Amy 2020/03/16
'Add by Sindy 2020/6/15
Dim m_FileName As String
Dim m_Page As Integer '1~2 1:上頁 2:下頁
Dim m_PRow As Integer '1~11,12~22 一張傳票只能列印11列
'2020/6/15 END
Dim strPrinter As String 'Add By Sindy 2020/7/8
Dim i As Integer 'Add by Amy 2022/01/18

Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 ' 只可選(用單純下拉預設會錯)
End Sub

'Mark by Amy 2024/07/18 不再使用[連續報表紙]-套印
'Private Sub Command1_Click()
'   'Add By Sindy 2020/7/8
'   If Trim(CboComp) = MsgText(601) Then
'      MsgBox "公司別不可空白！", , MsgText(5)
'      Exit Sub
'   End If
'   '2020/7/8 END
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   'Add by Amy 2023/03/22 加印表機判斷
'   If InStr(Combo1, "LQ") = 0 And Combo1 <> "PDFCreator" And Combo1 <> "PDF reDirect v2" Then
'        MsgBox "印表機只能選點「陣印表機」 " & vbCrLf & _
'                        "或　PDF Creator 或 PDF reDirect v2"
'        Exit Sub
'   End If
'   Call SetPrinter(False)
'   'end 2023/03/21
'   m_Page = 0: m_PRow = 0 'Add By Sindy 2020/6/16
'   Screen.MousePointer = vbHourglass
'   '2014/1/16 cancel by sonia
'   'If Text5 = "1" Then
'   'Modify By Sindy 2020/6/15
'   'Modify by Amy 2023/03/21 避免混洧,將Check1預設勾選拿掉(不使用)
''   If Check1.Value = 1 Then
''      ProcessDataI_New '開Word以Printer印(造字印不出)
''   Else
'   '2020/6/15 END
'      '連續報表紙-套印
'      ProcessDataI
''   End If
'   'end 2013/03/21
'   'Else
'   '   ProcessData
'   'End If
'   Call SetPrinter(True) 'Add by Amy 2023/03/21 還原印表機
'   MsgBox "傳票已列印完成！" 'Add by Amy 2024/02/01
'   Screen.MousePointer = vbDefault
'   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
'End Sub

'A4 中一刀紙(AA=4空白紙開啟Excel印)-事務機
Private Sub Command3_Click()
    Dim strMsg As String, strNo(1) As String, strCmp As String

    If FormCheck(strMsg) = False Then
        If strMsg = MsgText(601) Then strMsg = MsgText(181)
        MsgBox strMsg, , MsgText(5)
        Exit Sub
    End If
    
    strExc(0) = GetAcc020Sql(strCmp)
    If adoacc020.State = adStateOpen Then adoacc020.Close
    adoacc020.CursorLocation = adUseClient
    adoacc020.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
    If adoacc020.RecordCount = 1 Then
        If "" & adoacc020.Fields("VNo") = "N" Then
            adoacc020.Close
            StatusClear
            MsgBox MsgText(9010)
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    i = 0
    adoacc020.MoveFirst
    Do While adoacc020.EOF = False
        strNo(i) = "" & adoacc020.Fields("VNo")
        i = i + 1
        adoacc020.MoveNext
    Loop
    If adoacc020.RecordCount = 1 Then strNo(1) = strNo(0)
    Call SetPrinter(False)
    If PrintVoucherExcel(Me.Name, strCmp, strNo(0), strNo(1)) = False Then
        Call SetPrinter(True)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call SetPrinter(True)
    MsgBox "傳票已列印完成！"
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

'Private Sub Command2_Click()
'Dim i As Integer
'Dim intHeight As Integer
'
'   Printer.PaperSize = PUB_GetPaperSize(9)
'   'Printer.FontSize = 10
'   Printer.Font = "新細明體"
'
'   Printer.Font.Size = 20
'      Printer.Font.Underline = False
'      Printer.FontBold = True
'      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("會 計 傳 票") / 2)
'      Printer.CurrentY = 300 + intHeight
'      Printer.Print "會 計 傳 票"
'      Printer.Font.Underline = True
'      Printer.FontBold = True
'      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("會 計 傳 票") / 2)
'      Printer.CurrentY = 800 + intHeight
'      Printer.Print "會 計 傳 票"
'
'      Printer.Font.Size = 12
'      Printer.Font.Underline = False
'      Printer.FontBold = False
'      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("中華民國 " & "會 計 傳 票") / 2)
'      Printer.CurrentY = 1300 + intHeight
'      Printer.Print "中華民國 " & "會 計 傳 票"
'      Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("No：" & "會 計 傳 票") - 1500
'      Printer.CurrentY = 1300 + intHeight
'      Printer.Print "No：" & "會 計 傳 票"
'      '方格
'      Printer.Line (400, 1600 + intHeight)-(11000, 7500 + intHeight), , B
'      '橫線
'      Printer.Line (400, 2000 + intHeight)-(11000, 2000 + intHeight)
'      Printer.Line (400, 7100 + intHeight)-(11000, 7100 + intHeight)
'      '直線
'      Printer.Line (3100, 1600 + intHeight)-(3100, 7100 + intHeight)
'      Printer.Line (4000, 1600 + intHeight)-(4000, 7500 + intHeight)
'      Printer.Line (5500, 1600 + intHeight)-(5500, 7500 + intHeight)
'      Printer.Line (7000, 1600 + intHeight)-(7000, 7500 + intHeight)
'      Printer.CurrentX = 1100
'      Printer.CurrentY = 1700 + intHeight
'      Printer.Print "會  計  科  目"
'      Printer.CurrentX = 3300
'      Printer.CurrentY = 1700 + intHeight
'      Printer.Print "部門"
'      Printer.CurrentX = 4150
'      Printer.CurrentY = 1700 + intHeight
'      Printer.Print "借 方 金 額"
'      Printer.CurrentX = 5650
'      Printer.CurrentY = 1700 + intHeight
'      Printer.Print "貸 方 金 額"
'      Printer.CurrentX = 8300
'      Printer.CurrentY = 1700 + intHeight
'      Printer.Print "摘　　　　　要"
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 7200 + intHeight
'      Printer.Print "合　　　　　計"
'      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("核准　　　　　　會計　　　　　　出納　　　　　　製單　　　　　　") / 2)
'      Printer.CurrentY = 7600 + intHeight
'      Printer.Print "核准　　　　　　會計　　　　　　出納　　　　　　製單　　　　　　"
'
'   Printer.EndDoc
'End Sub

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
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   'Modify by Amy 2024/07/18 原:3435 不再使用連續報表紙套印
   Me.Height = 3200 '2895 '2700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
 
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo5.AddItem MsgText(1)
   Combo5.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo5 = MsgText(1)
   ComboAdd
   'Modify by Amy 2020/03/16 改為下拉式選單
   'Text5 = "1"
   strDefCmp = "1"
   Call Pub_SetCboCmp(CboComp, False, False, False, strDefCmp)
   'end 2020/03/16
   
 
'   'Add By Sindy 2020/6/15 (以Word印上下,若下方無資料不好控制表格不印,故先不使用)
'   m_FileName = "$$會計傳票.doc"
'   If Dir(App.path & "\" & strUserNum & "\" & m_FileName) <> "" Then
'      Kill App.path & "\" & strUserNum & "\" & m_FileName
'   End If
'   Call PUB_GetSampleFile(m_FileName, "M31-000009-0-00", , App.path & "\" & strUserNum & "\")
'   '2020/6/15 END
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add by Sindy 2020/7/8
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2020/7/8
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2020/7/8 END
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44g0 = Nothing
End Sub

'2014/1/16 cancel by sonia
'*************************************************
'  產生報表資料
'
'*************************************************
'Private Sub ProcessData()
'Dim strOrder1, strOrder2 As String
'Dim strSql As String, intCounter As Integer
'Dim strAmount As String, intLength As Integer
'
'On Error GoTo Checking
'   Me.MousePointer = vbHourglass
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   Select Case Combo2
'      Case strSort1
'         If Combo5 = MsgText(1) Then
'            strOrder2 = ", a0202 asc"
'         Else
'            strOrder2 = ", a0202 desc"
'         End If
'      Case strSort2
'         If Combo5 = MsgText(1) Then
'            strOrder2 = ", a0205 asc"
'         Else
'            strOrder2 = ", a0205 desc"
'         End If
'      Case Else
'         strOrder2 = MsgText(601)
'   End Select
'   Select Case Combo3
'      Case strSort1
'         If Combo4 = MsgText(1) Then
'            strOrder1 = " order by a0202 asc"
'         Else
'            strOrder1 = " order by a0202 desc"
'         End If
'      Case strSort2
'         If Combo4 = MsgText(1) Then
'            strOrder1 = " order by a0205 asc"
'         Else
'            strOrder1 = " order by a0205 desc"
'         End If
'      Case Else
'         strOrder1 = MsgText(601)
'   End Select
'   If Text5 <> MsgText(601) Then
'      strSql = " and a0201 = '" & Text5 & "'"
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text2 <> MsgText(601) Then
'      strSql = strSql & " and a0202 >= '" & Text2 & "'"
'   End If
'   If Text1 <> MsgText(601) Then
'      strSql = strSql & " and a0202 <= '" & Text1 & "'"
'   End If
'   If strSql <> MsgText(601) Then
'      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
'   End If
'
'   Printer.FontSize = intFontSize
'   adoacc020.CursorLocation = adUseClient
'   adoacc020.Open "select * from acc020" & strSql & " order by a0201 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc020.RecordCount = 0 Then
'      adoacc020.Close
'      StatusClear
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
'
'   Do While adoacc020.EOF = False
'      PrintHead
'      adoacc021.CursorLocation = adUseClient
'      adoacc021.Open "select * from acc021, acc010 where ax205 = a0101 and ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "' order by ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoacc021.EOF = False
'         Printer.CurrentX = 300
'         Printer.CurrentY = 2050 + intCounter * 300
'         Printer.Print IIf(IsNull(adoacc021.Fields("ax215").Value), MsgText(601), adoacc021.Fields("ax215").Value)
'         Printer.CurrentX = 700
'         Printer.CurrentY = 2050 + intCounter * 300
'         Printer.Print IIf(IsNull(adoacc021.Fields("a0102").Value), MsgText(601), adoacc021.Fields("a0102").Value)
'         Printer.CurrentX = 3600
'         Printer.CurrentY = 2050 + intCounter * 300
'         Printer.Print IIf(IsNull(adoacc021.Fields("ax212").Value), MsgText(601), adoacc021.Fields("ax212").Value)
'         strAmount = Format(IIf(IsNull(adoacc021.Fields("ax206").Value), 0, adoacc021.Fields("ax206").Value), FDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 8900 - intLength
'         Printer.CurrentY = 2050 + intCounter * 300
'         Printer.Print strAmount
'         strAmount = Format(IIf(IsNull(adoacc021.Fields("ax207").Value), 0, adoacc021.Fields("ax207").Value), FDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 10800 - intLength
'         Printer.CurrentY = 2050 + intCounter * 300
'         Printer.Print strAmount
'         intCounter = intCounter + 1
'         If intCounter > 14 Then
'            intCounter = 0
'            Printer.NewPage
'            PrintHead
'         End If
'         adoacc021.MoveNext
'      Loop
'      adoacc021.Close
'      adoaccsum.CursorLocation = adUseClient
'      adoaccsum.Open "select sum(ax206), sum(ax207) from acc021 where ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'      If adoaccsum.RecordCount <> 0 Then
'         strAmount = Format(IIf(IsNull(adoaccsum.Fields(0).Value), 0, adoaccsum.Fields(0).Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 8900 - intLength
'         Printer.CurrentY = 4500
'         Printer.Print strAmount
'         strAmount = Format(IIf(IsNull(adoaccsum.Fields(1).Value), 0, adoaccsum.Fields(1).Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 10800 - intLength
'         Printer.CurrentY = 4500
'         Printer.Print strAmount
'         Printer.NewPage
'      Else
'         Printer.NewPage
'      End If
'      adoaccsum.Close
'      intCounter = 0
'      adoacc020.MoveNext
'   Loop
'   adoacc020.Close
'   Printer.EndDoc
'   Me.MousePointer = vbDefault
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub
'2014/1/16 end

'Added by Morgan 2012/6/21 傳票不會跨日列印--瑞婷
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox2 = MaskEdBox1
   End If
End Sub

Private Sub MaskEdBox2_GotFocus()
   If MaskEdBox2 = MsgText(29) And MaskEdBox1 <> MsgText(29) Then
      MaskEdBox2 = MaskEdBox1
   End If
   MaskEdBoxInverse MaskEdBox2
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
   'ADD BY SONIA 2016/9/8
   If Text1 = "" Then
      Text1 = Text2
      If Len("" & Text1) > 0 Then
         Text1.SelStart = Len("" & Text1)
         Text1.SelLength = 0
      End If
   End If
   'END 2016/9/8
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      '92.8.25 modify by sonia
      'MsgBox MsgText(210), , MsgText(5)
      'Cancel = True
      'Text1.SetFocus
      'Exit Sub
      If Text2 <> MsgText(601) Then
         MsgBox MsgText(210), , MsgText(5)
         Cancel = True
         Text1.SetFocus
         Exit Sub
      End If
      '92.8.25 end
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2020/03/16 公司別改下拉選單
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   CloseIme
'End Sub

'Private Sub Text5_Change()
'  Text6 = A0802Query(Text5)
'End Sub

'2014/1/16 add by sonia
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

'Private Sub Text5_Validate(Cancel As Boolean)
'   If Text5 = MsgText(601) Then
'      MsgBox MsgText(10) & Label1, , MsgText(5)
'      Cancel = True
'      Text5.SetFocus
'      Exit Sub
'   Else
'      If Text5 <> "1" And Text5 <> "J" Then
'         MsgBox "只可輸入 1 或 J", vbCritical
'         Cancel = True
'         Text5.SetFocus
'         Exit Sub
'      End If
'   End If
'End Sub
'2014/1/16 end

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "傳票號碼"
   strSort2 = "傳票日期"
   Combo2.AddItem strSort1
   Combo2.AddItem strSort2
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   Printer.CurrentX = 5500
   Printer.CurrentY = 1000
   Printer.Print Val(Mid(CFDate(adoacc020.Fields("a0205").Value), 1, 3))
   Printer.CurrentX = 6500
   Printer.CurrentY = 1000
   Printer.Print Mid(CFDate(adoacc020.Fields("a0205").Value), 5, 2)
   Printer.CurrentX = 7500
   Printer.CurrentY = 1000
   Printer.Print Mid(CFDate(adoacc020.Fields("a0205").Value), 8, 2)
   Printer.CurrentX = 9500
   Printer.CurrentY = 1000
   Printer.Print adoacc020.Fields("a0202").Value
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
   Text2 = ""
   Text1 = ""
   Combo3 = ""
   Combo2 = ""
   'Modify by Amy 2020/03/16
   'Text5 = "1"
   'Text5.SetFocus
   CboComp = strDefCmp
   'end 2020/03/16
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProcessDataI()
Dim strOrder1 As String, strOrder2 As String
Dim strSql As String, intCounter As Integer
Dim strAmount As String, intLength As Integer
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo2
      Case strSort1
         If Combo5 = MsgText(1) Then
            strOrder2 = ", a0202 asc"
         Else
            strOrder2 = ", a0202 desc"
         End If
      Case strSort2
         If Combo5 = MsgText(1) Then
            strOrder2 = ", a0205 asc"
         Else
            strOrder2 = ", a0205 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0202 asc"
         Else
            strOrder1 = " order by a0202 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0205 asc"
         Else
            strOrder1 = " order by a0205 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   'Modify by Amy 2020/03/16 公司別改下拉 原:Text5
   If CboComp <> MsgText(601) Then
      strSql = " and a0201 = '" & Mid(CboComp, 1, Val(InStr(CboComp, "　")) - 1) & "'"
   End If
   'end 2020/03/16
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0202 >= '" & Text2 & "'"
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0202 <= '" & Text1 & "'"
   Else
      '92.8.25 modify by sonia
      'MsgBox MsgText(210), , MsgText(5)
      If Text2 <> MsgText(601) Then
         MsgBox MsgText(210), , MsgText(5)
         Exit Sub
      End If
      '92.8.25 end
   End If
   If strSql <> MsgText(601) Then
      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
   End If
   
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   adoacc020.CursorLocation = adUseClient
    'Modify By Cheng 2003/12/08
    '依傳票日期及編號排序
'   adoacc020.Open "select * from acc020" & strSQL & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
   adoacc020.Open "select * from acc020" & strSql & " Order By a0201, a0202 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc020.RecordCount = 0 Then
      adoacc020.Close
      StatusClear
      MsgBox MsgText(9010)
      Exit Sub
   End If
   
   'Modify by Morgan 2008/3/25 控制 9x 才自訂
   If pub_OS = "1" Then
      Printer.Height = 8775
      Printer.Width = 13000
   Else
      Printer.PaperSize = PUB_GetPaperSize(9)
   End If
   'end 2008/3/25
   Printer.FontSize = 10
   Printer.Font = "新細明體"
      
   Do While adoacc020.EOF = False
      PrintHeadI
      
      adoacc021.CursorLocation = adUseClient
      adoacc021.Open "select * from acc021, acc010 where ax205 = a0101 and ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "' order by ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc021.EOF = False
         '作帳公司
         Printer.CurrentX = 300
         Printer.CurrentY = 3050 + intCounter * 300 - intDefault
         Printer.Print IIf(IsNull(adoacc021.Fields("ax215").Value), MsgText(601), adoacc021.Fields("ax215").Value)
         '會計科目
         If adoacc021.Fields("ax206").Value <> 0 Then
            Printer.CurrentX = 700
         Else
            Printer.CurrentX = 1100
         End If
         Printer.CurrentY = 3050 + intCounter * 300 - intDefault
         Printer.Print IIf(IsNull(adoacc021.Fields("a0102").Value), MsgText(601), adoacc021.Fields("a0102").Value)
         '部門別
         Printer.CurrentX = 3250
         Printer.CurrentY = 3050 + intCounter * 300 - intDefault
         Printer.Print IIf(IsNull(adoacc021.Fields("ax204").Value), "", adoacc021.Fields("ax204").Value)
         '借方金額
         strAmount = Format(IIf(IsNull(adoacc021.Fields("ax206").Value), 0, adoacc021.Fields("ax206").Value), FDollar)
         If Val(strAmount) = 0 Then
            strAmount = ""
         End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 5250 - intLength
         Printer.CurrentY = 3050 + intCounter * 300 - intDefault
         Printer.Print strAmount
         '貸方金額
         strAmount = Format(IIf(IsNull(adoacc021.Fields("ax207").Value), 0, adoacc021.Fields("ax207").Value), FDollar)
         If Val(strAmount) = 0 Then
            strAmount = ""
         End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 7100 - intLength
         Printer.CurrentY = 3050 + intCounter * 300 - intDefault
         Printer.Print strAmount
         '摘要
         Printer.CurrentX = 7350
         Printer.CurrentY = 3050 + intCounter * 300 - intDefault
         If IsNull(adoacc021.Fields("ax212").Value) Then
            Printer.Print ""
         Else
            Printer.Print StrToStr(adoacc021.Fields("ax212").Value, 20)
         End If
         intCounter = intCounter + 1
         If intCounter > 14 Then
            adoacc021.MoveNext
            If adoacc021.EOF = False Then
               adoacc021.MovePrevious
               intCounter = 0
               Printer.NewPage
               PrintHeadI
            Else
               adoacc021.MovePrevious
            End If
         End If
         adoacc021.MoveNext
      Loop
      adoacc021.Close
      adoaccsum.CursorLocation = adUseClient
      adoaccsum.Open "select sum(ax206), sum(ax207) from acc021 where ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccsum.RecordCount <> 0 Then
         '借方合計
         strAmount = Format(IIf(IsNull(adoaccsum.Fields(0).Value), 0, adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 5250 - intLength
         Printer.CurrentY = 7450
         Printer.Print strAmount
         '貸方合計
         strAmount = Format(IIf(IsNull(adoaccsum.Fields(1).Value), 0, adoaccsum.Fields(1).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 7100 - intLength
         Printer.CurrentY = 7450
         Printer.Print strAmount
         Printer.NewPage
      Else
         Printer.NewPage
      End If
      adoaccsum.Close
      intCounter = 0
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   Printer.EndDoc
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHeadI()
   'Add by Amy 2014/12/11 '+J公司抬頭
   Dim intPaperWidth As Long
   Dim stCmpNo As String 'Add by Amy 2020/03/16
   Dim stCmpName As String 'Add by Amy 2020/03/18
   
   'Modify by Amy 2020/03/16 改下拉
   stCmpNo = Mid(CboComp, 1, Val(InStr(CboComp, "　")) - 1)
   'Modify By Sindy 2020/7/14 Mark
'   If stCmpNo = "J" Or stCmpNo = "L" Then
   '2020/7/14 END
       Printer.FontSize = 20
       Printer.Font = "新細明體"
       'Add by Amy 2020/03/18
       stCmpName = A0802Query(stCmpNo)
       intPaperWidth = 10415
       Printer.CurrentX = 500 + ((intPaperWidth - Printer.TextWidth(stCmpName)) / 2) '1300
       Printer.CurrentY = 300
       Printer.Print stCmpName
       'end 2020/03/18
        
       Printer.FontSize = 10
       Printer.Font = "新細明體"
'   End If
   'end 2020/03/16
   'end 2014/12/11
   
   '傳票日期
   Printer.CurrentX = 5600
   Printer.CurrentY = 1800
   Printer.Print Val(Mid(CFDate(adoacc020.Fields("a0205").Value), 1, 3))
   Printer.CurrentX = 6450
   Printer.CurrentY = 1800
   Printer.Print Mid(CFDate(adoacc020.Fields("a0205").Value), 5, 2)
   Printer.CurrentX = 7400
   Printer.CurrentY = 1800
   Printer.Print Mid(CFDate(adoacc020.Fields("a0205").Value), 8, 2)
   '傳票號碼
   Printer.CurrentX = 9600
   Printer.CurrentY = 1750
   Printer.Print adoacc020.Fields("a0202").Value
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Amy 2022/01/18 +stMsg,公司別從Command1搬過來
Public Function FormCheck(Optional ByRef stMsg As String = "") As Boolean
   'Modify by Amy 2022/01/18 從Command1搬過來
   stMsg = ""
   If Trim(CboComp) = MsgText(601) Then
        stMsg = "公司別不可空白！"
        Exit Function
    End If
    '92.8.25 modify by sonia
    '傳票起號不是空,迄號為空
    If Text2 <> MsgText(601) And Text1 = MsgText(601) Then
         stMsg = MsgText(210)
        Exit Function
    End If
    '92.8.25 end
    'end 2022/01/18
    
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   
   FormCheck = False
End Function

'Add By Sindy 2020/6/15
'*************************************************
'  用Word範本,套印
'
'*************************************************
Private Sub ProcessDataI_New()
Dim strOrder1 As String, strOrder2 As String
Dim intCounter As Integer, intPage As Integer
Dim stCmpNo As String, stCmpName As String, strDDate As String, strDNo As String
Dim strItem As String, strDept As String, strAmt1 As String, strAmt2 As String, strNote As String
Dim bolLastRow As Boolean, bolLastPage As Boolean
   
On Error GoTo Checking
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo2
      Case strSort1
         If Combo5 = MsgText(1) Then
            strOrder2 = ", a0202 asc"
         Else
            strOrder2 = ", a0202 desc"
         End If
      Case strSort2
         If Combo5 = MsgText(1) Then
            strOrder2 = ", a0205 asc"
         Else
            strOrder2 = ", a0205 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0202 asc"
         Else
            strOrder1 = " order by a0202 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0205 asc"
         Else
            strOrder1 = " order by a0205 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   strSql = ""
   'Modify by Amy 2020/03/16 公司別改下拉 原:Text5
   If Trim(CboComp) <> MsgText(601) Then
      strSql = strSql & " and a0201 = '" & Mid(CboComp, 1, Val(InStr(CboComp, "　")) - 1) & "'"
   End If
   'end 2020/03/16
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0202 >= '" & Text2 & "'"
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0202 <= '" & Text1 & "'"
   Else
      '92.8.25 modify by sonia
      'MsgBox MsgText(210), , MsgText(5)
      If Text2 <> MsgText(601) Then
         MsgBox MsgText(210), , MsgText(5)
         Exit Sub
      End If
      '92.8.25 end
   End If
   If strSql <> MsgText(601) Then
      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
   End If
   
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   adoacc020.CursorLocation = adUseClient
    'Modify By Cheng 2003/12/08
    '依傳票日期及編號排序
'   adoacc020.Open "select * from acc020" & strSQL & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
   adoacc020.Open "select * from acc020" & strSql & " Order By a0201, a0202 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc020.RecordCount = 0 Then
      adoacc020.Close
      StatusClear
      MsgBox MsgText(9010)
      Exit Sub
   Else
      adoacc020.MoveFirst
   End If
   
   '切換印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_RestorePrinter Combo1
   
   Printer.EndDoc
   'Printer.PaperSize = PUB_GetPaperSize(9)
   Printer.PaperSize = 9
   'Printer.FontSize = 10
   Printer.Font = "新細明體"
   
   '傳票主檔
   intPage = 0
   Do While adoacc020.EOF = False
      intPage = intPage + 1
      Frmacc0000.StatusBar1.Panels(1).Text = "處理 " & intPage & " / " & adoacc020.RecordCount & "..."
      '公司名稱
      stCmpNo = "" & adoacc020.Fields("a0201").Value 'Mid(cboComp, 1, Val(InStr(cboComp, "　")) - 1)
      stCmpName = A0802Query(stCmpNo)
      
      '傳票日期
      strDDate = Val(Mid(CFDate(adoacc020.Fields("a0205").Value), 1, 3)) & "年" & _
                  Mid(CFDate(adoacc020.Fields("a0205").Value), 5, 2) & "月" & _
                  Mid(CFDate(adoacc020.Fields("a0205").Value), 8, 2) & "日"
      '傳票號碼
      strDNo = adoacc020.Fields("a0202").Value
      
      If adoacc021.State = adStateOpen Then
         adoacc021.Close
      End If
      adoacc021.CursorLocation = adUseClient
      adoacc021.Open "select * from acc021, acc010 where ax205 = a0101 and ax201 = '" & adoacc020.Fields("a0201").Value & "' and ax202 = '" & adoacc020.Fields("a0202").Value & "' order by ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc021.RecordCount > 0 Then adoacc021.MoveFirst: intCounter = 0
      Do While adoacc021.EOF = False
         intCounter = intCounter + 1 '目前資料筆數
         strItem = "": strDept = "": strAmt1 = "": strAmt2 = "": strNote = "" '清變數值
         '作帳公司
         'Printer.Print IIf(IsNull(adoacc021.Fields("ax215").Value), MsgText(601), adoacc021.Fields("ax215").Value)
         '會計科目
         strItem = IIf(IsNull(adoacc021.Fields("a0102").Value), MsgText(601), adoacc021.Fields("a0102").Value)
         If adoacc021.Fields("ax206").Value = 0 Then
            strItem = "　　" & strItem '貸方會計科目往內縮排
         End If
         '部門別
         strDept = IIf(IsNull(adoacc021.Fields("ax204").Value), "", adoacc021.Fields("ax204").Value)
         '借方金額
         strAmt1 = Format(IIf(IsNull(adoacc021.Fields("ax206").Value), 0, adoacc021.Fields("ax206").Value), FDollar)
         If Val(strAmt1) = 0 Then strAmt1 = ""
         '貸方金額
         strAmt2 = Format(IIf(IsNull(adoacc021.Fields("ax207").Value), 0, adoacc021.Fields("ax207").Value), FDollar)
         If Val(strAmt2) = 0 Then strAmt2 = ""
         '摘要
         If Not IsNull(adoacc021.Fields("ax212").Value) Then
            strNote = StrToStr(adoacc021.Fields("ax212").Value, 20)
         End If
         
         '呼叫Word物件:
         bolLastRow = False
         If intCounter = adoacc021.RecordCount Then
            bolLastRow = True
         End If
         bolLastPage = False
         If intPage = adoacc020.RecordCount Then
            bolLastPage = True
         End If
         If pCallPrint(adoacc020.Fields("a0201").Value, stCmpName, strDDate, _
               strDNo, strItem, strDept, strAmt1, strAmt2, strNote, bolLastRow) = False Then
            GoTo Checking
         End If
         
         adoacc021.MoveNext
      Loop
      adoacc021.Close
      
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   Printer.EndDoc
   
   '還原印表機
   PUB_SetOsDefaultPrinter strPrinter
   PUB_RestorePrinter strPrinter
   
   Frmacc0000.StatusBar1.Panels(1).Text = "列印完畢!"
   'StatusClear
   Exit Sub
   
Checking:
   adoacc021.Close
   adoacc020.Close
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

Private Function pCallPrint(strCmpNo As String, stCmpName As String, _
   strDDate As String, strDNo As String, strItem As String, strDept As String, _
   strAmt1 As String, strAmt2 As String, strNote As String, bolLastRow As Boolean) As Boolean
   
Dim strFileName As String
Dim strName As String
Dim strText As String
Dim i As Integer, j As Integer
Dim strAmt1Tot As String, strAmt2Tot As String
Dim strShapes As String
Dim bolSmall As Boolean
Dim intHeight As Integer
   
On Error GoTo ErrHand
   
   pCallPrint = True
   'm_Page 1~2 1:上頁 2:下頁
   'm_PRow 1~16,17~32 一張傳票只能列印16列
   If m_Page = 0 Then m_Page = 1
   If m_Page = 1 Then
      intHeight = 0
   Else
      intHeight = 8500
   End If
   m_PRow = m_PRow + 1
'標題,整個表格:
   'If m_PRow = 1 Or m_PRow > 17 Then
   If m_PRow = 1 Or m_PRow > 16 Then
      'If m_PRow > 17 Then
      If m_PRow > 16 Then
         m_Page = m_Page + 1
         If m_Page > 2 Then
            Printer.NewPage
            m_Page = 1
            intHeight = 0 'm_Page=1
         Else
            intHeight = 8500 'm_Page=2
         End If
         m_PRow = 1
      End If
      Printer.Font.Size = 20
      Printer.Font.Underline = False
      Printer.FontBold = False 'True
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(stCmpName) / 2)
      Printer.CurrentY = 300 + intHeight
      Printer.Print stCmpName
      Printer.Font.Underline = True
      Printer.FontBold = False 'True
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("會 計 傳 票") / 2)
      Printer.CurrentY = 800 + intHeight
      Printer.Print "會 計 傳 票"
      
      Printer.Font.Size = 12
      Printer.Font.Underline = False
      Printer.FontBold = False
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("中華民國 " & strDDate) / 2)
      Printer.CurrentY = 1300 + intHeight
      Printer.Print "中華民國 " & strDDate
      Printer.Font.Size = 16
      Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("No：" & strDNo) - 500 '1500
      Printer.CurrentY = 1300 + intHeight
      Printer.Print "No：" & strDNo
      Printer.Font.Size = 12
      '方格
      Printer.Line (400, 1600 + intHeight)-(11000, 7500 - 300 + intHeight), , B
      '橫線
      Printer.Line (400, 2000 + intHeight)-(11000, 2000 + intHeight)
      Printer.Line (400, 7200 - 300 + intHeight)-(11000, 7200 - 300 + intHeight)
      '直線
      Printer.Line (3300, 1600 + intHeight)-(3300, 7200 - 300 + intHeight)
      Printer.Line (4000, 1600 + intHeight)-(4000, 7500 - 300 + intHeight)
      Printer.Line (5500, 1600 + intHeight)-(5500, 7500 - 300 + intHeight)
      Printer.Line (7000, 1600 + intHeight)-(7000, 7500 - 300 + intHeight)
      Printer.CurrentX = 1100
      Printer.CurrentY = 1700 + intHeight
      Printer.Print "會  計  科  目"
      Printer.CurrentX = 3400
      Printer.CurrentY = 1700 + intHeight
      Printer.Print "部門"
      Printer.CurrentX = 4150
      Printer.CurrentY = 1700 + intHeight
      Printer.Print "借 方 金 額"
      Printer.CurrentX = 5650
      Printer.CurrentY = 1700 + intHeight
      Printer.Print "貸 方 金 額"
      Printer.CurrentX = 8300
      Printer.CurrentY = 1700 + intHeight
      Printer.Print "摘　　　　　要"
      Printer.CurrentX = 1200
      Printer.CurrentY = 7250 - 300 + intHeight
      Printer.Print "合　　　　　計"
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("核准　　　　　　會計　　　　　　出納　　　　　　製單　　　　　　") / 2)
      Printer.CurrentY = 7600 - 300 + intHeight
      Printer.Print "核准　　　　　　會計　　　　　　出納　　　　　　製單　　　　　　"
   End If
'明細:
   Printer.Font.Size = 10
   '會計科目
   Printer.CurrentX = 450
   Printer.CurrentY = 2050 + (300 * (m_PRow - 1)) + intHeight
   Printer.Print strItem
   '部門
   Printer.CurrentX = 3350
   Printer.CurrentY = 2050 + (300 * (m_PRow - 1)) + intHeight
   Printer.Print strDept
   '借方金額
   Printer.CurrentX = 5450 - Printer.TextWidth(strAmt1)
   Printer.CurrentY = 2050 + (300 * (m_PRow - 1)) + intHeight
   Printer.Print strAmt1
   '貸方金額
   Printer.CurrentX = 6950 - Printer.TextWidth(strAmt2)
   Printer.CurrentY = 2050 + (300 * (m_PRow - 1)) + intHeight
   Printer.Print strAmt2
   '摘要
   Printer.CurrentX = 7050
   Printer.CurrentY = 2050 + (300 * (m_PRow - 1)) + intHeight
   'TextWidth("P124790000/藍德工業股份有限公司/新")
   If TextWidth(strNote) > 2970 Then '肆/美無痕生物科技股份有 1090601^
      Printer.Font.Size = 10 '縮小
   End If
   Printer.Print strNote
   Printer.Font.Size = 10 '12 '正常
'合計:
   If bolLastRow = True Then
      If adoaccsum.State = adStateOpen Then
         adoaccsum.Close
      End If
      adoaccsum.CursorLocation = adUseClient
      adoaccsum.Open "select sum(ax206), sum(ax207) from acc021 where ax201 = '" & strCmpNo & "' and ax202 = '" & strDNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccsum.RecordCount <> 0 Then
         '借方合計
         strAmt1Tot = Format(IIf(IsNull(adoaccsum.Fields(0).Value), 0, adoaccsum.Fields(0).Value), FDollar)
         '貸方合計
         strAmt2Tot = Format(IIf(IsNull(adoaccsum.Fields(1).Value), 0, adoaccsum.Fields(1).Value), FDollar)
      End If
      Printer.CurrentX = 5450 - Printer.TextWidth(strAmt1Tot)
      Printer.CurrentY = 7250 - 300 + intHeight
      Printer.Print strAmt1Tot
      Printer.CurrentX = 6950 - Printer.TextWidth(strAmt2Tot)
      Printer.CurrentY = 7250 - 300 + intHeight
      Printer.Print strAmt2Tot
      If m_Page = 1 Then
         m_Page = m_Page + 1
      Else
         m_Page = 0
         Printer.NewPage
      End If
      m_PRow = 0
   End If
   
   Exit Function
   
ErrHand:
   If Err.Number <> 0 Then
      pCallPrint = False
      MsgBox (Err.Description)
   End If
End Function

'Add by Amy 2022/01/18 語法從拆出來
Private Function GetAcc020Sql(ByRef stCmp As String) As String
    Dim strSql As String

    strSql = "": stCmp = ""
    '公司別
    If Trim(CboComp) <> MsgText(601) Then
        stCmp = Mid(CboComp, 1, Val(InStr(CboComp, "　")) - 1)
        strSql = strSql & " and a0201 = '" & stCmp & "'"
    End If
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
       strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
       strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
    End If
    '傳票起號
    If Text2 <> MsgText(601) Then
       strSql = strSql & " and a0202 >= '" & Text2 & "'"
    End If
    '傳票迄號
    If Text1 <> MsgText(601) Then
       strSql = strSql & " and a0202 <= '" & Text1 & "'"
    End If
    If strSql <> MsgText(601) Then
       strSql = " Where " & Mid(strSql, 5, Len(strSql) - 4)
    End If
    GetAcc020Sql = "Select Nvl(Min(a0202),'N') as VNo From acc020" & strSql & _
                   " Union Select Nvl(Max(a0202),'N') as VNo From acc020" & strSql & _
                   " Order by VNo "
End Function
'end 2022/01/18

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
