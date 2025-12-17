VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1470 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員帳款明細表應收規費明細表"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   5160
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel"
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
      Left            =   2640
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   2160
      Width           =   2300
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1680
      Width           =   3495
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
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   2160
      Width           =   2300
   End
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
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1572
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
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   945
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
   Begin VB.Label Label7 
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
      Left            =   360
      TabIndex        =   13
      Top             =   1695
      Width           =   975
   End
   Begin VB.Label lblSalesName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員名稱"
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
      Left            =   2340
      TabIndex        =   11
      Top             =   660
      Width           =   1350
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(預設三個月以前)"
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
      TabIndex        =   10
      Top             =   1368
      Width           =   2196
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "元以上"
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
      TabIndex        =   9
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "規費底限"
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
      TabIndex        =   8
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      TabIndex        =   7
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label5 
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
      TabIndex        =   6
      Top             =   240
      Width           =   252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳款日期"
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
      TabIndex        =   5
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1470"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0k0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt107 As New ADODB.Recordset
'Dim dllaccrpt107 As Object 'Mark by Amy 2016/09/14
'Add by Amy 2016/09/14
Private Const ciTitleFontSize = 14
Private Const ciFontSize = 12
Private Const ciStartX = 0
Private Const ciStartY = 500
Dim i As Integer, iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrint As String
Dim prnPrint As Printer
Dim strFieldN(), intWidth()
Dim PLeft(0 To 6) As Integer

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   'Add by Amy 2016/09/14 +印表機
   For Each prnPrint In Printers
      If prnPrint.DeviceName = Combo1 Then
         Set Printer = prnPrint
      End If
   Next
   'end 2016/09/14
   Screen.MousePointer = vbHourglass
   Accrpt107Delete
   ProduceData
   If adoaccrpt107.State = adStateOpen Then
      adoaccrpt107.Close
   End If
   adoaccrpt107.CursorLocation = adUseClient
   'Modify by Amy 2016/09/14 避免2個人同時執行加Where,依智權人員之ST15部門+智權人員編號+客戶編號+收據抬頭+收據號碼排序
   strExc(0) = "select * from accrpt107 Where R10701='" & strUserNum & "' " & _
                     "Order by R10707,R10708,R10703,R10704,R10709"
   adoaccrpt107.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt107.RecordCount <> 0 Then
      'Modify by Amy 2016/09/14不使用AccReport
      'dllaccrpt107.Acc1470 ReportTitle(107), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      PrintA4
      If Combo1 <> strPrint Then PUB_RestorePrinter strPrint
   End If
   adoaccrpt107.Close
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

'Add by Amy 2016/09/22 +Excel
Private Sub Command2_Click()
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Accrpt107Delete
    ProduceData
    If adoaccrpt107.State = adStateOpen Then
        adoaccrpt107.Close
    End If
    adoaccrpt107.CursorLocation = adUseClient
    strExc(0) = "select * from accrpt107 Where R10701='" & strUserNum & "' " & _
                     "Order by R10707,R10708,R10703,R10704,R10709"
    adoaccrpt107.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt107.RecordCount <> 0 Then
        If SaveExcel = True Then MsgBox "Excel檔案已產生於" & strExcelPath & "!"
    End If
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
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
'Add by Amy 2016/09/14
Dim strSQLP As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 3120
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2016/09/14 +印表機-瑞婷
   PUB_SetPrinter Me.Name, Combo1, strPrint ''Modified by Morgan 2017/11/8 改呼叫公用函數,原程式移除
   'end 2016/09/14
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Text = ShowDate(ACDate(ServerDate), -90)
   MaskEdBox2.Mask = DFormat
   lblSalesName = ""
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   'Set dllaccrpt107 = CreateObject("AccReport.ReportSelect") 'Mark by Amy 2016/09/14
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2016/09/14若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Set dllaccrpt107 = Nothing 'Mark by Amy 2016/09/14
   Set Frmacc1470 = Nothing
End Sub

Private Sub Text1_Change()
   If Len(Text1) = 5 Then
      lblSalesName = StaffQuery(Text1)
   Else
      lblSalesName = MsgText(601)
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSql As String

On Error GoTo Checking
   If MaskEdBox1.Text <> MsgText(601) Then
      strSql = " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) Then
      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   'Modify by Morgan 2007/10/1 智權人員範圍改成一個
   'If Text1 <> MsgText(601) Then
   '   strSQL = strSQL & " and a0k20 >= '" & Text1 & "'"
   'End If
   'If Text2 <> MsgText(601) Then
   '   strSQL = strSQL & " and a0k20 <= '" & Text2 & "'"
   'End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0k20= '" & Text1 & "'"
   End If
   'end 2007/10/1
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0k07 >= " & Val(Text3) & ""
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt107.CursorLocation = adUseClient
   'Modify by Amy 2016/09/14 避免2個人同時執行加Where
   adoaccrpt107.Open "select * from accrpt107 Where R10701='" & strUserNum & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0k0.CursorLocation = adUseClient
   'Modified by Morgan 2011/11/4 考慮拆收據情形
   'adoacc0k0.Open "select sum(nvl(a0k06, 0) - nvl(a0k17, 0)), sum(nvl(a0k07, 0) - nvl(a0k18, 0)), a0k20, a0k03, a0k04 from acc0k0 where (a0k09 is null or a0k09 = 0) and a0k23 = '000' and (a0k07 - nvl(a0k18, 0)) > 0" & strSql & " group by a0k20, a0k03, a0k04 having sum(a0k06 - a0k17) > 0", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2016/09/14 +st15,a0k01 原:group by a0k20, a0k03, a0k04
   'Modify by Amy 2017/06/08 修正語法(因acc1u0 會多筆,造成多加總)且加a0k37及拿掉a0j04的判斷改Group 及 having
'   strExc(0) = "select sum(nvl(a0j09, 0) - nvl(a1u04, 0) - nvl(a1u07,0) + nvl(a1u09,0))" & _
'      ", sum(nvl(a0j10, 0) - nvl(a1u05, 0) - nvl(a1u08,0) + nvl(a1u10,0)), a0k20, a0k03, a0k04,st15,a0k01" & _
'      " from acc0k0,acc0j0,acc1u0,Staff where (a0k09 is null or a0k09 = 0)" & strSql & _
'      " and a0j13(+)=a0k01 and a0j04 = '000' and a1u02(+)=a0j13 and a1u03(+)=a0j01 And a0k20=St01(+)" & _
'      " group by st15,a0k20, a0k03, a0k04,a0k01" & _
'      " having sum(nvl(a0j10, 0) - nvl(a1u05, 0) - nvl(a1u08,0) + nvl(a1u10,0)) > 0"
   strExc(0) = "Select Sum(Nvl(a0j09,0)),Sum(Nvl(a0j10,0)), a0k20, a0k03, a0k04,st15 From (" & _
                    "Select Sum(Nvl(a0j09, 0) - Nvl(a1u04, 0) - Nvl(a1u07,0) + Nvl(a1u08,0)) as a0j09, Sum(Nvl(a0j10, 0) - Nvl(a1u05, 0) - Nvl(a1u09,0) + Nvl(a1u10,0)) as a0j10, a0k20, a0k03, a0k04,st15,a0k01 " & _
                    "From acc0k0,acc0j0,Staff, " & _
                    "(Select a0k01 as Eno,a1u03,Sum(Nvl(a1u04,0)) as a1u04, Sum(Nvl(a1u05,0)) as a1u05, Sum(Nvl(a1u07,0)) as a1u07, Sum(Nvl(a1u08,0)) as a1u08, Sum(Nvl(a1u09,0)) as a1u09, Sum(Nvl(a1u10,0)) as a1u10 " & _
                    "From acc0k0,acc1u0,Staff " & _
                    "Where (a0k09 is null or a0k09 = 0) And a0k01=a1u02(+) And a0k20=St01(+) And a0k37 is null " & strSql & _
                    " Group by st15,a0k20, a0k03, a0k04,a0k01,a1u03) Acc1u0 " & _
                    "Where  (a0k09 is null or a0k09 = 0) And a0k01= a0j13(+) And a0k20=St01(+) And a0j13= Eno(+) And a0j01= a1u03(+)  And a0k37 is null " & strSql & _
                    " Group by st15,a0k20, a0k03, a0k04,a0k01 Having Sum(Nvl(a0j10, 0) - Nvl(a1u05, 0) - Nvl(a1u09,0) + Nvl(a1u10,0)) > 0" & _
                    ") Group by st15,a0k20, a0k03, a0k04"
   adoacc0k0.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   'end 2011/11/4
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      adoaccrpt107.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0k0.EOF = False
      adoaccrpt107.AddNew
      adoaccrpt107.Fields("r10701").Value = strUserNum
      'Modify by Amy 2016/09/14 +員工編號
      If IsNull(adoacc0k0.Fields(2).Value) Then
         adoaccrpt107.Fields("r10702").Value = Null
         adoaccrpt107.Fields("r10708").Value = Null
      Else
         adoaccrpt107.Fields("r10702").Value = StaffQuery(adoacc0k0.Fields(2).Value)
         adoaccrpt107.Fields("r10708").Value = adoacc0k0.Fields(2).Value
      End If
      'end 2016/09/14
      If IsNull(adoacc0k0.Fields(3).Value) Then
         adoaccrpt107.Fields("r10703").Value = Null
      Else
         adoaccrpt107.Fields("r10703").Value = adoacc0k0.Fields(3).Value
      End If
      If IsNull(adoacc0k0.Fields(4).Value) Then
         adoaccrpt107.Fields("r10704").Value = Null
      Else
         adoaccrpt107.Fields("r10704").Value = adoacc0k0.Fields(4).Value
      End If
      If IsNull(adoacc0k0.Fields(0).Value) Then
         adoaccrpt107.Fields("r10706").Value = 0
      Else
         adoaccrpt107.Fields("r10706").Value = adoacc0k0.Fields(0).Value
      End If
      If IsNull(adoacc0k0.Fields(1).Value) Then
         adoaccrpt107.Fields("r10705").Value = 0
      Else
         adoaccrpt107.Fields("r10705").Value = adoacc0k0.Fields(1).Value
      End If
      'Add by Amy 2016/09/14 +st15/a0k01
      If IsNull(adoacc0k0.Fields("st15").Value) Then
         adoaccrpt107.Fields("r10707").Value = Null
      Else
         adoaccrpt107.Fields("r10707").Value = adoacc0k0.Fields("st15").Value
      End If
      'Mark by Amy 2017/06/08
'      If IsNull(adoacc0k0.Fields("a0k01").Value) Then
'         adoaccrpt107.Fields("r10709").Value = Null
'      Else
'         adoaccrpt107.Fields("r10709").Value = adoacc0k0.Fields("a0k01").Value
'      End If
      'end 2016/09/14
      adoaccrpt107.UpdateBatch
      adoacc0k0.MoveNext
   Loop
   'Add by Amy 2016/09/14
   If adoacc0k0.RecordCount > 0 Then
      UpdateSum
   End If
   adoacc0k0.Close
   adoaccrpt107.Close
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
Private Sub Accrpt107Delete()
   adoTaie.Execute "delete from accrpt107 Where R10701='" & strUserNum & "'"
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
   MaskEdBox2.Text = ShowDate(ACDate(ServerDate), -90)
   MaskEdBox2.Mask = DFormat
   Text1 = ""
   
   'Text2 = ""
   Text3 = ""
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
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
   'Remove by Morgan 2007/10/2
   'If Text2 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If
   'end 2007/10/2
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add by Amy 2016/09/14
'增加業務人員小計/部門合計/總計
Private Sub UpdateSum()
    Dim strU As String
    
    '業務人員小計
    strU = "Insert Into Accrpt107 (R10701,R10704,R10705,R10706,R10707,R10708) " & _
            "Select '" & strUserNum & "',' 小    計 ',Sum(R10705),Sum(R10706),R10707,R10708||'Z' " & _
            "From accrpt107 Where R10701='" & strUserNum & "' Group by R10707,R10708 "
    cnnConnection.Execute strU
    '部門合計
    strU = "Insert Into Accrpt107 (R10701,R10704,R10705,R10706,R10707) " & _
            "Select '" & strUserNum & "',a0902||'合計',Sum(R10705),Sum(R10706),R10707||'Z' " & _
            "From accrpt107,Acc090 Where R10701='" & strUserNum & "' And R10707=A0901 And Instr(R10704,' 小    計 ')=0 " & _
            "Group by R10707,a0902 "
    cnnConnection.Execute strU
    '總計
    strU = "Insert Into Accrpt107 (R10701,R10704,R10705,R10706,R10707) " & _
            "Select '" & strUserNum & "','總      計',Sum(R10705),Sum(R10706),'ZZZZZZ' " & _
            "From accrpt107 Where R10701='" & strUserNum & "' And Instr(R10704,' 小    計 ')>0 "
    cnnConnection.Execute strU
End Sub

Private Sub GetPleft() '明細表邊界
    ReDim strFieldN(5)
    ReDim intWidth(5)
    
    strFieldN = Array("智權人員", "客戶編號", "收據抬頭", "未收規費", "未收服務費")
    intWidth = Array(0, 1300, 2800, 7000, 9000)
       
    Printer.Font.Name = "新細明體"
    Printer.Font.Size = ciFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    
    Erase PLeft
    
    For i = LBound(strFieldN) To UBound(strFieldN)
        PLeft(i) = intWidth(i)
    Next i
End Sub

Private Sub PrintNewLine()
    iPrint = iPrint + lngLineHeight
    If iPrint >= (lngPageHeight - 4 * lngLineHeight) Then
        Printer.CurrentX = ciStartX
        Printer.CurrentY = iPrint
        
        iPage = iPage + 1
        Printer.NewPage
        PrintHeader
    End If
    
End Sub

Private Sub PrintLine(Optional intChoose As Integer = 0)
    Select Case intChoose
        Case 0 '抬頭
            Printer.Line (PLeft(0), iPrint)-(lngPageWidth - 200, iPrint)
        Case 1 '小計
            If iPrint > 2500 Then
                iPrint = iPrint + 50
                Printer.Line (PLeft(0), iPrint)-(lngPageWidth - 200, iPrint)
            End If
    End Select
    iPrint = iPrint + 150
End Sub

Private Sub PrintHeader()
    Dim strPTmp As String
    Dim pa1 As Integer
   
    iPrint = ciStartY
    Printer.Font.Size = ciTitleFontSize
    Printer.Font.Bold = True
    Printer.Font.Underline = False
    
    '報表抬頭
    strPTmp = ReportTitle(107)
    Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print strPTmp
    
    PrintNewLine
    PrintNewLine
    
    Printer.Font.Size = ciFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    
    strPTmp = "帳款日期：" & IIf(Val(FCDate(MaskEdBox1.Text)) > 0, MaskEdBox1.Text, "096/01/01") & _
                    "～ " & MaskEdBox2.Text
    pa1 = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
    Printer.CurrentX = pa1
    Printer.CurrentY = iPrint
    Printer.Print strPTmp
    PrintNewLine
    
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "列印人員：" & strUserName
    
    Printer.CurrentX = lngPageWidth - 3000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & CFDate(strSrvDate(2))
    PrintNewLine
    
    Printer.CurrentX = lngPageWidth - 3000
    Printer.CurrentY = iPrint
    Printer.Print "頁　　次：" & Printer.Page
    PrintNewLine
    
    For i = 0 To UBound(strFieldN)
       '顯示／欄位
        If strFieldN(i) <> "" Then
            strPTmp = strFieldN(i)
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iPrint
            Printer.Print strPTmp
        End If
    Next i
    
    PrintNewLine
    
    PrintLine

End Sub

Private Sub PrintA4()
    Dim intPrintX As Integer
    Dim strTemp As String, strOldTemp As String
    
    Printer.Orientation = 1 '1.直印 2.橫印
    Printer.PaperSize = 9 'A4
       
    lngPageHeight = Printer.ScaleHeight
    lngPageWidth = Printer.ScaleWidth
    lngLineHeight = 300
    
    iPage = 0
    GetPleft
    PrintHeader '列印表頭
    
    With adoaccrpt107
        Do While .EOF = False
            iPage = iPage + 1
            If InStr("" & .Fields("R10704"), " 小    計 ") > 0 Or InStr(strOldTemp, " 小    計 ") > 0 Then
                PrintLine 1
            ElseIf InStr("" & .Fields("R10704"), "合計") > 0 Then
                PrintLine 1
            ElseIf InStr(strOldTemp, "合計") > 0 Then
                iPrint = iPrint + 50
                PrintLine 1
                PrintNewLine
            End If
            For i = 0 To UBound(strFieldN)
                intPrintX = PLeft(i)
                Select Case i
                    Case 0 '智權人員
                        strTemp = "" & .Fields("R10702")
                    Case 1 '客戶編號
                        strTemp = "" & .Fields("R10703")
                    Case 2 '收據抬頭
                        strTemp = PUB_StrToStr_byVal("" & .Fields("R10704"), 32)
                    Case 3 '未收規費
                        strTemp = Format(Val("" & .Fields("R10705")), "###,##0.00")
                        intPrintX = PLeft(i + 1) - Printer.TextWidth(strTemp) - 500
                    Case 4 '未收服務費
                        strTemp = Format(Val("" & .Fields("R10706")), "###,##0.00")
                        intPrintX = lngPageWidth - Printer.TextWidth(strTemp) - 1500
                End Select
                Printer.CurrentX = intPrintX
                Printer.CurrentY = iPrint
                Printer.Print strTemp
            Next i
            strOldTemp = "" & .Fields("R10704")
            .MoveNext
            PrintNewLine
        Loop
    End With
    Printer.EndDoc
End Sub

'Add by Amy 2016/09/22 +Excel
Private Function SaveExcel() As Boolean
    Dim xlsAnnuity As New Excel.Application
    Dim wksAnnuity As New Worksheet
    Dim intField As Integer, intCounter As Integer, intTitleRow As Integer
    Dim strFileName As String
    Dim strTemp As String
    Dim strStart As String, strSum As String, strTotal As String

On Error GoTo ErrHnd

    strFileName = strExcelPath & "智權人員應收規費明細表" & ACDate(ServerDate) & ServerTime & MsgText(43)
    If Dir(strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strFileName
    End If

    ReDim strFieldN(5)
    ReDim intWidth(5)
    intField = 65: intCounter = 1: SaveExcel = False

    strFieldN = Array("智權人員", "客戶編號", "收據抬頭", "未收規費", "未收服務費")
    intWidth = Array(13, 13, 28, 13, 13)
    
    xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAnnuity.Workbooks.add
    Set wksAnnuity = xlsAnnuity.Worksheets(1)
   
    wksAnnuity.Range(Chr(intField) & intCounter).Value = ReportTitle(107)
    intCounter = intCounter + 1
    wksAnnuity.Range(Chr(intField) & intCounter).Value = IIf(Val(FCDate(MaskEdBox1.Text)) > 0, MaskEdBox1.Text, "096/01/01") & _
                                                             "～ " & MaskEdBox2.Text
    intCounter = intCounter + 1
    wksAnnuity.Range(Chr(intField) & intCounter).Value = "列印人員：" & StaffQuery(strUserNum)
    wksAnnuity.Range(Chr(UBound(strFieldN) + intField - 1) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
    intCounter = intCounter + 1
        
    For i = LBound(strFieldN) To UBound(strFieldN)
        wksAnnuity.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
        wksAnnuity.Range(Chr(i + intField) & intCounter).Value = strFieldN(i)
        wksAnnuity.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
    Next i
    intTitleRow = intCounter: intCounter = intCounter + 1: strStart = intCounter
        
    With adoaccrpt107
        .MoveFirst
        Do While .EOF = False
            If InStr("" & .Fields("R10704"), " 小    計 ") > 0 Then
                Call GetTotal(0, wksAnnuity, intCounter, "" & .Fields("R10704"), strStart)
                strSum = strSum & "," & intCounter
            ElseIf InStr("" & .Fields("R10704"), "合計") > 0 Then
                Call GetTotal(1, wksAnnuity, intCounter, "" & .Fields("R10704"), Mid(strSum, 2))
                strSum = "": strTotal = strTotal & "," & intCounter
                intCounter = intCounter + 1
            ElseIf InStr("" & .Fields("R10704"), "總      計") > 0 Then
                Call GetTotal(2, wksAnnuity, intCounter, "" & .Fields("R10704"), Mid(strTotal, 2))
            Else
                For i = LBound(strFieldN) To UBound(strFieldN)
                    wksAnnuity.Range(Chr(i + intField) & intCounter).Value = "" & .Fields(i + 1)
                    If i = GetValue("未收規費") Or i = GetValue("未收服務費") Then
                        wksAnnuity.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "#,##0.00"
                    End If
                Next i
            End If
            intCounter = intCounter + 1
            If InStr("" & .Fields("R10704"), " 小    計 ") > 0 Or InStr("" & .Fields("R10704"), "合計") > 0 Then
                strStart = intCounter
            End If
            .MoveNext
        Loop
    End With
    
    wksAnnuity.PageSetup.PrintTitleRows = "$1:$" & intTitleRow
    
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    adoaccrpt107.Close
    SaveExcel = True
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    Exit Function
    
ErrHnd:
    xlsAnnuity.Visible = True
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    adoaccrpt107.Close
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub GetTotal(ByVal intChoose As Integer, ByRef Wks As Worksheet, ByVal intCounter As Integer, ByVal stRecName As String, ByVal stTot As String)
    Dim intField As Integer, j As Integer
    Dim stData As String, stTmp As String
    Dim arrRow
    
    intField = 65
    
    If intChoose = 0 Then
        For i = GetValue("收據抬頭") To GetValue("未收服務費")
            If i = GetValue("收據抬頭") Then
                stData = stRecName
                Wks.Range(Chr(i + intField) & intCounter).Value = stRecName
            Else
                stData = "=Sum(" & Chr(i + intField) & stTot & ":" & Chr(i + intField) & intCounter - 1 & ")"
                Wks.Range(Chr(i + intField) & intCounter).Formula = stData
            End If
        Next i
    Else
        arrRow = Split(stTot, ",")
        For i = GetValue("收據抬頭") To GetValue("未收服務費")
            If i = GetValue("收據抬頭") Then
                stData = stRecName
                Wks.Range(Chr(i + intField) & intCounter).Value = stRecName
            Else
                stData = ""
                For j = LBound(arrRow) To UBound(arrRow)
                    stData = stData & "," & Chr(i + intField) & arrRow(j)
                Next j
                stData = "=Sum(" & Mid(stData, 2) & ")"
                Wks.Range(Chr(i + intField) & intCounter).Formula = stData
                Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "#,##0.00"
            End If
        Next i
    End If
End Sub
