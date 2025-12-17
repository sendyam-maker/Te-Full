VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc4410 
   AutoRedraw      =   -1  'True
   Caption         =   "日計表"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
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
      Left            =   1350
      TabIndex        =   0
      Top             =   210
      Width           =   3500
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   12
      Top             =   3060
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3630
      TabIndex        =   11
      Top             =   3060
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   10
      Top             =   3420
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3630
      TabIndex        =   9
      Top             =   3420
      Visible         =   0   'False
      Width           =   1212
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
      Left            =   1350
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   1830
      Width           =   3450
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
      TabIndex        =   3
      Top             =   1080
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   600
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
      TabIndex        =   2
      Top             =   600
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
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      Left            =   390
      TabIndex        =   15
      Top             =   2700
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "1."
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
      Left            =   750
      TabIndex        =   14
      Top             =   3060
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3030
      Picture         =   "Frmacc4410.frx":0000
      Stretch         =   -1  'True
      Top             =   3060
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "2."
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
      Left            =   750
      TabIndex        =   13
      Top             =   3420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3030
      Picture         =   "Frmacc4410.frx":0442
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   375
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
      Left            =   330
      TabIndex        =   8
      Top             =   1860
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   90
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   270
      Top             =   2550
      Visible         =   0   'False
      Width           =   4695
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
      Height          =   252
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Top             =   600
      Width           =   972
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
      TabIndex        =   4
      Top             =   210
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc4410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc021 As New ADODB.Recordset
Public adoaccrpt401 As New ADODB.Recordset
Dim strSort1, strSort2 As String
Dim dllaccrpt401 As Object
Dim strPrinter As String 'Add By Sindy 2013/6/4
'Added by Lydia 2019/10/21 改成Printer輸出, 列印用
Dim mPrtOrt As Integer  '原本預設印表機的列印方向
Private Const ciTitleFontSize = 14, cInX = 5
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 150
Dim ciFontSize As Integer '報表內容字型大小
Dim mRptTitle As String '報表抬頭
Dim strTitle As String, strTitle2 As String '欄位抬頭/起始位置
Dim PLeft(0 To cInX) As Integer '欄位起始位置陣列
Dim PTitle(0 To cInX) As String   '欄位抬頭陣列
Dim iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPageLine As Integer '頁面資料列
Dim strAcAmt As String, strAcTot As String '應收金額(by日),合計
Dim strDCamt As String, strDctot  As String '貸方金額(by日),合計

'Add by Amy 2020/04/14
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
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp, True)
    End If
End Sub
'end 2020/04/14

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo5.SetFocus
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo4.SetFocus
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo2.SetFocus
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command1.SetFocus
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
   Dim HasShowMsg As Boolean 'Add by Amy 2020/04/14
   
   'Modify by Amy 2020/04/14
   If FormCheck(HasShowMsg) = False Then
      If HasShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   'end 2020/04/14
   'Modified by Lydia 2019/10/21 改成Printer輸出
'   Screen.MousePointer = vbHourglass
'   Accrpt401Delete
'   ProduceData
'   PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2013/6/4
'   If adoaccrpt401.State = adStateOpen Then
'      adoaccrpt401.Close
'   End If
'   adoaccrpt401.CursorLocation = adUseClient
'   adoaccrpt401.Open "select * from accrpt401", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt401.RecordCount <> 0 Then
'      '2014/1/22 modify by sonia
'      'dllaccrpt401.Acc4410 ReportTitle(401), Text5, Text6, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      dllaccrpt401.Acc4410 ReportTitle(401), IIf(Text5 = "2", "J", Text5), IIf(Text5 = "", "台一　專利商標/智權", Text6), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      '2014/1/22 end
'   End If
'   adoaccrpt401.Close
'   PUB_SetOsDefaultPrinter strPrinter 'Add By Sindy 2013/6/4
'   FormClear
'   Screen.MousePointer = vbDefault
   PUB_RestorePrinter Combo1.Text
   mPrtOrt = Printer.Orientation
   Screen.MousePointer = vbHourglass
   Accrpt401Delete
   ProduceData
   Call PrintRpt4410
   Screen.MousePointer = vbDefault
   FormClear
   Printer.Orientation = mPrtOrt
   PUB_RestorePrinter strPrinter
   'end 2019/10/21
   
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
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5280 '5250
   Me.Height = 2775 '2100
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/14 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, True, False, , 1)
   'end 2020/04/14
   ComboAdd
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo5.AddItem MsgText(1)
   Combo5.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo5 = MsgText(1)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   'Remove by Lydia 2019/10/21 改成Printer輸出
   'Set dllaccrpt401 = CreateObject("AccReport.ReportSelect")
   
   'Add By Cheng 2003/05/27
   '預設公司別
'2014/1/22 cancel by sonia
'   Me.Text5.Text = "1"
'   Text6 = A0802Query(Text5)
'2014/1/22 end
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
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
   
   'Set dllaccrpt401 = Nothing 'Remove by Lydia 2019/10/21 改成Printer輸出
   Set Frmacc4410 = Nothing
End Sub

'Mark by Amy 2020/04/14 公司別改下拉
'Private Sub Text5_Change()
'   '2014/1/22 modify by sonia
'   'Text6 = A0802Query(Text5)
'   Select Case Text5
'      Case "1"
'         Text6 = A0802Query(Text5)
'      Case "2"
'         Text6 = A0802Query("J")
'   End Select
'   '2014/1/22 end
'End Sub
'
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'End Sub
'
''2014/1/22 add by sonia
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/1/22 end

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "科目代號"
   strSort2 = "科目名稱"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo2.AddItem strSort1
   Combo2.AddItem strSort2
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2, strSql As String
Dim strCmp As String 'Add by Amy 2020/04/14

On Error GoTo Checking
   Me.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt401.CursorLocation = adUseClient
   'Modified by Lydia 2019/10/18 +操作者
   'adoaccrpt401.Open "select * from accrpt401", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoaccrpt401.Open "select * from accrpt401 where r40101=" & CNULL(strUserNum), adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   If Combo3 = strSort1 Then
      If Combo4 = MsgText(1) Then
         strOrder1 = " order by ax205 asc"
      Else
         strOrder1 = " order by ax205 desc"
      End If
      If Combo2 = strSort2 Then
         If Combo5 = MsgText(1) Then
            strOrder2 = ", a0102 asc"
         Else
            strOrder2 = ", a0102 desc"
         End If
      Else
         strOrder2 = MsgText(601)
      End If
   Else
      If Combo3 = strSort2 Then
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0102 asc"
         Else
            strOrder1 = " order by a0102 desc"
         End If
         If Combo2 = strSort1 Then
            If Combo5 = MsgText(1) Then
               strOrder2 = ", ax205 asc"
            Else
               strOrder2 = ", ax205 desc"
            End If
         Else
            strOrder2 = MsgText(601)
         End If
      Else
         strOrder1 = MsgText(601)
         strOrder2 = MsgText(601)
      End If
   End If
   adoacc021.CursorLocation = adUseClient
   'Modify By Cheng 2003/05/27
   '恢復公司別的限制
   'Modify by Amy 2020/04/14 改下拉 原:Text5
   If Trim(CboCmp) <> MsgText(601) Then
      strCmp = CboCmp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      If InStr(strCmp, "+") = 0 Then
            '2014/1/22 modify by sonia
            'strSql = " and ax201 = '" & Text5 & "'"
            strSql = " and ax201 = '" & strCmp & "'"
            '2014/1/22 end
       Else
            strSql = " and ax201 In ('" & Replace(strCmp, "+", "','") & "')"
       End If
   End If
   'end 2020/04/14
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   adoacc021.Open "select a0205, ax205, a0102, sum(ax206), sum(ax207) from acc021, acc020, acc010 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and acc021.ax205 = acc010.a0101 " & _
                  "" & strSql & " group by a0205, ax205, a0102" & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc021.RecordCount = 0 Then
      adoacc021.Close
      adoaccrpt401.Close
      MsgBox MsgText(28), , MsgText(5)
        Me.MousePointer = vbDefault
      Exit Sub
   End If
   Do While adoacc021.EOF = False
      adoaccrpt401.AddNew
      adoaccrpt401.Fields("r40101").Value = strUserNum
      If IsNull(adoacc021.Fields(0).Value) Then
         adoaccrpt401.Fields("r40102").Value = Null
      Else
         adoaccrpt401.Fields("r40102").Value = Val(adoacc021.Fields(0).Value)
      End If
      If IsNull(adoacc021.Fields(1).Value) Then
         adoaccrpt401.Fields("r40103").Value = Null
      Else
         adoaccrpt401.Fields("r40103").Value = adoacc021.Fields(1).Value
      End If
      If IsNull(adoacc021.Fields(2).Value) Then
         adoaccrpt401.Fields("r40104").Value = Null
      Else
         adoaccrpt401.Fields("r40104").Value = adoacc021.Fields(2).Value
      End If
      If IsNull(adoacc021.Fields(3).Value) Then
         adoaccrpt401.Fields("r40105").Value = 0
      Else
         adoaccrpt401.Fields("r40105").Value = adoacc021.Fields(3).Value
      End If
      If IsNull(adoacc021.Fields(4).Value) Then
         adoaccrpt401.Fields("r40106").Value = 0
      Else
         adoaccrpt401.Fields("r40106").Value = adoacc021.Fields(4).Value
      End If
      adoaccrpt401.UpdateBatch
      adoacc021.MoveNext
   Loop
   adoacc021.Close
   adoaccrpt401.Close
   Me.MousePointer = vbDefault
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
Private Sub Accrpt401Delete()
   'Modified by Lydia 2019/10/18 +操作者
   'adoTaie.Execute "delete from accrpt401"
   adoTaie.Execute "delete from accrpt401 where r40101=" & CNULL(strUserNum)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/14 公司別改下拉
'   Text5 = ""
'   Text6 = "台一　專利商標/智權"  '2014/1/22 modify by sonia
   CboCmp = ""
   'end 2020/04/14
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Combo3 = ""
   Combo2 = ""
   CboCmp.SetFocus 'Modify by Amy 2020/04/14 公司別改下拉
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(HasShowMsg As Boolean) As Boolean
   'Add by Amy 2020/04/14
   Dim bCancel As Boolean
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bCancel)
      If bCancel = True Then
        HasShowMsg = True
        Exit Function
      End If
   End If
   'end 2020/04/14
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

'Added by Lydia 2019/10/21 列印-日記表
Private Sub PrintRpt4410()
Dim inP As Integer
Dim rsPrt As New ADODB.Recordset
Dim strGrp As String '傳票日期
Dim strTmp As String
    
    mRptTitle = ReportTitle(401)
    '改A4橫印
    strTitle = "傳票日期,科目代號,科目名稱,借方金額,貸方金額,結束"
    strTitle2 = "0,5,5,15,7,7"
    ciFontSize = 12
    strAcAmt = "0": strAcTot = "0"
    strDCamt = "0": strDctot = "0"
    
    '抓暫存檔資料
    strSql = "select * from accrpt401 where r40101='" & strUserNum & "' "
    strSql = strSql & "order by r40101, r40102, r40103  "
    inP = 1
    Set rsPrt = ClsLawReadRstMsg(inP, strSql)
    If inP = 1 Then
       If lngLineHeight = 0 Then SettingPrtSet  '設定印表機
       
       With rsPrt
          .MoveFirst
          iPage = iPage + 1
          PrintHeader
          Printer.Font.Size = ciFontSize
          Printer.FontBold = False
          Do While Not .EOF
            If strGrp <> "" And strGrp <> "" & .Fields("r40102") Then '傳票日期
                 Call PrintSubTotal(1)
                 strAcAmt = "0"
                 strDCamt = "0"
            End If
            strGrp = "" & .Fields("r40102")
            strAcAmt = Val(strAcAmt) + Val("" & .Fields("r40105"))
            strAcTot = Val(strAcTot) + Val("" & .Fields("r40105"))
            strDCamt = Val(strDCamt) + Val("" & .Fields("r40106"))
            strDctot = Val(strDctot) + Val("" & .Fields("r40106"))

            '列印內容
            For inP = 0 To cInX
               If PTitle(inP) = "" Then Exit For
               
               If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
                  Printer.CurrentX = PLeft(inP)
                  Printer.CurrentY = iPrint
                  
                  Select Case inP
                      Case 0 '傳票日期
                           Printer.Print ChangeTStringToTDateString("" & .Fields("r40102"))
                           
                      Case 1 '科目代號
                           Printer.Print "" & .Fields("r40103")
                           
                      Case 2 '科目名稱
                           Printer.Print "" & PUB_StrToStr("" & .Fields("r40104"), 30)
                           
                      Case 3, 4 '借方金額/貸方金額
                           strTmp = Format(Val(IIf(inP = 3, "" & .Fields("r40105"), "" & .Fields("r40106"))), "##,##0.00")
                           Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                           Printer.CurrentY = iPrint
                           Printer.Print strTmp
                           
                  End Select
               End If 'If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
            Next 'For inP = 1 To cInX

            PrintNewLine
JumpPrint:
             .MoveNext
          Loop

       Call PrintSubTotal(1)  '小計by日
       Call PrintSubTotal(2) '合計
       End With
       
       Printer.EndDoc
       ShowPrintOk
       Set RsTemp = Nothing
    Else
        MsgBox MsgText(28), , MsgText(5)
    End If
    Set rsPrt = Nothing
End Sub

'Added by Lydia 2019/10/21 列印-設定印表機
Private Sub SettingPrtSet()
Dim inX As Integer
Dim tmpArr As Variant, tmpArr2 As Variant

    '設定印表機
     Printer.EndDoc
     Printer.PaperSize = 9  'A4
     Printer.Orientation = 1 '1.直印
     
     lngPageHeight = Printer.ScaleHeight
     lngPageWidth = Printer.ScaleWidth
     lngLineHeight = 300
     Printer.Font.Name = "新細明體"
     Printer.Font.Size = ciFontSize
     Erase PLeft
     Erase PTitle
     tmpArr = Empty: tmpArr2 = Empty
     
     '設定欄位抬頭和位置
     If strTitle <> "" And strTitle2 <> "" Then
        tmpArr = Split(strTitle, ",")
        tmpArr2 = Split(strTitle2, ",")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" And Trim(tmpArr2(inX)) <> "" Then
                If Trim(tmpArr(inX)) <> "結束" Then PTitle(inX) = Trim(tmpArr(inX))
                
                If inX < 1 Then
                   PLeft(inX) = ciStartX
                Else
                   PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(Val(tmpArr2(inX)), "　")) + ciColGap
                End If
                
                If Trim(tmpArr(inX)) = "結束" Then Exit For
            End If
        Next
     End If
     
     iPage = 0
End Sub

'Added by Lydia 2019/10/21 列印-換行判斷
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
End Sub

'Added by Lydia 2019/10/21 列印-表頭
Private Sub PrintHeader()
Dim x1 As Integer, x2 As Integer, iPos As Integer
Dim strTmp As String
'Add by Amy 2020/04/14
Dim strCmp As String, ii As Integer
Dim arrCmp

iPrint = ciStartY
iPageLine = 0

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(mRptTitle)) / 2
Printer.CurrentY = iPrint
Printer.Print mRptTitle

Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
PrintNewLine
PrintNewLine

'Modify by Amy 2020/04/16 公司別改下拉
'strTmp = "公司別: 　　" & IIf(Text5 = "2", "J", Text5) & "　" & IIf(Text5 = "", "台一　專利商標/智權", Text6)
strCmp = CboCmp
If InStr(strCmp, "　") > 0 Then
    strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
End If
strCmp = GetAccReportCmpN(strCmp, True, True)
strTmp = "公司別: 　　" & strCmp
'end 2020/04/16
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print strTmp

PrintNewLine

'對齊-公司別
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print "傳票日期: 　" & MaskEdBox1.Text & " - " & MaskEdBox2.Text

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName
x1 = Printer.ScaleWidth - Printer.TextWidth(String(12, "　"))
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
PrintNewLine
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page
PrintNewLine

'列印欄位抬頭
For iPos = 0 To cInX
    If PTitle(iPos) <> "" And PTitle(iPos) <> "結束" Then
       If PTitle(iPos) = "科目名稱" Or InStr(PTitle(iPos), "金額") > 0 Then  '置中
          x2 = PLeft(iPos) + (PLeft(iPos + 1) - PLeft(iPos) - Printer.TextWidth(PTitle(iPos))) / 2
       Else
          x2 = PLeft(iPos)
       End If
       Printer.CurrentX = x2 'PLeft(iPos)
       Printer.CurrentY = iPrint
       Printer.Print PTitle(iPos)
    ElseIf iPos > 1 Then
        x1 = iPos '結束
        Exit For
    End If
Next
PrintNewLine
Printer.Line (PLeft(0), iPrint)-(PLeft(x1), iPrint)
iPrint = iPrint + 150
End Sub

'Added by Lydia 2019/10/21 列印-小計或合計
Private Sub PrintSubTotal(ByVal iKind As String)
'iKind = 1(小計) ; 2(合計)
Dim strTmp As String

    Printer.Line (PLeft(3), iPrint)-(PLeft(5), iPrint)
    iPrint = iPrint + 150
    Printer.CurrentX = PLeft(2) + Printer.TextWidth(String(8, "A"))
    Printer.CurrentY = iPrint
    Printer.Print IIf(iKind = "1", "日期小計:", "合計:")
    '借方金額
    strTmp = IIf(iKind = "1", strAcAmt, strAcTot)
    strTmp = Format(Val(strTmp), "##,##0.00")
    Printer.CurrentX = PLeft(4) - Printer.TextWidth(strTmp) - ciColGap '靠右
    Printer.CurrentY = iPrint
    Printer.Print strTmp
    '貸方金額
    strTmp = IIf(iKind = "1", strDCamt, strDctot)
    strTmp = Format(Val(strTmp), "##,##0.00")
    Printer.CurrentX = PLeft(5) - Printer.TextWidth(strTmp) - ciColGap '靠右
    Printer.CurrentY = iPrint
    Printer.Print strTmp
    
    PrintNewLine
    If iKind = "2" Then
        Printer.Line (PLeft(3), iPrint + 60)-(PLeft(5), iPrint + 60)
        Printer.Line (PLeft(3), iPrint + 120)-(PLeft(5), iPrint + 120)
        iPrint = iPrint + 200
        strTmp = "*** 結束 ***"
        Printer.Font.Bold = True
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
        Printer.CurrentY = iPrint
        Printer.Print strTmp
    End If
End Sub


