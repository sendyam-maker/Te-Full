VERSION 5.00
Begin VB.Form Frmacc4420 
   AutoRedraw      =   -1  'True
   Caption         =   "科目餘額表"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   5160
   Begin VB.ComboBox CboComp 
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
      Left            =   1320
      TabIndex        =   0
      Text            =   "CboComp"
      Top             =   240
      Width           =   3520
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
      Left            =   1380
      Style           =   2  '單純下拉式
      TabIndex        =   21
      Top             =   2490
      Width           =   3450
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1920
      TabIndex        =   20
      Top             =   600
      Width           =   2892
   End
   Begin VB.TextBox Text7 
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
      TabIndex        =   1
      Top             =   600
      Width           =   612
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
      TabIndex        =   10
      Top             =   1800
      Width           =   4692
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   3570
      TabIndex        =   9
      Top             =   4140
      Visible         =   0   'False
      Width           =   1212
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
      Left            =   930
      TabIndex        =   8
      Top             =   4140
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
      Left            =   3570
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   1212
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
      Left            =   930
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   1812
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
      Height          =   300
      Left            =   4200
      TabIndex        =   5
      Top             =   1320
      Width           =   612
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
      TabIndex        =   4
      Top             =   1320
      Width           =   612
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
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1572
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
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1572
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
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   210
      Top             =   3300
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2970
      Picture         =   "Frmacc4420.frx":0000
      Stretch         =   -1  'True
      Top             =   4140
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
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
      Left            =   690
      TabIndex        =   19
      Top             =   4140
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2970
      Picture         =   "Frmacc4420.frx":0442
      Stretch         =   -1  'True
      Top             =   3780
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
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
      Left            =   690
      TabIndex        =   18
      Top             =   3780
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
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
      Left            =   330
      TabIndex        =   17
      Top             =   3420
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "月份"
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
      Left            =   3240
      TabIndex        =   16
      Top             =   1320
      Width           =   612
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      TabIndex        =   15
      Top             =   1320
      Width           =   612
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   60
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
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
      TabIndex        =   14
      Top             =   960
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   13
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      TabIndex        =   12
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc4420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoaccrpt403 As New ADODB.Recordset
Dim strSort1, strSort2 As String
Dim dllaccrpt403 As Object
Dim strSql As String, strSQL1 As String, strSQL2 As String
Dim strPrinter As String 'Add By Sindy 2013/6/4

'Add by Amy 2020/04/07
Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboComp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(cboComp) = MsgText(601) Then Exit Sub
    
    strCmp = cboComp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label2 & MsgText(63), , MsgText(5)
        Cancel = True
        cboComp.SetFocus
        Exit Sub
    ElseIf Len(Trim(cboComp)) = 1 Then
        cboComp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/4/07

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo4.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
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
         'Modify by Amy 2020/04/07
         'Text5.SetFocus
         cboComp.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
   Dim strCmpNo As String, strCmpN As String 'Add by Amy 2020/04/08
   
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt403Delete
   ProduceData
   PUB_SetOsDefaultPrinter Combo1  'Add By Sindy 2013/6/4
   If adoaccrpt403.State = adStateOpen Then
      adoaccrpt403.Close
   End If
   adoaccrpt403.CursorLocation = adUseClient
   adoaccrpt403.Open "select * from accrpt403", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt403.RecordCount <> 0 Then
      'Moidfy by Amy 2020/04/14 公司別改下拉 原:Text5, Text6
      strCmpNo = cboComp
      If InStr(strCmpNo, "　") > 0 Then
            strCmpNo = Mid(strCmpNo, 1, Val(InStr(strCmpNo, "　")) - 1)
      End If
      strCmpN = GetAccReportCmpN(strCmpNo, True)
      dllaccrpt403.Acc4420 ReportTitle(403), strCmpNo, strCmpN, Text7, Text8, Text2, Text1, Text3, Text4, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'end 2020/04/14
   End If
   adoaccrpt403.Close
   PUB_SetOsDefaultPrinter strPrinter  'Add By Sindy 2013/6/4
   FormClear
   Screen.MousePointer = vbDefault
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
   Me.Width = 5250
   Me.Height = 3420 '2800
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/4/07
   cboComp.Clear
   cboComp.AddItem "", 0
   Call Pub_SetCboCmp(cboComp, False, False, False, , 1)
   'end 2020/04/07
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt403 = CreateObject("AccReport.ReportSelect")
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

   Set dllaccrpt403 = Nothing
   Set Frmacc4420 = Nothing
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

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'Mark by Amy 2020/04/07 改下拉
'Private Sub Text5_Change()
'   If Text5 = MsgText(601) Then
'      Exit Sub
'   End If
'   Text6 = A0802Query(Text5)
'End Sub
'
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'End Sub
'
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''2014/1/22 add by sonia
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
''2014/1/22 end
'end 2020/04/07

Private Sub Text7_Change()
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   Text8 = A0902Query(Text7)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Public Sub ProduceData()
Dim intYear, intMonth As Integer
Dim lngStartDate, lngEndDate As Long
Dim strOrder1, strOrder2 As String
Dim douDebit As Double, douCredit As Double
Dim strCmp As String 'Add by Amy 2020/04/07

On Error GoTo Checking
   strSql = ""
   strSQL1 = ""
   strSQL2 = ""
   Me.MousePointer = vbHourglass
   If Combo3 = strSort1 Then
      If Combo4 = MsgText(1) Then
         strOrder1 = " order by a0405 asc"
      Else
         strOrder1 = " order by a0405 desc"
      End If
      If Combo5 = strSort2 Then
         If Combo6 = MsgText(1) Then
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
         If Combo5 = strSort1 Then
            If Combo6 = MsgText(1) Then
               strOrder2 = ", a0405 asc"
            Else
               strOrder2 = ", a0405 desc"
            End If
         Else
            strOrder2 = MsgText(601)
         End If
      Else
         strOrder1 = MsgText(601)
         strOrder2 = MsgText(601)
      End If
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoacc010.CursorLocation = adUseClient
   'Modify by Amy 2020/04/07 改下拉 原:Text5
   If Trim(cboComp) <> MsgText(601) Then
      strCmp = cboComp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      strSQL1 = " and a0403 = '" & strCmp & "'"
      strSQL2 = " and ax201 = '" & strCmp & "'"
   End If
   'end 2020/04/07
   If Text7 <> MsgText(601) Then
      'Modify by Amy a0407=  '" & Text7 & "'" 改a0404,查錯欄位
      strSQL1 = strSQL1 & " and a0404 = '" & Text7 & "'"
      strSQL2 = strSQL2 & " and ax204 = '" & Text7 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0101 >= '" & Text2 & "'"
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0101 <= '" & Text1 & "'"
   End If
'   If strSQL <> MsgText(601) Then
'      strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 4)
'   End If
   If Val(Text3) <> 0 Then
      strSQL1 = strSQL1 & " and a0401 = " & Val(Text3) & ""
   End If
   If Val(Text4) <> 0 Then
      strSQL1 = strSQL1 & " and a0402 = " & Val(Text4) & ""
   End If
   If Len(Text4) > 1 Then
      strSQL2 = strSQL2 & " and a0205 >= " & Val(Text3 & Text4 & "01") & ""
      strSQL2 = strSQL2 & " and a0205 <= " & Val(Text3 & Text4 & "31") & ""
   Else
      strSQL2 = strSQL2 & " and a0205 >= " & Val(Text3 & "0" & Text4 & "01") & ""
      strSQL2 = strSQL2 & " and a0205 <= " & Val(Text3 & "0" & Text4 & "31") & ""
   End If
   If strSQL1 <> MsgText(601) Then
      strSQL1 = " where " & Mid(strSQL1, 5, Len(strSQL1) - 4)
   End If
   Accrpt403Delete
   adoaccrpt403.CursorLocation = adUseClient
   adoaccrpt403.Open "select * from accrpt403", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   adoacc010.Open "select * from acc010" & strSQL & "  order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc010.Open "select * from acc010, (select a0405, sum(a0408) as Balance from acc040 " & strSQL1 & " group by a0405) new1, (select ax205, sum(ax206) as Debit, sum(ax207) as Credit from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and (ax210 is null)" & strSQL2 & " group by ax205) new2 where a0101 = a0405 (+) and a0101 = ax205 (+)" & strSql & "  order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount = 0 Then
      adoacc010.Close
      adoaccrpt403.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc010.EOF = False
      adoaccrpt403.AddNew
      adoaccrpt403.Fields("r40301").Value = strUserNum
      adoaccrpt403.Fields("r40302").Value = adoacc010.Fields(0).Value
      If IsNull(adoacc010.Fields(1).Value) Then
         adoaccrpt403.Fields("r40303").Value = Null
      Else
         adoaccrpt403.Fields("r40303").Value = adoacc010.Fields(1).Value
      End If
'      adoacc040.CursorLocation = adUseClient
'      adoacc040.Open "select sum(a0408) from acc040 where a0405 = '" & adoacc010.Fields(0).Value & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc040.RecordCount <> 0 Then
'         If IsNull(adoacc040.Fields(0).Value) Then
'            adoaccrpt403.Fields("r40304").Value = 0
'         Else
'            adoaccrpt403.Fields("r40304").Value = adoacc040.Fields(0).Value
'         End If
'      Else
'         adoaccrpt403.Fields("r040304").Value = 0
'      End If
      If IsNull(adoacc010.Fields("Balance").Value) Then
         adoaccrpt403.Fields("r40304").Value = 0
      Else
         adoaccrpt403.Fields("r40304").Value = adoacc010.Fields("Balance").Value
      End If
'      adoacc040.Close
'      adoacc040.CursorLocation = adUseClient
'      adoacc040.Open "select sum(ax206), sum(ax207) from acc021, acc020 where ax205 = '" & adoacc010.Fields(0).Value & "' and (ax210 is null or ax210 = 0)" & strSQL2, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc040.RecordCount <> 0 Then
'         If IsNull(adoacc040.Fields(0).Value) Then
'            douDebit = 0
'         Else
'            douDebit = adoacc040.Fields(0).Value
'         End If
'         If IsNull(adoacc040.Fields(1).Value) Then
'            douCredit = 0
'         Else
'            douCredit = adoacc040.Fields(1).Value
'         End If
'         If adoacc010.Fields("a0103").Value = "1" Then
'            adoaccrpt403.Fields("r40304").Value = Val(adoaccrpt403.Fields("r40304").Value) + douDebit - douCredit
'         Else
'            adoaccrpt403.Fields("r40304").Value = Val(adoaccrpt403.Fields("r40304").Value) + douCredit - douDebit
'         End If
'      End If
'      adoacc040.Close
      If IsNull(adoacc010.Fields("Debit").Value) Then
         douDebit = 0
      Else
         douDebit = adoacc010.Fields("Debit").Value
      End If
      If IsNull(adoacc010.Fields("Credit").Value) Then
         douCredit = 0
      Else
         douCredit = adoacc010.Fields("Credit").Value
      End If
      If adoacc010.Fields("a0103").Value = "1" Then
         adoaccrpt403.Fields("r40304").Value = Val(adoaccrpt403.Fields("r40304").Value) + douDebit - douCredit
      Else
         adoaccrpt403.Fields("r40304").Value = Val(adoaccrpt403.Fields("r40304").Value) + douCredit - douDebit
      End If
      adoaccrpt403.UpdateBatch
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt403.Close
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
Private Sub Accrpt403Delete()
   adoTaie.Execute "delete from accrpt403"
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "科目代號"
   strSort2 = "科目名稱"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/07 公司別改下拉
'   Text5 = ""
'   Text6 = ""
   cboComp = ""
   'end 2020/04/07
   Text7 = ""
   Text8 = ""
   Text2 = ""
   Text1 = ""
   Text3 = ""
   Text4 = ""
   Combo3 = ""
   Combo5 = ""
   'Modify by Amy 2020/04/07
   'Text5.SetFocus
   cboComp.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Amy 2020/04/07
   Dim bolCancel As Boolean
   
   If Trim(cboComp) <> MsgText(601) Then
      CboComp_Validate (bolCancel)
      If bolCancel = False Then
        FormCheck = True
        Exit Function
      End If
   End If
   'end 2020/04/07
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

