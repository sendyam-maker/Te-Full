VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc3470 
   AutoRedraw      =   -1  'True
   Caption         =   "往來對象別票據彙總表"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2385
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
      Left            =   1320
      TabIndex        =   0
      Top             =   30
      Width           =   3525
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
      Left            =   960
      TabIndex        =   6
      Top             =   2835
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
      Left            =   3600
      TabIndex        =   7
      Top             =   2835
      Visible         =   0   'False
      Width           =   1212
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
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   1580
      Width           =   4692
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
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   800
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
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   800
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
      TabIndex        =   1
      Top             =   440
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1155
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
      Left            =   3240
      TabIndex        =   5
      Top             =   1155
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
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   105
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "收/開日期"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1155
      Width           =   1095
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
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   1155
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   135
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
      Left            =   360
      TabIndex        =   14
      Top             =   2580
      Visible         =   0   'False
      Width           =   975
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
      Left            =   720
      TabIndex        =   13
      Top             =   2835
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3470.frx":0000
      Stretch         =   -1  'True
      Top             =   2835
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   870
      Left            =   240
      Top             =   2505
      Visible         =   0   'False
      Width           =   4695
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
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   795
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      Left            =   240
      TabIndex        =   10
      Top             =   795
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "(1.客戶 2.廠商 3.員工)"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   435
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      Left            =   240
      TabIndex        =   8
      Top             =   435
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc3470"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt308 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSort1, strSort2 As String
Dim dllaccrpt308 As Object
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, False, True)
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
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label9 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

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
         Text1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
Dim strSelect As String
Dim bCancel As Boolean
   
   'Add By Sindy 2020/4/23
   If CboCmp.Text = MsgText(601) Then
      MsgBox Label9 & MsgText(52), , MsgText(5)
      Exit Sub
   End If
   Call CboCmp_Validate(bCancel)
   If bCancel = True Then
      Exit Sub
   End If
   '2020/4/23 END
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   Call SetCompN 'Add by Sindy 2020/04/23
   
   Screen.MousePointer = vbHourglass
   Accrpt308Delete
   ProduceData
   Select Case Text1
      Case Mid(ComboItem(131), 1, 1)
         strSelect = Mid(ComboItem(131), 4, 2)
      Case Mid(ComboItem(132), 1, 1)
         strSelect = Mid(ComboItem(132), 4, 2)
      Case Mid(ComboItem(133), 1, 1)
         strSelect = Mid(ComboItem(133), 4, 2)
      Case Else
         strSelect = MsgText(601)
   End Select
   If adoaccrpt308.State = adStateOpen Then
      adoaccrpt308.Close
   End If
   adoaccrpt308.CursorLocation = adUseClient
   adoaccrpt308.Open "select * from accrpt308", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt308.RecordCount <> 0 Then
      '20140120START Modify By eric
      dllaccrpt308.Acc3470 ReportTitle(308) & "-" & strCmpN, strSelect, Text2, Text3, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'dllaccrpt308.Acc3470 ReportTitle(308), strSelect, Text2, Text3, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      '20140120END
   End If
   adoaccrpt308.Close
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
   'Modify by Amy 2023/10/12 原W5250 H2500
   Me.Width = 5280
   Me.Height = 2850
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
   Combo4 = MsgText(1)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt308 = CreateObject("AccReport.ReportSelect")
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, False, False, False, , 1)
   'end 2020/04/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt308 = Nothing
   Set Frmacc3470 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text1 = Mid(ComboItem(131), 1, 1) Then
      If Len(Text2) = 6 Then
         Text2 = AfterZero(Text2)
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "往來對象"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2 As String
Dim strSql As String
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e06 asc"
         Else
            strOrder1 = " order by a0e06 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   
   '20140120START MODIFY By eric
   'Modify By Sindy 2020/4/23
   'strSql = " and a0e23 = '" & IIf(Text4 = "2", "J", "1") & "' " & strSql
   strSql = " and a0e23 = '" & strCmp & "' " & strSql
   '2020/4/23 END
   
   If Text1 <> MsgText(601) Then
     strSql = strSql & " and a0e05 = '" & Text1 & "'"
   End If
   'If Text1 <> MsgText(601) Then
   '  strSql = " and a0e05 = '" & Text1 & "'"
   'End If
   '20140120END
   
   
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0e06 >= '" & Text2 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0e06 <= '" & Text3 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   
   adoaccrpt308.CursorLocation = adUseClient
   adoaccrpt308.Open "select * from accrpt308", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select a0e05, a0e06 from acc0e0 where a0e25 = 0" & strSql & " group by a0e05, a0e06" & strOrder1, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      adoaccrpt308.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0e0.EOF = False
      adoaccrpt308.AddNew
      adoaccrpt308.Fields("r30801").Value = strUserNum
      If IsNull(adoacc0e0.Fields("a0e05").Value) Then
         adoaccrpt308.Fields("r30802").Value = Null
      Else
         adoaccrpt308.Fields("r30802").Value = adoacc0e0.Fields("a0e05").Value
         Select Case adoacc0e0.Fields("a0e05").Value
            Case Mid(ComboItem(131), 1, 1)
               If IsNull(adoacc0e0.Fields("a0e06").Value) Then
                  adoaccrpt308.Fields("r30804").Value = Null
               Else
                  adoaccrpt308.Fields("r30804").Value = CustomerQuery(adoacc0e0.Fields("a0e06").Value, 1)
               End If
            Case Mid(ComboItem(132), 1, 1)
               adoaccrpt308.Fields("r30804").Value = A0i02Query(adoacc0e0.Fields("a0e06").Value)
            Case Mid(ComboItem(133), 1, 1)
               adoaccrpt308.Fields("r30804").Value = StaffQuery(adoacc0e0.Fields("a0e06").Value)
            Case Else
               adoaccrpt308.Fields("r30804").Value = Null
         End Select
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         adoaccrpt308.Fields("r30803").Value = Null
      Else
         adoaccrpt308.Fields("r30803").Value = adoacc0e0.Fields("a0e06").Value
      End If
      adoaccsum.CursorLocation = adUseClient
      adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e05 = '" & Text1 & "' and a0e06 = '" & adoacc0e0.Fields("a0e06").Value & "' and a0e04 = '" & MsgText(18) & "' and a0e25 = 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt308.Fields("r30805").Value = 0
         Else
            adoaccrpt308.Fields("r30805").Value = Val(adoaccsum.Fields(0).Value)
         End If
      Else
         adoaccrpt308.Fields("r30805").Value = 0
      End If
      adoaccsum.Close
      adoaccsum.CursorLocation = adUseClient
      adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e05 = '" & Text1 & "' and a0e06 = '" & adoacc0e0.Fields("a0e06").Value & "' and a0e04 = '" & MsgText(19) & "' and a0e25 = 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt308.Fields("r30806").Value = 0
         Else
            adoaccrpt308.Fields("r30806").Value = Val(adoaccsum.Fields(0).Value)
         End If
      Else
         adoaccrpt308.Fields("r30806").Value = 0
      End If
      adoaccsum.Close
      adoaccrpt308.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt308.Close
'   adoTaie.Execute "delete from accrpt308 where r30802 is null"
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
Private Sub Accrpt308Delete()
   adoTaie.Execute "delete from accrpt308"
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text1 = Mid(ComboItem(131), 1, 1) Then
      If Len(Text3) = 6 Then
         Text3 = AfterZero(Text3)
      End If
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Combo3 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   '20140120START Modify By eric
'   Label10 = ""
'   Text4 = ""
'   Text4.SetFocus
   CboCmp.ListIndex = -1 'Add By Sindy 2020/4/23
   'Text1.SetFocus
   '20140120END
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
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
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

'Mark by Sindy 2020/4/23 公司別改下拉式選單
''20140120START ADD By eric
'Private Sub Text4_LostFocus()
'   If Text4.Text = "" Then
'      MsgBox "公司別不可空白 !"
'      Text4.SetFocus
'      Exit Sub
'   End If
'   If Text4.Text <> "1" And Text4.Text <> "2" Then
'      MsgBox "公司別僅能為 1 或 2 !"
'      Text4.Text = ""
'      Text4.SetFocus
'      Exit Sub
'   End If
'End Sub
''20140120END
'
''20140120START ADD By eric
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'   CloseIme
'End Sub
''20140120END
'
''20140120START ADD By eric
'Private Sub Text4_Change()
'  Label10.Caption = A0802Query(IIf(Text4 = "2", "J", "1"))
'End Sub
''20140120END
