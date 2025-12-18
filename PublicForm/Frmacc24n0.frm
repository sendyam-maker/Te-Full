VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24n0 
   AutoRedraw      =   -1  'True
   Caption         =   "FC專業收款明細表"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   6495
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   0
      Top             =   60
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   5265
      ScaleHeight     =   435
      ScaleWidth      =   630
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1125
      Width           =   612
   End
   Begin VB.CommandButton CmdPrint 
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
      Left            =   990
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1530
      Width           =   4692
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   5
      Left            =   4170
      TabIndex        =   6
      Top             =   405
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   4
      Left            =   3570
      TabIndex        =   5
      Top             =   405
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   3
      Left            =   2970
      TabIndex        =   4
      Top             =   405
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   2
      Left            =   2370
      TabIndex        =   3
      Top             =   405
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   1
      Left            =   1770
      TabIndex        =   2
      Top             =   405
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   1170
      TabIndex        =   1
      Top             =   405
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1170
      TabIndex        =   7
      Top             =   765
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
      Left            =   3090
      TabIndex        =   8
      Top             =   765
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.專利 2.商標 3.FCT爭議)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1890
      TabIndex        =   19
      Top             =   90
      Width           =   2790
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "(空白=全部)"
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
      Left            =   4890
      TabIndex        =   17
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.明細 2.合計)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1890
      TabIndex        =   16
      Top             =   1155
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "報表內容："
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
      Index           =   4
      Left            =   180
      TabIndex        =   15
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "專業單位"
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
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   60
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   13
      Top             =   765
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳款日期："
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
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
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
      Index           =   1
      Left            =   180
      TabIndex        =   11
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "Frmacc24n0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2021/03/29
Option Explicit

Dim adoAcc24N0 As New ADODB.Recordset
Dim oText As TextBox, m_bPrinter As Boolean
Dim i As Integer, intField As Integer, intRow As Integer, intTitleR As Integer
Dim intWidth1, intWidth2, strField1, strField2
Dim stST15 As String, strSql As String, strTitle As String, strFileName As String
Dim strDateS As String, strDateE As String, strDeptS As String, strDeptE As String, strSystem As String

Private Sub cmdPrint_Click()
    Dim hLocalFile As Long
    Dim strField As String
    Dim intChoose As Integer 'Add by Amy 2021/08/06
    
On Error GoTo ErrHnd
    If FormCheck = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    strSql = "": strDateS = "": strDateE = "": strDeptS = "": strDeptE = "": strSystem = ""
    intChoose = 0 'Add by Amy 2021/08/06
    If Text2 = "3" Then intChoose = 9 'Add by Amy 2021/08/06 FCT爭議
    
    '專業單位
    If Text2 <> MsgText(601) Then
        Select Case Text2
            Case "1"
                strDeptS = "F20"
                strDeptE = "F29"
            'Modify by Amy 2021/08/06 +3.FCT爭議
            Case "2", "3"
                strDeptS = "F10"
                strDeptE = "F19"
        End Select
    End If
    
    '帳款日期
    strDateS = FCDate(MaskEdBox1)
    strDateE = FCDate(MaskEdBox2)
    '系統別
    For Each oText In Text1
        If Trim(oText) <> MsgText(601) Then
            strSystem = strSystem & "," & oText
        End If
    Next
    If Left(strSystem, 1) = "," Then
        strSystem = Mid(strSystem, 2)
    End If
    
    'Memo 2021/06/15 發現11002 商標 F4107 1500未歸至日文組,F4105/F4107 歸至日文部/組,其他st16=null者歸至其他(於Pub_GetAccRecePayAmt修改)
    strSql = Pub_GetAccRecePayAmt(Me.Name, strDateS, strDateE, strDeptS, strDeptE, , strSystem, , True, intChoose)
    'Modify by Amy 2021/06/16 專利點數不會分給智權,故抓 分配人員或ax209(R018),商標無此限制
    '專利
    If Text2 = "1" Then
        strField = "a1n04 "
    '商標
    Else
        strField = "Nvl(a1n04,cp13) "
    End If
    
    '合計
    If Text11 = "2" Then
        strSql = "Select st02,Sum(ReceVal) ReceVal,Sum(ProVal) ProVal,Sum(Nvl(AccPoint,0)) as AccPoint," & strField & " as a1n04,StaffGroup,Nvl(Sst70,Decode(a1n04,'F4107','5','9')) as Sst70 " & _
                     "From (" & strSql & "),Staff Where " & strField & "=st01(+) " & _
                     "Group by StaffGroup,Nvl(Sst70,Decode(a1n04,'F4107','5','9'))," & strField & ",st02 Order by StaffGroup,Sst70,a1n04"
    '明細
    Else
        strSql = "Select  st02," & strField & " as a1n04,a1k13,a1k14,a1k15,a1k16,a1k01,sqldatet(DDate) as DDate,Sum(ReceVal) as ReceVal,Sum(ProVal) as ProVal,Sum(Nvl(AccPoint,0)) as AccPoint " & _
                     ",StaffGroup,Nvl(Sst70,Decode(a1n04,'F4107','5','9')) as Sst70 From (" & strSql & ") a,Staff " & _
                     "Where " & strField & "=st01(+) Group by st02," & strField & ",a1k13,a1k14,a1k15,a1k16,a1k01,sqldatet(DDate) ,StaffGroup,Nvl(Sst70,Decode(a1n04,'F4107','5','9')),cp13 " & _
                     "Order by StaffGroup,Sst70,a1n04,a1k13,a1k14,a1k15,a1k16"
    End If
   
    If adoAcc24N0.State = adStateOpen Then adoAcc24N0.Close
    adoAcc24N0.CursorLocation = adUseClient
    adoAcc24N0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoAcc24N0.RecordCount <> 0 Then
        InsertQueryLog (adoAcc24N0.RecordCount)
        If SaveExcel1 = True Then
            If strFileName <> MsgText(601) Then
                ShellExecute hLocalFile, "open", strFileName, vbNullString, vbNullString, 1
            End If
        End If
    Else
        InsertQueryLog (0)
        MsgBox "無資料產生"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHnd:
    If Err.Number = 70 Then
        MsgBox ChgSQL(strFileName) & "檔案已開啟！", vbCritical
    Else
        MsgBox Err.Description, vbCritical
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    'acccount也有用
    If IsObject(Forms(0)) Then
        ToolShow
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
    If KeyCode <> vbKeyEscape Then
       StatusView MsgText(101)
    End If
End Sub

Private Sub Form_Load()
    Dim intX As Integer
    Dim intY As Integer
    Dim sglWidth As Single
    Dim sglHeight As Single
   
    Me.Width = 6615
    Me.Height = 2730 'Modify by Amy 2023/10/11 原2400
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath4
    
    stST15 = Pub_StrUserSt15
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = DFormat
    
    FormClear
    
    Text11 = "2"
    StatusView MsgText(101)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    StatusClear
    Set Frmacc24n0 = Nothing
End Sub

'Add by Amy 2021/12/01 起日輸1號,迄日預帶當月底-秀玲
Private Sub MaskEdBox1_LostFocus()
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then Exit Sub
    If Right(Val(FCDate(MaskEdBox1.Text)), 2) = "01" Then
        MaskEdBox2.Text = CFDate(ACDate(GetLastDay(DBDATE(FCDate(MaskEdBox1)))))
    End If
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Dim strDate As String
    If MaskEdBox1 = MsgText(601) Or MaskEdBox1 = MsgText(29) Then Exit Sub
    
    strDate = Format(DBDATE(MaskEdBox1), "####/##/##")
    If IsDate(strDate) = False Then
        MsgBox "帳款起始日期格式錯誤！", , MsgText(5)
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    Dim strDate As String
    If MaskEdBox2 = MsgText(601) Or MaskEdBox2 = MsgText(29) Then Exit Sub
    
    strDate = Format(DBDATE(MaskEdBox2), "####/##/##")
    If IsDate(strDate) = False Then
        MsgBox "帳款迄止日期格式錯誤！", , MsgText(5)
        Cancel = True
        Exit Sub
    End If
End Sub

'系統類別
Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
End Sub

'系統類別
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'系統類別
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Text1(Index) = MsgText(601) Then Exit Sub
    
    If CheckSys(Text1(Index)) = MsgText(601) Then
        MsgBox MsgText(1107), , MsgText(5)
        Cancel = True
        Text1_GotFocus (Index)
        Exit Sub
    End If
End Sub

'報表內容
Private Sub Text11_GotFocus()
    TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
    If Text11 = MsgText(601) Then
        MsgBox "請輸入報表類別！", vbCritical
        Cancel = True
    End If
End Sub

Private Sub FormClear()
    Text2 = ""
    For Each oText In Text1
        oText = ""
    Next
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox2.Mask = DFormat
    Text11 = "1"
    '預設
    'Modify by Amy 2021/08/19 開放 江協理使用
    If Left(stST15, 2) = "F1" Or strUserNum = "98020" Then
        Text2 = "2"
    ElseIf Left(stST15, 2) = "F2" Then
        Text2 = "1"
    End If
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
    Dim bCancel As Boolean
        
    For Each oText In Text1
        If oText <> MsgText(601) Then
            Call Text1_Validate(oText.Index, bCancel)
            If bCancel = True Then
                Exit For
            End If
        End If
    Next
    If bCancel = True Then FormCheck = False: Exit Function
    
    If MaskEdBox1 = MsgText(601) Or MaskEdBox1 = MsgText(29) Then
        FormCheck = False
        MsgBox "帳款起始日期不可空白！", , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    End If
    Call MaskEdBox1_Validate(bCancel)
    If bCancel = True Then
        FormCheck = False
        Exit Function
    End If
    
    If MaskEdBox2 = MsgText(601) Or MaskEdBox2 = MsgText(29) Then
        FormCheck = False
        MsgBox "帳款迄止日期不可空白！", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Function
    End If
    Call MaskEdBox2_Validate(bCancel)
    If bCancel = True Then
        FormCheck = False
        Exit Function
    End If
    If Val(FCDate(MaskEdBox1)) > Val(FCDate(MaskEdBox2)) Then
        FormCheck = False
        MsgBox "帳款迄止日期不可大於起日！", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Function
    End If
    
    If Text2 = MsgText(601) Then
        FormCheck = False
        Text11.SetFocus
        MsgBox Label1(0) & "不可空白！", , MsgText(5)
        Exit Function
    'Add by Amy 2021/08/19 開放 江協理、葉易雲、洪琬姿 使用,商標人員只能查商標,專利人員只能查專利
    ElseIf stST15 <> "M51" Then
        '商標
        If Left(stST15, 2) = "F1" Or strUserNum = "98020" Then
            If Text2 = "1" Then
                FormCheck = False
                Text2.SetFocus
                MsgBox "不可查專利！", , MsgText(5)
                Exit Function
            End If
        '專利
        ElseIf Left(stST15, 2) = "F2" Then
            If Text2 <> "1" Then
                FormCheck = False
                Text2.SetFocus
                MsgBox "不可查商標！", , MsgText(5)
                Exit Function
            End If
        End If
    End If
    If Text11 = MsgText(601) Then
        FormCheck = False
        Text11.SetFocus
        MsgBox "報表內容不可空白！", , MsgText(5)
        Exit Function
    End If
    
    FormCheck = True
End Function

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strField1)
       If UCase(strField1(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub SetTitle(ByRef Wks As Worksheet)
    '專業單位
    If Text2 = "1" Then
        strTitle = "國外部專利處"
    'Modify by Amy 2021/08/06
    ElseIf Text2 = "2" Then
        strTitle = "國外部商標處"
    Else
        strTitle = "國外部商標處爭議案"
    End If
    'end 2021/08/06
    strTitle = strTitle & "專業點數明細表"
    
    '明細
    If Text11 = "1" Then
        ReDim strField1(5)
        ReDim intWidth1(UBound(strField1))
        strField1 = Array("點數分配人員", "本所案號", "請款編號", "收款日期", "已收金額", "已收點數", "財務點數")
        intWidth1 = Array(12.63, 16, 10.38, 8.88, 13, 13, 13)
    '合計
    Else
        ReDim strField1(3)
        ReDim intWidth1(UBound(strField1))
        strField1 = Array("點數分配人員", "", "已收金額", "已收點數", "財務點數")
        intWidth1 = Array(13, 15, 15, 15, 15)
    
    End If
    
    'Memo 表名設於頁首
    
    Wks.Range(Chr(intField) & intRow).Value = " 列印人員：" & StaffQuery(strUserNum)
    Wks.Range(Chr(intField + UBound(strField1)) & intRow).Value = " 列印日期：" & CFDate(strSrvDate(2))
    Wks.Range(Chr(intField + UBound(strField1) - 1) & intRow & ":" & Chr(intField + UBound(strField1)) & intRow).MergeCells = True
    Wks.Range(Chr(intField + UBound(strField1) - 1) & intRow & ":" & Chr(intField + UBound(strField1)) & intRow).HorizontalAlignment = xlRight
    intRow = intRow + 1
    
    For i = LBound(strField1) To UBound(strField1)
        Wks.Range(Chr(intField + i) & intRow).Value = strField1(i)
        Wks.Columns(Chr(intField + i)).ColumnWidth = intWidth1(i)
        If i = GetValue("已收金額") Or i = GetValue("財務點數") Then
            Wks.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlCenter
        Else
            Wks.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
        End If
    Next i
End Sub

Private Sub SetLine(intChoose As Integer, ByRef Xls As Excel.Application, ByRef Wks As Worksheet)
    Dim intS As Integer, intR As Integer
    
    Select Case intChoose
        Case 1 '小計
            If Text11 = "1" Then
                intS = GetValue("收款日期")
                intR = intRow
            Else
                intS = GetValue("")
                intR = intRow
            End If
        Case 2 '開始
            intS = GetValue("點數分配人員")
            intR = intRow
        Case 3 '結束
            intS = GetValue("點數分配人員")
            intR = intTitleR + 1
            Wks.Range(Chr(intField + intS) & intRow & ":" & Chr(intField + GetValue("請款編號")) & intRow).HorizontalAlignment = xlCenter
            Wks.Range(Chr(intField + intS) & intRow & ":" & Chr(intField + GetValue("請款編號")) & intRow).MergeCells = True
        Case 4 ' 組別小計
            intS = GetValue("點數分配人員")
            intR = intRow
            Wks.Range(Chr(intField + intS) & intR & ":" & Chr(intField + GetValue("請款編號")) & intRow).HorizontalAlignment = xlCenter
            Wks.Range(Chr(intField + intS) & intR & ":" & Chr(intField + GetValue("請款編號")) & intRow).MergeCells = True
    End Select
    Wks.Range(Chr(intField + intS) & intR & ":" & Chr(intField + UBound(strField1)) & intRow).Select
    Select Case intChoose
        Case 1
            Xls.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 2
            Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 3
            Xls.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 4
            Xls.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End Select
    
End Sub

Private Function SaveExcel1() As Boolean
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim strFormat As String, strSumR As String, strSumT As String
    Dim strOldPeo As String, strOldGroup As String, strGroupN As String, strST70 As String
    Dim intStartR As Integer, intPage As Integer, strTmp As String
    
On Error GoTo ErrHnd
    
    'Modify by Amy 2021/06/16 +專利 or 商標
    If Text2 = "1" Then
        strFileName = "專利"
    'Modify by Amy 2021/08/06 +商標爭議
    ElseIf Text2 = "2" Then
        strFileName = "商標"
    Else
        strFileName = "商標爭議"
    End If
    'end 2021/08/06
    strFileName = "FC" & strFileName & "專業收款明細表" & FCDate(MaskEdBox1) & "-" & FCDate(MaskEdBox2) & " " & ACDate(ServerDate) & ServerTime
    'end 2021/06/16
    If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
        MkDir strExcelPath
    End If
    If Dir(strExcelPath & strFileName & MsgText(43)) <> MsgText(601) Then
        Kill strExcelPath & strFileName & MsgText(43)
    ElseIf Dir(strExcelPath & strFileName & ".PDF") <> MsgText(601) Then
        Kill strExcelPath & strFileName & ".PDF"
    End If
    
    xlsAgentPoint.SheetsInNewWorkbook = 3 '工作表份數
    xlsAgentPoint.Workbooks.add
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    'xlsAgentPoint.Visible = True
    
    intField = 65: intRow = 1
    '表頭/欄位名稱
    Call SetTitle(wksrpt)
    Call SetLine(2, xlsAgentPoint, wksrpt)
    intTitleR = intRow
    intRow = intRow + 1
    
    intStartR = intRow
    wksrpt.Application.Selection.Font.Bold = False
    wksrpt.Application.Selection.Font.Size = 12
    '資料內容
    With adoAcc24N0
        Do While .EOF = False
            '*** 明細小計 ***
            If Text11 = "1" Then
                If strOldPeo <> "" & .Fields("a1n04") And strOldPeo <> MsgText(601) Then
                    For i = GetValue("收款日期") To GetValue("財務點數")
                        strFormat = ""
                        If i = GetValue("收款日期") Then
                            strTmp = "小計："
                            strSumR = strSumR & "," & Chr(intField + GetValue("已收金額")) & intRow
                        Else
                            strFormat = "#,##0"
                            If i = GetValue("已收點數") Or i = GetValue("財務點數") Then strFormat = "#,##0.000"
                            strTmp = "=Sum(" & Chr(intField + i) & intStartR & ":" & Chr(intField + i) & intRow - 1 & ")"
                        End If
                        wksrpt.Range(Chr(intField + i) & intRow).Value = strTmp
                        If strFormat <> MsgText(601) Then
                            wksrpt.Range(Chr(intField + i) & intRow).Font.Bold = True
                            wksrpt.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                            wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
                        Else
                            wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
                        End If
                    Next i
                    Call SetLine(1, xlsAgentPoint, wksrpt)
                    intRow = intRow + 1
                    intStartR = intRow
                End If
            End If
            '*** end 明細小計 ***
            
            '*** 組別合計 ***
            'Modify by Amy 2021/08/18 +3.FCT爭議
            If Text2 = "2" Or Text2 = "3" Then strST70 = "" & .Fields("Sst70")
            If strOldGroup <> "" & .Fields("StaffGroup") & strST70 And strOldGroup <> MsgText(601) Then
                For i = GetValue("收款日期") To GetValue("財務點數")
                    strTmp = "": strGroupN = ""
                    If i = GetValue("收款日期") Then
                        '專利
                        If Text2 = "1" Then
                            If strOldGroup = MsgText(601) Then
                                strGroupN = "其他　"
                            Else
                                strGroupN = PUB_GetFCPGrpName(strOldGroup, , False) & "　"
                            End If
                        '商標
                        Else
                            strTmp = Right(strOldGroup, 1)
                            If Val(strTmp) = 0 Then strTmp = ""
                            strGroupN = PUB_GetFCTGrpName(Left(strOldGroup, 1), strTmp, False)
                            If strTmp = MsgText(601) Then strGroupN = strGroupN & "其他"
                            strGroupN = strGroupN & "　"
                        End If
                        strTmp = strGroupN & "合計"
                        wksrpt.Range(Chr(intField + GetValue("點數分配人員")) & intRow).Value = strTmp
                        strSumT = strSumT & "," & Chr(intField + GetValue("已收金額")) & intRow
                    Else
                        '明細
                        If Text11 = "1" Then
                            strTmp = Mid(strSumR, 2)
                            strFormat = "#,##0"
                            If i <> GetValue("已收金額") Then
                                strTmp = Replace(strTmp, Chr(intField + GetValue("已收金額")), Chr(i + intField))
                                strFormat = "#,##0.000"
                            End If
                        '報表內容合計且為空白欄
                        ElseIf Text11 = "2" And i = 1 Then
                            strTmp = ""
                            wksrpt.Range(Chr(intField + i - 1) & intRow & ":" & Chr(intField + i) & intRow).HorizontalAlignment = xlCenter
                            wksrpt.Range(Chr(intField + i - 1) & intRow & ":" & Chr(intField + i) & intRow).MergeCells = True
                        '合計
                        Else
                            strFormat = "#,##0"
                            If i <> GetValue("已收金額") Then
                                strFormat = "#,##0.000"
                            End If
                            strTmp = Chr(intField + i) & intStartR & ":" & Chr(intField + i) & intRow - 1
                        End If
                        If strTmp <> MsgText(601) Then
                            wksrpt.Range(Chr(intField + i) & intRow).Value = "=Sum( " & strTmp & ")"
                            wksrpt.Range(Chr(intField + i) & intRow).Font.Bold = True
                            wksrpt.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                            wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
                        End If
                    End If
                Next i
                strSumR = ""
                Call SetLine(4, xlsAgentPoint, wksrpt)
                intRow = intRow + 1
                intStartR = intRow
            End If
            '*** end 組別合計 ***
            
            For i = LBound(strField1) To UBound(strField1)
                strTmp = "": strFormat = ""
                Select Case i
                    Case GetValue("點數分配人員")
                        strTmp = "" & .Fields("st02")
                    Case GetValue("本所案號")
                        If "" & .Fields("a1k13") <> MsgText(601) Then
                            strTmp = "" & .Fields("a1k13") & "-" & .Fields("a1k14") & "-" & _
                                                    .Fields("a1k15") & "-" & .Fields("a1k16")
                        End If
                    Case GetValue("請款編號")
                        '只顯示請請單號,傳票號不顯示
                        If Left("" & .Fields("a1k01"), 1) = "X" Then
                            strTmp = "" & .Fields("a1k01")
                        End If
                    Case GetValue("收款日期")
                        strTmp = "" & .Fields("DDate")
                    Case GetValue("已收金額")
                        strTmp = "" & .Fields("ReceVal")
                        strFormat = "#,##0"
                    Case GetValue("已收點數")
                        strTmp = "" & .Fields("ProVal")
                        strFormat = "#,##0.000"
                    Case GetValue("財務點數")
                        strTmp = "" & .Fields("AccPoint")
                        If strTmp <> 0 Then strTmp = Round(Val(strTmp) / 1000, 3)
                        strFormat = "#,##0.000"
                    Case Else
                        strTmp = "小計："
                End Select
                wksrpt.Range(Chr(intField + i) & intRow).Value = strTmp
                If strFormat <> MsgText(601) Then
                    wksrpt.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                    wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
                Else
                    wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
                End If
            Next i
            intRow = intRow + 1
            
            strOldGroup = "" & .Fields("StaffGroup")
            '商標需記錄 ST70
            If Text2 = "2" Or Text2 = "3" Then strOldGroup = strOldGroup & .Fields("Sst70")
            strOldPeo = "" & .Fields("a1n04")
            .MoveNext
        Loop
    End With
    '明細最後一個小計
    If Text11 = "1" Then
        For i = GetValue("收款日期") To GetValue("財務點數")
            strFormat = ""
            If i = GetValue("收款日期") Then
                strTmp = "小計："
                strSumR = strSumR & "," & Chr(intField + GetValue("已收金額")) & intRow
            Else
                strFormat = "#,##0"
                If i <> GetValue("已收金額") Then strFormat = "#,##0.000"
                strTmp = "=Sum(" & Chr(intField + i) & intStartR & ":" & Chr(intField + i) & intRow - 1 & ")"
            End If
            wksrpt.Range(Chr(intField + i) & intRow).Value = strTmp
            If strFormat <> MsgText(601) Then
                wksrpt.Range(Chr(intField + i) & intRow).Font.Bold = True
                wksrpt.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
            Else
                wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
            End If
        Next i
        Call SetLine(1, xlsAgentPoint, wksrpt)
        intRow = intRow + 1
    End If
    
    '分組最後一個合計
    For i = GetValue("收款日期") To GetValue("財務點數")
        strTmp = ""
        If i = GetValue("收款日期") Then
            '專利
            If Text2 = "1" Then
                If strOldGroup = MsgText(601) Then
                    strGroupN = "其他　"
                Else
                    strGroupN = PUB_GetFCPGrpName(strOldGroup, , False)
                End If
            '商標
            Else
                strTmp = Right(strOldGroup, 1)
                If Val(strTmp) = 0 Then strTmp = ""
                strGroupN = PUB_GetFCTGrpName(Left(strOldGroup, 1), strTmp, False)
                If strTmp = MsgText(601) Then strGroupN = strGroupN & "其他"
                strGroupN = strGroupN & "　"
            End If
            strTmp = strGroupN & "合計"
            wksrpt.Range(Chr(intField + GetValue("點數分配人員")) & intRow).Value = strTmp
            strSumT = strSumT & "," & Chr(intField + GetValue("已收金額")) & intRow
        Else
            '明細
            If Text11 = "1" Then
                strTmp = Mid(strSumR, 2)
                strFormat = "#,##0"
                If i <> GetValue("已收金額") Then
                    strTmp = Replace(strTmp, Chr(intField + GetValue("已收金額")), Chr(i + intField))
                    strFormat = "#,##0.000"
                End If
            '報表內容合計且為空白欄
            ElseIf Text11 = "2" And i = 1 Then
                strTmp = ""
                wksrpt.Range(Chr(intField + i - 1) & intRow & ":" & Chr(intField + i) & intRow).HorizontalAlignment = xlRight
                wksrpt.Range(Chr(intField + i - 1) & intRow & ":" & Chr(intField + i) & intRow).MergeCells = True
            '合計
            Else
                strFormat = "#,##0"
                If i <> GetValue("已收金額") Then
                    strFormat = "#,##0.000"
                End If
                strTmp = Chr(intField + i) & intStartR & ":" & Chr(intField + i) & intRow - 1
            End If
            If strTmp <> MsgText(601) Then
                wksrpt.Range(Chr(intField + i) & intRow).Value = "=Sum( " & strTmp & ")"
                wksrpt.Range(Chr(intField + i) & intRow).Font.Bold = True
                wksrpt.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
            End If
        End If
    Next i
    Call SetLine(4, xlsAgentPoint, wksrpt)
    intRow = intRow + 1
    
    '合計
    For i = GetValue("收款日期") To GetValue("財務點數")
        If i = GetValue("收款日期") Then
            wksrpt.Range(Chr(intField + GetValue("點數分配人員")) & intRow).Value = "合計："
        '明細
        Else
            If Text11 = "1" Then
                strTmp = Mid(strSumT, 2)
                strFormat = "#,##0"
                If i <> GetValue("已收金額") Then
                    strTmp = Replace(strTmp, Chr(intField + GetValue("已收金額")), Chr(i + intField))
                    strFormat = "#,##0.000"
                End If
            '報表內容合計且為空白欄
            ElseIf Text11 = "2" And i = 1 Then
                strTmp = ""
                wksrpt.Range(Chr(intField + i - 1) & intRow & ":" & Chr(intField + i) & intRow).HorizontalAlignment = xlRight
                wksrpt.Range(Chr(intField + i - 1) & intRow & ":" & Chr(intField + i) & intRow).MergeCells = True
            '合計
            Else
                strTmp = Mid(strSumT, 2)
                strFormat = "#,##0"
                If i <> GetValue("已收金額") Then
                    strTmp = Replace(strTmp, Chr(intField + GetValue("已收金額")), Chr(i + intField))
                    strFormat = "#,##0.000"
                End If
            End If
            If strTmp <> MsgText(601) Then
                wksrpt.Range(Chr(intField + i) & intRow).Value = "=Sum(" & strTmp & ")"
                wksrpt.Range(Chr(intField + i) & intRow).Font.Bold = True
                wksrpt.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                wksrpt.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
            End If
        End If
    Next i
    Call SetLine(1, xlsAgentPoint, wksrpt)
    Call SetLine(3, xlsAgentPoint, wksrpt)
      
    '設定
    wksrpt.PageSetup.PaperSize = 9 'A4
    wksrpt.PageSetup.Orientation = wdOrientLandscape '直印
    wksrpt.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
    wksrpt.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    wksrpt.PageSetup.CenterHeader = "&""-,粗體""&18" & strTitle & "&""-,標準""&12" & Chr(10) & "收款日期：" & MaskEdBox1 & " ~ " & MaskEdBox2
    wksrpt.PageSetup.RightHeader = Chr(10) & Chr(10) & "頁　　次：　  　　 " & "&P"
    wksrpt.PageSetup.TopMargin = xlsAgentPoint.InchesToPoints(0.98) '2.5
    wksrpt.PageSetup.BottomMargin = xlsAgentPoint.InchesToPoints(0.51) '1
    wksrpt.PageSetup.HeaderMargin = xlsAgentPoint.InchesToPoints(0.39)
    wksrpt.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0.51) '1
    wksrpt.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0.51) '1
    wksrpt.PageSetup.CenterHorizontally = True '垂直置中
    
    '判斷版本2007
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=-4143
    '版本2007以上
    Else
        xlsAgentPoint.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Kill strExcelPath & strFileName & ".xls"
    strFileName = strExcelPath & strFileName & ".pdf"
    SaveExcel1 = True
    Exit Function
    
ErrHnd:
    SaveExcel1 = False
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=-4143
    Else
        xlsAgentPoint.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Kill strExcelPath & strFileName & ".xls"
    Kill strExcelPath & strFileName & ".pdf"
    If Err.Number <> 0 Then
        MsgBox "資料產生有誤(錯誤:" & Err.Description & ")", vbCritical
    End If
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
    'Modify by Amy 2021/08/06 +3.FCT爭議
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_LostFocus()
    TextInverse Text2
End Sub
