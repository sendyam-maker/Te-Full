VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14b0 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶付款明細"
   ClientHeight    =   2844
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5412
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2844
   ScaleWidth      =   5412
   Begin VB.ComboBox CboCmp 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   30
      Width           =   3525
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   3450
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   2070
      Width           =   3450
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
      Left            =   3240
      TabIndex        =   5
      Top             =   750
      Width           =   1572
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
      Left            =   1320
      TabIndex        =   4
      Top             =   750
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      Default         =   -1  'True
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
      Left            =   330
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   1680
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   420
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
      Left            =   3240
      TabIndex        =   3
      Top             =   420
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
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "會回寫客戶回執記錄檔"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   120
      TabIndex        =   16
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "Label2(2)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   550
      Index           =   2
      Left            =   1380
      TabIndex        =   15
      Top             =   1110
      Width           =   3495
   End
   Begin VB.Label Label3 
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
      Left            =   330
      TabIndex        =   14
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   13
      Top             =   2415
      Width           =   1725
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
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2100
      Width           =   975
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
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   750
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   330
      TabIndex        =   10
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "付款日期"
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
      TabIndex        =   9
      Top             =   420
      Width           =   975
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
      Left            =   3000
      TabIndex        =   1
      Top             =   390
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc14b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/03/15 Form2.0已修改 (地址條改為標籤地址條套印/付款通知單改為開Word畫表格印)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0e0 As New ADODB.Recordset
'Mark by Amy 2022/03/14 不使用
'Public adoaccrpt111 As New ADODB.Recordset
'Public adoquery As New ADODB.Recordset
'Dim dllaccrpt111 As Object
Dim strSql As String
Dim strAmount As String
Dim intLength As Integer
'Add by Morgan 2006/11/1
Dim lngYo As Long '列印起始位置
Dim lngHalfHeight As Long '中一刀位置
Dim lngPageNo As Long '頁數
Dim strRetNo As String '回執單號
'預設印表機
Dim m_DefaultPrinter As String
Dim m_bolPrint As Boolean '是否有列印資料
Dim m_NoMatchMsg As String '付款金額與案件明細不符提醒訊息
Dim strPrinter As String, strPrinter2 As String 'Add By Sindy 2013/6/4
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17
Dim strAddrData As String, StrSQLa As String 'Add by Amy 2022/03/28

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
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

Private Sub Command1_Click()
Dim bCancel As Boolean 'Add by Amy 2014/01/27
   
   m_NoMatchMsg = ""
   m_bolPrint = False
   
   lngYo = 0 'Add by Morgan 2006/11/1
   'Add by Amy 2014/01/27 +公司別不可為空
   'Modify By Sindy 2020/4/23
   'If Text3 = MsgText(601) Then
   If CboCmp.Text = MsgText(601) Then
   '2020/4/23 END
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Exit Sub
   End If
   Call CboCmp_Validate(bCancel)
   If bCancel = True Then
      Exit Sub
   End If
   'end 2014/01/27
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   strSql = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = " and a0q01 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0q01 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0q03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0q03 <= '" & Text2 & "'"
   End If
   
   Call SetCompN 'Add by Sindy 2020/04/23
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2022/03/28 改套印
   strAddrData = ""
   'PUB_RestorePrinter Combo2 'Add By Sindy 2013/6/4
   PrintAddress
   'PUB_RestorePrinter strPrinter2 'Add By Sindy 2013/6/4
   'end 2022/03/15
   Screen.MousePointer = vbDefault
   
   'Mark by Amy 2022/03/14 與瑞婷確認已不使用
'   Screen.MousePointer = vbHourglass
'   Accrpt111Delete
'   If ProduceData Then
'      m_bolPrint = True
'      PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2013/6/4
'      If adoaccrpt111.State = adStateOpen Then
'         adoaccrpt111.Close
'      End If
'      Set dllaccrpt111 = CreateObject("AccReport.ReportSelect")
'      adoaccrpt111.CursorLocation = adUseClient
'      adoaccrpt111.Open "select * from accrpt111 Where R11108 Is Null Or substr(R11108, 1, 1) <> '3' ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccrpt111.RecordCount <> 0 Then
'         Screen.MousePointer = vbDefault
'         'Modify by Morgan 2006/11/10  加可取消
'         'MsgBox ReportTitle(111), , MsgText(5)
'         If MsgBox(ReportTitle(111), vbOKCancel, MsgText(5)) = vbCancel Then
'            GoTo NoPrint1
'         End If
'         'end 2006/11/10
'         Screen.MousePointer = vbHourglass
'         'Modify by Amy 2014/01/28 +公司別名稱
'         dllaccrpt111.Acc14b0 strCmpN & "," & ReportTitle(111), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)) & ",1"
'         Me.SetFocus 'Add by Morgan 2006/11/10
'      End If
'NoPrint1:
'      adoaccrpt111.Close
'      '寄分所
'      If adoaccrpt111.State = adStateOpen Then
'         adoaccrpt111.Close
'      End If
'      adoaccrpt111.CursorLocation = adUseClient
'      adoaccrpt111.Open "select * from accrpt111 Where  substr(R11108, 1, 1) = '3' ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccrpt111.RecordCount <> 0 Then
'         Screen.MousePointer = vbDefault
'         'Modify by Morgan 2006/11/10  加可取消
'         'MsgBox ReportTitle(111) & " (寄分所)", , MsgText(5)
'         If MsgBox(ReportTitle(111) & " (寄分所)", vbOKCancel, MsgText(5)) = vbCancel Then
'           GoTo NoPrint2
'         End If
'         'end 2006/11/10
'         Screen.MousePointer = vbHourglass
'         'Modify by Amy 2014/01/28 +公司別名稱
'         dllaccrpt111.Acc14b0 strCmpN & "," & ReportTitle(111), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)) & ",3"
'      End If
'NoPrint2:
'      adoaccrpt111.Close
'      PUB_SetOsDefaultPrinter strPrinter 'Add By Sindy 2013/6/4
'   End If
'   Screen.MousePointer = vbDefault
   'end 2022/03/14
   
   'Remove by Morgan 2008/7/23 不用印 -- 瑞婷
   'Screen.MousePointer = vbHourglass
   'PrintPaySign
   'Screen.MousePointer = vbDefault
   
   PUB_SetOsDefaultPrinter Combo1 'Add by Amy 2022/03/28
   PUB_RestorePrinter Combo1 'Add By Sindy 2013/6/4
   Screen.MousePointer = vbHourglass
   PrintPayNotice
   Screen.MousePointer = vbDefault
   
   'Add by Morgan 2008/7/23 非票據的銷退也要印回執 -- 瑞婷
   Screen.MousePointer = vbHourglass
   PrintPayNotice3
   Screen.MousePointer = vbDefault
   'end 2008/7/23
   PUB_SetOsDefaultPrinter strPrinter 'Add by Amy 2022/03/28
   PUB_RestorePrinter strPrinter 'Add By Sindy 2013/6/4
   
   'Add by Morgan 2008/10/8
   If m_NoMatchMsg <> "" Then
      MsgBox m_NoMatchMsg, vbOKOnly, "退費金額與銷退收據不符合"
   End If
   
   If m_bolPrint = False Then
      MsgBox MsgText(28), , MsgText(5)
   End If
   
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   'Set dllaccrpt111 = Nothing'Mark by Amy 2022/03/14
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   'Modify by Amy 2023/07/19
   'PUB_InitForm Me, 5490, 3165
   PUB_InitForm Me, 5628, 3408 '2790 '2085
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
   'Set dllaccrpt111 = CreateObject("AccReport.ReportSelect") 'Mark by Amy 2022//03/14
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
   PUB_SetPrinter Me.Name, Combo2, strPrinter2 'Add By Sindy 2013/6/4
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, False, False, False, , 1)
   CboCmp.ListIndex = 1
   'end 2020/04/17
   
   'Add by Amy  2014/01/29 公司別預帶1
'   Text3 = "1"
'   Text13 = A0802Query(Text3)
   'Removed by Morgan 2023/11/22 公司別應該要可選,否則L公司的沒法印回執
   'CboCmp.Enabled = False
   'CboCmp.Locked = True
   'end 2023/11/22
   'end 2014/01/29
   'Add by Amy 2022/03/28 因為報表程式,會忽略可能有回寫,故提醒
   Label4.Visible = False
   If Pub_StrUserSt03 = "M51" Then
     Label4.Visible = True
   End If
   Label2(2).Caption = "列印時，請勿開啟Word" & vbCrLf & _
                                   "可以使用「標籤地址條」"
   'end 2022/03/15
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Add By Sindy 2013/6/4
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2013/6/4 END
   
   'Set dllaccrpt111 = Nothing 'Mark by Amy 2022/03/14
   Set Frmacc14b0 = Nothing
End Sub

'*************************************************
'  產生報表資料
'  Memo by Amy 2022/03/14 與瑞婷確認不使用了
'*************************************************
Private Function ProduceData() As Boolean
'On Error GoTo Checking
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   If adoaccrpt111.State = adStateOpen Then
'      adoaccrpt111.Close
'   End If
'   adoaccrpt111.CursorLocation = adUseClient
'   adoaccrpt111.Open "select * from accrpt111", adoTaie, adOpenDynamic, adLockBatchOptimistic
'
'' 國內應付資料
'   If adoacc0e0.State = adStateOpen Then
'      adoacc0e0.Close
'   End If
'   adoacc0e0.CursorLocation = adUseClient
''   adoacc0e0.Open "select * from acc0e0, acc1p0 where a0e01 = a1p10 and a0e02 = a1p09 and a1p02 = 'C' AND A1P24 NOT IN ('1', '4')" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'   'adoacc0e0.Open "select * from acc0q0, acc1p0 where a0q01 = a1p18 and a0q03 = a1p15 and a1p02 = 'C' AND A1P24 NOT IN ('1', '4')" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'   'Modify by Amy 2014/01/27 +公司別
'   'adoacc0e0.Open "select * from acc0q0, acc1p0 where a1p04 = a0q17 and a1p02 = 'C' AND A1P24 NOT IN ('1', '4') and a1p05 = '2111'" & strSql & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
'   adoacc0e0.Open "select * from acc0q0, acc1p0 where a1p01='" & strCmp & "' And a1p01=a0q19 And a1p04 = a0q17 and a1p02 = 'C' AND A1P24 NOT IN ('1', '4') and a1p05 = '2111'" & strSql & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0e0.RecordCount = 0 Then
'      'MsgBox MsgText(28), , MsgText(5)
'   End If
'   If adoacc0e0.RecordCount = 0 Then
'      ProduceData = False
'      adoacc0e0.Close
'      Exit Function
'   Else
'      ProduceData = True
'   End If
'   Do While adoacc0e0.EOF = False
'      adoaccrpt111.AddNew
'      adoaccrpt111.Fields("r11101").Value = strUserNum
'      If IsNull(adoacc0e0.Fields("a0q03").Value) Then
'         adoaccrpt111.Fields("r11102").Value = Null
'      Else
'         adoaccrpt111.Fields("r11102").Value = adoacc0e0.Fields("a0q03").Value
'      '   Select Case adoacc0e0.Fields("a0q04").Value
'      '      Case Mid(ComboItem(92), 1, 1)
'      '         adoaccrpt111.Fields("r11103").Value = CustomerQuery(adoacc0e0.Fields("a0q03").Value, 1)
'      '      Case Mid(ComboItem(91), 1, 1)
'      '         adoaccrpt111.Fields("r11103").Value = A0i02Query(adoacc0e0.Fields("a0q03").Value)
'      '      Case Mid(ComboItem(93), 1, 1)
'      '         adoaccrpt111.Fields("r11103").Value = StaffQuery(adoacc0e0.Fields("a0q03").Value)
'      '   End Select
'      End If
'      If IsNull(adoacc0e0.Fields("a0q05").Value) Then
'         adoaccrpt111.Fields("r11103").Value = Null
'      Else
'         adoaccrpt111.Fields("r11103").Value = adoacc0e0.Fields("a0q05").Value
'      End If
'      If IsNull(adoacc0e0.Fields("a1p12").Value) Then
'         adoaccrpt111.Fields("r11104").Value = Null
'      Else
'         adoaccrpt111.Fields("r11104").Value = adoacc0e0.Fields("a1p12").Value
'      End If
'      If IsNull(adoacc0e0.Fields("a1p04").Value) Then
'         adoaccrpt111.Fields("r11110").Value = Null
'      Else
'         adoaccrpt111.Fields("r11110").Value = adoacc0e0.Fields("a1p04").Value
'      End If
'      adoaccrpt111.Fields("r11105").Value = adoacc0e0.Fields("a1p09").Value
'      If IsNull(adoacc0e0.Fields("a1p08").Value) Then
'         adoaccrpt111.Fields("r11106").Value = Null
'      Else
'         adoaccrpt111.Fields("r11106").Value = adoacc0e0.Fields("a1p08").Value
'      End If
'      If IsNull(adoacc0e0.Fields("a0q16").Value) Then
'         adoaccrpt111.Fields("r11107").Value = Null
'      Else
'         adoaccrpt111.Fields("r11107").Value = adoacc0e0.Fields("a0q16").Value
'      End If
'      If IsNull(adoacc0e0.Fields("A1P24").Value) Then
'         adoaccrpt111.Fields("R11108").Value = Null
'      Else
'         Select Case adoacc0e0.Fields("A1P24").Value
'            Case Mid(ComboItem(81), 1, 1)
'               adoaccrpt111.Fields("R11108").Value = ComboItem(81)
'            Case Mid(ComboItem(82), 1, 1)
'               adoaccrpt111.Fields("R11108").Value = ComboItem(82)
'            Case Mid(ComboItem(83), 1, 1)
'               adoaccrpt111.Fields("R11108").Value = ComboItem(83)
'            Case Mid(ComboItem(84), 1, 1)
'               adoaccrpt111.Fields("R11108").Value = ComboItem(84)
'         End Select
'      End If
'      If IsNull(adoacc0e0.Fields("A1P26").Value) Then
'         adoaccrpt111.Fields("R11109").Value = Null
'      Else
'         Select Case adoacc0e0.Fields("A1P26").Value
'            Case Mid(ComboItem(111), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(111)
'            Case Mid(ComboItem(112), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(112)
'            Case Mid(ComboItem(113), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(113)
'            Case Mid(ComboItem(114), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(114)
'            Case Mid(ComboItem(115), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(115)
'            Case Mid(ComboItem(116), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(116)
'            Case Mid(ComboItem(117), 1, 1)
'               adoaccrpt111.Fields("R11109").Value = ComboItem(117)
'         End Select
'      End If
'      adoaccrpt111.UpdateBatch
'      adoacc0e0.MoveNext
'   Loop
'   adoacc0e0.Close
'   adoaccrpt111.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Function
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Function

'*************************************************
'  刪除報表資料
'  Mark by Amy 2022/03/28 不使用
'*************************************************
Private Sub Accrpt111Delete()
'   adoTaie.Execute "delete from accrpt111"
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
   Text1 = ""
   Text2 = ""
   MaskEdBox1.SetFocus 'Modify by Amy 2014/01/29
   'Add by Amy 2014/01/29
   'Text3 = ""
   'Text13 = ""
   'Text3.SetFocus
   'end 2014/01/29
End Sub

'*************************************************
'  列印付款通知單
'
'*************************************************
Private Sub PrintPayNotice()
   'Modify by Morgan 2007/8/17 改簽收也要印
   'adoacc0e0.Open "select * from acc1p0, acc0e0, acc0q0 where a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24 NOT IN ('1', '4')" & strSQL & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2008/1/29 + a1p08>0
   'Modify by Amy 2014/01/27 +公司別
   'adoacc0e0.Open "select * from acc1p0, acc0e0, acc0q0 where a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24<>'4' and a1p08>0 " & strSql & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/07/06 +a0e07 因改為key
   'Modify by Amy 2025/01/15 因為慢,調語法,並過濾票據資料才印
   'ex:1140114 X88492000 G11401030 因目前改以[電匯]方式,而畫面無此選項,故使用預設[支票],無票據資料跑PUB_PrintReceipt_Doc無a0e01資料會錯
'   StrSQLa = "Select * From acc1p0, acc0e0, acc0q0 Where a1p01='" & strCmp & "' " & _
'                 "And a1p01=a0q19 And a1p01=a0e23 and a1p11=a0e07 And a1p04 = a0q17 " & _
'                 "And a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24<>'4' and a1p08>0 " & _
'                 strSql & " order by a1p04 asc"
    StrSQLa = "Select * From acc1p0, acc0e0, acc0q0 Where a0q19='" & strCmp & "' " & _
                 "And a1p01(+)=a0q19 And a1p01=a0e23(+) and a1p11=a0e07(+) And a1p04(+) = a0q17 " & _
                 "And a1p10 = a0e01(+) and a1p09 = a0e02(+)  and a1p02(+) = 'C' AND A1P24<>'4' and a1p08>0 " & _
                 strSql & " And a0e01 is not null order by a1p04 asc"
   If adoacc0e0.State <> adStateClosed Then adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount <> 0 Then
      Screen.MousePointer = vbDefault
      'Modify by Morgan 2006/11/10
      'MsgBox MsgText(100) & ReportTitle(1111), , MsgText(5)
      If MsgBox(MsgText(100) & ReportTitle(1111), vbOKCancel, MsgText(5)) = vbCancel Then
         GoTo NoPrint
      End If
      'end 2006/11/10
      Screen.MousePointer = vbHourglass
   Else
NoPrint:
      adoacc0e0.Close
      Exit Sub
   End If
   
   'Mark by Amy 2022/03/28 改開Word畫表格印
'   Printer.EndDoc '回復印表機預設值
'   'Add by Morgan 2006/11/2
'   lngHalfHeight = Printer.Height / 2 '中一刀起始位置
'   lngYo = 0 '列印起始位置
'   lngPageNo = 0
'   'end 2006/11/2
'end 2022/03/18
   m_bolPrint = True
   Do While adoacc0e0.EOF = False
      'Modify by Morgan 2006/11/1 加案件退費格式
      If adoacc0e0("a1p26") = "3" Then
         'PrintPayNotice1
         PUB_PrintReceipt_Doc Me.Name, "3", adoacc0e0, strCmp, "", m_NoMatchMsg
      Else
         'PrintPayNotice2
         PUB_PrintReceipt_Doc Me.Name, "1", adoacc0e0, strCmp, "", ""
      End If
      adoacc0e0.MoveNext
   Loop
   'Printer.EndDoc''Mark by Amy 2022/03/28 改開Word畫表格印
   adoacc0e0.Close
End Sub


'*************************************************
'  列印付款通知單 (案件退費格式)
'  Mark by Amy 2022/03/28 改共用function 開Word畫表格印
'*************************************************
Private Sub PrintPayNotice1()
'   Dim strCompName As String, strCaseDesc As String
'   Dim lngAmount As Long, lngAmt As Long
'   Dim adoacc0s0 As New ADODB.Recordset
'   Dim iCount As Integer
'
'   '公司別,抓退費明細(可能會有多張收據)
'   'Modify by Amy 2014/01/27 公司別
'   'strExc(0) = "select a0k01,a0k11,a0s06,a0s07,a0o01,a0s17,a0k04 from acc0o0, acc0s0, acc0k0 " & _
'      " where a0o03='" & adoacc0e0("a0q03") & "' and a0o11=" & adoacc0e0("a0q01") & " and a0s01(+)=a0o09 and a0o09 is not null and substr(a0s02,1,1)='E'" & _
'      " and a0k01(+)=a0s02 order by a0s02"
'    strExc(0) = "select a0k01,a0k11,a0s06,a0s07,a0o01,a0s17,a0k04 from acc0o0, acc0s0, acc0k0 " & _
'      " where a0o07='" & adoacc0e0("a0q19") & "' And a0o03='" & adoacc0e0("a0q03") & "' and a0o11=" & adoacc0e0("a0q01") & " and a0s01(+)=a0o09 and a0o09 is not null and substr(a0s02,1,1)='E'" & _
'      " and a0k01(+)=a0s02 order by a0s02"
'   intI = 1
'   'edit by nickc 2007/02/07 不用 dll 了
'   'Set adoacc0s0 = objLawDll.ReadRstMsg(intI, strExc(0))
'   Set adoacc0s0 = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      strCompName = A0802Query("" & adoacc0s0("a0k11"))
'      strCaseDesc = PUB_GetCaseInfo("" & adoacc0s0("a0k01"))
'   End If
'
'   If lngYo > 0 Then
'      Printer.NewPage
'      lngPageNo = lngPageNo + 1
'      lngYo = 0
'   Else
'      If lngPageNo = 0 Then
'         lngPageNo = lngPageNo + 1
'      Else
'         lngYo = lngHalfHeight
'      End If
'   End If
'
'   Printer.Font = "新細明體"
'   Printer.FontSize = 14
'   Printer.CurrentX = 3650
'   Printer.CurrentY = lngYo + 300
'   Printer.Print strCompName
'
'   Printer.CurrentX = 4000
'   Printer.CurrentY = lngYo + 800
'   Printer.Print "***  付款通知單  ***"
'
'   Printer.FontSize = 12
'   '製表日期
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 1000
'   Printer.Print "製表日期: " & CFDate(strSrvDate(2))
'
'   Printer.Line (200, lngYo + 1250)-(10000, lngYo + 1250)
'
'   '客戶名稱
'   If Not IsNull(adoacc0e0.Fields("a0q05").Value) Then
'      Printer.CurrentX = 200
'      Printer.CurrentY = lngYo + 1500
'      Printer.Print adoacc0e0.Fields("a0q05").Value & "     台 鑒:"
'   End If
'
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 1900
'   Printer.Print "　 　 茲 寄 上 應 付    台 端  ( 貴 公 司 )  之 票 據  ( 詳 述 如 下 )  ， 並 將 退 費 收 訖 憑 單"
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 2300
'   Printer.Print "填 妥 寄 回 ， 謝 謝 您 的 支 持 與 合 作 。"
'
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3000
'   Printer.Print "付款行庫: " & A0g02Query(adoacc0e0.Fields("a0e01").Value)
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3300
'   Printer.Print "付款帳號: " & adoacc0e0.Fields("a0e07").Value
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3600
'   Printer.Print "支票號碼: " & adoacc0e0.Fields("a0e02").Value
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3900
'   Printer.Print "到 期 日:   " & CFDate(adoacc0e0.Fields("a0e10").Value)
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 4200
'   lngAmount = Val("" & adoacc0e0.Fields("a0e11").Value)
'   Printer.Print "金　　額: " & "$" & Format(lngAmount, DDollar) & "**"
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 4500
'   Printer.Print "備　　註: " & ComboItem(113)
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 5200
'
'   'Modify by Morgan 2011/7/29 考慮多行
'   'Printer.Print "退 費 明 細 : " & strCaseDesc
'   Printer.Print "退 費 明 細 : "
'   iCount = 0
'   Do While strCaseDesc <> ""
'      Printer.CurrentX = 200 + Printer.TextWidth("退 費 明 細 : ")
'      Printer.CurrentY = lngYo + 5200 + iCount * 300
'      intI = getCutPos(strCaseDesc, Printer.TextWidth(String(33, "　")))
'      If intI = 0 Then
'         Printer.Print strCaseDesc
'         strCaseDesc = ""
'      Else
'         Printer.Print Left(strCaseDesc, intI)
'         iCount = iCount + 1
'         strCaseDesc = Mid(strCaseDesc, intI + 1)
'      End If
'   Loop
'   'End 2011/7/29
'
'
'   '退費收訖憑單
'   With adoacc0s0
'   If .RecordCount > 0 Then
'      Do While Not .EOF
'         'Modified by Morgan 2012/5/30 不必減稅款退費金額--瑞婷
'         'lngAmt = Val("" & .Fields("a0s06")) + Val("" & .Fields("a0s07")) - Val("" & .Fields("a0s17"))
'         lngAmt = Val("" & .Fields("a0s06")) + Val("" & .Fields("a0s07"))
'         If .AbsolutePosition = .RecordCount Then
'            'Modify by Morgan 2008/10/7 改彈訊息提醒
'            'lngAmt = lngAmount
'            'Modified by Morgan 2012/10/11
'            'If lngAmt <> lngAmount Then
'            If lngAmt - Val("" & .Fields("a0s17")) <> lngAmount Then
'               m_NoMatchMsg = m_NoMatchMsg & vbCrLf & "<" & adoacc0e0("a0q03") & ">" & adoacc0e0("a0q05")
'            End If
'         End If
'         PUB_PrintReceipt3 Me.adoacc0e0, adoacc0s0, lngYo, lngPageNo, lngAmt
'         'Modified by Morgan 2013/5/13
'         'lngAmount = lngAmount - lngAmt
'         lngAmount = lngAmount - (lngAmt - Val("" & .Fields("a0s17")))
'
'         .MoveNext
'      Loop
'
'   End If
'   End With
'
'   Set adoacc0s0 = Nothing
End Sub
'Add by Morgan 2011/7/29
Private Function getCutPos(p_Desc As String, p_lWidth As Long) As Integer
   Dim i As Integer
   For i = 1 To Len(p_Desc)
      If Printer.TextWidth(Left(p_Desc, i)) > p_lWidth Then
         getCutPos = i - 1
         Exit For
      End If
   Next
End Function

'*************************************************
'  列印付款通知單 (非案件退費格式)
'  Mark by Amy 2022/03/28 改共用function 開Word畫表格印
'*************************************************
Private Sub PrintPayNotice2()
'   If lngYo > 0 Then
'      Printer.NewPage
'      lngPageNo = lngPageNo + 1
'      lngYo = 0
'   Else
'      If lngPageNo = 0 Then
'         lngPageNo = lngPageNo + 1
'      Else
'         lngYo = lngHalfHeight
'      End If
'   End If
'
'   '付款通知單
'   Printer.Font = "新細明體"
'   Printer.FontSize = 14
'   Printer.CurrentX = 3650
'   Printer.CurrentY = lngYo + 300
'   Printer.Print strCmpN 'A0802Query(Text3) 'Modify by Amy 2014/01/28 改公司別 原:"1"
'
'   Printer.CurrentX = 4000
'   Printer.CurrentY = lngYo + 800
'   Printer.Print ReportTitle(1111)
'
'   Printer.FontSize = 12
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 1000
'   Printer.Print ReportSum(35) & CFDate(strSrvDate(2))
'
'   Printer.Line (200, 1250)-(10000, 1250)
'
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 1500
'   If IsNull(adoacc0e0.Fields("a0q05").Value) Then
'      Printer.Print ""
'   Else
'      Printer.Print adoacc0e0.Fields("a0q05").Value & ReportSum(43)
'   End If
'
'   Printer.CurrentX = 700
'   Printer.CurrentY = lngYo + 1900
'   Printer.Print ReportSum(44)
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 2300
'   Printer.Print ReportSum(45)
'
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3000
'   Printer.Print ReportSum(37) & A0g02Query(adoacc0e0.Fields("a0e01").Value)
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3300
'   Printer.Print ReportSum(38) & adoacc0e0.Fields("a0e07").Value
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3600
'   Printer.Print ReportSum(39) & adoacc0e0.Fields("a0e02").Value
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 3900
'   Printer.Print ReportSum(40) & CFDate(adoacc0e0.Fields("a0e10").Value)
'
'   strAmount = "$" & Format(adoacc0e0.Fields("a0e11").Value, DDollar) & "**"
'   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 4200
'   Printer.Print ReportSum(41) & strAmount
'
'   Printer.CurrentX = 1000
'   Printer.CurrentY = lngYo + 4500
'   Select Case adoacc0e0.Fields("a1p26").Value
'      Case "1"
'         Printer.Print ReportSum(42) & ComboItem(111)
'      Case "2"
'         Printer.Print ReportSum(42) & ComboItem(112)
'      Case "3"
'         Printer.Print ReportSum(42) & ComboItem(113)
'      Case "4"
'         Printer.Print ReportSum(42) & ComboItem(114)
'      Case "5"
'         Printer.Print ReportSum(42) & ComboItem(115)
'      Case "6"
'         Printer.Print ReportSum(42) & ComboItem(116)
'      Case "7"
'         'Modify by Morgan 2006/11/2
'         'Printer.Print ReportSum(42) & ComboItem(117)
'         If Not IsNull(adoacc0e0("a0q18")) Then
'            Printer.Print ReportSum(42) & adoacc0e0("a0q18")
'         Else
'            Printer.Print ReportSum(42) & ComboItem(117)
'         End If
'   End Select
'
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngYo + 5200
'   Printer.Print ReportSum(46)
'   Printer.CurrentX = 2000
'   Printer.CurrentY = lngYo + 5200
'   Printer.Print ReportSum(47)
'
'   '票據受領收據
'   PUB_PrintReceipt1 adoacc0e0, lngYo, lngPageNo
End Sub
'*************************************************
'  列印付款通知單 (抬頭及報表格式)
'
'*************************************************
Private Sub PrintNoticeHead()
Dim i As Integer

   Printer.FontSize = 14
   Printer.CurrentX = 3650
   Printer.CurrentY = lngYo + 300
   Printer.Print A0802Query("1")
   Printer.CurrentX = 4000
   Printer.CurrentY = lngYo + 800
   Printer.Print ReportTitle(1111)
   Printer.FontSize = 12
   Printer.CurrentX = 200
   Printer.CurrentY = lngYo + 1000
   Printer.Print ReportSum(35) & CFDate(strSrvDate(2))
   Printer.Line (200, 1250)-(10000, 1250)
   Printer.CurrentX = 700
   Printer.CurrentY = lngYo + 1900
   Printer.Print ReportSum(44)
   Printer.CurrentX = 200
   Printer.CurrentY = lngYo + 2300
   Printer.Print ReportSum(45)
   Printer.CurrentX = 200
   Printer.CurrentY = lngYo + 5200
   Printer.Print ReportSum(46)
   Printer.CurrentX = 2000
   Printer.CurrentY = lngYo + 5200
   Printer.Print ReportSum(47)
   Printer.CurrentX = 200
   Printer.CurrentY = lngYo + 7500
   Printer.Print ReportSum(50)
   Printer.FontSize = 14
   Printer.CurrentX = 3800
   Printer.CurrentY = lngYo + 9000
   Printer.Print ReportTitle(1114)
   Printer.FontSize = 12
   Printer.CurrentX = 200
   Printer.CurrentY = lngYo + 9000
   Printer.Print ReportSum(35) & CFDate(strSrvDate(2))
   Printer.Line (200, 9500)-(10000, 9500)
   Printer.CurrentX = 200
   Printer.CurrentY = lngYo + 10900
   Printer.Print ReportSum(51)
   Printer.CurrentX = 1000
   Printer.CurrentY = lngYo + 11500
   Printer.Print "    票           據           內            容"
   Printer.CurrentX = 5100
   Printer.CurrentY = lngYo + 11500
   Printer.Print " 蓋                                      章"
   Printer.Line (700, 11400)-(8500, 11400)
   Printer.Line (700, 11800)-(8500, 11800)
   Printer.Line (700, 13900)-(8500, 13900)
   Printer.Line (700, 14300)-(8500, 14300)
   Printer.Line (700, 11400)-(700, 14300)
   Printer.Line (5000, 11400)-(5000, 13900)
   Printer.Line (8500, 11400)-(8500, 14300)
End Sub

''*************************************************
''  列印付款簽收簿
''
''*************************************************
'Private Sub PrintPaySign()
'
'   Dim intCounter As Integer
'   Dim lngCounter As Long
'
'   adoacc0e0.CursorLocation = adUseClient
'   'Modify by Amy 2014/01/27 未使用但+公司別
'   'adoacc0e0.Open "select * from acc1p0, acc0e0, acc0q0 where a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24 = '1' AND A1P05 = '2111'" & strSql & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
'   adoacc0e0.Open "select * from acc1p0, acc0e0, acc0q0 where a1p01='" & Text3 & "' And a1p01=a0e23 And a1p01=a1q19 And a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24 = '1' AND A1P05 = '2111'" & strSql & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0e0.RecordCount <> 0 Then
'      Screen.MousePointer = vbDefault
'      'Modify by Morgan 2006/11/10  加可取消
'      'MsgBox MsgText(100) & ReportTitle(1112), , MsgText(5)
'      If MsgBox(MsgText(100) & ReportTitle(1112), vbOKCancel, MsgText(5)) = vbCancel Then
'         GoTo NoPrint
'      End If
'      'end 2006/11/10
'      Screen.MousePointer = vbHourglass
'   Else
'NoPrint:
'      adoacc0e0.Close
'      Exit Sub
'   End If
'
'   intCounter = 0
'   lngCounter = 1
'   Printer.EndDoc '清除印表機設定
'   Printer.Font = "新細明體"
'   PrintHead
'   Printer.CurrentX = 9500
'   Printer.CurrentY = 1100
'   Printer.Print lngCounter
'
'   Do While adoacc0e0.EOF = False
'      If intCounter > 5 Then
'         intCounter = 0
'         lngCounter = lngCounter + 1
'         Printer.NewPage
'         PrintHead
'         Printer.CurrentX = 9500
'         Printer.CurrentY = 1100
'         Printer.Print lngCounter
'      End If
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 2050 + 2000 * intCounter
'      Printer.Print ReportSum(37) & MidB(A0802Query(adoacc0e0.Fields("a1p01").Value), 1, 4)
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 2350 + 2000 * intCounter
'      Printer.Print ReportSum(38) & adoacc0e0.Fields("a0e07").Value
'      Printer.CurrentX = 150
'      Printer.CurrentY = 2650 + 2000 * intCounter
'      If IsNull(adoacc0e0.Fields("a0q05").Value) Then
'         Printer.Print ""
'      Else
'         Printer.Print MidB(adoacc0e0.Fields("a0q05").Value, 1, 8)
'      End If
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 2650 + 2000 * intCounter
'      Printer.Print ReportSum(39) & adoacc0e0.Fields("a0e02").Value
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 2950 + 2000 * intCounter
'      Printer.Print ReportSum(40) & CFDate(adoacc0e0.Fields("a0e10").Value)
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 3250 + 2000 * intCounter
'      Printer.Print ReportSum(41)
'      strAmount = "$" & Format(adoacc0e0.Fields("a0e11").Value, DDollar) & "**"
'      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = 4000 - intLength
'      Printer.CurrentY = 3250 + 2000 * intCounter
'      Printer.Print strAmount
'      Printer.CurrentX = 150
'      Printer.CurrentY = 3550 + 2000 * intCounter
'      Select Case adoacc0e0.Fields("a1p26").Value
'         Case "1"
'            Printer.Print ReportSum(42) & ComboItem(111)
'         Case "2"
'            Printer.Print ReportSum(42) & ComboItem(112)
'         Case "3"
'            Printer.Print ReportSum(42) & ComboItem(113)
'         Case "4"
'            Printer.Print ReportSum(42) & ComboItem(114)
'         Case "5"
'            Printer.Print ReportSum(42) & ComboItem(115)
'         Case "6"
'            Printer.Print ReportSum(42) & ComboItem(116)
'         Case "7"
'            Printer.Print ReportSum(42) & ComboItem(117)
'      End Select
'      intCounter = intCounter + 1
'      adoacc0e0.MoveNext
'   Loop
'   adoacc0e0.Close
'   Printer.EndDoc
'
'End Sub

'*************************************************
'  列印付款簽收簿 (抬頭及報表格式)
'
'*************************************************
Private Sub PrintHead()
Dim i As Integer

   Printer.FontSize = 14
   Printer.CurrentX = 3650
   Printer.CurrentY = 300
   Printer.Print A0802Query("1")
   Printer.CurrentX = 4000
   Printer.CurrentY = 800
   Printer.Print ReportTitle(1112)
   Printer.FontSize = 12
   Printer.CurrentX = 200
   Printer.CurrentY = 1100
   Printer.Print ReportSum(35) & CFDate(strSrvDate(2))
   Printer.CurrentX = 8200
   Printer.CurrentY = 1100
   Printer.Print ReportSum(36)
   Printer.CurrentX = 100
   Printer.CurrentY = 1600
   Printer.Print " 貴 寶 號"
   Printer.CurrentX = 1000
   Printer.CurrentY = 1600
   Printer.Print "    票           據           內            容"
   Printer.CurrentX = 5000
   Printer.CurrentY = 1600
   Printer.Print " 蓋                                      章"
   Printer.CurrentX = 8200
   Printer.CurrentY = 1600
   Printer.Print " 簽                  名"
   Printer.Line (100, 1400)-(10000, 1400)
   Printer.Line (100, 2000)-(10000, 2000)
   For i = 1 To 6
      Printer.Line (100, 2000 + 2000 * i)-(10000, 2000 + 2000 * i)
   Next i
   Printer.Line (100, 1400)-(100, 14000)
   Printer.Line (5000, 1400)-(5000, 14000)
   Printer.Line (8200, 1400)-(8200, 14000)
   Printer.Line (10000, 1400)-(10000, 14000)
End Sub

'*************************************************
'  列印地址條
'
'*************************************************
Private Sub PrintAddress()
   Dim intCounter As Integer
   Dim strName As String
   Dim intLine As Integer
   
   'Modify by Morgan 2008/8/15 +6(寄出地址特別)也要印
   'Modify by Amy 2014/01/27 +公司別
   'adoacc0e0.Open "select * from acc1p0, acc0e0, acc0q0 where a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02 and a1p02 = 'C' AND (A1P24='2' or A1P24='6') " & strSql & " order by a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Add by Amy 2020/07/06 +a0e07因改為key
   'Modify by Amy 2025/01/15 因為慢,調語法,並過濾票據資料才印
   'ex:1140114 X88492000 G11401030 因目前改以[電匯]方式,而畫面無此選項,故使用預設[支票],無票據資料不需印地址條-瑞婷
'   strExc(1) = "select * from acc1p0, acc0e0, acc0q0 where a1p01='" & strCmp & "' And a1p01=a0q19 " & _
'                    "And a1p01=a0e23 And a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02 and a1p11=a0e07 " & _
'                    "and a1p02 = 'C' AND (A1P24='2' or A1P24='6') " & strSql & " order by a1p04 asc"
   strExc(1) = "select * from acc1p0, acc0e0, acc0q0 where a0q19='" & strCmp & "' And a1p01(+)=a0q19 " & _
                    "And a1p01=a0e23(+) And a1p04(+) = a0q17 and a1p10 = a0e01(+) and a1p09 = a0e02(+) and a1p11=a0e07(+) " & _
                    "and a1p02(+) = 'C' AND (A1P24='2' or A1P24='6') " & strSql & " And a0e01 is not null order by a1p04 asc"
   If adoacc0e0.State <> adStateClosed Then adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
   'end 2014/01/27
   If adoacc0e0.RecordCount <> 0 Then
      Screen.MousePointer = vbDefault
      'Modify by Morgan 2006/11/10  加可取消
      'MsgBox MsgText(100) & ReportTitle(1113), , MsgText(5)
      If MsgBox(MsgText(100) & ReportTitle(1113), vbOKCancel, MsgText(5)) = vbCancel Then
         GoTo NoPrint
      End If
      'end 2006/11/10
      Screen.MousePointer = vbHourglass
   Else
NoPrint:
      adoacc0e0.Close
      Exit Sub
   End If
   
   'Add by Amy 2022/03/28
   With adoacc0e0
        Do While .EOF = False
            If strName <> "" & .Fields("a1p04") Then
                '                                                         地  址                                      票據抬頭
                strAddrData = strAddrData & .Fields("a0q16") & "$" & .Fields("a0q05") & "|"
                strName = "" & .Fields("a1p04")
            End If
            .MoveNext
        Loop
   End With
   
   PUB_SetOsDefaultPrinter Combo2
   PUB_RestorePrinter Combo2
   If PUB_XlsAccAddress(strAddrData) = False Then
        MsgBox "列印失敗！", vbCritical
   End If
   PUB_SetOsDefaultPrinter strPrinter2
   PUB_RestorePrinter strPrinter2
   'end 2022/03/15
   
   
   'Mark by amy 2022/03/28 改套印,以下不使用
'   intCounter = 0
'   'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
'   '9x
'   If pub_OS = "1" Then
'      Printer.Height = 2880
'      Printer.Width = 13000
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(2)
'   End If
'   'end 2008/3/25
'   'Modified by Lydia 2015/11/23 改為橫印
'   'Printer.Font = "@新細明體"
'   Printer.Font = "新細明體"
'   Printer.FontSize = 12
'
'   Do While adoacc0e0.EOF = False
'      If strName <> adoacc0e0.Fields("a1p04").Value Then
'         intLine = 0
'         If Not IsNull(adoacc0e0.Fields("a0q16").Value) Then
'            PUB_PrintAddress adoacc0e0.Fields("a0q16").Value, intCounter, intLine
'         End If
'         'Modify by Morgan 2007/10/17 下移兩格--瑞婷
'         'Printer.CurrentX = 100
'         Printer.CurrentX = 600
'         'end 2007/10/17
'         Printer.CurrentY = 1000 + 2200 * intCounter
'         If IsNull(adoacc0e0.Fields("a0q05").Value) Then
'            Printer.Print ""
'         Else
'            Printer.Print adoacc0e0.Fields("a0q05").Value & MsgText(104)
'         End If
'         Printer.NewPage
'
'      strName = adoacc0e0.Fields("a1p04").Value
'      End If
'      adoacc0e0.MoveNext
'   Loop
'   adoacc0e0.Close
'   Printer.Font = "新細明體"
'   Printer.EndDoc
   
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

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add by Morgan 2008/7/23
'退費收訖憑單
Private Sub PrintPayNotice3()
   Dim lngAmount As Long, lngAmt As Long
   Dim adoacc0s0 As New ADODB.Recordset
   'Modify by Amy 2014/01/27 +公司別
   'strExc(0) = "select * from acc0q0 where a0q04='2' and not exists(select * from acc1p0 where a1p04 = a0q17 and a1p02 = 'C' and a1p05 = '2111' and a1p08>0) " & strSql & " order by a0q17 asc"
   strExc(0) = "select * from acc0q0 where a0q19='" & strCmp & "' And a0q04='2' and not exists(select * from acc1p0 where a1p01=a0q19 And a1p04 = a0q17 and a1p02 = 'C' and a1p05 = '2111' and a1p08>0) " & strSql & " order by a0q17 asc"
   intI = 1
   Set adoacc0e0 = ClsLawReadRstMsg(intI, strExc(0))
   'Add by amy 2022/03/28 改開Word畫表格印
   If intI = 1 Then
        PUB_PrintReceipt_Doc Me.Name, "3.1", adoacc0e0, strCmp, "", m_NoMatchMsg, , True
        m_bolPrint = True
   End If
   Set adoacc0e0 = Nothing
   'end 2022/03/18
   
   'Mark by Amy 2022/03/28 改開Word畫表格印,以下不使用
'   If intI = 1 Then
'      Printer.EndDoc '回復印表機預設值
'      'Add by Morgan 2006/11/2
'      lngHalfHeight = Printer.Height / 2 '中一刀起始位置
'      lngYo = 0 '列印起始位置
'      lngPageNo = 0
'      'end 2006/11/2
'
'      With adoacc0e0
'      Do While Not .EOF
'         lngAmount = Val("" & adoacc0e0.Fields("a0q06").Value)
'         'Modify  by Amy 2014/01/27 +公司別
'         'strExc(0) = "select a0k01,a0k11,a0s06,a0s07,a0o01,a0s17,a0k04 from acc0o0, acc0s0, acc0k0 " & _
'         " where a0o03='" & .Fields("a0q03") & "' and a0o11=" & .Fields("a0q01") & " and a0s01(+)=a0o09 and a0o09 is not null and substr(a0s02,1,1)='E'" & _
'         " and a0k01(+)=a0s02 order by a0s02"
'         strExc(0) = "select a0k01,a0k11,a0s06,a0s07,a0o01,a0s17,a0k04 from acc0o0, acc0s0, acc0k0 " & _
'         " where a0o07='" & .Fields("A0q19") & "' And a0o03='" & .Fields("a0q03") & "' and a0o11=" & .Fields("a0q01") & " and a0s01(+)=a0o09 and a0o09 is not null and substr(a0s02,1,1)='E'" & _
'         " and a0k01(+)=a0s02 order by a0s02"
'         intI = 1
'         Set adoacc0s0 = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'         With adoacc0s0
'            Do While Not .EOF
'               'Modified by Morgan 2013/5/13 不用扣稅(同票據)
'               'lngAmt = Val("" & .Fields("a0s06")) + Val("" & .Fields("a0s07")) - Val("" & .Fields("a0s17"))
'               lngAmt = Val("" & .Fields("a0s06")) + Val("" & .Fields("a0s07"))
'
'               If .AbsolutePosition = .RecordCount Then
'                  'Modify by Morgan 2008/10/7 改彈訊息提醒
'                  'lngAmt = lngAmount
'                  'Modified by Morgan 2013/5/13
'                  'If lngAmt <> lngAmount Then
'                  If lngAmt - Val("" & .Fields("a0s17")) <> lngAmount Then
'                     m_NoMatchMsg = m_NoMatchMsg & vbCrLf & "<" & adoacc0e0("a0q03") & ">" & adoacc0e0("a0q05")
'                  End If
'               End If
'               PUB_PrintReceipt3 Me.adoacc0e0, adoacc0s0, lngYo, lngPageNo, lngAmt
'               m_bolPrint = True
'               'Modified by Morgan 2013/5/13
'               'lngAmount = lngAmount - lngAmt
'               lngAmount = lngAmount - (lngAmt - Val("" & .Fields("a0s17")))
'               .MoveNext
'            Loop
'         End With
'         End If
'         .MoveNext
'      Loop
'      End With
'      Printer.EndDoc
'   End If
'   Set adoacc0s0 = Nothing
'   Set adoacc0e0 = Nothing
End Sub

'Mark by Sindy 2020/4/23 公司別改下拉式選單
''Add by Amy 2014/01/27 +公司別
'Private Sub Text3_Change()
'    If Text3 = MsgText(601) Then
'        Text13 = ""
'        Exit Sub
'   End If
'   If Text3 = "1" Or Text3 = "J" Then
'        Text13 = A0802Query(Text3)
'   End If
'End Sub
'
'Private Sub Text3_GotFocus()
'    TextInverse Text3
'End Sub
'
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub Text3_Validate(Cancel As Boolean)
'    If Text3 = "" Then Exit Sub
'    If Text3 <> "1" And Text3 <> "J" Then
'        Text13 = ""
'        MsgBox "公司別輸入錯誤請確認 ！"
'        Cancel = True
'        Exit Sub
'    End If
'End Sub
''end 2014/01/27


