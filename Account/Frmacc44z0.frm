VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc44z0 
   AutoRedraw      =   -1  'True
   Caption         =   "會計師資料／客戶E-Mail資料查詢"
   ClientHeight    =   5292
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8688
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5292
   ScaleWidth      =   8688
   Begin VB.TextBox textA4905 
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
      Left            =   2055
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1710
      Width           =   6555
   End
   Begin VB.TextBox textA4903 
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
      Left            =   2055
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdCall 
      BackColor       =   &H00C0FFC0&
      Caption         =   "扣繳憑單查詢及列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5340
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   2250
      Width           =   2505
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrd1 
      Height          =   2565
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   8565
      _ExtentX        =   15113
      _ExtentY        =   4530
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V|事務所|會計師|電話|會計師信箱|客戶代號|客戶名稱／特殊抬頭|客戶／特殊抬頭E-Mail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6180
      TabIndex        =   1
      Top             =   210
      Width           =   765
   End
   Begin VB.OptionButton Option1 
      Caption         =   "會計師姓名："
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
      Index           =   0
      Left            =   300
      TabIndex        =   11
      Top             =   270
      Width           =   1755
   End
   Begin VB.OptionButton Option1 
      Caption         =   "      E-mail："
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
      Left            =   300
      TabIndex        =   12
      Top             =   1770
      Value           =   -1  'True
      Width           =   1755
   End
   Begin MSForms.TextBox textA4902 
      Height          =   345
      Left            =   2055
      TabIndex        =   0
      Top             =   210
      Width           =   2175
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "3836;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4912 
      Height          =   345
      Left            =   2055
      TabIndex        =   2
      Top             =   690
      Width           =   4005
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "7064;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(含客戶、特殊抬頭檔)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   2070
      TabIndex        =   10
      Top             =   2100
      Width           =   2040
   End
   Begin VB.Label lblA4901_C 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3420
      TabIndex        =   7
      Top             =   600
      Width           =   75
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   120
      Top             =   2490
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "事務所："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1065
      TabIndex        =   8
      Top             =   750
      Width           =   960
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "會計師電話："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   585
      TabIndex        =   6
      Top             =   1260
      Width           =   1440
   End
End
Attribute VB_Name = "Frmacc44z0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/17 Form2.0已修改
'Create by Lydia 2016/12/19 會計師客戶資料查詢
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim colA4901_1 As Integer '客戶代號
Dim colA4901_2 As Integer '收據抬頭


Private Sub cmdCall_Click()
Dim tmpList As String
Dim ii As Integer
   
    For ii = 1 To MGrd1.Rows - 1
       If MGrd1.TextMatrix(ii, 0) = "v" Then
          If Trim("" & MGrd1.TextMatrix(ii, colA4901_1)) <> "" Then
             '客戶代號|客戶名稱
             tmpList = tmpList & "1," & Trim(MGrd1.TextMatrix(ii, colA4901_1)) & "|" & Trim(MGrd1.TextMatrix(ii, colA4901_1 + 1)) & ";"
          ElseIf Trim("" & MGrd1.TextMatrix(ii, colA4901_2)) <> "" Then
             '收據抬頭
             tmpList = tmpList & "2," & Trim(MGrd1.TextMatrix(ii, colA4901_2)) & ";"
          End If
          '因為扣繳憑單查詢沒有回前畫面的功能,所以直接回復未勾選
          MGrd1.TextMatrix(ii, 0) = ""
       End If
    Next ii
    
    PUB_SetMSFGridColor Me.MGrd1, 15
    
    If tmpList = "" Then
       MsgBox "請勾選客戶資料!"
       Exit Sub
    Else
        Me.MousePointer = vbHourglass
        strUserLevel = Me.Name
        Frmacc44t0.Show
        Call Frmacc44t0.CallByA4901(tmpList)
        Me.Hide
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdQuery_Click()
Dim Cancel As Boolean

   With Frmacc44z0
      If .textA4902 = MsgText(601) And _
         .textA4912 = MsgText(601) And _
         .textA4903 = MsgText(601) And _
         .textA4905 = MsgText(601) Then
         MsgBox "會計師姓名、事務所、會計師電話、E-Mail至少要輸入一項！", , MsgText(5)
         strControlButton = MsgText(602)
         If .textA4902.Enabled = True Then .textA4902.SetFocus
         Exit Sub
      End If

  End With
  
  If QueryData = True Then
     cmdCall.Enabled = True
  Else
     cmdCall.Enabled = False
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()

   '表單初始化
   PUB_InitForm Me, 8800, 5700, strBackPicPath1
   tool3_enabled
   
   Frmacc44z0_Clear
   
   Option1(1).Value = True
   Call Option1_Click(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
       Cancel = 1
       Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
  
   strUserLevel = MsgText(601)
   Set Frmacc44z0 = Nothing
End Sub

Private Function QueryData() As Boolean
Dim rsAD As New ADODB.Recordset
Dim inA As Integer
Dim strMid As String
    
    'Modfiy by Amy 2024/12/25 避免查'(單引號)出現錯誤+ChgSQL
    '會計師姓名
    If textA4902 <> "" Then
       strMid = strMid & IIf(strMid <> "", "AND", "") & " INSTR(UPPER(A4902),'" & ChgSQL(UCase(textA4902)) & "') > 0 "
    End If
    '事務所
    If textA4912 <> "" Then
       strMid = strMid & IIf(strMid <> "", "AND", "") & " INSTR(UPPER(A4912),'" & ChgSQL(UCase(textA4912)) & "') > 0 "
    End If
    '會計師電話
    If textA4903 <> "" Then
       strMid = strMid & IIf(strMid <> "", "AND", "") & " INSTR(UPPER(A4903),'" & ChgSQL(UCase(textA4903)) & "') > 0 "
    End If
    'E-mail
    If textA4905 <> "" Then
       strMid = strMid & IIf(strMid <> "", "AND", "") & " INSTR(UPPER(A4905),'" & ChgSQL(UCase(textA4905)) & "') > 0 "
    End If
    
    Screen.MousePointer = vbHourglass
    'Modified by Lydia 2024/09/18 財務副本信箱CU200
    strSql = "SELECT '' v,A4912,A4902,A4903,A4905,CU01||CU02 CNO,DECODE(CU01,NULL,A4901,NVL(CU04,NVL(CU05,CU06))) CTITLE,decode(cu20,null,'',cu20||';')||decode(cu115,null,'',cu115||';')||decode(cu116,null,'',cu116||';')||decode(cu117,null,'',cu117||';')||decode(cu118,null,'',cu118||';')||decode(cu200,null,'',cu200||';') " & _
             "FROM ACC490,CUSTOMER WHERE (" & strMid & _
             ") AND SUBSTR(A4901,1,8)=CU01(+) AND SUBSTR(A4901,9,1)=CU02(+)"
    'Add By Sindy 2017/6/28
    '含客戶、特殊抬頭檔的E-mail
    
    If textA4905 <> "" Then
      'Modified by Lydia 2024/09/18 財務副本信箱CU200
      strSql = strSql & " union " & _
                        "SELECT '' v,A4912,A4902,A4903,A4905,CU01||CU02 CNO,NVL(CU04,NVL(CU05,CU06)) CTITLE,decode(cu20,null,'',cu20||';')||decode(cu115,null,'',cu115||';')||decode(cu116,null,'',cu116||';')||decode(cu117,null,'',cu117||';')||decode(cu118,null,'',cu118||';')||decode(cu200,null,'',cu200||';') " & _
                        "FROM ACC490,CUSTOMER WHERE INSTR(UPPER(cu20)||UPPER(cu115)||UPPER(cu116)||UPPER(cu117)||UPPER(cu118)||UPPER(cu200),'" & ChgSQL(UCase(textA4905)) & "') > 0 " & _
                        "AND SUBSTR(A4901(+),1,8)=CU01 AND SUBSTR(A4901(+),9,1)=CU02"
      strSql = strSql & " union " & _
                        "SELECT '' v,A4912,A4902,A4903,A4905,'' CNO,A4201 CTITLE,A4218 " & _
                        "FROM ACC490,ACC420 WHERE INSTR(UPPER(a4218),'" & ChgSQL(UCase(textA4905)) & "') > 0 " & _
                        "AND A4901(+)=A4201"
    End If
    strSql = strSql & " ORDER BY A4912,A4902,CNO,CTITLE"
    '2017/6/28 END
    'end 2024/12/25
    inA = 0
    Set rsAD = ClsLawReadRstMsg(inA, strSql)
    MGrd1.FixedCols = 0
    If inA = 1 Then
       QueryData = True
       Set MGrd1.Recordset = rsAD
       SetGrd (rsAD.RecordCount + 1)
    Else
       QueryData = False
       Set MGrd1.Recordset = rsAD
       SetGrd
    End If
    MGrd1.FixedCols = 3
    
    Screen.MousePointer = vbDefault
    Set rsAD = Nothing
End Function

Private Sub Frmacc44z0_Clear()
   With Frmacc44z0
      .textA4902 = ""
      .textA4903 = ""
      .textA4905 = ""
      .textA4912 = ""
      SetGrd
      cmdCall.Enabled = False
   End With
End Sub

Private Sub mgrd1_Click()
Dim intRow As Integer

   If MGrd1.row > 0 Then
      intRow = MGrd1.row
      GridClick MGrd1, intRow, 0
   End If
End Sub

Private Sub MGrd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MGrd1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrd1.col = nCol
   MGrd1.row = nRow
   If Me.MGrd1.row < 1 And Me.MGrd1.Text <> "V" Then

      If Me.MGrd1.Text = "電話" Then
         If m_blnColOrderAsc = True Then
            Me.MGrd1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrd1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGrd1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrd1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub Mgrd1_SelChange()
'Dim tmpRow As Integer
'tmpRow = MGrd1.MouseRow
'
'If tmpRow > 0 Then
'   GridClick Me.MGrd1, tmpRow, 0
'End If
End Sub

'Add By Sindy 2017/6/29
Private Sub Mgrd1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MGrd1.ToolTipText = ""
If MGrd1.MouseRow <> 0 And MGrd1.MouseCol > 0 Then
   If MGrd1.MouseCol = 7 And MGrd1.TextMatrix(MGrd1.MouseRow, 7) <> "" Then
      MGrd1.ToolTipText = MGrd1.TextMatrix(MGrd1.MouseRow, 7)
   ElseIf MGrd1.MouseCol = 6 And MGrd1.TextMatrix(MGrd1.MouseRow, 6) <> "" Then
      MGrd1.ToolTipText = MGrd1.TextMatrix(MGrd1.MouseRow, 6)
   End If
End If
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      textA4902.Enabled = True
      textA4912.Enabled = True
      textA4903.Enabled = True
      textA4905.Enabled = False: textA4905 = ""
      If textA4902.Visible = True Then textA4902.SetFocus
   Else 'E-Mail
      textA4902.Enabled = False: textA4902 = ""
      textA4912.Enabled = False: textA4912 = ""
      textA4903.Enabled = False: textA4903 = ""
      textA4905.Enabled = True
      If textA4905.Visible = True Then textA4905.SetFocus
   End If
End Sub

Private Sub textA4902_GotFocus()
   InverseTextBox textA4902
   OpenIme
End Sub

Private Sub textA4903_GotFocus()
   CloseIme
   TextInverse textA4903
End Sub

Private Sub textA4905_GotFocus()
   CloseIme
   TextInverse textA4905
End Sub


Private Sub textA4912_GotFocus()
   InverseTextBox textA4912
   OpenIme
End Sub

Private Sub SetGrd(Optional ByVal iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("v", "事務所", "會計師", "電話", "會計師信箱", "客戶代號", "名稱／特殊抬頭", "客戶／特殊抬頭E-Mail")
   arrGridHeadWidth = Array(200, 1000, 900, 900, 1000, 1000, 1400, 1600)
   MGrd1.Visible = False
   MGrd1.Cols = UBound(arrGridHeadText) + 1
   MGrd1.Rows = iR
   For iRow = 0 To MGrd1.Cols - 1
      MGrd1.row = 0
      MGrd1.col = iRow
      MGrd1.Text = arrGridHeadText(iRow)
      MGrd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrd1.CellAlignment = flexAlignCenterCenter
   Next
   
   If colA4901_1 = 0 Then
      colA4901_1 = PUB_MGridGetId("客戶代號", MGrd1)
      colA4901_2 = PUB_MGridGetId("名稱／特殊抬頭", MGrd1)
   End If
   
   PUB_SetMSFGridColor Me.MGrd1, 15
   MGrd1.Visible = True
End Sub
