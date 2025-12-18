VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081033 
   BorderStyle     =   1  '單線固定
   Caption         =   "智財顧問案重新計算各部門實際比例"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7905
   Begin VB.CommandButton CmdDesc 
      Caption         =   "分配比例明細"
      Height          =   375
      Left            =   5445
      TabIndex        =   4
      Top             =   60
      Width           =   1425
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   3
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   2
      Top             =   555
      Width           =   495
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   2
      Left            =   2595
      MaxLength       =   1
      TabIndex        =   1
      Top             =   555
      Width           =   345
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   1
      Left            =   1695
      MaxLength       =   6
      TabIndex        =   0
      Top             =   555
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   0
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "ACS"
      Top             =   555
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2205
      Left            =   90
      TabIndex        =   6
      Top             =   1800
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   3889
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V|總收文號　|收文日期　|智權人員　|收文費用　|收文點數　|簽約時數　|顧問期間"
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
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   3
      Left            =   5880
      TabIndex        =   16
      Top             =   570
      Width           =   1335
      Size            =   "2355;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   2
      Left            =   5130
      TabIndex        =   15
      Top             =   555
      Width           =   705
      Size            =   "1244;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1500
      X2              =   3090
      Y1              =   720
      Y2              =   720
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   14
      Top             =   1275
      Width           =   6615
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   13
      Top             =   915
      Width           =   5535
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   0
      Left            =   1140
      TabIndex        =   12
      Top             =   915
      Width           =   885
      Size            =   "1561;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "目前智權人員："
      Height          =   225
      Index           =   4
      Left            =   3810
      TabIndex        =   10
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "當事人1："
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   975
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1290
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "frm081033"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/06/03 Form2.0已修改 lblFM2(index)、Combo1 ;  MSHFlexGrid1改字型=新細明體-ExtB
'Create by Lydia 2021/06/03 智財顧問案重新計算各部門實際比例
Option Explicit
Dim intLastRow As Integer '記錄勾選最後一筆
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset

Private Sub CmdDesc_Click()
Dim intX As Integer
Dim strCP09 As String

   If PUB_CheckFormExist("frm081031_3") Then
       MsgBox "請先關閉〔智財顧問專業分配比例〕畫面！"
       Exit Sub
   End If
   If lblFM2(0).Caption = "" Then
       MsgBox "請輸入本所案號，並且查詢資料！", vbInformation
       Exit Sub
   End If
   
   With MSHFlexGrid1
       For intX = 1 To .Rows - 1
          If .TextMatrix(intX, 0) = "v" And "" & .TextMatrix(intX, 1) <> "" Then
               strCP09 = "" & .TextMatrix(intX, 1)
               Exit For
          End If
       Next
   End With
   
   If strCP09 = "" Then
       MsgBox "請先選取一道收文！", vbInformation
   Else
        Call frm081031_3.SetParent(Me, strCP09, "U")
        Me.Hide
        frm081031_3.Show
   End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call doQuery(True)
End Sub

Private Sub doQuery(ByVal bolMsg As Boolean)

    If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
        MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
        Exit Sub
    End If
    
    Call ClearForm(False)
    If txtCase(2) = "" Then txtCase(2) = "0"
    If txtCase(3) = "" Then txtCase(3) = "00"
     
    strTmpQ = "select lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11 as custno,nvl(cu04,nvl(cu05,cu06)) custname " & _
                    "from lawcase,customer where lc01='" & txtCase(0) & "' and lc02='" & txtCase(1) & "' and lc03='" & txtCase(2) & "' and lc04='" & txtCase(3) & "' and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        Exit Sub
    End If
    
    intQ = 0
    Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
    If rsQuery.Fields("lc05") <> "" Then intQ = 1
    Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
    If rsQuery.Fields("lc06") <> "" Then intQ = 2
    Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
    If rsQuery.Fields("lc07") <> "" Then intQ = 3
    Combo1.ListIndex = intQ - 1
    
    lblFM2(0).Caption = "" & rsQuery.Fields("custno")
    lblFM2(1).Caption = "" & rsQuery.Fields("custname")
    strTmpQ = PUB_GetAKindSalesNo(txtCase(0), txtCase(1), txtCase(2), txtCase(3))
    lblFM2(2).Caption = strTmpQ
    lblFM2(3).Caption = GetPrjSalesNM(strTmpQ)
    
    Call SetGrd(True) '清空
    '以本所案號抓出所有顧問期間已過期的ACS之智財顧問112進度列示於Grid中
    strTmpQ = "select '' as V,cp09,substr(sqldatet(cp05),1,10) cp05t,st02,cp16,cp18,cp15,substr(sqldatet(cp53)||'~'||sqldatet(cp54),1,22) as caserange " & _
                     "from caseprogress,staff where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' " & _
                     "and cp10='112' and cp159=0 and cp13=st01(+) and cp54 is not null and cp54<=" & strSrvDate(1)
    strTmpQ = strTmpQ & "order by cp05,cp04 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 1 Then
         MSHFlexGrid1.FixedCols = 0
         Set MSHFlexGrid1.Recordset = rsQuery
         Call SetGrd
    Else
         If bolMsg = True Then MsgBox "查無顧問期間已過期的進度！", vbInformation
    End If
End Sub

Private Sub Form_Load()
 
    MoveFormToCenter Me
    Call ClearForm(True)
    Call SetGrd(True)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081033 = Nothing
End Sub

Private Sub ClearForm(ByVal bolResetCase As Boolean)
Dim oObj
    
    If bolResetCase = True Then
        For Each oObj In txtCase
            If oObj.Index > 0 Then
               oObj.Text = ""
            End If
        Next
    End If
    
    For Each oObj In lblFM2
       oObj.Caption = ""
    Next

    Combo1.Clear
    
End Sub

Private Sub MSHFlexGrid1_Click()

   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) <> "" Then
       GridClick MSHFlexGrid1, intLastRow, 0, 0
   End If
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If Me.MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "V" Then
        If m_blnColOrderAsc = True Then
           Me.MSHFlexGrid1.Sort = 5 '字串昇冪
           m_blnColOrderAsc = False
        Else
           Me.MSHFlexGrid1.Sort = 6 '字串降冪
           m_blnColOrderAsc = True
        End If
   End If
End Sub

Private Sub txtCase_GotFocus(Index As Integer)
    TextInverse txtCase(Index)
End Sub

Private Sub txtCase_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCase_LostFocus(Index As Integer)
    If Index > 1 And Trim(txtCase(Index)) = "" Then
        If Index = 2 Then
             txtCase(2) = "0"
        ElseIf Index = 3 Then
             txtCase(3) = "00"
        End If
    End If
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iR As Integer
Dim strTmp As String
 
   arrGridHeadText = Array("V", "總收文號", "收文日期", "智權人員", "收文費用", "收文點數", "簽約時數", "顧問期間")
   arrGridHeadWidth = Array(260, 1000, 900, 900, 900, 900, 900, 1650)
        
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
   For iRow = 0 To MSHFlexGrid1.Cols - 1
      MSHFlexGrid1.row = 0
      MSHFlexGrid1.col = iRow
      MSHFlexGrid1.Text = arrGridHeadText(iRow)
      MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
   Next

   MSHFlexGrid1.Visible = True
End Sub


