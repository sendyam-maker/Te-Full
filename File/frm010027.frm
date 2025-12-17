VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010027 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請書"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdok 
      Caption         =   "申請書(&A)"
      Height          =   345
      Index           =   1
      Left            =   6435
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   2
      Left            =   7560
      TabIndex        =   6
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5310
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3855
      Left            =   210
      TabIndex        =   7
      Top             =   1740
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V | 收文日 | 總收文號 | 案件性質 | 相關收文號 | 承辦人 | 智權人員| 本所期限 | 法定期限"
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
      _Band(0).Cols   =   9
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Top             =   465
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2670
      MaxLength       =   1
      TabIndex        =   2
      Top             =   465
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1770
      MaxLength       =   6
      TabIndex        =   1
      Top             =   465
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Top             =   465
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "備註：欲按申請書時，建議不要同時使用Word軟體，因程式執行中會使用到Word。"
      ForeColor       =   &H000000C0&
      Height          =   345
      Index           =   8
      Left            =   4050
      TabIndex        =   17
      Top             =   1380
      Width           =   4680
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   780
      Width           =   900
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   1170
      TabIndex        =   15
      Top             =   780
      Width           =   7680
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "申  請  人："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   900
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   1170
      TabIndex        =   13
      Top             =   1080
      Width           =   7680
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1380
      Width           =   900
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   1170
      TabIndex        =   11
      Top             =   1380
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "申請案號："
      Height          =   255
      Index           =   4
      Left            =   4740
      TabIndex        =   10
      Top             =   480
      Width           =   900
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   5670
      TabIndex        =   9
      Top             =   480
      Width           =   1740
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1500
      X2              =   3060
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   510
      Width           =   900
   End
End
Attribute VB_Name = "frm010027"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 grd1/Label1(1)/Label1(3)/Label1(5)/Label1(7)
'Create By Sindy 2013/6/19
Option Explicit

Dim m_row As Integer, i As Integer
Dim strCP01 As String
Dim strCP02 As String
Dim strCP03 As String
Dim strCP04 As String
Dim m_Nation As String


Private Sub cmdOK_Click(Index As Integer)
Dim strUpdDate As String, strUpdTime As String
Dim Cancel As Boolean
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   Select Case Index
      Case 0 '查詢
         If Trim(txt1(0)) = "" Or Trim(txt1(1)) = "" Then
            MsgBox "本所案號不可空白！", vbCritical, "操作錯誤！"
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         doQuery
         Screen.MousePointer = vbDefault
      Case 1 '申請書
         If GRD1.TextMatrix(m_row, 0) = "V" Then
            If GRD1.TextMatrix(m_row, 9) = "102" Or _
               GRD1.TextMatrix(m_row, 9) = "103" Then
               'Modified by Lydia 2019/03/28 傳入收文號
               'Call PUB_GetApplBook(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04, grd1.TextMatrix(m_row, 9))
               Call PUB_GetApplBook(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04, GRD1.TextMatrix(m_row, 9), , , , , , "" & GRD1.TextMatrix(m_row, 2))
            ElseIf GRD1.TextMatrix(m_row, 9) = "301" Or _
               GRD1.TextMatrix(m_row, 9) = "501" Then
               frm090201_b_3.m_CP10 = Trim(GRD1.TextMatrix(m_row, 9))
               frm090201_b_3.lbl1(3).Caption = Trim(GRD1.TextMatrix(m_row, 2))
               frm090201_b_3.lbl1(7).Caption = strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04
               frm090201_b_3.lbl1(9).Caption = Label1(1).Caption
               frm090201_b_3.lbl1(15).Caption = Trim(GRD1.TextMatrix(m_row, 3))
               frm090201_b_3.Show vbModal
            Else
               MsgBox "目前只有（延展,變更,移轉,補換發註冊證）有申請書!!!"
            End If
         End If
      Case 2 '結束
         Unload Me
   Case Else
   End Select
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Sub

'清除欄位值
Sub ClearData()
   Label1(7) = ""
   Label1(1) = ""
   Label1(3) = ""
   Label1(5) = ""
   m_Nation = ""
   m_row = 0
End Sub

Sub doQuery()
Dim Cancel As Boolean
   
On Error GoTo ErrHnd
   
   Cancel = False
   Call txt1_Validate(0, Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   
   strCP01 = UCase(txt1(0))
   strCP02 = txt1(1)
   strCP03 = Left(txt1(2) & "0", 1)
   strCP04 = Left(txt1(3) & "00", 2)
   
   '清除欄位值
   Call ClearData
   strSql = "SELECT TM12,TM05||TM06||TM07,TM23||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,TM10,TM11,NA01" & _
                " From trademark, nation, Customer" & _
                " WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "'" & _
                " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
                " AND TM10=NA01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Label1(7) = "" & Trim(RsTemp(0))
      Label1(1) = "" & Trim(RsTemp(1))
      Label1(3) = "" & Trim(RsTemp(2))
      Label1(5) = "" & Trim(RsTemp(3))
      m_Nation = "" & Trim(RsTemp("NA01"))
   End If
   
   strSql = "SELECT ' ' as V,sqldatet(CP05) as 收文日,CP09 as 總收文號,DECODE('" & m_Nation & "','000',cpm03,cpm04) as 案件性質,CP43 as 相關收文號,s1.st02 as 承辦人," & _
            "s2.st02 as 智權人員,sqldatet(CP06) as 本所期限,sqldatet(CP07) as 法定期限,cp10 " & _
            "From CaseProgress,casepropertymap,staff s1,staff s2 " & _
            "WHERE CP01='" & strCP01 & "' and CP02='" & strCP02 & "' and CP03='" & strCP03 & "' and CP04='" & strCP04 & "' " & _
            "AND CP01=cpm01(+) AND CP10=cpm02(+) " & _
            "AND CP14=s1.st01(+) " & _
            "AND CP13=s2.st01(+) " & _
            "AND CP57 is null " & _
            "order by CP05 desc"
   CheckOC3
   GRD1.Rows = 2
   GRD1.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set GRD1.Recordset = AdoRecordSet3.Clone
         '選取第一筆資料
         GRD1.Visible = False
         GRD1.col = 0
         GRD1.row = 1
'         m_row = 1
'         grd1.TextMatrix(m_row, 0) = "V"
'         For i = 0 To grd1.Cols - 1
'            grd1.col = i
'            grd1.CellBackColor = &HFFC0C0
'         Next i
         GRD1.Visible = True
      Else
         MsgBox "無符合資料！", vbInformation
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010027 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim m_mouseRow As Integer
   
   GRD1.Visible = False
   m_mouseRow = GRD1.MouseRow
   GRD1.col = 0
   If m_mouseRow > 0 Then
      If m_row > 0 Then
         GRD1.row = m_row
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            If GRD1.CellBackColor = &HFFC0C0 Then
               GRD1.CellBackColor = &H80000018
               GRD1.TextMatrix(m_row, 0) = ""
            Else
               GRD1.CellBackColor = &HFFC0C0
               GRD1.TextMatrix(m_row, 0) = "V"
            End If
         Next i
      End If
      GRD1.row = m_mouseRow
      m_row = m_mouseRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         If GRD1.CellBackColor = &HFFC0C0 Then
            GRD1.CellBackColor = &H80000018
            GRD1.TextMatrix(m_row, 0) = ""
            m_row = 0
         Else
            GRD1.CellBackColor = &HFFC0C0
            GRD1.TextMatrix(m_row, 0) = "V"
         End If
      Next i
   End If
   GRD1.Visible = True
End Sub

Private Sub SetDataListWidth()
GRD1.Visible = False
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, m_i As Integer
   
   arrGridHeadText = Array("V", "收文日", "總收文號", "案件性質", "相關收文號" _
                         , "承辦人", "智權人員", "本所期限", "法定期限", "CP10")
   arrGridHeadWidth = Array(200, 900, 1000, 1200, 1000 _
                          , 1000, 1000, 900, 900, 0)
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      If iRow > 10 Then
         GRD1.ColWidth(iRow) = 0
      Else
         GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      End If
      GRD1.CellAlignment = flexAlignLeftCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If Index <> 0 Then Exit Sub
   
   Cancel = False
   If txt1(Index).Text <> "" Then
      If txt1(Index).Text <> "T" And txt1(Index).Text <> "FCT" Then
         MsgBox "系統別只可輸入T或FCT"
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
