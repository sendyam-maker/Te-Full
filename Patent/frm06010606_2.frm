VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010606_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "異議/舉發受理函輸入"
   ClientHeight    =   3492
   ClientLeft      =   1116
   ClientTop       =   1140
   ClientWidth     =   7332
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3492
   ScaleWidth      =   7332
   Begin VB.TextBox Text23 
      Height          =   270
      Left            =   1632
      MaxLength       =   3
      TabIndex        =   14
      Top             =   3096
      Width           =   684
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm06010606_2.frx":0000
      Left            =   1140
      List            =   "frm06010606_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   20
      Top             =   1140
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6372
      TabIndex        =   19
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4320
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5148
      TabIndex        =   17
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2700
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1140
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2340
      Width           =   5655
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   4
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   3
      Top             =   780
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   2
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   1
      Top             =   780
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4380
      TabIndex        =   0
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "對造案件數代號:"
      Height          =   252
      Left            =   216
      TabIndex        =   25
      Top             =   3096
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   180
      X2              =   7140
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   7140
      Y1              =   2250
      Y2              =   2250
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1140
      TabIndex        =   24
      Top             =   1860
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   4380
      TabIndex        =   23
      Top             =   1500
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   1140
      TabIndex        =   22
      Top             =   1500
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   780
      Width           =   768
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "(1:受理 2:對方延期)"
      Height          =   180
      Left            =   1620
      TabIndex        =   16
      Top             =   2700
      Width           =   1512
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "程序:"
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   2700
      Width           =   408
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   2340
      Width           =   768
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   10
      Top             =   1140
      Width           =   5460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "9631;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "對造號數:"
      Height          =   180
      Left            =   3540
      TabIndex        =   9
      Top             =   780
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   1140
      Width           =   768
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   1500
      Width           =   768
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   3540
      TabIndex        =   6
      Top             =   1500
      Width           =   588
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   1860
      Width           =   948
   End
End
Attribute VB_Name = "frm06010606_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/23 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strSales As String
Dim cp(10) As String
Dim intWhere As Integer
Dim pa(1 To 4) As String
Dim m_928Upd As Boolean '是否更新重新委任准駁
Dim m_928CP09 As String '重新委任收文號
Dim m_PA75 As String 'Added by Morgan 2012/11/5
'Added by Morgan 2017/5/10 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
Dim m_PA26 As String, m_PA27 As String, m_PA28 As String, m_PA29 As String, m_PA30 As String
'end 2017/5/10


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If Text6 = "" Then
            MsgBox "受理程序不可空白 !", vbCritical
            Text6.SetFocus
            Exit Sub
         End If
         
         'Added by Morgan 2024/12/2
         If Text23.Text = "" Then
            MsgBox "請輸入對造案件數代號!!!", vbInformation
            Text23.SetFocus
            Text23_GotFocus
            Exit Sub
         '台灣對造案件數代號第一碼必須為 N
         ElseIf Left(Text23.Text, 1) <> "N" Then
            MsgBox "對造案件數代號第一碼必須為 N !!!", vbInformation
            Text23.SetFocus
            Text23_GotFocus
            Exit Sub
         End If
         'end 2024/12/2
            
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Added by Morgan 2012/11/5
         If Left(m_PA75, 6) = "Y53309" Then
            MsgBox "本案需調卷轉承辦組報告並寄代！", vbInformation
         End If
         'end 2012/11/5
         
         Unload Me
         
         'Modified by Morgan 2017/5/10 電子公文
         'frm06010606_1.Show
         If m_DocNo <> "" Then
            Unload frm06010606_1
            frm060119.GoNext
         Else
            frm06010606_1.Show
         End If
         'end 2017/5/10
         
      Case 1
         frm06010606_1.Show
         Unload Me
      Case 2
         Unload frm06010606_1
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean

   Dim i As Integer, strTmp(1 To 7) As String
   Dim strCP20 As String, strCP16 As String, stCP09 As String
   Dim stCP10 As String 'Added by Morgan 2017/5/10
   
   pa(1) = Text2
   pa(2) = Text3
   pa(3) = Text4
   pa(4) = Text5
   
   m_928Upd = PUB_928Check(pa, m_928CP09) 'Add by Morgan 2007/7/18

   cnnConnection.BeginTrans
   
On Error GoTo CheckingErr
      
    'Added by Morgan 2024/12/2
    '更新原進度檔資料的對造號數
    strSql = "Update CaseProgress Set CP36='" & Text1 & Text23 & "' Where CP09='" & strReceiveNo & "' "
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql, intI
    'end 2024/12/2
    
   
   'Add by Morgan 2007/7/18
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09
   End If
   'end 2007/7/18
 
   strExc(0) = "SELECT CP36,CP37,CP38,CP39,CP40,CP41,CP42 FROM CASEPROGRESS WHERE CP09='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 1 To 7
         If Not IsNull(RsTemp.Fields(i - 1)) Then strTmp(i) = RsTemp.Fields(i - 1)
      Next
   End If
   If Text6 = "1" Then
      i = 爭議受理
   ElseIf Text6 = "2" Then
      i = 對方延期
   End If
   
   stCP10 = i 'Added by Morgan 2017/5/10
    
   '智權人員存國家檔FCP承辦智權人員
   stCP09 = AutoNo("C", 6)
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP13," & _
      "CP12,CP14,CP20,CP26,CP32,CP27,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) VALUES " & _
      "('" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "'," & TransDate(Label2(3), 2) & _
      "," & CNULL(ChgSQL(Text7)) & ",'" & stCP09 & "'," & _
      CNULL(Format(i)) & "," & CNULL(PUB_GetFCPSalesNo(Text2.Text, Text3.Text, Text4.Text, Text5.Text)) & "," & CNULL(cp(5)) & ",'" & strUserNum & "'," & _
      "'N','N','N'," & strSrvDate(1) & "," & CNULL(ChgSQL(strTmp(1))) & "," & _
      CNULL(ChgSQL(strTmp(2))) & "," & CNULL(ChgSQL(strTmp(3))) & "," & CNULL(ChgSQL(strTmp(4))) & "," & _
      CNULL(ChgSQL(strTmp(5))) & "," & CNULL(ChgSQL(strTmp(6))) & "," & CNULL(ChgSQL(strTmp(7))) & ",'" & strReceiveNo & "')"
      
   cnnConnection.Execute strSql, intI
   
   'Modify by Morgan 2007/7/23 CP20改抓CPM的設定
   'Modify by Morgan 2008/3/27 +pa75
   'Modify by Morgan 2008/4/10 +本所案號
   'Modified by Morgan 2017/5/10
   'strCP20 = PUB_GetCP20(Text2, Format(i), strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   strCP20 = PUB_GetCP20(Text2, Format(i), strCP16, m_PA26 & m_PA27 & m_PA28 & m_PA29 & m_PA30, m_PA75, pa(1) & pa(2) & pa(3) & pa(4))
   If strCP20 = "" Then
      strSql = "update caseprogress set cp20=NULL,cp16=" & strCP16 & ",cp17=0,cp18=" & strCP16 / 1000 & _
         " where cp09='" & stCP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2007/7/23
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, stCP09, pa(1), pa(2), pa(3), pa(4), stCP10
   'Added by Morgan 2021/6/11 紙本公文--何淑華
   Else
      PUB_FCPOAInform stCP09, pa(1), pa(2), pa(3), pa(4), stCP10
   End If
   'end 2017/5/10
   
   cnnConnection.CommitTrans
   
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(0) = cp(1)
      Case "英"
         Label2(0) = cp(2)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(0) = cp(3)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Text2 = strExc(1)
   Text3 = strExc(2)
   Text4 = strExc(3)
   Text5 = strExc(4)
   strReceiveNo = strExc(5)
   ReadPatent
   Combo1.ListIndex = 0
Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   Text7.Text = "（" & strTmp & "）智專一（二）字第號"
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         Text7 = m_DocWord & "字第" & m_DocNo & "號"
      ElseIf m_DocNo <> "" Then
         Text7 = Replace(Text7, "第號", "第" & m_DocNo & "號")
      End If
   End If
   'end 2017/5/10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2021/6/11
   Set frm06010606_2 = Nothing
End Sub

Private Sub ReadPatent()
 'Modified by Morgan 2024/12/2 Label2已改form2.0 變數要改不定型態才不會錯
 'Dim Lbl As LABEL, i As Integer, strTempName As String
 Dim Lbl
 Dim i As Integer, strTempName As String
 
   For Each Lbl In Label2
      Lbl = ""
   Next
   'Add By Cheng 2002/07/17
   Erase cp
   
   Label2(2) = strReceiveNo
   Label2(3) = frm06010606_1.Text5
   strExc(0) = "SELECT CP36,CP37,CP38,CP39,CP13,CP12,CPM03,PA75,pa26,pa27,pa28,pa29,pa30 FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT WHERE CP09='" & strReceiveNo & "'" & _
      " AND CP01=CPM01(+) AND CP10=CPM02(+) and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         'Modified by Morgan 2024/12/2 改和內專一樣,對造號不帶NXX另放於代號欄位
         'If Not IsNull(.Fields(0)) Then Text1 = .Fields(0)
         If Left(Right("" & .Fields(0), 3), 1) = "N" Then
            Text1 = Left(.Fields(0), Len(.Fields(0)) - 3)
            Text23 = Right("" & .Fields(0), 3)
         Else
            Text1 = "" & .Fields(0)
         End If
         If m_AppNo <> "" Then Text23 = Right(m_AppNo, 3)
         'end 2024/12/2
         For i = 1 To 5
            If Not IsNull(.Fields(i)) Then cp(i) = .Fields(i)
         Next
         If Not IsNull(.Fields(6)) Then Label2(1) = .Fields(6)
         m_PA75 = "" & .Fields("pa75") 'Added by Morgan 2012/11/5
         'Added by Morgan 2017/5/10
         m_PA26 = "" & .Fields("pa26")
         m_PA27 = "" & .Fields("pa27")
         m_PA28 = "" & .Fields("pa28")
         m_PA29 = "" & .Fields("pa29")
         m_PA30 = "" & .Fields("pa30")
         'end 2017/5/10
      End If
   End With
End Sub

Private Sub Text23_GotFocus()
    '欄位值反白
    TextInverse Me.Text23
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
    '轉換為大寫
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text7_GotFocus()
'   TextInverse Text7
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text7.IMEMode = 1
   OpenIme
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'將游標設定在機關文號欄的"專"的後面
With Me.Text7
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "專")
      If intPos > 0 Then
         .SelStart = intPos
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub Text7_LostFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text7.IMEMode = 2
   CloseIme
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If CheckLengthIsOK(Text7, Text7.MaxLength) = False Then
      Cancel = True
      Text7.SetFocus
      Text7_GotFocus
   End If
End Sub
