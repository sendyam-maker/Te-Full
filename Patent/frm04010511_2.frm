VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010511_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "消滅函／視為撤回輸入"
   ClientHeight    =   5748
   ClientLeft      =   -3840
   ClientTop       =   3300
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   8
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   7
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   6
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   5
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   6420
      TabIndex        =   4
      Top             =   660
      Width           =   1452
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm04010511_2.frx":0000
      Left            =   1080
      List            =   "frm04010511_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8376
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6324
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7152
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4212
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   7430
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4980
      TabIndex        =   12
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1020
      Width           =   768
   End
   Begin MSForms.Label Label8 
      Height          =   240
      Left            =   1800
      TabIndex        =   10
      Top             =   1050
      Width           =   6060
      VariousPropertyBits=   27
      Caption         =   "Label8"
      Size            =   "10689;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm04010511_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Label8)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 1
         frm04010511_1.Show
         Unload Me
      Case 2
         Unload frm04010511_1
         Unload Me
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
'Add By Cheng 2002/12/18
Dim rsA As New ADODB.Recordset
Dim StrSQLa  As String
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            strExc(5) = .TextMatrix(i, 3)
            strExc(4) = .TextMatrix(i, 9)
            'Add By Cheng 2002/12/18
            StrSQLa = "Select * From Patent Where " & ChgPatent(Me.Text2.Text & Me.Text3.Text & Me.Text4.Text & Me.Text5.Text)
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               'Modify by Morgan 2006/5/23 不必再控制公告號(93新法沒有)
               'If rsA("PA09").Value = "000" And (IsNull(rsA("PA14").Value) Or IsNull(rsA("PA15").Value)) Then
               '     MsgBox "本案件基本檔無公告日或公告號, 不可執行專利權消滅函輸入!!!", vbExclamation + vbOKOnly
               If rsA("PA09").Value = "000" And IsNull(rsA("PA14").Value) Then
                    MsgBox "本案件基本檔無公告日, 不可執行消滅函!!!", vbExclamation + vbOKOnly
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    Exit Sub
                End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If

   'Added by Morgan 2021/12/20
   '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
   If PUB_CheckFormExist("frm04010511_3") = False Then
      Set frm04010511_3 = Nothing
   End If
   'end 2021/12/20
   
   'Added by Morgan 2014/1/14
   frm04010511_3.m_AppNo = frm04010511_1.m_AppNo
   frm04010511_3.m_DocNo = frm04010511_1.m_DocNo
   frm04010511_3.m_DocWord = frm04010511_1.m_DocWord
   'end 2014/1/14
   'Add By Sindy 2016/10/5
   frm04010511_3.m_strIR01 = m_strIR01
   frm04010511_3.m_strIR02 = m_strIR02
   frm04010511_3.m_strIR03 = m_strIR03
   frm04010511_3.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010511_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      Case "日"
         Label8 = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010511_1.m_strIR01
   m_strIR02 = frm04010511_1.m_strIR02
   m_strIR03 = frm04010511_1.m_strIR03
   m_strIR04 = frm04010511_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   ReadPatent
   Combo1.ListIndex = 0
End Sub

' 90.07.05 只有一筆時直接跳下一畫面
Public Sub JumpIfOneRecord()
   If MSHFlexGrid1.Rows = 2 Then
      If IsEmptyText(MSHFlexGrid1.TextMatrix(1, 1)) = False Then
         MSHFlexGrid1.TextMatrix(1, 0) = "v"
         GridClick MSHFlexGrid1, 1, 0
         FormConfirm
      End If
   'Add By Cheng 2002/11/11
   '自動勾選該案號收文日最小的A類收文號資料, 直接進入輸入畫面
   ElseIf MSHFlexGrid1.Rows > 2 Then
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = "v"
         GridClick MSHFlexGrid1, MSHFlexGrid1.Rows - 1, 0
         FormConfirm
   End If
End Sub

Private Sub ReadPatent()
 Dim strTmp As String
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   Label8 = ""
   If pa(1) = "P" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         Label8 = pa(5)
         Text1 = pa(11)
      End If
   Else
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         Label8 = pa(5)
         Text1 = pa(11)
      End If
   End If
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   'Modify By Cheng 2002/03/05
   '案件性質為"準備程序"(211), "言詞辯論"(212)須另外處理,處理方式為:除非已取消收文日, 否則不管是否發文都要出現
   '其他案件性質則照原方式處理
'   strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
   'Modify By Cheng 2002/04/12
'   strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
' 91.09.13 modify by louis
'   strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
   'Modified by Lydia 2017/06/07 先抓非內部收文的新案進度,其他照原有規則
   'strExc(0) = "SELECT '',CP09," & strTmp & "," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
      "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
   strExc(0) = "SELECT '',CP09," & strTmp & "," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
      "CP10,decode(sign(instr('" & NewCasePtyList & "',cp10)),1,decode(cp27,19221111,99999999,11111111),DECODE(CP27,19221111,99999999,CP27)) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
      
   '處理案件性質為"準備程序"(211), "言詞辯論"(212)
   'Modify By Cheng 2002/04/12
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') AND ( CP05 IS NULL AND CP27 IS NOT NULL ) And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
' 91.09.13 modify by louis
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and ( cp09<'C' ) AND ( CP05 IS NULL AND CP27 IS NOT NULL ) And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
      "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and ( cp09<'C' ) AND ( CP05 IS NULL AND CP27 IS NOT NULL ) And cp01=cpm01(+) and " & _
      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
   'Modify By Cheng 2002/04/12
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
' 91.09.13 modify by louis
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and ( cp09<'C' ) And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
      strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
      "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and ( cp09<'C' ) And cp01=cpm01(+) and " & _
      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) " & _
      "ORDER BY SORTFIELD DESC,CP09 DESC "

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010511_2 = Nothing
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 600: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 2000: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .ColWidth(9) = 0
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub
