VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04010504_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函輸入"
   ClientHeight    =   5745
   ClientLeft      =   -3840
   ClientTop       =   3300
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
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
      ItemData        =   "frm04010504_2.frx":0000
      Left            =   1080
      List            =   "frm04010504_2.frx":000D
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
      _ExtentX        =   16007
      _ExtentY        =   7435
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   240
      Left            =   1800
      TabIndex        =   10
      Top             =   1050
      Width           =   6060
   End
End
Attribute VB_Name = "frm04010504_2"
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
         frm04010504_1.Show
         Unload Me
      Case 2
         Unload frm04010504_1
         Unload Me
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            strExc(5) = .TextMatrix(i, 3)
            'Modify By Cheng 2003/01/27
            '因加對造號數欄故位移一位
'            strExc(4) = .TextMatrix(i, 9)
            strExc(4) = .TextMatrix(i, 10)
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
   If PUB_CheckFormExist("frm04010504_3") = False Then
      Set frm04010504_3 = Nothing
   End If
   'end 2021/12/20
   
   'Added by Morgan 2014/1/14
   frm04010504_3.m_AppNo = frm04010504_1.m_AppNo
   frm04010504_3.m_DocNo = frm04010504_1.m_DocNo
   frm04010504_3.m_DocWord = frm04010504_1.m_DocWord
   frm04010504_3.m_DeadLine = frm04010504_1.m_DeadLine
   frm04010504_3.m_NewCP10 = frm04010504_1.m_NewCP10
   'end 2014/1/14
   'Add By Sindy 2016/10/5
   frm04010504_3.m_strIR01 = m_strIR01
   frm04010504_3.m_strIR02 = m_strIR02
   frm04010504_3.m_strIR03 = m_strIR03
   frm04010504_3.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010504_3.Show
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
   m_strIR01 = frm04010504_1.m_strIR01
   m_strIR02 = frm04010504_1.m_strIR02
   m_strIR03 = frm04010504_1.m_strIR03
   m_strIR04 = frm04010504_1.m_strIR04
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
   End If
End Sub

Private Sub ReadPatent()
Dim strTmp As String
'Add By Cheng 2002/06/28
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
Dim strSQLCondition As String '1:CP27 IS NOT NULL AND CP09<'C'; 2:CP57 IS NULL AND CP09<'C'
Dim ii As Integer

   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   Label8 = ""
   If pa(1) = "P" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         Label8 = pa(5)
         Text1 = pa(11)
      End If
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         Label8 = pa(5)
         Text1 = pa(11)
      End If
   End If
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   
   'Add By Cheng 2002/06/28
   strSQLCondition = "1"
   '搜尋案件性質為"延期"(404)且有"發文日"
   StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='404' AND CP27 IS NOT NULL "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Do While Not rsA.EOF
         '以上述資料的總收文號搜尋相關總收文號及案件性質為"延期受理"的資料
         StrSqlB = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='1004' AND CP43='" & rsA.Fields("CP09") & "'"
         rsB.CursorLocation = adUseClient
         rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
         If rsB.RecordCount <= 0 Then
            strSQLCondition = "2"
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            Exit Do
         End If
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
         rsA.MoveNext
      Loop
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   'Modify By Cheng 2002/06/27
   '多顯示相關總收文號
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
' 91.09.13 modify by louis (排序)
'   If strSQLCondition = "1" Then
'      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'         "CP10,CP43 from caseprogress,casepropertymap,CUSTOMER where " & _
'         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
'         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
'   Else
'      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'         "CP10,CP43 from caseprogress,casepropertymap,CUSTOMER where " & _
'         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and ( cp09<'C' ) AND CP57 IS NULL and cp01=cpm01(+) and " & _
'         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
'   End If
   If strSQLCondition = "1" Then
        'Modify By Cheng 2002/12/02
'      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'         "CP10,CP43,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
'         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
'         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
        'Modify By Cheng 2003/01/27
        '加顯示對造號數
'      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP43,CP64," & _
'         "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
'         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
'         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36," & _
         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP43,CP64," & _
         "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and ( cp09<'C' ) AND CP27 IS NOT NULL and cp01=cpm01(+) and " & _
         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' AND CP10<>'408' AND CP10<>'410' AND CP10<>'506') "
   Else
        'Modify By Cheng 2002/12/02
'      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'         "CP10,CP43,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
'         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and ( cp09<'C' ) AND CP57 IS NULL and cp01=cpm01(+) and " & _
'         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
        'Modify By Cheng 2003/01/27
        '加顯示對造號數
'      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
'         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP43,CP64," & _
'         "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
'         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and ( cp09<'C' ) AND CP57 IS NULL and cp01=cpm01(+) and " & _
'         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' ) "
      strExc(0) = "SELECT '',CP09," & strTmp & "," & _
         "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36," & _
         SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP43,CP64," & _
         "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
         ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and ( cp09<'C' ) AND CP57 IS NULL and cp01=cpm01(+) and " & _
         "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10<>'211' AND CP10<>'212' AND CP10<>'408' AND CP10<>'410' AND CP10<>'506') "
   End If
   
   '處理案件性質為"準備程序"(211), "言詞辯論"(212)
   'Modify By Cheng 2002/04/12
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') AND ( CP05 IS NULL AND CP27 IS NOT NULL ) And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
' 91.09.13 modify by louis (排序)
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10,CP43 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and ( cp09<'C' ) AND CP57 IS NULL And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
    'Modify By Cheng 2003/01/27
    '加顯示對造號數
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP43,CP64," & _
'      "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and ( cp09<'C' ) AND CP57 IS NULL And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) " & _
'      "ORDER BY SORTFIELD DESC "
   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP43,CP64," & _
      "CP10,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap,CUSTOMER where " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " and ( cp09<'C' ) AND CP57 IS NULL And cp01=cpm01(+) and " & _
      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' Or CP10='408' Or CP10='410' Or CP10='506') " & _
      "ORDER BY SORTFIELD DESC "

   'Modify By Cheng 2002/04/12
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
'   strExc(0) = strExc(0) + " union all select '',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64," & _
'      "CP10,CP43 from caseprogress,casepropertymap,CUSTOMER where " & _
'      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " and ( cp09<'C' ) And cp01=cpm01(+) and " & _
'      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) AND ( CP10='211' Or CP10='212' ) "
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   'ADD BY SONIA 2014/5/13 加相關總收號案件性質
   For ii = 1 To Me.MSHFlexGrid1.Rows - 1
       Me.MSHFlexGrid1.TextMatrix(ii, 2) = Me.MSHFlexGrid1.TextMatrix(ii, 2) & PUB_GetRelateCasePropertyName(Me.MSHFlexGrid1.TextMatrix(ii, 1), "1")
   Next ii
   'END 2014/5/13
   GridHead

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010504_2 = Nothing
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
        'Add By Cheng 2003/01/27
        '加對造號數
      .col = 4: .ColWidth(4) = 1500: .Text = "對造號數"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 600: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      'Add By Cheng 2002/12/02
      '加相關總收文號
      .col = 9: .ColWidth(9) = 1400: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .ColWidth(10) = 2000: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 0
        'Modify By Cheng 2002/12/02
'      'Add By Cheng 2002/06/27
'      .Col = 10: .ColWidth(10) = 1400: .Text = "相關總收文號"
'      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   '若案件性質非準備程序, 言詞辯論, 參加訴訟, 面詢, 閱卷
    'Modify By Cheng 2002/12/02
'    If Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 9) <> "211" And _
'        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 9) <> "212" And _
'        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 9) <> "506" Then
    'Modify By Cheng 2003/01/28
'    If Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 10) <> "211" And _
'        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 10) <> "212" And _
'        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 10) <> "506" Then
    If Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 11) <> "211" And _
        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 11) <> "212" And _
        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 11) <> "408" And _
        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 11) <> "410" And _
        Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 11) <> "506" Then
        'Modify By Cheng 2003/01/28
'        'Add By Cheng 2002/0627
'        If Len(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)) <= 0 Then
        If Len(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 7)) <= 0 Then
           GridClick MSHFlexGrid1, intLastRow, 0
           MsgBox "未發文資料, 不可輸入審查機關來函, 請重新選取!!!"
        End If
    End If
End Sub
