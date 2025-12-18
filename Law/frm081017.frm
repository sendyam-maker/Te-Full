VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081017 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   5820
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm081017.frx":0000
      Left            =   1224
      List            =   "frm081017.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   29
      Top             =   775
      Width           =   780
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1224
      MaxLength       =   50
      TabIndex        =   28
      Top             =   5085
      Width           =   4335
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm081017.frx":001D
      Left            =   1224
      List            =   "frm081017.frx":0027
      Style           =   2  '單純下拉式
      TabIndex        =   27
      Top             =   4740
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   26
      Top             =   5430
      Width           =   255
   End
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7224
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   8052
      TabIndex        =   2
      Top             =   70
      Width           =   1100
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   5025
      MaxLength       =   7
      TabIndex        =   0
      Top             =   4740
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2475
      Left            =   180
      TabIndex        =   30
      Top             =   2190
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4366
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   1224
      TabIndex        =   35
      Top             =   504
      Width           =   1368
      VariousPropertyBits=   27
      Size            =   "2413;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   1224
      TabIndex        =   34
      Top             =   1091
      Width           =   816
      VariousPropertyBits=   27
      Size            =   "1439;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   1224
      TabIndex        =   33
      Top             =   1362
      Width           =   816
      VariousPropertyBits=   27
      Size            =   "1439;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   8
      Left            =   1224
      TabIndex        =   32
      Top             =   1633
      Width           =   756
      VariousPropertyBits=   27
      Size            =   "1333;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   12
      Left            =   1224
      TabIndex        =   31
      Top             =   1905
      Width           =   795
      VariousPropertyBits=   27
      Size            =   "1402;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   14
      Left            =   5160
      TabIndex        =   25
      Top             =   1905
      Width           =   885
      VariousPropertyBits=   27
      Size            =   "1561;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   13
      Left            =   2070
      TabIndex        =   24
      Top             =   1905
      Width           =   1245
      VariousPropertyBits=   27
      Size            =   "2196;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   23
      Top             =   1633
      Width           =   1488
      VariousPropertyBits=   27
      Size            =   "2625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   10
      Left            =   5160
      TabIndex        =   22
      Top             =   1633
      Width           =   768
      VariousPropertyBits=   27
      Size            =   "1355;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   9
      Left            =   2070
      TabIndex        =   21
      Top             =   1633
      Width           =   1248
      VariousPropertyBits=   27
      Size            =   "2201;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   2070
      TabIndex        =   20
      Top             =   1362
      Width           =   6984
      VariousPropertyBits=   27
      Size            =   "12319;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   5808
      TabIndex        =   19
      Top             =   504
      Width           =   2040
      VariousPropertyBits=   27
      Size            =   "3598;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   18
      Top             =   504
      Width           =   576
      VariousPropertyBits=   27
      Size            =   "1016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   2070
      TabIndex        =   17
      Top             =   798
      Width           =   7032
      VariousPropertyBits=   27
      Size            =   "12404;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   2070
      TabIndex        =   16
      Top             =   1091
      Width           =   6984
      VariousPropertyBits=   27
      Size            =   "12319;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "是否列印定稿：　　 (N：不印)"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   5490
      Width           =   2445
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "日     期："
      Height          =   180
      Left            =   4224
      TabIndex        =   14
      Top             =   4800
      Width           =   765
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號："
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   5145
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "結      果："
      Height          =   180
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "案件性質  :"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   1670
      Width           =   864
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Left            =   4224
      TabIndex        =   10
      Top             =   1670
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "智權人員： "
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   1942
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "發  文  日："
      Height          =   180
      Left            =   4224
      TabIndex        =   8
      Top             =   1942
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "代  理  人："
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   1399
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "相關國家："
      Height          =   180
      Left            =   4224
      TabIndex        =   6
      Top             =   541
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "當  事  人："
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1128
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   541
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   835
      Width           =   900
   End
End
Attribute VB_Name = "frm081017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Label1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, intCols As Integer
Dim stName(1 To 3) As String, LcTmp As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String

Private Sub ComBack_Click()
   Unload Me
   frm081016.Show
End Sub

Private Sub Combo1_Click()
   Label1(3).Caption = stName(Combo1.ListIndex + 1)
End Sub

Private Sub ComSure_Click()
  Dim i As Integer
  Dim nSelect As Integer
  Dim strNP01 As String '總收文號
  
   With MSHFlexGrid1
        For i = 1 To .Rows - 1
            nSelect = 0
            If .TextMatrix(i, 0) = "v" Then
                'Add By Cheng 2002/01/02
                strNP01 = .TextMatrix(i, 1)
                nSelect = 1
                Exit For
            End If
        Next i
   End With
   
   If nSelect = 0 Then
      MsgBox "請點選收文號!", vbInformation, "代理人已收達/已提申"
      Exit Sub
   End If
   
   If Text2.Text = "" Then
      MsgBox "日期不可空白!", vbInformation, "代理人已收達/已提申"
      Text2.SetFocus
      Exit Sub
   Else
      If CheckIsTaiwanDate(Text2.Text) Then
         If Val(Text2.Text) + 19110000 > Val(GetTodayDate) Then
            MsgBox "日期不能超過系統日 !", vbCritical
            Text2.SetFocus
            Exit Sub
         End If
      Else
         Text2.SetFocus
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   strExc(3) = " WHERE CP09='" & LcTmp & "'"
   If Text2.Text <> "" Then
      If Combo2.ListIndex = 0 Then
         strExc(0) = "CP46"
      Else
         strExc(0) = "CP47"
      End If
      strExc(1) = "UPDATE CASEPROGRESS SET " & strExc(0) & "='" & ChangeTStringToWString(Text2.Text) & "'" & strExc(3)
      
      '911107 nick transation
      cnnConnection.Execute strExc(1)
      'If objLawDll.ExecSQL(1, strExc) = False Then
      '   DataErrorMessage 3
      '   Exit Sub
      'End If
   End If
   If Text4.Text <> "" Then
      strExc(1) = "UPDATE CASEPROGRESS SET CP45='" & Text4.Text & "'" & strExc(3)
      'strExc(2) = "UPDATE LAWCASE SET LC23='" & Text4.Text & "' WHERE " & strLc & "='" & Me.Tag & "'"
      
      '911107 nick transation
      cnnConnection.Execute strExc(1)
      
         strSql = "update caseprogress set cp45=" & CNULL(ChgSQL(Text4)) & " where cp09 in (select cp09 from caseprogress where cp45 is null and CP01 = '" & m_CP01 & "' AND  CP02 = '" & m_CP02 & "' AND CP03 = '" & m_CP03 & "' AND CP04 = '" & m_CP04 & "' and cp09<'C' AND cp44 in (select cp44 from caseprogress where cp09='" & LcTmp & "' ))"
         cnnConnection.Execute strSql
      'If objLawDll.ExecSQL(1, strExc) = False Then
      '   DataErrorMessage 3
      '   Exit Sub
      'End If
   End If
   
   'Add By Cheng 2002/01/02
   '若在結果欄選擇收達日
   If Me.Combo2.ListIndex = 0 Then
      strExc(1) = "UPDATE NextProgress SET NP06='Y' " & _
                  " Where NP01 = '" & strNP01 & "'" & _
                  " And NP02 = '" & m_CP01 & "'" & _
                  " And NP03 = '" & m_CP02 & "'" & _
                  " And NP04 = '" & m_CP03 & "'" & _
                  " And NP05 = '" & m_CP04 & "'" & _
                  " And Np07='997' "
      
      '911107 nick transation
      cnnConnection.Execute strExc(1)
      'If objLawDll.ExecSQL(1, strExc) = False Then
      '   DataErrorMessage 3
      '   Exit Sub
      'End If
      
   '若在結果欄選擇提申日
   ElseIf Me.Combo2.ListIndex = 1 Then
      strExc(1) = "UPDATE NextProgress SET NP06='Y' " & _
                  " Where NP01 = '" & strNP01 & "'" & _
                  " And NP02 = '" & m_CP01 & "'" & _
                  " And NP03 = '" & m_CP02 & "'" & _
                  " And NP04 = '" & m_CP03 & "'" & _
                  " And NP05 = '" & m_CP04 & "'" & _
                  " And (Np07='997' or NP07='998') "
      
      '911107 nick transation
      cnnConnection.Execute strExc(1)
      'If objLawDll.ExecSQL(1, strExc) = False Then
      '   DataErrorMessage 3
      '   Exit Sub
      'End If
      
   End If
   
 '911107 nick transation
  cnnConnection.CommitTrans
  
   Unload Me
   Unload frm081016
   frm081016.Show
   
 '911107 nick transation
     Exit Sub
CheckingErr:
    DataErrorMessage 3
     cnnConnection.RollbackTrans
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_CP01 = frm081016.Text8.Text
   m_CP02 = frm081016.Text7.Text
   If frm081016.Text6.Text <> "" Then
      m_CP03 = frm081016.Text6.Text
   Else
      m_CP03 = "0"
   End If
   If frm081016.Text5.Text <> "" Then
      m_CP04 = frm081016.Text5.Text
   Else
      m_CP04 = "00"
   End If
   Text1 = "N"
End Sub

Public Sub ShowData()
 Dim i As Integer, Lbl As Control
   For Each Lbl In Me.Label1
      Lbl.Caption = ""
   Next
   If Right(Me.Tag, 3) = "000" Then
      Label1(0).Caption = Left(Me.Tag, Len(Me.Tag) - 3)
   Else
      Label1(0).Caption = Me.Tag
   End If
   Label1(1).Caption = ChangeCustomerS(RsTemp.Fields(0).Value)
   'edit by nickc 2007/02/27 不用 dll
   'objPublicData.GetCustomer Label1(1).Caption, strExc(0)
   ClsPDGetCustomer Label1(1).Caption, strExc(0)
   Label1(2).Caption = strExc(0)
   For i = 1 To 3
      If IsNull(RsTemp.Fields(i).Value) = False Then stName(i) = RsTemp.Fields(i).Value
   Next
   Label1(4).Caption = RsTemp.Fields(4).Value
   'edit by nickc 2007/02/27 不用 dll
   'objPublicData.GetNation Label1(4).Caption, strExc(0)
   ClsPDGetNation Label1(4).Caption, strExc(0)
   Label1(5).Caption = strExc(0)
   RsTemp.Close
   Combo1.ListIndex = 0
   strExc(0) = "SELECT '',CP09,DECODE(LENGTH(CP05),NULL,NULL,SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)),CPM03,DECODE(LENGTH(CP27),NULL,NULL,SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2))," & _
               "DECODE(LENGTH(CP46),NULL,NULL,SUBSTR(CP46,1,4)-1911||'/'||SUBSTR(CP46,5,2)||'/'||SUBSTR(CP46,7,2))," & _
               "DECODE(LENGTH(CP47),NULL,NULL,SUBSTR(CP47,1,4)-1911||'/'||SUBSTR(CP47,5,2)||'/'||SUBSTR(CP47,7,2)) FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
      "CP27 IS NOT NULL AND CP47 IS NULL AND CP24 IS NULL AND " & _
      "CP01 ='" & m_CP01 & "' AND CP02 ='" & m_CP02 & "'" & _
      " AND CP03='" & m_CP03 & "' AND CP04 ='" & m_CP04 & "'" & _
      " AND CP09<'C' AND CP01 =CPM01(+) AND CP10=CPM02(+) ORDER BY CP27 DESC"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 1 Then Exit Sub
   Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
End Sub
Private Sub GetNum(strCP44 As String, strCP05 As String)
  Dim rsTmp As New ADODB.Recordset
  Dim strSql As String
  
  strSql = "SELECT CP45 FROM CASEPROGRESS WHERE CP09 =(SELECT MAX(CP09) FROM CASEPROGRESS WHERE " & _
           "CP01 ='" & m_CP01 & "' AND CP02 ='" & m_CP02 & "'" & _
           " AND CP03 ='" & m_CP03 & "' AND CP04 ='" & m_CP04 & "'" & _
           " AND CP09<'C' AND CP05 < '" & strCP05 & "'" & _
           " and cp44 ='" & strCP44 & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      If Not IsNull(rsTmp.Fields("CP45")) Then
         Text4.Text = rsTmp.Fields("CP45")
      End If
   End If
         
End Sub

Private Sub GridHead()
 Dim i As Integer
       
   With MSHFlexGrid1
      .row = 0
      .col = 0: .ColWidth(0) = 300: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 900: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1200: .Text = "代理人收達日"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1200: .Text = "代理人提申日"
      .CellAlignment = flexAlignCenterCenter
      
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081017 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
 Dim rs1 As New ADODB.Recordset, i As Integer
On Error Resume Next
   intCols = MSHFlexGrid1.Cols - 1
   ShowBar MSHFlexGrid1, intLastRow, intCols
   With MSHFlexGrid1
        ClearGrid
        .col = 0
       .row = intLastRow
        If .Text = "v" Then
           .Text = ""
        ElseIf .Text = "" Then
           .Text = "v"
        End If
           
   End With
   MSHFlexGrid1.col = 1
   LcTmp = MSHFlexGrid1.Text
   strExc(0) = "SELECT CP44,CP10,CP14,CP13,CP45,CP27,CP05 FROM CASEPROGRESS WHERE " & _
      "CP09='" & LcTmp & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   For i = 6 To 14
      Label1(i).Caption = ""
   Next
   Text4.Text = ""
   If intI = 1 Then
      strExc(0) = "SELECT FA05 FROM FAGENT WHERE " & ChgFagent(RsTemp.Fields(0).Value)
      'edit by nickc 2007/02/27 不用 dll
      'Set rs1 = objLawDll.ReadRstMsg(1, strExc(0))
      Set rs1 = ClsLawReadRstMsg(1, strExc(0))
      If Right(RsTemp.Fields(0).Value, 3) <> "000" Then
         Label1(6).Caption = RsTemp.Fields(0).Value
      Else
         Label1(6).Caption = Left(RsTemp.Fields(0).Value, Len(RsTemp.Fields(0).Value) - 3)
      End If
      Label1(7).Caption = rs1.Fields(0).Value
      
      Label1(8).Caption = RsTemp.Fields(1).Value
      'edit by nickc 2007/02/27 不用 dll
      'If objPublicData.GetCaseProperty("CFL", RsTemp.Fields(1).Value, strExc(0)) = True Then
      If ClsPDGetCaseProperty("CFL", RsTemp.Fields(1).Value, strExc(0)) = True Then
         Label1(9).Caption = strExc(0)
      End If
      
      Label1(10).Caption = RsTemp.Fields(2).Value
      If IsNull(RsTemp.Fields(2).Value) = False Then
      'edit by nickc 2007/02/27 不用 dll
      'If objPublicData.GetStaff(RsTemp.Fields(2).Value, strExc(0)) = True Then
      If ClsPDGetStaff(RsTemp.Fields(2).Value, strExc(0)) = True Then
         Label1(11).Caption = strExc(0)
      End If
      End If
      
      Label1(12).Caption = RsTemp.Fields(3).Value
      'edit by nickc 2007/02/27 不用 dll
      'If objPublicData.GetStaff(RsTemp.Fields(3).Value, strExc(0)) = True Then
      If ClsPDGetStaff(RsTemp.Fields(3).Value, strExc(0)) = True Then
         Label1(13).Caption = strExc(0)
      End If
      
      Text4.Text = RsTemp.Fields(4).Value
      Label1(14).Caption = ChangeWStringToTString(RsTemp.Fields(5).Value)
   End If
   If Text4.Text = "" Then
      GetNum GetNewFagent(RsTemp.Fields("CP44")), RsTemp.Fields("CP05")
   End If
   MSHFlexGrid1.col = 5
   If MSHFlexGrid1.Text = "" Then
      Combo2.ListIndex = 0
   Else
      Combo2.ListIndex = 1
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1.Text <> "" Then
      If Text1.Text <> "N" Then
         MsgBox "只可輸入 N 或空白 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2.Text <> "" Then
      Cancel = True
      If CheckIsTaiwanDate(Text2.Text) Then
         If Val(Text2.Text) + 19110000 > Val(GetTodayDate) Then
            MsgBox "日期不能超過系統日 !", vbCritical
            TextInverse Text2
         Else
            Cancel = False
         End If
      Else
         TextInverse Text2
      End If
   End If
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub
Private Sub ClearGrid()
 Dim i As Integer
   With MSHFlexGrid1
      .Visible = False
      For i = 1 To .Rows - 1
         .col = 0
         .row = i
         .Text = ""
      Next
      .Visible = True
   End With
End Sub

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text1.Enabled = True Then
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text2.Enabled = True Then
   Cancel = False
   Text2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

