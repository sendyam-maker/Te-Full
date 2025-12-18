VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060101_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件進度查詢"
   ClientHeight    =   5736
   ClientLeft      =   120
   ClientTop       =   996
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7530
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8376
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060101_2.frx":0000
      Left            =   960
      List            =   "frm060101_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1260
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3825
      Left            =   60
      TabIndex        =   0
      Top             =   1860
      Width           =   9195
      _ExtentX        =   16214
      _ExtentY        =   6752
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1620
      TabIndex        =   13
      Top             =   1320
      Width           =   7635
      VariousPropertyBits=   268435483
      Caption         =   "LblFM2"
      Size            =   "13467;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   960
      TabIndex        =   12
      Top             =   1620
      Width           =   8295
      VariousPropertyBits=   268435483
      Caption         =   "LblFM2"
      Size            =   "14631;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   5970
      TabIndex        =   11
      Top             =   720
      Width           =   3225
      VariousPropertyBits=   268435483
      Caption         =   "LblFM2"
      Size            =   "5689;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Top             =   990
      Width           =   3795
      VariousPropertyBits=   268435483
      Caption         =   "LblFM2"
      Size            =   "6694;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   9
      Top             =   720
      Width           =   3795
      VariousPropertyBits=   268435483
      Caption         =   "LblFM2"
      Size            =   "6694;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利號數"
      Height          =   180
      Index           =   2
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frm060101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/12 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strPatent As String
Dim pa(10) As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
Public fmParent As Form


Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer
   Select Case Index
      Case 0
         MSHFlexGrid1.col = 0
         For i = 1 To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.row = i
            If MSHFlexGrid1.Text = "v" Then
               MSHFlexGrid1.col = 1
               'Modified by Lydia 2015/12/31 會稿發文一定要有相關總收文號
               'If MSHFlexGrid1.Text <> "" Then fmParent.text1(6) = MSHFlexGrid1.Text
               If MSHFlexGrid1.Text <> "" Then
                  If TypeName(fmParent) = "frm060104_3" Or TypeName(fmParent) = "frm060104_g" Then
                     fmParent.txtCP43 = MSHFlexGrid1.Text
                  'Added by Lydia 2018/05/14 新案建檔：翻譯分案無紙化
                  ElseIf TypeName(fmParent) = "frm060102" Then
                     fmParent.txtTF(30).Text = MSHFlexGrid1.Text
                     fmParent.Chk02.Value = 0
                  'end 2018/05/14
                  'Added by Morgan 2024/11/15
                  ElseIf fmParent.Name = "frm06010604_3" Then
                     fmParent.m_UpdCP09 = MSHFlexGrid1.Text
                  'end 2024/11/15
                  Else
                     fmParent.Text1(6) = MSHFlexGrid1.Text
                  End If
               End If
               'end 2015/12/31
               Exit For
            End If
         Next
         
         If fmParent.Visible = False Then 'Added by Morgan 2024/11/15 可能會以強制表單方式呼叫
            fmParent.Show
         End If
      Case 1
         If fmParent.Visible = False Then 'Added by Morgan 2024/11/15 可能會以強制表單方式呼叫
            fmParent.Show
         End If
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(3) = pa(1)
      Case "英"
         Label2(3) = pa(2)
      Case "日"
         Label2(3) = pa(3)
   End Select
End Sub

Private Sub Form_Activate()
   Me.ZOrder 'Added by Morgan 2024/11/15 改獨立視窗後會被MdiMain蓋住,故增加此控制
End Sub

Private Sub Form_Load()
   'Modified by Morgan 2024/11/15 因改為獨立視窗,加控制顯示於mdiMain中央
   MoveFormToCenter Me, True
   intWhere = 國外_FC
   With fmParent
      'Added by Lydia 2015/12/31 會稿發文一定要有相關總收文號
      If .Name = "frm060104_3" Or .Name = "frm060104_g" Then
         strReceiveNo = .Label3(0)
         strPatent = .Text1 & .Text2 & .Text3 & .Text4
      'Added by Lydia 2018/05/14 新案建檔：翻譯分案無紙化
      ElseIf .Name = "frm060102" Then
         strReceiveNo = "AAA" '不傳入收文號
         strPatent = .Text1 & .Text2 & .Text3 & .Text4
      'end 2018/05/14
      'Added by Morgan 2024/11/15
      ElseIf .Name = "frm06010604_3" Then
         strReceiveNo = .Label3(5)
         strPatent = .Text2 & .Text3 & .Text4 & .Text5
      'end 2024/11/15
      Else
         strReceiveNo = .Label3(8)
         strPatent = .Label3(9)
      End If
   End With
   ReadPatent
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_2 = Nothing
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer
 Dim m_PA09 As String 'Added by Lydia 2016/1/13
   For Each Lbl In Label2
      Lbl = ""
   Next
   Label2(0) = strPatent
   'Modified by Lydia 2016/1/13 +PA09
   strExc(0) = "SELECT PA11,PA22,PA05,PA06,PA07,CU04,PA09 FROM PATENT,CUSTOMER WHERE " & ChgPatent(strPatent) & _
      " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      If Not IsNull(.Fields(0)) Then Label2(2) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
      For i = 1 To 3
         If Not IsNull(.Fields(i + 1)) Then
            pa(i) = .Fields(i + 1)
         Else
            pa(i) = ""
         End If
      Next
      If Not IsNull(.Fields(5)) Then Label2(4) = .Fields(5)
      m_PA09 = .Fields("PA09")
   End If
   End With
   'Modify by Morgan 2004/10/13 第一句語法需排除CP10=701以免重複
   'Modified by Lydia 2016/1/13 案件性質依國別顯示 DECODE(pa09,'000',CPM03,CPM04)) CPM03
   'strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & ",CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,CP40) FROM CASEPROGRESS,CASEPROPERTYMAP" & _
      " WHERE " & ChgCaseprogress(strPatent) & " AND CP09<>'" & strReceiveNo & "'" & _
      " AND CP10<>'701' AND cp01=cpm01 and cp10=cpm02 UNION " & _
      "SELECT '',CP09," & SQLDate("CP05") & ",CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      "NVL(CU04,NVL(CU05,CU06)) FROM CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      ChgCaseprogress(strPatent) & " AND CP09<>'" & strReceiveNo & "' AND " & _
      "CP10='701' AND CP01=CPM01 and CP10=CPM02 AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
   'Added by Lydia 2018/05/14 新案建檔：翻譯分案無紙化(限定B類補文件)
   If fmParent.Name = "frm060102" Then
       strExc(1) = " and substr(cp09,1,1)='B' and cp10='202' "
       
   'Added by Morgan 2024/11/15 一般來函只需列出未發文的AB類程序
   ElseIf fmParent.Name = "frm06010604_3" Then
       strExc(1) = " and cp158=0 and cp159=0 and cp09<'C'"
   'end 2024/11/15
   Else
       strExc(1) = ""
   End If
   'end 2018/05/14
   
   
   
   'Modified by Lydia 2018/05/14 +其他條件strexc(1)
   strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & IIf(m_PA09 = "000", "", " (CPM04)") & " CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,CP40) FROM CASEPROGRESS,CASEPROPERTYMAP" & _
      " WHERE " & ChgCaseprogress(strPatent) & " AND CP09<>'" & strReceiveNo & "'" & strExc(1) & _
      " AND CP10<>'701' AND cp01=cpm01 and cp10=cpm02 UNION " & _
      "SELECT '',CP09," & SQLDate("CP05") & "," & IIf(m_PA09 = "000", "", " (CPM04)") & " CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      "NVL(CU04,NVL(CU05,CU06)) FROM CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      ChgCaseprogress(strPatent) & " AND CP09<>'" & strReceiveNo & "' AND " & _
      "CP10='701' AND CP01=CPM01 and CP10=CPM02 AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "後金"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1400: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub
