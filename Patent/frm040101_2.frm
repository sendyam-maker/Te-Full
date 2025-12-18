VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040101_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件進度查詢"
   ClientHeight    =   5748
   ClientLeft      =   -756
   ClientTop       =   1452
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9348
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7500
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8328
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm040101_2.frx":0000
      Left            =   960
      List            =   "frm040101_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1125
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3912
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   6900
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
      Index           =   4
      Left            =   840
      TabIndex        =   13
      Top             =   1440
      Width           =   8220
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "14499;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1680
      TabIndex        =   12
      Top             =   1140
      Width           =   7350
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "12965;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   960
      TabIndex        =   11
      Top             =   900
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3598;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Top             =   660
      Width           =   3360
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5927;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   660
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3651;317"
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
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利號數"
      Height          =   180
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   720
   End
End
Attribute VB_Name = "frm040101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/3 改成Form2.0(Label..)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Public iGo As Integer
Dim strReceiveNo As String, strPatent As String
Dim pa(10) As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer
   Select Case Index
      Case 0
         MSHFlexGrid1.col = 0
         For i = 1 To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.row = i
            If MSHFlexGrid1.Text = "v" Then
               MSHFlexGrid1.col = 1
               If MSHFlexGrid1.Text <> "" Then
                  Select Case iGo
                     Case 4
                        frm040101_1.Text1(6) = MSHFlexGrid1.Text
                     Case 5
                        frm050101_2.txtCaseField(5) = MSHFlexGrid1.Text
                  End Select
               End If
               Exit For
            End If
         Next
   End Select
   
   Select Case iGo
      Case 4
         frm040101_1.Show
         frm040101_1.Set414Date 'Add by Morgan 2009/11/9
      Case 5
         frm050101_2.Show
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

Private Sub Form_Load()
   MoveFormToCenter Me
   'Add By Cheng 2002/07/17
   strReceiveNo = ""
   strPatent = ""
   Select Case iGo
      Case 4
         intWhere = 國內
         With frm040101_1
            strReceiveNo = .Label3(8)
            strPatent = .Label3(9)
         End With
      Case 5
         intWhere = 國外_CF
         With frm050101_2
            strReceiveNo = .lblReceiveCode
            'Modified by Morgan 2023/12/13
            'strPatent = Replace(.lblCaseCode, " - ", "")
            strPatent = Replace(.lblCaseCode, "-", "")
            'end 2023/12/13
         End With
   End Select
   ReadPatent
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040101_2 = Nothing
End Sub

Private Sub ReadPatent()
 Dim Lbl, i As Integer
   For Each Lbl In Label2
      Lbl.Caption = ""
   Next
   strExc(1) = "CPM03,"
   Select Case iGo
      Case 4
         Label2(0) = frm040101_1.Label3(9)
      Case 5
         Label2(0) = frm050101_2.lblCaseCode
   End Select
   strExc(0) = "SELECT PA11,PA22,PA05,PA06,PA07,CU04,PA09 FROM PATENT,CUSTOMER WHERE " & _
      ChgPatent(strPatent) & " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)"
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
      If Not IsNull(.Fields(6)) Then
         If .Fields(6) = 台灣國家代號 Then
            strExc(1) = "CPM03,"
         Else
            strExc(1) = "CPM04,"
         End If
      End If
   End If
   End With
   
   'Modify by Morgan 2004/10/11 第一句語法需排除CP10=701以免重複
   strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & strExc(1) & SQLDate("CP27") & _
      ",CP24,CP19,DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,CP40) FROM CASEPROGRESS," & _
      "CASEPROPERTYMAP WHERE " & ChgCaseprogress(strPatent) & " AND CP09<>'" & strReceiveNo & "'" & _
      " AND CP10<>'701' AND cp01=cpm01(+) and cp10=cpm02(+) UNION " & _
      "SELECT '',CP09," & SQLDate("CP05") & "," & strExc(1) & SQLDate("CP27") & _
      ",CP24,CP19,NVL(CU04,NVL(CU05,CU06)) FROM CASEPROGRESS,CASEPROPERTYMAP," & _
      "CUSTOMER WHERE " & ChgCaseprogress(strPatent) & " AND CP09<>'" & strReceiveNo & "'" & _
      " AND CP10='701' AND cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
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
