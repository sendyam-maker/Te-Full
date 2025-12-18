VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1103_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "相關卷號"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   996
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9312
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3252
      Left            =   96
      TabIndex        =   37
      Top             =   2304
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   5736
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      CausesValidation=   0   'False
      Height          =   320
      Index           =   2
      Left            =   6744
      TabIndex        =   10
      Top             =   864
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   320
      Index           =   1
      Left            =   5916
      TabIndex        =   9
      Top             =   864
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   320
      Index           =   0
      Left            =   5088
      TabIndex        =   8
      Top             =   864
      Width           =   800
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2376
      TabIndex        =   26
      Top             =   876
      Width           =   2532
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   0
         Left            =   0
         MaxLength       =   6
         TabIndex        =   1
         Top             =   0
         Width           =   1212
      End
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.TextBox txtSystem 
      Height          =   288
      Left            =   1536
      MaxLength       =   3
      TabIndex        =   0
      Top             =   876
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8325
      TabIndex        =   13
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   6240
      TabIndex        =   11
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7065
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraTF1 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1455
      TabIndex        =   19
      Top             =   516
      Visible         =   0   'False
      Width           =   2412
      Begin VB.Label lblTFCode 
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   36
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblTFCode 
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   22
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblTFCode 
         Height          =   255
         Index           =   1
         Left            =   930
         TabIndex        =   21
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblTFCode 
         Height          =   252
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.Frame fraElse1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1455
      TabIndex        =   15
      Top             =   516
      Width           =   2295
      Begin VB.Label Label8 
         Caption         =   " -"
         Height          =   255
         Left            =   1005
         TabIndex        =   35
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   " -"
         Height          =   255
         Left            =   615
         TabIndex        =   34
         Top             =   0
         Width           =   135
      End
      Begin VB.Label lblCode 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblCode 
         Height          =   255
         Index           =   1
         Left            =   750
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblCode 
         Height          =   252
         Index           =   2
         Left            =   1260
         TabIndex        =   16
         Top             =   0
         Width           =   432
      End
   End
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2376
      TabIndex        =   27
      Top             =   876
      Width           =   2652
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   0
         Left            =   0
         MaxLength       =   5
         TabIndex        =   4
         Top             =   0
         Width           =   972
      End
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   1
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   5
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   2
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   14
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   3
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   6
         Top             =   0
         Width           =   492
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1050
      TabIndex        =   7
      Top             =   1230
      Width           =   8070
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14235;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblNationName 
      Height          =   255
      Left            =   1590
      TabIndex        =   33
      Top             =   1590
      Width           =   3075
      VariousPropertyBits=   27
      Size            =   "5424;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSystem 
      Height          =   252
      Left            =   1056
      TabIndex        =   23
      Top             =   516
      Width           =   372
   End
   Begin VB.Label lblNation 
      Height          =   252
      Left            =   1056
      TabIndex        =   32
      Top             =   1596
      Width           =   492
   End
   Begin MSForms.Label lblCustomer 
      Height          =   255
      Left            =   930
      TabIndex        =   31
      Top             =   1950
      Width           =   8160
      VariousPropertyBits=   27
      Size            =   "14393;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "申請人："
      Height          =   255
      Left            =   90
      TabIndex        =   30
      Top             =   1950
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "申請國家："
      Height          =   252
      Left            =   96
      TabIndex        =   29
      Top             =   1596
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   28
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "相關之本所案號："
      Height          =   252
      Left            =   96
      TabIndex        =   25
      Top             =   876
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   96
      TabIndex        =   24
      Top             =   516
      Width           =   972
   End
End
Attribute VB_Name = "frm1103_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/12 改成Form2.0 (MSFlexGrid1改為MSHFlexGrid1)
'Memo by Morgan 2021/10/15 改成Form2.0 (cboCaseName,lblNationName...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

'intWhereComeFrom  1:frm1103_1     2:Others
Public intWhereComeFrom As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, intLeaveKind As Integer
'edit by nickc 2007/02/05 不用 dll 了
'Public obj011 As New cls011
Public m_form As Form
Dim bFrm1103 As Boolean
Dim m_bolCombine As Boolean, m_strCombine(3) As String 'Added by Morgan 2022/11/8
Dim m_strCaseNo(3) As String 'Added by Morgan 2023/12/27 本案號

Public Sub SetFrom1103()
   bFrm1103 = True
End Sub

Private Sub ReadRelationData()
Dim varSaveCursor, strRelation() As String, i As Integer
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   If lblSystem = 馬德里案 Then
      strExc(0) = lblSystem.Caption
      If lblTFCode(1) = "" Then
         strExc(1) = lblTFCode(0) & "0"
      Else
         strExc(1) = lblTFCode(0) & lblTFCode(1)
      End If
      If lblTFCode(2) = "" Then
         strExc(2) = "0"
      Else
         strExc(2) = lblTFCode(2)
      End If
      If lblTFCode(3) = "" Then
         strExc(3) = "00"
      Else
         strExc(3) = lblTFCode(3)
      End If
   Else
      strExc(0) = lblSystem.Caption
      strExc(1) = lblCode(0).Caption
      If lblCode(1) = "" Then
         strExc(2) = "0"
      Else
         strExc(2) = lblCode(1).Caption
      End If
      If lblCode(2) = "" Then
         strExc(3) = "00"
      Else
         strExc(3) = lblCode(2)
      End If
   End If
   
   'Added by Morgan 2023/12/27
   m_strCaseNo(0) = strExc(0)
   m_strCaseNo(1) = strExc(1)
   m_strCaseNo(2) = strExc(2)
   m_strCaseNo(3) = strExc(3)
   'end 2023/12/27
   
   'Modify by Morgan 2006/6/22
   'i = obj011.ReadCaseRelationData(strExc(0), strExc(1), strExc(2), strExc(3), strRelation())
   i = PUB_ReadCaseRelationData(strExc(0), strExc(1), strExc(2), strExc(3), strRelation())
   If i = 1 Then
      If Not SetRelationToLisBox(strRelation()) Then GoTo err1
   ElseIf i = -1 Then
      GoTo err1
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
err1:
   Screen.MousePointer = vbDefault
   intLeaveKind = 1
   bolLeave = True
   Unload Me
   Exit Sub
ErrHand:
   ErrorMsg
   Screen.MousePointer = vbDefault
End Sub

'Modify by Morgan 2006/6/30 加p_bolAdd控制是否為新增案號
Private Function SetRelationToLisBox(ByRef strRelation() As String, Optional ByVal p_bolAdd As Boolean = False) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNation As String, i As Integer, strCaseName As String
Dim strCaseCode As String, strTemp As String
If p_bolAdd = False Then
   'Modified by Lydia 2023/10/12
   'Me.MSFlexGrid1.Rows = 1 'Add by Morgan 2005/1/19 避免重複增加
   Me.MSHFlexGrid1.Rows = 2
End If
For i = 0 To UBound(strRelation, 2)
       'edit by nickc 2007/02/02 不用 dll 了
       'If objPublicData.CheckCaseCodeIsExist(strRelation(0, i), strRelation(1, i), strRelation(2, i), strRelation(3, i), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
       If ClsPDCheckCaseCodeIsExist(strRelation(0, i), strRelation(1, i), strRelation(2, i), strRelation(3, i), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
          If strRelation(0, i) = "LA" Then
             strNation = "000"
          End If
          If strCaseName1 <> "" Then
             strCaseName = strCaseName1
          ElseIf strCaseName2 <> "" Then
             strCaseName = strCaseName2
          Else
             strCaseName = strCaseName3
          End If
          If strRelation(0, i) = 馬德里案 Then
             strCaseCode = strRelation(0, i) + "-" + Left(strRelation(1, i), 5) + IIf(Right(strRelation(1, i), 1) = "0", "-0", "-" + Right(strRelation(1, i), 1)) + IIf(strRelation(2, i) = "0", "-0", "-" + strRelation(2, i)) + IIf(strRelation(3, i) = "00", "-00", "-" + strRelation(3, i))
          Else
             strCaseCode = strRelation(0, i) + "-" + strRelation(1, i) + IIf(strRelation(2, i) = "0", "-0", "-" + strRelation(2, i)) + IIf(strRelation(3, i) = "00", "-00", "-" + strRelation(3, i))
          End If
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.GetNation(strNation, strTemp) Then
          If ClsPDGetNation(strNation, strTemp) Then
            'Modified by Lydia 2023/10/12 MSFlexGrid1=>MSHFlexGrid1
            If i > 0 Then 'Added by Lydia 2023/10/12
               Me.MSHFlexGrid1.AddItem Me.MSHFlexGrid1.Rows
            End If 'Added by Lydia 2023/10/12
            Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 0) = ""
            Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 1) = strCaseCode
            Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 2) = strCaseName
            Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 3) = strTemp
            Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 4) = strCustomer
            'end 2023/10/12
             SetRelationToLisBox = True
          Else
             SetRelationToLisBox = False
             Exit For
          End If
       Else
          SetRelationToLisBox = False
          Exit For
       End If
Next

'Added by Lydia 2023/10/12
If SetRelationToLisBox = True Then
   Me.MSHFlexGrid1.FixedRows = 1
End If

End Function
'Added by Morgan 2012/12/14
'檢查案號是否有存在於其他群組
Private Function ChkNoOtherGroup(pCaseNo1 As String, pCaseNo2 As String) As Boolean
   Dim stSQL As String, intR As Integer, adoTmp As ADODB.Recordset
   
   '原來就有關聯
   stSQL = "select * from caserelation1 where cr01||cr02||cr03||cr04='" & pCaseNo1 & "' and cr05||cr06||cr07||cr08='" & pCaseNo2 & "'"
   intR = 1
   Set adoTmp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      ChkNoOtherGroup = True
   Else
      '新增的案號已有關聯
      stSQL = "select * from caserelation1 where cr01||cr02||cr03||cr04='" & pCaseNo2 & "'"
      intR = 1
      Set adoTmp = ClsLawReadRstMsg(intR, stSQL)
      If intR = 0 Then
         ChkNoOtherGroup = True
      End If
   End If
   Set adoTmp = Nothing
End Function


Private Sub cmdMove_Click(Index As Integer)
Dim intlastIndex As Integer, strCaseName As String, strCaseCode As String, varSaveCursor
Dim bolRt As Boolean, i As Integer
Dim blnChkValue As Boolean
Dim ii As Integer
Dim strRelation() As String 'Added by Morgan 2022/11/8

Select Case Index
             Case 0 '新增

'Removed by Morgan 2023/12/18 因為群組方式建關聯，本案也要列出，個案移除群組時操作才會一致
'                        If txtSystem = lblSystem Then
'                           If txtSystem = 馬德里案 Then
'                              For i = 0 To 3
'                                     If txtTFCode(i) <> lblTFCode(i) Then
'                                        Exit For
'                                     End If
'                              Next
'                              If i = 4 Then
'                                 ShowMsg MsgText(8004)
'                                 Exit Sub
'                              End If
'                           Else
'                              For i = 0 To 2
'                                     If txtCode(i) <> lblCode(i) Then
'                                        Exit For
'                                     End If
'                              Next
'                              If i = 3 Then
'                                 ShowMsg MsgText(8004)
'                                 Exit Sub
'                              End If
'                           End If
'                        End If
'end 2023/12/18

                        If txtSystem = 馬德里案 Then
                           bolRt = CheckKeyIn1(3)
                           If bolRt Then
                              strCaseCode = txtSystem + "-" + txtTFCode(0) + IIf(txtTFCode(1) = "", "-0", "-" + txtTFCode(1)) + IIf(txtTFCode(2) = "", "-0", "-" + txtTFCode(2)) + IIf(txtTFCode(3) = "", "-00", "-" + txtTFCode(3))
                           End If
                        Else
                           bolRt = CheckKeyIn2(2)
                           If bolRt Then
                              strCaseCode = txtSystem + "-" + txtCode(0) + IIf(txtCode(1) = "", "-0", "-" + txtCode(1)) + IIf(txtCode(2) = "", "-00", "-" + txtCode(2))
                           End If
                        End If
                        
                        'Added by Morgan 2012/12/14
                        'Modified by Morgan 2023/12/27
                        'If txtSystem = 馬德里案 Then
                        '   strExc(5) = lblSystem & lblTFCode(0) & Right("0" & lblTFCode(1), 1) & Right("0" & lblTFCode(2), 1) & Right("00" & lblTFCode(3), 2)
                        'Else
                        '   strExc(5) = lblSystem & lblCode(0) & Right("0" & lblCode(1), 1) & Right("00" & lblCode(2), 2)
                        'End If
                        strExc(5) = m_strCaseNo(0) & m_strCaseNo(1) & m_strCaseNo(2) & m_strCaseNo(3)
                        'end 2023/12/27
                        
                        If txtSystem = 馬德里案 Then
                           strExc(6) = txtSystem & txtTFCode(0) & Right("0" & txtTFCode(1), 1) & Right("0" & txtTFCode(2), 1) & Right("00" & txtTFCode(3), 2)
                           strExc(1) = txtTFCode(0) & Right("0" & txtTFCode(1), 1)
                           strExc(2) = Right("0" & txtTFCode(2), 1)
                           strExc(3) = Right("00" & txtTFCode(3), 2)
                        Else
                           strExc(6) = txtSystem & txtCode(0) & Right("0" & txtCode(1), 1) & Right("00" & txtCode(2), 2)
                           strExc(1) = txtCode(0)
                           strExc(2) = Right("0" & txtCode(1), 1)
                           strExc(3) = Right("00" & txtCode(2), 2)
                        End If
                        
                     If strExc(5) <> strExc(6) Then 'Added by Morgan 2023/18 本案改也會列出，重新加回時不必檢查
                     
                        If ChkNoOtherGroup(strExc(5), strExc(6)) = False Then
                           'Modified by Morgan 2022/11/8
                           'MsgBox "案號 " & strCaseCode & " 已另有相關群組，若要加入該群組，" & vbCrLf & vbCrLf & "請回前畫面並以此案號來建立關聯！", vbExclamation
                           'Exit Sub
                           If m_bolCombine = True Then
                              MsgBox "此案號已存在於另一相關群組，因本群組已有合併群組，故不可再行合併！"
                              
                           ElseIf MsgBox("案號 " & strCaseCode & " 已另有相關群組，是否要加入該群組？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                              intI = PUB_ReadCaseRelationData(txtSystem, strExc(1), strExc(2), strExc(3), strRelation())
                              If intI = 1 Then
                                 
                                 If SetRelationToLisBox(strRelation()) Then
                                    m_strCombine(0) = txtSystem
                                    m_strCombine(1) = strExc(1)
                                    m_strCombine(2) = strExc(2)
                                    m_strCombine(3) = strExc(3)
                                    m_bolCombine = True
                                    
                                    'Added by Morgan 2023/12/27
                                    bolRt = False
                                    txtSystem = ""
                                    For i = 0 To 2
                                           txtTFCode(i) = ""
                                           txtCode(i) = ""
                                    Next
                                    txtTFCode(i) = ""
                                    txtSystem.SetFocus
                                    'end 2023/12/27
                                 End If
                              End If
                           Else
                              Exit Sub
                           End If
                           'end 2022/11/8
                        End If
                        'end 2012/12/14
                        
                     End If 'Added by Morgan 2023/18
                     
                        varSaveCursor = Screen.MousePointer
                        Screen.MousePointer = vbHourglass
                        
                        If bolRt Then
                           'Modified by Lydia 2023/10/12 MSFlexGrid1=>MSHFlexGrid1
                           For i = 1 To Me.MSHFlexGrid1.Rows - 1
                               If Me.MSHFlexGrid1.TextMatrix(i, 1) = strCaseCode Then
                                   Exit For
                               End If
                           Next i
                           If i = Me.MSHFlexGrid1.Rows Then
                              strCaseName = Mid(cboCaseName.List(cboCaseName.ListIndex), 3)
                              Me.MSHFlexGrid1.AddItem Me.MSHFlexGrid1.Rows
                              Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 0) = ""
                              Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 1) = strCaseCode
                              Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 2) = strCaseName
                              Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 3) = lblNationName
                              Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 4) = lblCustomer
                            'end 2023/10/12
                              txtSystem = ""
                              For i = 0 To 2
                                     txtTFCode(i) = ""
                                     txtCode(i) = ""
                              Next
                              txtTFCode(i) = ""
                              txtSystem.SetFocus
                           Else
                              ShowMsg MsgText(8005)
                              txtSystem.SetFocus
                           End If
                        End If
                        Screen.MousePointer = varSaveCursor
                        
             Case 1 '刪除
                        blnChkValue = False
RemoveItem:
                        'Modified by Lydia 2023/10/12 MSFlexGrid1=>MSHFlexGrid1
                        For ii = 1 To Me.MSHFlexGrid1.Rows - 1
                            If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
                                blnChkValue = True
                                If Me.MSHFlexGrid1.Rows > 2 Then
                                    Me.MSHFlexGrid1.RemoveItem ii
                        'end 2023/10/12
                                Else
                                    InitialGrid
                                End If
                                GoTo RemoveItem
                            End If
                        Next ii
                        If blnChkValue = False Then ShowMsg MsgText(8006)
             Case 2 '清除
                        InitialGrid
End Select
End Sub

Private Function SaveData() As Boolean

   Dim i As Integer, strTmp As Variant, j As Integer, strRelation() As String
   Dim varSaveCursor
   Dim StrSQLa As String

   Screen.MousePointer = vbHourglass
   
On Error GoTo ErrHand

   cnnConnection.BeginTrans
   'Modified by Lydia 2023/10/12 MSFlexGrid1=>MSHFlexGrid1
'Modified by Morgan 2023/12/27 清除所有案號改為刪除群組(與逐案刪除為自群組剔除的概念一致)
'   If Me.MSHFlexGrid1.Rows <= 1 Then
'       If Me.lblSystem.Caption = 馬德里案 Then
'           StrSQLa = "Delete From CaseRelation1 Where CR01='" & frm1103_1.txtSystem.Text & "' And CR02='" & frm1103_1.txtTFCode(0).Text & IIf(frm1103_1.txtTFCode(1).Text = "", 0, frm1103_1.txtTFCode(1).Text) & "' And CR03='" & Left(frm1103_1.txtTFCode(2).Text & "0", 1) & "' And CR04='" & Left(frm1103_1.txtTFCode(3).Text & "00", 2) & "' "
'           cnnConnection.Execute StrSQLa
'           StrSQLa = "Delete From CaseRelation1 Where CR05='" & frm1103_1.txtSystem.Text & "' And CR06='" & frm1103_1.txtTFCode(0).Text & IIf(frm1103_1.txtTFCode(1).Text = "", 0, frm1103_1.txtTFCode(1).Text) & "' And CR07='" & Left(frm1103_1.txtTFCode(2).Text & "0", 1) & "' And CR08='" & Left(frm1103_1.txtTFCode(3).Text & "00", 2) & "' "
'           cnnConnection.Execute StrSQLa
'       Else
'           StrSQLa = "Delete From CaseRelation1 Where CR01='" & frm1103_1.txtSystem.Text & "' And CR02='" & frm1103_1.txtCode(0).Text & "' And CR03='" & Left(frm1103_1.txtCode(1).Text & "0", 1) & "' And CR04='" & Left(frm1103_1.txtCode(2).Text & "00", 2) & "' "
'           cnnConnection.Execute StrSQLa
'           StrSQLa = "Delete From CaseRelation1 Where CR05='" & frm1103_1.txtSystem.Text & "' And CR06='" & frm1103_1.txtCode(0).Text & "' And CR07='" & Left(frm1103_1.txtCode(1).Text & "0", 1) & "' And CR08='" & Left(frm1103_1.txtCode(2).Text & "00", 2) & "' "
'           cnnConnection.Execute StrSQLa
'       End If
   If Me.MSHFlexGrid1.Rows <= 2 Then
      If MsgBox("群組內" & IIf(Me.MSHFlexGrid1.Rows = 1, "已無", "僅存 1 ") & "案件，是否確定要刪除此群組？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         cnnConnection.RollbackTrans
         Screen.MousePointer = varSaveCursor
         Exit Function
      Else
         '先刪除子關係
         strSql = "delete from caserelation1 where (CR01,CR02,CR03,CR04) IN" + _
            " (SELECT B.CR05,B.CR06,B.CR07,B.CR08 FROM CASERELATION1 B WHERE B.CR01=" + CNULL(m_strCaseNo(0)) + _
            " AND B.CR02=" + CNULL(m_strCaseNo(1)) + " AND B.CR03=" + CNULL(m_strCaseNo(2)) + " AND B.CR04=" + CNULL(m_strCaseNo(3)) + ")"
         cnnConnection.Execute strSql, j
         '後刪除母關係
         strSql = "delete from caserelation1 Where cr01 = " + CNULL(m_strCaseNo(0)) + " And cr02 = " + CNULL(m_strCaseNo(1)) + " And cr03 = " + CNULL(m_strCaseNo(2)) + " And cr04 = " + CNULL(m_strCaseNo(3))
         cnnConnection.Execute strSql, j
      End If
'end 2023/12/27
   Else
      'Modified by Lydia 2023/10/12 MSFlexGrid1=>MSHFlexGrid1
      ReDim strRelation(3, Me.MSHFlexGrid1.Rows - 2)
      For i = 1 To Me.MSHFlexGrid1.Rows - 1
          strTmp = Split(Me.MSHFlexGrid1.TextMatrix(i, 1), "-")
      'end 2023/10/12
          If strTmp(0) = 馬德里案 Then
              strRelation(0, i - 1) = strTmp(0)
              strRelation(1, i - 1) = strTmp(1) & strTmp(2)
              strRelation(2, i - 1) = strTmp(3)
              strRelation(3, i - 1) = strTmp(4)
          Else
              For j = 0 To 3
                  strRelation(j, i - 1) = strTmp(j)
              Next
          End If
      Next
      
      'Added by Morgan 2022/11/8
      If m_bolCombine Then
            '先刪除子關係
            strSql = "delete from caserelation1 where (CR01,CR02,CR03,CR04) IN" + _
               " (SELECT B.CR05,B.CR06,B.CR07,B.CR08 FROM CASERELATION1 B WHERE B.CR01=" + CNULL(m_strCombine(0)) + _
               " AND B.CR02=" + CNULL(m_strCombine(1)) + " AND B.CR03=" + CNULL(m_strCombine(2)) + " AND B.CR04=" + CNULL(m_strCombine(3)) + ")"
            cnnConnection.Execute strSql, j
            '後刪除母關係
            strSql = "delete from caserelation1 Where cr01 = " + CNULL(m_strCombine(0)) + " And cr02 = " + CNULL(m_strCombine(1)) + " And cr03 = " + CNULL(m_strCombine(2)) + " And cr04 = " + CNULL(m_strCombine(3))
            cnnConnection.Execute strSql, j
      End If
      'end 2022/11/8
      
      'Modified by Morgan 2023/12/27
      'If lblSystem = 馬德里案 Then
      '   Call PUB_SaveCaseRelationData(lblSystem, lblTFCode(0) + IIf(lblTFCode(1) = "", "0", lblTFCode(1)), IIf(lblTFCode(2) = "", "0", lblTFCode(2)), IIf(lblTFCode(3) = "", "00", lblTFCode(3)), strRelation(), m_bolCombine)
      'Else
      '   Call PUB_SaveCaseRelationData(lblSystem, lblCode(0), IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2)), strRelation(), m_bolCombine)
      'End If
      Call PUB_SaveCaseRelationData(m_strCaseNo(0), m_strCaseNo(1), m_strCaseNo(2), m_strCaseNo(3), strRelation(), m_bolCombine)
      'end 2023/12/27
   End If
   cnnConnection.CommitTrans
   SaveData = True
   
ErrHand:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
   Screen.MousePointer = varSaveCursor
   
End Function

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
             Case 0 '確定
                        If SaveData Then
                           intLeaveKind = 1
                           bolLeave = True
                           Unload Me
                        End If
             Case 1, 2 '回前畫面, 結束
                        If Index = 2 Then
                           intLeaveKind = 0
                        Else
                           intLeaveKind = 1
                        End If
                        bolLeave = False
                        Unload Me
End Select
End Sub

Private Sub Form_Activate()
   txtSystem.SetFocus
   ReadRelationData
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    'edit by nickc 2007/02/05 不用 dll 了
    'Set obj011.Connection = cnnConnection
    If intWhereComeFrom <> 1 Then
        cmdok(0).Left = 7212
        cmdok(1).Left = 8040
        cmdok(2).Visible = False
    End If
    InitialGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If intWhereComeFrom = 1 Then
   If intLeaveKind = 1 Then
      If bFrm1103 = False Then
         m_form.Show
      Else
         frm1103_1.Show
      End If
   Else
      If bFrm1103 = False Then
         Unload m_form
      Else
        Unload frm1103_1
      End If
   End If
Else
   Where1103ComeFrom
End If
bFrm1103 = False
Set m_form = Nothing
End Sub

Private Sub lblNation_Change()
Dim strTemp As String

If lblNation = "" Then
   lblNationName = ""
Else
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetNation(lblNation, strTemp) Then
   If ClsPDGetNation(lblNation, strTemp) Then
      lblNationName = strTemp
   End If
End If
End Sub

Private Sub lblSystem_Change()
If lblSystem = 馬德里案 Then
   fraTF1.Visible = True
   fraElse1.Visible = False
Else
   fraTF1.Visible = False
   fraElse1.Visible = True
End If
End Sub

'Modified by Lydia 2023/10/12 MSFlexGrid1改為MSHFlexGrid1
Private Sub MSHFlexGrid1_Click()
Dim ii As Integer

With Me.MSHFlexGrid1
    If .row < 1 Then Exit Sub
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    For ii = 1 To .Rows - 1
        If ii <> .row Then .TextMatrix(ii, 0) = ""
    Next ii
    If .TextMatrix(.row, 0) = "" Then
        .TextMatrix(.row, 0) = "V"
    Else
        .TextMatrix(.row, 0) = ""
    End If
End With
End Sub


Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
lblNation = ""
lblCustomer = ""
End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
If txtSystem <> "" Then
    'Modify By Cheng 2002/11/19
    '在frm1103_1檢查即可
'   If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
'      ShowMsg MsgText(1056)
'      Cancel = True
'      txtSystem_GotFocus
'   End If
End If
End Sub

Private Sub txtCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
lblNation = ""
lblCustomer = ""
End Sub

Private Sub txtTFCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
lblNation = ""
lblCustomer = ""
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub

Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn2 (Index)
End Sub

Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNation As String

If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
        IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
        IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      lblNation = strNation
      lblCustomer = strCustomer
      CheckKeyIn1 = True
   End If
Else
   CheckKeyIn1 = True
End If
End Function

Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNation As String, i As Integer

If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      lblNation = strNation
      lblCustomer = strCustomer
      CheckKeyIn2 = True
   End If
Else
   CheckKeyIn2 = True
End If
End Function

'Add By Cheng 2003/08/14
Private Sub InitialGrid()
'Modified by Lydia 2023/10/12 MSFlexGrid1=>MSHFlexGrid1
With Me.MSHFlexGrid1
    .Cols = 5
    .Rows = 1
    .row = 0
    .col = 0: .Text = "V"
    .ColWidth(0) = 300: .ColAlignment(0) = flexAlignCenterCenter
    .col = 1: .Text = "本所案號"
    .ColWidth(1) = 1600: .ColAlignment(1) = flexAlignLeftCenter
    .col = 2: .Text = "案件名稱"
    .ColWidth(2) = 3600: .ColAlignment(2) = flexAlignLeftCenter
    .col = 3: .Text = "申請國家"
    .ColWidth(3) = 1000: .ColAlignment(3) = flexAlignLeftCenter
    .col = 4: .Text = "申請人"
    .ColWidth(4) = 3000: .ColAlignment(4) = flexAlignLeftCenter
End With
m_bolCombine = False 'Added by Morgan 2022/11/8
End Sub

'Modify by Morgan 2006/6/29
'相關卷號檔存檔
'Modified by Morgan 2023/12/14 +bolCombine
Public Function PUB_SaveCaseRelationData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, ByRef strRelation() As String, bolCombine As Boolean) As Boolean

   Dim i As Integer, j As Integer
   
   'Added by Morgan 2023/12/14
   '加入另一群組時本案自原群組刪除
   If bolCombine Then
      strSql = "Delete From CaseRelation1 Where CR01='" & strCode1 & "' And CR02='" & strCode2 & "' And CR03='" & strCode3 & "' And CR04='" & strCode4 & "' "
      cnnConnection.Execute strSql, j
      strSql = "Delete From CaseRelation1 Where CR05='" & strCode1 & "' And CR06='" & strCode2 & "' And CR07='" & strCode3 & "' And CR08='" & strCode4 & "' "
      cnnConnection.Execute strSql, j
   Else
   'end 2023/12/14
   
      '先刪除子關係
      strSql = "delete from caserelation1 where (CR01,CR02,CR03,CR04) IN" + _
         " (SELECT B.CR05,B.CR06,B.CR07,B.CR08 FROM CASERELATION1 B WHERE B.CR01=" + CNULL(strCode1) + _
         " AND B.CR02=" + CNULL(strCode2) + " AND B.CR03=" + CNULL(strCode3) + " AND B.CR04=" + CNULL(strCode4) + ")"
      cnnConnection.Execute strSql, j
      '後刪除母關係
      strSql = "delete from caserelation1 Where cr01 = " + CNULL(strCode1) + " And cr02 = " + CNULL(strCode2) + " And cr03 = " + CNULL(strCode3) + " And cr04 = " + CNULL(strCode4)
      cnnConnection.Execute strSql, j
   End If
   
'Modified by Morgan 2023/12/18 改以群組方式操作(本案也會列在清單內)
'   For i = 0 To UBound(strRelation, 2)
'      strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" + CNULL(strCode1) + "," + CNULL(strCode2) + "," + CNULL(strCode3) + "," + CNULL(strCode4) + _
'            "," + CNULL(strRelation(0, i)) + "," + CNULL(strRelation(1, i)) + "," + CNULL(strRelation(2, i)) + "," + CNULL(strRelation(3, i)) + ")"
'      cnnConnection.Execute strSql, j
'   Next
'
'   strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08)" + _
'      " select a.cr05,a.cr06,a.cr07,a.cr08,a.cr01,a.cr02,a.cr03,a.cr04 from caserelation1 a" + _
'      " where a.cr01=" + CNULL(strCode1) + " and a.cr02=" + CNULL(strCode2) + " and a.cr03=" + CNULL(strCode3) + " and a.cr04=" + CNULL(strCode4)
'   cnnConnection.Execute strSql, j
'
'   '本語法結果和上面一樣,改成一句語法
'   strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08)" & _
'      " select a.CR01,a.CR02,a.CR03,a.CR04,b.CR05,b.CR06,b.CR07,b.CR08" & _
'      " from caserelation1 a,caserelation1 b where a.cr05=" + CNULL(strCode1) + " and a.cr06=" + CNULL(strCode2) + " and a.cr07=" + CNULL(strCode3) + " and a.cr08=" + CNULL(strCode4) & _
'      " and b.cr01=a.cr05 and b.cr02=a.cr06 and b.cr03=a.cr07 and b.cr04=a.cr08 and not (b.cr05=a.cr01 and b.cr06=a.cr02 and b.cr07=a.cr03 and b.cr08=a.cr04)" & _
'      " and not exists(select * from caserelation1 c where c.CR01=a.CR01 and c.CR02=a.CR02 and c.CR03=a.CR03 and c.CR04=a.CR04 and c.CR05=b.CR05 and c.CR06=b.CR06 and c.CR07=b.CR07 and c.CR08=b.CR08)"
'   cnnConnection.Execute strSql, j
   For i = 1 To UBound(strRelation, 2)
      strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" + CNULL(strRelation(0, 0)) + "," + CNULL(strRelation(1, 0)) + "," + CNULL(strRelation(2, 0)) + "," + CNULL(strRelation(3, 0)) + _
            "," + CNULL(strRelation(0, i)) + "," + CNULL(strRelation(1, i)) + "," + CNULL(strRelation(2, i)) + "," + CNULL(strRelation(3, i)) + ")"
      cnnConnection.Execute strSql, j
   Next
   
   strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08)" + _
      " select a.cr05,a.cr06,a.cr07,a.cr08,a.cr01,a.cr02,a.cr03,a.cr04 from caserelation1 a" + _
      " where a.cr01=" + CNULL(strRelation(0, 0)) + " and a.cr02=" + CNULL(strRelation(1, 0)) + " and a.cr03=" + CNULL(strRelation(2, 0)) + " and a.cr04=" + CNULL(strRelation(3, 0))
   cnnConnection.Execute strSql, j
      
   '本語法結果和上面一樣,改成一句語法
   strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08)" & _
      " select a.CR01,a.CR02,a.CR03,a.CR04,b.CR05,b.CR06,b.CR07,b.CR08" & _
      " from caserelation1 a,caserelation1 b where a.cr05=" + CNULL(strRelation(0, 0)) + " and a.cr06=" + CNULL(strRelation(1, 0)) + " and a.cr07=" + CNULL(strRelation(2, 0)) + " and a.cr08=" + CNULL(strRelation(3, 0)) & _
      " and b.cr01=a.cr05 and b.cr02=a.cr06 and b.cr03=a.cr07 and b.cr04=a.cr08 and not (b.cr05=a.cr01 and b.cr06=a.cr02 and b.cr07=a.cr03 and b.cr08=a.cr04)" & _
      " and not exists(select * from caserelation1 c where c.CR01=a.CR01 and c.CR02=a.CR02 and c.CR03=a.CR03 and c.CR04=a.CR04 and c.CR05=b.CR05 and c.CR06=b.CR06 and c.CR07=b.CR07 and c.CR08=b.CR08)"
   cnnConnection.Execute strSql, j
'end 2023/12/18
   
   PUB_SaveCaseRelationData = True
End Function

'Copy from Dll011 by Morgan 2006/6/22
'讀取相關卷號檔
Public Function PUB_ReadCaseRelationData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, ByRef strRelation() As String) As Integer
   Dim i As Integer, j As Integer
On Error GoTo ErrHand
   PUB_ReadCaseRelationData = 0
   strSql = "select cr05,cr06,cr07,cr08 from caserelation1 where cr01=" + CNULL(strCode1) + " and cr02=" + CNULL(strCode2) + " and cr03=" + CNULL(strCode3) + " and cr04=" + CNULL(strCode4)
   'Added by Morgan 2023/12/18 改傳入案號也要列出
   strSql = strSql & " union select '" & strCode1 & "','" & strCode2 & "','" & strCode3 & "','" & strCode4 & "' from dual"
   'end 2023/12/18
   'Added by Morgan 2023/12/27 再加目前操作案號也要列出(併入群組時會與傳入案號不同)
   strSql = strSql & " union select '" & m_strCaseNo(0) & "','" & m_strCaseNo(1) & "','" & m_strCaseNo(2) & "','" & m_strCaseNo(3) & "' from dual"
   'end 2023/12/27
   
   strSql = strSql & " order by 1,2,3,4" 'Added by Morgan 2023/12/14
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         ReDim Preserve strRelation(3, j)
         For i = 0 To 3
            strRelation(i, j) = .Fields(i)
         Next
         .MoveNext
         j = j + 1
      Loop
      End With
      PUB_ReadCaseRelationData = 1
   End If
ErrHand:
   If Err.NUMBER <> 0 Then
      PUB_ReadCaseRelationData = -1
      MsgBox Err.Description
   End If
End Function
