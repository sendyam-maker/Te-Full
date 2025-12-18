VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_11 
   BorderStyle     =   1  '單線固定
   Caption         =   "介紹法務案件"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7185
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   2595
      Left            =   120
      TabIndex        =   34
      Top             =   4050
      Width           =   6945
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         Height          =   350
         Left            =   270
         Style           =   1  '圖片外觀
         TabIndex        =   37
         Top             =   90
         Width           =   765
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "回前畫面"
         Height          =   350
         Left            =   5580
         Style           =   1  '圖片外觀
         TabIndex        =   36
         Top             =   90
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1815
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   4
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|檔案名稱|副檔名說明|最後修改時間"
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(註:雙擊開啟)"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1230
         TabIndex        =   38
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboCaseType 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1140
      Style           =   2  '單純下拉式
      TabIndex        =   30
      Top             =   120
      Width           =   4065
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   180
      TabIndex        =   17
      Top             =   540
      Width           =   6855
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "TT"
         Top             =   210
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   21
         Text            =   "999999"
         Top             =   210
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   2610
         MaxLength       =   1
         TabIndex        =   20
         Text            =   "0"
         Top             =   210
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "00"
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txtCP10 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   18
         Text            =   "735"
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   2
         Left            =   3420
         TabIndex        =   29
         Top             =   270
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   1500
         X2              =   3210
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   6
         Left            =   150
         TabIndex        =   28
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   10
         Left            =   150
         TabIndex        =   27
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "介紹法務案件"
         Height          =   180
         Index           =   3
         Left            =   1710
         TabIndex        =   26
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "總收文號："
         Height          =   180
         Index           =   11
         Left            =   3420
         TabIndex        =   25
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "台一國際法律事務所"
         Height          =   180
         Index           =   1
         Left            =   4320
         TabIndex        =   24
         Top             =   270
         Width           =   1620
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   4320
         TabIndex        =   23
         Top             =   660
         Width           =   1725
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Height          =   2085
      Left            =   180
      TabIndex        =   5
      Top             =   1680
      Width           =   6855
      Begin VB.TextBox txtLOS04 
         Height          =   264
         Left            =   120
         MaxLength       =   70
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   810
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2220
         TabIndex        =   11
         Top             =   420
         Width           =   1905
         Begin VB.TextBox txtUserNo 
            Height          =   264
            Index           =   0
            Left            =   780
            MaxLength       =   6
            TabIndex        =   0
            Top             =   120
            Width           =   1035
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除>>"
            Height          =   285
            Index           =   0
            Left            =   30
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<<新增"
            Height          =   285
            Index           =   0
            Left            =   30
            TabIndex        =   12
            Top             =   120
            Width           =   735
         End
         Begin MSForms.Label lblName 
            Height          =   180
            Index           =   0
            Left            =   810
            TabIndex        =   14
            Top             =   450
            Width           =   1005
            VariousPropertyBits=   27
            Caption         =   "lblName"
            Size            =   "1773;317"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtCtrlDate 
         Height          =   300
         Left            =   1110
         MaxLength       =   7
         TabIndex        =   1
         Top             =   1260
         Width           =   1125
      End
      Begin MSForms.ComboBox cboLawMan 
         Height          =   300
         Left            =   1110
         TabIndex        =   2
         Top             =   1620
         Width           =   2385
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4207;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   630
         Index           =   0
         Left            =   1110
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   510
         Width           =   1125
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "1984;1111"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   4290
         TabIndex        =   33
         Top             =   150
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案源單號："
         Height          =   180
         Index           =   3
         Left            =   3390
         TabIndex        =   32
         Top             =   150
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法務人員："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1110
         TabIndex        =   9
         Top             =   150
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹日期："
         Height          =   180
         Index           =   7
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "管制日期："
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹人員："
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   6
         Top             =   510
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "確定"
      Height          =   350
      Index           =   1
      Left            =   5280
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   0
      Left            =   6150
      TabIndex        =   4
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件類型："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   180
      Width           =   900
   End
End
Attribute VB_Name = "frm090801_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/20 改成Form2.0 (cboLawMan,lstUsers,lblName)
'Created by Morgan 2020/4/17
Option Explicit

Public frmParent As Form '前一畫面
Public strLOS15 As String, strLOS04 As String, strLOS02 As String
Public bolIsPTCCase As Boolean, bolIsIPCase As Boolean, bolIsSuitCase As Boolean '是否PTC案,是否智財權案,是否訴訟案
Public m_strSaveFiles As String

Dim m_AttachPath As String

Private Sub cboCaseType_Click()
   If cboCaseType.Enabled = True And cboCaseType <> "" Then
      
      strExc(0) = Trim(Left(cboCaseType, 2))
      If strLOS02 <> strExc(0) Then
         strLOS02 = strExc(0)
         If Left(strLOS02, 1) = "A" Then
            txtCP10 = "735" '案件性質
            
         ElseIf strLOS02 = "B1" Then
            If MsgBox("是否需要智慧所配合？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
               strLOS02 = "A4"
               txtCP10 = "735" '案件性質
            Else
               txtCP10 = "736" '案件性質
            End If
         ElseIf Left(strLOS02, 1) = "B" Then
            txtCP10 = "736" '案件性質
         Else
            txtCP10 = ""
         End If
         If txtCP10 <> "" Then
            Call ClsPDGetCaseProperty("TT", txtCP10, strExc(0))
            Label2(3).Caption = strExc(0)
         End If
         SetFramePos
      End If
   End If
End Sub

Private Sub cboLawMan_Change()
   If Len(cboLawMan) = 5 Then
      SetcboLawMan cboLawMan
   End If
End Sub

Private Sub cboLawMan_GotFocus()
   'SendMessage cboLawMan.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

Private Sub cboLawMan_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboLawMan_Validate(Cancel As Boolean)
'   Dim ii As Integer
'   For ii = 0 To cboLawMan.ListCount - 1
'      If InStr(cboLawMan.List(ii), cboLawMan) > 0 Then
'         cboLawMan.ListIndex = ii
'      End If
'   Next
   SetcboLawMan cboLawMan.Text
End Sub

Private Sub cmdAdd_Click(Index As Integer)
   AddlstUsers Index
   txtLOS04 = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub

Private Sub cmdBack_Click()
   Me.Caption = Me.Tag
   Me.Height = 4230
   Frame4.Top = Me.Height + 500
End Sub

Private Sub cmdOK_Click(Index As Integer)
   If Index = 1 Then If TxtValidate = False Then Exit Sub
   
   With frmParent
   .iReturn = Index
   If Index = 1 Then
      .strTTCP10 = txtCP10 'TT收文性質
      .strIntroducer = txtLOS04 '介紹人
      .strCtrlDate = DBDATE(txtCtrlDate)
      .strLawMan = Left(cboLawMan, 5)  '法務人員
      .strTTSaveFiles = m_strSaveFiles
      .strLSourceType = strLOS02
   End If
   End With
   Unload Me
End Sub

Private Sub cmdRemove_Click(Index As Integer)
   RemovelstUsers Index
   txtLOS04 = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub

Private Sub Form_Load()
   Dim oLabel As LABEL
   
   m_AttachPath = App.path & "\" & strUserNum
   Me.Height = 4230
   
   MoveFormToCenter Me, True
   lstUsers(0).Clear
   txtUserNo(0) = ""
   lblName(0) = ""
   
   For Each oLabel In Label2
      oLabel.BackColor = &H8000000F
   Next
   SetcboLawMan
   SetCaseType
   ReadData
   Me.Tag = Me.Caption
End Sub
'Modified by Morgan 2021/7/27 取消 C 類並調整其他類說明
Private Sub SetCaseType()
   cboCaseType.Clear
   If Not (bolIsIPCase = True And bolIsSuitCase = True) Then
      cboCaseType.AddItem "A  法律案、一般訴訟案、法律顧問"
   End If
   cboCaseType.AddItem "B1 智財民事訴訟、智財刑事訴訟"
   cboCaseType.AddItem "B2 專利/商標行政訴訟、專利上訴..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   If Not frmParent Is Nothing Then
'      frmParent.Enabled = True
'      frmParent.Show
'      frmParent.ZOrder
'   End If
   Set frm090801_11 = Nothing
End Sub

Private Function ComposeListX(p_index As Integer) As String
   'Modified by Morgan 2022/1/21
   'strExc(1) = ""
   'If lstUsers(p_index).ListCount > 0 Then
   '   strExc(1) = PUB_Num2Id(lstUsers(p_index).ItemData(0))
   '   For intI = 1 To lstUsers(p_index).ListCount - 1
   '      strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ItemData(intI))
   '   Next
   'End If
   'ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
End Function

Private Sub MSHFlexGrid1_Click()
   If MSHFlexGrid1.MouseRow > 0 Then
      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = "" Then
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = "V"
      Else
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = ""
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
   Dim stFileName As String, hLocalFile As Long, idx As Integer
   Dim stFtpPath As String
   
   If MSHFlexGrid1.MouseRow > 0 Then
      idx = MSHFlexGrid1.MouseRow
      stFileName = MSHFlexGrid1.TextMatrix(idx, 3)
      If PUB_GetAttachFile_CPP("LOS" & strLOS15, stFileName, m_AttachPath, False) Then
         ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
      End If
   End If
End Sub

Private Sub txtCtrlDate_GotFocus()
   TextInverse txtCtrlDate
End Sub

Private Sub txtCtrlDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCtrlDate_Validate(Cancel As Boolean)
   If txtCtrlDate <> "" Then
      If ChkDate(txtCtrlDate) Then
         If txtCtrlDate < strSrvDate(2) Then
            MsgBox "管制日期不可早於系統日！", vbExclamation
            Cancel = True
         Else
            strExc(0) = TransDate(PUB_GetWorkDay1(txtCtrlDate.Text, True), 1)
            If txtCtrlDate <> strExc(0) Then
               MsgBox "管制日期將自動調整為工作日！", vbInformation
               txtCtrlDate = strExc(0)
            End If
         End If
      Else
         'txtCtrlDate_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub txtUserNo_Change(Index As Integer)
   Dim strTempName As String
   If Len(txtUserNo(Index)) = 5 Then
      If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
         lblName(Index) = strTempName
      End If
   Else
      lblName(Index) = ""
   End If
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" And lblName(Index) = "" Then
         If txtUserNo(Index) > "6" And txtUserNo(Index) < "F" Then
            If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
               lblName(Index) = strTempName
            End If
         ElseIf GetIdFromName(txtUserNo(Index), strExc(1)) Then
            lblName(Index) = txtUserNo(Index)
            txtUserNo(Index) = strExc(1)
         End If
         If lblName(Index) = "" Then
            MsgBox "員工編號輸入錯誤！", vbExclamation
            Cancel = True
         End If
      End If
   End If
End Sub

Private Function GetIdFromName(ByVal pName As String, ByRef pID As String) As Boolean
   strExc(0) = "select st01,st02 from staff where st02='" & ChgSQL(pName) & "' and st04='1' and st01>'6' and st01<'F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         pID = RsTemp.Fields("st01")
         GetIdFromName = True
      Else
         MsgBox "員工名稱重複，請直接輸入員工編號！"
      End If
   Else
      MsgBox "該員工名稱不存在！"
   End If
End Function

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      'Modified by Morgan 2022/1/22
      'For idx = 0 To lstUsers(p_idx).ListCount - 1
      '   If lstUsers(p_idx).ItemData(idx) = PUB_Id2Num(txtUserNo(p_idx)) Then
      '      MsgBox "員工已存在於介紹人員清單中！"
      '      txtUserNo(p_idx).SetFocus
      '      txtUserNo_GotFocus p_idx
      '      bFound = True
      '      Exit For
      '   End If
      'Next
      If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
         MsgBox "員工已存在於介紹人員清單中！"
         txtUserNo(p_idx).SetFocus
         txtUserNo_GotFocus p_idx
         bFound = True
      End If
      'end 2022/1/22
      
      If bFound = False Then
         '收文人放第1個(不可移除)，其他往後加
         lstUsers(p_idx).AddItem lblName(p_idx), lstUsers(p_idx).ListCount
         'Modified by Morgan 2022/1/22
         'lstUsers(p_idx).ItemData(lstUsers(p_idx).ListCount - 1) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = lstUsers(p_idx).Tag & "," & txtUserNo(p_idx)
         'end 2022/1/22
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   Dim idx As Integer, ii As Integer
   If lstUsers(p_idx).ListCount > 0 Then
      For ii = lstUsers(p_idx).ListCount - 1 To 0 Step -1
         If lstUsers(p_idx).Selected(ii) = True Then
            If ii = 0 Then
               MsgBox "第1介紹人(TT收文人員)不可移除！", vbExclamation
            'Modified by Morgan 2022/1/22
            'Else
            '   lstUsers(p_idx).RemoveItem ii
            'end 2022/1/22
               lstUsers(p_idx).Selected(ii) = False
            End If
         End If
      Next
   End If
   'Added by Morgan 2022/1/22
   lstUsers(p_idx).Tag = PUB_RemoveListBox2(lstUsers(p_idx), lstUsers(p_idx).Tag)
   'end 2022/1/22
End Sub

Public Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = ""
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0 order by instr('" & p_stNums & "',st01) desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         .MoveFirst
         Do While Not .EOF
            lstUsers(p_idx).AddItem "" & .Fields(1), 0
            'Modified byMorgan 2022/1/21
            'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
            lstUsers(p_idx).Tag = lstUsers(p_idx).Tag & "," & .Fields(0)
            'end 2022/1/21
            .MoveNext
         Loop
         End With
         lstUsers(p_idx).Tag = Mid(lstUsers(p_idx).Tag, 2) 'Added byMorgan 2022/1/21
      End If
   End If
   txtLOS04 = ComposeListX(p_idx)
End Sub

Private Sub SetcboLawMan(Optional pLawMan As String)
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   If pLawMan <> "" Then
      For intQ = 0 To cboLawMan.ListCount - 1
         If InStr(cboLawMan.List(intQ), pLawMan) > 0 Then
            cboLawMan.ListIndex = intQ
            Exit For
         End If
      Next
   Else
      cboLawMan.Clear
      'Modified by Morgan 2020/6/18 律師不要列--杜經理
      'Modified by Morgan 2020/7/17 律師改要列且列前面--麗真, 統一用部門,員工號排序--秀玲
      stSQL = "select st01||' '||st02 from staff where st04='1' and st03 like 'L%' and st01>'6' and st01<'F' order by st03,st01"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         Do While Not RsQ.EOF
            cboLawMan.AddItem RsQ(0)
            RsQ.MoveNext
         Loop
      End If
   End If
   Set RsQ = Nothing
End Sub
Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   

   If cboCaseType.ListIndex = -1 Then
      MsgBox "請選擇案件類型！", vbExclamation
      cboCaseType.SetFocus
      Exit Function
   End If
   
   If lstUsers(0).ListCount = 0 Then
      MsgBox "介紹人員不可空白！", vbExclamation
      txtUserNo(0).SetFocus
      txtUserNo_GotFocus 0
      Exit Function
   End If
   
   If txtCtrlDate <> "" Then
      txtCtrlDate_Validate bCancel
      If bCancel = True Then Exit Function
   End If
   
   If cboLawMan <> "" And cboLawMan.ListIndex = -1 Then
      MsgBox "法務人員輸入錯誤！", vbExclamation
      cboLawMan.SetFocus
      Exit Function
   End If
   
   
   TxtValidate = True
End Function

Private Sub ReadData()
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   Dim stCP10 As String
   Dim strShowType As String
   
   If strLOS15 <> "" Then
      stSQL = "select * from lawofficesource,caseprogress,servicepractice,customer,casepropertymap,staff" & _
         " where LOS15='" & strLOS15 & "' and cp09(+)=LOS10" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)" & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=los03"
         
       Label2(0) = strLOS15
       Label1(3).Visible = True
       Label2(0).Visible = True
   
   ElseIf strLOS02 <> "C" Then
   
      If Left(strLOS02, 1) = "B" Then
         stCP10 = "736"
      Else
         stCP10 = "735"
      End If
      stSQL = "select * from servicepractice,customer,casepropertymap" & _
         " where sp01='TT' and sp02='999999' and sp03='0' and sp04='00'" & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and cpm01(+)=sp01 and cpm02(+)='" & stCP10 & "'"
   End If
   
   If stSQL = "" Then
      Label2(4).Caption = strSrvDate(2) '介紹日期
   Else
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With RsQ
         Text1(6) = "" & .Fields("sp01")
         Text1(7) = "" & .Fields("sp02")
         Text1(8) = "" & .Fields("sp03")
         Text1(9) = "" & .Fields("sp04")
         Label2(1) = "" & .Fields("cu04") '申請人
         txtCP10 = "" & .Fields("cpm02") '案件性質
         Label2(3).Caption = "" & .Fields("cpm03")
         If strLOS15 = "" Then
            Label2(4).Caption = strSrvDate(2) '介紹日期
         Else
            strLOS02 = "" & .Fields("LOS02")
            Label2(2).Caption = "" & .Fields("cp09") 'TT總收文號
            Label2(4).Caption = TransDate(.Fields("los12"), 1) '介紹日期
            strLOS04 = "" & .Fields("LOS04") '介紹人
            If Not IsNull(.Fields("los16")) Then
               txtCtrlDate = TransDate("" & .Fields("los16"), 1) '管制日期
            End If
            If Not IsNull(.Fields("los03")) Then
               cboLawMan = .Fields("los03") & " " & .Fields("st02") '法務人員
            End If
         End If
         
         End With
      End If
   End If
   
   If strLOS04 <> "" Then SetlstUsers 0, strLOS04
   
   If strLOS02 <> "" Then
      cboCaseType.Enabled = False
      
      If strLOS02 = "A4" Then
         strShowType = "B1"
      'Modified by Morgan 2020/9/30
      'ElseIf strLOS02 = "A3" Then
      ElseIf Left(strLOS02, 1) = "A" Then
         strShowType = "A"
      'end 2020/9/3
      Else
         strShowType = strLOS02
      End If
      For intQ = 0 To cboCaseType.ListCount - 1
         If InStr(cboCaseType.List(intQ), strShowType) = 1 Then
            cboCaseType.ListIndex = intQ
         End If
      Next
      
   End If
   
   SetFramePos
   Set RsQ = Nothing
End Sub

Private Sub SetFramePos()
   Static lngPos As Long, lngHeight As Long
   
   If lngPos = 0 Then lngPos = Frame3.Top
   If lngHeight = 0 Then lngHeight = Me.Height
   
   '只有A(A1、A2)才能多人，A3、A4、B、C都只能1人(只有點數且只給第1人)
   If strLOS02 = "A1" Or strLOS02 = "A2" Then
      Frame2.Visible = True
   Else
      Frame2.Visible = False
   End If
   
   'Removed by Morgan 2021/7/27 取消C類, 不必再設定位置
   'If strLOS02 = "C" Then
   '   Frame1.Visible = False
   '   Frame3.Top = Frame1.Top
   '   Me.Height = Me.Height - Frame1.Height
   'Else
   '   Frame3.Top = lngPos
   '   Frame1.Visible = True
   '   Me.Height = lngHeight
   'End If
   'end 2021/7/27
End Sub

Public Sub SetReadOnly()
   cmdOK(1).Visible = False
   Frame3.Enabled = False
   Frame2.Visible = False
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 3500, 2300)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 0
   .FormatString = "V|檔案名稱|最後修改時間"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

'下載
Private Sub cmdSaveAtt_Click()
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer
  
   Dim strCP10 As String
   Dim strName2 As String
   
   
   stFolderPath = PUB_Getdesktop
   stFolderPath = PUB_GetFolder(Me.hWnd, stFolderPath, "請選取資料夾:")
   
   If Trim(stFolderPath) = "" Then Exit Sub
         
   stFileName = ""
   bMultiFile = False
   
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If (.TextMatrix(ii, 0) = "V" Or .TextMatrix(ii, 0) = "v") Then
         stFileName = Trim(.TextMatrix(ii, 3))
         Exit For
      End If
   Next ii
   
   Screen.MousePointer = vbHourglass
   If stFileName = "" Then
      MsgBox "請先勾選檔案！"
   Else
      For ii = 1 To .Rows - 1
         If (.TextMatrix(ii, 0) = "V" Or .TextMatrix(ii, 0) = "v") Then
            stFileName = Trim(.TextMatrix(ii, 3))
            stFullName = stFolderPath & "\" & stFileName
            
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋??", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            
            If stFullName <> "" Then
               If PUB_GetAttachFile_CPP("LOS" & strLOS15, stFileName, stFolderPath, False) = False Then
                  MsgBox "無法儲存檔案[ " & stFullName & " ]！", vbCritical
                  GoTo RunExit
               End If
            End If
         End If
      Next ii
      MsgBox "下載完成！"
   End If
   End With
RunExit:
   Screen.MousePointer = vbDefault
End Sub
