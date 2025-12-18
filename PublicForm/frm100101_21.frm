VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_21 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
   Caption         =   "國內潛在客戶資料查詢"
   ClientHeight    =   6000
   ClientLeft      =   1440
   ClientTop       =   2320
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8950
   Begin VB.CommandButton CmdOk1 
      Caption         =   "被介紹者"
      Height          =   400
      Index           =   3
      Left            =   5810
      Style           =   1  '圖片外觀
      TabIndex        =   47
      Top             =   45
      Width           =   950
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Index           =   0
      ItemData        =   "frm100101_21.frx":0000
      Left            =   5355
      List            =   "frm100101_21.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2616
      Width           =   1035
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   6780
      TabIndex        =   0
      Top             =   45
      Width           =   1230
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8040
      TabIndex        =   1
      Top             =   45
      Width           =   800
   End
   Begin VB.Label SpecCU 
      ForeColor       =   &H000000FF&
      Height          =   408
      Left            =   2568
      TabIndex        =   48
      Top             =   72
      Width           =   3468
   End
   Begin VB.Label Label1 
      Caption         =   "為關係企業"
      Height          =   252
      Index           =   8
      Left            =   7536
      TabIndex        =   46
      Top             =   2904
      Width           =   1044
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   28
      Left            =   7860
      TabIndex        =   42
      Top             =   2289
      Width           =   330
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "582;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   3
      Left            =   1215
      TabIndex        =   38
      Top             =   507
      Width           =   7335
      VariousPropertyBits=   671105051
      MaxLength       =   79
      Size            =   "12938;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   23
      Left            =   1215
      TabIndex        =   37
      Top             =   804
      Width           =   5415
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   26
      Left            =   1215
      TabIndex        =   36
      Top             =   1695
      Width           =   5415
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   25
      Left            =   1215
      TabIndex        =   35
      Top             =   1398
      Width           =   5415
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   24
      Left            =   1215
      TabIndex        =   34
      Top             =   1101
      Width           =   5415
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   27
      Left            =   1215
      TabIndex        =   33
      Top             =   1992
      Width           =   7335
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "12938;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   1
      Left            =   1215
      TabIndex        =   16
      Top             =   210
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1926;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   2
      Left            =   2295
      TabIndex        =   15
      Top             =   210
      Width           =   255
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   705
      Index           =   15
      Left            =   1215
      TabIndex        =   14
      Top             =   3180
      Width           =   7545
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13309;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   12
      Left            =   1215
      TabIndex        =   13
      Top             =   2586
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1508;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   4
      Left            =   1215
      TabIndex        =   12
      Top             =   2289
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1508;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   14
      Left            =   1215
      TabIndex        =   10
      Top             =   2883
      Width           =   1305
      VariousPropertyBits=   671105051
      MaxLength       =   12
      Size            =   "2302;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   11
      Left            =   4725
      TabIndex        =   9
      Top             =   2289
      Width           =   330
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "582;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   8
      Left            =   1215
      TabIndex        =   8
      Top             =   4476
      Width           =   2955
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "5212;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   9
      Left            =   5355
      TabIndex        =   7
      Top             =   4476
      Width           =   2955
      VariousPropertyBits=   671105051
      MaxLength       =   150
      Size            =   "5212;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   7
      Left            =   1215
      TabIndex        =   6
      Top             =   4179
      Width           =   2955
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "5212;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   6
      Left            =   5355
      TabIndex        =   5
      Top             =   3882
      Width           =   2955
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "5212;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   5
      Left            =   1215
      TabIndex        =   4
      Top             =   3882
      Width           =   2955
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "5212;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   855
      Index           =   10
      Left            =   1215
      TabIndex        =   3
      Top             =   4770
      Width           =   7545
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "13309;1508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPOC 
      Height          =   300
      Index           =   16
      Left            =   3036
      TabIndex        =   2
      Top             =   2880
      Width           =   1068
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1884;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   90
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5670
      Width           =   7140
      VariousPropertyBits=   -2147467233
      BackColor       =   16777215
      Size            =   "12594;529"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "ＰＳ：非研發處或專利處程序之潛在客戶, 三年內無往來記錄者, 系統會自動刪除．"
      ForeColor       =   &H000000FF&
      Height          =   720
      Index           =   7
      Left            =   6840
      TabIndex        =   44
      Top             =   840
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否寄發專利雙週報：      （N:不寄）"
      Height          =   180
      Index           =   4
      Left            =   6045
      TabIndex        =   43
      Top             =   2349
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱（中）："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   41
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱（日）："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱（英）："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   39
      Top             =   870
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "編　　號："
      Height          =   210
      Index           =   0
      Left            =   285
      TabIndex        =   32
      Top             =   225
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Index           =   19
      Left            =   285
      TabIndex        =   31
      Top             =   3195
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "開發日期：                    ( 西元 )"
      Height          =   180
      Index           =   13
      Left            =   285
      TabIndex        =   30
      Top             =   2646
      Width           =   2370
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   29
      Top             =   2312
      Width           =   1245
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2196;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國　　籍："
      Height          =   180
      Index           =   5
      Left            =   285
      TabIndex        =   28
      Top             =   2349
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "開發人員："
      Height          =   180
      Index           =   3
      Left            =   4425
      TabIndex        =   27
      Top             =   2646
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "狀　　態："
      Height          =   180
      Index           =   21
      Left            =   285
      TabIndex        =   26
      Top             =   2925
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否寄電子報：      （N:不寄）"
      Height          =   180
      Index           =   17
      Left            =   3465
      TabIndex        =   25
      Top             =   2349
      Width           =   2415
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "行動電話："
      Height          =   180
      Index           =   15
      Left            =   285
      TabIndex        =   24
      Top             =   4536
      Width           =   915
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "電　話2："
      Height          =   180
      Index           =   12
      Left            =   4515
      TabIndex        =   23
      Top             =   3942
      Width           =   810
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "E-MAIL："
      Height          =   180
      Index           =   11
      Left            =   4545
      TabIndex        =   22
      Top             =   4536
      Width           =   780
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "傳　真1："
      Height          =   180
      Index           =   10
      Left            =   390
      TabIndex        =   21
      Top             =   4239
      Width           =   810
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "電　話1："
      Height          =   180
      Index           =   9
      Left            =   390
      TabIndex        =   20
      Top             =   3942
      Width           =   810
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "地　　址："
      Height          =   180
      Index           =   28
      Left            =   300
      TabIndex        =   19
      Top             =   4860
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "與 "
      Height          =   252
      Index           =   107
      Left            =   2760
      TabIndex        =   18
      Top             =   2904
      Width           =   276
   End
   Begin MSForms.Label lbl1 
      Height          =   252
      Index           =   0
      Left            =   4128
      TabIndex        =   17
      Top             =   2904
      Width           =   3324
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5863;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; textCUID、lbl1(index)、txtPOC(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

Public cmdState As Integer
Dim strTmp As String
Dim rsContact As ADODB.Recordset
Dim m_bReadGrid As Boolean
Dim oText As Control
Dim idx As Integer


Private Sub DataGrid1_Click()
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   textCUID.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_21 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      'Add by Amy 2023/07/12
      Case 3 '被介紹者
         If CmdOk1(3).BackColor <> &HFFFF80 Then
            MsgBox "無被介紹者資料"
            Exit Sub
         End If
         If PUB_CheckFormExist("frm050705_1") Then
              MsgBox "請先關閉〔被介紹資料〕的畫面！", vbInformation
              Exit Sub
         End If
         cmdState = -1
         Me.Enabled = False
         If fnSaveParentForm(Me) = False Then
            Me.Enabled = True
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         Call ShowFrm050705_1
         Me.Enabled = True
         Screen.MousePointer = vbDefault
         Exit Sub
   End Select
End Sub

Sub StrMenu()
   Dim strKey  As String, strKey1 As String
   If Mid(Me.Tag, 10, 1) = "-" Then
      strKey = Left(Me.Tag, 9)
      strKey1 = Mid(Me.Tag, 11)
   Else
      strKey = Me.Tag
   End If
   
   'Add By Sindy 2011/01/03 檢查國內外權限
   If CheckSR12(strKey) = False Then
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   pub_QL05 = ";潛在客戶編號：" & strKey & "(國內潛在客戶基本資料)" 'Add By Sindy 2025/8/13
   
   'Added by Lydia 2024/07/01 新增國內潛在客戶之不得宣傳; 參考100101_10 'Added by Lydia 2023/01/19 往來紀錄中有「A14客戶名稱資訊不得宣傳」者，在申請人/代理人資料查詢首頁提示
   strExc(0) = "SELECT ac03 as memo FROM allcode where AC01='11' and ac02='A14' and exists (select * from contactrecord where instr(cr05,'A14')>0 and substr(cr03,1,8)='" & Mid(strKey, 1, 8) & "' and substr(cr03,9,1)='" & Mid(strKey, 9, 1) & "') "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       SpecCU.Caption = SpecCU.Caption & IIf(Trim(SpecCU.Caption) <> "", "；", "") & RsTemp.Fields("memo")
       SpecCU.Font.Size = 14
       SpecCU.AutoSize = True
   End If
   'end 2024/07/01
   
   strExc(0) = "select * from potcustomer1 where poc01='" & Left(strKey, 8) & "' and poc02='" & Mid(strKey, 9) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pub_QL04 <> "" Then InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2025/8/13
      ShowRecord RsTemp
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   'Add by Amy 2023/07/12 被介紹者
   CmdOk1(3).BackColor = &H8000000F
   If Pub_GetXYSource(2, Left(strKey, 8)) = True Then
      CmdOk1(3).BackColor = &HFFFF80
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset)
   Dim rsPOC As ADODB.Recordset
   Dim CUID(1 To 6) As String
   
   ClearField
   SetCtrlReadOnly True
   Set rsPOC = p_Rst.Clone
   With rsPOC
      If .RecordCount > 0 Then
         For Each oText In txtPOC
            idx = oText.Index
            oText.Text = "" & .Fields("POC" & Format(idx, "0#"))
         Next
         
         CUID(1) = "" & .Fields("POC17")
         CUID(2) = "" & .Fields("POC18")
         CUID(3) = "" & .Fields("POC19")
         CUID(4) = "" & .Fields("POC20")
         CUID(5) = "" & .Fields("POC21")
         CUID(6) = "" & .Fields("POC22")
         
         '國籍
         If Trim(txtPOC(4)) = "" Then
            LBL1(1).Caption = ""
         Else
            If ClsPDGetNation(Left(txtPOC(4), 3), strTmp) = True Then
               LBL1(1).Caption = strTmp
            End If
         End If
         
         '開發人員
         If Not IsNull(.Fields("POC13")) Then
            strExc(0) = "select st02 from staff where instr('" & .Fields("POC13") & "',st01)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTmp = RsTemp.GetString(, , , ",")
               SetList lstUsers(0), strTmp
            End If
         End If
         
         If Not IsNull(txtPOC(16)) And Trim(txtPOC(16)) <> "" Then
            GetCustData (txtPOC(16))
         End If
      End If
   End With
   UpdateCUID CUID, textCUID
End Sub

Private Sub ClearField()
   Dim oLabel As Control
   For Each oText In txtPOC
      oText.Text = Empty
   Next
   For Each oLabel In LBL1
      oLabel.Caption = Empty
   Next
   textCUID = ""
   lstUsers(0).Clear
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Control)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   'Modified by Lydia 2022/01/22 vbCrLf=> String(6, " ")
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtPOC
      oText.Locked = bLocked
   Next
End Sub

Private Function GetCustData(p_stCust As String) As Boolean
   Dim aiOrder(1 To 3) As Integer
   LBL1(0) = ""
   Select Case Left(p_stCust, 1)
      Case "X"
         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
'      Case Else
'         MsgBox "關係企業必須為 X 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   LBL1(0) = ""
   If intI = 1 Then
      For intI = 1 To 3
         If Not IsNull(RsTemp(intI)) Then
            LBL1(0) = RsTemp(intI)
            Exit For
         End If
      Next
      GetCustData = True
   End If
End Function

'Add by Amy 2023/07/12 顯示被介紹者資料
Private Sub ShowFrm050705_1()
   Dim stName As String
   
   If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   '中->英->日
   If Trim(txtPOC(3)) = MsgText(601) Then
      If Trim(txtPOC(23)) = MsgText(601) Then
         stName = txtPOC(27) '日
      Else
         stName = txtPOC(23)
         If Trim(txtPOC(24)) <> MsgText(601) Then
            stName = stName & " " & txtPOC(24)
         End If
         If Trim(txtPOC(25)) <> MsgText(601) Then
            stName = stName & " " & txtPOC(25)
         End If
         If Trim(txtPOC(26)) <> MsgText(601) Then
            stName = stName & " " & txtPOC(26)
         End If
      End If
   Else
      stName = txtPOC(3) '中
   End If
   frm050705_1.txtNo = Left(txtPOC(1), 8)
   frm050705_1.LBL1(0) = txtPOC(4) '國籍code
   frm050705_1.LBL1(1) = LBL1(1) '國籍
   frm050705_1.LBL1(3) = stName
   frm050705_1.SetParent Me
   frm050705_1.QueryData
   frm050705_1.Show
End Sub
