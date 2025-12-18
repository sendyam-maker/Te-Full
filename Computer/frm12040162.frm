VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040162 
   BorderStyle     =   1  '單線固定
   Caption         =   "查詢特殊置換字對照表"
   ClientHeight    =   1872
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7512
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1872
   ScaleWidth      =   7512
   Begin VB.ComboBox cboClass 
      Height          =   276
      ItemData        =   "frm12040162.frx":0000
      Left            =   1368
      List            =   "frm12040162.frx":0002
      TabIndex        =   2
      Text            =   "cboClass"
      Top             =   1440
      Width           =   1500
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6705
      Top             =   720
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":0320
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":063C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":0818
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":0E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":116C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":17A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":1AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040162.frx":1DDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7512
      _ExtentX        =   13250
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox Text2 
      Height          =   348
      Left            =   1368
      TabIndex        =   1
      Top             =   1080
      Width           =   708
      VariousPropertyBits=   746604571
      MaxLength       =   1
      ScrollBars      =   2
      Size            =   "1249;609"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox Text1 
      Height          =   348
      Left            =   1368
      TabIndex        =   0
      Top             =   720
      Width           =   708
      VariousPropertyBits=   746604571
      MaxLength       =   1
      ScrollBars      =   2
      Size            =   "1249;609"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label lblCode 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   1092
      Width           =   1104
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "說明："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   804
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "類　別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1476
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "替換字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   756
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "統一字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1116
      Width           =   720
   End
End
Attribute VB_Name = "frm12040162"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2023/08/15
Option Explicit

'0:新增 1:修改 2:查詢 3:瀏覽
Dim ActionEdit As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim adoR As New ADODB.Recordset
Dim strQ As String

Private Sub Action(Index As Integer)
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub
   Select Case Index
      Case 1 '新增
         FormReset
         Text1.SetFocus
         ActionEdit = 0
      Case 2 '修改
         Text1.SetFocus
         Text1_GotFocus
         ActionEdit = 1
      Case 3 '刪除
         If MsgBox("是否確定要刪除??", vbYesNo + vbDefaultButton2) = vbYes Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               FormReset
               RsAction 3
            End If
         Else
            Exit Sub
         End If
         
      Case 4 '查詢
         FormReset
         Text2.Enabled = False
         Text1.SetFocus
         ActionEdit = 2
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3
      Case 11 '按下確定
         Select Case ActionEdit
            Case 0, 1 '新增,修改
               If TxtValidate = True Then
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                  Else
                     ActionEdit = 3
                  End If
               End If
            Case 2
               If cboClass = MsgText(601) Then
                  MsgBox Label1(2) & "不可為空白！"
                  Exit Sub
               End If
               RsAction 4
               Text2.Enabled = True
         End Select
         
      Case 12 '按下取消
         Text1 = Text1.Tag
         If Text1 <> MsgText(601) Then
            RsAction 4
         End If
      Case 14 '結束
         If ActionEdit = 0 Or ActionEdit = 1 Then
            If MsgBox("尚未存檔，是否確定要結束??", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         Unload Me
         Exit Sub
   End Select
   TxtLock ActionEdit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2: Action 1 '新增
      Case vbKeyF3: Action 2 '修改
      Case vbKeyF5: Action 3 '刪除
      Case vbKeyF4: Action 4 '查詢
      Case vbKeyHome: Action 6 '第一筆
      Case vbKeyPageUp: Action 7 '前一筆
      Case vbKeyPageDown: Action 8 '後一筆
      Case vbKeyEnd: Action 9 '最後筆
      Case vbKeyF9: Action 11 '確定
      Case vbKeyF10: Action 12 '取消
      Case vbKeyEscape: Action 14 '結束
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   SetcboClass
   OpenTable
   lblMemo.Caption = lblMemo.Caption & vbCrLf & _
                                       "統一字及替換字都為key" & vbCrLf & _
                                       "類別為「1.符號」者" & vbCrLf & _
                                       "替換字請輸入[全型空白]"
                                       
End Sub

Private Sub FormReset()
   Text1.Text = ""
   Text2.Text = ""
   lblCode(1) = ""
   cboClass = ""
End Sub

Private Sub SetcboClass()
   cboClass.Clear
   cboClass.AddItem ""
   cboClass.AddItem "1.符號"
   cboClass.AddItem "2.文字"
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub

Private Sub OpenTable()
   strQ = "Select * From SpecWord Order By 1"
   
   adoR.CursorLocation = adUseClient
   adoR.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoR.RecordCount > 0 Then
      adoR.MoveFirst
      FormRefresh
   End If
End Sub

Private Sub FormRefresh()
   If Not (adoR.EOF And adoR.BOF) Then
      Text1 = adoR.Fields("SW01")
      Text1.Tag = Text1
      Text2 = adoR.Fields("SW02")
      Text2_Validate False
      cboClass = ClassVal(adoR.Fields("SW03"))
      ActionEdit = 3
   End If
End Sub

Private Function ClassVal(stSW03 As String) As String
   Select Case Val(stSW03)
      Case "1"
         ClassVal = "符號"
      Case "2"
         ClassVal = "文字"
   End Select
   ClassVal = stSW03 & "." & ClassVal
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set adoR = Nothing
   Set frm12040162 = Nothing
End Sub

Private Sub RsAction(ByVal pCmd As Integer)
 
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case pCmd
      Case 0 '第一筆
         adoR.MoveFirst
         FormRefresh
      Case 1 '前一筆
         adoR.MovePrevious
         If adoR.BOF Then
            DataErrorMessage 6
            adoR.MoveFirst
         End If
         FormRefresh
      Case 2 '後一筆
         adoR.MoveNext
         If adoR.EOF Then
            DataErrorMessage 7
            adoR.MoveLast
         End If
         FormRefresh
      Case 3 '最後筆
         adoR.MoveLast
         FormRefresh
      Case 4
         adoR.Find "SW01='" & Text1 & "'", , , 1
         If cboClass = "1" Then
            If adoR.EOF Then
               MsgBox "查無資料！"
            Else
               FormRefresh
            End If
         Else
            If adoR.EOF Then
               adoR.MoveFirst
               adoR.Find "SW02='" & Text1 & "'", , , 1
            End If
            If adoR.EOF Then
               MsgBox "查無資料！"
            Else
               FormRefresh
            End If
         End If
         
   End Select
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub TxtLock(ByVal pMode As Integer)
     
   Select Case pMode
   Case 0, 1 '新增,修改
      CmdSitu False
      
   Case 2 '查詢
      CmdSitu False
      
   Case 3 '瀏覽
      CmdSitu True
   End Select
      
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
   Dim ii As Integer, txt As TextBox
   Dim oButton As Button
 
   For ii = 1 To 4
      TBar1.Buttons(ii).Enabled = False
      TBar1.Buttons(ii + 5).Enabled = False
   Next
   TBar1.Buttons(11).Enabled = False
   TBar1.Buttons(12).Enabled = False
      
   If TF = True Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      End If
      If Text1 <> "" Then
         If m_bUpdate Then
          TBar1.Buttons(2).Enabled = True
         End If
         If m_bDelete Then
             TBar1.Buttons(3).Enabled = True
         End If
         
         For ii = 1 To 4
            TBar1.Buttons(ii + 5).Enabled = True
         Next
         TBar1.Buttons(4).Enabled = True
      End If
   Else
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
   End If
   TBar1.Buttons(14).Enabled = True
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.Text)
   OpenIme
End Sub

Private Sub Text2_GotFocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.Text)
   OpenIme
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   lblCode(1) = ""
   If Text2 = "　" Then
      lblCode(1) = "全型空白"
      Text2_GotFocus
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim stMsg As String
   
   If Text1 = MsgText(601) Then
      MsgBox Label1(0) & "不可空白！"
      Exit Function
   End If
   If Text2 = MsgText(601) Then
      MsgBox Label1(1) & "不可空白！"
      Exit Function
   End If
   If cboClass = MsgText(601) Then
      MsgBox Label1(2) & "不可空白！"
      Exit Function
   End If
   '符號
   If cboClass = "1" Then
      If Text2 <> "　" Then
         MsgBox Label1(0) & "只能輸全型空白！"
         Exit Function
      End If
   End If
   
   stMsg = ChkData
   If stMsg = "資料已存在！" Then
      MsgBox stMsg & vbCrLf & _
                     "不可重覆新增"
      Exit Function
   ElseIf stMsg <> MsgText(601) Then
      If MsgBox("資料庫已有下列資料" & vbCrLf & _
                        stMsg & vbCrLf & _
                        "不存檔！回前畫面請按[是]", vbYesNo + vbDefaultButton2) = vbYes Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Function ChkData() As String
   Dim RsQ As New ADODB.Recordset, intQ As Integer
   Dim stTxt As String
   
   If cboClass = "1" Then
      strQ = "Select SW01,SW02,1 Sort From SpecWord Where SW01='" & Text1 & "' "
   Else
      strQ = "Select Distinct SW01,SW02,Sort From (" & _
                    "Select SW01,SW02,1 Sort  From SpecWord Where SW01='" & Text1 & "' And SW02='" & Text2 & "' " & _
      " Union Select SW01,SW02,2 Sort  From SpecWord Where (SW01='" & Text1 & "' Or SW01='" & Text2 & "' ) " & _
      " Union Select SW01,SW02,2 Sort  From SpecWord Where (SW02='" & Text1 & "' Or SW02='" & Text2 & "' ) " & _
       ")"
   End If
   strQ = strQ & " Order by Sort "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      If RsQ.Fields("Sort") = "1" Then
         ChkData = "資料已存在！"
      Else
         Do While Not RsQ.EOF
            stTxt = stTxt & "," & RsQ.Fields("SW01") & "-->" & RsQ.Fields("SW02") & "(統一字)" & vbCrLf
            
            RsQ.MoveNext
         Loop
         If stTxt <> MsgText(601) Then
            ChkData = ChkData & Mid(stTxt, 2)
         End If
      End If
   End If
   Set RsQ = Nothing
End Function

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   If adoR("SW01") = Text1 Then
      adoR.Delete
      adoR.UpdateBatch
      FormDelete = True
   Else
      MsgBox "資料已異動請重新查詢！"
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function FormSave() As Boolean
   
On Error GoTo ErrHnd
   If ActionEdit = 0 Then
      adoR.AddNew
      'Added by Lydia 2024/10/15 KEY值從SW01+SW02改為SW04
      strSql = "select (max(sw04)+1) as MNO from specword "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         adoR.Fields("SW04") = "" & RsTemp.Fields("MNO")
      End If
      'end 2024/10/15
   End If
   adoR.Fields("SW01") = Text1
   adoR.Fields("SW02") = Text2
   adoR.Fields("SW03") = Left(cboClass, 1)
   adoR.UpdateBatch
   
   FormSave = True
   Exit Function
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function
