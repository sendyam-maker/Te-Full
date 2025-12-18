VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040158 
   BorderStyle     =   1  '單線固定
   Caption         =   "造字與UnidCode字對照表"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7515
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6660
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1350
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3285
      MaxLength       =   1
      TabIndex        =   0
      Top             =   900
      Width           =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6705
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040158.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
   Begin VB.Label lblCode 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label lblCode 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   930
      Width           =   1215
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   3285
      TabIndex        =   3
      Top             =   1290
      Width           =   600
      VariousPropertyBits=   738213915
      Size            =   "1058;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "造字："
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
      Left            =   2520
      TabIndex        =   2
      Top             =   930
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unicode字："
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
      Left            =   1935
      TabIndex        =   1
      Top             =   1350
      Width           =   1275
   End
End
Attribute VB_Name = "frm12040158"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/27 Form2.0 Lydia說不用改
'Created by Morgan 2013/6/28
Option Explicit

'0:新增 1:修改 2:查詢 3:瀏覽
Dim ActionEdit As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim adoRecord As New ADODB.Recordset
Private Sub Action(Index As Integer)
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub
   Select Case Index
      Case 1 '新增
         FormReset
         Text1.SetFocus
         ActionEdit = 0
      Case 2 '修改
         TextBox1.SetFocus
         TextBox1_GotFocus
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
               RsAction 4
         End Select
         
      Case 12 '按下取消
         Text1 = Text1.Tag
         If Text1 <> "" Then
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
   OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set adoRecord = Nothing
   Set frm12040158 = Nothing
End Sub

Private Sub FormReset()
   Text1.Text = ""
   TextBox1.Text = ""
   lblCode(0) = ""
   lblCode(1) = ""
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.Text)
   OpenIme
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   lblCode(0) = ""
   If Text1 <> "" Then
      lblCode(0) = Hex(Asc(Text1))
   End If
End Sub

Private Sub TextBox1_GotFocus()
   TextBox1.SelStart = 0
   TextBox1.SelLength = Len(TextBox1.Text)
   OpenIme
End Sub

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   If adoRecord("UM01") = Text1 Then
      adoRecord.Delete
      adoRecord.UpdateBatch
      FormDelete = True
   Else
      MsgBox "資料已異動請重新查詢！"
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   Dim stSQL As String
   stSQL = "Select * From unicodemap Order By 1"
   
   adoRecord.CursorLocation = adUseClient
   adoRecord.Open stSQL, adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoRecord.RecordCount > 0 Then
      adoRecord.MoveFirst
      FormRefresh
   End If
End Sub

Private Sub FormRefresh()
   If Not (adoRecord.EOF And adoRecord.BOF) Then
      Text1 = adoRecord("UM01")
      Text1_Validate False
      Text1.Tag = Text1
      'TextBox1 = StrConv("" & adoRecord("UM02"), vbFromUnicode)
      TextBox1 = U2Word("" & adoRecord("UM02"))
      TextBox1_Validate False
      ActionEdit = 3
   End If
End Sub

Private Sub RsAction(ByVal pCmd As Integer)
 
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case pCmd
      Case 0 '第一筆
         adoRecord.MoveFirst
         FormRefresh
      Case 1 '前一筆
         adoRecord.MovePrevious
         If adoRecord.BOF Then
            DataErrorMessage 6
            adoRecord.MoveFirst
         End If
         FormRefresh
      Case 2 '後一筆
         adoRecord.MoveNext
         If adoRecord.EOF Then
            DataErrorMessage 7
            adoRecord.MoveLast
         End If
         FormRefresh
      Case 3 '最後筆
         adoRecord.MoveLast
         FormRefresh
      Case 4
         adoRecord.Find "UM01='" & Text1 & "'", , , 1
         If adoRecord.EOF Then
            MsgBox "查無資料！"
         Else
            FormRefresh
         End If
   End Select
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Function TxtValidate() As Boolean
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim mem() As Byte
   
On Error GoTo ErrHnd
   If ActionEdit = 0 Then
      adoRecord.AddNew
   End If
   adoRecord.Fields("UM01") = Text1
   'adoRecord.Fields("UM02") = StrConv(TextBox1, vbUnicode)
   adoRecord.Fields("UM02") = lblCode(1)
   adoRecord.UpdateBatch
   
   FormSave = True
   Exit Function
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

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

Private Sub TextBox1_Validate(Cancel As Boolean)
   Dim strNew As String, ii As Integer
   lblCode(1) = ""
   If TextBox1 <> "" Then
      Text2.Text = TextBox1.Text
      lblCode(1) = Hex(AscW(TextBox1))
      If Text2.Text = TextBox1.Text Then
         MsgBox "請輸入Unicode字！"
         Cancel = True
      Else
'Modified by Morgan 2014/1/15
'         strNew = StrConv(TextBox1.Text, vbUnicode)
'         '測試發現若轉換後仍為1個字時LOW,HIGH BYTE要對調
'         If Len(strNew) = 1 Then
'            lblCode(1) = Mid(Hex(Asc(strNew)), 3) & Left(Hex(Asc(strNew)), 2)
'         Else
'            For ii = 1 To Len(strNew)
'               lblCode(1) = Hex(Asc(Mid(strNew, ii, 1))) & lblCode(1)
'            Next
'         End If
         lblCode(1) = Hex(AscW(TextBox1))
      End If
   End If
End Sub

Private Function Big2U(myText As String) As String
    'Convert Big 5 to uni-code
    Dim Li_Ind, Ls_Temp, Ls_Temp_Hex
 
    Ls_Temp = ""
    Big2U = ""
    For Li_Ind = 1 To Len(myText)
      If Big2U <> "" Then Big2U = Big2U & ";"
      Ls_Temp = Mid(myText, Li_Ind, 1)
      Ls_Temp_Hex = Hex(AscW(Ls_Temp))
      If Len(Ls_Temp_Hex) = 4 Then
          'Big2U = Big2U & "&#x" & Ls_Temp_Hex & ";"
          If AscW(Ls_Temp) > 0 Then
             Big2U = Big2U & AscW(Ls_Temp)
          Else
             Big2U = Big2U & (65536 + AscW(Ls_Temp))
          End If
      Else
          'Big2U = Big2U & Ls_Temp
          Big2U = Big2U & Format(Asc(Ls_Temp), "0####")
      End If
    Next
End Function

Private Function U2Word(myText As String) As String
    Dim idx As Integer
    Dim arrTmp
    
    U2Word = ""
    arrTmp = Split(myText, ";")
    For idx = LBound(arrTmp) To UBound(arrTmp)
      U2Word = U2Word & ChrW("&H" & arrTmp(idx))
    Next
End Function
