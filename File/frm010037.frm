VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010037 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件基本資料-打字室"
   ClientHeight    =   5340
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7884
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7884
   Begin VB.CommandButton cmdOpen 
      Caption         =   "瀏覽資料夾"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3240
      TabIndex        =   32
      Top             =   2160
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Caption         =   "中文本資訊"
      Height          =   3200
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   3015
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   34
         Top             =   2805
         Width           =   1800
         Caption         =   "圖式圖數："
         Size            =   "3175;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   173
         Left            =   2040
         TabIndex        =   10
         Top             =   2760
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   33
         Top             =   2445
         Width           =   1800
         ForeColor       =   16711680
         Caption         =   "申請專利範圍項數："
         Size            =   "3175;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   172
         Left            =   2040
         TabIndex        =   9
         Top             =   2400
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   64
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   65
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   66
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   67
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2PA 
         Height          =   330
         Index           =   68
         Left            =   2040
         TabIndex        =   8
         Top             =   1680
         Width           =   705
         VariousPropertyBits=   679495707
         MaxLength       =   4
         Size            =   "1235;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Txt2Tot6 
         Height          =   330
         Left            =   2040
         TabIndex        =   11
         Top             =   2040
         Width           =   705
         VariousPropertyBits=   679495711
         Size            =   "1235;582"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   285
         Width           =   1005
         Caption         =   "摘要頁數："
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   30
         Top             =   645
         Width           =   1245
         Caption         =   "說明書頁數："
         Size            =   "2196;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   1005
         Width           =   795
         Caption         =   "序列表："
         Size            =   "1411;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   27
         Top             =   1005
         Width           =   960
         ForeColor       =   16711680
         Caption         =   "不算超頁費"
         Size            =   "1693;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   26
         Top             =   1365
         Width           =   1800
         Caption         =   "申請專利範圍頁數："
         Size            =   "3175;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   1725
         Width           =   1005
         Caption         =   "圖式頁數："
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LBL2 
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   24
         Top             =   2085
         Width           =   1005
         Caption         =   "頁數總計："
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   5850
      TabIndex        =   12
      Top             =   120
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6840
      TabIndex        =   14
      Top             =   120
      Width           =   930
   End
   Begin MSForms.Label Label3 
      Height          =   300
      Index           =   1
      Left            =   2140
      TabIndex        =   22
      Top             =   1575
      Width           =   5535
      Caption         =   "1111"
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   300
      Index           =   0
      Left            =   2140
      TabIndex        =   21
      Top             =   1215
      Width           =   5535
      Caption         =   "111"
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Txt2PA 
      Height          =   330
      Index           =   75
      Left            =   1080
      TabIndex        =   20
      Top             =   1560
      Width           =   1000
      VariousPropertyBits=   679495711
      Size            =   "1764;582"
      Value           =   "111"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Txt2PA 
      Height          =   330
      Index           =   26
      Left            =   1080
      TabIndex        =   19
      Top             =   1200
      Width           =   1000
      VariousPropertyBits=   679495711
      Size            =   "1764;582"
      Value           =   "111"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   345
      Left            =   1080
      TabIndex        =   18
      Top             =   840
      Width           =   6735
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11880;617"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   495
      VariousPropertyBits=   679495707
      MaxLength       =   2
      Size            =   "873;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   2
      Left            =   2715
      TabIndex        =   1
      Top             =   480
      Width           =   375
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "661;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   1
      Left            =   1725
      TabIndex        =   0
      Top             =   480
      Width           =   975
      VariousPropertyBits=   679495707
      MaxLength       =   6
      Size            =   "1720;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   28
      Top             =   480
      Width           =   615
      VariousPropertyBits=   679495707
      MaxLength       =   3
      Size            =   "1085;582"
      Value           =   "FCP"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL2 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1598
      Width           =   1005
      Caption         =   "代理人："
      Size            =   "1773;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL2 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1238
      Width           =   1005
      Caption         =   "申請人1："
      Size            =   "1773;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL2 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   885
      Width           =   1005
      Caption         =   "案件名稱："
      Size            =   "1773;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL2 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   518
      Width           =   1005
      Caption         =   "本所案號："
      Size            =   "1773;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm010037"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2018/12/28 案件基本資料-打字室
'Memo by Lydia 2018/12/28 使用Form2.0 (Label、Combobox1和TextBox)
Option Explicit
Dim pa() As String '專利基本檔

Dim oThin As Control

Private Sub FormClear(Optional bolAll As Boolean = False)

    If bolAll = True Then
        txtFM2(1) = ""
        txtFM2(2) = ""
        txtFM2(3) = ""
    End If
    
    ComboBox1.Clear
    
    For Each oThin In Label3
         oThin.Caption = ""
    Next
    
    For Each oThin In Txt2PA
        oThin.Text = ""
        oThin.Tag = ""
    Next
    
    Txt2Tot6.Text = Empty
    
    Call TxtLock(True)
End Sub

Private Sub cmdFind_Click()
Dim Cancel As Boolean
Dim intWhere As Integer

    txtFM2_Validate 0, Cancel
    If Cancel = True Then
        Exit Sub
    End If
    If Len(txtFM2(1)) <> 6 Then
        MsgBox "本所案號請輸入6碼!! '", vbCritical
        txtFM2(1).SetFocus
        txtFM2_GotFocus 1
        Exit Sub
    End If
    If Trim(txtFM2(2)) = "" Then txtFM2(2) = "0"
    If Trim(txtFM2(3)) = "" Then txtFM2(3) = "00"
    
    Call FormClear
    
    Select Case txtFM2(0)
       Case "P"
          intWhere = 國內
       Case "CFP"
          intWhere = 國外_CF
       Case "FCP"
          intWhere = 國外_FC
       Case "ALL"
          intWhere = 國外_FC
    End Select
    
    For intI = 0 To 3
        pa(intI + 1) = txtFM2(intI).Text
    Next intI
    
    If Not PUB_ReadPatentDatabase(pa(), intWhere, False) Then
        Exit Sub
    End If
    
    ComboBox1.AddItem "中:" & pa(5)
    ComboBox1.AddItem "英:" & pa(6)
    'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
    ComboBox1.AddItem "外:" & pa(7)
    ComboBox1.ListIndex = 0
        
    For Each oThin In Txt2PA
       oThin.Text = "" & pa(oThin.Index)
       oThin.Tag = oThin.Text
    Next
    Call Txt2PA_Validate(64, False)
    
   '客戶名稱:中->英->日 ; 代理人名稱: 英->中->日
   strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA26,PA75," & _
                     " NVL(CU04,NVL(CU05,CU06)) CNAME,NVL(FA05,NVL(FA04,FA06)) FNAME" & _
                     " FROM PATENT,CUSTOMER,FAGENT" & _
                     " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
                     " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
    intI = 0
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        Label3(0).Caption = "" & RsTemp.Fields("CNAME")
        Label3(1).Caption = "" & RsTemp.Fields("FNAME")
    End If
    
    Call TxtLock(False)
    
    Txt2PA(64).SetFocus
    
    'Added by Lydia 2020/03/03 專利案件和English_Vers檔案：判斷檔案上傳目的地
    If PUB_ChkCPExist(pa, cnt專利案件, , strExc(1), , "D") = True Then
         cmdOpen.Caption = "瀏覽原始檔"
         cmdOpen.Tag = strExc(1)
    Else
         cmdOpen.Caption = "瀏覽資料夾"
         cmdOpen.Tag = ""
    End If
    'end 2020/03/03
End Sub

Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0 '確定
            If pa(1) & pa(2) & pa(3) & pa(4) <> txtFM2(0) & txtFM2(1) & txtFM2(2) & txtFM2(3) Then
                MsgBox "本所案號與資料不一致！", vbCritical, "資料檢核"
                Exit Sub
            End If
            If ChkDiff = True Then
               If FormSave = False Then
                   MsgBox "存檔失敗！", vbCritical
                   Exit Sub
               End If
            End If
            '清空
            Call FormClear(True)
        Case 1 '結束
            If ChkDiff = True Then
                If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
            Unload Me
    End Select
End Sub

Private Sub cmdOpen_Click()
Dim hLocalFile As Long

On Error GoTo ErrHand01
    
    If pa(5) & pa(6) & pa(7) = "" Then
        MsgBox "請先輸入本所案號後，尋找案件資料！", vbCritical
        Exit Sub
    End If
    If pa(1) & pa(2) & pa(3) & pa(4) <> txtFM2(0) & txtFM2(1) & txtFM2(2) & txtFM2(3) Then
        MsgBox "本所案號與資料不一致！", vbCritical
        Exit Sub
    End If
    
    'Added by Lydia 2020/03/03 開啟[原始檔區]
    If InStr(cmdOpen.Caption, "原始檔") > 0 Then
        If PUB_CheckFormExist("frm100101_M") Then
            MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
            Exit Sub
        End If
        If cmdOpen.Tag = "" Then
            MsgBox pa(1) & "-" & pa(2) & "在〔原始檔區〕的專利案件收文號不存在!", vbInformation
        Else
            frm100101_M.m_strKey = cmdOpen.Tag '多筆總收文號
            frm100101_M.SetParent Me
            If frm100101_M.QueryData = True Then
               frm100101_M.Show
               Me.Hide
            End If
        End If
    Else
    'end 2020/03/03
        If pa(1) <> "FCP" Then
            MsgBox "專利案件資料夾目前無" & pa(1) & "的指定路徑", vbCritical
            Exit Sub
        End If
        'Modified by Lydia 2024/07/22 改用變數
        'strExc(1) = "\\Typing2\專利案件\" & Left(Val(pa(2)), 3)
        strExc(1) = "\\" & strTyping2Path & "\專利案件\" & Left(Val(pa(2)), 3)
        
        If Dir(strExc(1), vbDirectory) <> "" Then
             ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
        Else
             MsgBox strExc(1) & " 資料夾不存在 ！", vbInformation
        End If
    End If 'Added by Lydia 2020/03/03
    
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         MsgBox "無法讀取" & strExc(1) & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
End Sub

Private Sub Form_Initialize()
    ReDim pa(0 To TF_PA) As String
End Sub

Private Sub Form_Load()
   
    MoveFormToCenter Me
    Call FormClear(True)
    
    Txt2PA(26).BackColor = &H8000000F
    Txt2PA(75).BackColor = &H8000000F
     Txt2Tot6.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010037 = Nothing
End Sub

Private Function ChkDiff() As Boolean
    ChkDiff = False
    
    For Each oThin In Txt2PA
        If oThin.Text <> oThin.Tag Then
            ChkDiff = True
        End If
    Next
    
End Function

Private Function FormSave() As Boolean
    FormSave = False
    
    strSql = ""
    For Each oThin In Txt2PA
        'Modified by Lydia 2019/01/10 +申請專利範圍項數(最初項數)、圖式圖數
        'If oThin.Index >= 64 And oThin.Index <= 68 Then
        If oThin.Index >= 64 Then
            If oThin.Text <> oThin.Tag Then
                strSql = strSql & ", PA" & Format(oThin.Index) & "=" & CNULL(oThin.Text, True)
            End If
        End If
    Next
    
    If strSql <> "" Then
        cnnConnection.BeginTrans
            strSql = "UPDATE PATENT SET " & Mid(strSql, 2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            strSql = "begin user_data.user_enabled:=0; " & strSql & "; end;"
            Pub_SeekTbLog strSql '新增log
            cnnConnection.Execute strSql
        cnnConnection.CommitTrans
    End If
    FormSave = True
    Exit Function
    
ErrHand:
    If Err.Number <> 0 Then
       cnnConnection.RollbackTrans
       MsgBox Err.Description
    End If
End Function

Private Sub TxtLock(ByVal bEnabled As Boolean)
    'Locked
    For Each oThin In Txt2PA
        If oThin.Index >= 64 And oThin.Index <= 68 Then
            oThin.Locked = bEnabled
        End If
    Next

End Sub

Private Sub Txt2PA_GotFocus(Index As Integer)
    TextInverse Txt2PA(Index)
End Sub

Private Sub Txt2PA_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Txt2PA_Validate(Index As Integer, Cancel As Boolean)
Dim intA As Integer
   '中文本-頁數總計
   If Index >= 64 And Index <= 68 Then
        intA = Val(Txt2PA(64)) + Val(Txt2PA(65)) + Val(Txt2PA(66)) + Val(Txt2PA(67)) + Val(Txt2PA(68))
        Txt2Tot6 = intA
   End If
End Sub

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
  KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtFM2_LostFocus(Index As Integer)
    If Index = 1 Then
        If txtFM2(Index).Text <> "" Then
           If Len(txtFM2(Index)) <> 6 Then
                 MsgBox "本所案號請輸入6碼!! '"
                 txtFM2(Index).SetFocus
                 txtFM2_GotFocus Index
           End If
        End If
    End If
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If txtFM2(Index) <> "FCP" And txtFM2(Index) <> "P" Then
            MsgBox "系統別請輸入FCP或P !! '"
            txtFM2(Index).SetFocus
            txtFM2_GotFocus Index
            Cancel = True
        End If
    End If
End Sub
