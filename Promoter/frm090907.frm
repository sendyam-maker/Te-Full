VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090907 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專-專利連結通知維護作業"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7905
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Height          =   375
      Left            =   5940
      TabIndex        =   5
      Top             =   195
      Width           =   855
   End
   Begin VB.TextBox txtPA177 
      Height          =   300
      Left            =   1470
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2010
      Width           =   405
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   6900
      TabIndex        =   6
      Top             =   195
      Width           =   855
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   270
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   3
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Top             =   285
      Width           =   495
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   2
      Left            =   2595
      MaxLength       =   1
      TabIndex        =   2
      Top             =   285
      Width           =   345
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   1
      Left            =   1695
      MaxLength       =   6
      TabIndex        =   1
      Top             =   285
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   0
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "FCP"
      Top             =   285
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   7740
      Y1              =   1860
      Y2              =   1860
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   17
      Top             =   1410
      Width           =   885
      Size            =   "1561;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   3
      Left            =   1980
      TabIndex        =   16
      Top             =   1410
      Width           =   5535
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1500
      X2              =   3090
      Y1              =   450
      Y2              =   450
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   15
      Top             =   675
      Width           =   6615
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   1
      Left            =   1980
      TabIndex        =   14
      Top             =   1050
      Width           =   5535
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   1050
      Width           =   885
      Size            =   "1561;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "申請人1："
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1095
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   690
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "專利連結通知：　　　 (Y：是)"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   330
      Width           =   945
   End
End
Attribute VB_Name = "frm090907"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/30 Form2.0已修改 lblFM2(index)、Combo1
'Create by Lydia 2021/07/30 外專-專利連結通知維護作業
Option Explicit

Dim strTmpQ As String
Dim mpa(1 To 4) As String '本所案號
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim strB960CP09 As String '藥品專利連結告代的收文號

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
        MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
        txtCase(1).SetFocus
        txtCase_GotFocus 1
        Exit Sub
    Else
        If Trim(lblFM2(0).Caption & lblFM2(2).Caption) = "" Then
            MsgBox "請先查詢本所案號！", vbExclamation, "檢核資料"
            CmdQuery.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txtCase(0).Text & txtCase(1).Text & txtCase(2).Text & txtCase(3).Text) <> Trim(mpa(1) & mpa(2) & mpa(3) & mpa(4)) Then
        MsgBox "請先查詢本所案號！", vbExclamation, "檢核資料"
        CmdQuery.SetFocus
        Exit Sub
    End If
    
    If FormSave = True Then
        MsgBox "存檔完成！", vbInformation
        Call ClearForm(True)
    End If
    
End Sub

Private Sub cmdQuery_Click()
    Call doQuery(True)
End Sub

Private Sub doQuery(ByVal bolMsg As Boolean)

    If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
        MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
        txtCase(1).SetFocus
        txtCase_GotFocus 1
        Exit Sub
    End If
    
    Call ClearForm(False)
    If txtCase(2) = "" Then txtCase(2) = "0"
    If txtCase(3) = "" Then txtCase(3) = "00"

    mpa(1) = txtCase(0): mpa(2) = txtCase(1):  mpa(3) = txtCase(2): mpa(4) = txtCase(3)
    
    strTmpQ = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa177,pa26 as custno,nvl(cu04,nvl(cu05,cu06)) custname,pa75 as fano, nvl(fa04,nvl(fa05,fa06)) as faname " & _
                    "from patent,customer,fagent where pa01='" & mpa(1) & "' and pa02='" & mpa(2) & "' and pa03='" & mpa(3) & "' and pa04='" & mpa(4) & "' " & _
                    "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) "
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        Exit Sub
    End If
    
    intQ = 0
    Combo1.AddItem "中：" & rsQuery.Fields("pa05"), 0
    If rsQuery.Fields("pa05") <> "" Then intQ = 1
    Combo1.AddItem "英：" & rsQuery.Fields("pa06"), 1
    If rsQuery.Fields("pa06") <> "" Then intQ = 2
    'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
    Combo1.AddItem "外：" & rsQuery.Fields("pa07"), 2
    If rsQuery.Fields("pa07") <> "" Then intQ = 3
    Combo1.ListIndex = intQ - 1
    
    lblFM2(0).Caption = "" & rsQuery.Fields("custno")
    lblFM2(1).Caption = "" & rsQuery.Fields("custname")
    lblFM2(2).Caption = "" & rsQuery.Fields("fano")
    lblFM2(3).Caption = "" & rsQuery.Fields("faname")
    txtPA177 = "" & rsQuery.Fields("pa177")
    txtPA177.Tag = txtPA177.Text
    
    Call PUB_ChkCPExist(mpa, "960", , strB960CP09, , "B")
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    Call ClearForm(True)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090907 = Nothing
End Sub

Private Sub ClearForm(ByVal bolResetCase As Boolean)
Dim oObj
    
    If bolResetCase = True Then
        For Each oObj In txtCase
            oObj.Text = ""
        Next
        txtCase(0) = "FCP"
    End If
    
    For Each oObj In lblFM2
       oObj.Caption = ""
    Next
    Combo1.Clear
    
    txtPA177.Text = ""
    txtPA177.Tag = ""
    strB960CP09 = ""
    For intQ = 1 To UBound(mpa)
         mpa(intQ) = ""
    Next intQ
End Sub

Private Sub txtCase_GotFocus(Index As Integer)
    TextInverse txtCase(Index)
End Sub

Private Sub txtCase_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCase_LostFocus(Index As Integer)
    If Index > 1 And Trim(txtCase(Index)) = "" Then
        If Index = 2 Then
             txtCase(2) = "0"
        ElseIf Index = 3 Then
             txtCase(3) = "00"
        End If
    End If
End Sub

Private Function FormSave() As Boolean
Dim strDiff As String
    
   If txtPA177.Text <> txtPA177.Tag Then
       strDiff = strDiff & ", PA177=" & CNULL(txtPA177.Text)
   End If
    
On Error GoTo ErrHandle
   
   If strDiff <> "" Then
       '再抓一次專利連結通知收文(B類收文960)
       Call PUB_ChkCPExist(mpa, "960", , strB960CP09, , "B")
       cnnConnection.BeginTrans
            '專利連結通知=Y：進度檔自動新增一專利連結通知收文(B類收文960)，自動上發文日
            If txtPA177.Text = "Y" Then
                 If strB960CP09 = "" Then
                     strB960CP09 = AutoNo("B", 6)
                     strExc(5) = PUB_GetFCPSalesNo(mpa(1), mpa(2), mpa(3), mpa(4))
                     strExc(6) = GetST15(strExc(5))
                     strSql = "Insert into CaseProgress(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32) " & _
                                 "values (" & CNULL(mpa(1)) & ", " & CNULL(mpa(2)) & ", " & CNULL(mpa(3)) & ", " & CNULL(mpa(4)) & ", " & _
                                 strSrvDate(1) & ", " & CNULL(strB960CP09) & ", '960', '" & strExc(6) & "', '" & strExc(5) & "','" & strUserNum & "', 'N', 'N', " & strSrvDate(1) & ", 'N' )"
                     cnnConnection.Execute strSql
                 End If
            Else
                 If strB960CP09 <> "" Then
                     strSql = "Update CaseProgress set cp57=" & strSrvDate(1) & " where cp09 = " & CNULL(strB960CP09)
                     cnnConnection.Execute strSql
                 End If
            End If
            '更新基本檔
            strDiff = Mid(strDiff, 2)
            strSql = "UPDATE PATENT SET " & strDiff & " WHERE PA01='" & txtCase(0) & "' AND PA02='" & txtCase(1) & "' AND PA03='" & txtCase(2) & "' AND PA04='" & txtCase(3) & "' "
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
       cnnConnection.CommitTrans
   End If
   FormSave = True

ErrHandle:
   If Err.Number <> 0 Then
        MsgBox Err.Description
        cnnConnection.RollbackTrans
   End If
   
End Function

Private Sub txtPA177_GotFocus()
    TextInverse txtPA177
End Sub

Private Sub txtPA177_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
End Sub
