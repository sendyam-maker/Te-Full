VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050204_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件性質統計"
   ClientHeight    =   3060
   ClientLeft      =   510
   ClientTop       =   2355
   ClientWidth     =   5100
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5100
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   8
      Left            =   1656
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2105
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   9
      Left            =   3270
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2105
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   10
      Left            =   1656
      TabIndex        =   4
      Top             =   1130
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   7
      Left            =   1656
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   6
      Left            =   3264
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1780
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   5
      Left            =   1656
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1780
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   4
      Left            =   1656
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1455
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   3
      Left            =   3264
      MaxLength       =   4
      TabIndex        =   3
      Top             =   805
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   2
      Left            =   1656
      MaxLength       =   4
      TabIndex        =   2
      Top             =   805
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   1
      Left            =   3264
      MaxLength       =   4
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   288
      Index           =   0
      Left            =   1656
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3048
      TabIndex        =   11
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3828
      TabIndex        =   12
      Top             =   10
      Width           =   1200
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Left            =   3000
      TabIndex        =   22
      Top             =   2430
      Width           =   2055
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3625;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      Caption         =   "代理人國籍："
      Height          =   180
      Left            =   330
      TabIndex        =   21
      Top             =   859
      Width           =   1125
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請國家："
      Height          =   180
      Left            =   330
      TabIndex        =   20
      Top             =   534
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   19
      Top             =   2159
      Width           =   1125
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   3030
      X2              =   3150
      Y1              =   2242
      Y2              =   2242
   End
   Begin VB.Label Label5 
      Caption         =   "(一次只統計一種系統類別)"
      Height          =   180
      Left            =   2430
      TabIndex        =   18
      Top             =   1184
      Width           =   2565
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   3030
      X2              =   3150
      Y1              =   1917
      Y2              =   1917
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3024
      X2              =   3144
      Y1              =   942
      Y2              =   942
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3024
      X2              =   3144
      Y1              =   617
      Y2              =   617
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "系統類別："
      Height          =   180
      Left            =   330
      TabIndex        =   17
      Top             =   1184
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "代理人："
      Height          =   180
      Left            =   330
      TabIndex        =   16
      Top             =   2484
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "日期："
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   15
      Top             =   1834
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "(1.收文 2.發文)"
      Height          =   180
      Left            =   2430
      TabIndex        =   14
      Top             =   1509
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "統計別："
      Height          =   180
      Left            =   330
      TabIndex        =   13
      Top             =   1509
      Width           =   1125
   End
End
Attribute VB_Name = "frm050204_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; lbl1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit

Public Bol050102_3 As Boolean
Dim bloKeyPreview As Boolean
Dim IntSysTotal As Integer
Dim ArySysName As Variant
Dim SysNums() As String
Dim strTemp As Variant
Dim strTemp1 As Variant
Dim IntSys As Integer
Dim i As Integer, j As Integer, s As Integer
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   blnClkSure = False
      '申請國家與國籍, 至少須輸入一項
      If Len(Me.txt1(0).Text) <= 0 And Len(Me.txt1(1).Text) <= 0 And Len(Me.txt1(2).Text) <= 0 And Len(Me.txt1(3).Text) <= 0 Then
         MsgBox "申請國家與國籍, 至少須輸入一項!!!", vbExclamation
         Me.txt1(0).SetFocus
         Exit Sub
      End If
      If (Len(Trim(txt1(0))) > 0 And Len(Trim(txt1(1))) <= 0) Or (Len(Trim(txt1(0))) <= 0 And Len(Trim(txt1(1))) > 0) Then
          s = MsgBox("申請國家區間不可空白", , "USER 輸入錯誤")
         txt1(0).SetFocus
         txt1_GotFocus (0)
          Exit Sub
      End If
      If (Len(Trim(txt1(2))) > 0 And Len(Trim(txt1(3))) <= 0) Or (Len(Trim(txt1(2))) <= 0 And Len(Trim(txt1(3))) > 0) Then
         s = MsgBox("代理人國籍區間不可空白", , "USER 輸入錯誤")
         txt1(2).SetFocus
         txt1_GotFocus (2)
         Exit Sub
      End If
      If Len(Trim(txt1(10))) = 0 Then
          s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
          txt1(10).SetFocus
          Exit Sub
      End If
      If Len(Trim(txt1(10))) <> 0 Then
          strTemp = Split(GetSystemKindByNick, ",")
          strTemp1 = txt1(10)
          s = 0
          For j = 0 To UBound(strTemp)
              If Me.txt1(10).Text = strTemp(j) Then
                  s = 1
              End If
          Next j
          If s = 0 Then
              s = MsgBox(strUserNum + " 沒有 " + Me.txt1(10).Text + " 的使用權限 ", , "USER 權限不足!!!")
              txt1(10).SetFocus
              txt1(10).SelStart = 0
              txt1(10).SelLength = Len(txt1(10))
              Exit Sub
          End If
      End If
      If Len(Trim(txt1(4))) = 0 Then
         s = MsgBox("統計別不可空白", , "USER 輸入錯誤")
         txt1(4).SetFocus
         Exit Sub
      End If
      If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
         Me.txt1(5).SetFocus
         txt1_GotFocus 5
         Exit Sub
      End If
      If PUB_CheckKeyInDate(Me.txt1(6)) = -1 Then
         Me.txt1(6).SetFocus
         txt1_GotFocus 6
         Exit Sub
      End If
      If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
         If Val(Me.txt1(5).Text) > Val(Me.txt1(6).Text) Then
            MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txt1(5).SetFocus
            txt1_GotFocus 5
            Exit Sub
         End If
      End If
      If Len(Trim(txt1(6))) = 0 Then
         s = MsgBox("日期區間不可空白", , "USER 輸入錯誤")
         txt1(5).SetFocus
         txt1_GotFocus (5)
         Exit Sub
      End If
      If Me.txt1(8).Text <> "" And Me.txt1(9).Text <> "" Then
         If Me.txt1(8).Text > Me.txt1(9).Text Then
            MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txt1(8).SetFocus
            txt1_GotFocus 8
            Exit Sub
         End If
      End If
      
      Me.Hide
      Me.Enabled = False
      Screen.MousePointer = vbHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
      'Add by Amy 2018/02/27 記錄此層
      If fnSaveParentForm(Me) = False Then
        Exit Sub
      End If
      frm050204_2.Show
      Screen.MousePointer = vbDefault
      Me.Enabled = True
Case 1
    Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    bloKeyPreview = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050204_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 4
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii < 8 Then
         KeyAscii = 0
      End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
    Dim strA As String
Dim strName As String
'Add By Cheng 2002/03/08
'申請國家與國籍,至少須輸入一項
If Index = 3 Then
   If Len(Me.txt1(0).Text) <= 0 And Len(Me.txt1(1).Text) <= 0 And Len(Me.txt1(2).Text) <= 0 And Len(Me.txt1(3).Text) <= 0 Then
      MsgBox "申請國家與國籍, 至少須輸入一項!!!", vbExclamation
      Me.txt1(0).SetFocus
      Exit Sub
   End If
End If

    'If txt1(Index) = "" Then Exit Sub
    Select Case Index
    Case 1
         If RunNick(txt1(0), txt1(1)) Then
            txt1(0).SetFocus
         End If
    Case 3
         If RunNick(txt1(2), txt1(3)) Then
            txt1(2).SetFocus
         End If
        'If Not objPublicData.GetNation(txt1(Index), strA) Then Cancel = True
    Case 4
        If InStr(1, "12 ", txt1(4)) = 0 Then
            s = MsgBox("統計別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
            txt1(4).SetFocus
            txt1(4).SelStart = 0
            txt1(4).SelLength = Len(txt1(4))
            Exit Sub
        End If
    Case 5, 6
      If blnClkSure = False Then
        If Trim(txt1(Index)) <> "" Then
            If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
               Me.txt1(Index).SetFocus
               txt1_GotFocus Index
               Exit Sub
            End If
        End If
         If Index = 6 Then
            If RunNick(txt1(5), txt1(6)) Then
               txt1(5).SetFocus
               txt1_GotFocus (5)
            End If
         End If
      Else
         blnClkSure = False
      End If
    Case 7
        strA = GetNewFagent(txt1(Index))
         'Modify By Cheng 2002/07/08
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'        If Not objPublicData.GetAgent(strA, strName) Then
         strTemp1 = Split(txt1(10) & " ", ",")
         If Trim(txt1(Index)) = "" Then lbl1 = "": Exit Sub
        If Not PUB_GetAgentName(IIf(Len(Trim(txt1(10))) <> 0, strTemp1(0), ""), strA, strName) Then
           lbl1 = "錯誤!!"
           txt1(7).SetFocus
           Exit Sub
        Else
           lbl1 = strName
           txt1(Index) = IIf(Len(strA) = 8, strA & "0", txt1(Index))
        End If
    Case 9
      If blnClkSure = False Then
         If RunNick(txt1(8), txt1(9)) Then
            txt1(8).SetFocus
            txt1_GotFocus (8)
         End If
      Else
         blnClkSure = False
      End If
    Case 10
        If Len(Trim(txt1(10))) <> 0 Then
            strTemp = Split(GetSystemKindByNick, ",")
            strTemp1 = txt1(10)
            s = 0
            For j = 0 To UBound(strTemp)
                If Me.txt1(10).Text = strTemp(j) Then
                    s = 1
                End If
            Next j
            If s = 0 Then
                s = MsgBox(strUserNum + " 沒有 " + Me.txt1(10).Text + " 的使用權限 ", , "USER 權限不足!!!")
                txt1(10).SetFocus
                txt1(10).SelStart = 0
                txt1(10).SelLength = Len(txt1(10))
                Exit Sub
            End If
        End If
    Case Else
    End Select
End Sub
