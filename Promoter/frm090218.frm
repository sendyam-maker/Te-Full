VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090218 
   BorderStyle     =   1  '單線固定
   Caption         =   "英文核稿查詢"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4020
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   1
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   960
      Width           =   1065
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   0
      Left            =   990
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   435
      Index           =   1
      Left            =   3030
      TabIndex        =   4
      Top             =   60
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   60
      Width           =   945
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   540
      Width           =   2790
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4921;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   1740
      X2              =   2490
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "完稿日："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文核稿人："
      Height          =   180
      Index           =   25
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frm090218"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; Combo1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim i As Integer, strSql As String, ADORECORDSET66 As New ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        If Val(Txt1(0)) > Val(Txt1(1)) Then
            MsgBox "輸入資料範圍錯誤,請重新輸入", vbInformation
            Txt1(0).SetFocus
            Txt1(0).SelStart = 0
            Txt1(0).SelLength = Len(Txt1(0))
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        Me.Hide
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
        frm090218_1.Show
        frm090218_1.StrMenu
        Me.Enabled = True
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'2010/10/21 MODIFY BY SONIA 改抓專利處英文顧問
'*****改此部門條件要改四個畫面frm090201_2,frm090218,frm090218_1,frm100101_F
'strSql = "select st01||' ==> '||st02 from staff where st04='1' and st03='F62' and st01<>'99998' order by Decode(ST01,'99998','00000',ST01) "
strSql = "select st01||' ==> '||st02 from staff where st04='1' and st03='P14' and st01<>'99998' order by Decode(ST01,'99998','00000',ST01) "
i = 1
Combo1.Clear
Combo1.AddItem "", 0
Set ADORECORDSET66 = New ADODB.Recordset
With ADORECORDSET66
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        .MoveFirst
        Do While .EOF = False
            Combo1.AddItem "" & .Fields(0), i
            i = i + 1
            .MoveNext
        Loop
        Combo1.Text = Combo1.List(0)
    Else
    End If
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090218 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_LostFocus(Index As Integer)
If Index = 2 Then
    If Val(Txt1(1)) > Val(Txt1(2)) Then
        MsgBox "輸入資料範圍錯誤,請重新輸入", vbInformation
        Txt1(1).SetFocus
        Txt1(1).SelStart = 0
        Txt1(1).SelLength = Len(Txt1(1))
        Exit Sub
    End If
End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If Txt1(Index).Text <> "" Then
      If CheckIsTaiwanDate(Txt1(Index).Text) = False Then
          Txt1(Index).SetFocus
          Txt1(Index).SelStart = 0
          Txt1(Index).SelLength = Len(Txt1(Index))
          Cancel = True
          Exit Sub
      Else
          Cancel = False
      End If
   End If
End Sub
