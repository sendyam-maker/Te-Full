VERSION 5.00
Begin VB.Form frm03010303_05 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品及服務資料輸入"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3630
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2430
      TabIndex        =   5
      Top             =   30
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1260
      TabIndex        =   4
      Top             =   30
      Width           =   1095
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   3
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   3
      Top             =   510
      Width           =   285
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   2
      Left            =   2445
      MaxLength       =   1
      TabIndex        =   2
      Top             =   510
      Width           =   180
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   1
      Left            =   1665
      MaxLength       =   6
      TabIndex        =   1
      Top             =   510
      Width           =   645
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   0
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Top             =   510
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   540
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2940
      X2              =   1410
      Y1              =   630
      Y2              =   630
   End
End
Attribute VB_Name = "frm03010303_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/10/15 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Create by Nickc 2006/06/13
Option Explicit
Dim tmpTM09 As String
Public ChkTG As Boolean

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
    If Trim(txt2(0)) = "" Then MsgBox "系統別不可空白！", vbCritical: txt2(0).SetFocus: Exit Sub
    Dim Cancel As Boolean
    Cancel = False
    txt2_Validate 0, Cancel
    If Cancel = True Then Exit Sub
    If Trim(txt2(1)) = "" Then MsgBox "流水號不可空白！", vbCritical: txt2(1).SetFocus: Exit Sub
    '2012/11/9 modify by sonia 加TF
    If Trim(txt2(0)) <> "T" And Trim(txt2(0)) <> "FCT" And Trim(txt2(0)) <> "CFT" And Trim(txt2(0)) <> "TF" Then MsgBox "系統別錯誤！", vbCritical: txt2(0).SetFocus: Exit Sub
    If Trim(txt2(2)) = "" Then txt2(2) = "0"
    If Trim(txt2(3)) = "" Then txt2(3) = "00"
    tmpTM09 = ""
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    strSql = "select * from trademark where tm01='" & txt2(0) & "' and tm02='" & txt2(1) & "' and tm03='" & txt2(2) & "' and tm04='" & txt2(3) & "' "
    CheckOC3
    AdoRecordSet3.CursorLocation = adUseClient
    AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If AdoRecordSet3.RecordCount <> 0 Then
        tmpTM09 = CheckStr(AdoRecordSet3.Fields("tm09"))
         '檢查案號
        frm03010303_04.Hide
        Set frm03010303_04.UpForm = Me
        frm03010303_04.TGKey = txt2(0) & "-" & txt2(1) & "-" & txt2(2) & "-" & txt2(3)
        frm03010303_04.AllClass = tmpTM09
'cancel by sonia 2014/4/24 嘉雯要求 T-192087
'        frm03010303_04.cmd2.Visible = False
'        frm03010303_04.txt2(0).Visible = False
'        frm03010303_04.Line1.Visible = False
'        frm03010303_04.txt2(1).Visible = False
'        frm03010303_04.txt2(2).Visible = False
'        frm03010303_04.txt2(3).Visible = False
'2014/4/24 END
        frm03010303_04.Caption = "商品及服務資料"
        frm03010303_04.Label2.Visible = True
        Me.Hide
        frm03010303_04.QueryData
        frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
    Else
        MsgBox "錯誤本所案號，資料不存在！", vbInformation
        Me.Enabled = True
        txt2(0).SetFocus
    End If
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    txt2(1).SetFocus   '2011/5/30 add by sonia
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm03010303_05 = Nothing
End Sub

Private Sub txt2_GotFocus(Index As Integer)
   InverseTextBox txt2(Index)
End Sub

Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0, 2
    KeyAscii = UpperCase(KeyAscii)
Case Else
    Select Case KeyAscii
    Case 48 To 57
    Case 8
    Case Else
            KeyAscii = 0
    End Select
End Select
End Sub

Private Sub txt2_Validate(Index As Integer, Cancel As Boolean)
Dim SeekSystem As Variant
Dim i As Integer
Dim IsSystemOk As Boolean
If Index = 0 Then
    SeekSystem = Split(GetSystemKindByNick, ",")
    IsSystemOk = False
    For i = 0 To UBound(SeekSystem)
        If txt2(0) = SeekSystem(i) Then
            IsSystemOk = True
            Exit For
        End If
    Next i
    If IsSystemOk = False Then MsgBox "此系統別不屬於此部門使用！", vbCritical, "錯誤！": Cancel = True: txt2(0).SelStart = 0: txt2(0).SelLength = Len(txt2(0))
End If
End Sub
