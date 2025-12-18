VERSION 5.00
Begin VB.Form frm04060304_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公開後實審輸入"
   ClientHeight    =   4095
   ClientLeft      =   630
   ClientTop       =   2700
   ClientWidth     =   5760
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5760
   Begin VB.TextBox txtTPG13 
      Height          =   264
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2130
      Width           =   2772
   End
   Begin VB.TextBox txtTPG14 
      Height          =   264
      Left            =   2100
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2490
      Width           =   345
   End
   Begin VB.TextBox txtTPG02New 
      Height          =   264
      Left            =   2100
      MaxLength       =   11
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "下一筆公開號(&N)"
      Height          =   400
      Index           =   1
      Left            =   1710
      TabIndex        =   10
      Top             =   120
      Width           =   1560
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "上一筆公開號(&P)"
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1560
   End
   Begin VB.TextBox text09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2772
   End
   Begin VB.TextBox txtTPG12 
      Height          =   264
      Left            =   3690
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3195
      Width           =   852
   End
   Begin VB.TextBox txtTPG11 
      Height          =   264
      Left            =   2100
      MaxLength       =   2
      TabIndex        =   6
      Top             =   3195
      Width           =   852
   End
   Begin VB.TextBox txtTPG10 
      Height          =   264
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2835
      Width           =   2772
   End
   Begin VB.TextBox txtTPG02 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1395
      Width           =   2772
   End
   Begin VB.TextBox txtTPG01 
      Height          =   264
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1770
      Width           =   2772
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3630
      TabIndex        =   11
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "是否本人申請:                 (Y:是  N: 否 )"
      Height          =   180
      Left            =   690
      TabIndex        =   23
      Top             =   2520
      Width           =   2880
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "新公開編號 :"
      Height          =   180
      Left            =   690
      TabIndex        =   22
      Top             =   1020
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "上下筆移動時會儲存此筆記錄 !!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   21
      Top             =   630
      Width           =   3405
   End
   Begin VB.Label Label10 
      Caption         =   "期"
      Height          =   255
      Left            =   4650
      TabIndex        =   20
      Top             =   3210
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "卷"
      Height          =   255
      Left            =   3090
      TabIndex        =   19
      Top             =   3210
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "本所案號  :"
      Height          =   180
      Left            =   690
      TabIndex        =   18
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "實審申請日 :"
      Height          =   180
      Left            =   690
      TabIndex        =   17
      Top             =   2130
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公報 :"
      Height          =   180
      Left            =   690
      TabIndex        =   16
      Top             =   3210
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "實審公開日 :"
      Height          =   180
      Left            =   690
      TabIndex        =   15
      Top             =   2850
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公開編號 :"
      Height          =   180
      Left            =   690
      TabIndex        =   14
      Top             =   1410
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Left            =   690
      TabIndex        =   13
      Top             =   1770
      Width           =   810
   End
End
Attribute VB_Name = "frm04060304_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Add by Morgan 2005/9/20
Option Explicit
' 設定編輯資料的模式 (新增或修改)
Public m_EditMode As String
Public m_Multi As Boolean '多筆
Public m_Force As Boolean '改公開編號
Const c_Caption As String = "國內公開後實審輸入"
Dim m_CurrTPG10 As String
Dim m_CurrTPG11 As String
Dim m_CurrTPG12 As String

Dim m_TPG13 As String, m_TPG01 As String, m_PaNo As String, m_CP27 As String, m_CP09 As String
Dim m_stTit As String, m_stMsg As String, m_Resp As VbMsgBoxResult
   
Private Sub FormReset()
   txtTPG02New = Empty
   txtTPG01.Tag = Empty
   txtTPG01 = Empty
   txtTPG10 = Empty
   txtTPG11 = Empty
   txtTPG12 = Empty
   txtTPG13 = Empty
   txtTPG14 = Empty
   text09 = Empty
   cmdMove(0).Visible = m_Multi
   cmdMove(1).Visible = m_Multi
   Label11.Visible = m_Multi
End Sub

Private Sub Form_Activate()
   If txtTPG02New.Visible = True Then
      txtTPG02New.SetFocus
   ElseIf txtTPG01.Enabled = True Then
      txtTPG01.SetFocus
   End If
End Sub

Private Sub txtTPG01_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG01.IMEMode = 2
   CloseIme
   TextInverse txtTPG01
End Sub

Private Sub txtTPG01_Validate(Cancel As Boolean)
   If txtTPG01 <> "" Then
      If txtTPG01 <> txtTPG01.Tag Then
         m_stTit = "檢查申請案號"
         m_stMsg = "申請案號輸入錯誤！"
         m_Resp = MsgBox(m_stMsg, vbOKOnly + vbExclamation, m_stTit)
         Cancel = True
      End If
   End If
End Sub

Private Sub txtTPG02New_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG02New.IMEMode = 2
   CloseIme
   TextInverse txtTPG02New
End Sub

Private Sub txtTPG02New_Validate(Cancel As Boolean)

   m_stTit = "資料檢核"
   If txtTPG02New <> "" Then
      If IsDataExist(txtTPG02New, m_TPG13, m_TPG01, m_PaNo, m_CP27, m_CP09) = False Then
         m_stMsg = "該公開編號未輸入公開無提實審資料！"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         Cancel = True
      Else
         txtTPG01.Tag = m_TPG01
         text09 = m_PaNo
         If m_TPG13 <> "" Then
            m_stMsg = "該公開編號已輸入公開後實審資料！"
            m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
            Cancel = True
         End If
      End If
   End If
   
End Sub

Private Sub txtTPG10_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG10.IMEMode = 2
   CloseIme
   TextInverse txtTPG10
End Sub

Private Sub txtTPG10_Validate(Cancel As Boolean)
   Dim m_stMsg As String
   Dim m_stTit As String
   Dim m_Resp
   Cancel = False
   If Trim(txtTPG10) <> Empty Then
      If CheckIsTaiwanDate(txtTPG10, False) = False Then
         Cancel = True
         m_stMsg = "請輸入正確的實審公開日"
         m_stTit = "資料檢核"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         txtTPG10.SetFocus
         txtTPG10_GotFocus
         GoTo EXITSUB
      End If
        '實審公開日不能大於系統日
      If DBDATE(txtTPG10) > strSrvDate(1) Then
         Cancel = True
         m_stMsg = "實審公開日不能大於系統日"
         m_stTit = "資料檢核"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         txtTPG10.SetFocus
         txtTPG10_GotFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub txtTPG11_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG11.IMEMode = 2
   CloseIme
   TextInverse txtTPG11
End Sub

Private Sub txtTPG11_Validate(Cancel As Boolean)
   If txtTPG11 <> "" Then
      If Not ChktxtTPG11 Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtTPG12_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG12.IMEMode = 2
   CloseIme
   TextInverse txtTPG12
End Sub

Private Sub txtTPG12_Validate(Cancel As Boolean)
   If txtTPG12 <> "" Then
      If Not ChktxtTPG12 Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtTPG13_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG13.IMEMode = 2
   CloseIme
   TextInverse txtTPG13
End Sub

Private Sub txtTPG13_Validate(Cancel As Boolean)
   If Trim(txtTPG13) <> Empty Then
      '2011/4/20 MODIFY BY SONIA
      'If CheckIsTaiwanDate(txtTPG13, False) = False Then
      If CheckIsTaiwanDate(txtTPG13, False) = False Or txtTPG13 < 780101 Then
         Cancel = True
         m_stMsg = "請輸入正確的實審申請日"
         m_stTit = "資料檢核"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         txtTPG13.SetFocus
         txtTPG13_GotFocus
         GoTo EXITSUB
      End If
        '實審申請日不能大於系統日
      If DBDATE(txtTPG13) > strSrvDate(1) Then
         Cancel = True
         m_stMsg = "實審申請日不能大於系統日"
         m_stTit = "資料檢核"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         txtTPG13.SetFocus
         txtTPG13_GotFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub txtTPG14_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPG14.IMEMode = 2
   CloseIme
   TextInverse txtTPG14
End Sub

Private Sub txtTPG14_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
      Case "78", "89", "8"
          KeyAscii = KeyAscii
      Case Else
          KeyAscii = 0
    End Select
End Sub

Private Sub UpdateState()
   Select Case m_EditMode
      Case strFind, strDel:
         txtTPG01.Locked = True
         txtTPG10.Locked = True
         txtTPG11.Locked = True
         txtTPG12.Locked = True
         txtTPG13.Locked = True
         txtTPG14.Locked = True
      Case Else:
         txtTPG01.Locked = False
         txtTPG10.Locked = False
         txtTPG11.Locked = False
         txtTPG12.Locked = False
         txtTPG13.Locked = False
         txtTPG14.Locked = False
   End Select
End Sub

Public Sub UpdateData()
   ' 先清除欄位內容
   FormReset
   ReadData txtTPG02
   ' 更新 Caption
   Select Case m_EditMode
      Case strAdd:
         Caption = c_Caption & " -- 新增"
         txtTPG01 = Empty
         txtTPG10 = m_CurrTPG10
         txtTPG11 = m_CurrTPG11
         txtTPG12 = m_CurrTPG12
         txtTPG14 = "Y"
         cmdMove(0).Visible = False
         cmdMove(1).Visible = False
         Label11.Visible = False
      Case strEdit:
         Caption = c_Caption & " -- 修改"

      Case strFind:
         Caption = c_Caption & " -- 查詢"
         
      Case strDel:
         Caption = c_Caption & " -- 刪除"
         
   End Select
   UpdateState
   If m_Force = True Then
      Label12.Visible = True
      txtTPG02New.Visible = True
      txtTPG01 = Empty
   Else
      Label12.Visible = False
      txtTPG02New.Visible = False
   End If
End Sub

' 使用者按下取消的按鍵
Private Sub cmdCancel_Click()
   
   If m_EditMode = strEdit Or m_EditMode = strAdd Then
      m_stTit = "詢問"
      m_stMsg = "你並未存檔, 確定離開嗎?"
      m_Resp = MsgBox(m_stMsg, vbYesNo + vbDefaultButton2 + vbQuestion, m_stTit)
      If m_Resp = vbNo Then
         txtTPG01.SetFocus
         txtTPG01_GotFocus
         Exit Sub
      End If
   End If
   m_Force = False
   Me.Hide
   frm04060304_1.Show
   frm04060304_1.SetInputTPG02 False
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtTPG02.BackColor = &H8000000F
   text09.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm04060304_2 = Nothing
End Sub

Private Function ChktxtTPG11() As Boolean
   Dim strTmp As String
   ChktxtTPG11 = True
   If Len(txtTPG10) = 6 Then
      strTmp = Left(txtTPG10, 2)
   Else
      strTmp = Left(txtTPG10, 3)
   End If
   If Val(txtTPG11) + 91 <> Val(strTmp) Then
      MsgBox "實審公開日期與公報卷數不符，請重新輸入 !", vbCritical
      ChktxtTPG11 = False
   End If
End Function

Private Function ChktxtTPG12() As Boolean
 Dim strTmp As String
 Dim i As Integer, j As Integer
   ChktxtTPG12 = True
   If Len(txtTPG10) = 6 Then
      j = Val(Mid(txtTPG10, 3, 2))
   Else
      j = Val(Mid(txtTPG10, 4, 2))
   End If
   i = (j - 1) * 2
   j = Val(Right(txtTPG10, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   End If
   
   '92年公報從5月開始
   If Val(txtTPG10) < 930000 Then i = i - 8
         
   If Val(txtTPG12) <> i Then
      MsgBox "實審公開日期與公報期數不符，請重新輸入 !", vbCritical
      ChktxtTPG12 = False
   End If

End Function

'使用者按下確定的按鍵
Private Sub cmdOK_Click()
     
   Select Case m_EditMode
      ' 新增或變更
      Case strAdd, strEdit:
         If CheckDataValid() = True Then
            If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            Select Case m_EditMode
               Case strAdd:
                  m_CurrTPG10 = txtTPG10
                  m_CurrTPG11 = txtTPG11
                  m_CurrTPG12 = txtTPG12
            End Select
            
            Me.Hide
            frm04060304_1.Show
            If m_Force = True Then
               frm04060304_1.textQuery = txtTPG02New
            End If
            If m_Multi = True Then
               frm04060304_1.buttonSearch_Click
            Else
               frm04060304_1.SetInputTPG02
            End If
         End If
      ' 刪除
      Case strDel:
         m_stTit = "詢問"
         m_stMsg = "是否要刪除此筆資料"
         m_Resp = MsgBox(m_stMsg, vbYesNo + vbExclamation + vbDefaultButton2, m_stTit)
         If m_Resp = vbYes Then
            If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            Me.Hide
            frm04060304_1.Show
            frm04060304_1.SetInputTPG02
         End If
      
      Case Else:
        Me.Hide
        frm04060304_1.Show
        frm04060304_1.SetInputTPG02 False
   End Select
   m_Force = False
   
EXITSUB:

End Sub

'上下筆
Private Sub cmdMove_Click(Index As Integer)

   Dim i As Integer
   
   If CheckDataValid() = False Then Exit Sub
   
   If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
               
   If Index = 0 Then
      '上一筆
      i = frm04060304_1.grdList.row - 1
      If i > 0 Then
         frm04060304_1.grdList.row = i
         frm04060304_1.grdList_SelChange
      Else
         MsgBox "已是第一筆了 !", vbInformation
      End If
      
   Else
      i = frm04060304_1.grdList.row + 1
      If i < frm04060304_1.grdList.Rows Then
         frm04060304_1.grdList.row = i
         frm04060304_1.grdList_SelChange
      Else
         MsgBox "已是最後一筆了 !", vbInformation
      End If
   End If
   
   Select Case m_EditMode
      Case strEdit
         frm04060304_1.buttonMod_Click
         If Me.txtTPG01.Enabled Then Me.txtTPG01.SetFocus
      Case strFind
         frm04060304_1.buttonQuery_Click
         
   End Select
   
End Sub

' 此模組在處理資料到資料庫的工作
Private Function OnWork() As Boolean
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans

   Select Case Me.m_EditMode
      Case strAdd, strEdit '新增,修改
         If txtTPG02New.Visible = True Then
            
            strSql = "UPDATE TPGAZETTE SET TPG10=" & TransDate(txtTPG10, 2) & _
               ",TPG11='" & Format(txtTPG11, "00") & "',TPG12='" & Format(txtTPG12, "00") & "'" & _
               ",TPG13=" & TransDate(txtTPG13, 2) & ",TPG14='" & txtTPG14 & "'" & _
               " WHERE TPG02='" & txtTPG02New & "'"
            
            cnnConnection.Execute strSql
            
            strSql = "UPDATE TPGAZETTE SET TPG10=NULL" & _
               ",TPG11=NULL,TPG12=NULL,TPG13=NULL,TPG14=NULL" & _
               " WHERE TPG02='" & txtTPG02 & "'"
            
            cnnConnection.Execute strSql
            
         Else
            strSql = "UPDATE TPGAZETTE SET TPG10=" & TransDate(txtTPG10, 2) & _
               ",TPG11='" & Format(txtTPG11, "00") & "',TPG12='" & Format(txtTPG12, "00") & "'" & _
               ",TPG13=" & TransDate(txtTPG13, 2) & ",TPG14='" & txtTPG14 & "'" & _
               " WHERE TPG02='" & txtTPG02 & "'"
            
            cnnConnection.Execute strSql
            
         End If
   
      Case strDel '刪除
         strSql = "UPDATE TPGAZETTE SET TPG10=NULL" & _
            ",TPG11=NULL,TPG12=NULL,TPG13=NULL,TPG14=NULL" & _
            " WHERE TPG02='" & txtTPG02 & "'"
         
         cnnConnection.Execute strSql
         
   End Select
   
   cnnConnection.CommitTrans
   OnWork = True

ErrorHandler:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
    
End Function

Private Function CheckDataValid() As Boolean
   
   Dim bCancel As Boolean
   
   m_stTit = "資料檢核"
   
   CheckDataValid = False
   '若為修改公開編號
   If txtTPG02New.Visible = True Then
      '若未輸入新公開編號
      If Me.txtTPG02New.Text = "" Then
         m_stMsg = "請輸入新公開編號"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         txtTPG02New.SetFocus
         txtTPG02New_GotFocus
         GoTo EXITSUB
      '若有輸入新的公開編號
      Else
         txtTPG02New_Validate bCancel
         If bCancel = True Then
            txtTPG02New.SetFocus
            txtTPG02New_GotFocus
            GoTo EXITSUB
         End If
      End If
   Else
      '讀取相關資料
      Call IsDataExist(txtTPG02, m_TPG13, m_TPG01, m_PaNo, m_CP27, m_CP09)
   End If
   
   If Trim(txtTPG01) = Empty Then
      m_stMsg = "請輸入申請案號"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      txtTPG01.SetFocus
      txtTPG01_GotFocus
      GoTo EXITSUB
   Else
      txtTPG01_Validate bCancel
      If bCancel = True Then
         txtTPG01.SetFocus
         txtTPG01_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   If Trim(txtTPG13) = Empty Then
      m_stMsg = "請輸入實審申請日"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      txtTPG13.SetFocus
      txtTPG13_GotFocus
      GoTo EXITSUB
   Else
      txtTPG13_Validate bCancel
      If bCancel = True Then
         txtTPG13.SetFocus
         txtTPG13_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   '檢查是否本人申請
   If Trim(txtTPG14) = Empty Then
      m_stMsg = "請輸入是否本人申請"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      txtTPG14.SetFocus
      txtTPG14_GotFocus
      GoTo EXITSUB
   End If
   
   If Trim(txtTPG10) = Empty Then
      m_stMsg = "請輸入實審公開日"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      txtTPG10.SetFocus
      txtTPG10_GotFocus
      GoTo EXITSUB
   Else
      txtTPG10_Validate bCancel
      If bCancel = True Then
         txtTPG10.SetFocus
         txtTPG10_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   If Trim(txtTPG11) = Empty Then
      m_stMsg = "請輸入公開卷數"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      txtTPG11.SetFocus
      txtTPG11_GotFocus
      GoTo EXITSUB
   Else
      txtTPG11_Validate bCancel
      If bCancel = True Then
         txtTPG11.SetFocus
         txtTPG11_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   If Trim(txtTPG12) = Empty Then
      m_stMsg = "請輸入公開期數"
      m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
      txtTPG12.SetFocus
      txtTPG12_GotFocus
      GoTo EXITSUB
   Else
      txtTPG12_Validate bCancel
      If bCancel = True Then
         txtTPG12.SetFocus
         txtTPG12_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   '若為本人申請且為本所案件
   If txtTPG14 = "Y" And text09.Text <> "" Then
      If m_CP09 = "" Then
         m_stMsg = "此案件為本所案件但未收文實審，是否要繼續？"
         m_Resp = MsgBox(m_stMsg, vbYesNo + vbDefaultButton2 + vbExclamation, m_stTit)
         If m_Resp = vbNo Then
            txtTPG14.SetFocus
            txtTPG14_GotFocus
            GoTo EXITSUB
         End If
      ElseIf m_CP27 = "" Then
         m_stMsg = "此案件為本所案件但未提實審，請與程序人員確認！"
         m_Resp = MsgBox(m_stMsg, vbOKOnly, m_stTit)
         txtTPG14.SetFocus
         txtTPG14_GotFocus
         GoTo EXITSUB
      ElseIf TransDate(m_CP27, 1) <> Format(txtTPG13) Then
         m_stMsg = "此案件為本所案件但實審申請日與本所發文日[" & TransDate(m_CP27, 1) & "]不符，是否要繼續？"
         m_Resp = MsgBox(m_stMsg, vbYesNo + vbDefaultButton2 + vbQuestion, m_stTit)
         If m_Resp = vbNo Then
            txtTPG14.SetFocus
            txtTPG14_GotFocus
            GoTo EXITSUB
         End If
      End If
    End If
    
   CheckDataValid = True
   
EXITSUB:
   
End Function

' 檢查此筆資料是否存在
Private Function IsDataExist(ByVal p_TPG02 As String, Optional ByRef p_TPG13 As String, Optional ByRef p_TPG01 As String, Optional ByRef p_PaNo As String, Optional ByRef p_CP27 As String, Optional ByRef p_CP09 As String) As Boolean
   
On Error GoTo ErrHnd

   p_TPG13 = ""
   p_TPG01 = ""
   p_PaNo = ""
   p_CP27 = ""
   p_CP09 = ""
   CheckOC3
   With AdoRecordSet3
      '2007/6/6 modify by sonia 加判斷 cp57 因為P-074206
      'strSQL = "SELECT TPG13,TPG01,PA01,PA02,PA03,PA04,CP27,CP09" & _
      '   " FROM TPGAZETTE,PATENT,CASEPROGRESS" & _
      '   " WHERE TPG02 = '" & p_TPG02 & "' AND TPG09='N'" & _
      '   " AND PA11(+)=TPG01 AND PA09(+)='000'" & _
      '   " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP10(+)='416'"
      strSql = "SELECT TPG13,TPG01,PA01,PA02,PA03,PA04,CP27,CP09" & _
         " FROM TPGAZETTE,PATENT,CASEPROGRESS" & _
         " WHERE TPG02 = '" & p_TPG02 & "' AND TPG09='N'" & _
         " AND PA11(+)=TPG01 AND PA09(+)='000'" & _
         " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP10(+)='416' AND CP57 IS NULL"
      '2007/6/6 end
         
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenDynamic
      If .RecordCount > 0 Then
         IsDataExist = True
         p_TPG01 = "" & .Fields("TPG01")
         p_TPG13 = "" & .Fields("TPG13")
         p_CP27 = "" & .Fields("CP27")
         p_CP09 = "" & .Fields("CP09")
         If IsNull(.Fields("PA01")) = False Then
            p_PaNo = .Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04")
         End If
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical

End Function
' 讀取相關資料
Private Sub ReadData(ByVal p_TPG02 As String)
   
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      strSql = "SELECT TPG01,TPG10,TPG11,TPG12,TPG13,TPG14,PA01,PA02,PA03,PA04" & _
         " FROM TPGAZETTE,PATENT" & _
         " WHERE TPG02 = '" & p_TPG02 & "'" & _
         " AND PA11(+)=TPG01 AND PA09(+)='000'"
         
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenDynamic
      If .RecordCount > 0 Then
         txtTPG01 = "" & .Fields("TPG01")
         txtTPG01.Tag = txtTPG01
         txtTPG10 = TransDate("" & .Fields("TPG10"), 1)
         txtTPG11 = "" & .Fields("TPG11")
         txtTPG12 = "" & .Fields("TPG12")
         txtTPG13 = TransDate("" & .Fields("TPG13"), 1)
         txtTPG14 = "" & .Fields("TPG14")
         If IsNull(.Fields("PA01")) = False Then
            text09 = .Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04")
         End If
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical

End Sub
