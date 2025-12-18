VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100122_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收/發文量比較查詢"
   ClientHeight    =   4980
   ClientLeft      =   285
   ClientTop       =   1755
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5445
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2865
      Width           =   372
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3531
      Width           =   372
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   24
      Left            =   2850
      MaxLength       =   7
      TabIndex        =   7
      Top             =   2199
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   23
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2199
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   22
      Left            =   2850
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1533
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   21
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1533
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2850
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1866
      Width           =   1332
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4530
      Width           =   372
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3690
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4470
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   60
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1200
      TabIndex        =   0
      Top             =   570
      Width           =   2892
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2532
      Width           =   372
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   12
      Top             =   3864
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1200
      Width           =   372
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   14
      Top             =   4197
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   17
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3198
      Width           =   372
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   13
      Top             =   3864
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1866
      Width           =   1332
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1950
      TabIndex        =   32
      Top             =   4220
      Width           =   1515
      Size            =   "2672;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否含多國案：             （Y：是）"
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   2925
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "比較時段3："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   2259
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "( ALL：指 P , T , CFP , CFT , FCP , FCT )"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1290
      TabIndex        =   29
      Top             =   930
      Width           =   3045
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "點件數：                 (1.件數 2.點數)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   3591
      Width           =   2640
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "比較時段2：                                                                      (基礎時段)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   1926
      Width           =   4980
   End
   Begin VB.Line Line7 
      X1              =   2610
      X2              =   2730
      Y1              =   2342
      Y2              =   2342
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "比較時段1："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   1593
      Width           =   990
   End
   Begin VB.Line Line6 
      X1              =   2640
      X2              =   2760
      Y1              =   1676
      Y2              =   1676
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "是否含FC資料：           （Ｙ：是）"
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   4575
      Width           =   2685
   End
   Begin VB.Line Line5 
      X1              =   2610
      X2              =   2730
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "（1. 收文   2. 發文）"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1740
      TabIndex        =   24
      Top             =   1242
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件性質：             (1.新申請案 2.全部)"
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   2592
      Width           =   3000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "業務區別："
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   3924
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   615
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1242
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "對　　象：             (1.各區 2.個人)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   3258
      Width           =   2640
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   4257
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2250
      X2              =   2370
      Y1              =   4007
      Y2              =   4007
   End
End
Attribute VB_Name = "frm100122_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; lbl1(0)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim strSql As String, i As Integer, j As Integer, s As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Public m_strSystemkindByUser As String '使用者可使用的系統類別
Dim m_arrSystemKindByUser

'92.04.16 nick
Public Sub PubShowNextData()
Dim m_arrSK
Dim ii As Integer
Dim jj As Integer
Dim blnOK As Boolean

Select Case cmdState
Case 0 '確定
    cmdState = -1
    '檢查系統類別
    If Me.txt1(3).Text = "" Then
        MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
        Me.txt1(3).SetFocus
        txt1_GotFocus 3
        Exit Sub
    Else
        If m_strSystemkindByUser = "" Then
            MsgBox "您無P,T,CFP,CFT,FCP,FCT的使用權限!!!", vbExclamation + vbOKOnly
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
        Else
            If Me.txt1(3).Text = "ALL" Then
                m_arrSK = Split(m_strSystemkindByUser, ",")
            Else
                m_arrSK = Split(Me.txt1(3).Text, ",")
            End If
            For ii = LBound(m_arrSK) To UBound(m_arrSK)
                blnOK = False
                For jj = LBound(m_arrSystemKindByUser) To UBound(m_arrSystemKindByUser)
                    If m_arrSK(ii) <> "" Then
                        If m_arrSK(ii) = m_arrSystemKindByUser(jj) Then
                            blnOK = True
                            Exit For
                        End If
                    Else
                        blnOK = True
                        Exit For
                    End If
                Next jj
                If blnOK = False Then
                    MsgBox "系統類別<" & m_arrSK(ii) & ">輸入錯誤!!!", vbExclamation + vbOKOnly
                    Me.txt1(3).SetFocus
                    txt1_GotFocus 3
                    Exit Sub
                End If
            Next ii
        End If
    End If
    '檢查查詢別
    If Me.txt1(0).Text = "" Then
        MsgBox "請輸入查詢別!!!", vbExclamation + vbOKOnly
        Me.txt1(0).SetFocus
        txt1_GotFocus 0
        Exit Sub
    End If
    '檢查統計時段
    If Me.txt1(1).Text = "" Then
        MsgBox "請輸入統計時段!!!", vbExclamation + vbOKOnly
        Me.txt1(1).SetFocus
        txt1_GotFocus 1
        Exit Sub
    End If
    If Me.txt1(2).Text = "" Then
        MsgBox "請輸入統計時段!!!", vbExclamation + vbOKOnly
        Me.txt1(2).SetFocus
        txt1_GotFocus 2
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
        Me.txt1(1).SetFocus
        txt1_GotFocus 1
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
        Me.txt1(2).SetFocus
        txt1_GotFocus 2
        Exit Sub
    End If
    If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
        MsgBox "統計時段區間輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txt1(1).SetFocus
        txt1_GotFocus 1
        Exit Sub
    End If
    '檢查比較時段1
    If Me.txt1(21).Text = "" Then
        MsgBox "請輸入比較時段1!!!", vbExclamation + vbOKOnly
        Me.txt1(21).SetFocus
        txt1_GotFocus 21
        Exit Sub
    End If
    If Me.txt1(22).Text = "" Then
        MsgBox "請輸入比較時段1!!!", vbExclamation + vbOKOnly
        Me.txt1(22).SetFocus
        txt1_GotFocus 22
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(21)) = -1 Then
        Me.txt1(21).SetFocus
        txt1_GotFocus 21
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(22)) = -1 Then
        Me.txt1(22).SetFocus
        txt1_GotFocus 22
        Exit Sub
    End If
    If Val(Me.txt1(21).Text) > Val(Me.txt1(22).Text) Then
        MsgBox "比較時段1區間輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txt1(21).SetFocus
        txt1_GotFocus 21
        Exit Sub
    End If
    '檢查比較時段2
    If Me.txt1(23).Text = "" Then
        MsgBox "請輸入比較時段2!!!", vbExclamation + vbOKOnly
        Me.txt1(23).SetFocus
        txt1_GotFocus 23
        Exit Sub
    End If
    If Me.txt1(24).Text = "" Then
        MsgBox "請輸入比較時段2!!!", vbExclamation + vbOKOnly
        Me.txt1(24).SetFocus
        txt1_GotFocus 24
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(23)) = -1 Then
        Me.txt1(23).SetFocus
        txt1_GotFocus 23
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(24)) = -1 Then
        Me.txt1(24).SetFocus
        txt1_GotFocus 24
        Exit Sub
    End If
    If Val(Me.txt1(23).Text) > Val(Me.txt1(24).Text) Then
        MsgBox "比較時段2區間輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txt1(23).SetFocus
        txt1_GotFocus 23
        Exit Sub
    End If
    '檢查案件性質
    If Me.txt1(13).Text = "" Then
        MsgBox "請輸入案件性質!!!", vbExclamation + vbOKOnly
        Me.txt1(13).SetFocus
        txt1_GotFocus 13
        Exit Sub
    End If
    '檢查對象
    If Me.txt1(17).Text = "" Then
        MsgBox "請輸入對象!!!", vbExclamation + vbOKOnly
        Me.txt1(17).SetFocus
        txt1_GotFocus 17
        Exit Sub
    End If
    '檢查點件數
    If Me.txt1(5).Text = "" Then
        MsgBox "請輸入點件數!!!", vbExclamation + vbOKOnly
        Me.txt1(5).SetFocus
        txt1_GotFocus 5
        Exit Sub
    End If
    '檢查業務區
    If Me.txt1(6).Text <> "" And Me.txt1(7).Text <> "" Then
        If Me.txt1(6).Text > Me.txt1(7).Text Then
            MsgBox "業務區別區間錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(6).SetFocus
            txt1_GotFocus 6
            Exit Sub
        End If
    End If
    '檢查智權人員
    lbl1(0).Caption = GetStaffName(Me.txt1(8).Text, False)
    If Me.txt1(8).Text <> "" And Me.lbl1(0).Caption = "" Then
        s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
        Me.txt1(8).SetFocus
        txt1_GotFocus 8
        Exit Sub
    End If
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    frm100122_2.Show
    frm100122_2.StrMenu
    Screen.MousePointer = vbDefault
    Me.Enabled = True
Case 1 '結束
    fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'add by nickc 2007/01/12
If Len(Trim(Me.txt1(3).Text)) = 0 Then
    Me.txt1(3).Text = "ALL"
End If
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
m_strSystemkindByUser = GetUserSystemKind
If m_strSystemkindByUser <> "" Then
    m_arrSystemKindByUser = Split(m_strSystemkindByUser, ",")
End If
txt1(3) = m_strSystemkindByUser
bolToEndByNick = False
cmdState = -1

txt1(9).Text = "Y" 'Added by Lydia 2016/09/06
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100122_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
Select Case Index
Case 0, 5, 13, 17 '查詢別, 點件數, 案件性質, 對象
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
'是否含FC資料
'Modified by Lydia 2016/09/06 +是否含多國案
Case 4, 9
    If KeyAscii <> 89 And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSK03 As String
Dim strTemp

Select Case Index
Case 1, 2, 21, 22, 23, 24 '統計時段, 比較時段1, 比較時段2
    If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
        Me.txt1(Index).SetFocus
        txt1_GotFocus Index
        Exit Sub
    End If
    If Index = 2 Or Index = 22 Or Index = 24 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
    End If
Case 7 '業務區別
    If RunNick(txt1(Index - 1), txt1(Index)) Then
        txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
    End If
Case Else
End Select
End Sub

'取得使用者可使用的系統類別
Private Function GetUserSystemKind() As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetUserSystemKind = ""
StrSQLa = "Select SG02 From Staff_Group, Staff Where ST11=SG01 And ST01='" & strUserNum & "' Group By SG02 Order By 1 "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    While Not rsA.EOF
        Select Case "" & rsA.Fields(0).Value
        Case "P", "T", "CFP", "CFT", "FCP", "FCT"
            GetUserSystemKind = GetUserSystemKind & rsA.Fields(0).Value & ","
        Case Else
        End Select
        rsA.MoveNext
    Wend
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim m_arrSK
Dim ii As Integer
Dim jj As Integer
Dim blnOK As Boolean
    
    Select Case Index
    Case 3 '系統類別
        If Me.txt1(Index).Text <> "" Then
            If m_strSystemkindByUser = "" Then
                Cancel = True
                MsgBox "您無P,T,CFP,CFT,FCP,FCT的使用權限!!!", vbExclamation + vbOKOnly
            Else
                If Me.txt1(Index).Text = "ALL" Then
                    m_arrSK = Split(m_strSystemkindByUser, ",")
                Else
                    m_arrSK = Split(Me.txt1(Index).Text, ",")
                End If
                For ii = LBound(m_arrSK) To UBound(m_arrSK)
                    blnOK = False
                    For jj = LBound(m_arrSystemKindByUser) To UBound(m_arrSystemKindByUser)
                        If m_arrSK(ii) <> "" Then
                            If m_arrSK(ii) = m_arrSystemKindByUser(jj) Then
                                blnOK = True
                                Exit For
                            End If
                        Else
                            blnOK = True
                            Exit For
                        End If
                    Next jj
                    If blnOK = False Then
                        Cancel = True
                        MsgBox "系統類別<" & m_arrSK(ii) & ">輸入錯誤!!!", vbExclamation + vbOKOnly
                        Exit For
                    End If
                Next ii
            End If
        End If
    Case 8 '智權人員
        lbl1(0).Caption = GetStaffName(Me.txt1(Index).Text, False)
        If Me.txt1(Index).Text <> "" And Me.lbl1(0).Caption = "" Then
            Cancel = True
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
        End If
    End Select
    If Cancel = True Then txt1_GotFocus Index
End Sub
