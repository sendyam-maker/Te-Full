VERSION 5.00
Begin VB.Form frm100119_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "條碼廠商號碼查詢"
   ClientHeight    =   1605
   ClientLeft      =   15
   ClientTop       =   3330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4860
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   900
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "廠商號碼："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   576
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3015
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3810
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   60
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   840
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1170
      Width           =   492
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2235
      TabIndex        =   16
      Top             =   864
      Width           =   2580
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   2
         Left            =   0
         MaxLength       =   6
         TabIndex        =   4
         Top             =   0
         Width           =   1212
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   3
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   5
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   4
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   6
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2220
      TabIndex        =   11
      Top             =   864
      Width           =   2580
      Begin VB.TextBox txtTFCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   972
      End
      Begin VB.TextBox txtTFCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   3
         Left            =   2040
         TabIndex        =   12
         Top             =   0
         Width           =   492
      End
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   3
      Top             =   864
      Width           =   732
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   1
      Top             =   528
      Width           =   972
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "（1.案件基本資料  2.案件進度）"
      Height          =   180
      Left            =   1380
      TabIndex        =   17
      Top             =   1230
      Width           =   2520
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   1230
      Width           =   1245
   End
End
Attribute VB_Name = "frm100119_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/07 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
Dim strSql As String, StrTest As String, strSQL9 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
   Case 0
     cmdState = -1
     If Option1(0).Value = True Then
        If Len(Trim(txt1(0))) = 0 Then
            s = MsgBox("廠商號碼不可空白", , "USER 輸入錯誤")
            txt1(0).SetFocus
            Exit Sub
        End If
        If Len(Trim(txt1(5))) = 0 Then
            s = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
            txt1(5).SetFocus
            Exit Sub
        End If
        strSql = "SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04 FROM SERVICEPRACTICE WHERE SP19='" & txt1(0) & "' "
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If Not IsNull(adoRecordset.Fields(0)) Then
                StrTest = adoRecordset.Fields(0)
            Else
                StrTest = ""
            End If
        Else
            s = MsgBox("沒有此廠商編號", , "搜尋不到")
            Exit Sub
        End If
        CheckOC
        If txt1(5) = "1" Then
            Me.Enabled = False
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
            pub_QL05 = pub_QL05 & ";" & Label3 & "1.案件基本資料" 'Add By Sindy 2010/11/16
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(0) 'Add By Sindy 2010/11/16
            frm100101_7.Show
            frm100101_7.Tag = StrTest
            frm100101_7.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
        Else
            If txt1(5) = "2" Then
                Me.Enabled = False
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
                pub_QL05 = pub_QL05 & ";" & Label3 & "2.案件進度" 'Add By Sindy 2010/11/16
                pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(0) 'Add By Sindy 2010/11/16
                frm100101_2.Show
                frm100101_2.Tag = StrTest
                frm100101_2.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
            End If
        End If
     Else
        If Option1(1).Value = True Then
            If Len(Trim(txt1(2))) = 0 Then
                s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
                txt1(2).SetFocus
                Exit Sub
            End If
            If Len(Trim(txt1(5))) = 0 Then
                s = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
                txt1(5).SetFocus
                Exit Sub
            End If
            strSQL9 = ""
            strSQL9 = txt1(1) & "-" & txt1(2) & "-"
            If Len(txt1(3)) = 0 Then
                strSQL9 = strSQL9 & "0-"
            Else
               strSQL9 = strSQL9 & txt1(3) & "-"
            End If
            If Len(txt1(4)) = 0 Then
                strSQL9 = strSQL9 & "00"
            Else
               strSQL9 = strSQL9 & txt1(4)
            End If
            If txt1(5) = "1" Then
                Me.Enabled = False
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
                pub_QL05 = pub_QL05 & ";" & Label3 & "1.案件基本資料" 'Add By Sindy 2010/11/16
                frm100101_7.Show
                frm100101_7.Tag = strSQL9
                frm100101_7.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
            Else
                If txt1(5) = "2" Then
                    Me.Enabled = False
                    If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                    End If
                    Screen.MousePointer = vbHourglass
                    ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
                    pub_QL05 = pub_QL05 & ";" & Label3 & "2.案件進度" 'Add By Sindy 2010/11/16
                    frm100101_2.Show
                    frm100101_2.Tag = strSQL9
                    frm100101_2.StrMenu
                    Screen.MousePointer = vbDefault
                    Me.Enabled = True
                End If
            End If
        Else
            Exit Sub
        End If
    End If
   Case 1
        fnCloseAllFrm100
   Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0
'     If Option1(0).Value = True Then
'        If Len(Trim(txt1(0))) = 0 Then
'            s = MsgBox("廠商號碼不可空白", , "USER 輸入錯誤")
'            txt1(0).SetFocus
'            Exit Sub
'        End If
'        If Len(Trim(txt1(5))) = 0 Then
'            s = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
'            txt1(5).SetFocus
'            Exit Sub
'        End If
'        strSql = "SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04 FROM SERVICEPRACTICE WHERE SP19='" & txt1(0) & "' "
'        CheckOC
'        adoRecordset.CursorLocation = adUseClient
'        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If Not IsNull(adoRecordset.Fields(0)) Then
'                StrTest = adoRecordset.Fields(0)
'            Else
'                StrTest = ""
'            End If
'        Else
'            s = MsgBox("沒有此廠商編號", , "搜尋不到")
'            Exit Sub
'        End If
'        CheckOC
'        If txt1(5) = "1" Then
'            Me.Enabled = False
'            Screen.MousePointer = vbHourglass
'            frm100101_7.Show
'            'frm100101_7.Hide
'
'            frm100101_7.Tag = StrTest
'            frm100101_7.StrMenu
'            Screen.MousePointer = vbDefault
'            Me.Hide
'            'frm100101_7.Show
'            Do
'            DoEvents
'            If bolToEndByNick = True Then Unload Me: Exit Sub
'            Loop Until Not frm100101_7.Visible
'            Unload frm100101_7
'            Me.Enabled = True
'            Me.Show
'        Else
'            If txt1(5) = "2" Then
'                Me.Enabled = False
'                Screen.MousePointer = vbHourglass
'                frm100101_2.Show
'                'frm100101_2.Hide
'
'                frm100101_2.Tag = StrTest
'                frm100101_2.StrMenu
'                Screen.MousePointer = vbDefault
'                Me.Hide
'                'frm100101_2.Show
'                Do
'                DoEvents
'                If bolToEndByNick = True Then Unload Me: Exit Sub
'                Loop Until Not frm100101_2.Visible
'                Unload frm100101_2
'                Me.Enabled = True
'                Me.Show
'            End If
'        End If
'     Else
'        If Option1(1).Value = True Then
'            If Len(Trim(txt1(2))) = 0 Then
'                s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
'                txt1(2).SetFocus
'                Exit Sub
'            End If
'            If Len(Trim(txt1(5))) = 0 Then
'                s = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
'                txt1(5).SetFocus
'                Exit Sub
'            End If
'            strSQL9 = ""
'            strSQL9 = txt1(1) & "-" & txt1(2) & "-"
'            If Len(txt1(3)) = 0 Then
'                strSQL9 = strSQL9 & "0-"
'            Else
'               strSQL9 = strSQL9 & txt1(3) & "-"
'            End If
'            If Len(txt1(4)) = 0 Then
'                strSQL9 = strSQL9 & "00"
'            Else
'               strSQL9 = strSQL9 & txt1(4)
'            End If
'            If txt1(5) = "1" Then
'                Me.Enabled = False
'                Screen.MousePointer = vbHourglass
'                frm100101_7.Show
'                'frm100101_7.Hide
'
'                frm100101_7.Tag = strSQL9
'                frm100101_7.StrMenu
'                Screen.MousePointer = vbDefault
'                Me.Hide
'                'frm100101_7.Show
'                Do
'                DoEvents
'                If bolToEndByNick = True Then Unload Me: Exit Sub
'                Loop Until Not frm100101_7.Visible
'                Unload frm100101_7
'                Me.Enabled = True
'                Me.Show
'            Else
'                If txt1(5) = "2" Then
'                    Me.Enabled = False
'                    Screen.MousePointer = vbHourglass
'                    frm100101_2.Show
'                    'frm100101_2.Hide
'
'                    frm100101_2.Tag = strSQL9
'                    frm100101_2.StrMenu
'                    Screen.MousePointer = vbDefault
'                    Me.Hide
'                    'frm100101_2.Show
'                    Do
'                    DoEvents
'                    If bolToEndByNick = True Then Unload Me: Exit Sub
'                    Loop Until Not frm100101_2.Visible
'                    Unload frm100101_2
'                    Me.Enabled = True
'                    Me.Show
'                End If
'            End If
'        Else
'            Exit Sub
'        End If
'    End If
'Case 1
'     Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
      MoveFormToCenter Me
   Option1(1).Value = False
   txt1(1) = "TB"
   txt1(1).Enabled = False
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100119_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
               txt1(0).SetFocus
               txt1_GotFocus (0)
           End If
      Case 1
           If Option1(1).Value = True Then
               txt1(2).SetFocus
               txt1_GotFocus (2)
           End If
      Case Else
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 1
           
      Case 5
           If InStr(1, "12", txt1(5)) = 0 Then
              s = MsgBox("查詢別只能 1 或 2 !!", , "USER 輸入錯誤")
              txt1(5).SetFocus
              txt1(5).SelStart = 0
              txt1(5).SelLength = Len(txt1(5))
              Exit Sub
           End If
      Case Else
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
      Case 0
          Option1(0).Value = True
      Case 1, 2, 3, 4
          Option1(1).Value = True
      Case Else
   End Select
End Sub
