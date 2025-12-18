VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090618 
   BorderStyle     =   1  '單線固定
   Caption         =   "季考核"
   ClientHeight    =   2250
   ClientLeft      =   2655
   ClientTop       =   2325
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4350
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   8
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1155
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   7
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1155
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   285
      Index           =   6
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1485
      Width           =   1020
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3015
      TabIndex        =   9
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2235
      TabIndex        =   8
      Top             =   30
      Width           =   756
   End
   Begin VB.TextBox Txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   5
      Left            =   3630
      MaxLength       =   1
      TabIndex        =   15
      Top             =   3465
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   4
      Left            =   1155
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "1"
      Top             =   1860
      Width           =   285
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   3
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   2
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   1
      Top             =   492
      Width           =   405
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   0
      Top             =   492
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   1125
      X2              =   1575
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   24
      Left            =   1860
      TabIndex        =   22
      Top             =   1230
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   21
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   1530
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   2250
      TabIndex        =   19
      Top             =   1500
      Width           =   1575
      VariousPropertyBits=   671107099
      Size            =   "2778;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "(1.全部 2.承辦人 3.繪圖人員)"
      Enabled         =   0   'False
      Height          =   180
      Left            =   4020
      TabIndex        =   18
      Top             =   3510
      Width           =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "(1.螢幕 2.報表)"
      Height          =   180
      Left            =   1530
      TabIndex        =   17
      Top             =   1905
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "對象："
      Enabled         =   0   'False
      Height          =   180
      Index           =   3
      Left            =   2670
      TabIndex        =   16
      Top             =   3510
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   1890
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   1470
      X2              =   2445
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Label Label2 
      Caption         =   "部門代號："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   870
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "年季："
      Height          =   180
      Index           =   0
      Left            =   156
      TabIndex        =   12
      Top             =   540
      Width           =   576
   End
   Begin VB.Label Label1 
      Caption         =   "季"
      Height          =   180
      Index           =   4
      Left            =   2508
      TabIndex        =   11
      Top             =   552
      Width           =   252
   End
   Begin VB.Label Label1 
      Caption         =   "年"
      Height          =   180
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      Top             =   552
      Width           =   240
   End
End
Attribute VB_Name = "frm090618"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/13 改成Form2.0 (lbl1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, k As Integer, s As Integer, TextOk As Boolean, SeekAction As Integer, SeekRec As Variant
Dim StrSQL6 As String, strTemp1 As Variant, SeekTemp As String, DELMenu() As String, DELTemp() As String, SeekBmk1 As Variant, SeekBmk2 As Variant, SeekBmk3 As Variant
Dim strTemp(0 To 7) As String, PLeft(0 To 7) As Integer, Page As Integer, iPrint As Integer, StrTemp95(0 To 1) As String
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
        s = MsgBox("年不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1(0).SelStart = 0
        txt1(0).SelLength = Len(txt1(0))
        Exit Sub
     Else
        If Len(txt1(1)) = 0 Then
            s = MsgBox("季不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1(1).SelStart = 0
            txt1(1).SelLength = Len(txt1(1))
            Exit Sub
        Else
            If Len(txt1(4)) = 0 Then
                s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                txt1(4).SetFocus
                txt1(4).SelStart = 0
                txt1(4).SelLength = Len(txt1(4))
                Exit Sub
            Else
                 If txt1(4) = "1" Then
                     Me.Hide
                 End If
                 Screen.MousePointer = vbHourglass
                 Me.Enabled = False
                 
                 'Added by Morgan 2019/3/20 +108考核(工程師)
                 strExc(1) = Trim(Val(txt1(0)) + 1911) & Format(3 * Val(txt1(1)) - 2, "00")
                 If ProSysState = "1" And strExc(1) >= Left(PUB_108RuleDate, 6) Then
                     frm090618_2.Show
                 Else
                 'end 2019/3/20
                 
                     frm090618_1.Show
                     
                 End If 'Added by Morgan 2019/3/20
                 Me.Enabled = True
                 Screen.MousePointer = vbDefault
            End If
        End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Activate()
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
'add by nickc 2005/04/18 游經理說預設系統日之前一季
   txt1(0) = Val(ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(Mid(ServerDate, 1, 6) & "01")))) \ 10000
   txt1(1).Text = Trim((Val(Format(Val(ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(Mid(ServerDate, 1, 6) & "01")))) Mod 10000, "0000")) \ 100 \ 3) + 1)
   'add by nickc 2008/03/24
   If Val(txt1(1).Text) > 4 Then txt1(1).Text = "4"
   
   txt1(0).Tag = txt1(0): txt1(1).Tag = txt1(1) 'Added by Morgan 2019/3/26 記錄前一季年月
   
If ProState <> "2" Then
'edit by nickc 2005/04/18 游經理說預設系統日之前一季
   txt1(6) = strUserNum
   txt1_Validate 6, False
   txt1(2).Enabled = False
   txt1(3).Enabled = False
   txt1(7).Enabled = False
   txt1(8).Enabled = False
   txt1(4).Enabled = False
   txt1(6).Enabled = False
   
   'Added by Morgan 2019/3/26
   txt1(2) = Pub_StrUserSt03
   txt1(3) = txt1(2)
   SetUserNumEnabled
   'end 2019/3/26
End If
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090618 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 0
   If Trim(txt1(Index)) <> "" Then
     If IsNumeric(txt1(0)) = False Then
        s = MsgBox("年請輸入數字!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1(0).SelStart = 0
        txt1(0).SelLength = Len(txt1(0))
        Cancel = True
        Exit Sub
     End If
     SetUserNumEnabled 'Added by Morgan 2019/3/26
   End If
   
Case 1
   If Trim(txt1(Index)) <> "" Then
      If IsNumeric(txt1(1)) = False Or Val(txt1(1)) < 1 Or Val(txt1(1)) > 4 Then
         s = MsgBox("季請輸入 1-4 的數字!!", , "USER 輸入錯誤")
         txt1(1).SetFocus
         txt1(1).SelStart = 0
         txt1(1).SelLength = Len(txt1(1))
         Cancel = True
         Exit Sub
      End If
      SetUserNumEnabled 'Added by Morgan 2019/3/26
   End If
Case 3
     If RunNick(txt1(Index - 1), txt1(Index)) Then
        txt1(Index).SetFocus
        txt1_GotFocus (Index)
        Cancel = True
        Exit Sub
     End If
Case 4
     If Trim(txt1(Index)) <> "" Then
     If IsNumeric(txt1(Index)) = False Or Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 2 Then
        s = MsgBox("顯示方式請輸入 1 或 2 的數字!!", , "USER 輸入錯誤")
        txt1(Index).SetFocus
        txt1(Index).SelStart = 0
        txt1(Index).SelLength = Len(txt1(Index))
        Cancel = True
        Exit Sub
     End If
     End If
Case 5
    If Trim(txt1(Index)) <> "" Then
     If IsNumeric(txt1(Index)) = False Or Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 3 Then
        s = MsgBox("對象請輸入 1-3 的數字!!", , "USER 輸入錯誤")
        txt1(Index).SetFocus
        txt1(Index).SelStart = 0
        txt1(Index).SelLength = Len(txt1(Index))
        Cancel = True
        Exit Sub
     End If
     End If
Case 6
    lbl1.Caption = "" 'Added by Morgan 2019/3/26
    If Trim(txt1(Index)) <> "" Then
          strSql = "select * from staff where st01='" & txt1(Index) & "' " & IIf(ProSysState = "1", " and st03>='P10' and st03<='P11' ", " and st03='P13' ")
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
               lbl1.Caption = AdoRecordSet3.Fields("st02").Value
          Else
              MsgBox "請輸入有效" & IIf(ProSysState = "1", " 承辦人 ", " 繪圖人員 ") & "員工編號！", , "錯誤！"
              If ProState = "1" Then Exit Sub
               txt1(6).SetFocus
               txt1(6).SelStart = 0
               txt1(6).SelLength = Len(txt1(6))
               Cancel = True
              Exit Sub
          End If
     End If
Case 7
     Select Case Trim(txt1(7))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(7).SetFocus
          txt1(7).SelStart = 0
          txt1(7).SelLength = Len(txt1(7))
          Cancel = True
          Exit Sub
     End Select
Case 8
     Select Case Trim(txt1(8))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          txt1(8).SelStart = 0
          txt1(8).SelLength = Len(txt1(8))
          Cancel = True
          Exit Sub
     End Select
     If RunNick(txt1(7), txt1(8)) Then
         txt1(8).SetFocus
         txt1_GotFocus (8)
         Cancel = True
         Exit Sub
      End If
      
Case Else
End Select
End Sub

'Added by Morgan 2019/3/25 108考核,開放工程師成員(P11)可以查詢所有人員最近一季的季考核成績(在王副總完成評分之後)
Private Sub SetUserNumEnabled()
   Dim stYM As String, stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim bolOpen As Boolean
   
   bolOpen = False
   '完成評分定義:開始評分後一週定義為完成評分時間
   If strSrvDate(1) >= PUB_108RuleDate And Pub_StrUserSt03 = "P11" And ProSysState = "1" And ProState <> "2" Then
      If Val(txt1(0)) = Val(txt1(0).Tag) And Val(txt1(1)) = Val(txt1(1).Tag) Then
         stSQL = "select ea21 from engineerassess where ea02=" & (Val(txt1(0).Tag) + 1911) & " and ea03=" & Val(txt1(1).Tag) & " and ea16>0 and SYSDATE>EA21+7 and rownum<2"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            bolOpen = True
         End If
      End If
      If bolOpen Then
         txt1(6).Enabled = True
      Else
         txt1(6) = strUserNum
         txt1(6).Enabled = False
      End If
   End If
   Set rsQuery = Nothing
End Sub

