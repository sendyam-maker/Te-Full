VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100105_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以收/發文量查詢"
   ClientHeight    =   5736
   ClientLeft      =   288
   ClientTop       =   1764
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   37
      Left            =   6258
      MaxLength       =   1
      TabIndex        =   25
      Top             =   3660
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   34
      Left            =   5985
      MaxLength       =   5
      TabIndex        =   28
      Top             =   3930
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   33
      Left            =   2150
      MaxLength       =   1
      TabIndex        =   35
      Top             =   5376
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   32
      Left            =   1400
      MaxLength       =   1
      TabIndex        =   29
      Top             =   4200
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   31
      Left            =   6600
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1230
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   30
      Left            =   3020
      MaxLength       =   9
      TabIndex        =   33
      Top             =   4740
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   29
      Left            =   3020
      MaxLength       =   9
      TabIndex        =   31
      Top             =   4470
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   28
      Left            =   1400
      MaxLength       =   1
      TabIndex        =   20
      Top             =   3120
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   27
      Left            =   3020
      MaxLength       =   4
      TabIndex        =   22
      Top             =   3390
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   26
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   21
      Top             =   3390
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   24
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   18
      Top             =   2850
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   25
      Left            =   3020
      MaxLength       =   4
      TabIndex        =   19
      Top             =   2850
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   23
      Left            =   2150
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1230
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   21
      Left            =   1400
      MaxLength       =   1
      TabIndex        =   26
      Top             =   3930
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   22
      Left            =   2235
      MaxLength       =   1
      TabIndex        =   27
      Top             =   3930
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   20
      Left            =   6600
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1500
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   19
      Left            =   2150
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1500
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   18
      Left            =   1400
      MaxLength       =   6
      TabIndex        =   13
      Top             =   2040
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   3020
      MaxLength       =   7
      TabIndex        =   2
      Top             =   420
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2150
      MaxLength       =   1
      TabIndex        =   4
      Top             =   960
      Width           =   492
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6750
      Style           =   1  '圖片外觀
      TabIndex        =   36
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7530
      Style           =   1  '圖片外觀
      TabIndex        =   37
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1400
      TabIndex        =   3
      Top             =   690
      Width           =   3975
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   23
      Top             =   3660
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   6600
      MaxLength       =   1
      TabIndex        =   5
      Top             =   960
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   14
      Top             =   2310
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1400
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1770
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1400
      MaxLength       =   1
      TabIndex        =   0
      Top             =   150
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   5865
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1770
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   1400
      MaxLength       =   1
      TabIndex        =   34
      Top             =   5040
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   3020
      MaxLength       =   4
      TabIndex        =   24
      Top             =   3660
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   3020
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2310
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   3020
      MaxLength       =   3
      TabIndex        =   11
      Top             =   1770
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   3020
      MaxLength       =   4
      TabIndex        =   17
      Top             =   2580
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1400
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2580
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1400
      MaxLength       =   7
      TabIndex        =   1
      Top             =   420
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   15
      Left            =   1400
      MaxLength       =   9
      TabIndex        =   30
      Top             =   4470
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   16
      Left            =   1400
      MaxLength       =   9
      TabIndex        =   32
      Top             =   4740
      Width           =   1245
   End
   Begin VB.Label lblKind 
      Height          =   252
      Left            =   1944
      TabIndex        =   73
      Top             =   5064
      Width           =   1236
   End
   Begin VB.Label Lbl_InfKind 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '單線固定
      Caption         =   "遊標移至此看選項"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3264
      TabIndex        =   72
      Top             =   5064
      Width           =   2268
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   3
      Left            =   2670
      TabIndex        =   52
      Top             =   2100
      Width           =   1605
      Size            =   "2831;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   7140
      TabIndex        =   50
      Top             =   1830
      Width           =   1605
      Size            =   "2831;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "收文量不含已取消收文案件"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4440
      TabIndex        =   71
      Top             =   195
      Width           =   2160
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否只統計新申請案：             Y: 僅查詢新申請案(含改請) "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   4400
      TabIndex        =   70
      Top             =   3705
      Width           =   4560
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "國際分類："
      Height          =   180
      Left            =   4920
      TabIndex        =   69
      Top             =   3990
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "(國際分類只查詢台灣已審定之發明新型專利案件)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   4860
      TabIndex        =   68
      Top             =   4260
      Width           =   3900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "只統計電子送件資料：             （Ｙ：是 ）"
      Height          =   180
      Index           =   6
      Left            =   300
      TabIndex        =   67
      Top             =   5412
      Width           =   3336
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "案件目前准駁：          （1.准 2.駁 3.未准(含駁) 4. 無准駁）"
      Height          =   180
      Left            =   160
      TabIndex        =   66
      Top             =   4260
      Width           =   4530
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   7140
      TabIndex        =   65
      Top             =   1290
      Width           =   1440
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "FCP工程師組別："
      Height          =   180
      Left            =   5160
      TabIndex        =   64
      Top             =   1290
      Width           =   1380
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "輸入智權人員條件時只統計接洽記錄單資料"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4800
      TabIndex        =   63
      Top             =   2100
      Width           =   3420
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   2775
      X2              =   2895
      Y1              =   4840
      Y2              =   4840
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   2775
      X2              =   2895
      Y1              =   4570
      Y2              =   4570
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人性質：            (A.律師事務所 B.公司直接委辦 C.其他)"
      Height          =   180
      Left            =   120
      TabIndex        =   62
      Top             =   3180
      Width           =   4875
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "本所委任之國外代理人"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4400
      TabIndex        =   61
      Top             =   3450
      Width           =   1800
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "委任本所之國外代理人"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4400
      TabIndex        =   60
      Top             =   2910
      Width           =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CF代理人國籍："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   59
      Top             =   3450
      Width           =   1290
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   2780
      X2              =   2900
      Y1              =   3490
      Y2              =   3490
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   2780
      X2              =   2900
      Y1              =   2950
      Y2              =   2950
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人國籍："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   58
      Top             =   2910
      Width           =   1290
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PCT進入國家階段         ：             （Ｙ：國家階段）"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   57
      Top             =   1290
      Width           =   4005
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "專利/商標種類："
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   3990
      Width           =   1305
   End
   Begin VB.Line Line7 
      X1              =   1995
      X2              =   2115
      Y1              =   4035
      Y2              =   4035
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "是否含來函資料：            （Ｙ：含來函）"
      Height          =   180
      Left            =   5100
      TabIndex        =   55
      Top             =   1560
      Width           =   3240
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否含內部收文資料    ：             （Ｙ：含內部收文）"
      Height          =   180
      Left            =   120
      TabIndex        =   54
      Top             =   1560
      Width           =   4185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Left            =   330
      TabIndex        =   53
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "不計件之案件是否統計：             （Ｙ：統計）"
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   1020
      Width           =   3645
   End
   Begin VB.Line Line5 
      X1              =   2780
      X2              =   2900
      Y1              =   520
      Y2              =   520
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否計算多國案件：             （Ｙ：計算）"
      Height          =   180
      Index           =   0
      Left            =   4920
      TabIndex        =   49
      Top             =   1020
      Width           =   3285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "代  理  人："
      Height          =   180
      Left            =   330
      TabIndex        =   48
      Top             =   4530
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   330
      TabIndex        =   47
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   46
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "業務區別："
      Height          =   180
      Left            =   330
      TabIndex        =   45
      Top             =   1830
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                                                (ALL：全部)"
      Height          =   180
      Left            =   330
      TabIndex        =   44
      Top             =   750
      Width           =   6210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日        期：                                                                       輸入民國年"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   330
      TabIndex        =   43
      Top             =   480
      Width           =   4995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查  詢  別：                （1. 收文  2. 發文）"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   330
      TabIndex        =   42
      Top             =   195
      Width           =   3150
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "統計條件："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   336
      TabIndex        =   41
      Top             =   5100
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申  請  人："
      Height          =   180
      Left            =   330
      TabIndex        =   40
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   4800
      TabIndex        =   39
      Top             =   1830
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   38
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2780
      X2              =   2900
      Y1              =   1870
      Y2              =   1870
   End
   Begin VB.Line Line2 
      X1              =   2780
      X2              =   2900
      Y1              =   2410
      Y2              =   2410
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   2780
      X2              =   2900
      Y1              =   2680
      Y2              =   2680
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   2780
      X2              =   2900
      Y1              =   3760
      Y2              =   3760
   End
End
Attribute VB_Name = "frm100105_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/19 改成Form2.0(lbl1(0),lbl1(3))
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit

Dim strSql As String, i As Integer, j As Integer, s As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add by Morgan2005/4/8
Dim m_bolOk As Boolean

'92.04.16 nick
Public Sub PubShowNextData()
   
Select Case cmdState
Case 0
   'Add by Morgan 2005/4/8
   Dim oText As TextBox
   For Each oText In txt1
      Call txt1_LostFocus(oText.Index)
      If m_bolOk = False Then
         Exit Sub
      End If
   Next
   
    cmdState = -1
    If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
        txt1_GotFocus 1
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
        txt1_GotFocus 2
        Exit Sub
    End If
    
    'Add by Morgan 2011/4/1
    If DBDATE(txt1(1)) < "19221111" Then
      MsgBox "查詢起始日期輸入錯誤！", vbExclamation
      txt1(1).SetFocus
      txt1_GotFocus 1
      Exit Sub
    End If
    
    If txt1(0) = "" Then
        MsgBox "查詢別不可空白 !", vbCritical
        txt1(0).SetFocus
        Exit Sub
    End If
    If txt1(2) = "" Then
        MsgBox "範圍不可空白 !", vbCritical
        txt1(2).SetFocus
        Exit Sub
    End If
    If txt1(17) = "" Then
        MsgBox "統計條件不可空白 !", vbCritical
        txt1(17).SetFocus
        Exit Sub
    End If
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
    frm100105_2.Show
    frm100105_2.StrMenu
    Screen.MousePointer = vbDefault
    Me.Enabled = True
Case 1
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'add by nickc 2007/01/12
If Len(Trim(Me.txt1(3).Text)) = 0 Then
    Me.txt1(3).Text = "ALL"
End If
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   txt1(3) = Systemkind_g
   bolToEndByNick = False
   If bolFNation = False Then
      Label7.Visible = False
      txt1(15).Visible = False
      Line3(3).Visible = False
      txt1(29).Visible = False
   End If
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100105_1 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      Case 18 '承辦人
         Me.lbl1(3).Caption = StaffQuery("" & Me.txt1(Index).Text)
      Case 31
          '2008/10/15 add by toni 增加FCP工程師組別
           lblName = PUB_GetFCPGrpName(txt1(31))
           'end 2008/10/15
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/05/09
   Select Case Index
      'edit by nick 2005/02/04
      'Case 19, 20, 4, 5
      'Modify by Morgan 2011/1/17 +33
      'Add by Lydia 2015/02/12 + 37
      Case 19, 20, 4, 5, 23, 33, 37
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      'Add by Morgan 2007/1/16
      Case 28 'FC代理人性質
         If KeyAscii <> 8 And KeyAscii <> Asc("A") And KeyAscii <> Asc("B") And KeyAscii <> Asc("C") Then
            KeyAscii = 0
         End If
      '2009/4/2 ADD BY SONIA
      Case 32 '准駁條件
         If KeyAscii <> 8 And (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'Add By Cheng 2002/07/08
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSK03 As String
Dim strTemp
   'Add by Morgan 2005/4/8 判斷輸入是否有錯
   m_bolOk = False
   
   Select Case Index
       Case 0
           If InStr(1, "12 ", txt1(0)) = 0 Then
              s = MsgBox("請輸入 1 或 2 !!", , "輸入錯誤")
              txt1(0).SetFocus
              txt1(0).SelStart = 0
              txt1(0).SelLength = Len(txt1(0))
              Exit Sub
           End If
       Case 1, 2
           If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
              Me.txt1(Index).SetFocus
              txt1_GotFocus Index
              Exit Sub
           End If
           If Index = 2 Then
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                   txt1(Index - 1).SetFocus
                   txt1_GotFocus (Index - 1)
                   Exit Sub
               End If
               '92.12.21 add by sonia 分所管理部人員只可查詢一天的統計資料
               StrSQLa = "SELECT ST05 FROM STAFF WHERE ST01 = '" & Trim(strUserNum) & "'"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic
               If rsA.RecordCount > 0 Then
                   If Left(rsA.Fields("ST05").Value, 1) = "C" Or Left(rsA.Fields("ST05").Value, 1) = "N" Or Left(rsA.Fields("ST05").Value, 1) = "K" Then
                       txt1(4) = "Y": txt1(5) = "Y": txt1(17) = "1"
                       Select Case Left(rsA.Fields("ST05").Value, 1)
                           Case "C"
                               txt1(6) = "S20": txt1(7) = "S29"
                           Case "N"
                               txt1(6) = "S30": txt1(7) = "S39"
                           Case "K"
                               txt1(6) = "S40": txt1(7) = "S49"
                       End Select
                       If txt1(1) <> txt1(2) Then
                           s = MsgBox("只可統計一天的資料！", , "錯誤！")
                           txt1(2).SetFocus
                           txt1_GotFocus (2)
                           Exit Sub
                       End If
                   End If
               End If
               '92.12.21 end
           End If
       Case 3 '系統類別
             'Modify By Cheng 2002/03/14
       '      'Add By Cheng 2002/01/07
       '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
            'Added by Lydia 2016/02/24 檢查跨部門權限
            txt1(Index) = Replace(txt1(Index), " ", "")
            If Len(Me.txt1(Index)) > 0 And Me.txt1(Index) <> "ALL" Then
               If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, Me.txt1(Index)) = False Then
                   txt1(Index).SetFocus
                   txt1_GotFocus Index
                   Exit Sub
               End If
            End If
            'end 2016/02/24
       Case 8 '智權人員
            If Len(txt1(8)) <> 0 Then
                 strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & txt1(8) & "'"
                 CheckOC
                 adoRecordset.CursorLocation = adUseClient
                 adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                 If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                     If Not IsNull(adoRecordset.Fields(0)) Then
                          lbl1(0).Caption = adoRecordset.Fields(0)
                     Else
                          lbl1(0).Caption = ""
                     End If
                 Else
                     lbl1(0).Caption = ""
                    s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                    txt1(Index).SetFocus
                    txt1_GotFocus (Index)
                    Exit Sub
                 End If
                 CheckOC
                'Add by Morgan 2005/1/11
                '若有輸入智權人員, 則是否含內部收文資料欄與是否含來函資料欄預設為NULL
                Me.txt1(19).Text = Empty
                Me.txt1(20).Text = Empty
             Else
                lbl1(0).Caption = ""
             End If
       Case 9
   '    Case 7, 10, 12, 14
       Case 7, 10, 12, 14, 22
             If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
             End If
       Case 15 '代理人
   'edit by nickc 2007/04/24 改成區間，不用了
   '          If Len(txt1(15)) <> 0 Then
   '             'Add By Cheng 2002/07/08
   '             strTemp = Split(IIf(Me.txt1(3).Text = "ALL", GetAllSysKind(Me.txt1(3)), Me.txt1(3).Text), ",")
   '             strSK03 = ""
   '             If Me.txt1(3).Text = "" Then
   '                StrSQLa = "Select SK03 From SystemKind Where SK01=''"
   '             Else
   '                StrSQLa = "Select SK03 From SystemKind Where SK01='" & IIf(Me.txt1(3).Text = "", "", strTemp(0)) & "'"
   '             End If
   '             rsA.CursorLocation = adUseClient
   '             rsA.Open StrSQLa, cnnConnection, adOpenStatic
   '             If rsA.RecordCount > 0 Then
   '                strSK03 = "" & rsA.Fields(0).Value
   '             End If
   '             If rsA.State <> adStateClosed Then rsA.Close
   '             Set rsA = Nothing
   '             'Modify By Cheng 2002/07/08
   '             If strSK03 = "0" Then
   '                strSQL = "SELECT NVL(FA04,DECODE(FA05||FA06||FA64||FA65,NULL,FA06,FA05||' '||FA06||' '||FA64||' '||FA65)) FROM FAGENT WHERE FA01='" & Left(GetNewFagent(txt1(15)), 8) & "' AND FA02='" & Right(GetNewFagent(txt1(15)), 1) & "' "
   '             Else
   '                strSQL = "SELECT DECODE(FA05||FA06||FA64||FA65,NULL,NVL(FA04,FA06),FA05||' '||FA06||' '||FA64||' '||FA65) FROM FAGENT WHERE FA01='" & Left(GetNewFagent(txt1(15)), 8) & "' AND FA02='" & Right(GetNewFagent(txt1(15)), 1) & "' "
   '             End If
   '              CheckOC
   '              adoRecordset.CursorLocation = adUseClient
   '              adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   '              If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   '                  If Not IsNull(adoRecordset.Fields(0)) Then
   '                      lbl1(1).Caption = adoRecordset.Fields(0)
   '                  Else
   '                      lbl1(1).Caption = ""
   '                      s = MsgBox("無此代理人 !!", , "輸入錯誤")
   '                      txt1(15).SetFocus
   '                      txt1(15).SelStart = 0
   '                      txt1(15).SelLength = Len(txt1(15))
   '                  End If
   '              Else
   '                  lbl1(1).Caption = ""
   '                  s = MsgBox("無此代理人 !!", , "輸入錯誤")
   '                  txt1(15).SetFocus
   '                  txt1(15).SelStart = 0
   '                  txt1(15).SelLength = Len(txt1(15))
   '              End If
   '              CheckOC
   '          Else
   '             lbl1(1).Caption = ""
   '          End If
       Case 16
   'edit by nickc 2007/04/24 改成區間，不用了
   '          If Len(txt1(16)) <> 0 Then
   '              strSQL = "SELECT NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) FROM CUSTOMER WHERE CU01='" & Left(GetNewFagent(txt1(16)), 8) & "' AND CU02='" & Right(GetNewFagent(txt1(16)), 1) & "' "
   '              CheckOC
   '              adoRecordset.CursorLocation = adUseClient
   '              adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   '              If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   '                  If Not IsNull(adoRecordset.Fields(0)) Then
   '                      lbl1(2).Caption = adoRecordset.Fields(0)
   '                  Else
   '                      lbl1(2).Caption = ""
   '                      s = MsgBox("無此申請人 !!", , "輸入錯誤")
   '                      txt1(16).SetFocus
   '                      txt1(16).SelStart = 0
   '                      txt1(16).SelLength = Len(txt1(16))
   '                  End If
   '              Else
   '                  lbl1(2).Caption = ""
   '                  s = MsgBox("無此申請人 !!", , "輸入錯誤")
   '                  txt1(16).SetFocus
   '                  txt1(16).SelStart = 0
   '                  txt1(16).SelLength = Len(txt1(16))
   '              End If
   '              CheckOC
   '          Else
   '             lbl1(2).Caption = ""
   '          End If
       Case 17
            'edit by nickc 2006/09/01 葉大增加代理人，有請作單
            'If InStr(1, "12345 ", txt1(17)) = 0 Then
                's = MsgBox("請輸入 1 或 2 或 3 或 4 或 5 !!", , "輸入錯誤")
            'edit by nickc 2007/11/27 加入FC代理人國籍
            'If InStr(1, "1234567 ", txt1(17)) = 0 Then
               's = MsgBox("請輸入 1 或 2 或 3 或 4 或 5 或 6  或 7 !!", , "輸入錯誤")
               's = MsgBox("請輸入 1 或 2 或 3 或 4 或 5 或 6  或 7 或 8 或 9 !!", , "輸入錯誤") 2008/10/15 add by toni
            'If InStr(1, "123456789 ", txt1(17)) = 0 Then
            'Modify By Sindy 2014/7/9 +A
            'Modified by Lydia 2017/02/02 +B
            'Modified by Lydia 2025/08/06 +C
            If InStr(1, "123456789ABC ", txt1(17)) = 0 Then
                s = MsgBox("請輸入 1 或 2 或 3 或 4 或 5 或 6  或 7 或 8 或 9 或 A 或 B 或 C !!", , "輸入錯誤")
                txt1(17).SetFocus
                txt1(17).SelStart = 0
                txt1(17).SelLength = Len(txt1(17))
                lblKind.Caption = "" 'Added by Lydia 2025/08/06
                Exit Sub
             'Added by Lydia 2025/08/06 若增加選項，必須一併調整
             Else
                Select Case txt1(17)
                   Case "1": lblKind.Caption = "業務區"
                   Case "2": lblKind.Caption = "智權人員"
                   Case "3": lblKind.Caption = "申請國家"
                   Case "4": lblKind.Caption = "申請人國籍"
                   Case "5": lblKind.Caption = "案件性質"
                   Case "6": lblKind.Caption = "FC代理人"
                   Case "7": lblKind.Caption = "CF代理"
                   Case "8": lblKind.Caption = "FC代理人國籍"
                   Case "9": lblKind.Caption = "FCP工程師組別"
                   Case "A": lblKind.Caption = "專利案件屬性"
                   Case "B": lblKind.Caption = "申請國洲別"
                   Case "C": lblKind.Caption = "承辦人"
                   Case Else: lblKind.Caption = ""
                End Select
             'end 2025/08/06
             End If
       Case 18 '承辦人
          'Add By Cheng 2002/05/09
          '若有輸入承辦人, 則是否含內部收文資料欄與是否含來函資料欄預設為"Y"
          lbl1(3) = GetPrjSalesNM(txt1(Index))
          If Trim(txt1(Index)) <> "" Then
              If Trim(lbl1(3)) = "" Then
                   s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                   txt1(Index).SetFocus
                   txt1_GotFocus (Index)
                   Exit Sub
              End If
          End If
          If Len("" & Me.txt1(Index).Text) > 0 Then
             Me.txt1(19).Text = "Y"
             Me.txt1(20).Text = "Y"
          End If
       'add by nickc 2007/04/24
       Case 29
             If txt1(15) <> "" Or txt1(Index) <> "" Then
                   If RunNick(txt1(15), txt1(Index)) Then
                      txt1(15).SetFocus
                      txt1_GotFocus 15
                      Exit Sub
                   End If
                   If Mid(txt1(15), 1, 6) <> Mid(txt1(Index), 1, 6) Then
                        MsgBox "申請人前六碼必須相同！", , "發生錯誤！"
                        txt1(15).SetFocus
                        txt1_GotFocus 15
                        Exit Sub
                   End If
             End If
       Case 30
             If txt1(16) <> "" Or txt1(Index) <> "" Then
                   If RunNick(txt1(16), txt1(Index)) Then
                      txt1(16).SetFocus
                      txt1_GotFocus 16
                      Exit Sub
                   End If
                   If Mid(txt1(16), 1, 6) <> Mid(txt1(Index), 1, 6) Then
                        MsgBox "代理人前六碼必須相同！", , "發生錯誤！"
                        txt1(16).SetFocus
                        txt1_GotFocus 16
                        Exit Sub
                   End If
             End If
       
       Case Else
   End Select
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   'If Len(Trim(txt1(0))) <> 0 And Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 And Len(Trim(txt1(17))) <> 0 Then
   '    CmdOk(0).Enabled = True
   '    'cmdOK(0).SetFocus
   'Else
   '    CmdOk(0).Enabled = False
   'End If
   'Add by Morgan 2005/4/8
   m_bolOk = True
End Sub

'Added by Lydia 2025/08/06 統計條件太長改成提示按鈕; 若增加選項，必須一併調整
Private Sub Lbl_InfKind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   '2025/08/06 增加C.承辦人
   Lbl_InfKind.ToolTipText = "統計條件：1.業務區  2.智權人員  3.申請國家  4.申請人國籍  5.案件性質  6.FC代理人  7.CF代理  8.FC代理人國籍  9.FCP工程師組別  A.專利案件屬性  B.申請國洲別  C.承辦人"
End Sub

