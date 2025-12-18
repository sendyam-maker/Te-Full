VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210139 
   BorderStyle     =   1  '單線固定
   Caption         =   "銷案／銷帳單"
   ClientHeight    =   6500
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6500
   ScaleWidth      =   8950
   Begin VB.Frame Frame5 
      BorderStyle     =   0  '沒有框線
      Height          =   195
      Left            =   870
      TabIndex        =   59
      Top             =   2790
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "改請"
         Height          =   255
         Index           =   4
         Left            =   3210
         TabIndex        =   10
         Top             =   0
         Width           =   705
      End
      Begin VB.CheckBox Check1 
         Caption         =   "退費"
         Height          =   255
         Index           =   3
         Left            =   2415
         TabIndex        =   9
         Top             =   0
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "銷案"
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   8
         Top             =   0
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "銷帳"
         Height          =   255
         Index           =   1
         Left            =   825
         TabIndex        =   7
         Top             =   0
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "轉帳"
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   885
      Left            =   1020
      TabIndex        =   50
      Top             =   4260
      Width           =   7875
      Begin VB.OptionButton Option2 
         Caption         =   "2.本所立新檔"
         Height          =   255
         Index           =   4
         Left            =   6510
         TabIndex        =   25
         Top             =   255
         Width           =   1395
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1.存舊卷內"
         Height          =   255
         Index           =   3
         Left            =   6510
         TabIndex        =   24
         Top             =   30
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "3.轉存相關卷內，卷號："
         Height          =   255
         Index           =   2
         Left            =   540
         TabIndex        =   22
         Top             =   570
         Width           =   2235
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2.資料退寄客戶，址："
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   20
         Top             =   263
         Width           =   2115
      End
      Begin VB.OptionButton Option2 
         Caption         =   "其他："
         Height          =   255
         Index           =   5
         Left            =   4380
         TabIndex        =   26
         Top             =   570
         Width           =   885
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1.資料銷毀"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   19
         Top             =   30
         Width           =   1215
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   2790
         TabIndex        =   23
         Top             =   540
         Width           =   1545
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "2725;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   21
         Top             =   240
         Width           =   3045
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "5371;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   5280
         TabIndex        =   27
         Top             =   540
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "不銷號："
         Height          =   225
         Left            =   5820
         TabIndex        =   57
         Top             =   60
         Width           =   765
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "銷號："
         Height          =   225
         Left            =   0
         TabIndex        =   56
         Top             =   60
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Height          =   615
      Left            =   1410
      TabIndex        =   49
      Top             =   5190
      Width           =   7515
      Begin VB.OptionButton Option3 
         Caption         =   "由財務處寄客戶，址："
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   31
         Top             =   323
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "智權人員親退客戶"
         Height          =   225
         Index           =   0
         Left            =   2400
         TabIndex        =   29
         Top             =   48
         Width           =   1836
      End
      Begin VB.OptionButton Option3 
         Caption         =   "其他："
         Height          =   255
         Index           =   3
         Left            =   4350
         TabIndex        =   33
         Top             =   323
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "銷應收帳款"
         Height          =   225
         Index           =   1
         Left            =   4350
         TabIndex        =   30
         Top             =   60
         Width           =   1305
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   2250
         TabIndex        =   32
         Top             =   300
         Width           =   2055
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "3625;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   480
         TabIndex        =   28
         Top             =   30
         Width           =   1755
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "3096;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   8
         Left            =   5220
         TabIndex        =   34
         Top             =   300
         Width           =   2235
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "3942;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "金額："
         Height          =   225
         Left            =   60
         TabIndex        =   55
         Top             =   60
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   345
      Left            =   1410
      TabIndex        =   48
      Top             =   5880
      Width           =   7515
      Begin VB.OptionButton Option4 
         Caption         =   "退回"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   35
         Top             =   30
         Width           =   915
      End
      Begin VB.OptionButton Option4 
         Caption         =   "其他："
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   36
         Top             =   30
         Width           =   915
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   1980
         TabIndex        =   37
         Top             =   30
         Width           =   5475
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9657;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   6810
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "1"
      Top             =   270
      Width           =   270
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   525
      Left            =   1380
      TabIndex        =   41
      Top             =   3630
      Width           =   7515
      Begin VB.OptionButton Option1 
         Caption         =   "已送件"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   15
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "其他："
         Height          =   255
         Index           =   5
         Left            =   5100
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "已收文未送件之改請(含異議、評定、撤銷或改請專利種類)"
         Height          =   255
         Index           =   4
         Left            =   30
         TabIndex        =   16
         Top             =   240
         Width           =   5055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "已作業未送件"
         Height          =   255
         Index           =   2
         Left            =   2730
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "已收文未作業"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "未收文"
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   915
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   6000
         TabIndex        =   18
         Top             =   210
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   2
      Left            =   7530
      TabIndex        =   39
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)　　份"
      Height          =   345
      Index           =   1
      Left            =   6060
      Style           =   1  '圖片外觀
      TabIndex        =   38
      Top             =   240
      Width           =   1395
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5010
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1515
      Left            =   30
      TabIndex        =   40
      Top             =   1200
      Width           =   8835
      _ExtentX        =   15575
      _ExtentY        =   2663
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   13
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V|收文日|總收文號|案件性質|相關收文號|承辦人|智權人員|本所期限|法定期限|發文日|取消收文|相關人|進度備註"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   3
      Top             =   315
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2490
      MaxLength       =   1
      TabIndex        =   2
      Top             =   315
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   1
      Top             =   315
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   990
      MaxLength       =   3
      TabIndex        =   0
      Top             =   315
      Width           =   525
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   180
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin MSForms.TextBox Text1 
      Height          =   615
      Index           =   0
      Left            =   870
      TabIndex        =   11
      Top             =   3000
      Width           =   7935
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "13996;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      Caption         =   "註：若非改請會寄發E-Mail至智權委辦區。"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   66
      Top             =   6270
      Width           =   6735
   End
   Begin VB.Label Label14 
      Caption         =   "備註：使用此作業時，建議不要同時使用Word軟體，因程式執行中會使用到Word。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   270
      TabIndex        =   65
      Top             =   30
      Width           =   7455
   End
   Begin VB.Label Label13 
      Caption         =   "註：案件性質可複選"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7260
      TabIndex        =   64
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label Label12 
      Caption         =   "註：若為改請會寄發E-Mail給財務處及電腦中心"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5070
      TabIndex        =   63
      Top             =   2790
      Width           =   3735
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   60
      X2              =   8865
      Y1              =   5850
      Y2              =   5850
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   60
      X2              =   8865
      Y1              =   5145
      Y2              =   5130
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   60
      X2              =   8865
      Y1              =   4185
      Y2              =   4170
   End
   Begin VB.Label Label11 
      Caption         =   "案件資料"
      Height          =   195
      Left            =   60
      TabIndex        =   62
      Top             =   4290
      Width           =   945
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   6420
      TabIndex        =   61
      Top             =   900
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   5490
      TabIndex        =   60
      Top             =   900
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "申請國家："
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      Caption         =   "通知單："
      Height          =   225
      Left            =   60
      TabIndex        =   58
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "收據處理方式："
      Height          =   225
      Left            =   60
      TabIndex        =   54
      Top             =   5940
      Width           =   1305
   End
   Begin VB.Label Label5 
      Caption         =   "款項處理方式："
      Height          =   225
      Left            =   60
      TabIndex        =   53
      Top             =   5220
      Width           =   1305
   End
   Begin VB.Label Label4 
      Caption         =   "處理方式："
      Height          =   225
      Left            =   60
      TabIndex        =   52
      Top             =   4500
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "案件進度說明："
      Height          =   225
      Left            =   60
      TabIndex        =   51
      Top             =   3630
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "理由："
      Height          =   225
      Left            =   60
      TabIndex        =   47
      Top             =   3060
      Width           =   615
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   11
      Left            =   60
      TabIndex        =   46
      Top             =   630
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "案件名稱："
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   990
      TabIndex        =   45
      Top             =   630
      Width           =   6210
      VariousPropertyBits=   27
      Size            =   "10954;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   44
      Top             =   900
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "申  請  人："
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   990
      TabIndex        =   43
      Top             =   900
      Width           =   4440
      VariousPropertyBits=   27
      Size            =   "7832;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1320
      X2              =   2880
      Y1              =   450
      Y2              =   450
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   42
      Top             =   360
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "本所案號："
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm210139"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; grd1=新細明體-ExtB、Label1(index)、Text1(index)
'Memo by Lydia 2019/07/01 表單名稱:銷案銷帳單=>銷案／銷帳單
'Create By Sindy 2013/3/15
Option Explicit

Dim m_row As Integer, i As Integer
Dim strCP01 As String
Dim strCP02 As String
Dim strCP03 As String
Dim strCP04 As String
Dim m_Nation As String
Dim m_TM34 As String
Dim m_FileName As String, m_TempFileName As String


Private Sub Check1_Click(Index As Integer)
   If Check1(4).Value = 1 Then '改請
      cmdOK(1).Caption = "E-Mail(&E)"
      txtPCnt.Text = 1
      txtPCnt.Visible = False
   Else
      cmdOK(1).Caption = "列印(&P)　　份"
      txtPCnt.Visible = True
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim bolChk As Boolean
Dim strName As String, strText As String
Dim intPCnt As Integer
Dim strCP10Nm As String
Dim strCP06 As String 'Add By Sindy 2021/9/23
Dim strP605CP09 As String, strP605Memo As String 'Added by Morgan 2016/11/14
Dim strTo As String 'Add by Amy 2024/05/15
Dim strSys As String, stComp As String 'Add By Sindy 2024/8/23
   
On Error GoTo ErrHand

   Select Case Index
      Case 0
         If Trim(Txt1(0)) = "" Or Trim(Txt1(1)) = "" Then
            MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         GRD1.MousePointer = flexHourglass
         doQuery
         GRD1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
      Case 1
         '檢查是否有勾選資料
         m_row = 0
         strCP10Nm = "": strCP06 = ""
         For i = 1 To GRD1.Rows - 1
            If Trim(GRD1.TextMatrix(i, 0)) <> "" Then
               If m_row = 0 Then m_row = i '取得勾選的第一筆資料列
               If strCP10Nm = "" Then
                  strCP10Nm = Trim(GRD1.TextMatrix(i, 3))
               Else
                  strCP10Nm = strCP10Nm & "，" & Trim(GRD1.TextMatrix(i, 3))
               End If
               
               'Add By Sindy 2021/9/23
               If strCP06 = "" And Trim(GRD1.TextMatrix(i, 7)) <> "" Then
                  strCP06 = DBDATE(Replace(Trim(GRD1.TextMatrix(i, 7)), "/", ""))
                  strCP06 = Left(strCP06, 4) - 1911 & "年" & Mid(strCP06, 5, 2) & "月" & Right(strCP06, 2) & "日"
               End If
               '2021/9/23 END
               
               If GRD1.TextMatrix(i, 14) = "P605" Then strP605CP09 = GRD1.TextMatrix(i, 2) 'Added by Morgan 2016/11/14
               
            End If
         Next i
         If m_row <> 0 Then
            If GRD1.TextMatrix(m_row, 2) = "" Then
               MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
               txt1_GotFocus 0
               Exit Sub
            Else
               '檢查資料
               '通知單
               bolChk = False
               For i = 0 To 4
                  If Check1(i).Value = 1 Then
                     bolChk = True
                     Exit For
                  End If
               Next i
               If bolChk = False Then
                  MsgBox "請勾選通知單！"
                  Exit Sub
               End If
               '理由
               If Trim(Text1(0).Text) = "" Then
                  MsgBox "理由不可空白！"
                  Text1(0).SetFocus
                  Exit Sub
               End If
               '案件進度說明
               bolChk = False
               For i = 0 To 5
                  If Option1(i).Value = True Then
                     bolChk = True
                     Exit For
                  End If
               Next i
               If bolChk = False Then
                  MsgBox "請點選案件進度說明！"
                  Exit Sub
               End If
               '案件資料處理方式
               bolChk = False
               For i = 0 To 5
                  If Option2(i).Value = True Then
                     bolChk = True
                     Exit For
                  End If
               Next i
               If bolChk = False Then
                  MsgBox "請點選案件資料處理方式！"
                  Exit Sub
               End If
               '款項處理方式
               bolChk = False
               For i = 0 To 3
                  If Option3(i).Value = True Then
                     bolChk = True
                     Exit For
                  End If
               Next i
               If bolChk = False Then
                  MsgBox "請點選款項處理方式！"
                  Exit Sub
               End If
               '收據處理方式
               bolChk = False
               For i = 0 To 1
                  If Option4(i).Value = True Then
                     bolChk = True
                     Exit For
                  End If
               Next i
               If bolChk = False Then
                  MsgBox "請點選收據處理方式！"
                  Exit Sub
               End If
               
               'Added by Morgan 2016/11/14
               'P新型案年費銷帳檢查若為一案兩請且發明案未核准且未閉卷(不必限定大陸案)確認並列印於理由 --秀玲
               strP605Memo = ""
               If strP605CP09 <> "" Then
                  strExc(0) = "select p2.pa01||'-'||p2.pa02||decode(p2.pa03||p2.pa04,'000','','-'||p2.pa03||'-'||p2.pa04) CaseNo,p2.pa16 from caseprogress, patent p1, casemap, patent p2 where cp09='" & strP605CP09 & "' and cp01='P' and cp10='605' and p1.pa01(+)=cp01 and p1.pa02(+)=cp02 and p1.pa03(+)=cp03 and p1.pa04(+)=cp04 and p1.pa08='2' and cm01(+)=p1.pa01 and cm02(+)=p1.pa02 and cm03(+)=p1.pa03 and cm04(+)=p1.pa04 and cm10='3' and p2.pa01(+)=cm05 and p2.pa02(+)=cm06 and p2.pa03(+)=cm07 and p2.pa04(+)=cm08 and p2.pa08='1' and p2.pa57 is null and nvl(p2.pa16,'2')='2'"
                  strExc(0) = strExc(0) & " union select p2.pa01||'-'||p2.pa02||decode(p2.pa03||p2.pa04,'000','','-'||p2.pa03||'-'||p2.pa04) CaseNo,p2.pa16 from caseprogress, patent p1, casemap, patent p2 where cp09='" & strP605CP09 & "' and cp01='P' and cp10='605' and p1.pa01(+)=cp01 and p1.pa02(+)=cp02 and p1.pa03(+)=cp03 and p1.pa04(+)=cp04 and p1.pa08='2' and cm05(+)=p1.pa01 and cm06(+)=p1.pa02 and cm07(+)=p1.pa03 and cm08(+)=p1.pa04 and cm10='3' and p2.pa01(+)=cm01 and p2.pa02(+)=cm02 and p2.pa03(+)=cm03 and p2.pa04(+)=cm04 and p2.pa08='1' and p2.pa57 is null and nvl(p2.pa16,'2')='2'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strP605Memo = "本案為一案兩請之新型案，目前發明案(" & RsTemp(0) & ")" & IIf("" & RsTemp("pa16") = "2", "被核駁但尚未閉卷", "尚未審定且未閉卷") & "。"
                     If MsgBox(strP605Memo & "是否確定要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        Exit Sub
                     End If
                  End If
               End If
               'end 2016/11/14
               
               Screen.MousePointer = vbHourglass
               GRD1.MousePointer = flexHourglass
               
               '產生報表
               '判斷word是否已開啟
               'Modify By Sindy 2016/1/30
               'If g_WordAp Is Nothing Then
               If TypeName(g_WordAp) <> "Application" Then
               '2016/1/30 END
RestarWord:
               Set g_WordAp = New Word.Application
               g_WordAp.Visible = False
      '         ElseIf g_WordAp.Visible = True Then
      '            bVisible = True
               End If
               m_TempFileName = Txt1(0) & Txt1(1) & IIf(Txt1(2) & Txt1(3) = "000", "", Txt1(2) & Txt1(3)) & "銷案銷帳單.doc"
               If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
                  Kill PUB_Getdesktop & "\" & m_TempFileName
               End If
               'Added by Morgan 2023/7/24
               'If Dir(App.path & "\" & strUserNum & "\" & m_FileName) = "" Then 'Modify By Sindy 2024/8/23 換樣版
                  Call PUB_GetSampleFile(m_FileName, "M51-000099-0-00", , App.path & "\" & strUserNum & "\")
                  Sleep 300 'Add By Sindy 2024/8/23
               'End If
               'end 2023/7/24
               g_WordAp.Documents.Open App.path & "\" & strUserNum & "\" & m_FileName
               g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & m_TempFileName
               g_WordAp.ActiveDocument.Close
               g_WordAp.Documents.Open PUB_Getdesktop & "\" & m_TempFileName
               With g_WordAp
                  .Selection.WholeStory
                  .Selection.Copy
                  For i = 0 To 33 '32 '31 '30
                     strName = ""
                     strText = ""
                     'Add By Sindy 2024/8/23
                     If i = 0 Then
                        strName = "公司抬頭"
                        strSys = CheckSys(strCP01) '系統種類
                        If (strSys = "3" Or strSys = "4" Or strSys = "7" Or strSys = "8") _
                           And strCP01 <> "ACS" Then
                           strText = CompNameQuery("L")
                        Else
                           stComp = PUB_GetReceiptComp(strCP01, strCP02, strCP03, strCP04)
                           If stComp = "" Then
                              strText = CompNameQuery("2")
                           Else
                              strText = CompNameQuery(stComp)
                           End If
                        End If
                     '2024/8/23 END
                     ElseIf i = 1 Then
                        strName = "通知單"
                        strText = IIf(Check1(0) = 0, "□", "■") & Check1(0).Caption & "　" & _
                                  IIf(Check1(1) = 0, "□", "■") & Check1(1).Caption & "　" & _
                                  IIf(Check1(2) = 0, "□", "■") & Check1(2).Caption & "　" & _
                                  IIf(Check1(3) = 0, "□", "■") & Check1(3).Caption & "　" & _
                                  IIf(Check1(4) = 0, "□", "■") & Check1(4).Caption & "　"
                     ElseIf i = 2 Then
                        strName = "本所案號"
                        strText = strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04)
                     ElseIf i = 3 Then
                        strName = "填表日期"
                        strText = Val(Left(strSrvDate(1), 4)) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Right(strSrvDate(1), 2) & "日"
'                     ElseIf i = 3 Then
'                        strName = "分所案號"
'                        strText = m_TM34
                     'Modify By Sindy 2021/9/23
                     ElseIf i = 4 Then
                        strName = "本所期限"
                        strText = strCP06
                        '2021/9/23 END
                     ElseIf i = 5 Then
                        strName = "案件收文日期"
                        strText = DBDATE(Trim(GRD1.TextMatrix(m_row, 1)))
                        strText = Val(Left(strText, 4)) - 1911 & "年" & Mid(strText, 5, 2) & "月" & Right(strText, 2) & "日"
                     ElseIf i = 6 Then
                        strName = "案件名稱"
                        'Modify By Sindy 2020/1/22 T-222628
                        'strText = Left(Trim(Label1(1)), 14)
                        strText = LeftB(Trim(Label1(1)), 30) & IIf(LenB(Trim(Label1(1))) > 30, "...", "")
                     ElseIf i = 7 Then
                        strName = "申請人"
                        'Add By Sindy 2015/10/15
                        If Trim(Label1(3)) <> "" Then
                        '2015/10/15 END
                           strText = Left(Trim(Right(Trim(Label1(3)), Len(Trim(Label1(3))) - 10)), 17)
                        End If
                     ElseIf i = 8 Then
                        strName = "單位"
                        strText = Replace(Trim(GRD1.TextMatrix(m_row, 13)), "主管", "")
                     'Modify By Sindy 2016/5/27
'                     ElseIf i = 8 Then
'                        strName = "接洽人"
'                        strText = Trim(grd1.TextMatrix(m_row, 6))
                     ElseIf i = 9 Then
                        strName = "操作者"
                        strText = strUserName
                     '2016/5/27 END
                     ElseIf i = 10 Then
                        strName = "案件性質"
                        strText = strCP10Nm 'Trim(grd1.TextMatrix(m_row, 3))
                     ElseIf i = 11 Then
                        strName = "理由"
                        strText = Trim(Text1(0))
                        If strP605Memo <> "" Then
                           strText = strText & vbCrLf & "＊" & strP605Memo
                        End If
                     ElseIf i = 12 Then
                        strName = "案件進度說明1"
                        strText = IIf(Option1(0).Value = False, "□", "■") & Option1(0).Caption & "　" & _
                                  IIf(Option1(1).Value = False, "□", "■") & Option1(1).Caption & "　" & _
                                  IIf(Option1(2).Value = False, "□", "■") & Option1(2).Caption & "　" & _
                                  IIf(Option1(3).Value = False, "□", "■") & Option1(3).Caption
                     ElseIf i = 13 Then
                        strName = "案件進度說明2"
                        strText = IIf(Option1(4).Value = False, "□", "■") & Option1(4).Caption & "　" & _
                                  IIf(Option1(5).Value = False, "□", "■") & Option1(5).Caption
                     ElseIf i = 14 Then
                        strName = "其他內容1"
                        strText = IIf(Text1(1).Enabled = True, Trim(Text1(1)), "")
                     ElseIf i = 15 Then
                        strName = "資料銷毀"
                        strText = IIf(Option2(0).Value = False, "□", "■") & Option2(0).Caption
                     ElseIf i = 16 Then
                        strName = "資料退寄客戶"
                        strText = IIf(Option2(1).Value = False, "□", "■") & Option2(1).Caption
                     ElseIf i = 17 Then
                        strName = "址1"
                        strText = IIf(Text1(2).Enabled = True, Trim(Text1(2)), "")
                     ElseIf i = 18 Then
                        strName = "轉存相關卷內"
                        strText = IIf(Option2(2).Value = False, "□", "■") & Option2(2).Caption
                     ElseIf i = 19 Then
                        strName = "卷號"
                        strText = IIf(Text1(3).Enabled = True, Trim(Text1(3)), "　　")
                     ElseIf i = 20 Then
                        strName = "存舊卷內"
                        strText = IIf(Option2(3).Value = False, "□", "■") & Option2(3).Caption
                     ElseIf i = 21 Then
                        strName = "本所立新檔"
                        strText = IIf(Option2(4).Value = False, "□", "■") & Option2(4).Caption
                     ElseIf i = 22 Then
                        strName = "其他2"
                        strText = IIf(Option2(5).Value = False, "□", "■") & Option2(5).Caption
                     ElseIf i = 23 Then
                        strName = "其他內容2"
                        strText = IIf(Text1(4).Enabled = True, Trim(Text1(4)), "")
                     ElseIf i = 24 Then
                        strName = "金額"
                        strText = IIf(Trim(Text1(5)) = "", "　　　　", Trim(Text1(5)))
                     ElseIf i = 25 Then
                        strName = "款項處理方式"
                        strText = IIf(Option3(0).Value = False, "□", "■") & Option3(0).Caption & "　" & _
                                  IIf(Option3(1).Value = False, "□", "■") & Option3(1).Caption
                     ElseIf i = 26 Then
                        strName = "由財務處寄客戶"
                        strText = IIf(Option3(2).Value = False, "□", "■") & Option3(2).Caption
                     ElseIf i = 27 Then
                        strName = "址2"
                        strText = IIf(Text1(6).Enabled = True, Trim(Text1(6)), "　　")
                     ElseIf i = 28 Then
                        strName = "其他3"
                        strText = IIf(Option3(3).Value = False, "□", "■") & Trim(Option3(3).Caption)
                     ElseIf i = 29 Then
                        strName = "其他內容3"
                        strText = IIf(Text1(8).Enabled = True, Trim(Text1(8)), "")
                     ElseIf i = 30 Then
                        strName = "收據處理方式"
                        strText = IIf(Option4(0).Value = False, "□", "■") & Option4(0).Caption & "　" & _
                                  IIf(Option4(1).Value = False, "□", "■") & Option4(1).Caption
                     ElseIf i = 31 Then
                        strName = "其他內容4"
                        strText = IIf(Text1(7).Enabled = True, Trim(Text1(7)), "")
                     'Add By Sindy 2021/3/9
                     ElseIf i = 32 Then
                        strName = "申請國家"
                        strText = Label1(5)
                     'Add By Sindy 2021/8/19
                     ElseIf i = 33 Then
                        strName = "單位主管意見"
                        If Check1(4).Value = 1 Then
                           strText = "改請單不需單位主管簽核"
                        End If
                     End If
                     
                     If Trim(strName) <> "" Then
                        'Add By Sindy 2021/8/19
                        If i = 33 Then
                           .ActiveDocument.Shapes("Text Box 124").Select
                        End If
                        '2021/8/19 END
                        .Selection.Find.ClearFormatting
                        .Selection.Find.Text = "|#" & strName & "#|"
                        .Selection.Find.Replacement.Text = ""
                        .Selection.Find.Forward = True
                        .Selection.Find.Wrap = wdFindContinue
                        .Selection.Find.Format = False
                        .Selection.Find.MatchCase = False
                        .Selection.Find.MatchWholeWord = False
                        .Selection.Find.MatchWildcards = False
                        .Selection.Find.MatchSoundsLike = False
                        .Selection.Find.MatchAllWordForms = False
                        .Selection.Find.MatchByte = True
                        .Selection.Find.Execute
                        .Selection.Delete
                        .Selection.TypeText strText
'                        Selection.Find.ClearFormatting
'                        With Selection.Find
'                            .Text = "|#其他內容3#|"
'                            .Replacement.Text = ""
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .MatchCase = False
'                            .MatchWholeWord = False
'                            .MatchWildcards = False
'                            .MatchSoundsLike = False
'                            .MatchAllWordForms = False
'                            .MatchByte = True
'                        End With
'                        Selection.Find.Execute
                         'Modify By Sindy 2021/9/23 標示網底
                        If strName = "本所期限" And strText <> "" Then
                           .Selection.HomeKey
                           .Selection.Find.ClearFormatting
                           With .Selection.Find
                               .Text = strText
                               .Replacement.Text = ""
                               .Forward = True
                               .Wrap = wdFindContinue
                               .Format = False
                               .MatchCase = False
                               .MatchWholeWord = False
                               .MatchWildcards = False
                               .MatchSoundsLike = False
                               .MatchAllWordForms = False
                               .MatchByte = True
                           End With
                           .Selection.Find.Execute
                           .Selection.Range.HighlightColorIndex = wdGray25
                        '2021/9/23 END
                        ElseIf i = 14 Or i = 17 Or i = 19 Or i = 23 Or i = 24 Or i = 27 Or i = 29 Or i = 31 Then
                           If Trim(strText) <> "" Then
                              .Selection.HomeKey
                              .Selection.Find.ClearFormatting
                              With .Selection.Find
                                  .Text = strText
                                  .Replacement.Text = ""
                                  .Forward = True
                                  .Wrap = wdFindContinue
                                  .Format = False
                                  .MatchCase = False
                                  .MatchWholeWord = False
                                  .MatchWildcards = False
                                  .MatchSoundsLike = False
                                  .MatchAllWordForms = False
                                  .MatchByte = True
                              End With
                              .Selection.Find.Execute
   '                           If Selection.Font.Underline = wdUnderlineNone Then
                                  .Selection.Font.Underline = wdUnderlineSingle '底線
   '                           Else
   '                               Selection.Font.Underline = wdUnderlineNone
   '                           End If
                           End If
                        End If
                     End If
                  Next i
               End With
'               If intRow = 1 Then
'                  g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & strFile
'               Else
'                  g_WordAp.ActiveDocument.Save 'PUB_Getdesktop & "\" & strFile
'               End If
'               g_WordAp.ActiveDocument.Close
'               g_WordAp.Documents.Open PUB_Getdesktop & "\" & strFile
               
               GRD1.MousePointer = flexDefault
               Screen.MousePointer = vbDefault
               
               If Check1(4).Value = 0 Then '非改請時則需列印
                  'Modify By Sindy 2016/1/30
'                  For intPCnt = 1 To Val(txtPCnt.Text)
'                     g_WordAp.PrintOut
'                  Next intPCnt
                  g_WordAp.PrintOut Background:=False, Copies:=Val(txtPCnt.Text), Collate:=True
                  '2016/1/30 END
               End If
                  
               g_WordAp.ActiveDocument.Save
               
               'Modify By Sindy 2016/1/30
               'g_WordAp.ActiveDocument.Close
               g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
               g_WordAp.Quit wdDoNotSaveChanges
               Set g_WordAp = Nothing
               '2016/1/30 END
               'g_WordAp.Quit
'               Set g_WordAp = Nothing
               
               If Check1(4).Value = 1 Then '改請
                  '改請時會寄發E-Mail給財務處及電腦中心
                  'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("程式管理人員")& ";83002"
                  If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
                      strTo = Pub_GetSpecMan("財務處應收處理人員")
                  Else
                     strTo = Pub_GetSpecMan("財務處總帳人員")
                  End If
                  strTo = strTo & ";" & Pub_GetSpecMan("程式管理人員")
                  PUB_SendMail strUserNum, strTo, "", strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & "改請通知單！", "Dear Sirs," & vbCrLf & vbCrLf & "　　改請通知單乙份，同附件！", , PUB_Getdesktop & "\" & m_TempFileName
                  'end 2024/05/15
                  MsgBox "檔案已存檔，並且放置：" & PUB_Getdesktop & "\" & m_TempFileName & vbCrLf & vbCrLf & vbCrLf & "改請程序不需主管簽核, 此改請單後續由電腦中心處理並列印, 請不要再列印, 謝謝！"
               Else
                  'Add By Sindy 2013/8/14
                  If Left(Trim(Pub_StrUserSt15), 1) = "S" Then
                     '2013/9/16 modify by sonia 取消通知財務處,瑞婷說杜副總可能退件或擋住,她不想追,故仍由邱素蓮將已核可之案件自行e-mail給財務處
                     'modify by sonia 2019/4/12邱素蓮調職改成莊敏惠73017
                     'modify by sonia 2019/5/15再改智權委辦區ip_transfer
                     PUB_SendMail strUserNum, "ip_transfer", "", strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & "非改請通知單！", "Dear Sirs," & vbCrLf & vbCrLf & "　　非改請通知單乙份，同附件！", , PUB_Getdesktop & "\" & m_TempFileName
                  End If
                  '2013/8/14 END
                  MsgBox "列印完成！檔案已存放在：" & PUB_Getdesktop & "\" & m_TempFileName
                  'ShowPrintOk
               End If
               Clipboard.Clear 'Modify By Sindy 2013/8/22 清除剪貼簿動作
            End If
         Else
            If GRD1.Rows = 2 And GRD1.TextMatrix(1, 2) = "" Then
               MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
               txt1_GotFocus 0
               Exit Sub
            Else
               MsgBox "請先選擇一筆要列印的資料！", vbCritical, "操作錯誤！"
               Exit Sub
            End If
         End If
      Case 2
         Unload Me
   Case Else
   End Select
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

'清除欄位值
Sub ClearData()
   Label1(1) = ""
   Label1(3) = ""
   Label1(5) = ""
   m_Nation = ""
   For i = 0 To 4
      Check1(i).Value = False
   Next i
   For i = 0 To 5
      Option1(i).Value = False
      Text1(1).Enabled = False
   Next i
   For i = 0 To 5
      Option2(i).Value = False
      Text1(2).Enabled = False
      Text1(3).Enabled = False
      Text1(4).Enabled = False
   Next i
   For i = 0 To 3
      Option3(i).Value = False
      Text1(6).Enabled = False
      Text1(8).Enabled = False
   Next i
   For i = 0 To 1
      Option4(i).Value = False
      Text1(7).Enabled = False
   Next i
   'Modify By Sindy 2013/8/22 理由欄位不清資料值
   'For i = 0 To 8
   For i = 1 To 8
      Text1(i).Text = ""
   Next i
End Sub

Sub doQuery()

On Error GoTo ErrHnd

   m_row = 0
   strCP01 = UCase(Txt1(0))
   strCP02 = Txt1(1)
   strCP03 = Left(Txt1(2) & "0", 1)
   strCP04 = Left(Txt1(3) & "00", 2)
   Txt1(0) = strCP01
   Txt1(1) = strCP02
   Txt1(2) = strCP03
   Txt1(3) = strCP04
   
   '清除欄位值
   Call ClearData
   
   '基本檔資料
   strSql = "SELECT TM12,TM05||TM06||TM07,TM23||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,TM10,TM34" & _
                " From Trademark, nation, Customer" & _
                " WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "'" & _
                " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
                " AND TM10=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT PA11,PA05||PA06||PA07,PA26||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,PA09,PA47 as TM34" & _
                " From Patent, nation, Customer" & _
                " WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "'" & _
                " AND SUBSTR(PA26,1,8)=CU01(+) AND decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)" & _
                " AND PA09=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT '',LC05||LC06||LC07,LC11||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,LC15,LC16 as TM34" & _
                " From LawCase, nation, Customer" & _
                " WHERE LC01='" & strCP01 & "' AND LC02='" & strCP02 & "' AND LC03='" & strCP03 & "' AND LC04='" & strCP04 & "'" & _
                " AND SUBSTR(LC11,1,8)=CU01(+) AND decode(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+)" & _
                " AND LC15=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT '',HC06,HC05||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),' ',' ',HC07 as TM34" & _
                " From HireCase, Customer" & _
                " WHERE HC01='" & strCP01 & "' AND HC02='" & strCP02 & "' AND HC03='" & strCP03 & "' AND HC04='" & strCP04 & "'" & _
                " AND SUBSTR(HC05,1,8)=CU01(+) AND decode(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+)"
   strSql = strSql & " Union " & _
                "SELECT SP11,SP05||SP06||SP07,SP08||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,SP09,SP28 as TM34" & _
                " From Servicepractice, nation, Customer" & _
                " WHERE SP01='" & strCP01 & "' AND SP02='" & strCP02 & "' AND SP03='" & strCP03 & "' AND SP04='" & strCP04 & "'" & _
                " AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)" & _
                " AND SP09=NA01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Label1(1) = "" & Trim(RsTemp(1)) '案件名稱
      Label1(3) = "" & Trim(RsTemp(2)) '申請人
      Label1(5) = "" & Trim(RsTemp(3)) '申請國家
      m_Nation = "" & Trim(RsTemp(4))
      m_TM34 = "" & Trim(RsTemp("TM34")) '分所案號
   Else
      m_Nation = "000"
   End If
   
   '進度檔資料
   'Added by Lydia 2023/12/28
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "SELECT ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(" & CNULL(m_Nation) & ",'000',CPM03,CPM04),CP10) as 案件性質," & _
               "NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限," & _
               "SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP64 as 進度備註,NVL(A0922,A0902) AS a0902,CP01||CP10 TAG1 " & _
               "FROM caseprogress,staff s1,staff s2,CUSTOMER,CASEPROPERTYMAP,acc090,ACC090NEW " & _
               "WHERE CP01='" & strCP01 & "' and CP02='" & strCP02 & "' and  CP03='" & strCP03 & "' and CP04='" & strCP04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
               "AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) " & _
               "AND CP01=CPM01(+) AND CP10=CPM02(+) AND S2.ST03=A0901(+) AND S2.ST93=A0921(+) " & _
               "order by CP05,CP09 asc"
   Else
   'end 2023/12/28
      strSql = "SELECT ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(" & CNULL(m_Nation) & ",'000',CPM03,CPM04),CP10) as 案件性質," & _
               "NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限," & _
               "SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP64 as 進度備註,a0902,CP01||CP10 TAG1 " & _
               "FROM caseprogress,staff s1,staff s2,CUSTOMER,CASEPROPERTYMAP,acc090 " & _
               "WHERE CP01='" & strCP01 & "' and CP02='" & strCP02 & "' and  CP03='" & strCP03 & "' and CP04='" & strCP04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & _
               "AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) " & _
               "AND CP01=CPM01(+) AND CP10=CPM02(+) AND S2.ST03=A0901(+) " & _
               "order by CP05,CP09 asc"
   End If
   CheckOC3
   GRD1.Rows = 2
   GRD1.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set GRD1.Recordset = AdoRecordSet3.Clone
         GRD1.row = 1
      Else
         MsgBox "無符合資料！", vbInformation
         Exit Sub
      End If
   End With
   
   'Add By Sindy 2015/1/12 以防未查詢就列印
   cmdOK(1).Enabled = True
   
   Exit Sub
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   Call settxtPCnt
   
   
   m_FileName = "銷案銷帳單_樣本.doc"
   'm_TempFileName = "$$TempDoc.doc"
   'PUB_OpenFtpFile "", m_FileName, Winsock1, "RptSample", False, , True
   'Removed by Morgan 2023/7/24 改列印時下載,否則同時開別的功能時可能會被刪除,且也不一定會列印
   'Call PUB_GetSampleFile(m_FileName, "M51-000099-0-00", , App.path & "\" & strUserNum & "\")
   'end 2023/7/24
   
   'Add By Sindy 2013/8/14
   Label15.Visible = False
   If Left(Trim(Pub_StrUserSt15), 1) = "S" Or Pub_StrUserSt03 = "M51" Then
      Label15.Visible = True
   End If
   '2013/8/14 END
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo ErrHand
   
'   If Not g_WordAp Is Nothing Then
'      g_WordAp.Quit
'CloseWord:
'      Set g_WordAp = Nothing
'   End If
   
   Set frm210139 = Nothing
   
   Exit Sub
   
'ErrHand:
'   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
'      GoTo CloseWord
'   ElseIf Err.Number <> 0 Then
'      MsgBox (Err.Description)
'   End If
End Sub

Private Sub settxtPCnt()
Dim strST06 As String
   
   strST06 = PUB_GetST06(strUserNum)
   If strST06 = "1" Then
      txtPCnt = "1"
   Else
      'Modify By Sindy 2015/5/13
      'txtPCnt = "2"
      txtPCnt = "1"
      '2015/5/13 END
      '2013/10/11 ADD BY SONIA 分所智權部已改發E-MAIL給邱素蓮,故改為一張,作業完成後小真會影印一份寄下去
      If Left(Trim(Pub_StrUserSt15), 1) = "S" Then
         txtPCnt = "1"
      End If
      '2013/10/11 END
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim m_mouseRow As Integer
   
'   grd1.Visible = False
'   m_mouseRow = grd1.row
'   grd1.col = 0
'   If m_mouseRow <> 0 Then
'       If m_row <> 0 Then
'           grd1.row = m_row
'            For i = 0 To grd1.Cols - 1
'                 grd1.col = i
'                 If grd1.CellBackColor = &HFFC0C0 Then
'                   grd1.CellBackColor = &H80000018
'                   grd1.TextMatrix(m_row, 0) = ""
'                 Else
'                   grd1.CellBackColor = &HFFC0C0 '&H80000018 '&H8080FF
'                   grd1.TextMatrix(m_row, 0) = "V"
'                 End If
'           Next i
'       End If
'       If m_row <> m_mouseRow Then
'           grd1.row = m_mouseRow
'           m_row = m_mouseRow
'            For i = 0 To grd1.Cols - 1
'                 grd1.col = i
'                 If grd1.CellBackColor = &HFFC0C0 Then
'                   grd1.CellBackColor = &H80000018
'                   grd1.TextMatrix(m_row, 0) = ""
'                   m_row = 0
'                 Else
'                   grd1.CellBackColor = &HFFC0C0
'                   grd1.TextMatrix(m_row, 0) = "V"
'                 End If
'           Next i
'       Else
'           m_row = 0
'       End If
'   End If
'   grd1.Visible = True
   
   GRD1.Visible = False
   m_mouseRow = GRD1.row
   GRD1.col = 0
   If m_mouseRow <> 0 Then
      GRD1.row = m_mouseRow
      If Trim(GRD1.TextMatrix(GRD1.row, 0)) = "" Then
         GRD1.TextMatrix(GRD1.row, 0) = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      Else
         GRD1.TextMatrix(GRD1.row, 0) = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &H80000018
         Next i
      End If
   End If
   GRD1.Visible = True
End Sub

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   GRD1.Visible = False
   arrGridHeadText = Array("V", "收文日", "總收文號", "案件性質", "相關收文號" _
             , "承辦人", "智權人員", "本所期限", "法定期限", "發文日" _
             , "取消收文", "相關人", "進度備註", "a0902", "tag1")
   arrGridHeadWidth = Array(180, 788, 938, 950, 938 _
                      , 593, 593, 788, 788, 788 _
                      , 788, 600, 800, 0, 0)
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignLeftCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
   If Option1(5).Value = False Then
      Text1(1).Enabled = False
      'Text1(1).Text = ""
   Else
      Text1(1).Enabled = True
   End If
End Sub

Private Sub Option2_Click(Index As Integer)
   If Option2(1).Value = False Then
      Text1(2).Enabled = False
      'Text1(2).Text = ""
   Else
      Text1(2).Enabled = True
   End If
   If Option2(2).Value = False Then
      Text1(3).Enabled = False
      'Text1(3).Text = ""
   Else
      Text1(3).Enabled = True
   End If
   If Option2(5).Value = False Then
      Text1(4).Enabled = False
      'Text1(4).Text = ""
   Else
      Text1(4).Enabled = True
   End If
End Sub

Private Sub Option3_Click(Index As Integer)
   If Option3(2).Value = False Then
      Text1(6).Enabled = False
      'Text1(6).Text = ""
   Else
      Text1(6).Enabled = True
   End If
   If Option3(3).Value = False Then
      Text1(8).Enabled = False
      'Text1(8).Text = ""
   Else
      Text1(8).Enabled = True
   End If
End Sub

Private Sub Option4_Click(Index As Integer)
   If Option4(1).Value = False Then
      Text1(7).Enabled = False
      'Text1(7).Text = ""
   Else
      Text1(7).Enabled = True
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
   Select Case Index
   Case 3, 5
      CloseIme
   Case Else
      OpenIme
   End Select
End Sub

'Modified by Lydia 2022/01/26 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
'   Case 5
'      KeyAscii = Pub_NumAscii(KeyAscii)
   Case 3
      KeyAscii = UpperCase(KeyAscii)
   Case Else
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strField As String

   If Text1(Index).Text = "" Then Exit Sub
   'Add By Sindy 2021/2/4
   If Index = 0 Then
      strField = "理由"
   ElseIf Index = 1 Then
      strField = "案件進度說明：其他"
   ElseIf Index = 2 Then
      strField = "資料退寄客戶，址："
   ElseIf Index = 3 Then
      strField = "轉存相關卷內，卷號："
   ElseIf Index = 4 Then
      strField = "處理方式：其他"
   ElseIf Index = 5 Then
      strField = "金額"
   ElseIf Index = 6 Then
      strField = "由財務處寄客戶，址："
   ElseIf Index = 7 Then
      strField = "收據處理方式：其他"
   ElseIf Index = 8 Then
      strField = "款項處理方式：其他"
   Else
      strField = ""
   End If
   '2021/2/4 END
   If Not CheckLengthIsOK(Text1(Index), Text1(Index).MaxLength, , strField) Then
      Cancel = True
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse Txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2015/1/12 以防未查詢就列印
'檢查本所案號若有異動時，則列印按鈕要鎖住，查詢後才能解開
Private Sub txt1_LostFocus(Index As Integer)
   If strCP01 = "" Or strCP02 = "" Then Exit Sub
   If strCP01 <> Txt1(0) Or strCP02 <> Txt1(1) Or strCP03 & strCP04 <> IIf(Txt1(2) & Txt1(3) = "000", "000", Txt1(2) & Txt1(3)) Then
      cmdOK(1).Enabled = False
      GRD1.Rows = 2
      GRD1.Clear
      SetDataListWidth
   End If
End Sub

Private Sub txtPCnt_GotFocus()
   TextInverse Me.txtPCnt
End Sub

Private Sub txtPCnt_KeyPress(KeyAscii As Integer)
   If KeyAscii <> vbKeyBack And KeyAscii <> Asc(1) And KeyAscii <> Asc(2) And KeyAscii <> Asc(3) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtPCnt_Validate(Cancel As Boolean)
   If txtPCnt = "" Then
      MsgBox "請輸入列印份數！", vbCritical
      Cancel = True
   End If
End Sub
