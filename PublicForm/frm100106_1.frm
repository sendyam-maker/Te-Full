VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100106_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以期限管制日查詢"
   ClientHeight    =   5832
   ClientLeft      =   4476
   ClientTop       =   3336
   ClientWidth     =   5328
   ControlBox      =   0   'False
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5832
   ScaleWidth      =   5328
   Begin VB.CommandButton cmdCusCase 
      BackColor       =   &H00C0FFC0&
      Caption         =   "個人客戶二個月之內期限未收文案件"
      Height          =   480
      Left            =   2065
      Style           =   1  '圖片外觀
      TabIndex        =   70
      Top             =   12
      Width           =   1650
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      Height          =   345
      Left            =   1350
      TabIndex        =   67
      Top             =   1560
      Width           =   3675
      Begin VB.TextBox txt6 
         Height          =   264
         Index           =   0
         Left            =   480
         MaxLength       =   7
         TabIndex        =   12
         Top             =   30
         Width           =   1212
      End
      Begin VB.TextBox txt6 
         Height          =   264
         Index           =   1
         Left            =   2310
         MaxLength       =   7
         TabIndex        =   13
         Top             =   30
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "迄"
         Height          =   210
         Index           =   4
         Left            =   1950
         TabIndex        =   69
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "起"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   68
         Top             =   30
         Width           =   375
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含FMP外專管制期限"
      Height          =   285
      Left            =   60
      TabIndex        =   66
      Top             =   150
      Width           =   3345
   End
   Begin VB.TextBox txt5 
      Height          =   264
      Index           =   14
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   20
      Top             =   3330
      Width           =   435
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   13
      Left            =   2370
      MaxLength       =   1
      TabIndex        =   32
      Text            =   " "
      Top             =   5550
      Width           =   345
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   11
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   30
      Text            =   " "
      Top             =   5220
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   12
      Left            =   2175
      MaxLength       =   6
      TabIndex        =   31
      Text            =   " "
      Top             =   5220
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   10
      Left            =   2175
      MaxLength       =   3
      TabIndex        =   29
      Text            =   " "
      Top             =   4890
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   9
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   28
      Text            =   " "
      Top             =   4890
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   8
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   27
      Text            =   " "
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   7
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   26
      Text            =   " "
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   6
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   25
      Text            =   " "
      Top             =   4230
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   5
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   24
      Text            =   " "
      Top             =   4230
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1515
      Left            =   45
      TabIndex        =   50
      Top             =   435
      Width           =   5100
      Begin VB.OptionButton OPT1 
         Caption         =   "承辦期限："
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1170
         Width           =   1245
      End
      Begin VB.OptionButton OPT1 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1245
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2610
         TabIndex        =   58
         Top             =   810
         Width           =   2300
         Begin VB.TextBox txt3 
            Height          =   288
            Index           =   3
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   10
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txt3 
            Height          =   288
            Index           =   2
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   9
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txt3 
            Height          =   288
            Index           =   1
            Left            =   0
            MaxLength       =   6
            TabIndex        =   8
            Top             =   0
            Width           =   972
         End
      End
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2610
         TabIndex        =   54
         Top             =   810
         Width           =   2300
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   2
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   57
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   1
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   56
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   55
            Top             =   0
            Width           =   1212
         End
      End
      Begin VB.TextBox txt3 
         Height          =   288
         Index           =   0
         Left            =   1770
         MaxLength       =   3
         TabIndex        =   7
         Top             =   810
         Width           =   732
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   0
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   4
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   1
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   2
         Top             =   150
         Width           =   1212
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   1
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   5
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   0
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   1
         Top             =   150
         Width           =   1212
      End
      Begin VB.OptionButton OPT1 
         Caption         =   "法定期限："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   1245
      End
      Begin VB.OptionButton OPT1 
         Caption         =   "本所期限："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "起"
         Height          =   210
         Index           =   0
         Left            =   1410
         TabIndex        =   65
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "起"
         Height          =   210
         Index           =   1
         Left            =   1410
         TabIndex        =   53
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "迄"
         Height          =   210
         Index           =   3
         Left            =   3240
         TabIndex        =   52
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "迄"
         Height          =   210
         Index           =   2
         Left            =   3240
         TabIndex        =   51
         Top             =   510
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3750
      Style           =   1  '圖片外觀
      TabIndex        =   36
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4530
      Style           =   1  '圖片外觀
      TabIndex        =   37
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   18
      Top             =   2724
      Width           =   2800
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   1
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   19
      Text            =   " "
      Top             =   3024
      Width           =   615
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   2
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   21
      Text            =   " "
      Top             =   3630
      Width           =   735
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   4
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   23
      Text            =   " "
      Top             =   3930
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "查詢"
      Height          =   1332
      Left            =   3465
      TabIndex        =   45
      Top             =   3900
      Width           =   1635
      Begin VB.OptionButton OPT2 
         Caption         =   "未收文"
         Height          =   228
         Index           =   0
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton OPT2 
         Caption         =   "已收文未發文"
         Height          =   252
         Index           =   1
         Left            =   180
         TabIndex        =   34
         Top             =   570
         Width           =   1400
      End
      Begin VB.OptionButton OPT2 
         Caption         =   "已收文已發文"
         Height          =   252
         Index           =   2
         Left            =   180
         TabIndex        =   35
         Top             =   960
         Width           =   1400
      End
   End
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   36
      TabIndex        =   38
      Top             =   1905
      Width           =   5130
      Begin VB.TextBox txt4 
         Height          =   270
         Index           =   0
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   14
         Top             =   150
         Width           =   1212
      End
      Begin VB.TextBox txt4 
         Height          =   270
         Index           =   1
         Left            =   3600
         MaxLength       =   9
         TabIndex        =   15
         Top             =   150
         Width           =   1212
      End
      Begin VB.TextBox txt4 
         Height          =   270
         Index           =   2
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   16
         Top             =   450
         Width           =   1212
      End
      Begin VB.TextBox txt4 
         Height          =   270
         Index           =   3
         Left            =   3600
         MaxLength       =   9
         TabIndex        =   17
         Top             =   450
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "迄"
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   44
         Top             =   450
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "迄"
         Height          =   180
         Index           =   0
         Left            =   3240
         TabIndex        =   43
         Top             =   150
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "起"
         Height          =   180
         Index           =   2
         Left            =   1320
         TabIndex        =   42
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "起"
         Height          =   180
         Index           =   3
         Left            =   1320
         TabIndex        =   41
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "代理人："
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   40
         Top             =   450
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "申請人："
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.TextBox txt5 
      Height          =   270
      Index           =   3
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   22
      Text            =   " "
      Top             =   3630
      Width           =   735
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   1
      Left            =   1740
      TabIndex        =   72
      Top             =   3930
      Width           =   1365
      VariousPropertyBits=   27
      Size            =   "2408;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   0
      Left            =   1740
      TabIndex        =   71
      Top             =   3030
      Width           =   1365
      VariousPropertyBits=   27
      Size            =   "2408;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人組別：         ( 1 電子電機 2 化學 3 日文 4 機械設計 5 其他 )"
      Height          =   180
      Index           =   10
      Left            =   60
      TabIndex        =   64
      Top             =   3375
      Width           =   5070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否依智權人員排序跳頁：            (Y:是)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   63
      Top             =   5580
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FCP管制人："
      Height          =   180
      Index           =   8
      Left            =   60
      TabIndex        =   62
      Top             =   5250
      Width           =   1020
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   1935
      X2              =   2055
      Y1              =   5340
      Y2              =   5340
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1935
      X2              =   2055
      Y1              =   5010
      Y2              =   5010
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1920
      X2              =   2040
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   61
      Top             =   4590
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   60
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1920
      X2              =   2040
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   59
      Top             =   4260
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   49
      Top             =   3060
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   48
      Top             =   3645
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                   (ALL：全部)"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   47
      Top             =   2775
      Width           =   4905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   46
      Top             =   3960
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1920
      X2              =   2040
      Y1              =   3780
      Y2              =   3780
   End
End
Attribute VB_Name = "frm100106_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/05/17 Form2.0已修改: lbl1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit
Dim strSql As String, s As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'add by nickc 2005/08/11 台灣新型修正申復未續辦明細
Dim bolIsTw12011202 As Boolean
'Add by Amy 2016/07/13
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim stST15 As String '登入者區別
'Dim bolCmdCusCase As Boolean '是否按「個人客戶二個月之內期限未收文案件」

'92.04.16 nick
Public Sub PubShowNextData()
Dim bolNPExcept As Boolean

Select Case cmdState
Case 0
     cmdState = -1
    'add by nickc 2005/08/11
    bolIsTw12011202 = False
    'ALL或P的已收文未發文
    If (txt5(0) = "ALL" Or InStr(1, GetAddStr(txt5(0)), "'P'") <> 0) And OPT2(1).Value = True Then
        If Trim(txt5(5)) = "" And Trim(txt5(6)) = "" Then
            If Trim(txt5(7)) = "" And Trim(txt5(8)) = "" Then
               bolIsTw12011202 = True
            ElseIf Trim(txt5(7)) = "" And Trim(txt5(8)) <> "" Then
               If Trim(txt5(8)) >= "000" Then
                  bolIsTw12011202 = True
               End If
            ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) = "" Then
               If Trim(txt5(7)) <= "000" Then
                  bolIsTw12011202 = True
               End If
            ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) <> "" Then
               If Trim(txt5(7)) <= "000" And Trim(txt5(8)) >= "000" Then
                  bolIsTw12011202 = True
               End If
            End If
        ElseIf Trim(txt5(5)) = "" And Trim(txt5(6)) <> "" Then
            If Trim(txt5(6)) >= "1202" Then
               If Trim(txt5(7)) = "" And Trim(txt5(8)) = "" Then
                  bolIsTw12011202 = True
               ElseIf Trim(txt5(7)) = "" And Trim(txt5(8)) <> "" Then
                  If Trim(txt5(8)) >= "000" Then
                     bolIsTw12011202 = True
                  End If
               ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) = "" Then
                  If Trim(txt5(7)) <= "000" Then
                     bolIsTw12011202 = True
                  End If
               ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) <> "" Then
                  If Trim(txt5(7)) <= "000" And Trim(txt5(8)) >= "000" Then
                     bolIsTw12011202 = True
                  End If
               End If
            End If
        ElseIf Trim(txt5(5)) <> "" And Trim(txt5(6)) = "" Then
            If Trim(txt5(5)) <= "1201" Then
               If Trim(txt5(7)) = "" And Trim(txt5(8)) = "" Then
                  bolIsTw12011202 = True
               ElseIf Trim(txt5(7)) = "" And Trim(txt5(8)) <> "" Then
                  If Trim(txt5(8)) >= "000" Then
                     bolIsTw12011202 = True
                  End If
               ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) = "" Then
                  If Trim(txt5(7)) <= "000" Then
                     bolIsTw12011202 = True
                  End If
               ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) <> "" Then
                  If Trim(txt5(7)) <= "000" And Trim(txt5(8)) >= "000" Then
                     bolIsTw12011202 = True
                  End If
               End If
            End If
        ElseIf Trim(txt5(5)) <> "" And Trim(txt5(6)) <> "" Then
            If Trim(txt5(5)) <= "1201" And Trim(txt5(6)) >= "1202" Then
               If Trim(txt5(7)) = "" And Trim(txt5(8)) = "" Then
                  bolIsTw12011202 = True
               ElseIf Trim(txt5(7)) = "" And Trim(txt5(8)) <> "" Then
                  If Trim(txt5(8)) >= "000" Then
                     bolIsTw12011202 = True
                  End If
               ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) = "" Then
                  If Trim(txt5(7)) <= "000" Then
                     bolIsTw12011202 = True
                  End If
               ElseIf Trim(txt5(7)) <> "" And Trim(txt5(8)) <> "" Then
                  If Trim(txt5(7)) <= "000" And Trim(txt5(8)) >= "000" Then
                     bolIsTw12011202 = True
                  End If
               End If
            End If
        End If
    End If
    
   'Add by Morgan 2008/10/2
   bolNPExcept = False
   If OPT2(0).Value = True Then
      If MsgBox("是否包含程序管制的案件性質？", vbYesNo + vbDefaultButton2) = vbNo Then
         bolNPExcept = True
      End If
   End If
    
     '本所期限
     If OPT1(0).Value = True Then
         If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
            Me.txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If Len(Trim(txt1(0))) = 0 And Len(Trim(txt1(1))) = 0 Then
            s = MsgBox("本所期限未輸入", , "USER 輸入錯誤")
            Me.txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         ElseIf Not nickChgRan(txt1(0), txt1(1), "本所期限") Then
            txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         Else

            Me.Enabled = False
            
            'Modify by Morgan 2009/7/22 加判斷專利處或電腦中心人員
            If (Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51") Then
              'Memo by Lydia 2019/11/14 專利處P1的顯示順序，查詢：已收文未發文frm100106_6(一案二申請，)->frm100106_7(已發文未收達)-
                                                                                 '->frm100106_9(代理人信件每日期限管制查詢)
                                                                                 '->bolIsTw12011202 =>frm100106_8(台灣新型修正申復未續辦明細)
                                                       '最後->frm100106_3(以期限管制日查詢by所限)
               'add by nick 2004/07/06 加入判斷要不要顯示  一案二申請
               If OPT2(1).Value = True And (OPT1(0).Value = True Or OPT1(1).Value = True) And (InStr(1, GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0)))), "'P'") <> 0 Or InStr(1, GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0)))), "'CFP'") <> 0) Then
                   Screen.MousePointer = vbHourglass
                   frm100106_6.Show
                   frm100106_6.Process
                   Screen.MousePointer = vbDefault
                   Me.Hide
                   Do
                     DoEvents
                     'add by nickc 2005/08/19
                     'If ArrFormByNick = "" Then Exit Sub
                   Loop Until Not frm100106_6.Visible
                   Unload frm100106_6
                   
                   'Add by Morgan 2005/4/13 已發文未收達
                   Screen.MousePointer = vbHourglass
                   frm100106_7.Show
                   frm100106_7.Process GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0))))
                   Screen.MousePointer = vbDefault
                   Me.Hide
                   Do
                     DoEvents
                     'add by nickc 2005/08/19
                     'If ArrFormByNick = "" Then Exit Sub
                   Loop Until Not frm100106_7.Visible
                   Unload frm100106_7
                   
                   
                   'Add by Toni 2008/8/15  代理人信件每日期限管制查詢
                   Screen.MousePointer = vbHourglass
                   Call frm100106_9.SetParent(Me) 'Add By Sindy 2019/5/30
                   frm100106_9.Show
                   '2012/4/24 MODIFY BY SONIA
                   'frm100106_9.Process
                   frm100106_9.Process GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0))))
                   '2012/4/24 END
                   Screen.MousePointer = vbDefault
                   Me.Hide
                   Do
                     DoEvents
                  Loop Until Not frm100106_9.Visible
                   Unload frm100106_9
                   
               End If
               'add by nickc 2005/08/11 秀台灣新型修正申復未續辦明細
               If bolIsTw12011202 = True Then
                   Screen.MousePointer = vbHourglass
                   frm100106_8.Show
                   frm100106_8.Process
                   Screen.MousePointer = vbDefault
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                   End If
                   Me.Hide
                   Do
                     DoEvents
                     'add by nickc 2005/08/19
                     If ArrFormByNick = "" Then Exit Sub
                   Loop Until Not frm100106_8.Visible
                   Unload frm100106_8
               End If
               
            End If
            
            Screen.MousePointer = vbHourglass
            frm100106_3.Show
            frm100106_3.Enabled = False
            If bolNPExcept = True Then frm100106_3.m_NpCon = strNpSqlOfNoSalesDuty 'Add by Morgan 2008/10/2
            If Not frm100106_3.StrMenu Then
               Me.Enabled = True
               Unload frm100106_3
               'edit by nick 2004/07/14 若查無資料，畫面都不見
               Me.Show
            Else
               frm100106_3.Enabled = True
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbDefault
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
        End If
     '法定期限
     ElseIf Me.OPT1(1).Value Then
         If PUB_CheckKeyInDate(Me.txt2(0)) = -1 Then
            Me.txt2(0).SetFocus
            txt2_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt2(1)) = -1 Then
            Me.txt2(1).SetFocus
            txt2_GotFocus 1
            Exit Sub
         End If
        If Len(Trim(txt2(0))) = 0 And Len(Trim(txt2(1))) = 0 Then
            s = MsgBox("法定期限未輸入", , "USER 輸入錯誤")
            Me.txt2(0).SetFocus
            txt2_GotFocus 0
            Exit Sub
        ElseIf Not nickChgRan(txt2(0), txt2(1), "法定期限") Then
            txt2(0).SetFocus
            txt2_GotFocus 0
            Exit Sub
        Else
            Me.Enabled = False
            
            'Modify by Morgan 2009/7/22 加判斷專利處或電腦中心人員
            If (Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51") Then
              'Memo by Lydia 2019/11/14 專利處P1的顯示順序，查詢：已收文未發文frm100106_6(一案二申請，)->frm100106_7(已發文未收達)
                                                                                 '->bolIsTw12011202 =>frm100106_8(台灣新型修正申復未續辦明細)
                                                       '最後->frm100106_2(以期限管制日查詢by法限)
               'add by nick 2004/07/06 加入判斷要不要顯示  一案二申請
               If OPT2(1).Value = True And (OPT1(0).Value = True Or OPT1(1).Value = True) And (InStr(1, GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0)))), "'P'") <> 0 Or InStr(1, GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0)))), "'CFP'") <> 0) Then
                   Screen.MousePointer = vbHourglass
                   frm100106_6.Show
                   frm100106_6.Process
                   Screen.MousePointer = vbDefault
                   Me.Hide
                   Do
                   DoEvents
                     'add by nickc 2005/08/19
                     'If ArrFormByNick = "" Then Exit Sub
                   Loop Until Not frm100106_6.Visible
                   Unload frm100106_6
                   
                   'Add by Morgan 2005/4/14 已收文未收達
                   Screen.MousePointer = vbHourglass
                   frm100106_7.Show
                   frm100106_7.Process GetAddStr(IIf(txt5(0).Text <> "ALL", txt5(0).Text, GetAllSysKind(txt5(0))))
                   Screen.MousePointer = vbDefault
                   Me.Hide
                   Do
                     DoEvents
                     'add by nickc 2005/08/19
                    ' If ArrFormByNick = "" Then Exit Sub
                   Loop Until Not frm100106_7.Visible
                   Unload frm100106_7
                   
               End If
               'add by nickc 2005/08/11 秀台灣新型修正申復未續辦明細
               If bolIsTw12011202 = True Then
                   Screen.MousePointer = vbHourglass
                   frm100106_8.Show
                   frm100106_8.Process
                   Screen.MousePointer = vbDefault
                   If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                   End If
                   Me.Hide
                   Do
                     DoEvents
                     'add by nickc 2005/08/19
                     If ArrFormByNick = "" Then Exit Sub
                   Loop Until Not frm100106_8.Visible
                   Unload frm100106_8
               End If
               
            End If
            
            Screen.MousePointer = vbHourglass
            frm100106_2.Show
            frm100106_2.Enabled = False
            If bolNPExcept = True Then frm100106_2.m_NpCon = strNpSqlOfNoSalesDuty 'Add by Morgan 2008/10/2
            If Not frm100106_2.StrMenu Then
                Me.Enabled = True
               Unload frm100106_2
               'edit by nick 2004/07/14 若查無資料，畫面都不見
               Me.Show
            Else
               frm100106_2.Enabled = True
                If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
                End If
               Screen.MousePointer = vbDefault
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
        End If
     'Add By Sindy 2013/8/9
     '承辦期限
     ElseIf OPT1(3).Value = True Then
         'Memo by Lydia 2019/11/14 顯示frm100106_3(以期限管制日查詢by所限)
         If PUB_CheckKeyInDate(Me.txt6(0)) = -1 Then
            Me.txt6(0).SetFocus
            txt6_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt6(1)) = -1 Then
            Me.txt6(1).SetFocus
            txt6_GotFocus 1
            Exit Sub
         End If
         If Len(Trim(txt6(0))) = 0 And Len(Trim(txt6(1))) = 0 Then
            s = MsgBox("承辦期限未輸入", , "USER 輸入錯誤")
            Me.txt6(0).SetFocus
            txt6_GotFocus 0
            Exit Sub
         ElseIf Not nickChgRan(txt6(0), txt6(1), "承辦期限") Then
            Me.txt6(0).SetFocus
            txt6_GotFocus 0
            Exit Sub
         Else
            Me.Enabled = False
            Screen.MousePointer = vbHourglass
            frm100106_3.Show
            frm100106_3.Enabled = False
            If bolNPExcept = True Then frm100106_3.m_NpCon = strNpSqlOfNoSalesDuty
            If Not frm100106_3.StrMenu Then
               Me.Enabled = True
               Unload frm100106_3
               '若查無資料，畫面都不見
               Me.Show
            Else
               frm100106_3.Enabled = True
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbDefault
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
        End If
     'Add By Cheng 2002/01/04
     '本所案號
     Else
        If Len(Trim(txt3(0))) = 0 And Len(Trim(txt3(1))) = 0 And Len(Trim(txt3(2))) = 0 And Len(Trim(txt3(3))) = 0 Then
            s = MsgBox("本所案號未輸入", , "USER 輸入錯誤")
            Exit Sub
        Else
                    
            Me.Enabled = False
            
            'Modify by Morgan 2009/7/22 加判斷專利處或電腦中心人員
            If (Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51") Then
              'Memo by Lydia 2019/11/14 專利處P1的顯示順序，bolIsTw12011202 =>frm100106_8(台灣新型修正申復未續辦明細)
                                                       '最後->frm100106_3(以期限管制日查詢by所限)
               'add by nickc 2005/08/11 秀台灣新型修正申復未續辦明細
               If bolIsTw12011202 = True Then
                   Screen.MousePointer = vbHourglass
                   frm100106_8.Show
                   'Add by Morgan 2010/8/18 修正本表單會被 unload 的 bug
                   If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                   End If
                   'end 2010/8/18
                   If frm100106_8.Process = True Then
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        Do
                          DoEvents
                        Loop Until Not frm100106_8.Visible
                   End If
                   Unload frm100106_8
               End If
               
            End If
            
            Screen.MousePointer = vbHourglass
            frm100106_3.Show
            frm100106_3.Enabled = False
            If bolNPExcept = True Then frm100106_3.m_NpCon = strNpSqlOfNoSalesDuty 'Add by Morgan 2008/10/2
            If Not frm100106_3.StrMenu Then
               Unload frm100106_3
               'edit by nick 2004/07/14 若查無資料，畫面都不見
               Me.Show
            Else
               frm100106_3.Enabled = True
                If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
                End If
               Screen.MousePointer = vbDefault
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
        End If
     End If
Case 1
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdCusCase_Click()
    Dim P_OrgPrint As String
    Dim PrinterIndex As Integer, i As Integer
    
On Error GoTo ErrHnd

    OPT2(0).Value = True
    'Modify by Amy 2016/08/16 +產生電子檔-設印表機
    If FormCheck = True Then
        P_OrgPrint = Printer.DeviceName
        PrinterIndex = -1
        For i = 0 To Printers.Count - 1
            If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
                PrinterIndex = i
                Exit For
            End If
        Next i
        If PrinterIndex < 0 Then
            MsgBox "請通知電腦中心安裝PDFCreator !!!"
            Exit Sub
        End If
        P_OrgPrint = PUB_GetOsDefaultPrinter '取得作業系統預設印表機
        PUB_RestorePrinter Printers(PrinterIndex).DeviceName '印表機指到PDFCreator
            
        Screen.MousePointer = vbHourglass
        frm100106_3.m_NpCon = strNpSqlOfNoSalesDuty
        If frm100106_3.StrMenu(False) Then
            frm100106_3.cmdState = 6
            frm100106_3.m_strFACUData = vbNo
            frm100106_3.m_blnSales = IIf(UCase(Left(frm100106_3.GetST05(strUserNum), 1)) = "S", True, False)
            frmPDF.Show
            frmPDF.StartProcess PUB_Getdesktop, "個人客戶二個月之內期限未收文案件" & strSrvDate(2) & ".PDF"
            frm100106_3.PrintData2
            frmPDF.EndtProcess
            Unload frmPDF
            Me.Enabled = True
            Unload frm100106_3
            Me.Show
            MsgBox "列印完成！PDF已產生於桌面！"
        Else
            Unload frm100106_3
            Me.Show
            MsgBox "查無資料~"
        End If
        If Printers(PrinterIndex).DeviceName <> P_OrgPrint Then PUB_RestorePrinter P_OrgPrint
        'end 2016/08/16
        Screen.MousePointer = vbDefault
    End If
    Exit Sub

ErrHnd:
    If Err.Number <> 0 Then
        PUB_RestorePrinter P_OrgPrint
        Screen.MousePointer = vbDefault
        MsgBox "錯誤 : " & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
'add by nickc 2007/01/12
If Len(Trim(Me.txt5(0).Text)) = 0 Then
    Me.txt5(0).Text = "ALL"
End If
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Private Sub Form_Initialize()
bolToEndByNick = False
Me.txt1(0).SetFocus
'92.04.16 nick
cmdState = -1
'add by nickc 2005/08/11
bolIsTw12011202 = False
End Sub

Private Sub Form_Load()
Dim rsA As New ADODB.Recordset
Dim strST05 As String
Dim bolQuery As Boolean
   stST15 = PUB_GetStaffST15(strUserNum, 1) 'Add by Amy 2016/07/13
   
   MoveFormToCenter Me
   OPT1(1).Value = False
   'txt2(0).Enabled = False
   'txt2(1).Enabled = False
   txt5(0) = Systemkind_g
   bolToEndByNick = False
   txt1_GotFocus (1)
   If bolFNation = False Then
       Label1(5).Visible = False
       Label3(2).Visible = False
       Label4(1).Visible = False
       txt4(2).Visible = False
       txt4(3).Visible = False
   End If
   Opt1_Click 0
   
   'Add by Morgan 2003/12/04
   Dim strDeptNo As String
   'Modified by Lydia 2023/12/22
   'strDeptNo = GetDeptNo()
   strDeptNo = Pub_StrUserSt03
   If strDeptNo = "F11" Or strDeptNo = "F23" Or Left(strDeptNo, 1) = "S" Then
      OPT2(0).Value = True
   Else
      OPT2(1).Value = True
   End If
   'End 2003/12/04
   
   'Add by Amy 2016/07/13 增加「個人客戶二個月之內期限未收文案件」鈕
    If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
       bolSpecMan = True
       strSpecCode = "總經理業務工作代理人員"
   '開放專利處部份智權同仁資料給彥葶代為處理
   ElseIf CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = "A8"
   End If
   'end 2016/07/13
   
   'Add by Morgan 2007/9/21
   '外專工程師要控管權限
   strST05 = PUB_GetST05(strUserNum)
   Select Case strST05
      Case "39" '外專工程師中級主管只可查該組
         txt5(14) = PUB_GetStaffST16(strUserNum)
         txt5(14).Locked = True
      Case "40" '外專工程師只可查本人
         txt5(1) = strUserNum
         lbl1(0) = strUserName
         txt5(1).Locked = True
   End Select
   
   'Add By Sindy 2013/8/9 可以查詢承辦期限的人員及電腦中心,才開放其查詢
   bolQuery = False
   strSql = "select distinct decode(st01,null,SR01,st01) from staff_right,staff" & _
            " where upper(sr02)='" & UCase(Me.Name) & "' and sr03='Y' and sr01=st05(+)" & _
            " and decode(st01,null,SR01,st01)='" & strUserNum & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       bolQuery = True
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If bolQuery = True Or Pub_StrUserSt03 = "M51" Then
      OPT1(3).Visible = True
      Frame4.Visible = True
   Else
      OPT1(3).Visible = False
      Frame4.Visible = False
   End If
   '2013/8/9 END
End Sub

'Add by Morgan 2003/12/04
'Mark by Lydia 2023/12/22 改成共用變數
'Private Function GetDeptNo() As String
'
'   Dim strSql As String
'   Dim rsA As New ADODB.Recordset
'
'   strSql = "Select ST03 From STAFF WHERE ST01='" & strUserNum & "'"
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetDeptNo = "" & rsA.Fields(0).Value
'   End If
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'End Function
'end 2023/12/22

Private Sub Form_Unload(Cancel As Integer)
Set frm100106_1 = Nothing
End Sub

Private Sub Opt1_Click(Index As Integer)
On Error Resume Next
   Select Case Index
      Case 0
         txt1(0).Enabled = True
         txt1(1).Enabled = True
         txt2(0).Enabled = False
         txt2(1).Enabled = False
         txt3(0).Enabled = False
         txt3(1).Enabled = False
         txt3(2).Enabled = False
         txt3(3).Enabled = False
         'Add By Sindy 2013/8/9
         txt6(0).Enabled = False
         txt6(1).Enabled = False
         OPT2(0).Enabled = True
         '2013/8/9 END
         If txt1(0).Visible = True Then txt1(0).SetFocus
      Case 1
         txt1(0).Enabled = False
         txt1(1).Enabled = False
         txt2(0).Enabled = True
         txt2(1).Enabled = True
         txt3(0).Enabled = False
         txt3(1).Enabled = False
         txt3(2).Enabled = False
         txt3(3).Enabled = False
         'Add By Sindy 2013/8/9
         txt6(0).Enabled = False
         txt6(1).Enabled = False
         OPT2(0).Enabled = True
         '2013/8/9 END
         If txt2(0).Visible = True Then txt2(0).SetFocus
      Case 2
         txt1(0).Enabled = False
         txt1(1).Enabled = False
         txt2(0).Enabled = False
         txt2(1).Enabled = False
         txt3(0).Enabled = True
         txt3(1).Enabled = True
         txt3(2).Enabled = True
         txt3(3).Enabled = True
         'Add By Sindy 2013/8/9
         txt6(0).Enabled = False
         txt6(1).Enabled = False
         OPT2(0).Enabled = True
         '2013/8/9 END
         If txt3(0).Visible = True Then txt3(0).SetFocus
      'Add By Sindy 2013/8/9
      Case 3
         txt1(0).Enabled = False
         txt1(1).Enabled = False
         txt2(0).Enabled = False
         txt2(1).Enabled = False
         txt3(0).Enabled = False
         txt3(1).Enabled = False
         txt3(2).Enabled = False
         txt3(3).Enabled = False
         txt6(0).Enabled = True
         txt6(1).Enabled = True
         OPT2(0).Enabled = False
         If txt6(0).Visible = True Then txt6(0).SetFocus
      '2013/8/9 END
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
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 1 Then
      If Not nickChgRan(txt1(0), txt1(1), "本所期限") Then
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
      End If
   End If
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
OPT1(0).Value = True
End Sub

Private Sub txt2_GotFocus(Index As Integer)
   txt2(Index).SelStart = 0
   txt2(Index).SelLength = Len(txt2(Index))
   CloseIme
End Sub

Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt2_LostFocus(Index As Integer)
Select Case Index
   Case 0, 1
   If PUB_CheckKeyInDate(Me.txt2(Index)) = -1 Then
      Me.txt2(Index).SetFocus
      txt2_GotFocus Index
      Exit Sub
   End If
   If Index = 1 Then
      If Not nickChgRan(txt2(0), txt2(1), "法定期限") Then
         txt2(0).SetFocus
         txt2_GotFocus (0)
         Exit Sub
      End If
   End If
End Select
End Sub

Private Sub txt2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
OPT1(1).Value = True
End Sub

Private Sub TXT3_GotFocus(Index As Integer)
   txt3(Index).SelStart = 0
   txt3(Index).SelLength = Len(txt3(Index))
   CloseIme
End Sub

Private Sub txt3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TXT3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.OPT1(2).Value = True
End Sub

Private Sub txt4_GotFocus(Index As Integer)
   txt4(Index).SelStart = 0
   txt4(Index).SelLength = Len(txt4(Index))
   CloseIme
End Sub

Private Sub txt4_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt4_LostFocus(Index As Integer)
'Add by Amy 2016/07/13
If Index = 0 And Trim(txt4(0)) <> MsgText(601) Then
    txt4(0) = Left(txt4(0).Text & String(9, "0"), 9)
    txt4(1).Text = Left(txt4(0).Text, 8) & "Z"
'end 2016/07/13
ElseIf Index = 1 Or Index = 3 Then
    If Mid(txt4(Index - 1), 1, 6) <> Mid(txt4(Index), 1, 6) Then
        s = MsgBox("前6碼必須相同！", , "錯誤！")
        txt4(Index - 1).SetFocus
        txt4_GotFocus (Index - 1)
        Exit Sub
    End If
    If RunNick(txt4(Index - 1), txt4(Index)) Then
        txt4(Index - 1).SetFocus
        txt4_GotFocus (Index - 1)
        Exit Sub
    End If
End If
End Sub

Private Sub txt5_GotFocus(Index As Integer)
   txt5(Index).SelStart = 0
   txt5(Index).SelLength = Len(txt5(Index))
   CloseIme
End Sub

Private Sub txt5_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2003/05/23
   Select Case Index
      Case 13 '是否依智權人員排序
         If KeyAscii <> 89 And KeyAscii <> 8 Then
           KeyAscii = 0
         End If
      'Add by Morgan 2007/9/21    2008/2/22加德文組by sonia
      Case 14
         If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") Then
            Beep
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt5_LostFocus(Index As Integer)
   Select Case Index
      Case 1
           lbl1(0) = GetPrjSalesNM(txt5(Index))
           If Trim(txt5(Index)) <> "" Then
              If Trim(lbl1(0).Caption) = "" Then
                  s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                  txt5(Index).SetFocus
                  txt5_GotFocus (Index)
                  Exit Sub
              End If
           End If
      Case 4
           lbl1(1) = GetPrjSalesNM(txt5(Index))
           If Trim(txt5(Index)) <> "" Then
              If Trim(lbl1(1).Caption) = "" Then
                  s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                  txt5(Index).SetFocus
                  txt5_GotFocus (Index)
                  Exit Sub
              End If
           End If
      Case 3, 6, 8, 10, 12
          If RunNick(txt5(Index - 1), txt5(Index)) Then
              txt5(Index - 1).SetFocus
              txt5_GotFocus (Index - 1)
              Exit Sub
          End If
      Case Else
   End Select
End Sub

Private Sub txt6_GotFocus(Index As Integer)
   txt6(Index).SelStart = 0
   txt6(Index).SelLength = Len(txt6(Index))
   CloseIme
End Sub

Private Sub txt6_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt6_LostFocus(Index As Integer)
   If PUB_CheckKeyInDate(Me.txt6(Index)) = -1 Then
      Me.txt6(Index).SetFocus
      txt6_GotFocus Index
      Exit Sub
   End If
   If Index = 1 Then
      If Not nickChgRan(txt6(0), txt6(1), "承辦期限") Then
         txt6(0).SetFocus
         txt6_GotFocus (0)
         Exit Sub
      End If
   End If
End Sub

Private Sub txt6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
OPT1(3).Value = True
End Sub

'Add by Amy 2016/07/13
Private Function FormCheck() As Boolean
    Dim bolCancel As Boolean
    Dim strDate As String, stCU01 As String, stCU02 As String
    
    strDate = Val(DBDATE(DateAdd("m", 2, Format(strSrvDate(1), "####/##/##")))) - 19110000
    FormCheck = False
        
    '本所期限
    If OPT1(0).Value = True Then
        If Trim(txt1(0)) = MsgText(601) Then
            MsgBox "本所期限(起)不可為空！", vbCritical
            txt1(0).SetFocus
            Exit Function
        End If
        If Trim(txt1(1)) = MsgText(601) Then
            MsgBox "本所期限(迄)不可為空！", vbCritical
            txt1(1).SetFocus
            Exit Function
        End If
        '期限不可超過系統日+2個月
        If Val(txt1(0)) > strDate Then
            MsgBox "本所期限(起)不可查超過系統日加2個月！", vbCritical
            txt1(0).SetFocus
            Exit Function
        End If
        If Val(txt1(1)) > strDate Then
            MsgBox "本所期限(迄)不可查超過系統日加2個月！", vbCritical
            txt1(1).SetFocus
            Exit Function
        End If
    End If
    '法定期限
    If OPT1(1).Value = True Then
        If Trim(txt2(0)) = MsgText(601) Then
            MsgBox "法定期限(起)不可為空！", vbCritical
            txt2(0).SetFocus
            Exit Function
        End If
        If Trim(txt2(1)) = MsgText(601) Then
            MsgBox "法定期限(迄)不可為空！", vbCritical
            txt2(1).SetFocus
            Exit Function
        End If
        '期限不可超過系統日+2個月
        If Val(txt2(0)) > strDate Then
            MsgBox "法定期限(起)不可查超過系統日加2個月！", vbCritical
            txt2(0).SetFocus
            Exit Function
        End If
        If Val(txt2(1)) > strDate Then
            MsgBox "法定期限(迄)不可查超過系統日加2個月！", vbCritical
            txt2(1).SetFocus
            Exit Function
        End If
    End If
    '本所案號
    If OPT1(2).Value = True Then
        If Trim(txt3(2)) = MsgText(601) Then txt3(2) = "0"
        If Trim(txt3(3)) = MsgText(601) Then txt3(3) = "00"
        If Trim(txt3(0)) = MsgText(601) Or Trim(txt3(1)) = MsgText(601) Then
            MsgBox "本所案號不可為空！", vbCritical
            If Trim(txt3(0)) = MsgText(601) Then
                txt3(0).SetFocus
            Else
                txt3(1).SetFocus
            End If
            Exit Function
        End If
        If ChkApplyLimit = False And Pub_StrUserSt03 <> "M51" Then
            MsgBox "您無權限查詢此本所案號資料！", vbCritical
            txt3(0).SetFocus
            Exit Function
        End If
    End If
    
    '申請人
    If OPT1(2).Value = False Then
        If Trim(txt4(0)) = MsgText(601) Then
            MsgBox "申請人(起)不可為空！", vbCritical
            txt4(0).SetFocus
            Exit Function
        End If
        If Trim(txt4(1)) = MsgText(601) Then
            MsgBox "申請人(迄)不可為空！", vbCritical
            txt4(1).SetFocus
            Exit Function
        End If
        If Left(txt4(0), 8) <> Left(txt4(1), 8) Then
            MsgBox "申請人編號前8碼必需相同！", vbCritical
            txt4(1).SetFocus
            Exit Function
        End If
        If ChkCusLimit = False Then
            txt4(0).SetFocus
            Exit Function
        End If
    End If
    
    FormCheck = True
    
End Function

'客戶權限控管(自己、同區及離職的帶人主管才能查詢)
Private Function ChkCusLimit() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strCUNo As String, m_CU01 As String, m_CU02 As String
    Dim m_CU13 As String '確認人員
    
    ChkCusLimit = False
    strCUNo = UCase(Left(txt4(0) & "000000000", 9))
    m_CU01 = Left(strCUNo, 8)
    m_CU02 = Right(strCUNo, 1)
            
    '非帶人主管
    If ChkPLimit(1, m_CU01, m_CU02) = False And Pub_StrUserSt03 <> "M51" Then
        m_CU13 = strUserNum
        
        '特殊權限人員
        If bolSpecMan = True Then
            If InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 Then
                m_CU13 = Pub_GetSpecMan("總經理員工編號")
            '開放專利處部份智權同仁資料給彥葶代為處理
            ElseIf InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
                m_CU13 = Pub_GetSpecMan("A7")
            End If
        End If
    
        '自己及同區才可查
        strQ = "Select CU01, CU02, CU13, CU12 From Customer Where CU01='" & m_CU01 & "' AND CU02='" & m_CU02 & "'"
        RsQ.CursorLocation = adUseClient
        RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
        If RsQ.RecordCount > 0 Then
            If "" & RsQ.Fields("CU12") <> stST15 Then
                If InStr(m_CU13, "" & RsQ.Fields("CU13")) = 0 Then
                    MsgBox "您無權限查詢此客戶！", vbCritical
                    Exit Function
                End If
            End If
        End If
        RsQ.Close
    End If
    
    ChkCusLimit = True
End Function

Private Function ChkApplyLimit() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim i As Integer
    Dim m_CU13 As String '確認人員
    Dim strCU13(4) As String, strCU12(4) As String
    
    ChkApplyLimit = False
    Select Case txt3(0)
        Case "CFP", "FCP", "P"  '專利
            strQ = "Select pa26 as S00,S0.cu13 as S01,S0.cu12 as S02,pa27 as S10,S1.cu13 as S11,S1.cu12 as S12,pa28 as S20,S2.cu13 as S21,S2.cu12 as S22" & _
                      ",pa29 as S30,S3.cu13 as S31,S3.cu12 as S32,pa30 as S40,S4.cu13 as S41,S4.cu12 as S42 " & _
                    "From Patent,Customer S0, Customer S1, Customer S2, Customer S3, Customer S4 " & _
                    "Where Pa01= '" & txt3(0) & "' and Pa02='" & txt3(1) & "' and Pa03='" & txt3(2) & "' and Pa04='" & txt3(3) & "' " & _
                    "And Substr(pa26,1,8)=S0.cu01(+) And Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=S0.cu02(+) " & _
                    "And Substr(pa27,1,8)=S1.cu01(+) And Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=S1.cu02(+) " & _
                    "And Substr(pa28,1,8)=S2.cu01(+) And Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=S2.cu02(+) " & _
                    "And Substr(pa29,1,8)=S3.cu01(+) And Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=S3.cu02(+) " & _
                    "And Substr(pa30,1,8)=S4.cu01(+) And Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=S4.cu02(+) "
            
        Case "CFT", "FCT", "T", "TF"  '商標
            strQ = "Select tm23 as S00,S0.cu13 as S01,S0.cu12 as S02,tm78 as S10,S1.cu13 as S11,S1.cu12 as S12,tm79 as S20,S2.cu13 as S21,S2.cu12 as S22" & _
                      ",tm80 as S30,S3.cu13 as S31,S3.cu12 as S32,tm81 as S40,S4.cu13 as S41,S4.cu12 as S42 " & _
                    "From TradeMark,Customer S0, Customer S1, Customer S2, Customer S3, Customer S4 " & _
                    "Where Tm01= '" & txt3(0) & "' and Tm02='" & txt3(1) & "' and Tm03='" & txt3(2) & "' and Tm04='" & txt3(3) & "' " & _
                    "And Substr(tm23,1,8)=S0.cu01(+) And Decode(Substr(tm23,9,1),null,'0',Substr(tm23,9,1))=S0.cu02(+) " & _
                    "And Substr(tm78,1,8)=S1.cu01(+) And Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=S1.cu02(+) " & _
                    "And Substr(tm79,1,8)=S2.cu01(+) And Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=S2.cu02(+) " & _
                    "And Substr(tm80,1,8)=S3.cu01(+) And Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=S3.cu02(+) " & _
                    "And Substr(tm81,1,8)=S4.cu01(+) And Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=S4.cu02(+) "
                    
        Case "CFL", "FCL", "L", "LIN" '法務
            strQ = "Select tm23 as S00,S0.cu13 as S01,S0.cu12 as S02,pa27 as S10,S1.cu13 as S11,S1.cu12 as S12,pa28 as S20,S2.cu13 as S21,S2.cu12 as S22" & _
                      ",pa29 as S30,S3.cu13 as S31,S3.cu12 as S32,pa30 as S40,S4.cu13 as S41,S4.cu12 as S42 " & _
                    "From LawCase,Customer S0, Customer S1, Customer S2, Customer S3, Customer S4 " & _
                    "Where Lc01= '" & txt3(0) & "' and Lc02='" & txt3(1) & "' and Lc03='" & txt3(2) & "' and Lc04='" & txt3(3) & "' " & _
                    "And Substr(lc11,1,8)=S0.cu01(+) And Decode(Substr(lc11,9,1),null,'0',Substr(lc11,9,1))=S0.cu02(+) " & _
                    "And Substr(lc43,1,8)=S1.cu01(+) And Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=S1.cu02(+) " & _
                    "And Substr(lc44,1,8)=S2.cu01(+) And Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=S2.cu02(+) " & _
                    "And Substr(lc45,1,8)=S3.cu01(+) And Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=S3.cu02(+) " & _
                    "And Substr(lc46,1,8)=S4.cu01(+) And Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=S4.cu02(+) "

         Case Else  '服務
            strQ = "Select hc05 as S00,S0.cu13 as S01,S0.cu12 as S02,hc24 as S10,S1.cu13 as S11,S1.cu12 as S12,hc25 as S20,S2.cu13 as S21,S2.cu12 as S22" & _
                      ",hc26 as S30,S3.cu13 as S31,S3.cu12 as S32,hc27 as S40,S4.cu13 as S41,S4.cu12 as S42 " & _
                    "From LawCase,Customer S0, Customer S1, Customer S2, Customer S3, Customer S4 " & _
                    "Where Hc01= '" & txt3(0) & "' and Hc02='" & txt3(1) & "' and Hc03='" & txt3(2) & "' and Hc04='" & txt3(3) & "' " & _
                    "And Substr(hc05,1,8)=S0.cu01(+) And Decode(Substr(hc05,9,1),null,'0',Substr(hc05,9,1))=S0.cu02(+) " & _
                    "And Substr(hc24,1,8)=S1.cu01(+) And Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=S1.cu02(+) " & _
                    "And Substr(hc25,1,8)=S2.cu01(+) And Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=S2.cu02(+) " & _
                    "And Substr(hc26,1,8)=S3.cu01(+) And Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=S3.cu02(+) " & _
                    "And Substr(hc27,1,8)=S4.cu01(+) And Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=S4.cu02(+) "
        
    End Select
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        For i = LBound(strCU13) To UBound(strCU13)
            strCU13(i) = "" & RsQ.Fields("S" & Format(i * 10 + "1", "00"))
            strCU12(i) = "" & RsQ.Fields("S" & Format(i * 10 + "2", "00"))
            If strCU12(i) <> MsgText(601) And strCU12(i) = stST15 Then
                ChkApplyLimit = True
                Exit Function
            End If
            If strCU13(i) <> MsgText(601) Then
                '帶人權限
                If ChkPLimit(2, strCU13(i)) = True Then
                    ChkApplyLimit = True
                    Exit Function
                '特殊權限人員
                ElseIf bolSpecMan = True Then
                    If InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 Then
                        m_CU13 = Pub_GetSpecMan("總經理員工編號")
                    '開放專利處部份智權同仁資料給彥葶代為處理
                    ElseIf InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
                        m_CU13 = Pub_GetSpecMan("A7")
                    End If
                    If InStr(m_CU13, strCU13(i)) > 0 Then
                        ChkApplyLimit = True
                        Exit Function
                    End If
                End If
                If strCU13(i) = strUserNum Then
                    ChkApplyLimit = True
                    Exit Function
                End If
            End If
        Next i
    End If
    RsQ.Close
End Function

'Add by Amy 2016/08/16 判斷印表機
Private Sub ChkPrinter()
    
End Sub
