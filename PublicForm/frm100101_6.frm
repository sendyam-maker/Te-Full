VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件資料及案件進度查詢（顧問基本資料）"
   ClientHeight    =   5490
   ClientLeft      =   530
   ClientTop       =   1280
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   9
      Left            =   6180
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "各項指示"
      Height          =   400
      Index           =   8
      Left            =   1155
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   45
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當5"
      Height          =   400
      Index           =   7
      Left            =   3390
      TabIndex        =   5
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當4"
      Height          =   400
      Index           =   6
      Left            =   2970
      TabIndex        =   6
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當3"
      Height          =   400
      Index           =   5
      Left            =   2550
      TabIndex        =   7
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當2"
      Height          =   400
      Index           =   4
      Left            =   2130
      TabIndex        =   8
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人資料"
      Height          =   400
      Index           =   3
      Left            =   3810
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   45
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   2
      Left            =   5115
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6990
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7890
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   45
      Width           =   800
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1035
      TabIndex        =   56
      Top             =   1456
      Width           =   7875
      VariousPropertyBits=   671105051
      Size            =   "13891;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1035
      TabIndex        =   55
      Top             =   532
      Width           =   2655
      VariousPropertyBits=   671105051
      Size            =   "4683;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   765
      Index           =   1
      Left            =   930
      TabIndex        =   11
      Top             =   3720
      Width           =   7710
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "13600;1349"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   0
      Left            =   1035
      TabIndex        =   10
      Top             =   2700
      Width           =   7605
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "13414;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "當事人5："
      Height          =   255
      Left            =   4590
      TabIndex        =   54
      Top             =   1171
      Width           =   810
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   20
      Left            =   5430
      TabIndex        =   53
      Top             =   1171
      Width           =   3510
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6191;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "當事人4："
      Height          =   255
      Left            =   60
      TabIndex        =   52
      Top             =   1171
      Width           =   810
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   19
      Left            =   1035
      TabIndex        =   51
      Top             =   1171
      Width           =   3510
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6191;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "當事人3："
      Height          =   255
      Left            =   4590
      TabIndex        =   50
      Top             =   863
      Width           =   810
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   18
      Left            =   5430
      TabIndex        =   49
      Top             =   863
      Width           =   3510
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6191;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "當事人2："
      Height          =   255
      Left            =   60
      TabIndex        =   48
      Top             =   863
      Width           =   810
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   17
      Left            =   1035
      TabIndex        =   47
      Top             =   863
      Width           =   3510
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6191;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   16
      Left            =   5040
      TabIndex        =   46
      Top             =   2130
      Width           =   2775
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "4895;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "接洽人："
      Height          =   255
      Left            =   4140
      TabIndex        =   45
      Top             =   2130
      Width           =   720
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   44
      Top             =   3420
      Width           =   7290
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "12859;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   14
      Left            =   7230
      TabIndex        =   43
      Top             =   3105
      Width           =   1380
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2434;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   13
      Left            =   4410
      TabIndex        =   42
      Top             =   3105
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2381;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label81 
      AutoSize        =   -1  'True
      Caption         =   "北所銷卷日："
      Height          =   255
      Left            =   60
      TabIndex        =   41
      Top             =   3105
      Width           =   1080
   End
   Begin VB.Label Label80 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷日："
      Height          =   255
      Left            =   3270
      TabIndex        =   40
      Top             =   3105
      Width           =   1080
   End
   Begin VB.Label Label79 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷員："
      Height          =   255
      Left            =   6090
      TabIndex        =   39
      Top             =   3105
      Width           =   1080
   End
   Begin VB.Label Label78 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷備註："
      Height          =   255
      Left            =   60
      TabIndex        =   38
      Top             =   3420
      Width           =   1260
   End
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   37
      Top             =   4860
      Width           =   945
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   36
      Top             =   4560
      Width           =   945
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   11
      Left            =   1035
      TabIndex        =   35
      Top             =   4560
      Width           =   7485
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "13203;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   12
      Left            =   1035
      TabIndex        =   34
      Top             =   4860
      Width           =   7485
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "13203;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   33
      Top             =   2430
      Width           =   2775
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "4895;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   9
      Left            =   6750
      TabIndex        =   32
      Top             =   1815
      Width           =   975
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   5430
      TabIndex        =   31
      Top             =   1815
      Width           =   975
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   7
      Left            =   5430
      TabIndex        =   30
      Top             =   555
      Width           =   3510
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6191;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   1140
      TabIndex        =   29
      Top             =   3105
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2381;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1035
      TabIndex        =   28
      Top             =   2430
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2355;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1035
      TabIndex        =   27
      Top             =   2130
      Width           =   2775
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "4895;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   2700
      TabIndex        =   26
      Top             =   1815
      Width           =   975
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1380
      TabIndex        =   25
      Top             =   1815
      Width           =   975
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   "--"
      Height          =   180
      Left            =   6510
      TabIndex        =   24
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "--"
      Height          =   180
      Left            =   2460
      TabIndex        =   23
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "當事人1："
      Height          =   255
      Left            =   4590
      TabIndex        =   22
      Top             =   555
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "首次聘任期間："
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   1815
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本次聘任期間："
      Height          =   255
      Left            =   4110
      TabIndex        =   19
      Top             =   1815
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件備註："
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   3780
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   255
      Left            =   60
      TabIndex        =   17
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷："
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   1479
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(Y:是)"
      Height          =   255
      Left            =   2580
      TabIndex        =   14
      Top             =   2430
      Width           =   465
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "閉卷日期："
      Height          =   255
      Left            =   4110
      TabIndex        =   13
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "閉卷原因："
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   2745
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/16 改成Form2.0 ;lbl1(index)、txt1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim StrTag As String, StrTag1 As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer


'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     fnCloseAllFrm100
'add by nickc 2005/05/30
Case 2
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_3.Show
     frm100108_3.Tag = txt1(2).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'add by nickc 2006/07/12
Case 3
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag1 ' StrTag    傳申請人代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Add By Sindy 2011/1/17
Case 4
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(lbl1(17).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(lbl1(17).Caption, 9) '當事人2
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 5
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(lbl1(18).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(lbl1(18).Caption, 9) '當事人3
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 6
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(lbl1(19).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(lbl1(19).Caption, 9) '當事人4
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 7
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(lbl1(20).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(lbl1(20).Caption, 9) '當事人5
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'2011/1/17 End
'Added by Lydia 2016/11/23
Case 8 '各項指示
    'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
    If PUB_CheckFormExist("frm12040159") Then
        MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
        Exit Sub
    End If
    'end 2020/05/05
    
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm12040159.SetParent "Q", Trim(Replace(txt1(2), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 9 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(2)
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Sub StrMenu()
Dim strSql  As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'edit by nickc 2006/07/12
'Dim strArr(T_HC) As String, i As Integer, StrOk(12) As String, StrOkTxt(1) As String
Dim strArr() As String, i As Integer, StrOk(12) As String, StrOkTxt(1) As String
'add  by nickc 2006/07/12
ReDim strArr(TF_HC) As String
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
Dim tmp01 As String, tmp02 As String
'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'End by Toni 20080926

Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
If Left(Me.Tag, 1) = "N" Then
   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSql = Me.Tag
End If
txt1(2).Text = Me.Tag
Str01 = SystemNumber(strSql, 1)
Str02 = SystemNumber(strSql, 2)
Str03 = SystemNumber(strSql, 3)
Str04 = SystemNumber(strSql, 4)

' 使用者沒有權限
'add by Toni 20080926 控制跨部門權限訊息 for 顧問基本資料查詢
'2008/10/2 modify by sonia
'If IsUserHasRightOfSystem(strUserNum, Str01) = False Then
'   If IsUserHasRightOfFunction("frm100101_1", strCrossDept, False) = False Then
'      strTit = "檢核資料"
'      strMsg = "您沒有使用該系統類別的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   End If
'End If
If CheckSR09(strUserNum, Str01, "Y", , Str01, Str02, Str03, Str04) = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'2008/10/2 end
'End 20080926

pub_QL05 = ";本所案號：" & Str01 & "-" & Str02 & "-" & Str03 & "-" & Str04 & _
           "(基本資料)" 'Add By Sindy 2025/8/7

'Add By Sindy 2011/1/17
cmdOK(4).Visible = False
cmdOK(5).Visible = False
cmdOK(6).Visible = False
cmdOK(7).Visible = False
'2011/1/17 End

'欲搜尋的SQL字串
strSql = "SELECT * FROM HIRECASE WHERE HC01='" & Str01 & "' AND HC02='" & Str02 & "' AND HC03='" & Str03 & "' AND HC04='" & Str04 & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/7
   'For i = 0 To 17
   For i = 0 To (TF_HC - 1) 'edit by nickc 2006/07/12 (T_HC - 1)
      Select Case i
      Case 9, 13, 14, 16, 17
           If IsNull(adoRecordset.Fields(i)) Then
               strArr(i + 1) = ""
           Else
               strArr(i + 1) = str(adoRecordset.Fields(i))
           End If
      Case Else
           If IsNull(adoRecordset.Fields(i)) Then
                strArr(i + 1) = ""
           Else
                strArr(i + 1) = adoRecordset.Fields(i)
           End If
      End Select
      DoEvents
   Next i
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
   ShowNoData
   Screen.MousePointer = vbDefault
       '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Dim strTemp As String    '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
intK = 18
'For i = 0 To 18
For i = 1 To TF_HC 'edit by nickc 2006/07/12 T_HC
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4)
         txt1(2) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4) 'Add By Sindy 2013/1/31
         strSql = "SELECT CP05,CP53,CP54 FROM CASEPROGRESS WHERE CP01='" & strArr(1) & "' AND CP02='" & strArr(2) & "' AND CP03='" & strArr(3) & "' AND CP04='" & strArr(4) & "' AND CP10='0'  ORDER BY CP05 "
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveFirst
            If Not IsNull(adoRecordset.Fields(1)) Then
                'edit by nick 2004/11/04
                'StrOk(2) = adoRecordset.Fields(1)
                StrOk(2) = ChangeWStringToTDateString("" & adoRecordset.Fields(1))
            Else
                StrOk(2) = ""
            End If
            If Not IsNull(adoRecordset.Fields(2)) Then
                'edit by nick 2004/11/04
                'StrOk(3) = adoRecordset.Fields(2)
                StrOk(3) = ChangeWStringToTDateString("" & adoRecordset.Fields(2))
            Else
                StrOk(3) = ""
            End If
            adoRecordset.MoveLast
            If Not IsNull(adoRecordset.Fields(1)) Then
                'edit by nick 2004/11/04
                'StrOk(8) = adoRecordset.Fields(1)
                StrOk(8) = ChangeWStringToTDateString("" & adoRecordset.Fields(1))
            Else
                StrOk(8) = ""
            End If
            If Not IsNull(adoRecordset.Fields(2)) Then
                'edit by nick 2004/11/04
                'StrOk(9) = adoRecordset.Fields(2)
                StrOk(9) = ChangeWStringToTDateString("" & adoRecordset.Fields(2))
            Else
                StrOk(9) = ""
            End If
         Else
            StrOk(2) = ""
            StrOk(3) = ""
            StrOk(8) = ""
            StrOk(9) = ""
         End If
         CheckOC
    Case 6
         StrOk(1) = strArr(i)
         txt1(3) = strArr(i) 'Add By Sindy 2013/1/31
    Case 7
         StrOk(4) = strArr(i)
    Case 9
         StrOk(5) = strArr(i)
    Case 19
         'edit by nickc 2006/07/12
         'StrOk(6) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(6) = ""
         Else
             StrOk(6) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 5
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Len(strArr(i)) = 9 Then
'              strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'         Else
'              strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(7) = strArr(i) + ""
'                     Else
'                          StrOk(7) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(7) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(7) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
        End If
        If tmp02 <> "" Then
            StrOk(7) = strArr(i) + "  " + tmp02
            'Add by Morgan 2004/1/14
            lbl1(7).ForeColor = vbBlack
         Else
            StrOk(7) = ""
            'Add by Morgan 2004/1/14
            lbl1(7).ForeColor = vbRed
             StrOk(7) = strArr(i)
         End If
         CheckOC
    Case 9
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(16) = strArr(i) + ""
              Else
                  StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
         Else
              StrOk(16) = ""
         End If
         CheckOC
    Case 12
         StrOkTxt(1) = strArr(i)
    Case 11
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                     StrOkTxt(0) = strArr(i) + ""
             Else
                     StrOkTxt(0) = strArr(i) + "  " + adoRecordset.Fields(0)
             End If
         Else
             StrOkTxt(0) = ""
         End If
         CheckOC
    Case 13
         'edit by nick 2004/10/05
         'StrOk(11) = GetPrjSalesNM(strArr(i)) & " " & strArr(14) & " " & strArr(15)
         StrOk(11) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(14))) & " " & Format(strArr(15), "##:##")
    Case 16
         'edit by nick 2004/10/05
         'StrOk(12) = GetPrjSalesNM(strArr(i)) & " " & strArr(17) & " " & strArr(18)
         StrOk(12) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(17))) & " " & Format(strArr(18), "##:##")
    'add by nickc 2006/07/12
    Case 20
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             lbl1(13) = ""
         Else
             lbl1(13) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 21
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               lbl1(14) = strArr(i) + ""
            Else
               lbl1(14) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            lbl1(14) = ""
         End If
         CheckOC
    Case 22
         lbl1(15) = strArr(i)
    'Add by Morgan 2008/8/4
    Case 23
         lbl1(16) = PUB_GetContact(strArr(5), strArr(i))
    'Add By Sindy 2011/1/17
    Case 24
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdOK(4).Visible = True
        End If
        If tmp02 <> "" Then
            lbl1(17).Caption = strArr(i) + "  " + tmp02
            lbl1(17).ForeColor = vbBlack
         Else
            lbl1(17).Caption = strArr(i)
            lbl1(17).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/1/17
    Case 25
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdOK(5).Visible = True
        End If
        If tmp02 <> "" Then
            lbl1(18).Caption = strArr(i) + "  " + tmp02
            lbl1(18).ForeColor = vbBlack
         Else
            lbl1(18).Caption = strArr(i)
            lbl1(18).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/1/17
    Case 26
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdOK(6).Visible = True
        End If
        If tmp02 <> "" Then
            lbl1(19).Caption = strArr(i) + "  " + tmp02
            lbl1(19).ForeColor = vbBlack
         Else
            lbl1(19).Caption = strArr(i)
            lbl1(19).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/1/17
    Case 27
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdOK(7).Visible = True
        End If
        If tmp02 <> "" Then
            lbl1(20).Caption = strArr(i) + "  " + tmp02
            lbl1(20).ForeColor = vbBlack
         Else
            lbl1(20).Caption = strArr(i)
            lbl1(20).ForeColor = vbRed
         End If
         CheckOC
    Case Else
    End Select
    DoEvents
Next i
For i = 0 To 12
   If i <> 0 And i <> 1 Then 'Add By Sindy 2013/1/31 +if
      lbl1(i) = StrOk(i)
   End If
Next i
txt1(0) = StrOkTxt(0)
txt1(1) = StrOkTxt(1)
'add by nickc 2006/07/12
StrTag1 = strArr(5)
'add by nickc 2005/05/30  檢查有無分割或相關卷號
     cmdOK(2).Visible = ChkDataByCR(txt1(2).Text)
End Sub

'edit by nickc 2005/05/30 改成與我們現在共用相同
'Private Sub cmdRef_Click()
'    Dim stTmp As String
'    stTmp = Right(Space(2) & txt1(2), 15)
'    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me

'92.04.16 nick
cmdState = -1

'Added by Lydia 2020/05/05 各項指示：顯示按鈕
If strSrvDate(1) >= 各項指示啟用日 Then
   cmdOK(8).Visible = True
Else
   cmdOK(8).Visible = False
End If
'end 2020/05/05
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100101_6 = Nothing
End Sub


'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub
