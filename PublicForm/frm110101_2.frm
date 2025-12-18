VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110101_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "解除期限"
   ClientHeight    =   6024
   ClientLeft      =   120
   ClientTop       =   348
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6024
   ScaleWidth      =   9312
   Begin VB.CheckBox ChkOutlook 
      Caption         =   "出OutLook草稿"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   67
      Top             =   30
      Width           =   1650
   End
   Begin VB.TextBox txtSalesNo 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   66
      Top             =   1725
      Width           =   855
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視回覆單"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2256
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "退回智權"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   3
      Left            =   3348
      TabIndex        =   64
      Top             =   0
      Width           =   1125
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   11
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3630
      Width           =   3090
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   3270
      Width           =   1290
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   8
      Left            =   6525
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2250
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   9
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   13
      Top             =   5640
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   10
      Left            =   6135
      MaxLength       =   1
      TabIndex        =   14
      Top             =   5655
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   7
      Left            =   5652
      MaxLength       =   7
      TabIndex        =   12
      Top             =   5355
      Width           =   1212
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   6
      Left            =   1476
      MaxLength       =   7
      TabIndex        =   11
      Top             =   5340
      Width           =   1212
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   2
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2940
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   5
      Left            =   6132
      MaxLength       =   1
      TabIndex        =   10
      Top             =   5055
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   4
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   9
      Top             =   5040
      Width           =   492
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8430
      TabIndex        =   17
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6450
      TabIndex        =   15
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7275
      TabIndex        =   16
      Top             =   0
      Width           =   1125
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2250
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   1
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2940
      Width           =   492
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   420
      Left            =   1080
      TabIndex        =   8
      Top             =   4605
      Width           =   7350
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "12965;741"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboReason 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   2565
      Width           =   6975
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14420;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboMemo 
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   3930
      Width           =   7335
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14420;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   270
      Left            =   1110
      TabIndex        =   18
      Top             =   675
      Width           =   7395
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13044;476"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   115
      Left            =   120
      TabIndex        =   63
      Top             =   4320
      Width           =   8220
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人："
      Height          =   180
      Index           =   1
      Left            =   132
      TabIndex        =   62
      Top             =   1290
      Width           =   930
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   1125
      TabIndex        =   61
      Top             =   1290
      Width           =   7365
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   9
      Left            =   5280
      TabIndex        =   60
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "審定號數："
      Height          =   180
      Left            =   4350
      TabIndex        =   59
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號："
      Height          =   180
      Left            =   132
      TabIndex        =   58
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "CF代理人："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   57
      Top             =   3330
      Width           =   930
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   2385
      TabIndex        =   56
      Top             =   3330
      Width           =   5970
      VariousPropertyBits=   27
      Size            =   "10530;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "是否管制下次期限：             (Y：管制)"
      Height          =   180
      Left            =   4905
      TabIndex        =   55
      Top             =   2295
      Width           =   2985
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函：            （N：不印）"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   54
      Top             =   5685
      Width           =   3000
   End
   Begin VB.Label Label13 
      Caption         =   "是否修改通知函內容：            （Y：Word）"
      Height          =   180
      Left            =   4335
      TabIndex        =   53
      Top             =   5685
      Width           =   3495
   End
   Begin VB.Label lblNation 
      Height          =   180
      Left            =   5820
      TabIndex        =   52
      Top             =   1035
      Width           =   2685
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   8
      Left            =   5310
      TabIndex        =   51
      Top             =   1035
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "申請國家："
      Height          =   180
      Left            =   4335
      TabIndex        =   50
      Top             =   1035
      Width           =   975
   End
   Begin VB.Label lblChildCase 
      BackColor       =   &H8000000B&
      Caption         =   "有子案或相關卷號"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2745
      TabIndex        =   49
      Top             =   2295
      Width           =   1575
   End
   Begin MSForms.Label lblPromoter 
      Height          =   180
      Left            =   1920
      TabIndex        =   48
      Top             =   1755
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   180
      Left            =   1935
      TabIndex        =   47
      Top             =   1035
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNextProgress 
      Height          =   180
      Left            =   1575
      TabIndex        =   46
      Top             =   2040
      Width           =   2655
   End
   Begin MSForms.Label lblSales 
      Height          =   180
      Left            =   6180
      TabIndex        =   45
      Top             =   1755
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      Caption         =   "下次法定期限："
      Height          =   180
      Index           =   4
      Left            =   4335
      TabIndex        =   44
      Top             =   5385
      Width           =   1335
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "下次本所期限："
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   43
      Top             =   5385
      Width           =   1260
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   4092
      TabIndex        =   42
      Top             =   2988
      Width           =   5000
      VariousPropertyBits=   27
      Caption         =   "後續准駁簡單報告：            （Y：核准以及C類來函簡單報告）"
      Size            =   "8819;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   6
      Left            =   960
      TabIndex        =   41
      Top             =   1755
      Width           =   855
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   5
      Left            =   5325
      TabIndex        =   40
      Top             =   1515
      Width           =   2895
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   4
      Left            =   1125
      TabIndex        =   39
      Top             =   1515
      Width           =   2895
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   3
      Left            =   1155
      TabIndex        =   38
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   2
      Left            =   975
      TabIndex        =   37
      Top             =   1035
      Width           =   855
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   1
      Left            =   5292
      TabIndex        =   36
      Top             =   384
      Width           =   3012
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   0
      Left            =   1092
      TabIndex        =   35
      Top             =   384
      Width           =   3012
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   34
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "是否修改指示信內容：            （Y：Word）"
      Height          =   180
      Left            =   4320
      TabIndex        =   33
      Top             =   5085
      Width           =   3495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：            （N：不印）"
      Height          =   180
      Left            =   135
      TabIndex        =   32
      Top             =   5085
      Width           =   3000
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   31
      Top             =   4650
      Width           =   900
   End
   Begin VB.Label Label21 
      Caption         =   "案件備註："
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   30
      Top             =   3975
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "機關文號："
      Height          =   180
      Index           =   2
      Left            =   4332
      TabIndex        =   29
      Top             =   384
      Width           =   972
   End
   Begin VB.Label Label14 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   1755
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   27
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   26
      Top             =   384
      Width           =   972
   End
   Begin VB.Label Label8 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   4350
      TabIndex        =   25
      Top             =   1755
      Width           =   900
   End
   Begin VB.Label Label9 
      Caption         =   "下一程序："
      Height          =   180
      Left            =   135
      TabIndex        =   24
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   23
      Top             =   1515
      Width           =   930
   End
   Begin VB.Label Label21 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   2
      Left            =   4335
      TabIndex        =   22
      Top             =   1515
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "解除期限日期："
      Height          =   180
      Left            =   135
      TabIndex        =   21
      Top             =   2295
      Width           =   1260
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "解除期限原因："
      Height          =   180
      Left            =   135
      TabIndex        =   20
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷：            （Y：閉卷）"
      Height          =   180
      Left            =   135
      TabIndex        =   19
      Top             =   2985
      Width           =   2460
   End
End
Attribute VB_Name = "frm110101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/5 改成Form2.0 (txtCP64,cboReason,cboMemo,cboCaseName,lblPetitionName,Label2,lblPromoter,lblSales)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'2010/8/3 日期欄已修改 by sonia
Option Explicit
 
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
'intWhere 國內,國外_CF,國外_FC
Dim intCaseKind As Integer, intWhere As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String, SCp(1 To 79) As String
'intLeaveKind 離開時，是 0:結束 1:回上一畫面 2:確定
Dim intLeaveKind As Integer
'edit by nickc 2007/02/02 不用 dll 了
'Dim obj011 As New prjTaieDll011.cls011
Dim strSql As String
'看卷是否已經閉卷
Dim BolFileClose As Boolean
'看卷是否確定閉卷
Dim BolFileCloseOk As Boolean
'下次本所期限，下次法定期限
Dim Nextdate1 As String, Nextdate2 As String
Dim bolIsChina As Boolean
'Add by Morgan 2004/8/6
Dim stNP09 As String
Dim m_bolFMP As Boolean 'Add by Morgan 2010/2/3
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/09 是否為寰華案
Public mPrev01 As Form 'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
Dim strNP01 As String
Dim strNP07 As String
Dim strNP22 As String
Dim m_boleOrderLetter As Boolean 'Added by Morgan 2015/11/3 指示信電子化
'Added by Lydia 2018/03/16 是否已觸發 Form Active 事件
Dim bolActive As Boolean
Dim m_strAF01 As String, m_strChildAF01 As String 'Added by Morgan 2018/8/22
Dim bolHas907 As Boolean, bolShow060318 As Boolean 'Add by Amy 2018/11/27
Dim strTi01 As String 'Add by Amy 2022/06/20
Dim m_strSaveFiles As String, m_strSaveFilesCP09 As String 'Add by Amy 2022/09/05 附件/附件總收文號
Public Pre_ProState As String 'Add by Amy 2023/02/14 登入之系統
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知
'Add by Amy 2025/06/02
Public intFCState As Integer '0-智權/1-FC商標/2-FC專利 發起之結案單
Dim strF0301 As String, bolInvoice As Boolean, strClose As String '結案單號/是否開請款單輸入/閉卷
Dim strOutLookType As String 'Add by Amy 2025/07/10 "0":寄 工程師+承辦 / "1":寄 承辦

Private Sub cboMemo_GotFocus()
'edit by nickc 2007/06/06 切換輸入法改用API
'Me.cboMemo.IMEMode = 1
OpenIme
End Sub

Private Sub cboReason_LostFocus()
'Mark by Amy 2025/06/16 不使用
'Dim ii As Integer
'Dim blnInput As Boolean
'
'    'Add By Cheng 2003/04/16
'    If Me.cboReason.Text <> "" Then
'        blnInput = False
'        For ii = 0 To Me.cboReason.ListCount - 1
'            If Left(Me.cboReason.Text, 2) = Left(Me.cboReason.List(ii), 2) Then
'                Me.cboReason.ListIndex = ii
'                blnInput = True
'            End If
'        Next ii
'        If blnInput = False Then
'            Me.cboReason.Text = ""
'        End If
'    End If
End Sub

'Add by Amy 2022/09/05 避免智權人員放空白回覆單 ex:T-184183、CFT-015329
Private Sub cmdFile_Click()
   Dim ii As Integer, jj As Integer, arrData As Variant
   Dim strMsg As String 'Add by Amy 2025/06/02
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2025/09/23 +if bug-T延展結案(無Flow003資料),有回覆單資料會無法顯示
   If (field(1) = "T" Or field(1) = "TF") And (strNP07 = "102" Or strNP07 = "109" Or strNP07 = "716") And strNP22 <> "" Then
      Call Pub_OpenReplayPDFOrMsg(intFCState, Me, cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4), strNP22, m_strSaveFiles, "", strMsg, True)
   Else
      Call Pub_OpenReplayPDFOrMsg(intFCState, Me, cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4), strF0301, m_strSaveFiles, "", strMsg, True)
   End If
   'end 2025/09/23
   If strMsg <> "" Then
      MsgBox strMsg
   End If
   Screen.MousePointer = vbDefault
   'end 2025/06/02
 
   'Mark by Amy 2025/06/02 改至Pub_OpenReplayPDFOrMsg,以下不執行
'    Screen.MousePointer = vbHourglass
'    frm100101_L.m_strKey = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
'    frm100101_L.SetParent Me
'    If frm100101_L.QueryData = True Then
'        arrData = Split(m_strSaveFiles, ":")
'        For jj = 0 To UBound(arrData)
'            For ii = frm100101_L.GRD1.Rows - 1 To 1 Step -1
'               If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), arrData(jj)) > 0 Then
'                  Exit For
'               End If
'            Next ii
'            If ii > 0 Then
'               Call frm100101_L.FrmCallOpenFile(ii, IIf(UBound(arrData) = jj, True, False))
'               If UBound(arrData) = jj Then
'                  frm100101_L.Show
'                  Me.Hide
'               End If
'            Else
'               Unload frm100101_L
'               Screen.MousePointer = vbDefault
'               MsgBox "有回覆單電子檔:" & m_strSaveFiles
'               Exit Sub
'            End If
'        Next jj
'    Else
'        Unload frm100101_L
'    End If
'    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim stLetter As String 'Add by Morgan 2004/9/27
   Dim strTmp As String, i As Integer, bolChk As Boolean
   'Add By Cheng 2002/07/31
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim stET02 As String
   'Add by Morgan 2007/1/22
   Dim bolMail As Boolean '是否通知智權人員
   Dim stCP13 As String '智權人員編號
   Dim stCP09 As String '收文號
   Dim strUpdDate As String, strUpdTime As String, strF0308 As String, strF0309 As String 'Add By Sindy 2015/1/14
   'Added by Lydia 2017/01/25 CFP案EPC母案解除期限一併產生子案指示信
   Dim strChildList As String '子案B類收文號
   Dim tmpArr As Variant, tmpNo As String
   Dim mCCase As String 'Added by Lydia 2017/05/08 記錄子案案號(避免Trigger一併更新子案)
   'Add by Amy 2018/06/25
   Dim strSubject As String, strTo As String '主旨/Mail to
   Dim strLetterJudge As String, strChildCP45 As String, strCP10 As String '指示信判發人員/子案彼所案號/案件性質 'Added by Morgan 2018/8/16
   Dim strMsg As String 'Add by Amy 2018/08/30
   Dim bolCancel As Boolean 'Added by Morgan 2018/9/20
   Dim strContent As String 'Add by Amy 2018/10/08 信內容
   Dim strTM22New As String 'add by sonia 2020/5/27
   Dim stCP14 As String 'Add by Amy 2020/09/07
   Dim bolNPScalable As Boolean 'Add by Amy 2020/11/16 有延展或延展(英國)同時未解除期限
   Dim strOldF0308 As String, strCmd(1) As String 'Add by Amy 2021/06/23 記錄原程序人員(CFT/CFC/S可職代操作)/更新語法
   Dim strPA146 As String 'Add by Amy 2025/08/07
   Dim strNotPay As String, strCCD08 As String, intI As Integer 'Add by Amy 2025/08/08
   Dim bolMailF0202_3 As Boolean  'Add by Amy 2025/08/19 解除期限後是否寄信給補看人員
   Dim bolOpen21H0Ok As Boolean 'Add by Amy 2025/10/20
    
    'Add by Amy 2023/02/14 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
    If Pre_ProState = "T" And (Index = 0 Or Index = 3) Then
        If Pub_ChkLock(0, Me.Name, "C", Me.Caption, cp(1) & cp(2) & cp(3) & cp(4)) = False Then
            Exit Sub
        End If
    End If
    
    Select Case Index
       Case 0
          'Add by Amy 2025/10/28 有請款項目 且請款單輸入已開啟需關閉
          If bolInvoice = True Then
            If PUB_CheckFormExist("Frmacc21h0") = True Or PUB_CheckFormExist("Frmacc21h01") = True Then
               MsgBox "請款單輸入已開啟請關閉", vbExclamation
               Exit Sub
            End If
          End If
          'end 2025/10/28
          'Modify by Amy 2024/10/04 鎖住[退回智權]鈕不可按,避免結案ti06又設Y
          If cmdOK(3).Visible = True Then cmdOK(3).Enabled = False
          'Add by Amy 2023/01/31 智權人員可輸需檢查必輸
          If txtSalesNo.Locked = False Then
             If Trim(txtSalesNo) = MsgText(601) Then
                MsgBox "智權人員不可空白 !", vbCritical
                If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
                txtSalesNo.SetFocus
                txtSalesNo_GotFocus
                Exit Sub
             End If
          End If
          If txtCaseField(0) = "" Then
             MsgBox "解除期限日期不可空白 !", vbCritical
             If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
             txtCaseField(0).SetFocus
             txtCaseField_GotFocus (0)
             Exit Sub
          End If
          'add by nickc 2008/05/30 再次檢查 解除期限原因
          'cboReason_LostFocus 'Mark by Amy 2025/06/16
            'Modify By Cheng 2003/04/15
'          If txtCaseField(8) = "" Then
          If Me.cboReason.Text = "" Then
             MsgBox "解除期限原因不可空白 !", vbCritical
             If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
'             txtCaseField(8).SetFocus
                Me.cboReason.SetFocus
             Exit Sub
          End If
          
          'Add by Morgan 2005/1/4國內及大陸案年費若閉卷欄位未輸入則提醒預設要輸入
          'Modified by Morgan 2023/12/4 +txtCaseField(1).Enabled(因有增加控制有未發文程序時不可閉卷)
          If field(1) = "P" And lblCaseField(3) = "605" And txtCaseField(1).Enabled Then
            If txtCaseField(1) = "" Then
                  If MsgBox("是否閉卷？", vbYesNo + vbDefaultButton1, "國內及大陸案年費解除期限提醒") = vbYes Then
                     txtCaseField(1).SetFocus
                     Exit Sub
                  End If
            End If
            
            'Added by Morgan 2019/11/26 年費閉卷催審期限提醒 Ex:P-121561--茹曣
            If txtCaseField(1) = "Y" Then
               strExc(0) = "select cpm03,cpm04 from nextprogress,caseprogress,casepropertymap where np02='" & field(1) & "' and np03='" & field(2) & "' and np04='" & field(3) & "' and np05='" & field(4) & "' and np06 is null and np07='411' and cp09(+)=np01 and cpm01(+)=cp01 and cpm02(+)=cp10"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If MsgBox("下一程序尚有「" & IIf(field(9) = "000", RsTemp("cpm03"), RsTemp("cpm04")) & "」催審期限，請確認是否閉卷？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
                     txtCaseField(1).SetFocus
                     Exit Sub
                  End If
               End If
            End If
            'end 2019/11/6
          End If
          'Add by Amy 2020/11/16 CFP 延展或延展(英國)沒有同時未解除期限者彈提醒,若有同時未解除期限且畫面閉卷=Y,不可存檔
          'Modified by Morgan 2023/12/4 +txtCaseField(1).Enabled(因有增加控制有未發文程序時不可閉卷)
          If field(1) = "CFP" And (lblCaseField(3) = "607" Or lblCaseField(3) = "613") And txtCaseField(1).Enabled Then
            strMsg = ""
            bolNPScalable = ChkNPScalable(lblCaseField(8), field(1), field(2), field(3), field(4), lblCaseField(3), strMsg)
            If bolNPScalable = False And txtCaseField(1) = "" Then
                If MsgBox("是否閉卷？", vbYesNo + vbDefaultButton1, "專利延展費解除期限提醒") = vbYes Then
                   txtCaseField(1).SetFocus
                   Exit Sub
                End If
            ElseIf bolNPScalable = True And txtCaseField(1) = "Y" Then
                MsgBox "下一程序尚有「" & strMsg & "」期限，不可閉卷"
                txtCaseField(1).SetFocus
                Exit Sub
            End If
          End If
          '2006/11/21 ADD BY SONIA 商標延展案若閉卷欄位未輸入則提醒預設要輸入
          'modify by sonia 2020/5/27 解除期限原因19他所延展未變更代理人時就不問
          'If (field(1) = "T" Or field(1) = "TF" Or field(1) = "FCT" Or field(1) = "CFT") And lblCaseField(3) = "102" And txtCaseField(1) = ""  Then
          'Modify by Amy 2020/11/16 +案件性質110,並判斷CFT 延展或延展(英國)沒有同時未解除期限者彈提醒,若有同時未解除期限且畫面閉卷=Y,不可存檔
          'If (field(1) = "T" Or field(1) = "TF" Or field(1) = "FCT" Or field(1) = "CFT") And lblCaseField(3) = "102" And txtCaseField(1) = "" And Left(cboReason.Text, 2) <> "19" Then
          'Modified by Morgan 2023/12/4 +txtCaseField(1).Enabled(因有增加控制有未發文程序時不可閉卷)
          If (field(1) = "T" Or field(1) = "TF" Or field(1) = "FCT" Or field(1) = "CFT") And (lblCaseField(3) = "102" Or lblCaseField(3) = "110") And Left(cboReason.Text, 2) <> "19" And txtCaseField(1).Enabled Then
            strMsg = ""
            bolNPScalable = ChkNPScalable(lblCaseField(8), field(1), field(2), field(3), field(4), lblCaseField(3), strMsg)
            If bolNPScalable = False And txtCaseField(1) = "" Then
                If MsgBox("是否閉卷？", vbYesNo + vbDefaultButton1, "商標延展案解除期限提醒") = vbYes Then
                   If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
                   txtCaseField(1).SetFocus
                   Exit Sub
                End If
            ElseIf bolNPScalable = True And txtCaseField(1) = "Y" Then
                MsgBox "下一程序尚有「" & strMsg & "」期限，不可閉卷"
                If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
                txtCaseField(1).SetFocus
                Exit Sub
            End If
          End If
          'end 2020/11/16
          '2006/11/21 END
          'add by sonia 2016/3/28 CFT異議案不可閉卷
          If field(1) = "CFT" And lblCaseField(3) = "601" And txtCaseField(1) <> "" Then
             MsgBox "CFT之異議期限解除，不可閉卷！", vbCritical
             If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
             txtCaseField(1).SetFocus
             Exit Sub
          End If
          'end 2016/3/28
          'add by sonia 2020/5/27
          If Left(cboReason.Text, 2) = "19" And ((field(1) <> "T" And field(1) <> "FCT") Or lblCaseField(3) <> "102") Then
             MsgBox "非T或FCT延展期限，不可選擇解除期限原因19 他所延展未變更代理人！", vbCritical
             If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
             cboReason.SetFocus
             Exit Sub
          End If
          
          '解除期限原因19他所延展未變更代理人時不可閉卷
          If Left(cboReason.Text, 2) = "19" And txtCaseField(1) <> "" Then
             MsgBox "解除期限原因19，他所延展未變更代理人，不可閉卷！", vbCritical
             If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
             txtCaseField(1).SetFocus
             Exit Sub
          End If
          'end 2020/5/27
          
          'Added by Morgan 2018/9/20
          'CF案要輸代理人,否則無法產生指示信 Ex.CFP-21588-1
          If Combo2.Enabled Then
            Combo2_Validate bolCancel
            If bolCancel Then
               If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
               Combo2.SetFocus
               Exit Sub
            End If
          End If
          'end 2018/9/20
         
          'Modified by Lydia 2019/06/21 改成模組
          'CheckFCPDualCase 'Added by Morgan 2015/5/18
          'Modified by Morgan 2024/12/4 要用下一程序的性質,cp(10)是相關收文號的性質 cp(10)->lblCaseField(3)
          Call Pub_ChkFCPDualCaseBYcancel(field(1), field(2), field(3), field(4), field(8), field(9), lblCaseField(3))
          
          'Add by Morgan 2007/1/22 若有A類收文未發文時提醒不閉卷
          bolMail = False
          'Modify by Amy 2020/09/07 其他系統別也都彈訊息 原:(field(1) = "P" Or field(1) = "PS" Or field(1) = "CFP" Or field(1) = "CPS")
          If txtCaseField(1).Text = "Y" Then
            'Modify by Amy 2020/09/07 +cp14
            strExc(0) = "select cp09,cp13,cp14 from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp27 is null and cp57 is null and cp09<'B'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               stCP09 = "" & RsTemp.Fields("cp09")
               '2008/1/16 modify by sonia 改依一般客戶智權人員之抓法 CFP-018376
               'stCP13 = "" & RsTemp.Fields("cp13")
               stCP13 = "" & PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4))
               stCP14 = "" & RsTemp.Fields("cp14") 'Add by Amy 2020/09/07
               '2008/1/16 end
               txtCaseField(1) = ""
               bolMail = True
               MsgBox "本案尚有收文號【" & stCP09 & "】未發文故將不閉卷，" & vbCrLf & "但會發Mail通知智權人員【" & GetStaffName(stCP13, True) & "】" & _
                            "及承辦人【" & GetStaffName(stCP14, True) & "】！", vbExclamation
            End If
          End If
          'end 2007/1/122
          
          BolFileCloseOk = False
          If txtCaseField(1) = "Y" Then
               'Modify by Amy 2018/08/30 +判斷及訊息文字
               strMsg = ""
               'Modify by Amy 2018/09/05 拿掉Pub_StrUserSt03 <> "P12" 都彈是否閉卷訊息
               'Modify by Amy 2021/01/11 +顯示下一程序案件性質名稱,若有多筆要串起來顯示
               'If txtCaseField(1).Tag = "不可閉卷" Then strMsg = "下一程序尚有未續辦案件，確定要閉卷？"
               If InStr(txtCaseField(1).Tag, "不可閉卷") > 0 Then
                    strMsg = "下一程序尚有未續辦案件如下：" & vbCrLf & Replace(txtCaseField(1).Tag, "不可閉卷", "") & vbCrLf & vbCrLf & "確定要閉卷？"
               End If
               If CheckCloseFile(strMsg) Then
               'end 2018/08/30
                  BolFileCloseOk = True
               Else
                  If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
                  Exit Sub
               End If
          End If
         
          'add by nickc 2005/04/22
          '2009/8/18 MODIFY BY SONIA T大陸部分核駁1205之復審401不辦時不必詢問T-151044
          'Pub_EndModCashMsg lblCaseField(8).Caption
          If field(1) = "T" And lblCaseField(3) = "401" And cp(10) = "1205" Then
          '2009/11/11 ADD BY SONIA T大陸部分勝部分敗1006之下一程序不辦時不必詢問T-151858
          ElseIf field(1) = "T" And cp(10) = "1006" Then
          'ADD BY SONIA 2016/8/31 P下一程序標準專利記錄請求110之不辦時不必詢問P-111047
          ElseIf field(1) = "P" And lblCaseField(3) = "110" Then
          'END 2016/8/31
          Else
          'Modified by Lydia 2015/02/12 P案非台灣案一律上結餘日,其餘皆不詢問
             If field(9) <> 台灣國家代號 And field(1) = "P" Then
                bolEndModCash = True  '自動上結餘日
             'add by sonia 2016/3/28 CFT異議案不必詢問是否計算結餘,不可結餘
             ElseIf field(1) = "CFT" And lblCaseField(3) = "601" Then
                bolEndModCash = False
             'end 2016/3/28
             'add by sonia 2024/12/20 CFP卷宗性質非申請時一律上結餘日,不必詢問
             ElseIf field(1) = "CFP" And field(23) <> "1" Then
                'add by sonia 2025/4/23 再加排除下一程序為補文件202、修正204、陳述意見205、補充說明206
                'bolEndModCash = True
                If lblCaseField(3) <> "202" And lblCaseField(3) <> "204" And lblCaseField(3) <> "205" And lblCaseField(3) <> "206" Then
                   bolEndModCash = True
                Else
                   Pub_EndModCashMsg lblCaseField(8).Caption, field(1), field(2), field(3), field(4)
                End If
                'end 2025/4/23
             '2024/12/20 CFP之PCT申請案一律上結餘日,不必詢問
             ElseIf field(1) = "CFP" Then
                strExc(0) = "select cp09 from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='109' and cp159=0"
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                'modify by sonia 2025/4/23 改為PCT案之下一程序為實體審查416、進入國家階段119,不必詢問
                If intI = 1 And (lblCaseField(3) = "416" Or lblCaseField(3) = "119") Then
                   bolEndModCash = True
                'add by sonia 2025/4/22 因2024/12/20加入上述控制其他情形就不會問了
                Else
                   Pub_EndModCashMsg lblCaseField(8).Caption, field(1), field(2), field(3), field(4) 'end 2025/4/22
                End If
             'end 2024/12/20
             Else
                If field(1) <> "P" Then Pub_EndModCashMsg lblCaseField(8).Caption, field(1), field(2), field(3), field(4)
             End If
            '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
            'Pub_EndModCashMsg lblCaseField(8).Caption
            'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
'            If Not (m_bolFMP = True And (lblCaseField(3) = "110" Or lblCaseField(3) = "203")) Then
'             ' Pub_EndModCashMsg lblCaseField(8).Caption, field(1), field(2), field(3), field(4)
'               'Add by Lydia 2015/02/10 P案非台灣案自動上結餘日
'                If field(9) <> 台灣國家代號 And field(1) = "P" Then
'                   If ReadPA21EndModCash(cp()) = True And InStr("110,203", lblCaseField(3)) = 0 And m_bolFMP = False Then
'                      bolEndModCash = True  '自動上結餘日
'                   Else
'                      Pub_EndModCashMsg lblCaseField(8).Caption, field(1), field(2), field(3), field(4)
'                   End If
'                Else
'                  Pub_EndModCashMsg lblCaseField(8).Caption, field(1), field(2), field(3), field(4)
'                End If
'               'end 2015/02/10
'            End If
          'end  'Modified by Lydia 2015/02/12
          End If
          '2009/8/18 end
          
          'Added by Lydia 2022/05/16 一案二請新型案上年費不續辦/閉卷時，判斷發明案尚未核准
          'Memo by Lydia 2022/05/18 彈訊息從「不續辦/閉卷」改為「不續辦」
          'Modified by Morgan 2024/12/4 要用下一程序的性質,cp(10)是相關收文號的性質 cp(10)->lblCaseField(3)
          If field(1) = "FCP" And field(8) = "2" And field(9) = "000" And lblCaseField(3) = "605" Then
               If PUB_IsDualApply(field, strExc, , , , , , True) = True Then
                   strExc(0) = "select pa16 from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' and pa08='1' "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                       '"是"按鈕'上不續辦/閉卷；"否"按鈕'新增以下
                       If "" & RsTemp.Fields("pa16") = "" Then
                          If MsgBox("為一案兩請且發明案尚未核准，是否不續辦？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                              strExc(5) = PUB_GetFCPSalesNo(field(1), field(2), field(3), field(4), lblCaseField(3))
                              strExc(6) = PUB_GetFCPProSup(strExc(5))
                              strExc(7) = PUB_GetFCPHandler(field(1), field(2), field(3), field(4), lblCaseField(3))
                              strExc(8) = PUB_GetFCPProSup(strExc(7))
                              '行事曆
                              If PUB_AddFCPStaffCalendar(CompWorkDay(4, strSrvDate(1)), "1", strExc(5) & "," & strExc(7), "一案二請因新型案接獲客戶不續辦通知且發明案尚未核准，待承辦通知客戶確認再上不續辦", strExc(5) & "," & strExc(7), "1", field(1), field(2), field(3), field(4)) = False Then
                                  Exit Sub
                              End If
                              strExc(1) = "【一案兩請】新型案接獲客戶年費不續辦通知且""發明案尚未核准""，請確認報告客戶。 Our Ref:" & field(1) & "-" & field(2) & IIf(field(3) <> "0", "-" & field(3), "") & IIf(field(4) <> "00", "-" & field(4), "") & "[INCOM.]"
                              strExc(2) = "1. 新型案接獲客戶不續辦通知且""發明案尚未核准""，請確認報告客戶。" & vbCrLf & _
                                               "2. 行事曆已自動新增一3天期限" & vbCrLf & _
                                               "3. 承辦確認後再行通知程序上不續辦"
                              strExc(3) = strExc(6) & ";" & strExc(7) & ";" & strExc(8) & ";backup"
                              PUB_SendMail strUserNum, strExc(5), "", strExc(1), strExc(2), , , , , , strExc(3)
                              Exit Sub
                          End If
                       End If
                   End If
               End If
          End If
          'end 2022/05/16
          
          Screen.MousePointer = vbHourglass
        'Modify By Cheng 2002/11/12
'          For i = 0 To 7
'          For i = 0 To 8
          For i = 0 To 10
                'Modified by Morgan 2021/5/6 +And i <> 3
                If i <> 8 And i <> 3 Then
                    If txtCaseField(i).Enabled Then
                       If CheckKeyIn(i) <> 1 Then
                          If cmdOK(3).Visible = True Then cmdOK(3).Enabled = True
                          txtCaseField(i).SetFocus
                          txtCaseField_GotFocus (i)
    '                      Exit For
                            Screen.MousePointer = vbDefault
                            Exit Sub
                       End If
                    End If
                End If
          Next
          'end 2024/10/04

         'Added by Morgan 2015/11/3 指示信電子化
         'P非臺灣案指示信都要彈修改畫面來確認送判的內容
         m_boleOrderLetter = False
         'Modified by Morgan 2015/12/15 外專程序結案除外
         'Modified by Morgan 2018/8/16 +CFP電子化
         If (field(1) = "P" Or (field(1) = "CFP" And strSrvDate(1) >= CFP指示信電子化啟用日)) And field(9) <> "000" And txtCaseField(4) = "" And Left(Pub_StrUserSt03, 1) <> "F" Then
            m_boleOrderLetter = True
         End If
         'end 2015/11/3
         'Modify by Amy 2021/06/25 取得案件表單主檔程序人員
         If UCase(mPrev01.Name) = UCase("frm210149_1") Then
            strOldF0308 = GetFlow003Data(Trim(mPrev01.txtF0301), , "F0308")
         End If
         'Add by Amy 2025/08/07 C類收文是否請款
         If field(1) = "FCP" Then strPA146 = field(146)

          '911106 nick transation
'          On Error GoTo CheckingErr
          cnnConnection.BeginTrans
          '**************************************************
          ' nick 900801 改
          '收文號
          'frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 0)
          'frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 10)
          '本所案號
          'cp(1), cp(2), cp(3), cp(4)
          'UPDATE 是否續辦為 N 和解除期限日期和解除期限原因
            'Modify By Cheng 2003/04/15
'          strSQL = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & txtCaseField(8) & "' WHERE NP01='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 0) & "' AND NP22='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 10) & "' "

'          strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & Left(Me.cboReason.Text, 2) & "' WHERE NP01='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0) & "' AND NP22='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 10) & "' "
           'Add by Lydia 2014/10/14 FMP案
          strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & Left(Me.cboReason.Text, 2) & "' WHERE NP01='" & strNP01 & "' AND NP22='" & strNP22 & "' "

          cnnConnection.Execute strSql
          'Add By Sindy 2013/5/24 CFT歐盟被異議案CFT-015285 : CFT案件若解除之下一程序性質為"異議答辯"時,若該NP01還有"緩衝期限"312時, 同時解除"緩衝期限"的期限.
          If field(1) = "CFT" And field(10) = "239" And lblCaseField(3) = "602" Then
           ' strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & Left(Me.cboReason.Text, 2) & "'" & _
                     " WHERE NP01='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0) & "'" & _
                     " AND NP07='312' AND NP06 is null"
                     
           'Add by Lydia 2014/10/14 FMP案
            strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & Left(Me.cboReason.Text, 2) & "'" & _
                     " WHERE NP01='" & strNP01 & "'" & _
                     " AND NP07='312' AND NP06 is null"
            cnnConnection.Execute strSql
          End If
          '2013/5/24 End
          
         '92.4.11 ADD BY SONIA
         If field(1) <> "CFT" Then
            ' 更新內商承辦人C類相關總收文號之發文日為系統日
'            strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(txtCaseField(0)) & " " & _
'                     "WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,STAFF " & _
'                     "WHERE CP09 = '" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0) & _
'                     "' AND CP09>'C' AND CP27 IS NULL " & _
'                     "AND CP14=ST01(+) AND ST03>='P20' AND ST03<='P29')"
           'Add by Lydia 2014/10/14 FMP案
            strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(txtCaseField(0)) & " " & _
                     "WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,STAFF " & _
                     "WHERE CP09 = '" & strNP01 & _
                     "' AND CP09>'C' AND CP27 IS NULL " & _
                     "AND CP14=ST01(+) AND ST03>='P20' AND ST03<='P29')"
            cnnConnection.Execute strSql
         End If
         '92.4.11 END
          
          'Added by Lydia 2017/05/08 記錄子案案號(避免Trigger一併更新子案)
          mCCase = ""
          If field(1) = "CFP" And field(3) & field(4) = "000" Then
            strExc(0) = "SELECT PA01||'-'||PA02||'-'||PA03||'-'||PA04 From PATENT " & _
                        "WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03||PA04 <> '" & field(3) & field(4) & "' AND NVL(PA57,'N') <> 'Y' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               mCCase = RsTemp.GetString(adClipString, , , ",")
            End If
          End If
          'end 2017/05/08
          
          If BolFileCloseOk Then
             Select Case Val(CheckSys(field(1)))
             Case 1
                  strSql = "UPDATE PATENT SET PA57='" & txtCaseField(1) & "',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & Left(Me.cboReason.Text, 2) & "',PA89='" & txtCaseField(2) & "',PA91='" & ChgSQL(cboMemo.Text) & "' WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03='" & field(3) & "' AND PA04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 2
                  strSql = "UPDATE TRADEMARK SET TM29='" & txtCaseField(1) & "',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & Left(Me.cboReason.Text, 2) & "',TM58='" & ChgSQL(cboMemo.Text) & "' WHERE TM01='" & field(1) & "' AND TM02='" & field(2) & "' AND TM03='" & field(3) & "' AND TM04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 3
                  strSql = "UPDATE LAWCASE SET LC08='" & txtCaseField(1) & "',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & Left(Me.cboReason.Text, 2) & "',LC27='" & ChgSQL(cboMemo.Text) & "' WHERE LC01='" & field(1) & "' AND LC02='" & field(2) & "' AND LC03='" & field(3) & "' AND LC04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 4
                  strSql = "UPDATE HIRECASE SET HC09='" & txtCaseField(1) & "',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & Left(Me.cboReason.Text, 2) & "',HC12='" & ChgSQL(cboMemo.Text) & "' WHERE HC01='" & field(1) & "' AND HC02='" & field(2) & "' AND HC03='" & field(3) & "' AND HC04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 5, 6, 7, 8
                  strSql = "UPDATE SERVICEPRACTICE SET SP15='" & txtCaseField(1) & "',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & Left(Me.cboReason.Text, 2) & "',SP18='" & ChgSQL(cboMemo.Text) & "' WHERE SP01='" & field(1) & "' AND SP02='" & field(2) & "' AND SP03='" & field(3) & "' AND SP04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case Else
             End Select
            
            'Added by Morgan 2016/8/24 與閉卷相同所有未發文CP都上取消收文日
            strSql = "UPDATE CASEPROGRESS SET CP26='N',CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & Left(Me.cboReason.Text, 2) & "' WHERE CP01='" & field(1) & "' AND CP02='" & field(2) & "' AND CP03='" & field(3) & "' AND CP04='" & field(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
            '排除FCP案的代辦退費(實審,再審和再審延期)
            If cp(1) = "FCP" Then
               'Modified by Morgan 2022/11/23 +排除續行母案再審的代辦退費 Ex:FCP-067213 --Winfrey
               strSql = strSql & "and cp09 not in (select a.cp09 from caseprogress a,caseprogress b where a.cp01='" & field(1) & "' and a.cp02='" & field(2) & "' and a.cp03='" & field(3) & "' and a.cp04='" & field(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10 in ('416','107','435') " & _
                        "union select a.cp09 from  caseprogress a,caseprogress b,nextprogress where a.cp01='" & field(1) & "' and a.cp02='" & field(2) & "' and a.cp03='" & field(3) & "' and a.cp04='" & field(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07='107' " & _
                        "union select a.cp09 from  caseprogress a,caseprogress b,caseprogress c where a.cp01='" & field(1) & "' and a.cp02='" & field(2) & "' and a.cp03='" & field(3) & "' and a.cp04='" & field(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10='107') "
            End If
            cnnConnection.Execute strSql, intI
            'end 2016/8/24
            
          Else
             Select Case Val(CheckSys(field(1)))
             Case 1
                  strSql = "UPDATE PATENT SET PA89='" & txtCaseField(2) & "',PA91='" & ChgSQL(cboMemo.Text) & "' WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03='" & field(3) & "' AND PA04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 2
                  strSql = "UPDATE TRADEMARK SET TM58='" & ChgSQL(cboMemo.Text) & "' WHERE TM01='" & field(1) & "' AND TM02='" & field(2) & "' AND TM03='" & field(3) & "' AND TM04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 3
                  strSql = "UPDATE LAWCASE SET LC27='" & ChgSQL(cboMemo.Text) & "' WHERE LC01='" & field(1) & "' AND LC02='" & field(2) & "' AND LC03='" & field(3) & "' AND LC04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 4
                  strSql = "UPDATE HIRECASE SET HC12='" & ChgSQL(cboMemo.Text) & "' WHERE HC01='" & field(1) & "' AND HC02='" & field(2) & "' AND HC03='" & field(3) & "' AND HC04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case 5, 6, 7, 8
                  strSql = "UPDATE SERVICEPRACTICE SET SP18='" & ChgSQL(cboMemo.Text) & "' WHERE SP01='" & field(1) & "' AND SP02='" & field(2) & "' AND SP03='" & field(3) & "' AND SP04='" & field(4) & "' "
                  cnnConnection.Execute strSql
             Case Else
             End Select
          End If
          'Add by Amy 2025/08/05 外專人員 [後續准駁簡單報告] =Y,則[C類收文是否請款 (pa146)] 設 N -Winfrey 請作
          'Modify by Amy 2025/08/07 改抓變數,C類收文是否請款 pa146 只有專利有
          If Left(PUB_GetST03(strUserNum), 2) = "F2" And txtCaseField(2) = "Y" And strPA146 <> "N" Then
            strSql = "PA146='N',pa91='" & CFDate(strSrvDate(2)) & " 不續辦:後續准駁簡單報告;'||pa91"
            strSql = "Update Patent SET " & strSql & " Where PA01='" & field(1) & "' And PA02='" & field(2) & "' And PA03='" & field(3) & "' And PA04='" & field(4) & "' "
            cnnConnection.Execute strSql
          End If
                        
          bolLeave = True
          'If BolFileClose = False Then
              'If BolFileCloseOk = True Then
                   Dim strAutoNum As String
                  'Modify By Cheng 2002/10/01
'                   If objPublicData.GetAutoNumber("B", strAutoNum, True, False) Then
                   'edit by nickc 2007/02/02 不用 dll 了
                   'If objPublicData.GetAutoNumber("B", strAutoNum, True, True) Then
                   If ClsPDGetAutoNumber("B", strAutoNum, True, True) Then
                        CheckOC
                        strSql = "select au01||(au02-1911) from autonumber where au01='B'"
                        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                        If Not adoRecordset.BOF Then adoRecordset.MoveFirst
                        'Modify By Cheng 2002/11/13
'                        If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: Exit Sub
                        If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: GoTo CheckingErr
                        'Modify By Sindy 2010/8/18 比對自動編號年度
                        'strAutoNum = CheckStr(adoRecordset.Fields(0).Value) & strAutoNum
                        strAutoNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) & strAutoNum
                        CheckOC
                        'Modify By Sindy 2015/5/19 +cp140
                        'Remove by Lydia 2017/01/25  B類收文模組化
                        'strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20,cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30,cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50,cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60,cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76,cp77,cp78,cp79,cp140) values "
                        
                        'Set SCp() = cp()
                        For i = 1 To 79
                           Select Case i
                           '文字null
                           'Modify By Cheng 2002/07/26
'                           Case 8, 11, 14, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63
                           '92.1.25 MODIFY BY SONIA 取消收文日及原因要存
                           'Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63
                           
'Remove by Morgan 2005/1/5 有case else 此處不必控制,否則加欄位時都要改
'                           Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 59, 60, 61, 62, 63
'                           '92.1.25 END
'                                SCp(i) = "null "
'2005/1/5 end

                           'Add By Cheng 2002/07/26
                           Case 14 '承辦人代號
                                SCp(i) = "'" & strUserNum & "'"
                           '文字畫面上
                           Case 64
                                SCp(i) = "'" & ChgSQL(Trim(txtCP64)) & "'"
                           Case 1, 2, 3, 4
                                SCp(i) = "'" & Trim(ChgSQL(field(i))) & "'"
                           Case 12
                                '2012/10/2 modify by sonia
                                'SCp(i) = "'" & Trim(ChgSQL(cp(i))) & "'"
                                 SCp(i) = "'" & GetSalesArea(PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4))) & "'"
                           '2011/6/1 add by sonia 自上面Case 12抽出來
                           Case 13
                                'Modify by Amy 2023/01/31 FCT案存畫面上智權人員
                                If txtSalesNo.Locked = False Then
                                    SCp(i) = "'" & txtSalesNo & "'"
                                Else
                                    SCp(i) = "'" & PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4)) & "'"
                                End If
                           '2011/6/1 end
                           Case 5, 27
                                SCp(i) = strSrvDate(1)
                           Case 9
                                SCp(i) = "'" & strAutoNum & "'"
                           '91.12.6 modify by sonia
                           'Case 26, 20, 32
                           '     SCp(i) = "'N'"
                           Case 20
                              If intWhere <> "2" Then
                                 SCp(i) = "'N'"
                                 '2013/8/13 add by sonia FMT要請款
                                 If cp(1) = "T" And Left(GetSalesArea(PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4))), 1) = "F" Then
                                    SCp(i) = "null "
                                 End If
                                 '2013/8/13 end
                              'Add by Morgan 2007/7/23 改抓CPM設定
                              'Modify by Amy 2025/06/03 +FG-陳亭妙 ex:FG-001323(亭妙已自行補上cp20=N)
                              ElseIf cp(1) = "FCP" Or cp(1) = "FG" Then
                                 SCp(i) = CNULL(PUB_GetCP20(cp(1), Replace(SCp(10), "'", "")))
                              Else
                                 SCp(i) = "null "
                              End If
                           Case 26, 32
                                SCp(i) = "'N'"
                           '91.12.6 end
                           Case 43
                                SCp(i) = "'" & cp(9) & "'"
                           Case 10
                                If (IsNull(txtCaseField(1)) Or txtCaseField(1) = "") Then
                                    Select Case Val(CheckSys(field(1)))
                                    Case 1, 5         'patent
                                       SCp(i) = "'907'"
                                    Case 2, 6         'trademark
                                       SCp(i) = "'703'"
                                    Case 3, 4, 7, 8   'lawcase & hirecase
                                       SCp(i) = "'991'"
                                    Case Else
                                    End Select
                                Else
                                    Select Case Val(CheckSys(field(1)))
                                    Case 1, 5         'patent
                                       SCp(i) = "'913'"
                                    Case 2, 6         'trademark
                                       SCp(i) = "'704'"
                                    Case 3, 4, 7, 8   'lawcase & hirecase
                                       SCp(i) = "'993'"
                                    Case Else
                                    End Select
                                End If
                           Case 65, 66, 67, 68, 69, 70
                                SCp(i) = ""
                           '92.1.25 ADD BY SONIA
                           Case 57
                                SCp(i) = ChangeTStringToWString(txtCaseField(0))
                           Case 58
                                SCp(i) = "'" & Trim(ChgSQL(Left(Me.cboReason.Text, 2))) & "'"
                           '92.1.25 END
                           '數字
                           'Add by Morgan 2005/1/5
                           Case 44
                              If Combo2 <> "" Then
                                 SCp(i) = "'" & Left(Combo2 & "00000000", 9) & "'"
                              Else
                                 SCp(i) = "null "
                              End If
                           'Add by Morgan 2007/1/8
                           Case 45
                              SCp(i) = CNULL(ChgSQL(Me.txtCaseField(11)))
                              
                           '2006/9/15 ADD BY SONIA
                           Case 30
                           '   SCp(i) = "'" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 10) & "'"
                            'Add by Lydia 2014/10/14 FMP案
                            SCp(i) = "'" & strNP22 & "'"
                           '2006/9/15 END
                           Case Else
                                SCp(i) = "null "
                           End Select
                        Next i
                        'Modified by Lydia 2017/01/25 B類收文模組化
'                        strSql = strSql & " ("
'                        For i = 1 To 79
'                            Select Case i
'                            Case 65, 66, 67, 68, 69, 70
'                            Case Else
'                                 strSql = strSql & SCp(i)
'                                 If i <> 79 Then
'                                    strSql = strSql & ","
'                                 End If
'                            End Select
'                        Next i
'                        'Add By Sindy 2015/5/19 +結案單電子化 : CP140
'                        If UCase(mPrev01.Name) = UCase("frm210149_1") Then
'                           strSql = strSql & ",'" & mPrev01.txtF0301 & "'"
'                        Else
'                           strSql = strSql & ",null"
'                        End If
'                        '2015/5/19 END
'                        strSql = strSql & ") "
                        strSql = GetInsBCP
                        'end 2017/01/25
                        cnnConnection.Execute strSql
                        
                        'Added by Morgan 2015/11/3 指示信電子化
                        If m_boleOrderLetter Then
                           'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                           'strLetterJudge = Pub_GetSpecMan("PS4") 'P案指示信判發人
                           strCP10 = Replace(SCp(10), "'", "")
                           strLetterJudge = PUB_GetLetterJudgeNew("2", field(1), strCP10, field(9), lblCaseField(3))
                           'Modified by Morgan 2018/8/20 +傳指示信主旨strSubject
                           'Modified by Morgan 2025/7/25 彼號改抓畫面上的
                           strSubject = PUB_GetSubject(field(1), field(2), field(3), field(4), strCP10, field(11), txtCaseField(11))
                           PUB_AddAppForm strAutoNum, True, strLetterJudge, strSubject
                           strSubject = ""
                           m_strAF01 = strAutoNum
                           'end 2018/8/20
                        End If
                        'end 2015/11/3
                        
                        'Added by Lydia 2017/01/25 CFP案解除期限程序若為EPC 准後已繳指定註冊費案,原僅產生母案之結案指示信,現請同時上結案的子案也一併帶出指示信
                        'Modified by Lydia 2017/05/08 +有子案 mccase
                        If field(1) = "CFP" And field(4) = "00" And field(9) = "221" And field(16) = "1" And mCCase <> "" Then
                           '和使用者確認過,以領證發文為已繳指定註冊費案 (by 89037)
                           If PUB_ChkCPExist(field, "601", 2) Then
                                'Modified  by Lydia 2017/05/08 改成先記錄子案案號
'                                strExc(0) = "SELECT PA01,PA02,PA03,PA04 From PATENT " & _
'                                            "WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03||PA04 <> '" & field(3) & field(4) & "' AND (PA57<>'Y' OR PA57 IS NULL) ORDER BY 3,4 "
'                                intI = 1
'                                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                                If intI = 1 Then
'                                   RsTemp.MoveFirst
'                                   strChildList = ""
'                                   Do While Not RsTemp.EOF
'                                      tmpNo = AutoNo("B", 6)
'                                      '子案的相關總收文為母案B類收文號
'                                      strSql = GetInsBCP(RsTemp.Fields("PA03"), RsTemp.Fields("PA04"), tmpNo, strAutoNum)
'
'                                      '子案B類收文
'                                      cnnConnection.Execute strSql
'                                      strChildList = strChildList & IIf(strChildList <> "", ",", "") & tmpNo
'                                      RsTemp.MoveNext
'                                   Loop
'                                End If
                                tmpArr = Empty
                                tmpArr = Split(mCCase, ",")
                                strChildList = ""
                                For i = 0 To UBound(tmpArr)
                                   If Trim(tmpArr(i)) <> "" Then
                                       tmpNo = AutoNo("B", 6)
                                      '子案的相關總收文為母案B類收文號
                                      'Modified by Morgan 2018/8/20 + strChildCP45
                                      strSql = GetInsBCP(SystemNumber(Trim(tmpArr(i)), 3), SystemNumber(Trim(tmpArr(i)), 4), tmpNo, strAutoNum, , strChildCP45)

                                      '子案B類收文
                                      cnnConnection.Execute strSql
                                      
                                      'Added by Morgan 2018/8/16
                                      If m_boleOrderLetter Then
                                          strSubject = PUB_GetSubject(field(1), field(2), SystemNumber(Trim(tmpArr(i)), 3), SystemNumber(Trim(tmpArr(i)), 4), strCP10, field(11), strChildCP45)
                                          PUB_AddAppForm tmpNo, False, strLetterJudge, strSubject
                                          strSubject = ""
                                      End If
                                      'end 2018/8/16
                                      
                                      strChildList = strChildList & IIf(strChildList <> "", ",", "") & tmpNo
                                   End If
                                Next i
                           End If
                        End If
                        'end 2017/01/25
                        
                        'Add by Sindy 2013/04/12 更新c類的代理人及彼所案號，要在新增c類之後
                        If Combo2.Enabled = False Then 'Added by Morgan 2025/10/17 要以畫面的代理人為準
                           Pub_UpdateFromMaxCP27 field(1), field(2), field(3), field(4)
                        End If
                        'end 2025/10/17
                        
                        bolLeave = True
                        intLeaveKind = 2
                        Me.Hide
                   Else
                       MsgBox ("自動給號錯誤")
                       Me.Hide
                   End If
               'End If
          'End If
          If Len(txtCaseField(6)) <> 0 Or Len(txtCaseField(7)) <> 0 Then
            'frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 0)
            'frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 10)
            '92.9.10 MODIFY BY SONIA
            'strSQL = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT NP01,NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",NP10,NP13,NP14," & GetNextProgressNo & ",'" & Trim(txtCP64) & "' FROM NEXTPROGRESS WHERE NP01='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 0) & "' AND NP22=" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 10) & " "
'            If (lblCaseField(3) = "605" Or lblCaseField(3) = "606" Or lblCaseField(3) = "607") And (field(1) = "FCP" Or field(1) = "CFP" Or field(1) = "P") Then
'               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT '" & strAutoNum & "',NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",NP10,NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0) & "' AND NP22=" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 10) & " "
'            Else
'               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT NP01,NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",NP10,NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0) & "' AND NP22=" & frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 10) & " "
'            End If

           'Add by Lydia 2014/10/14 FMP案
            If (lblCaseField(3) = "605" Or lblCaseField(3) = "606" Or lblCaseField(3) = "607") And (field(1) = "FCP" Or field(1) = "CFP" Or field(1) = "P") Then
               'modify by sonia 2020/5/28
               'strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT '" & strAutoNum & "',NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",NP10,NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22 & " "
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT '" & strAutoNum & "',NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",'" & PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4)) & "',NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22 & " "
            'Add by Amy 2023/01/31 FCT案可輸智權人員,以畫面上輸為主
            ElseIf txtSalesNo.Locked = False Then
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT NP01,NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",'" & txtSalesNo & "',NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22 & " "
            Else
               'strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT NP01,NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",NP10,NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22 & " "
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT NP01,NP02,NP03,NP04,NP05,NP07," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",'" & PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4)) & "',NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22 & " "
            End If
            'Add by Lydia 2014/10/14 FMP案.end
            cnnConnection.Execute strSql
          End If
          
          'add by nickc 2005/04/22
          Pub_UpdateEndModCash field(1), field(2), field(3), field(4)
          
          'Add By Sindy 2023/12/13 檢查接洽單的Flow是否要結束
          'Modify By Sindy 2024/11/20 +, Me, _
            IIf(lblCaseField(3) = "218" And lblCaseField(8) = "221" And field(1) = "CFP", True, False)
          Call PUB_UpdateCRLFlowClose(cp(140), cp(9), Me, _
            IIf(lblCaseField(3) = "218" And lblCaseField(8) = "221" And field(1) = "CFP", True, False))
          
          'Add By Sindy 2015/1/14 結案單電子化
          'Modify by Amy 2022/06/20 +frm210149,商標延展可顯示close.menu
          If UCase(mPrev01.Name) = UCase("frm210149_1") Or UCase(mPrev01.Name) = UCase("frm210149") Then
            intLeaveKind = 2 '0
            strUpdDate = strSrvDate(1)
            strUpdTime = Right("000000" & ServerTime, 6)
'            '記錄電子表單編號
'            strSql = "update caseprogress set cp140='" & mPrev01.txtF0301 & "' where cp09=" & SCp(9)
'            cnnConnection.Execute strSql
            '卷宗區 : 新增一筆結案單Close至卷宗區
            'Modify By Sindy 2020/2/19 電子檔名,本所案號使用函數 PUB_CaseNo2FileName
'            strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
'                     " values(" & SCp(9) & "," & _
'                             "'" & field(1) & Val(field(2)) & IIf(field(3) = "0" And field(4) = "00", "", "-" & field(3)) & IIf(field(4) = "00", "", "-" & field(4)) & "." & Replace(SCp(10), "'", "") & "." & EMP_結案單 & ".menu',0,'" & strUserNum & "'," & _
'                             strUpdDate & "," & strUpdTime & "," & _
'                             strUpdDate & "," & strUpdTime & ",'Y')"
            strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                     " values(" & SCp(9) & ",'" & PUB_CaseNo2FileName(field(1), field(2), field(3), field(4)) & _
                             "." & Replace(SCp(10), "'", "") & "." & EMP_結案單 & ".menu',0,'" & strUserNum & "'," & _
                             strUpdDate & "," & strUpdTime & "," & _
                             strUpdDate & "," & strUpdTime & ",'Y')"
            cnnConnection.Execute strSql
            
            'Modify by Amy 2022/06/20 +if
            If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                bolMailF0202_3 = True 'Add by Amy 2025/08/19
                'Modify by Amy 2021/06/23 職代需註明(代)
                strSql = ""
                'Modify by Amy 2025/06/02 目前只有內商人員不是掛在個人
                'If (field(1) = "CFT" Or field(1) = "CFC" Or field(1) = "S") And strOldF0308 <> MsgText(601) Then
                If Pub_StrUserSt03 <> "P21" And strOldF0308 <> MsgText(601) Then
                    If strOldF0308 <> strUserNum Then
                        strSql = " ,F0208='(代)' "
                    End If
                    '葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現
                    'If strUserNum = "78011" Or strUserNum = "80030" Then
                    'Modify by Amy 2025/06/09 + lblCaseField(8) <> "000"
                    'Modify by Amy 2025/08/19 +外商案之外商程序人員=補看人員者(目前st03=F12的二級主管為 湘嫺),補看不出現,也不發信給補看
                    If ((field(1) = "CFT" Or field(1) = "CFC" Or (field(1) = "S" And lblCaseField(8) <> "000")) And (strUserNum = "78011" Or strUserNum = "80030")) _
                      Or (intFCState = 1 And strSrvDate(1) >= FCT結案單電子化啟用日 And strUserNum = "79020") Then
                'end 2025/06/02
                        strCmd(0) = "update FLOW002 set " & _
                                "F0205='" & strUpdDate & "'" & _
                                ",F0206='" & strUpdTime & "'" & _
                                ",F0207='3',F0204='" & strUserNum & "'" & _
                                " where F0201='" & mPrev01.txtF0301 & "' and F0202='3' and F0207 is null "
                        strCmd(1) = "Update FLOW003 Set F0309=" & CNULL(Flow_歸檔) & " Where F0301='" & mPrev01.txtF0301 & "'"
                        bolMailF0202_3 = False 'Add by Amy 2025/08/19
                    End If
                End If
                '簽核檔-程序人員:3.已處理
                strSql = "update FLOW002 set " & _
                         "F0205='" & strUpdDate & "'" & _
                         ",F0206='" & strUpdTime & "'" & _
                         ",F0207='3',F0204='" & strUserNum & "'" & strSql & _
                         " where F0201='" & mPrev01.txtF0301 & "' and F0202='2' and F0207 is null "
                cnnConnection.Execute strSql
                'end 2021/06/23
                '讀取下一處理人員
                'Modified by Morgan 2015/11/3 +傳m_boleOrderLetter
                If GetNextProPerson_Flow(Trim(mPrev01.txtF0301), Trim(mPrev01.m_F0316), strF0308, strF0309, m_boleOrderLetter) = False Then GoTo CheckingErr
                '流程備註檔
                If Trim(mPrev01.txtNote.Text) <> "" Then
                   strSql = GetInsertFLOW004Sql(Trim(mPrev01.txtF0301), strUserNum, strUpdDate, strUpdTime, strF0309, ChgSQL(Trim(mPrev01.txtNote.Text)))
                   cnnConnection.Execute strSql
                End If
                'Add by Amy 2021/06/23 葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現
                If strCmd(0) <> MsgText(601) Then
                    cnnConnection.Execute strCmd(0)
                    cnnConnection.Execute strCmd(1)
                End If
            End If
            'end 2022/06/20
          End If
          '2015/1/14 END
         
'Removed by Morgan 2012/10/1 改在列印結案單時提醒並列印於接洽單上
'            'Add by Morgan 2009/10/15
'            '大陸案一案兩請:新型年費欲結案時,若該發明案尚未核准公告,則發E-MAIL告知智權同仁及其所屬區主管
'            If cp(10) = "605" And field(1) = "P" And field(9) = "020" And field(8) = "2" And Val(DBDATE(field(10))) >= 20091001 Then
'               strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1" & _
'                  " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "'" & _
'                  " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & field(1) & "' and cm06='" & field(2) & "' and cm07='" & field(3) & "' and cm08='" & field(4) & "') X" & _
'                  ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null and (pa16 is null or pa16='2')"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strExc(1) = field(1) & "-" & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4))
'                  strExc(1) = "提醒:" & strExc(1) & "大陸案為一案兩請,新型放棄續繳年費將同時放棄發明或實用新型間擇一選擇的權利。"
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " values ('" & strUserNum & "','" & lblCaseField(7) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(1)) & "','如旨')"
'                  cnnConnection.Execute strSql, intI
'
'                  strExc(0) = "select a0908 from staff,acc090 where st01='" & lblCaseField(7) & "' and a0901(+)=st15 and a0908<>'" & lblCaseField(7) & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                        " values ('" & strUserNum & "','" & RsTemp(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(1)) & "','如旨')"
'                     cnnConnection.Execute strSql, intI
'                  End If
'               End If
'            End If
'            'end 2009/10/15
          'Added by Lydia 2015/10/12 商標爭議案無結果掛催審期限
          'T,FCT,CFT,TF之602異答,604評答,606廢答之解除期限或取消收文,都要掛被異議(1602)、被評定(1604)、被撤銷(1606)之c類來函的下一程序305催審
          If (field(1) = "T" Or field(1) = "TF" Or field(1) = "FCT" Or field(1) = "CFT") And (lblCaseField(3) = "602" Or lblCaseField(3) = "604" Or lblCaseField(3) = "606") Then
             '管制人員為1602,1604,1606之承辦人
             strExc(0) = " select cp09,cp14 from caseprogress where cp09='" & cp(9) & "' and cp10 in (1602,1604,1606) "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
             '法限=系統日+1年,所限=法限
                strExc(9) = PUB_GetWorkDay1(CompDate(0, 1, strSrvDate(1)), True)
                strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                         "VALUES('" & RsTemp.Fields("cp09") & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','305','" & strExc(9) & "','" & strExc(9) & "','" & RsTemp.Fields("cp14") & "'," & GetNextProgressNo & ") "
                cnnConnection.Execute strSql, intI
             End If
          End If
          'END 2015/10/12
          'Add by Amy 2018/11/27 FCP領證不續辦時,未超過領證期限增加行事曆提醒
          'Modify by Amy 2019/01/23 +閉卷 不增加行事曆提醒 ex:FCP-050292
          If field(1) = "FCP" And txtCaseField(1) <> "Y" And Val(CADate(FCDate(lblCaseField(5)))) >= Val(strSrvDate(1)) And lblCaseField(3) = "601" And bolHas907 = False Then
              strExc(0) = CompDate(2, 1, CADate(FCDate(lblCaseField(5))))
              strExc(1) = PUB_GetFCPHandler(field(1), field(2), field(3), field(4))
              If PUB_AddFCPStaffCalendar(strExc(0), 1, strExc(1), "寄領證逾期通知函", strExc(1), "1", field(1), field(2), field(3), field(4), strExc(0)) Then
              End If
          End If
          'Add by Amy 2025/08/08 外專結案單有勾 未付帳款 及有輸 管制催款日,加 未付帳款 行事曆提醒
          If intFCState = "2" Then
             If ChkCCD03(1, Me.Name, strF0301, strNotPay, strCCD08) = True Then
                  strExc(1) = PUB_GetFCPHandler(field(1), field(2), field(3), field(4)) '程序管制人
                  strExc(2) = PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4)) '承辦案件管制人
                  strExc(0) = strExc(1) & "," & strExc(2)
                  If InStr(strNotPay, "追蹤欠款：") = 0 Then strNotPay = "追蹤欠款：" & strNotPay
                  strCCD08 = Val(strCCD08) + 19110000
                  If PUB_AddFCPStaffCalendar(strCCD08, 1, strExc(0), strNotPay, strExc(0), "1", field(1), field(2), field(3), field(4), strCCD08, , , strF0301) Then
                  End If
             End If
          End If
         'end 2025/08/08
         'add by sonia 2020/5/27 T及FCT延展，解除期限原因19他所延展未變更代理人時，專用期間欄位依各部門規定延長，並管制延長後延展期限
         If Left(cboReason.Text, 2) = "19" Then
            strExc(0) = "SELECT NVL(NA14,0) NA14,TM22 FROM TRADEMARK,NATION WHERE TM01='" & field(1) & "' AND TM02='" & field(2) & "' AND TM03='" & field(3) & "' AND TM04='" & field(4) & "' AND TM10=NA01(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTM22New = RsTemp.Fields("TM22") + RsTemp.Fields("NA14") * 10000
               '專用期間欄位依各部門規定延長
               If field(1) = "FCT" Then   'FCT僅專用期止日+NA14,起日不動
                  strSql = "UPDATE TRADEMARK SET TM22=" & strTM22New & " WHERE TM01='" & field(1) & "' AND TM02='" & field(2) & "' AND TM03='" & field(3) & "' AND TM04='" & field(4) & "' "
               Else                       'T 專用期起日及止日都+NA14
                  strSql = "UPDATE TRADEMARK SET TM21=TM21+" & RsTemp.Fields("NA14") * 10000 & ",TM22=" & strTM22New & " WHERE TM01='" & field(1) & "' AND TM02='" & field(2) & "' AND TM03='" & field(3) & "' AND TM04='" & field(4) & "' "
               End If
               cnnConnection.Execute strSql
               '管制延長後延展期限
               Dim strDate(0 To 3) As String
               strDate(1) = field(1)     '系統別
               strDate(2) = lblCaseField(8) '國家
               'Modified by Lydia 2020/07/13 用西元年月
               'strDate(3) = ChangeTStringToWString(strTM22New)  '下次法定期限
               strDate(3) = strTM22New
               GetCtrlDT strDate()
              ' strDate(0) = ChangeWStringToTString(strDate(0)) 'Mark by Lydia 2020/07/13  用西元年月
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT NP01,NP02,NP03,NP04,NP05,NP07," & strDate(0) & "," & strTM22New & ",'" & PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4)) & "',NP13,NP14," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM NEXTPROGRESS WHERE NP01='" & strNP01 & "' AND NP22=" & strNP22 & " "
               cnnConnection.Execute strSql
            End If
         End If
         'end 2020/5/27
          
          'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：輸入閉卷913自動收文「通知資訊變更961」,發一封Email給承辦工程師
          If txtCaseField(1) = "Y" And field(1) = "FCP" And m_PA177 = "Y" Then
             'Memo by Lydia 2025/04/02 模組內已去掉SCp(10)的單引號Replace
             If PUB_GetFCPlinkMC("6", TransDate(txtCaseField(0), 2), field, strAutoNum, SCp(10)) = True Then
             End If
          End If
          'end 2023/07/28
          
          'Added by Lydia 2024/11/21 不續辦則不用保留內商大陸之部份核駁商品異動
          If field(1) = "T" And lblCaseField(3) = "401" And lblCaseField(8) = "020" Then
             strSql = "delete from tmgoods where tg01='" & Mid(cp(9), 1, 3) & "' and tg02='" & Mid(cp(9), 4, 6) & "' and tg03='0' and tg04='00' "
             cnnConnection.Execute strSql
          End If
          'end 2024/11/21
          
          'Added by Lydia 2025/09/12 TF基礎案號設定：基礎案已閉卷、【703不續辦-續展】=>基礎案狀態通知Email
          If (field(1) = "T" Or field(1) = "CFT") And (txtCaseField(1) = "Y" Or lblCaseField(3) = "102" Or lblCaseField(3) = "110") Then
             strSql = PUB_GetTFbaseInfo(field(1), field(2), field(3), field(4), field(15), field(10), "2", field(12), IIf(txtCaseField(1) = "Y", "", strAutoNum))
          End If
          'end 2025/09/12
          
          cnnConnection.CommitTrans
          'Add by Amy 2018/11/27 FCP領證不續辦時,超過領證期限彈訊息提醒
          'Modify by Amy 2019/01/23 +閉卷 不彈訊息提醒 ex:FCP-050292
          bolShow060318 = False
          If field(1) = "FCP" And txtCaseField(1) <> "Y" And Val(CADate(FCDate(lblCaseField(5)))) < Val(strSrvDate(1)) And lblCaseField(3) = "601" And bolHas907 = False Then
              bolShow060318 = True
              MsgBox "需寄領證逾期通知函！"
          End If
          'Moidfy by Amy 2018/10/17  Amy ADD 2018/08/30  +frm210149_1 由待處理區做T延展結案
          'If UCase(mPrev01.Name) = UCase("frm210149") Then intLeaveKind = 1 'Modify by Amy 2018/10/08 原:intLeaveKind = 0
          
          '2015/8/10 add by sonia 專利案件閉卷時有新案翻譯尚未完稿要提醒(FCP-51551)
          'modify by sonia 2015/9/4 再加cp05>20150101否則舊案無完稿也會有訊息
          If txtCaseField(1) = "Y" And intCaseKind = 專利 Then
            strExc(0) = "select cp09,nvl(ep09,0) ep09,nvl(cp27,0) cp27 from caseprogress,engineerprogress where " & ChgCaseprogress(field(1) & field(2) & field(3) & field(4)) & " and cp10='201' and cp09=ep02(+) and cp05>20150101 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Val("" & RsTemp("ep09")) = 0 Then
                  MsgBox "此案新案翻譯進度尚未完稿！"
               ElseIf Val("" & RsTemp("cp27")) = 0 Then
                  MsgBox "此案新案翻譯進度尚未發文！"
               End If
            End If
          End If
          '2015/8/10 end
          
          'Add by Amy 2018/07/18 結案後發mail通知相關人員
          'Modfy by Amy 2018/08/31 CFP補看人員也不發mail
          'Modify By Sindy 2025/6/4 內專不只有P,CFP案還有PS,CPS
          'If strSrvDate(1) >= 非P結案電子化啟用日 And field(1) <> "P" And field(1) <> "CFP" Then
          If Left(PUB_GetST03(strUserNum), 2) <> "P1" Then '非內專
          '2025/6/4 END
            If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                'Add by Amy 2018/10/08 +內文
                strContent = "案件名稱：" & cboCaseName & vbCrLf & _
                             "　申請人：" & lblCaseField(2) & lblPetitionName & vbCrLf & _
                             "案件性質：" & lblCaseField(3) & lblNextProgress & vbCrLf & _
                             "本所期限：" & lblCaseField(4) & vbCrLf & _
                             "法定期限：" & lblCaseField(5)
                'end 2018/10/08
                strSubject = field(1) & "-" & field(2) & "-" & field(3) & "-" & field(4) & " 已結案！"
                If Left(field(1), 1) = "T" Then
                    'T字頭通知承辦人員
                    strExc(0) = ""
                    strExc(0) = GetLastCP14(field(1), field(2), field(3), field(4), strAutoNum)
                    If strExc(0) <> MsgText(601) Then
                        PUB_SendMail strUserNum, strExc(0), "", strSubject, strContent 'Modify by Amy 2018/10/08 +strContent
                    End If
                End If
            '*** 通知補看人員 ***
                'Add by Amy 2021/06/28+ if 需發mail(葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現,故不需發mail)
                'Modfiy by Amy 2025/06/02 +外專/外商 案件
                'If (strCmd(0) = MsgText(601) And (field(1) = "CFT" Or field(1) = "CFC" Or field(1) = "S")) Or (field(1) <> "CFT" And field(1) <> "CFC" And field(1) <> "S") then
                'Modify By Sindy 2025/6/4 + And strNation <> "000")
                'Modify by Amy 2025/06/09 FCT上線延後,故加intFCState = 1 And strSrvDate(1) >= FCT結案單電子化啟用日 判斷
                'Modify by Amy 2025/08/19 +bolMailF0202_3 是否寄補看人員,改抓變數,目前只要有run frm210149_1 (T延展不會,也不會有補看人員),預設都寄,除非有設定
                'If (strCmd(0) = MsgText(601) And (field(1) = "CFT" Or field(1) = "CFC" Or (field(1) = "S" And lblCaseField(8) <> "000"))) _
                   Or (field(1) <> "CFT" And field(1) <> "CFC" And field(1) <> "S" And intFCState <> 1) _
                   Or (intFCState = 1 And strSrvDate(1) >= FCT結案單電子化啟用日) Then
                    'Modify by Amy 2021/06/29 +本所案號
                    'Modify by Amy 2025/06/02 +lblCaseField(3) 畫面下一程序-FCT 爭議案內商結
                    strTo = GetF0202_3(field(1), field(2), field(3), field(4), lblCaseField(3))
                If strSrvDate(1) >= FCT結案單電子化啟用日 And strTo <> "" Then
                     'Memo by Amy 葉易雲/洪琬姿為承辦人員且又是補看人員,補看不出現,故不需發-於上面更新Flow002時已設
                     'CF案且補看人員為 葉易雲(78011) 不發
                     If strTo = Pub_GetSpecMan("CFT62") Then
                        bolMailF0202_3 = False
                     '外商案且補看人員為 湘嫺(79020) 不發,因目前FC補看人員都是湘嫺,故以部門判斷即可-Sindy
                     ElseIf PUB_GetST03(strTo) = "F12" Then
                        bolMailF0202_3 = False
                     End If
                End If
                
                If bolMailF0202_3 = True Then
                'end 2025/08/19
                    If strTo <> MsgText(601) Then
                        'Modify By Sindy 2025/6/4
                        strContent = GetEMailContent_Flow(Trim(mPrev01.txtF0301), strSubject, , strContent)
                        PUB_SendMail strUserNum, strTo, "", strSubject, strContent
'                        'Modify by Amy 2018/10/08 +strContent
'                        strContent = strContent & vbCrLf & vbCrLf & 結案單補看人員操作路徑
'                        'Modify by Amy 2018/08/27 內容改抓變數
'                        'Modify by Amy 2024/03/29 T字頭通知承辦及補看人員主旨相同,若承辦與補看人員同一人會造成誤解,故加文字區別
'                        'ex:TF-000750-3-08 結使用宣誓,最近一道A類cp14=承慧與補看人員同一人
'                        PUB_SendMail strUserNum, strTo, "", strSubject & "(審核/補看)", strContent
'                        'end 2018/10/08
                        '2025/6/4 END
                    End If
                End If
                'end 2021/06/28
            '*** End 通知補看人員 ***
            
                'Add by Amy 2025/06/02 外專請款通知,開啟Outlook
                If ChkOutlook.Value = vbChecked Then
                   strExc(9) = Replace(SCp(10), "'", "")
                   'Modify by Amy 2025/07/10 +strOutLookType(依Pub_ChkCloseInvoce函數回傳寄誰)
                   If Pub_CloseOutLook(Me.Name, strF0301, field(1), field(2), field(3), field(4), lblCaseField(8), strExc(9), strOutLookType, strMsg) = False Then
                     If strMsg <> "" Then
                        MsgBox strMsg
                        If InStr(strMsg, "無C類來函掛工程師,不需出草稿") > 0 Then
                           ChkOutlook.Value = vbUnchecked
                        End If
                     Else
                        MsgBox "開啟Outlook失敗,請洽電腦中心!"
                     End If
                   End If
                End If
                'end 2025/06/02
            End If
          End If
          
          '**************************************************************************************************
          If txtCaseField(4) <> "N" Then '指示信
            If txtCaseField(5) = "Y" Then
                bolChk = True
            Else
                bolChk = False
            End If
            
            If field(1) = "CFP" Then
                bolChk = True 'Memo by Lydia 2017/01/25 CFP案預設開Word維護
                'Add by Morgan 2004/4/27
                '當下一程序為208選取時,固定出處理狀況40的定稿
                'Modify by Morgan 2005/5/23 美國 結案原因2,5,10,11例外(跑一般結案)
                'If lblCaseField(3) = "208" Then
                If lblCaseField(3) = "208" And Not (lblCaseField(8).Caption = "101" And InStr("02,05,10,11", Left(Me.cboReason.Text, 2)) > 0) Then
                  strTmp = "40"
                Else
                
                  Select Case Left(Me.cboReason.Text, 2)
                     Case "10" '自行處理
                        strTmp = "31"
                          'Add By Cheng 2003/01/28
                        '若申請國家為日本
                        If lblCaseField(8).Caption = "011" Then
                          '自行處理 36
                          strTmp = "36"
                        End If
                     Case "02" '找其他代理人
                        strTmp = "32"
                          'Add By Cheng 2003/01/28
                        '若申請國家為日本
                        If lblCaseField(8).Caption = "011" Then
                          '找其他代理人 37
                          strTmp = "37"
                        End If
                     Case "05" '已遷移
                        strTmp = "33"
                          'Add By Cheng 2003/01/28
                        '若申請國家為日本
                        If lblCaseField(8).Caption = "011" Then
                          '已遷移 38
                          strTmp = "38"
                        End If
                     'Add by Morgan 2004/10/18
                     Case "09" '自請撤回
                        strTmp = "41"
                     Case "11" '倒閉
                        strTmp = "34"
                          'Add By Cheng 2003/01/28
                        '若申請國家為日本
                        If lblCaseField(8).Caption = "011" Then
                          '倒閉 39
                          strTmp = "39"
                        End If
                     Case Else
                        '一般 30
                        strTmp = "30"
                          'Add By Cheng 2003/01/28
                        '若申請國家為日本
                        If lblCaseField(8).Caption = "011" Then
                          '一般 30
                          strTmp = "35"
                        'Added by Morgan 2015/8/10
                        '印尼
                        'Removed by Morgan 2021/3/29 取消改回用一般--禧佩 Ex:CFP-030225
                        'ElseIf lblCaseField(8).Caption = "017" Then
                        '  '一般 30
                        '  strTmp = "42"
                        'end 2021/3/29
                        End If
                  End Select
                End If
            End If
            
            If field(1) = "P" Then
               '預設一般 30
               strTmp = "30"
               Select Case lblCaseField(3)
                   Case "601" '領證
                      strTmp = "31"
                      
                   Case "605", "606" '年費,維持費
                      '92.10.22 ADD BY SONIA
                      'strTmp = "32"
                      If lblCaseField(8).Caption = "013" And field(8) <> "1" Then
                         strTmp = "35"
                      Else
                         strTmp = "32"
                      End If
                      If Left(Me.cboReason.Text, 2) = "02" Then strTmp = "33" '找其他代理人
                      'Added by Morgan 2016/4/26
                      'FMP年費不續辦原因選10也用02的定稿
                      If m_bolFMP And Left(Me.cboReason.Text, 2) = "10" Then strTmp = "33"
                      
                      '92.10.22 END
                   '92.7.7 ADD BY SONIA
                   Case "408" '面詢
                      strTmp = "34"
                      
                   'Add by Morgan 2007/4/13 香港標準專利批准記錄請求
                   Case "111"
                      strTmp = "36"
                      
                   'Add by Morgan 2007/4/13 PCT實審
                   Case "416"
                      If field(9) = "056" Then
                        strTmp = "37"
                      Else
                        'Modify by Morgan 2007/8/15
                        'strTmp = "30"
                        strTmp = "38"
                      End If
                      
                   'Add by Morgan 2007/10/15
                   Case "119" '進入國家階段
                       If field(9) = "056" Then
                          strTmp = "39"
                       End If
                   
                   'Add by Morgan 2008/4/28 原先為 30 但會與再審的陳述意見(P75999)抓到同一指示信故改為 40
                   Case "503" '行政訴訟
                       If field(9) = "020" Then
                          strTmp = "40"
                       End If
               End Select
            End If
            
            'Modify by Morgan 2006/12/26 CFP指示信&小信封的代理人要一致，其他維持原樣
            If field(1) = "CFP" Then
               'Add by Morgan 2006/6/21
               Dim stFAgent As String
               If UCase(Trim(SCp(44))) <> "NULL" Then
                  stFAgent = Replace(SCp(44), "'", "")
                  stET02 = Replace(SCp(9), "'", "")
               Else
                  stFAgent = GetCP44(cp(9), stET02)
               End If
               StartLetter "14", stET02, strTmp
               'Modify by Morgan 2004/9/27
               'CFP解除期限加印傳真封面
               If bolChk = True Then
                  'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                  'NowPrint stET02, "01", "89", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01  'Memo by Lydia 2017/01/25 抓傳真封面 stLetter
                  'If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                  'end 2018/10/22
                  NowPrint stET02, "14", strTmp, bolChk, strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01 'Memo by Lydia 2017/01/25 傳入傳真封面,所以定稿檔不會存 nowprint ->'有附加資料時不存檔(因不止一份定稿內容會無法編輯)

                  'Added by Morgan 2018/8/20 CFP電子化
                  If m_strAF01 <> "" Then
                     If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                        Set frm1105_1.m_PrevForm = frm210149
                     End If
                     frm1105_1.m_RecNo = m_strAF01
                     'Modify By Sindy 2022/5/11 流水號要足6碼
                     'Val(field(2)) ==> field(2)
                     frm1105_1.m_PdfName = field(1) & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4)) & "." & Mid(SCp(10), 2, 3) & ".DATA.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/8/20
               
                  'Added by Lydia 2017/01/25 CFP案解除期限程序若為EPC 准後已繳指定註冊費案,原僅產生母案之結案指示信,現請同時上結案的子案也一併帶出指示信
                  If strChildList <> "" Then
                     'Added by Morgan 2018/8/22 CFP電子化
                     If m_boleOrderLetter And bolChk Then
                        bolChk = False
                        MsgBox "子案指示信請至待處理區作業！", vbExclamation
                     End If
                     'end 2018/8/22
                     
                     tmpArr = Empty
                     tmpArr = Split(strChildList, ",")
                     For i = 0 To UBound(tmpArr)
                        If Trim(tmpArr(i)) <> "" Then
                           If m_boleOrderLetter Then m_strChildAF01 = Trim(tmpArr(i)) 'Added by Morgan 2018/8/20 CFP電子化
                           
                           stLetter = ""
                           'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                           'NowPrint Trim(tmpArr(i)), "01", "89", False, strUserNum, , , True, stLetter, , , , , , , , , m_strChildAF01
                           'If m_strChildAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/31
                           'end 2018/10/22
                           NowPrint Trim(tmpArr(i)), "14", strTmp, bolChk, strUserNum, , stLetter, , , , , , , , , , , m_strChildAF01
                        End If
                     Next i
                  End If
                  'end 2017/01/25
               Else
                  'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                  'NowPrint stET02, "01", "89", bolChk, strUserNum, 0, , , , , , , , , m_strAF01
                  'end 2018/10/22
                  NowPrint stET02, "14", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
               End If
               
               'Removed by Morgan 2012/7/12 取消--禧佩
               'StrSQLa = "Select FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,FA01||FA02 From FAGENT WHERE FA01='" & Left(stFagent, 8) & "' AND FA02='" & Right(stFagent, 1) & "'"
               'rsA.CursorLocation = adUseClient
               'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               'If rsA.RecordCount > 0 Then
               '  If MsgBox("代理人名稱(中)：" & rsA.Fields(0).Value & Chr(10) & Chr(13) & _
               '            "　　　　　(英)：" & rsA.Fields(1).Value & Chr(10) & Chr(13) & _
               '            "　　　　　(日)：" & rsA.Fields(2).Value & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
               '            "是否列印代理人小信封？", vbExclamation + vbYesNo) = vbYes Then
               '     '列印地址條
               '     'Modify by Morgan 2006/10/17 改Call公用函數
               '     'PrintCase "" & rsA.Fields(3).Value
               '     PUB_PrintCase "" & rsA.Fields(3).Value
               '  End If
               'End If
               'If rsA.State <> adStateClosed Then rsA.Close
               'Set rsA = Nothing
               
            Else
            
               'Add by Morgan 2007/4/11 香港標準專利批准記錄請求
               If lblCaseField(3) = "111" Then
                  '抓香港案最後發文有代理人的收文號
                  strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE CP09<'C' AND CP01='" & field(1) & "' AND CP02='" & field(2) & "' AND CP03='" & field(3) & "' AND CP04='" & field(4) & "' AND CP27>0 AND CP44 IS NOT NULL ORDER BY CP27 DESC"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     stET02 = RsTemp.Fields(0)
                  Else
                     stET02 = field(1) & field(2) & field(3) & field(4) & "&111"
                  End If
               'Modified by Morgan 2016/10/28 補充說明(206)用一般指示信
               'ElseIf cp(9) > "C" Then
               ElseIf cp(9) > "C" And lblCaseField(3) <> "206" Then
               'end 2016/10/28
                  stET02 = cp(43)
               Else
                  stET02 = cp(9)
               End If
               If stET02 <> "" Then
                  StartLetter "14", stET02, strTmp
                  'Modified by Morgan 2015/11/3 指示信電子化
                  If m_boleOrderLetter Then
                     NowPrint stET02, "14", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strAutoNum
                     If bolChk = True Then
                        If UCase(mPrev01.Name) = UCase("frm210149_1") Then
                           Set frm1105_1.m_PrevForm = frm210149
                        End If
                        frm1105_1.m_RecNo = strAutoNum
                        'Modify By Sindy 2022/5/11 流水號要足6碼
                        'Val(field(2)) ==> field(2)
                        frm1105_1.m_PdfName = field(1) & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4)) & "." & Mid(SCp(10), 2, 3) & ".DATA.PDF"
                        frm1105_1.Show
                     End If
                  Else
                     NowPrint stET02, "14", strTmp, bolChk, strUserNum, 0
                  End If
                  'end 2015/11/3
               End If
            End If
            
          End If
          '92.1.7 add by sonia
          '**************************************************************************************************
          If txtCaseField(9) <> "N" Then '通知函
             If txtCaseField(10) = "Y" Then
                bolChk = True
             Else
                bolChk = False
             End If
             NowPrint cp(9), "14", "00", bolChk, strUserNum, 0
          End If
          '92.1.7 add by sonia
            'Modify By Cheng 2002/11/13
'          '911106 nick transation
'          cnnConnection.CommitTrans
          
          'Add by Morgan 2007/1/22
          If bolMail = True Then
'            MsgBox "收件者：" & GetPrjSalesNM(stCP13) & vbCrLf & vbCrLf & _
'                   "主　旨：" & lblCaseField(0).Caption & "案尚有收文號【" & stCP09 & "】未發文故無法閉卷！" & vbCrLf & vbCrLf & _
'                   "內　容：無", vbInformation
            'Modify by Amy 2020/09/07 內文增加文字,增加發mail 給未放文那道之A類承辦人
            PUB_SendMail strUserNum, stCP13 & ";" & stCP14, stCP09, lblCaseField(0).Caption & "案尚有收文號【" & stCP09 & "】未發文故暫時無法閉卷，請承辦確認後通知程序續行！"
          End If
          'end 2007/1/22
          
          'Added by Morgan 2022/10/17
          'FMP大陸閉卷請款函
          If m_bolFMP And field(9) = "020" And SCp(10) = "'913'" Then
             strUserNum = strFMPNum
             StartLetter2 "23", cp(9), "51"
             NowPrint cp(9), "23", "51", False, strUserNum
             strUserNum = strUser1Num
          End If
          'end 2022/10/17
          
          'Added by Lydia 2023/06/09 當寰華案在key閉卷按確認時，請判斷是否有相關香港案及澳門案未不續辦/閉卷，若有則發mail
          If m_bolFMP2 = True And txtCaseField(1) = "Y" And lblCaseField(8) = "020" Then
             'Modified by Lydia 2023/06/28 傳入案件性質SCp(10)
             'Modified by Lydia 2025/01/02 去掉案件性質SCp(10)的單引號Replace
             Call PUB_CloseMailto013044("1", field(1), field(2), field(3), field(4), Replace(SCp(10), "'", ""))
          End If
          Call PUB_SendMailCache
          'end 2023/06/09
          'Add by Amy 2025/08/19+有請款資料彈訊息詢問是否輸請款單
          If bolInvoice = True Then
            'Modify by Amy 2025/10/20 有請款項目直接開請款單輸入不詢問-薛經理
'            intI = MsgBox("開啟請款單輸入及Outlook草稿？" & vbCrLf & _
'                        "是：開啟請款單輸入及Outlook草稿" & vbCrLf & _
'                        "否：回待處理區列表", vbYesNo + vbDefaultButton2 + vbQuestion)
'            If intI = vbYes Then
               mPrev01.QueryData '更新前畫面資料
               mPrev01.SetButtonEnable (False) '鎖住前畫面按鈕
               mPrev01.SSTab1.Tab = 1
               Screen.MousePointer = vbHourglass
               'Modify by Amy 2025/11/11 原只帶第一個畫面,改外商有輸請款項目前3碼與未付款案件性質的金額相符,直接帶入第二個畫面
               strExc(9) = Replace(SCp(10), "'", "")
               bolOpen21H0Ok = Pub_Open21H0(strF0301, Me.Name, mPrev01, field(1), field(2), field(3), field(4), strMsg, strExc(9), txtCaseField(1))
               If bolOpen21H0Ok = False Then
                  MsgBox strMsg, vbExclamation
                  If InStr(strMsg, "通知電腦中心") > 0 Then mPrev01.SetButtonEnable (True) '開放前畫面按鈕
               End If
               '開啟Outlook草稿
               If Pub_CloseOutLook_T(Me.Name, strF0301, field(1), field(2), field(3), field(4), lblCaseField(8), strExc(9), "", strMsg) = False Then
                  MsgBox "開啟Outlook失敗,請洽電腦中心!"
               End If
               Screen.MousePointer = vbDefault
'            End If
             'end 2025/11/11
          End If
          'end 2025/08/19
'          Screen.MousePointer = vbDefault
          
'          'Modify by Amy 2018/10/08 待處理區結束自動 run下一筆
'          'Modify by Amy 2018/10/12 +frm210149_1自動從待處理區 run下一筆
'          If UCase(mPrev01.Name) = UCase("frm210149") Or UCase(mPrev01.Name) = UCase("frm210149_1") Then
'            tmpBol = fnCancelNowFormAndShowParentForm(Me)
'          Else
'            Unload Me
'          End If
'          Exit Sub
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 0 '結束
         Else
            intLeaveKind = 1 '回前畫面
         End If
'         'Modify by Amy 2018/10/08 待處理區結束自動 run下一筆
'         If UCase(mPrev01.Name) = UCase("frm210149") And intLeaveKind = 1 Then
'            tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Else
'            Unload Me
'         End If
      'Add by Amy 2020/05/21 +取消延展(Memo by Amy 2023/02/14 按鈕改顯示退回智權)
      Case 3
        'Modify by Amy 2025/01/20  再加文字,避免User不知按到此鈕,並鎖住確定鈕
        cmdOK(0).Enabled = False
        Screen.MousePointer = vbHourglass
        'Modify by Amy 2023/02/14 原:確定取消嗎?
        If MsgBox("確定取消結案，退回智權嗎？" & vbCrLf & _
          "否:回前畫面" & vbCrLf & _
          "是:確定取消,退回智權", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
        'end 2025/01/20
            Screen.MousePointer = vbDefault
            cmdOK(0).Enabled = True
            Exit Sub
        End If
        strSql = "Update T102Inform Set ti06='Y' Where ti02='" & strNP01 & "' And ti04='" & strNP22 & "' "
        Pub_SeekTbLog strSql 'Add by Amy 2024/10/04 加入以便知道為何已結案,ti06又設 Y
        cnnConnection.Execute strSql
        'Add by Amy 2024/10/08 DeBug用
        strSubject = GetPrjSalesNM(strUserNum) & "(" & strUserNum & ")操作延展結案" & vbCrLf & _
                                 "於[解除期限]作業，按[退回智權鈕]" & vbCrLf & _
                                 "本所案號：" & lblCaseField(0) & vbCrLf
        PUB_SendMail strUserNum, "A2004", "", "T102Inform.Ti06 已上Y 請確認是否User 按錯", strSubject, , , , , , , , , , , , , True
        strSubject = ""
        'end 2024/10/08
   End Select
   
   'Modify by Amy 2020/05/21 +if 不是按「取消延展」才詢問
   If Index <> 3 Then
        If bolLeave = False Then
           If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
              Screen.MousePointer = vbDefault
              'Cancel = 1
              Exit Sub
           End If
        End If
   End If
   Screen.MousePointer = vbDefault 'Added by Lydia 2024/11/12 Teresa反應不續辦一直跑漏斗
   
   Unload Me
   'Modify by Amy 2025/10/20 +if 結案單要Run 請款單輸入
   If bolInvoice = True And intLeaveKind = 2 Then
      '結案單有Run 請款單輸入,關閉此表單,前畫面(frm210149_1)需保留
   ElseIf intLeaveKind <> 0 Then
     'Add by Lydia 2014/10/14 FMP案
      'frm110101_1.Show
      'Modify By Sindy 2018/10/18 前一畫面:下一筆
      If UCase(mPrev01.Name) = UCase("frm210149_1") Then
         mPrev01.cmdQueryNext_Click
         '搬至cmdQueryNext
'         'Add by Amy 2025/07/31 國外結案單 需開啟frm210149_1的按鈕
'         If intFCState > 0 Then
'            mPrev01.SetButtonEnable (True)
'         End If
      ElseIf UCase(mPrev01.Name) = UCase("frm210149") Then
         'mPrev01.Show
         mPrev01.PubShowNextData
         'Add By Sindy 2018/10/18
         If mPrev01.Tag = "" Then
            mPrev01.Show
         End If
      '2018/10/18 END
      ElseIf intLeaveKind = 2 Then '2.確定
         mPrev01.Show
         'frm110101_1.Cleartxt
         mPrev01.Cleartxt
         'Add by Amy 2018/11/27 FCP領證不續辦時,超過領證期限彈訊息提醒
         If bolShow060318 = True Then
            Unload mPrev01
            ShowFrm060318.Show
         End If
      Else
         mPrev01.Show
      End If
      
   ElseIf intLeaveKind = 0 Then '0.結束
    ' Unload frm110101_1
      'Add By Sindy 2015/1/14 結案單電子化
      'Modify by Amy 2018/10/08 拿掉2018/08/30 UCase("frm210149") 判斷
      'Modify by Amy 2018/08/30 +frm210149 由待處理區做T延展結案
      If UCase(mPrev01.Name) = UCase("frm210149_1") Then
         mPrev01.cmdQueryNext_Click
         'Add by Amy 2025/07/31 國外結案單 需開啟frm210149_1的按鈕
         If intFCState > 0 Then
            mPrev01.SetButtonEnable (True)
         End If
      ElseIf UCase(mPrev01.Name) = UCase("frm210149") Then
         mPrev01.QueryData
         mPrev01.Show
      Else
         Unload mPrev01
      End If
   End If
   Exit Sub
     
CheckingErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub ReadAllData()
Dim i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
'edit by nickc 2007/02/02
'Dim intCaseKind As Integer, strReasonName() As String, strNPTemp(1 To T_NP) As String
'2015/8/10 modify by sonia 取消intCaseKind,因為在最上方已宣告
'Dim intCaseKind As Integer, strReasonName() As String, strNPTemp() As String
Dim strReasonName() As String, strNPTemp() As String
'add by nickc 2007/02/02
ReDim strNPTemp(1 To TF_NP) As String
'Dim strNP01 As String,strNP07 As String,strNP22 As String
'改為表單變數
'Add by Lydia 2014/10/14 FMP案
'Add By Cheng 2003/04/15
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSystem As String 'Add By Sindy 2015/1/13
Dim strTi06 As String, strTi05Item As String 'Add by Amy 2020/05/21
Dim m_CCM04 As String, strRCodeN As String 'Add by Amy 2025/06/02 閉卷原因 代碼/名稱

Nextdate1 = "": Nextdate2 = ""
strTi01 = "" 'Add by Amy 2022/06/20
'Add by Amy 2025/06/02
bolInvoice = False
ChkOutlook.Visible = False
'end 2025/06/02
strOutLookType = "" 'Add by Amy 2025/07/10
'Add by Lydia 2014/10/14 FMP案
'strNP01 = frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0)
'strNP07 = frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 9)
'strNP22 = frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 10)
If mPrev01.Name = "frm110101_3" Then
    'Memo by Amy 2025/07/29 P-135515 FMP 07銷香港註冊期限管制,[外專程序]電子結案單未上線前,從內專程式中「FMP解除期限」操作(請作單號:1031016-01 玲玲)
    '    畫面只是選1.主動補正 2.香港第一階段請求(判斷是否為FMP案,於此程式撰寫後才寫的Function),只是不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
    '    故[外專結案單電子化]上線後,由待處理區進入,程式不需做調整即會預設
    strNP01 = mPrev01.m101_3_NP01
    strNP07 = mPrev01.m101_3_NP07
    strNP22 = mPrev01.m101_3_NP22
    strSystem = mPrev01.txtSystem 'Modify By Sindy 2015/1/13
'Modify By Sindy 2015/1/13
ElseIf UCase(mPrev01.Name) = UCase("frm210149_1") Then
   'Memo by Amy 2025/07/29 P-135515 FMP 07銷香港註冊期限管制,[外專程序]電子結案單上線後,由待處理區進入,程式不需做調整即會預設(說明參看上面mPrev01.Name = "frm110101_3")
    strNP01 = mPrev01.m_F0303
    strNP07 = mPrev01.m_NP07
    strNP22 = mPrev01.m_F0304
    strSystem = mPrev01.m_CP01
    'Add by Amy 2025/06/02
    strF0301 = mPrev01.txtF0301
    intFCState = mPrev01.intFCState
    bolInvoice = mPrev01.bolInvoice
    cmdFile.Caption = "檢視回覆單"
    If intFCState > 0 Then
      cmdFile.Caption = "檢視電子檔"
      strClose = mPrev01.strClose
    End If
    'end 2025/06/02
'2015/1/13 END
'Add by Amy 2018/08/30 延展結案判斷
ElseIf UCase(mPrev01.Name) = UCase("frm210149") Then
    'Modified by Morgan 2019/1/18 欄位順序可能調整,改用變數
    'strNP01 = mPrev01.GRD1.TextMatrix(mPrev01.GRD1.row, 5)
    'strNP07 = mPrev01.GRD1.TextMatrix(mPrev01.GRD1.row, 12)
    'strNP22 = mPrev01.GRD1.TextMatrix(mPrev01.GRD1.row, 13)
    strNP01 = mPrev01.GRD1.TextMatrix(mPrev01.GRD1.row, mPrev01.idxCP09)
    strNP07 = mPrev01.GRD1.TextMatrix(mPrev01.GRD1.row, mPrev01.idxCP10)
    strNP22 = mPrev01.GRD1.TextMatrix(mPrev01.GRD1.row, mPrev01.idxNP22)
    'end 2019/1/19
    strSystem = "T"
Else
    strNP01 = mPrev01.grdDataList.TextMatrix(mPrev01.grdDataList.row, 0)
    strNP07 = mPrev01.grdDataList.TextMatrix(mPrev01.grdDataList.row, 9)
    strNP22 = mPrev01.grdDataList.TextMatrix(mPrev01.grdDataList.row, 10)
    strSystem = mPrev01.txtSystem 'Modify By Sindy 2015/1/13
End If

On Error GoTo ErrHand
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetSystemKind(frm110101_1.txtSystem, intCaseKind, , intWhere) = False Then
'Add by Lydia 2014/10/14 FMP案
'If ClsPDGetSystemKind(frm110101_1.txtSystem, intCaseKind, , intWhere) = False Then
'Modify By Sindy 2015/1/13 結案單電子化
'If ClsPDGetSystemKind(mPrev01.txtSystem, intCaseKind, , intWhere) = False Then
If ClsPDGetSystemKind(strSystem, intCaseKind, , intWhere) = False Then
'2015/1/13 END
   GoTo err1
End If
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 0), cp(), field(), intCaseKind, intWhere) Then
ReDim cp(TF_CP) As String
'Add by Lydia 2014/10/14 FMP案
'cp(9) = frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.row, 0)
cp(9) = strNP01

If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   'Add by Amy 2022/09/05 延展結案 +回覆單 鈕
   If UCase(mPrev01.Name) = UCase("frm210149") Then
        cmdFile.Visible = True
        cmdFile.Caption = "檢視回覆單"
        If PUB_ChkIsReplyFile(cp(1) & cp(2) & cp(3) & cp(4), m_strSaveFiles, , m_strSaveFilesCP09) = True Then
            If m_strSaveFiles <> "" Then
                cmdFile.Enabled = True
            Else
                 cmdFile.Caption = "卷宗區"
            End If
        Else
            cmdFile.Caption = "卷宗區"
        End If
   End If
   'end 2022/09/05
   'Add by Morgan 2010/2/3
   'Modify by Morgan 2010/2/24 要控制P案
   'If Left(cp(12), 1) = "F" Then
   'Modified by Morgan 2021/2/2
   'If Left(cp(12), 1) = "F" And cp(1) = "P" And field(10) <> "000" Then
   '   m_bolFMP = True
   'Else
   '   m_bolFMP = False
   'End If
   m_bolFMP = PUB_ChkIsFMP(field(1), field(2), field(3), field(4), field(9))
   'end 2021/2/2
   'end 2010/2/3
   
   'Added by Lydia 2023/06/09 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, field(1), field(2), field(3), field(4))
   End If
   'end 2023/06/09
   
   'Added by Lydia 2023/07/28 FCP專利連結通知
   If field(1) = "FCP" Then
      m_PA177 = field(177)
   Else
      m_PA177 = ""
   End If
   'end 2023/07/28
   
   If ReadNextProgressData(strNPTemp(), strNP01, strNP07, strNP22) = False Then GoTo err1
   'Add by Morgan 2007/4/10 若進度檔本所案號與下一程序不同時重讀基本檔
   If strNPTemp(2) & strNPTemp(3) & strNPTemp(4) & strNPTemp(5) <> cp(1) & cp(2) & cp(3) & cp(4) Then
      ReDim field(4) As String
      field(1) = strNPTemp(2)
      field(2) = strNPTemp(3)
      field(3) = strNPTemp(4)
      field(4) = strNPTemp(5)
      PUB_ReadCaseData field, intCaseKind, intPWhere
   End If
   'end 2007/4/10
   
   lblCaseField(0) = MergeString(field(1), field(2), field(3), field(4))
   lblCaseField(1) = strNPTemp(13)
    'Modify By Cheng 2003/01/16
    '移至顯示國家後再顯示下一程序名稱(依申請國家抓CPM03 or CPM04)
'   lblCaseField(3) = strNPTemp(7)
   ' 91.08.07  邱小姐說全部改成民國年  nickc
   'If intWhere <> 國外_CF Then
      lblCaseField(4) = ChangeTStringToTDateString(ChangeWStringToTString(strNPTemp(8)))
      lblCaseField(5) = ChangeTStringToTDateString(ChangeWStringToTString(strNPTemp(9)))
   'Else
   '   lblCaseField(4) = ChangeWStringToWDateString(strNPTemp(8))
   '   lblCaseField(5) = ChangeWStringToWDateString(strNPTemp(9))
   'End If
   lblCaseField(6) = cp(14)
   'Modify by Amy 2023/01/31 FCT案,智權人員改可以輸-陳金蓮
   'lblCaseField(7) = strNPTemp(10)
   txtSalesNo = strNPTemp(10)
   txtSalesNo.Tag = txtSalesNo
   txtSalesNo.Locked = True
   If field(1) = "FCT" Then
      txtSalesNo.Appearance = 1
      txtSalesNo.BorderStyle = 1
      txtSalesNo.BackColor = &H80000005
      txtSalesNo.Locked = False
   End If
   Call txtSalesNo_Validate(False)
   'end 2023/01/31
   txtCP64 = cp(64)
   'add by sonia 2017/4/19 外商案件不帶CP64
   If field(1) = "FCT" Or field(1) = "CFT" Or field(1) = "CFC" Or field(1) = "S" Then
      txtCP64 = ""
   End If
   'end 2017/4/19
   If intCaseKind = 顧問 Then
      SetNameToCombo cboCaseName, field(6), "", ""
   Else
      SetNameToCombo cboCaseName, field(5), field(6), field(7)
   End If
   Select Case intCaseKind
                Case 專利
                           lblCaseField(2) = field(26)
                           lblCaseField(8) = field(9)
                           cboMemo.Text = field(91)
                           txtCaseField(1) = field(57)
                           'Add by Morgan 2010/7/15
                           strExc(1) = ""
                           If field(75) <> "" Then
                              PUB_GetAgentName "1", field(75), strExc(1)
                           End If
                           Label2(1) = field(75) & " " & strExc(1)
                           'end 2010/7/15
                Case 商標
                           lblCaseField(2) = field(23)
                           lblCaseField(8) = field(10)
                           cboMemo.Text = field(58)
                           txtCaseField(1) = field(29)
                           'add by nickc 2007/07/11 加入審定號數
                           Label18.Visible = True
                           lblCaseField(9).Visible = True
                           lblCaseField(9) = field(15)
                           'Add by Morgan 2010/7/15
                           strExc(1) = ""
                           If field(44) <> "" Then
                              PUB_GetAgentName "1", field(44), strExc(1)
                           End If
                           Label2(1) = field(44) & " " & strExc(1)
                           'end 2010/7/15
                Case 法務
                           lblCaseField(2) = field(11)
                           cboMemo.Text = field(27)
                           txtCaseField(1) = field(8)
                           'Add by Morgan 2010/7/15
                           strExc(1) = ""
                           If field(22) <> "" Then
                              PUB_GetAgentName "1", field(22), strExc(1)
                           End If
                           Label2(1) = field(22) & " " & strExc(1)
                           'end 2010/7/15
                           lblCaseField(8) = field(15)   'add by sonia 2019/8/14
                Case 顧問
                           lblCaseField(2) = field(5)
                           cboMemo.Text = field(12)
                           txtCaseField(1) = field(9)
                           Label2(1) = "" 'Add by Morgan 2010/7/15
                           lblCaseField(8) = "000"   'add by sonia 2019/8/14
                Case Else
                           lblCaseField(2) = field(8)
                           lblCaseField(8) = field(9)
                           cboMemo.Text = field(18)
                           txtCaseField(1) = field(15)
                           'Add by Morgan 2010/7/15
                           strExc(1) = ""
                           If field(26) <> "" Then
                              PUB_GetAgentName "1", field(26), strExc(1)
                           End If
                           Label2(1) = field(26) & " " & strExc(1)
                           'end 2010/7/15
   End Select
    'Move By Cheng 2003/01/16
    lblCaseField(3) = strNPTemp(7)
   If txtCaseField(1) = "Y" Then txtCaseField(1).BackColor = vbRed: BolFileClose = True
   If Len(Trim(txtCaseField(0))) = 0 Then txtCaseField(0) = ChangeWStringToTString(GetTodayDate)
   If intCaseKind = 專利 Then
      txtCaseField(2) = field(89)
   Else
      txtCaseField(2).Enabled = False
   End If
   
   'modify by sonia 90.11.20
   If lblCaseField(8) < "010" Then
      txtCaseField(4) = "N"
      txtCaseField(4).Enabled = False
      
   'Added by Morgan 2016/8/9
   '解除(放棄專利權)請設定不出指示信 --玲玲
   ElseIf field(1) = "P" And lblCaseField(3) = "429" Then
      txtCaseField(4) = "N"
      txtCaseField(4).Enabled = False
   'end 2016/8/9
   
   Else
      txtCaseField(4).Enabled = True
   End If
   
   'Add by Morgan 2011/4/20 標準專利技術請求預設不出指示信--玲玲
   'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
   'If lblCaseField(3) = "110" Then
   If lblCaseField(3) = "110" Or (m_bolFMP = True And lblCaseField(3) = "203") Then
      txtCaseField(4) = "N"
   End If
      
   'add by sonia 92.1.7
   If field(1) <> "CFT" Then
      txtCaseField(9) = "N"
      txtCaseField(9).Enabled = False
      txtCaseField(10).Enabled = False
   Else
      txtCaseField(9).Enabled = True
      txtCaseField(10).Enabled = True
   End If
    'Modify By Cheng 2003/04/15
'   Label7 = ""
   'edit by nickc 2006/06/22 從 dll 內 copy 出
   'Select Case obj011.CheckChildCaseOrCaseRelation(field())
   Select Case CheckChildCaseOrCaseRelation(field())
                Case 1, 2
                           lblChildCase.Visible = True
                Case 0
                           lblChildCase.Visible = False
                Case -1, -2
                           GoTo err1
   End Select
   
   'Add by Morgan 2004/8/6
   stNP09 = strNPTemp(9)
      
   '92.10.7 ADD BY SONIA 全部不預設下次期限, 改由人工輸入
   '下次期限(本所和法定)
   If strNPTemp(9) <> "" Then
      
      If field(1) <> "P" And field(1) <> "T" Then '92.1.22 add by sonia
         Dim tmpSQL  As String
         Dim tmpRs As New ADODB.Recordset
         tmpSQL = "select cf12,cf28 from casefee where cf01='" & field(1) & "' and cf02='" & lblCaseField(8) & "' and cf03='" & lblCaseField(3) & "' "
         Set tmpRs = New ADODB.Recordset
         tmpRs.CursorLocation = adUseClient
         tmpRs.Open tmpSQL, cnnConnection, adOpenStatic, adLockReadOnly
         If tmpRs.RecordCount <> 0 Then
              If CheckStr(tmpRs.Fields(0).Value) <> "" Then
                   Nextdate2 = ChangeWStringToTString(ChangeWDateStringToWString(DateAdd("d", Val(CheckStr(tmpRs.Fields(0).Value)), ChangeWStringToWDateString(strNPTemp(9)))))
              Else
                  If CheckStr(tmpRs.Fields(1).Value) <> "" Then
                       Nextdate2 = ChangeWStringToTString(ChangeWDateStringToWString(DateAdd("M", Val(CheckStr(tmpRs.Fields(1).Value)), ChangeWStringToWDateString(strNPTemp(9)))))
                  Else
                       Nextdate1 = ""
                       Nextdate2 = ""
                  End If
              End If
         Else
              Nextdate1 = ""
              Nextdate2 = ""
         End If
         Dim strDate(0 To 3) As String
         If Nextdate2 <> "" Then
               strDate(1) = field(1)     '系統別
               strDate(2) = lblCaseField(8) '國家
               strDate(3) = ChangeTStringToWString(Nextdate2)  '下次法定期限
               GetCtrlDT strDate()
               Nextdate1 = ChangeWStringToTString(strDate(0))
         End If
         'strTemp = GetCaseFeeNextDays(cp(1), lblCaseField(8), lblCaseField(3))
         'If strTemp <> "" Then
         '      If intWhere <> 國外_CF Then
         '         strTemp1 = ChangeWStringToWDateString(ChangeTStringToWString(strNPTemp(9)))
         '         strTemp1 = DateAdd("D", Val(strTemp), strTemp1)
         '         Nextdate2 = ChangeWDateStringToTString(strTemp1)
         '         strTemp1 = DateAdd("D", -4, strTemp1)
         '         Nextdate1 = ChangeWDateStringToTString(strTemp1)
         '      Else
         '         strTemp1 = ChangeWStringToWDateString(strNPTemp(9))
         '         strTemp1 = DateAdd("D", Val(strTemp), strTemp1)
         '         Nextdate2 = ChangeWDateStringToWString(strTemp1)
         '         strTemp1 = DateAdd("D", -4, strTemp1)
         '         Nextdate1 = ChangeWDateStringToWString(strTemp1)
         '      End If
         'End If
      Else
      '92.1.17 P領證及年費案件不管制半年
         Nextdate1 = ""
         Nextdate2 = ""
      End If
   End If
   '92.10.7 ADD BY SONIA 全部不預設下次期限, 改由人工輸入
   Nextdate1 = ""
   Nextdate2 = ""
   '92.10.7 END
   If txtCaseField(1) <> "Y" Then
      txtCaseField(6) = Nextdate1
      txtCaseField(7) = Nextdate2
   Else
      txtCaseField(6) = ""
      txtCaseField(7) = ""
   End If
Else
err1:
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
If txtCaseField(1) = "Y" Then
   txtCaseField(6) = ""
   txtCaseField(7) = ""
   txtCaseField(6).Enabled = False
   txtCaseField(7).Enabled = False
Else
   txtCaseField(6) = Nextdate1
   txtCaseField(7) = Nextdate2
   txtCaseField(6).Enabled = True
   txtCaseField(7).Enabled = True
End If
'Add By Cheng 2003/04/15
Me.cboReason.Clear
'Modify by Amy 2025/06/02  +FC結案單電子化,「閉卷原因」改共用
'StrSQLa = "Select * From ReasonofRelief Order By ROR01 "
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    While Not rsA.EOF
'        Me.cboReason.AddItem "" & rsA("ROR01").Value & "--" & rsA("ROR02").Value
'        rsA.MoveNext
'    Wend
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'Modify by Amy 2025/08/25 國內與國外結案單原因代碼相同,但顯示名稱可能不同
'     ex:測式國外結案單操作 CFT-025335 原因15,國內無此代碼
'結案單電子化
strExc(9) = "0"
If UCase(mPrev01.Name) = UCase("frm210149_1") Then
   '以電子結案單顯示之名稱為主 ex:CFT 有智權部和外商的案子
   strExc(9) = intFCState
'非電子化 or 由解除期限進入
Else
   If field(1) = "FCT" Or ((field(1) = "T" Or field(1) = "CFT") And Left(PUB_GetST03(strUserNum), 1) = "F") Or (field(1) = "S" And lblCaseField(8) = "000") Then
      'Nvl(ROR03,ROR02) 商標專有名稱->原名稱
      strExc(9) = "1"
   ElseIf field(1) = "FCP" Or ((field(1) = "P" Or field(1) = "CFP") And Left(PUB_GetST03(strUserNum), 1) = "F") Or field(1) = "FG" Then
      'Nvl(ROR04,ROR02)  專利專有名稱->原名稱
      strExc(9) = "2"
   End If
End If
'end 2025/08/25
Call Pub_SetCloseReason(Val(strExc(9)), Me.Name, Me.cboReason)
'end 2025/06/02

'Add By Sindy 2015/1/14 結案單電子化
If UCase(mPrev01.Name) = UCase("frm210149_1") Then
   'Modify by Amy 2025/06/02 +FC結案單電子化,F0305/F0306 拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      '閉卷原因=結案記錄
      m_CCM04 = mPrev01.m_F0305
      strRCodeN = m_CCM04
      Call Pub_SetCloseReason(intFCState, Me.Name, , strRCodeN)
      If strRCodeN = MsgText(601) Then
         Me.cboReason = m_CCM04
      Else
         Me.cboReason = m_CCM04 & "--" & strRCodeN
      End If
      '專利
      If intFCState = 2 Then
         'Add by Amy 2025/07/15 承辦於結案單勾選「需管制6個月補繳期限」,於畫面上[是否管制下次期限 ] =Y
         If txtCaseField(8).Enabled = True Then
            If Pub_GetField("CloseCaseDetail", "ccd01='" & strF0301 & "' And ccd03='19'", "CCD03", True) = "19" Then
               Me.txtCaseField(8) = "Y"
               Call txtCaseField_KeyPress(8, 89)
            End If
         End If
         '外專人員可勾選「出Outlook草稿」
         If Left(PUB_GetST03(strUserNum), 2) = "F2" Then
            Me.ChkOutlook.Visible = True
            '外專承辦於結案單勾選需請款項目,「出Outlook草稿」預設勾選
            'Modify by Amy 2025/07/10 +strOutLookType
            If Pub_ChkCloseInvoce(Me.Name, strF0301, field(1), field(2), field(3), field(4), , strOutLookType) = True Then
               Me.ChkOutlook.Value = vbChecked
            End If
            'Add by Amy 2025/08/05 外專承辦於結案單勾選[後續准駁簡單報告],此畫面此欄位預設Y
            If ChkCCD03(2, Me.Name, strF0301) = True Then
               Me.txtCaseField(2) = "Y"
            End If
         End If
      End If
   Else
      'Modify By Sindy 2015/2/16
      '結案記錄
      'Modify by Amy 2025/06/16 改抓共用
'      If Val(mPrev01.m_F0305) = 99 Then
'         Me.cboReason.ListIndex = Me.cboReason.ListCount - 1
'      Else
'      '2015/2/16 END
'         Me.cboReason.ListIndex = Val(mPrev01.m_F0305) - 1
'      End If
      m_CCM04 = Left(mPrev01.m_F0305, 2)
      strRCodeN = m_CCM04
      If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         Call Pub_SetCloseReason(intFCState, Me.Name, , strRCodeN)
      Else
         Call Pub_SetCloseReason(Val(strExc(9)), Me.Name, , strRCodeN)
      End If
      If strRCodeN = MsgText(601) Then
         Me.cboReason = m_CCM04
      Else
         Me.cboReason = m_CCM04 & "--" & strRCodeN
      End If
      'end 2025/06/16
   End If
   'end 2025/06/02
   '備註
    Me.txtCP64.Text = Trim(Me.txtCP64.Text) & Trim(mPrev01.m_F0306)
   'add by sonia 2017/4/19 外商案件不帶CP64
   If field(1) = "FCT" Or field(1) = "CFT" Or field(1) = "CFC" Or field(1) = "S" Then
      Me.txtCP64.Text = ""
   End If
   'end 2017/4/19
End If
'2015/1/14 END
'Modify by Amy 2020/05/21 +取消延展鈕
cmdOK(3).Visible = False
'Add by Amy 2018/06/05 T延展案帶結案說明TI05 for  T延展電子化
'Modify by Amy 2022/06/20 +109 大陸商標核准後未發註冊證之前，被異議的案件可能會拖到10年沒結果 ex:T-110674
'Modify by Amy 2025/10/28 +intFCState,外商T案不顯示「退回智權」鈕,因不會寫 T102Inform
If (field(1) = "T" Or field(1) = "TF") And (strNP07 = "102" Or strNP07 = "109" Or strNP07 = "716") And intFCState = 0 Then
     'Modify by Amy 2022/06/20 +strTi01
     Me.txtCP64.Text = GetTi05(strTi06, strTi01) '備註
     strTi05Item = Me.txtCP64.Text
     '切割 ti05 項目為 1-11 帶入「解除期限原因」
     If InStr(strTi05Item, ";") > 0 Then
        strTi05Item = Mid(strTi05Item, 1, Val(InStr(strTi05Item, ";") - 1))
     End If
     'Modify by Amy 2020/05/28 空值會error
     'Modify by Amy 2022/11/03 bug 99其他沒抓到
     If strTi05Item <> MsgText(601) Then strTi05Item = Mid(strTi05Item, 1, Val(InStr(strTi05Item, ".") - 1))
     'If Val(strTi05Item) >= 1 And Val(strTi05Item) <= 11 Then
        Call SetReason(strTi05Item)
     'End If
     If strTi06 <> "Y" Then cmdOK(3).Visible = True
End If
'end 2020/05/21
'Add by Morgan 2005/1/5 加代理人並預設最新發文最大收文號的
'Modified by Morgan 2025/10/17 改判斷非台灣案--先還原,P案會錯(指示信與進度規則不同) Ex:P-118285
If Left(field(1), 2) = "CF" Then
'If lblCaseField(8) <> "000" Then
'end 2025/10/17
   Combo2.Enabled = True
   'Modify by Morgan 2007/1/8 加彼所案號
   'AddAgent Combo2, field, strTemp
   AddAgent Combo2, field, strTemp
   Me.txtCaseField(11) = strTemp
   'end 2007/1/8
Else
   Combo2.Clear
   Combo2.Enabled = False
End If

'Add by Amy 2023/02/14 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
If Pre_ProState = "T" Then
    Call Pub_ChkLock(3, Me.Name, "A", Me.Caption, cp(1) & cp(2) & cp(3) & cp(4))
End If
   
Screen.MousePointer = vbDefault
Exit Sub
ErrHand:
ErrorMsg
Screen.MousePointer = varSaveCursor
Resume Next
End Sub

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Dim strTempName As String
   Cancel = False
   strExc(0) = Combo2
   
   'Removed by Morgan 2025/10/17 已改為判斷是否台灣案，但應該不會觸發可不必檢查
   'If Left(field(1), 2) <> "CF" Then
   '   If strExc(0) <> "" Then
   '      MsgBox "非CF案時，必須空白 !", vbCritical
   '      Combo2 = ""
   '      Cancel = True
   '   End If
   'Else
   'end 2025/10/17
   
      If strExc(0) = "" Then
         'Modified by Morgan 2025/10/17
         'MsgBox "CF案時，不可空白 !", vbCritical
         MsgBox "非台灣案時，不可空白 !", vbCritical
         'end 2025/10/17
         Cancel = True
      Else
         If PUB_GetAgentName(field(1), strExc(0), strTempName) = True Then
            Combo2.Text = strExc(0)
            Label2(7).Caption = strTempName
         Else
            MsgBox "代理人代碼輸入錯誤！", vbExclamation
            Label2(7).Caption = ""
            Cancel = True
         End If
      End If
      If Cancel = False Then
         If PUB_CheckStatus(Combo2.Text) = False Then Cancel = True
      End If
      
   'End If 'Removed by Morgan 2025/10/17
   
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

   Select Case Index
             Case 1
                   If txtCaseField(1) = "Y" Then
                     txtCaseField(6) = ""
                     txtCaseField(7) = ""
                     txtCaseField(6).Enabled = False
                     txtCaseField(7).Enabled = False
                  Else
                     txtCaseField(6) = Nextdate1
                     txtCaseField(7) = Nextdate2
                     txtCaseField(6).Enabled = True
                     txtCaseField(7).Enabled = True
                  End If

             Case 2
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCustomer(lblCaseField(Index), strTemp) Then
                        If ClsPDGetCustomer(lblCaseField(Index), strTemp) Then
                           lblPetitionName = strTemp
                        End If
             Case 3
                        'Modify by Morgan 2006/10/17 加香港,澳門,PCT
                        'If lblCaseField(8) = 大陸國家代號 Then
                        If lblCaseField(8) = 大陸國家代號 Or lblCaseField(8) = "013" Or lblCaseField(8) = "044" Or lblCaseField(8) = "056" Then
                           bolIsChina = True
                        Else
                           bolIsChina = False
                        End If
                        
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                        If ClsPDGetCaseProperty(field(1), lblCaseField(Index), strTemp, bolIsChina) Then
                           lblNextProgress = strTemp
                        End If
             Case 6
                        If GetPrjSales(lblCaseField(Index)) <> lblCaseField(Index) Then
                           lblPromoter = GetPrjSales(lblCaseField(Index))
                        End If
             'Mark by Amy 2023/01/31 智權人員改為可輸入
'             Case 7
'                        If GetPrjSales(lblCaseField(Index)) <> lblCaseField(Index) Then
'                           lblSales = GetPrjSales(lblCaseField(Index))
'                        End If
             Case 8
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
                        If ClsPDGetNation(lblCaseField(Index), strTemp) Then
                           lblNation.Caption = strTemp
                        End If
   End Select
End Sub

Private Sub Form_Activate()
   'Added by Lydia 2018/03/16 避免重複執行
   If bolActive Then
      Exit Sub
   Else
      bolActive = True
   End If
   'end 2018/03/16
   
   BolFileClose = False
   ReadAllData
   'ReadAllData 有設txtCaseField(1) = "Y",會設顏色,於此會跳離開
   If txtCaseField(1) = "Y" Then
      MsgBox ("此案號已閉卷！！")
      bolLeave = True
      cmdok_Click (1)
      Exit Sub
   End If
   If UCase(mPrev01.Name) = UCase("frm210149_1") And strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState > 0 Then
      'P案寰華 or 外商案,預帶前畫面是否閉卷
      txtCaseField(1) = strClose
   End If
   
   'Added by Morgan 2015/11/3 指示信電子化
   'P非臺灣案指示信都要彈修改畫面來確認送判的內容
   'Modified by Morgan 2015/12/15 外專程序除外
   'Modified by Morgan 2018/8/16 +CFP電子化
   If (field(1) = "P" Or (field(1) = "CFP" And strSrvDate(1) >= CFP指示信電子化啟用日)) And field(9) <> "000" And Left(Pub_StrUserSt03, 1) <> "F" Then
      txtCaseField(5) = "Y"
      txtCaseField(5).Enabled = False
   End If
   'end 2015/11/3
   Call GetNextpro
   'Add by Amy 2018/11/27 判斷有FCP或 P台灣案已有領證不續辦進度,則不可輸管制下次期限
   txtCaseField(6).Enabled = True: txtCaseField(7).Enabled = True: txtCaseField(8).Enabled = True
   If ((field(1) = "P" And field(9) = "000") Or field(1) = "FCP") And lblCaseField(3) = "601" Then
       bolHas907 = ChkHas907
       If bolHas907 = True Then
           txtCaseField(6).Enabled = False
           txtCaseField(7).Enabled = False
           txtCaseField(8).Enabled = False '是否管制下次期限 於k
       End If
   End If
   'add by sonia 2020/6/10 香港111標準專利批准記錄請求的不續辦預設閉卷並提醒
   If field(1) = "P" And lblCaseField(8) = "013" And lblCaseField(3) = "111" Then
      txtCaseField(1) = "Y"
      MsgBox ("香港第二階段不續辦，本案已預設閉卷！！")
   End If
   'end 2020/6/10
   'Add by Amy  2023/02/17 FCT開放智權欄可輸,進入時游標設至智權欄-湘�A
   If txtSalesNo.Locked = False Then
        txtSalesNo.SetFocus
        txtSalesNo_GotFocus
   End If
   
   'Added by Morgan 2023/12/4
   '有未發文程序不可閉卷--郭
   If field(1) = "P" Or field(1) = "PS" Or field(1) = "CFP" Or field(1) = "CPS" Then
      strExc(0) = "select * from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and  cp04='" & field(4) & "' and cp158=0 and cp159=0 and cp12 not like 'F%'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtCaseField(1).Enabled = False
      End If
      
      'Added by Morgan 2024/12/3
      '大陸一案兩請之新型案，在處理年費不續辦時，倘若發明案未不續辦或閉卷，則新型年費僅能不續辦，不能閉卷--郭
      If field(1) = "P" And field(9) = "020" And field(8) = "2" And lblCaseField(3) = "605" And txtCaseField(1).Enabled Then
         If PUB_IsDualApplyCom(field, strExc) = True Then
            txtCaseField(1).Enabled = False
         End If
      End If
      'end 2024/12/3
      
   End If
   'end 2023/12/4
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   'Add by Amy 2018/10/17 拿掉結束鈕避免全部Form 被關掉
   'If UCase(TypeName(mPrev01)) = UCase("frm210149_1") Or UCase(TypeName(mPrev01)) = UCase("frm210149") Then
   '    cmdOK(2).Visible = False
   'End If
   'Memo by Amy 2025/08/05 將[不續辦但准通知] 改為[後續准駁簡單報告]
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If bolLeave = False Then
'   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
'      Cancel = 1
'   End If
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   PUB_SendMailCache 'Add by Morgan 2009/10/15
   
'   If intLeaveKind <> 0 Then
'     'Add by Lydia 2014/10/14 FMP案
'      'frm110101_1.Show
'      mPrev01.Show
'      If intLeaveKind = 2 Then
'         'frm110101_1.Cleartxt
'         mPrev01.Cleartxt
'      'Modify By Sindy 2018/10/18
'      Else
'         mPrev01.PubShowNextData
'      End If
'      '2018/10/18 END
'   ElseIf intLeaveKind = 0 Then
'    ' Unload frm110101_1
'      'Add By Sindy 2015/1/14 結案單電子化
'      'Modify by Amy 2018/10/08 拿掉2018/08/30 UCase("frm210149") 判斷
'      'Modify by Amy 2018/08/30 +frm210149 由待處理區做T延展結案
'      'Mark by Amy 2018/10/12
''      If UCase(mPrev01.Name) = UCase("frm210149_1") Then
''         frm210149.Hide
''         frm210149.QueryData
''         frm210149.Show
''      End If
'      '2015/1/14 END
'      'Modify By Sindy 2018/10/17
'      If UCase(mPrev01.Name) = UCase("frm210149") Then
'         mPrev01.Show
'      Else
'      '2018/10/17 END
'         Unload mPrev01
'      End If
'   End If
   ShowEditForm 'Added by Morgan 2018/8/22
   'Add by Amy 2023/02/14 內商人員共用待處理區,避免同時處理同一筆資料,造成後續資料有問題
   If UCase(mPrev01.Name) = UCase("frm210149_1") Then '不是T延展,鎖定記錄寫於待處理區(待處理區全選資料,需於此處刪)
        Call Pub_ChkLock(3, mPrev01.Name, "D", , cp(1) & cp(2) & cp(3) & cp(4))
   ElseIf Pre_ProState = "T" Then
        Call Pub_ChkLock(3, Me.Name, "D", , cp(1) & cp(2) & cp(3) & cp(4))
   End If
   Pre_ProState = ""
   'end 2023/02/14
   intFCState = Empty 'Add by Amy 2025/06/02
   'Set mPrev01 = Nothing
   Set frm110101_2 = Nothing
End Sub

Private Sub txtCaseField_Change(Index As Integer)
If Index = 1 Then
   'Add By Sindy 2016/11/9
   'Mark by Amy 2018/09/05 都於確定時詢問是否閉卷
'   'Modify by Amy 2018/08/29 +Pub_StrUserSt03= "P12"
'   If txtCaseField(1) = "Y" And txtCaseField(1).Tag = "不可閉卷" And Pub_StrUserSt03 = "P12" Then
'      MsgBox "下一程序尚有未續辦案件不可閉卷！", vbCritical
'      txtCaseField(1) = ""
'      Exit Sub
'   End If
'   '2016/11/9 END
   'Add by Amy 2018/09/05 +判斷還有別張結案單未結則不可閉卷
   If txtCaseField(1) = "Y" And ChkOtherCloseNo = True Then
      MsgBox "尚有其他結案單未結，不可閉卷！", vbCritical
      txtCaseField(1) = ""
      Exit Sub
   End If
   'end 2018/09/05
   If txtCaseField(1) = "Y" Then
      txtCaseField(6) = ""
      txtCaseField(7) = ""
      txtCaseField(6).Enabled = False
      txtCaseField(7).Enabled = False
   Else
      txtCaseField(6) = Nextdate1
      txtCaseField(7) = Nextdate2
      txtCaseField(6).Enabled = True
      txtCaseField(7).Enabled = True
   End If
End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
             Case 1, 2, 4, 5, 9, 10
                  KeyAscii = UpperCase(KeyAscii)
                  
             'Add by Morgan 2004/8/6
             Case 8 '是否管制下次期限
                  KeyAscii = UpperCase(KeyAscii)
                  '是否管制下次期限=Y
                  If KeyAscii = 89 Then
                     If stNP09 <> "" Then
                        'Modify by Amy 2021/03/04 商標會抓錯,因商標field(9)=商品類別,不是申請國家
                        If GetNextCtrlDate(field(1), lblCaseField(8), lblCaseField(3), stNP09, Nextdate1, Nextdate2) = True Then
                           txtCaseField(6).Text = ChangeWStringToTString(Nextdate1)
                           txtCaseField(7).Text = ChangeWStringToTString(Nextdate2)
                        End If
                     End If
                  End If
                  
End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
   txtCaseField_GotFocus (Index)
End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           '2010/8/3 加val
                           If Val(txtCaseField(intIndex)) <= Val(GetTaiwanTodayDate) Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(8002)
                           End If
                         End If
             Case 1
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case 2, 5, 10
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case 4, 9
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 7
                        If txtCaseField(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              '2010/8/3 加val
                              If Val(txtCaseField(6)) <= Val(txtCaseField(7)) Then
                                 If CheckReKey(txtCaseField(intIndex)) Then
                                    CheckKeyIn = 1
                                 Else
                                    'Modify By Cheng 2002/11/12
'                                    CheckKeyIn = 0
                                    CheckKeyIn = 1
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        ElseIf txtCaseField(6) <> "" Then
                           ShowMsg MsgText(1033)
                            'Modify By Cheng 2002/11/12
'                           CheckKeyIn = 0
                        Else
                           CheckKeyIn = 1
                        End If
             Case 8
'                       strExc(0) = "select ror02 from reasonofrelief where ror01='" & txtCaseField(8) & "'"
'                       intI = 1
'                       Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                       If intI = 1 Then
'                           If IsNull(rsTemp.Fields(0)) Then
'                              Label7 = ""
'                           Else
'                              Label7 = rsTemp.Fields(0).Value
'                           End If
'                           CheckKeyIn = 1
'                       Else
'                           Label7 = ""
'                           MsgBox "解除期限原因代號錯誤，請重新輸入 !", vbCritical
'                       End If
                        'Add by Morgan 2004/8/6
                        CheckKeyIn = 1
             Case 11
               If CheckLengthIsOK(txtCaseField(11).Text, txtCaseField(11).MaxLength) = True Then
                  CheckKeyIn = 1
               End If
             Case Else
               CheckKeyIn = 1
End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   TextInverse txtCaseField(Index)
   If Index = 3 Then
      'edit by nickc 2007/06/06 切換輸入法改用API
      'txtCaseField(Index).IMEMode = 1
      OpenIme
   Else
      'edit by nickc 2007/06/06 切換輸入法改用API
      'txtCaseField(Index).IMEMode = 2
      CloseIme
   End If
End Sub
'讀取下一程序檔是否存在
'Public Function ReadNextProgressData(ByRef np() As String, ByVal strNP01 As String, ByVal strNP07 As String, ByVal strNP22 As String) As Boolean
'Add by Lydia 2014/10/14 FMP案
Public Function ReadNextProgressData(ByRef np() As String, ByVal tmpNP01 As String, ByVal tmpNP07 As String, ByVal tmpNP22 As String) As Boolean
  
  Dim strSql As String
  Dim rsRecordset As New ADODB.Recordset
  Dim i As Integer
  Dim intWhere As Integer
  'Add by Lydia 2014/10/14 FMP案
'   strSql = "SELECT * FROM NEXTPROGRESS WHERE " & _
            "NP01='" & strNP01 & "'" & _
            " AND NP07 ='" & strNP07 & "'" & _
            " AND NP22 ='" & strNP22 & "'"

   strSql = "SELECT * FROM NEXTPROGRESS WHERE " & _
            "NP01='" & tmpNP01 & "'" & _
            " AND NP07 ='" & tmpNP07 & "'" & _
            " AND NP22 ='" & tmpNP22 & "'"
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection, adOpenStatic
If rsRecordset.RecordCount > 0 Then
   For i = 0 To TF_NP - 1 'edit by nickc 2007/02/02 T_NP - 1
          np(i + 1) = IIf(IsNull(rsRecordset.Fields(i)), "", rsRecordset.Fields(i))
   Next
   If ClsPDGetSystemKind(np(2), , , intWhere) = True Then
      '91.08.07   邱小姐說畫面上都用民國年   nick
      'If intWhere <> 國外_CF Then
         'np(8) = ChangeWStringToTString(np(8))
         'np(9) = ChangeWStringToTString(np(9))
      'End If
   Else
      ReadNextProgressData = False
      Exit Function
   End If

Else
   ReadNextProgressData = False
   Exit Function
End If

rsRecordset.Close
ReadNextProgressData = True

End Function

'Modify by Morgan 2006/12/26 加收文號參數
'取得代理人
Private Function GetCP44(strCP09 As String, ByRef p_CP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCP44 = ""
StrSQLa = "Select * From CaseProgress Where CP09 ='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
'若有資料
If rsA.RecordCount > 0 Then
    GetCP44 = "" & rsA("CP44").Value
    p_CP09 = strCP09
    '若無代理人, 再抓相關總收文號的代理人
    If GetCP44 = "" Then
        p_CP09 = "" & rsA("CP43").Value
        StrSQLa = "Select * From CaseProgress Where CP09 ='" & rsA("CP43").Value & "' "
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            GetCP44 = "" & rsA("CP44").Value
            
        End If
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 2003/01/16
'傳定稿例外欄位
Private Sub StartLetter(ByVal ET01 As String, ByVal strReceiveNo As String, ByVal ET03 As String)
Dim strTxt(1 To 5) As String, strTemp As String
Dim ii As Integer
    
    ii = 1
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    
    If field(1) = "P" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(cp(1), cp(10), strTemp, bolIsChina) Then
      If ClsPDGetCaseProperty(cp(1), cp(10), strTemp, bolIsChina) Then
      End If
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','案件性質分類','" & strTemp & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下一程序名稱','" & Me.lblNextProgress.Caption & "')"
      ii = ii + 1
      
      'Added by Morgan 2016/11/17
      '年費結案代理人不是案件代理人時指示信不要帶轉寄官方文件段落
      If lblCaseField(3) = "605" Then
         strExc(0) = "select substr(cp27||cp44,9) from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10<>'605' and cp10<>'421' and cp09<'B' and cp27>0 and cp44 is not null"
         strExc(0) = strExc(0) & " union select substr(cp27||cp44,9) from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='605' and cp27>0 and cp44 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 1 Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','非案件代理人不印','♀')"
               ii = ii + 1
            End If
         End If
      End If
      'end 2016/11/17
    End If
    
    If ii = 1 Then Exit Sub
    'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(ii, strTxt) Then
    If Not ClsLawExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
End Sub

'Add by Morgan 2004/8/6
'取得下次管制期限
Private Function GetNextCtrlDate(ByVal stCF01 As String, ByVal stCF02 As String, ByVal stCF03 As String, ByVal stNP09 As String, ByRef stDate1 As String, ByRef stDate2 As String) As Boolean

   Dim stDate(0 To 3) As String
On Error GoTo ErrHnd
   stDate1 = "": stDate2 = ""
   strSql = "select cf12,cf28 from casefee where cf01='" & stCF01 & "' and cf02='" & stCF02 & "' and cf03='" & stCF03 & "'"
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      If Val("" & adoRecordset.Fields("cf12")) > 0 Then
         stDate2 = ChangeWDateStringToWString(DateAdd("d", Val(adoRecordset.Fields("cf12")), ChangeWStringToWDateString(stNP09)))
      ElseIf Val("" & adoRecordset.Fields("cf28")) > 0 Then
         stDate2 = ChangeWDateStringToWString(DateAdd("m", Val(adoRecordset.Fields("cf28")), ChangeWStringToWDateString(stNP09)))
         'Modify by Morgan 2004/12/15 原法定若為月底則延期後也要是月底
         PUB_LastDayConvert stNP09, stDate2
      End If
      If stDate2 <> "" Then
            stDate(1) = stCF01     '系統別
            stDate(2) = stCF02 '國家
            stDate(3) = stDate2  '下次法定期限
            'Add by Morgan 2010/2/3 FMP案提前10天
            If m_bolFMP Then
               stDate(0) = CompDate(2, -10, stDate2)
            'Added by Morgan 2018/10/3 非FMP非台灣的專利案年費及實審的所期限也改法限-10天
            ElseIf stCF01 = "P" And stCF02 <> "000" And InStr("416,605,606", stCF03) > 0 Then
               stDate(0) = CompDate(2, -10, stDate2)
            'end 2018/10/3
            Else
            'end 2010/2/3
               GetCtrlDT stDate()
            End If
            stDate1 = stDate(0)
            'Modify by Morgan 2010/6/22 控制FCP不用抓上班日
            If cp(1) <> "FCP" Then 'Memo by Lydia 2020/07/13 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天;因為FCP案有輸入非工作日的需求,所以排除
               stDate1 = PUB_GetWorkDay1(stDate1, True) 'Add by Morgan 2010/2/3
            End If
            GetNextCtrlDate = True
      End If
      
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
End Function

          
'Add by Lydia 2015/02/10 判斷是否發證
'Private Function ReadPA21EndModCash(ByRef cp() As String) As Boolean
'Dim bStr01 As String
'Dim rsB As New ADODB.Recordset
'
'ReadPA21EndModCash = False
'bStr01 = "Select PA21 From PATENT where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04='" & cp(4) & "' "
'intI = 1
'Set rsB = ClsLawReadRstMsg(intI, bStr01)
'If intI = 1 Then
'   If rsB!pa21 > 0 Then
'      '發證後有再收文（Ａ類），於結案時自動上結餘日
'       bStr01 = "select cp09 from caseprogress where cp05>= " & rsB!pa21 & " and cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
'                "and substr(cp09,1,1) = 'A' "
'          intI = 1
'          Set rsB = ClsLawReadRstMsg(intI, bStr01)
'          If intI = 1 Then ReadPA21EndModCash = True
'   Else
'      '發證前未計算過結餘，結案時自動上結餘日
'       bStr01 = "select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
'                "and cp109 > 0 "
'          intI = 1
'          Set rsB = ClsLawReadRstMsg(intI, bStr01)
'          If intI = 0 Then ReadPA21EndModCash = True
'   End If
'End If
'
'End Function

'Added by Morgan 2015/5/18
'FCP台灣新型年費解除期限一案兩請提醒
'Remove by Lydia 2019/06/21 改成模組Pub_ChkFCPDualCaseBYcancel
'Private Sub CheckFCPDualCase()
'
'   If cp(10) = "605" And field(1) = "FCP" And field(9) = "000" And field(8) = "2" Then
'      '若發明案尚未審定或核駁且未閉卷時，提醒使用者
'      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)" & _
'         " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "'" & _
'         " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & field(1) & "' and cm06='" & field(2) & "' and cm07='" & field(3) & "' and cm08='" & field(4) & "') X" & _
'         ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa08='1' AND pa57 is null and (pa16 is null or pa16='2')"
'
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strExc(2) = "此案為一案兩請且發明案 " & RsTemp(0) & " 尚未審定，請將卷宗交業務承辦告知客戶新型專利權若因未繳年費而當然消滅者，則將不予專利！"
'         MsgBox strExc(2), vbExclamation
'      End If
'   End If
'
'End Sub
'end 2019/06/21

'Added by Lydia 2017/01/25 新增B類收文模組化
'Modified by Morgan 2018/8/20 +cCP45
Private Function GetInsBCP(Optional ByVal cCP03 As String, Optional ByVal cCP04 As String, Optional ByVal cCP09 As String, Optional ByVal cCP43 As String, Optional ByVal cCP44 As String, Optional ByRef cCP45 As String) As String
Dim intC As Integer
Dim strCP44 As String, strCP45 As String 'Added by Lydia 2017/12/01

    GetInsBCP = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20,cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30,cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50,cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60,cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76,cp77,cp78,cp79,cp140) values "
    GetInsBCP = GetInsBCP & " ("
    For intC = 1 To UBound(SCp)
        Select Case intC
            Case 3
               If cCP03 <> "" Then
                  GetInsBCP = GetInsBCP & CNULL(cCP03) & ","
               Else
                  GetInsBCP = GetInsBCP & SCp(intC) & ","
               End If
            Case 4
               If cCP04 <> "" Then
                  GetInsBCP = GetInsBCP & CNULL(cCP04) & ","
               Else
                  GetInsBCP = GetInsBCP & SCp(intC) & ","
               End If
            Case 9
               If cCP09 <> "" Then
                  GetInsBCP = GetInsBCP & CNULL(cCP09) & ","
               Else
                  GetInsBCP = GetInsBCP & SCp(intC) & ","
               End If
            Case 43
               If cCP43 <> "" Then
                  GetInsBCP = GetInsBCP & CNULL(cCP43) & ","
               Else
                  GetInsBCP = GetInsBCP & SCp(intC) & ","
               End If
            'Added by Lydia 2017/12/01 CF代理人+彼所案號
            Case 44, 45
               If cCP04 <> "" And cCP04 <> "00" And strCP44 = "" And strCP45 = "" Then
                  If PUB_GetCP44(cp(1), cp(2), cCP03, cCP04, strCP44, strExc(5), strCP45) = True Then
                      '抓子案的最後代理人/彼所案號 (EPC案母案解除期限和上閉卷，子案一併上閉卷和出結案指信，相關總收文號放母案的閉卷收文號，但是CF代理人要代子案的最後代理人　ex.CFP-26333)
                      strExc(4) = CNULL(IIf(intC = 44, strCP44, strCP45))
                  Else
                      strExc(4) = IIf(intC = 44, SCp(44), SCp(45)) '母案
                  End If
               Else
                  If cCP04 = "" Or cCP04 = "00" Then
                     strExc(4) = IIf(intC = 44, SCp(44), SCp(45)) '母案
                  Else
                      '抓子案的最後代理人/彼所案號 (EPC案母案解除期限和上閉卷，子案一併上閉卷和出結案指信，相關總收文號放母案的閉卷收文號，但是CF代理人要代子案的最後代理人　ex.CFP-26333)
                      strExc(4) = CNULL(IIf(intC = 44, strCP44, strCP45))
                  End If
               End If
               GetInsBCP = GetInsBCP & strExc(4) & ","
            'end 2017/12/01
            Case 65, 66, 67, 68, 69, 70 'create id,time
            Case Else
                GetInsBCP = GetInsBCP & SCp(intC) & IIf(intC <> 79, ",", "")
        End Select
    Next intC
    'Modify by Amy 2022/06/20 +商標延展結案,可顯示close.menu
    If UCase(mPrev01.Name) = UCase("frm210149") And (field(1) = "T" Or field(1) = "TF") _
      And (strNP07 = "102" Or strNP07 = "109" Or strNP07 = "716") Then
      GetInsBCP = GetInsBCP & ",'T-" & strTi01 & "'"
    ElseIf UCase(mPrev01.Name) = UCase("frm210149_1") Then
       GetInsBCP = GetInsBCP & ",'" & mPrev01.txtF0301 & "'"
    Else
       GetInsBCP = GetInsBCP & ",null"
    End If
    GetInsBCP = GetInsBCP & ") "
    
    cCP45 = strCP45 'Added by Morgan 2018/8/20
End Function

'Add by Amy 2018/06/05 取得T延展結案說明
'Modify by Amy 2020/05/21 +回傳ti06
'Modify by Amy 2022/06/20 +回傳ti01
Private Function GetTi05(ByRef stTi06 As String, ByRef stTi01 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    GetTi05 = ""
    'Modify by Amy 2022/06/20 +ti01
    strQ = "Select ti01,ti05,ti06 From T102InForm Where ti02='" & cp(9) & "' And ti04='" & strNP22 & "' "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        GetTi05 = "" & RsQ.Fields("ti05")
        stTi06 = "" & RsQ.Fields("ti06")
        stTi01 = "" & RsQ.Fields("ti01")
    End If
    'end 2022/06/20
    RsQ.Close
End Function
'end 2020/05/21

'Add by Amy 2018/08/29 檢查下一程序是否尚有未續辦案件(從frm210149_1搬過來-Sindy 2016/11/9)
Private Sub GetNextpro()
    Dim rsA As New ADODB.Recordset
    Dim stTmp As String, intC As Integer 'Add by Amy 2020/01/11
    
    'Modify by Amy 2018/09/05 +strNpSqlOfNoSalesDuty-外商阿蓮
    'Modify by Amy 2021/01/11 +顯示下一程序案件性質名稱，若有多筆要串起來顯示
'    strSql = "Select np01 From nextprogress" & _
'                " where np02='" & field(1) & "' and np03='" & field(2) & "'" & _
'                " and np04='" & field(3) & "' and np05='" & field(4) & "'" & _
'                " and np06 is null and np22<>" & strNP22 & strNpSqlOfNoSalesDuty
    strSql = "Select Nvl(Decode(" & lblCaseField(8) & ",'000',cpm03,cpm04),np07) as np07N From NextProgress,CaseProPertyMap" & _
                " Where np02='" & field(1) & "' and np03='" & field(2) & "'  and np04='" & field(3) & "' and np05='" & field(4) & "'" & _
                " and np02=cpm01(+) and np07=cpm02(+) and np06 is null and np22<>" & strNP22 & strNpSqlOfNoSalesDuty
    rsA.CursorLocation = adUseClient
    rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    txtCaseField(1).Tag = "可閉卷"
    If rsA.RecordCount > 0 Then
        Do While rsA.EOF = False
            stTmp = stTmp & "/ " & rsA.Fields("np07n")
            If intC Mod 4 = 0 And intC <> 0 Then stTmp = stTmp & vbCrLf
            intC = intC + 1
            rsA.MoveNext
        Loop
        If stTmp <> MsgText(601) Then stTmp = Mid(stTmp, 2)
        txtCaseField(1).Tag = stTmp & "不可閉卷"
    End If
    'end 2021/01/11
    If rsA.State <> adStateClosed Then rsA.Close
End Sub

'Add by Amy 2018/09/05 判斷同案號是否有其他張結案單未結
Private Function ChkOtherCloseNo() As Boolean
    Dim rsA As New ADODB.Recordset
    
    ChkOtherCloseNo = False
    strSql = "Select np01 From Nextprogress" & _
                " where np02='" & field(1) & "' and np03='" & field(2) & "'" & _
                " and np04='" & field(3) & "' and np05='" & field(4) & "'" & _
                " and length(np24)=8 and np06 is null and np22<>" & strNP22
    rsA.CursorLocation = adUseClient
    rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        ChkOtherCloseNo = True
    End If
    If rsA.State <> adStateClosed Then rsA.Close
End Function

'Add by Amy 2018/11/27 是否已有領進不續辦進度
Private Function ChkHas907() As Boolean
    Dim rsA As New ADODB.Recordset
    
    ChkHas907 = False
    strSql = "Select * From CaseProgress Where CP43='" & cp(9) & "' And CP10='907'"
    rsA.CursorLocation = adUseClient
    rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        ChkHas907 = True
    End If
    If rsA.State <> adStateClosed Then rsA.Close
End Function

'Add by Amy 2020/05/21 Ti05有對應項目帶入
Private Sub SetReason(ByVal stItem As String)
    Dim ii As Integer
    For ii = 0 To Me.cboReason.ListCount - 1
        If Left(Me.cboReason.List(ii), 2) = stItem Then
            Me.cboReason = Me.cboReason.List(ii)
            Exit For
        End If
    Next ii
End Sub

'Added by Morgan 2022/10/17
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   Dim strTxt(1 To 2) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, ET02, ET03, strUserNum
   
   j = 0
   '請款函備註
   strExc(0) = PUB_GetDebitNotePS(field(1) & field(2) & field(3) & field(4), "913", field(75), field(26))
   If strExc(0) <> "" Then
      j = j + 1
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & ChgSQL(strExc(0)) & "')"
      j = j + 1
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有請款函備註時不印','♀')"
   End If
   
   If Not ClsLawExecSQL(j, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add by Amy 2023/01/31
Private Sub txtSalesNo_GotFocus()
    TextInverse txtSalesNo
End Sub

Private Sub txtSalesNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesNo_Validate(Cancel As Boolean)
    Dim stST02 As String, stDept As String
    
    If txtSalesNo = MsgText(601) Then Exit Sub
    
    If field(1) = "FCT" Then
        If PUB_GetStaffNameDept(txtSalesNo, stST02, stDept, True, True) = True Then
            lblSales = stST02
            '操作人員為外商時,判斷輸入之智權人員與操作員部門不同彈提醒,可操作(內商會操作FCT爭議案) ex:FCT-049992
            If Left(PUB_GetStaffST15(txtSalesNo, 1), 2) <> Left(Pub_StrUserSt15, 2) And txtSalesNo <> txtSalesNo.Tag And Left(Pub_StrUserSt15, 2) = "F1" Then
                MsgBox "輸入之智權人員非同部門"
            End If
        Else
            lblSales = ""
            Cancel = True
        End If
    Else
        '避免改為可輸入後,會彈訊息,故非 FCT 案照舊抓法
        If GetPrjSales(txtSalesNo) <> txtSalesNo Then
            lblSales = GetPrjSales(txtSalesNo)
        End If
    End If
End Sub
'end 2023/01/31

