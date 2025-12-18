VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030402 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF期限通知函"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8244
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   8244
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2460
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   795
      Left            =   4410
      TabIndex        =   38
      Top             =   2730
      Visible         =   0   'False
      Width           =   3765
      Begin VB.TextBox txtUKMoney 
         Height          =   300
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   7
         Top             =   120
         Width           =   1065
      End
      Begin VB.TextBox txtUKChgMoney 
         Height          =   300
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   11
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox txtUKScore 
         Height          =   300
         Left            =   3090
         MaxLength       =   8
         TabIndex        =   8
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtUKChgScore 
         Height          =   300
         Left            =   3090
         MaxLength       =   8
         TabIndex        =   12
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "英國變更費用："
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   450
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "英國延展費用："
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Left            =   2520
         TabIndex        =   40
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Left            =   2520
         TabIndex        =   39
         Top             =   450
         Width           =   540
      End
   End
   Begin VB.TextBox textCaseFee 
      Height          =   300
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   13
      Top             =   3520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtCScore 
      Height          =   300
      Left            =   3660
      MaxLength       =   8
      TabIndex        =   14
      Top             =   3520
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtChgScore 
      Height          =   300
      Left            =   3660
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3180
      Width           =   690
   End
   Begin VB.TextBox txtScore 
      Height          =   300
      Left            =   3660
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2820
      Width           =   690
   End
   Begin VB.TextBox textConvert 
      Enabled         =   0   'False
      Height          =   264
      Left            =   8400
      MaxLength       =   8
      TabIndex        =   17
      Top             =   3660
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox textChgMoney 
      Height          =   300
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   9
      Top             =   3180
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6240
      TabIndex        =   18
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7200
      TabIndex        =   19
      Top             =   72
      Width           =   912
   End
   Begin VB.TextBox textMoney 
      Height          =   300
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2820
      Width           =   1425
   End
   Begin VB.TextBox textCF15 
      Height          =   300
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2460
      Width           =   732
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2100
      Width           =   6492
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   2
      Top             =   660
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   660
      Width           =   732
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   1
      Top             =   660
      Width           =   1092
   End
   Begin VB.Label lblFee1 
      Caption         =   "費用："
      Height          =   252
      Left            =   240
      TabIndex        =   28
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label lblFee1s 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   2820
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.TextBox txtOldAddr 
      Height          =   300
      Left            =   1560
      TabIndex        =   15
      Top             =   3840
      Width           =   6495
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      MaxLength       =   180
      Size            =   "11456;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtNewAddr 
      Height          =   300
      Left            =   1560
      TabIndex        =   16
      Top             =   4200
      Width           =   6495
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      MaxLength       =   180
      Size            =   "11456;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM06 
      Height          =   285
      Left            =   1560
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1380
      Width           =   6495
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   285
      Left            =   1560
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1020
      Width           =   6495
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   285
      Left            =   1560
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1740
      Width           =   6495
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Caption         =   "新地址："
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   4230
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "原註冊地址："
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   3870
      Width           =   1275
   End
   Begin VB.Label lblCase1 
      Caption         =   "第八及十五條費用："
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   3525
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblCase2 
      Caption         =   "點數："
      Height          =   255
      Left            =   3090
      TabIndex        =   33
      Top             =   3525
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label11 
      Caption         =   "點數："
      Height          =   255
      Left            =   3090
      TabIndex        =   32
      Top             =   3180
      Width           =   600
   End
   Begin VB.Label Label10 
      Caption         =   "點數："
      Height          =   255
      Left            =   3090
      TabIndex        =   31
      Top             =   2820
      Width           =   600
   End
   Begin VB.Label Label9 
      Caption         =   "轉類費用 :"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8070
      TabIndex        =   30
      Top             =   3660
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "變更費用："
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3180
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "下一程序："
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   2460
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "商品類別："
      Height          =   252
      Left            =   240
      TabIndex        =   24
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "案件中文名稱："
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱："
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   21
      Top             =   1740
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   240
      TabIndex        =   20
      Top             =   660
      Width           =   1275
   End
End
Attribute VB_Name = "frm030402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/12/06  改成Form2.0 ; textTM05、textTM06、textTM07、txtOldAddr、txtNewAddr
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_TM05 As String 'Add By Sindy 2010/02/25
Dim m_TM10 As String
Dim m_NP07 As String
Dim m_NP08 As String
Dim m_NP09 As String
Dim m_TM11 As String '申請日
'add by nickc 2005/05/31
Dim oClsPrtForm001 As New ClsPrtForm001
Dim m_NP01 As String
Dim m_NP22 As String 'Add by Morgan 2008/5/19
Dim m_DATE As String  '2006/1/9 ADD BY SONIA 其他日期(定稿用:延展使用宣誓可辦開始日期)
Dim m_TM08 As String, m_TM23 As String 'Add by Morgan 2008/5/19
Dim m_TM24 As String 'Add By Sindy 2020/6/3
Dim m_tm25 As String 'Add By Sindy 2020/6/8
Dim m_102cp25 As String   '2011/9/23 ADD BY SONIA
Dim m_TM21 As String 'add by sonia 2022/10/14
Dim m_cpHave102 As String 'add by sonia 2022/10/14

Private Sub cmdExit_Click()
    'Add By Cheng 2003/03/26
'mvoe to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    'Modify By Cheng 2003/08/20
'    '移至Form_Load時做
''    '刪除暫存資料
'    'edit by nick 2004/10/11
'    PUB_DeleteCaseCloseSheet strUserNum
    Unload Me
End Sub

Private Sub SetInputEntry()
   textTM01.SetFocus
End Sub

Private Sub Clear()
   textTM01 = Empty
   textTM02 = Empty
   textTM03 = Empty
   textTM04 = Empty
   textTM05 = Empty
   textTM06 = Empty
   textTM07 = Empty
   textTM09 = Empty
   textCF15 = Empty
   textCF15_2 = Empty
   textMoney = Empty
   txtOldAddr = Empty 'Add By Sindy 2020/6/3
   txtNewAddr = Empty 'Add By Sindy 2020/6/3
   'Add by Morgan 2008/5/19
   textChgMoney = Empty
   txtScore = Empty
   txtChgScore = Empty
   'end 2008/5/19
   
   'Add by Morgan 2022/4/25
   txtUKMoney = Empty
   txtUKScore = Empty
   txtUKChgMoney = Empty
   txtUKChgScore = Empty
   'end 2022/4/25
   
   Call SetCase105 'Added by Lydia 2016/03/04
   
'edit by nickc 2005/08/25 取消轉類費用
'   textConvert = Empty
End Sub

Private Sub cmdok_Click()
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/15 清除查詢印表記錄檔欄位
      If OnProcess = True Then
         Clear
         SetInputEntry
      End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    textTM05.BackColor = &H8000000F
    textTM06.BackColor = &H8000000F
    textTM07.BackColor = &H8000000F
    textTM09.BackColor = &H8000000F
    textCF15_2.BackColor = &H8000000F
    'Add By Cheng 2003/08/20
    '刪除暫存資料
    PUB_DeleteCaseCloseSheet strUserNum
    If Caption = "馬德里催使用宣誓函" Then
        textCF15.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '列印接洽接案單
    PUB_PrintCaseCloseSheet strUserNum
    'Modify By Cheng 2003/08/20
    '移至Form_Load時做
'    '刪除暫存資料
    'edit by nick 2004/10/11
    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/19
   Set frm030402 = Nothing
End Sub

Private Sub lblFee1_Click()
   If lblFee1.Tag = "Y" Then
      SetOldPrice True
   End If
End Sub

Private Sub lblFee1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblFee1, lblFee1s
End Sub

Private Sub lblFee1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblFee1, lblFee1s
End Sub

Private Sub textChgMoney_GotFocus()
    TextInverse Me.textChgMoney
End Sub

Private Sub textChgMoney_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textChgMoney) = False Then
      If IsNumeric(textChgMoney) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "變更費用請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textChgMoney_GotFocus
      End If
   End If
End Sub

'edit by nickc 2005/08/25 取消轉類費用
'Private Sub textConvert_GotFocus()
'    TextInverse Me.textConvert
'End Sub
'
'Private Sub textConvert_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   If IsEmptyText(textConvert) = False Then
'      If IsNumeric(textConvert) = False Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "轉類費用請輸入數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textConvert_GotFocus
'      End If
'   End If
'End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
         ' 檢查系統類別
      If IsCorrectSysKind(textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textTM01
         Case "CFT":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   End If
EXITSUB:
End Sub

Private Sub textTM04_LostFocus()
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strTit As String
Dim strMsg As String
Dim nResponse As Integer
   
    ' 清除資料
    textTM05 = Empty
    textTM06 = Empty
    textTM07 = Empty
    textTM09 = Empty
    ' 設定本所案號
    m_TM01 = textTM01
    m_TM02 = textTM02
    m_TM03 = textTM03
    If IsEmptyText(m_TM03) = True Then: m_TM03 = "0"
    m_TM04 = textTM04
    If IsEmptyText(m_TM04) = True Then: m_TM04 = "00"
   
    ' 查詢商標基本檔
    strSql = "SELECT * FROM TRADEMARK " & _
             "WHERE TM01 = '" & m_TM01 & "' AND " & _
                   "TM02 = '" & m_TM02 & "' AND " & _
                   "TM03 = '" & m_TM03 & "' AND " & _
                   "TM04 = '" & m_TM04 & "' "
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        ' 案件中文名稱
        m_TM05 = "" 'Add By Sindy 2010/02/25
        If IsNull(rsTmp.Fields("TM05")) = False Then
           textTM05 = rsTmp.Fields("TM05")
           m_TM05 = rsTmp.Fields("TM05") 'Add By Sindy 2010/02/25
        End If
        ' 案件英文名稱
        If IsNull(rsTmp.Fields("TM06")) = False Then
           textTM06 = rsTmp.Fields("TM06")
        End If
        ' 案件日文名稱
        If IsNull(rsTmp.Fields("TM07")) = False Then
           textTM07 = rsTmp.Fields("TM07")
        End If
        ' 商品類別
        If IsNull(rsTmp.Fields("TM09")) = False Then
           textTM09 = rsTmp.Fields("TM09")
        End If
        ' 申請國家
        If IsNull(rsTmp.Fields("TM10")) = False Then
           m_TM10 = rsTmp.Fields("TM10")
        End If
        'Add by Morgan 2008/5/19
        m_TM08 = "" & rsTmp.Fields("TM08")
        m_TM23 = "" & rsTmp.Fields("TM23")
        'end 2008/5/19
        'Add By Sindy 2020/6/3
        m_TM24 = "" & rsTmp.Fields("TM24")
        m_tm25 = "" & rsTmp.Fields("TM25")
        txtOldAddr = m_tm25
        'If Trim(txtOldAddr.Text) <> "" Then txtOldAddr.Enabled = False
        '2020/6/3 END
    Else
        strTit = "資料檢核"
        strMsg = "本所案號不存在"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Me.textTM01.SetFocus
        textTM01_GotFocus
   End If
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Cancel = False
   PUB_LabelActive lblFee1, lblFee1s, False 'Add by Morgan 2008/5/20
   Frame1.Visible = False 'Added by Morgan 2022/4/22
   
   'Add By Sindy 2009/07/15
   ' 設定本所案號
   m_TM01 = textTM01
   m_TM02 = textTM02
   m_TM03 = textTM03
   If IsEmptyText(m_TM03) = True Then: m_TM03 = "0"
   m_TM04 = textTM04
   If IsEmptyText(m_TM04) = True Then: m_TM04 = "00"
   '2009/07/15 End
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      ' 取得案件性質名稱
      If m_TM10 > "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      End If
      If IsEmptyText(textCF15_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
         GoTo EXITSUB:
      End If
        
      strSql = "SELECT * FROM NEXTPROGRESS " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & textCF15 & " AND " & _
                     "(NP06 IS NULL OR NP06 = '' OR NP06 = ' ') "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "沒有符合該下一程序的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
      'Add by Morgan 2008/5/19
      Else
         Call SetCase105 'Added by Lydia 2016/03/04
         SetOldPrice False
         'Added by Morgan 2022/4/26 歐盟延展費要顯示英國相關費用欄位
         If m_TM10 = "239" And textCF15 = "102" Then
            strSql = "SELECT * FROM NEXTPROGRESS " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = '110' AND NP06 IS NULL "
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               Frame1.Visible = True
            End If
         End If
         'end 2022/4/26
      End If
      If rsTmp.State <> adStateClosed Then rsTmp.Close
      Set rsTmp = Nothing
   End If
EXITSUB:
End Sub

' 費用
Private Sub textMoney_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textMoney) = False Then
      If IsNumeric(textMoney) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMoney_GotFocus
      End If
   End If
End Sub

' 檢查資料輸入是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 本所案號
   If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   ' 下一程序不可空白
   If IsEmptyText(textCF15) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入下一程序"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   'Added by Lydia 2016/03/04 +費用判斷
   If IsEmptyText(textMoney) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入" & Replace(lblFee1.Caption, "：", "")
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 列印資料
Private Sub OnPrintData(ByRef rsTmp As ADODB.Recordset)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If OnProcess() = False Then
      strTit = "列印定稿"
      strMsg = "沒有符合條件的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

Private Function OnProcess() As Boolean
Dim strSql As String
Dim rsTmp As ADODB.Recordset
   
   OnProcess = False
   
   'Add by Morgan 2008/5/20
   If textMoney <> "" And txtScore <> "" Then
      If Val(textMoney) < Val(txtScore) * 1000 Then
         MsgBox "點數輸入錯誤！"
         txtScore.SetFocus
         txtScore_GotFocus
         Exit Function
      End If
   End If
   If textChgMoney <> "" And txtChgScore <> "" Then
      If Val(textChgMoney) < Val(txtChgScore) * 1000 Then
         MsgBox "變更點數輸入錯誤！"
         txtChgScore.SetFocus
         txtChgScore_GotFocus
         Exit Function
      End If
   End If
   'end 2008/5/20
   
   'Add by Morgan 2022/4/25
   If txtUKMoney <> "" And txtUKScore <> "" Then
      If Val(txtUKMoney) < Val(txtUKScore) * 1000 Then
         MsgBox "點數輸入錯誤！"
         txtUKScore.SetFocus
         txtUKScore_GotFocus
         Exit Function
      End If
   End If
   If txtUKChgMoney <> "" And txtUKChgScore <> "" Then
      If Val(txtUKChgMoney) < Val(txtUKChgScore) * 1000 Then
         MsgBox "變更點數輸入錯誤！"
         txtUKChgScore.SetFocus
         txtUKChgScore_GotFocus
         Exit Function
      End If
   End If
   'end 2022/4/25
   

   ' 設定本所案號
   m_TM01 = textTM01
   m_TM02 = textTM02
   m_TM03 = textTM03
   If IsEmptyText(m_TM03) = True Then: m_TM03 = "0"
   m_TM04 = textTM04
   If IsEmptyText(m_TM04) = True Then: m_TM04 = "00"
   
   pub_QL05 = pub_QL05 & ";" & Label3 & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 'Add By Sindy 2010/10/15
   ' 查詢商標基本檔
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      OnProcess = False
      rsTmp.Close
      Set rsTmp = Nothing
      GoTo EXITSUB:
   Else
      'Add By Sindy 2023/2/20 將原註冊地址再存回案件之申請地址1(英)TM25
      If "" & rsTmp.Fields("TM25") <> txtOldAddr Then
         strSql = "UPDATE TRADEMARK SET tm25='" & ChgSQL(Trim(txtOldAddr)) & "' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      '2023/2/20 END
   End If
    m_TM10 = "" & rsTmp("TM10").Value
    m_TM11 = "" & rsTmp("TM11").Value
    m_TM21 = "" & rsTmp("TM21").Value    'add by sonia 2022/10/14
   rsTmp.Close
   
   pub_QL05 = pub_QL05 & ";" & Label6 & textCF15 'Add By Sindy 2010/10/15
   ' 查詢下一程序基本檔
   'Modified by Lydia 2019/12/16 請改為不控制NP09>=系統日，但若資料<系統日則提醒
   'strSql = "SELECT * FROM NEXTPROGRESS " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & textCF15 & " AND " & _
                     "(NP06 IS NULL OR NP06 = '' OR NP06 = ' ') " & _
                     "AND NP09 >= " & strSrvDate(1) & " Order BY NP09 ASC "
   'modify by sonia 2022/10/14 +CASEPROGRESS C1,CASEPROGRESS C2
   strSql = "SELECT NEXTPROGRESS.*,C2.CP10 CP10,C2.CP53 CP53 FROM NEXTPROGRESS,CASEPROGRESS C1,CASEPROGRESS C2 " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & textCF15 & " AND " & _
                     "(NP06 IS NULL OR NP06 = '' OR NP06 = ' ') AND NP01=C1.CP09(+) and C1.CP43=C2.CP09(+) " & _
                     " Order BY NP09 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/10/15
      OnProcess = False
      rsTmp.Close
      Set rsTmp = Nothing
      'ADD BY SONIA 2015/10/29 CFT-012359
      MsgBox "此案無此下一程序未過期的期限！"
      textTM02.SetFocus
      textTM02_GotFocus
      'end 2015/10/29
      GoTo EXITSUB:
   'Added by Lydia 2019/12/16
   Else   'Modified by Lydia 2019/12/16 請改為不控制NP09>=系統日，但若資料<系統日則提醒
      rsTmp.MoveFirst
      If "" & rsTmp.Fields("np09") < strSrvDate(1) Then
          If MsgBox("此期限已過期，法定期限：" & ChangeWStringToTDateString("" & rsTmp.Fields("np09")) & "，仍要通知嗎？", vbExclamation + vbYesNo + vbDefaultButton1, " 檢核資料") = vbYes Then
               MsgBox "過期期限不會經智權人員報價確認，定稿會直接產生，請務必自行修改定稿內容！", vbInformation, "檢核資料"
          Else
               textTM02.SetFocus
               textTM02_GotFocus
               GoTo EXITSUB:
          End If
      End If
   'end 2019/12/16
   End If
   
   InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/10/15
   rsTmp.MoveFirst
   'Modify By Sindy 2009/07/15
   'Do While rsTmp.EOF = False
   If rsTmp.RecordCount > 0 Then
   '2009/07/15 End
      m_NP22 = "" & rsTmp("np22") 'Add by Morgan 2008/5/19
      m_NP07 = Empty
      m_NP08 = Empty
      m_NP09 = Empty
      'add by nickc 2005/05/31
      m_NP01 = Empty
      If IsNull(rsTmp.Fields("NP01")) = False Then
         m_NP01 = rsTmp.Fields("NP01")
      End If
      ' 下一程序代碼
      If IsNull(rsTmp.Fields("NP07")) = False Then
         m_NP07 = rsTmp.Fields("NP07")
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("NP09")) = False Then
         m_NP09 = rsTmp.Fields("NP09")
      End If
      ' 本所期限
      '2005/10/18 MODIFY BY SONIA
      '延展定稿之本所期限改為法定期限-國家檔之延展時間(月)NA15,不抓下一程序檔
      '2006/1/9 MODIFY BY SONIA本所期限仍抓下一程序檔,另加法定期限-國家檔之延展時間(月)NA15去印定稿之其他日期m_DATE
      If IsNull(rsTmp.Fields("NP08")) = False Then
         m_NP08 = rsTmp.Fields("NP08")
      End If
      m_DATE = ""
      Select Case textCF15
         Case "102"
            m_DATE = DBDATE(DateAdd("m", -GetDelayTime(m_TM10), ChangeWStringToWDateString(m_NP09)))
            If m_TM10 = "011" Then m_DATE = CompDate(2, 1, m_DATE) 'Added by Morgan 2023/7/11 日本可辦延展期限起日要+1天(翌日起),回覆單要同步修改--May
            
         '2007/5/1 ADD BY SONIA 使用宣誓都抓一年
         Case "105"
            m_DATE = DBDATE(DateAdd("YYYY", -1, ChangeWStringToWDateString(m_NP09)))
            'add by sonia 2022/10/14 阿根廷使用宣誓抓專用期起日或上次延展專用期起日
            If m_TM10 = "118" Then
               m_DATE = m_TM21
               m_cpHave102 = ""
               If "" & rsTmp("CP10") = "102" And Val("" & rsTmp("CP53")) > 0 Then
                  m_DATE = "" & rsTmp("CP53")
                  m_cpHave102 = "經延展後"
               End If
            End If
            'end 2022/10/14
         '2007/5/1 END
      End Select
      '2006/1/9 END
      '2005/10/18 END
      'Add By Cheng 2003/03/26
      '新增列印接洽結案單資料
      pub_AddressListSN = pub_AddressListSN + 1
      PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & rsTmp("NP22").Value, "" & rsTmp("NP02").Value, "" & rsTmp("NP03").Value, "" & rsTmp("NP04").Value, "" & rsTmp("NP05").Value
      ' 列印定稿
      PrintLetter
      
      'add by nickc 2005/05/31
      'add by nickc 2005/09/27
      If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
         Call oClsPrtForm001.PrintReturnSheet(m_NP01, m_NP07, m_NP09, False)
      End If
      
      rsTmp.MoveNext
   'Modify By Sindy 2009/07/15
   'Loop
   End If
   '2009/07/15 End
   
   OnProcess = True
   
EXITSUB:
End Function

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textMoney_GotFocus()
   InverseTextBox textMoney
End Sub

'' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
'Private Sub InsExpField(ET01 As String, ET02 As String, ET03 As String)
'
'   ' 清除定稿例外欄位檔原有資料
'   EndLetter ET01, ET02, ET03, strUserNum
'   ' 法定期限
'   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','法定期限','" & DBDATE(m_NP09) & "')"
'   cnnConnection.Execute strSql
'   ' 本所期限
'   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','本所期限','" & DBDATE(m_NP08) & "')"
'   cnnConnection.Execute strSql
'   ' 其他日期   2006/1/9 ADD BY SONIA
'   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','其他日期','" & DBDATE(m_DATE) & "')"
'   cnnConnection.Execute strSql
'   '2006/1/9 END
'   ' 費用
'   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','費用','" & textMoney & "')"
'   cnnConnection.Execute strSql
'
'   'Add by Morgan 2008/5/19
'   If txtScore <> "" Then
'      ' 費用點數
'      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'            "','費用點數','" & txtScore & "')"
'      cnnConnection.Execute strSql
'   End If
'
'   '若有輸入變更費用
'   If Me.textChgMoney.Text <> "" Then
'        '附件
'        'Modify by Morgan 2008/5/19 將敘述移到定稿內
'        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'              "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'              "','變更費用','" & textChgMoney.Text & "')"
'        cnnConnection.Execute strSql
'   End If
'   'Add by Morgan 2008/5/19
'   If txtChgScore <> "" Then
'      ' 費用點數
'      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'            "','變更費用點數','" & txtChgScore & "')"
'      cnnConnection.Execute strSql
'   End If
'
'   Select Case m_NP07
'      Case "105" ' 使用宣誓
'         ' 申請國家
'         Select Case m_TM10
'            Case "030" ' 菲律賓
'               If ET03 = "05" Then
'                  'Add By Cheng 2004/02/10
'                  '應呈提第五, 十, 十五年使用宣誓書
'                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                        "','提使用宣誓書年度','" & Get105Year(m_NP09) & "')"
'                  cnnConnection.Execute strSql
'                  'End
'               End If
'            '2011/9/23 ADD BY SONIA
'            Case "046" ' 柬埔寨
'               If ET03 = "12" Then
'                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                        "','延展核准日','" & DBDATE(m_102cp25) & "')"
'                  cnnConnection.Execute strSql
'               End If
'            '2011/9/23 END
'         End Select
'   End Select
'
'End Sub

'Modify by Morgan 2008/5/19
'改先報價給智權人員,程式有大幅度調整
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strET01 As String
Dim strKey As String
Dim strET03 As String
Dim strSql As String
Dim rsTmp As ADODB.Recordset
'Add By Sindy 2012/1/16
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/16 End
   
   strET01 = "10"
   strKey = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&" & m_NP07
   
   Select Case m_NP07
        Case "102" '延展
         ' 申請國家
         Select Case m_TM10
            Case "011" '日本
               strET03 = "09"
            
            Case "012" '韓國
               strET03 = "07"
            
            Case "013" ' 香港
               strET03 = "02"
            
            Case "231" ' 德國
               strET03 = "03"
               
            'add by nickc 2006/05/16 ，因為葉芳如改內容
            Case "101" '美國
               strET03 = "11"
               
            'Add By Sindy 2014/10/9
            Case "030" '菲律賓
               strET03 = "13"
               
            'Add By Sindy 2017/4/18
            Case "048" '緬甸
               strET03 = "15"
            '2017/4/18 END
            
            'Add By Sindy 2020/6/3
            Case "104" '墨西哥
               strET03 = "16"
               
            'Add By Sindy 2020/10/15
            Case "239" '歐洲聯盟
               strET03 = "19"
               
            'Added by Lydia 2024/03/19
            Case "025" '伊朗"
               strET03 = "23"
               
            Case "032" '敘利亞
               strET03 = "24"
            'end 2024/03/19
            Case Else:
               strET03 = "01"
               
         End Select
      
      Case "105" ' 使用宣誓
         ' 申請國家
         Select Case m_TM10
            Case "046" ' 柬埔寨
               strET03 = "04"
                '2011/9/23 ADD BY SONIA 延展後使用宣誓用另一定稿,同時抓延展核准日
                m_102cp25 = ""
                strSql = "SELECT * FROM CASEPROGRESS " & _
                            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                  "CP02 = '" & m_TM02 & "' AND " & _
                                  "CP03 = '" & m_TM03 & "' AND " & _
                                  "CP04 = '" & m_TM04 & "' AND " & _
                                  "CP10 = '102' AND CP27 IS NOT NULL "
                Set rsTmp = New ADODB.Recordset
                rsTmp.CursorLocation = adUseClient
                rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If rsTmp.RecordCount > 0 Then
                   strET03 = "12"
                   m_102cp25 = "" & rsTmp("CP25").Value
                End If
                rsTmp.Close
                '2011/9/23 END
            
            Case "030" '菲律賓
                strET03 = "05"
                If m_TM11 <> "" And m_NP09 <> "" Then
                    'Modified by Morgan 2025/3/17
                    'If Format(DateAdd("yyyy", 3, ChangeWStringToWDateString(m_TM11)), "yyyy/MM/dd") = ChangeWStringToWDateString(m_NP09) Then
                    If Format(DateAdd("yyyy", 3, ChangeWStringToWDateString(m_TM11)), "yyyyMMdd") = m_NP09 Then
                    'end 2025/3/17
                        'Modified by Lydia 2025/04/14 菲律賓3年使用宣誓書
                        'strET03 = strET01
                        strET03 = 10
                    End If
                End If
            
            'add by sonia 2020/12/29
            Case "104" '墨西哥
                strET03 = "20"
            'end 2020/12/29
            
            Case "101" '美國第六年
                strET03 = "08"
                
            'Add By Sindy 2017/3/6
            Case "318" '莫三比克共和國
                strET03 = "14"
            '2017/3/6 END
            
            'add by sonia 2019/10/17
            Case "118"  '阿根廷第六年 2022/10改新定稿
                strET03 = "18"
            'end 2019/10/17
                
            'add by sonia 2022/10/19
            Case "110"  '海地
                strET03 = "21"
            
            Case "112"  '波多黎各
                strET03 = "22"
            'end 2022/10/19
                
         End Select
      
      'Add By Sindy 2019/10/18
      Case "109" ' 緩審延展
         strET03 = "17"
      '2019/10/18 END
      
      Case "702" ' 刊登廣告
         strET03 = "06"
         
      Case Else:
   End Select
   
   If strET03 <> "" Then
'      'Modify by Morgan 2008/10/24
'      '97.11.17 改為報價通知
'      If Val(strSrvDate(1)) < 20081117 Then
'         InsExpField strET01, strKey, strET03
'         'Add By Sindy 2012/1/16
'         If intPWhere = 國內 Then
'            bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_NP07 = "102", , bolPlusPaper)
'            If bolEmail Then
'               '判斷是否EMail同時寄紙本
'               If Not bolPlusPaper Then
'                  iCopy = 1
'               End If
'               NowPrint strKey, strET01, strET03, False, strUserNum, 0, , , , iCopy, , True, True
'               MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
'            Else
'               NowPrint strKey, strET01, strET03, False, strUserNum, 0
'            End If
'         Else
'         '2012/1/16 End
'            NowPrint strKey, strET01, strET03, False, strUserNum, 0
'         End If
'      Else
         PUB_AddLetterCache m_NP01, m_NP22, m_NP01, strET01, strET03
         InsExpField1 m_NP01, m_NP22, strET03
         strExc(0) = CompWorkDay(5, strSrvDate(1))
         strExc(1) = DBDATE(m_NP08)
         '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
         If Val(strExc(1)) <= Val(strExc(0)) Then
            PUB_Cache2Letter m_NP01, m_NP22, False, False
         End If
'      End If
   'Add By Sindy 2015/2/25 要報價但沒有定稿時提醒
   Else
      MsgBox "本案要報價但沒有系統的定稿，請注意！", vbExclamation
   '2015/2/25 END
   End If
End Sub

'2005/10/18 ADD BY SONIA
Private Function GetDelayTime(strTM10 As String) As Integer
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

StrSQLa = "Select NA15 From Nation Where NA01='" & strTM10 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetDelayTime = Val("0" & rsA.Fields(0).Value)
Else
   GetDelayTime = 0
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'2005/10/18 END

'Add by Morgan 2008/5/16
'寫例外欄位到暫存檔
Private Sub InsExpField1(NP01 As String, NP22 As String, Optional ET03 As String)
Dim tmpArr As Variant, j As Integer 'Add By Sindy 2020/8/25

   ' 法定期限
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & NP01 & "'," & NP22 & ",'法定期限','" & DBDATE(m_NP09) & "','')"
   cnnConnection.Execute strSql
   
   ' 本所期限
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & NP01 & "'," & NP22 & ",'本所期限','" & DBDATE(m_NP08) & "','')"
   cnnConnection.Execute strSql
   
   ' 其他日期
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & NP01 & "'," & NP22 & ",'其他日期','" & DBDATE(m_DATE) & "','')"
   cnnConnection.Execute strSql
   
   ' 費用
   'Added by Lydia 2016/03/04 催美國第6年使用宣誓
   If m_NP07 = "105" And m_TM10 = "101" Then
        strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                 "VALUES ('" & NP01 & "'," & NP22 & ",'第八條費用','" & textMoney & "','Y')"
        cnnConnection.Execute strSql
        If txtScore <> "" Then
           strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                 "VALUES ('" & NP01 & "'," & NP22 & ",'第八條費用點數','" & txtScore & "','')"
           cnnConnection.Execute strSql
        End If
        strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                 "VALUES ('" & NP01 & "'," & NP22 & ",'第八及十五條費用','" & textCaseFee & "','Y')"
        cnnConnection.Execute strSql
        If txtCScore <> "" Then
           strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                 "VALUES ('" & NP01 & "'," & NP22 & ",'第八及十五條費用點數','" & txtCScore & "','')"
           cnnConnection.Execute strSql
        End If
   Else
   'end 2016/03/04
        strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                 "VALUES ('" & NP01 & "'," & NP22 & ",'費用','" & textMoney & "','Y')"
        cnnConnection.Execute strSql
        If txtScore <> "" Then
           ' 費用點數
           strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                 "VALUES ('" & NP01 & "'," & NP22 & ",'費用點數','" & txtScore & "','')"
           cnnConnection.Execute strSql
        End If
   End If

   '若有輸入變更費用
   If Me.textChgMoney.Text <> "" Then
      '附件
      'Modify by Morgan 2008/5/19 將敘述移到定稿內
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'變更費用','" & textChgMoney & "','Y')"
      cnnConnection.Execute strSql
      
      'Add By Sindy 2020/6/3
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'原註冊地址','" & txtOldAddr & "','')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'新地址','" & txtNewAddr & "','')"
      cnnConnection.Execute strSql
      '2020/6/3 END
   End If
   
   If txtChgScore <> "" Then
      ' 費用點數
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'變更費用點數','" & txtChgScore & "','')"
      cnnConnection.Execute strSql
   End If
   
   'Added by Morgan 2022/4/25
   If txtUKMoney <> "" Then
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'英國費用','" & txtUKMoney & "','Y')"
      cnnConnection.Execute strSql
   End If
   If txtUKScore <> "" Then
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'英國費用點數','" & txtUKScore & "','')"
      cnnConnection.Execute strSql
   End If
   If txtUKChgMoney <> "" Then
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'英國變更費用','" & txtUKChgMoney & "','Y')"
      cnnConnection.Execute strSql
   End If
   If txtUKChgScore <> "" Then
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & NP01 & "'," & NP22 & ",'英國變更費用點數','" & txtUKChgScore & "','')"
      cnnConnection.Execute strSql
   End If
   'end 2022/4/25
   
   Select Case m_NP07
      'Add By Sindy 2010/02/25
      Case "102" '延展
         Select Case m_TM10
            'Add By Sindy 2021/3/17
            Case "239" '歐洲聯盟
               '1730.通知英國再註冊
               strSql = "select cp30 from caseprogress where cp01='" & m_TM01 & "'" & _
                  " and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "'" & _
                  " and cp10='1730'"
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'英國脫歐號數','" & adoRecordset.Fields("cp30") & "','')"
                  cnnConnection.Execute strSql
               End If
               
            '239.歐洲聯盟27個會員國
            Case "201", "231", "203", "204", "211", "207", "209", "208", "206", "212", "214", "220", "217", "213", "216", "223", "242", "230", "240", "241", "219", "236", "222", "232", "234", "226", "228"
               '若同一件商標在取得歐盟註冊前已有會員國註冊者，在
               '取得歐盟註冊後，會員國之註冊可不必辦理延展，故在催詢各會員國
               '延展之定稿上必須提醒客戶該商標已有歐盟註冊，以利
               '客戶斟酌是否有必要辦理會員國延展
               strSql = "select * from trademark where tm23='" & m_TM23 & "'" & _
                  " and tm05='" & m_TM05 & "' and tm16='1' and tm10='239' "
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  'Modify By Sindy 2020/8/25 催延展之商標案，若其申請人有相同商標名稱之歐盟案，且二案之商標類別至少有一個是重疊的，定稿才要帶出
                  'Modify By Sindy 2021/7/12 請加入卷宗性質是申請的條件
                  tmpArr = Split(textTM09, ",")
                  For j = 0 To UBound(tmpArr)
                     strSql = "select * from trademark where tm23='" & m_TM23 & "'" & _
                        " and tm05='" & m_TM05 & "' and tm16='1' and tm10='239' and instr(tm09,'" & tmpArr(j) & "')>0" & _
                        " and tm28='1'"
                     intI = 1
                     Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                  '2020/8/25 END
                        'modify by sonia 2017/11/23
                        'strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                              "VALUES ('" & NP01 & "'," & NP22 & ",'歐洲聯盟','本商標已取得歐洲聯盟註冊。','')"
      '                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      '                        "VALUES ('" & NP01 & "'," & NP22 & ",'歐洲聯盟','" & Chr(13) & "　　本商標已取得歐洲聯盟註冊。','')"
                        'Modify By Sindy 2020/8/25 同時帶出歐盟案之商品類別以供承辦人檢查
                        strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                              "VALUES ('" & NP01 & "'," & NP22 & ",'歐洲聯盟','" & Chr(13) & "　　本商標已取得歐洲聯盟註冊(第" & adoRecordset.Fields("tm09") & "類)。','')"
                        cnnConnection.Execute strSql
                        Exit For 'Add By Sindy 2020/8/25
                     End If
                  Next j
               End If
            'add by sonia 2022/11/7 CFT-010682
            Case "102"   '加拿大
               strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                     "VALUES ('" & NP01 & "'," & NP22 & ",'歐洲聯盟','" & Chr(13) & "　　依加拿大實務，審查員當依職權審查商標延展案之指定商品/服務是否符合現行國際分類，若商品或服務類別須重新分類時，得發給審查報告，要求申請人答辯及補繳延展之跨類規費，併此說明。','')"
               cnnConnection.Execute strSql
            'end 2022/11/7
            'add by sonia 2023/11/16 CFT-016572
            Case "019"   '泰國
               strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                     "VALUES ('" & NP01 & "'," & NP22 & ",'歐洲聯盟','" & Chr(13) & "　　依泰國實務，審查員會依職權審查商標延展案之指定商品/服務是否符合現行國際分類，若審查員發現商品或服務類別應重新分類時，得核發審查報告，要求商標權人依法修正指定商品/服務敘述或補繳延展的增類規費，屆時會衍生相關答辯費用，併此說明。','')"
               cnnConnection.Execute strSql
            'end 2023/11/16
         End Select
      '2010/02/25 End
      
      Case "105" ' 使用宣誓
         ' 申請國家
         Select Case m_TM10
            Case "030" ' 菲律賓
               If ET03 = "05" Then
                  '應呈提第五, 十, 十五年使用宣誓書
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'提使用宣誓書年度','" & PUB_Get030To105Year(Me.textTM01.Text, Me.textTM02.Text, Me.textTM03.Text, Me.textTM04.Text, m_NP09) & "','')"
                  cnnConnection.Execute strSql
               End If
            '2011/9/23 ADD BY SONIA
            Case "046" ' 柬埔寨
               If ET03 = "12" Then
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'延展核准日','" & DBDATE(m_102cp25) & "','')"
                  cnnConnection.Execute strSql
               End If
            '2011/9/23 END
            'add by sonia 2022/10/14
            Case "118" ' 阿根廷
               If ET03 = "18" Then
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'公開與否','" & m_cpHave102 & "','')"
                  cnnConnection.Execute strSql
               End If
            'end 2022/10/14
         End Select
   End Select
End Sub

Private Sub txtChgScore_GotFocus()
   TextInverse txtChgScore
End Sub

Private Sub txtChgScore_Validate(Cancel As Boolean)
   If txtChgScore <> "" Then
      If Not IsNumeric(txtChgScore) Then
         MsgBox "請輸入數字！"
         Cancel = True
         txtChgScore_GotFocus
      End If
   End If
End Sub

'地址
Private Sub txtOldAddr_GotFocus()
   OpenIme
   InverseTextBox txtOldAddr
End Sub
'Modified by Lydia 2022/12/06 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtOldAddr_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub txtOldAddr_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(txtOldAddr, txtOldAddr.MaxLength) = False Then
      Cancel = True
      txtOldAddr_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub
Private Sub txtNewAddr_GotFocus()
   OpenIme
   InverseTextBox txtNewAddr
End Sub
'Modified by Lydia 2022/12/06 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtNewAddr_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub txtNewAddr_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(txtNewAddr, txtNewAddr.MaxLength) = False Then
      Cancel = True
      txtNewAddr_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub txtScore_GotFocus()
   TextInverse txtScore
End Sub

Private Sub txtScore_Validate(Cancel As Boolean)
   If txtScore <> "" Then
      If Not IsNumeric(txtScore) Then
         MsgBox "請輸入數字！"
         Cancel = True
         txtScore_GotFocus
      End If
   End If
End Sub

Private Sub SetOldPrice(bolShow As Boolean)
   PUB_LabelActive lblFee1, lblFee1s, False
   If m_TM23 <> "" And m_TM10 <> "" And m_TM08 <> "" And textCF15 <> "" Then
      If PUB_GetOldPrice(m_TM23, m_TM10, m_TM08, textCF15, RsTemp, , , "2") = True Then
         PUB_LabelActive lblFee1, lblFee1s
         If bolShow = True Then
            Set frm880014.grdDataList.Recordset = RsTemp
            Set frm880014.fmParent = Me
            frm880014.Show vbModal
         End If
      End If
   End If
End Sub
'Added by Lydia 2016/03/04 CFT美國催第六年使用宣誓之設定
Private Sub SetCase105()
   '下一程序105使用宣誓且申請國家為101美國, 原"費用"欄改為"第八條費用",增加"第八及十五條費用"
   lblCase1.Visible = False: lblCase2.Visible = False
   textCaseFee.Visible = False: txtCScore.Visible = False
   lblFee1.Caption = "費用："
   If textCF15 = "105" And m_TM10 = "101" Then
        lblCase1.Visible = True: lblCase2.Visible = True
        textCaseFee.Visible = True: txtCScore.Visible = True
        lblFee1.Caption = "第八條費用："
   End If
End Sub

Private Sub txtUKChgMoney_GotFocus()
   InverseTextBox txtUKChgMoney
End Sub

Private Sub txtUKChgMoney_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(txtUKChgMoney) = False Then
      If IsNumeric(txtUKChgMoney) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtUKChgMoney_GotFocus
      End If
   End If
End Sub

Private Sub txtUKChgScore_GotFocus()
   InverseTextBox txtUKChgScore
End Sub

Private Sub txtUKChgScore_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(txtUKChgScore) = False Then
      If IsNumeric(txtUKChgScore) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtUKChgScore_GotFocus
      End If
   End If
End Sub

Private Sub txtUKMoney_GotFocus()
   InverseTextBox txtUKMoney
End Sub

Private Sub txtUKMoney_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(txtUKMoney) = False Then
      If IsNumeric(txtUKMoney) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtUKMoney_GotFocus
      End If
   End If
End Sub

Private Sub txtUKScore_GotFocus()
   InverseTextBox txtUKScore
End Sub

Private Sub txtUKScore_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(txtUKScore) = False Then
      If IsNumeric(txtUKScore) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtUKScore_GotFocus
      End If
   End If
End Sub
