VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030202_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5760
   ClientLeft      =   4710
   ClientTop       =   2280
   ClientWidth     =   9350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9350
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Text            =   "C:\temp\電子送件"
      Top             =   1230
      Width           =   7065
   End
   Begin VB.CommandButton CmdPath 
      Caption         =   "<="
      Height          =   315
      Left            =   8640
      TabIndex        =   15
      Top             =   1230
      Width           =   345
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   585
      Left            =   5100
      TabIndex        =   12
      Top             =   510
      Width           =   4035
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   13
         Top             =   210
         Width           =   3150
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   14
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdExtent 
      Caption         =   "延期(&D)"
      Height          =   400
      Left            =   5100
      TabIndex        =   10
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   4
      Top             =   780
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   3
      Top             =   780
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   780
      Width           =   732
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   810
      Value           =   -1  'True
      Width           =   1332
   End
   Begin VB.OptionButton radio 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   1332
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textCP09 
      Height          =   264
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   5
      Top             =   480
      Width           =   2892
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8280
      TabIndex        =   9
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7320
      TabIndex        =   8
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "發文資料(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6060
      TabIndex        =   11
      Top             =   60
      Width           =   1212
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   1
      Top             =   780
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4035
      Left            =   90
      TabIndex        =   18
      Top             =   1650
      Width           =   9165
      _ExtentX        =   16157
      _ExtentY        =   7108
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPath1 
      AutoSize        =   -1  'True
      Caption         =   "電子檔存放路徑："
      Height          =   180
      Left            =   150
      TabIndex        =   17
      Top             =   1290
      Width           =   1440
   End
End
Attribute VB_Name = "frm030202_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/18 grdList : MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
'Memo by Lydia 2021/09/01 改成Form2.0 ; grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 使用者所選取的查詢方式是收文號還是本所案號
Dim m_KeySel As Integer
' 使用者所選取的收文號
Public m_CP09 As String
' 使用者所選取的列其位置
Dim m_CurrSel As Integer
' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
'Add By Cheng 2002/04/22
Dim m_bln_NoData As Boolean
'Add By Cheng 2003/03/19
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Added by Lydia 2016/01/19
Dim m_TM10 As String '申請國家
Dim m_TM16 As String '已准駁
'Added by Lydia 2019/08/28
Dim m_TM22 As String '專用期間(止日)
'Add By Sindy 2024/8/14
Public bolIsEMPFlow As Boolean '是否為電子承辦簽核
Public m_EEP01 As String
Dim bolFirst As Boolean
'2024/8/14 END


Public Sub Clear()
   textCP09 = Empty
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   InitialGrdList
   radio(0).Value = True
   radio(1).Value = False
   radio_Click 0
End Sub

Public Sub Clear1()
   textCP09 = Empty
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   InitialGrdList
   cmdQuery.Default = True
   textTM02.SetFocus   'add by sonia 2016/9/9
End Sub

Private Sub cmdExtent_Click()
Dim frmNext As Form
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   If CheckDataValid = True Then
      
      'Add By Cheng 2002/07/15
      '所點選的案件性質不可為"延期"
      If PUB_CPKindDelay(Me.grdList.TextMatrix(Me.grdList.row, 7), "T") Then
         Exit Sub
      End If
      
      'Add By Cheng 2002/07/12
      '若案件已閉卷, 不可發文
      If PUB_CaseClosedCP09(Me.grdList.TextMatrix(Me.grdList.row, 7)) = True Then
         Exit Sub
      End If
        
      '2006/3/20 ADD BY SONIA 若專用期間已過期但發文案件性質非延展,補正時, 不可發文
      If Me.grdList.TextMatrix(Me.grdList.row, 8) <> "102" And Me.grdList.TextMatrix(Me.grdList.row, 8) <> "201" Then
         StrSQLa = "SELECT TM22 FROM TRADEMARK,CASEPROGRESS WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 7) & "'"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'edit by nickc 2006/06/26 半年內皆可以補繳
            'If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ServerDate Then
                'MsgBox "此案件專用期間已過, 不可執行發文作業!!!", vbExclamation + vbOKOnly
            If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(ServerDate))) Then
               MsgBox "此案件專用期間已過半年, 不可執行發文作業!!!", vbExclamation + vbOKOnly
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               Exit Sub
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
      '2006/3/20 END
      
      'Add By Sindy 2024/8/14
      '檢查是否有承辦歷程是否有產生承辦單可以發文
      If PUB_IsEmpFlowIsSend(m_CP09) = False Then
         Exit Sub
      End If
      '2024/8/14 END
      
      Set frmNext = frm030202_11
      ' 顯示下一個畫面
      If IsObject(frmNext) = True Then
         frmNext.SetData 0, m_CP09, True
         Me.Hide
         frmNext.Show
         frmNext.QueryData
      End If
   End If
End Sub

Private Sub cmdok_Click()
'Add By Cheng 2002/07/11
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
   
   ' 檢查是否資料以完全輸入
   If CheckDataValid = True Then
      'Added by Lydia 2016/01/19 台灣案的註冊費發文時,判斷案件為已准
      If (Me.grdList.TextMatrix(Me.grdList.row, 10) = "FCT" Or Me.grdList.TextMatrix(Me.grdList.row, 10) = "T") And m_TM10 = 台灣國家代號 And m_TM16 <> "1" And Me.grdList.TextMatrix(Me.grdList.row, 8) = "717" Then
         MsgBox "本案尚未核准不可繳註冊費!", vbCritical
         Exit Sub
      End If
      'end 2016/01/19
      
      'Add By Cheng 2002/07/12
      '若案件已閉卷, 不可發文
      'Modify By Sindy 2021/3/23 FCT案之回代720及催款901不管是否閉卷都開放發文
      'Modify By Sindy 2021/3/31 + 外商發文722,不管是否閉卷都開放發文
      If Me.grdList.TextMatrix(Me.grdList.row, 8) <> "720" And _
         Me.grdList.TextMatrix(Me.grdList.row, 8) <> "901" And _
         Me.grdList.TextMatrix(Me.grdList.row, 8) <> "722" Then
      '2021/3/23 END
         If PUB_CaseClosedCP09(Me.grdList.TextMatrix(Me.grdList.row, 7)) = True Then
            Exit Sub
         End If
      End If
      
      'Added by Lydia 2015/11/24 管控延展案102,系統日不得早於"延展期滿前6個月"的第一天
      If Me.grdList.TextMatrix(Me.grdList.row, 8) = "102" And Not IsNull(Me.grdList.TextMatrix(Me.grdList.row, 9)) Then
         'Modified by Lydia 2017/06/01 延展期滿日期改用模組控制 ;因為下午可發次日,所以多判斷前一工作天
         'If strSrvDate(1) < CompWorkDay(2, CompDate(1, -6, Me.grdList.TextMatrix(Me.grdList.row, 9)), 1) Then
         If strSrvDate(1) < CompWorkDay(2, PUB_Get102DeadLine("3", Me.grdList.TextMatrix(Me.grdList.row, 9)), 1) Then
            MsgBox "延展案發文時,系統日不得早於延展期滿前6個月的第一天!", vbCritical
            Exit Sub
         End If
      End If
      'end 2015/11/24
      
      'Add By Sindy 2024/8/14
      '檢查是否有承辦歷程是否有產生承辦單可以發文
      If PUB_IsEmpFlowIsSend(m_CP09) = False Then
         Exit Sub
      End If
      '2024/8/14 END
        
      '2006/3/20 ADD BY SONIA 若專用期間已過期但發文案件性質非延展,補正時, 不可發文
      'modify by sonia 2019/5/28 +剔除延期303(FCT-043529的補正延期)
      'Modified by Lydia 2019/08/28 修改不限案件性質, 而且改為僅提醒仍可選擇要不要繼續發文
'      If Me.grdList.TextMatrix(Me.grdList.row, 8) <> "102" And Me.grdList.TextMatrix(Me.grdList.row, 8) <> "201" And Me.grdList.TextMatrix(Me.grdList.row, 8) <> "303" Then
'         StrSQLa = "SELECT TM22 FROM TRADEMARK,CASEPROGRESS WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 7) & "'"
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            'edit by nickc 2006/06/26 半年內皆可以補繳
'            'If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ServerDate Then
'                'MsgBox "此案件專用期間已過, 不可執行發文作業!!!", vbExclamation + vbOKOnly
'            If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(ServerDate))) Then
'               MsgBox "此案件專用期間已過半年, 不可執行發文作業!!!", vbExclamation + vbOKOnly
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
'               Exit Sub
'            End If
'         End If
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'      'add by sonia 2019/5/28 延期303可能為延展補正延期(FCT-043529),若已過專用期則提醒不必限制
'      ElseIf Me.grdList.TextMatrix(Me.grdList.row, 8) = "303" Then
'         StrSQLa = "SELECT TM22 FROM TRADEMARK,CASEPROGRESS WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 7) & "'"
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ServerDate Then
'               MsgBox "此案件專用期間已過, 請確認是否仍要發文延期 !!!"
'            End If
'         End If
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'      'end 2019/5/28
'      End If
'      '2006/3/20 END
      If m_TM22 <> "" And m_TM22 < strSrvDate(1) Then '修改不限案件性質, 而且改為僅提醒仍可選擇要不要繼續發文
         If MsgBox("此案件專用期間已過, 請確認是否繼續發文？", vbExclamation + vbYesNo + vbDefaultButton2, "專用期間已過") = vbNo Then
             Exit Sub
         End If
      End If
      'end 2019/08/28
      
      'Add By Cheng 2002/07/11
      '檢查所點選的案件進度資料當案件性質為自請撤回"306"或自請撤銷"307"時, 其相關總收文號若為空白, 則不可進入下一畫面
      StrSQLa = "Select * From CaseProgress Where CP09='" & m_CP09 & "' AND (CP10='306' OR CP10='307') AND CP43 IS NULL "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         MsgBox "該自撤案件未輸入相關總收文號, 請先補齊資料!!!", vbExclamation + vbOKOnly
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         Exit Sub
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      
      'Added by Lydia 2020/07/22 檢查路徑
      If txtPath1.Visible = True And txtPath1.Text <> "" Then
          If Dir(txtPath1.Text, vbDirectory) = "" Then
              MsgBox "電子檔存放路徑不存在！", vbOKOnly, "檢核資料"
              txtPath1.SetFocus
              txtPath1_GotFocus
              Exit Sub
          End If
      End If
      'end 2020/07/22
      
      'add by sonia 2021/4/23 若下一程序有相同案件性質未續辦則提醒
      StrSQLa = "Select * From CaseProgress,nextprogress Where CP09='" & m_CP09 & "' AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND CP10=NP07(+) AND NP06 IS NULL AND NP01 IS NOT NULL"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         MsgBox "此案件性質於下一程序檔仍有未續辦期限，請自行確認是否消期限!!!", vbExclamation + vbOKOnly
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      'end 2021/4/23
      
      '91.7.16此段錯誤, 正確的控制在DisplayNextForm的ShowMaintainForm m_CP09
      'Modify By Cheng 2002/07/11
      '檢查是否要顯示商檔基本檔資料維護的畫面暫時取消
'      ' 檢查是否要顯示商標基本檔資料維護的畫面
'      If CheckJumpFrm020501() = True Then
'         DisplayFrm020501
'      Else
         DisplayNextForm
'      End If
   End If
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
Initial
InitialGrdList
UpdateCtrlState

'設定印表機
SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, m_OriPrinterName, False, SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

'Added by Lydia 2020/07/22 預設自動上傳到卷宗區的指定資料夾
If GetSetting("TAIE", "FCT", UCase(Me.Name) & "Dir", "") <> "" Then
    txtPath1.Text = GetSetting("TAIE", "FCT", UCase(Me.Name) & "Dir", "")
Else
    txtPath1.Text = PUB_Getdesktop '預設個人桌面
End If
'end 2020/07/22

End Sub

Private Sub Initial()
   ' 預設由收文號來取得資料
   'modify by sonia 2016/9/9 改本所案號,同時改form之radio(1).Value為true及欄位tabindex順序
   'm_KeySel = 0
   m_KeySel = 1
End Sub

' 按下結束離開按紐
Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2003/01/29
    '列印地址條
'move to unload by nick 2004/10/22
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    'Add By Cheng 2003/02/05
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
    Unload Me
End Sub
' 按下查詢按紐
Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ' 先檢查該輸入的資料是否有全部輸入
   Select Case m_KeySel
      ' 依收文號
      Case 0:
         If IsEmptyText(textCP09) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      ' 依本所案號
      Case 1:
         If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   ' 查詢資料
   'Modify By Cheng 2002/04/22
   QueryData
'   If QueryData() = False Then
   If m_bln_NoData = True Then
      strTit = "資料查詢"
      strMsg = "沒有符合條件的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Modify By Cheng 2002/09/17
'      textTM01.SetFocus
      If Me.radio(0).Value Then
         Me.textCP09.SetFocus
         textCP09_GotFocus
      Else
         textTM01.SetFocus
         textTM01_GotFocus
      End If
      cmdQuery.Default = True
   Else
      cmdOK.Default = True
   End If
EXITSUB:
End Sub

' 查詢資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim rsTmp As New ADODB.Recordset
      
   'Add By Cheng 2002/04/22
   m_bln_NoData = True
   
   m_CP09 = Empty
   InitialGrdList
   
   'Add By Sindy 2024/8/14 控管多筆未發文時,後面的資料不能直接進入發文作業
   If bolFirst = True Then
      bolIsEMPFlow = False
   End If
   bolFirst = True
   '2024/8/14 END
   
   ' 組成SQL語法
   Select Case m_KeySel
      ' 依收文號
      Case 0:
         ' 檢查案件進度檔, 系統類別必須為FCT, 且必須為未輸入發文日, 且未輸入取消收文日期的A,B類收文號
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP09 = '" & textCP09 & "' "
                        
      ' 依本所案號
      Case 1:
         strCP01 = Trim(textTM01)
         strCP02 = Trim(textTM02)
         strCP03 = Trim(textTM03)
         If IsEmptyText(strCP03) = True Then: strCP03 = "0"
         strCP04 = Trim(textTM04)
         If IsEmptyText(strCP04) = True Then: strCP04 = "00"
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & strCP01 & "' AND " & _
                        "CP02 = '" & strCP02 & "' AND " & _
                        "CP03 = '" & strCP03 & "' AND " & _
                        "CP04 = '" & strCP04 & "' "
   End Select
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   ' 列出所有資料
   If rsTmp.RecordCount > 0 Then
'      'Add By Cheng 2002/04/22
'      m_bln_NoData = False
      
      ListData rsTmp
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 列出所有符合條件的資料
Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
Dim nRow As Integer
Dim m_TM29 As String 'Add By Sindy 2022/1/11 是否閉卷
Dim i As Integer, intQRow As Integer
   
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   
   'Added by Lydia 2016/01/19
   m_TM10 = "": m_TM16 = ""
   m_TM22 = ""  'Added by Lydia 2019/08/28
   
   strExc(1) = rsTmp.Fields("CP01")
   strExc(2) = rsTmp.Fields("CP02")
   strExc(3) = rsTmp.Fields("CP03")
   strExc(4) = rsTmp.Fields("CP04")
   If ClsPDGetSystemKind(rsTmp.Fields("CP01"), intI) Then
      'Modified by Lydia 2016/01/19 + 目前准駁TM16
      If intI = 2 Then '商標
         'Modified by Lydia 2019/08/28 +TM22
         strExc(0) = "SELECT TM10,TM16,TM22,TM29 FROM TRADEMARK WHERE TM01='" & strExc(1) & "' AND TM02='" & strExc(2) & "' AND TM03='" & strExc(3) & "' AND TM04='" & strExc(4) & "'"
      'Add By Sindy 2015/10/19
      ElseIf intI = 3 Then '法務
         'Modified by Lydia 2019/08/28 +TM22
         strExc(0) = "SELECT LC15,'','' as TM22,LC08 as TM29 FROM LAWCASE WHERE LC01='" & strExc(1) & "' AND LC02='" & strExc(2) & "' AND LC03='" & strExc(3) & "' AND LC04='" & strExc(4) & "'"
      '2015/10/19 END
      Else
         'Modified by Lydia 2019/08/28 +TM22
         strExc(0) = "SELECT SP09,'','' as TM22,SP15 as TM29 FROM SERVICEPRACTICE WHERE SP01='" & strExc(1) & "' AND SP02='" & strExc(2) & "' AND SP03='" & strExc(3) & "' AND SP04='" & strExc(4) & "'"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) < "010" Then
            grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 0)
         Else
            grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 1)
         End If
         'Added by Lydia 2016/01/19
         m_TM10 = "" & RsTemp.Fields(0)
         m_TM16 = "" & RsTemp.Fields(1)
         'Added by Lydia 2019/08/28
         m_TM22 = "" & RsTemp.Fields(2)
         m_TM29 = "" & RsTemp.Fields("TM29") 'Add By Sindy 2022/1/11 是否閉卷
      End If
   End If
   
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      ' 系統類別必須為"FCT"
      If IsNull(rsTmp.Fields("CP01")) = False Then
         Select Case rsTmp.Fields("CP01")
            '2009/10/20 modify by sonia 加入FMT的外商發文722
            'modify by sonia 2015/9/4 再加入非台灣案之719告代720回代, TS-001292
            'Modify By Sindy 2015/10/19 +FCL,LIN的901催款
            Case "FCT", "S", "T", "TS", "FCL", "LIN":
            Case Else: GoTo EXITSUB
         End Select
      End If
      
      'Add By Sindy 2015/10/19
      If rsTmp.Fields("CP01") = "FCL" Or rsTmp.Fields("CP01") = "LIN" Then
         If rsTmp.Fields("CP10") <> "901" Then
            GoTo NextRecord
         End If
      End If
      '2015/10/19 END
      
      '2009/10/20 ADD BY SONIA 加入FMT的外商發文722
      '2010/5/24 MODIFY BY SONIA 加入T的回覆代理人720
      'MODIFY BY SONIA 2015/9/4 加入所有非台灣案的719告知代理人,720回覆代理人,722外商發文
      'If rsTmp.Fields("CP01") = "T" And rsTmp.Fields("CP10") <> "722" And rsTmp.Fields("CP10") <> "720" And rsTmp.Fields("CP10") <> "719" Then
      If RsTemp.Fields(0) <> "000" Then
         If rsTmp.Fields("CP10") <> "722" And rsTmp.Fields("CP10") <> "720" And rsTmp.Fields("CP10") <> "719" Then
            GoTo NextRecord
         End If
      End If
      '2009/10/20 END
      
      '收文號不為A,B類的不予計入
      '2009/6/9 CANCEL BY SONIA 開放C類來函也可以由此發文
      'Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
      '   Case "A", "B":
      '   Case Else: GoTo NextRecord
      'End Select
      '2009/6/9 end
      ' 尚未輸入發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then
         If IsEmptyText(rsTmp.Fields("CP27")) = False Then
            If rsTmp.Fields("CP27") <> "0" Then: GoTo NextRecord
         End If
      End If
      
      'Modify By Sindy 2022/1/12
      '案件性質：720、901、722的發文，包含FCT及T(外商收文的)的發文程式
      '1. 未閉卷案件，只要取消收文都不出現；
      '2. 已閉卷案件，不考慮是否取消收文，只要沒有發文日的，都要出現。
      If (rsTmp.Fields("CP01") = "FCT" Or rsTmp.Fields("CP01") = "T") And _
         (rsTmp.Fields("CP10") = "720" Or rsTmp.Fields("CP10") = "901" Or rsTmp.Fields("CP10") = "722") Then
         If m_TM29 <> "Y" Then
            ' 尚未輸入取消收文日期
            If IsNull(rsTmp.Fields("CP57")) = False Then
               If IsEmptyText(rsTmp.Fields("CP57")) = False Then
                  If rsTmp.Fields("CP57") <> "0" Then: GoTo NextRecord
               End If
            End If
         End If
      Else
         ' 尚未輸入取消收文日期
         If IsNull(rsTmp.Fields("CP57")) = False Then
            If IsEmptyText(rsTmp.Fields("CP57")) = False Then
               If rsTmp.Fields("CP57") <> "0" Then: GoTo NextRecord
            End If
         End If
      End If
      '2022/1/12 END
      
      'Add By Cheng 2002/09/17
      m_bln_NoData = False
               
      grdList.Rows = grdList.Rows + 1
      nRow = grdList.Rows - 1
      ' 收文日欄位
      If IsNull(rsTmp.Fields("CP05")) = False Then
         grdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP05")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         strExc(1) = rsTmp.Fields("CP01")
         strExc(2) = rsTmp.Fields("CP02")
         strExc(3) = rsTmp.Fields("CP03")
         strExc(4) = rsTmp.Fields("CP04")
         If RsTemp.Fields(0) < "010" Then
            grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 0)
         Else
            grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 承辦人
      If IsNull(rsTmp.Fields("CP14")) = False Then
         grdList.TextMatrix(nRow, 3) = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         grdList.TextMatrix(nRow, 4) = GetStaffName(rsTmp.Fields("CP13"))
      End If
      'Add By Sindy 2012/5/3
      If IsNull(rsTmp.Fields("CP67")) = False Then
         grdList.TextMatrix(nRow, 5) = Format(rsTmp.Fields("CP67"), "##:##")
      End If
      '2012/5/3 End
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         grdList.TextMatrix(nRow, 6) = rsTmp.Fields("CP64")
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         grdList.TextMatrix(nRow, 7) = rsTmp.Fields("CP09")
      End If
      '2006/3/20 ADD BY SONIA
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         grdList.TextMatrix(nRow, 8) = rsTmp.Fields("CP10")
      End If
      '2006/3/20 END
      'Add By Sindy 2010/12/27 判斷有相關總收文號才做
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         '案件性質
         grdList.TextMatrix(nRow, 2) = grdList.TextMatrix(nRow, 2) & PUB_GetRelateCasePropertyName(grdList.TextMatrix(nRow, 7), "1")
      End If
      '2010/12/27 End
      'Modified by Lydia 2015/11/24 +CP07
      '法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         grdList.TextMatrix(nRow, 9) = rsTmp.Fields("CP07")
      End If
      'end 2015/11/24
      'Added by Lydia 2016/01/19 + CP01
      If IsNull(rsTmp.Fields("CP01")) = False Then
         grdList.TextMatrix(nRow, 10) = rsTmp.Fields("CP01")
      End If
      'end 2016/01/19
NextRecord:
      rsTmp.MoveNext
   Loop
   
   'Added by Lydia 2022/03/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
   If grdList.Rows >= 2 Then
        grdList.FixedRows = 1
   End If
   'end 2022/03/18

   ' 顯示符合的所有資料
   grdList.Refresh
   
   'Modify By Sindy 2024/8/14
   ' 設定第一筆為被選取的狀態
   'grdList_SetSelection 1
   If grdList.Rows >= 2 Or (bolIsEMPFlow = True And m_EEP01 <> "") Then
      If (bolIsEMPFlow = True And m_EEP01 <> "") Then
         For i = 1 To grdList.Rows - 1
            If grdList.TextMatrix(i, 7) = m_EEP01 Then
               intQRow = i
               Exit For
            End If
         Next i
      Else
        intQRow = 1 '若有資料游標停在第一筆
      End If
   End If
   If intQRow > 0 Then
      grdList_SetSelection intQRow
      If bolIsEMPFlow = True Then Call cmdok_Click
   End If
   '2024/8/14 END
   
EXITSUB:
End Sub

' 更新控制項的狀態
Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textCP09, True
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         textTM02_2.Visible = False
      Case 1:
         EnableTextBox textCP09, False
         EnableTextBox textTM01, True
         EnableTextBox textTM02, True
         EnableTextBox textTM03, True
         EnableTextBox textTM04, True
         textTM01_Validate False
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PUB_PrintAddressList strUserNum, Me.Combo1.Text
    '刪除地址條列表資料
    PUB_DeleteAddressList strUserNum
    '初始化序號
    pub_AddressListSN = 0
    'Add By Cheng 2003/02/05
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    'Add By Cheng 2003/01/28
'還原預設印表機
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
'Add By Cheng 2002/07/19
Set frm030202_01 = Nothing
End Sub

' 使用者按下所選取的項目
Public Sub radio_Click(Index As Integer)
   '************ 90.11.23  nick
   If frm030202_01.Visible = True Then
   m_KeySel = Index
   UpdateCtrlState
   ' 90.07.25 modify
   Select Case Index
      Case 0:
         textCP09.SetFocus
      Case 1:
         textTM01.SetFocus
   End Select
   End If
   '**********************************
End Sub

Private Sub textCP09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   '2005/8/29 ADD BY SONIA
   cmdQuery.Default = True
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 檢查系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         '2009/10/20 modify by sonia 加入FMT的外商發文722
         'modify by sonia 2015/9/4 再加入非台灣案之719告代720回代, TS-001292
         'Modify By Sindy 2015/10/19 +FCL,LIN的901催款
         Case "FCT", "S", "T", "TS", "FCL", "LIN":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   Dim nIndex As Integer
   grdList.Clear
   grdList.Rows = 1
   'Modified by Lydia 2015/11/24 +CP07
   'grdList.Cols = 9 '8
   'Modified by Lydia 2016/01/19 +CP01
   'grdList.Cols = 10
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文日"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "案件性質"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "承辦人"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "智權人員"
   grdList.ColWidth(4) = 1000
   'Add By Sindy 2012/5/3
   grdList.col = 5
   grdList.Text = "收文時間"
   grdList.ColWidth(5) = 1000
   '2012/5/3 End
   grdList.col = 6 '5
   grdList.Text = "進度備註"
   grdList.ColWidth(6) = 1200
   ' 收文號欄位 (隱藏欄位)
   grdList.col = 7 '6
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   '2006/3/20 ADD BY SONIA
   grdList.col = 8 '7
   grdList.Text = "CP10"
   grdList.ColWidth(8) = 0
   'Modified by Lydia 2015/11/24 +CP07
   grdList.col = 9
   grdList.Text = "CP07"
   grdList.ColWidth(9) = 0
   'Added by Lydia 2016/01/19 +CP01
   grdList.col = 10
   grdList.Text = "CP01"
   grdList.ColWidth(10) = 0
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
   End If
End Sub

Private Sub grdList_SelChange()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         m_CP09 = grdList.TextMatrix(grdList.row, 7)
      End If
   End If
   grdList_ShowSelection
   cmdOK.SetFocus
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If

   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

' 顯示下一個畫面
Public Sub DisplayNextForm()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCP10 As String
   Dim strCP01 As String
   Dim frmNext As Form
   
   strCP10 = Empty
   ' 組成SQL語法
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 列出所有資料
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP01")) = False Then
         strCP01 = rsTmp.Fields("CP01")
      End If
      If IsNull(rsTmp.Fields("CP10")) = False Then
         strCP10 = rsTmp.Fields("CP10")
      End If
      'Add By Sindy 2024/12/11
      If IsNull(rsTmp.Fields("CP14")) = True Then
         MsgBox "未分案不可發文!!!", vbExclamation + vbOKOnly
         Exit Sub
      End If
      '2024/12/11 END
   End If
   rsTmp.Close
   
   Select Case strCP01
      'Add By Sindy 2015/10/21
      Case "FCL", "LIN"
         frm071006.SetParent Me
         frm071006.Show
         Me.Hide
         Exit Sub
      Case Else
      '2015/10/21 END
         ' 其它
         Set frmNext = frm030202_16
         Select Case strCP10
            ' 申請, 延展, 補換發證書, 申請英文證明
            Case "101", "102", "103", "304": Set frmNext = frm030202_03
            ' 變更, 更正 2007/6/7 加減縮商品
            Case "301", "302", "313": Set frmNext = frm030202_07
            ' 移轉
            Case "501": Set frmNext = frm030202_08
            ' 授權, 再授權, 終止授權, 終止再授權  2009/10/14加徵求同意書724
            Case "502", "503", "504", "505", "724": Set frmNext = frm030202_09
            'Modify By Cheng 2002/06/05
      '      ' 補正
      '      Case "201": Set frmNext = frm030202_10
            ' 補正, 放棄專用權 2009/12/3 加電話回覆209
            Case "201", "206", "209": Set frmNext = frm030202_10
            ' 延期
            Case "303": Set frmNext = frm030202_11
            ' 自請撤回, 自請撤銷
            Case "306", "307": Set frmNext = frm030202_12
            ' 設定質權, 撤銷設定質權
            Case "506", "507": Set frmNext = frm030202_13
            ' 異議, 評定, 廢止, 評定專用權, 參加評定, 自評專用權, 禁止處分
            Case "601", "603", "605", "607", "608", "609", "616": Set frmNext = frm030202_14
            ' 補充理由, 訴願, 再訴願, 行政訴訟, 參加行政訴訟, 再審之訴
            Case "202", "612", "401", "402", "403", "404", "405": Set frmNext = frm030202_15
            ' 異議答辯, 評定答辯, 廢止答辯, 補充答辯, 參加被評定, 撤銷禁止處分, 修正, 刊登廣告, 第一期註冊費, 第二期註冊費, 其它
      '      Case "602", "604", "606", "613", "610", "617", "203", "702": Set frmNext = frm030202_16
            Case "602", "604", "606", "613", "610", "617", "203", "702", "715", "716": Set frmNext = frm030202_16
            ' 補理由書
            Case "611": Set frmNext = frm030202_17
         End Select
   End Select
   
   ' 顯示下一個畫面
   If IsObject(frmNext) = True Then
      frmNext.SetData 0, m_CP09, True
      '*********** 901121     nick
      If Me.Visible = True Then
         cmdQuery.Default = True
      End If
      '****************************
      Me.Hide
      frmNext.Show
      frmNext.QueryData
      
      'Modify By Cheng 2002/07/11
      '檢查是否要顯示商檔基本檔資料維護的畫面暫時取消
      ' 顯示商標基本資料的畫面
      'Modify By Sindy 2012/8/7 FCT案碰上未輸申請案號之案件，在分案及發文時會彈出基本資料供補輸，煩請控制案件性質為”回代理人函”及”告代理人函”時，不要彈，謝謝,蓮
      If strCP10 <> "720" And strCP10 <> "719" Then
      '2012/8/7 End
         ShowMaintainForm m_CP09
      End If
   End If
End Sub

' 顯示商標基本檔檔案畫面要求輸入
Private Sub DisplayFrm020501()
   frm020501.SetSystem 0
   frm020501.Show
End Sub

' 若為新案件且非新申請案且卷宗性質為"申請"時, 若商標基本檔的申請案號欄位是空白, 則先切換至商標基本資料維護
Private Function CheckJumpFrm020501() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bShowFrm020501 As Boolean
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim strTM12 As String
   Dim strTM28 As String
   Dim strCP10 As String
   Dim strCP31 As String
   
   bShowFrm020501 = False
   
   strTM12 = Empty
   strCP10 = Empty
   strCP31 = Empty
   
   ' 查詢案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   
   If IsNull(rsTmp.Fields("CP01")) = False Then: strTM01 = rsTmp.Fields("CP01")
   If IsNull(rsTmp.Fields("CP02")) = False Then: strTM02 = rsTmp.Fields("CP02")
   If IsNull(rsTmp.Fields("CP03")) = False Then: strTM03 = rsTmp.Fields("CP03")
   If IsNull(rsTmp.Fields("CP04")) = False Then: strTM04 = rsTmp.Fields("CP04")
   ' 案件性質
   If IsNull(rsTmp.Fields("CP10")) = False Then
      If IsEmptyText(rsTmp.Fields("CP10")) = False Then
         strCP10 = rsTmp.Fields("CP10")
      End If
   End If
   ' 是否為新案件欄位
   If IsNull(rsTmp.Fields("CP31")) = False Then
      If IsEmptyText(rsTmp.Fields("CP31")) = False Then
         strCP31 = rsTmp.Fields("CP31")
      End If
   End If
   rsTmp.Close
   
   ' 查詢商標基本檔
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   ' 卷宗性質
   If IsNull(rsTmp.Fields("TM28")) = False Then
      If IsEmptyText(rsTmp.Fields("TM28")) = False Then
         strTM28 = rsTmp.Fields("TM28")
      End If
   End If
   rsTmp.Close
   
   ' 判斷是否要顯示商標基本檔檔案維護的畫面
   If strTM28 = "1" Then
      If UCase(strCP31) = "Y" Then
         If strCP10 <> "101" Then
            bShowFrm020501 = True
         End If
      End If
   End If
   
   'CheckJumpFrm020501 = bShowFrm020501
   CheckJumpFrm020501 = False
EXITSUB:
   Set rsTmp = Nothing
End Function
' 檢查是否已選取資料
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If grdList.Rows <= 1 Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(m_CP09) = True Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_TM01 = strData
      Case 1: m_TM02 = strData
      Case 2: m_TM03 = strData
      Case 3: m_TM04 = strData
   End Select
End Sub
' 更新查詢的方式由本所案號來查詢
Public Sub SetQueryFromTM()
   textTM01 = m_TM01
   textTM02 = m_TM02
   textTM03 = m_TM03
   textTM04 = m_TM04
   radio_Click 1
End Sub

Private Sub textCP09_GotFocus()
   InverseTextBox textCP09
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub
'2005/8/29 ADD BY SONIA
Private Sub textTM02_2_KeyPress(KeyAscii As Integer)
   cmdQuery.Default = True
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub
'2005/8/29 ADD BY SONIA
Private Sub textTM02_KeyPress(KeyAscii As Integer)
   cmdQuery.Default = True
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

'Added by Lydia 2020/07/22
Private Sub cmdPath_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.pdf"
      .Filter = "PDF檔案 (*.pdf)|*.pdf"
      If GetSetting("TAIE", "FCT", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCT", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCT", UCase(Me.Name) & "Dir", sFile(0)
            txtPath1.Text = sFile(0)
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
                SaveSetting "TAIE", "FCT", UCase(Me.Name) & "Dir", Left(.FileName, InStrRev(.FileName, "\") - 1)
            End If
            txtPath1.Text = Left(.FileName, InStrRev(.FileName, "\") - 1)
         End If
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Added by Lydia 2020/07/22
Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub

