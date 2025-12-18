VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210133 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件結案單"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00C0C0FF&
      Caption         =   "回覆單匯入"
      Height          =   345
      Left            =   6870
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   30
      Width           =   1065
   End
   Begin VB.CommandButton cmdFlowEmp 
      Caption         =   "簽核人員"
      Height          =   285
      Left            =   3420
      TabIndex        =   4
      Top             =   420
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   3
      Left            =   7980
      TabIndex        =   15
      Top             =   30
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      Height          =   345
      Left            =   5010
      TabIndex        =   29
      Top             =   30
      Width           =   1815
      Begin VB.CommandButton cmdSend 
         Caption         =   "送出(&E)"
         Height          =   345
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5010
      TabIndex        =   25
      Top             =   30
      Width           =   1485
      Begin VB.TextBox txtPCnt 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   0
         Left            =   750
         MaxLength       =   1
         TabIndex        =   26
         Text            =   "2"
         Top             =   30
         Width           =   270
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "列印(&P)　　份"
         Height          =   345
         Index           =   0
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "結案原因"
      Height          =   2655
      Left            =   180
      TabIndex        =   17
      Top             =   3060
      Width           =   8600
      Begin VB.OptionButton Option1 
         Caption         =   "自請撤回"
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   41
         Top             =   2310
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "認為本所收費太高"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   40
         Top             =   2100
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "對本所服務不滿意"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   39
         Top             =   1890
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶自行處理"
         Height          =   315
         Index           =   10
         Left            =   4350
         TabIndex        =   36
         Top             =   210
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "其他"
         Height          =   255
         Index           =   5
         Left            =   4350
         TabIndex        =   35
         Top             =   840
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶另案重提"
         Height          =   255
         Index           =   3
         Left            =   4350
         TabIndex        =   34
         Top             =   1140
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "已轉由他所處理"
         Height          =   255
         Index           =   1
         Left            =   4350
         TabIndex        =   33
         Top             =   540
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶已倒閉"
         Height          =   255
         Index           =   11
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶認為核駁（敗訴、被異議、被評定）合理"
         Height          =   315
         Index           =   8
         Left            =   150
         TabIndex        =   8
         Top             =   1440
         Width           =   4100
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶無法再提供主管機關所要求的資料"
         Height          =   315
         Index           =   7
         Left            =   150
         TabIndex        =   7
         Top             =   1140
         Width           =   4100
      End
      Begin VB.OptionButton Option1 
         Caption         =   "放棄（停產、銷路不佳）"
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   6
         Top             =   840
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶已遷移，無法聯絡"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   540
         Width           =   2325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "以上項目不需填結案理由"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   150
         TabIndex        =   42
         Top             =   1770
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "以上項目若無回覆單須進一步填理由"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4350
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   2880
      End
      Begin MSForms.TextBox Text1 
         Height          =   960
         Left            =   4350
         TabIndex        =   37
         Top             =   1650
         Width           =   4155
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "7329;1685"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   2
      Left            =   3990
      TabIndex        =   10
      Top             =   30
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1695
      Left            =   180
      TabIndex        =   16
      Top             =   1320
      Width           =   8595
      _ExtentX        =   15177
      _ExtentY        =   3006
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   3
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
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   3
      Top             =   405
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   405
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Top             =   405
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   0
      Top             =   405
      Width           =   525
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   345
      Left            =   5010
      TabIndex        =   27
      Top             =   30
      Width           =   1875
      Begin VB.TextBox txtPCnt 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   1
         Left            =   1230
         MaxLength       =   1
         TabIndex        =   28
         Text            =   "2"
         Top             =   30
         Width           =   270
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "無期限閉卷(&P)　　份"
         Height          =   345
         Index           =   1
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSForms.Label lblCU01Nm 
      Height          =   255
      Left            =   1160
      TabIndex        =   32
      Top             =   1020
      Width           =   4035
      Size            =   "7117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseNm 
      Height          =   255
      Left            =   1160
      TabIndex        =   31
      Top             =   720
      Width           =   7770
      VariousPropertyBits=   27
      Size            =   "13705;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "label2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   210
      TabIndex        =   30
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   11
      Left            =   210
      TabIndex        =   24
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "申  請  人："
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   23
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   22
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   6210
      TabIndex        =   21
      Top             =   1020
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號："
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   20
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   6210
      TabIndex        =   19
      Top             =   450
      Width           =   1740
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1470
      X2              =   3030
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   18
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm210133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; Text1、grd1改字型=新細明體-ExtB、Label1(1)=>lblCaseNm、Label1(3)=>lblCU01Nm (Printer列印未改)
'Memo by Lydia 2019/07/01 表單名稱:期限資料結案單=>案件結案單
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim m_row As Integer, i As Integer, j As Integer
Dim strNP02 As String
Dim strNP03 As String
Dim strNP04 As String
Dim strNP05 As String
Dim m_Nation As String
Dim m_CurCP(1 To 4) As String '現在資料的本所號
Dim m_iDiscount As Integer '可減免退費金額
Dim m_iYear1 As Integer '減免退費起始年度
Dim m_iYear2 As Integer '減免退費終止年度
Dim m_PA08 As String, m_PA10 As String 'Added by Morgan 2012/10/1
'Add by Amy 2014/05/23
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2014/6/19
'Add By Sindy 2015/1/5
'********************************************************
Public m_strSaveFiles As String '新增附件
Public m_F0301 As String '表單編號
Dim m_F0303 As String '總收文號或本所案號
Dim m_F0304 As String '下一程序序號
Dim m_F0305 As String '結案理由
Dim m_F0308 As String '下一處理人員
Dim m_F0309 As String
Public m_SetFlowEmp1 As String '設定簽核人員1
Public m_NP01 As String, m_NP22 As String
Dim m_CP13 As String, m_F0316 As String
Dim m_AttachPath As String '附件暫存區
Dim m_CU10 As String 'Add By Sindy 2015/4/8 申請人1國籍
Dim m_FCfagent As String 'Add By Sindy 2015/4/8 FC代理人
'********************************************************
Dim bolNoFlow As Boolean 'Add by Amy 2020/05/18
Dim colNP07 As Integer 'Added by Lydia 2021/01/26

'Add By Sindy 2014/6/19
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdFile_Click()
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.lblCaseNo = strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05
   frm090801_8.Show vbModal
End Sub

'Modify By Sindy 2023/5/19
Private Function GetF0316(strNP10 As String) As String
   'If m_F0316 = "" Then m_F0316 = m_CP13 '智權人員
   If strNP10 <> "" Then
      GetF0316 = strNP10
   Else
      GetF0316 = strUserNum 'Add By Sindy 2016/9/23 P-092509 吳中一,操作人員是王俊剴
   End If
   '若智權人員已離職,則以Login人員代替
   'If ChkStaffST04(GetF0316, False) = True Then
   If ChkStaffST04(GetF0316, False) = True Or Left(GetF0316, 1) <= "6" Then
      GetF0316 = strUserNum
   End If
   'Add By Sindy 2021/11/10 判斷是否為MCT
   If ChkMCTF0XSales(PUB_GetAKindSalesNo(UCase(txt1(0)), txt1(1), Left(txt1(2) & "0", 1), Left(txt1(3) & "00", 2)), strUserNum) = True Then
      GetF0316 = PUB_GetAKindSalesNo(UCase(txt1(0)), txt1(1), Left(txt1(2) & "0", 1), Left(txt1(3) & "00", 2))
   End If
   '2021/11/10 END
   'Add By Sindy 2023/5/19 專利處程序人員代填郭經理或P1004的結案單，都以填表人的表單簽核走流程。
   If PUB_GetStaffST15(strUserNum, "1") = "P12" Then
      GetF0316 = strUserNum
   End If
   '2023/5/19 END
   'Add by Amy 2025/08/27 目前智權為P2006(商標智權人員),新增與退回之簽核人員不一致,Sindy與秀玲討論後都抓P2006之簽核人員
   '  ex:CFT-022430 (蒲璇已操作) 無期限之結案單,Amy 測式[退回]再送出,發現新增時抓蒲璇之簽核人員,退回抓P2006之簽核人員
   If m_CP13 = "P2006" Then
      GetF0316 = m_CP13
   End If
   'end 2025/08/27
End Function

Private Sub cmdFlowEmp_Click()
   If m_row > 0 Then
      If GRD1.TextMatrix(m_row, 13) <> "" Then 'NP01
         m_F0316 = GRD1.TextMatrix(m_row, 11) '智權人員
      End If
   End If
   m_F0316 = GetF0316(m_F0316) 'Modify By Sindy 2023/5/19
   
   frm210133_1.m_SetFlowKind = Flow_結案單
   frm210133_1.lblST01 = m_F0316
   frm210133_1.lblST02 = GetPrjSalesNM(m_F0316)
   frm210133_1.doQuery
   frm210133_1.Show vbModal
End Sub

'Modify by Amy 2025/08/13 原:Private
Public Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 2
         m_F0301 = "" 'Add By Sindy 2015/1/5
         If Trim(txt1(0)) = "" Or Trim(txt1(1)) = "" Then
            MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         GRD1.MousePointer = flexHourglass
         doQuery
         GRD1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
         
      Case 0
         '檢查條件
         If TxtValidate(0) = False Then Exit Sub
'         If m_row <> 0 Then
'            If grd1.TextMatrix(m_row, 14) = "" Then
'                MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
'                txt1_GotFocus 0
'                Exit Sub
'            Else
''               If Option1(5) = True And Trim(Text1) = "" Then
''                  MsgBox "結案理由點選其他時，請輸入說明！"
''                  Text1.SetFocus
''                  Exit Sub
''               End If
'               bolChk = False
'               For i = 0 To 11
'                  If Option1(i) = True Then
'                     bolChk = True
'                     Exit For
'                  End If
'               Next i
'               If bolChk = False Then
'                  MsgBox "請勾選結案理由！"
'                  Exit Sub
'               End If
'
'               m_iDiscount = 0: strSpecial = 0
'               m_CurCP(1) = strNP02: m_CurCP(2) = strNP03: m_CurCP(3) = strNP04: m_CurCP(4) = strNP05
'               '辦理減免退費提醒
'               If PUB_GetCaseDiscStat(strNP02 & strNP03 & strNP04 & strNP05) = "Y" Then
'                  Call PUB_CheckYearFeeReturn(m_CurCP, False, m_iDiscount, m_iYear1, m_iYear2)
'               End If
'               If m_iDiscount > 0 Then strSpecial = "1"
'
'               'Modify by Amy 2014/05/23 +if
'               If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
'                    '開放專利處部份智權(A7)同仁資料給彥葶(A8)代為處理
'                    If InStr(Pub_GetSpecMan("A7"), grd1.TextMatrix(m_row, 11)) > 0 Then
'                    Else
'                        MsgBox "您無權限將此客戶案件任意結案！", vbExclamation, "操作錯誤！"
'                        Exit Sub
'                    End If
'               'end 2014/05/23
'               '2011/7/6 ADD BY SONIA 檢查操作人員與結案期限智權人員 CFP-018696
'               ElseIf grd1.TextMatrix(m_row, 11) = strUserNum Or Pub_StrUserSt03 = "M51" Or grd1.TextMatrix(m_row, 11) < "6" Then
'               'Added by Morgan 2012/6/21
'               '國外部,同部門(前兩碼同)都可印
'               ElseIf Left(Pub_StrUserSt03, 1) = "F" Then
'                  If Left(PUB_GetStaffST15(grd1.TextMatrix(m_row, 11), "1"), 2) <> Left(PUB_GetStaffST15(strUserNum, "1"), 2) Then
'                     MsgBox "非相同部門客戶案件不可任意結案！"
'                     Exit Sub
'                  End If
'               'end 2012/6/21
'               'add by sonia 2014/10/30 美珍可操作林總案件
'               ElseIf strUserNum = "77027" And grd1.TextMatrix(m_row, 11) = "94007" Then
'               'end 2014/10/30
'               Else
'                  '若原智權人員離職則該區同仁都可結案
'                  '2011/7/14 modify by sonia 台中林協理客戶開放可由該區其他人結案
'                  If GetStaffName(grd1.TextMatrix(m_row, 11)) = "" Or grd1.TextMatrix(m_row, 11) = "68096" Then
'                     If PUB_GetStaffST15(grd1.TextMatrix(m_row, 11), "1") = PUB_GetStaffST15(strUserNum, "1") Then
'                     Else
'                        'modify by sonia 2013/7/2 帶人主管也可以結案
'                        'MsgBox "非相同業務區客戶案件不可任意結案！"
'                        'Exit Sub
'                        stIdList = PUB_GetSalesList(strUserNum, PUB_GetStaffST15(strUserNum, "1"), PUB_GetStaffST15(strUserNum, "1"), PUB_GetST06(strUserNum))
'                        If Pub_StrST52 = True And InStr(stIdList, grd1.TextMatrix(m_row, 11)) > 0 Then
'                        Else
'                           MsgBox "非相同業務區客戶案件不可任意結案！"
'                           Exit Sub
'                        End If
'                        '2013/7/2 end
'                     End If
'                  Else
'                     MsgBox "非本人案件不可任意結案！"
'                     Exit Sub
'                  End If
'               End If
'               '2011/7/6 END
'
'               'Added by Morgan 2012/9/28
'               '一案兩請的新型年費填寫結案單，若發明案尚未審定或核駁且未閉卷時，要提醒並由使用者確認後將提醒文字印在結案單之備註欄
'               If strNP02 = "P" And (m_Nation = "000" Or (m_Nation = "020" And Val(m_PA10) >= 20091001)) And m_PA08 = "2" And grd1.TextMatrix(m_row, colnp07) = "605" Then
'                  strExc(0) = "select pa01,pa02,pa03,pa04 from casemap,patent where cm10='3' and cm01='" & strNP02 & "' and cm02='" & strNP03 & "' and cm03='" & strNP04 & "' and cm04='" & strNP05 & "'" & _
'                     " and pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa08='1' and pa57 is null and (pa16 is null or pa16='2')" & _
'                     " union select pa01,pa02,pa03,pa04 from casemap,patent where cm10='3' and cm05='" & strNP02 & "' and cm06='" & strNP03 & "' and cm07='" & strNP04 & "' and cm08='" & strNP05 & "'" & _
'                     " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa08='1' and pa57 is null and (pa16 is null or pa16='2')"
'
'                  intI = 1
'                  Set Rs = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If m_Nation = "020" Then
'                        strExc(2) = "※此案為大陸之一案兩請，實用新型放棄繳年費，則大陸發明案就不予專利。"
'                     Else
'                        strExc(2) = "※此案為一案兩請，發明專利審定前，新型專利權若因未繳年費而當然消滅者，則將不予專利。"
'                     End If
'                     If MsgBox(strExc(2) & vbCrLf & "是否確認要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation, "一案兩請新型結案提醒") = vbYes Then
'                        Text1 = strExc(2) & Text1
'                     Else
'                        Exit Sub
'                     End If
'                  End If
'               End If
'               'end 2012/9/28
'
''2012/1/16 cancel by sonia 中所新人要印給主管簽
''               '2012/1/2 add by sonia 楊特助說即然要做結案通知,就不要再允許列印
''               If strNP02 = "T" Or strNP02 = "TF" Then
''                  If grd1.TextMatrix(m_row, colnp07) = "102" Or grd1.TextMatrix(m_row, colnp07) = "716" Then
''                     MsgBox "內商延展或第二期註冊費案件請改至 業務期限資料查詢 做結案通知！"
''                     Exit Sub
''                  End If
''               End If
''               '2012/1/2 end
               Call PrintData(Index) 'Modify By Sindy 2013/7/16
'            End If
'         Else
'            If grd1.Rows = 2 And grd1.TextMatrix(1, 13) = "" Then
'                MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
'                txt1_GotFocus 0
'                Exit Sub
'            Else
'                MsgBox "請先選擇一筆要列印的資料！", vbCritical, "操作錯誤！"
'                Exit Sub
'            End If
'         End If
         
      Case 1
'         If txt1(0) = "" Or txt1(1) = "" Then
'            MsgBox "請輸入本所案號！"
'            If txt1(1) = "" Then txt1(1).SetFocus
'            If txt1(0) = "" Then txt1(0).SetFocus
'            Exit Sub
'         End If
'         If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then
'            MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
'            txt1_GotFocus 0
'            Exit Sub
'         End If
'         bolChk = False
'         For i = 0 To 11
'            If Option1(i) = True Then
'               bolChk = True
'               Exit For
'            End If
'         Next i
'         If bolChk = False Then
'            MsgBox "請勾選結案理由！"
'            Exit Sub
'         End If
         '檢查條件
         If TxtValidate(1) = False Then Exit Sub
         Call PrintData(Index) 'Modify By Sindy 2013/7/16
      Case 3
         Unload Me
   Case Else
   End Select
End Sub

'Add By Sindy 2013/7/16
Private Sub PrintData(Index As Integer)
Dim ii As Integer
Dim intStarY As Integer
Dim ii_Row As Integer
Dim strText1 As String
Dim arrText1, iiiii As Integer 'Add By Sindy 2010/9/15
Dim strTitle As String 'Add by Amy 2020/03/27

   'Add By Sindy 2015/3/24 杜經理提加msg
   'Modify By Sindy 2015/4/8 +if 申請人1國籍非台灣且有FC代理人且操作者部門為Fxx者不必詢問
   If Not (m_CU10 <> "000" And m_FCfagent <> "" And Mid(Pub_StrUserSt03, 1, 1) = "F") Then
   '2015/4/8 END
      If m_strSaveFiles = "" Then
         'Modify By Sindy 2019/8/12
         'If MsgBox("缺回覆單，確定是否繼續列印？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         If MsgBox("回覆單未匯入，要匯入回覆單？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         '2019/8/12 END
            Exit Sub
         End If
      End If
   End If

   'Add By Sindy 2015/5/18 欲上傳回覆單,先檢查回覆單狀況
   If m_strSaveFiles <> "" Then
      If PUB_ChkIsReplyFile(strNP02 & strNP03 & strNP04 & strNP05, , , , IIf(Index = 1, "", "" & GRD1.TextMatrix(m_row, 14))) = True Then
         MsgBox "此案號該期限回覆單已上傳，不可作業，若有疑問請通知電腦心中協助處理！", vbCritical
         Exit Sub
      End If
   End If
   '2015/5/18 END

   'Add By Sindy 2015/2/9
   '儲存回覆單：統一先存本所案號收進系統，等產生B類文號時再掛進回覆單
   'Modify By Sindy 2015/5/18
   'If PUB_UpdReplyFile(m_strSaveFiles, "", strNP02, strNP03, strNP04, strNP05) = False Then Exit Sub
   If PUB_UpdReplyFile(m_strSaveFiles, "", strNP02, strNP03, strNP04, strNP05, , IIf(Index = 1, "", "" & GRD1.TextMatrix(m_row, 14))) = False Then Exit Sub
   '2015/2/9 END

   Screen.MousePointer = vbHourglass
   GRD1.MousePointer = flexHourglass

   For ii = 1 To Val(txtPCnt(Index))
      Printer.Orientation = 1
      Printer.PaperSize = vbPRPSA4
      Printer.Font.Size = 16
      Printer.Font.Bold = True
      Printer.FontName = "細明體"

      'Modify by Amy 2020/03/27 取消 台一國際專利商標事務所 文字
      strTitle = "結　　　案　　　記　　　錄　　　單"
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTitle) / 2)
      Printer.CurrentY = 500
      Printer.Print strTitle
      Printer.Font.Bold = False
      'end 2020/03/27

      ' 列印國內案件接洽及結案記錄單
      'Modify By Sindy 2012/6/19 +bolFullPrint = False
      If Index = 1 Then '無期限閉卷
         'Modify By Sindy 2015/5/1 +, IIf(m_strSaveFiles <> "", True, False) 是否有電子回覆單
         g_PrtForm001.PrintForm "000000", "" & strNP02, "" & strNP03, "" & strNP04, "" & strNP05, , "", , True, , False, IIf(m_strSaveFiles <> "", True, False)
      Else
         If GRD1.TextMatrix(m_row, 13) = "" Then
            'Modify By Sindy 2015/5/1 +, IIf(m_strSaveFiles <> "", True, False) 是否有電子回覆單
            g_PrtForm001.PrintForm "" & GRD1.TextMatrix(m_row, 14), "" & strNP02, "" & strNP03, "" & strNP04, "" & strNP05, , "", , True, Trim(Left(GRD1.TextMatrix(m_row, 4), 4)), False, IIf(m_strSaveFiles <> "", True, False)
         Else
            'Modify By Sindy 2015/5/1 +, IIf(m_strSaveFiles <> "", True, False) 是否有電子回覆單
            g_PrtForm001.PrintForm "" & GRD1.TextMatrix(m_row, 14), "" & strNP02, "" & strNP03, "" & strNP04, "" & strNP05, , "", , True, , False, IIf(m_strSaveFiles <> "", True, False)
         End If
      End If

      intStarY = 5500
      Printer.Line (500, intStarY)-(11200, intStarY + 7000), , B '4200
      Printer.Line (500, intStarY)-(1100, intStarY + 7000), , B '4200

      Printer.CurrentX = 700
      Printer.CurrentY = intStarY + (600 * 1)
      Printer.Print "結"
      Printer.CurrentX = 700
      Printer.CurrentY = intStarY + (600 * 2)
      Printer.Print "案"
      Printer.CurrentX = 700
      Printer.CurrentY = intStarY + (600 * 3)
      Printer.Print "記"
      Printer.CurrentX = 700
      Printer.CurrentY = intStarY + (600 * 4)
      Printer.Print "錄"

      'Modify by Amy 2022/03/16 調整位置,因2022/03/08 杜經理請作「1.對本所服務不滿意」(index0)/「3.認為本所收費太高」(index2) /「9.自請撤回」(index9) 不使用
'      Printer.CurrentX = 1700
'      Printer.CurrentY = intStarY + (400 * 1)
'      Printer.Print IIf(Option1(0) = False, "□", "■") & Option1(0).Caption; '對本所服務不滿意
'      Printer.CurrentX = 1700
'      Printer.CurrentY = intStarY + (400 * 2)
'      Printer.Print IIf(Option1(1) = False, "□", "■") & Option1(1).Caption
'      Printer.CurrentX = 1700
'      Printer.CurrentY = intStarY + (400 * 3)
'      Printer.Print IIf(Option1(2) = False, "□", "■") & Option1(2).Caption '認為本所收費太高
'      Printer.CurrentX = 1700
'      Printer.CurrentY = intStarY + (400 * 4)
'      Printer.Print IIf(Option1(3) = False, "□", "■") & Option1(3).Caption
'      Printer.CurrentX = 1700
'      Printer.CurrentY = intStarY + (400 * 5)
'      Printer.Print IIf(Option1(4) = False, "□", "■") & Option1(4).Caption
'      Printer.CurrentX = 1700
'      Printer.CurrentY = intStarY + (400 * 6)
'      Printer.Print IIf(Option1(5) = False, "□", "■") & Option1(5).Caption
'      Printer.CurrentX = 1900
'      Printer.CurrentY = intStarY + (400 * 7)
      '*** 左半項目 ***
      Printer.CurrentX = 1500
      Printer.CurrentY = intStarY + (400 * 1)
      Printer.Print IIf(Option1(11) = False, "□", "■") & Option1(11).Caption; '客戶已倒閉
      Printer.CurrentX = 1500
      Printer.CurrentY = intStarY + (400 * 2)
      Printer.Print IIf(Option1(4) = False, "□", "■") & Option1(4).Caption '客戶已遷移...
      Printer.CurrentX = 1500
      Printer.CurrentY = intStarY + (400 * 3)
      Printer.Print IIf(Option1(6) = False, "□", "■") & Option1(6).Caption '放棄...
      Printer.CurrentX = 1500
      Printer.CurrentY = intStarY + (400 * 4)
      Printer.Print IIf(Option1(7) = False, "□", "■") & Option1(7).Caption '客戶無法再提供主管機關...
      Printer.CurrentX = 1500
      Printer.CurrentY = intStarY + (400 * 5)
      Printer.Print IIf(Option1(8) = False, "□", "■") & Option1(8).Caption '客戶認為核駁...
      '*** end 左半項目 ***
      'Printer.Font.Underline = True
      '其他說明欄位
      'Add by Amy 2022/04/08
      Printer.CurrentX = 1500
      Printer.CurrentY = intStarY + (400 * 7)
      'end 2022/04/08
      ii_Row = 0
      If Trim(Text1) > "" Then
         strText1 = Trim(Text1)
         'Modify By Sindy 2010/9/15
         arrText1 = Split(strText1, vbCrLf)
         If UBound(arrText1) > 0 Then
            For iiiii = 0 To UBound(arrText1)
               Printer.CurrentX = 1700 'Modify by Amy 原:1900
               Printer.CurrentY = intStarY + (400 * 7) + (300 * ii_Row)
               Printer.Print arrText1(iiiii)
               ii_Row = ii_Row + 1
            Next iiiii
         '2010/9/15 End
         Else
            If Len(strText1) > 38 Then
               Do While Len(strText1) > 38
                  Printer.CurrentX = 1700 'Modify by Amy 原:1900
                  Printer.CurrentY = intStarY + (400 * 7) + (300 * ii_Row)
                  Printer.Print Mid(strText1, 1, 38)
                  strText1 = Right(strText1, Len(strText1) - 38)
                  ii_Row = ii_Row + 1
               Loop
               Printer.CurrentX = 1700 'Modify by Amy 原:1900
               Printer.CurrentY = intStarY + (400 * 7) + (300 * ii_Row)
               Printer.Print strText1
            Else
               Printer.Print strText1
            End If
         End If
      End If
      'Printer.Font.Underline = False
     
'      Printer.CurrentX = 5700
'      Printer.CurrentY = intStarY + (400 * 1)
'      Printer.Print IIf(Option1(6) = False, "□", "■") & Option1(6).Caption
'      Printer.CurrentX = 5700
'      Printer.CurrentY = intStarY + (400 * 2)
'      Printer.Print IIf(Option1(7) = False, "□", "■") & Option1(7).Caption
'      Printer.CurrentX = 5700
'      Printer.CurrentY = intStarY + (400 * 3)
'      Printer.Print IIf(Option1(8) = False, "□", "■") & Option1(8).Caption
'      Printer.CurrentX = 5700
'      Printer.CurrentY = intStarY + (400 * 4)
'      Printer.Print IIf(Option1(9) = False, "□", "■") & Option1(9).Caption '自請撤回
'      Printer.CurrentX = 5700
'      Printer.CurrentY = intStarY + (400 * 5)
'      Printer.Print IIf(Option1(10) = False, "□", "■") & Option1(10).Caption
'      Printer.CurrentX = 5700
'      Printer.CurrentY = intStarY + (400 * 6)
'      Printer.Print IIf(Option1(11) = False, "□", "■") & Option1(11).Caption
      '*** 右半項目 ***
       Printer.CurrentX = 6800
      Printer.CurrentY = intStarY + (400 * 1)
      Printer.Print IIf(Option1(10) = False, "□", "■") & Option1(10).Caption '客戶自行處理
      Printer.CurrentX = 6800
      Printer.CurrentY = intStarY + (400 * 2)
      Printer.Print IIf(Option1(1) = False, "□", "■") & Option1(1).Caption '已轉由他所處理
      Printer.CurrentX = 6800
      Printer.CurrentY = intStarY + (400 * 3)
      Printer.Print IIf(Option1(5) = False, "□", "■") & Option1(5).Caption '其他
      Printer.CurrentX = 6800
      Printer.CurrentY = intStarY + (400 * 4)
      Printer.Print IIf(Option1(3) = False, "□", "■") & Option1(3).Caption '客戶另案重提
      '*** end 右半項目 ***
      'end 2022/03/16
      
      'intStarY = 10000
      intStarY = 13500
      Printer.Line (500, intStarY)-(11200, intStarY + 2500), , B
      Printer.Line (500, intStarY)-(11200, intStarY + 500), , B
      Printer.CurrentX = 1200
      Printer.CurrentY = 13000 '10100
      'Modify By Sindy 2024/10/17
      'Printer.Print "專 業 主 管"
      Printer.Print "業 務 主 管"
      '2024/10/17 END
      Printer.CurrentX = 3900
      Printer.CurrentY = 13000 '10100
      'Modify By Sindy 2024/10/17
      'Printer.Print "業 務 主 管"
      Printer.Print "專 業 主 管"
      '2024/10/17 END
      Printer.CurrentX = 6800
      Printer.CurrentY = 13000 '10100
      'Modify By Sindy 2024/10/17
      If InStr(txt1(0), "L") > 0 Then
         Printer.Print "承辦律師"
      Else
      '2024/10/17 END
         Printer.Print "主　　任"
      End If
      Printer.CurrentX = 9500
      Printer.CurrentY = 13000 '10100
      'Modify By Sindy 2024/10/17
      If InStr(txt1(0), "L") > 0 Then
         Printer.Print "協辦助理"
      Else
      '2024/10/17 END
         Printer.Print "程序人員"
      End If
      Printer.Line (500, intStarY)-(3200, intStarY + 2500), , B
      Printer.Line (500, intStarY)-(5900, intStarY + 2500), , B
      Printer.Line (500, intStarY)-(8600, intStarY + 2500), , B

      Printer.EndDoc
   Next ii
   ShowPrintOk
   'Add By Sindy 2015/2/9 刪除匯入來源的回覆單
   Call PUB_DelPCOrgFile(m_strSaveFiles): m_strSaveFiles = ""
   '2015/2/9 END
   GRD1.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

'清除欄位值
Sub ClearData()
   Dim opt
   Label1(7) = ""
   lblCaseNm.Caption = ""
   lblCU01Nm.Caption = ""
   Label1(5) = ""
   m_Nation = ""
   For i = 0 To 11
      Option1(i).Value = False
   Next i
   Text1 = ""
   m_strSaveFiles = ""
   m_F0303 = ""
   m_F0304 = ""
   m_F0305 = ""
   GRD1.Clear: SetDataListWidth 'Add By Sindy 2015/1/8
   cmdFile.Enabled = False 'Modify by Amy 2018/09/03  原:true 因智權人員未先輸本案查詢,直接按回覆單匯入
   m_CU10 = "" 'Add By Sindy 2015/4/8 申請人1國籍
   m_FCfagent = "" 'Add By Sindy 2015/4/8 FC代理人
End Sub

'Modify By Sindy 2014/6/19
'Sub doQuery()
Public Function doQuery() As Boolean
'2014/6/19 END
Dim intRow As Integer
Dim m_IsClose As String 'Add By Sindy 2015/1/21
Dim strChkLimitEmp As String, stIdLeader As String 'Add By Sindy 2015/2/5
Dim stIdList As String       'add by sonia 2013/7/2
Dim Rs As New ADODB.Recordset
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer, strMsg As String, arrTmp 'Add by Amy 2025/05/19
   
On Error GoTo ErrHnd
   
   doQuery = False 'Add By Sindy 2014/6/19
   m_row = 0
   'Modify by Amy 2018/08/27 原預設列印,改送出鈕為預設-文雄
   Frame2.Visible = False
   Frame3.Visible = False
   Frame4.Visible = True: cmdFlowEmp.Visible = True 'Add By Sindy 2014/12/31 結案單電子化
   'end 2018/08/27
   '清除欄位值
   Call ClearData
   
   'Add By Sindy 2015/1/5
   If m_F0301 <> "" Then
      cmdOK(2).Visible = False
      'Modify by Amy 2025/05/19 +if FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
'      If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         'Modify By Sindy 2025/7/29 增加串NP的SQL; 因CB4027634(IDS)CFP-034735,在主檔此文號是P案(P-134557)
         strSql = "SELECT FLOW003.*,cp09,cp01,cp02,cp03,cp04,CloseCaseMain.*" & _
               " From FLOW003,caseprogress,CloseCaseMain" & _
               " WHERE F0301='" & m_F0301 & "' and CCM02=cp09(+) And F0301=CCM01(+) and CCM03 is null" & _
               " union SELECT FLOW003.*,np01,np02,np03,np04,np05,CloseCaseMain.*" & _
               " From FLOW003,nextprogress,CloseCaseMain" & _
               " WHERE F0301='" & m_F0301 & "' and CCM02=np01(+) and CCM03=np22(+) And F0301=CCM01(+) and CCM03 is not null"
'      Else
'         strSql = "SELECT FLOW003.*,cp09,cp01,cp02,cp03,cp04" & _
'               " From FLOW003,caseprogress" & _
'               " WHERE F0301='" & m_F0301 & "' and F0303=cp09(+)"
'      End If
      'end 2025/05/19
      intI = 1
      Set Rs = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(Rs.Fields("cp09")) Then
            strNP02 = Rs.Fields("cp01")
            strNP03 = Rs.Fields("cp02")
            strNP04 = Rs.Fields("cp03")
            strNP05 = Rs.Fields("cp04")
         Else
            'Modify by Amy 2025/05/19 +if FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
            If strSrvDate(1) >= FCP結案單電子化啟用日 Then
               strNP02 = Left(Rs.Fields("CCM02"), Len(Rs.Fields("CCM02")) - 9)
               strNP03 = Mid(Rs.Fields("CCM02"), Len(strNP02) + 1, 6)
               strNP04 = Mid(Rs.Fields("CCM02"), Len(strNP02) + 7, 1)
               strNP05 = Right(Rs.Fields("CCM02"), 2)
            Else
               strNP02 = Left(Rs.Fields("F0303"), Len(Rs.Fields("F0303")) - 9)
               strNP03 = Mid(Rs.Fields("F0303"), Len(strNP02) + 1, 6)
               strNP04 = Mid(Rs.Fields("F0303"), Len(strNP02) + 7, 1)
               strNP05 = Right(Rs.Fields("F0303"), 2)
            End If
            'end 2025/05/19
         End If
         'Modify by Amy 2025/05/19 +if FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
         If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            m_F0303 = Rs.Fields("CCM02") '總收文號/案號
            m_F0304 = "" & Rs.Fields("CCM03") '下一程序號
            '結案理由代碼
            If "" & Rs.Fields("CCM04") <> "" Then
               If Rs.Fields("CCM04") = "99" Then
                  Option1(5).Value = True
               Else
                  If Val(Rs.Fields("CCM04")) <= 5 Then
                     Option1(Val(Rs.Fields("CCM04")) - 1).Value = True
                  Else
                     Option1(Val(Rs.Fields("CCM04"))).Value = True
                  End If
               End If
            End If
            '結案理由
            Text1 = "" & Rs.Fields("CCM05")
         Else
            m_F0303 = Rs.Fields("F0303")
            m_F0304 = "" & Rs.Fields("F0304")
            
            If "" & Rs.Fields("F0305") <> "" Then
               If Rs.Fields("F0305") = "99" Then
                  Option1(5).Value = True
               Else
                  If Val(Rs.Fields("F0305")) <= 5 Then
                     Option1(Val(Rs.Fields("F0305")) - 1).Value = True
                  Else
                     Option1(Val(Rs.Fields("F0305"))).Value = True
                  End If
               End If
            End If
            Text1 = "" & Rs.Fields("F0306")
         End If
         'end 2025/05/19
         m_F0316 = "" & Rs.Fields("F0316")
      End If
      
'*** 下載回覆單 (此處有改, 需確認frm210133_F 是否也要改) ***
      'Modify By Sindy 2015/5/18
      'If PUB_ChkIsReplyFile(strNP02, strNP03, strNP04, strNP05, m_strSaveFiles) = True Then
       'Modify By Sindy 2025/5/9 + , , m_AttachPath
       'Modify by Amy 2025/05/19 退回 從一般作業->目前表單進入->修改 檔案會下至 c:\App.path\員編 與Form_Load 設定不符
      If m_AttachPath = "" Then
         If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
            MsgBox "附件資料夾建立失敗" & vbCrLf & _
                           strExc(9) & vbCrLf & "請洽電腦中心!"
         End If
      End If
      'end 2025/05/19
      If PUB_ChkIsReplyFile(strNP02 & strNP03 & strNP04 & strNP05, m_strSaveFiles, , , m_F0301, , m_AttachPath) = True Then
      '2015/5/18 ENd
         If m_strSaveFiles <> "" Then
            'Modify by Amy 2025/05/19 多檔會出現無法開啟的訊息,因未切割檔案
            'Modify By Sindy 2015/5/18
'            If PUB_GetAttachFile_CPP(strNP02 & strNP03 & strNP04 & strNP05, m_strSaveFiles, m_AttachPath) = False Then
'            'If PUB_GetAttachFile_CPP(m_F0301, m_strSaveFiles, m_AttachPath) = False Then
'            '2015/5/18 ENd
'               MsgBox "無法儲存欲開啟的檔案[ " & m_strSaveFiles & " ]！"
'            End If
            arrTmp = Split(m_strSaveFiles, "&")
            strExc(9) = m_AttachPath
            If Right(strExc(9), 1) <> "\" Then strExc(9) = strExc(9) & "\"
            For ii = LBound(arrTmp) To UBound(arrTmp)
               'Memo by Amy 2025/05/19 避免多筆未串路徑而抓不到資料,故先取代再統一加
               If PUB_GetAttachFile_CPP(strNP02 & strNP03 & strNP04 & strNP05, "" & Replace(arrTmp(ii), strExc(9), ""), m_AttachPath) = False Then
                  strMsg = strMsg & ";" & arrTmp(ii)
               End If
            Next ii
            If strMsg <> "" Then
               MsgBox "無法儲存欲開啟的檔案" & vbCrLf & Replace(Mid(strMsg, 2), ";", vbCrLf) & " ！"
            End If
            'end 2025/05/19
         End If
      End If
'*** End 下載回覆單 (此處有改, 需確認frm210133_F 是否也要改) ***
   Else
   '2015/1/5 END
      strNP02 = UCase(txt1(0))
      strNP03 = txt1(1)
      strNP04 = Left(txt1(2) & "0", 1)
      strNP05 = Left(txt1(3) & "00", 2)
      
      'Add By Sindy 2015/1/5
      If UCase(TypeName(m_PrevForm)) = UCase("frm210145") Then
         '檢查是否有結案單
         strSql = "SELECT NP01,NP24 FROM NextProgress,FLOW003" & _
                  " WHERE NP02='" & strNP02 & "' and NP03='" & strNP03 & "' and NP04='" & strNP04 & "' and NP05='" & strNP05 & "'" & _
                  " and np24 is not null and np24=f0301(+) and F0301 is not null"
         intI = 1
         Set Rs = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'MsgBox "結案處理中，請至相關程式查詢！", vbCritical
            MsgBox "結案處理中，請至案件目前表單查詢！", vbCritical
            Exit Function
         End If
'         '檢查是否還有該文號未結案
'         strSql = "SELECT NP01,NP24 FROM NextProgress" & _
'                  " WHERE NP02='" & strNP02 & "' and NP03='" & strNP03 & "' and NP04='" & strNP04 & "' and NP05='" & strNP05 & "'" & _
'                  " and np06 is null and np24 is null " & strNpSqlOfNoSalesDuty
'         intI = 1
'         Set Rs = ClsLawReadRstMsg(intI, strSql)
'         '有未結,繼續結案...
'         If intI = 0 Then
'         End If
      ElseIf UCase(TypeName(m_PrevForm)) = UCase("frm100123") Then
         'Modify By Sindy 2020/5/19
         If m_NP22 = 0 Then
            m_NP01 = "": m_NP22 = "" '檢查無期限
         End If
         '2020/5/19 END
         '檢查是否已有結案單
         If ChkFlowFormExists(Flow_結案單, m_NP01, m_NP22, strNP02, strNP03, strNP04, strNP05, "F0301", m_F0301) = True Then
            'MsgBox "結案處理中，請至相關程式查詢！", vbCritical
            MsgBox "結案處理中，請至案件目前表單查詢！", vbCritical
            Exit Function
         End If
      End If
      '2015/1/5 END
   End If
   'Modify By Sindy 2015/3/18 開放紙本也可放回覆單
'   If strNP02 = "P" Or strNP02 = "CFP" Then
'      cmdFile.Visible = True
'   Else
'      cmdFile.Visible = False
'   End If
   
   'Modify By Sindy 2015/1/21 +讀取是否閉卷欄
   strSql = "SELECT TM12,TM05||TM06||TM07,TM23||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,TM10,' ',TM11,TM29,CU10,TM44 as PA75" & _
                " From Trademark, nation, Customer" & _
                " WHERE TM01='" & strNP02 & "' AND TM02='" & strNP03 & "' AND TM03='" & strNP04 & "' AND TM04='" & strNP05 & "'" & _
                " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
                " AND TM10=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT PA11,PA05||PA06||PA07,PA26||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,PA09,PA08,PA10,PA57,CU10,PA75" & _
                " From Patent, nation, Customer" & _
                " WHERE PA01='" & strNP02 & "' AND PA02='" & strNP03 & "' AND PA03='" & strNP04 & "' AND PA04='" & strNP05 & "'" & _
                " AND SUBSTR(PA26,1,8)=CU01(+) AND decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)" & _
                " AND PA09=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT '',LC05||LC06||LC07,LC11||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,LC15,' ',0,LC08,CU10,LC22 as PA75" & _
                " From LawCase, nation, Customer" & _
                " WHERE LC01='" & strNP02 & "' AND LC02='" & strNP03 & "' AND LC03='" & strNP04 & "' AND LC04='" & strNP05 & "'" & _
                " AND SUBSTR(LC11,1,8)=CU01(+) AND decode(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+)" & _
                " AND LC15=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT '',HC06,HC05||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),' ',' ',' ',0,HC09,CU10,'' as PA75" & _
                " From HireCase, Customer" & _
                " WHERE HC01='" & strNP02 & "' AND HC02='" & strNP03 & "' AND HC03='" & strNP04 & "' AND HC04='" & strNP05 & "'" & _
                " AND SUBSTR(HC05,1,8)=CU01(+) AND decode(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+)"
   strSql = strSql & " Union " & _
                "SELECT SP11,SP05||SP06||SP07,SP08||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,SP09,' ',0,SP15,CU10,SP26 as PA75" & _
                " From Servicepractice, nation, Customer" & _
                " WHERE SP01='" & strNP02 & "' AND SP02='" & strNP03 & "' AND SP03='" & strNP04 & "' AND SP04='" & strNP05 & "'" & _
                " AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)" & _
                " AND SP09=NA01(+)"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Label1(7) = "" & Trim(Rs(0))
      lblCaseNm.Caption = "" & Trim(Rs(1))
      lblCU01Nm.Caption = "" & Trim(Rs(2))
      Label1(5) = "" & Trim(Rs(3))
      m_Nation = "" & Trim(Rs(4))
      m_PA08 = "" & Trim(Rs(5))
      m_PA10 = "" & Trim(Rs(6))
      m_IsClose = "" & Trim(Rs("TM29")) 'Add By Sindy 2015/1/21
      m_CP13 = ShowCurrCP13(strNP02, strNP03, strNP04, strNP05, m_Nation) 'Add By Sindy 2015/1/22
      m_CU10 = Mid("" & Trim(Rs("CU10")), 1, 3) 'Add By Sindy 2015/4/8 申請人1國籍
      m_FCfagent = "" & Trim(Rs("PA75")) 'Add By Sindy 2015/4/8 FC代理人
   End If
   If m_IsClose = "Y" Then
      MsgBox "此案已閉卷，不可操作此作業！", vbExclamation
      Frame2.Visible = False
      Frame3.Visible = False
      Frame4.Visible = False
      txt1(1).SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2014/12/31 結案單電子化
   'Mofieid by Morgan 2015/10/27 +P非臺灣案
   'If txt1(0) = "P" And m_Nation = "000" Then
   'Modified by Morgan 2015/12/3 +排除外專承辦組
   'Modify by Amy 2018/06/06 除法務顧問其他皆電子化
   'If txt1(0) = "P" And Pub_StrUserSt03 <> "F23" Then
   'Modify by Amy 2018/08/27 改送出鈕為預設,加系統別FC開頭也使用列印鈕
   'If Not (txt1(0) = "CFL" Or txt1(0) = "FCL" Or txt1(0) = "L" Or txt1(0) = "LIN" Or txt1(0) = "LA") Then
   'Modify by Amy 2018/09/03 +排除外專承辦組
   'Modify by Amy 2025/06/13 +ACS 先印紙本,待教威想程序人員及補看人員(電子化)再修改-秀玲
   'Memo by Amy  2025/06/13 FC 若上線後,應不會從Promoter 的這支程式操作,故程式先不用改-Sindy
   If Left(txt1(0), 2) = "FC" Or txt1(0) = "CFL" Or txt1(0) = "FCL" Or txt1(0) = "L" Or txt1(0) = "LIN" Or txt1(0) = "LA" Or (txt1(0) = "P" And Pub_StrUserSt03 = "F23") _
     Or txt1(0) = "ACS" Then
      Frame2.Visible = True
      Frame3.Visible = False
      Frame4.Visible = False: cmdFlowEmp.Visible = False
   End If
   'end 2018/08/27
   
   '2011/7/6 MODIFY BY SONIA 加NP10以判斷結案智權人員
   strSql = "SELECT ' ' AS V,decode(substr(cp09,1,1),'C',DECODE(cp05,'','',SUBSTR(cp05,1,4)-1911||'/'||SUBSTR(cp05,5,2)||'/'||SUBSTR(cp05,7,2)),'') as 來函收文日," & _
            "decode(substr(cp09,1,1),'C',DECODE('" & m_Nation & "','000',C2.cpm03,C2.cpm04),'') as 來函性質,decode(substr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||DECODE('" & m_Nation & "','000',C1.cpm03,C1.cpm04) as 下一程序," & _
            "DECODE(np08,'','',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2)) as 本所期限," & _
            "DECODE(np09,'','',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2)) as 法定期限," & _
            "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,np10,rownum as sort,np01,np22" & _
            " FROM NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
            " WHERE NP02='" & strNP02 & "' AND NP03='" & strNP03 & "' AND NP04='" & strNP04 & "' AND NP05='" & strNP05 & "'" & _
            " and np01=cp09(+) and np10=st01(+)" & _
            " and np02=C1.cpm01(+) and np07=C1.cpm02(+) and cp01=C2.cpm01(+) and cp10=C2.cpm02(+)" & _
            " and np06 is null " & strNpSqlOfNoSalesDuty
   'Modify By Sindy 2015/1/20
   If m_F0301 <> "" Then
      'Modify By Sindy 2017/8/11 ex:P-112100 length(NP24)=9:曾經收文過
      'strSql = strSql & " and NP24='" & m_F0301 & "'"
      strSql = strSql & " and (NP24='" & m_F0301 & "' or length(NP24)=9)"
   Else
      'Modify By Sindy 2017/8/11 ex:P-112100 length(NP24)=9:曾經收文過
      'strSql = strSql & " and NP24 is null"
      strSql = strSql & " and (NP24 is null or length(NP24)=9)"
   End If
   '2015/1/20 END
   strSql = strSql & " ORDER BY CP05 DESC, NP01 DESC, NP08 DESC"
   CheckOC3
   GRD1.Rows = 2
   GRD1.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         doQuery = True 'Add By Sindy 2014/6/19
         Set GRD1.Recordset = AdoRecordSet3.Clone
         intRow = GRD1.Rows - 1
         For i = 1 To GRD1.Rows - 1
            If strNP02 = "FCT" Or strNP02 = "T" Then
               If m_Nation < "010" And Trim(GRD1.TextMatrix(i, 10)) = "715" Then
                  intRow = intRow + 1
                  GRD1.AddItem ("")
                  GRD1.TextMatrix(intRow, 4) = "717 " & GetCaseTypeName(strNP02, "717", 0)
                  GRD1.TextMatrix(intRow, 10) = "717"
                  GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                  Call SetRowData(intRow, i)
               End If
            End If
         Next i
         GRD1.col = 12
         GRD1.Sort = 5 '字串昇冪
         SetDataListWidth
         GRD1.Visible = True
      Else
         If Trim(lblCaseNm) = "" Then
            MsgBox "無案件資料！", vbInformation
            'Add By Sindy 2025/5/9
            txt1(1).SetFocus
            Exit Function
            '2025/5/9 END
         Else
            'Add By Sindy 2015/1/20
            'Modified by Lydia 2017/07/25 不只一項引用
            'If UCase(TypeName(m_PrevForm)) = UCase("frm210147_1") Then
            If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
               doQuery = True
            Else
            '2015/1/20 END
               MsgBox "無期限資料！", vbInformation
            End If
            Frame2.Visible = False
            Frame3.Visible = True '無期限閉卷
         End If
      End If
   End With
   
   'Add By Sindy 2015/1/5 選取表單資料列
   If m_F0301 <> "" And m_F0303 <> "" And Val(m_F0304) > 0 Then
      GRD1.Visible = False
      For i = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(i, 13) = m_F0303 And _
            GRD1.TextMatrix(i, 14) = m_F0304 Then
            m_row = i
            GRD1.TextMatrix(m_row, 0) = "V"
            GRD1.row = m_row
            For j = 0 To GRD1.Cols - 1
               GRD1.col = j
               GRD1.CellBackColor = &HFFC0C0
            Next j
            Exit For
         End If
      Next i
      GRD1.Visible = True
   Else
      '只有一筆時,選取資料列
      If GRD1.Rows = 2 And (GRD1.TextMatrix(1, 13) <> "") Then
         GRD1.Visible = False
         m_row = 1
         GRD1.TextMatrix(m_row, 0) = "V"
         GRD1.row = m_row
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = &HFFC0C0
         Next j
         GRD1.Visible = True
      End If
   End If
   
   'Add by Amy 2020/05/20 +T/TF延展、續展、第二期註冊費,已有結案單彈訊息
   If (strNP02 = "T" Or strNP02 = "TF") And (Trim(GRD1.TextMatrix(m_row, colNP07)) = "102" Or Trim(GRD1.TextMatrix(m_row, colNP07)) = "716") _
    And ChkT102Inform(GRD1.TextMatrix(m_row, 13), GRD1.TextMatrix(m_row, 14)) = True Then
        MsgBox "此結案單已存在，不可重覆作業！", vbCritical, "操作錯誤！"
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        txt1(1).SetFocus
        Exit Function
   End If
   'end 2020/05/20
   
   'Add By Sindy 2015/2/9 檢查回覆單狀況
   If m_F0301 = "" Then
      If Frame4.Visible = True Then '電子結案單
         'Modify By Sindy 2015/5/18 電子結案單已用表單編號代替總收文號=CPP01,已不會重覆檔案,因此這裡不用檢查電子檔
'         If PUB_ChkIsReplyFile(strNP02, strNP03, strNP04, strNP05) = True Then
'            MsgBox "此案號回覆單已上傳，不可作業，若有疑問請通知電腦心中協助處理！", vbCritical
'            Exit Function
'         End If
      Else '紙本
         'Modify By Sindy 2015/5/18
         'If PUB_ChkIsReplyFile(strNP02, strNP03, strNP04, strNP05) = True Then
         If PUB_ChkIsReplyFile(strNP02 & strNP03 & strNP04 & strNP05) = True Then
         '2015/5/18 END
            MsgBox "此案號回覆單已上傳，可能已列印過！", vbInformation
            If Frame3.Visible = True Then '無期限閉卷
               cmdFile.Enabled = False
            End If
         End If
      End If
   End If
   '2015/2/9 END
   
   'Modify By Sindy 2015/2/4 原本在執行時才檢查,改在查詢後檢查權限
   strChkLimitEmp = m_CP13
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 11) <> "" Then '有智權人員
         strChkLimitEmp = GRD1.TextMatrix(i, 11)
         If GRD1.TextMatrix(i, 11) = strUserNum Then
            strChkLimitEmp = strUserNum
            Exit For 'Add by Sindy 2018/1/9 +
         End If
      End If
   Next i
   'Modify by Amy 2014/05/23 +if
   If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
        '開放專利處部份智權(A7)同仁資料給彥葶(A8)代為處理
        If InStr(Pub_GetSpecMan("A7"), strChkLimitEmp) > 0 Then
        Else
            MsgBox "您無權限將此客戶案件任意結案！", vbExclamation, "操作錯誤！"
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
            txt1(1).SetFocus
            Exit Function
        End If
   'end 2014/05/23
   '2011/7/6 ADD BY SONIA 檢查操作人員與結案期限智權人員 CFP-018696
   'Modify by Amy 2017/04/17 +MCTF同組可操作
   'modify by sonia 2019/3/15 GetMCTF0XCode改用ChkMCTF0XSales
   'ElseIf strChkLimitEmp = strUserNum Or Pub_StrUserSt03 = "M51" Or strChkLimitEmp < "6" Or strChkLimitEmp = GetMCTF0XCode(strUserNum) Then
   ElseIf strChkLimitEmp = strUserNum Or Pub_StrUserSt03 = "M51" Or strChkLimitEmp < "6" Or ChkMCTF0XSales(strChkLimitEmp, strUserNum) = True Then
   'Add by Amy 2024/12/24 智權 P2006的案子讓相關人員可結案 ex:CFT-021323 讓 a6034結案
   ElseIf strChkLimitEmp = "P2006" Then
      If (txt1(0) = "CFT" And InStr(Pub_GetSpecMan("P2006業績CFT案人員"), strUserNum) = 0) _
        Or (txt1(0) = "T" And InStr(Pub_GetSpecMan("P2006業績T案人員"), strUserNum) = 0) Then
         MsgBox "您無權限結" & GetPrjSalesNM(strChkLimitEmp) & "(P2006)之案件！"
         Frame2.Visible = False
         Frame3.Visible = False
         Frame4.Visible = False
         txt1(1).SetFocus
         Exit Function
      End If
   'Added by Morgan 2012/6/21
   '國外部,同部門(前兩碼同)都可印
   ElseIf Left(Pub_StrUserSt03, 1) = "F" Then
      If Left(PUB_GetStaffST15(strChkLimitEmp, "1"), 2) <> Left(PUB_GetStaffST15(strUserNum, "1"), 2) Then
         MsgBox "非相同部門客戶案件不可任意結案！"
         Frame2.Visible = False
         Frame3.Visible = False
         Frame4.Visible = False
         txt1(1).SetFocus
         Exit Function
      End If
   'end 2012/6/21
   'add by sonia 2014/10/30 美珍可操作林總案件
   'Modify by Amy 2015/02/04 改為特殊設定(總經理業務工作代理人員)
   'ElseIf strUserNum = "77027" And strChkLimitEmp = "94007" Then
   ElseIf InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), strChkLimitEmp) > 0 Then
   'end 2014/10/30
   Else
      '若原智權人員離職則該區同仁都可結案
      '2011/7/14 modify by sonia 台中林協理客戶開放可由該區其他人結案
'      If GetStaffName(strChkLimitEmp) = "" Or strChkLimitEmp = "68096" Then
'         If PUB_GetStaffST15(strChkLimitEmp, "1") = PUB_GetStaffST15(strUserNum, "1") Then
'         Else
            'MsgBox "非相同業務區客戶案件不可任意結案！"
            'Exit Function
            'modify by sonia 2013/7/2 帶人主管也可以結案
            stIdList = PUB_GetSalesList(strUserNum, PUB_GetStaffST15(strUserNum, "1"), PUB_GetStaffST15(strUserNum, "1"), PUB_GetST06(strUserNum))
            'Add By Sindy 2016/10/14 雅娟休假,玲玲要代操作 ex:P-105999
            'Modify by Amy 2021/02/23 Pub_SetSAManageEmpCombo 傳入表單名
            If CheckIsPersonRest(strChkLimitEmp, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then '當事人是否請假
               Call Pub_SetSAManageEmpCombo(strUserNum, , , stIdLeader, , , Me.Name) 'Add By Sindy 2015/2/5 區主管也可以結案 stIdLeader:若為主管,其所帶的人員ID
            Else
            '2016/10/14 END
               Call Pub_SetSAManageEmpCombo(strChkLimitEmp, , , stIdLeader, , , Me.Name) 'Add By Sindy 2015/2/5 區主管也可以結案 stIdLeader:若為主管,其所帶的人員ID
            End If
            'end 2021/02/23
            'Modify By Sindy 2015/3/11 原智權人員離職則該區同仁都可結案
            '                          帶人主管也可以結案
            'Modified by Lydia 2017/07/24 有權限的人員也可結案 (InStr(stIdList, strChkLimitEmp) > 0 And strUserNum <> strChkLimitEmp)
            If ((GetStaffName(strChkLimitEmp) = "" Or strChkLimitEmp = "68096") And PUB_GetStaffST15(strChkLimitEmp, "1") = PUB_GetStaffST15(strUserNum, "1")) Or _
               (Pub_StrST52 = True And InStr(stIdList, strChkLimitEmp) > 0) Or _
               InStr(stIdLeader, strChkLimitEmp) > 0 Or _
               (InStr(stIdList, strChkLimitEmp) > 0 And strUserNum <> strChkLimitEmp) Then
               If MsgBox("非本人客戶案件，確定要結案嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Frame2.Visible = False
                  Frame3.Visible = False
                  Frame4.Visible = False
                  txt1(1).SetFocus
                  Exit Function
               End If
            Else
               'add by sonia 2024/12/4 法律所案件則客戶檔之智權人員(介紹人)也可以操作
               If InStr(strNP02, "L") > 0 Then
                  strSql = "Select CU13, CU12, ST04, A0908 From Lawcase, Customer, Staff, acc090 Where substr(LC11,1,8)=CU01 And substr(LC11,9,1)=CU02 And CU13=ST01 and st15=a0901 And LC01='" & strNP02 & "' And LC02='" & strNP03 & "' And LC03='" & strNP04 & "' And LC04='" & strNP05 & "' "
                  strSql = strSql & " union Select CU13, CU12, ST04, A0908 From Hirecase, Customer, Staff, acc090 Where substr(HC05,1,8)=CU01 And substr(HC05,9,1)=CU02 And CU13=ST01 and st15=a0901 And HC01='" & strNP02 & "' And HC02='" & strNP03 & "' And HC03='" & strNP04 & "' And HC04='" & strNP05 & "' "
                  rsA.CursorLocation = adUseClient
                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     If "" & rsA("CU13").Value = strUserNum Then
                        strChkLimitEmp = strUserNum
                     ElseIf "" & rsA("A0908").Value = strUserNum Then
                        strChkLimitEmp = strUserNum
                     End If
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               'Modify By Sindy 2025/6/3
'                  If strChkLimitEmp = strUserNum Then Exit Function
               End If
               'end 2024/12/4
               If strChkLimitEmp <> strUserNum Then
               '2025/6/3 END
                  'Modify by Amy 2021/06/17 +(離職人員)
                  MsgBox "非本人、非相同業務區(離職人員)客戶案件不可任意結案！"
                  Frame2.Visible = False
                  Frame3.Visible = False
                  Frame4.Visible = False
                  txt1(1).SetFocus
                  Exit Function
               End If
            End If
            '2013/7/2 end
'         End If
'      Else
'         MsgBox "非本人案件不可任意結案！"
'         Frame2.Visible = False
'         Frame3.Visible = False
'         Frame4.Visible = False
'         txt1(1).SetFocus
'         Exit Function
'      End If
   End If
   '2011/7/6 END
    cmdFile.Enabled = True 'Add by Amy 2018/09/03 智權人員未先輸本案查詢,直接按回覆單匯入
    
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetRowData(intRow As Integer, i As Integer)
   GRD1.TextMatrix(intRow, 5) = Trim(GRD1.TextMatrix(i, 5))
   GRD1.TextMatrix(intRow, 6) = Trim(GRD1.TextMatrix(i, 6))
   GRD1.TextMatrix(intRow, 7) = Trim(GRD1.TextMatrix(i, 7))
   GRD1.TextMatrix(intRow, 14) = Trim(GRD1.TextMatrix(i, 14))
End Sub

'Add By Sindy 2015/1/5
Public Function TxtValidate(Index As Integer) As Boolean
Dim strSpecial As String
Dim bolChk As Boolean
Dim stIdList As String         'add by sonia 2013/7/2
Dim Rs As New ADODB.Recordset
Dim Cancel As Boolean
   
   TxtValidate = False
   
   Select Case Index
      Case 0
         If m_row <> 0 Then
            If GRD1.TextMatrix(m_row, 14) = "" Then
                MsgBox "請先查詢要執行的資料！", vbCritical, "操作錯誤！"
                txt1_GotFocus 0
                Exit Function
            Else
               m_iDiscount = 0: strSpecial = 0
               m_CurCP(1) = strNP02: m_CurCP(2) = strNP03: m_CurCP(3) = strNP04: m_CurCP(4) = strNP05
               '辦理減免退費提醒
               If PUB_GetCaseDiscStat(strNP02 & strNP03 & strNP04 & strNP05) = "Y" Then
                  Call PUB_CheckYearFeeReturn(m_CurCP, False, m_iDiscount, m_iYear1, m_iYear2)
               End If
               If m_iDiscount > 0 Then strSpecial = "1"
               
'               'Modify by Amy 2014/05/23 +if
'               If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
'                    '開放專利處部份智權(A7)同仁資料給彥葶(A8)代為處理
'                    If InStr(Pub_GetSpecMan("A7"), grd1.TextMatrix(m_row, 11)) > 0 Then
'                    Else
'                        MsgBox "您無權限將此客戶案件任意結案！", vbExclamation, "操作錯誤！"
'                        Exit Function
'                    End If
'               'end 2014/05/23
'               '2011/7/6 ADD BY SONIA 檢查操作人員與結案期限智權人員 CFP-018696
'               ElseIf grd1.TextMatrix(m_row, 11) = strUserNum Or Pub_StrUserSt03 = "M51" Or grd1.TextMatrix(m_row, 11) < "6" Then
'               'Added by Morgan 2012/6/21
'               '國外部,同部門(前兩碼同)都可印
'               ElseIf Left(Pub_StrUserSt03, 1) = "F" Then
'                  If Left(PUB_GetStaffST15(grd1.TextMatrix(m_row, 11), "1"), 2) <> Left(PUB_GetStaffST15(strUserNum, "1"), 2) Then
'                     MsgBox "非相同部門客戶案件不可任意結案！"
'                     Exit Function
'                  End If
'               'end 2012/6/21
'               'add by sonia 2014/10/30 美珍可操作林總案件
'               'Modify by Amy 2015/02/04 改為特殊設定(總經理業務工作代理人員)
'               'ElseIf strUserNum = "77027" And GRD1.TextMatrix(m_row, 11) = "94007" Then
'               ElseIf InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), grd1.TextMatrix(m_row, 11)) > 0 Then
'               'end 2014/10/30
'               Else
'                  '若原智權人員離職則該區同仁都可結案
'                  '2011/7/14 modify by sonia 台中林協理客戶開放可由該區其他人結案
'                  If GetStaffName(grd1.TextMatrix(m_row, 11)) = "" Or grd1.TextMatrix(m_row, 11) = "68096" Then
'                     If PUB_GetStaffST15(grd1.TextMatrix(m_row, 11), "1") = PUB_GetStaffST15(strUserNum, "1") Then
'                     Else
'                        'modify by sonia 2013/7/2 帶人主管也可以結案
'                        'MsgBox "非相同業務區客戶案件不可任意結案！"
'                        'Exit Sub
'                        stIdList = PUB_GetSalesList(strUserNum, PUB_GetStaffST15(strUserNum, "1"), PUB_GetStaffST15(strUserNum, "1"), PUB_GetST06(strUserNum))
'                        If Pub_StrST52 = True And InStr(stIdList, grd1.TextMatrix(m_row, 11)) > 0 Then
'                        Else
'                           MsgBox "非相同業務區客戶案件不可任意結案！"
'                           Exit Function
'                        End If
'                        '2013/7/2 end
'                     End If
'                  Else
'                     MsgBox "非本人案件不可任意結案！"
'                     Exit Function
'                  End If
'               End If
'               '2011/7/6 END
               
'2012/1/16 cancel by sonia 中所新人要印給主管簽
               'Mark by Amy 2020/05/18 內商延展、續展、第二期註冊費 改此執行
'               '2012/1/2 add by sonia 楊特助說即然要做結案通知,就不要再允許執行
'               'Modify By Sindy 2016/8/1
'               If strNP02 = "T" Or strNP02 = "TF" Then
'                  'If grd1.TextMatrix(m_row, colnp07) = "102" Or grd1.TextMatrix(m_row, colnp07) = "716" Then
'                  If GRD1.TextMatrix(m_row, colnp07) = "102" Then
'                     'MsgBox "內商延展或第二期註冊費案件請改至 業務期限資料查詢 做結案通知！"
'                     'Modify by Amy 2018/06/06 拿掉 以簡省紙張之列印 文字
'                     'Modified by Lydia 2019/07/03 更名
'                     'MsgBox "內商延展期限結案，請改至智權部作業之智權人員期限資料查詢，" & vbCrLf & vbCrLf & "按「內商延展結案通知」操作，謝謝合作！"
'                     'Modified by Lydia 2019/08/22 +路徑 (之=> ->程序作業->)
'                     MsgBox "內商延展期限結案，請改至智權部作業->程序作業->期限資料查詢，" & vbCrLf & vbCrLf & "按「內商延展結案通知」操作，謝謝合作！"
'                     Exit Function
'                  End If
'               End If
'               '2016/8/1 END
'               '2012/1/2 end
               'end 2020/05/18
            End If
         Else
            If GRD1.Rows = 2 And GRD1.TextMatrix(1, 13) = "" Then
                MsgBox "請先查詢要執行的資料！", vbCritical, "操作錯誤！"
                txt1_GotFocus 0
                Exit Function
            Else
                MsgBox "請先選擇一筆要執行的資料！", vbCritical, "操作錯誤！"
                Exit Function
            End If
         End If
         
      Case 1
         If txt1(0) = "" Or txt1(1) = "" Then
            MsgBox "請輸入本所案號！"
            If txt1(1) = "" Then txt1(1).SetFocus
            If txt1(0) = "" Then txt1(0).SetFocus
            Exit Function
         End If
         If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then
            MsgBox "請先查詢要執行的資料！", vbCritical, "操作錯誤！"
            txt1_GotFocus 0
            Exit Function
         End If
   End Select
   
'   If Option1(5) = True And Trim(Text1) = "" Then
'      MsgBox "結案理由點選其他時，請輸入說明！"
'      Text1.SetFocus
'      Exit Sub
'   End If
   bolChk = False
   For i = 0 To 11
      If Option1(i) = True Then
         bolChk = True
         Exit For
      End If
   Next i
   If bolChk = False Then
      MsgBox "請勾選結案理由！"
      Exit Function
   End If
   
   'Add By Sindy 2015/11/27
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   Cancel = False
   Call txtPCnt_Validate(Index, False)
   If Cancel = True Then
      Exit Function
   End If
   '2015/11/27 END
   
   'Added by Morgan 2016/6/7
   '從上面移下來，要放存檔前否則發生操作錯誤會導致備註重複 Ex.P-101521
   If Index = 0 And m_row <> 0 Then
      'Added by Morgan 2012/9/28
      '一案兩請的新型年費填寫結案單，若發明案尚未審定或核駁且未閉卷時，要提醒並由使用者確認後將提醒文字印在結案單之備註欄
      If strNP02 = "P" And (m_Nation = "000" Or (m_Nation = "020" And Val(m_PA10) >= 20091001)) And m_PA08 = "2" And GRD1.TextMatrix(m_row, colNP07) = "605" Then
         strExc(0) = "select pa01,pa02,pa03,pa04 from casemap,patent where cm10='3' and cm01='" & strNP02 & "' and cm02='" & strNP03 & "' and cm03='" & strNP04 & "' and cm04='" & strNP05 & "'" & _
            " and pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa08='1' and pa57 is null and (pa16 is null or pa16='2')" & _
            " union select pa01,pa02,pa03,pa04 from casemap,patent where cm10='3' and cm05='" & strNP02 & "' and cm06='" & strNP03 & "' and cm07='" & strNP04 & "' and cm08='" & strNP05 & "'" & _
            " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa08='1' and pa57 is null and (pa16 is null or pa16='2')"
         
         intI = 1
         Set Rs = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If m_Nation = "020" Then
               strExc(2) = "※此案為大陸之一案兩請，實用新型放棄繳年費，則大陸發明案就不予專利。"
            Else
               strExc(2) = "※此案為一案兩請，發明專利審定前，新型專利權若因未繳年費而當然消滅者，則將不予專利。"
            End If
            If MsgBox(strExc(2) & vbCrLf & "是否確認要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation, "一案兩請新型結案提醒") = vbYes Then
               Text1 = strExc(2) & Text1
            Else
               Exit Function
            End If
         End If
      End If
      'end 2012/9/28
   End If
   'end 2016/6/7
   
   'Added by Lydia 2021/01/26 CFP,CFT英國脫歐案，彈提醒：歐盟案件，除非確定延展（英國）或延展費（英國）不辦，否則不可單獨將委任代理人結案！
   strExc(1) = ""
   If m_Nation = "239" Then
        If strNP02 = "CFP" And GRD1.TextMatrix(m_row, colNP07) = "613" Then
            strExc(1) = " and np07='444' "
        ElseIf strNP02 = "CFP" And GRD1.TextMatrix(m_row, colNP07) = "444" Then
            strExc(1) = " and np07='613' "
        ElseIf strNP02 = "CFT" And GRD1.TextMatrix(m_row, colNP07) = "110" Then
            strExc(1) = " and np07='710' "
        ElseIf strNP02 = "CFT" And GRD1.TextMatrix(m_row, colNP07) = "710" Then
            strExc(1) = " and np07='110' "
        End If
        If strExc(1) <> "" Then
            strExc(1) = "select np01,np07,cpm03 from nextprogress,casepropertymap where np02=cpm01(+) and np07=cpm02(+) and np02='" & strNP02 & "' and np03='" & strNP03 & "' and np04='" & strNP04 & "' and np05='" & strNP05 & "'" & _
                             " and np06 is null and np24 is null " & strExc(1)
            intI = 1
            Set Rs = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
                 If GRD1.TextMatrix(m_row, colNP07) = "444" Or GRD1.TextMatrix(m_row, colNP07) = "710" Then
                     strExc(2) = "歐盟案件，除非確定" & Rs.Fields("cpm03") & "不辦，否則不可單獨將" & Mid(GRD1.TextMatrix(m_row, 4), 5) & "結案！"
                 Else
                     strExc(2) = "英國脫歐案之" & Mid(GRD1.TextMatrix(m_row, 4), 5) & "結案，請一併將" & Rs.Fields("cpm03") & "結案！"
                 End If
                 MsgBox strExc(2), vbInformation + vbOKOnly, "英國脫歐案"
            End If
        End If
   End If
   'end 2021/01/26
   
   TxtValidate = True
End Function

'Add By Sindy 2015/1/5 送出
Private Sub cmdSend_Click()
Dim strUpdDate As String, strUpdTime As String
Dim bolModify  As Boolean
Dim SignPerson As String
Dim strTemp As Variant, ii As Integer
Dim fs, f
Dim stFileName As String, stReName As String
Dim strSubject As String, strContent As String
Dim Rs As New ADODB.Recordset
Dim Cancel As Boolean
'Add by Amy 2018/06/06
Dim strF0202_2 As String, strF0202_3 As String '程序人員/補看人員
Dim strMsg As String
Dim opt As OptionButton, stTmp As String 'Add by Amy 2020/05/18

On Error GoTo ErrHand

   'Add by Amy 2020/05/20 直接由結案單進入,做內商延展、續展、第二期註冊費 之結案
   bolNoFlow = False
   If strNP02 = "T" Or strNP02 = "TF" Then 'NP01
        If GRD1.TextMatrix(m_row, colNP07) = "102" Or GRD1.TextMatrix(m_row, colNP07) = "716" Then
            bolNoFlow = True
            m_NP01 = GRD1.TextMatrix(m_row, 13)
            m_NP22 = GRD1.TextMatrix(m_row, 14)
        End If
   End If
   'end 2020/05/20
   
   'Modify by Amy 2020/05/18 +if 非內商延展、續展、第二期註冊費,才run簽核流程
   If bolNoFlow = False Then
        '無期限
        'Modify By Sindy 2018/9/5
        'If m_row = 0 Then
        If Frame3.Visible = True Then '無期限閉卷
        '2018/9/5 END
           m_F0303 = strNP02 & strNP03 & strNP04 & strNP05
           m_F0304 = ""
        '有期限
        Else
           'Add By Sindy 2018/9/5 未選取按送出會當掉
           If m_row = 0 Then
              MsgBox "請勾選欲解除期限的資料！"
              Exit Sub
           End If
           '2018/09/05 END
           If GRD1.TextMatrix(m_row, 13) <> "" Then 'NP01
              m_F0303 = GRD1.TextMatrix(m_row, 13)
              m_F0304 = GRD1.TextMatrix(m_row, 14) 'NP22
              m_F0316 = GRD1.TextMatrix(m_row, 11) '智權人員
           Else
              '會有此狀況嗎?
              m_F0303 = strNP02 & strNP03 & strNP04 & strNP05
              m_F0304 = ""
           End If
        End If
        'Modify By Sindy 2023/5/19
        m_F0316 = GetF0316(m_F0316)
'        'If m_F0316 = "" Then m_F0316 = m_CP13 '智權人員
'        If m_F0316 = "" Then m_F0316 = strUserNum 'Add By Sindy 2016/9/23 P-092509 吳中一,操作人員是王俊剴
'        '若智權人員已離職,則以Login人員代替
'        'If ChkStaffST04(m_F0316, False) = True Then
'        If ChkStaffST04(m_F0316, False) = True Or Left(m_F0316, 1) <= "6" Then
'           m_F0316 = strUserNum
'        End If
'        'Add By Sindy 2021/11/10 判斷是否為MCT
'        If ChkMCTF0XSales(PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05), strUserNum) = True Then
'         m_F0316 = PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05)
'        End If
'        '2021/11/10 END
        '2023/5/19 END
   End If
   
   'Modify by Amy 2022/03/08 程式優化-杜燕文
   'Add by Amy 2018/09/05 未匯入回覆單則一定要輸入結案理由之說明
'   If m_strSaveFiles = "" And Trim(Text1) = MsgText(601) Then
'     MsgBox "請匯入回覆單！若無法取得請在結案理由記載與客戶聯繫的過程與溝通的結論。"
'     Text1.SetFocus
'     Exit Sub
'   End If
'   'Add by Amy 2018/09/03 勾選其他,理由一定要輸-文雄
'   If Option1(5).Value = True And Trim(Text1) = MsgText(601) Then
'      MsgBox "勾選其他，結案說明不可空白！"
'      Text1.SetFocus
'      Exit Sub
'   End If
'   'end 2018/09/03
   ' 若無回覆單,理由不可空白-------->                                        2.已轉由他所處理        /   4.客戶另案重提         /    10.客戶自行處理              /      12.其他
   strMsg = ""
   If m_strSaveFiles = "" And Trim(Text1) = MsgText(601) And (Option1(1).Value = True Or Option1(3).Value = True Or Option1(10).Value = True Or Option1(5).Value = True) Then
      If Option1(1).Value = True Then
            strMsg = Option1(1).Caption
      ElseIf Option1(3).Value = True Then
            strMsg = Option1(3).Caption
      ElseIf Option1(10).Value = True Then
            strMsg = Option1(10).Caption
      Else
            strMsg = Option1(5).Caption
      End If
      
      MsgBox "點選" & strMsg & "，結案說明不可空白！"
      Text1.SetFocus
      Exit Sub
   End If
   'end 2022/03/08
   
   'Modify by Amy 2020/05/18 +if 非內商延展、續展、第二期註冊費,才run簽核流程
   If bolNoFlow = False Then
        If GetFLOW001Person(m_F0316, Flow_結案單, True) = "" Then
           MsgBox "無設定簽核人員，不可使用電子表單流程！"
           Exit Sub
        End If
        
        'Add by Amy 2018/06/06  從下面搬上來先檢查
        strF0202_2 = GetSignOffEmp("NP", strNP02, strNP03, m_Nation, strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05) 'Modify by Amy 2020/03/06 +本所案號
        If strF0202_2 = MsgText(601) Then
           MsgBox "無設定程序人員，請通知電腦中心！"
           Exit Sub
         End If
         'Modify by Amy 2021/06/29 +本所案號 for CFT/CFC/S 判斷補看人員
         strF0202_3 = GetF0202_3(strNP02, strNP03, strNP04, strNP05)
         If strF0202_3 = MsgText(601) Then
           MsgBox "無設定補看人員，請通知電腦中心！"
           Exit Sub
         End If
         'end 2018/06/06
    End If
    'end 2020/05/18
   
   '檢查條件
   If Frame3.Visible = True Then '無期限閉卷
      '檢查是否有資料重覆
      If ChkFlowFormExists(Flow_結案單, "", "", strNP02, strNP03, strNP04, strNP05, , , m_F0301) = True Then
         MsgBox "此結案單已存在，不可重覆作業！", vbCritical, "操作錯誤！"
         txt1(0).SetFocus
         Exit Sub
      End If
      If TxtValidate(1) = False Then Exit Sub
   Else
      '檢查是否有資料重覆
      If ChkFlowFormExists(Flow_結案單, GRD1.TextMatrix(m_row, 13), GRD1.TextMatrix(m_row, 14), strNP02, strNP03, strNP04, strNP05, , , m_F0301) = True Then
         MsgBox "此結案單已存在，不可重覆作業！", vbCritical, "操作錯誤！"
         txt1(0).SetFocus
         Exit Sub
      End If
      If TxtValidate(0) = False Then Exit Sub
   End If
   
   'Add By Sindy 2015/11/27
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Sub
   End If
   '2015/11/27 END
   
   'Added by Lydia 2021/10/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   'end 2021/10/07
   
   If m_strSaveFiles = "" Then
      'Modify By Sindy 2019/8/12
      'If MsgBox("缺回覆單，確定是否繼續結案？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
      If MsgBox("回覆單未匯入，要匯入回覆單？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
      '2019/8/12 END
         Exit Sub
      End If
   'Add by Amy 2025/06/19 檔案開啟要關閉,避免後面刪檔會錯
   Else
      If Pub_FileIsOpen(Me.Name, m_strSaveFiles, strExc(9)) = True Then
         MsgBox "檔案正在使用中,需關閉之檔案如下:" & vbCrLf & _
                           Replace(strExc(9), ";", vbCrLf)
         Exit Sub
      End If
   'end 2025/06/19
   End If
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   'Modify by Amy 2020/05/18 +if 內商延展、續展、第二期註冊費,不需 run簽核流程
   If bolNoFlow = True Then
        If m_strSaveFiles <> MsgText(601) Then
            If PUB_UpdReplyFile(m_strSaveFiles, "", strNP02, strNP03, strNP04, strNP05, , m_NP22) = False Then
                MsgBox "檔案上傳有誤！"
                cnnConnection.RollbackTran
            End If
        End If
        strSql = ""
        For Each opt In Option1
            If opt.Value = True Then
                If opt.Index = 5 Then
                    strSql = strSql & ";99.其他：" & Text1
                Else
                    'Modify by Amy 2022/03/10 bug 項目已拿掉會error
                    'stTmp = Format(Mid(opt.Caption, 1, Val(InStr(opt.Caption, " ")) - 1), "00") & "."
                    'strSql = strSql & ";" & stTmp & Mid(opt.Caption, Val(InStr(opt.Caption, " ")) + 1)
                    'Modify by Amy 2022/04/26 項目對應錯誤
                    If opt.Index <= 4 Then
                        stTmp = Format(opt.Index + 1, "00") & "."
                    Else
                        stTmp = Format(opt.Index, "00") & "."
                    End If
                    'end 2022/04/26
                    strSql = strSql & ";" & stTmp & opt.Caption
                End If
            End If
        Next
        If Text1 <> MsgText(601) And Option1(5).Value = False Then
            strSql = strSql & ";" & Text1
        End If
        strSql = "Insert into t102inform (ti01,ti02,ti03,ti04,ti05) Values (to_number(to_char(sysdate, 'YYYYMMDD')),'" & m_NP01 & "','" & strUserNum & "'," & m_NP22 & "," & CNULL(ChgSQL(Mid(strSql, 2))) & ") "
        cnnConnection.Execute strSql
        cnnConnection.CommitTrans
        '已結案清除案號記錄
        If m_strSaveFiles <> MsgText(601) Then
            Call PUB_DelPCOrgFile(m_strSaveFiles) '刪除原檔
            m_strSaveFiles = ""
        End If
   '簽核流程
   Else
       strUpdDate = strSrvDate(1)
       strUpdTime = Right("000000" & ServerTime, 6)
       bolModify = False
       If m_F0301 <> "" Then bolModify = True
       
       If bolModify = False Then
          '表單編號自動給號
          m_F0301 = AutoNo_FLOW("CLS", 5)
          '檢查是否還有自動給號資料不一致的問題
          strSql = "select AU03 from autonumber where AU01='CLS'"
          intI = 1
          Set Rs = ClsLawReadRstMsg(intI, strSql)
          If intI = 1 Then
             If Val(Rs.Fields("AU03")) <> Val(Right(m_F0301, Len(m_F0301) - 3)) Then
                MsgBox "系統自動給號(" & m_F0301 & ")更新有誤，請洽電腦中心！", vbInformation, "系統錯誤"
                m_F0301 = ""
                GoTo ErrHand
                Exit Sub
             End If
          End If
       End If
       
       '新增表單主檔
       If Option1(5).Value = True Then
          m_F0305 = "99"
       Else
          For ii = 0 To 4
             If Option1(ii).Value = True Then
                m_F0305 = Format(ii + 1, "00")
                Exit For
             End If
          Next ii
          For ii = 6 To 11
             If Option1(ii).Value = True Then
                m_F0305 = Format(ii, "00")
                Exit For
             End If
          Next ii
       End If
       'Modify by Amy 2025/05/19 +if FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
       If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         If bolModify = False Then
            strSql = "insert into FLOW003(F0301,F0302,F0307,F0310,F0311,F0312,F0316)" & _
                     " values('" & m_F0301 & "','" & Flow_結案單 & "','" & strUserNum & "','" & strUserNum & "'," & strUpdDate & "," & strUpdTime & ",'" & m_CP13 & "')"
            cnnConnection.Execute strSql, intI
            '結案單主檔
            strSql = "Insert into CloseCaseMain(CCM01,CCM02,CCM03,CCM04,CCM05,CCM11,CCM12,CCM13) " & _
                        "Values ('" & m_F0301 & "','" & m_F0303 & "'," & CNULL(m_F0304) & ",'" & m_F0305 & "'," & CNULL(ChgSQL(Text1)) & "" & _
                        ",'" & strUserNum & "'," & strUpdDate & "," & strUpdTime & ")"
            cnnConnection.Execute strSql, intI
         Else
            '結案單主檔
            strSql = "Update CloseCaseMain Set " & _
                            "CCM02='" & m_F0303 & "'" & _
                           ",CCM03=" & CNULL(m_F0304) & "" & _
                           ",CCM04='" & m_F0305 & "'" & _
                           ",CCM05='" & Text1 & "'" & _
                           " Where CCM01='" & m_F0301 & "' "
            cnnConnection.Execute strSql, intI
         End If
       Else
         If bolModify = False Then
            'Modify By Sindy 2023/5/22 F0316=m_F0316 => F0316=m_CP13 因內專程序會幫雅娟、P1004代填, F0316存智權人員ID
            strSql = "insert into FLOW003(F0301,F0302,F0303,F0304,F0305,F0306,F0307,F0310,F0311,F0312,F0316)" & _
                     " values('" & m_F0301 & "','" & Flow_結案單 & "','" & m_F0303 & "'," & CNULL(m_F0304) & "," & _
                             "'" & m_F0305 & "','" & Text1 & "','" & strUserNum & "','" & strUserNum & "'," & strUpdDate & "," & strUpdTime & ",'" & m_CP13 & "')"
         Else
            strSql = "update FLOW003 set" & _
                     " F0303='" & m_F0303 & "'" & _
                     ",F0304=" & CNULL(m_F0304) & "" & _
                     ",F0305='" & m_F0305 & "'" & _
                     ",F0306='" & Text1 & "'" & _
                     " where F0301='" & m_F0301 & "'"
         End If
         cnnConnection.Execute strSql, intI
       End If
       'end 2025/05/19
       
       '新增表單簽核檔
       strSql = "delete From FLOW002 where F0201=" & CNULL(m_F0301)
       cnnConnection.Execute strSql
       SignPerson = GetFLOW001Person(m_F0316, Flow_結案單)
       '簽核人員
       If SignPerson <> "" Then
          strTemp = Split(SignPerson, ",")
          For ii = 0 To UBound(strTemp)
             If strTemp(ii) <> "" Then
                If ii = 0 And m_SetFlowEmp1 <> "" Then strTemp(ii) = m_SetFlowEmp1 '簽核人員1若有調整,已調整的為主
                strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(m_F0301) & ",'1'," & (ii + 1) & "," & CNULL(CStr(strTemp(ii))) & ")"
                cnnConnection.Execute strSql
             End If
          Next ii
       End If
       '*****若是代他人填單,簽核檔中若自己也是簽核人員之一時,一併確認掉
       strSql = "update FLOW002 set " & _
                "F0205='" & strUpdDate & "'" & _
                ",F0206='" & strUpdTime & "'" & _
                ",F0207='1'" & _
                " where F0201='" & m_F0301 & "' and F0204='" & strUserNum & "' and F0207 is null"
       cnnConnection.Execute strSql
       '*****END
       '程序人員
       'Modfy by Amy 2018/06/06 抓取程序人員搬至上面
        strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(m_F0301) & ",'2',1," & CNULL(Left(strF0202_2, 5)) & ")"
        cnnConnection.Execute strSql
       
       '補看人員
       'Modify by Amy 2018/06/06 抓取補看人員搬至上面
       strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(m_F0301) & ",'3',1," & CNULL(strF0202_3) & ")"
       cnnConnection.Execute strSql
       'end 2018/06/06
       
       '更新下一程序
       If Val(m_F0304) > 0 Then
          strSql = "Update NextProgress Set NP24='" & m_F0301 & "' WHERE NP01='" & m_F0303 & "' and NP22=" & m_F0304
          cnnConnection.Execute strSql
       End If
       
       '記錄重送訊息
       If bolModify = True Then
          strSql = GetInsertFLOW004Sql(Trim(m_F0301), strUserNum, strUpdDate, strUpdTime, Flow_重送, "")
          cnnConnection.Execute strSql
       End If
       
       '儲存回覆單：統一先存本所案號收進系統，等產生B類文號時再掛進回覆單
       If bolModify = True Then
          'Modify By Sindy 2015/5/18
          'If PUB_ChkIsReplyFile(strNP02, strNP03, strNP04, strNP05) = True Then
          If PUB_ChkIsReplyFile(strNP02 & strNP03 & strNP04 & strNP05, , , , m_F0301) = True Then
          '2015/5/18 END
             'Modify By Sindy 2015/5/18
             'PUB_DelFtpFile2 strNP02 & strNP03 & strNP04 & strNP05, " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"  'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
             'strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & strNP02 & strNP03 & strNP04 & strNP05 & "' and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"
             PUB_DelFtpFile2 strNP02 & strNP03 & strNP04 & strNP05, " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"   'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
             strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & strNP02 & strNP03 & strNP04 & strNP05 & "' and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"
             '2015/5/18 END
             cnnConnection.Execute strSql
          End If
       End If
       'Modify By Sindy 2015/5/18
       'If PUB_UpdReplyFile(m_strSaveFiles, "", strNP02, strNP03, strNP04, strNP05) = False Then Exit Sub
       If PUB_UpdReplyFile(m_strSaveFiles, "", strNP02, strNP03, strNP04, strNP05, , m_F0301) = False Then Exit Sub
       '2015/5/18 END
       
       '讀取下一處理人員
       If GetNextProPerson_Flow(m_F0301, m_F0316, m_F0308, m_F0309) = False Then GoTo ErrHand
       
       cnnConnection.CommitTrans
       
       'Add By Sindy 2015/1/27 刪除匯入來源的回覆單
       Call PUB_DelPCOrgFile(m_strSaveFiles): m_strSaveFiles = ""
       '2015/1/27 END
       
       '發E-Mail通知下一處理主管(多審核主管用)
       If m_F0309 = Flow_主管審核中 Then
          If bolModify = True Then
             strContent = GetEMailContent_Flow(m_F0301, strSubject, Flow_重送)
          Else
             strContent = GetEMailContent_Flow(m_F0301, strSubject)
          End If
    '      MsgBox "收件者：" & m_F0308 & GetPrjSalesNM(m_F0308) & vbCrLf & vbCrLf & _
    '             "主　旨：" & strSubject & vbCrLf & vbCrLf & _
    '             "內　容：" & strContent, vbInformation
          'Modify By Sindy 2016/10/12 + 含特殊職代
          PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True
       'Add by Amy 2018/06/06 智權主管填自己的結案單,要發信通知承辦人員
       ElseIf m_F0309 = Flow_處理中 And (txt1(0) = "CFT" Or txt1(0) = "CFC" Or txt1(0) = "S") Then
         'Modify By Sindy 2025/6/4
         strContent = GetEMailContent_Flow(m_F0301, strSubject)
         PUB_SendMail strUserNum, Left(strF0202_2, 5), "", strSubject, strContent, , , , , , , , , , , , , True
'         strSubject = txt1(0) & "-" & txt1(1) & txt1(2) & txt1(3) & "結案單通知"
'         'Modify by Amy 2018/08/27 路徑改用變數
'         PUB_SendMail strUserNum, Left(strF0202_2, 5), "", strSubject, 結案單外商CF操作路徑, , , , , , , , , , , , , True
         '2025/6/4 END
       End If
   End If
   'end 2020/05/18
       
    Screen.MousePointer = vbDefault
       
    If bolModify = True Then
        If TypeName(m_PrevForm) <> "Nothing" Then
            If UCase(TypeName(m_PrevForm)) = UCase("frm210147_1") Then
                m_PrevForm.cmdExit_Click
                Set m_PrevForm = Nothing
            End If
        End If
        Unload Me
    Else
        strMsg = txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3)
        '清除欄位值
        m_F0301 = ""
        txt1(0) = "": txt1(1) = "": txt1(2) = "": txt1(3) = ""
        Call ClearData
        cmdOK(2).Default = True
        'Add by Amy 2020/05/18 1.  由期限彈跳選多筆來結案時,彈訊息後跳下一筆
        If TypeName(m_PrevForm) <> "Nothing" Then
            If UCase(TypeName(m_PrevForm)) = UCase("frm100123") Then
                MsgBox strMsg & " 案號已送出結案單！", vbInformation
                Call cmdok_Click(3)
            End If
        '由案件結案單進入，按送出後Focus要停在系統類別欄
        Else
            txt1(0).SetFocus
        End If
    End If
   Exit Sub
   
ErrHand:
   If bolModify = False Then m_F0301 = ""
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 送出失敗！" & strMsg & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2015/1/23
   'Add by Amy 2021/06/17 +提示文字
   Label2.Caption = "1.本人、帶人主管、區主管才能操作！" & vbCrLf & _
                                "2.離職人員開放同區人員可操作！"
   'Modify by Amy 2018/08/27 原預設列印,改送出鈕為預設-文雄
   Frame2.Visible = False '列印
   Frame3.Visible = False
   Frame4.Visible = True: cmdFlowEmp.Visible = True 'Add By Sindy 2014/12/31 結案單電子化
   'end 2018/08/27
   cmdFile.Enabled = False 'Add by Amy 2018/09/03 智權人員未先輸本案查詢,直接按回覆單匯入
   'Modify  by Amy 2014/05/23 開放專利處部份智權同仁資料給彥葶代為處理
    If CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = "A8"
    End If
    'end 2014/05/23
   Call settxtPCnt
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210133 = Nothing
   
   'Add By Sindy 2014/6/19
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
      'Modify by Amy 2025/08/20+if 不是由外商系統收件區進入
      If UCase(TypeName(m_PrevForm)) <> UCase("frm06010612") Then
         'Modify by Amy 2025/08/13 +frm210133_2
         If UCase(TypeName(m_PrevForm)) <> UCase("frm210147_1") And UCase(TypeName(m_PrevForm)) <> UCase("frm210133_2") Then
            'm_PrevForm.Show
            'If UCase(TypeName(m_PrevForm)) = UCase("frm100123") Then
               m_PrevForm.PubShowNextData
            'End If
         End If
      End If
      Set m_PrevForm = Nothing
   End If
   '2014/6/19 END
   bolNoFlow = False 'Add by Amy 2020/05/18
End Sub

Private Sub settxtPCnt()
   Dim strST06 As String
   strST06 = PUB_GetST06(strUserNum)
   If strST06 = "1" Then
      txtPCnt(0) = "1"
      txtPCnt(1) = "1"
   Else
      txtPCnt(0) = "2"
      txtPCnt(1) = "2"
   End If
End Sub

Private Sub grd1_SelChange()
Dim m_mouseRow As Integer

   GRD1.Visible = False
   m_mouseRow = GRD1.MouseRow
   GRD1.col = 0
   If m_mouseRow <> 0 Then
      'Add By Sindy 2018/9/5
      If GRD1.TextMatrix(m_mouseRow, 4) <> "" And GRD1.TextMatrix(m_mouseRow, 5) <> "" Then
      '2018/9/5 END
       If m_row <> 0 Then
           GRD1.row = m_row
            For i = 0 To GRD1.Cols - 1
                 GRD1.col = i
                 If GRD1.CellBackColor = &HFFC0C0 Then
                   GRD1.CellBackColor = &H80000018
                   GRD1.TextMatrix(m_row, 0) = ""
                   m_row = 0 'Add By Sindy 2018/9/5
                 Else
                   GRD1.CellBackColor = &HFFC0C0 '&H80000018 '&H8080FF
                   GRD1.TextMatrix(m_row, 0) = "V"
                 End If
           Next i
       End If
       If m_row <> m_mouseRow Then
           GRD1.row = m_mouseRow
           m_row = m_mouseRow
            For i = 0 To GRD1.Cols - 1
                 GRD1.col = i
                 If GRD1.CellBackColor = &HFFC0C0 Then
                   GRD1.CellBackColor = &H80000018
                   GRD1.TextMatrix(m_row, 0) = ""
                   m_row = 0
                 Else
                   GRD1.CellBackColor = &HFFC0C0
                   GRD1.TextMatrix(m_row, 0) = "V"
                 End If
           Next i
       'Modify By Sindy 2018/9/5 Mark
'       Else
'           m_row = 0
       End If
      End If
   End If
   GRD1.Visible = True
End Sub

'Add By Sindy 2015/11/27
Private Sub Text1_GotFocus()
   OpenIme
   Text1.SetFocus
End Sub
'備註
Private Sub Text1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(Text1, Text1.MaxLength) = False Then
      Cancel = True
      Text1_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub
'2015/11/27 END

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub SetDataListWidth()
GRD1.Visible = False
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, m_i As Integer

   arrGridHeadText = Array("V", "來函收文日", "來函性質", "來函總收文號", "下一程序" _
             , "本所期限", "法定期限", "智權人員", "相關人", "備註" _
             , "NP07", "NP10", "Sort", "np01", "np22")
   arrGridHeadWidth = Array(200, 1000, 1000, 1000, 1500 _
                      , 800, 800, 800, 1000, 3000 _
                      , 0, 0, 0, 0, 0)
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      If iRow > 10 Then
         GRD1.ColWidth(iRow) = 0
      Else
         GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      End If
      GRD1.CellAlignment = flexAlignLeftCenter
   Next
   'Added by Lydia 2021/01/26 取得欄位位置
   If colNP07 = 0 Then
       colNP07 = PUB_MGridGetId("NP07", GRD1)
   End If
   'end 2021/01/26
   
   GRD1.Visible = True
End Sub

Private Sub txtPCnt_GotFocus(Index As Integer)
   TextInverse Me.txtPCnt(Index)
End Sub

Private Sub txtPCnt_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> vbKeyBack And KeyAscii <> Asc(1) And KeyAscii <> Asc(2) And KeyAscii <> Asc(3) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtPCnt_Validate(Index As Integer, Cancel As Boolean)
   If txtPCnt(Index) = "" Then
      MsgBox "請輸入列印份數！", vbCritical
      Call txtPCnt(Index).SetFocus
      Cancel = True
   End If
End Sub

