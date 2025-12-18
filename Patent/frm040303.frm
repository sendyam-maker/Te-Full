VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040303 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳年費/實體審查通知函"
   ClientHeight    =   5712
   ClientLeft      =   8148
   ClientTop       =   1476
   ClientWidth     =   5148
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   5148
   Begin VB.TextBox TxtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   3210
      MaxLength       =   7
      TabIndex        =   38
      Top             =   3195
      Width           =   975
   End
   Begin VB.TextBox TxtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   1980
      MaxLength       =   7
      TabIndex        =   37
      Top             =   3195
      Width           =   975
   End
   Begin VB.TextBox TxtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   3210
      MaxLength       =   7
      TabIndex        =   36
      Top             =   2865
      Width           =   975
   End
   Begin VB.TextBox TxtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1980
      MaxLength       =   7
      TabIndex        =   35
      Top             =   2865
      Width           =   975
   End
   Begin VB.TextBox TxtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   3210
      MaxLength       =   7
      TabIndex        =   34
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox TxtDate 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1980
      MaxLength       =   7
      TabIndex        =   30
      Top             =   2220
      Width           =   975
   End
   Begin VB.CheckBox chkKind 
      Height          =   225
      Index           =   3
      Left            =   1935
      TabIndex        =   29
      Top             =   1058
      Width           =   285
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   1485
      TabIndex        =   22
      Top             =   1560
      Width           =   3570
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   630
         MaxLength       =   5
         TabIndex        =   26
         Top             =   30
         Width           =   795
      End
      Begin VB.OptionButton Option3 
         Caption         =   "20號"
         Height          =   180
         Index           =   1
         Left            =   2835
         TabIndex        =   25
         Top             =   60
         Width           =   720
      End
      Begin VB.OptionButton Option3 
         Caption         =   "10號"
         Height          =   180
         Index           =   0
         Left            =   2070
         TabIndex        =   24
         Top             =   60
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年月："
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   27
         Top             =   60
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日："
         Height          =   180
         Index           =   10
         Left            =   1665
         TabIndex        =   23
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CheckBox chkKind 
      Enabled         =   0   'False
      Height          =   225
      Index           =   0
      Left            =   1935
      TabIndex        =   19
      Top             =   780
      Width           =   285
   End
   Begin VB.CheckBox chkKind 
      Height          =   225
      Index           =   2
      Left            =   3105
      TabIndex        =   18
      Top             =   480
      Value           =   1  '核取
      Width           =   285
   End
   Begin VB.CheckBox Check2 
      Caption         =   "只印期限表"
      Height          =   315
      Left            =   270
      TabIndex        =   9
      Top             =   4470
      Width           =   2040
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只印地址條"
      Height          =   315
      Left            =   300
      TabIndex        =   8
      Top             =   5370
      Width           =   1905
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1950
      TabIndex        =   7
      Top             =   4965
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1950
      MaxLength       =   8
      TabIndex        =   6
      Top             =   4065
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   3336
      TabIndex        =   10
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   4164
      TabIndex        =   11
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   4
      Left            =   3510
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3735
      Width           =   375
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   3
      Left            =   3270
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3735
      Width           =   255
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   2
      Left            =   2430
      MaxLength       =   6
      TabIndex        =   3
      Top             =   3735
      Width           =   855
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1950
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "P"
      Top             =   3735
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3810
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整批"
      Height          =   180
      Index           =   0
      Left            =   276
      TabIndex        =   0
      Top             =   1350
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.CheckBox chkKind 
      Height          =   225
      Index           =   1
      Left            =   1935
      TabIndex        =   17
      Top             =   480
      Value           =   1  '核取
      Width           =   285
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   42
      Top             =   72
      Visible         =   0   'False
      Width           =   1800
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3175;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "程序人員："
      Height          =   180
      Left            =   264
      TabIndex        =   41
      Top             =   144
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   2
      X1              =   3000
      X2              =   3150
      Y1              =   3323
      Y2              =   3323
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   1
      X1              =   3000
      X2              =   3150
      Y1              =   2993
      Y2              =   2993
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   0
      X1              =   3000
      X2              =   3150
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FMP　法定期限："
      Height          =   180
      Index           =   14
      Left            =   495
      TabIndex        =   40
      Top             =   3240
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P 　　本所期限："
      Height          =   180
      Index           =   13
      Left            =   495
      TabIndex        =   39
      Top             =   2910
      Width           =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家：非台灣(大陸,香港,澳門,PCT...)"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   495
      TabIndex        =   33
      Top             =   2580
      Width           =   3345
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家：台灣"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   495
      TabIndex        =   32
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P 　　本所期限："
      Height          =   180
      Index           =   12
      Left            =   495
      TabIndex        =   31
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "4.TW-SUPA"
      Height          =   180
      Index           =   11
      Left            =   2295
      TabIndex        =   28
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "執行日期："
      Height          =   180
      Index           =   8
      Left            =   495
      TabIndex        =   21
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "3.主張國外優先權"
      Height          =   180
      Index           =   0
      Left            =   2295
      TabIndex        =   20
      Top             =   810
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1.年費"
      Height          =   180
      Index           =   4
      Left            =   2295
      TabIndex        =   13
      Top             =   510
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "列印地址條印表機："
      Height          =   180
      Left            =   150
      TabIndex        =   16
      Top             =   4995
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日："
      Height          =   180
      Index           =   5
      Left            =   510
      TabIndex        =   15
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2.實體審查"
      Height          =   180
      Index           =   3
      Left            =   3465
      TabIndex        =   14
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "通知函類別："
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   12
      Top             =   495
      Width           =   1080
   End
End
Attribute VB_Name = "frm040303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Const ET01 As String = "12"
'********* 90.12.03    nick
Dim tmpNP22 As String
'**********************

'Remove by Morgan 2008/8/13 改開窗定稿
'' 暫存申請人
'Dim m_CustList() As String
'Dim m_CustListCount As Integer

''  91.08.07  nick  暫存本所案號
'Dim m_CP() As String
''   end
'end 2008/8/13

' 預設印表機
Dim m_DefaultPrinter As String
'Add By Cheng 2002/03/06
Dim m_PA46 As String
Dim m_PA09 As String
Dim m_NP09 As String
'Add By Cheng 2002/05/16
Dim PLeft(0 To 7) As Integer
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'Add By Cheng 2002/10/24
Dim m_strPA08 As String '記錄專利種類
Dim m_strPA10 As String '申請日 Add by Morgan 2009/10/7
'Add by Morgan 2005/5/16
Dim m_CurCP(1 To 4) As String '現在資料的本所號
Dim m_iDiscount As Integer '可減免退費金額
Dim m_iYear1 As Integer '減免退費起始年度
Dim m_iYear2 As Integer '減免退費終止年度
Dim m_NP22 As String
Dim m_Select As String
Dim m_bFirstYear As Boolean '是否繳第一次年費
'Dim pa26 As String

'Add by Morgan 2009/12/7
'Dim strPrint As String 'Remove by Morgan 2009/12/22 改國外部自行列印
Dim m_bolFMP As Boolean
Dim m_lngRefund As Long 'Add by Morgan 2011/7/6 預繳未退金額
Dim m_LD18 As String 'Added by Morgan 2014/6/12
Dim m_PA26 As String 'Added by Morgan 2014/6/12
Dim m_FMP_ET02 As String 'Added by Morgan 2014/8/20
Dim m_SelArea As String 'Add by Lydia 2015/01/27
'Remove by Morgan 2008/8/13 改開窗定稿
'' 清除申請人代碼暫存區
'Private Sub ClearCustList()
'   If m_CustListCount > 0 Then
'      Erase m_CustList
'      Erase m_CP
'   End If
'   m_CustListCount = 0
'End Sub
'Added by Morgan 2018/3/7 開其他的表單可能會清除全域變數值，故表單改用區域變數
Dim m_FMP2openSQL As String, m_FMP2open As Boolean
'Added by Lydia 2019/08/30
Dim m_Date1 As String, m_Date2 As String 'P案所限區間(大陸案)
Dim m_FMPDate1 As String, m_FMPDate2 As String 'FMP案所限區間(大陸案)
Dim m_DateTW1 As String, m_DateTW2 As String 'Added by Lydia 2019/12/16 P案所限區間(台灣案)
'Added by Morgan 2025/1/16
Dim rsQuery As ADODB.Recordset
Dim mSeqNo As String, stVTBX As String
'end 2025/1/16

Private Sub Check1_Click()
   Check2.Value = Abs(Check1.Value - 1) * Check2.Value
End Sub

Private Sub Check2_Click()
   Check1.Value = Abs(Check2.Value - 1) * Check1.Value
End Sub

Private Sub chkKind_Click(Index As Integer)
   Dim index1 As Integer, index2 As Integer
   'Added by Lydia 2015/04/20 + TW-SUPA
   If chkKind(3).Value = 1 Then Option1(1).Value = True
    If chkKind(Index).Enabled = True And Option1(1).Value = True Then
       For index1 = 0 To 3
          If index1 <> Index Then
             chkKind(index1).Value = Abs(chkKind(Index).Value - 1) * chkKind(index1).Value
          End If
       Next
    End If

End Sub

'Modify by Morgan 2006/4/3 重整
Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         Screen.MousePointer = vbHourglass
         '選擇下次繳費日區間(整批)
         If Option1(0).Value = True Then
            If FormCheck = True Then
               'Added by Lydia 2025/07/25
               If cntFrm040303New = "Y" Then
                  Process_New
               Else
               'end 2025/07/25
                  Process
               End If
            End If
         '選擇本所案號
         Else
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
            If FormCheck3 = True Then
               'Modify by Morgan 2007/9/4
               'Process3
               'Added by Lydia 2015/04/20
               If chkKind(3).Value = 1 Then 'TW-SUPA
                  Process5
               'end 2015/04/20
               ElseIf chkKind(0).Value = 1 Then '主張國內優先權
                  pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(0) 'Add By Sindy 2010/11/29
                  Process4 '通知主張國外優先權
               Else
                  Process3 '年費, 實審
               End If
               'end 2007/9/4
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1 '結束
         Unload Me
   End Select
End Sub

'Add By Cheng 2002/05/16
'選擇下次繳費日時才列印期限表(選擇本所案號時不列印此報表)
Private Sub Process1()
Dim Rs As New ADODB.Recordset
Dim intPage As Integer
Dim strDate As String
Dim strNation As String
Dim ii As Integer
Dim jj As Integer
Dim arrJJ
Dim intMaxJJ As Integer
Dim kk As Integer
Dim arrKK
Dim intMaxKK As Integer
Dim Prn As Printer
Dim iPrint As Integer
Dim iPrint1 As Integer
Dim strDeadLineCon As String
Dim strDLCon As String
'Dim bolIsYearFee As Boolean
'Add by Morgan 2006/4/17
Dim strYears As String

'92.04.03 nick add left join
'strSQL = "Select NVL(R04030305,''),PA09,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA11,PA72,DECODE(PA08,'1',NA21,'2',NA23,'3',NA25,'')  From R040303,PATENT,NATION WHERE R04030301=PA01 AND R04030302=PA02 AND R04030303=PA03 AND R04030304=PA04 AND PA09=NA01(+) AND ID='" & strUserNum & "' " & _
         "ORDER BY 1,2,3 "
'Modify by Morgan 2006/2/20 進度檔有1604 or 1606 or 1907 or 413案件性質發文時不印
'strSQL = "Select NVL(R04030305,''),PA09,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA11,PA72,DECODE(PA08,'1',NA21,'2',NA23,'3',NA25,'')  From R040303,PATENT,NATION WHERE R04030301=PA01(+) AND R04030302=PA02(+) AND R04030303=PA03(+) AND R04030304=PA04(+) AND PA09=NA01(+) AND ID='" & strUserNum & "' " & _
'         "ORDER BY 1,2,3 "
'Modify by Morgan 2006/4/13 加R04030306以判斷是年費或實審, PA08,PA14 判斷新型適用新法或舊法
'Modify by Morgan 2007/4/4 加判斷未收文恢復權利414,P61469
'Modify by Morgan 2007/10/24 +閉卷的都不印
strSql = "Select NVL(R04030305,''),PA09,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA11,PA72,DECODE(PA08,'1',NA21,'2',NA23,'3',NA25,''),R04030306, PA08, PA14  From R040303,PATENT,NATION WHERE R04030301=PA01(+) AND R04030302=PA02(+) AND R04030303=PA03(+) AND R04030304=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND ID='" & strUserNum & "' " & _
         " and not exists(select * from caseprogress A where cp01=R04030301 and cp02=R04030302 and cp03=R04030303 and cp04=R04030304" & _
         " and cp10 in ('1604','1606','1907','413') and cp27 is not null and not exists(select * from caseprogress B where B.cp01=A.cp01 and B.cp02=A.cp02 and B.cp03=A.cp03 and B.cp04=A.cp04 and B.cp10='414' and B.cp05>A.cp05 and B.cp57 is null))" & _
         " ORDER BY 1,2,3 "
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
Rs.CursorLocation = adUseClient
Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Rs.RecordCount > 0 Then
   intPage = 1
   '搜尋 Printer
   For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next

   GetPrintLeft
   PrintTitle intPage, ChangeTStringToTDateString(Rs.Fields(0).Value - 19110000), NationQuery(Rs.Fields(1).Value, 1)
   ii = 0
   iPrint = 2700
   iPrint1 = 2700
   Rs.MoveFirst
   strDate = Rs.Fields(0).Value
   strNation = "" & Rs.Fields(1).Value

'Remove by Morgan 2006/4/13 期限表合併,改用R04030306判斷
'   'Add by Morgan 2004/6/9 年費期限表的期限內容有時會印成實體審查，尚未確認原因
'   If m_Select = "1" Then
'      bolIsYearFee = True
'   Else
'      bolIsYearFee = False
'   End If

   While Not Rs.EOF
      If strDate <> Rs.Fields(0).Value Or strNation <> Rs.Fields(1).Value Then
         intPage = intPage + 1
         Printer.NewPage
         PrintTitle intPage, ChangeTStringToTDateString(Rs.Fields(0).Value - 19110000), NationQuery(Rs.Fields(1).Value, 1)
         ii = 0
         iPrint = 2700
         iPrint1 = 2700
         strDate = Rs.Fields(0).Value
         strNation = "" & Rs.Fields(1).Value
      End If
      If ii >= 40 Then
         intPage = intPage + 1
         Printer.NewPage
         PrintTitle intPage, ChangeTStringToTDateString(Rs.Fields(0).Value - 19110000), NationQuery(Rs.Fields(1).Value, 1)
         ii = 0
         iPrint = 2700
         iPrint1 = 2700
         strDate = Rs.Fields(0).Value
         strNation = "" & Rs.Fields(1).Value
      End If
        '通知函類別選擇年費
      'If m_Select = "1" Then
      'Modify by Morgan 2006/4/13 期限表合併,改用R04030306判斷
      'If bolIsYearFee Then
      If "" & Rs("R04030306") = "1" Then
         'Modify by Morgan 2006/4/17 台灣新型新舊法判斷
         strYears = "" & Rs.Fields(4).Value
         If Rs("PA08") = "2" And Rs("PA09") = "000" And Val("" & Rs("PA14")) < 20040700 Then
            strYears = "1,2,3,4,5,6,7,8,9,10,11,12"
         Else
            strYears = "" & Rs.Fields(5).Value
         End If

         'Modify By Cheng 2003/02/12
         '若有繳年費記錄
         If Rs.Fields(4).Value <> "" Then
             arrJJ = Split(Rs.Fields(4).Value, ",")
             jj = UBound(arrJJ)
             intMaxJJ = Val("0" & arrJJ(0))
             If jj > 0 Then
                For jj = 1 To UBound(Split(Rs.Fields(4).Value, ","))
                   If intMaxJJ < Val("0" & arrJJ(jj)) Then
                      intMaxJJ = Val("0" & arrJJ(jj))
                   End If
                Next jj
             End If
             arrKK = Split(strYears, ",")
             kk = UBound(arrKK)
             intMaxKK = Val("0" & arrKK(0))
             If kk > 0 Then
                For kk = 1 To UBound(Split(strYears, ","))
                   If intMaxKK < Val("0" & arrKK(kk)) Then
                      intMaxKK = Val("0" & arrKK(kk))
                   End If
                Next kk
             End If
             If intMaxJJ + 1 <= intMaxKK Then
                strDLCon = "第" & intMaxJJ + 1 & "年"
             Else
                strDLCon = ""
             End If
         '若無繳年費記錄
         Else
             strDLCon = ""
         End If
      'Add by Morgan 2006/5/15
      ElseIf "" & Rs("R04030306") = "3" Then
         strDLCon = "進入國家階段"
      'Added by Morgan 2024/11/6
      ElseIf "" & Rs("R04030306") = "4" Then
         strDLCon = "補償期年費"
      '通知函類別選擇實體審查
      Else
         strDLCon = "實體審查"
      End If
      '列印左半邊
      If ii < 20 Then
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "" & Rs.Fields(2)
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print Left("" & Rs.Fields(3), 11)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "" & strDLCon
         Printer.CurrentX = PLeft(4) - 300
         Printer.CurrentY = iPrint
         Printer.Print "｜"
         iPrint = iPrint + 300

         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print String(250, "-")
         iPrint = iPrint + 300
      '列印右半邊
      Else
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint1
         Printer.Print "" & Rs.Fields(2)
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint1
         Printer.Print Left("" & Rs.Fields(3), 11)
         Printer.CurrentX = PLeft(6)
         Printer.CurrentY = iPrint1
         Printer.Print "" & strDLCon
         iPrint1 = iPrint1 + 300
         iPrint1 = iPrint1 + 300
      End If
      Rs.MoveNext
      ii = ii + 1
   Wend
   Printer.EndDoc
End If
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
End Sub

Private Sub GetPrintLeft()
PLeft(0) = 200
PLeft(1) = 1800 + 100
PLeft(2) = 3200 + 100
PLeft(3) = 4600
PLeft(4) = 6000 + 100
PLeft(5) = 7600 + 200
PLeft(6) = 9000 + 200
PLeft(7) = 10400
End Sub

'Add By Cheng 2002/12/20
Private Sub GetPrintLeft_1()
PLeft(0) = 200
PLeft(1) = 2200
PLeft(2) = 6000
PLeft(6) = 8000
End Sub
'Modified by Morgan 2012/10/5 +bolIsTripleFee:補繳3倍年費期限表
Private Sub PrintTitle(Page As Integer, strDate As String, strNation As String, Optional bolIsTripleFee As Boolean)
'Page : 頁數
'strDate : 期限日期
Dim i As Integer
   
i = 500
If Page = 1 Then Printer.Orientation = vbPRORPortrait
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 3500
Printer.CurrentY = i
If bolIsTripleFee Then
   Printer.Print "補繳3倍年費期限表"
Else
   Printer.Print "繳年費/實體審查期限表"
End If
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 800
Printer.Print "列印人　 : " & strUserName
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 800
Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1100
Printer.Print "期限日期 : " & strDate
Printer.CurrentX = 4500
Printer.CurrentY = i + 1100
Printer.Print "申請國家 : " & strNation
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 1100
Printer.Print "頁　　次 : " & Page
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1400
Printer.Print String(250, "-")

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = i + 1700
Printer.Print "申請案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = i + 1700
Printer.Print "期限內容"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = i + 1700
Printer.Print "備註"

Printer.CurrentX = PLeft(4)
Printer.CurrentY = i + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = i + 1700
Printer.Print "申請案號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = i + 1700
Printer.Print "期限內容"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = i + 1700
Printer.Print "備註"

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 2000
Printer.Print String(250, "-")

Printer.CurrentX = PLeft(0)
Printer.CurrentY = 2700 + 300 * 40 - 100
Printer.Print String(250, "-")
Printer.CurrentX = PLeft(0)
Printer.CurrentY = 2700 + 300 * 41 - 100
Printer.Print "期限日期 : " & strDate
Printer.CurrentX = 4500
Printer.CurrentY = 2700 + 300 * 41 - 100
Printer.Print "申請國家 : " & strNation
End Sub

'Add By Cheng 2002/12/20
Private Sub PrintTitle_1(Page As Integer)
'Page : 頁數
'strDate : 期限日期
Dim i As Integer
   
i = 500
If Page = 1 Then Printer.Orientation = vbPRORPortrait
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 4000
Printer.CurrentY = i
Printer.Print "專利權消滅清單"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 800
Printer.Print "列印人　 : " & strUserName
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 800
Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))

'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = i + 1100
'Printer.Print "期限日期 : " & strDate
'Printer.CurrentX = 4500
'Printer.CurrentY = i + 1100
'Printer.Print "申請國家 : " & strNation
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 1100
Printer.Print "頁　　次 : " & Page
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1400
Printer.Print String(250, "-")

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = i + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = i + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = i + 1700
Printer.Print "智權人員"

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 2000
Printer.Print String(250, "-")
End Sub

Private Sub Form_Load()

   Me.Height = 5280 'Added by Lydia 2019/12/16 隱藏地址條, 因為已不再列印
   
   MoveFormToCenter Me
   intWhere = 國內
   Option1_Click 0
  
'Add by Lydia 2015/01/27 開放外專程序人員操作FMP寰華案件。當非FMP寰華權限,不可看寰華案=>回傳SQL
FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
If Pub_StrUserSt03 = "M51" Then
   If MsgBox("電腦中心人員請注意你現在是要看FMP寰華案嗎?", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
      FMP2openSQL = Replace(FMP2openSQL, "not", "")
   Else
      MsgBox "現在本報表不可查FMP寰華案"
   End If
End If
'end 2015/01/27

'Added by Morgan 2018/3/7 複製到區域變數
m_FMP2open = FMP2open
m_FMP2openSQL = FMP2openSQL
'end 2018/3/7

   'Added by Morgan 2025/1/15
   If FMP2open = False And strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label3.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label3)
   End If
   'end 2025/1/15
   
'Added by Lydia 2015/04/20
'Me.Height = 4500
'Me.Height = 5730 'Remove by Lydia 2019/12/16
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2020/4/10
   Set frm040303 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim txt As TextBox, i As Integer
On Error Resume Next
   
   For Each txt In Text1
      txt.Enabled = False
   Next
   
   Text2.Enabled = False
   Frame2.Enabled = False
   Text3.Enabled = False
   
   Select Case Index
      Case 0
         'Add by Morgan 2006/4/4
         Check2.Enabled = True
         chkKind(1).Value = vbChecked
         chkKind(2).Value = vbChecked
         'Add by Morgan 2007/9/4
         'Modified by Morgan 2015/5/27 目前還沒決定要通知,先取消,若將來要通知時需討論此種信函要新增的案件性質-- 郭雅娟
         'chkKind(0).Value = vbChecked
         chkKind(0).Value = vbUnchecked
         chkKind(0).Enabled = False
         'end 2015/5/27
         'Remove by Morgan 2007/9/6 改可讓User自行勾選--郭
         'chkKind(0).Enabled = False
         'chkKind(1).Enabled = False
         'chkKind(2).Enabled = False
         
         'Remove by Morgan 2009/12/23 鎖住
         'For i = 5 To 8
         '   Text1(i).Enabled = True
         'Next
         
         'Add by Morgan 2009/12/7
         Frame2.Enabled = True
         Text3.Enabled = True
         If Text3 = "" Then
            strExc(0) = Right(strSrvDate(1), 2)
            If Val(strExc(0)) <= 10 Then
               Text3 = Left(strSrvDate(1), 6) - 191100
               Option3(0).Value = True
            ElseIf Val(strExc(0)) <= 20 Then
               Text3 = Left(strSrvDate(1), 6) - 191100
               Option3(1).Value = True
            Else
               Text3 = Left(CompDate(1, 1, strSrvDate(1)), 6) - 191100
               Option3(0).Value = True
            End If
            
         End If
         'Added by Lydia 2025/07/25
         If cntFrm040303New = "Y" Then
            Call PUB_SetDateConFrm040303(Text3.Text, IIf(Option3(0).Value = True, "1", "2"), m_Date1, m_Date2, m_FMPDate1, m_FMPDate2, m_DateTW1, m_DateTW2)
            Call SetDateTextBox
         Else
         'end 2025/07/25
            'Modified by Lydia 2019/08/30
            'SetDateCondition
            SetDateCondition Text3.Text, True
            'END 2009/12/7
         End If
      Case 1
         For i = 2 To 4
            Text1(i).Enabled = True
         Next
         Text2.Enabled = True
         
         'Add by Morgan 2006/4/4
         Check2.Value = vbUnchecked
         Check2.Enabled = False
         chkKind(1).Value = vbChecked
         'modify by sonia 2014/7/3 玲玲說勾本所案號時仍維持三種類別自動勾選
         'chkKind(2).Value = vbUnchecked
         'chkKind(0).Value = vbUnchecked 'Add by Morgan 2007/9/4
         chkKind(2).Value = vbChecked
         'Modified by Morgan 2014/8/20  改要跑才勾--玲玲
         'chkKind(0).Value = vbChecked
         chkKind(0).Value = vbUnchecked
         'end 2014/8/20
         'end 2014/7/3
         
         chkKind(0).Enabled = True
         chkKind(1).Enabled = True
         chkKind(2).Enabled = True
   End Select
End Sub

Private Sub Option2_Click(Index As Integer)
   'Added by Lydia 2025/07/25
   If cntFrm040303New = "Y" Then
      Call PUB_SetDateConFrm040303(Text3.Text, IIf(Option3(0).Value = True, "1", "2"), m_Date1, m_Date2, m_FMPDate1, m_FMPDate2, m_DateTW1, m_DateTW2)
      Call SetDateTextBox
   Else
   'end 2025/07/25
      'Modified by Lydia 2019/08/30
      'SetDateCondition
      SetDateCondition Text3.Text, True
   End If
End Sub

Private Sub Option3_Click(Index As Integer)
   'Added by Lydia 2025/07/25
   If cntFrm040303New = "Y" Then
      Call PUB_SetDateConFrm040303(Text3.Text, IIf(Option3(0).Value = True, "1", "2"), m_Date1, m_Date2, m_FMPDate1, m_FMPDate2, m_DateTW1, m_DateTW2)
      Call SetDateTextBox
   Else
   'end 2025/07/25
      'Modified by Lydia 2019/08/30
      'SetDateCondition
      SetDateCondition Text3.Text, True
   End If
End Sub

'Added by Morgan 2012/1/5
Private Sub Text1_Change(Index As Integer)
   Select Case Index
      Case 1, 2, 3, 4
         Text2.Text = ""
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 6 '下次繳費日
         'Modify By Cheng 2002/09/10
         If blnClkSure = False Then
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Text1(Index - 1).SetFocus
            End If
         Else
            blnClkSure = False
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Option1(0).Value = True Then
      Select Case Index
         Case 5, 6, 7, 8
            If Text1(Index) <> "" Then
               Cancel = Not ChkDate(Text1(Index))
            End If
      End Select
   End If
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Text2_GotFocus()
Dim strCase As String
   
   TextInverse Text2
   If Option1(0).Value = True Then Exit Sub
   
   If chkKind(1).Value = vbChecked Then
      m_Select = "1"
   Else
      m_Select = "2"
   End If
   
   Select Case m_Select
      Case "1"
         'Modify by Morgan 2006/5/15
         'strCase = 年費
         'Modified by Morgan 2012/10/23 +維持費
         'Modified by Morgan 2024/11/1 +615補償期年費
         strCase = 年費 & "," & 維持費 & "," & 延展費 & ",119,615"
      Case "2"
         strCase = 實體審查
      Case Else
         MsgBox "通知函類別不可空白！", vbOKOnly, "錯誤！"
         Exit Sub
   End Select
   If Text1(1).Text = "" Or Text1(2).Text = "" Then
      MsgBox "本所案號不可空白，請重新輸入 !", vbCritical
      Text1(2).SetFocus
   Else
      'Add By Cheng 2002/03/06
      strExc(0) = "SELECT PA09,PA46,PA08,PA10,PA26,pa75 FROM PATENT WHERE PA01='" & Text1(1).Text & "' AND PA02='" & Text1(2).Text & "' AND PA03='" & IIf(Len(Text1(3).Text) > 0, Text1(3).Text, "0") & "' AND PA04='" & IIf(Len(Text1(4).Text) > 0, Text1(4).Text, "00") & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_PA09 = "" & RsTemp.Fields(0).Value
         m_PA46 = "" & RsTemp.Fields(1).Value
         m_strPA08 = "" & RsTemp.Fields("PA08")
         m_strPA10 = "" & RsTemp.Fields("PA10")
         m_PA26 = "" & RsTemp.Fields("pa26") 'Added by Morgan 2014/6/12
      Else
         m_PA09 = ""
         m_PA46 = ""
         m_strPA08 = ""
         m_strPA10 = ""
         m_PA26 = "" 'Added by Morgan 2014/6/12
      End If

      'Added by Moragn 2014/8/20 若年費及實審都有勾選時先全抓再判斷案件性質留下正確的勾選--玲玲
      If m_Select = "1" And chkKind(2).Value = vbChecked Then
         strCase = strCase & ",416"
      End If
      'end 2014/8/20
      
      strExc(0) = "SELECT NP08,np22,NP09,NP07 FROM NEXTPROGRESS WHERE " & ChgNextProgress(Text1(1).Text & Text1(2).Text & Text1(3).Text & Text1(4).Text) & " AND NP06 IS NULL AND NP07 IN (" & strCase & ") order by np09"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text2.Text = ChangeWStringToTString(CheckStr(RsTemp.Fields(0)))
         tmpNP22 = CheckStr(RsTemp.Fields(1).Value)
         m_NP09 = "" & RsTemp.Fields(2).Value
         
         'Added by Morgan 2014/8/20
         If RsTemp.Fields("NP07") = "416" Then
            If chkKind(1).Value = vbChecked Then
               chkKind(1).Value = vbUnchecked
            End If
         Else
            If chkKind(2).Value = vbChecked Then
               chkKind(2).Value = vbUnchecked
            End If
         End If
         'end 2014/8/20
         
      'Add By Cheng 2002/03/06
      Else
         Text2.Text = ""
         tmpNP22 = ""
         m_NP09 = ""
      End If
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   'Modify by Morgan 2007/9/4 主張國外優先權不必輸下次繳費日
   'Modified by Lydia 2015/04/20 +TW-SUPA不必輸下次繳費日
  ' If chkKind(0).Value <> 1 Then
   If chkKind(0).Value + chkKind(3).Value = 0 Then
      Cancel = Not ChkDate(Text2.Text)
      If Cancel Then TextInverse Text2
   End If
End Sub

'Remove by Morgan 2008/8/13 改開窗定稿
'Private Sub PrintAddress()
'   Dim nPageNo As String
'   Dim strCust As String
'   Dim nPos As Integer
'
'   ' 流水號
'   nPageNo = 1
'   For nPos = 0 To m_CustListCount - 1
'      strCust = m_CustList(nPos)
'
''Remove by Morgan 2004/11/15 要取九碼否則會印錯
''      If Len(strCust) > 8 Then
''         strCust = Left(strCust, 8)
''      Else
''         strCust = strCust & String(8 - Len(strCust), "0")
''      End If
'
'      ' 列印地址條
'      Load frm083014
'      '****** 90.11.29  nick
'      frm083014.Hide
'      '**********************
'      '****** 91.08.07   nick  加入本所案號
'      frm083014.text1(6).Text = m_CP(nPos)
'      '************************************
'      'Add By Cheng 2002/12/20
'      '傳本所案號
'      frm083014.SetCaseNo m_CP(nPos)
'      '含不寄雜誌的客戶
'      frm083014.text1(5).Text = "Y"
'
'      frm083014.text1(0).Text = strCust
'      ' 只印一份
'      frm083014.text1(3).Text = "1"
'      ' 印中文
'      frm083014.text1(4).Text = "1"
'      ' 地址條流水號
'      frm083014.SetPageNo nPageNo
'      ' 設定印表機
'      frm083014.SetPrinter cmbPrinter.List(cmbPrinter.ListIndex)
'      frm083014.cmdPrint_Click
'      frm083014.cmdBack_Click
'      ' 流水號遞增
'      nPageNo = nPageNo + 1
'   Next nPos
'
'   ' 清除申請人串列
'   ClearCustList
'End Sub

'依不同專利種類取得該年年費屆滿日期
'Modified by Morgan 2023/6/5 +strPA25
Private Function getPA72Year(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, Optional strPA25 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPayYear As String '記錄各專利種類的繳費年度
Dim arrPayYear '記錄各專利種類的繳費年度陣列
Dim strMaxPayYear As String '記錄各專利種類的最大繳費年度
Dim ii As Integer
Dim arrPA72 '記錄已繳費年度陣列
Dim strMaxPA72 As String '已繳費年度
Dim strPayDATE As String '記錄各專利種類的年費起算日

getPA72Year = ""
strPayYear = ""
StrSQLa = "Select * From Patent,Nation Where " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04) & " And PA09=NA01(+) "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   strPA25 = "" & rsA("pa25") 'Added by Morgan 2023/6/5
    Select Case "" & rsA("PA08").Value
    Case "1" '發明
        If "" & (rsA("NA21").Value) <> "" Then
            strPayYear = "" & (rsA("NA21").Value)
            strPayDATE = "" & (rsA("NA06").Value)
        End If
    Case "2" '新型
        If "" & (rsA("NA23").Value) <> "" Then
            strPayYear = "" & (rsA("NA23").Value)
            strPayDATE = "" & (rsA("NA08").Value)
        End If
    Case "3" '設計
        If "" & (rsA("NA25").Value) <> "" Then
            strPayYear = "" & (rsA("NA25").Value)
            strPayDATE = "" & (rsA("NA10").Value)
        End If
    End Select
    
    m_strPA08 = "" & rsA("PA08").Value
    
    If strPayYear <> "" Then
        arrPayYear = Split(strPayYear, ",")
        strMaxPayYear = "0"
        For ii = LBound(arrPayYear) To UBound(arrPayYear)
                If Val(strMaxPayYear) < Val(arrPayYear(ii)) Then
                    '取得設定的最大繳費年度
                    strMaxPayYear = arrPayYear(ii)
                End If
        Next ii
        If "" & rsA("PA72").Value = "" Then
            '已繳費年度
            strMaxPA72 = "0"
        Else
            arrPA72 = Split("" & rsA("PA72").Value, ",")
            strMaxPA72 = "0"
            For ii = LBound(arrPA72) To UBound(arrPA72)
                If Val(strMaxPA72) < Val(arrPA72(ii)) Then
                    '取得目前最大已繳費年度
                    strMaxPA72 = arrPA72(ii)
                End If
            Next ii
        End If
    End If
    
   Select Case strPayDATE
      Case 申請日
                 getPA72Year = CompDate(0, strMaxPA72, rsA("PA10").Value)
      Case 公開日
                 getPA72Year = CompDate(0, strMaxPA72, rsA("PA12").Value)
      Case 准駁日
                 getPA72Year = CompDate(0, strMaxPA72, rsA("PA20").Value)
      Case 公告日
                 getPA72Year = CompDate(0, strMaxPA72, rsA("PA14").Value)
      Case 發證日
                 getPA72Year = CompDate(0, strMaxPA72, rsA("PA21").Value)
   End Select
   'Modify by Morgan 2008/7/16 台灣才要減1天--郭
   If "" & rsA("PA09") = "000" Then
      getPA72Year = CompDate(2, -1, getPA72Year)
   End If
End If

If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'依不同專利種類取得下次繳費年度
'Modified by Morgan 2023/6/5 +strPA25
Private Function getPA72NextYear(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, Optional p_stMaxFeeYear As String, Optional p_bFirstYear As Boolean, Optional ByRef strPA25 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPayYear As String '記錄各專利種類的繳費年度
Dim arrPayYear '記錄各專利種類的繳費年度陣列
Dim strMaxPayYear As String '記錄各專利種類的最大繳費年度
Dim ii As Integer
Dim arrPA72 '記錄已繳費年度陣列
Dim strMaxPA72 As String '下一次繳費年度
Dim strPA14 As String
'Dim strPA25 As String 'Removed by Morgan 2023/6/5
Dim strEffDate As String '有效專用日期

getPA72NextYear = ""
strPayYear = ""
StrSQLa = "Select * From Patent,Nation Where " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04) & " And PA09=NA01(+) "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    strPA25 = "" & rsA("pa25") 'Added by Morgan 2023/6/5
    strPA14 = "" & rsA("PA14")
    Select Case "" & rsA("PA08").Value
    Case "1" '發明
        If "" & (rsA("NA21").Value) <> "" Then
            strPayYear = "" & (rsA("NA21").Value)
        End If
    Case "2" '新型
        If "" & (rsA("NA23").Value) <> "" Then
            strPayYear = "" & (rsA("NA23").Value)
        End If
    Case "3" '設計
        If "" & (rsA("NA25").Value) <> "" Then
            strPayYear = "" & (rsA("NA25").Value)
        End If
    End Select
    
   'Add by Morgan 2007/11/6 舊法新型專用期12年
   If rsA("pa09") = "000" And rsA("pa08") = "2" And Val("" & rsA("pa14")) > 0 And Val("" & rsA("pa14")) < 20040701 Then
      strPayYear = "1,2,3,4,5,6,7,8,9,10,11,12"
   End If
   'end 2007/11/6
                           
    m_strPA08 = "" & rsA("PA08").Value
    
    If strPayYear <> "" Then
        arrPayYear = Split(strPayYear, ",")
        strMaxPayYear = "0"
        For ii = LBound(arrPayYear) To UBound(arrPayYear)
                If Val(strMaxPayYear) < Val(arrPayYear(ii)) Then
                    '取得設定的最大繳費年度
                    strMaxPayYear = arrPayYear(ii)
                End If
        Next ii
        
        If "" & rsA("PA72").Value = "" Then
            p_bFirstYear = True 'Add by Morgan 2008/5/7
            '下一次繳費年度
            'Modify by Morgan 2008/5/7
            'strMaxPA72 = "1"
            strMaxPA72 = arrPayYear(LBound(arrPayYear))
            'end 2008/5/7
            If Val(strMaxPA72) <= Val(strMaxPayYear) Then
                getPA72NextYear = strMaxPA72
            End If
        Else
            arrPA72 = Split("" & rsA("PA72").Value, ",")
            strMaxPA72 = "0"
            For ii = LBound(arrPA72) To UBound(arrPA72)
                If Val(strMaxPA72) < Val(arrPA72(ii)) Then
                    '取得目前最大已繳費年度
                    strMaxPA72 = arrPA72(ii)
                End If
            Next ii
            '下一次繳費年度
            strMaxPA72 = Val(strMaxPA72) + 1
            If Val(strMaxPA72) <= Val(strMaxPayYear) Then
                getPA72NextYear = strMaxPA72
            End If
        End If
        'Add by Morgan 2007/11/20 新版年費要抓下兩年費用
        If rsA("pa09") = "000" Then
            strPA14 = "" & rsA("pa14")
            strPA25 = "" & rsA("pa25")
            p_stMaxFeeYear = strMaxPA72
            If strPA14 <> "" And strPA25 <> "" Then
                For ii = Val(strMaxPA72) To Val(strMaxPayYear) - 1
                   strEffDate = CompDate(0, ii, strPA14)
                   If strEffDate > strPA25 Then
                      Exit For
                   Else
                      p_stMaxFeeYear = ii + 1
                   End If
                Next
            End If
        End If
        'end 2007/11/20
    End If
End If

If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 2002/05/16
'選擇下次繳費日且為年費通知函時, 才列印專利權消滅清單(選擇本所案號時不列印此報表)
Private Sub Process2()
Dim Rs As New ADODB.Recordset
Dim intPage As Integer
Dim strDate As String
Dim strNation As String
Dim ii As Integer
Dim jj As Integer
Dim arrJJ
Dim intMaxJJ As Integer
Dim kk As Integer
Dim arrKK
Dim intMaxKK As Integer
Dim Prn As Printer
Dim iPrint As Integer
Dim iPrint1 As Integer
Dim strDeadLineCon As String
Dim strDLCon As String
Dim strLstNo As String

'92.04.03 nick add left join
'strSQL = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA05,DECODE(PA09,'000',CPM03,CPM04),ST02 From R040303,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP " & _
        " WHERE R04030301=CP01 AND R04030302=CP02 AND R04030303=CP03 AND R04030304=CP04 AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND (CP10 ='1604' OR CP10='1606' OR CP10='1907') AND ID='" & strUserNum & "' " & _
        " GROUP BY CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA05,DECODE(PA09,'000',CPM03,CPM04),ST02 " & _
        " ORDER BY 1 "
'Modify by Morgan 2006/2/17 加413自請撤回
'Modify by Morgan 2007/4/4 加判斷未收文恢復權利414,P61469
'Modify by Morgan 2008/11/14 +有收文放棄專利權429也要印在消滅清單上 P-66851--敏惠
'Modified by Morgan 2011/11/28 恢復權利改發文日也要判斷(原來只判斷收文日但有發生例外 Ex.P-78503)
strSql = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 C01,PA05 C02,DECODE(PA09,'000',CPM03,CPM04) C03,ST02 C04,1 C05 From R040303,CASEPROGRESS A,PATENT,STAFF,CASEPROPERTYMAP " & _
        " WHERE R04030301=CP01(+) AND R04030302=CP02(+) AND R04030303=CP03(+) AND R04030304=CP04(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND (CP10 ='1604' OR CP10='1606' OR CP10='1907' OR CP10='413' OR CP10='429') AND ID='" & strUserNum & "' " & _
        " and not exists(select * from caseprogress B where B.cp01=A.cp01 and B.cp02=A.cp02 and B.cp03=A.cp03 and B.cp04=A.cp04 and B.cp10='414' and ((B.cp05>A.cp05 and B.cp27 is null) or B.cp27>A.cp27) and B.cp57 is null) GROUP BY CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA05,DECODE(PA09,'000',CPM03,CPM04),ST02 "
'Add by Morgan 2007/10/24 +閉卷的也要印
strSql = strSql & " UNION Select PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA05,'閉卷',ST02,2 From R040303,PATENT,STAFF,CUSTOMER " & _
        " WHERE R04030301=PA01(+) AND R04030302=PA02(+) AND R04030303=PA03(+) AND R04030304=PA04(+) AND PA57='Y' AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND ST01(+)=CU13 AND ID='" & strUserNum & "'"
'end 2007/10/24
'Add by Morgan 2020/8/6 +核駁期限已逾期超過3月未辦也要印
strSql = strSql & " UNION Select PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA05,'核駁已逾法限3個月未辦',ST02,3 From R040303,PATENT,STAFF,CUSTOMER " & _
        " WHERE R04030307='X' AND R04030301=PA01(+) AND R04030302=PA02(+) AND R04030303=PA03(+) AND R04030304=PA04(+) AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND ST01(+)=CU13 AND ID='" & strUserNum & "'"
'end 2020/8/6

strSql = strSql & " ORDER BY 1,5"

If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
Rs.CursorLocation = adUseClient
Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Rs.RecordCount > 0 Then
   intPage = 1
   '搜尋 Printer
   For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next
   
   GetPrintLeft_1
   PrintTitle_1 intPage
   ii = 0
   iPrint = 2700
   iPrint1 = 2700
   Rs.MoveFirst
   While Not Rs.EOF
      'Modify by Morgan 2007/10/24 案號相同時印一筆就好
      If Rs.Fields(0) = strLstNo Then
         GoTo NextRec
      Else
         strLstNo = Rs.Fields(0)
      End If
      'end 2007/10/24
      
      If ii >= 40 Then
         intPage = intPage + 1
         Printer.NewPage
         PrintTitle_1 intPage
         ii = 0
         iPrint = 2700
         iPrint1 = 2700
      End If
      
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "" & Rs.Fields(0)
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print Left("" & Rs.Fields(1), 9)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "" & Rs.Fields(2)
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint
         Printer.Print "" & Rs.Fields(3)
         iPrint = iPrint + 300
         
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print String(250, "-")
         iPrint = iPrint + 300
         
      ii = ii + 1
NextRec:
      Rs.MoveNext
   Wend
   Printer.EndDoc
End If
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
End Sub

'Add By Cheng 2003/10/13
'取得年費期限
Private Function GetNowNP09(strCaseNo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
 
GetNowNP09 = ""
StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(strCaseNo) & " And NP06 Is Null And NP07=" & 年費
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetNowNP09 = "" & rsA("NP09").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add by Morgan 2004/5/18
'取得下一繳費年度
'依不同專利種類取得下次繳費年度
'Modified by Morgan 2023/6/5 +strPA25
Private Function getNextPayYear(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, ByRef strNextPayDate As String, Optional ByRef strPA25 As String) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim strPayYear As String '記錄各專利種類的繳費年度
   Dim arrPayYear '記錄各專利種類的繳費年度陣列
   Dim ii As Integer
   Dim arrPA72 '記錄已繳費年度陣列
   Dim strMaxPA72 As String   '已繳費年度
   Dim strPayDATE As String '記錄各專利種類的年費起算日
   Dim strNP07 As String '案件性質
   Dim strPA09 As String '國家
   
   getNextPayYear = ""
   strPayYear = ""
   strMaxPA72 = "0"
   
   StrSQLa = "Select * From Patent,Nation Where " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04) & " And PA09=NA01(+) "
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   
   If rsA.RecordCount > 0 Then
      strPA25 = "" & rsA("pa25") 'Added by Morgan 2023/6/5
      strPA09 = "" & rsA("pa09")
      If "" & rsA("PA72").Value <> "" Then
         arrPA72 = Split("" & rsA("PA72").Value, ",")
         For ii = UBound(arrPA72) To LBound(arrPA72) Step -1
             If Val(arrPA72(ii)) > 0 Then
                 '取得目前最大已繳費年度
                 strMaxPA72 = arrPA72(ii)
                 Exit For
             End If
         Next ii
      'Remove by Morgan 2009/2/12 P-75730 的下次繳費年無法取得,改再下面抓下次繳費日時判斷是否為延展費控制
      ''2008/9/10 add by sonia P-074053催第一次時因為pa72空的,strMaxPA72會變為0
      'Else
      '   If "" & (rsA("NA23").Value) <> "" Then
      '      arrPA72 = Split("" & rsA("NA23").Value, ",")
      '      strMaxPA72 = "" & arrPA72(0)
      '   End If
      ''2008/9/10 end
      'end 2009/2/12
      End If
   
      m_strPA08 = "" & rsA("PA08").Value
       
      Select Case m_strPA08
         Case "1" '發明
             If "" & (rsA("NA21").Value) <> "" Then
                 strPayYear = "" & (rsA("NA21").Value)
                 strPayDATE = "" & (rsA("NA06").Value)
             End If
             strNP07 = "" & rsA("NA20").Value
         Case "2" '新型
             If "" & (rsA("NA23").Value) <> "" Then
                 strPayYear = "" & (rsA("NA23").Value)
                 strPayDATE = "" & (rsA("NA08").Value)
             End If
             strNP07 = "" & rsA("NA22").Value
         Case "3" '設計
             If "" & (rsA("NA25").Value) <> "" Then
                 strPayYear = "" & (rsA("NA25").Value)
                 strPayDATE = "" & (rsA("NA10").Value)
             End If
             strNP07 = "" & rsA("NA24").Value
      End Select
      
      If strPayYear <> "" Then
         arrPayYear = Split(strPayYear, ",")
         For ii = LBound(arrPayYear) To UBound(arrPayYear)
            If Val(arrPayYear(ii)) > Val(strMaxPA72) Then
               '取得設定的下次繳費年度
               getNextPayYear = arrPayYear(ii)
               Exit For
            End If
         Next ii
      End If
      
      'Modify by Morgan 2009/2/12 香港新型,設計為延展算法不同,非台灣案法限不必減1天
      Select Case strPayDATE
         Case 申請日
            If strNP07 = "605" Then
               strNextPayDate = CompDate(0, strMaxPA72, rsA("PA10").Value)
            Else
               strNextPayDate = CompDate(0, getNextPayYear, rsA("PA10").Value)
            End If
         Case 公開日
                    strNextPayDate = CompDate(0, strMaxPA72, rsA("PA12").Value)
         Case 准駁日
                    strNextPayDate = CompDate(0, strMaxPA72, rsA("PA20").Value)
         Case 公告日
                    strNextPayDate = CompDate(0, strMaxPA72, rsA("PA14").Value)
         Case 發證日
                    strNextPayDate = CompDate(0, strMaxPA72, rsA("PA21").Value)
      End Select
      
      If strPA09 = "000" Then
         strNextPayDate = CompDate(2, -1, strNextPayDate)
      End If
   End If
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Morgan 2005/5/17 台灣新增年費通知紀錄
Private Function UpdateAI() As Boolean
On Error GoTo ErrHnd

   strSql = "begin"
   strSql = strSql & " DELETE FROM ANNUITYINFORM WHERE AI01=" & m_NP22 & ";"
   strSql = strSql & " INSERT INTO ANNUITYINFORM (AI01,AI02,AI03,AI04) values (" & m_NP22 & ",'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')));"
   strSql = strSql & "end;"
   adoTaie.Execute strSql
   UpdateAI = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description
End Function

Private Function CheckCPExists(ByRef p_CP() As String) As Boolean
   'Modify by Morgan 2007/10/24 加429
   strSql = "select * from caseprogress A where cp01='" & p_CP(1) & "' and cp02='" & p_CP(2) & "' and cp03='" & p_CP(3) & "' and cp04='" & p_CP(4) & "' and cp10 in ('1604','1606','1907','413','429') and cp27 is not null"
   'Add by Morgan 2007/4/4 加判斷未收文恢復權利414,P61469
   'Modified by Morgan 2011/11/28 恢復權利改發文日也要判斷(原來只判斷收文日但有發生例外 Ex.P-78503)
   strSql = strSql & " and not exists(select * from caseprogress B where B.cp01=A.cp01 and B.cp02=A.cp02 and B.cp03=A.cp03 and B.cp04=A.cp04 and ((B.cp05>A.cp05 and B.cp27 is null) or B.cp27>A.cp27) and B.cp10='414' and B.cp57 is null)"
   'Added by Morgan 2015/9/7
   '自請撤回413要排除相關收文號非申請程序者 P-104761
   strSql = strSql & " and not exists(select * from caseprogress B where A.CP10='413' and B.cp09=A.cp43 and instr('" & NewCasePtyList & "',b.cp10)=0)"
   
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         CheckCPExists = True
      End If
   End With
End Function

'Add by Morgan 2006/4/3
'Memo by Lydia 2025/07/25 原程式保留不變，另外改用Process_New
Private Sub Process()
   Dim strCase As String
   Dim strTmp As String, strTmp2 As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTxt(1 To 20) As String
   Dim rsTemp10 As New ADODB.Recordset
   Dim strFee As String, strPoint As String
   Dim ii As Integer, jj As Integer
   Dim strPA72NextYear As String
   Dim strPA72Year As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim strCP09 As String '收文號
   Dim Prn As Printer
   Dim blnSitu1 As Boolean '下一程序檔是否有本所案號+案件性質本所期限在前半年且是否續辦為N的資料
   Dim strOldNP09 As String '半年前法定期限
   Dim stSitu As String '定稿處理狀況
   Dim idx As Integer
   Dim bolPrint As Boolean, bolPrint1 As Boolean
   Dim stMsg As String
   Dim strNP07 As String, strNP08 As String, strNP09 As String 'Add by Morgan 2006/5/15
   Dim strET03 As String
   Dim strMaxFeeYear As String '最大可繳費年度
   Dim bolDiscount As Boolean '是否可減免
   Dim iCopy As Integer
   Dim strNextYearFee As String '下次繳費金額
   Dim strPA75 As String
   Dim strNP23 As String 'Add by Morgan 2010//1/18 約定期限
   Dim iPlusFee As Integer 'Added by Morgan 2013/1/8 服務費外加金額(目前專利處大對台年費+500)
   Dim bolDualCaseUtility As Boolean, strInventionCaseNo As String, strInventionPA11 As String, strInventionPA77 As String 'Added by Morgan 2017/9/20 是否一案兩請新型案,發明案本所案號,發明案申請號,發明案彼所案號
   'Added by Lydia 2019/08/30
   Dim m_Except01 As String '和碩案之指定客戶(提早催期限)
   Dim strCon1 As String
   'Added by Lydia 2022/08/12 和碩案之期限(區分變數)
   Dim m_Ex1Date1 As String, m_Ex1Date2 As String 'P案所限區間(大陸案)
   Dim m_Ex1FMPDate1 As String, m_Ex1FMPDate2 As String 'FMP案所限區間(大陸案)
   Dim m_Ex1DateTW1 As String, m_Ex1DateTW2 As String 'P案所限區間(台灣案)
   'end 2022/08/12
   'Added by Lydia 2019/12/16
   Dim strPartA As String '台灣案SQL
   Dim strPartB As String '非台灣案SQL
   'Added by Lydia 2022/08/12 信邦案之指定客戶(晚催期限，法定期限前一個月通知)
   Dim m_ExceptB As String
   Dim m_ExBDate1 As String, m_ExBDate2 As String 'P案所限區間(大陸案)
   Dim m_ExBFMPDate1 As String, m_ExBFMPDate2 As String 'FMP案所限區間(大陸案)
   Dim m_ExBDateTW1 As String, m_ExBDateTW2 As String 'P案所限區間(台灣案)
   'end 2022/08/12
   'Added by Lydia 2022/08/25 大亞(X60601000)法定期限前4個月通知
   Dim m_ExceptC As String
   Dim m_ExCDate1 As String, m_ExCDate2 As String 'P案所限區間(大陸案)
   Dim m_ExCFMPDate1 As String, m_ExCFMPDate2 As String 'FMP案所限區間(大陸案)
   Dim m_ExCDateTW1 As String, m_ExCDateTW2 As String 'P案所限區間(台灣案)
   'end 2022/08/25
   'Added by Lydia 2023/11/09 康舒科技(X00497070)年費期限通知由三個月前改為一個月前
   Dim m_ExceptD As String
   Dim m_ExDDate1 As String, m_ExDDate2 As String 'P案所限區間(大陸案)
   Dim m_ExDFMPDate1 As String, m_ExDFMPDate2 As String 'FMP案所限區間(大陸案)
   Dim m_ExDDateTW1 As String, m_ExDDateTW2 As String 'P案所限區間(台灣案)
   'end 2023/11/09
   'Added by Lydia 2024/04/22 立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前
   'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
   Dim m_ExceptF As String
   Dim m_ExFDate1 As String, m_ExFDate2 As String 'P案所限區間(大陸案)
   Dim m_ExFFMPDate1 As String, m_ExFFMPDate2 As String 'FMP案所限區間(大陸案)
   Dim m_ExFDateTW1 As String, m_ExFDateTW2 As String 'P案所限區間(台灣案)
   'end 2024/04/22
   'Added by Lydia 2024/07/23 X38120000/ X38120030碩天科技/寧遠縣碩寧電子，中國專利年費通知程序，提早期限前兩個月前通知，原本只有一個月前通知。
   Dim m_ExceptG As String
   Dim m_ExGDate1 As String, m_ExGDate2 As String 'P案所限區間(大陸案)
   Dim m_ExGFMPDate1 As String, m_ExGFMPDate2 As String 'FMP案所限區間(大陸案)
   Dim m_ExGDateTW1 As String, m_ExGDateTW2 As String 'P案所限區間(台灣案)
   'end 2024/04/22
   
   Dim strPA25 As String 'Added by Morgan 2023/6/5
   Dim strSort As String 'Added by Morgan 2025/1/16
   
   
   bolPrint = False
   blnClkSure = False
                    
   '刪除暫存資料
   cnnConnection.Execute "Delete From R040303 Where ID='" & strUserNum & "'"
   '刪除接洽結案單暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   
'Remove by Morgan 2008/8/13 改開窗定稿
'   '清除
'   ClearCustList
'   '搜尋預設印表機
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   
   'Added by Lydia 2019/08/30 客戶X70017000和碩聯合科技股份有限公司及其關係企業(編號X70017010、X70017020、X70017030)
   '                                      所有P及CFP案年費的通知時間均提早為法定期限前6個月，實審的通知時間則是提早為申請日＋１年。
   '                                      若往後此客戶X70017新建關係企業，由智權人員通知設定。
   'Modified by Lydia 2022/02/23 改成共用模組取得
   'm_Except01 = "X70017000,X70017010,X70017020,X70017030"
   'Modified by Lydia 2022/02/08 長庚體系逐漸要將其專利案件回歸到產學中心，顯然之後皆須遵循其規則進行通知，故除顧服組客戶外，本所其他非顧服組也建議需一併設
   ''strCon1 = strCon1 & " AND INSTR('X70017000,X70017010,X70017020,X70017030',PA26)=0 "
   'm_Except01 = m_Except01 & ",X69365020,X75299000,X75299020,X69365060,X69365010,X69365030,X69365000,X69365040,X69365050"
   intI = Pub_Getfrm040303Except("X70017000", m_Except01)
   'end 2022/02/23
   
   'Added by Lydia 2022/08/12 因信邦電子(X39056000)承辦人反應本所專利年費、維持費繳交期限提前通知日期過早(現為三個月)，要求本所在法定期限前一個月通知即可
   intI = Pub_Getfrm040303Except("X39056000", m_ExceptB)
   'end 2022/08/12
   
   'Added by Lydia 2022/08/25 大亞電線電纜股份有限公司(X60601000)因為有導入TIPS智財管理制度，目前被要求專利案件年費需於法定期限前4個月通知，故客戶來電希望本所協助調整期限通知。
   intI = Pub_Getfrm040303Except("X60601000", m_ExceptC)
   'end 2022/08/25
   
   'Added by Lydia 2023/11/09 康舒科技(X00497070)年費期限通知由三個月前改為一個月前
   'intI = Pub_Getfrm040303Except("X00497070", m_ExceptD)  'Mark by Lydia 2024/12/05 (12/4)現因該公司內部問題，希望本所回復為原先三個月期限通知。
   'end 2023/11/09
   
   'Added by Lydia 2024/04/22 立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前
   'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
   intI = Pub_Getfrm040303Except("X01506000", m_ExceptF)
   'end 2024/04/22
   
   'Added by Lydia 2024/07/23 X38120000/ X38120030碩天科技/寧遠縣碩寧電子，中國專利年費通知程序，提早本所期限前兩個月前通知
   intI = Pub_Getfrm040303Except("X38120000", m_ExceptG)
   'end 2024/07/23
   
   'Modified by Lydia 2022/08/12 +IIf(m_ExceptB <> "", "," & m_ExceptB, "")
   'Modified by Lydia 2022/08/25 +IIf(m_ExceptC <> "", "," & m_ExceptC, "")
   'Modified by Lydia 2023/11/09 +IIf(m_ExceptD <> "", "," & m_ExceptD, "")
   'Modified by Lydia 2024/04/22 +IIf(m_ExceptF <> "", "," & m_ExceptF, "")
   'Modified by Lydia 2024/07/23 +IIf(m_ExceptG <> "", "," & m_ExceptG, "")
   strCon1 = strCon1 & " AND INSTR('" & m_Except01 & IIf(m_ExceptB <> "", "," & m_ExceptB, "") & IIf(m_ExceptC <> "", "," & m_ExceptC, "") & _
         IIf(m_ExceptD <> "", "," & m_ExceptD, "") & IIf(m_ExceptF <> "", "," & m_ExceptF, "") & IIf(m_ExceptG <> "", "," & m_ExceptG, "") & "',PA26)=0 "
   
   '可同時跑兩種通知函
   For idx = 1 To 2
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
      If chkKind(idx).Value = 1 Then
         m_Select = idx
         '判斷通知函類別
         Select Case m_Select
            Case "1"
               'Modified by Morgan 2012/10/23 +香港維持費
               'strCase = 年費
               'stMsg = "【年費】"
               'Modified by Morgan 2024/11/1 +615補償期年費
               strCase = 年費 & "," & 維持費 & ",615"
               stMsg = "【年費 維持費 補償期年費】"
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(4) 'Add By Sindy 2010/11/29
            Case "2"
               strCase = 實體審查
               stMsg = "【實體審查】"
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(3) 'Add By Sindy 2010/11/29
         End Select
         
         '申請國家條件
         strTmp = ""
         'Modify Morgan 2007/4/25
         '改可選台灣或非台灣
         'Remove by Lydia 2019/12/16 一併跑非大陸案
         'If Option2(0).Value = True Then
         '   pub_QL05 = pub_QL05 & ";" & Label1(1) & Option2(0).Caption 'Add By Sindy 2010/11/29
         '   strTmp = " AND PA09='000'"
         ''Else
          '  pub_QL05 = pub_QL05 & ";" & Label1(1) & Option2(1).Caption 'Add By Sindy 2010/11/29
          '  strTmp = " AND PA09<>'000'"
         'End If
         'end 2007/4/25
        'Add by Lydia 2015/01/27 +fmp寰華控制sql (m_selarea)
         Call ChangeSel(1) '將SQL改為對應NP
         
         '若為年費通知函, 不論是否閉卷皆要出現
         If m_Select = "1" Then
                'Modify by Morgan 2009/12/7 FMP案用法限條件且排除所限小於99/2/15的(已催過)
                'strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np08||NP09) in (select np02,np03,np04,np05,min(np08||NP09) FROM NEXTPROGRESS WHERE " & _
                   " NP02='P' and NP07 in (" & strCase & ",119) AND NP08 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
                   " AND NP06 IS NULL group by np02,np03,np04,np05)),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                   "NP05=pa04(+) " & strTmp & " AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & _
                   " ORDER BY PA09,PA01,PA02,PA03,PA04"
                'Modified by Lydia 2019/12/16
                'pub_QL05 = pub_QL05 & ";" & Label1(6) & text1(5) & "-" & text1(6) 'Add By Sindy 2010/11/29
                pub_QL05 = pub_QL05 & ";申請國家：台灣P案所限" & TxtDate(0) & "-" & TxtDate(1)

                'Modified by Morgan 2013/6/5 修正案號重複問題,Ex:P-094958
                'strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np08||NP09) in (select np02,np03,np04,np05,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                   " NP02='P' and NP07 in (" & strCase & ",119) AND NP08 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
                   " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)<>'F' group by np02,np03,np04,np05)),PATENT,CUSTOMER,FAGENT" & _
                   " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) " & strTmp & " AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
                'Modified by Morgan 2013/6/26 +CU12,CU13
                
                'Modified by Morgan 2014/7/1  外層語法也要+AND NP06 IS NULL(ex: P-87410 103/10/08 重複)
                'Modified by Morgan 2017/12/11 台灣案不要排除F部門(P-97258原業務原誤掛外商人員造成期限沒催到)
                'Modified by Morgan 2018/10/3 NVL(PA22,'') -> PA22
                'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
                'Modified by Lydia 2019/12/16 改成共用句; Option2=>opt2 , text1(5)=>mdate1, text1(6)=>mdate2, strtmp => na01
                'strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                   " NP02='P' and NP07 in (" & strCase & ",119) AND NP08 BETWEEN " & TransDate(text1(5).Text, 2) & " AND " & TransDate(text1(6).Text, 2) & _
                   " AND NP06 IS NULL AND st01(+)=NP10" & IIf(Option2(1), " and substr(st03,1,1)<>'F'", "") & " group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                   " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) " & strTmp & " AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                'end 2014/7/1
                'end 2013/6/5
                strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                   " NP02='P' and NP07 in (" & strCase & ",119) AND NP08 BETWEEN mdate1 AND mdate2 " & _
                   " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                   " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                   
                'Added by Lydia 2019/08/30 指定客戶的案件
                If m_Except01 <> "" Then '和碩: 所有P及CFP案年費的通知時間均提早為法定期限前6個月
                    'Modified by Lydia 2022/08/12 +傳入客戶編號
                    SetDateCondition Text3.Text, False, 6, "X70017000"
                    If m_Date1 = "" Or m_Date2 = "" Then
                        MsgBox m_Except01 & "的通知時間有錯！", vbCritical
                        Exit Sub
                    End If
                    'Added by Lydia 2022/08/12 區分變數
                    m_Ex1Date1 = m_Date1: m_Ex1Date2 = m_Date2
                    m_Ex1FMPDate1 = m_FMPDate1: m_Ex1FMPDate2 = m_FMPDate2
                    m_Ex1DateTW1 = m_DateTW1: m_Ex1DateTW2 = m_DateTW2
                    'end 2022/08/12
                    'Modified by Lydia 2019/12/16 改成共用句; Option2(1)=>opt2 , m_Date1=>exdate1, m_Date2=>exdate2, strtmp => na01
                    'strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_Date1) & " AND " & DBDATE(m_Date2) & _
                       " AND NP06 IS NULL AND st01(+)=NP10" & IIf(Option2(1), " and substr(st03,1,1)<>'F'", "") & " group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) " & strTmp & " AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_Except01 & "',PA26)>0"
                    strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN exdate1 AND exdate2 " & _
                       " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_Except01 & "',PA26)>0"
                End If
                'Added by Lydia 2022/08/12 指定客戶的案件
                If m_ExceptB <> "" Then '信邦案之指定客戶(晚催期限，法定期限前一個月通知)
                    'Modified by Morgan 2023/7/13 改法限前2個月--文雄
                    'SetDateCondition Text3.Text, False, 1, "X39056000"
                    SetDateCondition Text3.Text, False, 2, "X39056000"
                    'end 2023/7/13
                    If m_Date1 = "" Or m_Date2 = "" Then
                        MsgBox m_ExceptB & "的通知時間有錯！", vbCritical
                        Exit Sub
                    End If
                    '區分變數
                    m_ExBDate1 = m_Date1: m_ExBDate2 = m_Date2
                    m_ExBFMPDate1 = m_FMPDate1: m_ExBFMPDate2 = m_FMPDate2
                    m_ExBDateTW1 = m_DateTW1: m_ExBDateTW2 = m_DateTW2

                    strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN exbdate1 AND exbdate2 " & _
                       " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_ExceptB & "',PA26)>0"
                End If
                'end 2022/08/12
                
                'Added by Lydia 2022/08/25指定客戶的案件
                If m_ExceptC <> "" Then '大亞案之指定客戶：法定期限前4個月通知
                    SetDateCondition Text3.Text, False, 4, "X60601000"
                    If m_Date1 = "" Or m_Date2 = "" Then
                        MsgBox m_ExceptC & "的通知時間有錯！", vbCritical
                        Exit Sub
                    End If
                    '區分變數
                    m_ExCDate1 = m_Date1: m_ExCDate2 = m_Date2
                    m_ExCFMPDate1 = m_FMPDate1: m_ExCFMPDate2 = m_FMPDate2
                    m_ExCDateTW1 = m_DateTW1: m_ExCDateTW2 = m_DateTW2

                    strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN excdate1 AND excdate2 " & _
                       " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_ExceptC & "',PA26)>0"
                End If
                'end 2022/08/25
                
                'Added by Lydia 2023/11/09 指定客戶的案件
                If m_ExceptD <> "" Then '康舒科技(X00497070)年費期限通知由三個月前改為一個月前
                    SetDateCondition Text3.Text, False, 1, "X00497070"
                    If m_Date1 = "" Or m_Date2 = "" Then
                        MsgBox m_ExceptD & "的通知時間有錯！", vbCritical
                        Exit Sub
                    End If
                    '區分變數
                    m_ExDDate1 = m_Date1: m_ExDDate2 = m_Date2
                    m_ExDFMPDate1 = m_FMPDate1: m_ExDFMPDate2 = m_FMPDate2
                    m_ExDDateTW1 = m_DateTW1: m_ExDDateTW2 = m_DateTW2

                    strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN exddate1 AND exddate2 " & _
                       " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_ExceptD & "',PA26)>0"
                End If
                'end 2023/11/09
                
                'Added by Lydia 2024/04/22 指定客戶的案件
                If m_ExceptF <> "" Then '立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                    SetDateCondition Text3.Text, False, 1, "X01506000"
                    If m_Date1 = "" Or m_Date2 = "" Then
                        MsgBox m_ExceptF & "的通知時間有錯！", vbCritical
                        Exit Sub
                    End If
                    '區分變數
                    m_ExFDate1 = m_Date1: m_ExFDate2 = m_Date2
                    m_ExFFMPDate1 = m_FMPDate1: m_ExFFMPDate2 = m_FMPDate2
                    m_ExFDateTW1 = m_DateTW1: m_ExFDateTW2 = m_DateTW2

                    strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN exfdate1 AND exfdate2 " & _
                       " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_ExceptF & "',PA26)>0"
                End If
                'end 2024/04/22
                'Added by Lydia 2024/07/23 指定客戶的案件
                If m_ExceptG <> "" Then 'X38120000/ X38120030碩天科技/寧遠縣碩寧電子，中國專利年費通知程序，提早期限前兩個月前通知
                    SetDateCondition Text3.Text, False, 2, "X38120000"
                    If m_Date1 = "" Or m_Date2 = "" Then
                        MsgBox m_ExceptG & "的通知時間有錯！", vbCritical
                        Exit Sub
                    End If
                    '區分變數
                    m_ExGDate1 = m_Date1: m_ExGDate2 = m_Date2
                    m_ExGFMPDate1 = m_FMPDate1: m_ExGFMPDate2 = m_FMPDate2
                    m_ExGDateTW1 = TxtDate(0): m_ExGDateTW2 = TxtDate(1) '台灣案保持3個月

                    strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN exgdate1 AND exgdate2 " & _
                       " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                       " AND INSTR('" & m_ExceptG & "',PA26)>0"
                End If
                'end 2024/07/23
                
                'Added by Lydia 2019/12/16 處理台灣案SQL
                strPartA = strExc(0)
                strPartA = Replace(strPartA, "mdate1", DBDATE(TxtDate(0)))   '台灣P案期限
                strPartA = Replace(strPartA, "mdate2", DBDATE(TxtDate(1)))   '台灣P案期限
                strPartA = Replace(strPartA, "opt2", "")
                strPartA = Replace(strPartA, "na01", "AND PA09='000' ")
                'Modified by Lydia 2022/08/12 m_DateTW 改為 m_Ex1DateTW
                strPartA = Replace(strPartA, "exdate1", DBDATE(m_Ex1DateTW1))  '和碩台灣P案期限
                strPartA = Replace(strPartA, "exdate2", DBDATE(m_Ex1DateTW2))  '和碩台灣P案期限
                'end 2019/12/16
                'Added by Lydia 2022/08/12
                strPartA = Replace(strPartA, "exbdate1", DBDATE(m_ExBDateTW1))  '信邦台灣P案期限
                strPartA = Replace(strPartA, "exbdate2", DBDATE(m_ExBDateTW2))  '信邦台灣P案期限
                'end 2022/08/12
                'Added by Lydia 2022/08/25
                strPartA = Replace(strPartA, "excdate1", DBDATE(m_ExCDateTW1))  '大亞(X60601000)台灣P案期限
                strPartA = Replace(strPartA, "excdate2", DBDATE(m_ExCDateTW2))  '大亞(X60601000)台灣P案期限
                'end 2022/08/25
                'Added by Lydia 2023/11/09
                strPartA = Replace(strPartA, "exddate1", DBDATE(m_ExDDateTW1))  '康舒科技(X00497070)台灣P案期限
                strPartA = Replace(strPartA, "exddate2", DBDATE(m_ExDDateTW2))  '康舒科技(X00497070)台灣P案期限
                'end 2023/11/09
                'Added by Lydia 2024/04/22 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                strPartA = Replace(strPartA, "exfdate1", DBDATE(m_ExFDateTW1))  '立德電子(X01506000)、江蘇領先(X01506010)台灣P案期限
                strPartA = Replace(strPartA, "exfdate2", DBDATE(m_ExFDateTW2))  '立德電子(X01506000)、江蘇領先(X01506010)台灣P案期限
                'end 2023/04/22
                'Added by Lydia 2024/07/23
                strPartA = Replace(strPartA, "exgdate1", DBDATE(m_ExGDateTW1))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 台灣P案期限
                strPartA = Replace(strPartA, "exgdate2", DBDATE(m_ExGDateTW2))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 台灣P案期限
                'end 2024/07/23
                
            'Modified by Lydia 2019/12/16 一併跑非大陸案
            'If Option2(1) Then '非台灣案
                    'pub_QL05 = pub_QL05 & ";" & Label1(7) & text1(7) & "-" & text1(8) 'Add By Sindy 2010/11/29
                    pub_QL05 = pub_QL05 & ";申請國家：非台灣P案所限" & TxtDate(2) & "-" & TxtDate(3)
                    pub_QL05 = pub_QL05 & ";申請國家：非台灣FMP案所限" & TxtDate(4) & "-" & TxtDate(5)
            'end 2019/12/16
                    'Modified by Morgan 2013/6/5 修正案號會重複問題
                    'strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np08||NP09) in (select np02,np03,np04,np05,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(text1(7).Text) & " AND " & DBDATE(text1(8).Text) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05)),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) " & strTmp & " AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
                    'Modified by Morgan 2014/7/1  外層語法也要+AND NP06 IS NULL
                    'Modified by Morgan 2015/7/8 Y52323 法限抓4個月 -->Get605InformPeriod4NonTwCase要同步修改 Morgan 2017/1/13
                    'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
                    'Modified by Lydia 2019/12/16 改變欄位
                    'strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(Text1(7).Text) & " AND " & DBDATE(Text1(8).Text) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) " & strTmp & " AND NVL(PA75,'Y')<>'Y52323000'" & _
                       " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                    'strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, Text1(7).Text) & " AND " & CompDate(1, 1, Text1(8).Text) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) " & strTmp & " AND PA75='Y52323000'" & _
                       " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                    strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(TxtDate(4)) & " AND " & DBDATE(TxtDate(5)) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                       " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                    strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(TxtDate(4))) & " AND " & CompDate(1, 1, DBDATE(TxtDate(5))) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                       " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                       " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                       
                    'Added by Lydia 2019/08/30 指定客戶的案件
                    If m_Except01 <> "" Then '和碩: 所有P及CFP案年費的通知時間均提早為法定期限前6個月
                        'Modified by Lydia 2022/08/12 改變數
                        'If m_FMPDate1 = "" Or m_FMPDate2 = "" Then
                        If m_Ex1FMPDate1 = "" Or m_Ex1FMPDate2 = "" Then
                            MsgBox m_Except01 & "的通知時間有錯！", vbCritical
                            Exit Sub
                        End If
                        'Modified by Lydia 2019/12/16 strTmp=>PA09<>'000'
                        'Modifed by Lydia 2022/08/12 m_FMPDate=>m_Ex1FMPDate
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_Ex1FMPDate1) & " AND " & DBDATE(m_Ex1FMPDate2) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_Except01 & "',PA26)>0"
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(m_Ex1FMPDate1)) & " AND " & CompDate(1, 1, DBDATE(m_Ex1FMPDate2)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_Except01 & "',PA26)>0"
                    End If
                    'Added by Lydia 2022/08/12 指定客戶的案件
                    If m_ExceptB <> "" Then ' '信邦案之指定客戶(晚催期限，法定期限前一個月通知)
                        If m_ExBFMPDate1 = "" Or m_ExBFMPDate2 = "" Then
                            MsgBox m_ExceptB & "的通知時間有錯！", vbCritical
                            Exit Sub
                        End If
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_ExBFMPDate1) & " AND " & DBDATE(m_ExBFMPDate2) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptB & "',PA26)>0"
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(m_ExBFMPDate1)) & " AND " & CompDate(1, 1, DBDATE(m_ExBFMPDate2)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptB & "',PA26)>0"
                    End If
                    'end 2022/08/12
                    'Added by Lydia 2022/08/25 指定客戶的案件
                    If m_ExceptC <> "" Then ' 大亞案之指定客戶：法定期限前4個月通知
                        If m_ExCFMPDate1 = "" Or m_ExCFMPDate2 = "" Then
                            MsgBox m_ExceptC & "的通知時間有錯！", vbCritical
                            Exit Sub
                        End If
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_ExCFMPDate1) & " AND " & DBDATE(m_ExCFMPDate2) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptC & "',PA26)>0"
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(m_ExCFMPDate1)) & " AND " & CompDate(1, 1, DBDATE(m_ExCFMPDate2)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptC & "',PA26)>0"
                    End If
                    'end 2022/08/25
                    'Added by Lydia 2023/11/09 指定客戶的案件
                    If m_ExceptD <> "" Then '康舒科技(X00497070)年費期限通知由三個月前改為一個月前
                        If m_ExDFMPDate1 = "" Or m_ExDFMPDate2 = "" Then
                            MsgBox m_ExceptD & "的通知時間有錯！", vbCritical
                            Exit Sub
                        End If
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_ExDFMPDate1) & " AND " & DBDATE(m_ExDFMPDate2) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptD & "',PA26)>0"
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(m_ExDFMPDate1)) & " AND " & CompDate(1, 1, DBDATE(m_ExDFMPDate2)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptD & "',PA26)>0"
                    End If
                    'end 2023/11/09
                    'Added by Lydia 2024/04/22 指定客戶的案件
                    If m_ExceptF <> "" Then '立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                        If m_ExFFMPDate1 = "" Or m_ExFFMPDate2 = "" Then
                            MsgBox m_ExceptF & "的通知時間有錯！", vbCritical
                            Exit Sub
                        End If
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_ExFFMPDate1) & " AND " & DBDATE(m_ExFFMPDate2) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptF & "',PA26)>0"
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(m_ExFFMPDate1)) & " AND " & CompDate(1, 1, DBDATE(m_ExFFMPDate2)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptF & "',PA26)>0"
                    End If
                    'end 2024/04/22
                    'Added by Lydia 2024/07/23 指定客戶的案件
                    If m_ExceptG <> "" Then 'X38120000/ X38120030碩天科技/寧遠縣碩寧電子，中國專利年費通知程序，提早期限前兩個月前通知
                        If m_ExGFMPDate1 = "" Or m_ExGFMPDate2 = "" Then
                            MsgBox m_ExceptG & "的通知時間有錯！", vbCritical
                            Exit Sub
                        End If
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(m_ExGFMPDate1) & " AND " & DBDATE(m_ExGFMPDate2) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptG & "',PA26)>0"
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(m_ExGFMPDate1)) & " AND " & CompDate(1, 1, DBDATE(m_ExGFMPDate2)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptG & "',PA26)>0"
                    End If
                    'end 2024/07/23
            'End If 'Remove by Lydia 2019/12/16 一併跑非大陸案
            
                'Added by Lydia 2019/12/16 處理非台灣案SQL
                strPartB = strExc(0)
                strPartB = Replace(strPartB, "mdate1", DBDATE(TxtDate(2))) '大陸P案期限
                strPartB = Replace(strPartB, "mdate2", DBDATE(TxtDate(3))) '大陸P案期限
                strPartB = Replace(strPartB, "opt2", " and substr(st03,1,1)<>'F' ") '限制非FMP案
                strPartB = Replace(strPartB, "na01", "AND PA09<>'000' ")    '限制非台灣案
                'Modified by Lydia 2022/08/12 m_Date 改為 m_Ex1Date
                strPartB = Replace(strPartB, "exdate1", DBDATE(m_Ex1Date1))  '和碩大陸P案期限
                strPartB = Replace(strPartB, "exdate2", DBDATE(m_Ex1Date2))  '和碩大陸P案期限
                'end 2019/12/16
                'Added by Lydia 2022/08/12
                strPartB = Replace(strPartB, "exbdate1", DBDATE(m_ExBDate1))  '信邦大陸P案期限
                strPartB = Replace(strPartB, "exbdate2", DBDATE(m_ExBDate2))  '信邦大陸P案期限
                'end 2022/08/12
                'Added by Lydia 2022/08/25
                strPartB = Replace(strPartB, "excdate1", DBDATE(m_ExCDate1))  '大亞(X60601000)大陸P案期限
                strPartB = Replace(strPartB, "excdate2", DBDATE(m_ExCDate2))  '大亞(X60601000)大陸P案期限
                'end 2022/08/25
                'Added by Lydia 2023/11/09
                strPartB = Replace(strPartB, "exddate1", DBDATE(m_ExDDate1))  '康舒科技(X00497070)大陸P案期限
                strPartB = Replace(strPartB, "exddate2", DBDATE(m_ExDDate2))  '康舒科技(X00497070)大陸P案期限
                'end 2023/11/09
                'Added by Lydia 2024/04/22 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                strPartB = Replace(strPartB, "exfdate1", DBDATE(m_ExFDate1))  '立德電子(X01506000)、江蘇領先(X01506010)大陸P案期限
                strPartB = Replace(strPartB, "exfdate2", DBDATE(m_ExFDate2))  '立德電子(X01506000)、江蘇領先(X01506010)大陸P案期限
                'end 2024/04/22
                'Added by Lydia 2024/07/23
                strPartB = Replace(strPartB, "exgdate1", DBDATE(m_ExGDate1))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 大陸P案期限
                strPartB = Replace(strPartB, "exgdate2", DBDATE(m_ExGDate2))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 大陸P案期限
                'end 2024/07/23

         '實體審查通知函
         Else
                'Modify by Morgan 2009/12/7 FMP案用法限條件且排除所限小於99/2/15的(已催過)
                'strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa08,pa10 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np08||NP09) in (select np02,np03,np04,np05,min(np08||NP09) FROM NEXTPROGRESS WHERE " & _
                   " NP02='P' and NP07=" & strCase & " AND NP08 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
                   " AND NP06 IS NULL group by np02,np03,np04,np05)),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                   "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL)" & strTmp & " AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & _
                   " ORDER BY PA09,PA01,PA02,PA03,PA04"
                'Modified by Lydia 2019/12/16
                'pub_QL05 = pub_QL05 & ";" & Label1(6) & Text1(5) & "-" & Text1(6) 'Add By Sindy 2010/11/29
                pub_QL05 = pub_QL05 & ";申請國家：台灣P案所限" & TxtDate(0) & "-" & TxtDate(1)
                'Modified by Morgan 2017/12/11 台灣案不要排除F部門
                'Modified by Morgan 2018/10/3 NVL(PA22,'') -> PA22
                'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
                'Modified by Lydia 2019/12/16 改成共用句; Option2=>opt2 , text1(5)=>mdate1, text1(6)=>mdate2, strtmp => na01
                'strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                   " NP02='P' and NP07=" & strCase & " AND NP08 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & _
                   " AND NP06 IS NULL AND st01(+)=NP10" & IIf(Option2(1), " and substr(st03,1,1)<>'F'", "") & " group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                   " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL)" & strTmp & " AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                   "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                   "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                   "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                   " NP02='P' and NP07=" & strCase & " AND NP08 BETWEEN mdate1 AND mdate2 " & _
                   " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                   " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                   "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                   
                    'Added by Lydia 2019/08/30 指定客戶的案件
                    If m_Except01 <> "" Then '和碩: 實審的通知時間則是提早為申請日＋１年，落在系統日期的1-10日或11-月底(20號)，與畫面不同的原因：是怕執行日期在10號採用畫面的日期可能會缺資料
                        'Added by Lydia 2022/08/12 如果只執行實審，補抓期限
                        If m_Ex1Date1 = "" Then
                            SetDateCondition Text3.Text, False, 6, "X70017000"
                            m_Ex1Date1 = m_Date1: m_Ex1Date2 = m_Date2
                            m_Ex1FMPDate1 = m_FMPDate1: m_Ex1FMPDate2 = m_FMPDate2
                            m_Ex1DateTW1 = m_DateTW1: m_Ex1DateTW2 = m_DateTW2
                        End If
                        'end 2022/08/12
                        If Val(Right(strSrvDate(1), 2)) <= 10 Then
                            strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "01" & " AND " & Mid(strSrvDate(1), 1, 6) & "10 "
                        Else
                            strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "11" & " AND " & Mid(strSrvDate(1), 1, 6) & "31 "
                        End If
                        'Modified by Lydia 2019/12/16 改成共用句; Option2=>opt2 , strtmp => na01
                        'strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND st01(+)=NP10" & IIf(Option2(1), " and substr(st03,1,1)<>'F'", "") & " group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL)" & strTmp & " AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_Except01 & "',PA26)>0 AND NVL(PA10,0)>0" & strExc(1)
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_Except01 & "',PA26)>0 AND NVL(PA10,0)>0" & strExc(1)
                    End If
                    'Added by Lydia 2022/08/12 指定客戶的案件
                    If m_ExceptB <> "" Then '信邦案之指定客戶：法定期限前一個月
                        '如果只執行實審，補抓期限
                        If m_ExBDate1 = "" Then
                           'Modified by Morgan 2023/7/13 改法限前2個月--文雄
                            'SetDateCondition Text3.Text, False, 1, "X39056000"
                            SetDateCondition Text3.Text, False, 2, "X39056000"
                            'end 2023/7/13
                            m_ExBDate1 = m_Date1: m_ExBDate2 = m_Date2
                            m_ExBFMPDate1 = m_FMPDate1: m_ExBFMPDate2 = m_FMPDate2
                            m_ExBDateTW1 = m_DateTW1: m_ExBDateTW2 = m_DateTW2
                        End If
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exbdate1 AND exbdate2 " & _
                           "AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptB & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2022/08/12
                    'Added by Lydia 2022/08/25 指定客戶的案件
                    If m_ExceptC <> "" Then '大亞案之指定客戶：法定期限前4個月通知
                        '如果只執行實審，補抓期限
                        If m_ExCDate1 = "" Then
                            SetDateCondition Text3.Text, False, 4, "X60601000"
                            m_ExCDate1 = m_Date1: m_ExCDate2 = m_Date2
                            m_ExCFMPDate1 = m_FMPDate1: m_ExCFMPDate2 = m_FMPDate2
                            m_ExCDateTW1 = m_DateTW1: m_ExCDateTW2 = m_DateTW2
                        End If
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN excdate1 AND excdate2 " & _
                           "AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptC & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2022/08/25
                    'Added by Lydia 2023/11/09 指定客戶的案件
                    If m_ExceptD <> "" Then '康舒科技(X00497070)年費期限通知由三個月前改為一個月前
                        '如果只執行實審，補抓期限
                        If m_ExDDate1 = "" Then
                            SetDateCondition Text3.Text, False, 1, "X00497070"
                            m_ExDDate1 = m_Date1: m_ExDDate2 = m_Date2
                            m_ExDFMPDate1 = m_FMPDate1: m_ExDFMPDate2 = m_FMPDate2
                            m_ExDDateTW1 = m_DateTW1: m_ExDDateTW2 = m_DateTW2
                        End If
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exddate1 AND exddate2 " & _
                           "AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptD & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2023/11/09
                    'Added by Lydia 2024/22 指定客戶的案件
                    If m_ExceptF <> "" Then '立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                        '如果只執行實審，補抓期限
                        If m_ExFDate1 = "" Then
                            SetDateCondition Text3.Text, False, 1, "X01506000"
                            m_ExFDate1 = m_Date1: m_ExFDate2 = m_Date2
                            m_ExFFMPDate1 = m_FMPDate1: m_ExFFMPDate2 = m_FMPDate2
                            m_ExFDateTW1 = m_DateTW1: m_ExFDateTW2 = m_DateTW2
                        End If
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exfdate1 AND exfdate2 " & _
                           "AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptF & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2024/04/22
                    'Added by Lydia 2024/07/23 指定客戶的案件
                    If m_ExceptG <> "" Then 'X38120000/ X38120030碩天科技/寧遠縣碩寧電子，中國專利年費通知程序，提早期限前兩個月前通知
                        '如果只執行實審，補抓期限
                        If m_ExGDate1 = "" Then
                            SetDateCondition Text3.Text, False, 2, "X38120000"
                            m_ExGDate1 = m_Date1: m_ExGDate2 = m_Date2
                            m_ExGFMPDate1 = m_FMPDate1: m_ExGFMPDate2 = m_FMPDate2
                            m_ExGDateTW1 = TxtDate(0): m_ExGDateTW2 = TxtDate(1) '台灣案保持3個月
                        End If
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exgdate1 AND exgdate2 " & _
                           "AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & m_ExceptG & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2024/07/23
            
                    'Added by Lydia 2019/12/16 處理台灣案SQL
                    strPartA = strExc(0)
                    strPartA = Replace(strPartA, "mdate1", DBDATE(TxtDate(0)))
                    strPartA = Replace(strPartA, "mdate2", DBDATE(TxtDate(1)))
                    strPartA = Replace(strPartA, "opt2", "")
                    strPartA = Replace(strPartA, "na01", "AND PA09='000' ")
                    'end 2019/12/16
                    'Added by Lydia 2022/08/12
                    strPartA = Replace(strPartA, "exbdate1", DBDATE(m_ExBDateTW1))  '信邦台灣P案期限
                    strPartA = Replace(strPartA, "exbdate2", DBDATE(m_ExBDateTW2))  '信邦台灣P案期限
                    'end 2022/08/12
                    'Added by Lydia 2022/08/25
                    strPartA = Replace(strPartA, "excdate1", DBDATE(m_ExCDateTW1))  '大亞(X60601000)台灣P案期限
                    strPartA = Replace(strPartA, "excdate2", DBDATE(m_ExCDateTW2))  '大亞(X60601000)台灣P案期限
                    'end 2022/08/25
                    'Added by Lydia 2023/11/09
                    strPartA = Replace(strPartA, "exddate1", DBDATE(m_ExDDateTW1))  '康舒科技(X00497070)台灣P案期限
                    strPartA = Replace(strPartA, "exddate2", DBDATE(m_ExDDateTW2))  '康舒科技(X00497070)台灣P案期限
                    'end 2023/11/09
                    'Added by Lydia 2024/04/22 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                    strPartA = Replace(strPartA, "exfdate1", DBDATE(m_ExFDateTW1))  '立德電子(X01506000)、江蘇領先(X01506010)台灣P案期限
                    strPartA = Replace(strPartA, "exfdate2", DBDATE(m_ExFDateTW2))  '立德電子(X01506000)、江蘇領先(X01506010)台灣P案期限
                    'end 2024/04/22
                    'Added by Lydia 2024/07/23
                    strPartA = Replace(strPartA, "exgdate1", DBDATE(m_ExGDateTW1))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 台灣P案期限
                    strPartA = Replace(strPartA, "exgdate2", DBDATE(m_ExGDateTW2))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 台灣P案期限
                    'end 2024/07/23
                     
                'Modified by Lydia 2019/12/16 一併跑非大陸案
                'If Option2(1) Then '非台灣案
                    'pub_QL05 = pub_QL05 & ";" & Label1(7) & Text1(7) & "-" & Text1(8) 'Add By Sindy 2010/11/29
                    pub_QL05 = pub_QL05 & ";申請國家：非台灣P案所限" & TxtDate(2) & "-" & TxtDate(3)
                    pub_QL05 = pub_QL05 & ";申請國家：非台灣FMP案所限" & TxtDate(4) & "-" & TxtDate(5)
                'end 2019/12/16
                    'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
                    'Modified by Lydia 2019/12/16 改變欄位
                    'strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07=" & strCase & " AND NP09 BETWEEN " & DBDATE(Text1(7).Text) & " AND " & DBDATE(Text1(8).Text) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                       "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL)" & strTmp & " AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                    strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                       "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                       "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                       "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                       " NP02='P' and NP07=" & strCase & " AND NP09 BETWEEN " & DBDATE(TxtDate(4)) & " AND " & DBDATE(TxtDate(5)) & _
                       " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                       "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                       "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
                    
                    'Added by Lydia 2019/08/30 指定客戶的案件
                    If m_Except01 <> "" Then '和碩: 實審的通知時間則是提早為申請日＋１年，落在系統日期的1-10日或11-月底(20號)
                        If Val(Right(strSrvDate(1), 2)) <= 10 Then
                            strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "01" & " AND " & Mid(strSrvDate(1), 1, 6) & "10 "
                        Else
                            strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "11" & " AND " & Mid(strSrvDate(1), 1, 6) & "31 "
                        End If
                        'Modified by Lydia 2019/12/6 strTmp => PA09<>'000'
                         strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                            "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                            "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                            "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                            " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND st01(+)=NP10 " & _
                            "and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                            "),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                            "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                            " AND INSTR('" & m_Except01 & "',PA26)>0 AND NVL(PA10,0)>0" & strExc(1)
                    End If
                    'Added by Lydia 2022/08/12 指定客戶的案件
                    If m_ExceptB <> "" Then  '信邦案之指定客戶：法定期限前一個月
                         strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                            "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                            "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                            "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                            " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exbdate1 AND exbdate2 " & _
                            "AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                            "),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                            "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                            " AND INSTR('" & m_ExceptB & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2022/08/12
                    'Added by Lydia 2022/08/25 指定客戶的案件
                    If m_ExceptC <> "" Then  '大亞案(X60601000)：法定期限前4個月
                         strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                            "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                            "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                            "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                            " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN excdate1 AND excdate2 " & _
                            "AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                            "),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                            "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                            " AND INSTR('" & m_ExceptC & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2022/08/25
                    'Added by Lydia 2023/11/09 指定客戶的案件
                    If m_ExceptD <> "" Then  '康舒科技(X00497070)年費期限通知由三個月前改為一個月前
                         strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                            "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                            "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                            "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                            " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exddate1 AND exddate2 " & _
                            "AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                            "),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                            "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                            " AND INSTR('" & m_ExceptD & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2023/11/09
                    'Added by Lydia 2024/04/22 指定客戶的案件
                    If m_ExceptF <> "" Then  '康舒科技(X00497070)年費期限通知由三個月前改為一個月前
                         strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                            "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                            "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                            "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                            " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exfdate1 AND exfdate2 " & _
                            "AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                            "),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                            "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                            " AND INSTR('" & m_ExceptF & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2024/04/22
                   'Added by Lydia 2024/07/23 指定客戶的案件
                    If m_ExceptG <> "" Then  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子，中國專利年費通知程序，提早期限前兩個月前通知
                         strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                            "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                            "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                            "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                            " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL AND NP09 BETWEEN exgdate1 AND exgdate2 " & _
                            "AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                            "),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                            "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                            " AND INSTR('" & m_ExceptG & "',PA26)>0 AND NVL(PA10,0)>0"
                    End If
                    'end 2024/07/23
                
                    'Added by Lydia 2019/12/16 處理非台灣案SQL
                    strPartB = strExc(0)
                    strPartB = Replace(strPartB, "mdate1", DBDATE(TxtDate(2)))
                    strPartB = Replace(strPartB, "mdate2", DBDATE(TxtDate(3)))
                    strPartB = Replace(strPartB, "opt2", " and substr(st03,1,1)<>'F' ")
                    strPartB = Replace(strPartB, "na01", "AND PA09<>'000' ")
                    'end 2019/12/16
                    'Added by Lydia 2022/08/12
                    strPartB = Replace(strPartB, "exbdate1", DBDATE(m_ExBDateTW1))  '信邦台灣P案期限
                    strPartB = Replace(strPartB, "exbdate2", DBDATE(m_ExBDateTW2))  '信邦台灣P案期限
                    'end 2022/08/12
                    'Added by Lydia 2022/08/25
                    strPartB = Replace(strPartB, "excdate1", DBDATE(m_ExCDateTW1))  '大亞(X60601000)台灣P案期限
                    strPartB = Replace(strPartB, "excdate2", DBDATE(m_ExCDateTW2))  '大亞(X60601000)信邦台灣P案期限
                    'end 2022/08/25
                    'Added by Lydia 2023/11/09
                    strPartB = Replace(strPartB, "exddate1", DBDATE(m_ExDDateTW1))  '康舒科技(X00497070)台灣P案期限
                    strPartB = Replace(strPartB, "exddate2", DBDATE(m_ExDDateTW2))  '康舒科技(X00497070)台灣P案期限
                    'end 2023/11/09
                    'Added by Lydia 2024/04/22 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
                    strPartB = Replace(strPartB, "exfdate1", DBDATE(m_ExFDateTW1))  '立德電子(X01506000)、江蘇領先(X01506010)台灣P案期限
                    strPartB = Replace(strPartB, "exfdate2", DBDATE(m_ExFDateTW2))  '立德電子(X01506000)、江蘇領先(X01506010)台灣P案期限
                    'end 2024/04/22
                    'Added by Lydia 2024/07/23
                    strPartB = Replace(strPartB, "exgdate1", DBDATE(m_ExGDateTW1))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 台灣P案期限
                    strPartB = Replace(strPartB, "exgdate2", DBDATE(m_ExGDateTW2))  'X38120000/ X38120030碩天科技/寧遠縣碩寧電子 台灣P案期限
                    'end 2024/07/23

                'End If 'Remove by Lydia 2019/12/16 一併跑非大陸案
         End If
         
         'Added by  Lydia 2019/12/16 組合SQL
         strExc(0) = strPartA & " Union " & strPartB
         
         'Modified by Morgan 2013/6/26
         'strExc(0) = strExc(0) & " ORDER BY PA09,PA01,PA02,PA03,PA04"
         'Memo by Morgan 2015/9/1 整批列印定稿有另外控制列印順序(同接洽人一起,另年費逾期補繳通知也有)
         'Added by Morgan 2018/10/3 配合調整非FMP非台灣的年費與實審期限的所限,要剔除已整批催過的期限(過渡期避免重複催)
         'Modified by Morgan 2019/9/16 只跑期限表時不必剔除已整批催過的期限
         'Removed by Morgan 2021/8/16 過渡期早過，取消此檢查以避免漏催(曾發生改不續辦的原期限管制下次期限而沒催到)--8/16 有跟玲玲確認
         'If Check2.Value = 0 Then
         '   strExc(0) = "SELECT * FROM (" & strExc(0) & ") WHERE NOT EXISTS(select * from caseprogress y,letterprogress z" & _
         '      " where y.cp43=np01 and y.cp30=np22 and y.cp10='1913' and z.lp01(+)=y.cp09 and z.lp32='Y')"
         'End If
         'end 2021/8/16
         'end 2019/9/16
         'end 2018/10/3
         
         'Modified by Morgan 2025/1/16 +PID,排序語法改放變數(後面要用)
         'strExc(0) = strExc(0) & " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
         strExc(0) = "SELECT X.*,''PID FROM (" & strExc(0) & ") X"
         strSort = " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
         strExc(0) = strExc(0) & strSort
         'end 2025/1/16
         
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         
         'Added by Morgan 2025/1/16
         If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
            Combo1.Tag = ""
            Set rsQuery = PUB_CreateRecordset(rsTemp1, , , 300, Me.Name, mSeqNo)
            With rsQuery
               .MoveFirst
               Do While Not .EOF
                  .Fields("PID") = PUB_GetPHandler(.Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04"))
                  .MoveNext
               Loop
               .UpdateBatch
               
               stVTBX = "select R001 as " & .Fields(0).Name
               For intI = 2 To .Fields.Count
                  stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
               Next
               stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
            End With
            strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "'" & strSort
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strSql)
            Combo1.Tag = Combo1
         End If
         'end 2025/1/16
            
         If intI = 1 Then
            With rsTemp1
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/29
            Do While Not .EOF
               Erase strTxt
               strReceiveNo = .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)
               'Add by Morgan 2006/5/15
               strNP07 = "" & .Fields("NP07")
               strNP08 = "" & .Fields("NP08")
               strNP09 = "" & .Fields("NP09")
               'end 2006/5/15
               strNP23 = "" & .Fields("NP23") 'Add by Morgan 2010/1/18
               
               'Add by Morgan 2009/10/7
               m_PA09 = "" & .Fields("pa09").Value
               m_PA46 = "" & .Fields("pa46").Value
               m_strPA08 = "" & .Fields("PA08")
               m_strPA10 = "" & .Fields("PA10")
               m_PA26 = "" & .Fields("pa26") 'Added by Morgan 2014/6/12
               
               'Add by Morgan 2009/12/7
               strPA75 = "" & .Fields("PA75")
               strTmp2 = ""
               If .Fields("FMP") = "Y" Then
                  m_bolFMP = True
                  iCopy = 1
               Else
                  m_bolFMP = False
                  iCopy = 0
               End If
               
               'Modify by Morgan 2006/4/13 加R04030306欄位存類別
               'Modify by Morgan 2006/5/15 119進入國家階段 類別放3
               'Modified by Morgan 2020/8/6 指定欄位名稱
               'Modified by Morgan 2024/11/6 615補償期年費 類別放4
               If strNP07 = "119" Then
                  cnnConnection.Execute "Insert Into R040303(R04030301,R04030302,R04030303,R04030304,R04030305,ID,R04030306) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "','" & .Fields(3).Value & "','" & .Fields(4).Value & "','" & strUserNum & "','3')"
               ElseIf strNP07 = "615" Then
                  cnnConnection.Execute "Insert Into R040303(R04030301,R04030302,R04030303,R04030304,R04030305,ID,R04030306) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "','" & .Fields(3).Value & "','" & .Fields(4).Value & "','" & strUserNum & "','4')"
               Else
                  cnnConnection.Execute "Insert Into R040303(R04030301,R04030302,R04030303,R04030304,R04030305,ID,R04030306) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "','" & .Fields(3).Value & "','" & .Fields(4).Value & "','" & strUserNum & "','" & m_Select & "')"
               End If
            
               'Add by Morgan 2005/5/16
               m_CurCP(1) = .Fields(0): m_CurCP(2) = .Fields(1): m_CurCP(3) = .Fields(2): m_CurCP(4) = .Fields(3)
               m_NP22 = .Fields("np22"): m_iDiscount = 0
               
               'Added by Morgan 2020/8/6
               '年費通知加檢查是否核駁期限已逾期超過3月未辦的案件
               If m_Select = "1" Then
                  If ChkIsOverLimited(.Fields("NP01")) = True Then
                     strSql = "update R040303 set R04030307='X' where R04030301='" & m_CurCP(1) & "' and R04030302='" & m_CurCP(2) & "' and R04030303='" & m_CurCP(3) & "' and R04030304='" & m_CurCP(4) & "' and R04030306='1'"
                     cnnConnection.Execute strSql, intI
                     GoTo NoLetter
                  End If
               End If
               'end 2020/8/6
               
               'Add by Morgan 2004/12/13 控制只印地址條or期限表
               If Check1.Value = vbChecked Or Check2.Value = vbChecked Then GoTo NoLetter
               
               'Add by Morgan 2006/2/17
               '繳年費通知若已上閉卷,但下一程序仍有年費期限且進度檔內有1604 or 1606 or 1907 or 413案件性質時不出定稿只印清單--玲玲
               'Modify by Morgan 2006/2/20
               '不管是否閉卷都不出定稿,也不印期限表(語法控制),只印消滅清單--郭
               'If .Fields("PA57") = "Y" Then
               'Add by Morgan 2007/10/24
               '閉卷的都不印通知信但要印在清單上
               If .Fields("PA57") = "Y" Then
                  GoTo NoLetter
               Else
               'end 2007/10/24
                  'Modified by Lydia 2019/12/16 改判斷PA09
                  'If Option2(0).Value = True Then 'Add by Morgan 2010/11/30 非台灣沒閉卷的都要通知--991126請作單
                  If "" & .Fields("PA09") Then
                     If CheckCPExists(m_CurCP) = True Then GoTo NoLetter
                  End If
               End If
               'End If
               '2006/2/20 end
               
               '收文號
               strCP09 = ""
               'Modify by Morgan 2006/5/24 不用重抓
'               'Modify by Morgan 2005/11/14 收文號改抓NP01
'               '年費
'               If m_Select = "1" Then
'                  StrSQLa = "Select NP01 From NextProgress Where " & ChgNextProgress(strReceiveNo) & " And NP06 is null and NP07 in (" & 年費 & "," & 延展費 & ",119)"
'               '實審
'               Else
'                  StrSQLa = "Select NP01 From NextProgress Where " & ChgNextProgress(strReceiveNo) & " And NP06 is null and NP07='" & 實體審查 & "'"
'               End If
'               rsA.CursorLocation = adUseClient
'               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'               If rsA.RecordCount > 0 Then
'                   strCP09 = "" & rsA.Fields(0).Value
'               End If
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
               strCP09 = "" & .Fields("NP01")
               'END 2006/5/25
               
               'Added by Morgan 2013/8/7
               'Modified by Morgan 2014/6/12 +申請國家,配合定稿轉pdf要有收文號改先新增進度
               'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
               'Modified by Morgan 2016/11/8 +傳是否大宗發文(pbolBulk=True)
               If PUB_AddCP1913(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04"), .Fields("NP08"), .Fields("NP09"), .Fields("NP01"), .Fields("NP22"), m_PA09, m_PA26, m_LD18, strPA75, , , True) = False Then
                  MsgBox "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
                  Exit Sub
               End If
               'end 2013/8/7
               
               'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
               If m_Select = "1" Or m_Select = "2" Then
                  '實體審查
                  If m_Select <> "1" Then
                  'Modified by Lydia 2015/01/07 採共用模組
'                     strSql = "SELECT YF06+YF07 FROM PATENTYEARFEE WHERE YF01='" & .Fields(6) & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='416' AND YF05=1"
'                     If rsTemp10.State <> adStateClosed Then rsTemp10.Close
'                     rsTemp10.CursorLocation = adUseClient
'                     rsTemp10.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                     strFee = ""
'                     If rsTemp10.RecordCount > 0 Then
'                         strFee = "" & rsTemp10.Fields(0).Value
'                     End If
'                     If rsTemp10.State <> adStateClosed Then rsTemp10.Close
'                     Set rsTemp10 = Nothing
                     strFee = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, "416", "1", "1", "1")
                     '申請國家為大陸,是否為PCT案件為"Y",則定稿之案件性質為06,否則為05
                     'Modify by Morgan 2006/5/18 加PCT
                     'If .Fields(6) = 大陸國家代號 Then
                     'Modify by Morgan 2009/7/9 +澳門044
                     'If .Fields(6) = 大陸國家代號 Or .Fields(6) = "056" Then
                     If .Fields(6) = 大陸國家代號 Or .Fields(6) = "056" Or .Fields(6) = "044" Then
                        
                        If .Fields(6) = "056" Then
                           strTmp = "14"
                        'Add by Morgan 2009/7/9  澳門
                        ElseIf .Fields(6) = "044" Then
                           strTmp = "22"
                        Else
                           strTmp = IIf(.Fields("PA46").Value = "Y", "06", "05")
                        End If
                        '刪除定稿暫存資料
                        EndLetter ET01, strCP09, strTmp, strUserNum
                        '新增定稿暫存資料
                        strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                        strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','法定期限'," & CNULL(.Fields(5)) & ")"
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
                           
                        'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                        strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "416", "1", "1", "1")
                        strPoint = Round(Val(strPoint) / 1000, 1)
                        PUB_UpdateLP2930 m_LD18, strFee, strPoint
                        'end 2015/8/28
                        
                        'Add by Morgan 2005/11/16
                        strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','下一程序','416')"
                        'Add by Morgan 2009/7/9
                        strExc(0) = Pub_Get416Period(.Fields("PA08"), .Fields("PA09"))
                        strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','提實審期限','" & strExc(0) & "')"

                        If Not ClsLawExecSQL(5, strTxt) Then
                            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                        End If
                        NowPrint strCP09, ET01, strTmp, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                        'Add by Morgan 2009/12/7
                        If m_bolFMP Then
                           strUserNum = strFMPNum
                           strTmp2 = "51"
                           EndLetter ET01, strCP09, strTmp2, strUserNum
                           strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                           strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                           If m_PA46 = "Y" Then
                              strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & _
                                 "','PCT案','♀')"
                           Else
                              strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & _
                                 "','非PCT案','♀')"
                           End If
                           If Not ClsLawExecSQL(3, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strTmp2, False, strUserNum
                           strUserNum = strUser1Num
                        End If
                        
                     '申請國家為台灣,定稿之案件性質為07
                     ElseIf .Fields(6) = 台灣國家代號 Then
                        
                        '大-->台 催實體審查定稿定稿 20080916 ADD BY TONI
                        If PUB_CheckCuNation(rsTemp1.Fields("pa26"), rsTemp1.Fields("pa01"), rsTemp1.Fields("pa02"), rsTemp1.Fields("pa03"), rsTemp1.Fields("pa04")) = "1" Then
                              strET03 = "20"
                              '刪除定稿暫存資料
                              EndLetter ET01, strCP09, strET03, strUserNum
                              '新增定稿暫存資料
                              strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                              strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(.Fields(5)) & ")"
                              strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"

                              strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','416')"
                              If Not ClsLawExecSQL(4, strTxt) Then
                                  MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                              End If
                              NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                              'END BY TONI 20080916
                        Else
                           strET03 = "07"
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(.Fields(5)) & ")"
                           strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"

                           strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','416')"
                           If Not ClsLawExecSQL(4, strTxt) Then
                               MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                        End If
                     End If
                  'Add by Morgan 2006/5/15
                  ElseIf strNP07 = "119" Then
                     strET03 = "13"
                     '刪除定稿暫存資料
                     EndLetter ET01, strCP09, strET03, strUserNum
                     '新增定稿暫存資料
                     strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(strNP08) & ")"
                     strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
                     strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序'," & CNULL(strNP07) & ")"
                    
                     If Not ClsLawExecSQL(3, strTxt) Then
                         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                     End If
                     NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                  
                  'Added by Morgan 2024/11/6
                  ElseIf strNP07 = "615" Then
                     strET03 = "23"
                     '刪除定稿暫存資料
                     EndLetter ET01, strCP09, strET03, strUserNum
                     '新增定稿暫存資料
                     strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(strNP08) & ")"
                     strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
                     strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序'," & CNULL(strNP07) & ")"
                     strExc(1) = ""
                     If PUB_GetCNExtDays(m_CurCP(), , intI) Then
                        If intI > 0 Then strExc(1) = intI
                     End If
                     strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','補償天數'," & CNULL(strExc(1)) & ")"
                     strFee = PUB_GetCN615Fee(m_CurCP())
                     strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
                     If Not ClsLawExecSQL(5, strTxt) Then
                         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                     End If
                     NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                     
                     'Added by Morgan 2025/3/10
                     If m_bolFMP Then
                        strUserNum = strFMPNum
                        m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&615"
                        strTmp2 = "53"
                        EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                        strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                        strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strFee & "')"
                        strExc(1) = PUB_GetUSXRate
                        strExc(2) = ""
                        If Val(strExc(1)) <> 0 Then
                           strExc(2) = Fix(strFee / Val(strExc(1)))
                        End If
                        strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                                 
                        If Not ClsLawExecSQL(4, strTxt) Then
                            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                        End If
                        NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                        strUserNum = strUser1Num
                     End If
                     'end 2025/3/10
                     
                  'end 2024/11/6
                  '年費
                  Else
                     '本所期限是否已逾期且未超過7個月
                     blnSitu1 = False
                     StrSQLa = "Select * From Nextprogress Where " & ChgNextProgress(strReceiveNo) & " And NP07=" & 年費 & " And NP06='N' and NP08>0 AND NP09>0"
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        Do While Not rsA.EOF
                           If IsNull(rsA("NP08").Value) = False Then
                              If rsA("NP08").Value < strSrvDate(1) And DateDiff("m", ChangeWStringToWDateString(rsA("NP08").Value), ChangeWStringToWDateString(.Fields(4))) <= 7 Then
                                 strOldNP09 = "" & rsA("NP09").Value
                                 blnSitu1 = True
                                 Exit Do
                              End If
                           End If
                           rsA.MoveNext
                         Loop
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     
                     '補繳期限(有不續辦,該本所期限已過,期限差7個月內)
                     If blnSitu1 = True Then
                        If "" & .Fields(6).Value = "020" Then
                           'Added by Morgan 2015/8/28 非台灣信函進度要存報價(逾期原定稿只有點數)
                           strPA72NextYear = getPA72NextYear(m_CurCP(1), m_CurCP(2), m_CurCP(3), m_CurCP(4), , , strPA25)
                           If strPA72NextYear <> "" Then
                              strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                              strPoint = Round(Val(strPoint) / 1000, 1)
                           Else
                              strPoint = ""
                           End If
                           PUB_UpdateLP2930 m_LD18, "", strPoint
                           'end 2015/8/28
                           
                           strET03 = "09"
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           'Add by Morgan 2009/10/14
                           '98/10/1以後的一案兩請案,新型年費定稿加提醒
                           If m_PA09 = "020" And m_strPA08 = "2" And Val(m_strPA10) >= 20091001 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 'Added by Morgan 2012/8/30
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 'end 2012/8/30
                                 End If
                              End If
                           End If
                           
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','半年前法定期限'," & CNULL(strOldNP09) & ")"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           strPA72Year = GetNowNP09(strReceiveNo)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/5
                           'Removed by Morgan 2023/6/5 取消,起算日相同,不會發生
                           'If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                           '   If strPA25 < CompDate(1, 6, strPA72Year) Then
                           '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           '         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                           '      ii = ii + 1
                           '   End If
                           'End If
                           'end 2023/6/5
                           'end 2023/6/5
                           
                           'Add by Morgan 2005/11/16
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           ii = ii + 1
                           
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                               MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                           
                        ElseIf "" & .Fields(6).Value = "000" Then
                           strET03 = "08"
                        
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           'Added by Morgan 2012/9/21
                           '102新法一案兩請案,新型年費定稿加提醒
                           If m_strPA08 = "2" And Val(m_strPA10) >= 20130101 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 '未核准
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 End If
                              End If
                           End If
                           'end 2012/9/21
                           
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','半年前法定期限'," & CNULL(strOldNP09) & ")"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           strPA72Year = GetNowNP09(strReceiveNo)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/5
                           If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                              If strPA25 < CompDate(1, 6, strPA72Year) Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                                 ii = ii + 1
                              End If
                           End If
                           'end 2023/6/5
                           
                           'Remove by Morgan 2011/7/6 定稿已改用共用文字欄位
                           'Add by Morgan 2004/6/4 年費收費標準
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                           '   " SELECT '" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','年費收費標準',FTM05 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='21' AND FTM03='000' AND FTM04='02'"
                           'ii = ii + 1
                           'end 2011/7/6
                                    
                           'Add by Morgan 2005/11/16
                           'Modified by Morgan 2022/9/1 不可傳入606否則回覆單會多帶
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605')"
                           'end 2022/9/1
                           ii = ii + 1
                           
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                           
                        End If
                            
                     '正常期限
                     Else
                        '大陸
                        '大陸案若系統日>法定期限, 則定稿之案件性質為03
                        'Modify by Morgan 2008/3/20 +澳門(044)
                        'If CheckStr(.Fields(6)) = "020" Then
                        If CheckStr(.Fields(6)) = "020" Or CheckStr(.Fields(6)) = "044" Then
                           '取得下次繳費年度
                           strPA72NextYear = getPA72NextYear(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, , m_bFirstYear)
                           If CheckStr(.Fields(6)) = "044" Then
                              'Add by Morgan 2008/5/7 +繳第一次年費(無繳費記錄)
                              If m_bFirstYear = True Then
                                 strET03 = "18"
                              Else
                                 strET03 = "17"
                              End If
                           Else
                              'Modify By Sindy 2009/05/22 改定稿格式
                              'strET03 = IIf(strSrvDate(1) > "" & .Fields(5).Value, "03", "02")
                              strET03 = IIf(strSrvDate(1) > "" & .Fields(5).Value, "03", "21")
                           End If
                           
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           
                           'Add by Morgan 2010/1/18 FMP約定期限
                           If m_bolFMP Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','約定期限','" & strNP23 & "')"
                              ii = ii + 1
                           End If
                           
                           '計算該年年費屆滿日期strPA72Year
                           strPA72Year = getPA72Year(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strPA25)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/5
                           'Removed by Morgan 2023/6/5 取消,起算日相同,不會發生
                           'If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                           '   If strPA25 < CompDate(1, 6, strPA72Year) Then
                           '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           '         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                           '      ii = ii + 1
                           '   End If
                           'End If
                           'end 2023/6/5
                           'end 2023/6/5
                           
                           strNextYearFee = ""
                           If strPA72NextYear <> "" Then
                           'Modified by Lydia 2015/01/07 採共用模組
                           '   strNextYearFee = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, "Y00000001", "605", strPA72NextYear, strPA72NextYear)
                              strNextYearFee = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用','" & strNextYearFee & "')"
                              ii = ii + 1
                           End If
                           
                           'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                           If strPA72NextYear <> "" Then
                              strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                              strPoint = Round(Val(strPoint) / 1000, 1)
                           Else
                              strPoint = ""
                           End If
                           PUB_UpdateLP2930 m_LD18, strNextYearFee, strPoint
                           'end 2015/8/28
                              
                           'Add by Morgan 2005/11/16
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           ii = ii + 1
                           
                           'Add by Morgan 2009/10/7
                           '98/10/1以後的一案兩請案,新型年費定稿加提醒
                           '大陸一案兩請申請日輸入時必須互相檢查發明及新型之申請日是否為同一天,若不是,則show訊息告知user。
                           bolDualCaseUtility = False 'Added by Morgan 2017/9/20
                           If m_PA09 = "020" And m_strPA08 = "2" And Val(m_strPA10) >= 20091001 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14,pa11,PA77" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 'Added by Morgan 2017/9/20
                                 strInventionCaseNo = "" & RsTemp("C1")
                                 strInventionPA11 = "" & RsTemp("pa11")
                                 strInventionPA77 = "" & RsTemp("pa77")
                                 'end 2017/9/20
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    bolDualCaseUtility = True 'Added by Morgan 2017/9/20
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                       "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 'Added by Morgan 2012/8/30
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    bolDualCaseUtility = True 'Added by Morgan 2017/9/20
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 'end 2012/8/30
                                 End If
                              End If
                           End If
                           
                           
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                           'Add by Morgan 2009/12/7
                           If m_bolFMP Then
                              strUserNum = strFMPNum
                              
                              'Modified by Morgan 2014/8/20 FMP案有年費代理人,改傳本所案號+案件性質
                              'm_FMP_ET02 = strCP09
                              m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&605"
                              'end 2014/8/20
                              
                              'Removed by Morgan 2022/9/20 定稿已合併
                              '付款後辦案
                              'If CU72FA39("", strPA75) Then
                              '   strTmp2 = "53"
                              'Else
                              'end 2022/9/20
                              
                                 'Added by Morgan 2022/9/30
                                 If CheckStr(.Fields(6)) = "044" Then
                                    strTmp2 = "54"
                                 Else
                                 'end 2022/9/30
                                    strTmp2 = "52"
                                 End If
                                 
                              'End If 'Removed by Morgan 2022/9/20
                              
                              EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                              strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                              strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                              If m_PA46 = "Y" Then
                                 strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & _
                                    "','PCT案','♀')"
                              Else
                                 strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & _
                                    "','非PCT案','♀')"
                              End If
                              strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','下次年費年度','" & strPA72NextYear & "')"
                              strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strNextYearFee & "')"
                              strExc(1) = PUB_GetUSXRate
                              strExc(2) = ""
                              If Val(strExc(1)) <> 0 Then
                                 strExc(2) = Fix(strNextYearFee / Val(strExc(1)))
                              End If
                              strTxt(6) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                             
                              ii = 6
                              'Added by Morgan 2017/9/20
                              If bolDualCaseUtility = True Then
                                 strTxt(7) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','一案兩請新型案要印','♀')"
                                 strTxt(8) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案本所案號','" & strInventionCaseNo & "')"
                                 strTxt(9) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案申請號','" & ChgSQL(strInventionPA11) & "')"
                                 strTxt(10) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案彼所案號','" & ChgSQL(strInventionPA77) & "')"
                                 ii = 10
                              End If
                              'end 2017/9/20
                                    
                              If Not ClsLawExecSQL(ii, strTxt) Then
                                 MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                              End If
                              NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                              strUserNum = strUser1Num
                           End If
                        'Add by Morgan 2004/5/14
                        '香港
                        ElseIf CheckStr(.Fields(6)) = "013" Then
                           '取得已繳費年度及專利種類
                           strPA72NextYear = getNextPayYear(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strPA72Year, strPA25)
                           
                           Select Case m_strPA08
                              Case "1" '標準專利
                                 stSitu = "12"
                              Case "2" '短期專利
                                 stSitu = "11"
                              Case "3" '外觀設計
                                 stSitu = "10"
                           End Select
                           
                           If strNP07 = 維持費 Then stSitu = "02" 'Added by Morgan 2012/10/23
                           
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, stSitu, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/5
                           'Removed by Morgan 2023/6/5 取消,起算日相同,不會發生
                           'If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                           '   If strPA25 < CompDate(1, 6, strPA72Year) Then
                           '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           '         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                           '      ii = ii + 1
                           '   End If
                           'End If
                           'end 2023/6/5
                           'end 2023/6/5
                           
                           If strPA72NextYear <> "" Then
                           'Modified by Lydia 2015/01/07 採共用模組
'                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                                 "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','費用','" & Val(PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, "Y00000001", strNP07, strPA72NextYear, strPA72NextYear)) & "')"
                              strNextYearFee = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, strNP07, strPA72NextYear, strPA72NextYear, "1")
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','費用','" & Val(strNextYearFee) & "')"
                              ii = ii + 1
                           End If
                           
                           'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                           If strPA72NextYear <> "" Then
                              strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, strNP07, strPA72NextYear, strPA72NextYear, "1")
                              strPoint = Round(Val(strPoint) / 1000, 1)
                           Else
                              strPoint = ""
                           End If
                           PUB_UpdateLP2930 m_LD18, strNextYearFee, strPoint
                           'end 2015/8/28
                           
                           'Add by Morgan 2005/11/16
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','下一程序','605,607')"
                           ii = ii + 1
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, stSitu, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                           
                           'Add by Morgan 2022/9/30
                            If m_bolFMP Then
                               strUserNum = strFMPNum
                               m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&605"
                               strTmp2 = "54"
                               EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                               strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                               strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                               strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','下次年費年度','" & strPA72NextYear & "')"
                               strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strNextYearFee & "')"
                               strExc(1) = PUB_GetUSXRate
                               strExc(2) = ""
                               If Val(strExc(1)) <> 0 Then
                                  strExc(2) = Fix(strNextYearFee / Val(strExc(1)))
                               End If
                               strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                               ii = 5
                               If Not ClsLawExecSQL(ii, strTxt) Then
                                  MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                               End If
                               NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                               strUserNum = strUser1Num
                            End If
                            'end 2022/9/30
                            
                        '台灣
                        ElseIf CheckStr(.Fields(6)) = "000" Then
                           iPlusFee = 0 'Added by Morgan 2013/1/8
                           
                           '大-->台 催年費定稿 20090916 ADD BY TONI
                           If PUB_CheckCuNation(rsTemp1.Fields("pa26"), rsTemp1.Fields("pa01"), rsTemp1.Fields("pa02"), rsTemp1.Fields("pa03"), rsTemp1.Fields("pa04")) = "1" Then
                              strET03 = "19"
                              'Added by Morgan 2013/1/8 專利處大對台年費服務費+500 --郭雅娟 (113.7.12 接洽單也同步增加此規則 frm090801_new)
                              strExc(1) = PUB_GetStaffST15(PUB_GetAKindSalesNo(rsTemp1.Fields("PA01"), rsTemp1.Fields("PA02"), rsTemp1.Fields("PA03"), rsTemp1.Fields("PA04")), "1")
                              If Left(strExc(1), 2) = "P1" Then
                                 iPlusFee = 500
                              End If
                              'end 2013/1/8
                           Else
                              'Modify by Morgan 2008/1/7 一率改用新定稿
                              'strET03 = "01"
                              strET03 = "15"
                           End If
                           'END BY TONI
                           
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           '計算該年年費屆滿日期strPA72Year
                           strPA72Year = getPA72Year(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strPA25)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/1
                           If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                              If strPA25 < CompDate(1, 6, strPA72Year) Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                                 ii = ii + 1
                              End If
                           End If
                           'end 2023/6/1
                           
                           'Added by Morgan 2012/9/21
                           If m_strPA08 = "2" And Val(m_strPA10) >= 20130101 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 '未核准
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 End If
                              End If
                           End If
                           
                           If DBDATE(strPA72Year) >= 20120101 Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','102新法不印','♀')"
                              ii = ii + 1
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','102新法要印','♀')"
                              ii = ii + 1
                           End If
                           
                           '取得下次繳費年度
                           strPA72NextYear = getPA72NextYear(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strMaxFeeYear)
                           If strPA72NextYear <> "" Then
                              'Modify by Morgan 2007/11/6
                              'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              '   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用','" & Val(PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, "Y00000001", "605", strPA72NextYear, strPA72NextYear)) & "')"
                              'ii = ii + 1
                              
                              '服務費,規費
                              'Modified by Lydia 2015/01/07 採共用模組
'                              strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & .Fields("PA09").Value & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & strPA72NextYear
'                              intI = 1
'                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                              If intI = 1 Then
'                                 strExc(1) = "" & RsTemp("YF06")
'                                 strExc(2) = "" & RsTemp("YF07")
'                              Else
'                                 strExc(1) = ""
'                                 strExc(2) = ""
'                              End If
                              strExc(0) = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1", strExc(1), strExc(2))
                              If strExc(0) = "0" Then strExc(1) = "": strExc(2) = ""
                              
                              If strExc(1) <> "" Then
                                 strExc(1) = Val(strExc(1)) + iPlusFee 'Added by Morgan 2013/1/8
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','服務費','" & strExc(1) & "')"
                                 ii = ii + 1
                              End If
                              
                              '年費是否可減免
                              If PUB_GetCaseDiscStat(.Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)) = "Y" Then
                                 bolDiscount = True
                              Else
                                 bolDiscount = False
                              End If
                           
                              If Val(strExc(2)) > 0 Then
                                 '減免
                                 If Val(strPA72NextYear) < 7 Then
                                    If bolDiscount = True Then
                                       If Val(strPA72NextYear) < 4 Then
                                          strExc(2) = Val(strExc(2)) - 800
                                       Else
                                          strExc(2) = Val(strExc(2)) - 1200
                                       End If
                                    End If
                                 End If
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','規費','" & strExc(2) & "')"
                                 ii = ii + 1
                              End If
   
                              strExc(3) = Val(strExc(1)) + Val(strExc(2))
                              If Val(strExc(3)) > 0 Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','費用','" & strExc(3) & "')"
                                 ii = ii + 1
                              End If
                              'end 2007/11/6

                              'Add by Morgan 2007/11/20 下兩年的費用也要印
                              strExc(5) = strExc(2) '規費累計
                              strExc(6) = strExc(3) '費用累計
                              'Added by Lydia 2024/08/15
                              Dim strBaseYear As String
                              strBaseYear = strPA72NextYear
                              'end 2024/08/15
                              For jj = 1 To 2
                                 strPA72NextYear = Val(strPA72NextYear) + 1
                                 If Val(strPA72NextYear) <= Val(strMaxFeeYear) Then
                                 'Modified by Lydia 2015/01/07 採共用模組
'                                    strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & .Fields("PA09").Value & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & strPA72NextYear
'                                    intI = 1
'                                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                                    If intI = 1 Then
'                                       strExc(2) = "" & RsTemp("YF07")
'                                    Else
'                                       strExc(2) = ""
'                                    End If
                                    strExc(0) = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1", , strExc(2))
                                    If strExc(0) = "0" Then strExc(2) = ""
                                    
                                    If Val(strExc(2)) > 0 Then
                                       '減免
                                       If Val(strPA72NextYear) < 7 Then
                                          If bolDiscount = True Then
                                             If Val(strPA72NextYear) < 4 Then
                                                strExc(2) = Val(strExc(2)) - 800
                                             Else
                                                strExc(2) = Val(strExc(2)) - 1200
                                             End If
                                          End If
                                       End If
                                       strExc(5) = Val(strExc(5)) + Val(strExc(2))
                                       
                                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                          "','規費" & jj & "','" & strExc(5) & "')"
                                       ii = ii + 1
                                       '費用累計
                                       'Added by Lydia 2024/08/15 重抓服務費; ex.訊強電子 (惠州 )X41570060, P-117332
                                       strExc(0) = " select '1' as ord1, ys07 from patentyearspec where ys01='" & m_PA09 & "' and ys03='" & m_PA26 & "' and ys02='" & m_strPA08 & "' and ys04='605' and ys05='" & strBaseYear & "' and ys06='" & strPA72NextYear & "' " & _
                                                   " union select '2' as ord1, yf06 as ys07 from patentyearfee where yf01='" & m_PA09 & "' and yf03='" & m_PA26 & "' and yf02='" & m_strPA08 & "' and yf04='605' and yf05='" & strPA72NextYear & "' " & _
                                                   " order by 1"
                                       intI = 1
                                       Set rsTemp10 = ClsLawReadRstMsg(intI, strExc(0))
                                       If intI = 1 Then
                                          strExc(1) = Val("" & rsTemp10.Fields("ys07"))
                                       End If
                                       'end 2024/08/15
                                       strExc(6) = Val(strExc(1)) + Val(strExc(5))
                                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                          "','費用" & jj & "','" & strExc(6) & "')"
                                       ii = ii + 1
                                    End If
                                 Else
                                    Exit For
                                 End If
                              Next
                              
                              'end 2007/11/20
                           End If
                           
                           'Remove by Morgan 2011/7/6 定稿已改用共用文字欄位
                           ''Add by Morgan 2004/6/4 年費收費標準
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                           '   " SELECT '" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','年費收費標準',FTM05 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='21' AND FTM03='000' AND FTM04='02'"
                           'ii = ii + 1
                           'end 2011/7/6
                           
'Remove by Morgan 2011/8/3 新定稿已不再使用
'                           'Add by Morgan 2005/5/17 辦理減免退費提醒
'                           If PUB_GetCaseDiscStat(strReceiveNo) = "Y" Then
'                              If PUB_CheckYearFeeReturn(m_CurCP, False, m_iDiscount, m_iYear1, m_iYear2) = True Then
'                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                                    "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','減免退費起迄年','" & IIf(m_iYear1 = m_iYear1, m_iYear1, m_iYear1 & "年至第" & m_iYear2) & "')"
'                                 ii = ii + 1
'                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                                    "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','減免退費金額','" & m_iDiscount & "')"
'                                 ii = ii + 1
'                              End If
'                           End If
'                           '2005/5/16 end
                           
                           'Add by Morgan 2011/7/6
                           If PUB_ChkRefund(m_CurCP, m_lngRefund) = True Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','未退金額','" & m_lngRefund & "')"
                              ii = ii + 1
                           End If
                     
                           
                           'Add by Morgan 2005/11/16
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           ii = ii + 1
                           '2005/11/16 END
   
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                           'Add by Morgan 2005/5/17 台灣新增年費通知紀錄
                           Call UpdateAI
                           
                        End If
                     End If
                  End If
               End If
               
               '列印接洽結案單
               pub_AddressListSN = pub_AddressListSN + 1
               'Modify by Morgan 2005/5/18
               'PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, CheckStr(.Fields(12).Value), CheckStr(.Fields(0).Value), CheckStr(.Fields(1).Value), CheckStr(.Fields(2).Value), CheckStr(.Fields(3).Value)
               If m_iDiscount > 0 Then
                  PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, CheckStr(.Fields(12).Value), CheckStr(.Fields(0).Value), CheckStr(.Fields(1).Value), CheckStr(.Fields(2).Value), CheckStr(.Fields(3).Value), "1"
               Else
                  PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, CheckStr(.Fields(12).Value), CheckStr(.Fields(0).Value), CheckStr(.Fields(1).Value), CheckStr(.Fields(2).Value), CheckStr(.Fields(3).Value)
               End If
               
'Add by Morgan 2004/12/13 跳過定稿
NoLetter:

               '只印期限表時不印地址條
               If Check2.Value = vbUnchecked Then
               
'Remove by Morgan 2008/8/13 改開窗定稿
'                  '儲存客戶編號
'                  If IsEmptyText("" & .Fields(8)) = False Then
'                      ReDim Preserve m_CustList(m_CustListCount + 1)
'                      ReDim Preserve m_CP(m_CustListCount + 1)
'                      m_CustList(m_CustListCount) = .Fields(8)
'                      m_CP(m_CustListCount) = CheckStr(.Fields(0).Value) & "-" & CheckStr(.Fields(1).Value) & "-" & CheckStr(.Fields(2).Value) & "-" & CheckStr(.Fields(3).Value)
'                      m_CustListCount = m_CustListCount + 1
'                  End If

'Remove by Morgan 2009/12/22 改國外部自行列印
'                  'Add by Morgan 2009/12/7 FMP要印地址條
'                  If m_bolFMP Then
'                     pub_AddressListSN = pub_AddressListSN + 1
'                     PUB_AddNewAddressList strUserNum, m_CurCP(1), m_CurCP(2), m_CurCP(3), m_CurCP(4), "" & pub_AddressListSN, "0", 實體審查
'                  End If
'end 2009/12/22

               End If
               .MoveNext
            Loop
            End With
            bolPrint = True
         Else
            InsertQueryLog (0) 'Add By Sindy 2010/11/29
            MsgBox "無符合條件之" & stMsg & "資料可列印 !", vbInformation
         End If
      End If
      
   Next idx
  
   m_LD18 = "" 'Added by Morgan 2015/5/20
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
   'Add by Morgan 2007/9/4 主張國外優先權(未收文主張國外優先權,自撤)
   'Modified by Lydia 2019/12/16
   'If Option2(0).Value = True And chkKind(0).Value = 1 Then
   '   pub_QL05 = pub_QL05 & ";" & Label1(1) & Option2(0).Caption 'Add By Sindy 2010/11/29
   '   pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(0)  'Add By Sindy 2010/11/29
   '   pub_QL05 = pub_QL05 & ";" & Label1(6) & text1(5) & "-" & text1(6) 'Add By Sindy 2010/11/29
   If chkKind(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";申請國家：台灣"
      pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(0)
      pub_QL05 = pub_QL05 & ";台灣P案所限" & TxtDate(0) & "-" & TxtDate(1)
   'end 2019/12/16
      stMsg = "【主張國外優先權】"
     'Add by Lydia 2015/01/27 +fmp寰華控制sql (m_selarea)
      Call ChangeSel(2) '將SQL改為對應PA
      'Modified by Lydia 2019/12/16 改變欄位
      'strExc(0) = "select pa01,pa02,pa03,pa04,pa26" & _
         " from patent a where pa01='P' and pa09='000'" & _
         " and to_char(add_months(to_date(pa10,'yyyymmdd'),9),'yyyymmdd') between " & DBDATE(text1(5)) & " and " & DBDATE(text1(6)) & _
         " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
         " and cp10 in ('106','413') and cp57 is null)" & _
         " and not ( exists(select * from casemap,patent b where cm05=a.pa01 and cm06=a.pa02 and cm07=a.pa03 and cm08=a.pa04" & _
         " and b.pa01(+)=cm01 and b.pa02(+)=cm02 and b.pa03(+)=cm03 and b.pa04(+)=cm04 and b.pa10<a.pa10)" & _
         " or exists(select * from casemap,patent b where cm01=a.pa01 and cm02=a.pa02 and cm03=a.pa03 and cm04=a.pa04" & _
         " and b.pa01(+)=cm05 and b.pa02(+)=cm06 and b.pa03(+)=cm07 and b.pa04(+)=cm08 and b.pa10<a.pa10))" & m_SelArea
      'Modified by Morgan 2025/1/16 +PID
      strExc(0) = "select pa01,pa02,pa03,pa04,pa26,'' PID" & _
         " from patent a where pa01='P' and pa09='000'" & _
         " and to_char(add_months(to_date(pa10,'yyyymmdd'),9),'yyyymmdd') between " & DBDATE(TxtDate(0)) & " and " & DBDATE(TxtDate(1)) & _
         " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
         " and cp10 in ('106','413') and cp57 is null)" & _
         " and not ( exists(select * from casemap,patent b where cm05=a.pa01 and cm06=a.pa02 and cm07=a.pa03 and cm08=a.pa04" & _
         " and b.pa01(+)=cm01 and b.pa02(+)=cm02 and b.pa03(+)=cm03 and b.pa04(+)=cm04 and b.pa10<a.pa10)" & _
         " or exists(select * from casemap,patent b where cm01=a.pa01 and cm02=a.pa02 and cm03=a.pa03 and cm04=a.pa04" & _
         " and b.pa01(+)=cm05 and b.pa02(+)=cm06 and b.pa03(+)=cm07 and b.pa04(+)=cm08 and b.pa10<a.pa10))" & m_SelArea
      intI = 1
      Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
      
      'Added by Morgan 2025/1/16
      If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
         Combo1.Tag = ""
         Set rsQuery = PUB_CreateRecordset(rsTemp1, , , 300, Me.Name, mSeqNo)
         With rsQuery
            .MoveFirst
            Do While Not .EOF
               .Fields("PID") = PUB_GetPHandler(.Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04"))
               .MoveNext
            Loop
            .UpdateBatch
            
            stVTBX = "select R001 as " & .Fields(0).Name
            For intI = 2 To .Fields.Count
               stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
            Next
            stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
         End With
         strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strSql)
         Combo1.Tag = Combo1
      End If
      'end 2025/1/16
      
      If intI = 1 Then
         With rsTemp1
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/29
         Do While Not .EOF
            '只印地址條or期限表
            If Check1.Value = vbChecked Or Check2.Value = vbChecked Then GoTo NoLetter1
         
            If Me.Check1.Value = vbUnchecked Then
               strCP09 = .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
               strTmp = "16"
               NowPrint strCP09, ET01, strTmp, False, strUserNum, 0, , , , , , , , , , , , m_LD18
            End If
NoLetter1:
            '只印期限表時不印地址條
            If Check2.Value = vbUnchecked Then
'Remove by Morgan 2008/8/13 改開窗定稿
'               '儲存客戶編號
'               If "" & .Fields("PA26") <> "" Then
'                  ReDim Preserve m_CustList(m_CustListCount + 1)
'                  ReDim Preserve m_CP(m_CustListCount + 1)
'                  m_CustList(m_CustListCount) = .Fields("PA26")
'                  m_CP(m_CustListCount) = .Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04")
'                  m_CustListCount = m_CustListCount + 1
'               End If
            End If
            .MoveNext
         Loop
         End With
         bolPrint1 = True
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/11/29
         MsgBox "無符合條件之待通知" & stMsg & "資料可列印 !", vbInformation
      End If
   End If
   
   If bolPrint = True Then
      '只印地址條or期限表時不印結案單
      If Me.Check1.Value = vbUnchecked And Me.Check2.Value = vbUnchecked Then
          '列印接洽結案單
         PUB_PrintCaseCloseSheet strUserNum
      End If
      '只印地址條時不印期限表
      If Me.Check1.Value = vbUnchecked Then
          MsgBox "請更換紙張，按確定後開始列印期限表!", vbOKOnly + vbInformation, "列印期限表"
          '列印繳年費/實體審查期限表
          Process1
          '列印專利權消滅清單
          'Modify by Morgan 2006/8/16 改一起跑所以不必再判斷
          'If m_Select = "1" Then Process2
          Process2
      End If
   End If

   If bolPrint = True Or bolPrint1 = True Then
      '只印期限表時不印地址條
      If Me.Check2.Value = vbUnchecked Then
      
'Remove by Morgan 2008/7/18 改開窗定稿紙不必再印地址條
'         If m_CustListCount > 0 Then
'            If MsgBox("按確定後開始列印地址條!", vbOKCancel + vbInformation, "列印地址條") = vbOK Then
'               PrintAddress
'            End If
'         End If

'Remove by Morgan 2009/12/22 改國外部自行列印
'         'Add by Morgan 2009/12/7 FMP要印地址條
'         '列印地址條
'         PUB_PrintAddressList strUserNum, cmbPrinter.Text
'         PUB_RestorePrinter strPrint
'         '刪除地址條列表資料
'         PUB_DeleteAddressList strUserNum

      End If
      MsgBox "列印結束 !", vbInformation
   End If
End Sub

'Add by Morgan 2006/4/3
Private Function FormCheck() As Boolean
   'Add by Morgan 2007/9/5
   'Added by Lydia 2015/04/20 +TW-SUPA
   If chkKind(0).Value + chkKind(1).Value + chkKind(2).Value + chkKind(3).Value = 0 Then
      MsgBox "請勾選通知函類別!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   
'Modify by Morgan 2009/12/7
'   If Me.Text1(5).Text = "" Then
'      MsgBox "請輸入下次繳費起日!!!", vbExclamation + vbOKOnly
'      Me.Text1(5).SetFocus
'      Text1_GotFocus 5
'      Exit Function
'   End If
'   If Me.Text1(6).Text = "" Then
'      MsgBox "請輸入下次繳費迄日!!!", vbExclamation + vbOKOnly
'      Me.Text1(6).SetFocus
'      Text1_GotFocus 6
'      Exit Function
'   End If
'   If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
'      Me.Text1(5).SetFocus
'      Text1_GotFocus 5
'      Exit Function
'   End If
'   If PUB_CheckKeyInDate(Me.Text1(6)) = -1 Then
'      Me.Text1(6).SetFocus
'      Text1_GotFocus 6
'      Exit Function
'   End If
'   If Me.Text1(5).Text <> "" And Me.Text1(6).Text <> "" Then
'      If Val(Me.Text1(5).Text) > Val(Me.Text1(6).Text) Then
'         MsgBox "下次繳費日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
'         blnClkSure = True
'         Me.Text1(5).SetFocus
'         Text1_GotFocus 5
'         Exit Function
'      End If
'   End If
   If Text3 = "" Then
      MsgBox "請輸入執行日期年月!!!", vbExclamation + vbOKOnly
      Text3.SetFocus
      Exit Function
   Else
      'Modified by Lydia 2019/12/16
'      If text1(5) = "" Then
'         MsgBox "無法計算P案本所期限起日!!!", vbExclamation + vbOKOnly
'         Exit Function
'      ElseIf text1(6) = "" Then
'         MsgBox "無法計算P案本所期限迄日!!!", vbExclamation + vbOKOnly
'         Exit Function
'      ElseIf Option2(1) Then
'         If text1(7) = "" Then
'            MsgBox "無法計算FMP案法定期限起日!!!", vbExclamation + vbOKOnly
'            Exit Function
'         ElseIf text1(8) = "" Then
'            MsgBox "無法計算FMP案法定期限迄日!!!", vbExclamation + vbOKOnly
'            Exit Function
'         End If
'      End If
      If TxtDate(0) & TxtDate(1) = "" Then
          MsgBox "無法計算台灣P案本所期限區間!!!", vbExclamation + vbOKOnly
      ElseIf TxtDate(2) & TxtDate(3) = "" Then
          MsgBox "無法計算非台灣P案本所期限區間!!!", vbExclamation + vbOKOnly
      ElseIf TxtDate(4) & TxtDate(5) = "" Then
          MsgBox "無法計算非台灣FMP案法定期限區間!!!", vbExclamation + vbOKOnly
      End If
      'end 2019/12/16
   End If
'end 2009/12/7

   'Add by Morgan 2007/5/10
   'Remove by Lydia 2019/1/16
   'If Option2(0).Value = False And Option2(1).Value = False Then
   '   MsgBox "請選擇申請國家！", vbExclamation
   '   Exit Function
   'End If
   ''end 2007/5/10
   'end 2019/12/16
   FormCheck = True
End Function

'Add by Morgan 2006/4/3
Private Sub Process3()
   Dim strCase As String
   Dim strTmp As String, strTmp2 As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTxt(1 To 20) As String
   Dim rsTemp10 As New ADODB.Recordset
   Dim strFee As String, strPoint As String
   Dim ii As Integer, jj As Integer
   Dim strPA72NextYear As String
   Dim strPA72Year As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim strCP09 As String '收文號
   Dim Prn As Printer
   Dim blnSitu1 As Boolean '下一程序檔是否有本所案號+案件性質本所期限在前半年且是否續辦為N的資料
   Dim strOldNP09 As String '半年前法定期限
   Dim stSitu As String '定稿處理狀況
   Dim idx As Integer
   Dim bolPrint As Boolean
   Dim strNP07 As String, strNP08 As String, strNP09 As String 'Add by Morgan 2006/5/15
   Dim strET03 As String
   Dim strMaxFeeYear As String '最大可繳費年度
   Dim bolDiscount As Boolean '是否可減免
   Dim strNextYearFee As String '下次繳費金額
   Dim strPA75 As String
   Dim iCopy As Integer
   Dim strNP23 As String 'Add by Morgan 2010//1/18 約定期限
   Dim iPlusFee As Integer 'Added by Morgan 2013/1/8 服務費外加金額(目前專利處大對台年費+500)
   Dim strNP22 As String 'Added by Morgan 2013/8/7
   Dim bolDualCaseUtility As Boolean, strInventionCaseNo As String, strInventionPA11 As String, strInventionPA77 As String 'Added by Morgan 2017/9/20 是否一案兩請新型案,發明案本所案號,發明案申請號,發明案彼所案號
   Dim strPA25 As String 'Added by Morgan 2023/6/1
   
   bolPrint = False
   blnClkSure = False
   
   '刪除暫存資料
   cnnConnection.Execute "Delete From R040303 Where ID='" & strUserNum & "'"
   '刪除接洽結案單暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   
'Remove by Morgan 2008/8/13 改開窗定稿
'   '清除
'   ClearCustList
'   '搜尋預設印表機
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   
   
   For idx = 1 To 2
      If chkKind(idx).Value = 1 Then
         m_Select = idx
         Exit For
      End If
   Next
   
    strTmp = Text1(1) & Text1(2)
    If Text1(3).Text = "" Then
        strTmp = strTmp & "0"
    Else
        strTmp = strTmp & Text1(3).Text
    End If
    If Text1(4).Text = "" Then
        strTmp = strTmp & "00"
    Else
        strTmp = strTmp & Text1(4).Text
    End If
    'Add by Morgan 2005/5/16
    m_CurCP(1) = Text1(1): m_CurCP(2) = Text1(2)
    m_CurCP(3) = Right("0" & Text1(3), 1)
    m_CurCP(4) = Right("00" & Text1(4), 2)
    'm_NP22 = tmpNP22 'Removed by Morgan 2024/11/7 下面抓到資料再設定,也比較正確
    m_iDiscount = 0
    pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & m_CurCP(1) & "-" & m_CurCP(2) & "-" & m_CurCP(3) & "-" & m_CurCP(4) 'Add By Sindy 2010/11/29
    
    strReceiveNo = strTmp
    
    If Len(Text2.Text) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & Text2 'Add By Sindy 2010/11/29
    End If
    
    'Add by Lydia 2015/01/27 +fmp寰華控制sql (m_selarea)
     Call ChangeSel(1) '將SQL改為對應NP
     
    '取得收文號
    strCP09 = ""
    'Modify by Morgan 2005/11/14 收文號改抓NP01
    '年費
    If m_Select = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(4) 'Add By Sindy 2010/11/29
      'Modify by Morgan 2006/5/15 加119進入國家階段
      'StrSQLa = "Select NP01 From NextProgress Where " & ChgNextProgress(strReceiveNo) & " And NP06 is null and NP07 in (" & 年費 & "," & 延展費 & ")"
      'Modified by Morgan 2012/10/23 +香港維持費
      'Modified by Morgan 2023/6/1 +pa25
      'Modified by Morgan 2024/11/1 +615補償期年費
      StrSQLa = "Select NP01,NP07,NP08,NP09,st03,pa75,NP23,NP22,PA25 From NextProgress,staff,patent Where " & ChgNextProgress(strReceiveNo) & " And NP06 is null and NP07 in (" & 年費 & "," & 維持費 & "," & 延展費 & ",119,615) AND st01(+)=NP10 and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05" & m_SelArea
    '實審
    Else
      pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(3) 'Add By Sindy 2010/11/29
      'Modified by Morgan 2023/6/1 +pa25
      StrSQLa = "Select NP01,NP07,NP08,NP09,st03,pa75,NP23,NP22,PA25 From NextProgress,staff,patent Where " & ChgNextProgress(strReceiveNo) & " And NP06 is null and NP07='" & 實體審查 & "' AND st01(+)=NP10 and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05" & m_SelArea
    End If
    
    'Added by Morgan 2024/11/7
    If Text2 <> "" Then
      StrSQLa = StrSQLa & " and np08=" & DBDATE(Text2)
    End If
    'end 2024/11/7
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/11/29
        strCP09 = "" & rsA.Fields(0).Value
        strNP07 = "" & rsA.Fields(1)
        strNP08 = "" & rsA.Fields(2)
        strNP09 = "" & rsA.Fields(3)
        strNP23 = "" & rsA.Fields("NP23")
        strNP22 = "" & rsA.Fields("NP22")
        strPA25 = "" & rsA.Fields("PA25") 'Added by Morgan 2023/6/1
        m_NP22 = strNP22 'Added by Morgan 2024/11/7
        
        'Add by Morgan 2009/12/8
        strPA75 = "" & rsA.Fields("pa75")
        If Left(rsA.Fields("st03"), 1) = "F" Then
            m_bolFMP = True
            iCopy = 1
        Else
            m_bolFMP = False
            iCopy = 0
        End If
        
        
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   
     'Add by Morgan 2007/4/25
     If strCP09 = "" Then
         InsertQueryLog (0) 'Add By Sindy 2010/11/29
         MsgBox "無可列印資料！"
         Exit Sub
     End If
    '若為年費通知函, 則不論是否閉卷資料都要出來
'    If m_Select = "1" And strNP07 <> "119" Then
'        strExc(0) = "SELECT COUNT(*) FROM PATENT WHERE " & ChgPatent(strReceiveNo)
'    '實體審查通知函
'    Else
'        strExc(0) = "SELECT COUNT(*) FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND (PA57<>'Y' OR PA57 IS NULL)"
'    End If

   'add by toni 20080916
    If m_Select = "1" And strNP07 <> "119" Then
        strExc(0) = "SELECT * FROM PATENT WHERE " & ChgPatent(strReceiveNo) & ""
    '實體審查通知函
    Else
        strExc(0) = "SELECT * FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND (PA57<>'Y' OR PA57 IS NULL)"
    End If
   'end by Toni

    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       '只列印地址條
       If Check1.Value = vbChecked Or Check2.Value = vbChecked Then GoTo NoLetter
               
      'Added by Morgan 2013/8/7
      'Modified by Morgan 2014/6/12 +申請國家,配合定稿轉pdf要有收文號改先新增進度
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      If PUB_AddCP1913(m_CurCP(1), m_CurCP(2), m_CurCP(3), m_CurCP(4), strNP08, strNP09, strCP09, strNP22, m_PA09, m_PA26, m_LD18, strPA75) = False Then
         MsgBox "新增進度檔【通知期限】失敗！", vbCritical
      'Added by Morgan 2016/7/28
      '單筆自動檢核
      Else
         strSql = "update letterprogress set lp27='QPGMR',lp28=sysdate,lp32=null where lp01='" & m_LD18 & "'"
         cnnConnection.Execute strSql, intI
      'end 2016/7/28
      End If
      'end 2013/8/7
      
      '大陸/PCT實體審查
      If m_Select <> "1" Then
           'Add By Cheng 2002/03/05
           'Modified by Lydia 2015/01/07 採共用模組
'           strSql = "SELECT YF06+YF07 FROM PATENTYEARFEE WHERE YF01='" & m_PA09 & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='416' AND YF05=1"
'           If rsTemp10.State <> adStateClosed Then rsTemp10.Close
'           rsTemp10.CursorLocation = adUseClient
'           rsTemp10.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'           strFee = ""
'           If rsTemp10.RecordCount > 0 Then
'               strFee = "" & rsTemp10.Fields(0).Value
'           End If
'           If rsTemp10.State <> adStateClosed Then rsTemp10.Close
'           Set rsTemp10 = Nothing
            strFee = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, "416", "1", "1", "1")
            
           '若申請國家為大陸時, 是否為PCT案件為"Y",則定稿之案件性質為06,否則為05
           
           'Modify by Morgan 2009/7/9 +澳門044
           'If m_PA09 = "020" Or m_PA09 = "056" Then
           If m_PA09 = "020" Or m_PA09 = "056" Or m_PA09 = "044" Then
               If m_PA09 = "056" Then
                  strTmp = "14"
               ElseIf m_PA09 = "044" Then
                  strTmp = "22"
               Else
                  If m_PA46 = "Y" Then
                     strTmp = "06"
                  Else
                     strTmp = "05"
                  End If
               End If
               EndLetter ET01, strCP09, strTmp, strUserNum
               strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
               strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','費用'," & strFee & ")"
               
               'Added by Morgan 2015/8/28 非台灣信函進度要存報價
               strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "416", "1", "1", "1")
               strPoint = Round(Val(strPoint) / 1000, 1)
               PUB_UpdateLP2930 m_LD18, strFee, strPoint
               'end 2015/8/28
                        
               'Add by Morgan 2005/11/16
               strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','下一程序','416')"
               'Add by Morgan 2006/6/15
               strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
                               
               'Add by Morgan 2009/7/9
               strExc(0) = Pub_Get416Period(RsTemp("PA08"), RsTemp("PA09"))
               strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','提實審期限','" & strExc(0) & "')"
               
               If Not ClsLawExecSQL(5, strTxt) Then
                   MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
               End If
               NowPrint strCP09, ET01, strTmp, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
               
               'Add by Morgan 2009/12/7
               If m_bolFMP Then
                  strUserNum = strFMPNum
                  strTmp2 = "51"
                  EndLetter ET01, strCP09, strTmp2, strUserNum
                  strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                  strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                  If m_PA46 = "Y" Then
                     strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & _
                        "','PCT案','♀')"
                  Else
                     strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & _
                        "','非PCT案','♀')"
                  End If
                  If Not ClsLawExecSQL(3, strTxt) Then
                     MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                  End If
                  NowPrint strCP09, ET01, strTmp2, False, strUserNum
                  strUserNum = strUser1Num
               End If
               
           Else
               '大--台 催實體審查定稿 20080916 add by Toni
               If PUB_CheckCuNation(RsTemp.Fields("PA26"), RsTemp.Fields("PA01"), RsTemp.Fields("PA02"), RsTemp.Fields("PA03"), RsTemp.Fields("PA04")) = "1" Then
                  strET03 = "20"
                  '刪除定稿暫存資料
                  EndLetter ET01, strCP09, strET03, strUserNum
                 '新增定稿暫存資料
                 strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                 strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
                 'Add by Morgan 2005/11/16
                 strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','416')"
                
                 If Not ClsLawExecSQL(3, strTxt) Then
                     MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                 End If
                 NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                 'end by Toni 20080916
               Else
                  strET03 = "07"
                  '刪除定稿暫存資料
                  EndLetter ET01, strCP09, strET03, strUserNum
                  '新增定稿暫存資料
                  strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                  strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
                  'Add by Morgan 2005/11/16
                  strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','416')"
               
                  If Not ClsLawExecSQL(3, strTxt) Then
                     MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                  End If
                  NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
               End If
           End If
           
       'Add by Morgan 2006/5/15
       ElseIf strNP07 = "119" Then
          strET03 = "13"
          '刪除定稿暫存資料
          EndLetter ET01, strCP09, strET03, strUserNum
          '新增定稿暫存資料
          strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                          "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(strNP08) & ")"
          strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                          "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
          strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                          "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序'," & CNULL(strNP07) & ")"
         
          If Not ClsLawExecSQL(3, strTxt) Then
              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
          End If
          NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
       
       'Added by Morgan 2024/11/6
       ElseIf strNP07 = "615" Then
            strET03 = "23"
            '刪除定稿暫存資料
            EndLetter ET01, strCP09, strET03, strUserNum
            '新增定稿暫存資料
            strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(strNP08) & ")"
            strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
            strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序'," & CNULL(strNP07) & ")"
            strExc(1) = ""
            If PUB_GetCNExtDays(m_CurCP(), , intI) Then
               If intI > 0 Then strExc(1) = intI
            End If
            strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','補償天數'," & CNULL(strExc(1)) & ")"
            strFee = PUB_GetCN615Fee(m_CurCP())
            strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
            If Not ClsLawExecSQL(5, strTxt) Then
                MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
            End If
            NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
            
            'Added by Morgan 2025/3/10
            If m_bolFMP Then
               strUserNum = strFMPNum
               m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&615"
               strTmp2 = "53"
               EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
               strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
               strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
               strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strFee & "')"
               strExc(1) = PUB_GetUSXRate
               strExc(2) = ""
               If Val(strExc(1)) <> 0 Then
                  strExc(2) = Fix(strFee / Val(strExc(1)))
               End If
               strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                        
               If Not ClsLawExecSQL(4, strTxt) Then
                   MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
               End If
               NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
               strUserNum = strUser1Num
            End If
            'end 2025/3/10
       '年費
       Else
           'Add By Cheng 2003/10/13
           'Begin
           blnSitu1 = False
           'Modify by Morgan 2004/5/18
           '抓不續辦時加續期費(延展費)
           'strSQLA = "Select * From Nextprogress Where " & ChgNextProgress(strReceiveNo) & " And NP07=" & 年費 & " And NP06='N' "
           StrSQLa = "Select * From Nextprogress Where " & ChgNextProgress(strReceiveNo) & " And NP07 in (" & 年費 & "," & 延展費 & ") And NP06='N' "
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               Do While Not rsA.EOF
                   If IsNull(rsA("NP08").Value) = False Then
                       '92.11.3 MODIFY BY SONIA
                       'If rsA("NP08").Value < strSrvDate(1) And DateDiff("m", ChangeWStringToWDateString(rsA("NP08").Value), ChangeWStringToWDateString(strSrvDate(1))) <= 7 Then
                       'Modify by Morgan 2006/6/14
                       'If rsA("NP08").Value < strSrvDate(1) And DateDiff("m", ChangeWStringToWDateString(rsA("NP08").Value), ChangeWStringToWDateString(Text2.Text)) <= 7 Then
                       If rsA("NP08").Value < strSrvDate(1) And DateDiff("m", ChangeWStringToWDateString(rsA("NP08").Value), ChangeWStringToWDateString(TransDate(Text2.Text, 2))) <= 7 Then
                       '92.11.3 END
                           strOldNP09 = "" & rsA("NP09").Value
                           blnSitu1 = True
                           Exit Do
                       End If
                   End If
                   rsA.MoveNext
               Loop
           End If
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
           If blnSitu1 = True Then
               If m_PA09 = "020" Then
                   strET03 = "09"
                   
                  'Added by Morgan 2015/8/28 非台灣信函進度要存報價(逾期原定稿只有點數)
                  strPA72NextYear = getPA72NextYear(m_CurCP(1), m_CurCP(2), m_CurCP(3), m_CurCP(4), , , strPA25)
                  If strPA72NextYear <> "" Then
                     strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                     strPoint = Round(Val(strPoint) / 1000, 1)
                  Else
                     strPoint = ""
                  End If
                  PUB_UpdateLP2930 m_LD18, "", strPoint
                  'end 2015/8/28
                   
                   '刪除定稿暫存資料
                   EndLetter ET01, strCP09, strET03, strUserNum
                   '新增定稿暫存資料
                   ii = 1
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','半年前法定期限'," & CNULL(strOldNP09) & ")"
                   ii = ii + 1
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                   ii = ii + 1
                   strPA72Year = GetNowNP09(strReceiveNo)
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                   ii = ii + 1
                   
                  'Added by Morgan 2023/6/5
                  'Removed by Morgan 2023/6/5 取消,起算日相同,不會發生
                  'If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                  '   If strPA25 < CompDate(1, 6, strPA72Year) Then
                  '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  '         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                  '      ii = ii + 1
                  '   End If
                  'End If
                  'end 2023/6/5
                           
                   'Add by Morgan 2005/11/16
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                   ii = ii + 1
                   
                  'Add by Morgan 2009/10/19
                  '98/10/1以後的一案兩請案,新型年費定稿加提醒
                  If m_PA09 = "020" And m_strPA08 = "2" And Val(m_strPA10) >= 20091001 Then
                     strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                        " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                        " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                        ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                             "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                           ii = ii + 1
                        'Added by Morgan 2012/8/30
                        '已核准未公告
                        ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                             "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                           ii = ii + 1
                        'end 2012/8/30
                        End If
                     End If
                  End If
                           
                   If Not ClsLawExecSQL(ii - 1, strTxt) Then
                       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                   End If
                   NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                      
               ElseIf m_PA09 = "000" Then
                   strET03 = "08"
                   '刪除定稿暫存資料
                   EndLetter ET01, strCP09, strET03, strUserNum
                   '新增定稿暫存資料
                   ii = 1
                   
                  'Added by Morgan 2012/9/21
                  '102新法一案兩請案,新型年費定稿加提醒
                  If m_strPA08 = "2" And Val(m_strPA10) >= 20130101 Then
                     strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                        " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                        " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                        ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        '未核准
                        If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                             "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                           ii = ii + 1
                        '已核准未公告
                        ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                             "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                           ii = ii + 1
                        End If
                     End If
                  End If
                  'end 2012/9/21
                   
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','半年前法定期限'," & CNULL(strOldNP09) & ")"
                   ii = ii + 1
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                   ii = ii + 1
                   strPA72Year = GetNowNP09(strReceiveNo)
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                   ii = ii + 1
                   
                  'Added by Morgan 2023/6/5
                  If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                     If strPA25 < CompDate(1, 6, strPA72Year) Then
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                        ii = ii + 1
                     End If
                  End If
                  'end 2023/6/5
                           
                   'Remove by Morgan 2011/7/6 定稿已改用共用文字欄位
                   ''Add by Morgan 2004/6/4 年費收費標準
                   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                   '     " SELECT '" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','年費收費標準',FTM05 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='21' AND FTM03='000' AND FTM04='02'"
                   'ii = ii + 1
                   'end 2011/7/6
                   
                   'Add by Morgan 2005/11/16
                   'Modified by Morgan 2022/9/1 不可傳入606否則回覆單會多帶
                   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605')"
                   'end 2022/9/1
                   ii = ii + 1
                   If Not ClsLawExecSQL(ii - 1, strTxt) Then
                       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                   End If
                   NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
               End If
           'End
           Else
               '大陸
               'Modify by Morgan 2008/3/20 +澳門(044)
               'If m_PA09 = "020" Then
               If m_PA09 = "020" Or m_PA09 = "044" Then
                  '取得下次繳費年度
                   strPA72NextYear = getPA72NextYear(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, , m_bFirstYear, strPA25)
                   
                  'Modify by Morgan 2008/3/20 +澳門(044)
                  If m_PA09 = "044" Then
                     'Add by Morgan 2008/5/7 +繳第一次年費(無繳費記錄)
                     If m_bFirstYear = True Then
                        strET03 = "18"
                     Else
                        strET03 = "17"
                     End If
                  Else
                     'Modify By Sindy 2009/05/22 改定稿格式
                     'strET03 = IIf(strSrvDate(1) > m_NP09, "03", "02")
                     strET03 = IIf(strSrvDate(1) > m_NP09, "03", "21")
                  End If
                     
                   EndLetter ET01, strCP09, strET03, strUserNum
                   ii = 1
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                   ii = ii + 1
                   
                   'Add by Morgan 2010/1/18 FMP約定期限
                   If m_bolFMP Then
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','約定期限','" & strNP23 & "')"
                     ii = ii + 1
                   End If
                   
                   
                   '92.1.14 ADD BY SONIA 計算該年年費屆滿日期strPA72Year
                   strPA72Year = getPA72Year(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text)
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                   ii = ii + 1
                   '92.1.14 END
                   
                  'Added by Morgan 2023/6/5
                  'Removed by Morgan 2023/6/5 取消,起算日相同,不會發生
                  'If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                  '   If strPA25 < CompDate(1, 6, strPA72Year) Then
                  '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  '         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                  '      ii = ii + 1
                  '   End If
                  'End If
                  'end 2023/6/5
                  'end 2023/6/5
                           
                   strNextYearFee = ""
                   If strPA72NextYear <> "" Then
                   'Modified by Lydia 2015/01/07 採共用模組
                    ' strNextYearFee = PUB_GetYF0607(m_PA09, m_strPA08, "Y00000001", "605", strPA72NextYear, strPA72NextYear)
                     strNextYearFee = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                       "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用','" & strNextYearFee & "')"
                       ii = ii + 1
                   End If
                   
                   
                  'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                  If strPA72NextYear <> "" Then
                     strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                     strPoint = Round(Val(strPoint) / 1000, 1)
                  Else
                     strPoint = ""
                  End If
                  PUB_UpdateLP2930 m_LD18, strNextYearFee, strPoint
                  'end 2015/8/28
                           
                   'Add by Morgan 2005/11/16
                   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                   ii = ii + 1
                   
                  'Add by Morgan 2009/10/7
                  '98/10/1以後的一案兩請案,新型年費定稿加提醒
                  If m_PA09 = "020" And m_strPA08 = "2" And Val(m_strPA10) >= 20091001 Then
                     strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14,pa11,pa77" & _
                        " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                        " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                        ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        'Added by Morgan 2017/9/20
                        strInventionCaseNo = "" & RsTemp("C1")
                        strInventionPA11 = "" & RsTemp("pa11")
                        strInventionPA77 = "" & RsTemp("pa77")
                        'end 2017/9/20
                        If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                           bolDualCaseUtility = True 'Added by Morgan 2017/9/20
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                             "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                           ii = ii + 1
                        'Added by Morgan 2012/8/30
                        '已核准未公告
                        ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                           bolDualCaseUtility = True 'Added by Morgan 2017/9/20
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                             "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                           ii = ii + 1
                        'end 2012/8/30
                        End If
                     End If
                  End If
                  'end 2009/10/7
                   
                  If Not ClsLawExecSQL(ii - 1, strTxt) Then
                     MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                  End If
                  
                  NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                  
                  'Add by Morgan 2009/12/7
                  If m_bolFMP Then
                     strUserNum = strFMPNum
                     
                     'Modified by Morgan 2014/8/20 FMP案有年費代理人,改傳本所案號+案件性質
                     'm_FMP_ET02 = strCP09
                     m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&605"
                     'end 2014/8/20
                     
                     'Removed by Morgan 2022/9/20 定稿已合併
                     ''付款後辦案
                     'If CU72FA39("", strPA75) Then
                     '   strTmp2 = "53"
                     'Else
                     'end 2022/9/20
                     
                        'Added by Morgan 2022/9/19
                        If m_PA09 = "044" Then
                           strTmp2 = "54"
                        Else
                        'end 2022/9/19
                           strTmp2 = "52"
                        End If
                        
                     'End If 'Removed by Morgan 2022/9/20
                     
                     EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                     strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                     strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                     If m_PA46 = "Y" Then
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & _
                           "','PCT案','♀')"
                     Else
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & _
                           "','非PCT案','♀')"
                     End If
                     strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','下次年費年度','" & strPA72NextYear & "')"
                     strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strNextYearFee & "')"
                     strExc(1) = PUB_GetUSXRate
                     strExc(2) = ""
                     If Val(strExc(1)) <> 0 Then
                        strExc(2) = Fix(strNextYearFee / Val(strExc(1)))
                     End If
                     strTxt(6) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                        " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                    
                     ii = 6
                     'Added by Morgan 2017/9/20
                     If bolDualCaseUtility = True Then
                        strTxt(7) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','一案兩請新型案要印','♀')"
                        strTxt(8) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案本所案號','" & strInventionCaseNo & "')"
                        strTxt(9) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案申請號','" & ChgSQL(strInventionPA11) & "')"
                        strTxt(10) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案彼所案號','" & ChgSQL(strInventionPA77) & "')"
                        ii = 10
                     End If
                     'end 2017/9/20
                              
                     If Not ClsLawExecSQL(ii, strTxt) Then
                        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                     End If
                     NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                     strUserNum = strUser1Num
                  End If
                'Add by Morgan 2004/5/14
                ElseIf m_PA09 = "013" Then '香港
                
                   '取得已繳費年度及專利種類
                   strPA72NextYear = getNextPayYear(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, strPA72Year, strPA25)
                   Select Case m_strPA08
                      Case "1" '標準專利
                         stSitu = "12"
                      Case "2" '短期專利
                         stSitu = "11"
                      Case "3" '外觀設計
                         stSitu = "10"
                   End Select
                   
                   If strNP07 = 維持費 Then stSitu = "02" 'Added by Morgan 2012/10/23
                   
                   '刪除定稿暫存資料
                   EndLetter ET01, strCP09, stSitu, strUserNum
                   
                    '新增定稿暫存資料
                    ii = 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                    
                    ii = ii + 1
                    
                     'Added by Morgan 2023/6/1
                     'Removed by Morgan 2023/6/5 取消,起算日相同,不會發生
                     'If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                     '   If strPA25 < CompDate(1, 6, strPA72Year) Then
                     '      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     '         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                     '      ii = ii + 1
                     '   End If
                     'End If
                     'end 2023/6/5
                     'end 2023/6/1
                    
                    If strPA72NextYear <> "" Then
                    'Modified by Lydia 2015/01/07 採共用模組
'                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                                        "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','費用','" & Val(PUB_GetYF0607(m_PA09, m_strPA08, "Y00000001", strNP07, strPA72NextYear, strPA72NextYear)) & "')"
                        strNextYearFee = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, strNP07, strPA72NextYear, strPA72NextYear, "1")
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                        "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','費用','" & Val(strNextYearFee) & "')"
                        ii = ii + 1
                    End If
                    
                     'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                     If strPA72NextYear <> "" Then
                        strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, strNP07, strPA72NextYear, strPA72NextYear, "1")
                        strPoint = Round(Val(strPoint) / 1000, 1)
                     Else
                        strPoint = ""
                     End If
                     PUB_UpdateLP2930 m_LD18, strNextYearFee, strPoint
                     'end 2015/8/28
                           
                    'Add by Morgan 2005/11/16
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','下一程序','605,607')"
                    ii = ii + 1
                    
                    If Not ClsLawExecSQL(ii - 1, strTxt) Then
                        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                    End If
                    NowPrint strCP09, ET01, stSitu, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                                        
                    'Add by Morgan 2022/9/30
                     If m_bolFMP Then
                        strUserNum = strFMPNum
                        m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&605"
                        strTmp2 = "54"
                        EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                        strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                        strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','下次年費年度','" & strPA72NextYear & "')"
                        strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strNextYearFee & "')"
                        strExc(1) = PUB_GetUSXRate
                        strExc(2) = ""
                        If Val(strExc(1)) <> 0 Then
                           strExc(2) = Fix(strNextYearFee / Val(strExc(1)))
                        End If
                        strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                        ii = 5
                        If Not ClsLawExecSQL(ii, strTxt) Then
                           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                        End If
                        NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                        strUserNum = strUser1Num
                     End If
                     'end 2022/9/30
               Else
                   '台灣
                   If m_PA09 = "000" Then
                     iPlusFee = 0 'Added by Morgan 2013/1/8
                     
                     'for 催年費大-->台 定稿 20080916 add by toni
                     If PUB_CheckCuNation(RsTemp.Fields("PA26"), RsTemp.Fields("PA01"), RsTemp.Fields("PA02"), RsTemp.Fields("PA03"), RsTemp.Fields("PA04")) = "1" Then
                        strET03 = "19"
                        'Added by Morgan 2013/1/8 專利處大對台年費服務費+500 --郭雅娟
                        strExc(1) = PUB_GetStaffST15(PUB_GetAKindSalesNo(RsTemp.Fields("PA01"), RsTemp.Fields("PA02"), RsTemp.Fields("PA03"), RsTemp.Fields("PA04")), "1")
                        If Left(strExc(1), 2) = "P1" Then
                           iPlusFee = 500
                        End If
                        'end 2013/1/8
                     Else
                        'Modif by Morgan 2008/1/7 一率改用新定稿
                        'strET03 = "01"
                        strET03 = "15"
                     End If
                     'end by Toni 20080916
                     
                       EndLetter ET01, strCP09, strET03, strUserNum
                       ii = 1
                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                       "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(TransDate(Text2.Text, 2)) & ")"
                       ii = ii + 1
                       '92.1.14 MODIFY BY SONIA 計算該年年費屆滿日期strPA72Year
                       strPA72Year = getPA72Year(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, strPA25)
                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                       "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                       ii = ii + 1
                       
                        'Added by Morgan 2023/6/1
                        If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                           If strPA25 < CompDate(1, 6, strPA72Year) Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                              ii = ii + 1
                           End If
                        End If
                        'end 2023/6/1
                           
                        'Added by Morgan 2012/9/21
                        '102新法一案兩請案,新型年費定稿加提醒
                        If m_strPA08 = "2" And Val(m_strPA10) >= 20130101 Then
                           strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                              " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                              " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                              ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              '未核准
                              If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                 ii = ii + 1
                              '已核准未公告
                              ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                   "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                 ii = ii + 1
                              End If
                           End If
                        End If
                        'end 2012/9/21
                        
                        If DBDATE(strPA72Year) >= 20130101 Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','102新法不印','♀')"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','102新法要印','♀')"
                           ii = ii + 1
                        End If
                        'end 2012/9/21
                       
                       '92.1.14 END
                       
                       'Add By Cheng 2002/10/24
                       '取得下次繳費年度
                       strPA72NextYear = getPA72NextYear(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, strMaxFeeYear)
                        If Val(strPA72NextYear) > 0 Then
                           'Modify by Morgan 2007/9/21
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           '                "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用','" & Val(PUB_GetYF0607(m_PA09, m_strPA08, "Y00000001", "605", strPA72NextYear, strPA72NextYear)) & "')"
                           'ii = ii + 1

                           '服務費,規費
                           'Modified by Lydia 2015/01/07 採共用模組
'                           strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & m_PA09 & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & strPA72NextYear
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              strExc(1) = "" & RsTemp("YF06")
'                              strExc(2) = "" & RsTemp("YF07")
'                           Else
'                              strExc(1) = ""
'                              strExc(2) = ""
'                           End If
                            strExc(0) = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1", strExc(1), strExc(2))
                            If strExc(0) = "0" Then strExc(1) = "": strExc(2) = ""
                              
                           If strExc(1) <> "" Then
                              strExc(1) = Val(strExc(1)) + iPlusFee 'Added by Morgan 2013/1/8
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','服務費','" & strExc(1) & "')"
                              ii = ii + 1
                           End If
                           '年費是否可減免
                           If PUB_GetCaseDiscStat(Text1(1) & Text1(2) & Text1(3) & Text1(4)) = "Y" Then
                              bolDiscount = True
                           Else
                              bolDiscount = False
                           End If
                           
                           If Val(strExc(2)) > 0 Then
                              '減免
                              If Val(strPA72NextYear) < 7 Then
                                 If bolDiscount = True Then
                                    If Val(strPA72NextYear) < 4 Then
                                       strExc(2) = Val(strExc(2)) - 800
                                    Else
                                       strExc(2) = Val(strExc(2)) - 1200
                                    End If
                                 End If
                              End If
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','規費','" & strExc(2) & "')"
                              ii = ii + 1
                           End If

                           strExc(3) = Val(strExc(1)) + Val(strExc(2))
                           If Val(strExc(3)) > 0 Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','費用','" & strExc(3) & "')"
                              ii = ii + 1
                           End If
                           'end 2007/9/21
                           
                           'Add by Morgan 2007/11/20 下兩年的費用也要印
                           strExc(5) = strExc(2) '規費累計
                           strExc(6) = strExc(3) '費用累計
                           'Added by Lydia 2024/08/15
                           Dim strBaseYear As String
                           strBaseYear = strPA72NextYear
                           'end 2024/08/15
                           For jj = 1 To 2
                              strPA72NextYear = Val(strPA72NextYear) + 1
                              If Val(strPA72NextYear) <= Val(strMaxFeeYear) Then
                              'Modified by Lydia 2015/01/07 採共用模組
'                                 strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & m_PA09 & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & strPA72NextYear
'                                 intI = 1
'                                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                                 If intI = 1 Then
'                                    strExc(2) = "" & RsTemp("YF07")
'                                 Else
'                                    strExc(2) = ""
'                                 End If
                                strExc(0) = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1", , strExc(2))
                                If strExc(0) = "0" Then strExc(2) = ""
                                
                                 If Val(strExc(2)) > 0 Then
                                    '減免
                                    If Val(strPA72NextYear) < 7 Then
                                       If bolDiscount = True Then
                                          If Val(strPA72NextYear) < 4 Then
                                             strExc(2) = Val(strExc(2)) - 800
                                          Else
                                             strExc(2) = Val(strExc(2)) - 1200
                                          End If
                                       End If
                                    End If
                                    '規費累計
                                    strExc(5) = Val(strExc(5)) + Val(strExc(2))
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                       "','規費" & jj & "','" & strExc(5) & "')"
                                    ii = ii + 1
                                    
                                    '費用累計
                                    'Added by Lydia 2024/08/15 重抓服務費; ex.訊強電子 (惠州 )X41570060, P-117332
                                    strExc(0) = " select '1' as ord1, ys07 from patentyearspec where ys01='" & m_PA09 & "' and ys03='" & m_PA26 & "' and ys02='" & m_strPA08 & "' and ys04='605' and ys05='" & strBaseYear & "' and ys06='" & strPA72NextYear & "' " & _
                                                " union select '2' as ord1, yf06 as ys07 from patentyearfee where yf01='" & m_PA09 & "' and yf03='" & m_PA26 & "' and yf02='" & m_strPA08 & "' and yf04='605' and yf05='" & strPA72NextYear & "' " & _
                                                " order by 1"
                                    intI = 1
                                    Set rsTemp10 = ClsLawReadRstMsg(intI, strExc(0))
                                    If intI = 1 Then
                                       strExc(1) = Val("" & rsTemp10.Fields("ys07"))
                                    End If
                                    'end 2024/08/15
                                    strExc(6) = Val(strExc(1)) + Val(strExc(5))
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                       "','費用" & jj & "','" & strExc(6) & "')"
                                    ii = ii + 1
                                 End If
                              Else
                                 Exit For
                              End If
                           Next
                           'end 2007/11/20
                           
                        End If
                       
                      'Remove by Morgan 2011/7/6 定稿已改用共用文字欄位
                      ''Add by Morgan 2004/6/4 年費收費標準
                      'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
                      '   " SELECT '" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','年費收費標準',FTM05 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='21' AND FTM03='000' AND FTM04='02'"
                      'ii = ii + 1
                      'end 2011/7/6
                     
'Remove by Morgan 2011/8/3 新定稿已不再使用
'                     'Add by Morgan 2005/5/17 辦理減免退費提醒
'                     If PUB_GetCaseDiscStat(strReceiveNo) = "Y" Then
'                         If PUB_CheckYearFeeReturn(m_CurCP, False, m_iDiscount, m_iYear1, m_iYear2) = True Then
'                            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                               "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','減免退費起迄年','" & IIf(m_iYear1 = m_iYear1, m_iYear1, m_iYear1 & "年至第" & m_iYear2) & "')"
'
'                            ii = ii + 1
'                            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                               "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','減免退費金額','" & m_iDiscount & "')"
'                            ii = ii + 1
'                         End If
'                     End If
'                     '2005/5/16 end
                     
                     'Add by Morgan 2011/7/6
                     If PUB_ChkRefund(m_CurCP, m_lngRefund) = True Then
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','未退金額','" & m_lngRefund & "')"
                        ii = ii + 1
                     End If
                     
                      'Add by Morgan 2005/11/16
                      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                         "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                      ii = ii + 1
                      
                      If Not ClsLawExecSQL(ii - 1, strTxt) Then
                         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                      End If
                      NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                      Call UpdateAI 'Add by Morgan 2005/5/17 台灣新增年費通知紀錄
                   End If
               End If
           End If
       End If
               
NoLetter:

'Remove by Morgan 2008/8/13 改開窗定稿
'       '若為年費通知函, 不論是否閉卷資料都要出來
'       If m_Select = "1" And strNP07 <> "119" Then
'           strSQL = "SELECT PA26,pa01||'-'||pa02||'-'||pa03||'-'||pa04 FROM PATENT WHERE " & ChgPatent(strReceiveNo)
'       '實體審查通知函
'       Else
'           strSQL = "SELECT PA26,pa01||'-'||pa02||'-'||pa03||'-'||pa04 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND (PA57<>'Y' OR PA57 IS NULL)"
'       End If
'       Set rsTmp = New ADODB.Recordset
'       rsTmp.CursorLocation = adUseClient
'       rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'       If rsTmp.RecordCount > 0 Then
'           If IsNull(rsTmp.Fields("PA26")) = False Then
'               If IsEmptyText(rsTmp.Fields("PA26")) = False Then
'                   ReDim Preserve m_CustList(m_CustListCount + 1)
'                   ReDim Preserve m_CP(m_CustListCount + 1)
'                   m_CustList(m_CustListCount) = rsTmp.Fields("PA26")
'                   m_CP(m_CustListCount) = CheckStr(rsTmp.Fields(1).Value)
'                   m_CustListCount = m_CustListCount + 1
'               End If
'           End If
'       End If
'       rsTmp.Close
'       Set rsTmp = Nothing
       
       '只印地址條時不印接洽結案單
       'Modify by Morgan 2005/5/18 控制是否有減免
       'If Me.Check1.Value = vbUnchecked Then g_PrtForm001.PrintForm TmpNp22, Text1(1), Text1(2), Text1(3), Text1(4)
      If Me.Check1.Value = vbUnchecked Then
         If m_iDiscount > 0 Then
            g_PrtForm001.PrintForm m_NP22, Text1(1), Text1(2), Text1(3), Text1(4), , "1"
         Else
            g_PrtForm001.PrintForm m_NP22, Text1(1), Text1(2), Text1(3), Text1(4)
         End If
      End If
      bolPrint = True
   End If
   
   If bolPrint = True Then
'Remove by Morgan 2008/8/13 改開窗定稿
'      '只印期限表時不印地址條
'      If Me.Check2.Value = vbUnchecked Then
'         If m_CustListCount > 0 Then
'            If MsgBox("按確定後開始列印地址條!", vbOKCancel + vbInformation, "列印地址條") = vbOK Then
'               PrintAddress
'            End If
'         End If
'      End If
      MsgBox "列印結束 !", vbInformation
   Else
      MsgBox "無符合條件之資料可列印 !", vbInformation
   End If
End Sub

'Add by Morgan 2006/4/3
Private Function FormCheck3() As Boolean
   'Add by Morgan 2007/9/5
   'Added by Lydia 2015/04/20 + TW-SUPA
   If chkKind(0).Value + chkKind(1).Value + chkKind(2).Value + chkKind(3).Value = 0 Then
      MsgBox "請勾選通知函類別!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
      
   '檢查本所案號
   If Text1(2).Text = "" Then
       MsgBox "本所案號不可空白，請重新輸入 !", vbCritical
       Me.Text1(2).SetFocus
       Text1_GotFocus 2
       Exit Function
   End If
   'Modified by Lydia 2015/04/20 +TW-SUPA不必輸下次繳費日
  ' If chkKind(0).Value <> 1 Then
   If chkKind(0).Value + chkKind(3).Value = 0 Then
      If Me.Text2.Text = "" Then
          MsgBox "請輸入下次繳費日!!!", vbExclamation + vbOKOnly
          Me.Text2.SetFocus
          Text2_GotFocus
          Exit Function
      End If
      If PUB_CheckKeyInDate(Me.Text2) = -1 Then
          Me.Text2.SetFocus
          Text2_GotFocus
          Exit Function
      End If
   End If
   FormCheck3 = True
End Function

'Add by Morgan 2007/9/4
'通知主張國外優先權
Private Sub Process4()
   Dim strCP09 As String '收文號
   Dim Prn As Printer
   Dim strTmp As String
   
'Remove by Morgan 2008/8/13 改開窗定稿
'   ClearCustList
'   '搜尋預設印表機
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   
   m_CurCP(1) = Text1(1)
   m_CurCP(2) = Text1(2)
   m_CurCP(3) = Right("0" & Text1(3), 1)
   m_CurCP(4) = Right("00" & Text1(3), 2)
   pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & m_CurCP(1) & "-" & m_CurCP(2) & "-" & m_CurCP(3) & "-" & m_CurCP(4) 'Add By Sindy 2010/11/29
    'Add by Lydia 2015/01/27 +fmp寰華控制sql (m_selarea)
     Call ChangeSel(2) '將SQL改為對應PA
     
   strExc(0) = "select pa26,pa01||'-'||pa02||'-'||pa03||'-'||pa04 from patent where pa01='" & m_CurCP(1) & "' and pa02='" & m_CurCP(2) & "' and pa03='" & m_CurCP(3) & "' and pa04='" & m_CurCP(4) & "' and pa10>0 and pa09='000'" & m_SelArea
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/11/29
      If Me.Check1.Value = vbUnchecked Then
         strCP09 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&000"
         strTmp = "16"
         NowPrint strCP09, ET01, strTmp, False, strUserNum, 0, , , , , , , , , , , , m_LD18
      End If
'Remove by Morgan 2008/8/13 改開窗定稿
'      If "" & RsTemp.Fields("PA26") <> "" Then
'         ReDim Preserve m_CustList(m_CustListCount + 1)
'         ReDim Preserve m_CP(m_CustListCount + 1)
'         m_CustList(m_CustListCount) = RsTemp.Fields("PA26")
'         m_CP(m_CustListCount) = CheckStr(RsTemp.Fields(1).Value)
'         m_CustListCount = m_CustListCount + 1
'         If m_CustListCount > 0 Then
'            If MsgBox("按確定後開始列印地址條!", vbOKCancel + vbInformation, "列印地址條") = vbOK Then
'               PrintAddress
'            End If
'         End If
'      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/29
      MsgBox "無符合條件之資料可列印 !", vbInformation
   End If
End Sub
'Modified by Lydia 2015/07/20 改共用 Pub_Get416Period
'Add by Morgan 2009/7/9
'提實審期間
'Private Function Get416Period(PA08 As String, PA09 As String) As String
'   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer, stCol As String
'   Dim iNo As Integer
'   Select Case PA08
'      Case "1"
'         stCol = "na27"
'      Case "2"
'         stCol = "na29"
'      Case Else
'         stCol = "na31"
'   End Select
'   stSQL = "select " & stCol & " from nation where na01='" & PA09 & "'"
'   intR = 1
'   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      If Not IsNull(adoRst(0)) Then
'         If adoRst(0) Mod 12 = 0 Then
'            iNo = adoRst(0) / 12
'            Get416Period = PUB_ChgNumber2Chinese(Format(iNo)) & "年"
'         Else
'            iNo = adoRst(0)
'            Get416Period = PUB_ChgNumber2Chinese(Format(iNo)) & "個月"
'         End If
'      End If
'   End If
'   Set adoRst = Nothing
'End Function

'Moddified by Lydia 2019/08/30 改成傳入變數,計算區間
'Private Sub SetDateCondition_Old()
'   If Text3 <> "" Then
'      '10號
'      If Option3(0).Value = True Then
'         strExc(0) = Text3 & "10"
'         '台灣
'         If Option2(0) Then
'            'P:+3個月的1~20號
'            strExc(2) = CompDate(1, 3, strExc(0))
'            strExc(1) = Left(strExc(2), 6) & "01"
'            strExc(2) = Left(strExc(2), 6) & "20"
'            text1(5) = TransDate(strExc(1), 1)
'            text1(6) = TransDate(strExc(2), 1)
'            'FMP
'            text1(7) = ""
'            text1(8) = ""
'         '非台灣
'         Else
'            'P:下月16號-該月底
'            'Modified by Morgan 2017/1/13 領證發文也要用寫共用
'            'strExc(2) = CompDate(2, -1, Left(CompDate(1, 2, strExc(0)), 6) & "01")
'            'strExc(1) = Left(strExc(2), 6) & "16"
'            If Get605InformPeriod4NonTwCase(strExc(0), False, strExc(1), strExc(2)) = False Then Exit Sub
'            'end 2017/1/13
'            text1(5) = TransDate(strExc(1), 1)
'            text1(6) = TransDate(strExc(2), 1)
'            'FMP:+3月
'            'Modified by Morgan 2017/1/13 領證發文也要用寫共用
'            'strExc(1) = Left(CompDate(1, 3, strExc(0)), 6) & "11"
'            'strExc(2) = Left(strExc(1), 6) & "20"
'            '第一次要包含過渡的資料
'            If strExc(0) = "990110" Then
'               text1(7) = "990319"
'            Else
'               text1(7) = TransDate(strExc(1), 1)
'            End If
'            If Get605InformPeriod4NonTwCase(strExc(0), True, strExc(1), strExc(2)) = False Then Exit Sub
'            text1(7) = TransDate(strExc(1), 1)
'            'end 2017/1/13
'
'            text1(8) = TransDate(strExc(2), 1)
'         End If
'
'      '20號
'      Else
'         strExc(0) = Text3 & "20"
'         '台灣
'         If Option2(0) Then
'            'P:+3個月的21號~月底
'            strExc(2) = CompDate(2, -1, Left(CompDate(1, 4, strExc(0)), 6) & "01")
'            strExc(1) = Left(strExc(2), 6) & "21"
'            text1(5) = TransDate(strExc(1), 1)
'            text1(6) = TransDate(strExc(2), 1)
'            'FMP
'            text1(7) = ""
'            text1(8) = ""
'         Else
'            'P:下下月1號-15號
'            'Modified by Morgan 2017/1/13 領證發文也要用寫共用
'            'strExc(1) = Left(CompDate(1, 2, strExc(0)), 6) & "01"
'            'strExc(2) = Left(strExc(1), 6) & "15"
'            If Get605InformPeriod4NonTwCase(strExc(0), False, strExc(1), strExc(2)) = False Then Exit Sub
'            'end 2017/1/13
'            text1(5) = TransDate(strExc(1), 1)
'            text1(6) = TransDate(strExc(2), 1)
'            'FMP:+3月
'            'Modified by Morgan 2017/1/13 領證發文也要用寫共用
'            'strExc(1) = Left(CompDate(1, 3, strExc(0)), 6) & "21"
'            'strExc(2) = Left(CompDate(1, 1, strExc(1)), 6) & "10"
'            If Get605InformPeriod4NonTwCase(strExc(0), True, strExc(1), strExc(2)) = False Then Exit Sub
'            'end 2017/1/13
'            text1(7) = TransDate(strExc(1), 1)
'            text1(8) = TransDate(strExc(2), 1)
'         End If
'      End If
'   End If
'End Sub
'Modified by Lydia 2022/08/12 +指定客戶bCU01
Private Sub SetDateCondition(ByVal pYYMM As String, ByVal bText As Boolean, Optional ByVal pMon As Integer = 3, Optional ByVal bCU01 As String)
'pMon 催期限月數(預設3個月)
Dim pDate1 As String, pDate2 As String, pDate3 As String

   If pYYMM <> "" Then
      m_Date1 = "": m_Date2 = ""
      m_FMPDate1 = "": m_FMPDate2 = ""
      '10號
      If Option3(0).Value = True Then
         pDate1 = pYYMM & "10"
         '台灣
         'Memo by Lydia 2019/12/16 不區分台灣案及非台灣案，畫面請分別列出台灣案與非台灣案的本所期限、法定期限通知區段。
         'If Option2(0) Then 'Remove by Lydia 2019/12/16 一併跑非大陸案
            'P:+3個月的1~20號
            pDate3 = CompDate(1, pMon, pDate1)
            pDate2 = Left(pDate3, 6) & "01"
            pDate3 = Left(pDate3, 6) & "20"
            'Modified by Lydia 2019/12/16
            'm_Date1 = TransDate(pDate2, 1)
            'm_Date2 = TransDate(pDate3, 1)
            ''FMP
            'm_FMPDate1 = ""
            'm_FMPDate2 = ""
            m_DateTW1 = TransDate(pDate2, 1)
            m_DateTW2 = TransDate(pDate3, 1)
         '非台灣
         'Else  'Remove by Lydia 2019/12/16 一併跑非大陸案
            'P:下月16號-該月底
            '領證發文也要用寫共用
            'Modified by Lydia 2021/10/18 和碩案件在法限-6個月時就催; ex. P-112859
            'If Get605InformPeriod4NonTwCase(pDate1, False, pDate2, pDate3) = False Then Exit Sub
            'Modified by Lydia 2022/08/12 改傳入指定客戶
            'If Get605InformPeriod4NonTwCase(pDate1, False, pDate2, pDate3, IIf(bText = False, "X70017000", "")) = False Then Exit Sub
            If Get605InformPeriod4NonTwCase(pDate1, False, pDate2, pDate3, IIf(bText = False And bCU01 <> "", bCU01, "")) = False Then Exit Sub
            m_Date1 = TransDate(pDate2, 1)
            m_Date2 = TransDate(pDate3, 1)
            'FMP:+3月
            '第一次要包含過渡的資料
            If pDate1 = "990110" Then
               m_FMPDate1 = "990319"
            Else
               m_FMPDate1 = TransDate(pDate2, 1)
            End If
            'Modified by Lydia 2021/10/18 和碩案件在法限-6個月時就催; ex. P-112859
            'If Get605InformPeriod4NonTwCase(pDate1, True, pDate2, pDate3) = False Then Exit Sub
            'Modified by Lydia 2022/08/12 改傳入指定客戶
            'If Get605InformPeriod4NonTwCase(pDate1, True, pDate2, pDate3, IIf(bText = False, "X70017000", "")) = False Then Exit Sub
            If Get605InformPeriod4NonTwCase(pDate1, True, pDate2, pDate3, IIf(bText = False And bCU01 <> "", bCU01, "")) = False Then Exit Sub
            m_FMPDate1 = TransDate(pDate2, 1)
            m_FMPDate2 = TransDate(pDate3, 1)
            
         'End If 'Remove by Lydia 2019/12/16 一併跑非大陸案

      '20號
      Else
         pDate1 = pYYMM & "20"
         '台灣
         'If Option2(0) Then 'Remove by Lydia 2019/12/16 一併跑非大陸案
            'P:+3個月的21號~月底
            pDate3 = CompDate(2, -1, Left(CompDate(1, pMon + 1, pDate1), 6) & "01")
            pDate2 = Left(pDate3, 6) & "21"
            'Modified by Lydia 2019/12/16
            'm_Date1 = TransDate(pDate2, 1)
            'm_Date2 = TransDate(pDate3, 1)
            ''FMP
            'm_FMPDate1 = ""
            'm_FMPDate2 = ""
            m_DateTW1 = TransDate(pDate2, 1)
            m_DateTW2 = TransDate(pDate3, 1)
         'Else 'Remove by Lydia 2019/12/16 一併跑非大陸案
            'P:下下月1號-15號
            '領證發文也要用寫共用
            'Modified by Lydia 2021/10/18 和碩案件在法限-6個月時就催; ex. P-112859
            'If Get605InformPeriod4NonTwCase(pDate1, False, pDate2, pDate3) = False Then Exit Sub
            'Modified by Lydia 2022/08/12 改傳入指定客戶
            'If Get605InformPeriod4NonTwCase(pDate1, False, pDate2, pDate3, IIf(bText = False, "X70017000", "")) = False Then Exit Sub
            If Get605InformPeriod4NonTwCase(pDate1, False, pDate2, pDate3, IIf(bText = False And bCU01 <> "", bCU01, "")) = False Then Exit Sub
            m_Date1 = TransDate(pDate2, 1)
            m_Date2 = TransDate(pDate3, 1)
            'FMP:+3月
            'Modified by Lydia 2021/10/18 和碩案件在法限-6個月時就催; ex. P-112859
            'If Get605InformPeriod4NonTwCase(pDate1, True, pDate2, pDate3) = False Then Exit Sub
            If Get605InformPeriod4NonTwCase(pDate1, True, pDate2, pDate3, IIf(bText = False, "X70017000", "")) = False Then Exit Sub
            m_FMPDate1 = TransDate(pDate2, 1)
            m_FMPDate2 = TransDate(pDate3, 1)
         'End If 'Remove by Lydia 2019/12/16 一併跑非大陸案
      End If
      '設定在TextBox
      If bText = True Then
          'Modified by Lydia 2019/12/16
          'Text1(5) = m_Date1
          'Text1(6) = m_Date2
          'Text1(7) = m_FMPDate1
          'Text1(8) = m_FMPDate2
          TxtDate(0) = m_DateTW1
          TxtDate(1) = m_DateTW2
          TxtDate(2) = m_Date1
          TxtDate(3) = m_Date2
          TxtDate(4) = m_FMPDate1
          TxtDate(5) = m_FMPDate2
          'end 2019/12/16
      End If
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 <> "" Then
      If Not ChkDate(Text3 & "01") Then
         Cancel = True
      Else
         'Added by Lydia 2025/07/25
         If cntFrm040303New = "Y" Then
            Call PUB_SetDateConFrm040303(Text3.Text, IIf(Option3(0).Value = True, "1", "2"), m_Date1, m_Date2, m_FMPDate1, m_FMPDate2, m_DateTW1, m_DateTW2)
            Call SetDateTextBox
         Else
         'end 2025/07/25
            'Modified by Lydia 2019/08/30
            'SetDateCondition
            SetDateCondition Text3.Text, True
         End If
      End If
   End If
End Sub
'Add by Lydia 2015/01/27 開放外專程序人員操作FMP寰華案件。當非FMP寰華權限,不可看寰華案=>回傳SQL
Private Sub ChangeSel(iR As Integer)
'Modified by Lydia 2015/04/21
Select Case iR
Case 1
'If iR = 1 Then
    'Modified by Morgan 2018/3/7
    'm_SelArea = Replace(FMP2openSQL, "f0.CP01", "NP02")
    m_SelArea = Replace(m_FMP2openSQL, "f0.CP01", "NP02")
    'end 2018/3/7
    m_SelArea = Replace(m_SelArea, "f0.CP02", "NP03")
    m_SelArea = Replace(m_SelArea, "f0.CP03", "NP04")
    m_SelArea = Replace(m_SelArea, "f0.CP04", "NP05")
'Else
Case 2
    'Modified by Morgan 2018/3/7
    'm_SelArea = Replace(FMP2openSQL, "f0.CP", "PA")
    m_SelArea = Replace(m_FMP2openSQL, "f0.CP", "PA")
'End If
Case 3
    'Modified by Morgan 2018/3/7
    'm_SelArea = Replace(FMP2openSQL, "f0.CP", "f0.PA")
    m_SelArea = Replace(m_FMP2openSQL, "f0.CP", "f0.PA")
End Select
'end 2015/04/21
End Sub
'Added by Lydia 2015/04/20 P案每月5日TW-SUPA通知定稿改在"繳年費/實體審查通知函"。
'在案件進度也產生一道「通知TW-SUPA期限」記錄,以電子形式發文=>同台灣P案e化作業
Private Sub Process5()
   Dim strCP09 As String '收文號
   Dim Prn As Printer
   Dim strTmp As String
   Dim strTxt(1 To 3) As String
   Dim m_UsPatent As String '美國發明案號
   Dim m_RtnDate As String '客戶回覆期限
   Dim m_TwAppNo As String '台灣案申請號
   Dim mET01 As String, mPA09 As String, mPA26 As String, mPA75 As String, mCP43 As String
   Dim stRefNat As String  '相關案申請國家 Added by Morgan 2021/6/11
   
   m_CurCP(1) = Text1(1)
   m_CurCP(2) = Text1(2)
   m_CurCP(3) = Right("0" & Text1(3), 1)
   m_CurCP(4) = Right("00" & Text1(3), 2)
   pub_QL05 = pub_QL05 & ";" & Label1(11).Caption
   pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & m_CurCP(1) & "-" & m_CurCP(2) & "-" & m_CurCP(3) & "-" & m_CurCP(4)
   If Len(Text2.Text) <> 0 Then
     pub_QL05 = pub_QL05 & ";" & Label1(5) & Text2
   End If
         
   Call ChangeSel(3) '將SQL改為對應PA  (m_selarea)
   '原本在frm040322作業,現在改到這裡,同時新增信函進度檔(每月5日會自動發文)
    mET01 = "18"
    strTmp = "02"
    
    'Modified by Morgan 2021/6/8 +日本並修改語法(發文日早的優先)
    strExc(0) = "select pd01||'-'||pd02||decode(pd03||pd04,'000','','-'||pd03||'-'||pd04) UsNo,c1.cp27,pd06,f0.pa09,f0.pa26,f0.pa75,c3.cp09 PCP43" & _
       ",p2.pa09 RefNat from patent f0,pridate,caseprogress c1,caseprogress c3,patent p2" & _
       " where f0.pa01='" & m_CurCP(1) & "' and f0.pa02='" & m_CurCP(2) & "' and f0.pa03='" & m_CurCP(3) & "' and f0.pa04='" & m_CurCP(4) & "'" & _
       " and f0.pa08='1' and f0.pa09='000' and f0.pa16 is null and pd06(+)=f0.pa11 and pd07(+)=f0.pa09" & m_SelArea & _
       " and p2.pa01(+)=pd01 and p2.pa02(+)=pd02 and p2.pa03(+)=pd03 and p2.pa04(+)=pd04 and p2.pa09 in('101','011') and p2.pa08='1'" & _
       " and c1.cp01(+)=pd01 and c1.cp02(+)=pd02 and c1.cp03(+)=pd03 and c1.cp04(+)=pd04 and c1.cp10='101' and c1.cp27>0 and c1.cp57 is null" & _
       " and c3.cp01(+)=f0.pa01 and c3.cp02(+)=f0.pa02 and c3.cp03(+)=f0.pa03 and c3.cp04(+)=f0.pa04 and c3.cp10='101'" & _
       " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C2 WHERE C2.CP01=f0.PA01 AND C2.CP02=f0.PA02" & _
       " AND C2.CP03=f0.PA03 AND C2.CP04=f0.PA04 AND C2.CP10='1202') order by c1.cp27 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stRefNat = RsTemp.Fields("RefNat") 'Added by Morgan 2021/6/11
      m_UsPatent = RsTemp.Fields("UsNo")
      m_RtnDate = CompDate(1, 5, RsTemp.Fields("cp27"))
      m_TwAppNo = "" & RsTemp.Fields("pd06")
      mPA09 = "" & RsTemp.Fields("PA09")
      mPA26 = "" & RsTemp.Fields("PA26")
      mPA75 = "" & RsTemp.Fields("PA75")
      mCP43 = "" & RsTemp.Fields("PCP43")
      '新增信函進度檔
      If PUB_AddCP1915(m_CurCP(1), m_CurCP(2), m_CurCP(3), m_CurCP(4), mPA09, mPA26, m_LD18, mPA75, mCP43) = False Then
         MsgBox "新增進度檔【通知TW-SUPA期限】失敗！作業中斷！", vbCritical
         Exit Sub
      End If
      strCP09 = m_LD18
      'StartLetter
       EndLetter mET01, strCP09, strTmp, strUserNum
       
      'Added by Morgan 2021/6/11
      If stRefNat = "011" Then
         strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & mET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & _
            "','日本發明案','" & m_UsPatent & "')"
      Else
      'end 2021/6/11
         strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & mET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & _
            "','美國發明案','" & m_UsPatent & "')"
      End If
       strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & mET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & _
          "','回覆期限','" & m_RtnDate & "')"
       strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & mET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & _
          "','台灣案申請號','" & m_TwAppNo & "')"

        
       If Not ClsLawExecSQL(3, strTxt) Then
          MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
       End If
      '---------

      NowPrint strCP09, mET01, strTmp, False, strUserNum, 0, , , , , , , , , , , , m_LD18

      InsertQueryLog (RsTemp.RecordCount)
   Else
      InsertQueryLog (0)
      MsgBox "無符合條件之資料可列印 !", vbInformation
   End If
End Sub

'Added by Lydia 2015/04/20 新增進度檔【通知TW-SUPA期限】
Private Function PUB_AddCP1915(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, Optional pCountry As String, Optional pCustNo As String, Optional ByRef pNewCP09 As String, Optional pFcAgentNo As String, Optional pCP43 As String) As Boolean
   Dim stSQL As String, iR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCP09 As String, stCP13 As String, stCP12 As String
   
On Error GoTo ErrHnd
   stSQL = "select * from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp10='1915'"
   iR = 1
   Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      pNewCP09 = rsQuery("cp09")
   Else
      cnnConnection.BeginTrans
On Error GoTo ErrHnd2
      '收文號
      'Modified by Morgan 2018/7/26
      'stCP09 = AutoNo("C", 6)
      stCP09 = AutoNo("D", 6)
      'end 2018/7/26
      stCP13 = PUB_GetAKindSalesNo(pCP01, pCP02, pCP03, pCP04)
      stCP12 = GetSalesArea(stCP13)
      stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
         ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43 ) values ('" & pCP01 & "'" & _
         ",'" & pCP02 & "','" & pCP03 & "','" & pCP04 & "'," & strSrvDate(1) & _
         ",'" & stCP09 & "','1915','" & stCP12 & "'" & _
         ",'" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','" & pCP43 & "')"
      cnnConnection.Execute stSQL, intI
      '台灣案新增信函進度 '非直寄
      If pCountry = "000" Then
         PUB_AddLetterProgress stCP09, 0, True, "", False, pCustNo, "1915", pFcAgentNo
      End If
      pNewCP09 = stCP09
      cnnConnection.CommitTrans
   End If
   
   PUB_AddCP1915 = True
   Exit Function
   
ErrHnd2:
   cnnConnection.RollbackTrans
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function
'Added by Morgan 2020/8/6
'檢查是否核駁期限已逾期超過3月未辦的案件
Private Function ChkIsOverLimited(pNP01 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   'Modified by Morgan 2025/5/20 排除有部分准駁的案件 Ex:P-097468 --玲玲
   stSQL = "select a.* from caseprogress a,patent,caseprogress b,caseprogress c" & _
      " where a.cp09='" & pNP01 & "' and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa57 is null" & _
      " and b.cp01(+)=pa01 and b.cp02(+)=pa02 and b.cp03(+)=pa03 and b.cp04(+)=pa04 and b.cp05>a.cp27 and b.cp10='1002'" & _
      " and add_months(to_date(b.cp07,'yyyymmdd'),3)<sysdate" & _
      " and c.cp09(+)=b.cp43 and c.cp10 in ('501','503','504','505','506','507','508','804')" & _
      " and not exists(select * from nextprogress y where y.np01=b.cp09 and y.np06||''='Y')" & _
      " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1009')"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      ChkIsOverLimited = True
   End If
End Function

'Added by Lydia 2025/07/25
Private Sub SetDateTextBox()
   TxtDate(0) = m_DateTW1
   TxtDate(1) = m_DateTW2
   TxtDate(2) = m_Date1
   TxtDate(3) = m_Date2
   TxtDate(4) = m_FMPDate1
   TxtDate(5) = m_FMPDate2
End Sub

'Added by Lydia 2025/07/25 整批列印年費和實體審查：改用formdate控制例外通知的客戶
Private Sub Process_New()
Dim strCase As String
Dim strTmp As String, strTmp2 As String, rsTemp1 As New ADODB.Recordset, rsAD As New ADODB.Recordset
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strTxt(1 To 20) As String
Dim strFee As String, strPoint As String
Dim ii As Integer, jj As Integer
Dim strPA72NextYear As String
Dim strPA72Year As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP09 As String '收文號
Dim Prn As Printer
Dim blnSitu1 As Boolean '下一程序檔是否有本所案號+案件性質本所期限在前半年且是否續辦為N的資料
Dim strOldNP09 As String '半年前法定期限
Dim stSitu As String '定稿處理狀況
Dim idx As Integer
Dim bolPrint As Boolean, bolPrint1 As Boolean
Dim stMsg As String
Dim strNP07 As String, strNP08 As String, strNP09 As String 'Add by Morgan 2006/5/15
Dim strET03 As String
Dim strMaxFeeYear As String '最大可繳費年度
Dim bolDiscount As Boolean '是否可減免
Dim iCopy As Integer
Dim strNextYearFee As String '下次繳費金額
Dim strPA75 As String
Dim strNP23 As String 'Add by Morgan 2010//1/18 約定期限
Dim iPlusFee As Integer 'Added by Morgan 2013/1/8 服務費外加金額(目前專利處大對台年費+500)
Dim bolDualCaseUtility As Boolean, strInventionCaseNo As String, strInventionPA11 As String, strInventionPA77 As String 'Added by Morgan 2017/9/20 是否一案兩請新型案,發明案本所案號,發明案申請號,發明案彼所案號
Dim strPartA As String '台灣案SQL
Dim strPartB As String '非台灣案SQL
Dim strPA25 As String 'Added by Morgan 2023/6/5
Dim strSort As String 'Added by Morgan 2025/1/16
Dim intEx As Integer, m_ExCuList As String '例外通知(提早催通知):所有例外客戶編號(含關係企業)
Dim tmpArr1 As Variant
Dim tmpArrData As Variant '(0)代表X編號|(1)客戶編號(含關係企業)|(2)P案所限區間(大陸案)m_Date1|(3)m_Date2|(4)FMP案所限區間(大陸案)m_FMPDate1|(5)m_FMPDate2|(6)P案所限區間(台灣案)m_DateTW1|(7)m_DateTW2
Dim strCon1 As String

   bolPrint = False
   blnClkSure = False
                    
   '刪除暫存資料
   cnnConnection.Execute "Delete From R040303 Where ID='" & strUserNum & "'"
   '刪除接洽結案單暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   
   '取得例外通知的客戶相關控制資料
   intI = Pub_Getfrm040303ExceptNew("P", Text3.Text, IIf(Option3(0).Value = True, "1", "2"), "ALL", m_ExCuList, tmpArr1)
      
   strCon1 = strCon1 & " AND INSTR('" & m_ExCuList & "',PA26)=0 "
   
   '可同時跑兩種通知函
   For idx = 1 To 2
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
      If chkKind(idx).Value = 1 Then
         m_Select = idx
         '判斷通知函類別
         Select Case m_Select
            Case "1"
               'Modified by Morgan 2012/10/23 +香港維持費
               'Modified by Morgan 2024/11/1 +615補償期年費
               strCase = 年費 & "," & 維持費 & ",615"
               stMsg = "【年費 維持費 補償期年費】"
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(4)
            Case "2"
               strCase = 實體審查
               stMsg = "【實體審查】"
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(3)
         End Select
         
         '申請國家條件
         strTmp = ""
         Call ChangeSel(1) '將SQL改為對應NP  'Add by Lydia 2015/01/27 +fmp寰華控制sql (m_selarea) 'Memo by Lydia 2019/12/16 一併跑非大陸案
         
         '若為年費通知函, 不論是否閉卷皆要出現
         If m_Select = "1" Then
'----------台灣P案
            pub_QL05 = pub_QL05 & ";申請國家：台灣P案所限" & TxtDate(0) & "-" & TxtDate(1)
            
            'Modified by Morgan 2014/7/1  外層語法也要+AND NP06 IS NULL(ex: P-87410 103/10/08 重複)
            'Modified by Morgan 2017/12/11 台灣案不要排除F部門(P-97258原業務原誤掛外商人員造成期限沒催到)
            'Modified by Morgan 2018/10/3 NVL(PA22,'') -> PA22
            'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
            'Modified by Lydia 2019/12/16 改成共用句; Option2=>opt2 , text1(5)=>mdate1, text1(6)=>mdate2, strtmp => na01
            strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
               "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
               "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
               "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
               " NP02='P' and NP07 in (" & strCase & ",119) AND NP08 BETWEEN mdate1 AND mdate2 " & _
               " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
               " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
               "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
            '例外通知的客戶處理
            If m_ExCuList <> "" Then
               For intEx = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intEx)) <> "" Then
                     tmpArrData = Split(tmpArr1(intEx), "|")
                     '台灣P案
                     If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(6)) <> "" And Trim(tmpArrData(7)) <> "" Then
                          strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                             "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                             "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                             "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                             " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN tmpDateS-" & intEx & " AND tmpDateE-" & intEx & _
                             " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                             " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) na01 AND " & _
                             "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                             " AND INSTR('" & tmpArrData(1) & "',PA26)>0"
                     End If
                  End If
               Next intEx
            End If

            '處理台灣案SQL
            strPartA = strExc(0)
            strPartA = Replace(strPartA, "mdate1", DBDATE(TxtDate(0)))   '一般：台灣P案期限
            strPartA = Replace(strPartA, "mdate2", DBDATE(TxtDate(1)))   '一般：台灣P案期限
            strPartA = Replace(strPartA, "opt2", "")
            strPartA = Replace(strPartA, "na01", "AND PA09='000' ")
            '例外通知的客戶處理
            If m_ExCuList <> "" Then
               For intEx = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intEx)) <> "" Then
                     tmpArrData = Split(tmpArr1(intEx), "|")
                     If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(6)) <> "" And Trim(tmpArrData(7)) <> "" Then
                          'X38120000/X38120030碩天科技/寧遠縣碩寧電子，台灣案保持3個月
                          If InStr("X38120000,", Trim(tmpArrData(0))) > 0 And Trim(tmpArrData(0)) <> "" Then
                             strPartA = Replace(strPartA, "tmpDateS-" & intEx, DBDATE(TxtDate(0)))
                             strPartA = Replace(strPartA, "tmpDateE-" & intEx, DBDATE(TxtDate(1)))
                          Else
                             strPartA = Replace(strPartA, "tmpDateS-" & intEx, DBDATE(Trim(tmpArrData(6))))    '例外：台灣P案期限
                             strPartA = Replace(strPartA, "tmpDateE-" & intEx, DBDATE(Trim(tmpArrData(7))))    '例外：台灣P案期限
                          End If
                     End If
                  End If
               Next intEx
            End If
            
            'Modified by Lydia 2019/12/16 一併跑非大陸案
                pub_QL05 = pub_QL05 & ";申請國家：非台灣P案所限" & TxtDate(2) & "-" & TxtDate(3)
                pub_QL05 = pub_QL05 & ";申請國家：非台灣FMP案所限" & TxtDate(4) & "-" & TxtDate(5)
            'end 2019/12/16
            
'----------FMP案
            'Modified by Morgan 2013/6/5 修正案號會重複問題
            'Modified by Morgan 2014/7/1  外層語法也要+AND NP06 IS NULL
            'Modified by Morgan 2015/7/8 Y52323 法限抓4個月 -->Get605InformPeriod4NonTwCase要同步修改 Morgan 2017/1/13
            'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
            'Modified by Lydia 2019/12/16 改變欄位
            strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
               "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
               "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
               "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
               " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(TxtDate(4)) & " AND " & DBDATE(TxtDate(5)) & _
               " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
               " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
               " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
            'Y52323000杉村萬國特許法律事務所
            strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
               "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
               "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
               "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
               " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(TxtDate(4))) & " AND " & CompDate(1, 1, DBDATE(TxtDate(5))) & _
               " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
               " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
               " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
             '例外通知的客戶處理
             If m_ExCuList <> "" Then
                For intEx = 0 To UBound(tmpArr1)
                   If Trim(tmpArr1(intEx)) <> "" Then
                      'FMP案
                      tmpArrData = Split(tmpArr1(intEx), "|")
                      If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(4)) <> "" And Trim(tmpArrData(5)) <> "" Then
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & DBDATE(tmpArrData(4)) & " AND " & DBDATE(tmpArrData(5)) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND NVL(PA75,'Y')<>'Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & tmpArrData(1) & "',PA26)>0"
                        '一般語法已有抓取，例外不用
                        'strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,np23 from nextprogress WHERE " & _
                           "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07 in (" & strCase & ",119) AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(tmpArrData(4))) & " AND " & CompDate(1, 1, DBDATE(tmpArrData(5))) & _
                           " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) AND NP06 IS NULL),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND PA09<>'000' AND PA75='Y52323000'" & _
                           " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & tmpArrData(1) & "',PA26)>0"
                      End If
                   End If
                Next intEx
             End If

'----------大陸P案(非FMP)
             'Added by Lydia 2019/12/16 處理非台灣案SQL: 大陸P案(非FMP)
             strPartB = strExc(0)
             strPartB = Replace(strPartB, "mdate1", DBDATE(TxtDate(2))) '大陸P案期限
             strPartB = Replace(strPartB, "mdate2", DBDATE(TxtDate(3))) '大陸P案期限
             strPartB = Replace(strPartB, "opt2", " and substr(st03,1,1)<>'F' ") '限制非FMP案
             strPartB = Replace(strPartB, "na01", "AND PA09<>'000' ")    '限制非台灣案
             '例外通知的客戶處理
             If m_ExCuList <> "" Then
                For intEx = 0 To UBound(tmpArr1)
                   If Trim(tmpArr1(intEx)) <> "" Then
                      tmpArrData = Split(tmpArr1(intEx), "|")
                      If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(2)) <> "" And Trim(tmpArrData(3)) <> "" Then
                           strPartB = Replace(strPartB, "tmpDateS-" & intEx, DBDATE(Trim(tmpArrData(2))))    '例外：大陸P案期限
                           strPartB = Replace(strPartB, "tmpDateE-" & intEx, DBDATE(Trim(tmpArrData(3))))    '例外：大陸P案期限
                      End If
                   End If
                Next intEx
             End If
             
         Else '實體審查通知函
'**********實體審查通知函

'----------台灣P案
            'Modify by Morgan 2009/12/7 FMP案用法限條件且排除所限小於99/2/15的(已催過)
            pub_QL05 = pub_QL05 & ";申請國家：台灣P案所限" & TxtDate(0) & "-" & TxtDate(1)
            'Modified by Morgan 2017/12/11 台灣案不要排除F部門
            'Modified by Morgan 2018/10/3 NVL(PA22,'') -> PA22
            'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
            'Modified by Lydia 2019/12/16 改成共用句; Option2=>opt2 , text1(5)=>mdate1, text1(6)=>mdate2, strtmp => na01
            strExc(0) = "SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
               "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
               "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
               "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
               " NP02='P' and NP07=" & strCase & " AND NP08 BETWEEN mdate1 AND mdate2 " & _
               " AND NP06 IS NULL AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
               " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
               "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
               
            '例外通知的客戶處理
            If m_ExCuList <> "" Then
               For intEx = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intEx)) <> "" Then
                     tmpArrData = Split(tmpArr1(intEx), "|")
                     '非和碩案：用台灣P案
                     If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(6)) <> "" And Trim(tmpArrData(7)) <> "" Then
                        strExc(1) = ""
                        '和碩: 實審的通知時間則是提早為申請日＋１年，落在系統日期的1-10日或11-月底(20號)，與畫面不同的原因：是怕執行日期在10號採用畫面的日期可能會缺資料
                        If InStr("X70017000,", tmpArrData(0)) > 0 Then
                           If Val(Right(strSrvDate(1), 2)) <= 10 Then
                               strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "01" & " AND " & Mid(strSrvDate(1), 1, 6) & "10 "
                           Else
                               strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "11" & " AND " & Mid(strSrvDate(1), 1, 6) & "31 "
                           End If
                        End If
                        strExc(0) = strExc(0) & " UNION SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,PA22," & _
                           " NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'N' FMP,NP23,cu12,cu13 FROM " & _
                           " (SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           " (np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           " NP02='P' and NP07=" & strCase & " AND NP06 IS NULL " & IIf(strExc(1) <> "", "", "AND NP09 BETWEEN tmpDateS-" & intEx & " AND tmpDateE-" & intEx) & _
                           " AND st01(+)=NP10 opt2 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT" & _
                           " WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) na01 AND " & _
                           " SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & tmpArrData(1) & "',PA26)>0 AND NVL(PA10,0)>0" & strExc(1)
                     End If
                  End If
               Next intEx
            End If

            '處理台灣案SQL
            strPartA = strExc(0)
            strPartA = Replace(strPartA, "mdate1", DBDATE(TxtDate(0)))   '一般：台灣P案期限
            strPartA = Replace(strPartA, "mdate2", DBDATE(TxtDate(1)))   '一般：台灣P案期限
            strPartA = Replace(strPartA, "opt2", "")
            strPartA = Replace(strPartA, "na01", "AND PA09='000' ")
            '例外通知的客戶處理
            If m_ExCuList <> "" Then
               For intEx = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intEx)) <> "" Then
                     tmpArrData = Split(tmpArr1(intEx), "|")
                     If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(6)) <> "" And Trim(tmpArrData(7)) <> "" Then
                          If InStr("X38120000,", Trim(tmpArrData(0))) > 0 And Trim(tmpArrData(0)) <> "" Then
                             strPartA = Replace(strPartA, "tmpDateS-" & intEx, DBDATE(TxtDate(0)))
                             strPartA = Replace(strPartA, "tmpDateE-" & intEx, DBDATE(TxtDate(1)))
                          Else
                             strPartA = Replace(strPartA, "tmpDateS-" & intEx, DBDATE(Trim(tmpArrData(6))))    '例外：台灣P案期限
                             strPartA = Replace(strPartA, "tmpDateE-" & intEx, DBDATE(Trim(tmpArrData(7))))    '例外：台灣P案期限
                          End If
                     End If
                  End If
               Next intEx
            End If
            
            'Modified by Lydia 2019/12/16 一併跑非大陸案
                pub_QL05 = pub_QL05 & ";申請國家：非台灣P案所限" & TxtDate(2) & "-" & TxtDate(3)
                pub_QL05 = pub_QL05 & ";申請國家：非台灣FMP案所限" & TxtDate(4) & "-" & TxtDate(5)
            'end 2019/12/16
            
'----------FMP案
            'Modified by Lydia 2019/08/30 排除指定客戶的案件=>strCon1
            strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
               "NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
               "(SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
               "(np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
               " NP02='P' and NP07=" & strCase & " AND NP09 BETWEEN " & DBDATE(TxtDate(4)) & " AND " & DBDATE(TxtDate(5)) & _
               " AND NP06 IS NULL AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07)),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
               "NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
               "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & strCon1
            '例外通知的客戶處理
            If m_ExCuList <> "" Then
               For intEx = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intEx)) <> "" Then
                     tmpArrData = Split(tmpArr1(intEx), "|")
                     '非和碩案：用台灣P案
                     If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(6)) <> "" And Trim(tmpArrData(7)) <> "" Then
                        strExc(1) = ""
                        '和碩: 實審的通知時間則是提早為申請日＋１年，落在系統日期的1-10日或11-月底(20號)
                        If InStr("X70017000,", tmpArrData(0)) > 0 Then
                           If Val(Right(strSrvDate(1), 2)) <= 10 Then
                               strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "01" & " AND " & Mid(strSrvDate(1), 1, 6) & "10 "
                           Else
                               strExc(1) = " AND PA10+10000 BETWEEN " & Mid(strSrvDate(1), 1, 6) & "11" & " AND " & Mid(strSrvDate(1), 1, 6) & "31 "
                           End If
                        End If
                        strExc(0) = strExc(0) & " UNION ALL SELECT PA01,PA02,PA03,PA04,np08,np09,PA09,NVL(PA22,'')," & _
                           " NVL(PA26,'') pa26,PA85,CU64,FA31,np22,NP07,PA46,PA57,NP01,pa72,pa08,pa10,PA75,'Y' FMP,NP23,cu12,cu13 FROM " & _
                           " (SELECT np02,np03,np04,np05,NP08,NP09,np22,NP07,NP01,NP23 from nextprogress WHERE " & _
                           " (np02,np03,np04,np05,np07,np08||NP09) in (select np02,np03,np04,np05,np07,min(np08||NP09) FROM NEXTPROGRESS,staff WHERE " & _
                           "  NP02='P' and NP07=" & strCase & " AND NP06 IS NULL " & IIf(strExc(1) <> "", "", "AND NP09 BETWEEN tmpDateS-" & intEx & " AND tmpDateE-" & intEx) & _
                           " AND st01(+)=NP10 and substr(st03,1,1)='F' AND NP08>20100215 group by np02,np03,np04,np05,np07) " & _
                           " ),PATENT,CUSTOMER,FAGENT WHERE NP02=pa01(+) AND NP03=pa02(+) AND NP04=pa03(+) AND " & _
                           " NP05=pa04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND PA09<>'000' AND " & _
                           " SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & m_SelArea & _
                           " AND INSTR('" & tmpArrData(1) & "',PA26)>0 AND NVL(PA10,0)>0" & strExc(1)
                     End If
                  End If
               Next intEx
            End If

'----------大陸P案(非FMP)
             'Added by Lydia 2019/12/16 處理非台灣案SQL
             strPartB = strExc(0)
             strPartB = Replace(strPartB, "mdate1", DBDATE(TxtDate(2)))
             strPartB = Replace(strPartB, "mdate2", DBDATE(TxtDate(3)))
             strPartB = Replace(strPartB, "opt2", " and substr(st03,1,1)<>'F' ")
             strPartB = Replace(strPartB, "na01", "AND PA09<>'000' ")
             '例外通知的客戶處理
             If m_ExCuList <> "" Then
                For intEx = 0 To UBound(tmpArr1)
                   If Trim(tmpArr1(intEx)) <> "" Then
                      tmpArrData = Split(tmpArr1(intEx), "|")
                      '非和碩案：用台灣P案
                      If Trim(tmpArrData(1)) <> "" And Trim(tmpArrData(6)) <> "" And Trim(tmpArrData(7)) <> "" Then
                           strPartB = Replace(strPartB, "tmpDateS-" & intEx, DBDATE(Trim(tmpArrData(6))))    '例外：台灣P案期限
                           strPartB = Replace(strPartB, "tmpDateE-" & intEx, DBDATE(Trim(tmpArrData(7))))    '例外：台灣P案期限
                      End If
                   End If
                Next intEx
             End If
         End If
         
         'Added by  Lydia 2019/12/16 組合SQL
         strExc(0) = strPartA & " Union " & strPartB
         
         'Modified by Morgan 2013/6/26
         'strExc(0) = strExc(0) & " ORDER BY PA09,PA01,PA02,PA03,PA04"
         'Memo by Morgan 2015/9/1 整批列印定稿有另外控制列印順序(同接洽人一起,另年費逾期補繳通知也有)
         'Added by Morgan 2018/10/3 配合調整非FMP非台灣的年費與實審期限的所限,要剔除已整批催過的期限(過渡期避免重複催)
         'Modified by Morgan 2019/9/16 只跑期限表時不必剔除已整批催過的期限
         'Removed by Morgan 2021/8/16 過渡期早過，取消此檢查以避免漏催(曾發生改不續辦的原期限管制下次期限而沒催到)--8/16 有跟玲玲確認
         'If Check2.Value = 0 Then
         '   strExc(0) = "SELECT * FROM (" & strExc(0) & ") WHERE NOT EXISTS(select * from caseprogress y,letterprogress z" & _
         '      " where y.cp43=np01 and y.cp30=np22 and y.cp10='1913' and z.lp01(+)=y.cp09 and z.lp32='Y')"
         'End If
         'end 2021/8/16
         'end 2019/9/16
         'end 2018/10/3
         
         'Modified by Morgan 2025/1/16 +PID,排序語法改放變數(後面要用)
         'strExc(0) = strExc(0) & " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
         strExc(0) = "SELECT X.*,''PID FROM (" & strExc(0) & ") X"
         strSort = " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
         strExc(0) = strExc(0) & strSort
         'end 2025/1/16
         
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         
         'Added by Morgan 2025/1/16
         If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
            Combo1.Tag = ""
            Set rsQuery = PUB_CreateRecordset(rsTemp1, , , 300, Me.Name, mSeqNo)
            With rsQuery
               .MoveFirst
               Do While Not .EOF
                  .Fields("PID") = PUB_GetPHandler(.Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04"))
                  .MoveNext
               Loop
               .UpdateBatch
               
               stVTBX = "select R001 as " & .Fields(0).Name
               For intI = 2 To .Fields.Count
                  stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
               Next
               stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
            End With
            strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "'" & strSort
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strSql)
            Combo1.Tag = Combo1
         End If
         'end 2025/1/16
            
         If intI = 1 Then
            With rsTemp1
            InsertQueryLog (.RecordCount)
            Do While Not .EOF
               Erase strTxt
               strReceiveNo = .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)
               strNP07 = "" & .Fields("NP07")
               strNP08 = "" & .Fields("NP08")
               strNP09 = "" & .Fields("NP09")
                strNP23 = "" & .Fields("NP23")
               
               m_PA09 = "" & .Fields("pa09").Value
               m_PA46 = "" & .Fields("pa46").Value
               m_strPA08 = "" & .Fields("PA08")
               m_strPA10 = "" & .Fields("PA10")
               m_PA26 = "" & .Fields("pa26")
               
               strPA75 = "" & .Fields("PA75")
               strTmp2 = ""
               If .Fields("FMP") = "Y" Then
                  m_bolFMP = True
                  iCopy = 1
               Else
                  m_bolFMP = False
                  iCopy = 0
               End If
               
               'Modify by Morgan 2006/4/13 加R04030306欄位存類別
               'Modify by Morgan 2006/5/15 119進入國家階段 類別放3
               'Modified by Morgan 2020/8/6 指定欄位名稱
               'Modified by Morgan 2024/11/6 615補償期年費 類別放4
               If strNP07 = "119" Then
                  cnnConnection.Execute "Insert Into R040303(R04030301,R04030302,R04030303,R04030304,R04030305,ID,R04030306) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "','" & .Fields(3).Value & "','" & .Fields(4).Value & "','" & strUserNum & "','3')"
               ElseIf strNP07 = "615" Then
                  cnnConnection.Execute "Insert Into R040303(R04030301,R04030302,R04030303,R04030304,R04030305,ID,R04030306) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "','" & .Fields(3).Value & "','" & .Fields(4).Value & "','" & strUserNum & "','4')"
               Else
                  cnnConnection.Execute "Insert Into R040303(R04030301,R04030302,R04030303,R04030304,R04030305,ID,R04030306) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "','" & .Fields(3).Value & "','" & .Fields(4).Value & "','" & strUserNum & "','" & m_Select & "')"
               End If
            
               'Add by Morgan 2005/5/16
               m_CurCP(1) = .Fields(0): m_CurCP(2) = .Fields(1): m_CurCP(3) = .Fields(2): m_CurCP(4) = .Fields(3)
               m_NP22 = .Fields("np22"): m_iDiscount = 0
               
               'Added by Morgan 2020/8/6
               '年費通知加檢查是否核駁期限已逾期超過3月未辦的案件
               If m_Select = "1" Then
                  If ChkIsOverLimited(.Fields("NP01")) = True Then
                     strSql = "update R040303 set R04030307='X' where R04030301='" & m_CurCP(1) & "' and R04030302='" & m_CurCP(2) & "' and R04030303='" & m_CurCP(3) & "' and R04030304='" & m_CurCP(4) & "' and R04030306='1'"
                     cnnConnection.Execute strSql, intI
                     GoTo NoLetter
                  End If
               End If
               'end 2020/8/6
               
               '控制只印地址條or期限表
               If Check1.Value = vbChecked Or Check2.Value = vbChecked Then GoTo NoLetter
               
               '閉卷的都不印通知信但要印在清單上
               If .Fields("PA57") = "Y" Then
                  GoTo NoLetter
               Else
                  'Modified by Lydia 2019/12/16 改判斷PA09
                  'If Option2(0).Value = True Then 'Add by Morgan 2010/11/30 非台灣沒閉卷的都要通知--991126請作單
                  If "" & .Fields("PA09") Then
                     If CheckCPExists(m_CurCP) = True Then GoTo NoLetter
                  End If
               End If
               
               '收文號
               strCP09 = "" & .Fields("NP01")
  
               'Added by Morgan 2013/8/7
               'Modified by Morgan 2014/6/12 +申請國家,配合定稿轉pdf要有收文號改先新增進度
               'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
               'Modified by Morgan 2016/11/8 +傳是否大宗發文(pbolBulk=True)
               If PUB_AddCP1913(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04"), .Fields("NP08"), .Fields("NP09"), .Fields("NP01"), .Fields("NP22"), m_PA09, m_PA26, m_LD18, strPA75, , , True) = False Then
                  MsgBox "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
                  Exit Sub
               End If
               'end 2013/8/7
               
               'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
               If m_Select = "1" Or m_Select = "2" Then
                  '實體審查
                  If m_Select <> "1" Then

                     'Modified by Lydia 2015/01/07 採共用模組
                     strFee = PUB_GetYF0607(m_PA09, m_strPA08, m_PA26, "416", "1", "1", "1")
                     '申請國家為大陸,是否為PCT案件為"Y",則定稿之案件性質為06,否則為05
                     'Modify by Morgan 2006/5/18 加PCT
                     'Modify by Morgan 2009/7/9 +澳門044
                     If .Fields(6) = 大陸國家代號 Or .Fields(6) = "056" Or .Fields(6) = "044" Then
                        
                        If .Fields(6) = "056" Then
                           strTmp = "14"
                        'Add by Morgan 2009/7/9  澳門
                        ElseIf .Fields(6) = "044" Then
                           strTmp = "22"
                        Else
                           strTmp = IIf(.Fields("PA46").Value = "Y", "06", "05")
                        End If
                        '刪除定稿暫存資料
                        EndLetter ET01, strCP09, strTmp, strUserNum
                        '新增定稿暫存資料
                        strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                        strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','法定期限'," & CNULL(.Fields(5)) & ")"
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
                           
                        'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                        strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "416", "1", "1", "1")
                        strPoint = Round(Val(strPoint) / 1000, 1)
                        PUB_UpdateLP2930 m_LD18, strFee, strPoint
                        'end 2015/8/28
                        
                        'Add by Morgan 2005/11/16
                        strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','下一程序','416')"
                        'Add by Morgan 2009/7/9
                        strExc(0) = Pub_Get416Period(.Fields("PA08"), .Fields("PA09"))
                        strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & strCP09 & "','" & strTmp & "','" & strUserNum & "','提實審期限','" & strExc(0) & "')"
                           
                        If Not ClsLawExecSQL(5, strTxt) Then
                            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                        End If
                        NowPrint strCP09, ET01, strTmp, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                        'Add by Morgan 2009/12/7
                        If m_bolFMP Then
                           strUserNum = strFMPNum
                           strTmp2 = "51"
                           EndLetter ET01, strCP09, strTmp2, strUserNum
                           strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                           strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                           If m_PA46 = "Y" Then
                              strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & _
                                 "','PCT案','♀')"
                           Else
                              strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & strCP09 & "','" & strTmp2 & "','" & strUserNum & _
                                 "','非PCT案','♀')"
                           End If
                           If Not ClsLawExecSQL(3, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strTmp2, False, strUserNum
                           strUserNum = strUser1Num
                        End If
                        
                     '申請國家為台灣,定稿之案件性質為07
                     ElseIf .Fields(6) = 台灣國家代號 Then
                        
                        '大-->台 催實體審查定稿定稿 20080916 ADD BY TONI
                        If PUB_CheckCuNation(rsTemp1.Fields("pa26"), rsTemp1.Fields("pa01"), rsTemp1.Fields("pa02"), rsTemp1.Fields("pa03"), rsTemp1.Fields("pa04")) = "1" Then
                              strET03 = "20"
                              '刪除定稿暫存資料
                              EndLetter ET01, strCP09, strET03, strUserNum
                              '新增定稿暫存資料
                              strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                              strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(.Fields(5)) & ")"
                              strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"

                              strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','416')"
                              If Not ClsLawExecSQL(4, strTxt) Then
                                  MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                              End If
                              NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
    
                        Else
                           strET03 = "07"
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(.Fields(5)) & ")"
                           strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"

                           strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','416')"
                           If Not ClsLawExecSQL(4, strTxt) Then
                               MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                        End If
                     End If
                  'Add by Morgan 2006/5/15
                  ElseIf strNP07 = "119" Then
                     strET03 = "13"
                     '刪除定稿暫存資料
                     EndLetter ET01, strCP09, strET03, strUserNum
                     '新增定稿暫存資料
                     strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(strNP08) & ")"
                     strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
                     strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                     "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序'," & CNULL(strNP07) & ")"
                    
                     If Not ClsLawExecSQL(3, strTxt) Then
                         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                     End If
                     NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                  
                  'Added by Morgan 2024/11/6
                  ElseIf strNP07 = "615" Then
                     strET03 = "23"
                     '刪除定稿暫存資料
                     EndLetter ET01, strCP09, strET03, strUserNum
                     '新增定稿暫存資料
                     strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(strNP08) & ")"
                     strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strNP09) & ")"
                     strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序'," & CNULL(strNP07) & ")"
                     strExc(1) = ""
                     If PUB_GetCNExtDays(m_CurCP(), , intI) Then
                        If intI > 0 Then strExc(1) = intI
                     End If
                     strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','補償天數'," & CNULL(strExc(1)) & ")"
                     strFee = PUB_GetCN615Fee(m_CurCP())
                     strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用'," & CNULL(strFee) & ")"
                     If Not ClsLawExecSQL(5, strTxt) Then
                         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                     End If
                     NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                     
                     'Added by Morgan 2025/3/10
                     If m_bolFMP Then
                        strUserNum = strFMPNum
                        m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&615"
                        strTmp2 = "53"
                        EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                        strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                        strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                        strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strFee & "')"
                        strExc(1) = PUB_GetUSXRate
                        strExc(2) = ""
                        If Val(strExc(1)) <> 0 Then
                           strExc(2) = Fix(strFee / Val(strExc(1)))
                        End If
                        strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                                 
                        If Not ClsLawExecSQL(4, strTxt) Then
                            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                        End If
                        NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                        strUserNum = strUser1Num
                     End If
                     'end 2025/3/10
                     
                  'end 2024/11/6
                  '年費
                  Else
                     '本所期限是否已逾期且未超過7個月
                     blnSitu1 = False
                     StrSQLa = "Select * From Nextprogress Where " & ChgNextProgress(strReceiveNo) & " And NP07=" & 年費 & " And NP06='N' and NP08>0 AND NP09>0"
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        Do While Not rsA.EOF
                           If IsNull(rsA("NP08").Value) = False Then
                              If rsA("NP08").Value < strSrvDate(1) And DateDiff("m", ChangeWStringToWDateString(rsA("NP08").Value), ChangeWStringToWDateString(.Fields(4))) <= 7 Then
                                 strOldNP09 = "" & rsA("NP09").Value
                                 blnSitu1 = True
                                 Exit Do
                              End If
                           End If
                           rsA.MoveNext
                         Loop
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     
                     '補繳期限(有不續辦,該本所期限已過,期限差7個月內)
                     If blnSitu1 = True Then
                        If "" & .Fields(6).Value = "020" Then
                           'Added by Morgan 2015/8/28 非台灣信函進度要存報價(逾期原定稿只有點數)
                           strPA72NextYear = getPA72NextYear(m_CurCP(1), m_CurCP(2), m_CurCP(3), m_CurCP(4), , , strPA25)
                           If strPA72NextYear <> "" Then
                              strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                              strPoint = Round(Val(strPoint) / 1000, 1)
                           Else
                              strPoint = ""
                           End If
                           PUB_UpdateLP2930 m_LD18, "", strPoint
                           'end 2015/8/28
                           
                           strET03 = "09"
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           'Add by Morgan 2009/10/14
                           '98/10/1以後的一案兩請案,新型年費定稿加提醒
                           If m_PA09 = "020" And m_strPA08 = "2" And Val(m_strPA10) >= 20091001 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 'Added by Morgan 2012/8/30
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 'end 2012/8/30
                                 End If
                              End If
                           End If
                           
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','半年前法定期限'," & CNULL(strOldNP09) & ")"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           strPA72Year = GetNowNP09(strReceiveNo)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1

                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           ii = ii + 1
                           
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                               MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                           
                        ElseIf "" & .Fields(6).Value = "000" Then
                           strET03 = "08"
                        
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           'Added by Morgan 2012/9/21
                           '102新法一案兩請案,新型年費定稿加提醒
                           If m_strPA08 = "2" And Val(m_strPA10) >= 20130101 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 '未核准
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 End If
                              End If
                           End If
                           'end 2012/9/21
                           
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','半年前法定期限'," & CNULL(strOldNP09) & ")"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           strPA72Year = GetNowNP09(strReceiveNo)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                           "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/5
                           If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                              If strPA25 < CompDate(1, 6, strPA72Year) Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                                 ii = ii + 1
                              End If
                           End If
                           'end 2023/6/5

                           'Modified by Morgan 2022/9/1 不可傳入606否則回覆單會多帶
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605')"
                           'end 2022/9/1
                           ii = ii + 1
                           
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                           
                        End If
                            
                     '正常期限
                     Else
                        '大陸
                        '大陸案若系統日>法定期限, 則定稿之案件性質為03
                        'Modify by Morgan 2008/3/20 +澳門(044)
                        If CheckStr(.Fields(6)) = "020" Or CheckStr(.Fields(6)) = "044" Then
                           '取得下次繳費年度
                           strPA72NextYear = getPA72NextYear(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, , m_bFirstYear)
                           If CheckStr(.Fields(6)) = "044" Then
                              'Add by Morgan 2008/5/7 +繳第一次年費(無繳費記錄)
                              If m_bFirstYear = True Then
                                 strET03 = "18"
                              Else
                                 strET03 = "17"
                              End If
                           Else
                              'Modify By Sindy 2009/05/22 改定稿格式
                              'strET03 = IIf(strSrvDate(1) > "" & .Fields(5).Value, "03", "02")
                              strET03 = IIf(strSrvDate(1) > "" & .Fields(5).Value, "03", "21")
                           End If
                           
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           
                           'Add by Morgan 2010/1/18 FMP約定期限
                           If m_bolFMP Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','約定期限','" & strNP23 & "')"
                              ii = ii + 1
                           End If
                           
                           '計算該年年費屆滿日期strPA72Year
                           strPA72Year = getPA72Year(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strPA25)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                                                     
                           strNextYearFee = ""
                           If strPA72NextYear <> "" Then
                              'Modified by Lydia 2015/01/07 採共用模組
                              strNextYearFee = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','費用','" & strNextYearFee & "')"
                              ii = ii + 1
                           End If
                           
                           'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                           If strPA72NextYear <> "" Then
                              strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1")
                              strPoint = Round(Val(strPoint) / 1000, 1)
                           Else
                              strPoint = ""
                           End If
                           PUB_UpdateLP2930 m_LD18, strNextYearFee, strPoint
                           'end 2015/8/28
                              
                           'Add by Morgan 2005/11/16
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           ii = ii + 1
                           
                           'Add by Morgan 2009/10/7
                           '98/10/1以後的一案兩請案,新型年費定稿加提醒
                           '大陸一案兩請申請日輸入時必須互相檢查發明及新型之申請日是否為同一天,若不是,則show訊息告知user。
                           bolDualCaseUtility = False 'Added by Morgan 2017/9/20
                           If m_PA09 = "020" And m_strPA08 = "2" And Val(m_strPA10) >= 20091001 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14,pa11,PA77" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 'Added by Morgan 2017/9/20
                                 strInventionCaseNo = "" & RsTemp("C1")
                                 strInventionPA11 = "" & RsTemp("pa11")
                                 strInventionPA77 = "" & RsTemp("pa77")
                                 'end 2017/9/20
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    bolDualCaseUtility = True 'Added by Morgan 2017/9/20
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                       "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 'Added by Morgan 2012/8/30
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    bolDualCaseUtility = True 'Added by Morgan 2017/9/20
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 'end 2012/8/30
                                 End If
                              End If
                           End If
                           
                           
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                           'Add by Morgan 2009/12/7
                           If m_bolFMP Then
                              strUserNum = strFMPNum
                              
                              'Modified by Morgan 2014/8/20 FMP案有年費代理人,改傳本所案號+案件性質
                              'm_FMP_ET02 = strCP09
                              m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&605"
                              'end 2014/8/20
                              
                              'Removed by Morgan 2022/9/20 定稿已合併
                              '付款後辦案
                              'If CU72FA39("", strPA75) Then
                              '   strTmp2 = "53"
                              'Else
                              'end 2022/9/20
                              
                                 'Added by Morgan 2022/9/30
                                 If CheckStr(.Fields(6)) = "044" Then
                                    strTmp2 = "54"
                                 Else
                                 'end 2022/9/30
                                    strTmp2 = "52"
                                 End If
                                 
                              'End If 'Removed by Morgan 2022/9/20
                              
                              EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                              strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                              strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                              If m_PA46 = "Y" Then
                                 strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & _
                                    "','PCT案','♀')"
                              Else
                                 strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & _
                                    "','非PCT案','♀')"
                              End If
                              strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','下次年費年度','" & strPA72NextYear & "')"
                              strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strNextYearFee & "')"
                              strExc(1) = PUB_GetUSXRate
                              strExc(2) = ""
                              If Val(strExc(1)) <> 0 Then
                                 strExc(2) = Fix(strNextYearFee / Val(strExc(1)))
                              End If
                              strTxt(6) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                             
                              ii = 6
                              'Added by Morgan 2017/9/20
                              If bolDualCaseUtility = True Then
                                 strTxt(7) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','一案兩請新型案要印','♀')"
                                 strTxt(8) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案本所案號','" & strInventionCaseNo & "')"
                                 strTxt(9) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案申請號','" & ChgSQL(strInventionPA11) & "')"
                                 strTxt(10) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','發明案彼所案號','" & ChgSQL(strInventionPA77) & "')"
                                 ii = 10
                              End If
                              'end 2017/9/20
                                    
                              If Not ClsLawExecSQL(ii, strTxt) Then
                                 MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                              End If
                              NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                              strUserNum = strUser1Num
                           End If
                        'Add by Morgan 2004/5/14
                        '香港
                        ElseIf CheckStr(.Fields(6)) = "013" Then
                           '取得已繳費年度及專利種類
                           strPA72NextYear = getNextPayYear(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strPA72Year, strPA25)
                           
                           Select Case m_strPA08
                              Case "1" '標準專利
                                 stSitu = "12"
                              Case "2" '短期專利
                                 stSitu = "11"
                              Case "3" '外觀設計
                                 stSitu = "10"
                           End Select
                           
                           If strNP07 = 維持費 Then stSitu = "02" 'Added by Morgan 2012/10/23
                           
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, stSitu, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           
                          
                           If strPA72NextYear <> "" Then
                              'Modified by Lydia 2015/01/07 採共用模組
                              strNextYearFee = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, strNP07, strPA72NextYear, strPA72NextYear, "1")
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','費用','" & Val(strNextYearFee) & "')"
                              ii = ii + 1
                           End If
                           
                           'Added by Morgan 2015/8/28 非台灣信函進度要存報價
                           If strPA72NextYear <> "" Then
                              strPoint = PUB_GetYF06(m_PA09, m_strPA08, m_PA26, strNP07, strPA72NextYear, strPA72NextYear, "1")
                              strPoint = Round(Val(strPoint) / 1000, 1)
                           Else
                              strPoint = ""
                           End If
                           PUB_UpdateLP2930 m_LD18, strNextYearFee, strPoint
                           'end 2015/8/28
                           
                           'Add by Morgan 2005/11/16
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & stSitu & "','" & strUserNum & "','下一程序','605,607')"
                           ii = ii + 1
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, stSitu, False, strUserNum, 0, , , , iCopy, , , , , , , , m_LD18
                           
                           'Add by Morgan 2022/9/30
                            If m_bolFMP Then
                               strUserNum = strFMPNum
                               m_FMP_ET02 = m_CurCP(1) & m_CurCP(2) & m_CurCP(3) & m_CurCP(4) & "&605"
                               strTmp2 = "54"
                               EndLetter ET01, m_FMP_ET02, strTmp2, strUserNum
                               strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','本所期限','" & strNP08 & "')"
                               strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  "('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','法定期限','" & strNP09 & "')"
                               strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','下次年費年度','" & strPA72NextYear & "')"
                               strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','台幣報價','" & strNextYearFee & "')"
                               strExc(1) = PUB_GetUSXRate
                               strExc(2) = ""
                               If Val(strExc(1)) <> 0 Then
                                  strExc(2) = Fix(strNextYearFee / Val(strExc(1)))
                               End If
                               strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                  " ('" & ET01 & "','" & m_FMP_ET02 & "','" & strTmp2 & "','" & strUserNum & "','美金報價','" & strExc(2) & "')"
                               ii = 5
                               If Not ClsLawExecSQL(ii, strTxt) Then
                                  MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                               End If
                               NowPrint m_FMP_ET02, ET01, strTmp2, False, strUserNum
                               strUserNum = strUser1Num
                            End If
                            'end 2022/9/30
                            
                        '台灣
                        ElseIf CheckStr(.Fields(6)) = "000" Then
                           iPlusFee = 0 'Added by Morgan 2013/1/8
                           
                           '大-->台 催年費定稿 20090916 ADD BY TONI
                           If PUB_CheckCuNation(rsTemp1.Fields("pa26"), rsTemp1.Fields("pa01"), rsTemp1.Fields("pa02"), rsTemp1.Fields("pa03"), rsTemp1.Fields("pa04")) = "1" Then
                              strET03 = "19"
                              'Added by Morgan 2013/1/8 專利處大對台年費服務費+500 --郭雅娟 (113.7.12 接洽單也同步增加此規則 frm090801_new)
                              strExc(1) = PUB_GetStaffST15(PUB_GetAKindSalesNo(rsTemp1.Fields("PA01"), rsTemp1.Fields("PA02"), rsTemp1.Fields("PA03"), rsTemp1.Fields("PA04")), "1")
                              If Left(strExc(1), 2) = "P1" Then
                                 iPlusFee = 500
                              End If
                              'end 2013/1/8
                           Else
                              'Modify by Morgan 2008/1/7 一率改用新定稿
                              'strET03 = "01"
                              strET03 = "15"
                           End If
                           'END BY TONI
                           
                           '刪除定稿暫存資料
                           EndLetter ET01, strCP09, strET03, strUserNum
                           '新增定稿暫存資料
                           ii = 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','本所期限'," & CNULL(.Fields(4)) & ")"
                           ii = ii + 1
                           '計算該年年費屆滿日期strPA72Year
                           strPA72Year = getPA72Year(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strPA25)
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','法定期限'," & CNULL(strPA72Year) & ")"
                           ii = ii + 1
                           
                           'Added by Morgan 2023/6/1
                           If Val(strPA25) > 0 And Val(strPA72Year) > 0 Then
                              If strPA25 < CompDate(1, 6, strPA72Year) Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                    "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','即將屆滿','♀')"
                                 ii = ii + 1
                              End If
                           End If
                           'end 2023/6/1
                           
                           'Added by Morgan 2012/9/21
                           If m_strPA08 = "2" And Val(m_strPA10) >= 20130101 Then
                              strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14" & _
                                 " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_CurCP(1) & "' and cm02='" & m_CurCP(2) & "' and cm03='" & m_CurCP(3) & "' and cm04='" & m_CurCP(4) & "'" & _
                                 " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_CurCP(1) & "' and cm06='" & m_CurCP(2) & "' and cm07='" & m_CurCP(3) & "' and cm08='" & m_CurCP(4) & "') X" & _
                                 ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 '未核准
                                 If IsNull(RsTemp("pa16")) Or RsTemp("pa16") = "2" Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請要印','♀')"
                                    ii = ii + 1
                                 '已核准未公告
                                 ElseIf RsTemp("pa16") = "1" And IsNull(RsTemp("pa14")) Then
                                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                      "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','一案兩請發明已准未公告要印','♀')"
                                    ii = ii + 1
                                 End If
                              End If
                           End If
                           
                           If DBDATE(strPA72Year) >= 20120101 Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','102新法不印','♀')"
                              ii = ii + 1
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','102新法要印','♀')"
                              ii = ii + 1
                           End If
                           
                           '取得下次繳費年度
                           strPA72NextYear = getPA72NextYear(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value, strMaxFeeYear)
                           If strPA72NextYear <> "" Then

                              '服務費,規費
                              'Modified by Lydia 2015/01/07 採共用模組
                              strExc(0) = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1", strExc(1), strExc(2))
                              If strExc(0) = "0" Then strExc(1) = "": strExc(2) = ""
                              
                              If strExc(1) <> "" Then
                                 strExc(1) = Val(strExc(1)) + iPlusFee 'Added by Morgan 2013/1/8
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','服務費','" & strExc(1) & "')"
                                 ii = ii + 1
                              End If
                              
                              '年費是否可減免
                              If PUB_GetCaseDiscStat(.Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)) = "Y" Then
                                 bolDiscount = True
                              Else
                                 bolDiscount = False
                              End If
                           
                              If Val(strExc(2)) > 0 Then
                                 '減免
                                 If Val(strPA72NextYear) < 7 Then
                                    If bolDiscount = True Then
                                       If Val(strPA72NextYear) < 4 Then
                                          strExc(2) = Val(strExc(2)) - 800
                                       Else
                                          strExc(2) = Val(strExc(2)) - 1200
                                       End If
                                    End If
                                 End If
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','規費','" & strExc(2) & "')"
                                 ii = ii + 1
                              End If
   
                              strExc(3) = Val(strExc(1)) + Val(strExc(2))
                              If Val(strExc(3)) > 0 Then
                                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','費用','" & strExc(3) & "')"
                                 ii = ii + 1
                              End If
                              'end 2007/11/6

                              'Add by Morgan 2007/11/20 下兩年的費用也要印
                              strExc(5) = strExc(2) '規費累計
                              strExc(6) = strExc(3) '費用累計
                              'Added by Lydia 2024/08/15
                              Dim strBaseYear As String
                              strBaseYear = strPA72NextYear
                              'end 2024/08/15
                              For jj = 1 To 2
                                 strPA72NextYear = Val(strPA72NextYear) + 1
                                 If Val(strPA72NextYear) <= Val(strMaxFeeYear) Then
                                    'Modified by Lydia 2015/01/07 採共用模組
                                    strExc(0) = PUB_GetYF0607(.Fields("PA09").Value, m_strPA08, m_PA26, "605", strPA72NextYear, strPA72NextYear, "1", , strExc(2))
                                    If strExc(0) = "0" Then strExc(2) = ""
                                    
                                    If Val(strExc(2)) > 0 Then
                                       '減免
                                       If Val(strPA72NextYear) < 7 Then
                                          If bolDiscount = True Then
                                             If Val(strPA72NextYear) < 4 Then
                                                strExc(2) = Val(strExc(2)) - 800
                                             Else
                                                strExc(2) = Val(strExc(2)) - 1200
                                             End If
                                          End If
                                       End If
                                       strExc(5) = Val(strExc(5)) + Val(strExc(2))
                                       
                                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                          "','規費" & jj & "','" & strExc(5) & "')"
                                       ii = ii + 1
                                       '費用累計 --'Added by Lydia 2024/08/15 重抓服務費; ex.訊強電子 (惠州 )X41570060, P-117332
                                       strExc(0) = " select '1' as ord1, ys07 from patentyearspec where ys01='" & m_PA09 & "' and ys03='" & m_PA26 & "' and ys02='" & m_strPA08 & "' and ys04='605' and ys05='" & strBaseYear & "' and ys06='" & strPA72NextYear & "' " & _
                                                   " union select '2' as ord1, yf06 as ys07 from patentyearfee where yf01='" & m_PA09 & "' and yf03='" & m_PA26 & "' and yf02='" & m_strPA08 & "' and yf04='605' and yf05='" & strPA72NextYear & "' " & _
                                                   " order by 1"
                                       intI = 1
                                       Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
                                       If intI = 1 Then
                                          strExc(1) = Val("" & rsAD.Fields("ys07"))
                                       End If
                                       'end 2024/08/15
                                       strExc(6) = Val(strExc(1)) + Val(strExc(5))
                                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                          "','費用" & jj & "','" & strExc(6) & "')"
                                       ii = ii + 1
                                    End If
                                 Else
                                    Exit For
                                 End If
                              Next
                              
                              'end 2007/11/20
                           End If
                           
                           If PUB_ChkRefund(m_CurCP, m_lngRefund) = True Then
                              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','未退金額','" & m_lngRefund & "')"
                              ii = ii + 1
                           End If

                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & ET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "','下一程序','605,606')"
                           ii = ii + 1
    
                           If Not ClsLawExecSQL(ii - 1, strTxt) Then
                              MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                           End If
                           NowPrint strCP09, ET01, strET03, False, strUserNum, 0, , , , , , , , , , , , m_LD18
                           '台灣新增年費通知紀錄
                           Call UpdateAI
                           
                        End If
                     End If
                  End If
               End If
               
               '列印接洽結案單
               pub_AddressListSN = pub_AddressListSN + 1

               If m_iDiscount > 0 Then
                  PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, CheckStr(.Fields(12).Value), CheckStr(.Fields(0).Value), CheckStr(.Fields(1).Value), CheckStr(.Fields(2).Value), CheckStr(.Fields(3).Value), "1"
               Else
                  PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, CheckStr(.Fields(12).Value), CheckStr(.Fields(0).Value), CheckStr(.Fields(1).Value), CheckStr(.Fields(2).Value), CheckStr(.Fields(3).Value)
               End If
               
'跳過定稿
NoLetter:

               .MoveNext
            Loop
            End With
            bolPrint = True
         Else
            InsertQueryLog (0) 'Add By Sindy 2010/11/29
            MsgBox "無符合條件之" & stMsg & "資料可列印 !", vbInformation
         End If
      End If

   Next idx

   
   m_LD18 = "" 'Added by Morgan 2015/5/20
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
   'Add by Morgan 2007/9/4 主張國外優先權(未收文主張國外優先權,自撤)
   'Modified by Lydia 2019/12/16
   If chkKind(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";申請國家：台灣"
      pub_QL05 = pub_QL05 & ";" & Label1(2) & Label1(0)
      pub_QL05 = pub_QL05 & ";台灣P案所限" & TxtDate(0) & "-" & TxtDate(1)
   'end 2019/12/16
      stMsg = "【主張國外優先權】"
     'Add by Lydia 2015/01/27 +fmp寰華控制sql (m_selarea)
      Call ChangeSel(2) '將SQL改為對應PA

      'Modified by Morgan 2025/1/16 +PID
      strExc(0) = "select pa01,pa02,pa03,pa04,pa26,'' PID" & _
         " from patent a where pa01='P' and pa09='000'" & _
         " and to_char(add_months(to_date(pa10,'yyyymmdd'),9),'yyyymmdd') between " & DBDATE(TxtDate(0)) & " and " & DBDATE(TxtDate(1)) & _
         " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
         " and cp10 in ('106','413') and cp57 is null)" & _
         " and not ( exists(select * from casemap,patent b where cm05=a.pa01 and cm06=a.pa02 and cm07=a.pa03 and cm08=a.pa04" & _
         " and b.pa01(+)=cm01 and b.pa02(+)=cm02 and b.pa03(+)=cm03 and b.pa04(+)=cm04 and b.pa10<a.pa10)" & _
         " or exists(select * from casemap,patent b where cm01=a.pa01 and cm02=a.pa02 and cm03=a.pa03 and cm04=a.pa04" & _
         " and b.pa01(+)=cm05 and b.pa02(+)=cm06 and b.pa03(+)=cm07 and b.pa04(+)=cm08 and b.pa10<a.pa10))" & m_SelArea
      intI = 1
      Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
      
      'Added by Morgan 2025/1/16
      If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
         Combo1.Tag = ""
         Set rsQuery = PUB_CreateRecordset(rsTemp1, , , 300, Me.Name, mSeqNo)
         With rsQuery
            .MoveFirst
            Do While Not .EOF
               .Fields("PID") = PUB_GetPHandler(.Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04"))
               .MoveNext
            Loop
            .UpdateBatch
            
            stVTBX = "select R001 as " & .Fields(0).Name
            For intI = 2 To .Fields.Count
               stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
            Next
            stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
         End With
         strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strSql)
         Combo1.Tag = Combo1
      End If
      'end 2025/1/16
      
      If intI = 1 Then
         With rsTemp1
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/29
         Do While Not .EOF
            '只印地址條or期限表
            If Check1.Value = vbChecked Or Check2.Value = vbChecked Then GoTo NoLetter1
         
            If Me.Check1.Value = vbUnchecked Then
               strCP09 = .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
               strTmp = "16"
               NowPrint strCP09, ET01, strTmp, False, strUserNum, 0, , , , , , , , , , , , m_LD18
            End If
NoLetter1:

            .MoveNext
         Loop
         End With
         bolPrint1 = True
      Else
         InsertQueryLog (0) '
         MsgBox "無符合條件之待通知" & stMsg & "資料可列印 !", vbInformation
      End If
   End If
   
   If bolPrint = True Then
      '只印地址條or期限表時不印結案單
      If Me.Check1.Value = vbUnchecked And Me.Check2.Value = vbUnchecked Then
          '列印接洽結案單
         PUB_PrintCaseCloseSheet strUserNum
      End If
      '只印地址條時不印期限表
      If Me.Check1.Value = vbUnchecked Then
          MsgBox "請更換紙張，按確定後開始列印期限表!", vbOKOnly + vbInformation, "列印期限表"
          '列印繳年費/實體審查期限表
          Process1
          '列印專利權消滅清單
          Process2
      End If
   End If
   
   Set rsTemp1 = Nothing
   Set rsAD = Nothing
   
   If bolPrint = True Or bolPrint1 = True Then
      MsgBox "列印結束 !", vbInformation
   End If
End Sub

