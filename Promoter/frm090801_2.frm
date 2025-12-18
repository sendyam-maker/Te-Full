VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "下一程序期限資料"
   ClientHeight    =   4640
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4640
   ScaleWidth      =   8360
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   180
      TabIndex        =   18
      Top             =   1410
      Width           =   8085
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   1710
         MaxLength       =   7
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   7050
         MaxLength       =   2
         TabIndex        =   2
         Top             =   0
         Width           =   435
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   3690
         MaxLength       =   2
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   435
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   21
         Top             =   30
         Visible         =   0   'False
         Width           =   1635
         VariousPropertyBits=   27
         Caption         =   "延緩公告月數/日期："
         Size            =   "2884;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   12
         Left            =   6600
         TabIndex        =   19
         Top             =   30
         Visible         =   0   'False
         Width           =   1185
         VariousPropertyBits=   27
         Caption         =   "至 第　　　年"
         Size            =   "2090;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   4290
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   3750
         VariousPropertyBits=   27
         Caption         =   "繳費年度：第　　年"
         Size            =   "6615;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "含結案(&Y)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   2790
      TabIndex        =   4
      Top             =   30
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "不含已結案(&N)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   3
      Left            =   4140
      TabIndex        =   5
      Top             =   30
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7170
      TabIndex        =   7
      Top             =   30
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "帶入接洽記錄單(&O)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   0
      Left            =   5490
      TabIndex        =   6
      Top             =   30
      Width           =   1680
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2835
      Left            =   60
      TabIndex        =   3
      Top             =   1770
      Width           =   8265
      _ExtentX        =   14587
      _ExtentY        =   4992
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
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   8
      Left            =   1110
      TabIndex        =   17
      Top             =   510
      Width           =   1740
      VariousPropertyBits=   27
      Size            =   "3069;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   16
      Top             =   510
      Width           =   1740
      VariousPropertyBits=   27
      Size            =   "3069;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   4470
      TabIndex        =   15
      Top             =   510
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "申請案號："
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   14
      Top             =   1110
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "3387;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   4470
      TabIndex        =   13
      Top             =   1110
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "申請國家："
      Size            =   "1587;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   1110
      TabIndex        =   12
      Top             =   1110
      Width           =   3300
      VariousPropertyBits=   27
      Size            =   "5821;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   1110
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
      Index           =   1
      Left            =   1110
      TabIndex        =   10
      Top             =   810
      Width           =   7140
      VariousPropertyBits=   27
      Size            =   "12594;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   9
      Top             =   810
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
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   510
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
Attribute VB_Name = "frm090801_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/20 改成Form2.0 (grd1,Label1)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Public strNP02 As String
Public strNP03 As String
Public strNP04 As String
Public strNP05 As String
Public m_strCP10_1 As String
Public m_strCP10_2 As String
Public m_strCP10_3 As String
Public m_strCP10_4 As String
'Add By Sindy 2022/9/1
'Public m_strCP10Data(1 To 20) As String '案件性質先開陣列到20個
Public m_intCaseCnt As Integer '案件性質N個
'2022/9/1 END
Public m_strCP06 As String, m_strCP07 As String
Public m_PstrCP06 As String, m_PstrCP07 As String 'Added by Lydia 2025/02/20 P領證/年費的法限,所限
Public m_Note1 As String, m_Note2 As String
Public m_strGetNP01 As String 'Add By Sindy 2015/9/17
Public bolOK As Boolean 'True: 確定  False: 取消
Public strQType As String 'Add By Sindy 2015/4/1 0.一般 1.延期下一程序
'Public m_strNP23 As String 'Add By Sindy 2015/4/2 約定期限
Public m_strNP15 As String 'Added by Lydia 2021/04/15 帶回接洽單之下一程序備註

Dim m_row As Integer, i As Integer
Dim m_CurCP(1 To 4) As String '現在資料的本所號
Dim m_iDiscount As Integer '可減免退費金額
Dim m_iYear1 As Integer '減免退費起始年度
Dim m_iYear2 As Integer '減免退費終止年度
Dim m_MORE As String
Dim strYear As String '抓下次繳費年度
Dim m_Nexttimes As String '抓下次繳費次數
Dim strFeeType As String, strYF15 As String
Dim strKey(0 To 5) As String
Dim m_CaseFee(1 To 2) As String
Dim aryCaseFee As Variant
Dim m_iFixNo As Integer '修法次數
Dim m_Nation As String ' 國家代碼
Dim m_PA08 As String '專利種類
Dim bClose_Y As Boolean, bClose_N As Boolean
Dim m_CP01_Fee
Dim iR As Integer 'Modified by Lydia 2015/01/06
Dim m_TM21 As String '專用期間(起日) Add By Sindy 2015/4/1
Dim m_TM22 As String 'Added by Lydia 2018/06/26 專用期間(止日)
Dim m_NA77 As String '延展是否可延期 Add By Sindy 2015/4/1
Dim m_bUsIDS As Boolean 'Added by Morgan 2020/12/22 是否CFP美國案有勾選IDS
Dim m_PrevForm As Form '前一畫面


Public Sub SetParent(ByRef fm As Form, Optional ByVal mLength As Integer = 0, Optional ByVal bDesc As Boolean = False, Optional ByVal iTitle As String = "")
   Set m_PrevForm = fm
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim k As Integer
Dim bol_Chk As Boolean
Dim Cancel As Boolean
Dim dblDate As Double
'Dim varTemp As Variant 'Add By Sindy 2015/4/2
Dim strCP10List As String 'Added by Lydia 2020/11/19
Dim stIDSNote As String, bolDone As Boolean, bolSkip As Boolean 'Added by Morgan 2020/12/22
Dim stTot As String, stFee As String, stDot As String 'Added by Lydia 2020/12/29

   m_strCP06 = ""
   m_strCP07 = ""
'   m_strNP23 = "" 'Add By Sindy 2015/4/2
   m_strCP10_1 = ""
   m_strCP10_2 = ""
   m_strCP10_3 = ""
   m_strCP10_4 = ""
   'Add By Sindy 2022/9/1
   Erase m_strCP10Data
   m_intCaseCnt = 0
   '2022/9/1 END
   m_Note1 = ""
   m_strGetNP01 = "" 'Add By Sindy 2015/9/17
   m_Note2 = ""
   m_strNP15 = "" 'Added by Lydia 2021/04/15
   'Added by Lydia 2025/02/20
   m_PstrCP06 = ""
   m_PstrCP07 = ""
   'end 2025/02/20
   
   Select Case Index
   Case 0
      '檢查資料
      bol_Chk = False
      For i = 1 To GRD1.Rows - 1
         If Text22.Visible = True Then
            Cancel = False
            Text22_Validate Cancel
            If Cancel = True Then Exit Sub
         End If
         If Text20.Visible = True Then
            Cancel = False
            Text20_Validate Cancel
            If Cancel = True Then Exit Sub
         End If
         If Trim(GRD1.TextMatrix(i, 0)) = "V" And GRD1.RowHeight(i) > 0 Then
            If strQType = "0" Then 'Add By Sindy 2015/4/2 +if 非欲延期期限資料才需要檢查
               If Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "412" Then
                  If Trim(Text20.Text) = "" Then
                     MsgBox "請輸入延緩公告月數/日期！", vbInformation
                     Call Text20_GotFocus
                     Exit Sub
                  Else
                     If Len(Text20) = 1 Then
                        m_Note1 = "延緩公告：" & Text20.Text & "個月"
                     Else
                        m_Note1 = "延緩公告至 " & ChangeTStringToTDateString(Text20.Text)
                     End If
                  End If
               Else
                  'Modify By Sindy 2015/12/30 + (strNP02 = "P" Or strNP02 = "CFP")
                  'Modified by Lydia 2020/11/19 +CFP英國脫歐案 (strNP02 = "CFP" And Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "613")
                  If ((strNP02 = "P" Or strNP02 = "CFP") And (Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "605" Or Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "606" Or Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "607")) Or _
                     (strNP02 = "P" And Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "601") Or (strNP02 = "CFP" And Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "613") Then
                  '2015/12/30 END
                     m_Note2 = Trim(Label1(10).Caption)
                     'Added by Morgan 2020/12/8
                     'CFP延展費(英國)要用法限判斷，因為歐盟可能已經先繳
                     If strNP02 = "CFP" And Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "613" Then
                        m_Note2 = PUB_GetUKYr(DBDATE(GRD1.TextMatrix(i, 7)), strNP02, strNP03, strNP04, strNP05)
                     End If
                     'end 2020/12/8
                     
                     If Text22.Visible = True Then
                        If Trim(Text22.Text) = "" Then
                           'Text22.Text = strYear
                           MsgBox "請輸入至第幾年！", vbInformation
                           Call Text22_GotFocus
                           Exit Sub
                        Else
                           m_Note2 = m_Note2 & "至第 " & Text22.Text & " 年"
                        End If
                     End If
                  End If
                  '2011/8/18 add by sonia P非台灣的領證601年費605案起迄年度不同時提醒但仍可作業
                  '2011/10/26 modify by sonia 最後一年不必檢查 P-067266
                  'If strNP02 = "P" And m_Nation <> "000" And (Trim(GRD1.TextMatrix(i, 11)) = "601" Or Trim(GRD1.TextMatrix(i, 11)) = "605") And Val(Text21) <> Val(Text22) Then
                  If strNP02 = "P" And m_Nation <> "000" And (Trim(GRD1.TextMatrix(i, 11)) = "601" Or Trim(GRD1.TextMatrix(i, 11)) = "605") And Val(Text21) <> Val(Text22) And Text22.Visible = True Then
                     If MsgBox("非台灣案繳費年度起迄年度不同，是否確認一次繳多年？", vbYesNo + vbDefaultButton2) = vbNo Then
                        Call Text22_GotFocus
                        Exit Sub
                     End If
                  End If
                  '2011/8/18 end
                  
                  'Added by Morgan 2022/11/11
                  '台灣領證繳費年度超過專用期時提醒--蕭茹曣
                  If strNP02 = "P" And m_Nation = "000" And Trim(GRD1.TextMatrix(i, 11)) = "601" And Val(Text21) <> Val(Text22) And Text22.Visible = True Then
                     If GetMoneyDate(Val(m_PA08) + 10, m_Nation, strKey, , , m_TM22) = True Then     '抓專用期起止日
                        If m_TM22 <> "" Then
                           '最近公告日
                           If Len(Text20) > 3 Then '延緩公告日
                              strExc(1) = DBDATE(Text20)
                           Else
                              strExc(0) = Right(strSrvDate(1), 2)
                              If Val(strExc(0)) < "21" Then
                                 strExc(1) = Left(strSrvDate(1), 6) & "21"
                              ElseIf Val(strExc(0)) < "11" Then
                                 strExc(1) = Left(strSrvDate(1), 6) & "11"
                              Else
                                 strExc(1) = Left(CompDate(1, 1, strSrvDate(1)), 6) & "01"
                              End If
                              If Val(Text20) > 0 Then '延緩公告月數
                                 strExc(1) = CompDate(1, Val(Text20), strExc(1))
                              End If
                           End If
                           '繳費後有效期(最近公告日+繳費迄年-1年-1天)超過專用期時提醒
                           strExc(2) = CompDate(2, -1, CompDate(0, Val(Text22) - 1, strExc(1)))
                           If strExc(2) > m_TM22 Then
                              If MsgBox("繳費迄年可能超過專用期，請確認是否收文？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                  Call Text22_GotFocus
                                  Exit Sub
                              End If
                           End If
                        End If
                     End If
                  End If
                  'end 2022/11/11
                  
                  'Added by Lydia 2021/04/15 CFP和CFT英國脫歐委任代理之後續處理：將下一程序備註欄之「脫歐英國案代理人：Y…」印在案件說明處理事項欄
                  If ((strNP02 = "CFP" And Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "444") Or (strNP02 = "CFT" And Left(Trim(GRD1.TextMatrix(i, 4)), 3) = "710")) And _
                          Trim(GRD1.TextMatrix(i, 10)) <> "" And InStr(Trim(GRD1.TextMatrix(i, 10)), "脫歐英國案代理人：") > 0 Then
                      m_strNP15 = m_strNP15 & vbCrLf & Mid(Trim(GRD1.TextMatrix(i, 10)), InStr(Trim(GRD1.TextMatrix(i, 10)), "脫歐英國案代理人："), 18)
                  End If
                  'end 2021/04/15
               End If
            End If
            bol_Chk = True
         End If
      Next i
      If bol_Chk = False Then
         MsgBox "請勾選資料！", vbInformation
         Exit Sub
      End If
      
      If strQType = "0" Then 'Add By Sindy 2015/4/2 +if 非欲延期期限資料才需要檢查
         dblDate = 0
         For i = 1 To GRD1.Rows - 1
            If Trim(GRD1.TextMatrix(i, 0)) = "V" And GRD1.RowHeight(i) > 0 Then
               If dblDate <> 0 Then
                  If dblDate <> Val(DBDATE(Trim(GRD1.TextMatrix(i, 6)))) + Val(DBDATE(Trim(GRD1.TextMatrix(i, 7)))) Then
                     MsgBox "點選資料期限不同，帶最小期限日期至接洽記錄單！", vbInformation
                     Exit For
                  End If
               End If
               dblDate = Val(DBDATE(Trim(GRD1.TextMatrix(i, 6)))) + Val(DBDATE(Trim(GRD1.TextMatrix(i, 7))))
               strCP10List = strCP10List & "," & Trim(GRD1.TextMatrix(i, 4)) 'Added by Lydia 2020/11/19 記錄案件性質
            End If
         Next i
        'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：脫歐案性質需單獨收文 'Memo by Lydia 2020/12/01 若點選「延展費(英國)」/「延展(英國)」但未點選「委任代理人(CFP.444, CFT.710)」時提醒並自動勾選。
        strCP10List = Mid(strCP10List, 2)
        'Modified by Lydia 2020/12/01
        'If InStr(strCP10List, ",") > 0 Then
        If strCP10List <> "" And (strNP02 = "CFP" Or strNP02 = "CFT") Then
            strExc(1) = ""
            'Added by Lydia 2020/12/01
            If strNP02 = "CFP" And InStr(strCP10List, "613 ") > 0 And InStr(strCP10List, "607 ") > 0 Then
                 Call ClsPDGetCaseProperty(strNP02, "613", strExc(0))
                 strExc(1) = "613 " & strExc(0)
                 strExc(1) = "脫歐案性質〔" & strExc(1) & "〕需單獨收文！"
            ElseIf strNP02 = "CFP" And InStr(strCP10List, "444 ") > 0 And InStr(strCP10List, "607 ") > 0 Then
                 Call ClsPDGetCaseProperty(strNP02, "444", strExc(0))
                 strExc(1) = "444 " & strExc(0)
                 strExc(1) = "脫歐案性質〔" & strExc(1) & "〕需單獨收文！"
            ElseIf strNP02 = "CFT" And InStr(strCP10List, "110 ") > 0 And InStr(strCP10List, "102 ") > 0 Then
                 Call ClsPDGetCaseProperty(strNP02, "110", strExc(0))
                 strExc(1) = "110 " & strExc(0)
                 strExc(1) = "脫歐案性質〔" & strExc(1) & "〕需單獨收文！"
            ElseIf strNP02 = "CFT" And InStr(strCP10List, "710 ") > 0 And InStr(strCP10List, "102 ") > 0 Then
                 Call ClsPDGetCaseProperty(strNP02, "710", strExc(0))
                 strExc(1) = "710 " & strExc(0)
                 strExc(1) = "脫歐案性質〔" & strExc(1) & "〕需單獨收文！"
            'end 2020/12/01
            'Modified by Lydia 2020/12 若點選「延展費(英國)」/「延展(英國)」但未點選「委任代理人(CFP.444, CFT.710)」時提醒並自動勾選
            'If strNP02 = "CFP" And InStr(strCP10List, "613 ") > 0 Then
            ElseIf strNP02 = "CFP" And InStr(strCP10List, "613 ") > 0 And InStr(strCP10List, "444 ") = 0 Then
                 Call ClsPDGetCaseProperty(strNP02, "613", strExc(0))
                 strExc(1) = "613 " & strExc(0)
                 'Added by Lydia 2020/12/01
                 Call ClsPDGetCaseProperty(strNP02, "444", strExc(0))
                 strExc(1) = "〔" & strExc(1) & "〕需與〔" & "444 " & strExc(0) & "〕一併收文！"
                 'end 2020/12/01
            'Modified by Lydia 2020/12 若點選「延展費(英國)」/「延展(英國)」但未點選「委任代理人(CFP.444, CFT.710)」時提醒並自動勾選
            'ElseIf strNP02 = "CFT" And InStr(strCP10List, "110 ") > 0 Then
            ElseIf strNP02 = "CFT" And InStr(strCP10List, "110 ") > 0 And InStr(strCP10List, "710 ") = 0 Then
                 Call ClsPDGetCaseProperty(strNP02, "110", strExc(0))
                 strExc(1) = "110 " & strExc(0)
                 'Added by Lydia 2020/12/01
                 Call ClsPDGetCaseProperty(strNP02, "710", strExc(0))
                 strExc(1) = "〔" & strExc(1) & "〕需與〔" & "710 " & strExc(0) & "〕一併收文！"
                 'end 2020/12/01
            End If
            If strExc(1) <> "" Then
                'Modified by Lydia 2020/12/01
                'MsgBox "脫歐案性質〔" & strExc(1) & "〕需單獨收文！", vbCritical, "英國脫歐案管制"
                MsgBox strExc(1), vbCritical, "英國脫歐案管制"
                Exit Sub
            End If
        End If
        'end 2020/11/19
        
         dblDate = 0
         For i = 1 To GRD1.Rows - 1
            If Trim(GRD1.TextMatrix(i, 0)) = "V" And GRD1.RowHeight(i) > 0 Then
               dblDate = Val(DBDATE(Trim(GRD1.TextMatrix(i, 6)))) + Val(DBDATE(Trim(GRD1.TextMatrix(i, 7))))
               For k = 1 To GRD1.Rows - 1
                  If Trim(GRD1.TextMatrix(k, 0)) = "" And GRD1.RowHeight(k) > 0 Then
                     If dblDate > Val(DBDATE(Trim(GRD1.TextMatrix(k, 6)))) + Val(DBDATE(Trim(GRD1.TextMatrix(k, 7)))) Then
                        If MsgBox("尚有期限較近的資料未點選，是否重新點選？", vbYesNo + vbDefaultButton2) = vbYes Then
                           Exit Sub
                        End If
                        GoTo RunNext
                     End If
                  End If
               Next k
            End If
         Next i
      End If
      
RunNext:
      '回傳資料
      m_intCaseCnt = 0
      For i = 1 To GRD1.Rows - 1
         If Trim(GRD1.TextMatrix(i, 0)) = "V" And GRD1.RowHeight(i) > 0 Then
            'Added by Morgan 2020/12/22
            'CFP美國IDS只收文一道且要將備註帶到說明
            bolSkip = False
            If m_bUsIDS = True And Trim(Left(GRD1.TextMatrix(i, 4), 4)) = "214" Then
               If bolDone = False Then
                  bolDone = True
                  stIDSNote = "IDS相關案號:"
               Else
                  bolSkip = True
               End If
               stIDSNote = stIDSNote & vbCrLf & "　　" & GRD1.TextMatrix(i, 10)
            End If
            If bolSkip = False Then
            'end 2020/12/22
            
               m_intCaseCnt = m_intCaseCnt + 1
               If Val(m_strCP06) = 0 Or _
                  Val(ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 6)))) < Val(m_strCP06) Then
                  m_strCP06 = ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 6)))  '本所期限
                  m_strCP07 = ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 7)))  '法定期限
'                  m_strNP23 = ChangeWStringToTString(Trim(grd1.TextMatrix(i, 20))) '約定期限
                  'm_strGetNP01 = Trim(GRD1.TextMatrix(i, 13)) 'Add By Sindy 2015/9/15 NP01.總收文號
               End If
               
               'Added by Lydia 2025/02/20 P領證/年費的法限,所限; ex.P-124982同時收年費和代辦退費,期限抓到代辦退費,造成年費預設為逾期
               If strNP02 = "P" And InStr("601,605,", Trim(Left(GRD1.TextMatrix(i, 4), 4))) > 0 Then
                  If Val(m_PstrCP06) = 0 Or Val(ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 6)))) < Val(m_PstrCP06) Then
                     m_PstrCP06 = ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 6)))  '本所期限
                     m_PstrCP07 = ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 7)))  '法定期限
                  End If
               End If
               'end 2025/02/20
               
               'Modify By Sindy 2022/12/15
               'If strSrvDate(1) >= 接洽單電子收文啟用日 Then
               If UCase(TypeName(m_PrevForm)) = UCase("frm090801_New") Then
               '2022/12/15 ENd
                  m_strCP10Data(m_intCaseCnt) = Trim(GRD1.TextMatrix(i, 4))
                  'Add By Sindy 2015/4/2 欲延期期限
                  If m_intCaseCnt = 1 Then
                     If strQType = "1" Then
                        m_strGetNP01 = Trim(GRD1.TextMatrix(i, 13)) 'Add By Sindy 2015/9/4 NP01=延期案的相關總收文號
                        m_Note2 = "延期案件性質:" & m_strCP10Data(m_intCaseCnt)
      '                  varTemp = Split(m_strCP10_1, " ")
      '                  m_Note2 = "延期案件性質:" & varTemp(1)
                     End If
                     '2015/4/2 END
                  End If
                  'Added by Lydia 2020/12/29 CFP案件收文時自動帶入已報價之費用
                  If strQType = "0" Then
                      'Modified by Lydia 2021/01/05 改共用模組
                      'If CheckCFPtoFee(Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot) = True Then
                      'Modified by Morgan 2021/8/17 +參數提醒CFP領證是否應含其他費用
                      If PUB_CheckCFPtoFee(m_Nation, strNP02, Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot, True) = True Then
                         m_strCP10Data(m_intCaseCnt) = m_strCP10Data(m_intCaseCnt) & "," & Val(stTot) & "," & Val(stFee) & "," & Val(stDot)
                      End If
                  End If
                  'end 2020/12/29
               Else
               '2022/9/1 END
                  If m_intCaseCnt = 1 Then
                     m_strCP10_1 = Trim(GRD1.TextMatrix(i, 4))
                     'Add By Sindy 2015/4/2 欲延期期限
                     If strQType = "1" Then
                        m_strGetNP01 = Trim(GRD1.TextMatrix(i, 13)) 'Add By Sindy 2015/9/4 NP01=延期案的相關總收文號
                        m_Note2 = "延期案件性質:" & m_strCP10_1
      '                  varTemp = Split(m_strCP10_1, " ")
      '                  m_Note2 = "延期案件性質:" & varTemp(1)
                     End If
                     '2015/4/2 END
                     'Added by Lydia 2020/12/29 CFP案件收文時自動帶入已報價之費用
                     If strQType = "0" Then
                         'Modified by Lydia 2021/01/05 改共用模組
                         'If CheckCFPtoFee(Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot) = True Then
                         'Modified by Morgan 2021/8/17 +參數提醒CFP領證是否應含其他費用
                         If PUB_CheckCFPtoFee(m_Nation, strNP02, Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot, True) = True Then
                            m_strCP10_1 = m_strCP10_1 & "," & Val(stTot) & "," & Val(stFee) & "," & Val(stDot)
                         End If
                     End If
                     'end 2020/12/29
                  ElseIf m_intCaseCnt = 2 Then
                     m_strCP10_2 = Trim(GRD1.TextMatrix(i, 4))
                     'Added by Lydia 2020/12/29 CFP案件收文時自動帶入已報價之費用
                     If strQType = "0" Then
                         'Modified by Lydia 2021/01/05 改共用模組
                         'If CheckCFPtoFee(Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot) = True Then
                         'Modified by Morgan 2021/8/17 +參數提醒CFP領證是否應含其他費用
                         If PUB_CheckCFPtoFee(m_Nation, strNP02, Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot, True) = True Then
                            m_strCP10_2 = m_strCP10_2 & "," & Val(stTot) & "," & Val(stFee) & "," & Val(stDot)
                         End If
                     End If
                     'end 2020/12/29
                  ElseIf m_intCaseCnt = 3 Then
                     m_strCP10_3 = Trim(GRD1.TextMatrix(i, 4))
                     'Added by Lydia 2020/12/29 CFP案件收文時自動帶入已報價之費用
                     If strQType = "0" Then
                         'Modified by Lydia 2021/01/05 改共用模組
                         'If CheckCFPtoFee(Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot) = True Then
                         'Modified by Morgan 2021/8/17 +參數提醒CFP領證是否應含其他費用
                         If PUB_CheckCFPtoFee(m_Nation, strNP02, Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot, True) = True Then
                            m_strCP10_3 = m_strCP10_3 & "," & Val(stTot) & "," & Val(stFee) & "," & Val(stDot)
                         End If
                     End If
                     'end 2020/12/29
                  ElseIf m_intCaseCnt = 4 Then
                     m_strCP10_4 = Trim(GRD1.TextMatrix(i, 4))
                     'Added by Lydia 2020/12/29 CFP案件收文時自動帶入已報價之費用
                     If strQType = "0" Then
                         'Modified by Lydia 2021/01/05 改共用模組
                         'If CheckCFPtoFee(Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot) = True Then
                         'Modified by Morgan 2021/8/17 +參數提醒CFP領證是否應含其他費用
                         If PUB_CheckCFPtoFee(m_Nation, strNP02, Trim(GRD1.TextMatrix(i, 13)), Trim(GRD1.TextMatrix(i, 14)), Trim(GRD1.TextMatrix(i, 4)), Trim(GRD1.TextMatrix(i, 10)), stTot, stFee, stDot, True) = True Then
                            m_strCP10_4 = m_strCP10_4 & "," & Val(stTot) & "," & Val(stFee) & "," & Val(stDot)
                         End If
                     End If
                     'end 2020/12/29
                  Else
                     Exit For
                  End If
               End If
            End If 'Added by Morgan 2020/12/22
         End If
            
      Next i
      If stIDSNote <> "" Then m_Note1 = m_Note1 & stIDSNote  'Added by Morgan 2020/12/22
      
      bolOK = True
   Case 1 '回前畫面
      bolOK = False
      
   Case 2 '含結案
      For i = 1 To GRD1.Rows - 1
         GRD1.RowHeight(i) = 255
      Next i
      Exit Sub
      
   Case 3 '不含已結案
      For i = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(i, 5) = "Y" Then
            GRD1.RowHeight(i) = 0
            GRD1.row = i
            For k = 0 To GRD1.Cols - 1
               GRD1.col = k
               If GRD1.CellBackColor = &HFFC0C0 Then
                 GRD1.CellBackColor = &H80000018
                 GRD1.TextMatrix(i, 0) = ""
               End If
            Next k
         Else
            GRD1.RowHeight(i) = 255
         End If
      Next i
      Exit Sub
      
   End Select
   Me.Hide
End Sub

Public Function doQuery() As Boolean
Dim intRow As Integer
Dim strSpecial As String, strGetCP10 As String
Dim strPA10 As String 'Added by Morgan 2022/6/13
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass 'Add By Sindy 2014/5/26
   doQuery = False
   bClose_Y = False
   bClose_N = False
   '延緩公告
   Label1(9).Visible = False
   Text20.Visible = False
   Text20.Text = ""
   '年費
   Label1(10).Visible = False
   Label1(12).Visible = False
   Text21.Visible = False
   Text21.Text = ""
   Text22.Visible = False
   Text22.Text = ""
   Label1(7) = ""
   Label1(1) = ""
   Label1(3) = ""
   Label1(5) = ""
   m_Nation = ""
   m_PA08 = ""
   
   Label1(8) = strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05
   
   '讀取案件資料:
   'Modify By Sindy 2015/4/1 +TM21,NA77
   'Modified by Lydia 2018/06/26 +TM22
   'Modified by Morgan 2022/6/13 +TM11(PA10)
   strSql = "SELECT TM12,TM05||TM06||TM07,TM23||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,TM10,' ',TM21,TM22,NA77,TM11" & _
                " From Trademark, nation, Customer" & _
                " WHERE TM01='" & strNP02 & "' AND TM02='" & strNP03 & "' AND TM03='" & strNP04 & "' AND TM04='" & strNP05 & "'" & _
                " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
                " AND TM10=NA01(+)"
   'Modified by Lydia 2018/06/26 +PA25
   strSql = strSql & " Union " & _
                "SELECT PA11,PA05||PA06||PA07,PA26||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,PA09,PA08,PA24,PA25,NA77,PA10" & _
                " From Patent, nation, Customer" & _
                " WHERE PA01='" & strNP02 & "' AND PA02='" & strNP03 & "' AND PA03='" & strNP04 & "' AND PA04='" & strNP05 & "'" & _
                " AND SUBSTR(PA26,1,8)=CU01(+) AND decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)" & _
                " AND PA09=NA01(+)"
   'Modified by Lydia 2018/06/26 0,NA77=>0,0,NA77
   strSql = strSql & " Union " & _
                "SELECT '',LC05||LC06||LC07,LC11||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,LC15,' ',0,0,NA77,0" & _
                " From LawCase, nation, Customer" & _
                " WHERE LC01='" & strNP02 & "' AND LC02='" & strNP03 & "' AND LC03='" & strNP04 & "' AND LC04='" & strNP05 & "'" & _
                " AND SUBSTR(LC11,1,8)=CU01(+) AND decode(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+)" & _
                " AND LC15=NA01(+)"
   'Modified by Lydia 2018/06/26 ' ',' ',' ',0,= >' ',' ',' ',0,0
   strSql = strSql & " Union " & _
                "SELECT '',HC06,HC05||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),' ',' ',' ',0,0,'',0" & _
                " From HireCase, Customer" & _
                " WHERE HC01='" & strNP02 & "' AND HC02='" & strNP03 & "' AND HC03='" & strNP04 & "' AND HC04='" & strNP05 & "'" & _
                " AND SUBSTR(HC05,1,8)=CU01(+) AND decode(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Label1(7) = "" & Trim(RsTemp(0))
      Label1(1) = "" & Trim(RsTemp(1))
      Label1(3) = "" & Trim(RsTemp(2))
      Label1(5) = "" & Trim(RsTemp(3))
      m_Nation = "" & Trim(RsTemp(4))
      m_PA08 = "" & Trim(RsTemp(5))
      m_TM21 = "" & Trim(RsTemp(6)) '專用期間(起日)
      'Modified by Lydia 2108/06/26
      'm_NA77 = "" & Trim(RsTemp(7)) '延展是否可延期
      m_TM22 = "" & Trim(RsTemp.Fields("TM22")) '專用期間(止日)
      m_NA77 = "" & Trim(RsTemp.Fields("NA77")) '延展是否可延期
      strPA10 = "" & Trim(RsTemp.Fields("TM11")) 'Added by Morgan 2022/6/13
   End If
   
   'Add By Sindy 2015/4/1 查詢1.延期的下一程序
   '預設值 *****
   cmdOK(2).Visible = True
   cmdOK(3).Visible = True
   Frame1.Visible = True
   Me.Caption = "下一程序期限資料"
   '***** END
   If strQType = "1" Then
      cmdOK(2).Visible = False
      cmdOK(3).Visible = False
      Frame1.Visible = False
      Me.Caption = "欲延期期限資料"
      strSql = "SELECT ' ' AS V,decode(substr(cp09,1,1),'C',DECODE(cp05,'','',SUBSTR(cp05,1,4)-1911||'/'||SUBSTR(cp05,5,2)||'/'||SUBSTR(cp05,7,2)),'') as 來函收文日," & _
               "decode(substr(cp09,1,1),'C',decode('" & m_Nation & "','000',C2.cpm03,C2.cpm04),'') as 來函性質,decode(substr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||decode('" & m_Nation & "','000',C1.cpm03,C1.cpm04) as 下一程序,decode(np06,'N','Y','') as 結案," & _
               "DECODE(np08,'','',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2)) as 本所期限," & _
               "DECODE(np09,'','',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2)) as 法定期限," & _
               "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,rownum as sort,np01,np22,np08,np02,np03,np04,np05,np23" & _
               " FROM NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
               " WHERE NP02='" & strNP02 & "' AND NP03='" & strNP03 & "' AND NP04='" & strNP04 & "' AND NP05='" & strNP05 & "'" & _
               " and np01=cp09(+)" & _
               " and np10=st01(+)" & _
               " and np02=C1.cpm01(+) and np07=C1.cpm02(+)" & _
               " and cp01=C2.cpm01(+) and cp10=C2.cpm02(+)" & _
               " and NP06 is null" & strNpSqlOfNoSalesDuty
      If strNP02 = "P" Or strNP02 = "CFP" Then
         strSql = strSql & " and np07 in('107','204','205','206','501','804','424')"
      ElseIf strNP02 = "T" Then
         If m_Nation = "000" Then '台灣案
            strSql = strSql & " and np07 not in('102','716','717')"
         Else '大陸案都不可延期
            strSql = strSql & " and np07 is null"
         End If
      ElseIf strNP02 = "CFT" Then
         strSql = strSql & " and np07 in('" & IIf(Val(m_TM21) = 0, "105", "") & "','" & IIf(Trim(m_NA77) = "", "102", "") & "')"
      End If
      'Add By Sindy 2015/12/17 + 進度檔的已收文未發文未取消收文的資料
      'Modify By Sindy 2016/10/17 + 剔除C類來函,只抓AB類收文
      strSql = strSql & " union " & _
               "SELECT ' ' AS V,DECODE(c2.cp05,'','',SUBSTR(c2.cp05,1,4)-1911||'/'||SUBSTR(c2.cp05,5,2)||'/'||SUBSTR(c2.cp05,7,2)) as 來函收文日," & _
               "decode('" & m_Nation & "','000',m2.cpm03,m2.cpm04) as 來函性質,c1.cp43 as 來函總收文號,c1.cp10||' '||decode('" & m_Nation & "','000',m1.cpm03,m1.cpm04) as 下一程序,'' as 結案," & _
               "DECODE(c1.cp06,'','',SUBSTR(c1.cp06,1,4)-1911||'/'||SUBSTR(c1.cp06,5,2)||'/'||SUBSTR(c1.cp06,7,2)) as 本所期限," & _
               "DECODE(c1.cp07,'','',SUBSTR(c1.cp07,1,4)-1911||'/'||SUBSTR(c1.cp07,5,2)||'/'||SUBSTR(c1.cp07,7,2)) as 法定期限," & _
               "st02 As 智權人員, c1.cp40||c1.cp41||c1.cp42 As 相關人, c1.cp64 As 備註,c1.cp10 np07,rownum as sort,c1.cp09 np01,0 np22,c1.cp06 np08,c1.cp01 np02,c1.cp02 np03,c1.cp03 np04,c1.cp04 np05,0 np23" & _
               " FROM CaseProgress c1,Staff,CasePropertyMap m1,CaseProgress c2,CasePropertyMap m2" & _
               " WHERE c1.CP01='" & strNP02 & "' AND c1.CP02='" & strNP03 & "' AND c1.CP03='" & strNP04 & "' AND c1.CP04='" & strNP05 & "'" & _
               " and c1.cp27 is null and c1.cp57 is null" & _
               " and c1.cp13=st01(+)" & _
               " and c1.cp01=m1.cpm01(+) and c1.cp10=m1.cpm02(+)" & _
               " and c1.cp43=c2.cp09(+) and c1.cp09<'C'" & _
               " and c2.cp01=m2.cpm01(+) and c2.cp10=m2.cpm02(+)"
      If strNP02 = "P" Or strNP02 = "CFP" Then
         strSql = strSql & " and c1.cp10 in('107','204','205','206','501','804','424')"
      ElseIf strNP02 = "T" Then
         If m_Nation = "000" Then '台灣案
            'Modify By Sindy 2016/10/17 + 剔除303.延期
            strSql = strSql & " and c1.cp10 not in('102','716','717','303')"
         Else '大陸案都不可延期
            strSql = strSql & " and c1.cp10 is null"
         End If
      ElseIf strNP02 = "CFT" Then
         strSql = strSql & " and c1.cp10 in('" & IIf(Val(m_TM21) = 0, "105", "") & "','" & IIf(Trim(m_NA77) = "", "102", "") & "')"
      End If
      '2015/12/17 END
      strSql = strSql & " order by np08 asc,np01 asc,np02 asc,np03 asc,np04 asc,np05 asc"
      CheckOC3
      SetDataListWidth
      GRD1.Rows = 2
      GRD1.Clear
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            Set GRD1.Recordset = AdoRecordSet3.Clone
            '只有一筆資料,預設勾選
            If .RecordCount = 1 Then
               iR = 1
               Call grd1_SelChange
            Else
               iR = 0
            End If
            doQuery = True
         Else
            'MsgBox "無符合資料！", vbInformation
         End If
      End With
   Else
   '2015/4/1 END
      
      '讀取下一程序:查詢此案未收文或不辦案件(剔除程序管控之案件性質)
      'Add By Sindy 2012/6/29
      If m_Nation = "238" Then '馬德里
         'Modify By Sindy 2015/4/2 +,np23
         strSql = "SELECT ' ' AS V,decode(substr(cp09,1,1),'C',DECODE(cp05,'','',SUBSTR(cp05,1,4)-1911||'/'||SUBSTR(cp05,5,2)||'/'||SUBSTR(cp05,7,2)),'') as 來函收文日," & _
                  "decode(substr(cp09,1,1),'C',C2.cpm04,'') as 來函性質,decode(substr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||C1.cpm04 as 下一程序,decode(np06,'N','Y','') as 結案," & _
                  "DECODE(np08,'','',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2)) as 本所期限," & _
                  "DECODE(np09,'','',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2)) as 法定期限," & _
                  "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,rownum as sort,np01,np22,np08,np02,np03,np04,np05,np23" & _
                  " FROM NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
                  " WHERE NP02='" & strNP02 & "' AND substr(NP03,1,5)='" & Left(strNP03, 5) & "'" & _
                  " and np01=cp09(+)" & _
                  " and np10=st01(+)" & _
                  " and np02=C1.cpm01(+) and np07=C1.cpm02(+)" & _
                  " and cp01=C2.cpm01(+) and cp10=C2.cpm02(+)" & _
                  " and (NP06 is null OR NP06='N')" & strNpSqlOfNoSalesDuty & _
                  " order by np08 asc,np01 asc,np02 asc,np03 asc,np04 asc,np05 asc"
      '2012/6/29 End
      Else
         'Modify By Sindy 2012/6/29 +,np08,np02,np03,np04,np05及order by np08 asc改為order by np08 asc,np01 asc,np02 asc,np03 asc,np04 asc,np05 asc
         'Modify By Sindy 2015/4/2 +,np23
         strSql = "SELECT ' ' AS V,decode(substr(cp09,1,1),'C',DECODE(cp05,'','',SUBSTR(cp05,1,4)-1911||'/'||SUBSTR(cp05,5,2)||'/'||SUBSTR(cp05,7,2)),'') as 來函收文日," & _
                  "decode(substr(cp09,1,1),'C',decode('" & m_Nation & "','000',C2.cpm03,C2.cpm04),'') as 來函性質,decode(substr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||decode('" & m_Nation & "','000',C1.cpm03,C1.cpm04) as 下一程序,decode(np06,'N','Y','') as 結案," & _
                  "DECODE(np08,'','',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2)) as 本所期限," & _
                  "DECODE(np09,'','',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2)) as 法定期限," & _
                  "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,rownum as sort,np01,np22,np08,np02,np03,np04,np05,np23" & _
                  " FROM NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
                  " WHERE NP02='" & strNP02 & "' AND NP03='" & strNP03 & "' AND NP04='" & strNP04 & "' AND NP05='" & strNP05 & "'" & _
                  " and np01=cp09(+)" & _
                  " and np10=st01(+)" & _
                  " and np02=C1.cpm01(+) and np07=C1.cpm02(+)" & _
                  " and cp01=C2.cpm01(+) and cp10=C2.cpm02(+)" & _
                  " and (NP06 is null OR NP06='N')" & strNpSqlOfNoSalesDuty & _
                  " order by np08 asc,np01 asc,np02 asc,np03 asc,np04 asc,np05 asc"
      End If
      CheckOC3
      SetDataListWidth
      GRD1.Rows = 2
      GRD1.Clear
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            Set GRD1.Recordset = AdoRecordSet3.Clone
            'Add By Sindy 2012/6/29 馬德里使用宣誓要過濾重覆的子案
            Dim strTmpNP02 As String, strTmpNP03 As String, strTmpNP04 As String, strTmpNP05 As String, strTmpNP07 As String
            For i = 1 To GRD1.Rows - 1
               If i <= GRD1.Rows - 1 Then
                  If Trim(GRD1.TextMatrix(i, 16)) = "TF" And Trim(GRD1.TextMatrix(i, 11)) = "105" Then
                     If Trim(GRD1.TextMatrix(i, 17)) = strTmpNP03 And Trim(GRD1.TextMatrix(i, 19)) = strTmpNP05 And _
                        Trim(GRD1.TextMatrix(i, 11)) = strTmpNP07 Then
                        GRD1.RemoveItem i
                        i = i - 1
                     End If
                  End If
                  strTmpNP02 = Trim(GRD1.TextMatrix(i, 16))
                  strTmpNP03 = Trim(GRD1.TextMatrix(i, 17))
                  strTmpNP04 = Trim(GRD1.TextMatrix(i, 18))
                  strTmpNP05 = Trim(GRD1.TextMatrix(i, 19))
                  strTmpNP07 = Trim(GRD1.TextMatrix(i, 11))
               End If
            Next i
            '2012/6/29 End
            intRow = GRD1.Rows - 1
            For i = 1 To GRD1.Rows - 1
               If Trim(GRD1.TextMatrix(i, 5)) = "" Then bClose_N = True
               If Trim(GRD1.TextMatrix(i, 5)) = "Y" Then bClose_Y = True
               '*****
               strGetCP10 = ""
               strExc(0) = "select cp10 from caseprogress where cp09='" & Trim(GRD1.TextMatrix(i, 13)) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strGetCP10 = "" & RsTemp.Fields(0)
               End If
               '*****
               If strNP02 = "FCT" Then
                  If Trim(GRD1.TextMatrix(i, 11)) = "715" Then
                     GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 0)
                  ElseIf Trim(GRD1.TextMatrix(i, 11)) = "403" And m_Nation = "000" Then
                     intRow = intRow + 1
                     GRD1.AddItem ("")
                     GRD1.TextMatrix(intRow, 4) = "204 " & GetCaseTypeName(strNP02, "204", 0)
                     GRD1.TextMatrix(intRow, 11) = "204"
                     GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                     Call SetRowData(intRow, i)
                  End If
                  If m_Nation < "010" Then
                     If Trim(GRD1.TextMatrix(i, 11)) = "715" Then
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "717 " & GetCaseTypeName(strNP02, "717", 0)
                        GRD1.TextMatrix(intRow, 11) = "717"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                     Else
                        If Trim(GRD1.TextMatrix(i, 11)) = "403" And m_Nation = "000" Then
                           intRow = intRow + 1
                           GRD1.AddItem ("")
                           GRD1.TextMatrix(intRow, 4) = "205 " & GetCaseTypeName(strNP02, "205", 0)
                           GRD1.TextMatrix(intRow, 11) = "205"
                           GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                           Call SetRowData(intRow, i)
                        Else
                           GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 0)
                        End If
                     End If
                  Else
                     GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 1)
                  End If
               Else
                  If m_Nation < "010" Then
                     '*****
                     m_CurCP(1) = strNP02: m_CurCP(2) = strNP03: m_CurCP(3) = strNP04: m_CurCP(4) = strNP05
                     m_iDiscount = 0: strSpecial = ""
                     '辦理減免退費提醒
                     If PUB_GetCaseDiscStat(strNP02 & strNP03 & strNP04 & strNP05) = "Y" Then
                        Call PUB_CheckYearFeeReturn(m_CurCP, False, m_iDiscount, m_iYear1, m_iYear2)
                     End If
                     If m_iDiscount > 0 Then strSpecial = "1"
                     '*****
                     GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 0)
                     If ((strNP02 = "T" And Trim(GRD1.TextMatrix(i, 11)) = "403") Or (strNP02 = "P" And Trim(GRD1.TextMatrix(i, 11)) = "503")) Then
                        If (strNP02 = "T") Then
                           intRow = intRow + 1
                           GRD1.AddItem ("")
                           GRD1.TextMatrix(intRow, 4) = "204 " & GetCaseTypeName(strNP02, "204", 0)
                           GRD1.TextMatrix(intRow, 11) = "204"
                           GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                           Call SetRowData(intRow, i)
                           intRow = intRow + 1
                           GRD1.AddItem ("")
                           GRD1.TextMatrix(intRow, 4) = "205 " & GetCaseTypeName(strNP02, "205", 0)
                           GRD1.TextMatrix(intRow, 11) = "205"
                           GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-2"
                           Call SetRowData(intRow, i)
                        Else
                           intRow = intRow + 1
                           GRD1.AddItem ("")
                           GRD1.TextMatrix(intRow, 4) = "211 " & GetCaseTypeName(strNP02, "211", 0)
                           GRD1.TextMatrix(intRow, 11) = "211"
                           GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                           Call SetRowData(intRow, i)
                           intRow = intRow + 1
                           GRD1.AddItem ("")
                           GRD1.TextMatrix(intRow, 4) = "212 " & GetCaseTypeName(strNP02, "212", 0)
                           GRD1.TextMatrix(intRow, 11) = "212"
                           GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-2"
                           Call SetRowData(intRow, i)
                        End If
                     ElseIf (strNP02 = "T" And Trim(GRD1.TextMatrix(i, 11)) = "715") Then
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "717 " & GetCaseTypeName(strNP02, "717", 0)
                        GRD1.TextMatrix(intRow, 11) = "717"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                     '2016/3/22 ADD BY SONIA
'                     ElseIf (strNP02 = "T" And Trim(GRD1.TextMatrix(i, 11)) = "201") Then
'                        intRow = intRow + 1
'                        GRD1.AddItem ("")
'                        GRD1.TextMatrix(intRow, 4) = "303 " & GetCaseTypeName(strNP02, "303", 0)
'                        GRD1.TextMatrix(intRow, 11) = "303"
'                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
'                        Call SetRowData(intRow, i)
                     ElseIf (strNP02 = "T" And Trim(GRD1.TextMatrix(i, 11)) = "202") Then
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "201 " & GetCaseTypeName(strNP02, "201", 0)
                        GRD1.TextMatrix(intRow, 11) = "201"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "206 " & GetCaseTypeName(strNP02, "206", 0)
                        GRD1.TextMatrix(intRow, 11) = "206"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-2"
                        Call SetRowData(intRow, i)
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "211 " & GetCaseTypeName(strNP02, "211", 0)
                        GRD1.TextMatrix(intRow, 11) = "211"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
'                        intRow = intRow + 1
'                        GRD1.AddItem ("")
'                        GRD1.TextMatrix(intRow, 4) = "303 " & GetCaseTypeName(strNP02, "303", 0)
'                        GRD1.TextMatrix(intRow, 11) = "303"
'                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-2"
'                        Call SetRowData(intRow, i)
                     '2016/3/22 END
                     ElseIf (strNP02 = "P" And Trim(GRD1.TextMatrix(i, 11)) = "601") Then
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "412 " & GetCaseTypeName(strNP02, "412", 0)
                        GRD1.TextMatrix(intRow, 11) = "412" '延緩公告
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                        '延緩公告
                        Label1(9).Visible = True
                        Text20.Visible = True
                     ElseIf (strNP02 = "P" And Trim(GRD1.TextMatrix(i, 11)) = "205") Then
                        If strGetCP10 = "1202" Then
                           intRow = intRow + 1
                           GRD1.AddItem ("")
                           GRD1.TextMatrix(intRow, 4) = "204 " & GetCaseTypeName(strNP02, "204", 0)
                           GRD1.TextMatrix(intRow, 11) = "204"
                           GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                           Call SetRowData(intRow, i)
                        End If
                     ElseIf strSpecial = "1" Then
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "919 減免退費"
                        GRD1.TextMatrix(intRow, 11) = "919"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                     End If
                  Else
                     '*****
                     '美國發明領證案檢查有無公開費期限
                     m_MORE = ""
                     If strNP02 = "CFP" And Trim(GRD1.TextMatrix(i, 11)) = "601" Then
                        strExc(0) = "SELECT count(*) FROM NEXTPROGRESS WHERE " & _
                                          ChgNextProgress(strNP02 & strNP03 & strNP04 & strNP05) & _
                                          " And NP06 Is Null And NP07 = 217 And NP09 = " & _
                                          ChangeTStringToWString(ChangeTDateStringToTString(Trim(GRD1.TextMatrix(i, 7))))
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If RsTemp.Fields(0) > 0 Then m_MORE = "Y"
                        End If
                     End If
                     '*****
                     If strNP02 = "CFT" And Trim(GRD1.TextMatrix(i, 11)) = "312" Then
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "303 " & GetCaseTypeName(strNP02, "303", 1)
                        GRD1.TextMatrix(intRow, 11) = "303"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
'cancel by sonia 2025/5/9 陳蒲璇通知取消
'                     'Add By Sindy 2017/4/18 CFT,701領證,緬甸時;要增加顯示702刊登廣告
'                     ElseIf strNP02 = "CFT" And Trim(GRD1.TextMatrix(i, 11)) = "701" And m_Nation = "048" Then
'                        intRow = intRow + 1
'                        GRD1.AddItem ("")
'                        GRD1.TextMatrix(intRow, 4) = "702 " & GetCaseTypeName(strNP02, "702", 1)
'                        GRD1.TextMatrix(intRow, 11) = "702"
'                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
'                        Call SetRowData(intRow, i)
'                     '2017/4/18 END
'end 2025/5/9
                     ElseIf strNP02 = "CFP" And m_MORE = "Y" Then
                        GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 1)
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "217 公開費"
                        GRD1.TextMatrix(intRow, 11) = "217"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                     'Modified by Morgan 2021/7/21 IDS除外 Ex:CFP-030892 --玫音
                     'ElseIf strNP02 = "CFP" And strGetCP10 = "1006" And m_Nation = "101" Then
                     ElseIf strNP02 = "CFP" And strGetCP10 = "1006" And m_Nation = "101" And Trim(GRD1.TextMatrix(i, 11)) <> "214" Then
                     'end 2021/7/21
                        GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 1)
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "424 " & GetCaseTypeName(strNP02, "424", 0)
                        GRD1.TextMatrix(intRow, 11) = "424"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                        'Add by Amy 2017/07/20 +501訴願-甄妮 ex:CFP-026898
                        intRow = intRow + 1
                        GRD1.AddItem ("")
                        GRD1.TextMatrix(intRow, 4) = "501 " & GetCaseTypeName(strNP02, "501", 0)
                        GRD1.TextMatrix(intRow, 11) = "501"
                        GRD1.TextMatrix(intRow, 12) = CStr(GRD1.TextMatrix(i, 12)) & "-1"
                        Call SetRowData(intRow, i)
                        'end 2017/07/20
                     Else
                        GRD1.TextMatrix(i, 4) = Trim(GRD1.TextMatrix(i, 11)) & " " & GetCaseTypeName(strNP02, Trim(GRD1.TextMatrix(i, 11)), 1)
                     End If
                  End If
               End If
               '下次繳費年度(次數)
               If (strNP02 = "P" Or strNP02 = "CFP") And _
                  (Trim(GRD1.TextMatrix(i, 11)) = "605" Or Trim(GRD1.TextMatrix(i, 11)) = "606" Or Trim(GRD1.TextMatrix(i, 11)) = "607") Then
                  '設定本所案號
                  strKey(0) = Trim(GRD1.TextMatrix(i, 13))
                  strKey(1) = strNP02
                  strKey(2) = strNP03
                  strKey(3) = strNP04
                  strKey(4) = strNP05
                  '取得繳年費的資料
                  If GetMoneyDate(m_PA08, m_Nation, strKey, m_CaseFee(1), m_CaseFee(2), , , m_iFixNo) = True Then
                     '取得下次繳費次數/年度
                     m_Nexttimes = PUB_Getnexttimes(strNP02, strNP03, strNP04, strNP05, strYear)
                     If m_Nexttimes <> "" Then
                        If Trim(GRD1.TextMatrix(i, 11)) = "605" Then '605.年費
                           aryCaseFee = Split(m_CaseFee(2), ",")
                           Label1(10).Visible = True
                           Label1(10).Caption = "繳費年度：第 " & strYear & " 年"
                           Text21.Text = strYear
                           If Val(strYear) < Val(aryCaseFee(UBound(aryCaseFee))) Then
                              Label1(12).Visible = True '至第　　　年
                              Text22.Visible = True
                           End If
                        '606.維持費 607.延展費
                        ElseIf Trim(GRD1.TextMatrix(i, 11)) = "606" Or Trim(GRD1.TextMatrix(i, 11)) = "607" Then
                           '年度說明
                           'Modified by Morgan 2022/6/13 俄羅斯設計案2015/1/1以前提申案件除了延展費外仍要繳年費(繳費紀錄為年費)
                           'strFeeType = PUB_GetNa20Na22Na24(m_Nation, m_PA08)
                           strFeeType = PUB_GetNa20Na22Na24(m_Nation, m_PA08, strPA10)
                           If strFeeType = Trim(GRD1.TextMatrix(i, 11)) Then
                           'end 2022/6/13
                              strYF15 = PUB_GetYF15(m_Nation, m_PA08, "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
                              Label1(10).Visible = True
                              Label1(10).Caption = "繳費次數：" & strYF15
                              Text21.Text = strYF15
                           End If
                        End If
                     End If
                  End If
               ElseIf strNP02 = "P" And Trim(GRD1.TextMatrix(i, 11)) = "601" Then
                  'Added by Lydia 2018/06/13 取得繳年費資料
                  strKey(0) = Trim(GRD1.TextMatrix(i, 13))
                  strKey(1) = strNP02
                  strKey(2) = strNP03
                  strKey(3) = strNP04
                  strKey(4) = strNP05
                  If GetMoneyDate(m_PA08, m_Nation, strKey, m_CaseFee(1), m_CaseFee(2), , , m_iFixNo) = True Then
                  End If
                  'end 2018/06/13
                  Label1(10).Visible = True
                  Label1(10).Caption = "繳費年度：第 1 年"
                  Text21.Text = 1
                  Label1(12).Visible = True '至第　　　年
                  Text22.Visible = True
                  'Add By Sindy 2010/6/28
                  '讀取領證及繳年費的起迄預設值 ex.P-091077
                  strExc(0) = "Select cp53,cp54 " & _
                                      "From NextProgress, CaseProgress " & _
                                    "Where np01 = CP09 " & _
                                         "and np01 ='" & Trim(GRD1.TextMatrix(i, 13)) & "' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If Not IsNull(RsTemp.Fields(0)) Then
                        Label1(10).Caption = "繳費年度：第 " & RsTemp.Fields(0) & " 年"
                        Text21.Text = RsTemp.Fields(0)
                     End If
                     If Not IsNull(RsTemp.Fields(1)) Then
                        Text22 = RsTemp.Fields(1)
                     End If
                  End If
                  '2010/6/28 End
               End If
            Next i
            GRD1.col = 12
            'Modify By Sindy 2012/6/29 Mark
            'grd1.Sort = 5 '字串昇冪
            '2012/6/29 End
            If bClose_Y = False Or bClose_N = False Then
               cmdOK(2).Enabled = False
               cmdOK(3).Enabled = False
            Else
               cmdOK(2).Enabled = True
               cmdOK(3).Enabled = True
            End If
            If bClose_N = True Then
               Call cmdok_Click(3)
            Else
               Call cmdok_Click(2)
            End If
            'Modified by Lydia 2015/01/06 只有一筆資料,預設勾選
            If intRow = 1 Then
               iR = 1
               Call grd1_SelChange
            Else
               iR = 0
            End If
            doQuery = True
         Else
            'MsgBox "無符合資料！", vbInformation
         End If
      End With
   End If
   
   Screen.MousePointer = vbDefault 'Add By Sindy 2014/5/26
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetRowData(intRow As Integer, i As Integer)
   GRD1.TextMatrix(intRow, 5) = Trim(GRD1.TextMatrix(i, 5))
   GRD1.TextMatrix(intRow, 6) = Trim(GRD1.TextMatrix(i, 6))
   GRD1.TextMatrix(intRow, 7) = Trim(GRD1.TextMatrix(i, 7))
   GRD1.TextMatrix(intRow, 14) = Trim(GRD1.TextMatrix(i, 14))
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm090801_2 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim m_mouseRow As Integer
Dim iRow As Integer, lColor As Long 'Added by Morgan 2020/12/22

GRD1.Visible = False
'm_mouseRow = GRD1.MouseRow
'Modified by Lydia 2015/01/06 預設勾選
If iR = 1 Then
m_mouseRow = 1
Else
m_mouseRow = GRD1.MouseRow
End If
GRD1.col = 0
If m_mouseRow <> 0 Then
   'Modify By Sindy 2015/4/2 欲延期期限只能點選一筆資料
   If strQType = "1" Then
      '將上次點選的資料列,變成未選取狀態
      If m_row <> 0 And m_row <> m_mouseRow Then
      '   grd1.row = m_row
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
'            If grd1.CellBackColor = &HFFC0C0 Then
              GRD1.CellBackColor = &H80000018
              GRD1.TextMatrix(m_row, 0) = ""
'            Else
'              grd1.CellBackColor = &HFFC0C0 '&H80000018 '&H8080FF
'              grd1.TextMatrix(m_row, 0) = "V"
'            End If
         Next i
      End If
   End If
   '2015/4/2 END
'    If m_row <> m_mouseRow Then
        GRD1.row = m_mouseRow
        m_row = m_mouseRow
         For i = 0 To GRD1.Cols - 1
              GRD1.col = i
              If GRD1.CellBackColor = &HFFC0C0 Then
                GRD1.CellBackColor = &H80000018
                GRD1.TextMatrix(m_row, 0) = ""
              Else
                GRD1.CellBackColor = &HFFC0C0
                GRD1.TextMatrix(m_row, 0) = "V"
              End If
        Next i
        
      'Added by Morgan 2020/12/22
      'CFP美國IDS若有多筆時需全選，期限設最早的，只收文一道但在備註中註明，接洽單說明列出所有案號及國家
      m_bUsIDS = False
      If strNP02 = "CFP" And m_Nation = "101" And Left(Trim(GRD1.TextMatrix(m_row, 4)), 3) = "214" Then
         If GRD1.TextMatrix(m_row, 0) = "V" Then
            m_bUsIDS = True
         End If
         For iRow = 1 To GRD1.Rows - 1
            If m_row <> iRow Then
               GRD1.row = iRow
               If Left(Trim(GRD1.TextMatrix(iRow, 4)), 3) = "214" And GRD1.RowHeight(iRow) > 0 Then
                  If GRD1.TextMatrix(m_row, 0) = "V" Then
                     lColor = &HFFC0C0
                     GRD1.TextMatrix(iRow, 0) = "V"
                  Else
                     lColor = &H80000018
                     GRD1.TextMatrix(iRow, 0) = ""
                  End If
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = lColor
                  Next i
               End If
            End If
         Next iRow
      End If
      'end 2020/12/22
'    Else
'        m_row = 0
'    End If
End If
GRD1.Visible = True
End Sub

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, m_i As Integer
   
   GRD1.Visible = False
   'Modify By Sindy 2012/6/29 +, "np08", "np02", "np03", "np04", "np05"
   'Modify By Sindy 2015/4/2 +,np23
   arrGridHeadText = Array("V", "來函收文日", "來函性質", "來函總收文號", "下一程序" _
             , "結案", "本所期限", "法定期限", "智權人員", "相關人", "備註" _
             , "NP07", "Sort", "np01", "np22", "np08", "np02", "np03", "np04", "np05", "np23")
   arrGridHeadWidth = Array(200, 1000, 1000, 1000, 1500 _
                      , 500, 800, 800, 800, 1000, 4000 _
                      , 800, 800, 800, 0, 0, 0, 0, 0, 0, 0)
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
   GRD1.Visible = True
End Sub

Private Sub Text20_GotFocus()
   TextInverse Text20
   CloseIme
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'2010/4/23 ADD BY SONIA
Private Sub Text20_Validate(Cancel As Boolean)
   If Text20 <> "" Then
      If Len(Text20) = 1 Then
         'Modified by Morgan 2016/3/11 105/3/9日起延緩公告最長改6個月(原3個月)
         If Val(Text20) < 1 Or Val(Text20) > 6 Then
            MsgBox "延緩公告月數只可輸入1~6！", vbExclamation
            Text20_GotFocus
            Cancel = True
         End If
         'end 2016/3/11
      Else
         If ChkDate(Text20) = False Then
            Text20_GotFocus
            Cancel = True
         ElseIf Val(Text20) < Val(strSrvDate(2)) Then
            MsgBox "延緩日期不可小於系統日！", vbExclamation
            Text20_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub
'2010/4/23 END

Private Sub Text22_GotFocus()
   TextInverse Text22
   CloseIme
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
Dim nPos As Integer
Dim m_strCP81 As String   '2011/8/15 add by sonia
Dim bFound As Boolean

   Cancel = False
   'Move by Lydia 2018/06/26 NextCheck:從下面移上來
   '2011/8/15 add by sonia 台灣設計可減免者領證為1-3年
   m_strCP81 = PUB_GetCaseDiscStat(strNP02 & strNP03 & strNP04 & strNP05)
   If strNP02 = "P" And Trim(GRD1.TextMatrix(1, 11)) = "601" And m_Nation = "000" And m_PA08 = "3" And m_strCP81 = "Y" And Val(Text22) < 3 Then
      MsgBox "此案符合台灣設計可減免者1-3年年費, 領證年度請改為1-3年!", vbCritical
      Call Text22_GotFocus
      Cancel = True
   End If
   '2011/8/15 end
   
   If Trim(m_CaseFee(2)) <> "" And Val(Text22) > 0 Then
       'Modified Lydia 2018/06/13 P-116281的國內接洽單收領證601輸入繳費年度為第2-1年
      'If Val(strYear) > Val(Text22) Then
      If Val(strYear) > Val(Text22) Or Val(Text21) > Val(Text22) Then
         MsgBox "繳費年度(迄)不可小於繳費年度(起)輸入錯誤，請查明後再輸入!", vbCritical
         Call Text22_GotFocus
         Cancel = True
      End If
     
      aryCaseFee = Split(m_CaseFee(2), ",")
      ' 找尋輸入的年度是否有在字串內
      For nPos = 0 To UBound(aryCaseFee)
         If Val(Text22) = Val(aryCaseFee(nPos)) Then
            '2011/8/18 MODIFY BY SONIA
            'Exit Sub
            'Modified by Lydia 2018/06/26
            'GoTo NextCheck
            bFound = True
            Exit For
            '2011/8/18 END
         End If
      Next nPos
      
      If bFound = False Then 'Added by Lydia 2018/06/26
           MsgBox "繳費年度輸入錯誤，請查明後再輸入!", vbCritical
      'Added by Lydia 2018/06/26 國內接洽記錄單之繳費年度檢查,不可超過專利權止日.(ex. P-91086)
      Else
            If nPos > 0 Then
                strExc(1) = CompDate(0, Val(aryCaseFee(nPos - 1)), m_CaseFee(1)) '判斷繳費年度迄-1是否有超過專用期(多繳)
            'Added by Lydia 2018/06/29 第一年年費(ex.P118508收領證和繳年費601)
            Else
                strExc(1) = CompDate(0, Val(aryCaseFee(nPos)), m_CaseFee(1))
            End If
            'end 2018/06/29
            
            'Modified by Lydia 2018/06/29
            'If Val(m_TM22) > 0 And Val(strExc(1)) > Val(m_TM22) Then
            If Val(m_TM22) > 0 And Val(strExc(1)) > 0 And Val(strExc(1)) > Val(m_TM22) Then
                 MsgBox "繳費年度大於應繳年度，請查明後再輸入!", vbCritical
            Else
                 Exit Sub
            End If
      End If
      'end 2018/06/26
      
      Call Text22_GotFocus
      Cancel = True
      Exit Sub  '2011/8/15 add by sonia
   End If
   
'NextCheck: 'Move by Lydia 2018/06/26 檢查移到上面
   
End Sub


