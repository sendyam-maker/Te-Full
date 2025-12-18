VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060206 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專非台灣案已達約定期限通知"
   ClientHeight    =   5352
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9384
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5352
   ScaleWidth      =   9384
   Begin VB.CommandButton cmdHelp 
      Caption         =   "事件說明(&H)"
      Height          =   400
      Left            =   1908
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   60
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發 E-Mail(S)"
      Height          =   400
      Index           =   1
      Left            =   3096
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   60
      Width           =   1155
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060206.frx":0000
      Left            =   6120
      List            =   "frm060206.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   510
      Width           =   3195
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   915
      TabIndex        =   9
      Top             =   5025
      Width           =   3675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   4308
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   0
      Top             =   510
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   7632
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   6072
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   1500
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5208
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8484
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4095
      Left            =   45
      TabIndex        =   2
      Top             =   870
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7218
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frm060206.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin MSForms.Label lblUserName 
      Height          =   255
      Left            =   2100
      TabIndex        =   14
      Top             =   540
      Width           =   1710
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4815
      TabIndex        =   12
      Top             =   570
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   5055
      Width           =   975
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   7
      Top             =   5160
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   555
      Width           =   900
   End
End
Attribute VB_Name = "frm060206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblUserName
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
'Add by Morgan 2009/12/3
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset
Dim bLvlX As Boolean, bLvl4 As Boolean, bLvl5 As Boolean
Dim stNumList1(1 To 5) As String

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_iCols As Integer, m_iPrtCols As Integer
Dim stDept As String 'Modified by Morgan 2012/6/13 改全域變數


Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   grdDataList.FixedCols = 3
End Sub

'Add Sindy 2023/8/9
Private Sub cmdHelp_Click()
   frm060206_1.Show vbModal
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
   Dim StrToMail(1 To 6) As String 'Added by Lydia 2017/01/19
   
On Error GoTo ErrorHandler
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.CellBackColor = grdDataList.BackColor
         grdDataList.col = 3
         lngColor = grdDataList.CellBackColor
         For ii = 1 To 2
            grdDataList.col = ii
            grdDataList.CellBackColor = lngColor
         Next
         
         Dim Str01 As String
         
         StrTag = grdDataList.TextMatrix(i, 8)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If
         
         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Right(StrTag, Len(StrTag) - 1)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show
         Select Case cmdState
            'Added by Lydia 2017/01/19
            Case 1 '發E-Mail
                Me.Enabled = False
                Me.Hide
                '本所案號
                StrToMail(1) = StrTag
                '案件名稱
                StrToMail(2) = grdDataList.TextMatrix(i, 11)
                '收文日
                StrToMail(3) = ""
                '案件性質名稱
                StrToMail(4) = grdDataList.TextMatrix(i, 9)
                '法限
                StrToMail(5) = grdDataList.TextMatrix(i, 2)
                '所限
                StrToMail(6) = grdDataList.TextMatrix(i, 1)
                '內文
                strExc(1) = "           本所案號：" + StrToMail(1) + vbCrLf + vbCrLf + _
                            "           案件名稱：" + StrToMail(2) + vbCrLf + vbCrLf + _
                            "           案件性質：" + StrToMail(4) + vbCrLf + vbCrLf + _
                            "           本所期限：" + StrToMail(6) + "           法定期限：" + StrToMail(5) + vbCrLf + vbCrLf
                StrTag = ""
                '(R05管制人21,R06承辦人23,R07智權人員22)
                '管制人
                StrTag = StrTag & grdDataList.TextMatrix(i, 21) & "-"
                '智權人員
                StrTag = StrTag & grdDataList.TextMatrix(i, 22) & "-"
                '承辦人
                'Added by Lydia 2017/02/24 新案翻譯發email的承辦人要去抓"核稿人"(若核稿人為所內工程師F外翻編號，請轉為FCP所內編號)，若核稿人為空白，則發e-mail對象的承辦人為空白。
                'Modified by Lydia 2021/04/29 因P案翻譯只抓【達核稿】和FCP案不同，所以
                If "" & grdDataList.TextMatrix(i, 25) = "201" Then
                   strExc(2) = PUB_GetEP04id("" & grdDataList.TextMatrix(i, 24), True)
                   StrTag = StrTag & strExc(2) & "-"
                Else
                   StrTag = StrTag & grdDataList.TextMatrix(i, 23) & "-"
                End If
                'end 2017/02/24
                Call frm100106_4.SetParent(Me, StrToMail(5)) 'Added by Lydia 2020/03/11 傳入前一畫面和法定期限
                frm100106_4.txt1(1) = strExc(1)
                Screen.MousePointer = vbHourglass
                frm100106_4.Show
                '狀態+表單名稱
                frm100106_4.strFRname = grdDataList.TextMatrix(i, 15) + "-" & Me.Name
                frm100106_4.strCaseNo = StrToMail(1) 'Added by Lydia 2020/05/18 傳入本所案號
                frm100106_4.Tag = StrTag
                frm100106_4.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Sub
            'end 2017/01/19
            
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFP", "FCP", "P"   '專利
                     frm100101_3.Show
                     frm100101_3.Tag = StrTag
                     frm100101_3.StrMenu
                     
                  Case "FG"
                     frm100101_B.Show
                     frm100101_3.Tag = StrTag
                     frm100101_B.StrMenu
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
          End Select
         Exit For
      End If
   Next i
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   PUB_RestorePrinter cboPrinter.Text
   DoPrint
   PUB_RestorePrinter strPrinter
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmdQuery_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   doQuery
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function GetNumList(p_UID As String, Optional iLevel As Integer) As String
   Dim stRtn As String, stSQL As String
   
   stSQL = "select ''''||st01||'''' from staff"
   Select Case iLevel
      Case 2
         stSQL = stSQL & " where ST52='" & p_UID & "'"
      Case 3
         stSQL = stSQL & " where ST53='" & p_UID & "'"
      Case 4
         stSQL = stSQL & " where ST54='" & p_UID & "'"
      Case 5
         stSQL = stSQL & " where ST55='" & p_UID & "'"
      Case Else
         stSQL = stSQL & " where ST52='" & p_UID & "'"
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      stRtn = RsTemp.GetString(adClipString, , , ",")
      stRtn = Left(stRtn, Len(stRtn) - 1)
   End If
   GetNumList = stRtn
End Function

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery()
Dim stVTB11 As String, stVTB2 As String
Dim stVTB1 As String, stConNA16 As String, stConNP10 As String, stConNA51 As String
Dim stNumList As String
Dim ii As Integer, stIdList
Dim stUserID As String
'Added by Morgan 2012/6/14
Dim stConCP06 As String, stDate1 As String, stDate2 As String
Dim stConNA51P As String 'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
   
   stDate1 = strSrvDate(1) - 10000
   stDate2 = CompWorkDay(4, strSrvDate(1))
   stConCP06 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate2
   
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！"
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   'Added by Morgan 2015/11/13
   ElseIf Pub_StrUserSt03 = "F22" And txtUsernum <> strUserNum Then
      If PUB_GetST03(txtUsernum) <> Pub_StrUserSt03 Then
         MsgBox "員工編號錯誤！", vbExclamation, "權限不足"
         Exit Sub
      End If
   'end 2015/11/13
   End If
   
   stUserID = txtUsernum
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
   Else
      stDept = GetST15(stUserID)
   End If
   
   bLvl4 = CheckLevel(stUserID, "N") '第四級管制人(+FCP,FG未分案將到期)
   bLvl5 = CheckLevel(stUserID, "O") '第五級管制人(+FCP,FG未分案已逾期)
   
   If Not (stDept = "F22" Or stDept = "F23" Or bLvl4 = True Or bLvl5 = True) Then Exit Sub
   
   stNumList = "'" & stUserID & "'"
   stNumList1(1) = stNumList
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stIdList = Split(stNumList, ",")
   'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
   '代理人Y51333010=Pub_GetSpecMan("北京銀龍FCP案承辦業務") ,NA51 = decode(pa75," & midstr & ",na51)
   Dim midStr As String, stConSP26 As String
   'Modified by Lydia 2016/02/03改成回傳case句
   'midStr = Pub_GetSpecMan("北京銀龍FCP案承辦業務")
   midStr = Pub_GetSpecFCP
   
   If InStr(stNumList, ",") > 0 Then
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
        stConNA16 = " AND NA16 in (" & stNumList & ") "
      Else
        stConNA16 = " AND NVL(NA79,NA16) in (" & stNumList & ") "
      End If
      'end 2017/02/13
      
      'Modified by Morgan 2012/4/26 FMP 管制智權人員改和 FCP 一樣抓 NA51
      'stConNP10 = " AND NP10 in (" & stNumList & ") "
      stConNA51 = " AND NA51 in (" & stNumList & ") "
      'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
      'Modified by Lydia 2016/02/03
       'stConNA51P = " AND decode(pa75,'Y51333010','" & midStr & "',na51) in (" & stNumList & ") "
       'stConSP26 = " AND decode(sp26,'Y51333010','" & midStr & "',na51) in (" & stNumList & ") "
       stConNA51P = " AND decode(pa75," & midStr & ",na51) in (" & stNumList & ") "
       stConSP26 = " AND decode(sp26," & midStr & ",na51) in (" & stNumList & ") "
   Else
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
        stConNA16 = " AND NA16=" & stNumList
      Else
        stConNA16 = " AND NVL(NA79,NA16)=" & stNumList
      End If
      'end 2017/02/13
      
      'Modified by Morgan 2012/4/26 FMP 管制智權人員改和 FCP 一樣抓 NA51
      'stConNP10 = " AND NP10=" & stNumList
      stConNA51 = " AND NA51=" & stNumList
      'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
      'Modified by Lydia 2016/02/03
      'stConNA51P = " AND decode(pa75,'Y51333010','" & midStr & "',na51)=" & stNumList
      'stConSP26 = " AND decode(sp26,'Y51333010','" & midStr & "',na51)=" & stNumList
      stConNA51P = " AND decode(pa75," & midStr & ",na51)=" & stNumList
      stConSP26 = " AND decode(sp26," & midStr & ",na51)=" & stNumList
   End If
   
   '代碼1(R02):A=達本所,D=未收文,H=達約定
   '代碼2(R03):(數字的)0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿 ==> 目前只看到使用:2 (Add By Sindy 2023/8/8 加註解)
   
   'Added by Morgan 2012/6/14
   '清除暫存檔
   strSql = "delete R060206 where R01='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   
   '程序
   If stDept = "F22" Then
      '約定期限＜＝系統日之未收文FMP案件--管制人-H2(達約定,管制人)
      'Modified by Morgan 2012/6/15 +,NA51,PA75,NP08,NP09
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75," & midstr & ",na51)
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2017/02/20 NA16 C02= > NVL(NA79,NA16) C02
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) 取消  and np07<>'605' 條件
      stVTB1 = " SELECT NP01 C01,NVL(NA79,NA16) C02,NP22 C03,decode(pa75," & midStr & ",na51) NA51,PA75,NP08,NP09" & _
         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE (NP02='P' or NP02='CFP') and NP06 is null AND NP23<=" & strSrvDate(1) & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16
         
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'      'Modified by Lydia 2014/11/14 年費代理人 pa75->pa76
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NVL(NA79,NA16) NA16,NP22,decode(pa76," & midStr & ",na51) NA51,PA76,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA16
'      'Modified by Lydia 2014/11/14  pa75->pa76
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NVL(NA79,NA16) NA16,NP22,decode(pa76," & midStr & ",na51) NA51,PA76,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, Customer, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA16
'      'Modified by Lydia 2014/11/14  pa75->cu96
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NVL(NA79,NA16) NA16,NP22,decode(cu96," & midStr & ",na51) NA51,CU96,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT,Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA16
'      'Modified by Lydia 2014/11/14  pa75->cu96
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NVL(NA79,NA16) NA16,NP22,decode(c1.cu96," & midStr & ",na51) NA51,c1.CU96,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER c1,CUSTOMER c2,Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA16
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NVL(NA79,NA16) NA16,NP22,decode(pa75," & midStr & ",na51) NA51,PA75,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT,Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16
'
'end 2020/5/12
      
      'Modified by Lydia 2014/11/14
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2017/02/20 NA16 C02= > NVL(NA79,NA16) C02
      stVTB2 = "SELECT NP01 C01,NVL(NA79,NA16) C02,NP22 C03,decode(sp26," & midStr & ",na51) NA51,SP26,NP08,NP09" & _
         " From NEXTPROGRESS,CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06='PS' AND NP23<=" & strSrvDate(1) & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConNA16
      'Modified by Lydia 2014/11/14
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT '" & strUserNum & "','H' EV1,'2' EV2,C01,C02,NULL,NA51,NP08,NP09,NULL,NULL,C03,PA75" & _
         " FROM (" & stVTB1 & " UNION " & stVTB2 & ")"
      cnnConnection.Execute strSql, intI
      
   '承辦或主管
   Else
      
      '約定期限＜＝系統日之未收文FMP案件--智權人員-H3(達約定,智權人員)
      'Modified by Morgan 2012/6/15 +NA51,PA75,NP08,NP09
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75," & midstr & ",na51),stConNA51->stConNA51P
      'Modified by Lydia 2016/02/03
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) 取消  and np07 <>'605' 條件
      stVTB1 = " SELECT NP01 C01,NA16 C02,NP22 C03,decode(pa75," & midStr & ",na51) NA51,PA75,NP08,NP09" & _
         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE (NP02='P' or NP02='CFP') and NP06 is null AND NP23<=" & strSrvDate(1) & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P
         
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'      'Modified by Lydia 2014/11/14  pa75->pa76
'      'Modified by Lydia 2016/02/03
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NA16,NP22,decode(pa76," & midStr & ",na51) NA51,PA76,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07 ='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA51P
'      'Modified by Lydia 2014/11/14  pa75->pa76
'      'Modified by Lydia 2016/02/03
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NA16,NP22,decode(pa76," & midStr & ",na51) NA51,PA76,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, CUSTOMER, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07 ='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA51P
'      'Modified by Lydia 2014/11/14  pa75->cu96
'      'Modified by Lydia 2016/02/03
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NA16,NP22,decode(cu96," & midStr & ",na51) NA51,CU96,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, CUSTOMER, FAGENT, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07 ='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA51P
'      'Modified by Lydia 2014/11/14  pa75->cu96
'      'Modified by Lydia 2016/02/03
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NA16,NP22,decode(c1.cu96," & midStr & ",na51) NA51,c1.CU96,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT,CUSTOMER c1,CUSTOMER c2, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07 ='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA51P
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stVTB1 = stVTB1 & " UNION ALL SELECT NP01,NA16,NP22,decode(pa75," & midStr & ",na51) NA51,PA75,NP08,NP09" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, CUSTOMER, FAGENT, Nation" & _
'         " WHERE (NP02='P' or NP02='CFP') and NP06 is null and np07 ='605' AND NP23<=" & strSrvDate(1) & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P
'
'end 2020/5/12

      'Modified by Lydia 2014/11/14 pa75->sp26
      'Modified by Lydia 2016/02/03
      stVTB2 = " SELECT NP01 C01,NA16 C02,NP22 C03,decode(sp26," & midStr & ",na51) NA51,SP26,NP08,NP09" & _
         " From NEXTPROGRESS,CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06='PS' AND NP23<=" & strSrvDate(1) & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F' AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConSP26
      'Modified by Lydia 2014/11/14 na51
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT '" & strUserNum & "','H' EV1,'2' EV2,C01,C02,NULL,NA51,NP08,NP09,NULL,NULL,C03,PA75" & _
         " FROM (" & stVTB1 & " UNION " & stVTB2 & ")"
      cnnConnection.Execute strSql, intI
   End If
      
   'Added by Morgan 2012/6/14
   Combo1.Clear
   If stDept = "F22" Then
      Combo1.AddItem "紅色(*)：表示逾管控期限"
      Combo1.AddItem "綠色(v): 表示當日期限"
      Combo1.AddItem "藍色: 表示點選資料"
      Combo1.ListIndex = 0
      
      '達本所未完稿
      'Modified by Lydia 2016/09/14 CP57||CP27 IS NULL => CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/26 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/ '" & strUserNum & "','A' EV1,'2' EV2,CP09,NVL(NA79,NA16) NA16,CP14,CP13,CP06,CP07,CP48,NULL,0,PA75" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation,STAFF,engineerprogress Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0" & _
         " AND SUBSTR(CP12,1,1)='F' AND ST01(+)=CP14 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09 is null" & stConCP06 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute strSql, intI
      
      'Modified by Lydia 2016/09/14 CP57||CP27 IS NULL => CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/26 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/ '" & strUserNum & "','A' EV1,'2' EV2,CP09,NVL(NA79,NA16) NA16,CP14,CP13,CP06,CP07,CP48,NULL,0,SP26" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation,STAFF,engineerprogress Where (CP01='PS' OR CP01='CPS') AND CP158=0 AND CP159=0" & _
         " AND SUBSTR(CP12,1,1)='F' AND ST01(+)=CP14 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09 is null" & stConCP06 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute strSql, intI
      
      '達本所已完稿未核稿
       'Modified by Lydia 2016/09/14 CP57||CP27 IS NULL => CP158=0 AND CP159=0
       'Modified by Lydia 2016/09/26 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
       'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/ '" & strUserNum & "','A' EV1,'2' EV2,CP09,NVL(NA79,NA16) NA16,EP04,CP13,CP06,CP07,CP48,NULL,0,PA75" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation,STAFF,engineerprogress Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0" & _
         " AND SUBSTR(CP12,1,1)='F' and cp10='201' AND ST01(+)=EP04 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09>0 and ep33 is null" & stConCP06 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute strSql, intI
      
      '未收文且 2個工作天 後達本所期限者(不含當日) --管制人-D2(未收文,管制人)
      '非年費
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75," & midstr & ",na51)
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) 取消  and np07<>'605' 條件
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NVL(NA79,NA16) NA16,NULL,decode(pa75," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,PA75" & _
         " From NEXTPROGRESS, PATENT, FAGENT, Nation,STAFF" & _
         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F'" & strNpSqlOfNoSalesDuty & _
         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16 & _
         " AND NOT EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928' AND NP07='202')"
      cnnConnection.Execute strSql, intI
      
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'      '年費,個案有年費代理人 Y
'      'Modified by Lydia 2014/11/14  pa75->pa76
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NVL(NA79,NA16) NA16,NULL,decode(pa76," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,PA76" & _
'         " From NEXTPROGRESS, PATENT, FAGENT, Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute strSql, intI
'      '年費,個案有年費代理人 X
'      'Modified by Lydia 2014/11/14  pa75->pa76
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NVL(NA79,NA16) NA16,NULL,decode(pa76," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,PA76" & _
'         " From NEXTPROGRESS, PATENT, Customer, Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA16
'      cnnConnection.Execute strSql, intI
'      '年費,個案無有年費代理,客戶有年費代理人 Y
'      'Modified by Lydia 2014/11/14  pa75->cu96
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NVL(NA79,NA16) NA16,NULL,decode(cu96," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,CU96" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER,FAGENT,Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute strSql, intI
'      '年費,個案無有年費代理,客戶有年費代理人 X
'      'Modified by Lydia 2014/11/14  pa75->cu96
'      'Modified by Lydia 2016/02/03
'      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NA01,NULL,decode(c1.cu96," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,c1.CU96" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER c1,CUSTOMER c2,Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA16
'      cnnConnection.Execute strSql, intI
'
'      '年費,無年費代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
'      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NVL(NA79,NA16) NA16,NULL,decode(pa75," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,PA75" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER,FAGENT,Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute strSql, intI
'
'end 2020/5/12

      '服務業務
      'Modified by Lydia 2014/11/14  pa75->sp26
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2017/02/20 NA16= > NVL(NA79,NA16) NA16
      strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NVL(NA79,NA16) NA16,NULL,decode(sp26," & midStr & ",na51) NA51,NP08,NP09,NULL,NULL,NP22,SP26" & _
         " From NEXTPROGRESS, SERVICEPRACTICE, FAGENT, Nation,STAFF" & _
         " WHERE NP02||NP06 in ('PS','CPS') and st01(+)=NP10 and substr(st15,1,1)='F' " & strNpSqlOfNoSalesDuty & _
         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConNA16
      cnnConnection.Execute strSql, intI
   Else
      Combo1.AddItem "藍色: 表示點選資料"
      Combo1.ListIndex = 0
   End If
   'end 2012/6/14
   
   
   'Modified by Morgan 2012/4/26 FMP 管制智權人員改和 FCP 一樣抓 NA51
   'Modified by Morgan 2012/6/15
   'strExc(0) = " SELECT '' V,NVL(lpad(SQLDateT(NP08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(NP09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(NP23),9,' '),'2') 約定期限,S1.ST02 管制人,S2.ST02 智權人員,'' 承辦人,'達約定' 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(INSTR('020,013',PA09),0,CPM03,CPM04) 案件性質,NP15 備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",'' 未收款,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) 代理人,NA03 代理人國籍,NP02,NP03,NP04,NP05" & _
      " FROM (" & stVTB1 & ") X,NEXTPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP, FAGENT,NATION" & _
      " WHERE  NP01(+)=C01 AND NP22(+)=C03" & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=C02 AND S2.ST01(+)=NA51 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10"
      
   'strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(NP08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(NP09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(NP23),9,' '),'2') 約定期限,S1.ST02 管制人,S2.ST02 智權人員,'' 承辦人,'達約定' 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(INSTR('020,013',SP09),0,CPM03,CPM04) 案件性質,NP15 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",'' 未收款,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) 代理人,NA03 代理人國籍,NP02,NP03,NP04,NP05" & _
      " FROM (" & stVTB2 & ") X,NEXTPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP, FAGENT,NATION" & _
      " WHERE  NP01(+)=C01 AND NP22(+)=C03" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=C02 AND S2.ST01(+)=NA51 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10"
   
   
   '刪除達約定與未收文重複者
   strSql = "DELETE R060206 R1 WHERE R01='" & strUserNum & "' AND R02='H'" & _
      " AND EXISTS(SELECT * FROM R060206 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R12=R1.R12 AND R2.R02='D')"
   cnnConnection.Execute strSql, intI
   
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+CP10
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   strExc(0) = " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'2' 約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",'' 未收款,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) 代理人,NA03 代理人國籍,R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND NVL(SUBSTR(R13,1,1),'Y')='Y' AND CP09(+)=R04" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+CP10
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'2' 約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",'' 未收款,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) 代理人,NA03 代理人國籍,R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND NVL(SUBSTR(R13,1,1),'Y')='Y' AND CP09(+)=R04" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+NP07
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(NP23),9,' '),'2')  約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,NP15 備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",'' 未收款,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) 代理人,NA03 代理人國籍,R02,R03,NP02,NP03,NP04,NP05,R05,R07,R06,R04,NP07" & _
      " FROM R060206,NEXTPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'A' AND NVL(SUBSTR(R13,1,1),'Y')='Y' AND NP01(+)=R04 AND NP22(+)=R12" & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+NP07
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(NP23),9,' '),'2')  約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,NP15 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",'' 未收款,RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65) 代理人,NA03 代理人國籍,R02,R03,NP02,NP03,NP04,NP05,R05,R07,R06,R04,NP07" & _
      " FROM R060206,NEXTPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'A' AND NVL(SUBSTR(R13,1,1),'Y')='Y' AND NP01(+)=R04 AND NP22(+)=R12" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+CP10
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'2' 約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",'' 未收款,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90) 代理人,NA03 代理人國籍,R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND SUBSTR(R13,1,1)='X' AND CP09(+)=R04" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND CU01(+)=SUBSTR(R13,1,8) AND CU02(+)=SUBSTR(R13,9) AND NA01(+)=CU10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+CP10
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'2' 約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",'' 未收款,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90) 代理人,NA03 代理人國籍,R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND SUBSTR(R13,1,1)='X' AND CP09(+)=R04" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND CU01(+)=SUBSTR(R13,1,8) AND CU02(+)=SUBSTR(R13,9) AND NA01(+)=CU10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+NP07
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(NP23),9,' '),'2')  約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,NP15 備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",'' 未收款,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90) 代理人,NA03 代理人國籍,R02,R03,NP02,NP03,NP04,NP05,R05,R07,R06,R04,NP07" & _
      " FROM R060206,NEXTPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'A' AND SUBSTR(R13,1,1)='X' AND NP01(+)=R04 AND NP22(+)=R12" & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND CU01(+)=SUBSTR(R13,1,8) AND CU02(+)=SUBSTR(R13,9) AND NA01(+)=CU10"
      
   'Modified by Lydia 2017/01/19 +R06 承辦人員工編號
   'Modified by Lydia 2017/02/24 +R04 收文號,+NP07
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(NP23),9,' '),'2')  約定期限,S1.ST02 管制人,S3.ST02 智權人員,S2.ST02 承辦人" & _
      ",DECODE(R02,'A','達本所','D','未收文','H','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,NP15 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",'' 未收款,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90) 代理人,NA03 代理人國籍,R02,R03,NP02,NP03,NP04,NP05,R05,R07,R06,R04,NP07" & _
      " FROM R060206,NEXTPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'A' AND SUBSTR(R13,1,1)='X' AND NP01(+)=R04 AND NP22(+)=R12" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07" & _
      " AND CU01(+)=SUBSTR(R13,1,8) AND CU02(+)=SUBSTR(R13,9) AND NA01(+)=CU10"
   'end 2012/6/14
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   
   If RsTemp Is Nothing Then Exit Sub
   If RsTemp.RecordCount = 0 Then
      Set m_adoRst = RsTemp.Clone
      SetRst2Grid
      MsgBox "查無資料！", vbInformation
      LblCnt.Caption = "共 0 筆"
   Else
      Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300, Me.Name) 'Modify by Amy 2014/06/06 +FormName
      '更新未收文未收款
      SetXRecord
      Select Case stDept
         Case "F22" '程序
            'Modified by Lydia 2020/09/07 一併改用員工編號
            'm_stSort = "本所期限 asc,管制人 asc,代理人國籍 asc,本所案號 asc"
            m_stSort = "本所期限 asc,R05 asc,代理人國籍 asc,本所案號 asc"
         Case "F23" '智權人員
            'Modified by Lydia 2020/09/07 承辦組查詢會出錯
            'm_stSort = "本所期限 asc,智權人員 asc,代理人國籍 asc,本所案號 asc"
            m_stSort = "本所期限 asc,R07 asc,代理人國籍 asc,本所案號 asc"
         'F21,F81
         Case Else
            'Modified by Lydia 2020/09/07 承辦組查詢會出錯; err. 無法在其定義長度是不明或過長的資料行執行
            'm_stSort = "本所期限 asc,智權人員 asc,代理人國籍 asc,本所案號 asc"
            m_stSort = "本所期限 asc,R07 asc,代理人國籍 asc,本所案號 asc"
      End Select
      m_adoRst.Sort = m_stSort
      SetRst2Grid
      SetGrid
      RecordShow
      SetColor
      m_blnColOrderAsc = True
   End If
End Sub

Private Sub SetGrid()
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      'Modified by Morgan 2012/6/13 +|承辦人 |事件
      .FormatString = "V|本所期限 |法定期限 |約定期限 |管制人 |智權人員 |承辦人 |事件　 |本所案號　　|案件性質 |備註　　　　|案件名稱　　|未收款|代理人　　|代理人國籍"
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = 0
         If intI > 14 Then
            .ColWidth(intI) = 0
         End If
      Next
      .ColAlignment(12) = flexAlignRightTop
      .ColAlignment(1) = flexAlignRightTop
      .ColAlignment(2) = flexAlignRightTop
      .ColAlignment(3) = flexAlignRightTop
      .ColAlignment(4) = flexAlignRightTop
      
      'Added by Morgan 2012/6/13
      If stDept <> "F22" Then
         .ColWidth(6) = 0
         .ColWidth(7) = 0
      End If
      'end 2012/6/13
      
      .Visible = True
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtUsernum = strUserNum
   
   'Modified by Morgan 2015/11/13
   '改外專程序可看該組其他人員資料
   Select Case Pub_StrUserSt03
   Case "M51", "F22"
      txtUsernum.Enabled = True
   End Select
   'end 2015/11/13
   
   PUB_SetPrinter Me.Name, cboPrinter, strPrinter
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
   
   Combo1.Clear
   Combo1.AddItem "藍色: 表示點選資料"
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If Not bolUnloading Then 'Add by Morgan 2011/3/11
      If cboPrinter.Text <> cboPrinter.Tag Then
         PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinter.Name, "0", "0", Me.cboPrinter.Text
      End If
      If CheckUse("frm060204", strExec, False) = True Then
         strSql = "select * from executelog where el01='frm060204' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI <> 1 Then
            pub_bolInformCheck = True
            Load frm060204
            frm060204.cmdQuery(0).Value = True
            pub_bolInformCheck = False
         End If
      End If
   End If
   Set frm060206 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCol As Integer
    
    If grdDataList.Rows = 1 Or grdDataList.Recordset Is Nothing Then Exit Sub 'Added by Morgan 2021/4/9 無資料(備註)或未查詢時排序會錯
    
    iCol = grdDataList.MouseCol
    If grdDataList.MouseRow < 1 Then
      grdDataList.Visible = False
      ChgEmptyDate True
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc," & m_stSort
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc," & m_stSort
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      SetColor
      grdDataList.Visible = True
    End If
End Sub

Private Sub ChgEmptyDate(Optional p_bolBeforeSort As Boolean)
   Dim ii As Integer, jj As Integer
   With grdDataList
   If .Rows > 1 Then
      For ii = 1 To .Rows - 1
         For jj = 1 To 3
            If p_bolBeforeSort Then
               If .TextMatrix(ii, jj) = "" Then
                  .TextMatrix(ii, jj) = "2"
               End If
            Else
               If .TextMatrix(ii, jj) = "2" Then
                  .TextMatrix(ii, jj) = ""
               End If
            End If
         Next
      Next
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            .col = 0
            .CellBackColor = .BackColor
            .col = 3
            lngColor = .CellBackColor
            For ii = 1 To 2
               .col = ii
               .CellBackColor = lngColor
            Next
         Else
            .Text = "V"
            For ii = 0 To 2
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

Private Function SetXRecord()
   Dim iRow As Integer
   With m_adoRst
      .MoveFirst
      Do While Not .EOF
         strExc(0) = "select nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)),0) from acc1k0" & _
            " where a1k13 = '" & .Fields(17) & "' and a1k14 = '" & .Fields(18) & "' and a1k15 = '" & .Fields(19) & "' and a1k16 = '" & .Fields(20) & "' and (a1k29 is null or a1k29 = '')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp(0) > 0 Then
               .Fields(12) = Format(RsTemp(0), "#,###")
            End If
         End If
         .MoveNext
      Loop
      .UpdateBatch
   End With
End Function

Private Sub SetColor()
   Dim lngToday As Long, lngCP06 As Long, lngNP23 As Long, stType As String
   Dim ii As Integer, jj As Integer, dblCnt As Double
   
   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      ChgEmptyDate False
      lngToday = Val(strSrvDate(2)) 'Added by Morgan 2012/6/15
      For ii = 1 To .Rows - 1
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            .CellAlignment = flexAlignRightTop
            .CellFontSize = 9
         Next
         
         'Added by Morgan 2012/6/15
         '程序組要變色
         If stDept = "F22" Then
            .RowHeight(ii) = 255
            lngCP06 = Val(Replace(.TextMatrix(ii, 1), "/", ""))
            lngNP23 = Val(Replace(.TextMatrix(ii, 3), "/", ""))
            stType = .TextMatrix(ii, 15)
            
            '逾管控期限
            If ((stType = "A" Or stType = "D") And lngCP06 > 0 And lngCP06 < lngToday) Or _
               (stType = "H" And lngNP23 > 0 And lngNP23 < lngToday) Then
               .TextMatrix(ii, 8) = "*" & .TextMatrix(ii, 8)
               For jj = 1 To .Cols - 1
                  .col = jj
                  '紅
                  .CellBackColor = &HFF&
               Next
            '當日期限
            ElseIf ((stType = "A" Or stType = "D") And lngCP06 > 0 And lngCP06 = lngToday) Or _
               (stType = "H" And lngNP23 > 0 And lngNP23 = lngToday) Then
               .TextMatrix(ii, 8) = "v" & .TextMatrix(ii, 8)
               For jj = 1 To .Cols - 1
                  .col = jj
                  '綠
                  .CellBackColor = &HC000&
               Next
            ElseIf stType <> "H" Then
               strExc(1) = .TextMatrix(ii, 16)
               Select Case strExc(1)
                  Case "2" '管制人
                     strExc(2) = .TextMatrix(ii, 21)
                  Case "3" '智權人員
                     strExc(2) = .TextMatrix(ii, 22)
                  Case Else
                     strExc(2) = ""
               End Select
               
               If strExc(2) <> "" Then
                  '本人或第二級才看
                  If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 Then
                     .RowHeight(ii) = 0
                  End If
               End If
            End If
         End If
         'end 2012/6/15
            
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With

   LblCnt.Caption = "共 " & dblCnt & " 筆"

End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      
      'Added by Morgan 2012/6/13
      'ReDim strTemp(1 To m_iCols)
      If stDept = "F22" Then
         m_iPrtCols = m_iCols - 2
      Else
         m_iPrtCols = m_iCols
      End If
      ReDim strTemp(1 To m_iPrtCols)
      'end 2012/6/13
      
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            If iCol = 9 Then '案件性質
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol), 4)
            ElseIf iCol = 10 Then '備註
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol), 6)
            ElseIf iCol = 11 Then '案件名稱
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol), 6)
            ElseIf iCol = 13 Then '代理人
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol), 10)
            Else
               strTemp(iCol) = .TextMatrix(iRow, iCol)
            End If
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   m_iCols = 14
   ReDim PLeft(1 To m_iCols)
   PLeft(1) = ciStartX
   For intI = 2 To m_iCols
      If grdDataList.ColWidth(intI - 1) > 0 Then
         PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(grdDataList.TextMatrix(0, intI - 1)) + ciColGap
      Else
         PLeft(intI) = PLeft(intI - 1)
      End If
   Next
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print String(130, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      If Me.grdDataList.ColWidth(iCol) > 0 Then 'Added by Morgan 2012/6/13
         If iCol = 12 Then
            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
         Else
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
         End If
      End If
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Replace(Me.Caption, "通知", "清單")
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "員工編號：" & txtUsernum & " 姓名：" & lblUserName
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(130, "-")
End Sub

Sub PrintPageHeader1()
   
    Call PrintNewLine(False, 1)
    For intI = 1 To m_iPrtCols
      If Me.grdDataList.ColWidth(intI) > 0 Then 'Added by Morgan 2012/6/13
         Printer.CurrentX = PLeft(intI)
         Printer.CurrentY = iPrint
         Printer.Print grdDataList.TextMatrix(0, intI)
      End If
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub


