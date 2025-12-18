VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210126 
   BorderStyle     =   1  '單線固定
   Caption         =   "回覆單"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8604
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   8604
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   180
      ScaleHeight     =   204
      ScaleWidth      =   324
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "電子檔(&S)"
      Height          =   435
      Index           =   3
      Left            =   6660
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   435
      Index           =   2
      Left            =   7620
      TabIndex        =   9
      Top             =   90
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "結案單"
      Height          =   195
      Index           =   1
      Left            =   3330
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "案件回覆單"
      Height          =   195
      Index           =   0
      Left            =   3330
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   435
      Index           =   1
      Left            =   5640
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   4620
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3825
      Left            =   60
      TabIndex        =   10
      Top             =   960
      Width           =   8475
      _ExtentX        =   14944
      _ExtentY        =   6752
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
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   3
      Top             =   375
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   2
      Top             =   375
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   1
      Top             =   375
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   375
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件回覆單請用白紙列印(已含電子表頭)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4935
      TabIndex        =   13
      Top             =   750
      Width           =   3180
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1500
      X2              =   2955
      Y1              =   510
      Y2              =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm210126"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; grd1改字型=新細明體-ExtB; 原本就用Word列印
'Memo by Lydia 2019/07/01 表單名稱:案件回覆單列印=>回覆單
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim m_row As Integer, i As Integer
Dim strNP02 As String
Dim strNP03 As String
Dim strNP04 As String
Dim strNP05 As String
Dim m_Nation As String

Private Sub cmdok_Click(Index As Integer)
Dim strData As String
Dim m_NP07s As String 'Modified by Lydia 2015/01/22
'Modified by Lydia 2015/01/22 傳多筆案件性質
If Index = 1 Or Index = 3 Then
   For intI = 1 To GRD1.Rows - 1
       If GRD1.TextMatrix(intI, 0) = "V" Then
          m_NP07s = m_NP07s & GRD1.TextMatrix(intI, 11) & ","
       End If
   Next intI
End If

'Added by Lydia 2025/08/07 MCT不需要產生回覆單紙本及電子檔
If Index = 1 Or Index = 3 Then
   '若取消MCT控制，請一併取消ClsPrtForm001的控制
   If Mid(PUB_GetAKindSalesNo(txt1(0), txt1(1), Mid(txt1(2) & "0", 1, 1), Mid(txt1(3) & "00", 1, 2)), 1, 4) = "MCTF" Then
       MsgBox "大陸來的商標案件不需要產生回覆單紙本及電子檔！", vbInformation
       Exit Sub
   End If
End If
'end 2025/08/07

Select Case Index
Case 0
        If Trim(txt1(0)) = "" Or Trim(txt1(1)) = "" Then
            MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        GRD1.MousePointer = flexHourglass
        doQuery
        GRD1.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
Case 1
        If m_row <> 0 Then
'            If Option1(0).Value = True Then    '回覆單
                Screen.MousePointer = vbHourglass
                GRD1.MousePointer = flexHourglass
                'Modify by Morgan 2011/5/24 +控制智權人員印的回覆單要含信頭信尾
                'MsgBox "請更換''本所定稿紙''後按確定開始列印！", vbInformation, "注意！"
                'Modified by Lydia 2015/01/22 傳多筆案件性質
                'Call g_PrtForm001.PrintReturnSheet(grd1.TextMatrix(m_row, 13), grd1.TextMatrix(m_row, 11), DBDATE(grd1.TextMatrix(m_row, 7)), , , , , strNP02 & strNP03 & strNP04 & strNP05, True)
                'Modified by Lydia 2020/05/25 改用Word產生列印
                'Call g_PrtForm001.PrintReturnSheet(grd1.TextMatrix(m_row, 13), m_NP07s, DBDATE(grd1.TextMatrix(m_row, 7)), , , , True, strNP02 & strNP03 & strNP04 & strNP05, True)
                Call g_PrtForm001.PrintReturnSheet(GRD1.TextMatrix(m_row, 13), m_NP07s, DBDATE(GRD1.TextMatrix(m_row, 7)), , , , True, strNP02 & strNP03 & strNP04 & strNP05, True, False)

                GRD1.MousePointer = flexDefault
                Screen.MousePointer = vbDefault
'            Else '結案單
'                Screen.MousePointer = vbHourglass
'                grd1.MousePointer = flexHourglass
'                MsgBox "請更換''案件接洽及結案紀錄單格式紙張''後按確定開始列印！", vbInformation, "注意！"
'                '新增列印接洽結案單資料
'                pub_AddressListSN = pub_AddressListSN + 1
'                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, grd1.TextMatrix(m_row, 14), "" & strNP02, "" & strNP03, "" & strNP04, "" & strNP05
'                PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'                '刪除暫存資料
'                PUB_DeleteCaseCloseSheet strUserNum
'                ShowPrintOk
'                grd1.MousePointer = flexDefault
'                Screen.MousePointer = vbDefault
'            End If
        Else
            If GRD1.Rows = 2 And GRD1.TextMatrix(1, 13) = "" Then
                MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
                txt1_GotFocus 0
            Else
                MsgBox "請先選擇一筆要列印的資料！", vbCritical, "操作錯誤！"
            End If
        End If
'Add By Sindy 2012/3/27
Case 3 '電子檔
        If m_row <> 0 Then
            Screen.MousePointer = vbHourglass
            GRD1.MousePointer = flexHourglass
            'Modify by Morgan 2011/5/24 +控制智權人員印的回覆單要含信頭信尾
            'MsgBox "請更換''本所定稿紙''後按確定開始列印！", vbInformation, "注意！"
            'Call g_PrtForm001.PrintReturnSheet(GRD1.TextMatrix(m_row, 13), GRD1.TextMatrix(m_row, 11), DBDATE(GRD1.TextMatrix(m_row, 7)), , , , , strNP02 & strNP03 & strNP04 & strNP05, True, True, Me)
            'Modified by Lydia 2015/01/22 傳多筆案件性質
            'Call g_PrtForm001.PrintReturnSheet(grd1.TextMatrix(m_row, 13), grd1.TextMatrix(m_row, 11), DBDATE(grd1.TextMatrix(m_row, 7)), , True, strData, , strNP02 & strNP03 & strNP04 & strNP05, True)
            'Modified by Lydia 2020/05/25
'            Call g_PrtForm001.PrintReturnSheet(GRD1.TextMatrix(m_row, 13), m_NP07s, DBDATE(GRD1.TextMatrix(m_row, 7)), , True, strData, True, strNP02 & strNP03 & strNP04 & strNP05, True)
'            '引用basLetter裡的變數及函數
'            m_strFilePath = PUB_Getdesktop & "\" & strNP02 & strNP03 & strNP04 & strNP05 & "-案件回覆單.doc"
'            m_bolSave2File = True
'            strPrintType = "1"
'            m_MySt(1) = strNP02
'            m_MySt(2) = strNP03
'            m_MySt(3) = strNP04
'            m_MySt(4) = strNP05
'            'Modify By Sindy 2012/7/12 不寫入定稿資料檔
'            'ExportToMsWord strData, True, True, False, 0
'            ExportToMsWord strData, True, False, False, 0
'            'End
            Call g_PrtForm001.PrintReturnSheet(GRD1.TextMatrix(m_row, 13), m_NP07s, DBDATE(GRD1.TextMatrix(m_row, 7)), , False, strData, True, strNP02 & strNP03 & strNP04 & strNP05, True, True)
            'end 2020/05/25
            GRD1.MousePointer = flexDefault
            Screen.MousePointer = vbDefault
            MsgBox "電子檔已存於 [ " & m_strFilePath & " ]！"
        Else
            If GRD1.Rows = 2 And GRD1.TextMatrix(1, 13) = "" Then
                MsgBox "請先查詢要列印的資料！", vbCritical, "操作錯誤！"
                txt1_GotFocus 0
            Else
                MsgBox "請先選擇一筆要列印的資料！", vbCritical, "操作錯誤！"
            End If
        End If
Case 2
        Unload Me
Case Else
End Select
End Sub

Sub doQuery()
On Error GoTo ErrHnd
m_row = 0
strNP02 = UCase(txt1(0))
strNP03 = txt1(1)
strNP04 = Left(txt1(2) & "0", 1)
strNP05 = Left(txt1(3) & "00", 2)
   
   m_Nation = ""
   strSql = "SELECT TM12,TM05||TM06||TM07,TM23||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,TM10,' '" & _
                " From Trademark, nation, Customer" & _
                " WHERE TM01='" & strNP02 & "' AND TM02='" & strNP03 & "' AND TM03='" & strNP04 & "' AND TM04='" & strNP05 & "'" & _
                " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
                " AND TM10=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT PA11,PA05||PA06||PA07,PA26||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,PA09,PA08" & _
                " From Patent, nation, Customer" & _
                " WHERE PA01='" & strNP02 & "' AND PA02='" & strNP03 & "' AND PA03='" & strNP04 & "' AND PA04='" & strNP05 & "'" & _
                " AND SUBSTR(PA26,1,8)=CU01(+) AND decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)" & _
                " AND PA09=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT '',LC05||LC06||LC07,LC11||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,LC15,' '" & _
                " From LawCase, nation, Customer" & _
                " WHERE LC01='" & strNP02 & "' AND LC02='" & strNP03 & "' AND LC03='" & strNP04 & "' AND LC04='" & strNP05 & "'" & _
                " AND SUBSTR(LC11,1,8)=CU01(+) AND decode(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+)" & _
                " AND LC15=NA01(+)"
   strSql = strSql & " Union " & _
                "SELECT '',HC06,HC05||' '||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),' ',' ',' '" & _
                " From HireCase, Customer" & _
                " WHERE HC01='" & strNP02 & "' AND HC02='" & strNP03 & "' AND HC03='" & strNP04 & "' AND HC04='" & strNP05 & "'" & _
                " AND SUBSTR(HC05,1,8)=CU01(+) AND decode(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      m_Nation = "" & Trim(RsTemp(4))
   End If
   
   If m_Nation < "010" Then
      strSql = "SELECT ' ' AS V,decode(substr(cp09,1,1),'C',DECODE(cp05,'','',SUBSTR(cp05,1,4)-1911||'/'||SUBSTR(cp05,5,2)||'/'||SUBSTR(cp05,7,2)),'') as 來函收文日," & _
                   "decode(substr(cp09,1,1),'C',C2.cpm03,'') as 來函性質,decode(substr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||C1.cpm03 as 下一程序,decode(np06,'N','Y','') as 結案," & _
                   "DECODE(np08,'','',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2)) as 本所期限," & _
                   "DECODE(np09,'','',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2)) as 法定期限," & _
                   "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,rownum as sort,np01,np22" & _
                   " FROM NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
                   " WHERE NP02='" & strNP02 & "' AND NP03='" & strNP03 & "' AND NP04='" & strNP04 & "' AND NP05='" & strNP05 & "'" & _
                   " and np01=cp09(+)" & _
                   " and np10=st01(+)" & _
                   " and np02=C1.cpm01(+) and np07=C1.cpm02(+)" & _
                   " and cp01=C2.cpm01(+) and cp10=C2.cpm02(+)" & _
                   " and np06 is null " & strNpSqlOfNoSalesDuty & _
                   " ORDER BY CP05 DESC, NP01 DESC, NP08 DESC "
   Else
      strSql = "SELECT ' ' AS V,decode(substr(cp09,1,1),'C',DECODE(cp05,'','',SUBSTR(cp05,1,4)-1911||'/'||SUBSTR(cp05,5,2)||'/'||SUBSTR(cp05,7,2)),'') as 來函收文日," & _
                   "decode(substr(cp09,1,1),'C',C2.cpm04,'') as 來函性質,decode(substr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||C1.cpm04 as 下一程序,decode(np06,'N','Y','') as 結案," & _
                   "DECODE(np08,'','',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2)) as 本所期限," & _
                   "DECODE(np09,'','',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2)) as 法定期限," & _
                   "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,rownum as sort,np01,np22" & _
                   " FROM NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
                   " WHERE NP02='" & strNP02 & "' AND NP03='" & strNP03 & "' AND NP04='" & strNP04 & "' AND NP05='" & strNP05 & "'" & _
                   " and np01=cp09(+)" & _
                   " and np10=st01(+)" & _
                   " and np02=C1.cpm01(+) and np07=C1.cpm02(+)" & _
                   " and cp01=C2.cpm01(+) and cp10=C2.cpm02(+)" & _
                   " and np06 is null " & strNpSqlOfNoSalesDuty & _
                   " ORDER BY CP05 DESC, NP01 DESC, NP08 DESC "
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
         SetDataListWidth
         GRD1.Visible = True
      Else
            MsgBox "無符合資料！", vbInformation
      End If
   End With
 
  
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210126 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim m_mouseRow As Integer
GRD1.Visible = False
m_mouseRow = GRD1.MouseRow
GRD1.col = 0
If m_mouseRow <> 0 Then
'    If m_row <> 0 Then
''        grd1.row = m_row
'         For i = 0 To grd1.Cols - 1
'              grd1.col = i
'              If grd1.CellBackColor = &HFFC0C0 Then
'                grd1.CellBackColor = &H80000018
'                grd1.TextMatrix(m_row, 0) = ""
'              Else
'                grd1.CellBackColor = &HFFC0C0 '&H80000018 '&H8080FF
'                grd1.TextMatrix(m_row, 0) = "V"
'              End If
'        Next i
'    End If
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
'    Else
'        m_row = 0
'    End If
End If
GRD1.Visible = True
End Sub

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
                  , "結案", "本所期限", "法定期限", "智權人員", "相關人", "備註" _
                  , "NP07", "Sort", "np01", "np22")
        arrGridHeadWidth = Array(200, 1000, 1000, 1000, 1500 _
                           , 0, 800, 800, 800, 1000, 3000 _
                           , 800, 800, 800, 0)
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
