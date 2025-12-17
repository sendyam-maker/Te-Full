VERSION 5.00
Begin VB.Form frm160102 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人人事資料明細列印"
   ClientHeight    =   4310
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4310
   ScaleWidth      =   6220
   Begin VB.CheckBox Check1 
      Caption         =   "列印時，含專長資料"
      Height          =   225
      Left            =   1410
      TabIndex        =   5
      Top             =   2100
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   1980
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frm160102.frx":0000
      Top             =   2550
      Width           =   3075
   End
   Begin VB.PictureBox tmpPic 
      AutoRedraw      =   -1  'True
      DragMode        =   1  '自動
      Height          =   2900
      Left            =   0
      ScaleHeight     =   2860
      ScaleWidth      =   2280
      TabIndex        =   17
      Top             =   4350
      Width           =   2320
      Begin VB.Image tmpImg 
         BorderStyle     =   1  '單線固定
         Height          =   960
         Left            =   420
         Stretch         =   -1  'True
         Top             =   390
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      Height          =   500
      Left            =   2610
      ScaleHeight     =   46
      ScaleMode       =   3  '像素
      ScaleWidth      =   46
      TabIndex        =   16
      Top             =   4350
      Width           =   500
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   60
      ScaleHeight     =   460
      ScaleWidth      =   650
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "電子化線上確認(&E)"
      Default         =   -1  'True
      Height          =   435
      Index           =   2
      Left            =   2460
      TabIndex        =   7
      Top             =   90
      Width           =   1785
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2340
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1680
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   435
      Index           =   0
      Left            =   4290
      TabIndex        =   8
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   5235
      TabIndex        =   9
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   0
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2970
      MaxLength       =   3
      TabIndex        =   1
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1350
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   3120
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1350
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   270
      TabIndex        =   10
      Top             =   3630
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "繳回日期："
      Height          =   180
      Left            =   1410
      TabIndex        =   14
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1410
      TabIndex        =   13
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   1410
      TabIndex        =   12
      Top             =   1380
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2910
      X2              =   3300
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line2 
      X1              =   3060
      X2              =   3300
      Y1              =   1500
      Y2              =   1500
   End
End
Attribute VB_Name = "frm160102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/24 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2009/01/09
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 62) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim m_QueryCnt As Integer
Public intChoose As Integer 'Add By Sindy 2011/9/2 0.人事系統 1.出缺勤系統
Dim m_Device, m_iPages As Integer  'Add By Sindy 2011/9/2
Dim douExtRate As Double '字型位置縮放比
Dim sW As Integer, sH As Integer 'Add By Sindy 2012/6/20
Dim m_HaveSpecialty As Boolean 'Add By Sindy 2014/3/12 記錄有無專長資料
Dim m_ShowSpecialty As Boolean 'Add By Sindy 2014/3/12 記錄是否要顯示或列印專長資料
'Add By Sindy 2022/1/23
Dim strPrinter As String
Dim strTempFile As String
'2022/1/23 END


Private Sub GetSql()
   m_StrSQL = ""
   If txt1(0) <> "" Then
       'Modify By Sindy 2023/12/22 部門調整改抓ST93
       m_StrSQL = m_StrSQL & " and st93>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       'Modify By Sindy 2023/12/22 部門調整改抓ST93
       m_StrSQL = m_StrSQL & " and st93<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       m_StrSQL = m_StrSQL & " and st01>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
       m_StrSQL = m_StrSQL & " and st01<='" & txt1(3) & "' "
   End If
End Sub

Public Sub cmdok_Click(Index As Integer)
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim i As Integer, intRow As Integer
Dim strMeSql As String

m_ShowSpecialty = False 'Add By Sindy 2014/3/12
Select Case Index
Case 0, 2
        If intChoose = 1 Then '出缺勤系統
            m_ShowSpecialty = True 'Add By Sindy 2014/3/12
        Else
            If Check1.Value = 1 Then m_ShowSpecialty = True 'Add By Sindy 2014/3/12
            If txt1(4) = "" Then
                MsgBox "繳回日期不可以空白！", vbInformation, "操作錯誤！"
                txt1(4).SetFocus
                Exit Sub
            End If
            Call txt1_Validate(4, Cancel)
            If Cancel = True Then Exit Sub
        End If
        
        GetSql
        If Index = 0 Then '列印
            If intChoose = 1 Then '出缺勤系統
               txt1(0) = "": txt1(1) = ""
               For i = 1 To 2
                  If i = 1 Then
                     '第一筆先讀取自己的
                     strMeSql = "SELECT * FROM ABS013 WHERE B1301='05' and B1302='" & strUserNum & "' and B1303='" & strUserNum & "'"
                  ElseIf i = 2 Then
                     '簽核他人的
                     strMeSql = "SELECT * FROM ABS013 WHERE B1301='05' and B1302<>'" & strUserNum & "' and B1303='" & strUserNum & "' order by B1302 asc"
                  End If
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strMeSql)
                  If intI = 1 Then
                     With RsTemp
                        intRow = intRow + .RecordCount
                        .MoveFirst
                        Do While Not .EOF
                           txt1(2) = RsTemp.Fields("B1302"): txt1(3) = RsTemp.Fields("B1302")
                           txt1(4) = Val(RsTemp.Fields("B1315")) - 19110000 'Add By Sindy 2012/4/9
                           frm180203.m_B1301 = "" & RsTemp.Fields("B1301")
                           frm180203.m_B1302 = "" & RsTemp.Fields("B1302")
                           frm180203.m_B1303 = "" & RsTemp.Fields("B1303")
                           GetSql
                           'Call StrMenu1(False)
                           Call StrMenu_Word(False)
                           If bolfrm180203ExitForm = True Then GoTo GoToEnd
                           .MoveNext
                           'Add By Sindy 2011/11/4 檢查是否已確認完畢資料
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strMeSql)
                           If intI <> 1 Then Exit Do
                           '2011/11/4 End
                        Loop
                     End With
                  End If
               Next i
GoToEnd:
               If intRow = 0 Then nResponse = MsgBox("無資料!", vbExclamation + vbOKOnly, "員工個人資料明細確認")
            Else
               Screen.MousePointer = vbHourglass
               'Call StrMenu1(False)
               Call StrMenu_Word(False)
            End If
            
        ElseIf Index = 2 Then '電子化線上確認
            m_ShowSpecialty = True 'Add By Sindy 2014/4/11 要列印專長資料
            '檢查資料有無存在
            strSql = "SELECT count(*) FROM ABS013,Staff WHERE B1301='05' and B1302=ST01(+) " & m_StrSQL
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.Fields(0) > 0 Then
                  '提示訊息
                  strTit = "詢問"
                  strMsg = "有" & RsTemp.Fields(0) & "筆資料已存在,確定是否要重新產生資料?"
                  nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbQuestion, strTit)
                  If nResponse = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
               End If
            End If
            Screen.MousePointer = vbHourglass
            'Call StrMenu1(True)
            Call StrMenu_Word(True)
        End If
Case 1
        Unload Me
End Select
Screen.MousePointer = vbDefault
End Sub

Sub StrMenu_Word(bolEPrint As Boolean)
Dim strST14 As String 'Add By Sindy 2011/9/2
Dim strSendEmailTo As String 'Add By Sindy 2011/9/2
Dim intECnt As Integer, intPCnt As Integer 'Add By Sindy 2011/9/15
Dim bolRpt As Boolean

strTempFile = App.path & "\$$個人人事資料明細列印"
If Dir(strTempFile & "*.doc") <> "" Or Dir(strTempFile & "*.pdf") <> "" Then
   Kill strTempFile & "*.*"
End If

'員工基本資料
'Memo By Sindy 2023/12/22 修改抓新部門程式
m_str = "SELECT a0922,a.ac03,b.ac03,c.ac03,decode(ST22,'F','女','M','男','') as FM,d.ac03,Staff.* " & _
                "FROM Staff,acc090NEW,SalaryData,allcode a,allcode b,allcode c,allcode d " & _
              "WHERE ST04='1' and ST01=SD01 and ((SD02 not in('P','F') or SD02 is null) or ST01='68007') " & _
                   "AND ST93=a0921(+) " & _
                   "AND a.ac01(+)='01' AND ST20=a.ac02(+) " & _
                   "AND b.ac01(+)='02' AND ST21=b.ac02(+) " & _
                   "AND c.ac01(+)='03' AND ST37=c.ac02(+) " & _
                   "AND d.ac01(+)='06' AND ST27=d.ac02(+) " & m_StrSQL & _
             "ORDER BY ST93,ST01 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   '設定使用者所選擇的印表機成預設印表機
   PUB_SetOsDefaultPrinter Combo1
   
   If bolEPrint = True Then '電子化線上確認
      On Error GoTo ErrHand
      cnnConnection.BeginTrans
      '刪除資料
      strSql = "DELETE FROM ABS013 WHERE B1301='05' and B1302 in(SELECT B1302 FROM ABS013,Staff WHERE B1301='05' and B1302=ST01(+) " & m_StrSQL & ")"
      cnnConnection.Execute strSql
   End If
   
   With m_rs
      m_rs.MoveFirst
      
      '預設值
      iLine = 1
      strType = "" '切頁條件
      
      Do While Not m_rs.EOF
         'Add By Sindy 2011/9/2
         If Not IsNull(m_rs.Fields("ST14")) Then
            strST14 = m_rs.Fields("ST14")
         Else
            strST14 = ""
         End If
         'Modify By Sindy 2023/1/10 And Mid(m_rs.Fields("ST01"), 4, 1) <> "9"
         If bolEPrint = True And strST14 <> "99997" And Mid(m_rs.Fields("ST01"), 4, 1) <> "9" Then '執行E化並且有公司信箱者,產生待確認資料
            'Modify By Sindy 2012/4/9 +B1315
            strSql = "insert into ABS013(B1301,B1302,B1303,B1315) " & _
                     "values('05'," & CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & _
                     CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & DBDATE(txt1(4)) & ")"
            cnnConnection.Execute strSql
            
            intECnt = intECnt + 1
            '記錄要發通知確認的E-Mail人員
            strSendEmailTo = strSendEmailTo & CheckStr(m_rs.Fields("ST01")) & ";"
            GoTo GoToNext
         End If
         '2011/9/2 End
         
         'Add By Sindy 2022/1/24
         If intChoose = 1 Then '出缺勤系統
            Load frmpic002
            frmpic002.Label1.Caption = "電子檔產生中...請稍候..."
            frmpic002.Show
            frmpic002.ZOrder 0
         End If
         '2022/1/24 END
         
         intPCnt = intPCnt + 1
         For m_i = 1 To 62 '60
             strTemp(m_i) = ""
         Next m_i
         
         strTemp(1) = "部　　   門：" & CheckStr(m_rs.Fields("ST93")) & "　" & CheckStr(m_rs.Fields(0))
         strTemp(2) = "姓　　   名：" & CheckStr(m_rs.Fields("ST01")) & "　" & CheckStr(m_rs.Fields("ST02"))
         strTemp(3) = "職　　   稱：" & CheckStr(m_rs.Fields("ST20")) & "　" & CheckStr(m_rs.Fields(1))
         strTemp(4) = "職　　   位：" & CheckStr(m_rs.Fields("ST21")) & "　" & CheckStr(m_rs.Fields(2))
         strTemp(5) = "性　　   別：" & CheckStr(m_rs.Fields(4))
         strTemp(6) = "血　　   型：" & CheckStr(m_rs.Fields("ST25"))
         strTemp(7) = "出    生   地：" & CheckStr(m_rs.Fields(5))
         strTemp(8) = "身 份 證 號：" & CheckStr(m_rs.Fields("ST26"))
         strTemp(9) = "出 生 日 期：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(m_rs.Fields("ST23"))))
         strTemp(10) = "入 所 日 期：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(m_rs.Fields("ST13"))))
         strTemp(11) = "最 高 學 歷：" & CheckStr(m_rs.Fields("ST37")) & "　" & CheckStr(m_rs.Fields(3))
         strTemp(12) = "畢 業 學 校：" & CheckStr(m_rs.Fields("ST38"))
         strTemp(13) = "科　   　系：" & CheckStr(m_rs.Fields("ST39"))
         strTemp(14) = "通 訊 電 話：" & CheckStr(m_rs.Fields("ST09"))
         strTemp(15) = "傳 真 號 碼：" & CheckStr(m_rs.Fields("ST10"))
         strTemp(16) = "行 動 電 話：" & CheckStr(m_rs.Fields("ST19"))
         strTemp(17) = "通 訊 地 址：" & CheckStr(m_rs.Fields("ST33")) & "　" & CheckStr(m_rs.Fields("ST08"))
         strTemp(18) = "戶 籍 電 話：" & CheckStr(m_rs.Fields("ST35"))
         strTemp(19) = "戶 籍 地 址：" & CheckStr(m_rs.Fields("ST36")) & "　" & CheckStr(m_rs.Fields("ST34"))
         strTemp(21) = "結 婚 日 期：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(m_rs.Fields("ST41"))))
         
         '眷屬資料
         m_QueryCnt = 0
         'Modify By Sindy 2010/5/6 已有刪除日期的資料不顯示
   '            m_str = "select decode(SR05,'F','女','M','男','') as FM,nvl(sr08,'') as T08,staff_relation.* " & _
   '                          "from staff_relation " & _
   '                          "where sr01='" & CheckStr(m_rs.Fields("ST01")) & "' " & _
   '                          "order by sr03,sr06 "
         m_str = "select decode(SR05,'F','女','M','男','') as FM,nvl(sr08,'') as T08,staff_relation.* " & _
                       "from staff_relation " & _
                       "where sr01='" & CheckStr(m_rs.Fields("ST01")) & "' and (sr12 is null or sr12=0) " & _
                       "order by sr03,sr06 "
         If m_rs2.State = 1 Then m_rs2.Close
         m_rs2.CursorLocation = adUseClient
         m_rs2.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs2.EOF And Not m_rs2.BOF Then
            m_rs2.MoveFirst
            Do While Not m_rs2.EOF
               Select Case CheckStr(m_rs2.Fields("sr03"))
               Case "1"
                       'Modify by Morgan 2009/7/7 sr12 改放 'Y'
                       'If IsNull(m_rs2.Fields("sr13")) = False Then
                       If IsNull(m_rs2.Fields("sr13")) = True Then
                         strTemp(34) = "父 親 姓 名：" & CheckStr(m_rs2.Fields("sr04"))
                         strTemp(35) = "電　　   話：" & CheckStr(m_rs2.Fields("sr09"))
                         strTemp(36) = "通 訊 地 址：" & CheckStr(m_rs2.Fields("sr10")) & "　" & CheckStr(m_rs2.Fields("sr11"))
                       Else
                         strTemp(34) = "父 親 姓 名：" & CheckStr(m_rs2.Fields("sr04")) & " (歿)"
                         strTemp(35) = "電　　   話："
                         strTemp(36) = "通 訊 地 址："
                       End If
               Case "2"
                       'Modify by Morgan 2009/7/7 sr12 改放 'Y'
                       'If IsNull(m_rs2.Fields("sr13")) = False Then
                       If IsNull(m_rs2.Fields("sr13")) = True Then
                         strTemp(37) = "母 親 姓 名：" & CheckStr(m_rs2.Fields("sr04"))
                         strTemp(38) = "電　　   話：" & CheckStr(m_rs2.Fields("sr09"))
                         strTemp(39) = "通 訊 地 址：" & CheckStr(m_rs2.Fields("sr10")) & "　" & CheckStr(m_rs2.Fields("sr11"))
                       Else
                         strTemp(37) = "母 親 姓 名：" & CheckStr(m_rs2.Fields("sr04")) & " (歿)"
                         strTemp(38) = "電　　   話："
                         strTemp(39) = "通 訊 地 址："
                       End If
               Case "3"
                       strTemp(20) = "配 偶 姓 名：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
               Case "4"
                       If strTemp(22) = "" And strTemp(23) = "" And strTemp(24) = "" Then
                         strTemp(22) = "子 女 １   ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                         strTemp(23) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                         strTemp(24) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                       ElseIf strTemp(25) = "" And strTemp(26) = "" And strTemp(27) = "" Then
                         strTemp(25) = "子 女 ２   ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                         strTemp(26) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                         strTemp(27) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                       ElseIf strTemp(28) = "" And strTemp(29) = "" And strTemp(30) = "" Then
                         strTemp(28) = "子 女 ３   ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                         strTemp(29) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                         strTemp(30) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                       Else
                         strTemp(31) = "子 女 ４   ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                         strTemp(32) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                         strTemp(33) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                       End If
                End Select
                '健保眷屬
                'Modify by Morgan 2009/6/29
                'If CheckStr(m_rs2.Fields("T08")) = "" Then
                If CheckStr(m_rs2.Fields("T08")) = "Y" Then
                   m_QueryCnt = m_QueryCnt + 1
                   If strTemp(40) = "" And strTemp(41) = "" Then
                       strTemp(40) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(41) = CheckStr(m_rs2.Fields("sr07"))
                   ElseIf strTemp(42) = "" And strTemp(43) = "" Then
                       strTemp(42) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(43) = CheckStr(m_rs2.Fields("sr07"))
                   ElseIf strTemp(44) = "" And strTemp(45) = "" Then
                       strTemp(44) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(45) = CheckStr(m_rs2.Fields("sr07"))
                   ElseIf strTemp(46) = "" And strTemp(47) = "" Then
                       strTemp(46) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(47) = CheckStr(m_rs2.Fields("sr07"))
                   ElseIf strTemp(48) = "" And strTemp(49) = "" Then
                       strTemp(48) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(49) = CheckStr(m_rs2.Fields("sr07"))
                   ElseIf strTemp(50) = "" And strTemp(51) = "" Then
                       strTemp(50) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(51) = CheckStr(m_rs2.Fields("sr07"))
                   ElseIf strTemp(52) = "" And strTemp(53) = "" Then
                       strTemp(52) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(53) = CheckStr(m_rs2.Fields("sr07"))
                   Else
                       strTemp(54) = CheckStr(m_rs2.Fields("sr04"))
                       strTemp(55) = CheckStr(m_rs2.Fields("sr07"))
                   End If
                End If
                m_rs2.MoveNext
            Loop
         End If
         If strTemp(20) = "" Then strTemp(20) = "配 偶 姓 名："
         If strTemp(22) = "" Then strTemp(22) = "子 女 １   ："
         If strTemp(23) = "" Then strTemp(23) = "性　　   別："
         If strTemp(24) = "" Then strTemp(24) = "出 生 日 期："
         If strTemp(25) = "" Then strTemp(25) = "子 女 ２   ："
         If strTemp(26) = "" Then strTemp(26) = "性　　   別："
         If strTemp(27) = "" Then strTemp(27) = "出 生 日 期："
         If strTemp(28) = "" Then strTemp(28) = "子 女 ３   ："
         If strTemp(29) = "" Then strTemp(29) = "性　　   別："
         If strTemp(30) = "" Then strTemp(30) = "出 生 日 期："
         If strTemp(31) = "" Then strTemp(31) = "子 女 ４   ："
         If strTemp(32) = "" Then strTemp(32) = "性　　   別："
         If strTemp(33) = "" Then strTemp(33) = "出 生 日 期："
         If strTemp(34) = "" Then strTemp(34) = "父 親 姓 名："
         If strTemp(35) = "" Then strTemp(35) = "電　　   話："
         If strTemp(36) = "" Then strTemp(36) = "通 訊 地 址："
         If strTemp(37) = "" Then strTemp(37) = "母 親 姓 名："
         If strTemp(38) = "" Then strTemp(38) = "電　　   話："
         If strTemp(39) = "" Then strTemp(39) = "通 訊 地 址："
         'Add By Sindy 98/04/13
         strTemp(56) = "E-Mail：" & CheckStr(m_rs.Fields("ST18"))
         '98/04/13 End
         
         'Add By Sindy 2014/3/12 +專長資料
         m_HaveSpecialty = False
         m_str = "select * from staff_specialty " & _
                 "where ss01='" & CheckStr(m_rs.Fields("ST01")) & "'"
         If m_rs2.State = 1 Then m_rs2.Close
         m_rs2.CursorLocation = adUseClient
         m_rs2.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs2.EOF And Not m_rs2.BOF Then
            m_rs2.MoveFirst
            strTemp(57) = CheckStr("" & m_rs2.Fields("SS02"))
            If Trim(strTemp(57)) <> "" Then m_HaveSpecialty = True
            strTemp(58) = CheckStr("" & m_rs2.Fields("SS03"))
            If Trim(strTemp(58)) <> "" Then m_HaveSpecialty = True
            strTemp(59) = CheckStr("" & m_rs2.Fields("SS04"))
            If Trim(strTemp(59)) <> "" Then m_HaveSpecialty = True
            strTemp(60) = CheckStr("" & m_rs2.Fields("SS05"))
            If Trim(strTemp(60)) <> "" Then m_HaveSpecialty = True
            strTemp(61) = CheckStr("" & m_rs2.Fields("SS06"))
            If Trim(strTemp(61)) <> "" Then m_HaveSpecialty = True
            strTemp(62) = CheckStr("" & m_rs2.Fields("SS07"))
            If Trim(strTemp(62)) <> "" Then m_HaveSpecialty = True
         End If
         '2014/3/12 END
         
         bolRpt = WordEdit(CheckStr(m_rs.Fields("ST01"))) '列印報表
         strType = CheckStr(m_rs.Fields("ST01"))
GoToNext:
         m_rs.MoveNext
      Loop
   End With
   
   If bolRpt = True Then
      If intChoose <> 1 Then
         g_WordAp.ActiveDocument.PrintOut
      Else
         g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strTempFile & ".pdf", ExportFormat:=17, OpenAfterExport:=False
      End If
      g_WordAp.Documents.Close
      g_WordAp.Quit
      Set g_WordAp = Nothing
   End If
   
   PUB_SetOsDefaultPrinter strPrinter '復原系統預設印表機
Else
   Screen.MousePointer = vbDefault
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If

'Add By Sindy 2011/9/2
If intChoose = 1 Then '出缺勤系統
'   SetPic m_iPages + 1
'   frm180203.m_ImageW = m_Device.Width
'   frm180203.m_ImageH = m_Device.Height
'   frm180203.m_iPages = m_iPages + 1
   frm180203.Caption = "員工個人資料明細確認"
   If Dir(strTempFile & ".pdf") <> "" Then
      frm180203.WebBrowser1.Navigate strTempFile & ".pdf"
   Else
      frm180203.WebBrowser1.Navigate "about:blank"
   End If
   Unload frmpic002
   'Me.Hide
   frm180203.Show vbModal '強制回應表單
   Unload Me
Else
   If bolEPrint = True Then
      cnnConnection.CommitTrans
      '發通知確認的E-Mail
      If strSendEmailTo <> "" Then
         If Right(strSendEmailTo, 1) = ";" Then strSendEmailTo = Left(strSendEmailTo, Len(strSendEmailTo) - 1)
         'Modify By Sindy 2013/3/5
         'PUB_SendMail strUserNum, strSendEmailTo, "", "員工個人資料明細待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行員工個人資料明細確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
         'Modify By Sindy 2014/3/11
         If m_StrSQL <> "" Then
            PUB_SendMail strUserNum, strSendEmailTo, "", "員工個人資料明細待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行員工個人資料明細確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
         Else
         '2014/3/11 END
            PUB_SendMail strUserNum, "taie_alluser@taie.com.tw", "", "員工個人資料明細待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行員工個人資料明細確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
            '2013/3/5 End
         End If
      End If
   End If
   
   'ShowPrintOk
   If intECnt = 0 Then
      MsgBox "列印紙本 " & intPCnt & " 筆完成!!", , "列印成功"
   Else
      MsgBox "列印紙本 " & intPCnt & " 筆及產生電子檔 " & intECnt & " 筆完成!!", , "列印成功"
   End If
End If

Exit Sub

ErrHand:
   If Err.Number <> 0 Then
      If bolEPrint = True Then '電子化線上確認
         cnnConnection.RollbackTrans
         MsgBox " 更新失敗！" & vbCrLf & Err.Description
      End If
   End If
'2011/9/2 End
End Sub

'Add By Sindy 2022/2/23
Private Function WordEdit(strST01 As String) As Boolean
'+信頭
Dim stFileName As String
Dim iPicNo As Integer
Dim iPicNo2 As Integer
Dim oShape

Dim rsTmp As New ADODB.Recordset
Dim i As Integer, k As Integer
Dim strFAX As String
Dim strNo As String, strNote As String 'Add By Sindy 2019/3/21
Dim strManyTData As String 'Add By Sindy 2019/9/24
Dim intCnt As Integer
Dim strText As String
Dim bolFirst As Boolean
   
On Error GoTo ERRORSECTION1
   
   WordEdit = True
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Visible = False
   g_WordAp.WindowState = wdWindowStateMaximize 'wdWindowStateMinimize  wdWindowStateMaximize
   With g_WordAp
   
      If Dir(strTempFile & ".doc") = "" Then
         bolFirst = True
         g_WordAp.Documents.add.SaveAs strTempFile & ".doc"
      Else
         bolFirst = False
         .Selection.InsertBreak Type:=wdPageBreak
      End If
      
      '標題
      .Selection.Font.Name = "新細明體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1.5)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1.5)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      '配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行高
      .Selection.ParagraphFormat.LineSpacing = 20 '行高
      .Selection.Font.Size = 14
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
      .Selection.TypeText "個人人事資料明細表"
      .Selection.TypeParagraph
      .Selection.Font.Size = 12
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
      If m_ShowSpecialty = True Then
         strText = "頁　　次：1 / 2"
      Else
         strText = "頁　　次：1 / 1"
      End If
      .Selection.TypeText "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)) & _
         "　　　　　　　　　　　　　　　　　　　　　　　" & strText
      .Selection.TypeParagraph
      
      '明細
      Dim m_j As Integer ', i As Integer, intCnt As Integer
      
      '員工基本資料
      For m_j = 1 To 19
         If m_j <= 16 And m_j <> 12 Then
            strText = convForm(CheckStr(strTemp(m_j)), 36) '30
            If m_j = 1 Then
               strText = strText & "　　　　　　　　　　　　　　　|#右代表圖#|"
            End If
         Else
            strText = strTemp(m_j)
         End If
         If m_j = 3 Or m_j = 5 Or m_j = 9 Or m_j = 14 Then
            m_j = m_j + 1
            strText = strText & "" & strTemp(m_j)
         ElseIf m_j = 16 Then
            strText = strText & "" & strTemp(56)
         End If
         .Selection.TypeParagraph
         .Selection.TypeText strText
      Next m_j
      .Selection.TypeParagraph
      .Selection.TypeText "***接續***"
      
      '插入照片
      m_MySt(1) = "000": m_MySt(2) = strST01: m_MySt(3) = "0": m_MySt(4) = "00"
      Call PUB_AddInPicToWordR(g_WordAp)
      
      Call WordFindText(g_WordAp, "***接續***")
      '.Selection.TypeParagraph
      .Selection.TypeText String(125, "-")
      
      '配偶
      .Selection.TypeParagraph
      .Selection.TypeText convForm(CheckStr(strTemp(20)), 30) & strTemp(21)
      .Selection.TypeParagraph
      
      '子女
      intCnt = 21
      For m_j = 1 To 4
         strText = ""
         For i = 1 To 3
            intCnt = intCnt + 1
            If i = 3 Then
               strText = strText & strTemp(intCnt)
            Else
               strText = strText & convForm(CheckStr(strTemp(intCnt)), 25) '30
            End If
         Next i
         .Selection.TypeParagraph
         .Selection.TypeText strText
      Next m_j
      .Selection.TypeParagraph
      
      '父母親
      For m_j = 1 To 6
         If m_j = 1 Or m_j = 4 Then
            strText = convForm(CheckStr(strTemp(m_j + 33)), 30)
         Else
            strText = strTemp(m_j + 33)
         End If
         If m_j = 1 Or m_j = 4 Then
            m_j = m_j + 1
            strText = strText & strTemp(m_j + 33)
         End If
         .Selection.TypeParagraph
         .Selection.TypeText strText
      Next m_j
      
      .Selection.TypeParagraph
      .Selection.TypeText String(125, "-")
      
      '健保眷屬
      .Selection.TypeParagraph
      .Selection.TypeText convForm(CheckStr(""), 16) & convForm(CheckStr("姓　　名"), 16) & convForm(CheckStr("身份證字號"), 16) & convForm(CheckStr("姓　　名"), 16) & convForm(CheckStr("身份證字號"), 16)
      .Selection.TypeParagraph
      .Selection.TypeText ""
      If m_QueryCnt = 0 Then
         .Selection.TypeText "健保眷屬明細：　(無)"
      Else
         .Selection.TypeText convForm(CheckStr("健保眷屬明細："), 16)
      End If
      intCnt = 39
      For m_j = 1 To 4
         strText = ""
         For i = 1 To 4
            intCnt = intCnt + 1
            strText = strText & convForm(CheckStr(strTemp(intCnt)), 16)
         Next i
         .Selection.TypeText strText
      Next m_j
      
      '頁尾
      If bolFirst = True Then
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.Font.Name = "新細明體"
         .Selection.Font.Size = 12
         'Add By Sindy 2011/9/2
         If intChoose = 1 Then '出缺勤系統
            .Selection.TypeText "請於 " & ChangeTStringToTDateString(txt1(4)) & " 前核對後執行確認；如資料有誤，請以E-Mail通知人事處。"
         Else
            .Selection.TypeText "請於 " & ChangeTStringToTDateString(txt1(4)) & " 前核對後簽章繳回人事處！"
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End If
      
      'Add By Sindy 2014/3/12 +第二頁-專長資料
      If m_ShowSpecialty = True Then
         .Selection.InsertBreak Type:=wdPageBreak
         .Selection.Font.Size = 14
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
         .Selection.TypeText "個人人事資料明細表"
         .Selection.TypeParagraph
         .Selection.Font.Size = 12
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
   '      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
         strText = "頁　　次：2 / 2"
         .Selection.TypeText "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)) & _
            "　　　　　　　　　　　　　　　　　　　　　　　" & strText
         .Selection.TypeParagraph
         
         .Selection.TypeParagraph
         .Selection.TypeText "教育與訓練："
         .Selection.TypeParagraph
         .Selection.TypeText IIf(strTemp(57) <> "", strTemp(57), "（無）")
         .Selection.TypeParagraph
         
         .Selection.TypeParagraph
         .Selection.TypeText "語文能力檢定："
         .Selection.TypeParagraph
         .Selection.TypeText IIf(strTemp(58) <> "", strTemp(58), "（無）")
         .Selection.TypeParagraph
                 
         .Selection.TypeParagraph
         .Selection.TypeText "智財專業證照：" '專業資格／技能證照 Modify By Sindy 2024/4/17
         .Selection.TypeParagraph
         .Selection.TypeText IIf(strTemp(59) <> "", strTemp(59), "（無）")
         .Selection.TypeParagraph
                 
         .Selection.TypeParagraph
         .Selection.TypeText "著作／發明／創作："
         .Selection.TypeParagraph
         .Selection.TypeText IIf(strTemp(60) <> "", strTemp(60), "（無）")
         .Selection.TypeParagraph
         
         .Selection.TypeParagraph
         .Selection.TypeText "非智財專業證照：" '電腦技能證照 Modify By Sindy 2024/4/17
         .Selection.TypeParagraph
         .Selection.TypeText IIf(strTemp(61) <> "", strTemp(61), "（無）")
         .Selection.TypeParagraph
         
         .Selection.TypeParagraph
         .Selection.TypeText "專業／社會團體會員："
         .Selection.TypeParagraph
         .Selection.TypeText IIf(strTemp(62) <> "", strTemp(62), "（無）")
         .Selection.TypeParagraph
      End If
      '2014/3/12 END
      
   End With
   g_WordAp.Documents.Save
   
   Exit Function
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
'            If bolRetry = False Then
'               bolRetry = True
'               Resume
'            End If
         'Add By Sindy 2013/1/28
         Case 5152
            Resume Next
         '2013/1/28 End
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
            WordEdit = False
      End Select
   End If
End Function

'尋找Word檔中文字:Find/清除/置換文字 或 貼上
Private Sub WordFindText(g_WordAp As Word.Application, strFindText As String, Optional strReplaceText As String = "")
Dim bolResult As Boolean
   
   If Trim(strFindText) = "" Then Exit Sub
   With g_WordAp
'      .Selection.WholeStory
'      .Selection.Copy
      .Selection.GoTo what:=wdGoToPage, which:=wdGoToPrevious, Count:=3
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strFindText
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      bolResult = .Selection.Find.Execute
      If bolResult = True Then
         .Selection.Delete
         If strReplaceText = "複製圖片" Then
            .Selection.Paste 'Format '(wdSingleCellText)
         Else
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strReplaceText
         End If
      End If
   End With
End Sub

Sub StrMenu1(bolEPrint As Boolean)
Dim strST14 As String 'Add By Sindy 2011/9/2
Dim strSendEmailTo As String 'Add By Sindy 2011/9/2
Dim intECnt As Integer, intPCnt As Integer 'Add By Sindy 2011/9/15

Set Printer = Printers(Combo1.ListIndex)

'Add By Sindy 2011/9/2
m_iPages = 0
If intChoose = 1 Then '出缺勤系統
   Set m_Device = Picture1
   m_Device.AutoRedraw = True
   m_Device.Width = 9048 '11899
   m_Device.Height = 12000 '5700 '16838
   m_Device.AutoSize = True
   douExtRate = 0.7 'm_Device.Height / 8142 '16836
   DelPic
Else
   Set m_Device = Printer
   m_Device.EndDoc
   m_Device.Orientation = 1 '1.直印 2.橫印
   'm_Device.PaperSize = PUB_GetPaperSize(5) '催款單2
   m_Device.PaperSize = 9  'PDF (A4) Modify By Sindy 2012/6/21
   douExtRate = 1
End If

'員工基本資料
m_str = "SELECT a0922,a.ac03,b.ac03,c.ac03,decode(ST22,'F','女','M','男','') as FM,d.ac03,Staff.* " & _
                "FROM Staff,acc090NEW,SalaryData,allcode a,allcode b,allcode c,allcode d " & _
              "WHERE ST04='1' and ST01=SD01 and ((SD02 not in('P','F') or SD02 is null) or ST01='68007') " & _
                   "AND ST93=a0921(+) " & _
                   "AND a.ac01(+)='01' AND ST20=a.ac02(+) " & _
                   "AND b.ac01(+)='02' AND ST21=b.ac02(+) " & _
                   "AND c.ac01(+)='03' AND ST37=c.ac02(+) " & _
                   "AND d.ac01(+)='06' AND ST27=d.ac02(+) " & m_StrSQL & _
             "ORDER BY ST93,ST01 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   If bolEPrint = True Then '電子化線上確認
      On Error GoTo ErrHand
      cnnConnection.BeginTrans
      '刪除資料
      strSql = "DELETE FROM ABS013 WHERE B1301='05' and B1302 in(SELECT B1302 FROM ABS013,Staff WHERE B1301='05' and B1302=ST01(+) " & m_StrSQL & ")"
      cnnConnection.Execute strSql
   End If
   
    With m_rs
        m_rs.MoveFirst
        
        '預設值
        iLine = 1
        strType = "" '切頁條件
        
        Do While Not m_rs.EOF
            
            'Add By Sindy 2011/9/2
            strST14 = ""
            If Not IsNull(m_rs.Fields("ST14")) Then strST14 = m_rs.Fields("ST14")
            'Modify By Sindy 2023/1/10 And Mid(m_rs.Fields("ST01"), 4, 1) <> "9"
            If bolEPrint = True And strST14 <> "99997" And Mid(m_rs.Fields("ST01"), 4, 1) <> "9" Then '執行E化並且有公司信箱者,產生待確認資料
               'Modify By Sindy 2012/4/9 +B1315
               strSql = "insert into ABS013(B1301,B1302,B1303,B1315) " & _
                        "values('05'," & CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & _
                        CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & DBDATE(txt1(4)) & ")"
               cnnConnection.Execute strSql
               
               intECnt = intECnt + 1
               '記錄要發通知確認的E-Mail人員
               strSendEmailTo = strSendEmailTo & CheckStr(m_rs.Fields("ST01")) & ";"
               GoTo GoToNext
            End If
            '2011/9/2 End
            
            intPCnt = intPCnt + 1
            For m_i = 1 To 62 '60
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = "部　　   門：" & CheckStr(m_rs.Fields("ST93")) & "　" & CheckStr(m_rs.Fields(0))
            strTemp(2) = "姓　　   名：" & CheckStr(m_rs.Fields("ST01")) & "　" & CheckStr(m_rs.Fields("ST02"))
            strTemp(3) = "職　　   稱：" & CheckStr(m_rs.Fields("ST20")) & "　" & CheckStr(m_rs.Fields(1))
            strTemp(4) = "職　　   位：" & CheckStr(m_rs.Fields("ST21")) & "　" & CheckStr(m_rs.Fields(2))
            strTemp(5) = "性　　   別：" & CheckStr(m_rs.Fields(4))
            strTemp(6) = "血　　   型：" & CheckStr(m_rs.Fields("ST25"))
            strTemp(7) = "出    生   地：" & CheckStr(m_rs.Fields(5))
            strTemp(8) = "身 份 證 號：" & CheckStr(m_rs.Fields("ST26"))
            strTemp(9) = "出 生 日 期：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(m_rs.Fields("ST23"))))
            strTemp(10) = "入 所 日 期：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(m_rs.Fields("ST13"))))
            strTemp(11) = "最 高 學 歷：" & CheckStr(m_rs.Fields("ST37")) & "　" & CheckStr(m_rs.Fields(3))
            strTemp(12) = "畢 業 學 校：" & CheckStr(m_rs.Fields("ST38"))
            strTemp(13) = "科　   　系：" & CheckStr(m_rs.Fields("ST39"))
            strTemp(14) = "通 訊 電 話：" & CheckStr(m_rs.Fields("ST09"))
            strTemp(15) = "傳 真 號 碼：" & CheckStr(m_rs.Fields("ST10"))
            strTemp(16) = "行 動 電 話：" & CheckStr(m_rs.Fields("ST19"))
            strTemp(17) = "通 訊 地 址：" & CheckStr(m_rs.Fields("ST33")) & "　" & CheckStr(m_rs.Fields("ST08"))
            strTemp(18) = "戶 籍 電 話：" & CheckStr(m_rs.Fields("ST35"))
            strTemp(19) = "戶 籍 地 址：" & CheckStr(m_rs.Fields("ST36")) & "　" & CheckStr(m_rs.Fields("ST34"))
            strTemp(21) = "結 婚 日 期：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(m_rs.Fields("ST41"))))
            
            '眷屬資料
            m_QueryCnt = 0
            'Modify By Sindy 2010/5/6 已有刪除日期的資料不顯示
'            m_str = "select decode(SR05,'F','女','M','男','') as FM,nvl(sr08,'') as T08,staff_relation.* " & _
'                          "from staff_relation " & _
'                          "where sr01='" & CheckStr(m_rs.Fields("ST01")) & "' " & _
'                          "order by sr03,sr06 "
            m_str = "select decode(SR05,'F','女','M','男','') as FM,nvl(sr08,'') as T08,staff_relation.* " & _
                          "from staff_relation " & _
                          "where sr01='" & CheckStr(m_rs.Fields("ST01")) & "' and (sr12 is null or sr12=0) " & _
                          "order by sr03,sr06 "
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
                m_rs2.MoveFirst
                Do While Not m_rs2.EOF
                    Select Case CheckStr(m_rs2.Fields("sr03"))
                    Case "1"
                            'Modify by Morgan 2009/7/7 sr12 改放 'Y'
                            'If IsNull(m_rs2.Fields("sr13")) = False Then
                            If IsNull(m_rs2.Fields("sr13")) = True Then
                              strTemp(34) = "父 親 姓 名：" & CheckStr(m_rs2.Fields("sr04"))
                              strTemp(35) = "電　　   話：" & CheckStr(m_rs2.Fields("sr09"))
                              strTemp(36) = "通 訊 地 址：" & CheckStr(m_rs2.Fields("sr10")) & "　" & CheckStr(m_rs2.Fields("sr11"))
                            Else
                              strTemp(34) = "父 親 姓 名：" & CheckStr(m_rs2.Fields("sr04")) & " (歿)"
                              strTemp(35) = "電　　   話："
                              strTemp(36) = "通 訊 地 址："
                            End If
                    Case "2"
                            'Modify by Morgan 2009/7/7 sr12 改放 'Y'
                            'If IsNull(m_rs2.Fields("sr13")) = False Then
                            If IsNull(m_rs2.Fields("sr13")) = True Then
                              strTemp(37) = "母 親 姓 名：" & CheckStr(m_rs2.Fields("sr04"))
                              strTemp(38) = "電　　   話：" & CheckStr(m_rs2.Fields("sr09"))
                              strTemp(39) = "通 訊 地 址：" & CheckStr(m_rs2.Fields("sr10")) & "　" & CheckStr(m_rs2.Fields("sr11"))
                            Else
                              strTemp(37) = "母 親 姓 名：" & CheckStr(m_rs2.Fields("sr04")) & " (歿)"
                              strTemp(38) = "電　　   話："
                              strTemp(39) = "通 訊 地 址："
                            End If
                    Case "3"
                            strTemp(20) = "配 偶 姓 名：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                    Case "4"
                            If strTemp(22) = "" And strTemp(23) = "" And strTemp(24) = "" Then
                              strTemp(22) = "子 女 １     ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                              strTemp(23) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                              strTemp(24) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            ElseIf strTemp(25) = "" And strTemp(26) = "" And strTemp(27) = "" Then
                              strTemp(25) = "子 女 ２     ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                              strTemp(26) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                              strTemp(27) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            ElseIf strTemp(28) = "" And strTemp(29) = "" And strTemp(30) = "" Then
                              strTemp(28) = "子 女 ３     ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                              strTemp(29) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                              strTemp(30) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            Else
                              strTemp(31) = "子 女 ４     ：" & CheckStr(m_rs2.Fields("sr04")) & IIf(Not IsNull(m_rs2.Fields("sr13")) = True, " (歿)", "")
                              strTemp(32) = "性　　   別：" & CheckStr(m_rs2.Fields("FM"))
                              strTemp(33) = "出 生 日 期：" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            End If
                     End Select
                     '健保眷屬
                     'Modify by Morgan 2009/6/29
                     'If CheckStr(m_rs2.Fields("T08")) = "" Then
                     If CheckStr(m_rs2.Fields("T08")) = "Y" Then
                        m_QueryCnt = m_QueryCnt + 1
                        If strTemp(40) = "" And strTemp(41) = "" Then
                            strTemp(40) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(41) = CheckStr(m_rs2.Fields("sr07"))
                        ElseIf strTemp(42) = "" And strTemp(43) = "" Then
                            strTemp(42) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(43) = CheckStr(m_rs2.Fields("sr07"))
                        ElseIf strTemp(44) = "" And strTemp(45) = "" Then
                            strTemp(44) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(45) = CheckStr(m_rs2.Fields("sr07"))
                        ElseIf strTemp(46) = "" And strTemp(47) = "" Then
                            strTemp(46) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(47) = CheckStr(m_rs2.Fields("sr07"))
                        ElseIf strTemp(48) = "" And strTemp(49) = "" Then
                            strTemp(48) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(49) = CheckStr(m_rs2.Fields("sr07"))
                        ElseIf strTemp(50) = "" And strTemp(51) = "" Then
                            strTemp(50) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(51) = CheckStr(m_rs2.Fields("sr07"))
                        ElseIf strTemp(52) = "" And strTemp(53) = "" Then
                            strTemp(52) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(53) = CheckStr(m_rs2.Fields("sr07"))
                        Else
                            strTemp(54) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(55) = CheckStr(m_rs2.Fields("sr07"))
                        End If
                     End If
                     m_rs2.MoveNext
                Loop
            End If
            If strTemp(20) = "" Then strTemp(20) = "配 偶 姓 名："
            If strTemp(22) = "" Then strTemp(22) = "子 女 １     ："
            If strTemp(23) = "" Then strTemp(23) = "性　　   別："
            If strTemp(24) = "" Then strTemp(24) = "出 生 日 期："
            If strTemp(25) = "" Then strTemp(25) = "子 女 ２     ："
            If strTemp(26) = "" Then strTemp(26) = "性　　   別："
            If strTemp(27) = "" Then strTemp(27) = "出 生 日 期："
            If strTemp(28) = "" Then strTemp(28) = "子 女 ３     ："
            If strTemp(29) = "" Then strTemp(29) = "性　　   別："
            If strTemp(30) = "" Then strTemp(30) = "出 生 日 期："
            If strTemp(31) = "" Then strTemp(31) = "子 女 ４     ："
            If strTemp(32) = "" Then strTemp(32) = "性　　   別："
            If strTemp(33) = "" Then strTemp(33) = "出 生 日 期："
            If strTemp(34) = "" Then strTemp(34) = "父 親 姓 名："
            If strTemp(35) = "" Then strTemp(35) = "電　　   話："
            If strTemp(36) = "" Then strTemp(36) = "通 訊 地 址："
            If strTemp(37) = "" Then strTemp(37) = "母 親 姓 名："
            If strTemp(38) = "" Then strTemp(38) = "電　　   話："
            If strTemp(39) = "" Then strTemp(39) = "通 訊 地 址："
            'Add By Sindy 98/04/13
            strTemp(56) = "E-Mail：" & CheckStr(m_rs.Fields("ST18"))
            '98/04/13 End
            
            'Add By Sindy 2014/3/12 +專長資料
            m_HaveSpecialty = False
            m_str = "select * from staff_specialty " & _
                    "where ss01='" & CheckStr(m_rs.Fields("ST01")) & "'"
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
                m_rs2.MoveFirst
                strTemp(57) = CheckStr("" & m_rs2.Fields("SS02"))
                If Trim(strTemp(57)) <> "" Then m_HaveSpecialty = True
                strTemp(58) = CheckStr("" & m_rs2.Fields("SS03"))
                If Trim(strTemp(58)) <> "" Then m_HaveSpecialty = True
                strTemp(59) = CheckStr("" & m_rs2.Fields("SS04"))
                If Trim(strTemp(59)) <> "" Then m_HaveSpecialty = True
                strTemp(60) = CheckStr("" & m_rs2.Fields("SS05"))
                If Trim(strTemp(60)) <> "" Then m_HaveSpecialty = True
                strTemp(61) = CheckStr("" & m_rs2.Fields("SS06"))
                If Trim(strTemp(61)) <> "" Then m_HaveSpecialty = True
                strTemp(62) = CheckStr("" & m_rs2.Fields("SS07"))
                If Trim(strTemp(62)) <> "" Then m_HaveSpecialty = True
            End If
            '2014/3/12 END
            
            If iLine > 50 Or iLine = 1 Or _
               strType <> strTemp(1) Then
               If (strType <> "" And strType <> strTemp(1)) Then
                  '小計
               End If
               'If .AbsolutePosition <> .RecordCount Then
                  If strType <> "" Then
                     'Add By Sindy 2011/9/2
                     If intChoose = 1 Then '出缺勤系統
                        m_iPages = m_iPages + 1
                        If m_iPages > 1 Then
                           SetPic m_iPages - 1
                        End If
                        '2010/9/16 End
                     Else
                        m_Device.NewPage
                     End If
                  End If
                  iLine = 1
                  Call PrintTitle(CheckStr(m_rs.Fields("ST01")))  '列印表頭
               'End If
            End If
            
            PrintDetail '列印表中
            strType = CheckStr(m_rs.Fields("ST01"))
GoToNext:
            m_rs.MoveNext
        Loop
    End With
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
'Add By Sindy 2011/9/2
If intChoose = 1 Then '出缺勤系統
   SetPic m_iPages + 1
   frm180203.m_ImageW = m_Device.Width
   frm180203.m_ImageH = m_Device.Height
   frm180203.m_iPages = m_iPages + 1
   frm180203.Caption = "員工個人資料明細確認"
   Me.Hide
   frm180203.Show vbModal '強制回應表單
   Unload Me
Else
   If bolEPrint = True Then
      cnnConnection.CommitTrans
      '發通知確認的E-Mail
      If strSendEmailTo <> "" Then
         If Right(strSendEmailTo, 1) = ";" Then strSendEmailTo = Left(strSendEmailTo, Len(strSendEmailTo) - 1)
         'Modify By Sindy 2013/3/5
         'PUB_SendMail strUserNum, strSendEmailTo, "", "員工個人資料明細待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行員工個人資料明細確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
         'Modify By Sindy 2014/3/11
         If m_StrSQL <> "" Then
            PUB_SendMail strUserNum, strSendEmailTo, "", "員工個人資料明細待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行員工個人資料明細確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
         Else
         '2014/3/11 END
            PUB_SendMail strUserNum, "taie_alluser@taie.com.tw", "", "員工個人資料明細待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行員工個人資料明細確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
            '2013/3/5 End
         End If
      End If
   End If
   m_Device.EndDoc
   'ShowPrintOk
   If intECnt = 0 Then
      MsgBox "列印紙本 " & intPCnt & " 筆完成!!", , "列印成功"
   Else
      MsgBox "列印紙本 " & intPCnt & " 筆及產生電子檔 " & intECnt & " 筆完成!!", , "列印成功"
   End If
End If

Exit Sub

ErrHand:
   If bolEPrint = True Then '電子化線上確認
      cnnConnection.RollbackTrans
      MsgBox " 更新失敗！" & vbCrLf & Err.Description
   End If
'2011/9/2 End
End Sub

Sub PrintTitle(strST01 As String)
GetPleft

m_Device.Font.Size = 12 * douExtRate
m_Device.Font.Underline = False
m_Device.FontBold = False

'm_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth("個人人事資料明細表") / 2)
m_Device.CurrentX = 4500 * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
m_Device.Print "個人人事資料明細表"

m_Device.Font.Size = 10 * douExtRate 'Add By Sindy 2012/6/20
iLine = iLine + 1
'm_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
'Modify By Sindy 2012/6/20 改列印在左邊
m_Device.CurrentX = PLeft(1) * douExtRate '9000 * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
m_Device.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

'Add By Sindy 2014/3/12 +頁次
m_Device.CurrentX = (m_Device.ScaleWidth - 3000) * douExtRate '9000 * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
If m_ShowSpecialty = True Then
   m_Device.Print "頁　　次：1 / 2"
Else
   m_Device.Print "頁　　次：1 / 1"
End If
'2014/3/12 END

'Add By Sindy 2012/6/20
If ReadPhoto(strST01) = True Then
   'm_Device.PaintPicture tmpImg, 6350, 12600, intWidth - 100, intHeight - 100
   m_Device.PaintPicture tmpImg, (8500 * douExtRate), ((iLine + 2) * 400) * douExtRate, sW, sH
End If
'2012/6/20 End

iLine = iLine + 1
'm_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
'Modify By Sindy 2012/6/20 改列印在左邊
m_Device.CurrentX = PLeft(1) * douExtRate '9000 * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
'm_Device.Print "頁　　次：1" '& m_Device.Page

m_Device.Font.Size = 12 * douExtRate 'Add By Sindy 2012/6/20
iLine = iLine + 1
End Sub

Sub GetPleft()
'員工基本資料
PLeft(1) = 1000
'Modify By Sindy 2012/6/20 為了放照片,移動位置
PLeft(2) = 4500 '7000
'眷屬資料
PLeft(3) = 4000
PLeft(4) = 7000
'健保眷屬明細
PLeft(5) = 3000
PLeft(6) = 4500
PLeft(7) = 6500
PLeft(8) = 8500
End Sub

Sub PrintDetail()
Dim m_j As Integer, i As Integer, intCnt As Integer

'員工基本資料
For m_j = 1 To 19
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print strTemp(m_j)
   If m_j = 3 Or m_j = 5 Or m_j = 9 Or m_j = 14 Then
      m_j = m_j + 1
      m_Device.CurrentX = PLeft(2) * douExtRate
      m_Device.CurrentY = (iLine * 400) * douExtRate
      m_Device.Print strTemp(m_j)
   End If
   'Add By Sindy 98/04/13
   If m_j = 16 Then
      m_Device.CurrentX = PLeft(2) * douExtRate
      m_Device.CurrentY = (iLine * 400) * douExtRate
      m_Device.Print strTemp(56)
   End If
   '98/04/13 End
   iLine = iLine + 1
Next m_j
m_Device.CurrentX = 500 * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
m_Device.Print String(140, "-")
iLine = iLine + 1
'配偶
m_Device.CurrentX = PLeft(1) * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
m_Device.Print strTemp(20)
m_Device.CurrentX = PLeft(3) * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
m_Device.Print strTemp(21)
iLine = iLine + 2
'子女
intCnt = 21
For m_j = 1 To 4
   For i = 1 To 3
      If i = 1 Then m_Device.CurrentX = PLeft(1) * douExtRate
      If i = 2 Then m_Device.CurrentX = PLeft(3) * douExtRate
      If i = 3 Then m_Device.CurrentX = PLeft(4) * douExtRate
      m_Device.CurrentY = (iLine * 400) * douExtRate
      intCnt = intCnt + 1
      m_Device.Print strTemp(intCnt)
   Next i
   iLine = iLine + 1
Next m_j
'父母親
iLine = iLine + 1
For m_j = 1 To 6
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print strTemp(m_j + 33)
   If m_j = 1 Or m_j = 4 Then
      m_j = m_j + 1
      m_Device.CurrentX = PLeft(3) * douExtRate
      m_Device.CurrentY = (iLine * 400) * douExtRate
      m_Device.Print strTemp(m_j + 33)
   End If
   iLine = iLine + 1
Next m_j
m_Device.CurrentX = 500 * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
m_Device.Print String(140, "-")
iLine = iLine + 1
'健保眷屬
For m_j = 1 To 4
   m_Device.CurrentX = PLeft(m_j + 4) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   If m_j = 1 Or m_j = 3 Then m_Device.Print "姓　　名"
   If m_j = 2 Or m_j = 4 Then m_Device.Print "身 份 證 字 號"
Next m_j
iLine = iLine + 1
m_Device.CurrentX = PLeft(1) * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate
If m_QueryCnt = 0 Then
   m_Device.Print "健保眷屬明細：　(無)"
Else
   m_Device.Print "健保眷屬明細："
End If
intCnt = 39
For m_j = 1 To 4
   For i = 1 To 4
      If i = 1 Then m_Device.CurrentX = PLeft(5) * douExtRate
      If i = 2 Then m_Device.CurrentX = PLeft(6) * douExtRate
      If i = 3 Then m_Device.CurrentX = PLeft(7) * douExtRate
      If i = 4 Then m_Device.CurrentX = PLeft(8) * douExtRate
      m_Device.CurrentY = (iLine * 400) * douExtRate
      intCnt = intCnt + 1
      m_Device.Print strTemp(intCnt)
   Next i
   iLine = iLine + 1
Next m_j

iLine = iLine + 1
m_Device.CurrentX = PLeft(1) * douExtRate
m_Device.CurrentY = (iLine * 400) * douExtRate

'Add By Sindy 2011/9/2
If intChoose = 1 Then '出缺勤系統
   m_Device.Print "請於 " & ChangeTStringToTDateString(txt1(4)) & " 核對後執行確認；如資料有誤，請以E-Mail通知人事處。"
Else
   m_Device.Print "請於 " & ChangeTStringToTDateString(txt1(4)) & " 核對後簽章繳回人事處！"
End If
iLine = iLine + 1

'Add By Sindy 2014/3/12 +第二頁-專長資料
If m_ShowSpecialty = True Then
   If intChoose = 1 Then '出缺勤系統
      m_iPages = m_iPages + 1
      SetPic m_iPages
   Else
      m_Device.NewPage
   End If
   iLine = 1
   m_Device.Font.Size = 12 * douExtRate
   m_Device.Font.Underline = False
   m_Device.FontBold = False
   
   m_Device.CurrentX = 4500 * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "個人人事資料明細表"
   
   m_Device.Font.Size = 10 * douExtRate
   iLine = iLine + 1
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   m_Device.CurrentX = (m_Device.ScaleWidth - 3000) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "頁　　次：2 / 2"
   
   m_Device.Font.Size = 12 * douExtRate
   iLine = iLine + 2
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "教育與訓練："
   Call OffLineProcess(strTemp(57))
'   iLine = iLine + 1
'   m_Device.CurrentX = PLeft(1) * douExtRate
'   m_Device.CurrentY = (iLine * 400) * douExtRate
'   m_Device.Print IIf(strTemp(57) = "", "（無）", strTemp(57))
   iLine = iLine + 2
   If iLine > 38 Then
      If intChoose = 1 Then '出缺勤系統
         m_iPages = m_iPages + 1
         SetPic m_iPages
      Else
         m_Device.NewPage
      End If
      iLine = 3
   End If
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "語文能力檢定："
   Call OffLineProcess(strTemp(58))
'   iLine = iLine + 1
'   m_Device.CurrentX = PLeft(1) * douExtRate
'   m_Device.CurrentY = (iLine * 400) * douExtRate
'   m_Device.Print IIf(strTemp(58) = "", "（無）", strTemp(58))
   iLine = iLine + 2
   If iLine > 38 Then
      If intChoose = 1 Then '出缺勤系統
         m_iPages = m_iPages + 1
         SetPic m_iPages
      Else
         m_Device.NewPage
      End If
      iLine = 3
   End If
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "智財專業證照：" '專業資格／技能證照 Modify By Sindy 2024/4/17
   Call OffLineProcess(strTemp(59))
'   iLine = iLine + 1
'   m_Device.CurrentX = PLeft(1) * douExtRate
'   m_Device.CurrentY = (iLine * 400) * douExtRate
'   m_Device.Print IIf(strTemp(59) = "", "（無）", strTemp(59)) 'Replace(strTemp(59), vbCrLf, vbCrLf & "　　　　 "))
   iLine = iLine + 2
   If iLine > 38 Then
      If intChoose = 1 Then '出缺勤系統
         m_iPages = m_iPages + 1
         SetPic m_iPages
      Else
         m_Device.NewPage
      End If
      iLine = 3
   End If
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "著作／發明／創作："
   Call OffLineProcess(strTemp(60))
'   iLine = iLine + 1
'   m_Device.CurrentX = PLeft(1) * douExtRate
'   m_Device.CurrentY = (iLine * 400) * douExtRate
'   m_Device.Print IIf(strTemp(60) = "", "（無）", strTemp(60))
   iLine = iLine + 2
   If iLine > 38 Then
      If intChoose = 1 Then '出缺勤系統
         m_iPages = m_iPages + 1
         SetPic m_iPages
      Else
         m_Device.NewPage
      End If
      iLine = 3
   End If
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "非智財專業證照：" '電腦技能證照 Modify By Sindy 2024/4/17
   Call OffLineProcess(strTemp(61))
'   iLine = iLine + 1
'   m_Device.CurrentX = PLeft(1) * douExtRate
'   m_Device.CurrentY = (iLine * 400) * douExtRate
'   m_Device.Print IIf(strTemp(61) = "", "（無）", strTemp(61))
   iLine = iLine + 2
   If iLine > 38 Then
      If intChoose = 1 Then '出缺勤系統
         m_iPages = m_iPages + 1
         SetPic m_iPages
      Else
         m_Device.NewPage
      End If
      iLine = 3
   End If
   m_Device.CurrentX = PLeft(1) * douExtRate
   m_Device.CurrentY = (iLine * 400) * douExtRate
   m_Device.Print "專業／社會團體會員："
   Call OffLineProcess(strTemp(62))
'   iLine = iLine + 1
'   m_Device.CurrentX = PLeft(1) * douExtRate
'   m_Device.CurrentY = (iLine * 400) * douExtRate
'   m_Device.Print IIf(strTemp(62) = "", "（無）", strTemp(62))
End If
'2014/3/12 END
End Sub

'折行處理
Private Sub OffLineProcess(strPrText As String)
Dim PrintWidthWord As Integer
Dim ii As Integer
Dim PrintDetailTxt(40) As String
Dim w2 As Integer
Dim PrintDetailTemp As Variant
Dim oJ As Integer
Dim bolOK As Boolean
   
   If strPrText = "" Then
      iLine = iLine + 1
      m_Device.CurrentX = PLeft(1) * douExtRate
      m_Device.CurrentY = (iLine * 400) * douExtRate
      m_Device.Print "（無）"
   Else
      PrintWidthWord = 82
      '清除陣列值
      For ii = 0 To 40
         PrintDetailTxt(ii) = ""
      Next ii
      w2 = 0
      PrintDetailTemp = Split(strPrText, vbCrLf)
      For oJ = 0 To UBound(PrintDetailTemp)
         If PrintDetailTemp(oJ) <> "" Then
            w2 = w2 + 1
            PrintDetailTxt(w2) = PrintDetailTemp(oJ)
            If PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord + 2) <> PrintDetailTemp(oJ) Then
               PrintDetailTxt(w2) = "": w2 = w2 - 1
               bolOK = True
               Do While bolOK = True
                  w2 = w2 + 1
                  PrintDetailTxt(w2) = PrintDetailTxt(w2) & RTrim(PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord) & Chr(13)) & Chr(10)
                  PrintDetailTemp(oJ) = Replace(PrintDetailTemp(oJ), PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord), "")
                  If PUB_StrToStr_byVal(PrintDetailTemp(oJ), PrintWidthWord) = PrintDetailTemp(oJ) Then
                     w2 = w2 + 1
                     PrintDetailTxt(w2) = PrintDetailTxt(w2) & PrintDetailTemp(oJ)
                     bolOK = False
                  End If
              Loop
            Else
               PrintDetailTxt(w2) = RTrim(PrintDetailTemp(oJ))
            End If
         End If
      Next oJ
      For ii = 1 To w2
         iLine = iLine + 1
         m_Device.CurrentX = PLeft(1) * douExtRate
         m_Device.CurrentY = (iLine * 400) * douExtRate
         m_Device.Print PrintDetailTxt(ii)
         If iLine > 38 Then
            If intChoose = 1 Then '出缺勤系統
               m_iPages = m_iPages + 1
               SetPic m_iPages
            Else
               m_Device.NewPage
            End If
            iLine = 3
         End If
      Next ii
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
'   strSystemKind = GetSystemKindByNick
'   strSql = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
'   For i = 0 To Printers.Count - 1
'      Set Printer = Printers(i)
'      Combo1.AddItem Printer.DeviceName, j
'      j = j + 1
'      If Printer.DeviceName = strSql Then
'         SeekPrint = i
'      End If
'   Next i
'
'   Set Printer = Printers(SeekPrint)
'   Combo1.Text = Combo1.List(SeekPrint)

   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2022/1/23
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strText As String
   
   'Me.Form=D
   '一進入系統,檢查是否有須要開啟此作業
   'Modify By Sindy 2015/7/2
   If pub_CallNextABSForm = True Then
      strText = ChkIsAbsenceMustPro
      Me.Hide
      If InStr(1, strText, "G") > 0 Then
         If TypeName(Tmpfrm210148) <> "Nothing" Then
            Tmpfrm210148.Show
         End If
      ElseIf InStr(1, strText, "H") > 0 Then
         If TypeName(Tmpfrm210147) <> "Nothing" Then
            Tmpfrm210147.Show
         End If
      Else
         pub_CallNextABSForm = False
      End If
   End If
   
   Set frm160102 = Nothing
   If pub_CallNextABSForm = False Then
      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
   End If
End Sub

'Add By Sindy 2011/9/2
Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

'Add By Sindy 2011/9/2
Private Sub SetPic(idx As Integer)
Dim strPicFileName As String
   
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 4
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 4
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index)) = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

'Add By Sindy 2012/6/20 載入照片
Private Function ReadPhoto(strST01 As String) As Boolean
Dim PicRs As New ADODB.Recordset
Dim file_num As Integer
Dim bytes() As Byte
Dim IsWmf As Boolean
Dim pWidth As Integer '圖片寬度
Dim pHeight As Integer '圖片高度
Dim dblTmp As Double
   
   Screen.MousePointer = vbHourglass
   
   ReadPhoto = False
   
   '清圖片
   tmpPic.Picture = LoadPicture()
   tmpImg.Picture = LoadPicture()
   G_SeekPicColor.Picture = LoadPicture()
   G_SeekPicColor.Width = 0
   G_SeekPicColor.Height = 0
   
   DoEvents
   Set PicRs = New ADODB.Recordset
   PicRs.CursorLocation = adUseClient
   PicRs.Open "select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02 from ImgByteFile,staff S1,staff S2 where ibf05='3' and ibf01='000' and ibf02='" & strST01 & "' and ibf03='0' and ibf04='00' and ibf07=s1.st01(+) and ibf10=s2.st01(+) ", cnnConnection, adOpenStatic, adLockOptimistic
   If PicRs.RecordCount <> 0 Then
      ReadPhoto = True
      
      PicRs.MoveFirst
      If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Or CheckStr(PicRs.Fields("ibf06")) = "6" Then
         IsWmf = True
      Else
         IsWmf = False
      End If
      'Add By Sindy 2017/8/10
'      If "" & PicRs.Fields("IBF15") <> "" Then
         Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic." & IIf(IsWmf, "wmf", "jpg"), UCase("ImgByteFile"))
'      Else
'      '2017/8/10 END
'         ReDim bytes(Val(PicRs.Fields("ibf13").Value))
'         bytes() = PicRs.Fields("ibf14").GetChunk(Val(PicRs.Fields("ibf13").Value))
'         file_num = FreeFile
'         If IsWmf = False Then
'            Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
'         Else
'            Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
'         End If
'         Put #file_num, , bytes()
'         Close #file_num
'      End If
      
      G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.jpg")
      pWidth = G_SeekPicColor.Width
      pHeight = G_SeekPicColor.Height
      sH = 0: sW = 0
      If pWidth < pHeight Then '以高的比例
         dblTmp = pHeight / tmpPic.Height
         sH = tmpPic.Height
      Else '以寬的比例
         dblTmp = pWidth / tmpPic.Width
         sW = tmpPic.Width
      End If
      If sW = 0 Then
         sW = pWidth / dblTmp
         'Add By Sindy 2012/7/27
         If sW > tmpPic.Width Then
            '寬度等比例縮小後還是大於圖框寬,再以寬的比例縮放
            dblTmp = sW / tmpPic.Width
            sW = tmpPic.Width
            sH = sH / dblTmp
         End If
         '2012/7/27 End
      ElseIf sH = 0 Then
         sH = pHeight / dblTmp
         'Add By Sindy 2012/7/27
         If sH > tmpPic.Height Then
            '高度等比例縮小後還是大於圖框高,再以高的比例縮放
            dblTmp = sH / tmpPic.Height
            sH = tmpPic.Height
            sW = sW / dblTmp
         End If
         '2012/7/27 End
      End If
      tmpImg.Width = sW: tmpImg.Height = sH
      Set tmpImg.Picture = G_SeekPicColor
      'tmpPic.PaintPicture G_SeekPicColor, ((tmpPic.Width - sW) / 2) / 2, ((tmpPic.Height - sH) / 2), sW, sH
      tmpPic.PaintPicture G_SeekPicColor, IIf(tmpPic.ScaleWidth / 2 - (sW / 2) < 0, 0, tmpPic.ScaleWidth / 2 - (sW / 2)), IIf(tmpPic.ScaleHeight / 2 - (sH / 2) < 0, 0, tmpPic.ScaleHeight / 2 - (sH / 2)), sW, sH
      Set tmpPic.Picture = tmpPic.Image
      
      If Dir(App.path & "\NowPic.jpg") <> "" Then
         Kill App.path & "\NowPic.jpg"
      End If
      If Dir(App.path & "\NowPic.wmf") <> "" Then
         Kill App.path & "\NowPic.wmf"
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Function
