VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmacc4350 
   AutoRedraw      =   -1  'True
   Caption         =   "應收/付轉傳票作業"
   ClientHeight    =   5304
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5172
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5304
   ScaleWidth      =   5172
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   4815
      _ExtentX        =   8488
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      ItemData        =   "Frmacc4350.frx":0000
      Left            =   120
      List            =   "Frmacc4350.frx":0002
      TabIndex        =   6
      Top             =   4416
      Width           =   4815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1416
      ItemData        =   "Frmacc4350.frx":0004
      Left            =   120
      List            =   "Frmacc4350.frx":0006
      TabIndex        =   5
      Top             =   2460
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行(&E)"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   276
      Width           =   1584
      _ExtentX        =   2794
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1590
      Width           =   4815
      _ExtentX        =   8488
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "注意：請通知其他人員暫停相關資料的維護動作！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   4620
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "應收/付日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc4350"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoprimary As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim stChkDate As String, i As Integer 'Add by Amy 2025/05/05

Private Sub Command1_Click()
   Dim stMsg As String, arrTmp 'Add by Amy 2025/05/05
   
   'Add by Amy 2025/05/05 避免有資料未傳送,且畫面 應收/付日期 又未包含下列日期,造成傳票號與傳票日有問題-秀玲
   If stChkDate <> "" Then
      arrTmp = Split(stChkDate, ";")
      For i = LBound(arrTmp) To UBound(arrTmp)
         If FCDate(MaskEdBox1) >= arrTmp(i) And FCDate(MaskEdBox1) <= arrTmp(i) Then
         Else
            stMsg = "尚有之前日期資料未傳送"
            Exit For
         End If
      Next i
      If stMsg <> "" Then
         MsgBox stMsg & vbCrLf & "請確認！"
         Exit Sub
      End If
   End If
   'end 2025/05/05

   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0b10 from acc0b0 where a0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      MsgBox MsgText(197), , MsgText(5)
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   adoTaie.Execute "update acc0b0 set a0b10 = '01'"
   TransferTable
   adoTaie.Execute "update acc0b0 set a0b10 = null"
End Sub

Private Sub Form_Load()
Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
Dim stDate As String 'Add by Amy 2024/08/01
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5300
   Me.Height = 5715
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'modify by sonia 2013/6/25 瑞婷說改預設系統日之前一個工作天
   'MaskEdBox1.Text = Mid(ACDate(ServerDate), 1, 3) & "/" & Mid(ACDate(ServerDate), 4, 2) & "/" & Mid(ACDate(ServerDate), 6, 2)
   'MaskEdBox1.Mask = DFormat
   'MaskEdBox2.Text = Mid(ACDate(ServerDate), 1, 3) & "/" & Mid(ACDate(ServerDate), 4, 2) & "/" & Mid(ACDate(ServerDate), 6, 2)
   'MaskEdBox2.Mask = DFormat
   'Modify by Amy 2024/08/01 北所休颱風假時,預設為颱風假前一個工作天
   '        1130724-25 全台休颱風假,7/26上班時,日期預設 7/25 佩瑄 未注意直接按執行鈕,因還有7/23之應收付(分攤)資料要導致不連號
   stDate = GetTyphoon(CompWorkDay(2, strSrvDate(1), 1))
   MaskEdBox1 = CFDate(ChangeWStringToTString(stDate))
   'end 2024/08/01
   MaskEdBox2 = MaskEdBox1
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   '2013/6/25 END
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select distinct a1p01,a1p22 from acc1p0 where a1p27 = '" & MsgText(602) & "' order by a1p22 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      If IsNull(adoquery.Fields("a1p22").Value) = False Then
         List2.AddItem adoquery.Fields("a1p01").Value & "公司" & adoquery.Fields("a1p22").Value & " --> 已異動，請重新傳送"
      End If
      adoquery.MoveNext
   Loop
   adoquery.Close
   'Add by Amy 2022/09/21 +執行日之前尚有未傳送之資料
   strExc(1) = "Select Distinct a1p18,a1p04 From acc1p0 Where a1p18 < " & Val(ChangeTDateStringToTString(MaskEdBox1)) & " And a1p22 IS NULL Order by a1p18,a1p04 asc"
   If adoquery.State = adStateOpen Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      If IsNull(adoquery.Fields("a1p18").Value) = False Then
         List2.AddItem adoquery.Fields("a1p18").Value & " " & adoquery.Fields("a1p04").Value & " -->資料未傳送"
         'Add by Amy 2025/05/05 避免有資料未傳送,且畫面 應收/付日期 又未包含下列日期,造成傳票號與傳票日有問題-秀玲
         If InStr(stChkDate, adoquery.Fields("a1p18").Value) = 0 Then
            stChkDate = stChkDate & ";" & adoquery.Fields("a1p18").Value
         End If
         'end 2025/05/05
      End If
      adoquery.MoveNext
   Loop
   If stChkDate <> "" Then stChkDate = Mid(stChkDate, 2) 'Add by Amy 2025/05/05
   adoquery.Close
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Call PUB_GetLock("", "Frmacc4350") 'Add by Amy 2020/06/07
   Set Frmacc4350 = Nothing
End Sub

'Added by Morgan 2025/10/22 資料太長無法看到完整內容時用雙擊方式彈跳顯示
Private Sub List2_DblClick()
   If List2.Text <> "" Then
      MsgBox List2.Text, vbInformation
   End If
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
' 轉傳票處理
'
'*************************************************
Public Sub TransferTable()

Dim strAutoNo As String
Dim strSave As String
Dim strSerialNo As String
Dim strAccNo As String
Dim lngDebit, lngCredit As Long
Dim strDueDate As String, strNo As String
Dim strName As String
Dim intCounter As Integer
Dim douDebit As Double
Dim douCredit As Double
Dim m_A1P02 As String
Dim strOriAccNo As String
Dim strCaseNo As String
Dim w_douCredit As Double
'Add by Morgan 2007/8/10
Dim stA0206 As String, stA0207 As String, stA0208 As String, stAX210 As String
'Add by Morgan 2010/9/13
Dim lngThisDate As Long, lngEndDate As Long, iNum As Integer, stDelta As String
Dim stA1P04 As String, stA1P14 As String, stA1P22 As String, stA1P27 As String
Dim stA1P05_001 As String, stA1P05_002 As String
Dim tot_douDebit As Double  '2012/6/15 add by sonia
Dim m_balance As Boolean   '2012/6/15 add by sonia
'Add by Amy 2013/12/19 intSeq=回圈執行次數 / strCompany=公司別 / IdentifierStr=傳入AccAutoNo文字
Dim intSeq As Integer, ii As Integer, strCompany As String, IdentifierStr As String
Dim intErr As Integer 'Add by Amy 2014/01/20 計算錯誤
Dim strA1P18 As String 'Add by Amy 2017/05/02
'Add by Amy 2020/03/18
Dim arrTmp
Dim iCount As Integer
Dim strCmd As String 'Add by Amy 2023/03/07
   
On Error GoTo Checking
   adoTaie.BeginTrans
   List1.Clear
   intErr = 0 'Add by Amy 2014/01/20
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2020/03/18 抓作帳公司
   'Modify by Amy 2013/12/19
'   If strSrvDate(1) >= InvoiceStartDate Then
'        ProgressBar2.max = 12
'   Else
'        ProgressBar2.max = 6
'   End If
   'end 2013/12/19
   arrTmp = Split(GetBookKeepCmp, ",")
   ProgressBar2.max = 6 * (UBound(arrTmp) + 1)
   'end 2020/03/18
   ProgressBar2.Value = 0
   
'Modify by Amy 2020/03/18 抓作帳公司
'Add by Amy 2013/12/19
'If strSrvDate(1) >= InvoiceStartDate Then
'    intSeq = 2
'Else
'    intSeq = 1
'End If
intSeq = UBound(arrTmp)
'以作帳公司順序(a0801)跑
For ii = 0 To intSeq
   'Add by Morgan 2010/9/13
   '新增外幣匯差調整傳票
   StatusView "新增外幣匯差調整傳票"
   lngEndDate = Val(ChangeTDateStringToTString(MaskEdBox2))
   lngThisDate = Val(ChangeTDateStringToTString(MaskEdBox1))
   'Add 2013/12/19 +公司別 改strExc(0)
   strCompany = arrTmp(ii)
   If arrTmp(ii) = "1" Then
      IdentifierStr = MsgText(801)
   ElseIf arrTmp(ii) = "L" Then
      IdentifierStr = MsgText(820) 'LD
   Else
      IdentifierStr = MsgText(819) 'JD
   End If
   iCount = 0
'end 2020/03/18
   List1.AddItem strCompany & "-公司轉檔開始："
   'Add by Morgan 2010/11/4 先檢查是否有未過帳且需作外幣匯差調整的資料
'   strExc(0) = "select 1 from acc1p0,acc1x0 where a1p18>=" & lngThisDate & " and a1p18<=" & lngEndDate & _
'         " and a1p19||'' in ('USD','EUR','RMB')" & _
'         " and a1p05||'' in ('110205','110222','1917','113002')" & _
'         " and a1p07>0 and a1x01(+)=a1p19 and a1x02<>a1p20" & _
'         " and not exists(select * from acc021 where ax202=a1p22 and ax210>0) and rownum<2"
   'Modified by Morgan 2015/4/30 幣別改抓銀存匯率輸入的(原只抓 'USD','EUR','RMB')
   'Modified by Morgan 2021/1/20 +1918 --瑞婷
   'Modified by Morgan 2022/5/24 +110208 --婉莘
   strExc(0) = "select 1 from acc1p0,acc1x0 where a1p18>=" & lngThisDate & " and a1p18<=" & lngEndDate & _
         " and a1p19||'' in (select a1x01 from acc1x0)" & _
         " and a1p05||'' in ('110205','110222','1917','1918','113002','110208')" & _
         " and a1p07>0 and a1x01(+)=a1p19 and a1x02<>a1p20 And a1p01||'' ='" & strCompany & "' " & _
         " and not exists(select * from acc021 where ax202=a1p22 and ax210>0 And ax201='" & strCompany & "') and rownum<2"
   'end 201312/19
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   'end 2010/11/4
      'Add by Morgan 2010/10/6 檢查傳票是否已過帳
      'Modify by Amy 2013/12/19 +公司別
'      strExc(0) = "select a1p22 from acc1p0 where a1p04>='外幣匯差調整" & lngThisDate & "'" & _
'         " and a1p04<='外幣匯差調整" & lngEndDate & "99' and a1p22 is not null and rownum<2" & _
'         " and exists(select * from acc021 where ax202=a1p22 and ax210>0)"
      strExc(0) = "select a1p22 from acc1p0 where a1p04>='外幣匯差調整" & lngThisDate & "'" & _
         " and a1p04<='外幣匯差調整" & lngEndDate & "99' and a1p22 is not null And a1p01||''='" & strCompany & "' and rownum<2" & _
         " and exists(select * from acc021 where ax202=a1p22 and ax210>0 And ax201='" & strCompany & "' )"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modify by Morgan 2010/11/4
         'MsgBox "外幣匯差調整傳票已過帳，請取消過帳後再作業！"
         'Screen.MousePointer = vbDefault
         'GoTo Checking
         MsgBox "外幣匯差調整傳票已過帳，請自行新增尚未過帳之外幣匯差調整資料！"
      'end 2010/10/6
      Else
      
         Do While lngThisDate <= lngEndDate
            'Moify by Morgan 2010/9/16 考慮科目可能會更改,改為一天只做一張外幣匯差調整傳票
            'strExc(0) = "select * from ( select a1p18,a1p05,a1p19,a1p20,sum(a1p21) tot" & _
               " from acc1p0 where a1p18=" & lngThisDate & " and a1p22 is null" & _
               " and a1p19||'' in ('USD','EUR','RMB')" & _
               " and a1p05||'' in ('110205','110222','1917','113002')" & _
               " and a1p07>0 group by a1p18,a1p05,a1p19,a1p20" & _
               ") x,acc1x0,(select a1p18 y1,max(substr(a1p04,-2)) y2 from acc1p0" & _
               " where a1p18=" & lngThisDate & " and instr(a1p04,'外幣匯差調整')>0 group by a1p18) y" & _
               " where a1x01(+)=a1p19 and a1x02<>a1p20 and y1(+)=a1p18"
            stA1P04 = "外幣匯差調整" & lngThisDate & "01"
            stA1P22 = ""
            stA1P27 = ""
            'Modify by Amy 2013/12/19 +公司別
            'strExc(0) = "select a1p22 from acc1p0 where a1p04='" & stA1P04 & "' order by a1p22"
            strExc(0) = "select a1p22 from acc1p0 where a1p01='" & strCompany & "' And a1p04='" & stA1P04 & "' order by a1p22"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               '有傳票號碼
               If Not IsNull(RsTemp(0)) Then
                  stA1P22 = RsTemp(0)
                  stA1P27 = "Y"
               End If
               'Modify by Amy 2013/12/19 +公司別
               'strSql = "delete from acc1p0 where a1p04='" & stA1P04 & "'"
               strSql = "delete from acc1p0 where a1p01='" & strCompany & "' And a1p04='" & stA1P04 & "'"
               adoTaie.Execute strSql, intI
            End If
            
            'Modify by Amy 2013/12/19 +公司別
            'strExc(0) = "select a1p18,a1p05,a1p19,a1p20,a1x02,sum(a1p21) tot" & _
               ",sum(a1p07-round(a1p21*a1x02,2)) delta" & _
               " from acc1p0,acc1x0 where a1p18=" & lngThisDate & _
               " and a1p19||'' in ('USD','EUR','RMB')" & _
               " and a1p05||'' in ('110205','110222','1917','113002')" & _
               " and a1p07>0 and a1x01(+)=a1p19 and a1x02<>a1p20" & _
               " and not exists(select * from acc021 where ax202=a1p22 and ax210>0)" & _
               " group by a1p18,a1p05,a1p19,a1p20,a1x02"
            'Modified by Morgan 2015/4/30 幣別改抓銀存匯率輸入的(原只抓 'USD','EUR','RMB')
            'modify by sonia 2018/9/11 加order by
            'Modified by Morgan 2021/1/20 +1918 --瑞婷
            'Modified by Morgan 2022/5/24 +110208 --婉莘
            strExc(0) = "select a1p18,a1p05,a1p19,a1p20,a1x02,sum(a1p21) tot" & _
               ",sum(a1p07-round(a1p21*a1x02,2)) delta" & _
               " from acc1p0,acc1x0 where a1p18=" & lngThisDate & _
               " and a1p19||'' in (select a1x01 from acc1x0)" & _
               " and a1p05||'' in ('110205','110222','1917','1918','113002','110208')" & _
               " and a1p07>0 and a1x01(+)=a1p19 and a1x02<>a1p20 And a1p01='" & strCompany & "' " & _
               " and not exists(select * from acc021 where ax202=a1p22 and ax210>0 And ax201='" & strCompany & "' )" & _
               " group by a1p18,a1p05,a1p19,a1p20,a1x02 order by a1p18,a1p05,a1p19,a1p20,a1x02"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
               'stA1P04 = "外幣匯差調整" & lngThisDate & Format(Val("" & .Fields("y2")) + 1, "00")
               iNum = 0
               Do While Not .EOF
                  stA1P14 = .Fields("a1p19") & .Fields("tot") & " X " & .Fields("a1p20") & " 匯差" & (.Fields("a1p20") - .Fields("a1x02"))
                  stDelta = Abs(.Fields("delta"))
                  If .Fields("a1p20") > .Fields("a1x02") Then
                     stA1P05_001 = "7128"
                     stA1P05_002 = .Fields("a1p05")
                  Else
                     stA1P05_001 = .Fields("a1p05")
                     stA1P05_002 = "7128"
                  End If
                  
                  iNum = iNum + 1
                  'Modify by Amy 2013/12/19 原:a1p01值'1'
                  strSql = "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p18,a1p22,a1p27,a1p28,a1p29)" & _
                     " values ('" & strCompany & "','F','" & Format(iNum, "000") & "','" & stA1P04 & "','" & stA1P05_001 & "','TOT'," & stDelta & ",0,'" & stA1P14 & "'," & .Fields("a1p18") & ",'" & stA1P22 & "','" & stA1P27 & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'))"
                  adoTaie.Execute strSql, intI
                  
                  iNum = iNum + 1
                  'Modify by Amy 2013/12/19 原:a1p01值'1'
                  strSql = "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p18,a1p22,a1p27,a1p28,a1p29)" & _
                     " values ('" & strCompany & "','F','" & Format(iNum, "000") & "','" & stA1P04 & "','" & stA1P05_002 & "','TOT',0," & stDelta & ",'" & stA1P14 & "'," & .Fields("a1p18") & ",'" & stA1P22 & "','" & stA1P27 & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'))"
                  adoTaie.Execute strSql, intI
                  iCount = iCount + 1 'Add by Amy 2020/03/18
                  .MoveNext
               Loop
               End With
            End If
            lngThisDate = Val(TransDate(CompDate(2, 1, lngThisDate), 1))
         Loop
         'end 2010/9/13
      End If
   End If
   'Add by Amy 2020/03/18
   If iCount > 0 Then
        List1.AddItem "外幣匯差調整完成 " & iCount & " 筆"
   End If
   
   
' 每月固定傳票轉檔
   StatusView MsgText(74)
   strName = ""
   adoprimary.CursorLocation = adUseClient
   'Modify by Amy 2013/12/19 +公司別
   'adoprimary.Open "select * from acc0d0, acc0d1 where a0d01 = axd01 and a0d02 = axd02 and (decode(length(axd04), 6, substr(axd04, 1, 4), 7, substr(axd04, 1, 5)) < " & Val(Mid(MaskEdBox1.Text, 1, 3)) & Mid(MaskEdBox1.Text, 5, 2) & " or axd04 is null) and axd11 <= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & " and axd12 >= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & " and axd03 <= " & Val(Mid(MaskEdBox1.Text, 8, 2)) & " order by a0d01 asc, a0d02 asc, a0d03 asc", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2023/3/1  因為2/25~2/28放假故axd03 <= " & Val(Mid(MaskEdBox1.Text, 8, 2))改為axd03 <= " & Val(Mid(MaskEdBox2.Text, 8, 2))
   adoprimary.Open "select * from acc0d0, acc0d1 where a0d01='" & strCompany & "'  And a0d01 = axd01 and a0d02 = axd02 and (decode(length(axd04), 6, substr(axd04, 1, 4), 7, substr(axd04, 1, 5)) < " & Val(Mid(MaskEdBox1.Text, 1, 3)) & Mid(MaskEdBox1.Text, 5, 2) & " or axd04 is null) and axd11 <= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & " and axd12 >= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & " and axd03 <= " & Val(Mid(MaskEdBox2.Text, 8, 2)) & " order by a0d01 asc, a0d02 asc, a0d03 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoprimary.RecordCount <> 0 Then
      ProgressBar1.max = adoprimary.RecordCount
   End If
   Do While adoprimary.EOF = False
      '最後一個月
      If Val(Mid(FCDate(MaskEdBox1.Text), 1, Len(FCDate(MaskEdBox1.Text)) - 2)) = adoprimary.Fields("axd12").Value Then
         '2012/6/15 add by sonia 若餘額=借方總額則改抓固定傳票明細金額 213傳票
         'If Val(adoprimary.Fields("a0d06").Value) <> 0 Then
         m_balance = False
         'modify by sonia 2018/9/11 加order by
         strExc(0) = "select sum(a0d06) from acc0d0 where a0d01 = '" & adoprimary.Fields("axd01").Value & "' and a0d02 = " & adoprimary.Fields("axd02").Value & " group by a0d02 order by a0d02"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val(RsTemp.Fields(0)) = adoprimary.Fields("axd14").Value Then m_balance = True
         End If
         If m_balance Then
            If Val(adoprimary.Fields("a0d06").Value) <> 0 Then
               douDebit = Val(adoprimary.Fields("a0d06").Value)
               douCredit = 0
            Else
               douCredit = Val(adoprimary.Fields("a0d07").Value)
               douDebit = 0
            End If
         ElseIf Val(adoprimary.Fields("a0d06").Value) <> 0 Then
         '2012/6/15 end
            douDebit = Val(adoprimary.Fields("axd14").Value)
            douCredit = 0
         Else
            If adoprimary.Fields("a0d05").Value = "2401" Then
               douCredit = Val(adoprimary.Fields("a0d07").Value)
               douDebit = 0
               w_douCredit = Val(adoprimary.Fields("a0d07").Value)
            Else
               douCredit = Val(adoprimary.Fields("axd14").Value) - w_douCredit
               douDebit = 0
               w_douCredit = 0
            End If
         End If
      Else
         If Val(adoprimary.Fields("a0d06").Value) <> 0 Then
            douDebit = Val(adoprimary.Fields("a0d06").Value)
            douCredit = 0
         Else
            douCredit = Val(adoprimary.Fields("a0d07").Value)
            douDebit = 0
         End If
      End If
      If strName <> (adoprimary.Fields("a0d01").Value & adoprimary.Fields("a0d02").Value) Then
         '2012/6/15 modify by sonia 改為最後一個月更新為0,非最後一個月則以餘額扣除借方總額
         'If douDebit <> 0 Then
         '   adoTaie.Execute "update acc0d1 set axd14 = axd14 - " & douDebit & " where axd01 = '" & adoprimary.Fields("axd01").Value & "' and axd02 = " & adoprimary.Fields("axd02").Value & ""
         'Else
         '   adoTaie.Execute "update acc0d1 set axd14 = axd14 - " & douCredit & " where axd01 = '" & adoprimary.Fields("axd01").Value & "' and axd02 = " & adoprimary.Fields("axd02").Value & ""
         'End If
         '2012/6/15 end
         If Val(Mid(FCDate(MaskEdBox1.Text), 1, Len(FCDate(MaskEdBox1.Text)) - 2)) = adoprimary.Fields("axd12").Value Then
            adoTaie.Execute "update acc0d1 set axd14 = 0 where axd01 = '" & adoprimary.Fields("axd01").Value & "' and axd02 = " & adoprimary.Fields("axd02").Value & ""
         Else
            tot_douDebit = 0
            strExc(0) = "select sum(a0d06) from acc0d0 where a0d01 = '" & adoprimary.Fields("axd01").Value & "' and a0d02 = " & adoprimary.Fields("axd02").Value & ""
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               tot_douDebit = Val(RsTemp.Fields(0))
            End If
            adoTaie.Execute "update acc0d1 set axd14 = axd14 - " & tot_douDebit & " where axd01 = '" & adoprimary.Fields("axd01").Value & "' and axd02 = " & adoprimary.Fields("axd02").Value & ""
         End If
         '2012/6/15 end
         strName = (adoprimary.Fields("a0d01").Value & adoprimary.Fields("a0d02").Value)
      End If
      'Modify by Amy 2017/05/02 若當月最後上班日為24日以前,請自動產生固定傳票日在當月最後上班日-瑞婷
'      'Modify by Amy 2014/05/13 +A0d11(對沖代號-其它)寫入a1p30
'      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p18,a1p30) values ('" & adoprimary.Fields("a0d01").Value & "', 'U', '" & adoprimary.Fields("a0d03").Value & "', '" & adoprimary.Fields("a0d02").Value & Val(FCDate(MaskEdBox1.Text)) & "', '" & adoprimary.Fields("a0d05").Value & "', '" & adoprimary.Fields("a0d08").Value & "', " & douDebit & ", " & douCredit & ", '" & adoprimary.Fields("a0d10").Value & "', " & _
'                      "" & Val(FCDate(MaskEdBox1.Text)) & ", '" & adoprimary.Fields("a0d11").Value & "')"
'                      '"" & Val(Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & IIf(Len(adoprimary.Fields("axd03").Value) = 1, "0" & adoprimary.Fields("axd03").Value, adoprimary.Fields("axd03").Value)) & ")"
'      adoTaie.Execute "update acc0d1 set axd04 = " & Val(FCDate(MaskEdBox1.Text)) & " where axd01 = '" & adoprimary.Fields("axd01").Value & "' and axd02 = '" & adoprimary.Fields("axd02").Value & "'"
       If ChkWorkDay(Val(FCDate(MaskEdBox1.Text)) + 19110000) = False Then
        '若畫面日期>當月最後一個工作日, 則產生的傳票改做當月最後一個工作日
        If Val(FCDate(MaskEdBox1.Text)) + 19110000 > Val(GetMonthStdDay(Val(Left(FCDate(MaskEdBox1.Text), 5)) + 191100, 1)) Then
            strA1P18 = Val(GetMonthStdDay(Val(Left(FCDate(MaskEdBox1.Text), 5)) + 191100, 1)) - 19110000
        Else
            '每月處理日非工作日傳票改做畫面日期下一個工作日
            strA1P18 = PUB_GetWorkDayAfterSysDate(Val(FCDate(MaskEdBox1.Text)) + 19110000, 1)
        End If
      Else
        strA1P18 = Val(FCDate(MaskEdBox1.Text))
      End If
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p18,a1p30) values ('" & adoprimary.Fields("a0d01").Value & "', 'U', '" & adoprimary.Fields("a0d03").Value & "', '" & adoprimary.Fields("a0d02").Value & Val(FCDate(MaskEdBox1.Text)) & "', '" & adoprimary.Fields("a0d05").Value & "', '" & adoprimary.Fields("a0d08").Value & "', " & douDebit & ", " & douCredit & ", '" & adoprimary.Fields("a0d10").Value & "', " & _
                      "" & Val(strA1P18) & ", '" & adoprimary.Fields("a0d11").Value & "')"
      adoTaie.Execute "update acc0d1 set axd04 = " & Val(strA1P18) & " where axd01 = '" & adoprimary.Fields("axd01").Value & "' and axd02 = '" & adoprimary.Fields("axd02").Value & "'"
      'end 2017/05/02
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoprimary.MoveNext
   Loop
   List1.AddItem "每月固定傳票轉檔完成 " & adoprimary.RecordCount & " 筆"
    adoprimary.Close
   ProgressBar1.Value = 0
   ProgressBar2.Value = ProgressBar2.Value + 1

' 應收/付傳票更新
   StatusView MsgText(129)
   strName = ""
   'Add by Morgan 2007/11/27 若有新增的分錄也要更新a1p22,a1p27
   'Modify by Amy 2013/12/19 +公司別
    'strSql = "update acc1p0 a set a.a1p27='Y'" & _
      ",a.a1p22=(select nvl(a.a1p22,b.a1p22) from acc1p0 b where b.a1p04=a.a1p04 and b.a1p27 = 'Y' and b.a1p22 is not null and rownum<2)" & _
      " where a.a1p04 in ( select b.a1p04 from acc1p0 b where b.a1p27 = 'Y' and b.a1p22 is NOT null)" & _
      " and (a.a1p22 is null or a.a1p27 is null)"
   strSql = "update acc1p0 a set a.a1p27='Y'" & _
      ",a.a1p22=(select nvl(a.a1p22,b.a1p22) from acc1p0 b where b.a1p01='" & strCompany & "' And b.a1p04=a.a1p04 and b.a1p27 = 'Y' and b.a1p22 is not null and rownum<2)" & _
      " where a.a1p04 in ( select b.a1p04 from acc1p0 b where b.a1p01='" & strCompany & "' And b.a1p27 = 'Y' and b.a1p22 is NOT null)" & _
      " and (a.a1p22 is null or a.a1p27 is null) And a.a1p01='" & strCompany & "' "
   adoTaie.Execute strSql, intI
   'end 2007/11/27
   
   adoprimary.CursorLocation = adUseClient
   '2005/11/4 MODFIY BY SONIA
   'adoprimary.Open "select a1p01, a1p02, a1p18, a1p22 from acc1p0 where a1p27 = 'Y' group by a1p01, a1p02, a1p18, a1p22", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2013/12/19 +公司別
   'adoprimary.Open "select a1p01, a1p02, a1p18, a1p22 from acc1p0 where a1p27 = 'Y' and a1p22 is NOT null group by a1p01, a1p02, a1p18, a1p22", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2018/9/11 加order by
   adoprimary.Open "select a1p01, a1p02, a1p18, a1p22 from acc1p0 where a1p01='" & strCompany & "' And a1p27 = 'Y' and a1p22 is NOT null group by a1p01, a1p02, a1p18, a1p22 order by a1p01, a1p02, a1p18, a1p22", adoTaie, adOpenStatic, adLockReadOnly
   If adoprimary.RecordCount <> 0 Then
      ProgressBar1.max = adoprimary.RecordCount
   End If
   Do While adoprimary.EOF = False
      stA0206 = "null": stA0207 = "null": stA0208 = "null": stAX210 = "null"
      'Add by Morgan 2007/8/10 紀錄舊資料
      strExc(0) = "select * from acc020,acc021 where a0201 = '" & adoprimary.Fields("a1p01").Value & "' and a0202 = '" & adoprimary.Fields("a1p22").Value & "' and ax201(+)=a0201 and ax202(+)=a0202 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         If Not IsNull(.Fields("a0206")) Then stA0206 = .Fields("a0206")
         If Not IsNull(.Fields("a0207")) Then stA0207 = .Fields("a0207")
         If Not IsNull(.Fields("a0208")) Then stA0208 = "'" & .Fields("a0208") & "'"
         If Not IsNull(.Fields("ax210")) Then stAX210 = .Fields("ax210")
         End With
      End If
      'End 2007/8/10
      
      adoTaie.Execute "delete from acc020 where a0201 = '" & adoprimary.Fields("a1p01").Value & "' and a0202 = '" & adoprimary.Fields("a1p22").Value & "'"
      adoTaie.Execute "delete from acc021 where ax201 = '" & adoprimary.Fields("a1p01").Value & "' and ax202 = '" & adoprimary.Fields("a1p22").Value & "'"
      If strName <> (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p02").Value & adoprimary.Fields("a1p22").Value) Then
         intCounter = 0
         
         '檢查分錄是否平衡
         If adoacc1p0.State = adStateOpen Then
            adoacc1p0.Close
         End If
         adoacc1p0.CursorLocation = adUseClient
         adoacc1p0.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & adoprimary.Fields("a1p02").Value & "' and a1p22 = '" & adoprimary.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount <> 0 Then
            If IsNull(adoacc1p0.Fields(0).Value) Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p22").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextRecord
            End If
            If IsNull(adoacc1p0.Fields(1).Value) Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p22").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextRecord
            End If
            If adoacc1p0.Fields(0).Value <> adoacc1p0.Fields(1).Value Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p22").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextRecord
            End If
         Else
            List1.AddItem MsgText(209) & adoprimary.Fields("a1p22").Value
            intErr = intErr + 1 'Add by Amy 2014/01/20
            GoTo NextRecord
         End If
         adoacc1p0.Close
         
         strAutoNo = adoprimary.Fields("a1p22").Value
         adoacc1p0.CursorLocation = adUseClient
         adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & adoprimary.Fields("a1p02").Value & "' and a1p22 = '" & adoprimary.Fields("a1p22").Value & "' order by decode(a1p02, 'W', a1p08, 'L', a1p08, a1p03) asc", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount > 1 Then
            'Modify by Morgan 2007/8/10
            'adoTaie.Execute "insert into acc020 values ('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & ", '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '', 0, 0)"
            'Modify by Amy 2023/03/07 過帳後不更新 a0209/a0210/a0211,因過帳後有修改實績傳票,再跑 智權點數實績與結餘分析表 可能會一直彈訊息
            If stAX210 <> "null" Then
                strCmd = "Insert into acc020(a0201,a0202,a0205,a0206,a0207,a0208)" & _
                                " Values ('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & "," & stA0206 & "," & stA0207 & "," & stA0208 & ")"
            Else
                strCmd = "Insert into acc020(a0201,a0202,a0205,a0206,a0207,a0208,a0209,a0210,a0211)" & _
                                " Values ('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & "," & stA0206 & "," & stA0207 & "," & stA0208 & "," & strSrvDate(2) & ",to_char(sysdate,'HH24MISS'),'" & strUserNum & "')"
            End If
            adoTaie.Execute strCmd
            'end 2007/8/10
            Do While adoacc1p0.EOF = False
               If IsNull(adoacc1p0.Fields("a1p14").Value) Then
                  strRemark = ""
               Else
                  strRemark = Replace(adoacc1p0.Fields("a1p14").Value, "'", "''")
               End If
               If adoacc1p0.Fields("a1p02").Value = "W" Or adoacc1p0.Fields("a1p02").Value = "L" Then
                  strSerialNo = ZeroBeforeNo(Trim(str(intCounter)), 3)
               Else
                  strSerialNo = adoacc1p0.Fields("a1p03").Value
               End If
               If adoacc1p0.Fields("a1p05").Value = "2401" And adoacc1p0.Fields("a1p07").Value <> 0 And IsNull(adoacc1p0.Fields("a1p23").Value) = False Then
                  If adoquery.State = adStateOpen Then
                     adoquery.Close
                  End If
                  adoquery.CursorLocation = adUseClient
                  'Modify by Amy 2013/12/19 +公司別
                  'adoQuery.Open "select * from acc1p0 where a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p23").Value & "' or a1p23 = '" & adoacc1p0.Fields("a1p23").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                  adoquery.Open "select * from acc1p0 where a1p01='" & adoacc1p0.Fields("a1p01") & "' And a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p23").Value & "' or a1p23 = '" & adoacc1p0.Fields("a1p23").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                  If adoquery.RecordCount <> 0 Then
                     If IsNull(adoquery.Fields("a1p22").Value) Then
                        strOriAccNo = "null"
                     Else
                        strOriAccNo = "'" & adoquery.Fields("a1p22").Value & "'"
                     End If
                  Else
                     strOriAccNo = "null"
                  End If
                  adoquery.Close
               Else
                  If adoacc1p0.Fields("a1p05").Value = "2401" And adoacc1p0.Fields("a1p07").Value <> 0 And IsNull(adoacc1p0.Fields("a1p30").Value) = False Then
                     If adoquery.State = adStateOpen Then
                        adoquery.Close
                     End If
                     adoquery.CursorLocation = adUseClient
                     'Modify by Amy 2013/12/19 +公司別
                     'adoQuery.Open "select * from acc1p0 where a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p30").Value & "' or a1p30 = '" & adoacc1p0.Fields("a1p30").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                     adoquery.Open "select * from acc1p0 where  a1p01='" & adoacc1p0.Fields("a1p01") & "' And a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p30").Value & "' or a1p30 = '" & adoacc1p0.Fields("a1p30").Value & "') ", adoTaie, adOpenStatic, adLockReadOnly
                     If adoquery.RecordCount <> 0 Then
                        If IsNull(adoquery.Fields("a1p22").Value) Then
                           strOriAccNo = "null"
                        Else
                           strOriAccNo = "'" & adoquery.Fields("a1p22").Value & "'"
                        End If
                     Else
                        strOriAccNo = "null"
                     End If
                     adoquery.Close
                  Else
                     strOriAccNo = "null"
                  End If
               End If
               If IsNull(adoacc1p0.Fields("a1p17").Value) Then
                  strCaseNo = "null"
               Else
                  strCaseNo = "'" & Mid(adoacc1p0.Fields("a1p17").Value, 1, Len(adoacc1p0.Fields("a1p17").Value) - 9) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 8, 6) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 2, 1) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 1, 2) & "'"
               End If
'               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & strSerialNo & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value)) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value)) & ", " & strCaseNo & ", null, " & strOriAccNo & ", '" & strRemark & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value)) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               'Modify by Morgan 2007/8/10
               'adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & strSerialNo & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", " & CNULL(ChgSQL(IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value))) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value)) & ", " & strCaseNo & ", null, " & strOriAccNo & ", '" & strRemark & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value)) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & strSerialNo & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", " & CNULL(ChgSQL(IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value))) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value)) & ", " & strCaseNo & ", " & stAX210 & ", " & strOriAccNo & ", '" & strRemark & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value)) & ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               'end 2007/8/10
'               adoTaie.Execute "update acc1p0 set a1p27 = null, a1p03 = '" & strSerialNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & adoacc1p0.Fields("a1p04").Value & "'"
               '2006/1/5 MODIFY BY SONIA不更新A1P03
               'adoTaie.Execute "update acc1p0 set a1p27 = null, a1p03 = '" & strSerialNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & ChgSQL("" & adoacc1p0.Fields("a1p04").Value) & "'"
               adoTaie.Execute "update acc1p0 set a1p27 = null where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & ChgSQL("" & adoacc1p0.Fields("a1p04").Value) & "'"
               intCounter = intCounter + 1
               adoacc1p0.MoveNext
            Loop
         End If
         adoacc1p0.Close
         strName = (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p02").Value & adoprimary.Fields("a1p22").Value)
      End If
NextRecord:
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoprimary.MoveNext
   Loop
   List1.AddItem "應收/付傳票更新完成 " & adoprimary.RecordCount & " 筆"
   adoprimary.Close
   ProgressBar1.Value = 0
   ProgressBar2.Value = ProgressBar2.Value + 1
   
' 收票兌現傳票轉檔
   StatusView MsgText(79)
   strName = ""
   adoprimary.CursorLocation = adUseClient
   '2006/9/8 MODIFY BY SONIA 加 A0E04='R'
   'adoprimary.Open "select a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1) as typeno from acc1p0, acc0e0 where a1p09 = a0e02 and a1p02 = 'L' and a1p26 = '1' and a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1)", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2013/12/19 +公司別 修正寫法讓速度變快
   'adoprimary.Open "select a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1) as typeno from acc1p0, acc0e0 where a1p09 = a0e02 AND A0E04='R' and a1p02 = 'L' and a1p26 = '1' and a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1)", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2018/9/11 加order by
   adoprimary.Open "select a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1) as typeno from acc1p0, acc0e0 where  a1p09 = a0e02(+) AND A0E04='R' and a1p02||'' = 'L' and a1p26||'' = '1' and a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " And a1p01||''='" & strCompany & "' group by a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1) order by a1p01, a1p18, a0e20, substr(a1p04, length(a1p04), 1)", adoTaie, adOpenStatic, adLockReadOnly
   If adoprimary.RecordCount <> 0 Then
      ProgressBar1.max = adoprimary.RecordCount
   End If
   Do While adoprimary.EOF = False
      If strName <> (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p18").Value & adoprimary.Fields("a0e20").Value & adoprimary.Fields("typeno").Value) Then
         intCounter = 0
         adoacc1p0.CursorLocation = adUseClient
         '2012/10/11 MODIFY BY SONIA 瑞婷說改為輸資料的先後順序
         'adoacc1p0.Open "select * from acc1p0, acc0e0 where A0E04='R' and substr(a1p04, 1, length(a1p04) - 10) = a0e02 and substr(a1p04, length(a1p04) - 9, 9) = a0e01 and a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = 'L' and a1p26 = '1' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a0e20 = '" & adoprimary.Fields("a0e20").Value & "' and substr(a1p04, length(a1p04), 1) = '" & adoprimary.Fields("typeno").Value & "' and a1p22 is null order by a1p07 desc, a1p08 desc", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2013/12/19 修正寫法讓速度變快
         'adoacc1p0.Open "select * from acc1p0, acc0e0 where A0E04='R' and substr(a1p04, 1, length(a1p04) - 10) = a0e02 and substr(a1p04, length(a1p04) - 9, 9) = a0e01 and a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = 'L' and a1p26 = '1' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a0e20 = '" & adoprimary.Fields("a0e20").Value & "' and substr(a1p04, length(a1p04), 1) = '" & adoprimary.Fields("typeno").Value & "' and a1p22 is null order by a1p03, a1p29, a1p07 desc, a1p08 desc", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2020/07/23 +a0e07 因改為key
         'adoacc1p0.Open "select * from acc1p0, acc0e0 where A0E04='R' and substr(a1p04, 1, length(a1p04) - 10) = a0e02(+) and substr(a1p04, length(a1p04) - 9, 9) = a0e01(+) and a1p01||'' = '" & adoprimary.Fields("a1p01").Value & "' and a1p02||'' = 'L' and a1p26 = '1' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a0e20 = '" & adoprimary.Fields("a0e20").Value & "' and substr(a1p04, length(a1p04), 1) = '" & adoprimary.Fields("typeno").Value & "' and a1p22 is null order by a1p03, a1p29, a1p07 desc, a1p08 desc", adoTaie, adOpenStatic, adLockReadOnly
         adoacc1p0.Open "select * from acc1p0, acc0e0 where A0E04='R' and substr(a1p04, 1,7) = a0e02(+) and substr(a1p04, 8, 9) = a0e01(+) and substr(a1p04, 17,length(a1p04) - 17)=a0e07(+) and a1p01||'' = '" & adoprimary.Fields("a1p01").Value & "' and a1p02||'' = 'L' and a1p26 = '1' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a0e20 = '" & adoprimary.Fields("a0e20").Value & "' and substr(a1p04, length(a1p04), 1) = '" & adoprimary.Fields("typeno").Value & "' and a1p22 is null order by a1p03, a1p29, a1p07 desc, a1p08 desc", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount > 1 Then
            'Modify by Morgan 2009/10/2 自回圈外移進來以免跳號
            'Modify by Amy 2013/12/19 +針對不同公司別傳入字串產生編號
            'strAutoNo = AccAutoNo(MsgText(801), 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            'strSave = AccSaveAutoNo(MsgText(801), Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strAutoNo = AccAutoNo(IdentifierStr, 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strSave = AccSaveAutoNo(IdentifierStr, Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            adoTaie.Execute "insert into acc020 values('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '', 0, 0)"
            Do While adoacc1p0.EOF = False
               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & ZeroBeforeNo(CStr(intCounter), 3) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", '" & IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p17").Value), "", adoacc1p0.Fields("a1p17").Value) & "', null, null, '" & IIf(IsNull(adoacc1p0.Fields("a1p14").Value), "", adoacc1p0.Fields("a1p14").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value) & "'" & _
               ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & adoacc1p0.Fields("a1p04").Value & "'"
               'Debug.Print adoacc1p0.Fields("a1p04").Value
               intCounter = intCounter + 1
               adoacc1p0.MoveNext
            Loop
         End If
         adoacc1p0.Close
         strName = (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p18").Value & adoprimary.Fields("a0e20").Value & adoprimary.Fields("typeno").Value)
      End If
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoprimary.MoveNext
   Loop
   List1.AddItem "收票兌現傳票轉檔完成 " & adoprimary.RecordCount & " 筆"
   adoprimary.Close
   ProgressBar1.Value = 0
   ProgressBar2.Value = ProgressBar2.Value + 1
   
' 開票兌領傳票轉檔
   StatusView MsgText(79)
   strName = ""
   adoprimary.CursorLocation = adUseClient
   'Modify by Amy 2013/12/19 +公司別
   'adoprimary.Open "select a1p01, a1p18, a1p11 from acc1p0 where a1p02 = 'L' and a1p26 = '2' and a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by a1p01, a1p18, a1p11", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2018/9/11 加order by
   adoprimary.Open "select a1p01, a1p18, a1p11 from acc1p0 where a1p02||'' = 'L' and a1p26||'' = '2' and a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " And a1p01||''='" & strCompany & "' group by a1p01, a1p18, a1p11 order by a1p01, a1p18, a1p11", adoTaie, adOpenStatic, adLockReadOnly
   If adoprimary.RecordCount <> 0 Then
      ProgressBar1.max = adoprimary.RecordCount
   End If
   Do While adoprimary.EOF = False
      If strName <> (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p18").Value & adoprimary.Fields("a1p11").Value) Then
         intCounter = 0
         adoacc1p0.CursorLocation = adUseClient
'         adoacc1p0.Open "select * from acc1p0, acc0h0 where a1p10 = a0h01 and a1p11 = a0h02 and a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = 'L' and a1p26 = '2' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a1p11 = '" & adoprimary.Fields("a1p11").Value & "' and a1p08 = 0 and a1p22 is null order by a1p07 desc", adoTaie, adOpenStatic, adLockReadOnly
         adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = 'L' and a1p26 = '2' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a1p11 = '" & adoprimary.Fields("a1p11").Value & "' and a1p22 is null order by a1p07 desc, a1p08 desc", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount > 1 Then
            'Modify by Morgan 2009/10/2 自回圈外移進來以免跳號
            'Modify by Amy 2013/12/19 +針對不同公司別傳入字串產生編號
            'strAutoNo = AccAutoNo(MsgText(801), 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            'strSave = AccSaveAutoNo(MsgText(801), Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strAutoNo = AccAutoNo(IdentifierStr, 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strSave = AccSaveAutoNo(IdentifierStr, Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            adoTaie.Execute "insert into acc020 values('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '', 0, 0)"
            Do While adoacc1p0.EOF = False
               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & ZeroBeforeNo(CStr(intCounter), 3) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", '" & IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p17").Value), "", adoacc1p0.Fields("a1p17").Value) & "', null, null, '" & IIf(IsNull(adoacc1p0.Fields("a1p14").Value), "", adoacc1p0.Fields("a1p14").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value) & "'" & _
               ", " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & adoacc1p0.Fields("a1p04").Value & "'"
               intCounter = intCounter + 1
               adoacc1p0.MoveNext
            Loop
            adoacc1p0.MoveLast
'            adoaccsum.CursorLocation = adUseClient
'            adoaccsum.Open "select sum(a1p08) from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = 'L' and a1p26 = '2' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a1p11 = '" & adoprimary.Fields("a1p11").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) = False Then
'                  adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & ZeroBeforeNo(CStr(intCounter), 3) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a0h08").Value & "', 0, " & adoaccsum.Fields(0).Value & ", '" & IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p17").Value), "", adoacc1p0.Fields("a1p17").Value) & "', null, null, '" & IIf(IsNull(adoacc1p0.Fields("a1p14").Value), "", adoacc1p0.Fields("a1p14").Value) & "', null)"
'               Else
'                  adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & ZeroBeforeNo(CStr(intCounter), 3) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a0h08").Value & "', 0, 0, '" & IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p17").Value), "", adoacc1p0.Fields("a1p17").Value) & "', null, null, '" & IIf(IsNull(adoacc1p0.Fields("a1p14").Value), "", adoacc1p0.Fields("a1p14").Value) & "', null)"
'               End If
'            End If
'            adoaccsum.Close
'            adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "' where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = 'L' and a1p18 = " & adoprimary.Fields("a1p18").Value & " and a1p11 = '" & adoprimary.Fields("a1p11").Value & "' and a1p26 = '2'"
         End If
         adoacc1p0.Close
         strName = (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p18").Value & adoprimary.Fields("a1p11").Value)
      End If
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoprimary.MoveNext
   Loop
   List1.AddItem "開票兌領傳票轉檔完成 " & adoprimary.RecordCount & " 筆"
   adoprimary.Close
   ProgressBar1.Value = 0
   ProgressBar2.Value = ProgressBar2.Value + 1
   
'2005/11/15 ADD BY SONIA
' 結餘傳票轉檔
   StatusView MsgText(79)
   strName = ""
   adoprimary.CursorLocation = adUseClient
   'Modify by Amy 2013/12/19 +公司別
   'adoprimary.Open "select a1p01, A1P02, a1p28, a1p29, a1p18, a1p04 from acc1p0 where A1P02='S' AND a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by a1p01, a1p02, a1p18, a1p28, a1p29, a1p04", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2018/9/11 加order by
   adoprimary.Open "select a1p01, A1P02, a1p28, a1p29, a1p18, a1p04 from acc1p0 where A1P02||''='S' AND a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " And a1p01||''='" & strCompany & "' group by a1p01, a1p02, a1p18, a1p28, a1p29, a1p04 order by a1p01, a1p02, a1p18, a1p28, a1p29, a1p04", adoTaie, adOpenStatic, adLockReadOnly
   If adoprimary.RecordCount <> 0 Then
      ProgressBar1.max = adoprimary.RecordCount
   End If
   Do While adoprimary.EOF = False
      If strName <> (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p02").Value & adoprimary.Fields("a1p04").Value) Then
         intCounter = 0
         
         '檢查分錄是否平衡
         If adoacc1p0.State = adStateOpen Then
            adoacc1p0.Close
         End If
         adoacc1p0.CursorLocation = adUseClient
         m_A1P02 = adoprimary.Fields("a1p02").Value
         If adoprimary.Fields("a1p02").Value = "a" Then
            m_A1P02 = "C"
         End If
         adoacc1p0.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & ChgSQL("" & adoprimary.Fields("a1p04").Value) & "' and a1p22 is null", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount <> 0 Then
            If IsNull(adoacc1p0.Fields(0).Value) Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextSkip1
            End If
            If IsNull(adoacc1p0.Fields(1).Value) Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextSkip1
            End If
            If adoacc1p0.Fields(0).Value <> adoacc1p0.Fields(1).Value Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextSkip1
            End If
         Else
            List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
            intErr = intErr + 1 'Add by Amy 2014/01/20
            GoTo NextSkip1
         End If
         adoacc1p0.Close
         
         '取得分錄明細資料
         adoacc1p0.CursorLocation = adUseClient
         adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & ChgSQL("" & adoprimary.Fields("a1p04").Value) & "' and a1p22 is null order by decode(a1p02, 'W', a1p08, 'L', a1p08, a1p03) asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount > 1 Then
            'Modify by Amy 2013/12/19 +針對不同公司別傳入字串產生編號
            'strAutoNo = AccAutoNo(MsgText(801), 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            'strSave = AccSaveAutoNo(MsgText(801), Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strAutoNo = AccAutoNo(IdentifierStr, 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strSave = AccSaveAutoNo(IdentifierStr, Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            adoTaie.Execute "insert into acc020 values('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '', 0, 0)"
            Do While adoacc1p0.EOF = False
               If IsNull(adoacc1p0.Fields("a1p14").Value) Then
                  strRemark = ""
               Else
                  strRemark = Replace(adoacc1p0.Fields("a1p14").Value, "'", "''")
               End If
               If adoacc1p0.Fields("a1p02").Value = "W" Or adoacc1p0.Fields("a1p02").Value = "L" Then
                  strSerialNo = ZeroBeforeNo(Trim(str(intCounter)), 3)
               Else
                  strSerialNo = adoacc1p0.Fields("a1p03").Value
               End If
               If adoacc1p0.Fields("a1p05").Value = "2401" And adoacc1p0.Fields("a1p07").Value <> 0 And IsNull(adoacc1p0.Fields("a1p23").Value) = False Then
                  If adoquery.State = adStateOpen Then
                     adoquery.Close
                  End If
                  adoquery.CursorLocation = adUseClient
                  'Modify by Amy 2013/12/19 +公司別
                  'adoQuery.Open "select * from acc1p0 where a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p23").Value & "' or a1p23 = '" & adoacc1p0.Fields("a1p23").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                  adoquery.Open "select * from acc1p0 where a1p01='" & adoacc1p0.Fields("a1p01").Value & "' and a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p23").Value & "' or a1p23 = '" & adoacc1p0.Fields("a1p23").Value & "') ", adoTaie, adOpenStatic, adLockReadOnly
                  If adoquery.RecordCount <> 0 Then
                     If IsNull(adoquery.Fields("a1p22").Value) Then
                        strOriAccNo = "null"
                     Else
                        strOriAccNo = "'" & adoquery.Fields("a1p22").Value & "'"
                     End If
                  Else
                     strOriAccNo = "null"
                  End If
                  adoquery.Close
               Else
                  If adoacc1p0.Fields("a1p05").Value = "2401" And adoacc1p0.Fields("a1p07").Value <> 0 And IsNull(adoacc1p0.Fields("a1p30").Value) = False Then
                     If adoquery.State = adStateOpen Then
                        adoquery.Close
                     End If
                     adoquery.CursorLocation = adUseClient
                     'Modify by Amy 2013/12/19 +公司別
                     'adoQuery.Open "select * from acc1p0 where a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p30").Value & "' or a1p30 = '" & adoacc1p0.Fields("a1p30").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                     adoquery.Open "select * from acc1p0 where  a1p01='" & adoacc1p0.Fields("a1p01").Value & "' and a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p30").Value & "' or a1p30 = '" & adoacc1p0.Fields("a1p30").Value & "') ", adoTaie, adOpenStatic, adLockReadOnly
                     If adoquery.RecordCount <> 0 Then
                        If IsNull(adoquery.Fields("a1p22").Value) Then
                           strOriAccNo = "null"
                        Else
                           strOriAccNo = "'" & adoquery.Fields("a1p22").Value & "'"
                        End If
                     Else
                        strOriAccNo = "null"
                     End If
                     adoquery.Close
                  Else
                     strOriAccNo = "null"
                  End If
               End If
               If IsNull(adoacc1p0.Fields("a1p17").Value) Then
                  strCaseNo = "null"
               Else
                  strCaseNo = "'" & Mid(adoacc1p0.Fields("a1p17").Value, 1, Len(adoacc1p0.Fields("a1p17").Value) - 9) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 8, 6) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 2, 1) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 1, 2) & "'"
               End If
               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & strSerialNo & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", '" & ChgSQL(IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value)) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', " & strCaseNo & ", null, " & strOriAccNo & ", '" & strRemark & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value) & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               '2006/1/5 MODIFY BY SONIA不更新A1P03
               'adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "', a1p03 = '" & strSerialNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & ChgSQL("" & adoacc1p0.Fields("a1p04").Value) & "'"
               adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & ChgSQL("" & adoacc1p0.Fields("a1p04").Value) & "'"
               intCounter = intCounter + 1
               adoacc1p0.MoveNext
            Loop
         End If
         adoacc1p0.Close
         strName = (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p02").Value & adoprimary.Fields("a1p04").Value)
      End If
NextSkip1:
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoprimary.MoveNext
   Loop
   List1.AddItem "結餘傳票轉檔完成 " & adoprimary.RecordCount & " 筆"
   adoprimary.Close
   ProgressBar1.Value = 0
   ProgressBar2.Value = ProgressBar2.Value + 1
'2005/11/16 END
   
' 應收/付傳票轉檔
   StatusView MsgText(79)
   strName = ""
   adoprimary.CursorLocation = adUseClient
   '92.6.30 modify by sonia 國內應付系統之付款傳票改在最後產生 a1p02='C'改為'a'
   'adoprimary.Open "select a1p01, a1p02, a1p28, a1p29, a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p04 from acc1p0 where a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by a1p01, a1p02, a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p28, a1p29, a1p04", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2013/12/19 +公司別
   'adoprimary.Open "select a1p01, decode(a1p02, 'C', 'a', a1p02) AS A1P02, a1p28, a1p29, a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p04 from acc1p0 where a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " group by a1p01, decode(a1p02, 'C', 'a', a1p02), a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p28, a1p29, a1p04", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2018/9/11 加order by
   adoprimary.Open "select a1p01, decode(a1p02, 'C', 'a', a1p02) AS A1P02, a1p28, a1p29, a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p04 from acc1p0 where  a1p22 is null and a1p18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a1p18 <= " & Val(FCDate(MaskEdBox2.Text)) & " And a1p01||''='" & strCompany & "' group by a1p01, decode(a1p02, 'C', 'a', a1p02), a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p28, a1p29, a1p04 order by a1p01, decode(a1p02, 'C', 'a', a1p02), a1p18, decode(a1p02, 'I', substr(a1p04, length(a1p04) - 8, 9)||substr(a1p04, 1, length(a1p04) - 9), a1p04), a1p28, a1p29, a1p04", adoTaie, adOpenStatic, adLockReadOnly
   '92.6.30 end
   If adoprimary.RecordCount <> 0 Then
      ProgressBar1.max = adoprimary.RecordCount
   End If
   Do While adoprimary.EOF = False
      If strName <> (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p02").Value & adoprimary.Fields("a1p04").Value) Then
         intCounter = 0
         
         '檢查分錄是否平衡
         If adoacc1p0.State = adStateOpen Then
            adoacc1p0.Close
         End If
         adoacc1p0.CursorLocation = adUseClient
         '92.7.2 add by sonia
         m_A1P02 = adoprimary.Fields("a1p02").Value
         If adoprimary.Fields("a1p02").Value = "a" Then
            m_A1P02 = "C"
         End If
         'Debug.Print adoprimary.Fields("a1p04").Value
         '92.7.2 end
'         adoacc1p0.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & adoprimary.Fields("a1p04").Value & "' and a1p22 is null", adoTaie, adOpenStatic, adLockReadOnly
         adoacc1p0.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & ChgSQL("" & adoprimary.Fields("a1p04").Value) & "' and a1p22 is null", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc1p0.RecordCount <> 0 Then
            If IsNull(adoacc1p0.Fields(0).Value) Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextSkip
            End If
            If IsNull(adoacc1p0.Fields(1).Value) Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextSkip
            End If
            If adoacc1p0.Fields(0).Value <> adoacc1p0.Fields(1).Value Then
               List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
               intErr = intErr + 1 'Add by Amy 2014/01/20
               GoTo NextSkip
            End If
         Else
            List1.AddItem MsgText(209) & adoprimary.Fields("a1p04").Value
            intErr = intErr + 1 'Add by Amy 2014/01/20
            GoTo NextSkip
         End If
         adoacc1p0.Close
         
         '取得分錄明細資料
         adoacc1p0.CursorLocation = adUseClient
         '92.6.16 MODIFY BY SONIA
         'adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & adoprimary.Fields("a1p02").Value & "' and a1p04 = '" & adoprimary.Fields("a1p04").Value & "' and a1p22 is null order by decode(a1p02, 'W', a1p08, 'L', a1p08, a1p03) asc", adoTaie, adOpenStatic, adLockReadOnly
'         adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & adoprimary.Fields("a1p04").Value & "' and a1p22 is null order by decode(a1p02, 'W', a1p08, 'L', a1p08, a1p03) asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
         adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & adoprimary.Fields("a1p01").Value & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & ChgSQL("" & adoprimary.Fields("a1p04").Value) & "' and a1p22 is null order by decode(a1p02, 'W', a1p08, 'L', a1p08, a1p03) asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
         '92.6.16 END
         If adoacc1p0.RecordCount > 1 Then
            'Modify by Morgan 2009/10/2 自回圈外移進來以免跳號
            'Modify by Amy 2013/12/19 +針對不同公司別傳入字串產生編號
            'strAutoNo = AccAutoNo(MsgText(801), 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            'strSave = AccSaveAutoNo(MsgText(801), Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strAutoNo = AccAutoNo(IdentifierStr, 4, Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            strSave = AccSaveAutoNo(IdentifierStr, Mid(strAutoNo, 7, 4), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 1, 3)), Val(Mid(CFDate(adoprimary.Fields("a1p18").Value), 5, 2)))
            adoTaie.Execute "insert into acc020 values('" & adoprimary.Fields("a1p01").Value & "', '" & strAutoNo & "', " & adoprimary.Fields("a1p18").Value & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '', 0, 0)"
            Do While adoacc1p0.EOF = False
               If IsNull(adoacc1p0.Fields("a1p14").Value) Then
                  strRemark = ""
               Else
                  strRemark = Replace(adoacc1p0.Fields("a1p14").Value, "'", "''")
               End If
               If adoacc1p0.Fields("a1p02").Value = "W" Or adoacc1p0.Fields("a1p02").Value = "L" Then
                  strSerialNo = ZeroBeforeNo(Trim(str(intCounter)), 3)
               Else
                  strSerialNo = adoacc1p0.Fields("a1p03").Value
               End If
               If adoacc1p0.Fields("a1p05").Value = "2401" And adoacc1p0.Fields("a1p07").Value <> 0 And IsNull(adoacc1p0.Fields("a1p23").Value) = False Then
                  If adoquery.State = adStateOpen Then
                     adoquery.Close
                  End If
                  adoquery.CursorLocation = adUseClient
                  'Modify by Amy 2013/12/19 +公司別
                  'adoQuery.Open "select * from acc1p0 where a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p23").Value & "' or a1p23 = '" & adoacc1p0.Fields("a1p23").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                  adoquery.Open "select * from acc1p0 where a1p01='" & adoacc1p0.Fields("a1p01").Value & "' And a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p23").Value & "' or a1p23 = '" & adoacc1p0.Fields("a1p23").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                  If adoquery.RecordCount <> 0 Then
                     If IsNull(adoquery.Fields("a1p22").Value) Then
                        strOriAccNo = "null"
                     Else
                        strOriAccNo = "'" & adoquery.Fields("a1p22").Value & "'"
                     End If
                  Else
                     strOriAccNo = "null"
                  End If
                  adoquery.Close
               Else
                  If adoacc1p0.Fields("a1p05").Value = "2401" And adoacc1p0.Fields("a1p07").Value <> 0 And IsNull(adoacc1p0.Fields("a1p30").Value) = False Then
                     If adoquery.State = adStateOpen Then
                        adoquery.Close
                     End If
                     adoquery.CursorLocation = adUseClient
                     'Modify by Amy 2013/12/19 +公司別
                     'adoQuery.Open "select * from acc1p0 where a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p30").Value & "' or a1p30 = '" & adoacc1p0.Fields("a1p30").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                     adoquery.Open "select * from acc1p0 where a1p01='" & adoacc1p0.Fields("a1p01").Value & "' And a1p05 = '2401' and a1p08 <> 0 and (a1p04 = '" & adoacc1p0.Fields("a1p30").Value & "' or a1p30 = '" & adoacc1p0.Fields("a1p30").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
                     If adoquery.RecordCount <> 0 Then
                        If IsNull(adoquery.Fields("a1p22").Value) Then
                           strOriAccNo = "null"
                        Else
                           strOriAccNo = "'" & adoquery.Fields("a1p22").Value & "'"
                        End If
                     Else
                        strOriAccNo = "null"
                     End If
                     adoquery.Close
                  Else
                     strOriAccNo = "null"
                  End If
               End If
               If IsNull(adoacc1p0.Fields("a1p17").Value) Then
                  strCaseNo = "null"
               Else
                  strCaseNo = "'" & Mid(adoacc1p0.Fields("a1p17").Value, 1, Len(adoacc1p0.Fields("a1p17").Value) - 9) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 8, 6) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 2, 1) & Mid(adoacc1p0.Fields("a1p17").Value, Len(adoacc1p0.Fields("a1p17").Value) - 1, 2) & "'"
               End If
               'Debug.Print strSerialNo & " " & adoacc1p0.Fields("a1p03").Value
               'adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & adoacc1p0.Fields("a1p03").Value & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", '" & IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p17").Value), "", adoacc1p0.Fields("a1p17").Value) & "', null, null, '" & strRemark & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value) & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
'               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & strSerialNo & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", '" & IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', " & strCaseNo & ", null, " & strOriAccNo & ", '" & strRemark & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value) & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
               adoTaie.Execute "insert into acc021 values ('" & adoacc1p0.Fields("a1p01").Value & "', '" & strAutoNo & "', '" & strSerialNo & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p06").Value), "", adoacc1p0.Fields("a1p06").Value) & "', '" & adoacc1p0.Fields("a1p05").Value & "', " & IIf(IsNull(adoacc1p0.Fields("a1p07").Value), 0, adoacc1p0.Fields("a1p07").Value) & ", " & IIf(IsNull(adoacc1p0.Fields("a1p08").Value), 0, adoacc1p0.Fields("a1p08").Value) & ", '" & ChgSQL(IIf(IsNull(adoacc1p0.Fields("a1p15").Value), "", adoacc1p0.Fields("a1p15").Value)) & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p16").Value), "", adoacc1p0.Fields("a1p16").Value) & "', " & strCaseNo & ", null, " & strOriAccNo & ", '" & strRemark & "', '" & IIf(IsNull(adoacc1p0.Fields("a1p30").Value), "", adoacc1p0.Fields("a1p30").Value) & "', " & CNULL(IIf(IsNull(adoacc1p0.Fields("a1p31").Value), "", adoacc1p0.Fields("a1p31").Value)) & ")"
'               adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "', a1p03 = '" & strSerialNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & adoacc1p0.Fields("a1p04").Value & "'"
               '2006/1/5 MODIFY BY SONIA 不更新A1P03
               'adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "', a1p03 = '" & strSerialNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & ChgSQL("" & adoacc1p0.Fields("a1p04").Value) & "'"
               adoTaie.Execute "update acc1p0 set a1p22 = '" & strAutoNo & "' where a1p01 = '" & adoacc1p0.Fields("a1p01").Value & "' and a1p02 = '" & adoacc1p0.Fields("a1p02").Value & "' and a1p03 = '" & adoacc1p0.Fields("a1p03").Value & "' and a1p04 = '" & ChgSQL("" & adoacc1p0.Fields("a1p04").Value) & "'"
               intCounter = intCounter + 1
               adoacc1p0.MoveNext
            Loop
         End If
         adoacc1p0.Close
         strName = (adoprimary.Fields("a1p01").Value & adoprimary.Fields("a1p02").Value & adoprimary.Fields("a1p04").Value)
      End If
NextSkip:
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoprimary.MoveNext
   Loop
   List1.AddItem "應收/付傳票轉檔完成 " & adoprimary.RecordCount & " 筆"
   adoprimary.Close
   ProgressBar1.Value = 0
   ProgressBar2.Value = ProgressBar2.Value + 1
   'Add by Amy 2013/12/19
   List1.AddItem strCompany & "-公司轉檔結束"
Next ii
'end 2013/12/19
   StatusClear
   adoTaie.CommitTrans
   Screen.MousePointer = vbDefault
   If intErr > 0 Then
        MsgBox "轉檔有錯誤請參閱List！", , MsgText(21)
   Else
        MsgBox MsgText(78), , MsgText(21)
   End If
   Exit Sub
Checking:
   adoTaie.RollbackTrans
   If Err.Number = 0 Then
      Exit Sub
   End If
   'Resume Next 'Removed by Morgan 2021/11/1 不該有
   adoprimary.Close
   adoacc1p0.Close
   MsgBox Err.Description, , MsgText(5)
   Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2024/08/01 判斷傳入的日期北所是否為颱風假,則抓[非]颱風假的工作日
Private Function GetTyphoon(ByVal strDate As String) As String
   Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String, strWhere As String, m_Date As String
   
   GetTyphoon = strDate
   strQ = "Select * From WorkDay Where WD01=" & DBDATE(strDate)
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      '北所 放 颱風假
      If "" & RsQ.Fields("WD02") = "Y" Then
         '抓[非]颱風假的工作日
         strQ = "Select Max(WD01) From WorkDay Where WD01<=" & DBDATE(strDate) & " And WD02 is Null "
         intQ = 1
         Set RsQ = ClsLawReadRstMsg(intQ, strQ)
         If intQ = 1 Then
            GetTyphoon = RsQ.Fields(0)
         End If
      End If
   End If
   
   Set RsQ = Nothing
End Function

