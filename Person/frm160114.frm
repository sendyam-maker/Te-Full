VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160114 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式人事資料表"
   ClientHeight    =   3090
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4990
   Begin VB.Frame Frame2 
      Height          =   1550
      Left            =   90
      TabIndex        =   8
      Top             =   660
      Width           =   4755
      Begin VB.TextBox txt1 
         Height          =   285
         Index           =   0
         Left            =   2490
         MaxLength       =   5
         TabIndex        =   2
         Top             =   810
         Width           =   795
      End
      Begin VB.OptionButton Option1 
         Caption         =   "職災申報統計"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   1
         Top             =   870
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "員工類別統計"
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   510
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "年月："
         Height          =   260
         Left            =   1950
         TabIndex        =   9
         Top             =   870
         Width           =   590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   6
      Top             =   2430
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2940
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3900
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm160114.frx":0000
      Height          =   1665
      Left            =   270
      TabIndex        =   10
      Top             =   1410
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7691
      _ExtentY        =   2946
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm160114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create By Sindy 2012/2/13
Option Explicit

Dim m_StrSQL As String
Dim m_str  As String
Dim m_rs As New ADODB.Recordset
Dim m_i As Integer
Dim PLeft(1 To 11) As Integer
Dim strTemp(1 To 11) As String
Dim iPgae As Integer, iLine As Integer
Dim dblTotCnt As Double
Const A0925CName_114 As String = "decode(substr(st93,1,1),'B','01業務拓展部','F','02專利國外部','J','03專利日本部','L','04法律所','M','05管理部','P','06專利國內部','S','07智權部','T','08商標部','W','09顧問服務組','Y','10創新業務部',substr(st93,1,1))"


Private Sub cmdok_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0
         Set Printer = Printers(Combo1.ListIndex)
         Printer.Orientation = 1
         Printer.EndDoc
         If Option1(0).Value = True Then
            '含留職停薪,不含兼職(68007林信昌)
            '所別其他計北所(76028何尤玉),S01(74028邱素蓮)列入管理部
            '人數總計:含留職停薪,不含兼職
            dblTotCnt = 0
            'Modify By Sindy 2023/12/28 部門調整改抓ST93
            'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
            strSql = "select count(*) from staff,salarydata " & _
                     "where (st04='1' or sd02='S') and not(substr(st01,5,1)>='A') " & _
                     "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
                     "and st01=sd01(+) and sd02 in ('R','T','S') "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               dblTotCnt = RsTemp.Fields(0)
            End If
            
            Screen.MousePointer = vbHourglass
            If StrMenu1 = False And StrMenu2 = False And StrMenu3 = False Then
               ShowNoData
            Else
               Printer.EndDoc
               ShowPrintOk
            End If
            
         'Add By Sindy 2017/1/26
         ElseIf Option1(1).Value = True Then
            If txt1(0) = "" Then
               MsgBox "年月不可空白！", vbInformation, "操作錯誤！"
               txt1(0).SetFocus
               Exit Sub
            Else
               Cancel = False
               Call txt1_Validate(0, Cancel)
               If Cancel = True Then
                  txt1(0).SetFocus
                  Exit Sub
               End If
            End If
            
            Screen.MousePointer = vbHourglass
            If StrMenu4 = False Then
               'ShowNoData
            Else
               Printer.EndDoc
               ShowPrintOk
            End If
         '2017/1/26 END
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

'各所大部門性別
Function StrMenu1() As Boolean
Dim dblCnt As Double
Dim i As Integer, j As Integer, k As Integer
Dim dblCol_11 As Double, dblCol_12 As Double, dblCol_13 As Double, dblCol_14 As Double, dblCol_1Tot As Double
Dim dblCol_21 As Double, dblCol_22 As Double, dblCol_23 As Double, dblCol_24 As Double, dblCol_2Tot As Double

StrMenu1 = True

dblCol_11 = 0: dblCol_12 = 0: dblCol_13 = 0: dblCol_14 = 0: dblCol_1Tot = 0
dblCol_21 = 0: dblCol_22 = 0: dblCol_23 = 0: dblCol_24 = 0: dblCol_2Tot = 0

'Modify By Sindy 2015/10/1 +4法務部
'DECODE(ST03,'S01','4管理部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','2專業部','F','3國外部','M','4管理部','4管理部'))
'Modify By Sindy 2019/4/3 + ,'D01','6研發處','T10','7創新業務部','T20','7創新業務部'
'Modify By Sindy 2023/12/28 部門調整改抓ST93
'm_str = "select DECODE(ST03,'S01','5管理部','F31','4法務部','P31','4法務部','D01','6研發處','T10','7創新業務部','T20','7創新業務部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','4法務部','F','3國外部','M','5管理部','5管理部')), " & _
'        "DECODE(ST22,'M','1男','F','2女',ST22), " & _
'        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北'),count(*) NUM " & _
'        "From staff, salarydata " & _
'        "where (st04='1' or sd02='S') " & _
'        "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
'        "and st01=sd01(+) and sd02 in ('R','T','S') " & _
'        "GROUP BY DECODE(ST03,'S01','5管理部','F31','4法務部','P31','4法務部','D01','6研發處','T10','7創新業務部','T20','7創新業務部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','4法務部','F','3國外部','M','5管理部','5管理部')), " & _
'        "DECODE(ST22,'M','1男','F','2女',ST22), " & _
'        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北') "
'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
m_str = "select " & A0925CName_114 & ", " & _
        "DECODE(ST22,'M','1男','F','2女',ST22), " & _
        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北'),count(*) NUM " & _
        "From staff, salarydata " & _
        "where (st04='1' or sd02='S') " & _
        "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
        "and st01=sd01(+) and sd02 in ('R','T','S') and not(substr(st01,5,1)>='A') " & _
        "GROUP BY " & A0925CName_114 & ", " & _
        "DECODE(ST22,'M','1男','F','2女',ST22), " & _
        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北') " & _
        "ORDER BY " & A0925CName_114 & ", " & _
        "DECODE(ST22,'M','1男','F','2女',ST22), " & _
        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北') "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   With m_rs
      .MoveFirst
      PrintTitle
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "■各所大部門性別人數統計"
      iLine = iLine + 2
      PrintTitle1
      Do While Not .EOF
         For i = 1 To 10 '7 個部門
            For m_i = 1 To 11 '9 '8
               strTemp(m_i) = ""
            Next m_i
'            If i = 1 Then strTemp(1) = "智權部"
'            If i = 2 Then strTemp(1) = "專業部"
'            If i = 3 Then strTemp(1) = "國外部"
'            If i = 4 Then strTemp(1) = "法務部"
'            If i = 5 Then strTemp(1) = "管理部"
'            If i = 6 Then strTemp(1) = "研發處" 'Add By Sindy 2019/4/3
'            If i = 7 Then strTemp(1) = "創新業務部" 'Add By Sindy 2019/4/3
            strTemp(1) = Mid(.Fields(0), 3)
            For j = 1 To 2 '性別
               If j = 1 Then strTemp(2) = "男"
               If j = 2 Then strTemp(2) = "女"
               dblCnt = 0
               For m_i = 3 To 11 '9 '8
                   strTemp(m_i) = ""
               Next m_i
               For k = 1 To 4 '所別
                  If .EOF = True Then Exit For
                  If Val(Left(Trim(.Fields(0)), 2)) = i And Left(Trim(.Fields(1)), 1) = j Then
                     Select Case Left(Trim(.Fields(2)), 1)
                        Case "1"
                           strTemp(3) = CheckStr(.Fields(3))
                           dblCnt = dblCnt + Val(.Fields(3))
                           If j = 1 Then dblCol_11 = dblCol_11 + Val(.Fields(3))
                           If j = 2 Then dblCol_21 = dblCol_21 + Val(.Fields(3))
                        Case "2"
                           strTemp(4) = CheckStr(.Fields(3))
                           dblCnt = dblCnt + Val(.Fields(3))
                           If j = 1 Then dblCol_12 = dblCol_12 + Val(.Fields(3))
                           If j = 2 Then dblCol_22 = dblCol_22 + Val(.Fields(3))
                        Case "3"
                           strTemp(5) = CheckStr(.Fields(3))
                           dblCnt = dblCnt + Val(.Fields(3))
                           If j = 1 Then dblCol_13 = dblCol_13 + Val(.Fields(3))
                           If j = 2 Then dblCol_23 = dblCol_23 + Val(.Fields(3))
                        Case "4"
                           strTemp(6) = CheckStr(.Fields(3))
                           dblCnt = dblCnt + Val(.Fields(3))
                           If j = 1 Then dblCol_14 = dblCol_14 + Val(.Fields(3))
                           If j = 2 Then dblCol_24 = dblCol_24 + Val(.Fields(3))
                     End Select
                     If j = 1 Then dblCol_1Tot = dblCol_1Tot + Val(.Fields(3))
                     If j = 2 Then dblCol_2Tot = dblCol_2Tot + Val(.Fields(3))
                  Else
                     '下一筆資料列了
                     Exit For
                  End If
                  .MoveNext
               Next k
               strTemp(7) = dblCnt '合計
               strTemp(8) = Round(dblCnt / dblTotCnt * 100, 2) & "%" '比例
               PrintDetail
            Next j
         Next i
      Loop
      '合計
      Call SetLineation
      strTemp(1) = "合計"
      strTemp(2) = "男"
      strTemp(3) = dblCol_11
      strTemp(4) = dblCol_12
      strTemp(5) = dblCol_13
      strTemp(6) = dblCol_14
      strTemp(7) = dblCol_1Tot
      strTemp(8) = Round(dblCol_1Tot / dblTotCnt * 100, 2) & "%"
      PrintDetail
      strTemp(1) = ""
      strTemp(2) = "女"
      strTemp(3) = dblCol_21
      strTemp(4) = dblCol_22
      strTemp(5) = dblCol_23
      strTemp(6) = dblCol_24
      strTemp(7) = dblCol_2Tot
      strTemp(8) = Round(dblCol_2Tot / dblTotCnt * 100, 2) & "%"
      PrintDetail
      iLine = iLine + 3
   End With
Else
   StrMenu1 = False
   Exit Function
End If
End Function

Sub PrintTitle1()
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部門"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("台北所")
Printer.CurrentY = iLine * 300
Printer.Print "台北所"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("台中所")
Printer.CurrentY = iLine * 300
Printer.Print "台中所"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("台南所")
Printer.CurrentY = iLine * 300
Printer.Print "台南所"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("高雄所")
Printer.CurrentY = iLine * 300
Printer.Print "高雄所"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("合計")
Printer.CurrentY = iLine * 300
Printer.Print "合計"
Printer.CurrentX = PLeft(8) - Printer.TextWidth("比例")
Printer.CurrentY = iLine * 300
Printer.Print "比例"
iLine = iLine + 1
Call SetLineation
End Sub

Sub GetPleft()
PLeft(1) = 1000 '2000
PLeft(2) = 3000
PLeft(3) = 4500
PLeft(4) = 5500
PLeft(5) = 6500
PLeft(6) = 7500
PLeft(7) = 8500
PLeft(8) = 9500
PLeft(9) = 10500
End Sub

'Add By Sindy 2019/4/3
Sub GetPleft2()
PLeft(1) = 600 '2000
PLeft(2) = 1200
PLeft(3) = 2200
PLeft(4) = 3200
PLeft(5) = 4200
PLeft(6) = 5200
PLeft(7) = 6200
PLeft(8) = 7200
PLeft(9) = 8600
PLeft(10) = 9500
PLeft(11) = 10500
End Sub

'Add By Sindy 2017/2/2
Sub GetPleft4()
PLeft(1) = 500 '2000
PLeft(2) = 5800
PLeft(3) = 6800
PLeft(4) = 9000
PLeft(5) = 10500
End Sub

Sub PrintTitle()
GetPleft
Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Val(Left(strSrvDate(1), 4)) - 1911 & "年人事處員工類別統計") / 2)
Printer.CurrentY = 300
Printer.Print Val(Left(strSrvDate(1), 4)) - 1911 & "年人事處員工類別統計"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
End Sub

Sub PrintDetail()
Dim m_j As Integer
   
   For m_j = 1 To 11 '9 '8
      If m_j <= 2 Then
         Printer.CurrentX = PLeft(m_j)
      Else
         Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
      End If
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String

   MoveFormToCenter Me

   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i

   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm160114 = Nothing
End Sub

'各所學歷
Function StrMenu2() As Boolean
Dim dblCnt As Double
Dim i As Integer, j As Integer, k As Integer
Dim dblCol_11 As Double, dblCol_12 As Double, dblCol_13 As Double, dblCol_14 As Double, dblCol_1Tot As Double

StrMenu2 = True

dblCol_11 = 0: dblCol_12 = 0: dblCol_13 = 0: dblCol_14 = 0: dblCol_1Tot = 0

'decode(st37,'01','5研究所','02','5研究所','03','5研究所','04','5研究所','05','4大學','06','4大學' " & _
        ",'07','3專科','08','3專科','09','3專科','10','3專科','11','3專科','12','3專科' " & _
        ",'13','2高中','14','2高中','15','2高中','16','2高中','1國中(含)以下')
'Modify By Sindy 2023/12/28 部門調整改抓ST93
'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
m_str = "select decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北'),count(*) NUM " & _
        "From staff, salarydata, allcode " & _
        "where (st04='1' or sd02='S') " & _
        "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
        "and st37 is not null " & _
        "and st01=sd01(+) and sd02 in ('R','T','S') and not(substr(st01,5,1)>='A') " & _
        "and '03'=ac01(+) and st37=ac02(+) " & _
        "GROUP BY decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
        "DECODE(ST06,'5','1',ST06)||DECODE(ST06,'1','北','2','中','3','南','4','高','北') "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   With m_rs
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "■各所學歷人數統計"
      iLine = iLine + 2
      PrintTitle2
      For i = 1 To 5 '學歷
         For m_i = 1 To 11 '9 '8
            strTemp(m_i) = ""
         Next m_i
         If i = 1 Then strTemp(1) = "國中"
         If i = 2 Then strTemp(1) = "高中"
         If i = 3 Then strTemp(1) = "專科"
         If i = 4 Then strTemp(1) = "大學"
         If i = 5 Then strTemp(1) = "研究所"
         dblCnt = 0
         .MoveFirst
         Do While Not .EOF
            If Left(Trim(.Fields(0)), 1) = i Then
               Select Case Left(Trim(.Fields(1)), 1)
                  Case "1"
                     strTemp(3) = CheckStr(.Fields(2))
                     dblCnt = dblCnt + Val(.Fields(2))
                     dblCol_11 = dblCol_11 + Val(.Fields(2))
                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
                  Case "2"
                     strTemp(4) = CheckStr(.Fields(2))
                     dblCnt = dblCnt + Val(.Fields(2))
                     dblCol_12 = dblCol_12 + Val(.Fields(2))
                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
                  Case "3"
                     strTemp(5) = CheckStr(.Fields(2))
                     dblCnt = dblCnt + Val(.Fields(2))
                     dblCol_13 = dblCol_13 + Val(.Fields(2))
                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
                  Case "4"
                     strTemp(6) = CheckStr(.Fields(2))
                     dblCnt = dblCnt + Val(.Fields(2))
                     dblCol_14 = dblCol_14 + Val(.Fields(2))
                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
               End Select
            End If
            .MoveNext
         Loop
         strTemp(7) = dblCnt '合計
         strTemp(8) = Round(dblCnt / dblTotCnt * 100, 2) & "%" '比例
         PrintDetail
      Next i
      '合計
      Call SetLineation
      strTemp(1) = "合計"
      strTemp(2) = ""
      strTemp(3) = dblCol_11
      strTemp(4) = dblCol_12
      strTemp(5) = dblCol_13
      strTemp(6) = dblCol_14
      strTemp(7) = dblCol_1Tot
      strTemp(8) = Round(dblCol_1Tot / dblTotCnt * 100, 2) & "%"
      PrintDetail
      iLine = iLine + 3
   End With
Else
   StrMenu2 = False
   Exit Function
End If
End Function

Sub PrintTitle2()
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "學歷"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print ""
Printer.CurrentX = PLeft(3) - Printer.TextWidth("台北所")
Printer.CurrentY = iLine * 300
Printer.Print "台北所"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("台中所")
Printer.CurrentY = iLine * 300
Printer.Print "台中所"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("台南所")
Printer.CurrentY = iLine * 300
Printer.Print "台南所"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("高雄所")
Printer.CurrentY = iLine * 300
Printer.Print "高雄所"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("合計")
Printer.CurrentY = iLine * 300
Printer.Print "合計"
Printer.CurrentX = PLeft(8) - Printer.TextWidth("比例")
Printer.CurrentY = iLine * 300
Printer.Print "比例"
iLine = iLine + 1
Call SetLineation
End Sub

'Modify By Sindy 2023/12/28
'大部門學歷
Function StrMenu3() As Boolean
Dim dblCnt As Double
Dim i As Integer, j As Integer, k As Integer
Dim dblCol_11 As Double, dblCol_12 As Double, dblCol_13 As Double, dblCol_14 As Double, dblCol_15 As Double, dblCol_1Tot As Double

StrMenu3 = True

dblCol_11 = 0: dblCol_12 = 0: dblCol_13 = 0: dblCol_14 = 0: dblCol_15 = 0: dblCol_1Tot = 0

'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
m_str = "select " & A0925CName_114 & ", " & _
        "decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
        "count(*) NUM " & _
        "From staff, salarydata, allcode " & _
        "where (st04='1' or sd02='S') " & _
        "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
        "and st37 is not null " & _
        "and st01=sd01(+) and sd02 in ('R','T','S') and not(substr(st01,5,1)>='A') " & _
        "and '03'=ac01(+) and st37=ac02(+) " & _
        "GROUP BY " & A0925CName_114 & ", " & _
        "decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下') " & _
        "ORDER BY " & A0925CName_114 & ", " & _
        "decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下') "

'm_str = "select decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
'        "DECODE(ST03,'S01','5管理部','F31','4法務部','P31','4法務部','D01','6研發處','T10','7創新業務部','T20','7創新業務部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','4法務部','F','3國外部','M','5管理部','5管理部')),count(*) NUM " & _
'        "From staff, salarydata, allcode " & _
'        "where (st04='1' or sd02='S') " & _
'        "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
'        "and st37 is not null " & _
'        "and st01=sd01(+) and sd02 in ('R','T','S') " & _
'        "and '03'=ac01(+) and st37=ac02(+) " & _
'        "GROUP BY decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
'        "DECODE(ST03,'S01','5管理部','F31','4法務部','P31','4法務部','D01','6研發處','T10','7創新業務部','T20','7創新業務部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','4法務部','F','3國外部','M','5管理部','5管理部')) "

If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   With m_rs
      .MoveFirst
      'Add By Sindy 2023/12/28
      Printer.NewPage
      PrintTitle
      '2023/12/28 END
      'Call GetPleft2 'Add By Sindy 2019/4/3
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "■大部門學歷人數統計"
      iLine = iLine + 2
      PrintTitle3
      Do While Not .EOF
         For i = 1 To 10 '7 個部門
            For m_i = 1 To 11 '9 '8
               strTemp(m_i) = ""
            Next m_i
            strTemp(1) = Mid(.Fields(0), 3) '部門名稱
            dblCnt = 0
'            For m_i = 2 To 11 '9 '8
'                strTemp(m_i) = ""
'            Next m_i
            For k = 1 To 5 '學歷
               If .EOF = True Then Exit For
               If Val(Left(Trim(.Fields(0)), 2)) = i Then
                  Select Case Left(Trim(.Fields(1)), 1)
                     Case "1"
                        strTemp(2) = CheckStr(.Fields(2))
                        dblCnt = dblCnt + Val(.Fields(2))
                        dblCol_11 = dblCol_11 + Val(.Fields(2))
                     Case "2"
                        strTemp(3) = CheckStr(.Fields(2))
                        dblCnt = dblCnt + Val(.Fields(2))
                        dblCol_12 = dblCol_12 + Val(.Fields(2))
                     Case "3"
                        strTemp(4) = CheckStr(.Fields(2))
                        dblCnt = dblCnt + Val(.Fields(2))
                        dblCol_13 = dblCol_13 + Val(.Fields(2))
                     Case "4"
                        strTemp(5) = CheckStr(.Fields(2))
                        dblCnt = dblCnt + Val(.Fields(2))
                        dblCol_14 = dblCol_14 + Val(.Fields(2))
                     Case "5"
                        strTemp(6) = CheckStr(.Fields(2))
                        dblCnt = dblCnt + Val(.Fields(2))
                        dblCol_15 = dblCol_15 + Val(.Fields(2))
                  End Select
                  dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
               Else
                  '下一筆資料列了
                  Exit For
               End If
               .MoveNext
            Next k
            strTemp(7) = dblCnt '合計
            strTemp(8) = Round(dblCnt / dblTotCnt * 100, 2) & "%" '比例
            PrintDetail
         Next i
      Loop
      '合計
      Call SetLineation
      strTemp(1) = "合計"
      strTemp(2) = dblCol_11
      strTemp(3) = dblCol_12
      strTemp(4) = dblCol_13
      strTemp(5) = dblCol_14
      strTemp(6) = dblCol_15
      strTemp(7) = dblCol_1Tot
      strTemp(8) = Round(dblCol_1Tot / dblTotCnt * 100, 2) & "%"
      PrintDetail
      iLine = iLine + 3
   End With
Else
   StrMenu3 = False
   Exit Function
End If
End Function

Sub PrintTitle3()
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部門"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "國中"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("高中")
Printer.CurrentY = iLine * 300
Printer.Print "高中"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("專科")
Printer.CurrentY = iLine * 300
Printer.Print "專科"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("大學")
Printer.CurrentY = iLine * 300
Printer.Print "大學"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("研究所")
Printer.CurrentY = iLine * 300
Printer.Print "研究所"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("合計")
Printer.CurrentY = iLine * 300
Printer.Print "合計"
Printer.CurrentX = PLeft(8) - Printer.TextWidth("比例")
Printer.CurrentY = iLine * 300
Printer.Print "比例"
iLine = iLine + 1
Call SetLineation
End Sub
'Function StrMenu3() As Boolean
'Dim dblCnt As Double
'Dim i As Integer, j As Integer, k As Integer
'Dim dblCol_11 As Double, dblCol_12 As Double, dblCol_13 As Double, dblCol_14 As Double, dblCol_1Tot As Double
'Dim dblCol_15 As Double, dblCol_16 As Double, dblCol_17 As Double
'
'StrMenu3 = True
'
'dblCol_11 = 0: dblCol_12 = 0: dblCol_13 = 0: dblCol_14 = 0: dblCol_1Tot = 0: dblCol_15 = 0
''decode(st37,'01','5研究所','02','5研究所','03','5研究所','04','5研究所','05','4大學','06','4大學' " & _
'        ",'07','3專科','08','3專科','09','3專科','10','3專科','11','3專科','12','3專科' " & _
'        ",'13','2高中','14','2高中','15','2高中','16','2高中','1國中(含)以下')
''Modify By Sindy 2019/4/3 + ,'D01','6研發處','T10','7創新業務部','T20','7創新業務部'
''Modify By Sindy 2023/12/28 部門調整改抓ST93
'm_str = "select decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
'        "DECODE(ST03,'S01','5管理部','F31','4法務部','P31','4法務部','D01','6研發處','T10','7創新業務部','T20','7創新業務部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','4法務部','F','3國外部','M','5管理部','5管理部')),count(*) NUM " & _
'        "From staff, salarydata, allcode " & _
'        "where (st04='1' or sd02='S') " & _
'        "and st01>='6' and st01<'F' and substr(st01,4,1)<'9' and substr(st93,1,1)<>'R' " & _
'        "and st37 is not null " & _
'        "and st01=sd01(+) and sd02 in ('R','T','S') " & _
'        "and '03'=ac01(+) and st37=ac02(+) " & _
'        "GROUP BY decode(substr(st37,1,1),'0','5研究所','1','4大學','2','3專科','3','2高中','1國中(含)以下'), " & _
'        "DECODE(ST03,'S01','5管理部','F31','4法務部','P31','4法務部','D01','6研發處','T10','7創新業務部','T20','7創新業務部',DECODE(SUBSTR(ST03,1,1),'S','1智權部','P','2專業部','L','4法務部','F','3國外部','M','5管理部','5管理部')) "
'If m_rs.State = 1 Then m_rs.Close
'm_rs.CursorLocation = adUseClient
'm_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'If Not m_rs.EOF And Not m_rs.BOF Then
'   With m_rs
'      'Add By Sindy 2023/12/28
'      Printer.NewPage
'      PrintTitle
'      '2023/12/28 END
'      Call GetPleft2 'Add By Sindy 2019/4/3
'      Printer.CurrentX = PLeft(1)
'      Printer.CurrentY = iLine * 300
'      Printer.Print "■大部門學歷人數統計"
'      iLine = iLine + 2
'      PrintTitle3
'      For i = 1 To 5 '學歷
'         For m_i = 1 To 11 '9 '8
'            strTemp(m_i) = ""
'         Next m_i
'         If i = 1 Then strTemp(1) = "國中"
'         If i = 2 Then strTemp(1) = "高中"
'         If i = 3 Then strTemp(1) = "專科"
'         If i = 4 Then strTemp(1) = "大學"
'         If i = 5 Then strTemp(1) = "研究所"
'         dblCnt = 0
'         .MoveFirst
'         Do While Not .EOF
'            If Left(Trim(.Fields(0)), 1) = i Then
'               Select Case Left(Trim(.Fields(1)), 1)
'                  Case "1"
'                     strTemp(3) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_11 = dblCol_11 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  Case "2"
'                     strTemp(4) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_12 = dblCol_12 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  Case "3"
'                     strTemp(5) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_13 = dblCol_13 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  Case "4"
'                     strTemp(6) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_14 = dblCol_14 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  Case "5"
'                     strTemp(7) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_15 = dblCol_15 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  'Add By Sindy 2019/4/3
'                  Case "6"
'                     strTemp(8) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_16 = dblCol_16 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  Case "7"
'                     strTemp(9) = CheckStr(.Fields(2))
'                     dblCnt = dblCnt + Val(.Fields(2))
'                     dblCol_17 = dblCol_17 + Val(.Fields(2))
'                     dblCol_1Tot = dblCol_1Tot + Val(.Fields(2))
'                  '2019/4/3 END
'               End Select
'            End If
'            .MoveNext
'         Loop
'         strTemp(10) = dblCnt '合計
'         strTemp(11) = Round(dblCnt / dblTotCnt * 100, 2) & "%" '比例
'         PrintDetail
'      Next i
'      '合計
'      Call SetLineation
'      strTemp(1) = "合計"
'      strTemp(2) = ""
'      strTemp(3) = dblCol_11
'      strTemp(4) = dblCol_12
'      strTemp(5) = dblCol_13
'      strTemp(6) = dblCol_14
'      strTemp(7) = dblCol_15
'      strTemp(8) = dblCol_16 'Add By Sindy 2019/4/3
'      strTemp(9) = dblCol_17 'Add By Sindy 2019/4/3
'      strTemp(10) = dblCol_1Tot
'      strTemp(11) = Round(dblCol_1Tot / dblTotCnt * 100, 2) & "%"
'      PrintDetail
'      iLine = iLine + 3
'   End With
'Else
'   StrMenu3 = False
'   Exit Function
'End If
'End Function
'
'Sub PrintTitle3()
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine * 300
'Printer.Print "學歷"
'Printer.CurrentX = PLeft(2)
'Printer.CurrentY = iLine * 300
'Printer.Print ""
'Printer.CurrentX = PLeft(3) - Printer.TextWidth("智權部")
'Printer.CurrentY = iLine * 300
'Printer.Print "智權部"
'Printer.CurrentX = PLeft(4) - Printer.TextWidth("專業部")
'Printer.CurrentY = iLine * 300
'Printer.Print "專業部"
'Printer.CurrentX = PLeft(5) - Printer.TextWidth("國外部")
'Printer.CurrentY = iLine * 300
'Printer.Print "國外部"
''Add By Sindy 2015/10/1
'Printer.CurrentX = PLeft(6) - Printer.TextWidth("法務部")
'Printer.CurrentY = iLine * 300
'Printer.Print "法務部"
''2015/10/1 END
'Printer.CurrentX = PLeft(7) - Printer.TextWidth("管理部")
'Printer.CurrentY = iLine * 300
'Printer.Print "管理部"
''Add By Sindy 2019/4/3
'Printer.CurrentX = PLeft(8) - Printer.TextWidth("研發處")
'Printer.CurrentY = iLine * 300
'Printer.Print "研發處"
'Printer.CurrentX = PLeft(9) - Printer.TextWidth("創新業務部")
'Printer.CurrentY = iLine * 300
'Printer.Print "創新業務部"
''2019/4/3 END
'Printer.CurrentX = PLeft(10) - Printer.TextWidth("合計")
'Printer.CurrentY = iLine * 300
'Printer.Print "合計"
'Printer.CurrentX = PLeft(11) - Printer.TextWidth("比例")
'Printer.CurrentY = iLine * 300
'Printer.Print "比例"
'iLine = iLine + 1
'Call SetLineation
'End Sub

Sub SetLineation()
   Printer.CurrentX = 300 '1500
   Printer.CurrentY = iLine * 300
   Printer.Print String(150, "-")
   iLine = iLine + 1
End Sub

'Add By Sindy 2017/1/26
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0         1           2           3           4           5
   arrGridHeadText = Array("公司別", "公司名稱", "男生人數", "女生人數", "請假時數", "加班時數")
   arrGridHeadWidth = Array(800, 800, 800, 800, 1000, 1000)
   'GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   'GRD1.Visible = True
End Sub

'Add By Sindy 2017/1/26
'職災申報統計
Function StrMenu4() As Boolean
Dim intWorkDay As Integer, ii As Integer, jj As Integer
Dim strText As String
   
   StrMenu4 = False
   
   '先檢查輸入的年月,是否已薪資入帳
   m_str = "select br01 from BookRecord where br01=" & Left(DBDATE(txt1(0) & "01"), 6)
   intI = 1
   Set m_rs = ClsLawReadRstMsg(intI, m_str)
   If intI = 0 Then
      MsgBox "此月份尚未薪資入帳, 出缺勤資料尚有可能未完備！", vbInformation, "警示！"
      Exit Function
   End If
   m_rs.Close
   
   '先檢查輸入的年月,工作日數有幾天(不含休假日及例假日)
   intWorkDay = 0
   m_str = "SELECT count(*) FROM WORKDAY WHERE substr(WD01,1,6)=" & Left(DBDATE(txt1(0) & "01"), 6)
   intI = 1
   Set m_rs = ClsLawReadRstMsg(intI, m_str)
   If intI = 1 Then
      intWorkDay = Val("" & m_rs.Fields(0).Value)
   End If
   If intWorkDay = 0 Then
      MsgBox "此年月查無工作日數, 資料有誤！", vbInformation, "警示！"
      Exit Function
   End If
   m_rs.Close
   '公司別及男女人數
   GRD1.Clear
   SetGrd
   m_str = "select sm37 公司別,A0802 公司名稱,sum(decode(st22,'M',1,0)) 男生人數,sum(decode(st22,'F',1,0)) 女生人數,0 請假時數,0 加班時數" & _
           " From SalaryMonth,acc080,staff" & _
           " where sm02=" & Left(DBDATE(txt1(0) & "01"), 6) & " and sm37=a0801(+)" & _
           " and sm01=st01(+)" & _
           " and substr(st01,1,1)<>'F' and substr(st01,4,1)<>'9'" & _
           " group by sm37,A0802" & _
           " order by sm37 asc"
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount > 0 Then
      Set GRD1.Recordset = m_rs
   Else
      m_rs.Close
      Set m_rs = Nothing
      ShowNoData
      Exit Function
   End If
   m_rs.Close
   '計算請假時數
   m_str = "select sm37,sum((nvl(sa07,0) * 8)+nvl(sa08,0)) abshour" & _
           " From SalaryMonth,acc080,staff,staff_absence" & _
           " where sm02=" & Left(DBDATE(txt1(0) & "01"), 6) & " and sm37=a0801(+)" & _
           " and sm01=st01(+)" & _
           " and substr(st01,1,1)<>'F' and substr(st01,4,1)<>'9'" & _
           " and sm01=sa01(+)" & _
           " and sa02>=" & Left(DBDATE(txt1(0) & "01"), 6) & "01 and sa04<=" & Left(DBDATE(txt1(0) & "01"), 6) & "31" & _
           " group by sm37" & _
           " order by sm37 asc"
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount > 0 Then
      m_rs.MoveFirst
      Do While Not m_rs.EOF
         For ii = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(ii, 0) = m_rs.Fields(0) Then '公司別
               GRD1.TextMatrix(ii, 4) = m_rs.Fields(1)
               Exit For
            End If
         Next ii
         m_rs.MoveNext
      Loop
   End If
   m_rs.Close
   '計算加班時數
   m_str = "select sm37,sum(nvl(so05,0)+nvl(so06,0)) overtime" & _
           " From SalaryMonth,acc080,staff,staff_overtime" & _
           " where sm02=" & Left(DBDATE(txt1(0) & "01"), 6) & " and sm37=a0801(+)" & _
           " and sm01=st01(+)" & _
           " and substr(st01,1,1)<>'F' and substr(st01,4,1)<>'9'" & _
           " and sm01=so01(+)" & _
           " and so02>=" & Left(DBDATE(txt1(0) & "01"), 6) & "01 and so02<=" & Left(DBDATE(txt1(0) & "01"), 6) & "31" & _
           " group by sm37" & _
           " order by sm37 asc"
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount > 0 Then
      m_rs.MoveFirst
      Do While Not m_rs.EOF
         For ii = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(ii, 0) = m_rs.Fields(0) Then '公司別
               GRD1.TextMatrix(ii, 5) = m_rs.Fields(1)
               Exit For
            End If
         Next ii
         m_rs.MoveNext
      Loop
   End If
   m_rs.Close
   
   '列印出來
   iLine = 4
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "實際工作日數= " & intWorkDay & " 天 "
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   iLine = iLine + 2
   PrintTitle4
   For ii = 1 To GRD1.Rows - 1
      For jj = 1 To 5
         If jj = 1 Then strText = GRD1.TextMatrix(ii, 1) '公司名稱
         If jj = 2 Then strText = GRD1.TextMatrix(ii, 2) '男生
         If jj = 3 Then strText = GRD1.TextMatrix(ii, 3) '女生
         '總計工作日數=(總人數 * 實際工作日數) - (請假日數/8)
         If jj = 4 Then
            strText = ((Val(GRD1.TextMatrix(ii, 2)) + Val(GRD1.TextMatrix(ii, 3))) * intWorkDay) - Int((GRD1.TextMatrix(ii, 4) / 8))
         End If
         '總經歷工時=(總人數 * 實際工作日數 * 8) - 請假時數 + 加班時數
         If jj = 5 Then
            strText = ((Val(GRD1.TextMatrix(ii, 2)) + Val(GRD1.TextMatrix(ii, 3))) * intWorkDay * 8) - GRD1.TextMatrix(ii, 4) + GRD1.TextMatrix(ii, 5)
         End If
         If jj = 4 Or jj = 5 Then
            Printer.CurrentX = PLeft(jj) - Printer.TextWidth(strText)
         Else
            Printer.CurrentX = PLeft(jj)
         End If
         Printer.CurrentY = iLine * 300
         Printer.Print strText
      Next jj
      iLine = iLine + 1
   Next ii
   StrMenu4 = True
End Function

'Add By Sindy 2017/2/2
Sub PrintTitle4()
GetPleft4
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "公司名稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "男生"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "女生"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("總計工作日數")
Printer.CurrentY = iLine * 300
Printer.Print "總計工作日數"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("總經歷工時")
Printer.CurrentY = iLine * 300
Printer.Print "總經歷工時"
iLine = iLine + 1
Call SetLineation
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 1 Then txt1(0).SetFocus
End Sub

'Add By Sindy 2017/1/26
Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

'Add By Sindy 2017/1/26
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

'Add By Sindy 2017/1/26
Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
