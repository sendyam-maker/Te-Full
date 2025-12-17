VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180302 
   BorderStyle     =   1  '單線固定
   Caption         =   "近日請假公佈欄/工作所在地"
   ClientHeight    =   6670
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   6500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6670
   ScaleWidth      =   6500
   Tag             =   "加班資料"
   Begin TabDlg.SSTab SSTab1 
      Height          =   5625
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   6305
      _ExtentX        =   11113
      _ExtentY        =   9931
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "請假公佈欄"
      TabPicture(0)   =   "frm180302.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdPrint"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "GRD1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "工作所在地"
      TabPicture(1)   =   "frm180302.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GRD2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdPrint 
         Caption         =   "列印(&P)"
         Height          =   360
         Left            =   -70140
         TabIndex        =   11
         Top             =   5070
         Width           =   800
      End
      Begin VB.Frame Frame1 
         Caption         =   "設定"
         Height          =   600
         Left            =   -74850
         TabIndex        =   6
         Top             =   4920
         Width           =   4665
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   705
            Style           =   2  '單純下拉式
            TabIndex        =   7
            Top             =   180
            Width           =   3870
         End
         Begin VB.Label Label2 
            Caption         =   "印表機"
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   8
            Top             =   240
            Width           =   765
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm180302.frx":0038
         Height          =   4425
         Left            =   -74850
         TabIndex        =   9
         Top             =   420
         Width           =   6015
         _ExtentX        =   10601
         _ExtentY        =   7814
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
         Bindings        =   "frm180302.frx":004D
         Height          =   5085
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   6015
         _ExtentX        =   10601
         _ExtentY        =   8978
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "日期|部門|員工代號|姓名|地點"
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
         _Band(0).Cols   =   5
      End
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   4
      Top             =   270
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   2
      Top             =   270
      Width           =   945
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label3 
      Caption         =   "同仁同日有請假資料也有特殊工作地的資料，以綠色標註。"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   210
      TabIndex        =   12
      Top             =   6360
      Width           =   6135
   End
   Begin VB.Line Line4 
      X1              =   2130
      X2              =   2670
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "起迄日期："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   330
      Width           =   900
   End
End
Attribute VB_Name = "frm180302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/19 修改抓新部門程式
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/4
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_i As Integer
Dim PLeft(1 To 6) As Integer
Dim strTemp(1 To 6) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblPrevRow As Double, i As Integer
Dim strPrinter As String 'Added by Sindy 2021/1/26
Dim dblPrevRow2 As Double 'Added by Morgan 2021/6/2
Dim arrGridHeadText 'Add by Amy 2022/01/03

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim i As Integer

'Modify By Sindy 2021/1/26
'Set Printer = Printers(Combo1.ListIndex)
PUB_RestorePrinter Combo1
'2021/1/26 END
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

If GRD1.Rows - 1 > 1 Then
   iLine = 1
   strType = ""
   For i = 1 To GRD1.Rows - 1
      'Modify by Amy 2022/01/03 +部門欄
      For m_i = 1 To 6
          strTemp(m_i) = ""
      Next m_i
      
      '流水號
      strTemp(1) = i
      strTemp(2) = GRD1.TextMatrix(i, GetValue("請假日期"))
      strTemp(3) = GRD1.TextMatrix(i, GetValue("假別"))
      'Modify by Amy 2022/01/03 +部門
      strTemp(4) = GRD1.TextMatrix(i, GetValue("部門"))
      strTemp(5) = GRD1.TextMatrix(i, GetValue("員工代號"))
      strTemp(6) = GRD1.TextMatrix(i, GetValue("姓名"))
      'end 2022/01/03
      
      If iLine > 52 Or iLine = 1 Then
         If strType <> "" Then Printer.NewPage
         iLine = 1
         PrintTitle '列印表頭
      End If
      PrintDetail
      strType = GRD1.TextMatrix(i, 0)
   Next i
Else
   ShowNoData
   Exit Sub
End If
Printer.EndDoc
PUB_RestorePrinter strPrinter 'Modify By Sindy 2021/1/26
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1000
PLeft(3) = 2500
PLeft(4) = 3500
PLeft(5) = 5700
PLeft(6) = 6700 'Add by Amy 2021/01/03
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("近日請假名單") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "近日請假名單"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "起迄日期：" & ChangeTStringToTDateString(txt1(0)) & "　" & ChangeTStringToTDateString(txt1(1))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "序號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "請假日期"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "假別"
'Modify by Amy 2021/01/03
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "部門"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "員工代號"
Printer.CurrentX = PLeft(6)
'end 2021/01/03
Printer.CurrentY = iLine * 300
Printer.Print "姓名"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
'Modify by Amy 2022/01/03 +部門
For m_j = 1 To 6
   Printer.CurrentX = PLeft(m_j)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Double, jj As Double
Dim strDate As String
Dim Cancel As Boolean
Dim dblStarDate As Double
Dim dblEndDate As Double
Dim dblStarTime As Double
Dim dblEndTime As Double
Dim dblUpTime As Double
Dim dblDownTime As Double
Dim strB1014 As String
Dim strST14 As String 'Add By Sindy 2024/10/24
   
   dblPrevRow = 0
   GRD1.Clear
   SetGrd
   
   If txt1(0) & txt1(1) <> "" Then
      Cancel = False
      Call txt1_Validate(0, Cancel)
      If Cancel = True Then
         txt1(0).SetFocus
         Exit Sub
      End If
      Call txt1_Validate(1, Cancel)
      If Cancel = True Then
         txt1(1).SetFocus
         Exit Sub
      End If
   Else
      MsgBox "起迄日期不可以空白！", vbExclamation, "操作錯誤！"
      txt1(0).SetFocus
      Exit Sub
   End If
   
   'Add By Sindy 2020/4/13
'   If SSTab1.Tab = 1 Then
      Call QueryGrd2
'      Exit Sub
'   End If
   '2020/4/13 END
   
   txt1(0).Tag = txt1(0).Text
   txt1(1).Tag = txt1(1).Text
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   strSql = ""
   strDate = txt1(0)
   For i = txt1(0) To txt1(1)
      i = strDate
      If ChkWorkDay(DBDATE(strDate)) Then
         If strSql <> "" Then strSql = strSql & "union "
         'Modify By Sindy 2012/1/6 增加顯示主管代填的請假資料
         'Modify By Sindy 2013/5/1 只顯示人事處未先行作業的主管代填資料
         'Modify By Sindy 2013/8/30 +seqno
         'Modify By Sindy 2014/1/10 原程式:出差 decode(sb07,0,' ',decode(sb02,sb04,'SB'," & DBDATE(strDate) & ",'SB',decode(sb04," & DBDATE(strDate) & ",'SB'))) 起迄時間 ==> 改為 'SB' 起迄時間
         'Modify By Sindy 2014/8/28 請依協調會決議將「請假公布欄」之事假 . 病假 . 流產假改以「請假」替代, 可查詢日期不變
         'Modify By Sindy 2014/12/5 +生理假,產檢假,家庭照顧假改以「請假」替代
         'Modify By Sindy 2014/12/31 +健檢假改以「請假」替代
         'Modify By Sindy 2015/11/16 ,decode(sa08,0,' ',decode(sa02,sa04,'SA'," & DBDATE(strDate) & ",'SA',decode(sa04," & DBDATE(strDate) & ",'SA'))) 起迄時間 ==> 光檢查請假日數不代表不是請小時的假, 如:何金柱 104/11/16
         'Modify By Sindy 2017/6/14 +  and st04='1' 只顯示在職人員資料
         'Modify By Sindy 2020/2/5 +防疫照顧假改以「請假」替代
         'Modify By Sindy 2021/3/16 + ||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')'
         'Modify by Amy 2022/01/03 +部門-劉經理
         'Modify By Sindy 2023/12/19 +新部門
         'Modify By Sindy 2024/5/16 原抓 B1018='" & 主管代填 & "' => 再加 B1018 in('" & 會簽職代 & "','"& 主管審核中 &"','"& 重送 &"','"& 主管代填 &"')
         'Modify By Sindy 2024/5/29 要抓未簽准的假單是過去式及當日才顯示
         If strDate <= strSrvDate(2) Then
            strExc(10) = "'" & 會簽職代 & "','" & 主管審核中 & "','" & 重送 & "','" & 主管代填 & "'"
         Else
            strExc(10) = "'" & 主管代填 & "'"
         End If
         '2024/5/29 END
         strSql = strSql & "select '" & ChangeTStringToTDateString(strDate) & "'||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')' 日期,decode(ac03,'扣年終產假','產假','扣年終流產假','請假','事假','請假','病假','請假','流產假','請假','生理假','請假','產檢假','請假','家庭照顧假','請假','健檢假','請假','防疫照顧假','請假',ac03) 假別,nvl(A0922,'(舊)'||A0902) 部門,SA01 員工代號,ST02 員工姓名,decode(sa02,sa04,'SA'," & DBDATE(strDate) & ",'SA',decode(sa04," & DBDATE(strDate) & ",'SA')) 起迄時間,SA09 as seqno,nvl(st93,st03) st93 " & _
                           "From Staff_Absence,Staff,allcode,Acc090NEW,Acc090 Where SA01 = ST01 and " & DBDATE(strDate) & " between SA02 and SA04 and ac01='04' and sa06=ac02(+) and st04='1' and a0921=st93(+) and a0901=st03(+) " & _
                           "union " & _
                           "select '" & ChangeTStringToTDateString(strDate) & "'||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')' 日期,'出差' 假別,nvl(A0922,'(舊)'||A0902) 部門,SB01 員工代號,ST02 員工姓名,'SB' 起迄時間,SB10 as seqno,nvl(st93,st03) st93 " & _
                           "From Staff_Busi_Trip,Staff,Acc090NEW,Acc090 Where SB01 = ST01 and " & DBDATE(strDate) & " between SB02 and SB04 and st04='1' and a0921=st93(+) and a0901=st03(+) " & _
                           "union " & _
                           "select '" & ChangeTStringToTDateString(strDate) & "'||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')' 日期,'請假' 假別,nvl(A0922,'(舊)'||A0902) 部門,B1003 員工代號,ST02 員工姓名,decode(b1004,b1006,'B10'," & DBDATE(strDate) & ",'B10',decode(b1006," & DBDATE(strDate) & ",'B10')) 起迄時間,B1001 as seqno,nvl(st93,st03) st93 " & _
                           "From abs010,Staff,Acc090NEW,Acc090 Where B1003 = ST01 and " & DBDATE(strDate) & " between B1004 and B1006 and B1018 in(" & strExc(10) & ") and st04='1' and a0921=st93(+) and a0901=st03(+) " & _
                           "and not exists (select * from Staff_Absence where sa09=b1001) " & _
                           "and not exists (select * from Staff_Busi_Trip where sb10=b1001) "
         'Modify By Sindy 2013/6/26 增加顯示專利處人員外出記錄
         'Modify By Sindy 2016/11/25 原本程式是Mark起來的,但雅娟要求顯示,已不可考當除為何Mark
         '                討論後覺得當除Mark應該是因非事務所明定的人事規則,所以沒顯示,與劉經理確認過就同上需求顯示出來
         'Modify By Sindy 2023/12/19 +新部門
         strSql = strSql & "union " & _
                           "select '" & ChangeTStringToTDateString(strDate) & "'||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')' 日期,'外出' 假別,nvl(A0922,'(舊)'||A0902) 部門,og03 員工代號,ST02 員工姓名,og19||' ~ '||og20 起迄時間,' ' as seqno,nvl(st93,st03) st93 " & _
                           "From outgoing,Staff,Acc090NEW,Acc090 Where og03 = ST01 and og02=" & DBDATE(strDate) & " and st04='1' and a0921=st93(+) and a0901=st03(+) "
         '2013/6/26 END
      'Add By Sindy 2013/5/17 出差有含非工作日
      Else
         If strSql <> "" Then strSql = strSql & "union "
         'Modify By Sindy 2023/12/19 +新部門
         strSql = strSql & "select '" & ChangeTStringToTDateString(strDate) & "'||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')' 日期,'出差' 假別,nvl(A0922,'(舊)'||A0902) 部門,SB01 員工代號,ST02 員工姓名,' ' 起迄時間,SB10 as seqno,nvl(st93,st03) st93 " & _
                           "From Staff_Busi_Trip,Staff,Acc090NEW,Acc090 Where SB01 = ST01 and " & DBDATE(strDate) & " between SB02 and SB04 and st04='1' and a0921=st93(+) and a0901=st03(+) " & _
                           "union " & _
                           "select '" & ChangeTStringToTDateString(strDate) & "'||' ('||TO_CHAR(TO_DATE('" & DBDATE(strDate) & "','YYYY/MM/DD'),'D')||')' 日期,'請假' 假別,nvl(A0922,'(舊)'||A0902) 部門,B1003 員工代號,ST02 員工姓名,' ' 起迄時間,B1001 as seqno,nvl(st93,st03) st93 " & _
                           "From abs010,Staff,Acc090NEW,Acc090 Where B1003 = ST01 and B1002='" & 表單類別_出差 & "' and " & DBDATE(strDate) & " between B1004 and B1006 and B1018 in('" & 會簽職代 & "','" & 主管審核中 & "','" & 重送 & "','" & 主管代填 & "') and st04='1' and a0921=st93(+) and a0901=st03(+) " & _
                           "and not exists (select * from Staff_Absence where sa09=b1001) " & _
                           "and not exists (select * from Staff_Busi_Trip where sb10=b1001) "
      '2013/5/17 End
      End If
      strDate = ChangeWStringToTString(DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(CStr(i))))))
      If strDate > txt1(1) Then i = strDate
   Next i
   If strSql <> "" Then
      strSql = strSql & "order by 1,2,st93 "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set GRD1.Recordset = rsTmp
         'Add By Sindy 2013/6/26
          'Modify by Amy 2022/01/03 +GetValue
         For i = 1 To GRD1.Rows - 1
            dblStarDate = 0
            dblEndDate = 0
            dblStarTime = 0
            dblEndTime = 0
            dblUpTime = 0
            dblDownTime = 0
            strST14 = "" 'Add By Sindy 2024/10/24
            'Add By Sindy 2022/5/24 若同仁同日有請假資料也有特殊工作地的資料，於請假及工作地都同時加顏色提醒查詢人員。請問加入綠色標註
            For jj = 1 To GRD2.Rows - 1
               If DBDATE(Left(GRD1.TextMatrix(i, 0), 9)) = DBDATE(Left(GRD2.TextMatrix(jj, 0), 9)) And GRD1.TextMatrix(i, 3) = GRD2.TextMatrix(jj, 2) Then
                  GRD1.col = 4
                  GRD1.row = i
                  GRD1.CellBackColor = vbGreen '&HFF00
                  GRD2.col = 3
                  GRD2.row = jj
                  GRD2.CellBackColor = vbGreen
                  Exit For
               End If
            Next jj
            '2022/5/24 END
            'Add By Sindy 2021/3/16 轉換星期幾
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(1)", "(日)")
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(2)", "(一)")
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(3)", "(二)")
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(4)", "(三)")
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(5)", "(四)")
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(6)", "(五)")
            GRD1.TextMatrix(i, GetValue("請假日期")) = Replace(GRD1.TextMatrix(i, GetValue("請假日期")), "(7)", "(六)")
            '2021/3/16 END
            
            If GRD1.TextMatrix(i, GetValue("員工代號")) = "81040" Then
               'MsgBox GRD1.TextMatrix(i, GetValue("員工代號"))
            End If
            
            If GRD1.TextMatrix(i, GetValue("起迄時間")) = "SA" Then
               'Modify By Sindy 2013/8/30
               'Modify By Sindy 2017/6/23 + order by SA02 asc
               'Modify By Sindy 2024/10/24 + 抓ST14
               If GRD1.TextMatrix(i, GetValue("seqno")) <> "" Then
                  strExc(0) = "select * from Staff_Absence,staff where SA01=ST01(+) and SA01='" & GRD1.TextMatrix(i, GetValue("員工代號")) & "' and SA09='" & GRD1.TextMatrix(i, GetValue("seqno")) & "' order by SA02 asc"
               Else
               '2013/8/30 END
                  'Modify By Sindy 2024/10/24 + and SA09 is null
                  strExc(0) = "select * from Staff_Absence,staff where SA01=ST01(+) and SA01='" & GRD1.TextMatrix(i, GetValue("員工代號")) & "' and " & DBDATE(Trim(Mid(GRD1.TextMatrix(i, GetValue("請假日期")), 1, 9))) & " between SA02 and SA04 and SA09 is null order by SA02 asc"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Modify By Sindy 2017/6/23
                  RsTemp.MoveFirst
                  dblStarDate = RsTemp.Fields("SA02")
                  dblStarTime = RsTemp.Fields("SA03")
                  dblEndDate = RsTemp.Fields("SA04")
                  dblEndTime = RsTemp.Fields("SA05")
                  strST14 = "" & RsTemp.Fields("st14") 'Add By Sindy 2024/10/24
                  If RsTemp.RecordCount > 1 Then '多筆時再抓最後一筆假單的截止請假資料
                     RsTemp.MoveLast
                     dblEndDate = RsTemp.Fields("SA04")
                     dblEndTime = RsTemp.Fields("SA05")
'                     strExc(0) = "select * from Staff_Absence where SA01='" & GRD1.TextMatrix(i, 2) & "' and SA02=" & DBDATE(GRD1.TextMatrix(i, 0)) & " and SA04=" & DBDATE(GRD1.TextMatrix(i, 0))
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 0 Then GRD1.TextMatrix(i, 4) = "": GoTo ReadNext
                  End If
                  '2017/6/23 END
                  dblUpTime = Val("" & RsTemp.Fields("SA16"))
                  dblDownTime = Val("" & RsTemp.Fields("SA17"))
               End If
            ElseIf GRD1.TextMatrix(i, GetValue("起迄時間")) = "SB" Then
               'Modify By Sindy 2013/8/30
               'Modify By Sindy 2017/6/23 + order by SB02 asc
               'Modify By Sindy 2024/10/24 + 抓ST14
               If GRD1.TextMatrix(i, GetValue("seqno")) <> "" Then
                  strExc(0) = "select * from Staff_busi_trip,staff where SB01=ST01(+) and SB01='" & GRD1.TextMatrix(i, GetValue("員工代號")) & "' and SB10='" & GRD1.TextMatrix(i, GetValue("seqno")) & "' order by SB02 asc"
               Else
               '2013/8/30 END
                  'Modify By Sindy 2024/10/24 + and SB10 is null
                  strExc(0) = "select * from Staff_busi_trip,staff where SB01=ST01(+) and SB01='" & GRD1.TextMatrix(i, GetValue("員工代號")) & "' and " & DBDATE(Trim(Mid(GRD1.TextMatrix(i, GetValue("請假日期")), 1, 9))) & " between SB02 and SB04 and SB10 is null order by SB02 asc"
               End If
               '2017/6/23 END
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Modify By Sindy 2017/6/23
                  RsTemp.MoveFirst
                  dblStarDate = RsTemp.Fields("SB02")
                  dblStarTime = RsTemp.Fields("SB03")
                  dblEndDate = RsTemp.Fields("SB04")
                  dblEndTime = RsTemp.Fields("SB05")
                  strST14 = "" & RsTemp.Fields("st14") 'Add By Sindy 2024/10/24
                  If RsTemp.RecordCount > 1 Then '多筆時再抓最後一筆假單的截止請假資料
                     RsTemp.MoveLast
                     dblEndDate = RsTemp.Fields("SB04")
                     dblEndTime = RsTemp.Fields("SB05")
'                     strExc(0) = "select * from Staff_busi_trip where SB01='" & GRD1.TextMatrix(i, 2) & "' and SB02=" & DBDATE(GRD1.TextMatrix(i, 0)) & " and SB04=" & DBDATE(GRD1.TextMatrix(i, 0))
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 0 Then GRD1.TextMatrix(i, 4) = "": GoTo ReadNext
                  End If
                  '2017/6/23 END
                  dblUpTime = Val("" & RsTemp.Fields("SB17"))
                  dblDownTime = Val("" & RsTemp.Fields("SB18"))
               End If
            ElseIf GRD1.TextMatrix(i, GetValue("起迄時間")) = "B10" Then
               'Modify By Sindy 2013/8/30
               'Modify By Sindy 2017/6/23 + order by B1004 asc
               'Modify By Sindy 2024/10/24 mark: ABS010一定會有序號
'               If GRD1.TextMatrix(i, GetValue("seqno")) <> "" Then
                  strExc(0) = "select * from ABS010 where B1003='" & GRD1.TextMatrix(i, GetValue("員工代號")) & "' and B1001='" & GRD1.TextMatrix(i, GetValue("seqno")) & "' order by B1004 asc"
'               Else
'               '2013/8/30 END
'                  strExc(0) = "select * from ABS010 where B1003='" & GRD1.TextMatrix(i, GetValue("員工代號")) & "' and " & DBDATE(Trim(Mid(GRD1.TextMatrix(i, GetValue("請假日期")), 1, 9))) & " between B1004 and B1006 order by B1004 asc"
'               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Modify By Sindy 2017/6/23
                  RsTemp.MoveFirst
                  dblStarDate = RsTemp.Fields("B1004")
                  dblStarTime = RsTemp.Fields("B1005")
                  dblEndDate = RsTemp.Fields("B1006")
                  dblEndTime = RsTemp.Fields("B1007")
                  If RsTemp.RecordCount > 1 Then '多筆時再抓最後一筆假單的截止請假資料
                     RsTemp.MoveLast
                     dblEndDate = RsTemp.Fields("B1006")
                     dblEndTime = RsTemp.Fields("B1007")
'                     strExc(0) = "select * from ABS010 where B1003='" & GRD1.TextMatrix(i, 2) & "' and B1004=" & DBDATE(GRD1.TextMatrix(i, 0)) & " and B1006=" & DBDATE(GRD1.TextMatrix(i, 0))
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 0 Then GRD1.TextMatrix(i, 4) = "": GoTo ReadNext
                  End If
                  '2017/6/23 END
                  dblUpTime = Val("" & RsTemp.Fields("B1028"))
                  dblDownTime = Val("" & RsTemp.Fields("B1029"))
               End If
            End If
ReadNext:
            If dblStarDate > 0 Then
               '起迄日同一天
               If dblStarDate = dblEndDate Then
                  'Modify By Sindy 2014/1/10
                  If dblStarTime <= 900 And dblEndTime >= 1700 Then
                     GRD1.TextMatrix(i, GetValue("起迄時間")) = ""
                  Else
                  '2014/1/10 END
                     GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & Format(dblEndTime, "00:00")
                  End If
               Else
                  '第一天
                  If DBDATE(Trim(Mid(GRD1.TextMatrix(i, GetValue("請假日期")), 1, 9))) = dblStarDate Then
                     If dblStarTime > 900 Then
                        'Modify By Sindy 2016/5/12 跨天出差大陸或國外時第1天的結束是24:00
'                        strB1014 = ""
                        If GRD1.TextMatrix(i, GetValue("假別")) = "出差" Then
'                           strExc(0) = "select sb08 from staff_busi_trip where sB10='" & GRD1.TextMatrix(i, 5) & "'"
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              strB1014 = "" & RsTemp.Fields("sb08")
'                           Else
'                              strExc(0) = "select b1014 from ABS010 where B1001='" & GRD1.TextMatrix(i, 5) & "'"
'                              intI = 1
'                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                              If intI = 1 Then
'                                 strB1014 = "" & RsTemp.Fields("b1014")
'                              End If
'                           End If
'                        End If
'                        If GRD1.TextMatrix(i, 1) = "出差" And (strB1014 = "3" Or strB1014 = "4") Then
                           GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "24:00"
                        Else
                        '2016/5/12 END
                           If dblUpTime > 0 Then
                              If dblUpTime = 800 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "17:00"
                              ElseIf dblUpTime = 830 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "17:30"
                              ElseIf dblUpTime = 900 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "18:00"
                              'Add By Sindy 2021/8/13
                              ElseIf dblUpTime = 730 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "16:30"
                              ElseIf dblUpTime = 930 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "18:30"
                              ElseIf dblUpTime = 1000 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "19:00"
                              '2021/8/13 END
                              End If
                           Else
                              'Add By Sindy 2024/10/24
                              If strST14 = "99997" Then 'Add By Sindy 2024/10/24
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & Format(dblEndTime, "00:00")
                              Else
                              '2024/10/24 END
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & "17:00"
                              End If
                           End If
                        End If
                     Else
                        GRD1.TextMatrix(i, GetValue("起迄時間")) = ""
                     End If
                  '最後一天
                  ElseIf DBDATE(Trim(Mid(GRD1.TextMatrix(i, GetValue("請假日期")), 1, 9))) = dblEndDate Then
                     If dblEndTime < 1700 Then
                        'Modify By Sindy 2016/5/12 跨天出差大陸或國外時最後一天的起始是00:00
'                        strB1014 = ""
                        If GRD1.TextMatrix(i, GetValue("假別")) = "出差" Then
'                           strExc(0) = "select sb08 from staff_busi_trip where sB10='" & GRD1.TextMatrix(i, 5) & "'"
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              strB1014 = "" & RsTemp.Fields("sb08")
'                           Else
'                              strExc(0) = "select b1014 from ABS010 where B1001='" & GRD1.TextMatrix(i, 5) & "'"
'                              intI = 1
'                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                              If intI = 1 Then
'                                 strB1014 = "" & RsTemp.Fields("b1014")
'                              End If
'                           End If
'                        End If
'                        If GRD1.TextMatrix(i, 1) = "出差" And (strB1014 = "3" Or strB1014 = "4") Then
                           GRD1.TextMatrix(i, GetValue("起迄時間")) = "00:00" & " ~ " & Format(dblEndTime, "00:00")
                        Else
                        '2016/5/12 END
                           If dblDownTime > 0 Then
                              If dblDownTime = 1700 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "08:00" & " ~ " & Format(dblEndTime, "00:00")
                              ElseIf dblDownTime = 1730 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "08:30" & " ~ " & Format(dblEndTime, "00:00")
                              ElseIf dblDownTime = 1800 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "09:00" & " ~ " & Format(dblEndTime, "00:00")
                              'Add By Sindy 2021/8/13
                              ElseIf dblDownTime = 1630 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "07:30" & " ~ " & Format(dblEndTime, "00:00")
                              ElseIf dblDownTime = 1830 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "09:30" & " ~ " & Format(dblEndTime, "00:00")
                              ElseIf dblDownTime = 1900 Then
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "10:00" & " ~ " & Format(dblEndTime, "00:00")
                              '2021/8/13 END
                              End If
                           Else
                              'Add By Sindy 2024/10/24
                              If strST14 = "99997" Then 'Add By Sindy 2024/10/24
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = Format(dblStarTime, "00:00") & " ~ " & Format(dblEndTime, "00:00")
                              Else
                              '2024/10/24 END
                                 GRD1.TextMatrix(i, GetValue("起迄時間")) = "08:00" & " ~ " & Format(dblEndTime, "00:00")
                              End If
                           End If
                        End If
                     Else
                        GRD1.TextMatrix(i, GetValue("起迄時間")) = ""
                     End If
                  'Modify By Sindy 2014/1/10
                  '中間天數
                  Else
                     GRD1.TextMatrix(i, GetValue("起迄時間")) = ""
                  '2014/1/10 END
                  End If
               End If
            End If
         Next i
         '2013/6/26 END
      Else
         Screen.MousePointer = vbDefault
         'ShowNoData 'Removed by Morgan 2023/7/3 無資料不必彈訊息 --經理
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Sub
      End If
   Else
      Screen.MousePointer = vbDefault
      'ShowNoData 'Removed by Morgan 2023/7/3 無資料不必彈訊息 --經理
      'rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2020/4/13
Private Function QueryGrd2() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Double
Dim strDate As String
Dim Cancel As Boolean
Dim dblStarDate As Double
Dim dblEndDate As Double
Dim dblStarTime As Double
Dim dblEndTime As Double
Dim dblUpTime As Double
Dim dblDownTime As Double
Dim strB1014 As String
   
   'Modified by Morgan 2021/6/2
   'dblPrevRow = 0
   dblPrevRow2 = 0
   'end 2021/6/2
   
   GRD2.Clear
   SetGrd2
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass

   If txt1(0) <> "" Then
       strSql = strSql & " and SP01>=" & DBDATE(txt1(0))
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and SP01<=" & DBDATE(txt1(1))
   End If
   'Modify By Sindy 2021/3/16 + ||' ('||TO_CHAR(TO_DATE(SP01,'YYYY/MM/DD'),'D')||')'
   'Modify By Sindy 2023/12/19 +新部門
   strSql = "SELECT sqldateT(SP01)||' ('||TO_CHAR(TO_DATE(SP01,'YYYY/MM/DD'),'D')||')' 日期,nvl(A0922,'(舊)'||A0902) 部門,SP02 員工代號,st02 姓名," & SP03WorkPlace & " 地點,nvl(st93,st03) st93" & _
            " From STAFF_WORKPLACE, staff, acc090NEW, acc090" & _
            " where SP02=st01(+) and A0921(+)=st93 and A0901(+)=st03" & strSql & _
            " order by SP01,st93,SP02"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryGrd2 = True
      Set GRD2.Recordset = rsTmp
      'Add By Sindy 2021/3/16 轉換星期幾
      For i = 1 To GRD2.Rows - 1
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(1)", "(日)")
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(2)", "(一)")
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(3)", "(二)")
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(4)", "(三)")
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(5)", "(四)")
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(6)", "(五)")
         GRD2.TextMatrix(i, 0) = Replace(GRD2.TextMatrix(i, 0), "(7)", "(六)")
      Next i
      '2021/3/16 END
   Else
      Screen.MousePointer = vbDefault
      QueryGrd2 = False
      'ShowNoData 'Removed by Morgan 2023/7/3 無資料不必彈訊息 --經理
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
   '若有資料游標停在第一筆
   GRD2.Visible = False
   GRD2.col = 0
   GRD2.row = 1
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD2.Cols - 1
'         GRD2.col = i
'         GRD2.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD2.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer

   MoveFormToCenter Me
   
   'Modify By Sindy 2021/1/26
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
'   Set Printer = Printers(SeekPrint)
'   Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   '2021/1/26 END
   
   txt1(0).Text = strSrvDate(2) '當天
   txt1(1).Text = ChangeWStringToTString(CompWorkDay(2, strSrvDate(1))) '隔一天
   
   If Pub_StrUserSt03 = "M21" Or Pub_StrUserSt03 = "M51" Then
      cmdPrint.Visible = True
      Frame1.Visible = True
   Else
      cmdPrint.Visible = False
      Frame1.Visible = False
   End If
   
   cmdQuery_Click
   
   SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180302 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadWidth 'Mark by Amy 2022/01/03 arrGridHeadText
   Dim iRow As Integer
   
   'Modify By Sindy 2013/8/30 +seqno
   'Modify By Sindy 2024/1/11 +st93
   arrGridHeadText = Array("請假日期", "假別", "部門", "員工代號", "姓名", "起迄時間", "seqno", "st93")
   arrGridHeadWidth = Array(1200, 800, 1000, 800, 900, 1200, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'Add By Sindy 2020/4/13
Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2024/1/11 +st93
   arrGridHeadText = Array("日期", "部門", "員工代號", "姓名", "地點", "st93")
   arrGridHeadWidth = Array(1200, 1000, 800, 800, 1200, 0)
   GRD2.Visible = False
   GRD2.Cols = UBound(arrGridHeadText) + 1
   GRD2.Rows = 2
   For iRow = 0 To GRD2.Cols - 1
      GRD2.row = 0
      GRD2.col = iRow
      GRD2.Text = arrGridHeadText(iRow)
      GRD2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD2.CellAlignment = flexAlignCenterCenter
   Next
   GRD2.Visible = True
End Sub

'Add By Sindy 2020/4/13
Private Sub SSTab1_Click(PreviousTab As Integer)
'   If SSTab1.Tab = 0 Then
'      If (txt1(0).Tag <> txt1(0).Text Or _
'         txt1(1).Tag <> txt1(1).Text) And _
'         txt1(0).Tag <> "" And txt1(1).Tag <> "" Then
'         cmdQuery_Click
'      End If
'   Else
'      QueryGrd2
'   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
            If Val(txt1(Index)) > Val(txt1(Index + 1)) Then
               txt1(Index + 1) = txt1(Index)
            End If
            'Modify By Sindy 2012/8/3 開放電腦中心及人事處不須控管此條件
            '因如8/2蘇拉颱風假,人事處在8/3回來上班時需查詢8/2有多少人請假
            'Modify By Sindy 2021/7/6 開放71011王錦寬副總可往前查詢
            If Val(txt1(Index)) < Val(strSrvDate(2)) And _
               Pub_StrUserSt03 <> "M21" And _
               Pub_StrUserSt03 <> "M51" And _
               strUserNum <> "71011" Then
            '2012/8/3 End
               Call txt1_GotFocus(Index)
               Cancel = True
               MsgBox "起始日期不可小於系統日！", vbInformation, "輸入日期錯誤"
               Exit Sub
            End If
         ElseIf Index = 1 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblPrevRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         'Add By Sindy 2022/6/17 姓名有變色功能,在此均不變動顏色
         If i <> 4 Then
         '2022/6/17 END
            GRD1.CellBackColor = QBColor(15)
         End If
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
   For i = 0 To GRD1.Cols - 1
      GRD1.col = i
      'Add By Sindy 2022/6/17 姓名有變色功能,在此均不變動顏色
      If i <> 4 Then
      '2022/6/17 END
         GRD1.CellBackColor = &HFFC0C0
      End If
   Next i
End If
GRD1.Visible = True
End Sub

'請假公佈欄
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'Add By Sindy 2024/1/11
   If nCol = 2 Then '部門中文
      GRD1.col = 7 '部門代碼
   Else
   '2024/1/11 END
      GRD1.col = nCol
   End If
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
'      If Me.GRD1.Text = "日期" Then
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
'      End If
   End If
End Sub

Private Sub GRD2_SelChange()
GRD2.Visible = False
If GRD2.MouseRow <> 0 Then
   '上一筆資料列清除反白
   'Modified by Morgan 2021/6/2 dblPrevRow->dblPrevRow2
   If dblPrevRow2 > 0 Then
      GRD2.col = 2
      GRD2.row = dblPrevRow2
      For i = 0 To GRD2.Cols - 1
         GRD2.col = i
         'Add By Sindy 2022/6/17 姓名有變色功能,在此均不變動顏色
         If i <> 3 Then
         '2022/6/17 END
            GRD2.CellBackColor = QBColor(15)
         End If
      Next i
   End If
   '目前資料列反白
   GRD2.col = 0
   GRD2.row = GRD2.MouseRow
   dblPrevRow2 = GRD2.row
   For i = 0 To GRD2.Cols - 1
      GRD2.col = i
      'Add By Sindy 2022/6/17 姓名有變色功能,在此均不變動顏色
      If i <> 3 Then
      '2022/6/17 END
         GRD2.CellBackColor = &HFFC0C0
      End If
   Next i
End If
GRD2.Visible = True
End Sub

'工作所在地
Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD2, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'Add By Sindy 2024/1/11
   If nCol = 1 Then '部門中文
      GRD2.col = 5 '部門代碼
   Else
   '2024/1/11 END
      GRD2.col = nCol
   End If
   GRD2.row = nRow
   If Me.GRD2.row < 1 And Me.GRD2.Text <> "V" Then
'      If Me.GRD2.Text = "日期" Then
'         If m_blnColOrderAsc = True Then
'            Me.GRD2.Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD2.Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
         If m_blnColOrderAsc = True Then
            Me.GRD2.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD2.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
'      End If
   End If
End Sub

'Add by Amy 2022/01/03
Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(arrGridHeadText)
       If UCase(arrGridHeadText(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
