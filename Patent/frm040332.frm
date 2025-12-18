VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040332 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部FC案件不請款清單"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8445
   Begin VB.CheckBox Check2 
      Caption         =   "含已確認不請款"
      Height          =   225
      Left            =   4380
      TabIndex        =   20
      Top             =   1020
      Width           =   2000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FMP"
      Height          =   225
      Index           =   5
      Left            =   7080
      TabIndex        =   19
      Top             =   1260
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FMT"
      Height          =   225
      Index           =   4
      Left            =   5880
      TabIndex        =   18
      Top             =   1260
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   0
      Top             =   690
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3180
      MaxLength       =   7
      TabIndex        =   1
      Top             =   690
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   90
      TabIndex        =   12
      Top             =   30
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4035
      Left            =   60
      TabIndex        =   10
      Top             =   1530
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
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
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   2
      Left            =   7425
      TabIndex        =   9
      Top             =   90
      Width           =   885
   End
   Begin VB.CheckBox Check1 
      Caption         =   "S(台灣案)"
      Height          =   225
      Index           =   3
      Left            =   4380
      TabIndex        =   5
      Top             =   1260
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FG"
      Height          =   225
      Index           =   2
      Left            =   2955
      TabIndex        =   4
      Top             =   1260
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FCP"
      Height          =   225
      Index           =   1
      Left            =   1530
      TabIndex        =   3
      Top             =   1260
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FCT"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   1260
      Value           =   1  '核取
      Width           =   1425
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   405
      Index           =   1
      Left            =   6502
      TabIndex        =   8
      Top             =   90
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5580
      TabIndex        =   7
      Top             =   90
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   810
      Left            =   9435
      TabIndex        =   11
      Top             =   2325
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1429
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2940
      X2              =   3180
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   3180
      TabIndex        =   17
      Top             =   1020
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   1710
      TabIndex        =   16
      Top             =   1020
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "上次列印發文日："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   15
      Top             =   1020
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "發文日期："
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   14
      Top             =   720
      Width           =   915
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2940
      X2              =   3135
      Y1              =   810
      Y2              =   810
   End
End
Attribute VB_Name = "frm040332"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; grd1改字型=新細明體-ExtB; 尚未修改列印
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'列印控制
Dim PLeft(0 To 10) As Integer 'Modify by Amy 2016/12/07 +進度備註
Dim strTemp(0 To 8) As String
Dim iPrint As Integer
Dim Page As Integer
Dim SeekTemp1 As String, SeekTemp2 As String, SeekPrint As Integer, SeekPrintL As Integer, SeekTempPrint As String
Dim i As Integer, j As Integer, strTemp1 As Variant, strTemp2 As Variant, s As Integer


Sub SetGrid()
   With GRD1
      .Cols = 12 'Modify by Amy +顯示承辦人,2016/12/6再加進度備註 原10,
      .row = 0
      .col = 0: .Text = "本所案號"
      .ColWidth(0) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .Text = "收文日"
      .ColWidth(1) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .Text = "案件性質"
      .ColWidth(2) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .Text = "費用"
      .ColWidth(3) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .Text = "規費"
      .ColWidth(4) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .Text = "點數"
      .ColWidth(5) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .Text = "發文規費"
      .ColWidth(6) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .Text = "總收文號"
      .ColWidth(7) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .Text = "發文日"
      .ColWidth(8) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .Text = "系統別"
      .ColWidth(9) = 0
      .CellAlignment = flexAlignCenterCenter
      'Add by Amy 2013/09/05 +顯示承辦人
      .col = 10: .Text = "承辦人"
      .ColWidth(10) = 1000
      .CellAlignment = flexAlignCenterCenter
      'end 2013/09/05
      'add by sonia 2016/12/6 +顯示進度備註
      .col = 11: .Text = "進度備註"
      .ColWidth(11) = 2000
      .CellAlignment = flexAlignCenterCenter
      'end 2016/12/6
   End With
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strSql As String, strCP01 As String, strCP0110 As String
Dim strCP64 As String 'Add by Amy 2016/11/07

If Index = 0 Or Index = 1 Then
   If IsEmptyText(txt1(0)) = True Then
      txt1(0).SetFocus
      MsgBox "發文起始日期不可空白！", vbCritical, "錯誤！"
      Exit Sub
   End If
   If IsEmptyText(txt1(1)) = True Then
      txt1(1).SetFocus
      MsgBox "發文截止日期不可空白！", vbCritical, "錯誤！"
      Exit Sub
   End If
   If Val(txt1(0)) > Val(txt1(1)) Then
      txt1(0).SetFocus
      MsgBox "發文起始日期不可大於發文截止日期！", vbCritical, "錯誤！"
      Exit Sub
   End If
   'Modify by Amy 2016/08/23 +FMT/FMP
   If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And _
      Check1(3).Value = 0 And Check1(4).Value = 0 And Check1(5).Value = 0 Then
      MsgBox "請最少勾選一種系統別！", vbCritical, "錯誤！"
      Exit Sub
   End If
   
   'FCT、FCP、FG
   'Modify by Amy 2016/08/23 +FMT/FMP
   If Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1 _
     Or Check1(4).Value = 1 Or Check1(5).Value = 1 Then
      strCP01 = "" '系統別
      strCP0110 = "" '系統別+案件性質(延期+補收款)
      If Check1(0).Value = 1 Then
         If strCP01 <> "" Then strCP01 = strCP01 & ","
         If strCP0110 <> "" Then strCP0110 = strCP0110 & ","
         strCP01 = strCP01 & "'FCT'"
         strCP0110 = strCP0110 & "'FCT303','FCT705'"
      End If
      If Check1(1).Value = 1 Then
         If strCP01 <> "" Then strCP01 = strCP01 & ","
         If strCP0110 <> "" Then strCP0110 = strCP0110 & ","
         strCP01 = strCP01 & "'FCP'"
         strCP0110 = strCP0110 & "'FCP404','FCP911'"
      End If
      If Check1(2).Value = 1 Then
         If strCP01 <> "" Then strCP01 = strCP01 & ","
         If strCP0110 <> "" Then strCP0110 = strCP0110 & ","
         strCP01 = strCP01 & "'FG'"
         strCP0110 = strCP0110 & "'FG911'"
      End If
      'FMT(T非台灣) 'Add by Amy 2016/08/23
      If Check1(4).Value = 1 Then
         If strCP01 <> "" Then strCP01 = strCP01 & ","
         If strCP0110 <> "" Then strCP0110 = strCP0110 & ","
         strCP01 = strCP01 & "'T'"
         strCP0110 = strCP0110 & "'T303','T705'"
      End If
      'FMP (P非台灣) 'Add by Amy 2016/08/23
      If Check1(5).Value = 1 Then
         If strCP01 <> "" Then strCP01 = strCP01 & ","
         If strCP0110 <> "" Then strCP0110 = strCP0110 & ","
         strCP01 = strCP01 & "'P'"
         strCP0110 = strCP0110 & "'P404','P911'"
      End If
      'Add by Amy 2016/11/07 +未勾選「含已確認不請款」
      'modify by sonia 2016/12/6 加入cp64 is null or ,否則cp64為空者不會出來
      If Check2.Value = 0 Then
         strCP64 = " And (CP64 IS NULL OR (InStr(CP64,'取消費用')=0 And InStr(CP64,'及不請款')=0)) "
      End If
      
      'Modify by Amy +顯示承辦人join staff
      '指該收文號不請款但有發文規費...cp20||''<==加||''是為了影響index讓查詢速度變快
      'Modify by Amy 2016/08/23 FMT,FMP之檢查 CP84>0改為CP16>0,案件性質原只抓cpm03
      'Modify by Amy 2016/11/07 +strCP64
      'modify by sonia 2016/12/6 加未收費但有帳單cp61的控制Decode(cp01,'P',cp16,Decode(cp01,'T',cp16,cp84))>0改為((Decode(cp01,'P',cp16,Decode(cp01,'T',cp16,cp84))>0) or cp61 is not null)
      '                          +cp64,cp84改為decode(cp84,null,cp61,'0',cp61,cp84)
      'Modify by Amy 2016/12/07 +已做不請款確認顯示*
      'modify by sonia 2016/12/9 取消and (cp18>=0 or cp18 is null)控制 FCP-044942
      strSql = " select Decode(sign(InStr(CP64,'取消費用')),1,Decode(sign(InStr(CP64,'及不請款')),1,'*',''),'')||cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNO,(CP05-19110000),SUBSTR(Decode(cp01,'P',cpm04,Decode(cp01,'T',cpm04,cpm03)),1,6),cp16,cp17,cp18,decode(cp84,null,cp61,'0',cp61,cp84),CP09,(CP27-19110000),cp01,st02,cp64" & _
                     " From caseprogress a, casepropertymap,staff" & _
                     " where cp01 in(" & strCP01 & ")" & _
                     " and CP27>=" & (Val(txt1(0)) + 19110000) & " AND CP27<=" & (Val(txt1(1)) + 19110000) & _
                     " and cp01||cp10 not in(" & strCP0110 & ")" & _
                     " and SUBSTR(cp12,1,1)='F'" & _
                     " and cp20||''='N'" & _
                     " and ((Decode(cp01,'P',cp16,Decode(cp01,'T',cp16,cp84))>0) or cp61 is not null)" & strCP64 & _
                     " and CP60 IS NULL" & _
                     " and CP01=CPM01(+) AND CP10=CPM02(+) And cp14=st01(+)"
      
      'Add by Morgan 2010/6/3 排除FCP申請案有收文工程師提申
      'Modify by Amy 2016/08/23 +排除P申請案有收文工程師提申
      'strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10='940' and b.cp57 is null and a.cp01='FCP' and a.cp10 in ('101','102','103','105'))"
       strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10='940' and b.cp57 is null and a.cp01 in ('FCP','P') and a.cp10 in ('101','102','103','105'))"

      '指延期案件性質有發文規費但其相關總收文號不請款
      '2013/8/20 modify by sonia 剔除財務已做過不請款確認(frm040333)的資料,故加入and instr(c1.cp64,'取消費用')=0 and instr(c1.cp64,'及不請款')=0
      'Modify by Amy 2016/08/23 FMT,FMP之檢查 CP84>0改為CP16>0,案件性質原只抓cpm03
      'modify by sonia 2016/12/6 加入cp64 is null or ,否則cp64為空者不會出來
      'modify by sonia 2016/12/6 加未收費但有帳單cp61的控制Decode(C1.cp01,'P',C1.cp16,Decode(C1.cp01,'T',C1.cp16,C1.cp84))>0改為(Decode(C1.cp01,'P',C1.cp16,Decode(C1.cp01,'T',C1.cp16,C1.cp84))>0 or c1.cp61 is not null)
      '                          +C1.cp64,C1.cp84改為decode(C1.cp84,null,C1.cp61,'0',C1.cp61,C1.cp84)
      'Modify by Amy 2016/12/07 +已做不請款確認顯示*
      strSql = strSql & " Union All" & _
                     " select Decode(sign(InStr(C1.CP64,'取消費用')),1,Decode(sign(InStr(C1.CP64,'及不請款')),1,'*',''),'')||C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 CaseNO,(C1.CP05-19110000),SUBSTR(Decode(C1.cp01,'P',cpm04,Decode(C1.cp01,'T',cpm04,cpm03)),1,6),C1.cp16,C1.cp17,C1.cp18,decode(C1.cp84,null,C1.cp61,'0',C1.cp61,C1.cp84),C1.CP09,(C1.CP27-19110000),C1.cp01,st02,C1.cp64" & _
                     " from caseprogress C1,caseprogress C2,casepropertymap,staff" & _
                     " where C1.cp01||C1.cp10 in(" & strCP0110 & ")" & _
                     " and C1.CP27>=" & (Val(txt1(0)) + 19110000) & " AND C1.CP27<=" & (Val(txt1(1)) + 19110000) & _
                     " and SUBSTR(C1.cp12,1,1)='F'" & _
                     " and (Decode(C1.cp01,'P',C1.cp16,Decode(C1.cp01,'T',C1.cp16,C1.cp84))>0 or c1.cp61 is not null) and (c1.cp64 is null or (instr(c1.cp64,'取消費用')=0 and instr(c1.cp64,'及不請款')=0))" & _
                     " and C1.CP60 IS NULL" & _
                     " and C1.CP43=C2.CP09(+)" & _
                     " and C2.CP60 IS NULL" & _
                     " and C2.cp20='N'" & _
                     " and (C2.cp18>=0 or C2.cp18 is null)" & _
                     " and C1.CP01=CPM01(+) AND C1.CP10=CPM02(+) And C1.cp14=st01(+)"
   End If
   'S(台灣案)
   If Check1(3).Value = 1 Then
      If Trim(strSql) <> "" Then
         strSql = strSql & " Union All"
      End If
      '指該收文號不請款但有發文規費...cp20||''<==加||''是為了影響index讓查詢速度變快
      'modify by sonia 2016/12/6 +cp64,cp84改為decode(cp84,null,cp61,'0',cp61,cp84)
      'modify by sonia 2016/12/9 取消and (cp18>=0 or cp18 is null)控制 FCP-044942
      strSql = strSql & _
                     " select Decode(sign(InStr(CP64,'取消費用')),1,Decode(sign(InStr(CP64,'及不請款')),1,'*',''),'')||cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNO,(CP05-19110000),SUBSTR(CPM03,1,6),cp16,cp17,cp18,decode(cp84,null,cp61,'0',cp61,cp84),CP09,(CP27-19110000),cp01,st02,cp64" & _
                     " From caseprogress, servicepractice, casepropertymap,staff" & _
                     " where cp01='S'" & _
                     " and CP27>=" & (Val(txt1(0)) + 19110000) & " AND CP27<=" & (Val(txt1(1)) + 19110000) & _
                     " and cp01||cp10 not in('S705')" & _
                     " and SUBSTR(cp12,1,1)='F'" & _
                     " and cp20||''='N'" & _
                     " and cp84>0" & strCP64 & _
                     " and CP60 IS NULL" & _
                     " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
                     " and sp09='000'" & _
                     " and CP01=CPM01(+) AND CP10=CPM02(+) And cp14=st01(+)"
      '指延期案件性質有發文規費但其相關總收文號不請款
      '2013/8/20 modify by sonia 剔除財務已做過不請款確認(frm040333)的資料,故加入and instr(c1.cp64,'取消費用')=0 and instr(c1.cp64,'及不請款')=0
      'modify by sonia 2016/12/6 加入cp64 is null or ,否則cp64為空者不會出來,+cp64欄,C1.cp84改為decode(C1.cp84,null,C1.cp61,'0',C1.cp61,C1.cp84)
      strSql = strSql & " Union All" & _
                     " select C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 CaseNO,(C1.CP05-19110000),SUBSTR(CPM03,1,6),C1.cp16,C1.cp17,C1.cp18,decode(C1.cp84,null,C1.cp61,'0',C1.cp61,C1.cp84),C1.CP09,(C1.CP27-19110000),C1.cp01,st02,C1.cp64" & _
                     " from caseprogress C1,caseprogress C2, servicepractice, casepropertymap,staff" & _
                     " where C1.cp01||C1.cp10 in('S705')" & _
                     " and C1.CP27>=" & (Val(txt1(0)) + 19110000) & " AND C1.CP27<=" & (Val(txt1(1)) + 19110000) & _
                     " and SUBSTR(C1.cp12,1,1)='F'" & _
                     " and C1.cp84>0 and (c1.cp64 is null or (instr(c1.cp64,'取消費用')=0 and instr(c1.cp64,'及不請款')=0))" & _
                     " and C1.CP60 IS NULL" & _
                     " and C1.CP43=C2.CP09(+)" & _
                     " and C2.CP60 IS NULL" & _
                     " and C2.cp20='N'" & _
                     " and (C2.cp18>=0 or C2.cp18 is null)" & _
                     " and C1.cp01=sp01(+) and C1.cp02=sp02(+) and C1.cp03=sp03(+) and C1.cp04=sp04(+)" & _
                     " and sp09='000'" & _
                     " and C1.CP01=CPM01(+) AND C1.CP10=CPM02(+) And C1.cp14=st01(+)"
   End If
   strSql = strSql & " order by cp01,CaseNO"
End If

Select Case Index
Case 0    '查詢
         GRD1.Clear
         GRD1.Rows = 2
         SetGrid
         Screen.MousePointer = vbHourglass
         GRD1.MousePointer = flexArrowHourGlass
         CheckOC3
         With AdoRecordSet3
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount <> 0 Then
                  cmdOK(1).Enabled = True
               Else
                  ShowNoData
                  cmdOK(1).Enabled = False
               End If
               Set GRD1.Recordset = AdoRecordSet3
               SetGrid
               Screen.MousePointer = vbDefault
               GRD1.MousePointer = flexDefault
         End With
         CheckOC3
Case 1    '列印
         If Combo1.ListIndex >= SeekPrint Then
            j = Combo1.ListIndex + 1
         Else
            j = Combo1.ListIndex
         End If
         Set Printer = Printers(j)
         Screen.MousePointer = vbHourglass
         'Modify by Amy 2013/09/05 +if 判斷此次的發文迄日>上次列印發文迄日才記錄(記錄最大值)
         If Val(Label2(1).Caption) < Val(txt1(1).Text) Then
            'Modify by Amy 2014/07/14
'            SaveSetting "TAIE", "FCP", "DATE81", txt1(0).Text
'            SaveSetting "TAIE", "FCP", "DATE82", txt1(1).Text
            PUB_SaveLastDate Me.Name, "txt1(0)", txt1(0)
            PUB_SaveLastDate Me.Name, "txt1(1)", txt1(1)
            'end 2014/07/14
         End If
        'end 2013/09/05
        
'         Grd1.MousePointer = flexArrowHourGlass
'         CheckOC3
'         With AdoRecordSet3
'               .CursorLocation = adUseClient
'               .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'               If .RecordCount <> 0 Then
'                     Set grd2.Recordset = AdoRecordSet3
                     PrintData
'               End If
               'Modify by Amy 2014/07/14
'               Label2(0).Caption = GetSetting("TAIE", "FCP", "DATE81", "")
'               Label2(1).Caption = GetSetting("TAIE", "FCP", "DATE82", "")
               If PUB_GetLastDate(Me.Name, "txt1(0)") <> "" Then
                    Label2(0).Caption = PUB_GetLastDate(Me.Name, "txt1(0)")
               End If
               If PUB_GetLastDate(Me.Name, "txt1(1)") <> "" Then
                    Label2(1).Caption = PUB_GetLastDate(Me.Name, "txt1(1)")
               End If
               Screen.MousePointer = vbDefault
               'end 2014/07/14
'               Grd1.MousePointer = flexDefault
'         End With
         CheckOC3
Case 2
         Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Modify by Amy 2014/07/14 第一次DB可能沒資料所以先抓client
   If PUB_GetLastDate(Me.Name, "txt1(0)") <> "" Then
       Label2(0).Caption = PUB_GetLastDate(Me.Name, "txt1(0)")
   Else
       Label2(0).Caption = GetSetting("TAIE", "FCP", "DATE81", "")
   End If
   If PUB_GetLastDate(Me.Name, "txt1(1)") <> "" Then
       Label2(1).Caption = PUB_GetLastDate(Me.Name, "txt1(1)")
   Else
       Label2(1).Caption = GetSetting("TAIE", "FCP", "DATE82", "")
   End If
   'end 2014/07/14
   
   cmdOK(1).Enabled = False
   '先將資料放在暫存
   Screen.MousePointer = vbHourglass
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
       Set Printer = Printers(i)
       If Printer.DeviceName <> strSql Then
           Combo1.AddItem Printer.DeviceName, j
           j = j + 1
       End If
       If Printer.DeviceName = strSql Then
           SeekPrint = i
       End If
   Next i
   Combo1.Text = Combo1.List(0)
   DoEvents
   Screen.MousePointer = vbDefault
   SetGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Set frm040332 = Nothing
End Sub

'報表列印
Private Sub PrintData()
Dim ii As Integer
Dim SeekPrintKind As String
   
   Printer.Orientation = 2
   GetPleft
   'With grd2
   With GRD1
      Page = 1
      SeekPrintKind = .TextMatrix(1, 9)
      PrintTitle SeekPrintKind
      For ii = 1 To .Rows - 1
         If SeekPrintKind <> .TextMatrix(ii, 9) Then
            Printer.Font.Size = 12
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.Font.Size = 10
            Printer.NewPage
            Page = Page + 1
            SeekPrintKind = .TextMatrix(ii, 9)
            PrintTitle SeekPrintKind
         End If
         Printer.Font.Size = 10
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 0)
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 1)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 2)
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 3)
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 4)
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 5)
         Printer.CurrentX = PLeft(6)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 6)
         Printer.CurrentX = PLeft(7)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 7)
         Printer.CurrentX = PLeft(8)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 8)
         'Add by Amy 2013/09/05 +承辦人
         Printer.CurrentX = PLeft(9)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 10)
         'end 2013/09/05
         'Add by Amy 2016/12/07 +進度備註
         Printer.CurrentX = PLeft(10)
         Printer.CurrentY = iPrint
         Printer.Print PUB_StrToStr_byVal(.TextMatrix(ii, 11), 46)
         'end 2016/12/07
         iPrint = iPrint + 300
         If iPrint > 10000 And ii <> .Rows - 1 Then
            If SeekPrintKind = Replace(.TextMatrix(ii + 1, 9), "*", "") Then
               Printer.Font.Size = 12
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               Printer.NewPage
               Page = Page + 1
               PrintTitle SeekPrintKind
            End If
         End If
      Next ii
   End With
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = PLeft(0) + 1800
   PLeft(2) = PLeft(1) + 1000
   PLeft(3) = PLeft(2) + 1500
   PLeft(4) = PLeft(3) + 1000
   PLeft(5) = PLeft(4) + 1000
   PLeft(6) = PLeft(5) + 500
   PLeft(7) = PLeft(6) + 1200
   PLeft(8) = PLeft(7) + 1000
   'Add by Amy 2013/09/05+承辦人
   PLeft(9) = PLeft(8) + 1000
   'end 2013/09/05
   'Add by Amy 2016/12/07 +進度備註
   PLeft(10) = PLeft(9) + 1500
End Sub

Sub PrintTitle(oClass As String)
   GetPleft
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000
   Printer.CurrentY = iPrint
   Printer.Print Val(Left((Val(txt1(0)) + 19110000), 4)) - 1911 & "年起FC案件不請款但有發文規費案件明細"
   
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & GetPrjSalesNM(strUserNum)
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "發文日期：" & txt1(0) & " - " & txt1(1)
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   iPrint = iPrint + 300
   'Add by Amy 2016/11/07
   'Modify by Amy 2016/12/07 +*已做不請款確認
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   If Check2.Value = 0 Then
      Printer.Print "未含已確認不請款"
   Else
      Printer.Print "*已做不請款確認"
   End If
   'end 2016/11/07
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "費用"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "規費"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "發文規費"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "總收文號"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "發文日"
   'Add by Amy 2013/09/05 +承辦人
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   'end 2013/09/05
   'Add by Amy 2016/12/07 +進度備註
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "進度備註"
   'end 2016/12/07
   iPrint = iPrint + 300
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.Font.Size = 10
   iPrint = iPrint + 300
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   If Index = 1 Then
      If txt1(1) <> "" Then
         If Not ChkRange(txt1(0), txt1(1), "發文日期") Then
            txt1(0).SetFocus
            TextInverse txt1(0)
         End If
      End If
   End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index)) Then
            'Modify Amy 2013/09/05 取消發文日期不可小於或等於上次列印發文日的限制
'               If Index = 0 Then
'                  If Val(Label2(1).Caption) >= Val(txt1(Index).Text) Then
'                     MsgBox "發文日期不可小於或等於上次列印發文日，請重新輸入 !", vbCritical
'                     Cancel = True
'                  End If
'               End If
            Else
               Cancel = True
            End If
'         Else
'            MsgBox "發文日期不可空白，請重新輸入 !", vbCritical
'            Cancel = True
         End If
   End Select
   If Cancel Then TextInverse txt1(Index)
End Sub
