VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210127 
   BorderStyle     =   1  '單線固定
   Caption         =   "新申請案收文至發文件數日數比較表"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox txtCountDT 
      Height          =   285
      Left            =   2070
      MaxLength       =   7
      TabIndex        =   12
      Text            =   "950101"
      Top             =   900
      Width           =   735
   End
   Begin VB.CheckBox Check6 
      Caption         =   "累計未處理量＝　　　　 至統計條件止日收文未發文件數"
      Height          =   255
      Left            =   540
      TabIndex        =   11
      Top             =   930
      Value           =   1  '核取
      Width           =   4995
   End
   Begin VB.CheckBox Check5 
      Caption         =   "相對未處理量＝收文件數－發文件數－銷案件數"
      Height          =   255
      Left            =   540
      TabIndex        =   10
      Top             =   630
      Value           =   1  '核取
      Width           =   4290
   End
   Begin VB.CheckBox Check4 
      Caption         =   "各所分別統計"
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   750
      Value           =   1  '核取
      Width           =   1440
   End
   Begin VB.TextBox txtStatDate 
      Height          =   285
      Index           =   1
      Left            =   5730
      MaxLength       =   7
      TabIndex        =   5
      Top             =   30
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "設計"
      Height          =   255
      Index           =   1
      Left            =   5430
      TabIndex        =   9
      Top             =   330
      Value           =   1  '核取
      Width           =   870
   End
   Begin VB.CheckBox Check3 
      Caption         =   "發明+新型"
      Height          =   255
      Index           =   0
      Left            =   4290
      TabIndex        =   8
      Top             =   330
      Value           =   1  '核取
      Width           =   1110
   End
   Begin VB.CheckBox Check2 
      Caption         =   "非台灣"
      Height          =   255
      Index           =   1
      Left            =   1860
      TabIndex        =   7
      Top             =   330
      Value           =   1  '核取
      Width           =   960
   End
   Begin VB.CheckBox Check2 
      Caption         =   "台灣"
      Height          =   255
      Index           =   0
      Left            =   1065
      TabIndex        =   6
      Top             =   330
      Value           =   1  '核取
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CFT"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      Top             =   30
      Value           =   1  '核取
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "T"
      Height          =   255
      Index           =   2
      Left            =   2175
      TabIndex        =   2
      Top             =   30
      Value           =   1  '核取
      Width           =   405
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CFP"
      Height          =   255
      Index           =   1
      Left            =   1515
      TabIndex        =   1
      Top             =   30
      Value           =   1  '核取
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "P"
      Height          =   255
      Index           =   0
      Left            =   1065
      TabIndex        =   0
      Top             =   30
      Value           =   1  '核取
      Width           =   405
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   7320
      TabIndex        =   15
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox txtStatDate 
      Height          =   285
      Index           =   0
      Left            =   4845
      MaxLength       =   7
      TabIndex        =   4
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6510
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8130
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   30
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3735
      Left            =   60
      TabIndex        =   17
      Top             =   1980
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   21
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   21
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label6 
      Caption         =   "註："
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   105
      TabIndex        =   23
      Top             =   660
      Width           =   405
   End
   Begin VB.Label Label8 
      Caption         =   "平均會完整理工作天數＝會稿完成至發文平均工作天數"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   810
      TabIndex        =   25
      Top             =   1710
      Width           =   7335
   End
   Begin VB.Label Label7 
      Caption         =   "平均作業工作天數＝齊備至會稿平均工作天數，平均會稿工作天數＝會稿至會稿完成平均工作天數"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   810
      TabIndex        =   24
      Top             =   1470
      Width           =   7935
   End
   Begin VB.Label Label5 
      Caption         =   "平均工作天數＝收文至發文平均工作天數，平均齊備工作天數＝收文至齊備平均工作天數"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   810
      TabIndex        =   22
      Top             =   1230
      Width           =   7335
   End
   Begin VB.Line Line1 
      X1              =   5550
      X2              =   5730
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Left            =   3345
      TabIndex        =   21
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   105
      TabIndex        =   20
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   19
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收發文銷案期間："
      Height          =   180
      Left            =   3345
      TabIndex        =   18
      Top             =   60
      Width           =   1440
   End
End
Attribute VB_Name = "frm210127"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'Create by Morgan 2008/9/11
Option Explicit

'列印用變數
Dim m_i As Integer
Dim PLeft(1 To 21) As Integer
Dim strTemp(1 To 21) As String
Dim iLine As Integer
Dim strType As String
Dim iCnt As Integer


Private Function doQuery(Optional bolTot As Boolean = False) As Boolean
   Dim stVTable1 As String, stVTable2 As String
   Dim stConSysP As String, stConSysT As String
   Dim stConPA As String, stConTM As String
   Dim iRow As Integer
   Dim strST06 As String
   
On Error GoTo ErrHnd
   
   doQuery = False
   If Check4 = 1 And bolTot = False Then '各所分別統計
      strST06 = "st06"
   Else
      strST06 = "''"
   End If
   
   '系統別
   stConSysP = "''"
   If Check1(0) = 1 Then
      stConSysP = stConSysP & ",'P'"
   End If
   If Check1(1) = 1 Then
      stConSysP = stConSysP & ",'CFP'"
   End If
   
   stConSysT = "''"
   If Check1(2) = 1 Then
      stConSysT = stConSysT & ",'T'"
   End If
   If Check1(3) = 1 Then
      stConSysT = stConSysT & ",'CFT'"
   End If
   '申請國家
   stConPA = "": stConTM = ""
   '台灣
   If Check2(0) = 1 And Check2(1) = 0 Then
      stConPA = stConPA & " and pa09='000'"
      stConTM = stConTM & " and tm10='000'"
   '非台灣
   ElseIf Check2(0) = 0 And Check2(1) = 1 Then
      stConPA = stConPA & " and pa09<>'000'"
      stConTM = stConTM & " and tm10<>'000'"
   End If
   '專利種類
   '發明+新型
   If Check3(0) = 1 And Check3(1) = 0 Then
      stConPA = stConPA & " and pa08<'3'"
   '設計
   ElseIf Check3(0) = 0 And Check3(1) = 1 Then
      stConPA = stConPA & " and pa08='3'"
   End If
   
   'S1:收文件數,S2:發文件數,S3:無基礎案之主案-發文工作天數,S4:銷案件數,S5:累計未處理量,S6:無基礎案之主案-發文件數
   'S7:無基礎案之主案-齊備工作天,  S8:無基礎案之主案-有齊備日的發文件數
   'S9:無基礎案之主案-作業工作天,  S10:無基礎案之主案-有齊備日及會稿日的發文件數
   'S11:無基礎案之主案-會稿工作天,  S12:無基礎案之主案-有會稿日及會稿完成日的發文件數
   'S13:無基礎案之主案-會完整理工作天,  S14:無基礎案之主案-有會稿完成日及發文日的發文件數
   'S15:有基礎案之主案-發文工作天數,  S16:有基礎案之主案-發文件數
   'S17:有基礎案之主案-齊備工作天,  S18:有基礎案之主案-有齊備日的發文件數
   'S19:有基礎案之主案-作業工作天,  S20:有基礎案之主案-有齊備日及會稿日的發文件數
   'S21:有基礎案之主案-會稿工作天,  S22:有基礎案之主案-有會稿日及會稿完成日的發文件數
   'S23:有基礎案之主案-會完整理工作天,  S24:有基礎案之主案-有會稿完成日及發文日的發文件數
   strExc(1) = ""
   If stConSysP <> "''" Then
      '收文
      stVTable1 = "select " & strST06 & ",pa01,pa09,pa08,count(*) S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
         " and cp05>=" & DBDATE(txtStatDate(0)) & " and cp05<=" & DBDATE(txtStatDate(1)) & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " group by " & strST06 & ",pa01,pa09,pa08"
         
      '發文
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,count(*) S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " group by " & strST06 & ",pa01,pa09,pa08"
         
      '無基礎案之主案-發文工作天(收文日到發文日的工作天數)
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,count(*) S3,0 S4,0 S5,count(distinct cp01||cp02||cp03||cp04) S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and wd01>=cp05 and wd01<=cp27" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-齊備工作天(收文日到齊備日的工作天數)
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,count(*) S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null)" & _
         " and wd01>=cp05 and wd01<=nvl(ep36,ep06)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-有齊備日的發文件數
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,count(*) S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-作業工作天(齊備日到會稿日的工作天數)
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      '                          and wd01<=nvl(ep37,ep07) ==> and wd01<=ep07
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,count(*) S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null) and (ep07 is not null)" & _
         " and wd01>=nvl(ep36,ep06) and wd01<=ep07" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-有齊備日及會稿日的發文件數
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,count(*) S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null) and (ep07 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-會稿工作天(會稿日到會稿完成日的工作天數)
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      '                          and wd01>=nvl(ep37,ep07) ==> and wd01>=ep07
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,count(*) S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep07 is not null) and (ep08 is not null or ep38 is not null)" & _
         " and wd01>=ep07 and wd01<=nvl(ep38,ep08)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-有會稿日及會稿完成日的發文件數
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,count(*) S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep07 is not null) and (ep08 is not null or ep38 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '*****
      '無基礎案之主案-會完整理工作天(會稿完成日到發文日的工作天數)
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,count(*) S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep08 is not null or ep38 is not null)" & _
         " and wd01>=nvl(ep38,ep08) and wd01<=cp27" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '無基礎案之主案-有會稿完成日及發文日的發文件數
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,count(*) S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep08 is not null or ep38 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
         
      '銷案
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,count(*) S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
         " and cp57>=" & DBDATE(txtStatDate(0)) & " and cp57<=" & DBDATE(txtStatDate(1)) & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " group by " & strST06 & ",pa01,pa09,pa08"
         
      '累計未處理(95/1/1收文至今未發文未銷案)
      'Modify By Sindy 2011/2/9 將累計起算日期改成使用者可以自行輸入
      If Check6 = 1 Then
         stVTable1 = stVTable1 & " Union all" & _
            " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,count(*) S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
            " from caseprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
            " and instr('" & NewCasePtyList & "',cp10)>0 and cp27||cp57 is null and cp09<'B'" & _
            " and cp05>=" & ChangeTStringToWString(txtCountDT) & " and cp05<=" & DBDATE(txtStatDate(1)) & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
            " group by " & strST06 & ",pa01,pa09,pa08"
      End If
      
      '*****
      '有基礎案之主案-發文工作天(收文日到發文日的工作天數)
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,count(*) S15,count(distinct cp01||cp02||cp03||cp04) S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and wd01>=cp05 and wd01<=cp27" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-齊備工作天(收文日到齊備日的工作天數)
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,count(*) S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null)" & _
         " and wd01>=cp05 and wd01<=nvl(ep36,ep06)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-有齊備日的發文件數
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,count(*) S18,0 S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-作業工作天(齊備日到會稿日的工作天數)
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      '                          and wd01<=nvl(ep37,ep07) ==> and wd01<=ep07
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,count(*) S19,0 S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null) and (ep07 is not null)" & _
         " and wd01>=nvl(ep36,ep06) and wd01<=ep07" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-有齊備日及會稿日的發文件數
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,count(*) S20,0 S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep06 is not null or ep36 is not null) and (ep07 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-會稿工作天(會稿日到會稿完成日的工作天數)
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      '                          and wd01>=nvl(ep37,ep07) ==> and wd01>=ep07
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,count(*) S21,0 S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep07 is not null) and (ep08 is not null or ep38 is not null)" & _
         " and wd01>=ep07 and wd01<=nvl(ep38,ep08)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-有會稿日及會稿完成日的發文件數
      'Modify By Sindy 2018/8/31 and (ep07 is not null or ep37 is not null) ==> and (ep07 is not null) 取消ep37
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,count(*) S22,0 S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep07 is not null) and (ep08 is not null or ep38 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-會完整理工作天(會稿完成日到發文日的工作天數)
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,count(*) S23,0 S24" & _
         " from caseprogress,casemap,engineerprogress,patent,workday,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep08 is not null or ep38 is not null)" & _
         " and wd01>=nvl(ep38,ep08) and wd01<=cp27" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
      '有基礎案之主案-有會稿完成日及發文日的發文件數
      stVTable1 = stVTable1 & " Union all" & _
         " select " & strST06 & ",pa01,pa09,pa08,0 S1,0 S2,0 S3,0 S4,0 S5,0 S6,0 S7,0 S8,0 S9,0 S10,0 S11,0 S12,0 S13,0 S14,0 S15,0 S16,0 S17,0 S18,0 S19,0 S20,0 S21,0 S22,0 S23,count(*) S24" & _
         " from caseprogress,casemap,engineerprogress,patent,staff where cp01 in (" & stConSysP & ")" & _
         " and instr('" & NewCasePtyList & "',cp10)>0 and cp21 is null and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm01 is not null and cm10(+)='0'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp13=st01(+)" & stConPA & _
         " and ep02(+)=cp09 and (ep08 is not null or ep38 is not null)" & _
         " group by " & strST06 & ",pa01,pa09,pa08"
            
      strExc(1) = "select " & strST06 & " X0,pa01 X1,decode(pa09,'000',0,1) X2,decode(pa08,'3',1,0) X3" & _
         ",sum(S1) X4,sum(S2) X5,sum(S3) X6,sum(S4) X7,sum(S5) X8,sum(S6) X9,sum(S7) X10,sum(S8) X11,sum(S9) X12,sum(S10) X13,sum(S11) X14,sum(S12) X15,sum(S13) X16,sum(S14) X17" & _
         ",sum(S15) X18,sum(S16) X19,sum(S17) X20,sum(S18) X21,sum(S19) X22,sum(S20) X23,sum(S21) X24,sum(S22) X25,sum(S23) X26,sum(S24) X27" & _
         " from (" & stVTable1 & ") X group by " & strST06 & ",pa01,decode(pa09,'000',0,1),decode(pa08,'3',1,0)"
   End If
   
   strExc(2) = ""
   If stConSysT <> "''" Then
      '收文
      stVTable2 = "select " & strST06 & ",tm01,tm10,count(*) S1,0 S2,0 S3,0 S4,0 S5" & _
         " from caseprogress,trademark,staff where cp01 in (" & stConSysT & ")" & _
         " and cp10='101' and cp09<'B'" & _
         " and cp05>=" & DBDATE(txtStatDate(0)) & " and cp05<=" & DBDATE(txtStatDate(1)) & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cp13=st01(+)" & stConTM & _
         " group by " & strST06 & ",tm01,tm10"
         
      '發文
      stVTable2 = stVTable2 & " Union all" & _
         " select " & strST06 & ",tm01,tm10,0 S1,count(*) S2,0 S3,0 S4,0 S5" & _
         " from caseprogress,trademark,staff where cp01 in (" & stConSysT & ")" & _
         " and cp10='101' and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cp13=st01(+)" & stConTM & _
         " group by " & strST06 & ",tm01,tm10"
      
      '發文工作天(收文日到發文日的工作天數)
      stVTable2 = stVTable2 & " Union all" & _
         " select " & strST06 & ",tm01,tm10,0 S1,0 S2,count(*) S3,0 S4,0 S5" & _
         " from caseprogress,trademark,workday,staff where cp01 in (" & stConSysT & ")" & _
         " and cp10='101' and cp09<'B'" & _
         " and cp27>=" & DBDATE(txtStatDate(0)) & " and cp27<=" & DBDATE(txtStatDate(1)) & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cp13=st01(+)" & stConTM & _
         " and wd01>=cp05 and wd01<=cp27" & _
         " group by " & strST06 & ",tm01,tm10"
         
      '銷案
      stVTable2 = stVTable2 & " Union all" & _
         " select " & strST06 & ",tm01,tm10,0 S1,0 S2,0 S3,count(*) S4,0 S5" & _
         " from caseprogress,trademark,staff where cp01 in (" & stConSysT & ")" & _
         " and cp10='101' and cp09<'B'" & _
         " and cp57>=" & DBDATE(txtStatDate(0)) & " and cp57<=" & DBDATE(txtStatDate(1)) & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cp13=st01(+)" & stConTM & _
         " group by " & strST06 & ",tm01,tm10"
         
      '累計未處理
      'Modify By Sindy 2011/2/9 將累計起算日期改成使用者可以自行輸入
      If Check6 = 1 Then
         stVTable2 = stVTable2 & " Union all" & _
            " select " & strST06 & ",tm01,tm10,0 S1,0 S2,0 S3,0 S4,count(*) S5" & _
            " from caseprogress,trademark,staff where cp01 in (" & stConSysT & ")" & _
            " and cp10='101' and cp27||cp57 is null and cp09<'B'" & _
            " and cp05>=" & ChangeTStringToWString(txtCountDT) & " and cp05<=" & DBDATE(txtStatDate(1)) & _
            " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cp13=st01(+)" & stConTM & _
            " group by " & strST06 & ",tm01,tm10"
      End If
      
      strExc(2) = "select " & strST06 & " X0,tm01 X1,decode(tm10,'000',0,1) X2,0 X3" & _
         ",sum(S1) X4,sum(S2) X5,sum(S3) X6,sum(S4) X7,sum(S5) X8,0 X9,0 X10,0 X11,0 X12,0 X13,0 X14,0 X15,0 X16,0 X17" & _
         ",0 X18,0 X19,0 X20,0 X21,0 X22,0 X23,0 X24,0 X25,0 X26,0 X27" & _
         " from (" & stVTable2 & ") X group by " & strST06 & ",tm01,decode(tm10,'000',0,1)"
   End If
   
   If strExc(1) <> "" And strExc(2) <> "" Then
      strExc(0) = strExc(1) & " union all " & strExc(2)
   Else
      strExc(0) = strExc(1) & strExc(2)
   End If
   strExc(0) = strExc(0) & " order by 1,2,3"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If bolTot = True Then
            iRow = grdDataList.Rows - 1
         Else
            iRow = 0
         End If
         Do While Not .EOF
            iRow = iRow + 1
            grdDataList.Rows = iRow + 1
            '所別
            intI = 0
            If Check4 = 1 And bolTot = False Then
               If "" & .Fields("X0") = "1" Then
                  grdDataList.TextMatrix(iRow, intI) = "北所"
               ElseIf "" & .Fields("X0") = "2" Then
                  grdDataList.TextMatrix(iRow, intI) = "中所"
               ElseIf "" & .Fields("X0") = "3" Then
                  grdDataList.TextMatrix(iRow, intI) = "南所"
               ElseIf "" & .Fields("X0") = "4" Then
                  grdDataList.TextMatrix(iRow, intI) = "高所"
               Else
                  grdDataList.TextMatrix(iRow, intI) = "其他"
               End If
            ElseIf bolTot = True Then
               grdDataList.TextMatrix(iRow, intI) = "合計"
            End If
            '系統類別
            intI = intI + 1
            grdDataList.TextMatrix(iRow, intI) = "" & .Fields("X1")
            '申請國家
            intI = intI + 1
            If "" & .Fields("X1") = "T" Or "" & .Fields("X1") = "P" Then
               If .Fields("X2") = 0 Then
                  grdDataList.TextMatrix(iRow, intI) = "台灣"
               Else
                  grdDataList.TextMatrix(iRow, intI) = "非台灣"
               End If
            End If
            '專利種類
            intI = intI + 1
            If .Fields("X1") = "P" Or .Fields("X1") = "CFP" Then
               If .Fields("X3") = 0 Then
                  grdDataList.TextMatrix(iRow, intI) = "發明+新型"
               Else
                  grdDataList.TextMatrix(iRow, intI) = "設計"
               End If
            End If
            '收文件數
            intI = intI + 1
            If Val("" & .Fields("X4")) <> 0 Then
               grdDataList.TextMatrix(iRow, intI) = "" & .Fields("X4")
            End If
            '發文件數
            intI = intI + 1
            If Val("" & .Fields("X5")) <> 0 Then
               grdDataList.TextMatrix(iRow, intI) = "" & .Fields("X5")
            End If
            '銷案件數
            intI = intI + 1
            If Val("" & .Fields("X7")) <> 0 Then
               grdDataList.TextMatrix(iRow, intI) = "" & .Fields("X7")
            End If
            '相對未處理量
            If Check5 = 1 Then
               intI = intI + 1
               If Val("" & .Fields("X4")) - Val("" & .Fields("X5")) - Val("" & .Fields("X7")) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Val("" & .Fields("X4")) - Val("" & .Fields("X5")) - Val("" & .Fields("X7"))
               End If
            End If
            '累計未處理量
            If Check6 = 1 Then
               intI = intI + 1
               If Val("" & .Fields("X8")) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = "" & .Fields("X8")
               End If
            End If
            '無基礎案之主案-發文件數
            intI = intI + 1
            If .Fields("X1") = "P" Or .Fields("X1") = "CFP" Then
               '2010/12/13 modify by sonia 改為顯示無基礎案之主案件數
               'grdDataList.TextMatrix(iRow, intI) = Val("" & .Fields("X5")) - Val("" & .Fields("X9"))
               If Val("" & .Fields("X9")) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Val("" & .Fields("X9"))
               End If
            End If
'            '發文件數總工作天
'            intI = intI + 1
'            grdDataList.TextMatrix(iRow, intI) = "" & .Fields("X6")
            '無基礎案之主案-平均工作天數
            intI = intI + 1
            If Val("" & .Fields("X9")) > 0 Then
               If Round(Val("" & .Fields("X6")) / Val("" & .Fields("X9")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X6")) / Val("" & .Fields("X9")), 1)
               End If
            '2010/12/13 add by sonia
            ElseIf Val("" & .Fields("X5")) > 0 Then
               If Round(Val("" & .Fields("X6")) / Val("" & .Fields("X5")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X6")) / Val("" & .Fields("X5")), 1)
               End If
            '2010/12/13 end
            End If
'            '約當週數
'            intI = intI + 1
'            If grdDataList.TextMatrix(iRow, intI - 1) <> "" Then
'               grdDataList.TextMatrix(iRow, intI) = Round(Val(grdDataList.TextMatrix(iRow, intI - 1)) / 5, 1)
'            End If
            '無基礎案之主案-平均齊備工作天數
            intI = intI + 1
            If Val("" & .Fields("X11")) > 0 Then
               If Round(Val("" & .Fields("X10")) / Val("" & .Fields("X11")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X10")) / Val("" & .Fields("X11")), 1)
               End If
            End If
            '無基礎案之主案-平均作業工作天數
            intI = intI + 1
            If Val("" & .Fields("X13")) > 0 Then
               If Round(Val("" & .Fields("X12")) / Val("" & .Fields("X13")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X12")) / Val("" & .Fields("X13")), 1)
               End If
            End If
            '無基礎案之主案-平均會稿工作天數
            intI = intI + 1
            If Val("" & .Fields("X15")) > 0 Then
               If Round(Val("" & .Fields("X14")) / Val("" & .Fields("X15")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X14")) / Val("" & .Fields("X15")), 1)
               End If
            End If
            '無基礎案之主案-平均會完整理工作天數
            intI = intI + 1
            If Val("" & .Fields("X17")) > 0 Then
               If Round(Val("" & .Fields("X16")) / Val("" & .Fields("X17")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X16")) / Val("" & .Fields("X17")), 1)
               End If
            End If
                        
            '*****
            '有基礎案之主案-發文件數
            intI = intI + 1
            If .Fields("X1") = "P" Or .Fields("X1") = "CFP" Then
               If Val("" & .Fields("X19")) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Val("" & .Fields("X19"))
               End If
            End If
            '有基礎案之主案-平均工作天數
            intI = intI + 1
            If Val("" & .Fields("X19")) > 0 Then
               If Round(Val("" & .Fields("X18")) / Val("" & .Fields("X19")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X18")) / Val("" & .Fields("X19")), 1)
               End If
            ElseIf Val("" & .Fields("X5")) > 0 Then
               If Round(Val("" & .Fields("X18")) / Val("" & .Fields("X5")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X18")) / Val("" & .Fields("X5")), 1)
               End If
            End If
            '有基礎案之主案-平均齊備工作天數
            intI = intI + 1
            If Val("" & .Fields("X21")) > 0 Then
               If Round(Val("" & .Fields("X20")) / Val("" & .Fields("X21")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X20")) / Val("" & .Fields("X21")), 1)
               End If
            End If
            '有基礎案之主案-平均作業工作天數
            intI = intI + 1
            If Val("" & .Fields("X23")) > 0 Then
               If Round(Val("" & .Fields("X22")) / Val("" & .Fields("X23")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X22")) / Val("" & .Fields("X23")), 1)
               End If
            End If
            '有基礎案之主案-平均會稿工作天數
            intI = intI + 1
            If Val("" & .Fields("X25")) > 0 Then
               If Round(Val("" & .Fields("X24")) / Val("" & .Fields("X25")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X24")) / Val("" & .Fields("X25")), 1)
               End If
            End If
            '有基礎案之主案-平均會完整理工作天數
            intI = intI + 1
            If Val("" & .Fields("X27")) > 0 Then
               If Round(Val("" & .Fields("X26")) / Val("" & .Fields("X27")), 1) <> 0 Then
                  grdDataList.TextMatrix(iRow, intI) = Round(Val("" & .Fields("X26")) / Val("" & .Fields("X27")), 1)
               End If
            End If
            
            .MoveNext
         Loop
      End With
      If Check4 = 1 And bolTot = False Then '各所分別統計-合計
         Call doQuery(True)
      End If
   Else
      ShowNoData
   End If
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   If DoPrint = True Then
      MsgBox "列印完成", vbInformation
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function DoPrint() As Boolean
   Dim i As Integer
   
On Error GoTo ErrHnd
   
   Printer.Orientation = 2 '1.直印 2.橫印
   iLine = 1
   strType = ""
   For i = 1 To grdDataList.Rows - 1
      For m_i = 1 To 21
         strTemp(m_i) = ""
      Next m_i
      
      If Check4 = 1 Then '各所分別統計
         strTemp(1) = CheckStr(grdDataList.TextMatrix(i, 0))
      Else
         strTemp(1) = ""
      End If
      strTemp(2) = CheckStr(grdDataList.TextMatrix(i, 1))
      strTemp(3) = CheckStr(grdDataList.TextMatrix(i, 2))
      strTemp(4) = CheckStr(grdDataList.TextMatrix(i, 3))
      strTemp(5) = CheckStr(grdDataList.TextMatrix(i, 4))
      strTemp(6) = CheckStr(grdDataList.TextMatrix(i, 5))
      strTemp(7) = CheckStr(grdDataList.TextMatrix(i, 6))
      strTemp(8) = CheckStr(grdDataList.TextMatrix(i, 7))
      strTemp(9) = CheckStr(grdDataList.TextMatrix(i, 8))
      strTemp(10) = CheckStr(grdDataList.TextMatrix(i, 9))
      strTemp(11) = CheckStr(grdDataList.TextMatrix(i, 10))
      strTemp(12) = CheckStr(grdDataList.TextMatrix(i, 11))
      strTemp(13) = CheckStr(grdDataList.TextMatrix(i, 12))
      strTemp(14) = CheckStr(grdDataList.TextMatrix(i, 13))
      '*****
      strTemp(15) = CheckStr(grdDataList.TextMatrix(i, 14))
      strTemp(16) = CheckStr(grdDataList.TextMatrix(i, 15))
      strTemp(17) = CheckStr(grdDataList.TextMatrix(i, 16))
      strTemp(18) = CheckStr(grdDataList.TextMatrix(i, 17))
      strTemp(19) = CheckStr(grdDataList.TextMatrix(i, 18))
      strTemp(20) = CheckStr(grdDataList.TextMatrix(i, 19))
      strTemp(21) = CheckStr(grdDataList.TextMatrix(i, 20))
      
      If iLine > 33 Or iLine = 1 Then
         If i <> 1 Then Printer.NewPage
         iLine = 1
         PrintTitle '列印表頭
      Else
         If Check4 = 1 And strType <> strTemp(1) Then '各所分別統計
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iLine * 300
            Printer.Print String(205, "-")
            iLine = iLine + 1
         End If
      End If
      
      PrintDetail
      
      strType = strTemp(1)
   Next i
'   Printer.CurrentX = PLeft(1)
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(205, "-")
'   iLine = iLine + 1
   
   '備註
   If iLine > 33 Then
      Printer.NewPage
      iLine = 1
      PrintTitle '列印表頭
   End If
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "註："
   If Check5 = 1 Then
      iLine = iLine + 1
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "　　相對未處理量＝收文件數－發文件數－銷案件數"
   End If
   If Check6 = 1 Then
      iLine = iLine + 1
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "　　累計未處理量＝" & ChangeTStringToTDateString(txtCountDT) & "至統計條件止日收文未發文件數"
   End If
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　" & Label5.Caption
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　" & Label7.Caption
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　" & Label8.Caption
   iLine = iLine + 1
   
   Printer.EndDoc
   DoPrint = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Sub GetPleft()
   PLeft(1) = 300
   PLeft(2) = 1000
   PLeft(3) = 1700
   PLeft(4) = 2700
   PLeft(5) = 4500
   PLeft(6) = 5250
   PLeft(7) = 6000
   PLeft(8) = 6750
   PLeft(9) = 7500
   PLeft(10) = 8250
   PLeft(11) = 9000
   PLeft(12) = 9750
   PLeft(13) = 10500
   PLeft(14) = 11250
   PLeft(15) = 12000
   PLeft(16) = 12750
   PLeft(17) = 13500
   PLeft(18) = 14250
   PLeft(19) = 15000
   PLeft(20) = 15750
   PLeft(21) = 16500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 20
Printer.Font.Underline = False
Printer.FontBold = True

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("新申請案收文至發文件數日數比較表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "新申請案收文至發文件數日數比較表"

Printer.Font.Size = 11
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'2010/12/13 add by sonia
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 6500
Printer.CurrentY = 900
Printer.Print "收發文銷案期間：" & Format(ChangeTStringToTDateString(txtStatDate(0)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txtStatDate(1))
'2010/12/13 end
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5

Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "無基礎案之主案"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iLine * 300
Printer.Print "有基礎案之主案"

iLine = iLine + 1
Printer.CurrentX = PLeft(9) + 225
Printer.CurrentY = iLine * 300
Printer.Print String(64, "-")
Printer.CurrentX = PLeft(15) + 225
Printer.CurrentY = iLine * 300
Printer.Print String(64, "-")

iLine = iLine + 1
If Check4 = 1 Then '各所分別統計
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "所別"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "系統"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "申請國家"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "專利種類"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("收文")
Printer.CurrentY = iLine * 300
Printer.Print "收文"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("發文")
Printer.CurrentY = iLine * 300
Printer.Print "發文"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("銷案")
Printer.CurrentY = iLine * 300
Printer.Print "銷案"
If Check5 = 1 Then
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("相對未")
   Printer.CurrentY = iLine * 300
   Printer.Print "相對未"
End If
If Check6 = 1 Then
   Printer.CurrentX = PLeft(9) - Printer.TextWidth("累計未")
   Printer.CurrentY = iLine * 300
   Printer.Print "累計未"
End If
Printer.CurrentX = PLeft(10) - Printer.TextWidth("發文")
Printer.CurrentY = iLine * 300
Printer.Print "發文"
Printer.Font.Size = 8
Printer.CurrentX = PLeft(11) - Printer.TextWidth("平均工")
Printer.CurrentY = iLine * 300
Printer.Print "平均工"
Printer.CurrentX = PLeft(12) - Printer.TextWidth("平均齊備")
Printer.CurrentY = iLine * 300
Printer.Print "平均齊備"
Printer.CurrentX = PLeft(13) - Printer.TextWidth("平均作業")
Printer.CurrentY = iLine * 300
Printer.Print "平均作業"
Printer.CurrentX = PLeft(14) - Printer.TextWidth("平均會稿")
Printer.CurrentY = iLine * 300
Printer.Print "平均會稿"
Printer.Font.Size = 7
Printer.CurrentX = PLeft(15) - Printer.TextWidth("平均會完整")
Printer.CurrentY = iLine * 300
Printer.Print "平均會完整"
'*****
Printer.Font.Size = 11
Printer.CurrentX = PLeft(16) - Printer.TextWidth("發文")
Printer.CurrentY = iLine * 300
Printer.Print "發文"
Printer.Font.Size = 8
Printer.CurrentX = PLeft(17) - Printer.TextWidth("平均工")
Printer.CurrentY = iLine * 300
Printer.Print "平均工"
Printer.CurrentX = PLeft(18) - Printer.TextWidth("平均齊備")
Printer.CurrentY = iLine * 300
Printer.Print "平均齊備"
Printer.CurrentX = PLeft(19) - Printer.TextWidth("平均作業")
Printer.CurrentY = iLine * 300
Printer.Print "平均作業"
Printer.CurrentX = PLeft(20) - Printer.TextWidth("平均會稿")
Printer.CurrentY = iLine * 300
Printer.Print "平均會稿"
Printer.Font.Size = 7
Printer.CurrentX = PLeft(21) - Printer.TextWidth("平均會完整")
Printer.CurrentY = iLine * 300
Printer.Print "平均會完整"

Printer.Font.Size = 11
iLine = iLine + 1
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "類別"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("件數")
Printer.CurrentY = iLine * 300
Printer.Print "件數"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("件數")
Printer.CurrentY = iLine * 300
Printer.Print "件數"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("件數")
Printer.CurrentY = iLine * 300
Printer.Print "件數"
If Check5 = 1 Then
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("處理量")
   Printer.CurrentY = iLine * 300
   Printer.Print "處理量"
End If
If Check6 = 1 Then
   Printer.CurrentX = PLeft(9) - Printer.TextWidth("處理量")
   Printer.CurrentY = iLine * 300
   Printer.Print "處理量"
End If
Printer.CurrentX = PLeft(10) - Printer.TextWidth("件數")
Printer.CurrentY = iLine * 300
Printer.Print "件數"
Printer.Font.Size = 8
Printer.CurrentX = PLeft(11) - Printer.TextWidth("作天數")
Printer.CurrentY = iLine * 300
Printer.Print "作天數"
Printer.CurrentX = PLeft(12) - Printer.TextWidth("工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "工作天數"
Printer.CurrentX = PLeft(13) - Printer.TextWidth("工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "工作天數"
Printer.CurrentX = PLeft(14) - Printer.TextWidth("工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "工作天數"
Printer.Font.Size = 7
Printer.CurrentX = PLeft(15) - Printer.TextWidth("理工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "理工作天數"
'*****
Printer.Font.Size = 11
Printer.CurrentX = PLeft(16) - Printer.TextWidth("件數")
Printer.CurrentY = iLine * 300
Printer.Print "件數"
Printer.Font.Size = 8
Printer.CurrentX = PLeft(17) - Printer.TextWidth("作天數")
Printer.CurrentY = iLine * 300
Printer.Print "作天數"
Printer.CurrentX = PLeft(18) - Printer.TextWidth("工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "工作天數"
Printer.CurrentX = PLeft(19) - Printer.TextWidth("工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "工作天數"
Printer.CurrentX = PLeft(20) - Printer.TextWidth("工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "工作天數"
Printer.Font.Size = 7
Printer.CurrentX = PLeft(21) - Printer.TextWidth("理工作天數")
Printer.CurrentY = iLine * 300
Printer.Print "理工作天數"

Printer.Font.Size = 11
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(242, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer, m_row As Integer
   
   For m_j = 1 To 21 - iCnt
      If (m_j = 8 And Check5 = 0) Then
         m_row = m_j + iCnt
      ElseIf m_j = 9 And Check5 <> 0 And Check6 = 0 Then
         m_row = m_j + iCnt
      Else
         m_row = m_row + 1
      End If
      If m_row > 4 Then
         Printer.CurrentX = PLeft(m_row) - Printer.TextWidth(strTemp(m_j))
      Else
         Printer.CurrentX = PLeft(m_row)
      End If
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   cmdPrint.Enabled = False
   grdDataList.Clear
   InitGrid
   If ConstrainCheck = True Then
      Call doQuery
      cmdPrint.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210127 = Nothing
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bCancel As Boolean
   
   If Check1(0) + Check1(1) + Check1(2) + Check1(3) = 0 Then
      MsgBox "請勾選系統類別！"
      Exit Function
   End If
   
   If txtStatDate(0) = "" Then
      MsgBox "收發文銷案期間起日不可空白！"
      txtStatDate(0).SetFocus
      Exit Function
   Else
      txtStatDate_Validate 0, bCancel
      If bCancel = True Then Exit Function
   End If
   
   If txtStatDate(1) = "" Then
      MsgBox "收發文銷案期間迄日不可空白！"
      txtStatDate(1).SetFocus
      Exit Function
   Else
      txtStatDate_Validate 1, bCancel
      If bCancel = True Then Exit Function
   End If
   
   If Check6 = 1 Then
      If txtCountDT = "" Then
         MsgBox "累計未處理量的累計起算日不可空白！"
         txtCountDT.SetFocus
         Exit Function
      Else
         txtCountDT_Validate bCancel
         If bCancel = True Then Exit Function
      End If
   End If
   
   If Check2(0) + Check2(1) = 0 Then
      Check2(0).SetFocus
      MsgBox "請勾選申請國家！"
      Exit Function
   End If
   
   If Check1(0) + Check1(1) > 0 And Check3(0) + Check3(1) = 0 Then
      MsgBox "專利需勾選專利種類！"
      Check3(0).SetFocus
      Exit Function
   End If
   
   ConstrainCheck = True
End Function

Private Sub txtStatDate_GotFocus(Index As Integer)
   TextInverse txtStatDate(Index)
   CloseIme
End Sub

Private Sub txtStatDate_Validate(Index As Integer, Cancel As Boolean)
   If txtStatDate(Index) <> "" Then
      If Not ChkDate(txtStatDate(Index)) Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCountDT_GotFocus()
   TextInverse txtCountDT
   CloseIme
End Sub

Private Sub txtCountDT_Validate(Cancel As Boolean)
   If txtCountDT <> "" Then
      If Not ChkDate(txtCountDT) Then
         Cancel = True
      End If
   End If
End Sub

Private Sub InitGrid()
   Dim iCol As Integer
   
   iCnt = 0
   If Check5 = 0 Then iCnt = iCnt + 1
   If Check6 = 0 Then iCnt = iCnt + 1
   With grdDataList
      .Visible = False
      .WordWrap = True
      .Clear
      .Rows = 2: .FixedRows = 1
      .Cols = 21 - iCnt: .FixedCols = 6
      .RowHeight(0) = 950 '450
      iCol = 0
      If Check4 = 1 Then '各所分別統計
         .ColWidth(iCol) = 500
      Else
         .ColWidth(iCol) = 0
      End If
      .TextMatrix(0, iCol) = "所別"
      iCol = iCol + 1
      .ColWidth(iCol) = 480
      .TextMatrix(0, iCol) = "系統類別"
      iCol = iCol + 1
      .ColWidth(iCol) = 885
      .TextMatrix(0, iCol) = "申請國家"
      iCol = iCol + 1
      .ColWidth(iCol) = 980
      .TextMatrix(0, iCol) = "專利種類"
      iCol = iCol + 1
      .ColWidth(iCol) = 550
      .TextMatrix(0, iCol) = "收文件數"
      iCol = iCol + 1
      .ColWidth(iCol) = 550
      .TextMatrix(0, iCol) = "發文件數"
      iCol = iCol + 1
      .ColWidth(iCol) = 550
      .TextMatrix(0, iCol) = "銷案件數"
      If Check5 = 1 Then
         iCol = iCol + 1
         .ColWidth(iCol) = 675
         .TextMatrix(0, iCol) = "相對未處理量"
      End If
      If Check6 = 1 Then
         iCol = iCol + 1
         .ColWidth(iCol) = 675
         .TextMatrix(0, iCol) = "累計未處理量"
      End If
      iCol = iCol + 1
      .ColWidth(iCol) = 675
      .TextMatrix(0, iCol) = "無基礎案之主案發文件數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "無基礎案之主案平均工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "無基礎案之主案平均齊備工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "無基礎案之主案平均作業工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "無基礎案之主案平均會稿工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "無基礎案之主案平均會完整理工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 675
      .TextMatrix(0, iCol) = "有基礎案之主案發文件數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "有基礎案之主案平均工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "有基礎案之主案平均齊備工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "有基礎案之主案平均作業工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 900
      .TextMatrix(0, iCol) = "有基礎案之主案平均會稿工作天數"
      iCol = iCol + 1
      .ColWidth(iCol) = 850
      .TextMatrix(0, iCol) = "有基礎案之主案平均會完整理工作天數"
      .Visible = True
   End With
End Sub
