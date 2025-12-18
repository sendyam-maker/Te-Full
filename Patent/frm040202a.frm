VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040202a 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件結餘查詢"
   ClientHeight    =   6110
   ClientLeft      =   120
   ClientTop       =   990
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6110
   ScaleWidth      =   9320
   Begin VB.CommandButton CmdOk 
      Caption         =   "下一筆(&N)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7350
      TabIndex        =   32
      Top             =   70
      Width           =   1110
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8484
      TabIndex        =   4
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   0
      Left            =   6105
      TabIndex        =   3
      Top             =   75
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   2595
      Left            =   30
      TabIndex        =   1
      Top             =   3420
      Width           =   9255
      _ExtentX        =   16334
      _ExtentY        =   4568
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1245
      TabIndex        =   5
      Top             =   1395
      Width           =   7920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13970;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl3 
      Height          =   264
      Index           =   6
      Left            =   4650
      TabIndex        =   41
      Top             =   1800
      Width           =   1440
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2540;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "可結餘日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   3420
      TabIndex        =   40
      Top             =   1785
      Width           =   1170
   End
   Begin MSForms.Label lbl3 
      Height          =   270
      Index           =   5
      Left            =   1380
      TabIndex        =   39
      Top             =   2430
      Width           =   7755
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "13679;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   3420
      TabIndex        =   38
      Top             =   2085
      Width           =   990
   End
   Begin MSForms.Label lbl3 
      Height          =   270
      Index           =   4
      Left            =   4500
      TabIndex        =   37
      Top             =   2085
      Width           =   1440
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2540;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   195
      TabIndex        =   36
      Top             =   2430
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "lbl2"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   3810
      TabIndex        =   35
      Top             =   3210
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "退費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   3375
      TabIndex        =   34
      Top             =   3180
      Width           =   420
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3510
      TabIndex        =   33
      Top             =   240
      Width           =   990
   End
   Begin MSForms.Label lbl3 
      Height          =   270
      Index           =   3
      Left            =   7680
      TabIndex        =   31
      Top             =   2070
      Width           =   1440
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2540;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl3 
      Height          =   270
      Index           =   2
      Left            =   1380
      TabIndex        =   30
      Top             =   2085
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2805;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl3 
      Height          =   264
      Index           =   1
      Left            =   7680
      TabIndex        =   29
      Top             =   1770
      Width           =   1440
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2540;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl3 
      Height          =   264
      Index           =   0
      Left            =   1380
      TabIndex        =   28
      Top             =   1770
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "3528;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   7260
      TabIndex        =   27
      Top             =   2985
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   5370
      TabIndex        =   26
      Top             =   2970
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   3540
      TabIndex        =   25
      Top             =   2970
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1710
      TabIndex        =   24
      Top             =   2985
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "lbl2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   4
      Left            =   7590
      TabIndex        =   23
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "lbl2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   5700
      TabIndex        =   22
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "lbl2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   2
      Left            =   3810
      TabIndex        =   21
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "lbl2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   2010
      TabIndex        =   20
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "lbl2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   19
      Top             =   2985
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "實際收款金額    -    已作收入金額    -    實際支出費用    =      浮動準備金      +      結餘金額  "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   11
      Left            =   225
      TabIndex        =   18
      Top             =   2745
      Width           =   8775
   End
   Begin VB.Label Label1 
      Caption         =   "結算日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   6600
      TabIndex        =   17
      Top             =   2070
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "填表日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   6600
      TabIndex        =   16
      Top             =   1770
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   15
      Top             =   2085
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "結餘單號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   14
      Top             =   1770
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   510
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "中："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   1245
      TabIndex        =   12
      Top             =   510
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "英："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   1245
      TabIndex        =   11
      Top             =   810
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   1248
      TabIndex        =   10
      Top             =   1116
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   9
      Top             =   1425
      Width           =   780
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   0
      Left            =   1710
      TabIndex        =   8
      Top             =   510
      Width           =   6840
      VariousPropertyBits=   27
      Size            =   "12065;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   1
      Left            =   1710
      TabIndex        =   7
      Top             =   810
      Width           =   6840
      VariousPropertyBits=   27
      Size            =   "12065;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   2
      Left            =   1716
      TabIndex        =   6
      Top             =   1116
      Width           =   6840
      VariousPropertyBits=   27
      Size            =   "12065;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   3
      Left            =   1095
      TabIndex        =   2
      Top             =   180
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frm040202a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 lbl1()/lbl3()/Combo1/Grd1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strSql As String, i As Integer, j As Integer, s As Integer, strTemp As Variant, strLBL(2) As String
Dim StrTest As String, IntTest As Integer, IntTest2 As Integer, k As Boolean, intK As Integer

Sub SetGridWidth()        '顯示GRID上方字
With Grd1
    .Cols = 5
    .row = 0
    .col = 0
    .ColWidth(0) = 1000
    .Text = "收文號"
    .col = 1
    .ColWidth(1) = 1000
    .Text = "收文日"
    .col = 2
    .ColWidth(2) = 2000
    .Text = "案件性質"
    .col = 3
    .ColWidth(3) = 1000
    .Text = "發文日"
    .col = 4
    .ColWidth(4) = 1000
    .Text = "智權人員"
End With
End Sub

Sub StrMenu()
'add by nickc 代代理人
strSql = "select A240012 from acc240 where A240002='" & Trim(lbl3(0)) & "' "
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      lbl3(5) = CheckStr(.Fields(0))
   End If
End With

'edit by nickc 2005/07/21 換新 table
'edit by nickc 2005/09/16 加入可結餘日
'strSQL = "select cp09," & SQLDate("cp05") & ",decode(pa09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13) from caseprogress,staff,patent,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+) and cp01=PA01 AND cp02=PA02 AND cp03=PA03 AND cp04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "
'strSQL = strSQL + " union all select cp09," & SQLDate("cp05") & ",decode(Sp09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13) from caseprogress,staff,SERVICEPRACTICE,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and cp01=SP01 AND cp02=SP02 AND cp03=SP03 AND cp04=SP04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
'strSQL = strSQL + " union all select cp09," & SQLDate("cp05") & ",decode(TM10,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13) from caseprogress,staff,TRADEMARK,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and cp01=TM01 AND cp02=TM02 AND cp03=TM03 AND cp04=TM04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
'strSQL = strSQL + " union all select cp09," & SQLDate("cp05") & ",decode(LC15,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13) from caseprogress,staff,LAWCASE,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and cp01=LC01 AND cp02=LC02 AND cp03=LC03 AND cp04=LC04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "
'Modify by Morgan 2010/8/11 百年蟲
'strSql = "select distinct cp09," & SQLDate("cp05") & ",decode(pa09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,patent,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+) and cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "
'strSql = strSql + " union  select cp09," & SQLDate("cp05") & ",decode(Sp09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,SERVICEPRACTICE,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and A240005=SP01(+) AND A240006=SP02(+) AND A240007=SP03(+) AND A240008=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
'strSql = strSql + " union  select cp09," & SQLDate("cp05") & ",decode(TM10,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,TRADEMARK,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and A240005=TM01(+) AND A240006=TM02(+) AND A240007=TM03(+) AND A240008=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
'strSql = strSql + " union  select cp09," & SQLDate("cp05") & ",decode(LC15,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,LAWCASE,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and A240005=LC01(+) AND A240006=LC02(+) AND A240007=LC03(+) AND A240008=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "
strSql = "select distinct cp09,substrb(' '||sqldatet(cp05),-9),decode(pa09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,patent,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and a240005<>'CFP' and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+) and cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "
'2011/5/31 ADD BY SONIA CFP案
'modify by sonia 2024/12/26 CFP子案之閉卷或不續辦不顯示故加And Cp01||Cp10 Not In ('CFP907','CFP913')
strSql = strSql + " union  select distinct cp09,substrb(' '||sqldatet(cp05),-9),decode(pa09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,patent,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and a240005='CFP' and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and (cp64 is null or instr(cp64,'子案發文記錄')=0) And Cp01||Cp10 Not In ('CFP907','CFP913') and cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "
'2011/5/31 END
strSql = strSql + " union  select cp09,substrb(' '||sqldatet(cp05),-9),decode(Sp09,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,SERVICEPRACTICE,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and A240005=SP01(+) AND A240006=SP02(+) AND A240007=SP03(+) AND A240008=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
strSql = strSql + " union  select cp09,substrb(' '||sqldatet(cp05),-9),decode(TM10,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,TRADEMARK,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and A240005=TM01(+) AND A240006=TM02(+) AND A240007=TM03(+) AND A240008=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
strSql = strSql + " union  select cp09,substrb(' '||sqldatet(cp05),-9),decode(LC15,'000',cpm03,cpm04)," & SQLDate("cp27") & ",nvl(st02,cp13),nvl(sqldatet(cp109),'逾期未處理') as cp109 from caseprogress,staff,LAWCASE,casepropertymap,acc240 where A240002='" & Trim(lbl3(0)) & "' and A240002=cp59(+) and A240005=cp01(+) and A240006=cp02(+) and A240007=cp03(+) and A240008=cp04(+)  and A240005=LC01(+) AND A240006=LC02(+) AND A240007=LC03(+) AND A240008=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+)  "

strSql = strSql + " order by 2,1 "

With adoRecordset
   CheckOC
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      .MoveFirst
      lbl3(6).Caption = CheckStr(.Fields("cp109").Value)
      Set Grd1.Recordset = adoRecordset
      SetGridWidth
      'ADD BY SONIA 2016/8/31
      For i = 1 To Grd1.Rows - 1
         Grd1.TextMatrix(i, 2) = Grd1.TextMatrix(i, 2) & PUB_GetRelateCasePropertyName(Grd1.TextMatrix(i, 0), "1")
      Next i
      'END 2016/8/31
   Else
      lbl3(6).Caption = ""
      SetGridWidth
      s = MsgBox("沒有資料!!", , "錯誤!!")
      bolGoBackByNick = True
      Exit Sub
   End If
End With
End Sub

Private Sub cmdok_Click(Index As Integer)                  'COMMAND BATTON
Select Case Index
Case 0
     bolGoBackByNick = True
     Me.Hide
Case 1
     Me.Hide
Case 2
     bolToEndByNick = True
     Me.Hide
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetGridWidth
Label1(11).Caption = "實際收款金額    -    已作收入金額    -    實際支出費用    =      浮動準備金      +      結餘金額  "
'Combo1.Text = Combo1.List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm040202a = Nothing
End Sub
