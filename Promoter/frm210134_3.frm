VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210134_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "歷史記錄查詢"
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
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&X)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7080
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   90
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2655
      Left            =   60
      TabIndex        =   10
      Top             =   3030
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label LabCP13 
      Height          =   300
      Left            =   1470
      TabIndex        =   29
      Top             =   2700
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3916;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabTM05 
      Height          =   300
      Left            =   1470
      TabIndex        =   28
      Top             =   600
      Width           =   7440
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13123;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabTM06 
      Height          =   300
      Left            =   1470
      TabIndex        =   27
      Top             =   900
      Width           =   7440
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13123;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabTM07 
      Height          =   300
      Left            =   1470
      TabIndex        =   26
      Top             =   1200
      Width           =   7440
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13123;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabCP14 
      Height          =   300
      Left            =   1470
      TabIndex        =   25
      Top             =   1800
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3916;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "智權人員："
      Height          =   255
      Index           =   27
      Left            =   540
      TabIndex        =   24
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "齊備日："
      Height          =   255
      Index           =   25
      Left            =   540
      TabIndex        =   23
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label LabEP06 
      Height          =   255
      Left            =   1470
      TabIndex        =   22
      Top             =   2400
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "會稿日："
      Height          =   255
      Index           =   23
      Left            =   3780
      TabIndex        =   21
      Top             =   2400
      Width           =   930
   End
   Begin VB.Label LabEP07 
      Height          =   255
      Left            =   4740
      TabIndex        =   20
      Top             =   2400
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所期限："
      Height          =   255
      Index           =   21
      Left            =   540
      TabIndex        =   19
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label LabCP06 
      Height          =   255
      Left            =   1470
      TabIndex        =   18
      Top             =   2100
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "法定期限："
      Height          =   255
      Index           =   19
      Left            =   3780
      TabIndex        =   17
      Top             =   2100
      Width           =   930
   End
   Begin VB.Label LabCP07 
      Height          =   255
      Left            =   4740
      TabIndex        =   16
      Top             =   2100
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   255
      Index           =   17
      Left            =   540
      TabIndex        =   15
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   10
      Left            =   3780
      TabIndex        =   14
      Top             =   1800
      Width           =   930
   End
   Begin VB.Label LabCP48 
      Height          =   255
      Left            =   4740
      TabIndex        =   13
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱(日)："
      Height          =   255
      Index           =   16
      Left            =   210
      TabIndex        =   12
      Top             =   1200
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱(英)："
      Height          =   255
      Index           =   14
      Left            =   210
      TabIndex        =   11
      Top             =   900
      Width           =   1230
   End
   Begin VB.Label LabID 
      Height          =   255
      Left            =   1470
      TabIndex        =   8
      Top             =   300
      Width           =   1740
   End
   Begin VB.Label LabCP09 
      Height          =   255
      Left            =   4740
      TabIndex        =   7
      Top             =   300
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "總收文號："
      Height          =   255
      Index           =   6
      Left            =   3780
      TabIndex        =   6
      Top             =   300
      Width           =   930
   End
   Begin VB.Label LabCP05 
      Height          =   255
      Left            =   4740
      TabIndex        =   5
      Top             =   1500
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文日："
      Height          =   255
      Index           =   4
      Left            =   3780
      TabIndex        =   4
      Top             =   1500
      Width           =   930
   End
   Begin VB.Label LabCP10 
      Height          =   255
      Left            =   1470
      TabIndex        =   3
      Top             =   1500
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   2
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱(中)："
      Height          =   255
      Index           =   11
      Left            =   210
      TabIndex        =   1
      Top             =   600
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frm210134_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ;LabTM05、LabTM06、LabTM07、LabCP14、LabCP13
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/23 日期欄已修改
Option Explicit

'紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

arrGridHeadText = Array("異動日期", "異動人員", "管制期限", "已完成", "備註")
arrGridHeadWidth = Array(850, 1000, 850, 850, 2500)
grdDataList.MergeCells = flexMergeRestrictColumns
grdDataList.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grdDataList.Cols - 1
   grdDataList.row = 0
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText(iRow)
   grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grdDataList.CellAlignment = flexAlignLeftCenter
Next
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
Select Case cmdState
Case 1
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210134_3 = Nothing
End Sub

Public Sub Process(strText As String)
   
On Error GoTo ErrHnd
   
   strSql = "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(TM10,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,TM05,TM06,TM07" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,trademark,engineerprogress" & _
                " where cp09='" & strText & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,pa05 as TM05,pa06 as TM06,pa07 as TM07" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,patent,engineerprogress" & _
                " where cp09='" & strText & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(SP09,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,sp05 as TM05,sp06 as TM06,sp07 as TM07" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,servicepractice,engineerprogress" & _
                " where cp09='" & strText & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(CPM03,cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,hc06 as TM05,'' as TM06,'' as TM07" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,hirecase,engineerprogress" & _
                " where cp09='" & strText & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,lc05 as TM05,lc06 as TM06,lc07 as TM07" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,lawcase,engineerprogress" & _
                " where cp09='" & strText & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " order by 總收文號"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         LabID.Caption = "" & .Fields("本所案號")
         LabCP09.Caption = "" & .Fields("總收文號")
         LabTM05.Caption = "" & .Fields("TM05")
         LabTM06.Caption = "" & .Fields("TM06")
         LabTM07.Caption = "" & .Fields("TM07")
         LabCP10.Caption = "" & .Fields("案件性質")
         LabCP05.Caption = "" & .Fields("收文日")
         LabCP14.Caption = "" & .Fields("承辦人")
         LabCP48.Caption = "" & .Fields("承辦期限")
         LabCP06.Caption = "" & .Fields("本所期限")
         LabCP07.Caption = "" & .Fields("法定期限")
         LabEP06.Caption = "" & .Fields("齊備日")
         LabEP07.Caption = "" & .Fields("會稿日")
         LabCP13.Caption = "" & .Fields("智權人員")
      Else
         LabID.Caption = ""
         LabCP09.Caption = ""
         LabTM05.Caption = ""
         LabTM06.Caption = ""
         LabTM07.Caption = ""
         LabCP10.Caption = ""
         LabCP05.Caption = ""
         LabCP14.Caption = ""
         LabCP48.Caption = ""
         LabCP06.Caption = ""
         LabCP07.Caption = ""
         LabEP06.Caption = ""
         LabEP07.Caption = ""
         LabCP13.Caption = ""
      End If
   End With
   strSql = "select sqldatet(sc02) as 異動日期,st02 as 異動人員,sqldatet(sc05) as 管制期限,sc08 as 已完成,sc06 as 備註 from salescontroldate,staff where sc01='" & strText & "' and sc04=st01(+) order by sc02,sc03 "
   CheckOC3
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub
