VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210137_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文之案件性質及點數明細"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7065
   Begin VB.CommandButton cmdButton 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   6000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   45
      Width           =   915
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Index           =   0
      Left            =   4920
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   45
      Width           =   1035
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label Lb_People 
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   840
      Width           =   1290
      VariousPropertyBits=   27
      Size            =   "2275;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2. 收據狀態：N 暫不列印  Y待列印 "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   2325
      TabIndex        =   10
      Top             =   600
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1. 本所案號：＊閉卷 ●銷卷"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   2325
      TabIndex        =   9
      Top             =   360
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   2325
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Lb_Total 
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "合計點數："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Lb_Date 
      Caption         =   "(民國年月日)"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "點數結算日："
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frm210137_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/25 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Lb_People
' Create by Amy 2013/04/24
Option Explicit

Dim RbMain As New ADODB.Recordset, bp As New ADODB.Recordset
Dim i As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean

Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Case 1
            fnCloseAllFrm100
    End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Sub StrMenu()
Dim rsTmp As New ADODB.Recordset
Dim strSql  As String, i As Integer, dblSub As Double
Dim Str01() As String
Dim IntTemp1 As Long
Dim IntTemp2 As Long

   Str01() = Split(Me.Tag, ",") 'str01(0)=智權人姓名/str01(1)=智權人編號/str01(2)=業務區代碼/str01(3)=結算日起始日/str01(4)=結算日終止日/str01(5)點數合計
   Lb_People.Caption = Str01(0)
   Lb_Date.Caption = Str01(3) & " ~ " & Str01(4)
   Lb_Total.Caption = Str01(5)
   
   '專利/商標/法務/顧問/服務
   'Modify by Amy 2013/05/23 增加顯示A0K32
   'strSql = "Select SQLDATET2(CP05) as 收文日,Decode(PA23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號, NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,Decode(substr(CP60,1,1),'E',Decode(CP79,0,'收回',Decode(sign(CP75),1,'部分收回','未收')),CP60) CP60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家 " & _
               "  From CaseProgress, Patent,CasePropertyMap,Nation,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=Patent.PA01(+) and CaseProgress.CP02=Patent.PA02(+) And CaseProgress.CP03=Patent.PA03(+) And CaseProgress.CP04=Patent.PA04(+) And CP01=CPM01(+) And CP10=CPM02(+) And PA09=NA01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('P','CFP','CPS') And NVL(Recpt,0) >0 " & _
               "Union All Select SQLDATET2(CP05) as 收文日,Decode(TM28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(TM57,'')),null,'','●') AS 本所案號, NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家 " & _
               "  From CaseProgress, TradeMark,CasePropertyMap,Nation,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=TradeMark.TM01(+) and CaseProgress.CP02=TradeMark.TM02(+) And CaseProgress.CP03=TradeMark.TM03(+) And CaseProgress.CP04=TradeMark.TM04(+) And CP01=CPM01(+) And CP10=CPM02(+) And TM10=NA01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFT', 'FCT', 'T', 'TF') And NVL(Recpt,0) >0"
   
   'strSql = strSql + "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,Nvl(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家 " & _
               "  From CaseProgress, LawCase,CasePropertyMap,Nation,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=LawCase.LC01(+) and CaseProgress.CP02=LawCase.LC02(+) And CaseProgress.CP03=LawCase.LC03(+) And CaseProgress.CP04=LawCase.LC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And LC15=NA01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFL', 'FCL', 'L', 'LIN') And NVL(Recpt,0) >0 " & _
               "Union  All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號, HC06 AS 案件名稱,'' as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,'' AS 申請國家 " & _
               "  From CaseProgress, HireCase,CasePropertyMap,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=HireCase.HC01(+) and CaseProgress.CP02=HireCase.HC02(+) And CaseProgress.CP03=HireCase.HC03(+) And CaseProgress.CP04=HireCase.HC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('LA') And NVL(Recpt,0) >0 " & _
               "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號, NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家 " & _
               "  From CaseProgress, ServicePractice,CasePropertyMap,Nation,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=ServicePractice.SP01(+) and CaseProgress.CP02=ServicePractice.SP02(+) And CaseProgress.CP03=ServicePractice.SP03(+) And CaseProgress.CP04=ServicePractice.SP04(+) And CP01=CPM01(+) And CP10=CPM02(+) And SP09=NA01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 Not In ('P','CFP','CPS','CFT', 'FCT', 'T', 'TF','CFL', 'FCL', 'L', 'LIN','LA') And NVL(Recpt,0) >0"
   
   'strSql = "Select SQLDATET2(CP05) as 收文日,Decode(PA23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號, NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,Decode(substr(CP60,1,1),'E',Decode(CP79,0,'收回',Decode(sign(CP75),1,'部分收回','未收')),CP60) CP60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32 " & _
               "  From CaseProgress, Patent,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=Patent.PA01(+) and CaseProgress.CP02=Patent.PA02(+) And CaseProgress.CP03=Patent.PA03(+) And CaseProgress.CP04=Patent.PA04(+) And CP01=CPM01(+) And CP10=CPM02(+) And PA09=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('P','CFP','CPS') And NVL(Recpt,0) >0 " & _
               "Union All Select SQLDATET2(CP05) as 收文日,Decode(TM28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(TM57,'')),null,'','●') AS 本所案號, NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32 " & _
               "  From CaseProgress, TradeMark,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=TradeMark.TM01(+) and CaseProgress.CP02=TradeMark.TM02(+) And CaseProgress.CP03=TradeMark.TM03(+) And CaseProgress.CP04=TradeMark.TM04(+) And CP01=CPM01(+) And CP10=CPM02(+) And TM10=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFT', 'FCT', 'T', 'TF') And NVL(Recpt,0) >0"
   
   'strSql = strSql + "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,Nvl(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32 " & _
               "  From CaseProgress, LawCase,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=LawCase.LC01(+) and CaseProgress.CP02=LawCase.LC02(+) And CaseProgress.CP03=LawCase.LC03(+) And CaseProgress.CP04=LawCase.LC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And LC15=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFL', 'FCL', 'L', 'LIN') And NVL(Recpt,0) >0 " & _
               "Union  All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號, HC06 AS 案件名稱,'' as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,'' AS 申請國家,A0K32 " & _
               "  From CaseProgress, HireCase,CasePropertyMap,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=HireCase.HC01(+) and CaseProgress.CP02=HireCase.HC02(+) And CaseProgress.CP03=HireCase.HC03(+) And CaseProgress.CP04=HireCase.HC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('LA') And NVL(Recpt,0) >0 " & _
               "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號, NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32 " & _
               "  From CaseProgress, ServicePractice,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=ServicePractice.SP01(+) and CaseProgress.CP02=ServicePractice.SP02(+) And CaseProgress.CP03=ServicePractice.SP03(+) And CaseProgress.CP04=ServicePractice.SP04(+) And CP01=CPM01(+) And CP10=CPM02(+) And SP09=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 Not In ('P','CFP','CPS','CFT', 'FCT', 'T', 'TF','CFL', 'FCL', 'L', 'LIN','LA') And NVL(Recpt,0) >0 "
              
   'Modify by Amy 2013/05/27 增加排序 收文日,收文時間,總收文號
'   strSql = "Select SQLDATET2(CP05) as 收文日,Decode(PA23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號, NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,Decode(substr(CP60,1,1),'E',Decode(CP79,0,'收回',Decode(sign(CP75),1,'部分收回','未收')),CP60) CP60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
'               "  From CaseProgress, Patent,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "')  , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
'               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=Patent.PA01(+) and CaseProgress.CP02=Patent.PA02(+) And CaseProgress.CP03=Patent.PA03(+) And CaseProgress.CP04=Patent.PA04(+) And CP01=CPM01(+) And CP10=CPM02(+) And PA09=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('P','CFP','CPS') And NVL(Recpt,0) >0 " & _
'               "Union All Select SQLDATET2(CP05) as 收文日,Decode(TM28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(TM57,'')),null,'','●') AS 本所案號, NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
'               "  From CaseProgress, TradeMark,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "')  , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
'               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=TradeMark.TM01(+) and CaseProgress.CP02=TradeMark.TM02(+) And CaseProgress.CP03=TradeMark.TM03(+) And CaseProgress.CP04=TradeMark.TM04(+) And CP01=CPM01(+) And CP10=CPM02(+) And TM10=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFT', 'FCT', 'T', 'TF') And NVL(Recpt,0) >0"
'
'   strSql = strSql + "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,Nvl(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
'               "  From CaseProgress, LawCase,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
'               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=LawCase.LC01(+) and CaseProgress.CP02=LawCase.LC02(+) And CaseProgress.CP03=LawCase.LC03(+) And CaseProgress.CP04=LawCase.LC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And LC15=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFL', 'FCL', 'L', 'LIN') And NVL(Recpt,0) >0 " & _
'               "Union  All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號, HC06 AS 案件名稱,'' as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,'' AS 申請國家,A0K32,CP67,acc.CP09 " & _
'               "  From CaseProgress, HireCase,CasePropertyMap,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
'               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=HireCase.HC01(+) and CaseProgress.CP02=HireCase.HC02(+) And CaseProgress.CP03=HireCase.HC03(+) And CaseProgress.CP04=HireCase.HC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('LA') And NVL(Recpt,0) >0 " & _
'               "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號, NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
'               "  From CaseProgress, ServicePractice,CasePropertyMap,Nation,Acc0K0,( Select CP09,(CP18*1000-NVL(A1U07,0))/1000 RecPt From (Select CP09,CP18 From CaseProgress Where CP57 is null And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07 From CaseProgress,Acc1U0 Where CP57 is null And a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
'               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=ServicePractice.SP01(+) and CaseProgress.CP02=ServicePractice.SP02(+) And CaseProgress.CP03=ServicePractice.SP03(+) And CaseProgress.CP04=ServicePractice.SP04(+) And CP01=CPM01(+) And CP10=CPM02(+) And SP09=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP57 Is Null And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 Not In ('P','CFP','CPS','CFT', 'FCT', 'T', 'TF','CFL', 'FCL', 'L', 'LIN','LA') And NVL(Recpt,0) >0 " & _
'               " Order by 收文日,CP67,CP09"

   'Modify by Amy 2018/04/30 銷案不銷帳仍要計算,故拿掉cp57 is null -秀玲 P-122345 AA8013150
   'Modify By Sindy 2022/8/3 (CP18*1000-NVL(A1U07,0))/1000 => ((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000
   strSql = "Select SQLDATET2(CP05) as 收文日,Decode(PA23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號, NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,Decode(substr(CP60,1,1),'E',Decode(CP79,0,'收回',Decode(sign(CP75),1,'部分收回','未收')),CP60) CP60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
               "  From CaseProgress, Patent,CasePropertyMap,Nation,Acc0K0,( Select CP09,((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000 RecPt From (Select CP09,CP18,CP16,CP17 From CaseProgress Where cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "')  , (Select A1U03, A1U07, A1U09 From CaseProgress,Acc1U0 Where a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=Patent.PA01(+) and CaseProgress.CP02=Patent.PA02(+) And CaseProgress.CP03=Patent.PA03(+) And CaseProgress.CP04=Patent.PA04(+) And CP01=CPM01(+) And CP10=CPM02(+) And PA09=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('P','CFP','CPS') And NVL(Recpt,0) >0 " & _
               "Union All Select SQLDATET2(CP05) as 收文日,Decode(TM28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(TM57,'')),null,'','●') AS 本所案號, NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
               "  From CaseProgress, TradeMark,CasePropertyMap,Nation,Acc0K0,( Select CP09,((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000 RecPt From (Select CP09,CP18,CP16,CP17 From CaseProgress Where cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "')  , (Select A1U03, A1U07, A1U09 From CaseProgress,Acc1U0 Where a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=TradeMark.TM01(+) and CaseProgress.CP02=TradeMark.TM02(+) And CaseProgress.CP03=TradeMark.TM03(+) And CaseProgress.CP04=TradeMark.TM04(+) And CP01=CPM01(+) And CP10=CPM02(+) And TM10=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFT', 'FCT', 'T', 'TF') And NVL(Recpt,0) >0"
   
   strSql = strSql + "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 本所案號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,Nvl(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
               "  From CaseProgress, LawCase,CasePropertyMap,Nation,Acc0K0,( Select CP09,((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000 RecPt From (Select CP09,CP18,CP16,CP17 From CaseProgress Where cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07, A1U09 From CaseProgress,Acc1U0 Where a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=LawCase.LC01(+) and CaseProgress.CP02=LawCase.LC02(+) And CaseProgress.CP03=LawCase.LC03(+) And CaseProgress.CP04=LawCase.LC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And LC15=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('CFL', 'FCL', 'L', 'LIN','ACS') And NVL(Recpt,0) >0 "
   'Modify by Amy 2023/02/15 原:'' as 案件性質
   strSql = strSql + "Union  All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') AS 本所案號, HC06 AS 案件名稱,Nvl(cpm03,cpm04) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,'' AS 申請國家,A0K32,CP67,acc.CP09 " & _
               "  From CaseProgress, HireCase,CasePropertyMap,Acc0K0,( Select CP09,((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000 RecPt From (Select CP09,CP18,CP16,CP17 From CaseProgress Where cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07, A1U09 From CaseProgress,Acc1U0 Where a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=HireCase.HC01(+) and CaseProgress.CP02=HireCase.HC02(+) And CaseProgress.CP03=HireCase.HC03(+) And CaseProgress.CP04=HireCase.HC04(+) And CP01=CPM01(+) And CP10=CPM02(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 in('LA') And NVL(Recpt,0) >0 " & _
               "Union All Select SQLDATET2(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') AS 本所案號, NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,to_char(Acc.RecPt,'9990.00') AS 點數,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60, SQLDATET2(CP06) as 本所期限,SQLDATET2(CP27) as 發文日,NA03 AS 申請國家,A0K32,CP67,acc.CP09 " & _
               "  From CaseProgress, ServicePractice,CasePropertyMap,Nation,Acc0K0,( Select CP09,((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000 RecPt From (Select CP09,CP18,CP16,CP17 From CaseProgress Where cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') , (Select A1U03, A1U07, A1U09 From CaseProgress,Acc1U0 Where a1u03(+)=cp09 And a1u07>0 And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "') Where a1u03(+)=CP09 ) Acc " & _
               "  Where CaseProgress.CP09=Acc.CP09(+) And CaseProgress.CP01=ServicePractice.SP01(+) and CaseProgress.CP02=ServicePractice.SP02(+) And CaseProgress.CP03=ServicePractice.SP03(+) And CaseProgress.CP04=ServicePractice.SP04(+) And CP01=CPM01(+) And CP10=CPM02(+) And SP09=NA01(+) And CaseProgress.CP60=ACC0K0.A0K01(+) And CP12||''='" & Str01(2) & "' And cp05>=" & ChangeTStringToWString(Str01(3)) & " And cp05<=" & ChangeTStringToWString(Str01(4)) & " And CP13='" & Str01(1) & "' And CP01 Not In ('P','CFP','CPS','CFT', 'FCT', 'T', 'TF','CFL', 'FCL', 'L', 'LIN','ACS','LA') And NVL(Recpt,0) >0 " & _
               " Order by 收文日,CP67,CP09"
   
               
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
       Set grdDataList.Recordset = rsTmp
       SetDataListWidth
       Me.grdDataList.Visible = True
    For i = 1 To Me.grdDataList.Rows - 1
           grdDataList.row = i
           
           '收款情形
           IntTemp1 = 0
           IntTemp2 = 0
           Me.grdDataList.col = 5
           If Not IsNull(grdDataList.Text) And grdDataList.Text <> "" Then
               If Mid(grdDataList.Text, 1, 1) = "X" Then 'CP60開頭為X為外國請款
                   strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0 FROM ACC1K0 WHERE A1K01='" & grdDataList.Text & "'"
               
                  CheckOC2
                  adoRecordset1.CursorLocation = adUseClient
                  adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                      If Not IsNull(adoRecordset1.Fields(0)) Then
                          IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
                      End If
                      If Not IsNull(adoRecordset1.Fields(1)) Then
                          IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
                      End If
                      If Not IsNull(adoRecordset1.Fields(4)) Then
                          IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
                      End If
                      If Not IsNull(adoRecordset1.Fields(5)) Then
                          IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
                      End If
                      If IntTemp1 = IntTemp2 Then
                           grdDataList.Text = "收回"
                      Else
                           If IntTemp2 = 0 Then
                               grdDataList.Text = "未收"
                           Else
                               If IntTemp1 > IntTemp2 Then
                                   grdDataList.Text = "部分收回"
                               End If
                           End If
                       End If
                  End If
               End If
           End If
       Next i
      
   Else
       MsgBox ("無相關資料")
       tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End If
   
End Sub

Private Sub SetDataListWidth()
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "收文日"
   grdDataList.ColWidth(0) = 788
   grdDataList.col = 1: grdDataList.Text = "本所案號"
   grdDataList.ColWidth(1) = 1550
   grdDataList.col = 2: grdDataList.Text = "案件名稱"
   grdDataList.ColWidth(2) = 1300
   grdDataList.col = 3: grdDataList.Text = "案件性質"
   grdDataList.ColWidth(3) = 788
   grdDataList.col = 4: grdDataList.Text = "點數"
   grdDataList.ColWidth(4) = 600
   grdDataList.ColAlignment(4) = flexAlignRightCenter
   
   grdDataList.col = 5: grdDataList.Text = "收款情形"
   grdDataList.ColWidth(5) = 900
   grdDataList.col = 6: grdDataList.Text = "本所期限"
   grdDataList.ColWidth(6) = 788
   grdDataList.col = 7: grdDataList.Text = "發文日"
   grdDataList.ColWidth(7) = 788
   grdDataList.col = 8: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(8) = 788
   'Add by Amy 2013/05/23 增加AOK32
   grdDataList.col = 9: grdDataList.Text = "收據狀態"
   grdDataList.ColWidth(9) = 788
   'Add by Amy 2013/05/27 增加CP67和CP09
   grdDataList.col = 10: grdDataList.Text = "收文時間"
   grdDataList.ColWidth(10) = 0
   grdDataList.col = 11: grdDataList.Text = "收文號"
   grdDataList.ColWidth(11) = 0
End Sub
