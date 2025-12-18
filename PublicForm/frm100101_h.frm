VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_h 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利相關他國案"
   ClientHeight    =   5360
   ClientLeft      =   2930
   ClientTop       =   2510
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5360
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   420
      Index           =   4
      Left            =   7875
      TabIndex        =   5
      Top             =   30
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "下一筆"
      Height          =   420
      Index           =   3
      Left            =   6615
      TabIndex        =   4
      Top             =   30
      Width           =   1230
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印"
      Height          =   420
      Index           =   2
      Left            =   2595
      TabIndex        =   3
      Top             =   30
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "案件進度"
      Height          =   420
      Index           =   1
      Left            =   5340
      TabIndex        =   2
      Top             =   30
      Width           =   1260
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "案件基本資料"
      Height          =   420
      Index           =   0
      Left            =   3675
      TabIndex        =   1
      Top             =   30
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2985
      Left            =   45
      TabIndex        =   0
      Top             =   2340
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   5274
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   11
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   315
      Left            =   1170
      TabIndex        =   27
      Top             =   780
      Width           =   7035
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12409;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "黃色：相似案"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4200
      TabIndex        =   26
      Top             =   2055
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   5880
      TabIndex        =   25
      Top             =   2055
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專用期間："
      Height          =   180
      Index           =   9
      Left            =   150
      TabIndex        =   24
      Top             =   1995
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   1170
      TabIndex        =   23
      Top             =   1995
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3492;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利號數："
      Height          =   180
      Index           =   8
      Left            =   5415
      TabIndex        =   22
      Top             =   1725
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   7
      Left            =   6405
      TabIndex        =   21
      Top             =   1725
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3492;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公告號："
      Height          =   180
      Index           =   7
      Left            =   2805
      TabIndex        =   20
      Top             =   1725
      Width           =   720
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   3630
      TabIndex        =   19
      Top             =   1725
      Width           =   1695
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2990;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   18
      Top             =   1725
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1170
      TabIndex        =   17
      Top             =   1725
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2593;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1170
      TabIndex        =   16
      Top             =   1455
      Width           =   7080
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "12488;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   15
      Top             =   1455
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Index           =   4
      Left            =   4425
      TabIndex        =   14
      Top             =   1215
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   5370
      TabIndex        =   13
      Top             =   1200
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3492;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   12
      Top             =   1215
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1170
      TabIndex        =   11
      Top             =   1200
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3492;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   813
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日："
      Height          =   180
      Index           =   1
      Left            =   4485
      TabIndex        =   9
      Top             =   540
      Width           =   720
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   5325
      TabIndex        =   8
      Top             =   540
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3492;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      Top             =   540
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3492;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢案號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   540
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/24 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Combo1、lbl1(index) ; Printer列印未改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

'傳入的條件
Public KeyString As String
'要搜尋的種類
Public SearchKind As String
'紀錄作用按鍵
Public cmdState As Integer
Dim i As Integer, j As Integer
'edit by nickc 2007/01/23  少判斷第一年的  ' ',substr(lpad(pa72,200,' '),200,1),
'Private Const cntLstPayYearSQL As String = " decode( substr(lpad(pa72,200,' '),200,1),' ',' ',decode( substr(lpad(pa72,200,' '),199,1),',',substr(lpad(pa72,200,' '),200,1) ,decode( substr(lpad(pa72,200,' '),198,1),',',substr(lpad(pa72,200,' '),199,2) ,decode( substr(lpad(pa72,200,' '),197,1),',',substr(lpad(pa72,200,' '),198,3) ,decode( substr(lpad(pa72,200,' '),196,1),',',substr(lpad(pa72,200,' '),197,4) ) ) ) ) )"
'Private Const cntLstPayYearSQL2 As String = " decode( substr(lpad(p2.pa72,200,' '),200,1),' ',' ',decode( substr(lpad(p2.pa72,200,' '),199,1),',',substr(lpad(p2.pa72,200,' '),200,1) ,decode( substr(lpad(p2.pa72,200,' '),198,1),',',substr(lpad(p2.pa72,200,' '),199,2) ,decode( substr(lpad(p2.pa72,200,' '),197,1),',',substr(lpad(p2.pa72,200,' '),198,3) ,decode( substr(lpad(p2.pa72,200,' '),196,1),',',substr(lpad(p2.pa72,200,' '),197,4) ) ) ) ) )"
Private Const cntLstPayYearSQL As String = " decode( substr(lpad(pa72,200,' '),200,1),' ',' ',decode( substr(lpad(pa72,200,' '),199,1),' ',substr(lpad(pa72,200,' '),200,1),',',substr(lpad(pa72,200,' '),200,1) ,decode( substr(lpad(pa72,200,' '),198,1),',',substr(lpad(pa72,200,' '),199,2) ,decode( substr(lpad(pa72,200,' '),197,1),',',substr(lpad(pa72,200,' '),198,3) ,decode( substr(lpad(pa72,200,' '),196,1),',',substr(lpad(pa72,200,' '),197,4) ) ) ) ) )"
Private Const cntLstPayYearSQL2 As String = " decode( substr(lpad(p2.pa72,200,' '),200,1),' ',' ',decode( substr(lpad(p2.pa72,200,' '),199,1),' ',substr(lpad(pa72,200,' '),200,1),',',substr(lpad(p2.pa72,200,' '),200,1) ,decode( substr(lpad(p2.pa72,200,' '),198,1),',',substr(lpad(p2.pa72,200,' '),199,2) ,decode( substr(lpad(p2.pa72,200,' '),197,1),',',substr(lpad(p2.pa72,200,' '),198,3) ,decode( substr(lpad(p2.pa72,200,' '),196,1),',',substr(lpad(p2.pa72,200,' '),197,4) ) ) ) ) )"
Dim iPrint As Integer, Page As Integer
Dim PLeft(0 To 10) As Integer, StrLstPayYear As String, CaseNameCht As String
'Added by Lydia 2019/09/24
Dim strConTF As String '抓相似案的語法
Private Const cFixed As Integer = 6 '固定欄位
Dim colMemo As Integer '備註
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2024/2/7


'Add By Sindy 2024/2/7
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Sub StrMenu()
Select Case SearchKind
'Added by Lydia 2019/09/24 +相似案
Case "本所案號", "相似案"
   '有查詢權限，就可以印，因為印表權限，在別支有用過
   If CheckUse("frm100101_1", strFind, False) Then
       cmdok(2).Enabled = True
       cmdok(2).Visible = True
   Else
       cmdok(2).Enabled = False
       cmdok(2).Visible = False
   End If
   'Added by Lydia 2019/09/24
    If SearchKind = "相似案" Then
        lblMemo.Visible = True
    Else
        lblMemo.Visible = False
    End If
  
   StrMenu1
Case "客戶編號"
   '有查詢權限，就可以印，因為印表權限，在別支有用過
   If CheckUse("frm100102_1", strFind, False) Then
      cmdok(2).Enabled = True
      cmdok(2).Visible = True
   Else
      cmdok(2).Enabled = False
      cmdok(2).Visible = False
   End If
   StrMenu2
Case Else
End Select
End Sub

'以本所案號
Sub StrMenu1()
'Added by Lydia 2019/09/24
Dim intJ As Integer
Dim rs1 As New ADODB.Recordset

Me.Enabled = False
Screen.MousePointer = vbHourglass
'正常顯示
Me.grdDataList.Height = 3120
Me.grdDataList.Top = 2205
'基本資料
lbl1(0).Caption = KeyString
strSql = "select " & cntLstPayYearSQL & ",pa10,pa05,pa06,pa07,pa11,na03,na04,cu04,ptm03,ptm04,pa15,pa22,pa24,pa25 from patent,customer,nation,patenttrademarkmap where pa01='" & SystemNumber(KeyString, 1) & "' and pa02='" & SystemNumber(KeyString, 2) & "' and pa03='" & SystemNumber(KeyString, 3) & "' and pa04='" & SystemNumber(KeyString, 4) & "' and pa09=na01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and '1'=ptm01(+) and pa08=ptm02(+) "
StrLstPayYear = ""
CheckOC3
With AdoRecordSet3
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        lbl1(1) = ChangeWStringToTDateString("" & .Fields("pa10").Value)
        Combo1.Clear
        'Modified by Lydia 2021/12/24 +中、英、日
        Combo1.AddItem "中：" & .Fields("pa05").Value, 0
        Combo1.AddItem "英：" & .Fields("pa06").Value, 1
        'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
        Combo1.AddItem "外：" & .Fields("pa07").Value, 2
        Combo1.Text = "中：" & .Fields("pa05").Value
        'end 2021/12/24
        lbl1(2) = IIf(IsNull(.Fields("na03").Value), "" & .Fields("na04").Value, "" & .Fields("na03").Value)
        lbl1(3) = "" & .Fields("pa11").Value
        lbl1(4) = "" & .Fields("cu04").Value
        lbl1(5) = IIf(IsNull(.Fields("ptm03").Value), "" & .Fields("ptm04").Value, "" & .Fields("ptm03").Value)
        lbl1(6) = "" & .Fields("pa15").Value
        lbl1(7) = "" & .Fields("pa22").Value
        lbl1(8) = ChangeWStringToWDateString("" & .Fields("pa24").Value) & IIf(IsNull(.Fields("pa25").Value) And IsNull(.Fields("pa24").Value), "", " - ") & ChangeWStringToWDateString("" & .Fields("pa25").Value)
        StrLstPayYear = "" & .Fields(0).Value
    
    'Added by Morgan 2020/7/7 非專利案
    Else
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        SetDataListWidth
        ShowNoData
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    'end 2020/7/7
    
    End If
End With
CheckOC3
'add by nickc 2005/05/27 將所有關係都抓出來
Dim TmpRecCount As Long  '回傳比數
Dim TmpRecCount1 As Long  '回傳比數
Dim TmpRecCount2 As Long  '回傳比數
Dim TmpRecCount3 As Long  '回傳比數
Dim TmpRecCount4 As Long  '回傳比數
Dim TmpRecCount5 As Long  '回傳比數
'add by nickc 2005/06/07
Dim TmpRecCount6 As Long  '回傳比數
Dim TmpRecCount7 As Long  '回傳比數
'add by nickc 2006/06/20 加入分割
Dim TmpRecCount8 As Long  '回傳比數
Dim TmpRecCount9 As Long  '回傳比數
'add by nickc 2006/06/22 加入CaseRelation1
Dim TmpRecCount10 As Long  '回傳比數

Dim tmpCount As Integer '迴圈次
cnnConnection.Execute "delete from r100101_h where id='" & strUserNum & "' "
cnnConnection.Execute "insert into r100101_h select '" & SystemNumber(KeyString, 1) & "','" & SystemNumber(KeyString, 2) & "','" & SystemNumber(KeyString, 3) & "','" & SystemNumber(KeyString, 4) & "',0,'1','" & strUserNum & "' from dual "
cnnConnection.Execute "insert into r100101_h select '" & SystemNumber(KeyString, 1) & "','" & SystemNumber(KeyString, 2) & "','" & SystemNumber(KeyString, 3) & "','" & SystemNumber(KeyString, 4) & "',0,'2','" & strUserNum & "' from dual "
'edit by nickc 2007/09/17 改 proc
''''''''''''''TmpRecCount = 1
''''''''''''''tmpCount = 1
''''''''''''''Do While TmpRecCount <> 0
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount1
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount2
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount3
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount4
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cr01,cr02,cr03,cr04," & tmpCount & ",'2','" & strUserNum & "' from caserelation where cr05||cr06||cr07||cr08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cr01||cr02||cr03||cr04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount5
''''''''''''''   'add by nickc 2005/06/07
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount6
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount7
''''''''''''''   'edit by nickc 2005/06/07
''''''''''''''   'add by nickc 2006/06/20
''''''''''''''   cnnConnection.Execute "insert into r100101_h select dc01,dc02,dc03,dc04," & tmpCount & ",'3','" & strUserNum & "' from divisioncase where dc05||dc06||dc07||dc08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and dc01||dc02||dc03||dc04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount8
''''''''''''''   cnnConnection.Execute "insert into r100101_h select dc05,dc06,dc07,dc08," & tmpCount & ",'3','" & strUserNum & "' from divisioncase where dc01||dc02||dc03||dc04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and dc05||dc06||dc07||dc08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount9
''''''''''''''
'''''''''''''''edit by nickc 2006/11/21 專利處不用
''''''''''''''   'add  by nickc 2006/06/22
'''''''''''''''   cnnConnection.Execute "insert into r100101_h select cr01,cr02,cr03,cr04," & tmpCount & ",'2','" & strUserNum & "' from caserelation1 where cr05||cr06||cr07||cr08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cr01||cr02||cr03||cr04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount10
''''''''''''''
''''''''''''''   'edit by nickc 2006/06/20
''''''''''''''   'If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5) = 0 Then
''''''''''''''   'edit by nickc 2006/06/20
''''''''''''''   'If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7) = 0 Then
''''''''''''''   'edit by nickc 2006/06/22
''''''''''''''   'If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7 + TmpRecCount8 + TmpRecCount9) = 0 Then
''''''''''''''   If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7 + TmpRecCount8 + TmpRecCount9 + TmpRecCount10) = 0 Then
''''''''''''''      Exit Do
''''''''''''''   End If
''''''''''''''   tmpCount = tmpCount + 1
''''''''''''''Loop
 cnnConnection.Execute "begin   db_r100101_h('" & strUserNum & "'); end;"

'edit by nickc 2005/03/02  若國內外有，該組多國案不出來

'add by nickc 2006/06/20 所有有關聯的都出來，協理說的，包括分割
'edit by nickc 2006/08/25 加入銷卷
'strSQL = "select distinct '','',decode(pa23,'1','','N')||r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')," & SQLDate("pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),pa11,pa14,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & "," & cntLstPayYearSQL & ",decode(dc05||dc06||dc07||dc08,null,'','為 '||dc05||'-'||dc06||'-'||dc07||'-'||dc08|| ' 之分割案'),12 as bysort,'','',r001001||'-'||r001002||'-'||r001003||'-'||r001004||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A from patent,patenttrademarkmap,nation,r100101_h,divisioncase where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) and r001001=dc01(+) and r001002=dc02(+) and r001003=dc03(+) and r001004=dc04(+)"
'Modified by Morgan 2013/10/11 +pa08
'Modified by Lydia 2019/09/24 +別名
'strSql = "select distinct '','',replace(decode(pa23,'1','','N')||r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●'),'N---','')," & SQLDate("pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),pa11,pa15,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & "," & cntLstPayYearSQL & ",decode(dc05||dc06||dc07||dc08,null,'','為 '||dc05||'-'||dc06||'-'||dc07||'-'||dc08|| ' 之分割案'),r001005 as bysort,'','',r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A,pa08,pa09 from patent,patenttrademarkmap,nation,r100101_h,divisioncase where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) and r001001=dc01(+) and r001002=dc02(+) and r001003=dc03(+) and r001004=dc04(+)"
strSql = "select distinct '' as v01,'' as v02,replace(decode(pa23,'1','','N')||r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●'),'N---','') as v03" & _
             "," & SQLDate("pa10") & " as v04,nvl(ptm03,ptm04) as v05,nvl(na03,na04) as v06,pa11,pa15,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & " as v10 " & _
             "," & cntLstPayYearSQL & " as v11 ,nvl(decode(dc05||dc06||dc07||dc08,null,'','為 '||dc05||'-'||dc06||'-'||dc07||'-'||dc08|| ' 之分割案'),' ') as v12,r001005 as bysort,'' as v14,'' as v15,r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A,pa08,pa09 " & _
             "from patent,patenttrademarkmap,nation,r100101_h,divisioncase where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) and r001001=dc01(+) and r001002=dc02(+) and r001003=dc03(+) and r001004=dc04(+)"

'Added by Lydia 2019/09/24 抓相似案的資料
If SearchKind = "相似案" Then
     cnnConnection.Execute "delete from r100101_h3 where id ='" & strUserNum & "' "
     '逐筆抓相關案的相似案記錄
     strExc(1) = ""
     strExc(3) = ""
     strExc(0) = "SELECT * FROM R100101_H WHERE ID='" & strUserNum & "' AND (R001001='P' OR R001001='FCP') ORDER BY R001005, R001001,R001002 "
     intI = 1
     Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
     If intI = 1 Then
          rs1.MoveFirst
          Do While Not rs1.EOF
               If strExc(1) <> "" & rs1.Fields("R001001") & rs1.Fields("R001002") & rs1.Fields("R001003") & rs1.Fields("R001004") Then
                    intJ = intJ + 1
                    strExc(2) = ""
                    strConTF = "SELECT " & CNULL(strUserNum) & " as ID," & intJ & " as SEQNO, CP01,CP02,CP03,CP04,TF01,TF20,TF19 FROM CASEPROGRESS,TRANSFEE " & _
                                     "WHERE CP01='" & rs1.Fields("r001001") & "' AND CP02='" & rs1.Fields("r001002") & "'  AND CP03='" & rs1.Fields("r001003") & "'  AND CP04='" & rs1.Fields("r001004") & "'  AND CP159=0 AND CP09=TF01 "
                    strConTF = strConTF & "UNION " & _
                                     "SELECT " & CNULL(strUserNum) & " as ID," & intJ & " as SEQNO, CP01,CP02,CP03,CP04,TF01,TF20,TF19 FROM CASEPROGRESS,TRANSFEE " & _
                                     "WHERE CP09=TF01 AND CP159=0 AND TF20='" & rs1.Fields("R001001") & rs1.Fields("R001002") & rs1.Fields("R001003") & rs1.Fields("R001004") & "' "
                    'Modified by Lydia 2024/02/01 +TF20 is not null ; ex.FCP-070935先收新案翻譯後收其他翻譯
                    strConTF = strConTF & "UNION " & _
                                     "SELECT " & CNULL(strUserNum) & " as ID," & intJ & " as SEQNO, CP01,CP02,CP03,CP04,TF01,TF20,TF19 FROM CASEPROGRESS,TRANSFEE " & _
                                     "WHERE CP09=TF01 AND TF20 IN (SELECT TF20 FROM CASEPROGRESS,TRANSFEE WHERE CP159=0 AND CP09=TF01 AND TF20 is not null AND CP01='" & rs1.Fields("r001001") & "' AND CP02='" & rs1.Fields("r001002") & "'  AND CP03='" & rs1.Fields("r001003") & "'  AND CP04='" & rs1.Fields("r001004") & "' ) "
                    'Modified by Lydia 2024/02/01 +TF20 is not null
                    strConTF = strConTF & "UNION " & _
                                     "SELECT " & CNULL(strUserNum) & " as ID," & intJ & " as SEQNO, CP01,CP02,CP03,CP04,TF01,TF20,TF19 FROM CASEPROGRESS,TRANSFEE " & _
                                     "WHERE CP09=TF01 AND CP01||CP02||CP03||CP04 = (SELECT TF20 FROM CASEPROGRESS,TRANSFEE WHERE CP159=0 AND CP09=TF01 AND TF20 is not null AND CP01='" & rs1.Fields("r001001") & "' AND CP02='" & rs1.Fields("r001002") & "'  AND CP03='" & rs1.Fields("r001003") & "'  AND CP04='" & rs1.Fields("r001004") & "' ) "
                    cnnConnection.Execute "INSERT INTO R100101_H3 (ID,SEQNO,R001001,R001002,R001003,R001004,R001005,R001006,R001007) " & strConTF, intI
                    If intI = 1 Then '排除沒有相似案
                        strExc(2) = "delete from r100101_h3 where id ='" & strUserNum & "' and seqno=" & intJ
                    Else  '增加空白列->區隔相關案的不同相似案
                        strExc(2) = "INSERT INTO R100101_H3 (ID,SEQNO,R001001,R001002,R001003,R001004,R001005,R001006,R001007) values (" & _
                                         CNULL(strUserNum) & ", " & intJ & ", null, null, null, null, null, null, null) "
                        strExc(3) = strExc(3) & rs1.Fields("R001001") & rs1.Fields("R001002") & rs1.Fields("R001003") & rs1.Fields("R001004") & ","
                    End If
                    If strExc(2) <> "" Then cnnConnection.Execute strExc(2)
               End If
               strExc(1) = "" & rs1.Fields("R001001") & rs1.Fields("R001002") & rs1.Fields("R001003") & rs1.Fields("R001004")
               rs1.MoveNext
          Loop
          If strExc(3) <> "" Then
                strSql = strSql & " Union " & _
                            "select distinct '' as v01,'' as v02,replace(decode(pa23,'1','','N')||r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●'),'N---','') as v03" & _
                             "," & SQLDate("pa10") & " as v04,nvl(ptm03,ptm04) as v05,nvl(na03,na04) as v06,pa11,pa15,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & " as v10 " & _
                             "," & cntLstPayYearSQL & " as v11 ,nvl(decode(r001006,null,decode(r001001,null,'','相似母案'),'相似案號：'||r001006||'，'||'相似度：'||r001007),' ') as v12,90+Seqno+decode(r001001,null,0.5,0) as bysort,'' as v14,'' as v15,r001001||'-'||r001002||'-'||r001003||'-'||r001004||decode(pa57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A,pa08,pa09 " & _
                             "from patent,patenttrademarkmap,nation,r100101_h3 where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+) and id='" & strUserNum & "'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) "
          End If
     End If
End If
'end 2019/09/24

strSql = strSql & " order by bysort,A "

CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    grdDataList.FixedCols = 0 'Added Lydia 2019/09/24
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount <> 1 Then
        Set grdDataList.Recordset = adoRecordset
        SetDataListWidth
        grdDataList.FixedCols = cFixed 'Added by Lydia 2019/09/24 固定欄位
        CheckDesign 'Added by Morgan 2013/10/11 'Memo by Lydia 2019/09/24 一併將欄位底色改為空白
    Else
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        SetDataListWidth
        ShowNoData
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
End With
Screen.MousePointer = vbDefault
Me.Enabled = True
End Sub
Private Sub SetDataListWidth()
   If grdDataList.Cols < 12 Then grdDataList.Cols = 12
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.col = 1: grdDataList.Text = "狀態"
   grdDataList.ColWidth(1) = 0
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "本所案號"
   'Modified by Morgan 2013/10/11
   'GrdDataList.ColWidth(2) = 1500
   grdDataList.ColWidth(2) = 1700
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "申請日"
   grdDataList.ColWidth(3) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "種類"
   grdDataList.ColWidth(4) = 700
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(5) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "申請案號"
   grdDataList.ColWidth(6) = 1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 7: grdDataList.Text = "公告號"
   grdDataList.ColWidth(7) = 1250
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 8: grdDataList.Text = "專利號數"
   grdDataList.ColWidth(8) = 1250
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 9: grdDataList.Text = "專用期間"
   grdDataList.ColWidth(9) = 2000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 10: grdDataList.Text = "最後已繳年度"
   grdDataList.ColWidth(10) = 1250
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 11: grdDataList.Text = "備註"
   colMemo = 11 'Added by Lydia 2019/09/24
   
   'Modified by Lydia 2019/09/24 2500=>3200
   grdDataList.ColWidth(11) = 3200
   grdDataList.CellAlignment = flexAlignCenterCenter
   For intI = 12 To grdDataList.Cols - 1
      grdDataList.ColWidth(intI) = 0
   Next
End Sub


'以客戶編號
Sub StrMenu2()
Me.Enabled = False
Screen.MousePointer = vbHourglass
'不秀基本資料
Me.grdDataList.Height = 4830
Me.grdDataList.Top = 495
CheckOC3
'edit by nickc 2005/03/02 若國內外有，該組多國案不出來
'StrSql = "select distinct '','國外',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'StrSql = StrSql & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'StrSql = StrSql & " union select '','國外',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'StrSql = StrSql & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'StrSql = StrSql & " union select '','國外',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'StrSql = StrSql & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'StrSql = StrSql & " union select '','國外',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'StrSql = StrSql & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'StrSql = StrSql & " union select '','國外',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'StrSql = StrSql & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'StrSql = StrSql & "union select '','國內',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'StrSql = StrSql & "union select '','國內',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'StrSql = StrSql & "union select '','國內',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'StrSql = StrSql & "union select '','國內',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'StrSql = StrSql & "union select '','國內',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,11 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
''edit by nick 2005/02/18 改一組都出來
''StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
''                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''                            " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa26='" & KeyString & "' "
''StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa26='" & KeyString & "' "
''StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
''                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa27='" & KeyString & "' "
''StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa27='" & KeyString & "' "
''StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
''                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa28='" & KeyString & "' "
''StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa28='" & KeyString & "' "
''StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
''                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa29='" & KeyString & "' "
''StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa29='" & KeyString & "' "
''StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
''                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa30='" & KeyString & "' "
''StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress " & _
''         " Where c2.cr01 = c1.cr01 And c2.cr02 = c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c2.cr05=cp01 and c2.cr06=cp02 and c2.cr07=cp03 and c2.cr08=cp04 and cp21='Y') and p1.pa30='" & KeyString & "' "
'
'StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'                            " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa26='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa26='" & KeyString & "' "
'StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa27='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa27='" & KeyString & "' "
'StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa28='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa28='" & KeyString & "' "
'StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa29='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa29='" & KeyString & "' "
'StrSql = StrSql & "union select '','多國案',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa30='" & KeyString & "' "
'StrSql = StrSql & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa30='" & KeyString & "' "
'strSQL = "select distinct '','',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & " union select '','',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & " union select '','',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & " union select '','',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & " union select '','',cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'strSQL = strSQL & " union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'edit by nickc 2005/05/05
'strSQL = strSQL & "union select '','',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & "union select '','',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & "union select '','',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & "union select '','',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & "union select '','',cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'
'
'strSQL = strSQL & "union select '','',cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'
'strSQL = strSQL & "union select '','',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'                            " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa26='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa26='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa26='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa26='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa26='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa26='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa27='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa27='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa27='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa27='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa27='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa27='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa28='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa28='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa28='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa28='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa28='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa28='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa29='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa29='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa29='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa29='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa29='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa29='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa30='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa30='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa30='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','' as A,null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa30='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa30='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa30='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'edit by nickc 2005/05/27
'strSQL = "select distinct '','',decode(p2.pa23,'1','','N')||cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & " union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & " union select '','',decode(p2.pa23,'1','','N')||cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & " union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & " union select '','',decode(p2.pa23,'1','','N')||cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & " union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & " union select '','',decode(p2.pa23,'1','','N')||cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & " union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & " union select '','',decode(p2.pa23,'1','','N')||cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'strSQL = strSQL & " union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm01=p2.pa01(+) and cm02=p2.pa02(+) and cm03=p2.pa03(+) and cm04=p2.pa04(+) and cm10='0' and p1.pa01=cm05(+) and p1.pa02=cm06(+) and p1.pa03=cm07(+) and p1.pa04=cm08(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08 AS bysort2,p2.pa05,cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(p2.pa57,'Y','＊','') as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm05||'-'||cm06||'-'||cm07||'-'||cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap,patenttrademarkmap,nation,patent p1 where cm05=p2.pa01(+) and cm06=p2.pa02(+) and cm07=p2.pa03(+) and cm08=p2.pa04(+) and cm10='0' and p1.pa01=cm01(+) and p1.pa02=cm02(+) and p1.pa03=cm03(+) and p1.pa04=cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' "
'
''edit by nickc 2005/05/05 移除，薛跟秀玲說應該不用  edit by nick 2005/05/12 取消
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05,cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A" & _
'                               " from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa26='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05,cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A" & _
'                               " from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa27='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05,cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A" & _
'                               " from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa28='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05,cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A" & _
'                               " from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa29='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08 AS bysort2,p2.pa05,cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04||decode(p2.pa57,'Y','＊','') as A" & _
'                               " from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,12 as bysort,cm1.cm05||'-'||cm1.cm06||'-'||cm1.cm07||'-'||cm1.cm08||'Z' AS bysort2,null,'' as A  from patent p2,casemap cm1,patenttrademarkmap,nation,patent p1,casemap cm2 where cm2.cm01=p2.pa01(+) and cm2.cm02=p2.pa02(+) and cm2.cm03=p2.pa03(+) and cm2.cm04=p2.pa04(+) and cm1.cm10='0' and p1.pa01=cm1.cm01(+) and p1.pa02=cm1.cm02(+) and p1.pa03=cm1.cm03(+) and p1.pa04=cm1.cm04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and p1.pa30='" & KeyString & "' and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) "
'
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05,c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'                            " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa26='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa26='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa26='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null,'' as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa26='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa26='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa26='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05,c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa27='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa27='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa27='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null,'' as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa27='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa27='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa27='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05,c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa28='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa28='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa28='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null,'' as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa28='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa28='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa28='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05,c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa29='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa29='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa29='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null,'' as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa29='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa29='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa29='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','',decode(p2.pa23,'1','','N')||c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','')," & SQLDate("p2.pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),p2.pa11,p2.pa13,p2.pa22," & SQLDate("p2.pa24", False) & "||decode(p2.pa24,null,decode(p2.pa25,null,'','-'),'-')||" & SQLDate("p2.pa25", False) & "," & cntLstPayYearSQL2 & ",20 as bysort,d1.cr99 AS bysort2,p2.pa05,c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(p2.pa57,'Y','＊','') as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1," & _
'                            " patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa30='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa30='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa30='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & "union select '','','',null,null,null,null,null,null,null,null,20 as bysort,d1.cr99||'Z' AS bysort2,null,'' as A  from patent p2,caserelation c1,(select c3.cr01,c3.cr02,c3.cr03,c3.cr04,max(c3.cr05||'-'||c3.cr06||'-'||c3.cr07||'-'||c3.cr08) as cr99 from (select c6.cr01,c6.cr02,c6.cr03,c6.cr04,c6.cr05,c6.cr06,c6.cr07,c6.cr08 from caserelation c6 union select c7.cr05 as cr01,c7.cr06 as cr02,c7.cr07 as cr03,c7.cr08 as cr04,c7.cr05,c7.cr06,c7.cr07,c7.cr08 from caserelation c7) c3 group by c3.cr01,c3.cr02,c3.cr03,c3.cr04) D1,patenttrademarkmap,nation,patent p1 where c1.cr05=p2.pa01(+) and c1.cr06=p2.pa02(+) and c1.cr07=p2.pa03(+) and c1.cr08=p2.pa04(+) and p1.pa01=d1.cr01(+) and p1.pa02=d1.cr02(+) and p1.pa03=d1.cr03(+) and p1.pa04=d1.cr04(+) and p1.pa01=c1.cr01(+) and p1.pa02=c1.cr02(+) and p1.pa03=c1.cr03(+) and p1.pa04=c1.cr04(+) and '1'=ptm01(+) and p2.pa08=ptm02(+) and p2.pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2 " & _
'         " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y')) and p1.pa30='" & KeyString & "' and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select cm01||cm02||cm03||cm04 from casemap,patent where  pa30='" & KeyString & "' and pa01=cm05(+) and pa02=cm06(+) and pa03=cm07(+) and pa04=cm08(+) and cm10='0' union select cm2.cm01||'-'||cm2.cm02||'-'||cm2.cm03||'-'||cm2.cm04 from casemap cm1,casemap cm2,patent where cm1.cm10='0' and pa30='" & KeyString & "' and pa01=cm1.cm01(+) and pa02=cm1.cm02(+) and pa03=cm1.cm03(+) and pa04=cm1.cm04(+) " & _
'                           " and cm1.cm05=cm2.cm05(+) and cm1.cm06=cm2.cm06(+) and cm1.cm07=cm2.cm07(+) and cm1.cm08=cm2.cm08(+) and '0'=cm2.cm10(+) ) "
'strSQL = strSQL & " order by bysort,bysort2,A "
'edit by nickc 2005/05/27 end

Dim TmpRecCount As Long  '回傳比數
Dim TmpRecCount1 As Long  '回傳比數
Dim TmpRecCount2 As Long  '回傳比數
Dim TmpRecCount3 As Long  '回傳比數
Dim TmpRecCount4 As Long  '回傳比數
Dim TmpRecCount5 As Long  '回傳比數
'add by nickc 2005/06/07
Dim TmpRecCount6 As Long  '回傳比數
Dim TmpRecCount7 As Long  '回傳比數
'add by nickc 2006/06/20 加入分割
Dim TmpRecCount8 As Long  '回傳比數
Dim TmpRecCount9 As Long  '回傳比數
'add by nickc 2006/06/22 加入CaseRelation1
Dim TmpRecCount10 As Long  '回傳比數
Dim tmpCount As Integer '迴圈次

'Add By Sindy 2011/01/03 檢查國內外權限
If CheckSR12(KeyString) = False Then
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

cnnConnection.Execute "delete from r100101_h where id='" & strUserNum & "' "
strSql = "select distinct pa01,pa02,pa03,pa04,0,'1','" & strUserNum & "' from patent where  pa26='" & KeyString & "' "
strSql = strSql & " union select pa01,pa02,pa03,pa04,0,'1','" & strUserNum & "' from patent where pa27='" & KeyString & "' "
strSql = strSql & " union select pa01,pa02,pa03,pa04,0,'1','" & strUserNum & "' from patent where pa28='" & KeyString & "' "
strSql = strSql & " union select pa01,pa02,pa03,pa04,0,'1','" & strUserNum & "' from patent where pa29='" & KeyString & "' "
strSql = strSql & " union select pa01,pa02,pa03,pa04,0,'1','" & strUserNum & "' from patent where pa30='" & KeyString & "' "
cnnConnection.Execute "insert into r100101_h (" & strSql & ") "
'cnnConnection.Execute "insert into r100101_h select r001001,r001002,r001003,r001004,r001005,'2',id from r100101_h where id='" & strUserNum & "' "
'edit by nickc 2007/09/17 改 PROC
'''''''''''''TmpRecCount = 1
'''''''''''''tmpCount = 1
'''''''''''''Do While TmpRecCount <> 0
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount1
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount2
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount3
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount4
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cr01,cr02,cr03,cr04," & tmpCount & ",'1','" & strUserNum & "' from caserelation where cr05||cr06||cr07||cr08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cr01||cr02||cr03||cr04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount5
'''''''''''''   'add by nickc 2005/06/07
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount6
'''''''''''''   cnnConnection.Execute "insert into r100101_h select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount7
'''''''''''''   'edit by nickc 2005/06/07
'''''''''''''   'add by nickc 2006/06/20
'''''''''''''   cnnConnection.Execute "insert into r100101_h select dc01,dc02,dc03,dc04," & tmpCount & ",'3','" & strUserNum & "' from divisioncase where dc05||dc06||dc07||dc08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and dc01||dc02||dc03||dc04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount8
'''''''''''''   cnnConnection.Execute "insert into r100101_h select dc05,dc06,dc07,dc08," & tmpCount & ",'3','" & strUserNum & "' from divisioncase where dc01||dc02||dc03||dc04 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and dc05||dc06||dc07||dc08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount9
''''''''''''''edit by nickc 2006/11/21 專利處不用
'''''''''''''   'add by nickc 2006/06/22
''''''''''''''   cnnConnection.Execute "insert into r100101_h select cr01,cr02,cr03,cr04," & tmpCount & ",'1','" & strUserNum & "' from caserelation1 where cr05||cr06||cr07||cr08 in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') and cr01||cr02||cr03||cr04 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "') ", TmpRecCount10
'''''''''''''
'''''''''''''   'edit by nickc 2006/06/20
'''''''''''''   'If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7) = 0 Then
'''''''''''''   'edit by nickc 2006/06/22
'''''''''''''   'If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7 + TmpRecCount8 + TmpRecCount9) = 0 Then
'''''''''''''   If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7 + TmpRecCount8 + TmpRecCount9 + TmpRecCount10) = 0 Then
'''''''''''''      Exit Do
'''''''''''''   End If
'''''''''''''   tmpCount = tmpCount + 1
'''''''''''''Loop
''''''''''''''add by nickc 2005/05/31 分類
'''''''''''''cnnConnection.Execute "update r100101_h set r001005=0 where id='" & strUserNum & "' "
'''''''''''''Dim IsDataOK As Boolean
'''''''''''''Dim GroupCount As Integer
'''''''''''''Dim ChkDataRs As New ADODB.Recordset
''''''''''''''edit by nickc 2006/06/20 整段不做
'''''''''''''IsDataOK = False
'''''''''''''GroupCount = 10

'''''''''''''Do While IsDataOK = False
'''''''''''''      strSQL = "select * from r100101_h where id='" & strUserNum & "' and r001005=0  "
'''''''''''''      CheckOC
'''''''''''''      With adoRecordset
'''''''''''''          .CursorLocation = adUseClient
'''''''''''''          .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'''''''''''''          If .RecordCount <> 0 Then
'''''''''''''            .MoveFirst
'''''''''''''               cnnConnection.Execute "delete from r100101_h2 where id='" & strUserNum & "' "
'''''''''''''               cnnConnection.Execute "insert into r100101_h2 select '" & CheckStr(.Fields("R001001")) & "','" & CheckStr(.Fields("R001002")) & "','" & CheckStr(.Fields("R001003")) & "','" & CheckStr(.Fields("R001004")) & "',0,'1','" & strUserNum & "' from dual "
''''''''''''''               cnnConnection.Execute "insert into r100101_h2 select '" & CheckStr(.Fields("R001001")) & "','" & CheckStr(.Fields("R001002")) & "','" & CheckStr(.Fields("R001003")) & "','" & CheckStr(.Fields("R001004")) & "',0,'2','" & strUserNum & "' from dual "
'''''''''''''               TmpRecCount = 1
'''''''''''''               tmpCount = 1
''''''''''''''               Do While TmpRecCount <> 0
''''''''''''''                  Set ChkDataRs = New ADODB.Recordset
''''''''''''''                  If ChkDataRs.State = 1 Then ChkDataRs.Close
''''''''''''''                  ChkDataRs.CursorLocation = adUseClient
''''''''''''''                  ChkDataRs.Open "select * from r100101_h where id='" & strUserNum & "' and r001005=0  and r001001='" & CheckStr(.Fields("r001001").Value) & "' and r001002='" & CheckStr(.Fields("R001002")) & "' and r001003='" & CheckStr(.Fields("R001003")) & "' and r001004='" & CheckStr(.Fields("R001004")) & "'  ", cnnConnection, adOpenStatic, adLockReadOnly
''''''''''''''                  If ChkDataRs.RecordCount <> 0 Then
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount1
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount2
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='0' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount3
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='3' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount4
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cr01,cr02,cr03,cr04," & tmpCount & ",'1','" & strUserNum & "' from caserelation where cr05||cr06||cr07||cr08 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cr01||cr02||cr03||cr04 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount5
'''''''''''''                     'add by nickc 2005/06/07
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cm01,cm02,cm03,cm04," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm05||cm06||cm07||cm08 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cm01||cm02||cm03||cm04 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount6
'''''''''''''                     cnnConnection.Execute "insert into r100101_h2 select cm05,cm06,cm07,cm08," & tmpCount & ",'1','" & strUserNum & "' from casemap where cm10='4' and cm01||cm02||cm03||cm04 in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') and cm05||cm06||cm07||cm08 not in (select r001001||r001002||r001003||r001004 from r100101_h2 where  id='" & strUserNum & "') ", TmpRecCount7
'''''''''''''                     'add by nickc 2005/10/26 刪除沒有相關案的
''''''''''''''                     If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount6 + TmpRecCount7) = 0 Then
''''''''''''''                        cnnConnection.Execute "delete r100101_h where  r001001||r001002||r001003||r001004 in (select distinct r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "' and r001006='1') and id='" & strUserNum & "' and r001006='1' and R001005='0' "
''''''''''''''                     End If
''''''''''''''                     'add by nickc 2005/10/26 刪除沒有相關案的
''''''''''''''                     If (TmpRecCount5) = 0 Then
''''''''''''''                        cnnConnection.Execute "delete r100101_h where  r001001||r001002||r001003||r001004 in (select distinct r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "' and r001006='1') and id='" & strUserNum & "' and r001006='1' and R001005='0' "
''''''''''''''                     End If
'''''''''''''                     'edit by nickc 2005/06/07
'''''''''''''                     'If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5) = 0 Then
'''''''''''''                     'add by nickc 2005/10/26 刪除沒有相關案的
'''''''''''''                     If (TmpRecCount1 + TmpRecCount2 + TmpRecCount3 + TmpRecCount4 + TmpRecCount5 + TmpRecCount6 + TmpRecCount7) = 0 Then
''''''''''''''                        Exit Do
'''''''''''''                         cnnConnection.Execute "delete r100101_h where  r001001||r001002||r001003||r001004 in (select distinct r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "' and r001006='1') and id='" & strUserNum & "' and r001006='1' and R001005='0' "
'''''''''''''                     Else
'''''''''''''                        cnnConnection.Execute "update r100101_h set r001005=(select nvl(min(r001005),0) from r100101_h " & _
'''''''''''''                                                            " where r001001||r001002||r001003||r001004 in (select cm01||cm02||cm03||cm04 from casemap where cm05||cm06||cm07||cm08 in ( " & _
'''''''''''''                                                            " select r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'''''''''''''                                                            " union select cm05||cm06||cm07||cm08 from casemap where cm01||cm02||cm03||cm04 in ( " & _
'''''''''''''                                                            " select r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'''''''''''''                                                            " union select cr01||cr02||cr03||cr04 from caserelation where cr05||cr06||cr07||cr08 in (" & _
'''''''''''''                                                            " select r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "')) and r001005<>0 and id='" & strUserNum & "') " & _
'''''''''''''                                                            " where r001001||r001002||r001003||r001004 in (select cm01||cm02||cm03||cm04 from casemap where cm05||cm06||cm07||cm08 in ( " & _
'''''''''''''                                                            " select r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'''''''''''''                                                            " union select cm05||cm06||cm07||cm08 from casemap where cm01||cm02||cm03||cm04 in (" & _
'''''''''''''                                                            " select r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "') and cm10 in ('0','3','4') " & _
'''''''''''''                                                            " union select cr01||cr02||cr03||cr04 from caserelation where cr05||cr06||cr07||cr08 in (" & _
'''''''''''''                                                            " select r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "') ) and id='" & strUserNum & "' "
'''''''''''''                        TmpRecCount7 = 0
'''''''''''''                        cnnConnection.Execute "update r100101_h set r001005=" & GroupCount & " where  r001001||r001002||r001003||r001004 in (select distinct r001001||r001002||r001003||r001004 from r100101_h2 where id='" & strUserNum & "' ) and id='" & strUserNum & "' and r001005='0' ", TmpRecCount7
'''''''''''''                        If TmpRecCount7 <> 0 Then
'''''''''''''                           cnnConnection.Execute "insert into r100101_h (r001005,id) values (" & GroupCount + 1 & ",'" & strUserNum & "') "
'''''''''''''                        End If
'''''''''''''                        GroupCount = GroupCount + 10
'''''''''''''                     End If
'''''''''''''                     tmpCount = tmpCount + 1
'''''''''''''                     cnnConnection.Execute "delete r100101_h where id='" & strUserNum & "' and r001005=0  and r001001='" & CheckStr(.Fields("r001001").Value) & "' and r001002='" & CheckStr(.Fields("R001002")) & "' and r001003='" & CheckStr(.Fields("R001003")) & "' and r001004='" & CheckStr(.Fields("R001004")) & "'  "
''''''''''''''               Loop
'''''''''''''          Else
'''''''''''''               IsDataOK = True
'''''''''''''          End If
'''''''''''''      End With
'''''''''''''Loop
 cnnConnection.Execute "begin   db_r100101_h('" & strUserNum & "'); end;"

'edit by nickc 2006/06/20 所有相關聯的都要出來 協理說的，包括分割
'edit by nickc 2007/09/17 使用 PROC
'strSQL = "select distinct '','',decode(pa23,'1','','N')||cm01||'-'||cm02||'-'||cm03||'-'||cm04||decode(pa57,'Y','＊','')," & SQLDate("pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),pa11,pa14,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & "," & cntLstPayYearSQL & ",r001005 as bysort,'',pa05,cm01||'-'||cm02||'-'||cm03||'-'||cm04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A from patent,casemap,patenttrademarkmap,nation,r100101_h where cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and (cm10='0' or cm10='3' or cm10='4') and r001001=cm05(+) and r001002=cm06(+) and r001003=cm07(+) and r001004=cm08(+) and id='" & strUserNum & "' and r001006='1' and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) "
'strSQL = strSQL & "union select '','',decode(pa23,'1','','N')||cm05||'-'||cm06||'-'||cm07||'-'||cm08||decode(pa57,'Y','＊','')," & SQLDate("pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),pa11,pa14,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & "," & cntLstPayYearSQL & ",r001005 as bysort,'',pa05,cm05||'-'||cm06||'-'||cm07||'-'||cm08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A from patent,casemap,patenttrademarkmap,nation,r100101_h where cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) and (cm10='0' or cm10='3' or cm10='4') and r001001=cm01(+) and r001002=cm02(+) and r001003=cm03(+) and r001004=cm04(+) and id='" & strUserNum & "' and r001006='1' and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) "
'strSQL = strSQL & "union select '','',decode(pa23,'1','','N')||c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||decode(pa57,'Y','＊','')," & SQLDate("pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),pa11,pa14,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & "," & cntLstPayYearSQL & ",r001005 as bysort,'',pa05,c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A from r100101_h,patent,caserelation c1,patenttrademarkmap,nation where c1.cr05=pa01(+) and c1.cr06=pa02(+) and c1.cr07=pa03(+) and c1.cr08=pa04(+) and r001001=c1.cr01(+) and r001002=c1.cr02(+) and r001003=c1.cr03(+) and r001004=c1.cr04(+) and  id='" & strUserNum & "' and r001006='1'  and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) and exists (select * from caserelation C2,caseprogress cc1,caseprogress cc2  " & _
'                            " Where c2.cr01 = c1.cr01 And c2.cr02= c1.cr02 And c2.cr03 = c1.cr03 And c2.cr04 = c1.cr04 and c1.cr01=cc2.cp01 and c1.cr02=cc2.cp02 and c1.cr03=cc2.cp03 and c1.cr04=cc2.cp04 and c2.cr05=cc1.cp01 and c2.cr06=cc1.cp02 and c2.cr07=cc1.cp03 and c2.cr08=cc1.cp04 and (cc1.cp21='Y' or cc2.cp21='Y'))  and c1.cr05||'-'||c1.cr06||'-'||c1.cr07||'-'||c1.cr08 not in (select r001001||r001002||r001003||r001004 from r100101_h where  id='" & strUserNum & "' and r001006='1' ) "
'strSQL = strSQL & " union select '','','','','',null,null,null,null,null,null,r001005 as bysort,'','','' as A from r100101_h where  id='" & strUserNum & "' and r001006 is null  "
'Modified by Morgan 2013/10/11 +pa08
strSql = "select distinct '','',replace(decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04||decode(pa57,'Y','＊',''),'N---','')," & SQLDate("pa10") & ",nvl(ptm03,ptm04),nvl(na03,na04),pa11,pa15,pa22," & SQLDate("pa24", False) & "||decode(pa24,null,decode(pa25,null,'','-'),'-')||" & SQLDate("pa25", False) & "," & cntLstPayYearSQL & ",'',r001005 as bysort,'',pa05,pa01||'-'||pa02||'-'||pa03||'-'||pa04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as A,pa08,pa09 from patent,patenttrademarkmap,nation,r100101_h where r001001=pa01(+) and r001002=pa02(+) and r001003=pa03(+) and r001004=pa04(+)  and id='" & strUserNum & "' and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) order by bysort,A "

CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.FixedCols = 0 'Added by Lydia 2019/09/24
    If .RecordCount <> 0 Then
        Set grdDataList.Recordset = adoRecordset
        SetDataListWidth
        grdDataList.FixedCols = cFixed 'Added by Lydia 2019/09/24 固定欄位
        CheckDesign 'Added by Morgan 2013/10/11  'Memo by Lydia 2019/09/24 一併將欄位底色改為空白
    Else
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        SetDataListWidth
        ShowNoData
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
End With
Screen.MousePointer = vbDefault
Me.Enabled = True
End Sub

Private Sub cmdOK_Click(index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
SetDataListWidth
cmdState = -1

lblMemo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_PrevForm = Nothing 'Add By Sindy 2024/2/7
Set frm100101_h = Nothing
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim stCaseNo As String 'Added by Morgan 2013/10/11
Dim lngColor As Long 'Added by Lydia 2019/09/24

Select Case cmdState
Case 0 '案件基本資料
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        Dim Str01 As String
        grdDataList.col = 0
        grdDataList.Text = ""
        'Added by Lydia 2019/09/24  預設底色
        If Trim("" & grdDataList.TextMatrix(i, colMemo)) <> "" And Left("" & grdDataList.TextMatrix(i, colMemo), 2) = "相似" Then
            lngColor = &HFFFF&
        Else
            lngColor = QBColor(15)
        End If
        'end 2019/09/24
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            'Modified by Lydia 2019/09/24
            'grdDataList.CellBackColor = QBColor(15)
            grdDataList.CellBackColor = lngColor
        Next j
        grdDataList.col = 2
        stCaseNo = RplStrNew(grdDataList.Text)

        Str01 = SystemNumber(grdDataList, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Select Case Pub_RplStr(Str01)
                Case "CFP", "FCP", "P"   '專利
                      Screen.MousePointer = vbHourglass
                      frm100101_3.Show
                      frm100101_3.Tag = stCaseNo
                      frm100101_3.StrMenu
                      Screen.MousePointer = vbDefault
                Case "CFT", "FCT", "T", "TF"   '商標
                      Screen.MousePointer = vbHourglass
                      frm100101_4.Show
                      frm100101_4.Tag = stCaseNo
                      frm100101_4.StrMenu
                      Screen.MousePointer = vbDefault
                'Modify By Sindy 2009/07/24 增加LIN系統類別
                'modify by sonia 2019/7/29 +ACS系統類別
                Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
                      Screen.MousePointer = vbHourglass
                      frm100101_5.Show
                      frm100101_5.Tag = stCaseNo
                      frm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
                Case "LA"            '顧問
                      Screen.MousePointer = vbHourglass
                      frm100101_6.Show
                      frm100101_6.Tag = stCaseNo
                      frm100101_6.StrMenu
                      Screen.MousePointer = vbDefault
                Case Else                  '服務
                     Select Case Pub_RplStr(Str01)
                         Case "TB"    '條碼
                            Screen.MousePointer = vbHourglass
                            frm100101_7.Show
                            frm100101_7.Tag = stCaseNo
                            frm100101_7.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TM"
                            Screen.MousePointer = vbHourglass
                            frm100101_8.Show
                            frm100101_8.Tag = stCaseNo
                            frm100101_8.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TD"
                            Screen.MousePointer = vbHourglass
                            frm100101_9.Show
                            frm100101_9.Tag = stCaseNo
                            frm100101_9.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TC", "CFC"
                            Screen.MousePointer = vbHourglass
                            frm100101_A.Show
                            frm100101_A.Tag = stCaseNo
                            frm100101_A.StrMenu
                            Screen.MousePointer = vbDefault
                         Case Else
                            Screen.MousePointer = vbHourglass
                            frm100101_B.Show
                            frm100101_B.Tag = stCaseNo
                            frm100101_B.StrMenu
                            Screen.MousePointer = vbDefault
                      End Select
            End Select
        End If
        Me.Enabled = True
        Exit Sub
     End If
     Next i
     Me.Enabled = True
Case 1 '案件進度
     Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
        'Added by Lydia 2019/09/24  預設底色
        If Trim("" & grdDataList.TextMatrix(i, colMemo)) <> "" And Left("" & grdDataList.TextMatrix(i, colMemo), 2) = "相似" Then
            lngColor = &HFFFF&
        Else
            lngColor = QBColor(15)
        End If
        'end 2019/09/24
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            'Modified by Lydia 2019/09/24
            'grdDataList.CellBackColor = QBColor(15)
            grdDataList.CellBackColor = lngColor
        Next j
         grdDataList.col = 2
         stCaseNo = RplStrNew(grdDataList.Text)
        
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = stCaseNo
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault

            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 2 '列印
      If grdDataList.Rows = 2 And grdDataList.TextMatrix(1, 2) = "" Then MsgBox "沒資料可印！": Exit Sub
      Select Case SearchKind
      'Added by Lydia 2019/09/24 +相似案
      Case "本所案號", "相似案"
            PrintData1
      Case "客戶編號"
            PrintData2
      Case Else
      End Select
Case 3 '回前畫面
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 4 '結束
      'Add By Sindy 2024/2/7
      If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
         m_PrevForm.Show
      End If
      '2024/2/7 END
      
      fnCloseAllFrm100
Case Else
End Select
End Sub

'Remove by Lydia 2019/09/24 改成grdDataList_Click
'Private Sub grdDataList_SelChange()
'    grdDataList.Visible = False
'grdDataList.row = grdDataList.MouseRow
''空白不勾
'grdDataList.col = 2
'If Trim(grdDataList.Text) <> "" Then
'    grdDataList.col = 0
'    If grdDataList.row <> 0 Then
'        If grdDataList.Text = "V" Then
'             grdDataList.Text = ""
'             For i = 0 To grdDataList.Cols - 1
'                  grdDataList.col = i
'                  grdDataList.CellBackColor = QBColor(15)
'            Next i
'        Else
'             grdDataList.Text = "V"
'             For i = 0 To grdDataList.Cols - 1
'                 grdDataList.col = i
'                 grdDataList.CellBackColor = &HFFC0C0
'             Next i
'        End If
'    End If
'End If
'grdDataList.Visible = True
'End Sub

'Added by Lydia 2019/09/24
Private Sub GrdDataList_Click()
Dim intRow As Integer
Dim lngColor As Long
   With grdDataList
       If .MouseRow > 0 Then
         'Debug.Print .Text 'Added by Morgan 2019/10/24
         'Clipboard.SetText .Text 'Added by Morgan 2019/10/24
         
          intRow = .MouseRow
          .row = intRow
          .col = cFixed
          
          lngColor = .CellBackColor
          GridClick grdDataList, intRow, 0, 0, cFixed, "V", lngColor
       End If
   End With
End Sub

Sub PrintData1()
GetPleft
Printer.Orientation = 2
Page = 1
CaseNameCht = Combo1.Text
PrintTitle
'PrintMe
For i = 1 To Me.grdDataList.Rows - 1
      grdDataList.row = i
      grdDataList.col = 2
      If grdDataList.Text <> "" Then
         For j = 1 To 10
            Me.grdDataList.col = j
            Printer.CurrentX = PLeft(j - 1)
            Printer.CurrentY = iPrint
            Printer.Print grdDataList.Text
         Next j
         iPrint = iPrint + 300
         If iPrint >= 9000 Then
            Printer.NewPage
            Page = Page + 1
            PrintTitle
         End If
      End If
Next i
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintMe()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "多國案"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print lbl1(0)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print lbl1(1)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print lbl1(5)
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print lbl1(2)
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print lbl1(3)
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print lbl1(6)
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print lbl1(7)
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print lbl1(8)
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print StrLstPayYear
iPrint = iPrint + 300
End Sub
Sub PrintData2()
Dim IsMutiCountry As Boolean
Dim IsHeadle As Boolean
Dim IsNewPage As Boolean

GetPleft
Printer.Orientation = 2
Page = 1
iPrint = 300
CaseNameCht = ""
IsMutiCountry = False
IsHeadle = True
IsNewPage = False
For i = 1 To Me.grdDataList.Rows - 1
      grdDataList.row = i
      grdDataList.col = 2
      'If grdDataList.Text = "多國案" And IsMutiCountry = False Then
      '   IsMutiCountry = True

      'End If
      grdDataList.col = 2
      'If IsMutiCountry = True And Trim(grdDataList.Text) = "" And i <> grdDataList.Rows - 1 Then
      If Trim(grdDataList.Text) = "" And i <> grdDataList.Rows - 1 Then
         IsHeadle = True
         IsNewPage = True
         i = i + 1
         grdDataList.row = i
      End If
      'If grdDataList.Text = "多國案" And IsMutiCountry = True Then
         If IsHeadle = True Then
            If IsNewPage = True Then
               Printer.NewPage
               Page = Page + 1
            End If
            grdDataList.col = 13
            CaseNameCht = grdDataList.Text
            iPrint = 300
            PrintTitle
            IsHeadle = False
            IsNewPage = False
         End If
         For j = 1 To Me.grdDataList.Cols - 5
            Me.grdDataList.col = j
            Printer.CurrentX = PLeft(j - 1)
            Printer.CurrentY = iPrint
            Printer.Print grdDataList.Text
         Next j
         iPrint = iPrint + 300
         If iPrint >= 9000 Then
            Printer.NewPage
            Page = Page + 1
            iPrint = 300
            PrintTitle
         End If
      'End If
'      grdDataList.col = 1
'      'If IsMutiCountry = True And Trim(grdDataList.Text) = "" And i <> grdDataList.Rows - 1 Then
'      If Trim(grdDataList.Text) = "" And i <> grdDataList.Rows - 1 Then
'         IsHeadle = True
'         IsNewPage = True
'
'      End If
Next i
Printer.EndDoc
ShowPrintOk
End Sub

Sub ShowLine()
Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 3200
PLeft(3) = 4400
PLeft(4) = 5400
PLeft(5) = 6600
PLeft(6) = 8600
PLeft(7) = 10000 + 1000
PLeft(8) = 12000 + 1000
PLeft(9) = 14600 + 1000
PLeft(10) = 20000 + 1000
'add by nickc 2005/10/26
For i = 1 To 9
   PLeft(i) = PLeft(i) - 1000
Next i
End Sub

Sub PrintTitle()
      iPrint = 300
      Printer.Font.Name = "細明體"
      Printer.Font.Size = 16
      Printer.Font.Bold = True
      Printer.CurrentX = 2000
      Printer.CurrentY = iPrint
      Printer.Print "案件名稱：" & CaseNameCht
      Printer.Font.Size = 12
      Printer.Font.Bold = False
      iPrint = iPrint + 500
      Printer.CurrentX = 0
      Printer.CurrentY = iPrint
      Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
      Printer.CurrentX = 13000
      Printer.CurrentY = iPrint
      Printer.Print "頁    次：" & str(Page)
      iPrint = iPrint + 300
      ShowLine
      Printer.Font.Size = 12
'edit by nickc 2005/10/26
'      Printer.CurrentX = PLeft(0)
'      Printer.CurrentY = iPrint
'      Printer.Print "狀態"
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "相關案號"
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "申請日"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "種類"
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "申請國家"
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "申請案號"
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = iPrint
      Printer.Print "公告號"
      Printer.CurrentX = PLeft(7)
      Printer.CurrentY = iPrint
      Printer.Print "專利號數"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "專用期間"
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "最近已繳年度"
      iPrint = iPrint + 300
      ShowLine
End Sub
'Added by Morgan 2013/10/11
'衍生設計,集體新式樣檢查
 'Memo by Lydia 2019/09/24 一併將欄位底色改為空白
Private Sub CheckDesign()
   Dim iRow As Integer
   Dim arrNo() As String
   
   grdDataList.Visible = False 'Added by Lydia 2019/09/24
   
   For iRow = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(iRow, 2) <> "" And grdDataList.TextMatrix(iRow, 16) = "3" Then
         arrNo = Split(Pub_RplStr(grdDataList.TextMatrix(iRow, 2)), "-")
         strExc(0) = ""
         If grdDataList.TextMatrix(iRow, 17) = "000" Then
            If Len(grdDataList.TextMatrix(iRow, 6)) = 9 Then
               strExc(0) = "select cp10 from patent,caseprogress where pa11 like '" & grdDataList.TextMatrix(iRow, 6) & "%' and pa11<>'" & grdDataList.TextMatrix(iRow, 6) & "' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10 in ('125','308')"
            End If
         Else
            strExc(0) = "select cp10 from caseprogress where cp01='" & arrNo(0) & "' and cp02='" & arrNo(1) & "' and cp03<>'" & arrNo(2) & "' and cp04='" & arrNo(3) & "' and cp10 in ('105','305')"
         End If
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp(0) = "125" Or RsTemp(0) = "308" Then
                  grdDataList.TextMatrix(iRow, 2) = grdDataList.TextMatrix(iRow, 2) & "(衍)"
               Else
                  grdDataList.TextMatrix(iRow, 2) = grdDataList.TextMatrix(iRow, 2) & "(集)"
               End If
            End If
         End If
      End If
      'Added by Lydia 2019/09/24 將固定欄位的底色設為空白
      strExc(1) = ""
      If Trim("" & grdDataList.TextMatrix(iRow, colMemo)) <> "" And Left("" & grdDataList.TextMatrix(iRow, colMemo), 2) = "相似" Then
          strExc(1) = "Y" '判斷備註為相似案，底色設為黃色
      End If
      For intI = 0 To grdDataList.Cols - 1
         If intI = 0 Then grdDataList.row = iRow
         
         grdDataList.col = intI
         If strExc(1) = "Y" Then '底色設為黃色
             grdDataList.CellBackColor = &HFFFF&
         Else                             '底色設為白色
             grdDataList.CellBackColor = QBColor(15)
         End If
      Next intI
      'end 2019/09/24
   Next
   
   grdDataList.Visible = True 'Added by Lydia 2019/09/24
End Sub
'Added by Morgan 2013/10/11
Private Function RplStrNew(pString As String) As String
   RplStrNew = Replace(Replace(Pub_RplStr(pString), "(衍)", ""), "(集)", "")
End Function

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Added by Morgan 2019/10/24
   '按 Ctrl 點選欄位可複製內容
   If grdDataList.MouseRow > 0 Then
      If Shift = 2 Then
         Clipboard.Clear
         Clipboard.SetText grdDataList.Text
      End If
   End If
End Sub

