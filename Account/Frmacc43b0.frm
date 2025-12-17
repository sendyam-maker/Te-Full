VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc43b0 
   AutoRedraw      =   -1  'True
   Caption         =   "發票跨期轉開作業"
   ClientHeight    =   4680
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7570
   Begin VB.CommandButton cmdFind 
      Height          =   300
      Left            =   3700
      Picture         =   "Frmacc43b0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   680
      Width           =   350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "產生銷退折讓及新發票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4764
      TabIndex        =   1
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   0
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "傳票批次作業後會自動產生(已開發票未收款沖回傳票)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   25
      Top             =   4230
      Width           =   5325
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "未收款已申報發票，非銷帳但需要當期發票情況"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   24
      Top             =   120
      Width           =   4605
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "請款單號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   22
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label labAXC02 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labAXC02"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      TabIndex        =   21
      Top             =   1140
      Width           =   1560
   End
   Begin VB.Label NewA4302 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "NewA4302"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      TabIndex        =   20
      Top             =   3660
      Width           =   1080
   End
   Begin VB.Label NewA4301 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "NewA4301"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   19
      Top             =   3660
      Width           =   1080
   End
   Begin VB.Label labA4305 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4305"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      TabIndex        =   18
      Top             =   2940
      Width           =   945
   End
   Begin VB.Label labA4304 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4304"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   17
      Top             =   2940
      Width           =   945
   End
   Begin VB.Label labA4303 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4303"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   16
      Top             =   2490
      Width           =   945
   End
   Begin MSForms.Label labA0K04 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   2070
      Width           =   5415
      VariousPropertyBits=   19
      Caption         =   "labA0K04"
      Size            =   "9551;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label labA0K20 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   1620
      Width           =   1560
      VariousPropertyBits=   19
      Caption         =   "labA0K20"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label labA0K03 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K03"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   13
      Top             =   1620
      Width           =   990
   End
   Begin VB.Label OldA4302 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "OldA4302"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   12
      Top             =   1140
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "新發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4215
      TabIndex        =   11
      Top             =   3660
      Width           =   1440
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "新發票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   10
      Top             =   3660
      Width           =   1440
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "稅　　額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   9
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "銷 售 額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   855
      TabIndex        =   8
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "統一編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   855
      TabIndex        =   7
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   855
      TabIndex        =   6
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   5
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   855
      TabIndex        =   4
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "原發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   630
      TabIndex        =   3
      Top             =   1140
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   870
      TabIndex        =   2
      Top             =   660
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc43b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已修改
'Create by Sonia 2014/3/12
Option Explicit

Dim adoadodc1 As New ADODB.Recordset
Dim strAXC02 As String, strAXC03 As String


Private Sub CmdSave_Click()
Dim strA4302 As String
Dim strRetNo As String
Dim strSerialNo As String
Dim strDept As String
Dim strAccNo As String
Dim strSalesNo As String
Dim strRemark As String

   On Error GoTo Checking
   
   If Text1.Text = "" Then
      MsgBox "請輸入發票號碼！", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   If Text1.Tag <> Text1.Text Then
      MsgBox "請重新查詢此發票號碼的資料！", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   
   adoTaie.BeginTrans
   
   '更新原發票之發票轉開日期為系統日
   strSql = "update ACC430 set a4310=" & Val(strSrvDate(2)) & " where a4301='" & Text1.Text & "'"
   adoTaie.Execute strSql
   '更新原發票之收據號碼為AXC02||'轉'
   strSql = "update ACC431 set axc02=axc02||'轉' where axc01='" & Text1.Text & "'"
   adoTaie.Execute strSql
   '取得發票號碼及發票日期
   NewA4301 = GetInvNewNo(strA4302)
   NewA4302 = ChangeTStringToTDateString(strA4302)
   '新增新發票ACC430,ACC431
   strSql = "insert into ACC430 (A4301,A4302,A4303,A4304,A4305)" & _
            " values(" & CNULL(NewA4301) & "," & CNULL(strA4302, True) & "," & CNULL(labA4303) & "," & CNULL(Val(Format(labA4304, "##0")), True) & "," & CNULL(Val(Format(labA4305, "##0")), True) & ")"
   adoTaie.Execute strSql
   strSql = "insert into ACC431 (AXC01,AXC02,AXC03)" & _
            " values(" & CNULL(NewA4301) & "," & CNULL(strAXC02) & "," & CNULL(strAXC03) & ")"
   adoTaie.Execute strSql
   '讀取銷退折讓明細的品名,及沖轉傳票的各項資料
   strExc(0) = "select a0j02,na03,decode(sk02,'1','專利','2','商標','其他') Sys,nvl(a0j22,getcp10desc(cp01,cp10,a0j04)) cp10N,cp01,cp10,cpm11,a0k20,sn01,a0j04 from acc0j0,acc0k0,nation,caseprogress,systemkind,casepropertymap,salesno " & _
               " where a0j13='" & strAXC02 & "' and na01(+)=a0j04 and cp09(+)=a0j01 and sk01(+)=cp01 and a0k01(+)=a0j13 and cpm01(+)=cp01 and cpm02(+)=cp10 and a0k20=sn02(+) order by a0j25,a0j01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = RsTemp("na03") & RsTemp("Sys") & "/" & RsTemp("cp10N")
      
      '傳票摘要:智權人員名/客戶名稱
      'modify by sonia 2017/3/14 摘要加發票號碼
      strRemark = "" & RsTemp("sn01") & "/" & MidB(labA0K04, 1, 16) & "/原發票號碼" & Text1.Text & "/新發票號碼" & NewA4301
      
      If IsNull(RsTemp.Fields("cpm11").Value) = False Then
         'modify by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文組
         'If AccNoToSalesNo(RsTemp.Fields("cpm11").Value) = "" Then
         If AccNoToSalesNo(RsTemp.Fields("cpm11").Value, "" & RsTemp("a0j02")) = "" Then
            strSalesNo = IIf(IsNull(RsTemp("a0k20")), "", RsTemp("a0k20"))
         Else
            'modify by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文組
            'strSalesNo = AccNoToSalesNo(RsTemp("cpm11"))
            strSalesNo = AccNoToSalesNo(RsTemp("cpm11"), "" & RsTemp("a0j02"))
         End If
      Else
         strSalesNo = ""
      End If
      If RsTemp("a0j04") <> "000" And (Mid(RsTemp("cp01"), 1, 1) = "P" Or Mid(RsTemp("cp01"), 1, 1) = "T") Then
         If Mid(RsTemp("cp01"), 1, 1) = "P" Then
            strAccNo = "411103"
         Else
            strAccNo = "410103"
         End If
      Else
         If IsNull(RsTemp("cpm11")) Then
            strAccNo = "XXX"
         Else
            strAccNo = RsTemp("cpm11")
         End If
      End If
      If IsNull(RsTemp("cp01")) Then
         strDept = "null"
      Else
         'MODIFY BY SONIA 2016/1/5
         'Select Case Mid(strAccNo, 1, 4)
         '   Case "4101", "4151"
         '      strDept = "T"
         '   Case "4111"
         '      strDept = "P"
         '   Case "4121"
         '      strDept = "CFT"
         '   Case "4172"
         '      If RsTemp("cpm11") = "417202" Then
         '         strDept = "T"
         '      Else
         '         strDept = "FCT"
         '      End If
         '   Case "4131"
         '      strDept = "CFP"
         '   Case "4141"
         '      strDept = "L"
         '   Case "4171"
         '      strDept = "FCP"
         '   Case "4181"
         '      strDept = "L"
         '   Case "4161"
         '      strDept = "FCL"
         '   Case Else
         '      strDept = "TOT"
         'End Select
         If Left(strAccNo, 1) = "4" Then
            strDept = PUB_GETAccNODept(strAccNo, strDept)
         Else
            strDept = "TOT"
         End If
         'END 2016/1/5
      End If
   End If
   '新增銷退折讓明細檔
   strSql = "insert into acc460 (A4601,A4602,A4603,A4604,A4605) values('" & Text1.Text & "','" & strExc(1) & "','" & RsTemp("a0j02") & "'," & CNULL(Val(Format(labA4304, "##0")), True) & "," & CNULL(Val(Format(labA4305, "##0")), True) & ")"
   adoTaie.Execute strSql
   '新增客戶回執記錄
   strRetNo = AutoNo("H", 5, 1)
   strSql = "INSERT INTO ACC250 (A2501,A2502,A2503,A2504,A2505,A2506,A2513)" & _
      " VALUES('" & strRetNo & "','7','" & labA0K03 & "'," & Val(Format(labA4304, "##0")) + Val(Format(labA4305, "##0")) & ",'" & Text1.Text & "'" & _
      ",'" & strUserNum & "','" & labA0K04 & "')"
   adoTaie.Execute strSql
   '新增J公司沖轉傳票
   '借方2141應收未收款
   'modify by sonia 2017/3/15 部門改TOT原為strDept,加對沖本所案號,摘要加發票號碼
   strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='J' and a1p02 = 'Y' and a1p04 = '" & Text1.Text & "發票轉開'", 3)
   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18) " & _
            " values ('J', 'Y', '" & strSerialNo & "', '" & Text1.Text & "發票轉開', '2141', 'TOT', " & Val(Format(labA4304, "##0")) & ", 0, null, null, null, null, null, '" & strRemark & "', '" & labA0K03 & "', '" & strSalesNo & "', '" & RsTemp("a0j02") & "' ," & Val(strSrvDate(2)) & ")"
   adoTaie.Execute strSql
   '借方2119銷項稅額
   'modify by sonia 2017/3/15 部門改TOT原為strDept,加對沖本所案號,摘要加發票號碼
   strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='J' and a1p02 = 'Y' and a1p04 = '" & Text1.Text & "發票轉開'", 3)
   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18) " & _
            " values ('J', 'Y', '" & strSerialNo & "', '" & Text1.Text & "發票轉開', '2119', 'TOT', " & Val(Format(labA4305, "##0")) & ", 0, null, null, null, null, null, '" & strRemark & "', '" & labA0K03 & "', '" & strSalesNo & "', '" & RsTemp("a0j02") & "'," & Val(strSrvDate(2)) & ")"
   adoTaie.Execute strSql
   '貸方1133應收帳款
   'modify by sonia 2017/3/15 部門改TOT原為strDept,加對沖本所案號,摘要加發票號碼,1133應收帳款改發票全額
   'modify by sonia 2024/11/5 貸方1133應收帳款改用1141未入帳應收帳款(Frmacc43a0已於2017/4/5修改)
   strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='J' and a1p02 = 'Y' and a1p04 = '" & Text1.Text & "發票轉開'", 3)
   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18) " & _
            " values ('J', 'Y', '" & strSerialNo & "', '" & Text1.Text & "發票轉開', '1141', 'TOT', 0, " & Val(Format(labA4304, "##0")) + Val(Format(labA4305, "##0")) & ", null, null, null, null, null, '" & strRemark & "', '" & labA0K03 & "', '" & strSalesNo & "', '" & RsTemp("a0j02") & "'," & Val(strSrvDate(2)) & ")"
   adoTaie.Execute strSql
'cancel by sonia 2017/3/15
'   '貸方1135應收銷項稅額
'   strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='J' and a1p02 = 'Y' and a1p04 = '" & Text1.Text & "發票轉開'", 3)
'   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18) " & _
'            " values ('J', 'Y', '" & strSerialNo & "', '" & Text1.Text & "發票轉開', '1135', '" & strDept & "', 0, " & Val(Format(labA4305, "##0")) & ", null, null, null, null, null, '" & strRemark & "', '" & labA0K03 & "', '" & strSalesNo & "', null," & Val(strSrvDate(2)) & ")"
'   adoTaie.Execute strSql
   
   adoTaie.CommitTrans
   cmdSave.Enabled = False
   
   Label9.Visible = True: NewA4301.Visible = True
   Label10.Visible = True: NewA4302.Visible = True
   Exit Sub
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   Else
      adoTaie.RollbackTrans
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub cmdFind_Click()
   If Text1.Text = "" Then
      MsgBox "請輸入發票號碼！", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
      
   Call Frmacc43b0_Clear
   OpenTable

End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intCounter As Integer

   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Height = 5085
   Me.Width = 7695
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2 + 900
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   Call Frmacc43b0_Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc43b0 = Nothing

End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim dblAmt As Double, dblTax As Double
   
On Error GoTo Checking
   
   Text1.Tag = Text1.Text
   strAXC02 = "": strAXC03 = ""
   adoadodc1.CursorLocation = adUseClient
   strSql = "select sqldatet(a4302) a4302,a4303,a4304,a4305,a4317,sqldatet(a4310) a4310,axc02,axc03,a0k03,a0k04,st02,a0k37" & _
            " From acc430, acc431, acc0k0, staff Where a4301='" & Text1 & "'" & _
            " and a4301=axc01(+) and axc02=a0k01(+) and a0k20=st01(+)"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   If adoadodc1.RecordCount > 0 Then
      OldA4302.Caption = "" & adoadodc1.Fields("a4302")
      labAXC02.Caption = "" & adoadodc1.Fields("axc02")
      labA0K03.Caption = "" & adoadodc1.Fields("a0k03")
      labA0K04.Caption = "" & adoadodc1.Fields("a0k04")
      labA0K20.Caption = "" & adoadodc1.Fields("st02")
      labA4303.Caption = "" & adoadodc1.Fields("a4303")
      labA4304.Caption = Format("" & adoadodc1.Fields("a4304"), DDollar2)
      labA4305.Caption = Format("" & adoadodc1.Fields("a4305"), DDollar2)
      strAXC02 = "" & adoadodc1.Fields("axc02")
      strAXC03 = "" & adoadodc1.Fields("axc03")
      
      If "" & adoadodc1.Fields("a0k37") = "" Then
         If "" & adoadodc1.Fields("a4317") <> "" Then
            If "" & adoadodc1.Fields("a4310") = "" Then
               cmdSave.Enabled = True
            Else
               MsgBox "此發票已於 " & adoadodc1.Fields("a4310") & " 轉開新發票 !!"
            End If
         Else
            MsgBox "此發票為當期資料, 不可轉開 !!"
         End If
      Else
         MsgBox "此請款單已收款或銷退, 不可轉開 !!"
      End If
   Else
      If Trim(Text1) <> "" Then MsgBox "無此發票號碼!!"
   End If
   adoadodc1.Close
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc43b0_Clear()
   Label9.Visible = False: NewA4301.Visible = False
   Label10.Visible = False: NewA4302.Visible = False
   cmdSave.Enabled = False
   
   With Frmacc43b0
      .OldA4302.Caption = ""
      .labAXC02.Caption = ""
      .labA0K03.Caption = ""
      .labA0K20.Caption = ""
      .labA0K04.Caption = ""
      .labA4303.Caption = ""
      .labA4304.Caption = ""
      .labA4305.Caption = ""
      .NewA4301.Caption = ""
      .NewA4302.Caption = ""
   End With
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
   If Text1 <> "" And Text1.Enabled = True Then
      Call cmdFind_Click
   End If
End Sub

