VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1127 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "請款單開立發票作業"
   ClientHeight    =   5310
   ClientLeft      =   50
   ClientTop       =   310
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8760
   Begin VB.TextBox txtA4323 
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
      Left            =   5700
      MaxLength       =   1
      TabIndex        =   28
      Top             =   1620
      Width           =   500
   End
   Begin VB.CommandButton cmdFind 
      Height          =   300
      Left            =   3660
      Picture         =   "Frmacc1127.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   330
      Width           =   350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "發票存檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   2
      Top             =   480
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收據抬頭維護"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   1500
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
      MaxLength       =   9
      TabIndex        =   0
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "Label9(5)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1800
      Index           =   5
      Left            =   120
      TabIndex        =   33
      Top             =   4140
      Width           =   8215
   End
   Begin MSForms.Label labA4326 
      Height          =   450
      Left            =   2040
      TabIndex        =   32
      Top             =   2970
      Width           =   6000
      VariousPropertyBits=   19
      Caption         =   "labA4326"
      Size            =   "10583;794"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "發票備註："
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
      Index           =   4
      Left            =   870
      TabIndex        =   31
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label labA4319 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4319"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6180
      TabIndex        =   30
      Top             =   3870
      Width           =   945
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "發票上傳日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   4470
      TabIndex        =   29
      Top             =   3870
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "零  稅  率："
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
      Left            =   4470
      TabIndex        =   27
      Top             =   1650
      Width           =   1300
   End
   Begin VB.Label labA4112 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4112"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2250
      TabIndex        =   26
      Top             =   3870
      Width           =   945
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "目前使用發票日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   3870
      Width           =   2205
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "未收款沖帳傳票："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   24
      Top             =   3570
      Width           =   1935
   End
   Begin VB.Label labA4317 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4317"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2040
      TabIndex        =   23
      Top             =   3570
      Width           =   945
   End
   Begin VB.Label labA4302 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4302"
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
      Left            =   5670
      TabIndex        =   22
      Top             =   2580
      Width           =   945
   End
   Begin VB.Label labA4301 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA4301"
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
      TabIndex        =   21
      Top             =   2580
      Width           =   945
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
      Left            =   5670
      TabIndex        =   20
      Top             =   2100
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
      TabIndex        =   19
      Top             =   2100
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
      TabIndex        =   18
      Top             =   1650
      Width           =   945
   End
   Begin MSForms.Label labA0K04 
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   1230
      Width           =   4935
      VariousPropertyBits=   19
      Caption         =   "labA0K04"
      Size            =   "8705;450"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label labA0K20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K20"
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
      Left            =   5670
      TabIndex        =   16
      Top             =   780
      Width           =   990
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
      TabIndex        =   15
      Top             =   780
      Width           =   990
   End
   Begin VB.Label labA0K02 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "labA0K02"
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
      Left            =   5670
      TabIndex        =   14
      Top             =   300
      Width           =   990
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "發票日期："
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
      Left            =   4470
      TabIndex        =   13
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Index           =   0
      Left            =   870
      TabIndex        =   12
      Top             =   2580
      Width           =   1215
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
      Left            =   4470
      TabIndex        =   11
      Top             =   2100
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
      Left            =   870
      TabIndex        =   10
      Top             =   2100
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
      Left            =   870
      TabIndex        =   9
      Top             =   1650
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
      Left            =   870
      TabIndex        =   8
      Top             =   1230
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
      Left            =   4470
      TabIndex        =   7
      Top             =   780
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
      Left            =   870
      TabIndex        =   6
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收據日期："
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
      Left            =   4470
      TabIndex        =   5
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款單編號："
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
      Left            =   630
      TabIndex        =   4
      Top             =   300
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
Attribute VB_Name = "Frmacc1127"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 labA0k04
'Create by Sindy 2013/12/24
Option Explicit

Dim adoadodc1 As New ADODB.Recordset
Dim m_CU158 As String, m_CU15 As String 'Add By Sindy 2015/7/13
Dim m_CU178 As String 'Add by Amy 2019/07/23
Dim strA4326 As String 'Add by Amy 2022/08/29
Dim strA0K11 As String 'Add By Sindy 2024/9/25


Private Sub Command1_Click()
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   strUserLevel = Me.Name
   Frmacc11p0.Show
   Me.MousePointer = vbDefault
End Sub

'Modify By Sindy 2014/3/31
'Private Sub cmdSave_Click()
Public Sub CmdSave_Click()
'2014/3/31 END
Dim strAXC03 As String, strA4302 As String
Dim bolChkNoPrint As Boolean 'Add By Sindy 2023/9/4
Dim strCP01 As String 'Add By Sindy 2023/10/19
   
   On Error GoTo Checking
   
   If Text1.Text = "" Then
      MsgBox "請輸入請款單編號！", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   If Text1.Tag <> Text1.Text Then
      MsgBox "請先查詢此請款單編號的資料！", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   
   'add by sonia 2014/6/24 收據抬頭超過四個字且無統一編號時提醒
   'If Len(Trim(labA0K04)) >= 4 Then
   If labA4303 = "" Then
      'Modify By Sindy 2015/7/13 非個人或境外公司的其他情形且無統一編號時要提醒
      If m_CU158 <> "Y" And m_CU15 <> "0" Then
      '2015/7/13 END
         If MsgBox("此收據抬頭無統一編號，是否要補輸(客戶檔或收據抬頭檔)？若要補輸請選 是 ！", vbYesNo + vbCritical) = vbNo Then
         Else
            Exit Sub
         End If
      End If
   End If
   'end 2014/6/24
   
   'Add by Amy 2022/05/06 有統編(不是8個0-境外公司)且稅額為0 不可存檔 ex:BG29128906
   '應該是當下沒統編,稅額為0->按「抬頭資料維護」->統編帶入後稅額未重算,導致無法上傳盟立
   '與財務確認不會從此畫面按「抬頭資料維護,故先鎖住,仍加有統編,稅不可為0
   If labA4303 <> MsgText(601) And labA4305 <> "00000000" And Val(labA4305.Caption) = 0 Then
        MsgBox "有統一編號「稅額」不可為0，請確認！"
        Exit Sub
   End If
   
   If labA4301 <> "" Then
      adoTaie.BeginTrans
      
      '更新資料
      'Modify by Amy 2019/07/23 原只判斷labA4301 <> "" 就更新統編,加零稅率有改也更新
      strSql = ""
      strSql = strSql & ",A4303=" & CNULL(labA4303)
      If txtA4323 <> txtA4323.Tag Then strSql = strSql & ",A4323=" & CNULL(txtA4323)
      'end 2019/07/23
      If strSql <> MsgText(601) Then
            strSql = "update ACC430 set " & Mid(strSql, 2) & " Where A4301='" & labA4301 & "'"
      End If
      'end 2019/07/23
      adoTaie.Execute strSql
   Else
      'Modify By Sindy 2023/10/6
      'Add By Sindy 2023/10/19 ACS不管制
      strSql = "select * from acc0k0,acc0j0" & _
               " where a0k01='" & Text1 & "' and a0j13=a0k01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strA0K11 = "" & RsTemp.Fields("A0k11") 'Add By Sindy 2024/9/25
         If "" & RsTemp.Fields("a0j02") <> "" Then 'a0j02=本所案號
            strCP01 = Left(Trim(RsTemp.Fields("a0j02")), Len(Trim(RsTemp.Fields("a0j02"))) - 9)
         End If
      End If
      
      '檢查是否已經收款
      strAXC03 = ""
      strSql = "select a0m01 From acc0M0 where a0m02='" & Text1 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strAXC03 = RsTemp.Fields(0)
      End If
      
      If strCP01 <> "ACS" Then
      '2023/10/19 END
         'Add By Sindy 2023/9/4
         '檢查是否有呈報
         'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
         If strAXC03 = "" Then 'Add By Sindy 2024/10/14 未收款時,才檢查是否已呈報
            bolChkNoPrint = PUB_ChkCU144isN(Mid(labA0K03, 1, 8), Mid(labA0K03, 9, 1), "", strA0K11, , , "C")
            If bolChkNoPrint = True Then
               Text1.SetFocus
               Exit Sub
            End If
         End If
         '2023/9/4 END
      End If
      
      adoTaie.BeginTrans
      
      '先取得發票號碼及發票日期
      labA4301 = GetInvNewNo(strA4302)
      labA4302 = ChangeTStringToTDateString(strA4302)
      labA4326 = strA4326 'Add by Amy 2022/08/29
      
      '更新資料
      'Modify by Amy 2019/07/23 +A4323 零稅率
      'Modify by Amy 2022/08/29 +A4326 發票備註
      strSql = "insert into ACC430(A4301,A4302,A4303,A4304,A4305,A4323,A4326)" & _
               " values(" & CNULL(labA4301) & "," & CNULL(strA4302, True) & "," & CNULL(labA4303) & "," & CNULL(Val(Format(labA4304, "##0")), True) & "," & CNULL(Val(Format(labA4305, "##0")), True) & _
               "," & CNULL(txtA4323) & ",'" & ChgSQL(labA4326) & "')"
      adoTaie.Execute strSql
      strSql = "insert into ACC431(AXC01,AXC02,AXC03)" & _
               " values(" & CNULL(labA4301) & "," & CNULL(Text1) & "," & CNULL(strAXC03) & ")"
      adoTaie.Execute strSql
      
      'Modify By Sindy 2023/9/4
      '未收款(會在此處開發票一定是<未收款先開發票的>)且未列印(a0k19 is null or a0k19=0)請款單即開立發票，請款單列印次數自動上1
      strSql = "update acc0k0 set a0k19=1" & _
               " where a0k01='" & Text1 & "' and (a0k19 is null or a0k19=0)"
      adoTaie.Execute strSql, intI
      'Modify By Sindy 2023/11/21 取消暫不列印
      If intI > 0 Then
         strSql = "update acc0k0 set a0k32=null" & _
                  " where a0k01='" & Text1 & "' and a0k32 in('Y','N')"
         adoTaie.Execute strSql, intI
      End If
      '2023/11/21 END
      '2023/9/4 END
   End If
   adoTaie.CommitTrans
   
   'OpenTable
   cmdSave.Enabled = False
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
      MsgBox "請輸入請款單編號！", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   If Frmacc1127.Text1.Enabled = True Then
      strItemNo = ""
      Call Frmacc1127_Clear
      strA4326 = "" 'Add by Amy 2022/08/29
      OpenTable
   End If
End Sub

'Private Sub Form_Activate()
'   strFormName = Name
'   If strCompanyNo = MsgText(601) Then
'      Exit Sub
'   End If
'   If Adodc1.Recordset.RecordCount <> 0 Then
'      Adodc1.Recordset.MoveFirst
'   End If
'   Adodc1.Recordset.Find "custname = '" & strCompanyNo & "'", 0, adSearchForward, 1
'   If Adodc1.Recordset.EOF = False Then
'      FormShow
'      RecordShow
'   End If
'   strCompanyNo = MsgText(601)
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   KeyEnter KeyCode
'End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intCounter As Integer

   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Height = 5720 'Modify by Amy 2023/08/21 原:5520
   Me.Width = 8880
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2 + 900
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Call Frmacc1127_Clear
   OpenTable
   'Add by Amy 2017/09/29 +顯示a4112
   labA4112.Caption = GetA4112
   'Modify by Amy 2023/05/23 同frmacc11t0
   Label9(5).Caption = "收據抬頭為「可扣繳」資料則發票地址：(境外公司以此規則為主)" & vbCrLf & _
   "　先抓[客戶檔]中文地址(即營業登記地址)，無資料時再抓[收據抬頭檔]營業地址" & vbCrLf & _
                                    "收據抬頭為「不可扣繳」資料則發票地址：" & vbCrLf & _
   "　先抓[客戶檔]聯絡地址，無資料時抓中文地址(即營業登記地址)" & vbCrLf & _
   "　客戶檔[無]資料時再抓，[收據抬頭檔]郵寄地址，無資料時再抓營業地址" & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strItemNo = "" 'Add By Sindy 2017/2/14
   If UCase(strTitle) = UCase("Frmacc1140") Then
      'strItemNo = ""
      strCustNo = ""
      strTitle = ""
      tool1_enabled
      Frmacc1140.Enabled = True
      Frmacc1140.Show
   ElseIf UCase(strTitle) = UCase("Frmacc11d0") Then
      'strItemNo = ""
      strCustNo = ""
      strTitle = ""
      tool1_enabled
      Frmacc11d0.Enabled = True
      Frmacc11d0.Show
   ElseIf UCase(strTitle) = UCase("Frmacc1121") Then
      strTitle = ""
      tool3_enabled
      Frmacc1121.Enabled = True
      Frmacc1121.Show
   ElseIf UCase(strTitle) = UCase("Frmacc1220") Then
      'strItemNo = ""
      strTitle = ""
      tool3_enabled
      Frmacc1220.Enabled = True
      Frmacc1220.Show
   ElseIf UCase(strTitle) = UCase("Frmacc1230") Then
      'strItemNo = ""
      strTitle = ""
      tool3_enabled
      Frmacc1230.Enabled = True
      Frmacc1230.Show
   'Add By Sindy 2016/6/8
   ElseIf UCase(strTitle) = UCase("Frmacc12d0") Then
      'strItemNo = ""
      strTitle = ""
      tool3_enabled
      Frmacc12d0.Enabled = True
      Frmacc12d0.Show
   '2016/6/8 END
   Else
      'Add By Sindy 2016/6/29
      If UCase(strTitle) = UCase("Frmacc1240") Then
         Frmacc1240.Enabled = True
      End If
      '2016/6/29 END
      StatusClear
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
   End If
   Set Frmacc1127 = Nothing
End Sub

'Add by Amy 2022/05/09 原OpenTable_Old 原程式部分改寫至共用fuctnion
Private Sub OpenTable()
    Dim strField1, strField2, str0K0Data, str430Data
    Dim i As Integer, strData1 As String, strData2 As String, strMsg As String, stTP As String
        
    If strItemNo <> "" Then Text1.Text = strItemNo
    If Text1.Text = MsgText(601) Then Exit Sub
    Text1.Tag = Text1.Text
    Label9(3).Visible = False: labA4317.Visible = False: cmdSave.Enabled = False
    'Memo依序傳入需回傳的欄位(需與PUB_InvProc抓的欄位順序相同,避免抓不到 ex:此a4319,a4323 PUB_InvProc也要先抓a4319再抓a4323)
    strData1 = "a0k02,a0k03,a0k04,st02"
    'Modify by Amy 2022/08/29 +a4326
    strData2 = "a4301,a4302,a4303,a4304,a4305,a4317,a4319,a4323,a4326"
    strField1 = Split(strData1, ",")
    strField2 = Split(strData2, ",")
    m_CU158 = "": m_CU15 = ""
    Call PUB_InvProc(Text1, , , , Me.Name, strData1, strData2, m_CU158, m_CU15, strMsg)
    If strMsg <> MsgText(601) Then
        'Modify by Amy 2023/08/21 +不可開立發票
        If InStr(strMsg, "無此請款單編號") > 0 Or InStr(strMsg, "非J公司") > 0 Or InStr(strMsg, "不可開立發票") > 0 Then
            MsgBox strMsg, vbExclamation
            Exit Sub
        End If
    End If
    str0K0Data = Split(strData1, ",")
    str430Data = Split(strData2, ",")
    txtA4323.Enabled = False '2022/05/06與財務確目前從"客戶/代理人財務email資料維護"勾選零稅率.再開發票,故先鎖住,避免拿掉畫面零稅率不會重算稅-婉莘
    If InStr(strMsg, "ACS案") > 0 Then
        MsgBox strMsg, vbExclamation
    End If
    If InStr(strMsg, "此請款單已作廢") > 0 Then
        MsgBox strMsg, vbExclamation
    End If
    cmdSave.Enabled = True
    
    For i = LBound(strField1) To UBound(strField1)
        Select Case UCase(strField1(i))
            Case "A0K02" '收據日期
                labA0K02.Caption = str0K0Data(i)
            Case "A0K03" '客戶編號
                labA0K03.Caption = str0K0Data(i)
            Case "A0K04" '收據抬頭
                labA0K04.Caption = str0K0Data(i)
            Case "ST02" '業務人員
                labA0K20.Caption = str0K0Data(i)
        End Select
    Next i
    
    For i = LBound(strField2) To UBound(strField2)
        stTP = str430Data(i)
        Select Case UCase(strField2(i))
            Case "A4301" '發票號碼
                labA4301.Caption = stTP
            Case "A4302" '發票日期
                 labA4302.Caption = stTP
            Case "A4303" '統一編號
                labA4303.Caption = stTP
            Case "A4304" '銷售額
                If stTP <> MsgText(601) Then
                    stTP = Format(stTP, DDollar2)
                End If
                labA4304.Caption = stTP
            Case "A4305" '稅額
                If stTP <> MsgText(601) Then
                    stTP = Format(stTP, DDollar2)
                End If
                labA4305.Caption = stTP
            Case "A4317" '未收款沖帳傳票編號
                labA4317.Caption = stTP
            Case "A4319" '電子發票上傳日
                labA4319.Caption = stTP
            Case "A4323" '零稅率
                txtA4323 = stTP
                txtA4323.Tag = txtA4323
            Case "A4326" 'Add by Amy 2022/08/29 發票備註
                strA4326 = stTP
                If labA4301.Caption <> MsgText(601) Then
                    labA4326.Caption = stTP
                End If
        End Select
    Next i
    '有發票且未收款沖帳傳票編號不是空
    If labA4301.Caption <> MsgText(601) And labA4317.Caption <> MsgText(601) Then
        Label9(3).Visible = True: labA4317.Visible = True
    End If
    '有發票且電子發票上傳日為空 or 發票日小於電子發票上傳啟用日,鎖住發票儲存鈕
    If labA4301.Caption <> MsgText(601) And (labA4319.Caption <> MsgText(601) Or Val(FCDate(labA4302.Caption)) < TranInvoiceDate) Then
        cmdSave.Enabled = False
        'txtA4323.Enabled = False
    End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
'Memo by Amy 2022/05/06 不使用-原程式部分改寫至共用fucnion
Private Sub OpenTable_Old()
Dim dblAmt As Double, dblTax As Double
'Dim strCU11 As String

On Error GoTo Checking

   If strItemNo <> "" Then Text1.Text = strItemNo
   Text1.Tag = Text1.Text
   adoadodc1.CursorLocation = adUseClient
   strSql = "select sqldatet(a0k02) a0k02,a0k03,a0k04,a0k20,st02,sum(nvl(a0k06,0))+sum(nvl(a0k07,0)) - sum(nvl(A1uAmt,0)) Amt,a0k11,a0k09,cp01,cp10" & _
            " From acc0k0, staff, (select a1u02,sum(nvl(a1u07,0))+sum(nvl(a1u09,0)) A1uAmt from acc1u0 where a1u02='" & Text1 & "' group by a1u02)" & _
                 ",acc0j0,caseprogress" & _
            " Where a0k01='" & Text1 & "'" & _
            " and a0k20=st01(+)" & _
            " and a0k01=a1u02(+)"
   'Added by Sindy 2020/11/11 J公司請款單之收文
   strSql = strSql & " and a0j13=a0k01 and a0j01=cp09"
   'end 2020/11/11
   strSql = strSql & " group by a0k02,a0k03,a0k04,a0k20,st02,a0k11,a0k09,cp01,cp10"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Label9(3).Visible = False: labA4317.Visible = False
'   cmdSave.Enabled = False
   If adoadodc1.RecordCount > 0 Then
      If "" & adoadodc1.Fields("a0k11") = "J" Then
         labA0K02.Caption = "" & adoadodc1.Fields("a0k02")
         labA0K03.Caption = "" & adoadodc1.Fields("a0k03")
         labA0K04.Caption = "" & adoadodc1.Fields("a0k04")
         labA0K20.Caption = "" & adoadodc1.Fields("st02")
         dblAmt = "" & adoadodc1.Fields("Amt")
         dblTax = dblAmt - Round((dblAmt / 1.05), 0)

         '發票資料
         'Modify by Amy 2019/07/23 +a4323,a4319 零稅率/電子發票上傳時間
         strSql = "select a4301,sqldatet(a4302) a4302,a4303,a4304,a4305,a4317,a4323,Decode(a4319,111111,'',sqldatet(a4319)) a4319" & _
                  " From acc431,acc430" & _
                  " where axc02='" & Text1 & "'" & _
                  " and axc01=a4301"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            labA4301.Caption = "" & RsTemp.Fields("a4301")
            labA4302.Caption = "" & RsTemp.Fields("a4302")
            labA4303.Caption = "" & RsTemp.Fields("a4303")
            labA4304.Caption = Format("" & RsTemp.Fields("a4304"), DDollar2)
            labA4305.Caption = Format("" & RsTemp.Fields("a4305"), DDollar2)
            If "" & RsTemp.Fields("a4317") <> "" Then
               labA4317.Caption = "" & RsTemp.Fields("a4317")
               Label9(3).Visible = True: labA4317.Visible = True
            End If
            'Add by Amy 2019/07/23 +a4323零稅率/a4319上傳日期,電子發票已上傳不可修改資料
            labA4319.Caption = "" & RsTemp.Fields("a4319")
            txtA4323 = "" & RsTemp.Fields("a4323")
            txtA4323.Tag = txtA4323
            cmdSave.Enabled = True
            txtA4323.Enabled = True
            If labA4319 <> MsgText(601) Or FCDate(labA4302) < TranInvoiceDate Then
                cmdSave.Enabled = False
                txtA4323.Enabled = False
            End If
            'end 2019/07/23
         Else
            '統一編號
            'Modify By Sindy 2014/8/11 +and (cu80 is null or cu80='其他' or cu80='業務自行處理') and cu02='0'
'            strCU11 = ""
'            strSql = "select cu11" & _
'                     " From customer" & _
'                     " where cu04='" & labA0K04 & "' and (cu80 is null or cu80='其他' or cu80='業務自行處理') and cu02='0'" & _
'                     " and cu15<>'0'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               strCU11 = "" & RsTemp.Fields("cu11")
'            End If
'            If strCU11 = "" Then
'               'Modify By Sindy 2014/3/25 若A4202='04150022'者視為空值
'               strSql = "select a4202" & _
'                        " From acc420" & _
'                        " where a4201='" & labA0K04 & "' and A4202<>'04150022'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               If intI = 1 Then
'                  strCU11 = "" & RsTemp.Fields("a4202")
'               End If
'            End If
'            labA4303.Caption = strCU11

            'Modify By Sindy 2015/7/13 +抓是否為境外公司,個人或公司
            'labA4303.Caption = PUB_GetTaxNo(labA0K04) '統編
            labA4303.Caption = PUB_GetTaxNo(labA0K04, , m_CU158, m_CU15, m_CU178) '統編
            '2015/7/13 END

            '2014/12/2 modify by sonia 個人無稅額
            'labA4304.Caption = Format(dblAmt - dblTax, DDollar2)
            'labA4305.Caption = Format(dblTax, DDollar2)
            'Modify By Sindy 2017/1/9
            'If labA4303 <> "" Then
            If labA4303 <> "" And labA4303 <> "00000000" Then
            '2017/1/9 END
               labA4304.Caption = Format(dblAmt - dblTax, DDollar2)
               labA4305.Caption = Format(dblTax, DDollar2)
            Else
               labA4304.Caption = Format(dblAmt, DDollar2)
               labA4305.Caption = Format(0, DDollar2)
            End If
            '2014/12/2 end
            txtA4323 = m_CU178 'Add by Amy 2019/07/23 零稅率
            If Val("" & adoadodc1.Fields("a0k09")) > 0 Then
               MsgBox "此請款單已作廢!!"
               cmdSave.Enabled = False
               txtA4323.Enabled = False 'Add by Amy 2019/07/23 零稅率
            'Modify By Sindy 2018/3/29
            Else
               cmdSave.Enabled = True
               txtA4323.Enabled = True 'Add by Amy 2019/07/23 零稅率
            '2018/3/29 END
            End If
         End If

         'Added by Sindy 2020/11/11 J公司請款單之收文若為ACS案且案件性質706代收代付時，不可開立發票
         If adoadodc1.Fields("cp01") = "ACS" And adoadodc1.Fields("cp10") = "706" Then
            MsgBox "ACS案且案件性質706.代收代付，不可開立發票!!"
            cmdSave.Enabled = False
            txtA4323.Enabled = False
         End If
         'end 2020/11/11

      Else
         MsgBox "非J公司請款單!!"
      End If
   Else
      If Trim(Text1) <> "" Then MsgBox "無此請款單編號!!"
   End If
   adoadodc1.Close

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc1127_Clear()
   With Frmacc1127
      .labA0K02.Caption = ""
      .labA0K03.Caption = ""
      .labA0K04.Caption = ""
      .labA0K20.Caption = ""
      .labA4301.Caption = ""
      .labA4302.Caption = ""
      .labA4303.Caption = ""
      .labA4304.Caption = ""
      .labA4305.Caption = ""
      .labA4317.Caption = ""
      'Add by Amy 2019/07/23 零稅率
      .labA4319.Caption = ""
      .txtA4323 = ""
      .txtA4323.Tag = ""
      'end 2019/07/23
      .labA4112.Caption = ""  'Add by Amy 2022/08/24
      .labA4326.Caption = "" 'Add by Amy 2022/08/29 發票備註
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

'Add by Amy 2019/07/23
Private Sub txtA4323_GotFocus()
    TextInverse txtA4323
End Sub

'Add by Amy 2019/07/23
Private Sub txtA4323_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
        KeyAscii = 0
        Beep
    End If
End Sub
