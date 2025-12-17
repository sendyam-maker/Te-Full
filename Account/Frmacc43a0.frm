VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc43a0 
   AutoRedraw      =   -1  'True
   Caption         =   "已開發票未收款沖帳作業"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   4920
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   550
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   1740
      Width           =   3500
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   1110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   550
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1170
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc43a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'2014/3/14 create By Sonia
Option Explicit
Public adoacc430 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset

Private Sub Command1_Click()
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0b10 from acc0b0 where a0b04='J' and a0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      MsgBox MsgText(197), , MsgText(5)
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   
   adoTaie.Execute "update acc0b0 set a0b10 = '01' where a0b04='J' "
   Transfer
   adoTaie.Execute "update acc0b0 set a0b10 = null where a0b04='J' "
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5040
   Me.Height = 3000
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   '發票日期預設上一期起日(前一個單月的1日)至上一期止日(前一個雙用的最後一日)
   If Mid(strSrvDate(1), 5, 2) Mod 2 > 0 Then
      MaskEdBox1.Text = TransDate(CompDate(1, -2, (Left(strSrvDate(1), 6) & "01")), 1)
      MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 3) & "/" & Mid(MaskEdBox1.Text, 4, 2) & "/" & Mid(MaskEdBox1.Text, 6, 2)
      MaskEdBox2.Text = TransDate(CompDate(2, -1, (Left(strSrvDate(1), 6) & "01")), 1)
      MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, 3) & "/" & Mid(MaskEdBox2.Text, 4, 2) & "/" & Mid(MaskEdBox2.Text, 6, 2)
   Else
      MaskEdBox1.Text = TransDate(CompDate(1, -3, (Left(strSrvDate(1), 6) & "01")), 1)
      MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 3) & "/" & Mid(MaskEdBox1.Text, 4, 2) & "/" & Mid(MaskEdBox1.Text, 6, 2)
      MaskEdBox2.Text = TransDate(CompDate(2, -1, CompDate(1, -1, (Left(strSrvDate(1), 6)) & "01")), 1)
      MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, 3) & "/" & Mid(MaskEdBox2.Text, 4, 2) & "/" & Mid(MaskEdBox2.Text, 6, 2)
   End If
   
   '預設上一期的最後一個工作日
   MaskEdBox3.Text = TransDate(PUB_GetWorkDay1(DBDATE(Val(FCDate(MaskEdBox2.Text))), 1), 1)
   MaskEdBox3.Text = Mid(MaskEdBox3.Text, 1, 3) & "/" & Mid(MaskEdBox3.Text, 4, 2) & "/" & Mid(MaskEdBox3.Text, 6, 2)
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc43a0 = Nothing
End Sub

Private Sub Transfer()
Dim strAutoNo As String
Dim strSave As String
Dim strSerialNo As String
Dim strDept As String
Dim strAccNo As String
Dim strSalesNo As String
Dim strRemark As String

On Error GoTo Checking
   
   Screen.MousePointer = vbHourglass
   
   adoacc430.CursorLocation = adUseClient
   adoacc430.Open "select * from acc430,acc431,acc0k0 where a4302 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a4302 <= " & Val(FCDate(MaskEdBox2.Text)) & _
                  " and nvl(a4308,0)=0 and nvl(a4310,0)=0 and a4317 is null and a4301=axc01(+) and axc02=a0k01(+) and a0k01 is not null and a0k37 is null order by a4302, a4301", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc430.RecordCount <> 0 Then
      cnnConnection.BeginTrans
   Else
      adoacc430.Close
      Screen.MousePointer = vbDefault
      MsgBox "無符合條件資料待沖帳！", , MsgText(21)
      Exit Sub
   End If
   
   Do While adoacc430.EOF = False
      '讀取欲寫入傳票的各項資料
      'modify by sonia 2017/3/14 加對沖本所案號a0j02
      strExc(0) = "select cp01,cp10,cpm11,a0k03,a0k04,a0k20,sn01,a0j04,a0j02 from acc0j0,acc0k0,nation,caseprogress,casepropertymap,salesno " & _
                  " where a0j13='" & adoacc430.Fields("axc02") & "' and na01(+)=a0j04 and cp09(+)=a0j01 and a0k01(+)=a0j13 and cpm01(+)=cp01 and cpm02(+)=cp10 and a0k20=sn02(+) order by a0j25,a0j01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         
         '傳票摘要:智權人員名/客戶名稱
         'modify by sonia 2017/3/14 摘要加發票號碼
         strRemark = "" & RsTemp("sn01") & "/" & MidB(RsTemp("a0k04").Value, 1, 16) & "/" & adoacc430("a4301")
         
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
      
      '取得J公司的傳票自動編號(JD)MsgText(819)
      strAutoNo = AccAutoNo(MsgText(819), 4, Val(Mid(MaskEdBox3.Text, 1, 3)), Val(Mid(MaskEdBox3.Text, 5, 2)))
      strSave = AccSaveAutoNo(MsgText(819), Mid(strAutoNo, 7, 4), Val(Mid(MaskEdBox3.Text, 1, 3)), Val(Mid(MaskEdBox3.Text, 5, 2)))
      adoTaie.Execute "insert into acc020 (a0201,a0202,a0205,a0206,a0207,a0208)" & _
                      " values ('J', '" & strAutoNo & "', " & Val(FCDate(MaskEdBox3.Text)) & "," & strSrvDate(2) & ",to_char(sysdate,'HH24MISS'),'" & strUserNum & "')"
      '寫入傳票明細檔資料ACC021
      '借方1133應收帳款 modify by sonia 2017/4/5 改用1141未入帳應收帳款
      strSerialNo = GetSerialNo("select max(ax203) from acc021 where ax201='J' and ax202 = '" & strAutoNo & "'", 3)
      'modify by sonia 2017/3/14 部門改TOT原為strDept,加對沖本所案號,摘要加發票號碼,1133應收帳款改發票全額
      adoTaie.Execute "insert into acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212,ax214)" & _
                      " values ('J', '" & strAutoNo & "', '" & strSerialNo & "', 'TOT','1141'," & adoacc430.Fields("a4304") + adoacc430.Fields("a4305") & ",0, '" & RsTemp("a0k03") & "', '" & RsTemp("a0k20") & "','" & strRemark & "', '" & RsTemp("a0j02") & "')"
'cancel by sonia 2017/3/14 取消
'      '借方1135應收銷項稅額
'      If Val(adoacc430.Fields("a4305")) > 0 Then
'         strSerialNo = GetSerialNo("select max(ax203) from acc021 where ax201='J' and ax202 = '" & strAutoNo & "'", 3)
'         adoTaie.Execute "insert into acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212)" & _
'                         " values ('J', '" & strAutoNo & "', '" & strSerialNo & "', '" & strDept & "','1135'," & adoacc430.Fields("a4305") & ",0, '" & RsTemp("a0k03") & "', '" & RsTemp("a0k20") & "','" & strRemark & "')"
'      End If
      '貸方2141應收未收款
      strSerialNo = GetSerialNo("select max(ax203) from acc021 where ax201='J' and ax202 = '" & strAutoNo & "'", 3)
      'modify by sonia 2017/3/14 部門改TOT原為strDept,加對沖本所案號,摘要加發票號碼
      adoTaie.Execute "insert into acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212,ax214)" & _
                      " values ('J', '" & strAutoNo & "', '" & strSerialNo & "', 'TOT','2141',0," & adoacc430.Fields("a4304") & ",'" & RsTemp("a0k03") & "', '" & RsTemp("a0k20") & "','" & strRemark & "', '" & RsTemp("a0j02") & "')"
      '貸借方2119銷項稅額
      If Val(adoacc430.Fields("a4305")) > 0 Then
         strSerialNo = GetSerialNo("select max(ax203) from acc021 where ax201='J' and ax202 = '" & strAutoNo & "'", 3)
         'modify by sonia 2017/3/14 部門改TOT原為strDept,加對沖本所案號,摘要加發票號碼
         adoTaie.Execute "insert into acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax212,ax214)" & _
                         " values ('J', '" & strAutoNo & "', '" & strSerialNo & "', 'TOT','2119',0," & adoacc430.Fields("a4305") & ",'" & RsTemp("a0k03") & "', '" & RsTemp("a0k20") & "','" & strRemark & "', '" & RsTemp("a0j02") & "')"
      End If
      
      '更新發票檔A4317未收款沖帳傳票編號
      adoTaie.Execute "update acc430 set a4317='" & strAutoNo & "' where a4301='" & adoacc430.Fields("a4301") & "'"
      
      adoacc430.MoveNext
   Loop
   adoacc430.Close
   
   cnnConnection.CommitTrans
   
   Screen.MousePointer = vbDefault
   MsgBox "已開發票未收款沖帳作業已處理結束！", , MsgText(21)
   
Checking:
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   If adoacc430.State = adStateOpen Then
      adoacc430.Close
   End If
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   
   If Val(FCDate(MaskEdBox3.Text)) > Val(strSrvDate(2)) Then
      MsgBox "傳票日期不可大於系統日！"
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   
   If ChkWorkDay(Val(DBDATE(FCDate(MaskEdBox3.Text)))) = False Then
      MsgBox "傳票日期必須為工作日！"
      Cancel = True
   End If

End Sub
