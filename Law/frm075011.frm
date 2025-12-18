VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075011 
   BorderStyle     =   1  '單線固定
   Caption         =   "庭期資料維護"
   ClientHeight    =   5750
   ClientLeft      =   120
   ClientTop       =   570
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3405
      TabIndex        =   4
      Top             =   180
      Width           =   800
   End
   Begin VB.TextBox txtcp04 
      Height          =   270
      Left            =   2925
      MaxLength       =   2
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtcp03 
      Height          =   270
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
      Height          =   270
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtcp01 
      Height          =   270
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   550
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8388
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7560
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4230
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   7461
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   16
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      MergeCells      =   1
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
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Left            =   2190
      TabIndex        =   14
      Top             =   585
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1170
      TabIndex        =   13
      Top             =   915
      Width           =   7245
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12779;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "註：開庭日期欄若有**，表示已取消庭期！"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   60
      TabIndex        =   12
      Top             =   5550
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   210
      TabIndex        =   11
      Top             =   930
      Width           =   975
   End
   Begin VB.Label lbeCus 
      DataSource      =   "ORADC1"
      Height          =   255
      Left            =   1170
      TabIndex        =   10
      Top             =   585
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   255
      Left            =   210
      TabIndex        =   8
      Top             =   585
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   210
      TabIndex        =   7
      Top             =   255
      Width           =   975
   End
End
Attribute VB_Name = "frm075011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/17 改成Form2.0 ; cboCaseName、lbeCusName、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean
Dim Rs As New ADODB.Recordset
Dim m_Lawcase As Boolean

Private Sub cmdEnd_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
  SearchData
End Sub

Public Sub SearchData()
Dim LcTmp As String, strCus As String, strTempName As String
   
   If txtcp01.Text = Empty Or txtcp02.Text = Empty Then
      MsgBox "請輸入本所案號!", vbInformation, "庭期資料維護"
      Exit Sub
   End If
   MSHFlexGrid1.Clear
   Label1.Tag = GiveSymbol(txtcp01, txtcp02, txtcp03, txtcp04, LcTmp)
   If QueryDB = False Then
      MsgBox "基本檔無資料!", vbInformation, "庭期資料維護"
      lbeCus = ""
      lbeCusName = ""
      cboCaseName.Clear
      MSHFlexGrid1.Rows = 2
      m_Lawcase = False
      Exit Sub
   Else
      m_Lawcase = True
   End If
   If txtcp01 <> "LA" Then
      'Modify by Amy 2018/01/24 開庭別+調解庭,開庭種類+調解
      strExc(1) = "select '' v,decode(cp05,null,'',SUBSTR(CP05, 1, 4) - 1911 || '/' || SUBSTR(CP05, 5, 2) || '/' ||" + _
         "SUBSTR(CP05, 7, 2))  as 收受日,cdp01 as 收文號,decode(cdp03,null, '',decode(cdp18,null,'','**')||sqldatet(cdp03)) as 開庭日期" + _
         ",substr(sqltime(cdp04||'00'),1,5) as 時間,decode(cdp05,or01,or02) 法院,decode(cdp02, ST01, ST02) 開庭人員 ," + _
         "decode(cdp17,'1','民事庭','2','偵查庭','3','刑事庭','4','刑附民庭','5','行政庭','6','調解庭') 開庭別,decode(cdp06,'1','偵查','2','審查','3','言詞辯論','4','調查','5','調解') 開庭種類,lc05,lc06,lc07,lc11" + _
         " From lawcase, caseprogress, courtyardperiod, organization, staff" + _
         " where cdp01=cp09(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" + _
         " and " & ChgCaseprogress(LcTmp) & " and cdp05=or01(+) and cdp02=st01(+) and cp10(+) = '" & 通知開庭 & "'"
   Else
      'Modify by Amy 2018/01/24 開庭別+調解庭,開庭種類+調解
      strExc(1) = "select '' v,decode(cp05,null,'',SUBSTR(CP05, 1, 4) - 1911 || '/' || SUBSTR(CP05, 5, 2) || '/' ||" + _
         " SUBSTR(CP05, 7, 2))  as 收受日,cdp01 as 收文號,decode(cdp03,null,''," + _
         " decode(cdp18,null,'','**')||sqldatet(cdp03)) as 開庭日期," + _
         " substr(sqltime(cdp04||'00'),1,5) as 時間, decode(cdp05,or01,or02) 法院,decode(cdp02, ST01, ST02) 開庭人員 ," + _
         " decode(cdp17,'1','民事庭','2','偵查庭','3','刑事庭','4','刑附民庭','5','行政庭','6','調解') 開庭別,decode(cdp06,'1','偵查','2','審查','3','言詞辯論','4','調查','5','調解') 開庭種類,hc05,hc06 from caseprogress,hirecase," + _
         " courtyardperiod,organization,staff where cp01=Hc01(+) and cp02=Hc02(+) and cp03=Hc03(+) and cp04=Hc04(+)" + _
         " and " & ChgCaseprogress(LcTmp) & " and cdp01=cp09(+) and cdp05=or01(+) and cdp02=st01(+) and cp10 = '" & 通知開庭 & "'"
   End If
   intI = 0
   Set Rs = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      GridHead
      Set MSHFlexGrid1.Recordset = Rs
      cboCaseName.Clear
      If txtcp01 <> "LA" Then
         strCus = "" & Rs.Fields!LC11
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strCus, strTempName) Then lbeCus = strCus: lbeCusName = strTempName
         If ClsPDGetCustomer(strCus, strTempName) Then lbeCus = strCus: lbeCusName = strTempName
         If Not IsNull(Rs.Fields!lc05) Then cboCaseName.AddItem "中:" + Rs.Fields!lc05
         If Not IsNull(Rs.Fields!lc06) Then cboCaseName.AddItem "英:" + Rs.Fields!lc06
         If Not IsNull(Rs.Fields!lc07) Then cboCaseName.AddItem "日:" + Rs.Fields!lc07
      Else
         strCus = Rs.Fields!hc05
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strCus, strTempName) Then lbeCus = strCus: lbeCusName = strTempName
         If ClsPDGetCustomer(strCus, strTempName) Then lbeCus = strCus: lbeCusName = strTempName
         If Not IsNull(Rs.Fields!hc06) Then cboCaseName.AddItem "中:" + Rs.Fields!hc06
      End If
      cboCaseName.ListIndex = 0
   Else
     ' lbeCus = ""
     ' lbeCusName = ""
     ' cboCaseName.Clear
      MSHFlexGrid1.Rows = 2
      'cmdSure.Enabled = False
   End If

End Sub

Private Sub cmdSure_Click()
Dim i As Integer, blnChoese As Boolean
   
   If txtcp01.Text = Empty Or txtcp02.Text = Empty Then
      MsgBox "請輸入本所案號!", vbInformation, "庭期資料維護"
      Exit Sub
   End If
   If m_Lawcase = False Then
      MsgBox "基本檔無資料!", vbInformation, "庭期資料維護"
      Exit Sub
   End If
   
   blnChoese = False
   With MSHFlexGrid1
   For i = 1 To .Rows - 1
      If .TextMatrix(i, 0) = "v" Then
         blnChoese = True
         Exit For
      End If
   Next
   End With
   If Not blnChoese Then
      MsgBox "請點選輸入資料", vbCritical
      cmdSure.Enabled = False
      Exit Sub
   End If
   
   frm075012.Caption = Me.Caption
   frm075012.Show
   cmdSearch.SetFocus
   Me.Hide
End Sub

Private Sub Form_Activate()
   'txtcp01.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'cmdSearch.Enabled = False
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
End Sub
Private Sub GridHead()
Dim i As Integer

   With MSHFlexGrid1
      .Cols = 13
      blnOKtoShow = False
      .Visible = False
      .row = 0
      .col = 0
      .col = 0
      .ColWidth(0) = 200
      .col = 1
      .ColWidth(1) = 800
      .col = 2
      .ColWidth(2) = 1000
      .col = 3
      .ColWidth(3) = 1000
      .col = 4
      .ColWidth(4) = 900
      .col = 5
      .ColWidth(5) = 1200
      .col = 6
      .ColWidth(6) = 900
      .col = 7
      .ColWidth(7) = 900
      .col = 8
      .ColWidth(8) = 900
      .col = 9
      .ColWidth(9) = 0
      .col = 10
      .ColWidth(10) = 0
      .col = 11
      .ColWidth(11) = 0
      .col = 12
      .ColWidth(12) = 0
      
      intLastRow = 0
      blnOKtoShow = True
      
      .Visible = True
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm075011 = Nothing
End Sub

Private Sub lbeCus_Change()
Dim StrCusName As String
   
   If Len(lbeCus) > 8 Then
     'edit by nickc 2007/02/07 不用 dll 了
     'If objPublicData.GetCustomer(lbeCus, StrCusName) Then lbeCusName = StrCusName
     If ClsPDGetCustomer(lbeCus, StrCusName) Then lbeCusName = StrCusName
   End If

End Sub
Private Sub MSHFlexGrid1_Click()
   intCols = MSHFlexGrid1.Cols - 1
   If Not CheckGridChoese(MSHFlexGrid1, intLastRow, intCols) Then Exit Sub
   cmdSure.Enabled = True
   cmdSure.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then MSHFlexGrid1_Click
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
  
   txtcp01.Text = UCase(txtcp01)
   If IsEmptyText(txtcp01) = False Then
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      If CheckSys(txtcp01) <> "3" And CheckSys(txtcp01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         txtcp01_GotFocus
         Exit Sub
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         txtcp01_GotFocus
         Exit Sub
      End If
   End If
   If txtcp01 <> "" Then
       txtcp01 = UCase(txtcp01)
       If txtcp01 = "L" Or txtcp01 = "LA" Then
          blnCom1 = True
       Else
   '       DataErrorMessage 1, "系統類別"
   '       blnCom1 = False
   '       Cancel = True
          End If
   End If
   If Cancel Then TextInverse txtcp01

End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
   If txtcp02 <> "" Then
      blnCom2 = True
      cmdSearch.Enabled = True
   End If
   If Cancel Then TextInverse txtcp02

End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
   blnCom3 = True
   If Cancel Then TextInverse txtcp03

End Sub

Private Sub ChkCmd()
   If txtcp03 = "" Then blnCom3 = True: blnCom4 = True
   If blnCom1 And blnCom2 And blnCom3 And blnCom4 Then
         cmdSearch.SetFocus
   End If
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   blnCom4 = True
   ChkCmd
   If Cancel Then TextInverse txtcp04

End Sub
Private Function QueryDB() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
Dim bQuery As Boolean
   
   strTM01 = txtcp01.Text
   strTM02 = txtcp02.Text
   If txtcp03.Text <> Empty Then
      strTM03 = txtcp03.Text
   Else
      strTM03 = "0"
   End If
   If txtcp04.Text <> Empty Then
      strTM04 = txtcp04.Text
   Else
      strTM04 = "00"
   End If
   
   ' 依本所案號讀取基本檔案
   
   Select Case UCase(txtcp01.Text)
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(strTM01, strTM02, strTM03, strTM04)
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(strTM01, strTM02, strTM03, strTM04)
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(strTM01, strTM02, strTM03, strTM04)
      ' 讀取顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(strTM01, strTM02, strTM03, strTM04)
      ' 讀取服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(strTM01, strTM02, strTM03, strTM04)
   End Select
   QueryDB = bQuery
End Function

' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("TM05")
      End If
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("TM06")
      End If
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("TM23"), 0)
         lbeCus.Caption = rsTmp.Fields("TM23")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("SP05")
      End If
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("SP06")
      End If
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("SP08"), 0)
         lbeCus.Caption = rsTmp.Fields("SP08")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("PA07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("PA26"), 0)
         lbeCus.Caption = rsTmp.Fields("PA26")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("LC05")
      End If
      If IsNull(rsTmp.Fields("LC06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("LC06")
      End If
      If IsNull(rsTmp.Fields("LC07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("LC07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("LC11"), 0)
         lbeCus.Caption = rsTmp.Fields("LC11")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         cboCaseName.AddItem rsTmp.Fields("HC06")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("HC05"), 0)
         lbeCus.Caption = rsTmp.Fields("HC05")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
