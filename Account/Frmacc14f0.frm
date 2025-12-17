VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14f0 
   AutoRedraw      =   -1  'True
   Caption         =   "收文與收據資料檢核表"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   5160
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
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
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   840
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收文日期"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc14f0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoaccrpt101 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim lngCounter As Long
Dim dllaccrpt101 As Object

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt101Delete
   ProduceData
   If adoaccrpt101.State = adStateOpen Then
      adoaccrpt101.Close
   End If
   adoaccrpt101.CursorLocation = adUseClient
   adoaccrpt101.Open "select * from accrpt101", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt101.RecordCount <> 0 Then
      dllaccrpt101.Acc1411 ReportTitle(101), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
    'Modify By Cheng 2004/03/16
'   Else
'      MsgBox MsgText(28), , MsgText(5)
    'End
   End If
   adoaccrpt101.Close
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 1800
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = ""
   '92.10.7 CANCEL BY SONIA
   'MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   '92.10.7 CANCEL BY SONIA
   'MaskEdBox2.Text = CFDate(ACDate(ServerDate))
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt101 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt101 = Nothing
   Set Frmacc14f0 = Nothing
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt101Delete()
   adoTaie.Execute "delete from accrpt101"
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSql As String
Dim strNation As String
'Add By Cheng 2003/04/10
Dim StrSQLa As String

On Error GoTo Checking
   strSql = ""
   lngCounter = 0
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and cp05 >= " & Val(CADate(FCDate(MaskEdBox1.Text))) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and cp05 <= " & Val(CADate(FCDate(MaskEdBox2.Text))) & ""
   End If
    'Add By Cheng 2004/03/15
    '抓向客戶收款的資料
    strSql = strSql & " And CP20 Is Null "
    'End
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt101.CursorLocation = adUseClient
   adoaccrpt101.Open "select * from accrpt101", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoquery.CursorLocation = adUseClient
    'Modify By Cheng 2003/04/10
    '加收文日欄
'   adoquery.Open "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18 from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 is null and (cp16 is not null and cp16 <> 0)" & strSQL & " union " & _
'                 "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18 from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 and cp10 = cpm02 and cp60 = a0k01 (+) and a0k01 is null and (cp16 is not null and cp16 <> 0)" & strSQL & " order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Modify By Cheng 2004/03/16
    'FCP, FCT, FCL未發文的不抓
'   strSQLA = "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 is null and (cp16 is not null and cp16 <> 0) and cp57 is null" & strSQL & " union " & _
'             "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 and cp10 = cpm02 and cp60 = a0k01 (+) and substr(cp60, 1, 1) = 'E' and a0k01 is null and (cp16 is not null and cp16 <> 0) and cp57 is null" & strSQL & " order by cp05 asc, cp09 asc "
    'Modify By Cheng 2004/04/19
'   strSQLA = "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 is null and (cp16 is not null and cp16 <> 0) and cp57 is null And CP27=Decode(CP01,'FCP', Decode(CP27, Null, CP01, CP27),'FCT', Decode(CP27, Null, CP01, CP27), 'FCL', Decode(CP27, Null, CP01, CP27), CP27) " & strSQL & " union " & _
'             "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 and cp10 = cpm02 and cp60 = a0k01 (+) and substr(cp60, 1, 1) = 'E' and a0k01 is null and (cp16 is not null and cp16 <> 0) and cp57 is null And CP27=Decode(CP01,'FCP', Decode(CP27, Null, CP01, CP27),'FCT', Decode(CP27, Null, CP01, CP27), 'FCL', Decode(CP27, Null, CP01, CP27), CP27) " & strSQL & " order by cp05 asc, cp09 asc "
   '2008/10/21 modify by sonia 改抓智權人員姓名
   'strSQLA = "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 is null and (cp16 is not null and cp16 <> 0) and cp57 is null And 'CP27'||CP27='CP27'||Decode(CP01,'FCP', Decode(CP27, Null, CP01, CP27),'FCT', Decode(CP27, Null, CP01, CP27), 'FCL', Decode(CP27, Null, CP01, CP27), CP27) " & strSQL & " union " & _
             "select cp09, cp01, cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap, acc0k0 where cp01 = cpm01 and cp10 = cpm02 and cp60 = a0k01 (+) and substr(cp60, 1, 1) = 'E' and a0k01 is null and (cp16 is not null and cp16 <> 0) and cp57 is null And 'CP27'||CP27='CP27'||Decode(CP01,'FCP', Decode(CP27, Null, CP01, CP27),'FCT', Decode(CP27, Null, CP01, CP27), 'FCL', Decode(CP27, Null, CP01, CP27), CP27) " & strSQL & " order by cp05 asc, cp09 asc "
   StrSQLa = "select cp09, cp01, substr(st02,1,3) cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap, staff where cp01 = cpm01 and cp10 = cpm02 and cp13=st01(+) and cp60 is null and (cp16 is not null and cp16 <> 0) and cp57 is null And 'CP27'||CP27='CP27'||Decode(CP01,'FCP', Decode(CP27, Null, CP01, CP27),'FCT', Decode(CP27, Null, CP01, CP27), 'FCL', Decode(CP27, Null, CP01, CP27), CP27) " & strSql & " union " & _
             "select cp09, cp01, substr(st02,1,3) cp13, cpm03, cpm10, cp02, cp03, cp04, cp16, cp17, cp18, cp05 from caseprogress, casepropertymap, acc0k0, staff where cp01 = cpm01 and cp10 = cpm02 and cp13=st01(+) and cp60 = a0k01 (+) and substr(cp60, 1, 1) = 'E' and a0k01 is null and (cp16 is not null and cp16 <> 0) and cp57 is null And 'CP27'||CP27='CP27'||Decode(CP01,'FCP', Decode(CP27, Null, CP01, CP27),'FCT', Decode(CP27, Null, CP01, CP27), 'FCL', Decode(CP27, Null, CP01, CP27), CP27) " & strSql & " order by cp05 asc, cp09 asc "
    'End
    'End
   adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      adoaccrpt101.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoquery.EOF = False
      adoaccrpt101.AddNew
      adoaccrpt101.Fields("r10101").Value = strUserNum
      adoaccrpt101.Fields("r10111").Value = Counter
      adoaccrpt101.Fields("r10102").Value = adoquery.Fields("cp09").Value
      If IsNull(adoquery.Fields("cp01").Value) = False Then
         adoaccrpt101.Fields("r10103").Value = adoquery.Fields("cp01").Value & "-" & adoquery.Fields("cp02").Value & "-" & adoquery.Fields("cp03").Value & "-" & adoquery.Fields("cp04").Value
      End If
      If IsNull(adoquery.Fields("cp13").Value) = False Then
         adoaccrpt101.Fields("r10105").Value = adoquery.Fields("cp13").Value
      End If
      If IsNull(adoquery.Fields("cpm03").Value) = False Then
         adoaccrpt101.Fields("r10106").Value = adoquery.Fields("cpm03").Value
      Else
         If IsNull(adoquery.Fields("cpm10").Value) = False Then
            adoaccrpt101.Fields("r10106").Value = adoquery.Fields("cpm10").Value
         End If
      End If
      adocheck.CursorLocation = adUseClient
      
      'Modify by Morgan 2004/11/30 加抓申請人名稱
      'decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04)
'      adocheck.Open "select nvl(na03, na04) Nation, nvl(pa26, nvl(pa27, nvl(pa28, nvl(pa29, pa30)))) as Cust, cu10 from patent, nation, customer where pa09 = na01 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and pa01 = '" & adoquery.Fields("cp01").Value & "' and pa02 = '" & adoquery.Fields("cp02").Value & "' and pa03 = '" & adoquery.Fields("cp03").Value & "' and pa04 = '" & adoquery.Fields("cp04").Value & "' union " & _
'                    "select nvl(na03, na04) Nation, tm23 as Cust, cu10 from trademark, nation, customer where tm10 = na01 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and tm01 = '" & adoquery.Fields("cp01").Value & "' and tm02 = '" & adoquery.Fields("cp02").Value & "' and tm03 = '" & adoquery.Fields("cp03").Value & "' and tm04 = '" & adoquery.Fields("cp04").Value & "' union " & _
'                    "select nvl(na03, na04) Nation, lc11 as Cust, cu10 from lawcase, nation, customer where lc15 = na01 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and lc01 = '" & adoquery.Fields("cp01").Value & "' and lc02 = '" & adoquery.Fields("cp02").Value & "' and lc03 = '" & adoquery.Fields("cp03").Value & "' and lc04 = '" & adoquery.Fields("cp04").Value & "' union " & _
'                    "select '' Nation, hc05 as Cust, cu10 from hirecase, customer where substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and hc01 = '" & adoquery.Fields("cp01").Value & "' and hc02 = '" & adoquery.Fields("cp02").Value & "' and hc03 = '" & adoquery.Fields("cp03").Value & "' and hc04 = '" & adoquery.Fields("cp04").Value & "' union " & _
'                    "select nvl(na03, na04) Nation, nvl(sp08, nvl(sp58, sp59)) as Cust, cu10 from servicepractice, nation, customer where sp09 = na01 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and (sp01 not in ('S') or (sp01 in ('S') and sp09 >= '010')) and sp01 = '" & adoquery.Fields("cp01").Value & "' and sp02 = '" & adoquery.Fields("cp02").Value & "' and sp03 = '" & adoquery.Fields("cp03").Value & "' and sp04 = '" & adoquery.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocheck.Open "select nvl(na03, na04) Nation, nvl(pa26, nvl(pa27, nvl(pa28, nvl(pa29, pa30)))) as Cust, cu10,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) CustName from patent, nation, customer where pa09 = na01 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and pa01 = '" & adoquery.Fields("cp01").Value & "' and pa02 = '" & adoquery.Fields("cp02").Value & "' and pa03 = '" & adoquery.Fields("cp03").Value & "' and pa04 = '" & adoquery.Fields("cp04").Value & "' union " & _
                    "select nvl(na03, na04) Nation, tm23 as Cust, cu10,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) CustName from trademark, nation, customer where tm10 = na01 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and tm01 = '" & adoquery.Fields("cp01").Value & "' and tm02 = '" & adoquery.Fields("cp02").Value & "' and tm03 = '" & adoquery.Fields("cp03").Value & "' and tm04 = '" & adoquery.Fields("cp04").Value & "' union " & _
                    "select nvl(na03, na04) Nation, lc11 as Cust, cu10,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) CustName from lawcase, nation, customer where lc15 = na01 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and lc01 = '" & adoquery.Fields("cp01").Value & "' and lc02 = '" & adoquery.Fields("cp02").Value & "' and lc03 = '" & adoquery.Fields("cp03").Value & "' and lc04 = '" & adoquery.Fields("cp04").Value & "' union " & _
                    "select '' Nation, hc05 as Cust, cu10,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) CustName from hirecase, customer where substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and hc01 = '" & adoquery.Fields("cp01").Value & "' and hc02 = '" & adoquery.Fields("cp02").Value & "' and hc03 = '" & adoquery.Fields("cp03").Value & "' and hc04 = '" & adoquery.Fields("cp04").Value & "' union " & _
                    "select nvl(na03, na04) Nation, nvl(sp08, nvl(sp58, sp59)) as Cust, cu10,decode(cu04,Null,decode(cu05,Null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) CustName from servicepractice, nation, customer where sp09 = na01 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and (sp01 not in ('S') or (sp01 in ('S') and sp09 >= '010')) and sp01 = '" & adoquery.Fields("cp01").Value & "' and sp02 = '" & adoquery.Fields("cp02").Value & "' and sp03 = '" & adoquery.Fields("cp03").Value & "' and sp04 = '" & adoquery.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                    
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields("Cust").Value) = False Then
            adoaccrpt101.Fields("r10104").Value = adocheck.Fields("Cust").Value
         End If
         If IsNull(adocheck.Fields("Nation").Value) = False Then
            adoaccrpt101.Fields("r10107").Value = adocheck.Fields("Nation").Value
         End If
         If IsNull(adocheck.Fields("cu10").Value) = False Then
            strNation = adocheck.Fields("cu10").Value
         Else
            strNation = ""
         End If
         'Add by Morgan 2004/11/30 申請人名稱
         adoaccrpt101.Fields("r10112").Value = Left("" & adocheck.Fields("CustName").Value, 6)
      End If
      adocheck.Close
      If IsNull(adoquery.Fields("cp16").Value) = False Then
         adoaccrpt101.Fields("r10108").Value = adoquery.Fields("cp16").Value
      Else
         adoaccrpt101.Fields("r10108").Value = 0
      End If
      If IsNull(adoquery.Fields("cp17").Value) = False Then
         adoaccrpt101.Fields("r10109").Value = adoquery.Fields("cp17").Value
      Else
         adoaccrpt101.Fields("r10109").Value = 0
      End If
      If IsNull(adoquery.Fields("cp18").Value) = False Then
         adoaccrpt101.Fields("r10110").Value = adoquery.Fields("cp18").Value
      Else
         adoaccrpt101.Fields("r10110").Value = 0
      End If
      'Add By Cheng 2002/01/17
      '加寫入申請人
      'Modify by Morgan 2004/11/30 移到上面，不必重抓
      'adoaccrpt101.Fields("r10112").Value = Left("" & CaseCustNameQuery(CaseCustQuery(adoquery.Fields("cp09").Value)), 12)
      'Add By Cheng 2003/04/10
      
      '加寫入收文日
      adoaccrpt101.Fields("r10113").Value = ChangeWStringToWDateString("" & adoquery.Fields("cp05").Value)
      adoaccrpt101.UPDATE
      If Val(strNation) > 10 Then
         adoaccrpt101.Delete
      End If
      adoaccrpt101.UpdateBatch
      adoquery.MoveNext
   Loop
   adoquery.Close
   adoaccrpt101.Close
   adoTaie.Execute "delete from accrpt101 where substr(r10103, 1, 3) in ('FCT', 'FCP', 'FCL') and (r10107 is null or r10104 is null)"
   adoTaie.Execute "delete from accrpt101 where r10104 is null"
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
End Function

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   MaskEdBox2.Mask = ""
   '92.10.16 CANCEL BY SONIA
   'MaskEdBox2.Text = MaskEdBox1.Text
   MaskEdBox2.Mask = DFormat
End Sub
