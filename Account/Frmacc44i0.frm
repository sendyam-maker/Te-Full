VERSION 5.00
Begin VB.Form Frmacc44i0 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員別客戶扣繳稅款明細表"
   ClientHeight    =   3020
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3020
   ScaleWidth      =   5160
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   2520
      Width           =   3810
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   2
      Top             =   990
      Width           =   612
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1410
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   0
      Top             =   270
      Width           =   852
   End
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
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1950
      Width           =   4692
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2700
      TabIndex        =   13
      Top             =   660
      Width           =   915
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   180
      TabIndex        =   12
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
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
      Left            =   2460
      TabIndex        =   11
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "未收扣單 (Y)"
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
      TabIndex        =   10
      Top             =   1020
      Width           =   1590
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
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
      Left            =   3660
      TabIndex        =   9
      Top             =   690
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "是否產生Excel檔案(Y/N)"
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
      TabIndex        =   8
      Top             =   1410
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "智權人員編號"
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
      TabIndex        =   7
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      TabIndex        =   6
      Top             =   270
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc44i0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adostaff As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt114 As New ADODB.Recordset
Dim strSameName As String
Dim intCounter As Integer
Dim intPage As Integer
Dim strCustomerNo As String
Dim StrStaff As String
Dim lngCounter As Long
Private Const intDefault As Integer = 500
'Add by Morgan 2004/3/4
Dim PLeft(0 To 20) As Integer
Dim strPrinter As String 'Add By Sindy 2013/6/4
Dim prnPrint As Printer 'Add By Sindy 2014/3/6


Sub GetPleft()
    Erase PLeft
    '收款日期
    PLeft(0) = 300
    '收據號碼
    PLeft(1) = 1100
    '案件性質
    PLeft(2) = 2000
    '申請國家
    PLeft(3) = 3000
    '收款金額
    PLeft(4) = 4000
    '服務費
    PLeft(5) = 5000
    '可扣稅額
    PLeft(6) = 6000
    '收款扣繳額
    PLeft(7) = 7000
    '補扣繳額
    PLeft(8) = 8000
    '已收扣單金額
    PLeft(9) = 9000
    '調整稅額
    PLeft(10) = 10200
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      'Modify By Sindy 2025/3/21 mark
      'MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt114Delete
   ProduceData
   PUB_SetOsDefaultPrinter Combo1.Text 'Add By Sindy 2013/6/4
   PUB_RestorePrinter Combo1.Text 'Add by Sindy 2014/3/6
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify By Sindy 2020/5/8
   'adoquery.Open "select * from accrpt114 WHERE R11401='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "SELECT count(*) FROM" & _
                 " (SELECT * FROM accrpt114 WHERE R11401='" & strUserNum & "'" & _
                 " Union" & _
                 " select * from accrpt114_L WHERE R11401='" & strUserNum & "')", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 And adoquery.Fields(0) > 0 Then
      'Modify By Sindy 2020/5/8
      'FormPrint
      Call FormPrint("accrpt114")
      Call FormPrint("accrpt114_L")
      If Text3 = MsgText(602) Then
         'ExcelSave
         ExcelSaveMain
      End If
      '2020/5/8 END
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   PUB_SetOsDefaultPrinter strPrinter  'Add By Sindy 2013/6/4
   PUB_RestorePrinter strPrinter 'Add by Sindy 2014/3/6
   Screen.MousePointer = vbDefault
   FormClear
   MsgBox "執行完畢!!"
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
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
   Me.Height = 3420
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text3 = MsgText(602)
   'Add by Morgan 2004/3/4
   Text4 = MsgText(602)
   GetPleft
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Add By Sindy 2014/3/6
   '印表機設回預設印表機
   For Each prnPrint In Printers
      If prnPrint.DeviceName = strPrinter Then
         Set Printer = prnPrint
      End If
   Next
   '2014/3/6 END
   'Add By Sindy 2013/6/4
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2013/6/4 END
   
   Set Frmacc44i0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2020/6/19
Private Sub Text2_Validate(Cancel As Boolean)
   If Text2.Text <> "" Then
      Label5.Caption = ""
      Label5.Caption = GetPrjSalesNM(Text2)
      If Trim(Label5.Caption) = "" Then
         Text2.SetFocus
         MsgBox "員工編號輸入錯誤，無此人員！", , MsgText(5)
         Cancel = True
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = MsgText(602)
   Text4 = MsgText(602)
   Label5 = ""
   Text1.SetFocus
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
   Dim strSQL1 As String
   Dim strSQL2 As String
   'Add by Morgan 2004/3/4
   Dim stSQL As String
   'Add by Morgan 2004/3/17
   Dim stSalesNo As String
'   Dim stComp As String, stSID As String, stName As String
'   Dim arrSubtot(12 To 21) As String, arrTot(12 To 21) As String, idx As Integer
   Dim strSQL1k0 As String, stVTB As String 'Add By Sindy 2015/11/25
   'Add By Sindy 2017/3/10
   Dim m_CU16 As String '電話1
   Dim ConMan As String '聯絡人1
   Dim m_CU159 As String '會計備註
   Dim m_CU01 As String, m_CU02 As String
   '2017/3/10 END
   Dim rsA As New ADODB.Recordset
   Dim rsB As New ADODB.Recordset
   Dim stDept As String 'Add by Sindy 2017/3/13 部門
   Dim m_CU13Id As String, m_CU13Id_Tag As String, strR11402 As String, strR11408 As String, strR11421 As String
   Dim bolCancel As Boolean
   
On Error GoTo Checking
   
   lngCounter = 0
   
   '扣繳年度
   If Text1 <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a0k16=" & Val(Text1) & ""
      strSQL2 = strSQL2 & " and a1v09=" & Val(Text1) & ""
   End If
   
   '智權人員編號
   If Text2 <> MsgText(601) Then
      'Add By Sindy 2020/6/19 先檢查員編是否正確,以免資料輸錯
      Call Text2_Validate(bolCancel)
      If bolCancel = True Then
         Exit Sub
      End If
      
      'Modify By Sindy 2018/6/6 Mark,後面過濾掉
      'strSQL1 = strSQL1 & " and a0k20='" & Text2 & "'"
      
      'Add By Sindy 2015/11/25
      If Text2 = "F4101" Then
         'strSQL1k0 = strSQL1k0 & " and substr(s1.st15,1,2)='F3'"
         strSQL1k0 = strSQL1k0 & " and substr(cp12,1,2)='F3'"
      ElseIf Text2 = "F4102" Then
         'strSQL1k0 = strSQL1k0 & " and substr(s1.st15,1,2)='F2'"
         strSQL1k0 = strSQL1k0 & " and substr(cp12,1,2)='F2'"
      ElseIf Text2 = "F4103" Then
         'strSQL1k0 = strSQL1k0 & " and substr(s1.st15,1,2)='F1'"
         strSQL1k0 = strSQL1k0 & " and substr(cp12,1,2)='F1'"
      Else
         'strSQL1k0 = strSQL1k0 & " and cp13='" & Text2 & "'"
         strSQL1k0 = strSQL1k0 & " and cp13='" & Text2 & "'"
      End If
      '2015/11/25 END
   End If
   
   'Add by Morgan 2004/3/4 未收扣單(Y)
   If Text4 = MsgText(602) Then
      strSQL2 = strSQL2 & " and a1v15 is null "
   End If
   
   'Add By Sindy 2015/11/25
   stVTB = "SELECT a1k01,a1k35,a1k02" & _
           " From acc1k0,caseprogress,acc1v0,staff s1" & _
           " where a1k35 is not null" & _
           " and a1k01=cp60(+) and cp60 is not null" & _
           " and cp09=a1v01(+) and cp60=a1v02(+)" & _
           " and cp13=s1.st01(+)" & strSQL1k0 & strSQL2
   '2015/11/25 END
   
   'Modify by Morgan 2004/3/18
   '相同客戶(抬頭+客戶代碼前六碼相同者)跨不同智權人員時，先整理資料將此客戶全部改掛在智權人員編號最小且在職的智權人員上(單獨下智權人員條件不管)
   '分次收款時，收款日取最大者
   '同一收據抬頭不同客戶編號者不分別小計依起排序
   '每一收據抬頭下一列加印已收扣單資料，包括扣單編號、扣繳稅額及備註

   'Modify by Morgan 2006/10/18
   '日文聯絡人加部門 cu60-->cu114||' '||cu60
   'Modified by Morgan 2011/11/10 考慮拆收據情形,是否合併改抓a0j07
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modify by Sindy 2013/12/30 + and a0k11<>'J'
   'Modify By Sindy 2014/10/15 +R11422
   '                           +cu159
   'Modify By Sindy 2015/5/5 +and cu158 is null 屬境外公司者扣繳催收一律不用列出來
   'Modify By Sindy 2015/6/30 取消 CU158.境外公司 的控制
   '10000*ascii(st15)+100*ascii(substr(st15,2,1))+ascii(substr(st15,3,1)) c16 : 是為了sort 使用,轉成數字是為了save到table裡面
   'Modify By Sindy 2018/5/16 + and a1u02(+)=a1v02 ==> E106053152,E106053153 會出現重覆資料
   'Modify By Sindy 2018/5/16 substr(a0k03,1,6) c03 ==> substr(a0k03,1,8) c03
   'Modify By Sindy 2018/6/26 + R11424.收據日期
   'Modify By Sindy 2018/6/26 + R11425.公司別
   stSQL = "insert into accrpt114(r11401,r11417,R11421,r11402,r11403,r11404,r11405,r11406,r11407,r11408,r11409,r11410,r11411,r11412,r11418,r11419,r11413,r11420,r11415,r11414,R11422,R11424,R11425)" & _
           " select '" & strUserNum & "' c01, 0 c17, a0k22 c16" & _
           ",st04||a0k20 c02,substr(a0k03,1,8) c03, a0k04 c04, ConTel c05, ConMan c06, Ddate c07, ST02 c08, a0k01 c09, getcp10desc(cp01,cp10,a0j04) c10, na03, Ramount" & _
           ",decode(a0j07,'Y',Ramount,Fee1) Fee1, Fee2, Fee3, Fee4, Fee5, decode(a1v15,null,0, Fee3+Fee4) Fee6,cu159,a0k02,a0k11" & _
           " FROM ( select a0k01, a0k03, a0k04, a0k20, a1v01, a1v02, a1v15, nvl(cu16, cu17) as ConTel" & _
           ", nvl(cu58, nvl(cu59, cu114||' '||cu60)) as ConMan, st02, st04, a0k22,st15, a0l02 as Ddate" & _
           ", a1v04 Fee2, decode(a1v18,'1',a1v06,0) Fee3, decode(a1v18,null,a1v06,0) Fee4, a1v10 Fee5, cu159,a0k02,a0k11" & _
           " From acc0k0, acc1v0, customer, staff,(select a0m02,max(a0l02) a0l02 from acc0k0, acc0m0, acc0l0" & _
           " where a0m02(+) = a0k01 and a0l01(+) = a0m01 and a0m02 is not null and a0k11<>'J'" & strSQL1 & _
           " group by a0m02) a" & _
           " where a1v02=a0k01 and a1v06>0 " & strSQL1 & strSQL2 & _
           " and substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11<>'J'" & _
           " and a0k20 = st01(+) and a0m02(+) = a0k01" & _
           " ) X, ( select a1u02, a1u03" & _
           ", sum(nvl(a1u04, 0)+nvl(a1u05, 0)-nvl(a1u08, 0)-nvl(a1u10, 0)) as Ramount" & _
           ", sum(nvl(a1u04, 0)-nvl(a1u08, 0)) as Fee1" & _
           " From acc0k0, acc1v0, acc1u0" & _
           " where a1v02=a0k01 and a1v06>0 " & strSQL1 & strSQL2 & _
           " and a1v02=a1u02(+) and a1v01=a1u03(+) and a0k11<>'J'" & _
           " group by a1u02, a1u03 ) Y, acc0j0 Z,caseprogress,nation" & _
           " where a1u03(+)=a1v01 and a1u02(+)=a1v02 and a0j01(+) = a1v01 and a0j13(+)=a1v02 and Ramount<>0 and cp09(+)=a1v01 and na01(+)=a0j04 "
   'Add By Sindy 2015/11/25
   'and decode(substr(s1.st15,1,2),'F3','F4101','F2','F4102','F1','F4103',s1.st01)=s2.st01(+) ==> and decode(substr(cp12,1,2),'F3','F4101','F2','F4102','F1','F4103',s1.st01)=s2.st01(+)
   'Modify By Sindy 2018/5/29 union ==> union all 不然會有資料遺漏
   'Modify By Sindy 2020/6/4 + a1k37.公司別 = 改抓a1v03
   stSQL = stSQL & " union all " & _
           " select '" & strUserNum & "' c01,0 c17,cp12 c16" & _
           ",'1'||s2.st01 c02,substr(a1k28,1,8) c03,a1k35 c04,nvl(cu16, cu17) c05,nvl(cu58, nvl(cu59, cu114||' '||cu60)) c06,a0y02 c07,s2.ST02 c08,a1k01 c09,GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04),GETNA03DESCCaseNO(cp01,cp02,cp03,cp04),A1k30 Ramount" & _
           ",nvl(A1k30,0)-nvl(A1k09,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3,decode(a1v18,null,nvl(a1v06,0),0) Fee4,nvl(a1v10,0) Fee5,decode(a1v15,null,0,nvl(a1v06,0)) Fee6,cu159,a1k02,a1v03" & _
           " from (select a0z02,max(a0y02) a0y02 from (" & stVTB & ") V1,acc0z0,acc0y0" & _
           " where a0z02=a1k01 and a0y01=a0z01 group by a0z02) a" & _
           ",acc1k0,acc1v0,caseprogress,staff s1,staff s2,customer" & _
           " Where a1k35 Is Not Null" & _
           " and a0z02=a1k01" & _
           " and a1k01=a1v02(+)" & _
           " and a1v01=cp09(+) and a1v06>0" & _
           " and substr(a1k28,1,8)=cu01 and substr(a1k28,9,1)=cu02" & _
           " and cp13=s1.st01(+)" & _
           " and decode(substr(cp12,1,2),'F3','F4101','F2','F4102','F1','F4103',cp13)=s2.st01(+)"
   'and decode(substr(s1.st15,1,2),'F3','F4101','F2','F4102','F1','F4103',s1.st01)=s2.st01(+) ==> and decode(substr(cp12,1,2),'F3','F4101','F2','F4102','F1','F4103',s1.st01)=s2.st01(+)
   'Modify By Sindy 2018/5/29 union ==> union all 不然會有資料遺漏
   'Modify By Sindy 2020/6/4 + a1k37.公司別 = 改抓a1v03
   stSQL = stSQL & " union all " & _
           " select '" & strUserNum & "' c01,0 c17,cp12 c16" & _
           ",'1'||s2.st01 c02,substr(a1k28,1,8) c03,a1k35 c04,nvl(fa12, fa13) c05,nvl(fa07, nvl(fa08, fa78||' '||fa09)) c06,a0y02 c07,s2.ST02 c08,a1k01 c09,GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04),GETNA03DESCCaseNO(cp01,cp02,cp03,cp04),A1k30 Ramount" & _
           ",nvl(A1k30,0)-nvl(A1k09,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3,decode(a1v18,null,nvl(a1v06,0),0) Fee4,nvl(a1v10,0) Fee5,decode(a1v15,null,0,nvl(a1v06,0)) Fee6,fa118,a1k02,a1v03" & _
           " from (select a0z02,max(a0y02) a0y02 from (" & stVTB & ") V1,acc0z0,acc0y0" & _
           " where a0z02=a1k01 and a0y01=a0z01 group by a0z02) a" & _
           ",acc1k0,acc1v0,caseprogress,staff s1,staff s2,fagent" & _
           " Where a1k35 Is Not Null" & _
           " and a0z02=a1k01" & _
           " and a1k01=a1v02(+)" & _
           " and a1v01=cp09(+) and a1v06>0" & _
           " and substr(a1k28,1,8)=fa01 and substr(a1k28,9,1)=fa02" & _
           " and cp13=s1.st01(+)" & _
           " and decode(substr(cp12,1,2),'F3','F4101','F2','F4102','F1','F4103',cp13)=s2.st01(+)"
   '2015/11/25 END
   adoTaie.Execute stSQL
    
   'Modify By Sindy 2020/5/8
   '另存法律所資料
   strSql = "insert into accrpt114_L select * from accrpt114 where r11401='" & strUserNum & "' and r11425='L'"
   cnnConnection.Execute strSql
   '刪除法律所資料,只留智慧所
   strSql = "delete from accrpt114 where r11401='" & strUserNum & "' and r11425='L'"
   cnnConnection.Execute strSql
   
   '***********************************************************************
   'R11402:智權人員ID
   'R11403:客戶編號
   'R11404:收據抬頭
   'R11405:聯絡電話
   'R11406:聯絡人
   'R11408:智權人員Name
   'R11409:收據編號
   'R11421:業務區
   'R11422:會計備註
   'R11423:客戶編號
   'R11424:收據日期
   'R11425:公司別/Y.案源檔
   '***********************************************************************
   'Add By Sindy 2020/6/19 同收據抬頭有不同的智權人員
   '智慧所:
   strExc(0) = "select R11404 from(SELECT distinct R11404,substr(r11402,2) SID FROM accrpt114 WHERE R11401='" & strUserNum & "')" & _
               " GROUP BY R11404 having count(*)>1"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         strExc(0) = "select distinct r11404,r11402,r11408,r11421 from accrpt114" & _
                     " where r11401='" & strUserNum & "' and r11404='" & rsA.Fields("R11404") & "'" & _
                     " and r11409=(select max(a1.r11409) from accrpt114 a1 where a1.r11401='" & strUserNum & "' and a1.r11404='" & rsA.Fields("R11404") & "'" & _
                     " and a1.r11424=(select max(a2.r11424) from accrpt114 a2 where a2.r11401='" & strUserNum & "' and a2.r11404='" & rsA.Fields("R11404") & "'))"
         intI = 1
         Set rsB = ClsLawReadRstMsg(intI, strExc(0))
         strR11402 = "": strR11408 = "": strR11421 = ""
         If intI = 1 Then
            strR11402 = rsB.Fields("r11402")
            strR11408 = rsB.Fields("r11408")
            strR11421 = rsB.Fields("r11421")
            strSql = "update accrpt114 set" & _
                     " R11402=" & CNULL(strR11402) & _
                     ",R11408=" & CNULL(strR11408) & _
                     ",R11421=" & CNULL(strR11421) & _
                     " WHERE R11401='" & strUserNum & "'" & _
                     " and R11404='" & rsA.Fields("R11404") & "'"
            cnnConnection.Execute strSql
         End If
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   
   '法律所:
   '有案源檔智權人員抓第一個介紹人
   strExc(0) = "SELECT distinct R11404,R11409,R11402,st04||substr(LOS04,1,5) SID,st02,st15" & _
               " FROM accrpt114_L,acc0j0,lawofficesource,staff" & _
               " WHERE R11401='" & strUserNum & "' AND R11409=a0j13" & _
               " AND a0j01=los06" & _
               " AND substr(LOS04,1,5)=st01"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         strR11402 = rsA.Fields("SID")
         strR11408 = rsA.Fields("st02")
         strR11421 = rsA.Fields("st15")
         strSql = "update accrpt114_L set" & _
                  " R11402=" & CNULL(strR11402) & _
                  ",R11408=" & CNULL(strR11408) & _
                  ",R11421=" & CNULL(strR11421) & _
                  ",R11425='Y'" & _
                  " WHERE R11401='" & strUserNum & "'" & _
                  " and R11409='" & rsA.Fields("R11409") & "'"
         cnnConnection.Execute strSql
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   '無案源檔,同智慧所做法
   strExc(0) = "select R11404 from(SELECT distinct R11404,substr(r11402,2) SID FROM accrpt114_L WHERE R11401='" & strUserNum & "' AND R11425<>'Y')" & _
               " GROUP BY R11404 having count(*)>1"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         strExc(0) = "select distinct r11404,r11402,r11408,r11421 from accrpt114_L" & _
                     " where r11401='" & strUserNum & "' and r11404='" & rsA.Fields("R11404") & "' AND R11425<>'Y'" & _
                     " and r11409=(select max(a1.r11409) from accrpt114_L a1 where a1.r11401='" & strUserNum & "' and a1.r11404='" & rsA.Fields("R11404") & "' AND R11425<>'Y'" & _
                     " and a1.r11424=(select max(a2.r11424) from accrpt114_L a2 where a2.r11401='" & strUserNum & "' and a2.r11404='" & rsA.Fields("R11404") & "' AND R11425<>'Y'))"
         intI = 1
         Set rsB = ClsLawReadRstMsg(intI, strExc(0))
         strR11402 = "": strR11408 = "": strR11421 = ""
         If intI = 1 Then
            strR11402 = rsB.Fields("r11402")
            strR11408 = rsB.Fields("r11408")
            strR11421 = rsB.Fields("r11421")
            strSql = "update accrpt114_L set" & _
                     " R11402=" & CNULL(strR11402) & _
                     ",R11408=" & CNULL(strR11408) & _
                     ",R11421=" & CNULL(strR11421) & _
                     " WHERE R11401='" & strUserNum & "'" & _
                     " and R11404='" & rsA.Fields("R11404") & "' AND R11425<>'Y'"
            cnnConnection.Execute strSql
         End If
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   '2020/6/19 END
   
   '更新資料後,再過濾智權人員資料
   If Text2 <> MsgText(601) Then
      strSql = "delete from accrpt114 where r11401='" & strUserNum & "' and substr(R11402,2)<>'" & Text2 & "'"
      cnnConnection.Execute strSql
      'Add By Sindy 2020/6/19
      strSql = "delete from accrpt114_L where r11401='" & strUserNum & "' and substr(R11402,2)<>'" & Text2 & "'"
      cnnConnection.Execute strSql
      '2020/6/19 END
   End If
   
   'Modify By Sindy 2018/5/18 收據抬頭要排一起 ex:國立成功大學
   'order by r11403,r11404,r11402
   'Modify By Sindy 2018/5/18 原:r11404,r11403,r11402 ==> r11404,r11402,R11407,R11409
'   stSQL = "select a.*,substr(r11402,2) sid from accrpt114 a where r11401='" & strUserNum & "' and r11417=0" & _
'           " order by r11404,r11402,R11407,R11409"
'   adoacc0k0.CursorLocation = adUseClient
'   adoacc0k0.Open stSQL, adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoacc0k0.RecordCount = 0 Then
'      adoacc0k0.Close
'      If adoaccrpt114.State = adStateOpen Then adoaccrpt114.Close
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   End If
'   adoacc0k0.Close
   '檢查是否有資料
   stSQL = "select count(*) from (" & _
           " SELECT r11404 SID FROM accrpt114 WHERE r11401='" & strUserNum & "'" & _
           " Union All" & _
           " SELECT r11404 FROM accrpt114_L WHERE r11401='" & strUserNum & "'" & _
           ")"
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open stSQL, adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0k0.Fields(0) = 0 Then
      adoacc0k0.Close
      If adoaccrpt114.State = adStateOpen Then adoaccrpt114.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   adoacc0k0.Close
   
   'Add By Sindy 2017/3/10 依收據抬頭重新讀取電話號碼和聯絡人和會計備註
   'Modify By Sindy 2020/6/20 + accrpt114_L
   strExc(0) = "select r11403||'0' cuid,R11403,R11404 from accrpt114 WHERE R11401='" & strUserNum & "'" & _
               " group by r11403||'0',R11403,R11404" & _
               " union " & _
               "select r11403||'0' cuid,R11403,R11404 from accrpt114_L WHERE R11401='" & strUserNum & "'" & _
               " group by r11403||'0',R11403,R11404"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         Call GetTitleCustData(rsA.Fields("R11404"), rsA.Fields("cuid"), "", m_CU01, m_CU02, _
                            , , , , , , m_CU16, _
                            , "", , , , , , _
                            m_CU159, , , , , , , , , ConMan, , , , m_CU13Id)
         'If m_CU16 <> "" Or m_CU159 <> "" Or ConMan <> "" Then
            If m_CU01 = "" Then
               m_CU01 = rsA.Fields("R11403") '& "00"
               m_CU02 = "0"
            End If
            'Modify By Sindy 2020/12/10 + ChgSQL
            strSql = "update accrpt114 set" & _
                     " R11405=" & CNULL(m_CU16) & _
                     ",R11406=" & CNULL(ConMan) & _
                     ",R11422=" & CNULL(ChgSQL(m_CU159)) & _
                     ",R11423=" & CNULL(m_CU01 & m_CU02) & _
                     " WHERE R11401='" & strUserNum & "'" & _
                     " and R11404='" & rsA.Fields("R11404") & "'" & _
                     " and R11403='" & rsA.Fields("R11403") & "'"
            cnnConnection.Execute strSql
            'Add By Sindy 2020/6/20
            'Modify By Sindy 2020/12/10 + ChgSQL
            strSql = "update accrpt114_L set" & _
                     " R11405=" & CNULL(m_CU16) & _
                     ",R11406=" & CNULL(ConMan) & _
                     ",R11422=" & CNULL(ChgSQL(m_CU159)) & _
                     ",R11423=" & CNULL(m_CU01 & m_CU02) & _
                     " WHERE R11401='" & strUserNum & "'" & _
                     " and R11404='" & rsA.Fields("R11404") & "'" & _
                     " and R11403='" & rsA.Fields("R11403") & "'"
            cnnConnection.Execute strSql
            '2020/6/20 END
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   '2017/3/10 END
   
   Call CountData("accrpt114")
   Call CountData("accrpt114_L")
'    Set adoaccrpt114 = adoacc0k0.Clone
'    With adoacc0k0
'        '更新同客戶不同智權人員為最小且在職的智權人員
'        'Modify By Sindy 2017/6/7 不更新智權人員資料,業務區改抓 a0k22 or cp12
'        .MoveFirst
''        stComp = .Fields("r11403") & .Fields("r11404")
''        stSID = .Fields("r11402")
''        stName = .Fields("r11408")
''        stDept = .Fields("R11421") 'Add by Sindy 2017/3/13 部門
''        .MoveNext
''        Do While Not .EOF
''            If stComp = .Fields("r11403") & .Fields("r11404") Then
''                If stSID <> .Fields("r11402") Then
''                    .Fields("r11402") = stSID
''                    .Fields("r11408") = stName
''                    .Fields("R11421") = stDept 'Add by Sindy 2017/3/13 不更新部門,後續資料不會排在一起
''                End If
''            Else
''                stComp = .Fields("r11403") & .Fields("r11404")
''                stSID = .Fields("r11402")
''                stName = .Fields("r11408")
''                stDept = .Fields("R11421") 'Add by Sindy 2017/3/13 部門
''            End If
''            .MoveNext
''        Loop
''        .UpdateBatch
'
'        '加小計合計並排序
'        .Requery
'        'Modify By Sindy 2018/5/18 收據抬頭要排一起 ex:國立成功大學
'        '.Sort = "R11421, sid, r11403, r11404, r11407, r11409, r11410"
'        'Modify By Sindy 2018/5/29
'        '.Sort = "R11421, sid, r11404, r11403, r11407, r11409, r11410"
'        .Sort = "R11421, sid, r11404, r11407, r11409, r11410"
'        .MoveFirst
'        stComp = (.Fields("sid").Value & adoacc0k0.Fields("r11404").Value)
'        stSID = .Fields("sid")
'        Erase arrSubtot()
'        Erase arrTot()
'        Do While Not .EOF
'            .Fields("r11417").Value = Counter
'            '會加總
'            For idx = 12 To 20
'                If Not (idx = 14 And ("" & .Fields("r11414")) = "") Then
'                    arrSubtot(idx) = Format(Val(arrSubtot(idx)) + Val("" & .Fields("r114" & Format(idx, "00"))))
'                    arrTot(idx) = Format(Val(arrTot(idx)) + Val("" & .Fields("r114" & Format(idx, "00"))))
'                End If
'            Next
'            'Add By Sindy 2017/6/7 部門別
'            arrSubtot(21) = adoacc0k0.Fields("r11421").Value
'            arrTot(21) = adoacc0k0.Fields("r11421").Value
'            .MoveNext
'            If Not .EOF Then
'                If stComp <> (.Fields("sid").Value & .Fields("r11404").Value) Then
'                    Call AddSubTot(arrSubtot())
'                    stComp = (.Fields("sid").Value & .Fields("r11404").Value)
'                    Erase arrSubtot()
'                End If
'                If stSID <> .Fields("sid") Then
'                    Call AddSubTot(arrTot(), 2)
'                    stSID = .Fields("sid")
'                    Erase arrTot()
'                End If
'            Else
'                Call AddSubTot(arrSubtot())
'                Call AddSubTot(arrTot(), 2)
'            End If
'        Loop
'        .UpdateBatch
'        .Close
'    End With
'    If adoaccrpt114.State = adStateOpen Then adoaccrpt114.Close
'
'    stSQL = "Update accrpt114 set r11402=substr(r11402,2) where r11401='" & strUserNum & "'"
'    adoTaie.Execute stSQL
   '2020/5/8 END
   
   Set rsA = Nothing
   Set rsB = Nothing
   Exit Sub
      
Checking:
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
   'Resume
End Sub

'Modify By Sindy 2020/5/8
Private Sub CountData(strTableName As String)
Dim stSQL As String
Dim stComp As String, stSID As String
Dim arrSubtot(12 To 21) As String, arrTot(12 To 21) As String, idx As Integer

On Error GoTo Checking

'   stSQL = "select a.*,substr(r11402,2) sid from " & strTableName & " a where r11401='" & strUserNum & "' and r11417=0" & _
'           " order by r11404,r11402,R11407,R11409"
   stSQL = "select a.*,substr(r11402,2) sid from " & strTableName & " a where r11401='" & strUserNum & "'" & _
           " order by r11404,r11402,R11407,R11409"
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open stSQL, adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0k0.RecordCount > 0 Then
      Set adoaccrpt114 = adoacc0k0.Clone
      With adoacc0k0
          '更新同客戶不同智權人員為最小且在職的智權人員
          'Modify By Sindy 2017/6/7 不更新智權人員資料,業務區改抓 a0k22 or cp12
          .MoveFirst
          '加小計合計並排序
          .Requery
          'Modify By Sindy 2018/5/18 收據抬頭要排一起 ex:國立成功大學
          '.Sort = "R11421, sid, r11403, r11404, r11407, r11409, r11410"
          'Modify By Sindy 2018/5/29
          '.Sort = "R11421, sid, r11404, r11403, r11407, r11409, r11410"
          .Sort = "R11421, sid, r11404, r11407, r11409, r11410"
          .MoveFirst
          stComp = (.Fields("sid").Value & adoacc0k0.Fields("r11404").Value)
          stSID = .Fields("sid")
          Erase arrSubtot()
          Erase arrTot()
          Do While Not .EOF
              .Fields("r11417").Value = Counter
              '會加總
              For idx = 12 To 20
                  If Not (idx = 14 And ("" & .Fields("r11414")) = "") Then
                      arrSubtot(idx) = Format(Val(arrSubtot(idx)) + Val("" & .Fields("r114" & Format(idx, "00"))))
                      arrTot(idx) = Format(Val(arrTot(idx)) + Val("" & .Fields("r114" & Format(idx, "00"))))
                  End If
              Next
              'Add By Sindy 2017/6/7 部門別
              arrSubtot(21) = adoacc0k0.Fields("r11421").Value
              arrTot(21) = adoacc0k0.Fields("r11421").Value
              .MoveNext
              If Not .EOF Then
                  If stComp <> (.Fields("sid").Value & .Fields("r11404").Value) Then
                      Call AddSubTot(arrSubtot())
                      stComp = (.Fields("sid").Value & .Fields("r11404").Value)
                      Erase arrSubtot()
                  End If
                  If stSID <> .Fields("sid") Then
                      Call AddSubTot(arrTot(), 2)
                      stSID = .Fields("sid")
                      Erase arrTot()
                  End If
              Else
                  Call AddSubTot(arrSubtot())
                  Call AddSubTot(arrTot(), 2)
              End If
          Loop
          .UpdateBatch
          .Close
      End With
      If adoaccrpt114.State = adStateOpen Then adoaccrpt114.Close
   End If
   
   stSQL = "Update " & strTableName & " set r11402=substr(r11402,2) where r11401='" & strUserNum & "'"
   adoTaie.Execute stSQL
   
   'adoacc0k0.Close
   Set adoacc0k0 = Nothing 'Add By Sindy 2020/6/20
   Exit Sub
   
Checking:
   Set adoacc0k0 = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
   'Resume
End Sub

Private Sub AddSubTot(ByRef arrSubtot() As String, Optional iMode As Integer = 1)
    With adoaccrpt114
        .AddNew
        .Fields("r11401").Value = strUserNum
        .Fields("r11417").Value = Counter
        .Fields("r11412").Value = arrSubtot(12)
        .Fields("r11413").Value = arrSubtot(13)
        .Fields("r11414").Value = IIf(arrSubtot(14) = "", Null, arrSubtot(14))
        .Fields("r11415").Value = arrSubtot(15)
        '.Fields("R11416").Value = arrSubtot(16)
        .Fields("R11421").Value = arrSubtot(21)
        .Fields("r11418").Value = arrSubtot(18)
        .Fields("r11419").Value = arrSubtot(19)
        .Fields("r11420").Value = arrSubtot(20)
        If iMode = 1 Then
            .Fields("r11411").Value = ReportSum(24)
        Else
            .Fields("r11411").Value = ReportSum(25)
        End If
        .UpdateBatch
    End With
End Sub

''*************************************************
''  列印類別小計計算
''
''*************************************************
'Private Sub SubSelect(Optional ByVal iType As Integer = 1, Optional ByVal stSalesNo As String = "")
'
'   Dim stSQL As String
'
'   adoacc0k0.MovePrevious
'   adoaccrpt114.AddNew
'   adoaccrpt114.Fields("r11401").Value = strUserNum
'   adoaccrpt114.Fields("r11417").Value = Counter
'
'    If iType = 1 Then
'        adoaccrpt114.Fields("r11411").Value = ReportSum(24)
'        stSQL = "select sum(r11412) r12, sum(r11413) r13, sum(r11414) r14, sum(r11415) r15, sum(R11421) r16, sum(r11418) r18, sum(r11419) r19, sum(r11420) r20 from accrpt114 where r11401 = '" & strUserNum & "' and r11402||r11403||r11404 = '" & strSameName & "'"
'    ElseIf iType = 2 Then
'        adoaccrpt114.Fields("r11411").Value = ReportSum(25)
'        stSQL = "select sum(r11412) r12, sum(r11413) r13, sum(r11414) r14, sum(r11415) r15, sum(R11421) r16, sum(r11418) r18, sum(r11419) r19, sum(r11420) r20 from accrpt114 where r11401 = '" & strUserNum & "' and r11402='" & stSalesNo & "'"
'    End If
'    adoquery.CursorLocation = adUseClient
'   adoquery.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      adoaccrpt114.Fields("r11412").Value = 0 + Val("" & adoquery.Fields("r12").Value)
'      adoaccrpt114.Fields("r11413").Value = 0 + Val("" & adoquery.Fields("r13").Value)
'      adoaccrpt114.Fields("r11414").Value = adoquery.Fields("r14").Value
'      adoaccrpt114.Fields("r11415").Value = 0 + Val("" & adoquery.Fields("r15").Value)
'      adoaccrpt114.Fields("R11421").Value = 0 + Val("" & adoquery.Fields("r16").Value)
'      adoaccrpt114.Fields("r11418").Value = 0 + Val("" & adoquery.Fields("r18").Value)
'      adoaccrpt114.Fields("r11419").Value = 0 + Val("" & adoquery.Fields("r19").Value)
'      adoaccrpt114.Fields("r11420").Value = 0 + Val("" & adoquery.Fields("r20").Value)
'   End If
'   adoquery.Close
'   adoaccrpt114.UpdateBatch
'   adoacc0k0.MoveNext
'End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead(Optional stSalesName As String = "", Optional strTableName As String = "")
   'Add By Sindy 2020/6/19
   If strTableName = "accrpt114_L" Then '法律所
      Printer.FontSize = 16
      Printer.CurrentX = 1000
      Printer.CurrentY = 1000 - intDefault
      Printer.Print "(法律所)"
   End If
   '2020/6/19 END
   Printer.FontSize = 16
   Printer.CurrentX = 3000
   Printer.CurrentY = 1000 - intDefault
   Printer.Print ReportTitle(114)
   Printer.FontSize = 9
   Printer.CurrentX = 5000
   Printer.CurrentY = 1500 - intDefault
   Printer.Print "扣繳年度: "
   Printer.CurrentX = 5900
   Printer.CurrentY = 1500 - intDefault
   Printer.Print Text1
   If Text4 <> "Y" Then
        Printer.CurrentX = 5000
        Printer.CurrentY = 1800 - intDefault
        Printer.Print "（含已收扣單）"
   End If
   Printer.CurrentX = 300
   Printer.CurrentY = 2100 - intDefault
   Printer.Print "列印人員: "
   Printer.CurrentX = 1200
   Printer.CurrentY = 2100 - intDefault
   Printer.Print StaffQuery(strUserNum)
   Printer.CurrentX = 9000
   Printer.CurrentY = 2100 - intDefault
   Printer.Print "列印日期: "
   Printer.CurrentX = 10000
   Printer.CurrentY = 2100 - intDefault
   Printer.Print IIf(Mid(CFDate(strSrvDate(2)), 1, 1) = "0", Mid(CFDate(strSrvDate(2)), 2, 8), CFDate(strSrvDate(2)))
   Printer.CurrentX = 300
   Printer.CurrentY = 2400 - intDefault
   Printer.Print "智權人員: "
   Printer.CurrentX = 1200
   Printer.CurrentY = 2400 - intDefault
   
   'Modify by Morgan 2004/3/17
   '當印小計跳頁時，智權人員會空白
'   If IsNull(adoaccrpt114.Fields("r11402").Value) = False Then
'      'Printer.Print StaffQuery(adoaccrpt114.Fields("r11402").Value)
'      Printer.Print adoaccrpt114.Fields("r11408").Value
'   Else
'      Printer.Print ""
'   End If
   Printer.Print stSalesName
   
   Printer.CurrentX = 9000
   Printer.CurrentY = 2400 - intDefault
   Printer.Print "頁次: "
   Printer.CurrentX = 10000
   Printer.CurrentY = 2400 - intDefault
   Printer.Print intPage
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "收據抬頭: "
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt114.Fields("r11404").Value) = False Then
'      Printer.Print adoaccrpt114.Fields("r11404").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 7100
'   Printer.CurrentY = 2700
'   Printer.Print "聯絡電話: "
'   Printer.CurrentX = 8400
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt114.Fields("r11405").Value) = False Then
'      Printer.Print adoaccrpt114.Fields("r11405").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 9700
'   Printer.CurrentY = 2700
'   Printer.Print "聯絡人: "
'   Printer.CurrentX = 11000
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt114.Fields("r11406").Value) = False Then
'      Printer.Print adoaccrpt114.Fields("r11406").Value
'   Else
'      Printer.Print ""
'   End If

    'Modify by Morgan 2004/3/4
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "收款日期"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "收據號碼"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "申請國家"
   '數字欄位抬頭靠右
   Printer.CurrentX = PLeft(4) + 900 - Printer.TextWidth("收款金額")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "收款金額"
   Printer.CurrentX = PLeft(5) + 900 - Printer.TextWidth("服務費")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "服務費"
   Printer.CurrentX = PLeft(6) + 900 - Printer.TextWidth("可扣稅額")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "可扣稅額"
   Printer.CurrentX = PLeft(7) + 900 - Printer.TextWidth("收款扣繳額")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "收款扣繳額"
   Printer.CurrentX = PLeft(8) + 900 - Printer.TextWidth("補扣繳額")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "補扣繳額"
   Printer.CurrentX = PLeft(9) + 1100 - Printer.TextWidth("已收扣單金額")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "已收扣單金額"
   Printer.CurrentX = PLeft(10) + 900 - Printer.TextWidth("調整稅額")
   Printer.CurrentY = 3000 - intDefault
   Printer.Print "調整稅額"
   Printer.Line (300, 3400 - intDefault)-(12000, 3400 - intDefault)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt114Delete()
   adoTaie.Execute "delete from accrpt114 WHERE R11401='" & strUserNum & "'"
   adoTaie.Execute "delete from accrpt114_L WHERE R11401='" & strUserNum & "'" 'Add By Sindy 2020/5/8
End Sub

'列印扣單資料
Private Function PrintTaxData(ByVal stYear As String, ByVal stCustNo As String, ByVal stCustName As String, stSalesName As String, Optional ByVal stSalesNo As String = "", Optional strTableName As String = "") As Boolean
Dim stSQL As String, rsQuery As New ADODB.Recordset, stCon As String, bol1st As Boolean
'add by nickc 2007/02/08
Dim strAmount As String
Dim intLength As Integer
Dim stConA1k As String
   
On Error GoTo flgErr
   
'    'Modified by Morgan 2011/11/10 員工編號會有非數字要加單引號
'    If Trim(stSalesNo) <> "" Then
'      stCon = " And a0k20='" & stSalesNo & "'"
'      'Add By Sindy 2015/11/25
'      If Text2 = "F4101" Then
'         'stConA1k = " and substr(s1.st15,1,2)='F3'"
'         stConA1k = " and substr(cp12,1,2)='F3'"
'      ElseIf Text2 = "F4102" Then
'         'stConA1k = " and substr(s1.st15,1,2)='F2'"
'         stConA1k = " and substr(cp12,1,2)='F2'"
'      ElseIf Text2 = "F4103" Then
'         'stConA1k = " and substr(s1.st15,1,2)='F1'"
'         stConA1k = " and substr(cp12,1,2)='F1'"
'      Else
'         stConA1k = " and cp13='" & stSalesNo & "'"
'      End If
'      '2015/11/25 END
'    End If
'
'    'Modify By Sindy 2013/12/30 + and a0k11<>'J'
''    stSQL = " select a0w02, a0w05, a0w06 from acc0w0" & _
''            " where a0w01=" & stYear & " and a0w02 in (" & _
''            " select a1v15 from acc0k0, acc1v0" & _
''            " Where a1v02=a0k01 and a1v15 is not null " & stCon & _
''            " and a0k16 = " & stYear & " And substr(a0k03, 1, 6) = '" & stCustNo & "' And a0k04 = '" & stCustName & "' and a0k11<>'J'"
'    stSQL = " select a0w02, a0w05, a0w06 from acc0w0" & _
'            " where a0w01=" & stYear & " and a0w02 in (" & _
'            " select a1v15 from acc0k0, acc1v0" & _
'            " Where a1v02=a0k01 and a1v15 is not null " & stCon & _
'            " and a0k16 = " & stYear & " And a0k04 = '" & stCustName & "' and a0k11<>'J'" & _
'            " And substr(a0k03, 1, 6) = '" & stCustNo & "'"
'    'Add By Sindy 2015/11/25
'    stSQL = stSQL & " union " & _
'            " select a1v15 from acc1k0, acc1v0, caseprogress, staff s1" & _
'            " Where a1v02=a1k01 and a1v15 is not null " & _
'            " and a1v09 = " & stYear & " And a1k35 = '" & stCustName & "'" & _
'            " and a1v01=cp09(+) and cp13=s1.st01(+)" & _
'            " And substr(a1k28, 1, 6) = '" & stCustNo & "'" & stConA1k
'    '2015/11/25 END
'    stSQL = stSQL & ")"
    
    'Modify By Sindy 2018/6/6
    If Trim(stSalesNo) <> "" Then
      If stSalesNo = "F4101" Then
         stCon = " and (substr(R11421,1,2)='F3' or R11402='" & stSalesNo & "')"
      ElseIf stSalesNo = "F4102" Then
         stCon = " and (substr(R11421,1,2)='F2' or R11402='" & stSalesNo & "')"
      ElseIf stSalesNo = "F4103" Then
         stCon = " and (substr(R11421,1,2)='F1' or R11402='" & stSalesNo & "')"
      Else
         stCon = " and R11402='" & stSalesNo & "'"
      End If
    End If
    stSQL = "select a0w02, a0w05, a0w06 from acc0w0" & _
            " where a0w01=" & stYear & " and a0w02 in (" & _
            "select a1v15 from accrpt114, acc1v0" & _
            " Where r11401='" & strUserNum & "' and a1v02=R11409 and a1v15 is not null" & stCon & _
            " and a1v09=" & stYear & " And R11404='" & stCustName & "'" & _
            " And substr(R11403, 1, 8)='" & stCustNo & "')"
    With rsQuery
        .CursorLocation = adUseClient
        .Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        bol1st = True
        If .RecordCount > 0 Then
            Do While Not .EOF
                If intCounter > 28 Then
                   intCounter = 0
                   intPage = intPage + 1
                   Printer.NewPage
                   PrintHead stSalesName, strTableName
                End If
                If bol1st Then
                    Printer.CurrentX = 300
                    Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                    Printer.Print "已收扣單資料: "
                    bol1st = False
                End If
                
                Printer.CurrentX = 1600
                Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                Printer.Print "扣單編號"
                
                
                Printer.CurrentX = 2500
                Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                Printer.Print "" & .Fields("a0w02").Value
                
                Printer.CurrentX = 3700
                Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                Printer.Print "扣繳稅額"
                
                '扣繳稅額
                If IsNull(.Fields("a0w05").Value) = False Then
                  If .Fields("a0w05").Value = 0 Then
                      strAmount = "0"
                  Else
                      strAmount = Format(.Fields("a0w05").Value, DDollar)
                  End If
                   intLength = Printer.TextWidth(strAmount)
                   Printer.CurrentX = 5300 - intLength
                   Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                   Printer.Print strAmount
                End If
                
                '備註
                Printer.CurrentX = 5500
                Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                Printer.Print "給付總額"
                
                If IsNumeric("" & .Fields("a0w06")) Then
                    If .Fields("a0w06").Value = 0 Then
                      strAmount = "0"
                    Else
                        strAmount = Format(.Fields("a0w06").Value, DDollar)
                    End If
                     intLength = Printer.TextWidth(strAmount)
                     Printer.CurrentX = 7200 - intLength
                     Printer.CurrentY = 3500 + intCounter * 400 - intDefault
                     Printer.Print strAmount
                Else
                    Printer.Print "" & .Fields("a0w06").Value
                End If
                intCounter = intCounter + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    PrintTaxData = True
flgErr:
    Set rsQuery = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description
End Function

'*************************************************
' 列印報表
'
'*************************************************
'Modify By Sindy 2020/5/8 + strTableName As String
Public Sub FormPrint(strTableName As String)
Dim stSalesName As String
'add by nickc 2007/02/08
Dim strAmount As String
Dim intLength As Integer
Dim rsA As New ADODB.Recordset
   
   intCounter = 0
   intPage = 0
   strCustomerNo = ""
   StrStaff = ""
   adoaccrpt114.CursorLocation = adUseClient
   'Modify By Sindy 2018/3/14
   'Modify By Sindy 2020/5/8 + strTableName
'   adoaccrpt114.Open "select * from accrpt114 where r11401 = '" & strUserNum & "' order by r11417 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoaccrpt114.Open "select * from " & strTableName & " where r11401 = '" & strUserNum & "' order by r11417 asc", adoTaie, adOpenStatic, adLockReadOnly
   'adoaccrpt114.Open "select * from accrpt114 where r11401 = '" & strUserNum & "' order by r11402 asc,r11404 asc,r11403 asc,r11417 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2018/3/14 END
   Do While adoaccrpt114.EOF = False
      If StrStaff <> adoaccrpt114.Fields("r11402").Value Then
         stSalesName = adoaccrpt114.Fields("r11408").Value
         intCounter = 0
         intPage = intPage + 1
         If strCustomerNo <> "" Then
            Printer.NewPage
         End If
         PrintHead stSalesName, strTableName
         StrStaff = "" & adoaccrpt114.Fields("r11402").Value
      End If
      If strCustomerNo <> adoaccrpt114.Fields("r11404").Value Then
         Printer.CurrentX = 300
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print "收據抬頭: "
         Printer.CurrentX = 1200
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         If IsNull(adoaccrpt114.Fields("r11404").Value) = False Then
            Printer.Print adoaccrpt114.Fields("r11404").Value
         Else
            Printer.Print ""
         End If
         Printer.CurrentX = 4000
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print "聯絡電話: "
         Printer.CurrentX = 4900
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         If IsNull(adoaccrpt114.Fields("r11405").Value) = False Then
            Printer.Print adoaccrpt114.Fields("r11405").Value
         Else
            Printer.Print ""
         End If
         Printer.CurrentX = 7500
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print "聯絡人: "
         Printer.CurrentX = 8300
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         If IsNull(adoaccrpt114.Fields("r11406").Value) = False Then
            Printer.Print adoaccrpt114.Fields("r11406").Value
         Else
            Printer.Print ""
         End If
         'Add By Sindy 2014/10/15 +會計備註
         'Modify By Sindy 2017/3/10
         If IsNull(adoaccrpt114.Fields("r11422").Value) = False Then
            intCounter = intCounter + 1
            Printer.CurrentX = 300
            Printer.CurrentY = 3500 + intCounter * 400 - intDefault
            Printer.Print "會計備註: "
            Printer.CurrentX = 1200
            Printer.CurrentY = 3500 + intCounter * 400 - intDefault
            Printer.Print adoaccrpt114.Fields("r11422").Value
'         'Modify By Sindy 2016/5/11 帶出(收據抬頭email)的會計備註
'         ElseIf "" & adoaccrpt114.Fields("r11404").Value <> "" Then
'            strExc(0) = "SELECT a4217 FROM acc420 WHERE a4201='" & adoaccrpt114.Fields("r11404").Value & "' and a4217 is not null"
'            intI = 1
'            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               intCounter = intCounter + 1
'               Printer.CurrentX = 300
'               Printer.CurrentY = 3500 + intCounter * 400 - intDefault
'               Printer.Print "會計備註: "
'               Printer.CurrentX = 1200
'               Printer.CurrentY = 3500 + intCounter * 400 - intDefault
'               Printer.Print rsA.Fields("a4217").Value
'            End If
'            rsA.Close
'            Set rsA = Nothing
'         '2016/5/11 END
         End If
         '2014/10/15 END
         'Add By Sindy 2016/11/1 + 會計師資料
         strExc(0) = "SELECT * FROM acc490 WHERE a4901='" & adoaccrpt114.Fields("r11404").Value & "'" & _
                     " union " & _
                     "SELECT * FROM acc490 WHERE a4901='" & Left(Trim(adoaccrpt114.Fields("R11423").Value) & "000000", 9) & "'"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            intCounter = intCounter + 1
            Printer.CurrentX = 300
            Printer.CurrentY = 3500 + intCounter * 400 - intDefault
            Printer.Print "會計師資料:"
            Printer.CurrentX = 1300
            Printer.CurrentY = 3500 + intCounter * 400 - intDefault
            Printer.Print "" & rsA.Fields("a4912").Value & rsA.Fields("a4902").Value & _
                        IIf(Trim("" & rsA.Fields("a4903").Value) <> "", " 電話:" & rsA.Fields("a4903").Value, "") & _
                        IIf(Trim("" & rsA.Fields("a4904").Value) <> "", " 傳真:" & rsA.Fields("a4904").Value, "") & _
                        IIf(Trim("" & rsA.Fields("a4905").Value) <> "", " E-Mail:" & rsA.Fields("a4905").Value, "") & _
                        IIf(Trim("" & rsA.Fields("a4914").Value) <> "", " 備註:" & rsA.Fields("a4914").Value, "")
         End If
         rsA.Close
         Set rsA = Nothing
         '2016/11/1 END
         intCounter = intCounter + 1
         strCustomerNo = "" & adoaccrpt114.Fields("r11404").Value
         PrintTaxData Text1.Text, "" & adoaccrpt114.Fields("r11403"), strCustomerNo, stSalesName, Text2.Text
      End If
      If intCounter > 28 Then
         intCounter = 0
         intPage = intPage + 1
         Printer.NewPage
         PrintHead stSalesName, strTableName
      End If
      '收款日
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = 3500 + intCounter * 400 - intDefault
      If IsNull(adoaccrpt114.Fields("r11407").Value) = False Then
         Printer.Print IIf(Mid(CFDate(adoaccrpt114.Fields("r11407").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt114.Fields("r11407").Value), 2, 8), CFDate(adoaccrpt114.Fields("r11407").Value))
      Else
         Printer.Print ""
      End If
      '收據號碼
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = 3500 + intCounter * 400 - intDefault
      If IsNull(adoaccrpt114.Fields("r11409").Value) = False Then
         Printer.Print adoaccrpt114.Fields("r11409").Value
      Else
         Printer.Print ""
      End If
      '案件性質
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = 3500 + intCounter * 400 - intDefault
      If IsNull(adoaccrpt114.Fields("r11410").Value) = False Then
         Printer.Print StrConv(MidB(StrConv(adoaccrpt114.Fields("r11410").Value, vbFromUnicode), 1, 10), vbUnicode)
      Else
         Printer.Print ""
      End If
      '申請國家
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = 3500 + intCounter * 400 - intDefault
      If IsNull(adoaccrpt114.Fields("r11411").Value) = False Then
         Printer.Print adoaccrpt114.Fields("r11411").Value
      Else
         Printer.Print ""
      End If
      '收款金額
     If IsNull(adoaccrpt114.Fields("r11412").Value) = False Then
        If adoaccrpt114.Fields("r11412").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11412").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(4) + 900 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      '服務費
      If IsNull(adoaccrpt114.Fields("r11418").Value) = False Then
        If adoaccrpt114.Fields("r11418").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11418").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(5) + 900 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      '可扣稅額
      If IsNull(adoaccrpt114.Fields("r11419").Value) = False Then
        If adoaccrpt114.Fields("r11419").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11419").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(6) + 900 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      '收款扣繳額
      If IsNull(adoaccrpt114.Fields("r11413").Value) = False Then
        If adoaccrpt114.Fields("r11413").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11413").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(7) + 900 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      '補扣繳額
      If IsNull(adoaccrpt114.Fields("r11420").Value) = False Then
        If adoaccrpt114.Fields("r11420").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11420").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(8) + 900 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      '已收扣單金額
      If IsNull(adoaccrpt114.Fields("r11414").Value) = False Then
        If adoaccrpt114.Fields("r11414").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11414").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(9) + 1100 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      '調整稅額
      If IsNull(adoaccrpt114.Fields("r11415").Value) = False Then
        If adoaccrpt114.Fields("r11415").Value = 0 Then
            strAmount = "0"
        Else
            strAmount = Format(adoaccrpt114.Fields("r11415").Value, DDollar)
        End If
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = PLeft(10) + 900 - intLength
         Printer.CurrentY = 3500 + intCounter * 400 - intDefault
         Printer.Print strAmount
      End If
      
      If adoaccrpt114.Fields("r11411").Value = ReportSum(24) Then
         Printer.Line (300, 3500 + intCounter * 400 + 350 - intDefault)-(12000, 3500 + intCounter * 400 + 350 - intDefault)
      End If
      intCounter = intCounter + 1
      adoaccrpt114.MoveNext
   Loop
   adoaccrpt114.Close
   Printer.EndDoc
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
'  轉成Excel檔案-Main
'
'*************************************************
Private Sub ExcelSaveMain()
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
'Dim strTotalAmt1 As String
'Dim strTotalAmt2 As String
'Dim strTotalAmt3 As String
'Dim strTotalAmt4 As String
'Dim strTotalAmt5 As String
'Dim strTotalAmt6 As String
'Dim strTotalAmt7 As String 'Add By Sindy 2019/12/12
''Add by Morgan 2004/3/17
'Dim lngLstRow As Long
'Dim strS2TotalAmt1 As String
'Dim strS2TotalAmt2 As String
'Dim strS2TotalAmt3 As String
'Dim strS2TotalAmt4 As String
'Dim strS2TotalAmt5 As String
'Dim strS2TotalAmt6 As String
'Dim strS2TotalAmt7 As String 'Add By Sindy 2019/12/12
'Dim strDept As String 'Add By Sindy 2017/3/14
Dim i As Integer 'add By Sindy 2020/5/11
Dim strTableName As String, strCompName As String 'add By Sindy 2020/5/11
   
   lngCounter = 0
   
On Error GoTo flgErr
'   If Dir(strExcelPath & Mid(ReportTitle(1141), 6, 9) & strsrvdate(2) & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & Mid(ReportTitle(1141), 6, 9) & strsrvdate(2) & MsgText(43)
'   End If

   If Dir(strExcelPath & Mid(ReportTitle(1141), 6, 9) & strSrvDate(2) & ServerTime & MsgText(43)) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & Mid(ReportTitle(1141), 6, 9) & strSrvDate(2) & ServerTime & MsgText(43)
   End If
   xlsSalesPoint.SheetsInNewWorkbook = 2 '1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   
   'Modify By Sindy 2020/5/11 切二個工作表
   For i = 1 To 2
      If i = 1 Then
         strTableName = "accrpt114"
         strCompName = A0802Query("1", True)
      Else
         strTableName = "accrpt114_L"
         strCompName = A0802Query("L", True)
      End If
      '先檢查有無該公司資料
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
      adoquery.Open "SELECT count(*) FROM " & strTableName & " WHERE R11401='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 And adoquery.Fields(0) > 0 Then
         adoquery.Close
         'Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
         Set wksaccrpt114 = xlsSalesPoint.Worksheets(i)
         wksaccrpt114.Name = strCompName '工作表更名
         wksaccrpt114.Select '切換工作表
         wksaccrpt114.Columns("a:a").ColumnWidth = 13
         wksaccrpt114.Columns("b:b").ColumnWidth = 13
         wksaccrpt114.Columns("c:c").ColumnWidth = 8
         wksaccrpt114.Columns("d:d").ColumnWidth = 8
         wksaccrpt114.Columns("e:e").ColumnWidth = 8
         wksaccrpt114.Columns("f:f").ColumnWidth = 10
         wksaccrpt114.Columns("g:g").ColumnWidth = 10
         wksaccrpt114.Columns("h:h").ColumnWidth = 10
         wksaccrpt114.Columns("i:i").ColumnWidth = 10 'Add By Sindy 2019/12/11
         wksaccrpt114.Range("a1").Value = ReportTitle(1141)
         wksaccrpt114.Range("a1:i1").Select
         With wksaccrpt114.Range("a1:i1")
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlBottom
             .WrapText = False
             .Orientation = 0
             .AddIndent = False
             .ShrinkToFit = False
             .MergeCells = True
         End With
         wksaccrpt114.Range("a4").Value = ReportSum(76)
         wksaccrpt114.Range("b4").Value = Text1
         'Modify by Morgan 2004/4/29
         '"已扣金額"改為"催收稅額"
         'wksaccrpt114.Range("b6").Value = ReportSum(77)
         wksaccrpt114.Range("b6").Value = "催收稅額"
         wksaccrpt114.Range("b6").HorizontalAlignment = xlCenter
         wksaccrpt114.Range("c6").Value = "已收扣單"
         wksaccrpt114.Range("c6").HorizontalAlignment = xlCenter
         wksaccrpt114.Range("d6").Value = "已收現金"
         wksaccrpt114.Range("d6").HorizontalAlignment = xlCenter
         wksaccrpt114.Range("e6").Value = "列呆帳"
         wksaccrpt114.Range("e6").HorizontalAlignment = xlCenter
         wksaccrpt114.Range("f6").Value = "催收中"
         wksaccrpt114.Range("f6").HorizontalAlignment = xlCenter
         wksaccrpt114.Range("g6").Value = "轉列下年度"
         wksaccrpt114.Range("g6").HorizontalAlignment = xlCenter
         wksaccrpt114.Range("h6").Value = "合計"
         wksaccrpt114.Range("h6").HorizontalAlignment = xlCenter
         'Add By Sindy 2019/12/11
         wksaccrpt114.Range("i6").Value = "客戶家數"
         wksaccrpt114.Range("i6").HorizontalAlignment = xlCenter
         '2019/12/11 END
         
         Call ExcelSave(xlsSalesPoint, wksaccrpt114, strTableName) '明細
      End If
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
   Next i
   
    'Modify By Cheng 2003/06/09
'   xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(1141), 6, 9) & strsrvdate(2) & MsgText(43)
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(1141), 6, 9) & strSrvDate(2) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(1141), 6, 9) & strSrvDate(2) & ServerTime & MsgText(43), FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set xlsSalesPoint = Nothing
   StatusClear
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave(xlsSalesPoint As Excel.Application, wksaccrpt114 As Worksheet, strTableName As String)
Dim strTotalAmt1 As String
Dim strTotalAmt2 As String
Dim strTotalAmt3 As String
Dim strTotalAmt4 As String
Dim strTotalAmt5 As String
Dim strTotalAmt6 As String
Dim strTotalAmt7 As String 'Add By Sindy 2019/12/12
'Add by Morgan 2004/3/17
Dim lngLstRow As Long
Dim strS2TotalAmt1 As String
Dim strS2TotalAmt2 As String
Dim strS2TotalAmt3 As String
Dim strS2TotalAmt4 As String
Dim strS2TotalAmt5 As String
Dim strS2TotalAmt6 As String
Dim strS2TotalAmt7 As String 'Add By Sindy 2019/12/12
Dim strDept As String 'Add By Sindy 2017/3/14
   
On Error GoTo flgErr

   lngCounter = 7
   
' 智權人員
   adostaff.CursorLocation = adUseClient
   adostaff.Open "select * from acc090 where substr(a0901, 1, 1) = 'S' and a0901 <> 'S20' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adostaff.EOF = False
      adoaccrpt114.CursorLocation = adUseClient
      'Modify by Morgan 2004/3/17
      '改以 ST15 排序
      'adoaccrpt114.Open "select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st03 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      'adoaccrpt114.Open "select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoaccrpt114.Open "select distinct r11402 from " & strTableName & ", staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoaccrpt114.EOF = False
         If IsNull(adoaccrpt114.Fields(0).Value) Then
            wksaccrpt114.Range("a" & lngCounter).Value = PUB_ChkExcelZero(1)
         Else
            wksaccrpt114.Range("a" & lngCounter).Value = StaffQuery(adoaccrpt114.Fields(0).Value)
         End If
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select sum(a1u06) from acc1u0, acc0k0 where a1u02 = a0k01 and a0k16 = " & Val(Text1) & " and a0k20 = '" & adoaccrpt114.Fields(0).Value & "'"
         'Modify by Morgan 2004/3/5
         'adoquery.Open "select sum(r11413) from accrpt114  WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select sum(r11413+r11420) from " & strTableName & " WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
            'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之 MsgText(601)=> PUB_ChkExcelZero(1)
               wksaccrpt114.Range("b" & lngCounter).Value = PUB_ChkExcelZero(1)
            Else
               'Modify By Sindy 2015/6/30 取整數+Fix
               wksaccrpt114.Range("b" & lngCounter).Value = Fix(Val(adoquery.Fields(0).Value))
            End If
         Else
            wksaccrpt114.Range("b" & lngCounter).Value = PUB_ChkExcelZero(1)
         End If
         adoquery.Close
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(r11414) from " & strTableName & " WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               wksaccrpt114.Range("c" & lngCounter).Value = PUB_ChkExcelZero(1)
            Else
               wksaccrpt114.Range("c" & lngCounter).Value = Val(adoquery.Fields(0).Value)
            End If
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = PUB_ChkExcelZero(1)
         End If
         adoquery.Close
         wksaccrpt114.Range("d" & lngCounter).Value = 0
         wksaccrpt114.Range("e" & lngCounter).Value = 0
         'Modify by Morgan 2004/3/5
         'wksaccrpt114.Range("f" & lngCounter).Formula = "=(b" & lngCounter & "-c" & lngCounter & ")"
         wksaccrpt114.Range("f" & lngCounter).Formula = "=(b" & lngCounter & "-c" & lngCounter & "-d" & lngCounter & "-e" & lngCounter & "-g" & lngCounter & ")"
         wksaccrpt114.Range("g" & lngCounter).Value = 0
         wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
         
         'Add By Sindy 2019/12/12
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select count(distinct r11404) from " & strTableName & " WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               wksaccrpt114.Range("i" & lngCounter).Value = PUB_ChkExcelZero(1)
            Else
               wksaccrpt114.Range("i" & lngCounter).Value = Val(adoquery.Fields(0).Value)
            End If
         Else
            wksaccrpt114.Range("i" & lngCounter).Value = PUB_ChkExcelZero(1)
         End If
         adoquery.Close
         '2019/12/12 END
         
         lngCounter = lngCounter + 1
         strDept = adostaff.Fields("a0901").Value 'Add By Sindy 2017/3/14
         adoaccrpt114.MoveNext
      Loop
      adoaccsum.CursorLocation = adUseClient
      'Modify by Morgan 2004/3/17
      '改以 ST15 排序
      'adoaccsum.Open "select count(r11402) from (select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st03 = '" & adostaff.Fields("a0901").Value & "') new", adoTaie, adOpenStatic, adLockReadOnly
      'adoaccsum.Open "select count(r11402) from (select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "') new", adoTaie, adOpenStatic, adLockReadOnly
      adoaccsum.Open "select count(r11402) from (select distinct r11402 from " & strTableName & ", staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and r11421 = '" & adostaff.Fields("a0901").Value & "') new", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
            wksaccrpt114.Range("a" & lngCounter).Value = adostaff.Fields("a0902").Value & ReportSum(25)
            wksaccrpt114.Range("b" & lngCounter).Formula = "=sum(b" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":b" & (lngCounter - 1) & ")"
            wksaccrpt114.Range("c" & lngCounter).Formula = "=sum(c" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":c" & (lngCounter - 1) & ")"
            wksaccrpt114.Range("d" & lngCounter).Formula = "=sum(d" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":d" & (lngCounter - 1) & ")"
            wksaccrpt114.Range("e" & lngCounter).Formula = "=sum(e" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":e" & (lngCounter - 1) & ")"
            wksaccrpt114.Range("f" & lngCounter).Formula = "=sum(f" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":f" & (lngCounter - 1) & ")"
            wksaccrpt114.Range("g" & lngCounter).Formula = "=sum(g" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":g" & (lngCounter - 1) & ")"
            wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
            wksaccrpt114.Range("i" & lngCounter).Formula = "=sum(i" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":i" & (lngCounter - 1) & ")" 'Add By Sindy 2019/12/12
            strTotalAmt1 = strTotalAmt1 & "b" & lngCounter & ", "
            strTotalAmt2 = strTotalAmt2 & "c" & lngCounter & ", "
            strTotalAmt3 = strTotalAmt3 & "d" & lngCounter & ", "
            strTotalAmt4 = strTotalAmt4 & "e" & lngCounter & ", "
            strTotalAmt5 = strTotalAmt5 & "f" & lngCounter & ", "
            strTotalAmt6 = strTotalAmt6 & "g" & lngCounter & ", "
            strTotalAmt7 = strTotalAmt7 & "i" & lngCounter & ", " 'Add By Sindy 2019/12/12
            'Add by Morgan 2004/3/17
            '加台中所合計
            If Left(adostaff.Fields("a0901").Value, 2) = "S2" Then
                strS2TotalAmt1 = strS2TotalAmt1 & "b" & lngCounter & ", "
                strS2TotalAmt2 = strS2TotalAmt2 & "c" & lngCounter & ", "
                strS2TotalAmt3 = strS2TotalAmt3 & "d" & lngCounter & ", "
                strS2TotalAmt4 = strS2TotalAmt4 & "e" & lngCounter & ", "
                strS2TotalAmt5 = strS2TotalAmt5 & "f" & lngCounter & ", "
                strS2TotalAmt6 = strS2TotalAmt6 & "g" & lngCounter & ", "
                strS2TotalAmt7 = strS2TotalAmt7 & "i" & lngCounter & ", " 'Add By Sindy 2019/12/12
            End If
            If adostaff.Fields("a0901").Value = "S15" Then
            'If Left(strDept, 2) = "S1" And Left(strDept, 2) <> Left(adostaff.Fields("a0901").Value, 2) Then 'Modify By Sindy 2017/3/14
               lngCounter = lngCounter + 1
               wksaccrpt114.Range("a" & lngCounter).Value = ReportSum(105) '台北所合計 'ReportSum(63) & ReportSum(25)
               wksaccrpt114.Range("b" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt1, 1, Len(strTotalAmt1) - 2) & ")"
               wksaccrpt114.Range("c" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt2, 1, Len(strTotalAmt2) - 2) & ")"
               wksaccrpt114.Range("d" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt3, 1, Len(strTotalAmt3) - 2) & ")"
               wksaccrpt114.Range("e" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt4, 1, Len(strTotalAmt4) - 2) & ")"
               wksaccrpt114.Range("f" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt5, 1, Len(strTotalAmt5) - 2) & ")"
               wksaccrpt114.Range("g" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt6, 1, Len(strTotalAmt6) - 2) & ")"
               wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
               wksaccrpt114.Range("i" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt7, 1, Len(strTotalAmt7) - 2) & ")" 'Add By Sindy 2019/12/12
            'Add by Morgan 2004/3/17
            '加台中所合計
            'Modify By Sindy 2017/3/14
            'ElseIf adostaff.Fields("a0901").Value = "S23" Then
            'ElseIf Left(strDept, 2) = "S2" And Left(strDept, 2) <> Left(adostaff.Fields("a0901").Value, 2) Then
            ElseIf adostaff.Fields("a0901").Value = "S24" Then
            '2017/3/14 END
               lngCounter = lngCounter + 1
               wksaccrpt114.Range("a" & lngCounter).Value = ReportSum(106) '台中所合計
               wksaccrpt114.Range("b" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt1, 1, Len(strS2TotalAmt1) - 2) & ")"
               wksaccrpt114.Range("c" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt2, 1, Len(strS2TotalAmt2) - 2) & ")"
               wksaccrpt114.Range("d" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt3, 1, Len(strS2TotalAmt3) - 2) & ")"
               wksaccrpt114.Range("e" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt4, 1, Len(strS2TotalAmt4) - 2) & ")"
               wksaccrpt114.Range("f" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt5, 1, Len(strS2TotalAmt5) - 2) & ")"
               wksaccrpt114.Range("g" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt6, 1, Len(strS2TotalAmt6) - 2) & ")"
               wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
               wksaccrpt114.Range("i" & lngCounter).Formula = "=sum(" & Mid(strS2TotalAmt7, 1, Len(strS2TotalAmt7) - 2) & ")" 'Add By Sindy 2019/12/12
            End If
            lngCounter = lngCounter + 2
         End If
      End If
      adoaccsum.Close
      adoaccrpt114.Close
      adostaff.MoveNext
   Loop
   adostaff.Close
   
' 其他人員
   adostaff.CursorLocation = adUseClient
   'adostaff.Open "select * from acc090 where (substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = 'S20' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
   adostaff.Open "select * from acc090 where substr(a0901, 1, 1) <> 'S' or a0901 = 'S20' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
   lngLstRow = lngCounter
   Do While adostaff.EOF = False
      adoaccrpt114.CursorLocation = adUseClient
      'Modify by Morgan 2004/3/17
      'adoaccrpt114.Open "select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st03 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      'adoaccrpt114.Open "select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoaccrpt114.Open "select distinct r11402 from " & strTableName & ", staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoaccrpt114.EOF = False
         If IsNull(adoaccrpt114.Fields(0).Value) Then
            wksaccrpt114.Range("a" & lngCounter).Value = PUB_ChkExcelZero(1)
         Else
            wksaccrpt114.Range("a" & lngCounter).Value = StaffQuery(adoaccrpt114.Fields(0).Value)
         End If

         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select sum(a1u06) from acc1u0, acc0k0 where a1u02 = a0k01 and a0k16 = " & Val(Text1) & " and a0k20 = '" & adoaccrpt114.Fields(0).Value & "'"
         'Modify by Morgan 2004/3/5
         'adoquery.Open "select sum(r11413) from accrpt114  WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select sum(r11413+r11420) from " & strTableName & " WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               wksaccrpt114.Range("b" & lngCounter).Value = PUB_ChkExcelZero(1)
            Else
               'Modify By Sindy 2015/6/30 取整數+Fix
               wksaccrpt114.Range("b" & lngCounter).Value = Fix(Val(adoquery.Fields(0).Value))
            End If
         Else
            wksaccrpt114.Range("b" & lngCounter).Value = PUB_ChkExcelZero(1)
         End If
         adoquery.Close
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(r11414) from " & strTableName & " WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               wksaccrpt114.Range("c" & lngCounter).Value = PUB_ChkExcelZero(1)
            Else
               wksaccrpt114.Range("c" & lngCounter).Value = Val(adoquery.Fields(0).Value)
            End If
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = PUB_ChkExcelZero(1)
         End If
         wksaccrpt114.Range("d" & lngCounter).Value = 0
         wksaccrpt114.Range("e" & lngCounter).Value = 0
         'Modify by Morgan 2004/3/5
         'wksaccrpt114.Range("f" & lngCounter).Formula = "=(b" & lngCounter & "-c" & lngCounter & ")"
         wksaccrpt114.Range("f" & lngCounter).Formula = "=(b" & lngCounter & "-c" & lngCounter & "-d" & lngCounter & "-e" & lngCounter & "-g" & lngCounter & ")"
         wksaccrpt114.Range("g" & lngCounter).Value = 0
         wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
         adoquery.Close
         
         'Add By Sindy 2019/12/12
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select count(distinct r11404) from " & strTableName & " WHERE R11401='" & strUserNum & "' AND r11402 = '" & adoaccrpt114.Fields("r11402").Value & "' and r11421 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               wksaccrpt114.Range("i" & lngCounter).Value = PUB_ChkExcelZero(1)
            Else
               wksaccrpt114.Range("i" & lngCounter).Value = Val(adoquery.Fields(0).Value)
            End If
         Else
            wksaccrpt114.Range("i" & lngCounter).Value = PUB_ChkExcelZero(1)
         End If
         adoquery.Close
         '2019/12/12 END
         
         lngCounter = lngCounter + 1
         adoaccrpt114.MoveNext
      Loop
      'Modify by Morgan 2004/3/17
      '取消非SXX的個別小計改合計
'      adoaccsum.CursorLocation = adUseClient
'      'adoaccsum.Open "select count(r11402) from (select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st03 = '" & adostaff.Fields("a0901").Value & "') new", adoTaie, adOpenStatic, adLockReadOnly
'      adoaccsum.Open "select count(r11402) from (select distinct r11402 from accrpt114, staff where substr(r11402, 1, 5) = st01 and r11401 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "') new", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'            wksaccrpt114.Range("a" & lngCounter).Value = ReportSum(64) & ReportSum(25)
'            wksaccrpt114.Range("b" & lngCounter).Formula = "=sum(b" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":b" & (lngCounter - 1) & ")"
'            wksaccrpt114.Range("c" & lngCounter).Formula = "=sum(c" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":c" & (lngCounter - 1) & ")"
'            wksaccrpt114.Range("d" & lngCounter).Formula = "=sum(d" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":d" & (lngCounter - 1) & ")"
'            wksaccrpt114.Range("e" & lngCounter).Formula = "=sum(e" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":e" & (lngCounter - 1) & ")"
'            wksaccrpt114.Range("f" & lngCounter).Formula = "=sum(f" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":f" & (lngCounter - 1) & ")"
'            wksaccrpt114.Range("g" & lngCounter).Formula = "=sum(g" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":g" & (lngCounter - 1) & ")"
'            wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
'            strTotalAmt1 = strTotalAmt1 & "b" & lngCounter & ", "
'            strTotalAmt2 = strTotalAmt2 & "c" & lngCounter & ", "
'            strTotalAmt3 = strTotalAmt3 & "d" & lngCounter & ", "
'            strTotalAmt4 = strTotalAmt4 & "e" & lngCounter & ", "
'            strTotalAmt5 = strTotalAmt5 & "f" & lngCounter & ", "
'            strTotalAmt6 = strTotalAmt6 & "g" & lngCounter & ", "
'            lngCounter = lngCounter + 2
'         End If
'      End If
'      adoaccsum.Close
      adoaccrpt114.Close
      adostaff.MoveNext
   Loop
   adostaff.Close
   
    'Add by Morgan 2004/3/17
    '非SXX的合計
    If lngLstRow <> lngCounter Then
        wksaccrpt114.Range("a" & lngCounter).Value = ReportSum(25)
        wksaccrpt114.Range("b" & lngCounter).Formula = "=sum(b" & (lngLstRow) & ":b" & (lngCounter - 1) & ")"
        wksaccrpt114.Range("c" & lngCounter).Formula = "=sum(c" & (lngLstRow) & ":c" & (lngCounter - 1) & ")"
        wksaccrpt114.Range("d" & lngCounter).Formula = "=sum(d" & (lngLstRow) & ":d" & (lngCounter - 1) & ")"
        wksaccrpt114.Range("e" & lngCounter).Formula = "=sum(e" & (lngLstRow) & ":e" & (lngCounter - 1) & ")"
        wksaccrpt114.Range("f" & lngCounter).Formula = "=sum(f" & (lngLstRow) & ":f" & (lngCounter - 1) & ")"
        wksaccrpt114.Range("g" & lngCounter).Formula = "=sum(g" & (lngLstRow) & ":g" & (lngCounter - 1) & ")"
        wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
        wksaccrpt114.Range("i" & lngCounter).Formula = "=sum(i" & (lngLstRow) & ":i" & (lngCounter - 1) & ")" 'Add By Sindy 2019/12/12
        strTotalAmt1 = strTotalAmt1 & "b" & lngCounter & ", "
        strTotalAmt2 = strTotalAmt2 & "c" & lngCounter & ", "
        strTotalAmt3 = strTotalAmt3 & "d" & lngCounter & ", "
        strTotalAmt4 = strTotalAmt4 & "e" & lngCounter & ", "
        strTotalAmt5 = strTotalAmt5 & "f" & lngCounter & ", "
        strTotalAmt6 = strTotalAmt6 & "g" & lngCounter & ", "
        strTotalAmt7 = strTotalAmt7 & "i" & lngCounter & ", " 'Add By Sindy 2019/12/12
        lngCounter = lngCounter + 2
    End If
    'Add end
   
' 總所合計
   wksaccrpt114.Range("a" & lngCounter).Value = ReportSum(66) & ReportSum(25)
   wksaccrpt114.Range("b" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt1, 1, Len(strTotalAmt1) - 2) & ")"
   wksaccrpt114.Range("c" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt2, 1, Len(strTotalAmt2) - 2) & ")"
   wksaccrpt114.Range("d" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt3, 1, Len(strTotalAmt3) - 2) & ")"
   wksaccrpt114.Range("e" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt4, 1, Len(strTotalAmt4) - 2) & ")"
   wksaccrpt114.Range("f" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt5, 1, Len(strTotalAmt5) - 2) & ")"
   wksaccrpt114.Range("g" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt6, 1, Len(strTotalAmt6) - 2) & ")"
   wksaccrpt114.Range("h" & lngCounter).Formula = "=sum(c" & lngCounter & ":g" & lngCounter & ")"
   wksaccrpt114.Range("i" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt7, 1, Len(strTotalAmt7) - 2) & ")" 'Add By Sindy 2019/12/12
   
   Exit Sub
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Modify By Sindy 2025/3/21 訊息不明確
'   If Text1 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
   If Text1 = "" Then
      MsgBox "扣繳年度不可空白!!", vbExclamation
      Text1.SetFocus
      Exit Function
   End If
   'FormCheck = False
   FormCheck = True
   '2025/3/21 END
End Function
