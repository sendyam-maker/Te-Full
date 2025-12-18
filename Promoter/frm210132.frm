VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210132 
   BorderStyle     =   1  '單線固定
   Caption         =   "未列印收據查詢"
   ClientHeight    =   5720
   ClientLeft      =   3780
   ClientTop       =   3710
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   4
      Left            =   7140
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   3
      Left            =   6030
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   60
      Width           =   1080
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋(&Q)"
      Height          =   315
      Left            =   4620
      TabIndex        =   25
      Top             =   1170
      Width           =   765
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2670
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1515
      MaxLength       =   6
      TabIndex        =   2
      Top             =   470
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   5520
      TabIndex        =   22
      Top             =   570
      Width           =   3315
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   180
         Width           =   2550
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   23
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   264
      Left            =   7140
      TabIndex        =   18
      Top             =   5460
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   264
      Left            =   4770
      TabIndex        =   17
      Top             =   5460
      Width           =   1545
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   264
      Left            =   2130
      TabIndex        =   16
      Top             =   5460
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印通知(&S)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   7320
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   1410
      Width           =   1515
   End
   Begin VB.CheckBox Check2 
      Caption         =   "已送件"
      Height          =   255
      Left            =   780
      TabIndex        =   7
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印清單(&P)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   1410
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3435
      Left            =   30
      TabIndex        =   15
      Top             =   1980
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   6050
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   18
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1515
      MaxLength       =   9
      TabIndex        =   4
      Top             =   820
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   3135
      MaxLength       =   9
      TabIndex        =   5
      Top             =   820
      Width           =   1365
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5250
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   60
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   60
      Width           =   750
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   336
      Left            =   1500
      TabIndex        =   3
      Top             =   444
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;593"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   1515
      TabIndex        =   6
      Top             =   1170
      Width           =   3075
      VariousPropertyBits=   671105051
      Size            =   "5424;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2490
      TabIndex        =   31
      Top             =   510
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "已勾選列印收據　　　　張"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2640
      TabIndex        =   30
      Top             =   1740
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "申請人中文名稱："
      Height          =   180
      Index           =   19
      Left            =   45
      TabIndex        =   29
      Top             =   1230
      Width           =   1545
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Left            =   600
      TabIndex        =   28
      Top             =   880
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "業  務  區："
      Height          =   180
      Left            =   600
      TabIndex        =   27
      Top             =   180
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   600
      TabIndex        =   26
      Top             =   530
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "已通知列印收據　　　　張"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   2640
      TabIndex        =   24
      Top             =   1500
      Width           =   2160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2490
      X2              =   2610
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "規費："
      Height          =   180
      Left            =   4200
      TabIndex        =   21
      Top             =   5490
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "服務費："
      Height          =   180
      Index           =   0
      Left            =   1380
      TabIndex        =   20
      Top             =   5490
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "合計"
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   19
      Top             =   5490
      Width           =   360
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2940
      X2              =   3060
      Y1              =   990
      Y2              =   990
   End
End
Attribute VB_Name = "frm210132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原”未列印收據/請款單查詢”標題修改為”未列印收據查詢”
'原”客戶代號”改為”客戶編號”
'原”客戶名稱”改為”申請人名稱”
'原”客戶中文名稱”改為”申請人中文名稱”
'原”收據編號”改為”收據號碼”
'end 2021/07/27
'Memo by Lydia 2021/07/13 改成Form2.0 ; lblSalesName、Text6、GrdDataList改字型=新細明體-ExtB; Printer列印未改
'Memo By Eric  2013/12/31 增加顯示 收據公司別
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Public cmdState As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim stST15 As String, stST05 As String
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim m_dblPaperCnt As Double, m_dblAmt As Double 'Add By Sindy 2012/9/21
'Add by Amy 2014/05/21
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1
Dim strPrinter As String 'Added by Morgan 2020/10/30
'Add By Sindy 2023/6/12
Dim arrID
Dim bolAreaMan As Boolean '下拉選單有區主管
'2023/6/12 END


Private Sub SetDataListWidth()
   Dim iCol As Integer
   Dim i As Integer
   
   grdDataList.row = 0
   grdDataList.ColAlignment = flexAlignLeftCenter
   
   'Modified by Morgan 2011/12/23 調整欄位順序--辜
   iCol = 0
   grdDataList.col = iCol: grdDataList.Text = "V"
   grdDataList.ColWidth(iCol) = 200
   
   iCol = iCol + 1 '1
   grdDataList.col = iCol: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(iCol) = 750
   
   iCol = iCol + 1 '2
   grdDataList.col = iCol: grdDataList.Text = "介紹人"
   grdDataList.ColWidth(iCol) = 750
   
   iCol = iCol + 1 '3
   grdDataList.col = iCol: grdDataList.Text = "發文日"
   grdDataList.ColWidth(iCol) = 800
   
   iCol = iCol + 1 '4
   grdDataList.col = iCol: grdDataList.Text = "本所案號"
   grdDataList.ColWidth(iCol) = 1450
   
   iCol = iCol + 1 '5
   'Add By Sindy 2011/1/28
   grdDataList.col = iCol: grdDataList.Text = "案件名稱"
   grdDataList.ColWidth(iCol) = 1500
   
   iCol = iCol + 1 '6
   'Modified by Lydia 2021/07/27 申請人=>申請人名稱
   grdDataList.col = iCol: grdDataList.Text = "申請人名稱"
   grdDataList.ColWidth(iCol) = 800
   '2011/1/28 End
   
   iCol = iCol + 1 '7
   grdDataList.col = iCol: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(iCol) = 800
   
   '2013/12/31 start add by eric
   iCol = iCol + 1 '8
   grdDataList.col = iCol: grdDataList.Text = "公司"
   grdDataList.ColWidth(iCol) = 500
   '2013/12/31 end
   
   iCol = iCol + 1 '9
   grdDataList.col = iCol: grdDataList.Text = "案件性質"
   grdDataList.ColWidth(iCol) = 800
   
   iCol = iCol + 1 '10
   grdDataList.col = iCol: grdDataList.Text = "總收文號"
   grdDataList.ColWidth(iCol) = 0
   
   iCol = iCol + 1 '11
   grdDataList.col = iCol: grdDataList.Text = "服務費"
   grdDataList.ColWidth(iCol) = 750
   grdDataList.ColAlignment(iCol) = flexAlignRightCenter

   iCol = iCol + 1 '12
   grdDataList.col = iCol: grdDataList.Text = "規費"
   grdDataList.ColWidth(iCol) = 750
   grdDataList.ColAlignment(iCol) = flexAlignRightCenter
   
   iCol = iCol + 1 '13
   grdDataList.col = iCol: grdDataList.Text = "可列印"
   grdDataList.ColWidth(iCol) = 500
   
   iCol = iCol + 1 '14
   grdDataList.col = iCol: grdDataList.Text = "收據號碼"
   grdDataList.ColWidth(iCol) = 1000
   
   iCol = iCol + 1 '15
   grdDataList.col = iCol: grdDataList.Text = "收據抬頭"
   grdDataList.ColWidth(iCol) = 1200

   iCol = iCol + 1 '16
   grdDataList.col = iCol: grdDataList.Text = "收文日"
   grdDataList.ColWidth(iCol) = 800
   
   iCol = iCol + 1 '17
   grdDataList.col = iCol: grdDataList.Text = "已送件規費"
   grdDataList.ColWidth(iCol) = 0
   
   'Add By Sindy 2020/12/15
   iCol = iCol + 1 '18
   For i = iCol To grdDataList.Cols - 1
      grdDataList.ColWidth(i) = 0
   Next i
   '2020/12/15 END
End Sub

Private Sub Data_Process()
Dim i As Integer

On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   For i = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(i, 0) = "V" Then
         'Modified by Morgan 2011/12/23 調整欄位順序--辜
         'strExc(0) = "update acc0k0 set a0k02=" & strSrvDate(2) & ",a0k32='Y' where a0k01='" & grdDataList.TextMatrix(i, 9) & "'"
         'MODIFY BY SONIA 2014/1/6 加公司別, 7以後順延
         '2015/1/30 modify by sonia 加a0k32='N'條件,怕業務重覆按且財務處剛好在印收據又會更新為'Y'
         'strExc(0) = "update acc0k0 set a0k02=" & strSrvDate(2) & ",a0k32='Y' where a0k01='" & grdDataList.TextMatrix(i, 13) & "'"
         strExc(0) = "update acc0k0 set a0k02=" & strSrvDate(2) & ",a0k32='Y' where a0k01='" & grdDataList.TextMatrix(i, 14) & "' and a0k32='N'"
         cnnConnection.Execute strExc(0)
      End If
   Next i
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   MsgBox "執行完畢!!"
   Call Search_Process
   Exit Sub
ErrHand:
    cnnConnection.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox (Err.Description)
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim blChkData As Boolean
Dim i As Integer, j As Integer
Dim StrTag As String
   
   Select Case cmdState
      Case 0 '列印
         'Modified by Morgan 2011/12/23 調整欄位順序--辜
         'If grdDataList.Rows > 1 And grdDataList.TextMatrix(1, 2) <> "" Then
         If grdDataList.Rows > 1 And grdDataList.TextMatrix(1, 4) <> "" Then
            Call PrintData
         Else
            MsgBox "無資料!!"
            Exit Sub
         End If
         
      Case 1 '結束
         'Modified by Lydia 2019/07/02 從個人常用區進入後,無法結束1
         'Unload frm210132
         'Set frm210132 = Nothing
         Unload Me
         
      Case 2 '開放收據列印
         'Modified by Morgan 2011/12/23 調整欄位順序--辜
         'If grdDataList.Rows > 1 And grdDataList.TextMatrix(1, 2) <> "" Then
         If grdDataList.Rows > 1 And grdDataList.TextMatrix(1, 4) <> "" Then
            blChkData = False
            For i = 1 To grdDataList.Rows - 1
               If grdDataList.TextMatrix(i, 0) = "V" Then
                  blChkData = True
                  Exit For
               End If
            Next i
            If blChkData = True Then
               Call Data_Process
            Else
               MsgBox "無選取資料!!"
               Exit Sub
            End If
         Else
            MsgBox "無資料!!"
            Exit Sub
         End If
         
      'Modify By Sindy 2014/8/4 因為此作業Account和Promoter等都有呼叫到,以防在Account需要加入一堆Form,因此抽出來至Func
      Case 3, 4 '3.案件基本資料 4.案件進度
         Call frm210132_SubPubShowNextData(cmdState, Me)
      Case Else
   End Select
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   cmdOK(3).Enabled = False
   cmdOK(4).Enabled = False
   cmdOK(0).Enabled = False
   cmdOK(2).Enabled = False
   If ConstrainCheck = True Then
      Call Search_Process
      SetDataListWidth
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFind_Click()
   If Me.Text6.Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Sub
   End If
   frm090801_1.m_strCustChnName = Me.Text6.Text
   frm090801_1.lblName.Caption = Me.Text6.Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   If m_blnOneRec = True And m_strCustCode <> "" Then
      Me.Text1.Text = m_strCustCode
      'Modified by Lydia 2015/06/23 代號的迄改為尾碼Z
      'Me.Text2.Text = m_strCustCode
      Me.Text2.Text = IIf(Right(m_strCustCode, 3) = "000", Mid(m_strCustCode, 1, 6) & "ZZZ", IIf(Right(m_strCustCode, 1) = "0", Mid(m_strCustCode, 1, 8) & "Z", m_strCustCode))
      
      Me.Text6.Text = GetCustomerName(m_strCustCode)
      'Call Text1_Validate(False)
   End If
   'Added by Lydia 2015/06/25 輸入申請人名稱按Enter自動執行搜尋後,執行查詢
   If Me.Text1.Text <> "" And Me.Text2.Text <> "" Then
      Call cmdSearch_Click
   End If
   
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
   
   bolToEndByNick = False
   MoveFormToCenter Me
   
   'Add By Sindy 2014/8/4 Account系統不可使用進度及基本檔,因為牽扯很多Form,會導致需要加入一堆Form
   If UCase(App.Title) = "ACCOUNT" Then
      cmdOK(3).Visible = False
      cmdOK(4).Visible = False
   Else
      cmdOK(3).Visible = True
      cmdOK(4).Visible = True
   End If
   
   'Modified by Morgan 2020/10/30
   'strSql = Printer.DeviceName
   'SeekPrintL = Printer.Orientation
   'For i = 0 To Printers.Count - 1
   '   Set Printer = Printers(i)
   '   Combo1.AddItem Printer.DeviceName, j
   '   j = j + 1
   '   If Printer.DeviceName = strSql Then
   '      SeekPrint = i
   '   End If
   'Next i
   'Set Printer = Printers(SeekPrint)
   'Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   'end 2020/10/30
   
   '僅財務處及電腦中心人員可以列印
   If GetStaffDepartment(strUserNum) = "M31" Or GetStaffDepartment(strUserNum) = "M51" Then
      Frame1.Visible = True
      cmdOK(0).Visible = True
   Else
      Frame1.Visible = False
      cmdOK(0).Visible = False
   End If
   
   SetDataListWidth
   cmdState = -1
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   bolAreaMan = False 'Add By Sindy 2023/6/12
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   'Add By Sindy 2023/6/12
   '檢查當時是否需要為他人職代
   Combo3.Clear
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
         bolAreaMan = True
      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Added by Lydia 2021/05/20 Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2023/6/12 END
   
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'
'   Select Case strUserNum
'      '外商陳經理可看全所CFT,FCT,S,CFC
'      Case "68005"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
''cancel by sonia 2014/6/9
''      '蔣律師可看中所全部
''      Case "79037"
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSales.Enabled = True
''         txtSalesArea = "S2"
''         txtSalesArea1 = "S29"
''end 2014/6/9
'      '小真,杜副總可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      'Modify by Amy 2015/02/04 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
'      Case "65001", "68006"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         '副總預設所有智權人員
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
'      '王協理可看專利處
'      Case "71011"
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      'add by sonia 2016/12/21 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
'         txtSales = strUserNum
'      'end 2016/12/21
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'            '各區主管
'            Case "SM"
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               '原羅文旭72009可兼看中一區,94/7/1只可看S22
'               '2005/7/5林永生71003可看中所全部,但預設S23
'               If strUserNum = "71003" Then
'                  txtSalesArea = "S23"
'                  txtSalesArea1 = "S23"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'               '簡協理可看北所全部但預設S15
'               If strUserNum = "69005" Then
'                  txtSalesArea = "S15"
'                  txtSalesArea1 = "S15"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               End If
'               txtSales.Enabled = True
'            '加入外商主管  王宗珮、洪琬姿、葉易雲
'            Case "21", "26", "28"
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               txtSales.Enabled = True
'            '其他只能看自己
'            Case Else
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               'Added by Lydia 2017/07/25 多使用者權限,則增加部門範圍
'               strExc(1) = PUB_GetSalesList(strUserNum, , , , , strExc(2), strExc(3))
'               If strExc(3) <> "" And strExc(3) > txtSalesArea1 Then
'                  txtSalesArea1 = strExc(3)
'               End If
'               'end 2017/07/25
'         End Select
'   End Select
'
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'
'   'Add by Amy 2015/02/04 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify  by Amy 2014/05/21 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/02/04 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
'            txtSalesArea.Enabled = True: txtSalesArea = ""
'            txtSalesArea1.Enabled = True: txtSalesArea1 = ""
'            txtSales.Enabled = True
'        End If
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True: txtSales = ""
'   Else
'        txtSales = strUserNum
'   End If
'   'end 2014/05/21
'
'   'Add By Sindy 2020/7/1 記錄原操作人可以查詢的業務區及所別
'   'txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2020/7/1 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MenuEnabled
   Set frm210132 = Nothing
End Sub

Private Sub GrdDataList_Click()
Dim i As Integer
   
   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
      If grdDataList.Text = "V" Then
         grdDataList.Text = ""
         For i = 0 To grdDataList.Cols - 1
            If i <> 1 Then
               grdDataList.col = i
               grdDataList.CellBackColor = QBColor(15)
            End If
         Next i
         Call CountData 'Add By Sindy 2012/9/21
      Else
         grdDataList.Text = "V"
         For i = 0 To grdDataList.Cols - 1
            If i <> 1 Then
               grdDataList.col = i
               grdDataList.CellBackColor = &HFFC0C0
            End If
         Next i
         Call CountData 'Add By Sindy 2012/9/21
      End If
   End If
   grdDataList.Visible = True
End Sub

'Add By Sindy 2012/9/21 計算已勾選收據張數及金額
Private Sub CountData()
Dim i As Integer
Dim strNo As String
   
   m_dblPaperCnt = 0
   m_dblAmt = 0
   strNo = ""
   For i = 1 To grdDataList.Rows - 1
      'MODIFY BY SONIA 2014/1/6 加公司別, 7以後順延
      If grdDataList.TextMatrix(i, 0) = "V" And grdDataList.TextMatrix(i, 13) <> "Y" Then
         If strNo <> grdDataList.TextMatrix(i, 14) Then
            m_dblPaperCnt = m_dblPaperCnt + 1
            strNo = grdDataList.TextMatrix(i, 14)
         End If
         m_dblAmt = m_dblAmt + Val(Format(grdDataList.TextMatrix(i, 11), "######")) + Val(Format(grdDataList.TextMatrix(i, 12), "######"))
      End If
   Next i
   Label8 = "已勾選列印收據 " & m_dblPaperCnt & " 張, 共 " & m_dblAmt & " 元"
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   'Add By Sindy 2023/6/12
   If Combo3.Visible = True Then
      bolCancel = False
      Call Combo3_Validate(bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   'Add by Amy 2020/03/25 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus 'Add By Sindy 2020/7/15 讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
      If Combo3 = MsgText(601) Then
          Call Combo3_Validate(bolCancel)
          If bolCancel = True Then
              Combo3.SetFocus
              ConstrainCheck = False
              Exit Function
          End If
      ElseIf txtSales = MsgText(601) Then
          txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
      End If
   End If
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      'Modify by Amy 2020/03/25 +有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      'Modified by Lydia 2021/05/20 排除隱藏
      'ElseIf txtSales.Enabled = True Then
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
   End If
   '2023/6/12 END
'   Call txtSales_Validate(bolCancel)
'   If bolCancel = True Then
'      txtSales.SetFocus
'      txtSales_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol) = False Then
      If intErrCol = 0 Then
         txtSales.SetFocus
         txtSales_GotFocus
      ElseIf intErrCol = 1 Then
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
      Else
         txtSalesArea1.SetFocus
         txtSalesArea1_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
   End If
   
'   '2005/7/5 ADD BY SONIA 林永生71003檢查業務區範圍
'   If strUserNum = "71003" Then
'      If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'         MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'         MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
'
'   '2005/11/29 ADD BY SONIA 簡金泉69005檢查業務區範圍
''Removed by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
''   If strUserNum = "69005" Then
''      If txtSalesArea < "S1" Or txtSalesArea > "S19" Then
''         MsgBox "業務區起始條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''      If txtSalesArea1 < "S1" Or txtSalesArea1 > "S19" Then
''         MsgBox "業務區迄止條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea1.SetFocus
''         txtSalesArea1_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''   End If
''end 2019/12/30
'
'   'add by sonia 2016/12/21 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtSalesArea_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtSalesArea1_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'      Else
'         If Trim(txtSales) <> strUserNum Then
'            MsgBox "查專利處工程師時只可查自己的資料", vbExclamation
'            txtSales.SetFocus
'            txtSales_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'      End If
'   End If
'   'end 2016/12/21
'
'    'add by nickc 2008/01/18 加入外商主管  可以輸入相同組別的
'    If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") Then
'        If Trim(txtSales) = "" Then
'            MsgBox "智權人員不可以空白！", vbExclamation, "操作錯誤！"
'            txtSales.SetFocus
'            txtSales_GotFocus
'            ConstrainCheck = False
'            Exit Function
'        End If
'        If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txtSales) Then
'            MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
'            txtSales.SetFocus
'            txtSales_GotFocus
'            ConstrainCheck = False
'            Exit Function
'        End If
'    End If
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
End Function

Private Function Search_Process()
Dim strCon As String
Dim strSql As String, strInData As String
Dim stIdList As String, stConId As String, j As Long
Dim strConACC1K0 As String  '2011/11/16 ADD BY SONIA
Dim strConLoS As String 'Add By Sindy 2020/12/15
   
   Text3.Text = ""
   Text4.Text = ""
   Text5.Text = ""
   Label7 = "已通知列印收據　0　張"
   Label8 = "已勾選列印收據　0　張" 'Add By Sindy 2012/9/21
   
   Screen.MousePointer = vbHourglass
   grdDataList.Clear
   grdDataList.Rows = 2
   strCon = ""
   strConACC1K0 = ""  '2011/11/16 ADD BY SONIA
   strConLoS = "" 'Add By Sindy 2020/12/15
   
   '查詢權限：
   '所別
'cancel by sonia 2014/6/9
'   '2005/9/8 ADD BY SONIA 蔣律師要控制所別
'   If strUserNum = "79037" Then
'      strCon = strCon & " and st06 = '" & pub_strUserOffice & "' "
'   End If
'end 2014/6/9
   '2005/9/12 ADD BY SONIA 陳經理查詢所有智權人員要控制系統類別
   If strUserNum = "68005" And txtSales <> "68005" Then
      strCon = strCon & " and cp01 in ('CFT','FCT','S','CFC') "
      strConLoS = strConLoS & " and cp01 in ('CFT','FCT','S','CFC') " 'Add By Sindy 2020/12/15
   End If
   
   '區別
   'Modify By Sindy 98/03/11 若智權人員為80030時, 不限制區別
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   '2009/12/16 MODIFY BY SONIA 加巨京專利給郭雅娟79075看,所以不限制區別
   If txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   'Add by Amy 2014/05/21
   'Modify by Amy 2019/02/12 總經理業務工作代理人員
   ElseIf bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   '2011/9/20 add by sonia 查本人資料時不限制區別,因 98024 有調區
   ElseIf txtSales = strUserNum Then
   Else
      If txtSalesArea <> "" Then
         strCon = strCon & " and st15||'' >= '" & txtSalesArea & "' " 'Modify By Sindy 2021/8/4 cp12 => st15
         strConLoS = strConLoS & " and s2.st15||'' >= '" & txtSalesArea & "' " 'Add By Sindy 2020/12/15
      End If
      If txtSalesArea1 <> "" Then
         strCon = strCon & " and st15||'' <= '" & txtSalesArea1 & "' " 'Modify By Sindy 2021/8/4 cp12 => st15
         strConLoS = strConLoS & " and s2.st15||'' <= '" & txtSalesArea1 & "' " 'Add By Sindy 2020/12/15
      End If
   End If
   
   '智權人員
   If txtSales <> "" Then
        If (strUserNum <> "80030" And txtSales <> "80030") Then
            'Modify by Amy 2014/05/21 +if
            If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
                '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
                stIdList = PUB_GetSalesList(txtSales)
            Else
                'Add by Morgan 2010/1/29 若不是多員工編號時用 = 算符來加速查詢
                stIdList = PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, PUB_GetST06(Trim(txtSales)))
            End If
            'end 2014/05/21
            '2010/5/13 ADD BY SONIA 跨區帶人有考慮業務區條件
            If Pub_StrST52 Then
               strCon = ""
               strConLoS = "" 'Add By Sindy 2020/12/15
            End If
            '2010/5/13 END
            If InStr(stIdList, ",") = 0 Then
               stConId = " = " & stIdList & " "
            Else
               stConId = " in (" & stIdList & " ) "
            End If
            '2012/8/17 MODIFY BY SONIA 智權人員改抓收據檔
            'strCon = strCon & " and cp13||'' " & stConId & " "
            strCon = strCon & " and a0k20||'' " & stConId & " "
            strConLoS = strConLoS & " and s2.st01||'' " & stConId & " " 'Add By Sindy 2020/12/15
        Else
            '2008/3/31 ADD BY SONIA 查87027陳淑芳時同時查20001台中所
            'Modify By Sindy 98/02/27 查80030洪琬姿時同時查F4103
            If txtSales = "80030" Then
               strExc(0) = "select ST01 from STAFF where ST04<>'1' and ST03 like 'F1%' "
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
               strInData = "'80030','F4103'"
               If intI = 1 Then
                  adoRecordset.MoveFirst
                  Do While adoRecordset.EOF = False
                     strInData = strInData & ",'" & adoRecordset.Fields(0).Value & "'"
                     adoRecordset.MoveNext
                  Loop
               End If
               '2012/8/17 MODIFY BY SONIA 智權人員改抓收據檔
               'strCon = strCon & " and cp13||'' IN (" & strInData & ") "
               strCon = strCon & " and a0k20||'' IN (" & strInData & ") "
               strConLoS = strConLoS & " and s2.st01||'' IN (" & strInData & ") " 'Add By Sindy 2020/12/15
            Else
               '2012/8/17 MODIFY BY SONIA 智權人員改抓收據檔
               'strCon = strCon & " and cp13||'' = '" & txtSales & "' "
               strCon = strCon & " and a0k20||'' = '" & txtSales & "' "
               strConLoS = strConLoS & " and s2.st01||'' = '" & txtSales & "' " 'Add By Sindy 2020/12/15
            End If
        End If
   'Modify by Amy 2014/05/21
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            strCon = strCon & " and a0k20||'' in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') "
            strConLoS = strConLoS & " and s2.st01||'' in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') " 'Add By Sindy 2020/12/15
        End If
   'end 2014/05/21
   End If
   
   '客戶編號
   If Len(Text1) <> 0 And Len(Text2) <> 0 Then
      strCon = strCon & " and a0k03>='" & Text1 & "' AND a0k03<='" & Text2 & "' "
      strConLoS = strConLoS & " and a0k03>='" & Text1 & "' AND a0k03<='" & Text2 & "' " 'Add By Sindy 2020/12/15
      strConACC1K0 = strConACC1K0 & " and a0k03>='" & Text1 & "' AND a0k03<='" & Text2 & "' "
   End If
   '已送件
   If Check2.Value = 1 Then
      strCon = strCon & " and cp27||''>0 "
      strConLoS = strConLoS & " and cp27||''>0 " 'Add By Sindy 2020/12/15
   End If
   
   'Modify By Sindy 2011/1/28
'   strSql = "select '' as V,st02 as 智權人員,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,na03 as 申請國家, " & _
'               "decode(a0k23,'020',cpm04,cpm03) as 案件性質,cp09 as 總收文號,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,to_char(a0j09,'999,999,999') as 服務費,to_char(a0j10,'999,999,999') as 規費,a0k04 as 收據抬頭,sqldateT(cp27) as 發文日, " & _
'               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費 " & _
'               "From caseprogress, acc0k0, acc0j0, casepropertymap, nation, staff " & _
'               "where a0k01=cp60 " & _
'               "and cp01=cpm01(+) and cp10=cpm02(+) " & _
'               "and cp09=a0j01(+) " & _
'               "and a0j04=na01(+) " & _
'               "and cp13=st01(+) " & _
'               "and a0k32 is not null " & strCon & _
'               "order by cp12,cp13,a0k03,a0k01,cp09 "

   'Modified by Morgan 2011/10/31 考慮拆收據情形
   'strSql = "select '' as V,st02 as 智權人員,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人,na03 as 申請國家, " & _
               "decode(a0k23,'020',cpm04,cpm03) as 案件性質,cp09 as 總收文號,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,to_char(a0j09,'999,999,999') as 服務費,to_char(a0j10,'999,999,999') as 規費,a0k04 as 收據抬頭,sqldateT(cp27) as 發文日, " & _
               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費 " & _
               "From caseprogress, acc0k0, acc0j0, casepropertymap, nation, staff " & _
               "where a0k01=cp60 " & _
               "and cp01=cpm01(+) and cp10=cpm02(+) " & _
               "and cp09=a0j01(+) " & _
               "and a0j04=na01(+) " & _
               "and cp13=st01(+) " & _
               "and a0k32 is not null " & strCon & _
               "order by cp12,cp13,a0k03,a0k01,cp09 "
               
   'Modified by Morgan 2011/11/4 +本所案號排序 -- 辜(相同案件會有後收文未開收據程序也要排一起)
   '2011/11/16 MODIFY BY SONIA 未扣除銷帳E10022307
   'strSql = "select '' as V,st02 as 智權人員,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人,na03 as 申請國家, " & _
               "decode(a0k23,'020',cpm04,cpm03) as 案件性質,cp09 as 總收文號,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,to_char(a0j09,'999,999,999') as 服務費,to_char(a0j10,'999,999,999') as 規費,a0k04 as 收據抬頭,sqldateT(cp27) as 發文日, " & _
               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費 " & _
               "From acc0k0, acc0j0, caseprogress, casepropertymap, nation, staff " & _
               " where a0k32 is not null and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp01=cpm01(+) and cp10=cpm02(+) " & _
               " and a0j04=na01(+) " & _
               " and cp13=st01(+) " & strCon & _
               " order by cp12,cp13,a0k03,本所案號,a0k01,cp09 "
   'Modified by Morgan 2011/12/23 調整欄位順序--辜
   'strSql = "select '' as V,st02 as 智權人員,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人,na03 as 申請國家, " & _
               "decode(a0k23,'020',cpm04,cpm03) as 案件性質,cp09 as 總收文號,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,to_char(a0j09-NVL(A1U07,0),'999,999,999') as 服務費,to_char(a0j10-NVL(A1U09,0),'999,999,999') as 規費,a0k04 as 收據抬頭,sqldateT(cp27) as 發文日, " & _
               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費 " & _
               "From acc0k0, acc0j0, caseprogress, casepropertymap, nation, staff, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32 is not null and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp01=cpm01(+) and cp10=cpm02(+) " & _
               " and a0j04=na01(+) " & _
               " and cp13=st01(+) AND A0J01=A1U03(+) " & strCon & _
               " order by cp12,cp13,a0k03,本所案號,a0k01,cp09 "
   '2012/8/17 MODIFY BY SONIA 智權人員改抓收據檔
   'strSql = "select '' as V,st02 as 智權人員,sqldateT(cp27) as 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人,na03 as 申請國家, " & _
               "decode(a0k23,'020',cpm04,cpm03) as 案件性質,cp09 as 總收文號,to_char(a0j09-NVL(A1U07,0),'999,999,999') as 服務費,to_char(a0j10-NVL(A1U09,0),'999,999,999') as 規費,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,a0k04 as 收據抬頭, " & _
               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費 " & _
               "From acc0k0, acc0j0, caseprogress, casepropertymap, nation, staff, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32 is not null and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp01=cpm01(+) and cp10=cpm02(+) " & _
               " and a0j04=na01(+) " & _
               " and cp13=st01(+) AND A0J01=A1U03(+) " & strCon & _
               " order by cp12,cp13,a0k03,本所案號,a0k01,cp09 "
   '20131231 START modify by eric 增加 "公司"欄位
   'strSql = "select '' as V,st02 as 智權人員,sqldateT(cp27) as 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人,na03 as 申請國家, " & _
   '            "decode(a0k23,'020',cpm04,cpm03) as 案件性質,cp09 as 總收文號,to_char(a0j09-NVL(A1U07,0),'999,999,999') as 服務費,to_char(a0j10-NVL(A1U09,0),'999,999,999') as 規費,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,a0k04 as 收據抬頭, " & _
   '            "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費 " & _
   '            "From acc0k0, acc0j0, caseprogress, casepropertymap, nation, staff, " & _
   '            "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
   '            " where a0k32 is not null and a0j13(+)=a0k01" & _
   '            " and cp09(+)=a0j01 " & _
   '            " and cp01=cpm01(+) and cp10=cpm02(+) " & _
   '            " and a0j04=na01(+) " & _
   '            " and a0k20=st01(+) AND A0J01=A1U03(+) " & strCon & _
   '            " order by cp12,a0k20,a0k03,本所案號,a0k01,cp09 "
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Lydia 2021/07/27 申請人=>申請人名稱
   'Modified by Lydia 2023/11/13 +排除A0K40開立INVOICE , a0k32=Z不列印收據=>AND A0K32<>'Z'
   strSql = "select '' as V,st02 as 智權人員,'' as 介紹人,sqldateT(cp27) as 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人名稱,na03 as 申請國家,a0k11 as 公司, " & _
               "decode(a0k23,'000',CPM03,CPM04) as 案件性質,cp09 as 總收文號,to_char(a0j09-NVL(A1U07,0),'999,999,999') as 服務費,to_char(a0j10-NVL(A1U09,0),'999,999,999') as 規費,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,a0k04 as 收據抬頭, " & _
               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費,cp12,a0k20,a0k03 " & _
               "From acc0k0, acc0j0, caseprogress, casepropertymap, nation, staff, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL AND A0K32<>'Z' " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32 is not null AND A0K32<>'Z' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp01=cpm01(+) and cp10=cpm02(+) " & _
               " and a0j04=na01(+) " & _
               " and a0k20=st01(+) AND A0J01=A1U03(+) " & strCon
   'Add By Sindy 2020/12/15 + 案源資料檔
   'Modified by Lydia 2021/07/27 申請人=>申請人名稱
   'Modified by Lydia 2023/11/13 +排除A0K40開立INVOICE , a0k32=Z不列印收據=>AND A0K32<>'Z'
   strSql = strSql & " union all select '' as V,s1.st02 as 智權人員,GETSTAFFNAMELIST(LOS04) as 介紹人,sqldateT(cp27) as 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,'' as 案件名稱,'' as 申請人名稱,na03 as 申請國家,a0k11 as 公司, " & _
               "decode(a0k23,'000',CPM03,CPM04) as 案件性質,cp09 as 總收文號,to_char(a0j09-NVL(A1U07,0),'999,999,999') as 服務費,to_char(a0j10-NVL(A1U09,0),'999,999,999') as 規費,decode(a0k32,'N',' ',a0k32) as 可列印,a0k01 as 收據號碼,a0k04 as 收據抬頭, " & _
               "sqldateT(cp05) As 收文日, to_char(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0)),'999,999,999') As 已送件規費,cp12,a0k20,a0k03 " & _
               "From acc0k0, acc0j0, caseprogress, casepropertymap, nation, staff s1, staff s2, LawOfficeSource, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL AND A0K32<>'Z' " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32 is not null AND A0K32<>'Z' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp01=cpm01(+) and cp10=cpm02(+) " & _
               " and a0j04=na01(+) " & _
               " and cp13=s1.st01(+) AND A0J01=A1U03(+) " & _
               " and CP162 is not null and CP162=LOS15(+) " & _
               " and substr(LOS04,1,5)=s2.st01(+) " & strConLoS
   'strSql = strSql & " order by cp12,a0k20,a0k03,本所案號,a0k01,cp09 "
   strSql = strSql & " order by cp12,a0k20,a0k03,本所案號,收據號碼,總收文號 "
   '20131231 END
   'end 2011/10/31
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   grdDataList.Clear
   grdDataList.FixedCols = 0 'Added by Lydia 2021/07/27 清除固定欄位
   If intI = 1 Then
      Set grdDataList.Recordset = adoRecordset.Clone
      grdDataList.FixedCols = 5 'Added by Lydia 2021/07/27 固定欄位=4
      'Add By Sindy 2011/1/28
      grdDataList.Enabled = False
      Dim CaseNo
      For j = 1 To grdDataList.Rows - 1
         'Modified by Morgan 2011/12/23 調整欄位順序--辜
         'CaseNo = Split(grdDataList.TextMatrix(j, 2), "-")
         CaseNo = Split(grdDataList.TextMatrix(j, 4), "-")
         'Modified by Lydia 2021/07/27 申請人=>申請人名稱
         strSql = "SELECT substr(tm05,1,8) as 案件名稱,substr(cu04,1,4) as 申請人名稱 FROM trademark,customer WHERE tm01='" & CaseNo(0) & "' and tm02='" & CaseNo(1) & "' and tm03='" & CaseNo(2) & "' and tm04='" & CaseNo(3) & "' and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1) " & _
            "Union SELECT substr(pa05,1,8) as 案件名稱,substr(cu04,1,4) as 申請人名稱 FROM patent,customer WHERE pa01='" & CaseNo(0) & "' and pa02='" & CaseNo(1) & "' and pa03='" & CaseNo(2) & "' and pa04='" & CaseNo(3) & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) " & _
            "Union SELECT substr(sp05,1,8) as 案件名稱,substr(cu04,1,4) as 申請人名稱 FROM servicepractice,customer WHERE sp01='" & CaseNo(0) & "' and sp02='" & CaseNo(1) & "' and sp03='" & CaseNo(2) & "' and sp04='" & CaseNo(3) & "' and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) " & _
            "Union SELECT substr(hc06,1,8) as 案件名稱,substr(cu04,1,4) as 申請人名稱 FROM hirecase,customer WHERE hc01='" & CaseNo(0) & "' and hc02='" & CaseNo(1) & "' and hc03='" & CaseNo(2) & "' and hc04='" & CaseNo(3) & "' and cu01(+)=substr(hc05,1,8) and cu02(+)=substr(hc05,9,1) " & _
            "Union SELECT substr(lc05,1,8) as 案件名稱,substr(cu04,1,4) as 申請人名稱 FROM lawcase,customer WHERE lc01='" & CaseNo(0) & "' and lc02='" & CaseNo(1) & "' and lc03='" & CaseNo(2) & "' and lc04='" & CaseNo(3) & "' and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9,1) "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'Modified by Morgan 2011/12/23 調整欄位順序--辜
            'grdDataList.TextMatrix(j, 3) = "" & RsTemp.Fields(0)
            'grdDataList.TextMatrix(j, 4) = "" & RsTemp.Fields(1)
            grdDataList.TextMatrix(j, 5) = "" & RsTemp.Fields(0)
            grdDataList.TextMatrix(j, 6) = "" & RsTemp.Fields(1)
         End If
         'Added by Lydia 2021/07/27 固定欄位顏色改回預設
         For intI = 0 To 4
            grdDataList.row = j
            grdDataList.col = intI
            grdDataList.CellBackColor = QBColor(15)
         Next intI
         'end 2021/07/27
         
         '2015/7/28 add by sonia 已通知列印者本所案號欄變紅色
         If grdDataList.TextMatrix(j, 13) = "Y" Then
            grdDataList.row = j
            grdDataList.col = 4
            grdDataList.CellBackColor = &H8080FF '變紅
         End If
         '2015/7/28 end
      Next j
      grdDataList.Enabled = True
      '2011/1/28 End
      
      '計算合計金額
'Modified by Morgan 2011/10/31 考慮拆收據情形改語法
'      '服務費
'      strSql = "select sum(a0j09) " & _
'                  "From caseprogress, acc0k0, acc0j0, staff " & _
'                  "where a0k01=cp60 " & _
'                  "and cp09=a0j01(+) " & _
'                  "and cp13=st01(+) " & _
'                  "and a0k32 is not null " & strCon
'
'      intI = 1
'      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         Text3 = Format(adoRecordset.Fields(0), "##,###")
'      End If
'      If Text3 = "" Then Text3 = "0"
'      '規費
'      strSql = "select sum(a0j10) " & _
'                  "From caseprogress, acc0k0, acc0j0, staff " & _
'                  "where a0k01=cp60 " & _
'                  "and cp09=a0j01(+) " & _
'                  "and cp13=st01(+) " & _
'                  "and a0k32 is not null " & strCon
'      intI = 1
'      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         Text4 = Format(adoRecordset.Fields(0), "##,###")
'      End If
'      If Text4 = "" Then Text4 = "0"
'      '已送件規費
'      strSql = "select sum(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0))) " & _
                  "From caseprogress, acc0k0, acc0j0, staff " & _
                  "where a0k01=cp60 " & _
                  "and cp09=a0j01(+) " & _
                  "and cp13=st01(+) " & _
                  "and a0k32 is not null " & strCon

      '2012/8/17 MODIFY BY SONIA 智權人員改抓收據檔
      'strSql = "select sum(a0j09),sum(a0j10),sum(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0))) " & _
               " From acc0k0, acc0j0, caseprogress, staff " & _
               " where a0k32 is not null and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp13=st01(+) " & strCon
      '2012/9/19 MODIFY BY SONIA 未扣除銷帳
      'strSql = "select sum(a0j09),sum(a0j10),sum(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0))) " & _
               " From acc0k0, acc0j0, caseprogress, staff " & _
               " where a0k32 is not null and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and a0k20=st01(+) " & strCon
      'Modified by Lydia 2023/11/13 +排除A0K40開立INVOICE , a0k32=Z不列印收據=>AND A0K32<>'Z'
      strSql = "select sum(serFee),sum(Fee),sum(cp17) from (select sum(a0j09-NVL(A1U07,0)) as serFee,sum(a0j10-NVL(A1U09,0)) as Fee,sum(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0))) as cp17 " & _
               " From acc0k0, acc0j0, caseprogress, staff, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL AND A0K32<>'Z' " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32 is not null AND A0K32<>'Z' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and a0k20=st01(+) AND A0J01=A1U03(+) " & strCon
      'Add By Sindy 2020/12/15 + 案源資料檔
      'Modified by Lydia 2023/11/13 +排除A0K40開立INVOICE , a0k32=Z不列印收據=>AND A0K32<>'Z'
      strSql = strSql & " union all select sum(a0j09-NVL(A1U07,0)) as serFee,sum(a0j10-NVL(A1U09,0)) as Fee,sum(decode(nvl(cp27, 0), 0, 0, nvl(cp17, 0))) as cp17 " & _
               " From acc0k0, acc0j0, caseprogress, staff s1, staff s2, LawOfficeSource, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL AND A0K32<>'Z' " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32 is not null AND A0K32<>'Z' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp13=s1.st01(+) AND A0J01=A1U03(+) " & _
               " and CP162 is not null and CP162=LOS15(+) " & _
               " and substr(LOS04,1,5)=s2.st01(+) " & strConLoS & ")"
      '2020/12/15 END
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modified by Morgan 2011/10/31
         'Text5 = Format(adoRecordset.Fields(0), "##,###")
         Text3 = Format(adoRecordset.Fields(0), "##,###")
         Text4 = Format(adoRecordset.Fields(1), "##,###")
         Text5 = Format(adoRecordset.Fields(2), "##,###")
         'end 2011/10/31
      End If
      If Text3 = "" Then Text3 = "0" 'Add by Morgan 2011/10/31
      If Text4 = "" Then Text4 = "0" 'Add by Morgan 2011/10/31
      If Text5 = "" Then Text5 = "0"

      '已通知列印收據張數
      'Modified by Morgan 2011/10/31 考慮拆收據情形改語法
      'strSql = "select count(*) from (select distinct a0k01 " & _
                  "From caseprogress, acc0k0, acc0j0, staff " & _
                  "where a0k01=cp60 " & _
                  "and cp09=a0j01(+) " & _
                  "and cp13=st01(+) " & _
                  "and a0k32='Y' " & strCon & ")"
      '2012/8/17 MODIFY BY SONIA 智權人員改抓收據檔
      'strSql = " select count(distinct a0k01) " & _
               " From acc0k0, acc0j0, caseprogress, staff " & _
               " where a0k32='Y' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and cp13=st01(+) " & strCon
      '2012/9/19 MODIFY BY SONIA 扣除銷帳並加顯示金額
      'strSql = " select count(distinct a0k01) " & _
               " From acc0k0, acc0j0, caseprogress, staff " & _
               " where a0k32='Y' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and a0k20=st01(+) " & strCon
      'Modified by Lydia 2023/11/13 +排除A0K40開立INVOICE , a0k32=Z不列印收據=>AND A0K32<>'Z'
      strSql = " select count(distinct a0k01),to_char(sum(a0j09+a0j10-NVL(A1U07,0)-NVL(A1U09,0)),'999,999,999') " & _
               " From acc0k0, acc0j0, caseprogress, staff, " & _
               "(SELECT A1U03,SUM(NVL(A1U07,0)) A1U07,SUM(NVL(A1U09,0)) A1U09 FROM ACC0K0,ACC1U0 WHERE A0K32 IS NOT NULL AND A0K32<>'Z' " & strConACC1K0 & " AND A0K01=A1U02(+) GROUP BY A1U03) " & _
               " where a0k32='Y' and a0j13(+)=a0k01" & _
               " and cp09(+)=a0j01 " & _
               " and a0k20=st01(+) AND A0J01=A1U03(+) " & strCon
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         Label7 = "已通知列印收據 " & Val(adoRecordset.Fields(0)) & " 張, 共 " & Trim(adoRecordset.Fields(1)) & " 元"
      End If
      cmdOK(3).Enabled = True
      cmdOK(4).Enabled = True
      cmdOK(0).Enabled = True
      cmdOK(2).Enabled = True
   Else
      ShowNoData
      Me.Enabled = True
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End If
   Screen.MousePointer = vbDefault
End Function

Private Sub PrintData()
Dim i As Long, j As Long

'Modified by Morgan 2020/10/30
'Set Printer = Printers(Combo1.ListIndex)
PUB_RestorePrinter Combo1
'end 2020/10/30
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印

iLine = 1
strType = ""
For j = 1 To grdDataList.Rows
   For i = 1 To 10
      strTemp(i) = ""
   Next i
   If j = grdDataList.Rows Then
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print String(205, "-")
      iLine = iLine + 1
      strTemp(8) = "合計"
      strTemp(9) = Text3
      strTemp(10) = Text4
'      strTemp(10) = Text5
   Else
      'Modified by Morgan 2011/12/23 調整欄位順序--辜
      'strTemp(1) = CheckStr(grdDataList.TextMatrix(j, 13))
      'strTemp(2) = Left(CheckStr(grdDataList.TextMatrix(j, 12)) & "            ", 12)
      'strTemp(3) = CheckStr(grdDataList.TextMatrix(j, 1))
      'strTemp(4) = CheckStr(grdDataList.TextMatrix(j, 2))
      'strTemp(5) = Left(CheckStr(grdDataList.TextMatrix(j, 5)) & "     ", 5)
      'strTemp(6) = Left(CheckStr(grdDataList.TextMatrix(j, 6)) & "     ", 5)
      'strTemp(7) = CheckStr(grdDataList.TextMatrix(j, 9))
      'strTemp(8) = CheckStr(grdDataList.TextMatrix(j, 10))
      'strTemp(9) = CheckStr(grdDataList.TextMatrix(j, 11))
      strTemp(1) = CheckStr(grdDataList.TextMatrix(j, 3)) '發文日
      'MODIFY BY SONIA 2014/1/6 加公司別, 7以後順延
      strTemp(2) = Left(CheckStr(grdDataList.TextMatrix(j, 15)) & "            ", 12) '收據抬頭
      strTemp(3) = CheckStr(grdDataList.TextMatrix(j, 1)) '智權人員
      strTemp(4) = CheckStr(grdDataList.TextMatrix(j, 2)) '介紹人 Add By Sindy 2020/12/15
      strTemp(5) = CheckStr(grdDataList.TextMatrix(j, 4)) '本所案號
      strTemp(6) = Left(CheckStr(grdDataList.TextMatrix(j, 7)) & "     ", 5) '申請國家
      strTemp(7) = Left(CheckStr(grdDataList.TextMatrix(j, 9)) & "     ", 5) '案件性質
      strTemp(8) = CheckStr(grdDataList.TextMatrix(j, 14)) '收據號碼
      strTemp(9) = CheckStr(grdDataList.TextMatrix(j, 11)) '服務費
      strTemp(10) = CheckStr(grdDataList.TextMatrix(j, 12)) '規費
      'end 2011/12/23
      
   End If
   If iLine > 37 Or iLine = 1 Then
      If strType <> "" Then Printer.NewPage
      iLine = 1
      PrintTitle '列印表頭
   End If
   PrintDetail
   strType = strTemp(2)
Next j
Printer.EndDoc
PUB_RestorePrinter strPrinter 'Added by Morgan 2020/10/30
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 4700
PLeft(4) = 5800 '介紹人 Add By Sindy 2020/12/15
PLeft(5) = 6800
PLeft(6) = 8800
PLeft(7) = 10000
PLeft(8) = 11500
PLeft(9) = 14000
PLeft(10) = 15500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("未列印收據/請款單明細") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "未列印收據/請款單明細"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
If Trim(Text1.Text) <> "" Or Trim(Text2.Text) <> "" Then
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("客戶編號：" & Text1.Text & "-" & Text2.Text) / 2)
   Printer.CurrentY = 900
   Printer.Print "客戶編號：" & Text1.Text & "-" & Text2.Text
End If
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
If Trim(txtSalesArea.Text) <> "" Or Trim(txtSalesArea1.Text) <> "" Then
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 1200
   Printer.Print "業務區：" & txtSalesArea & " " & txtSalesArea1
End If
If Trim(txtSales.Text) <> "" Then
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("智權人員：" & txtSales & " " & lblSalesName) / 2)
   Printer.CurrentY = 1200
   Printer.Print "智權人員：" & txtSales & " " & lblSalesName
End If
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine = 6
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "發文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "收據抬頭"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "智權人員"
'Add By Sindy 2020/12/15
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "介紹人"
'2020/12/15 END
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "申請國家"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "案件性質"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "收據號碼"
Printer.CurrentX = PLeft(9) - Printer.TextWidth("服務費")
Printer.CurrentY = iLine * 300
Printer.Print "服務費"
Printer.CurrentX = PLeft(10) - Printer.TextWidth("規費")
Printer.CurrentY = iLine * 300
Printer.Print "規費"
'Printer.CurrentX = PLeft(10) - Printer.TextWidth("已送件規費")
'Printer.CurrentY = iLine * 300
'Printer.Print "已送件規費"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(205, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 10
   'If m_j = 8 Or m_j = 9 Or m_j = 10 Then
   If m_j = 9 Or m_j = 10 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
   Text2 = Text1
End Sub

Private Sub Text2_GotFocus()
   If Text1.Text <> "" Then
      Text2.Text = Left(Text1.Text, 6) & "ZZZ"
   End If
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   If Trim(txtSales) = "" Then
       lblSalesName = ""
   End If
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
If Trim(txtSalesArea1) <> "" Then
   If RunNick(txtSalesArea, txtSalesArea1) = True Then
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2023/6/12
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, , txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
End Sub

Private Sub Text6_GotFocus()
   'Added by Lydia 2015/06/25 輸入申請人名稱按Enter自動執行搜尋後,執行查詢
   cmdSearch.Default = False: cmdFind.Default = True
   
   TextInverse Me.Text6
   OpenIme
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If CheckLengthIsOK(Me.Text6.Text, 40) = False Then
      Cancel = True
   End If
End Sub

'Added by Lydia 2015/06/25 輸入申請人名稱按Enter自動執行搜尋後,執行查詢
Private Sub Text6_LostFocus()
   cmdSearch.Default = True: cmdFind.Default = False
End Sub

'Add By Sindy 2023/6/12
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo3_LostFocus()
   If Trim(Combo3) <> "" And Trim(Combo3) <> "全部" Then
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   ElseIf Trim(Combo3) <> "全部" Then
      txtSales = ""
   End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
Dim strEmp As String
Dim stTmp As String 'Add by Amy 2020/03/25
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      'Add by Amy 2020/03/25 只能輸入下拉選單中已有的人員
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      'Modify By Sindy 2020/6/15 Mark
'      If InStr(m_strListPer, stTmp) = 0 And stTmp <> strUserNum And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "不可輸入下拉選單以外的人員！"
'         Cancel = True
'         Combo3.SetFocus
'         Exit Sub
'      End If
      'end 2020/03/25
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales, True)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   'Modify by Amy 2023/05/09 +st05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        'Add by Amy 2020/03/25 下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2020/7/14
'        'If bolAreaMan = False And Pub_StrUserSt03 <> "M51" Then
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'        '2020/7/14 END
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
'        'end 2020/03/25
   '2024/8/5 END
   End If
   'end 2016/6/7
End Sub
'2023/6/12 END
