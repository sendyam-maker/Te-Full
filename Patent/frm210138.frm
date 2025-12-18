VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210138 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶案件預算預估表"
   ClientHeight    =   4812
   ClientLeft      =   3696
   ClientTop       =   1560
   ClientWidth     =   7476
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4812
   ScaleWidth      =   7476
   Begin VB.Frame Frame3 
      Height          =   380
      Left            =   1110
      TabIndex        =   44
      Top             =   3810
      Width           =   2600
      Begin VB.OptionButton Option1 
         Caption         =   "Excel"
         Height          =   180
         Index           =   1
         Left            =   780
         TabIndex        =   47
         Top             =   144
         Width           =   730
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Word"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   46
         Top             =   144
         Value           =   -1  'True
         Width           =   700
      End
      Begin VB.CheckBox Check4 
         Caption         =   "代表圖"
         Height          =   200
         Left            =   1530
         TabIndex        =   45
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "電子檔(勾選時，須副總簽字認可)"
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   43
      Top             =   4080
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CheckBox Check3 
      Caption         =   "紙本"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   42
      Top             =   4080
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Enabled         =   0   'False
      Height          =   540
      Left            =   120
      TabIndex        =   39
      Top             =   4260
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   40
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   41
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1350
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1350
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   16
      Left            =   1110
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1050
      Width           =   975
   End
   Begin VB.TextBox txtMail 
      Height          =   270
      Left            =   1110
      TabIndex        =   7
      Top             =   1956
      Width           =   6270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   0
      Top             =   450
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1110
      MaxLength       =   9
      TabIndex        =   1
      Top             =   750
      Width           =   1275
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   2250
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1050
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1536
      Left            =   120
      TabIndex        =   27
      Top             =   2268
      Width           =   7275
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   8
         Left            =   5565
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1050
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "網域"
         Height          =   255
         Index           =   6
         Left            =   5445
         TabIndex        =   14
         Top             =   180
         Width           =   705
      End
      Begin VB.CheckBox Check1 
         Caption         =   "條碼"
         Height          =   255
         Index           =   5
         Left            =   4725
         TabIndex        =   13
         Top             =   180
         Width           =   705
      End
      Begin VB.CheckBox Check1 
         Caption         =   "法務"
         Height          =   255
         Index           =   4
         Left            =   4005
         TabIndex        =   12
         Top             =   180
         Width           =   705
      End
      Begin VB.CheckBox Check1 
         Caption         =   "著作權"
         Height          =   255
         Index           =   3
         Left            =   3105
         TabIndex        =   11
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標"
         Height          =   255
         Index           =   2
         Left            =   2415
         TabIndex        =   10
         Top             =   180
         Width           =   705
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利"
         Height          =   255
         Index           =   1
         Left            =   1725
         TabIndex        =   9
         Top             =   180
         Width           =   705
      End
      Begin VB.CheckBox Check1 
         Caption         =   "全部"
         Height          =   255
         Index           =   0
         Left            =   1035
         TabIndex        =   8
         Top             =   180
         Width           =   705
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   9
         Left            =   1035
         MaxLength       =   1
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   5
         Left            =   1035
         MaxLength       =   4
         TabIndex        =   17
         Top             =   1065
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   6
         Left            =   2370
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1065
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   2
         Left            =   1035
         TabIndex        =   15
         Top             =   165
         Visible         =   0   'False
         Width           =   6060
      End
      Begin VB.Label Label3 
         Caption         =   "5.案戶案件案號)"
         Height          =   180
         Left            =   1530
         TabIndex        =   36
         Top             =   810
         Width           =   5235
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "(Y:分開)"
         Height          =   180
         Left            =   6045
         TabIndex        =   34
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "國內外是否分開列印:"
         Height          =   180
         Left            =   3840
         TabIndex        =   33
         Top             =   1095
         Width           =   1665
      End
      Begin VB.Line Line3 
         X1              =   2010
         X2              =   2250
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(1.本所案號 2.案件名稱 3.申請國家+本所案號 4.申請國家+案件名稱 "
         Height          =   180
         Left            =   1470
         TabIndex        =   31
         Top             =   600
         Width           =   5280
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "輸出順序："
         Height          =   180
         Left            =   60
         TabIndex        =   30
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         Height          =   180
         Left            =   60
         TabIndex        =   29
         Top             =   1095
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "系統別："
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2565
      MaxLength       =   9
      TabIndex        =   2
      Top             =   750
      Width           =   1275
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   6660
      Top             =   516
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   7860
      Top             =   456
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   6570
      TabIndex        =   21
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "寄電子檔(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5280
      TabIndex        =   20
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "110/9/23加註：CFT案寄琬姿及May，副本給江協理，註明彙整後再寄智權人員。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1080
      TabIndex        =   38
      Top             =   1692
      Width           =   6264
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "年費預算要下605~607"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3100
      TabIndex        =   37
      Top             =   1400
      Width           =   1710
   End
   Begin VB.Line Line2 
      X1              =   1890
      X2              =   2130
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   35
      Top             =   1400
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   32
      Top             =   1100
      Width           =   900
   End
   Begin VB.Line Line5 
      X1              =   2055
      X2              =   2295
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Left            =   150
      TabIndex        =   26
      Top             =   800
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2355
      X2              =   2595
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   150
      TabIndex        =   25
      Top             =   480
      Width           =   900
   End
   Begin MSForms.Label lblSalesName 
      Height          =   240
      Left            =   1965
      TabIndex        =   24
      Top             =   480
      Width           =   2115
      VariousPropertyBits=   27
      Caption         =   "Name"
      Size            =   "3731;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "輸出方式："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   3945
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail："
      Height          =   180
      Left            =   156
      TabIndex        =   22
      Top             =   2004
      Width           =   660
   End
End
Attribute VB_Name = "frm210138"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/05/02 Form2.0已修改 (lblSalesName/Printer 不使用無需修改)
'Memo By Sonia 2012/12/6 智權人員欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, k As Integer
Dim s As Integer, i As Integer, j As Integer
Dim strTemp(0 To 24) As String, StrTemp5(0 To 14) As String, iPrint As Integer, Page As Integer
Dim PLeft(0 To 20) As Integer, strTemp1 As Variant, strTemp2 As Variant, SeekPrint As Integer, SeekPrintL As Integer, IntTot As Integer
Dim m_ColCustName As String '申請人名稱欄位
Dim m_ColCustAdd As String '申請人地址欄位
Dim m_ColAgName As String '代理人名稱欄位
Dim m_ColAgAdd As String '代理人地址欄位
Dim m_strSQL1_2 As String, m_strSQL1_3 As String, m_strSQL1_4 As String, m_strSQL1_5 As String
Dim m_strSQL2_2 As String, m_strSQL2_3 As String, m_strSQL2_4 As String, m_strSQL2_5 As String
Dim m_strSQL3_2 As String, m_strSQL3_3 As String, m_strSQL3_4 As String, m_strSQL3_5 As String
Dim m_strSQL4_2 As String, m_strSQL4_3 As String, m_strSQL4_4 As String, m_strSQL4_5 As String
Dim m_strSQL5_2 As String, m_strSQL5_3 As String, m_strSQL5_4 As String, m_strSQL5_5 As String
Const strIdfTag As String = "DECODE(TM58,NULL,' ', DECODE( INSTR(TM58,'原為聯合商標',1,1),0,DECODE( INSTR(TM58,'原為服務標章',1,1),0,DECODE(INSTR(TM58,'原為聯合服務標章',1,1),0,'','C'),'B'),'A'))"
Dim g_WordAp As Word.Application
Dim IsOpenWord As Boolean, IsHaveData As Boolean, blnWordNewPage As Boolean
Dim o_PaperHeight As Integer
Dim tmpPrName As String, oldtmpPrName As String
Dim strTBF As String 'Add by Amy 2017/07/25 +欄位名(因與R050317_C 共用TempTB)
Dim intChoose As Integer  'Add by Amy 2022/04/29 1-Word/2-Excel
'Add by Amy 2022/05/02
Dim xlsCustPoint As New Excel.Application, wksrpt As New Worksheet
Dim bolShowPic As Boolean, IsOpenExcel(1 To 3) As Boolean '顯示代表圖/是否開啟Excel
Dim intXlsRow As Integer, intField As Integer, intTitleR As Integer  '目前列/欄位起始/抬頭列
Dim oldCaseNo As String, strWkName As String, ReportName As String, strXlsData As String  '用於 記錄本所案號 重覆者不需抓圖檔/工作表名/報表名/Excel資料
Dim SetTitle, setXlsWidth '記錄抬頭(陣列大小:專利 商標 其他 最大欄位數)/欄寬
Dim strXlsTp(4) As String, strFileN As String 'for Excel/檔名
Dim setFieldNo_txt() As String '記錄需設定文字欄位格式的編號3
Dim setFieldNo() As String '記錄需設定欄位格式的編號2
Dim setFieldType() As String '記錄設定欄位格式2
Dim intCntP As Integer, intCntT As Integer, intCntO As Integer 'Add by Amy 2022/07/19

Private Sub Check3_Click(Index As Integer)
'Mark by Amy 2022/05/02 和秀玲確認不會使用(程式copy 總簿,原本就只產生Word)
'If Index = 0 Then
'   'Add by Amy 2020/04/29
'   Option1(0).Value = vbChecked
'   Frame3.Enabled = False
'   'end 2022/04/29
'   If Check3(Index).Value = vbChecked Then
'      Frame1.Enabled = True '印表機
'      Check3(1).Value = 0
'   Else
'      Frame1.Enabled = False '印表機
'   End If
'ElseIf Index = 1 Then
   If Check3(Index).Value = vbChecked Then
      txtMail.Enabled = True
      Check3(0).Value = 0
      'Add by Amy 2020/04/29
      Call Option1_Click(0)
      Frame3.Enabled = True
      'end 2022/04/29
   Else
      txtMail.Enabled = False
      Frame3.Enabled = False 'Add by Amy 2022/04/29
   End If
'End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strTemp As String
Dim strMsg As String, oChk As CheckBox 'Add by Amy 2022/07/19
   
On Error GoTo ErrorHandler
   
Select Case Index
Case 0 '確定
    
   '假設沒有 word 文件
   IsOpenWord = False
   '假設沒有資料
   IsHaveData = False
   'Add by Amy 2022/04/29 +Excel
   IsOpenExcel(1) = False: IsOpenExcel(2) = False: IsOpenExcel(3) = False
   strFileN = "": bolShowPic = False: intChoose = 0: intField = 65
   'end 2022/04/29
   intCntP = 0: intCntT = 0: intCntO = 0 'Add by Amy 2022/07/19
         
   'Mark by Amy 2022/05/02 原本就沒使用
'   '紙本
'   If Check3(0).Value = vbChecked Then
'      Set Printer = Printers(Combo1.ListIndex)
'      '故意設定紙張屬性以便清除印表機狀態(相同印表機驅動程式會沿用原設定值,Ex.進紙槽)
'      Printer.PaperSize = 9
'      Printer.EndDoc
'   End If
   DoEvents
   'Modify  by Amy 2022/04/29 原檢查程式改至FromCheck
   If FormCheck = True Then
      '清除查詢印表記錄檔欄位
      ClearQueryLog (Me.Name)
      '智權人員
      If Len(Trim(txt1(11))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label5 & Trim(txt1(11)) & lblSalesName
      '客戶編號
      If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1 & Trim(txt1(0)) & "-" & Trim(txt1(1))
      End If
      '本所期限
      If Len(Trim(txt1(16))) <> 0 Or Len(Trim(txt1(17))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label2(5) & Trim(txt1(16)) & "-" & Trim(txt1(17))
      End If
      'Add By Sindy 2012/12/20
      '案件性質
      If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label2(2) & Trim(txt1(3)) & "-" & Trim(txt1(4))
      End If
      '2012/12/20 End
      'E -MAIL
      If Len(Trim(txtMail)) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label14 & Trim(txtMail)
      '系統類別
      If Len(Trim(txt1(2))) <> 0 And Me.Check1(0).Value = vbChecked Then
         pub_QL05 = pub_QL05 & ";系統類別：" & Trim(txt1(2))
      Else
         strTemp = ""
         If Me.Check1(1).Value = vbChecked Then
            strTemp = strTemp & ",CFP,FCP,P"
         End If
         If Me.Check1(2).Value = vbChecked Then
            strTemp = strTemp & ",CFT,FCT,T,TF"
         End If
         If Me.Check1(3).Value = vbChecked Then
            strTemp = strTemp & ",TC,CFC"
         End If
         If Me.Check1(4).Value = vbChecked Then
            strTemp = strTemp & ",L,CFL,FCL,LA,LIN"
         End If
         If Me.Check1(5).Value = vbChecked Then
            strTemp = strTemp & ",TB"
         End If
         If Me.Check1(6).Value = vbChecked Then
            strTemp = strTemp & ",TD"
         End If
         pub_QL05 = pub_QL05 & ";系統類別：" & Mid(strTemp, 2, Len(strTemp))
      End If
      '系統別
      pub_QL05 = pub_QL05 & ";" & Label2(0)
      If Check1(0).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(0).Caption
      If Check1(1).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(1).Caption
      If Check1(2).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(2).Caption
      If Check1(3).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(3).Caption
      If Check1(4).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(4).Caption
      If Check1(5).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(5).Caption
      If Check1(6).Value = vbChecked Then pub_QL05 = pub_QL05 & "," & Check1(6).Caption
      '國內外是否分開列印
      If Len(Trim(txt1(8))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label8 & Trim(txt1(8)) & Label9
      '輸出順序
      If Len(Trim(txt1(9))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label10 & Trim(txt1(9)) & Label11
      '申請國家
      If Len(Trim(txt1(5))) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label4 & Trim(txt1(5)) & "-" & Trim(txt1(6))
      End If
      
      'Add by Amy 222/04/29
      If Option1(0).Value = True Then
            intChoose = 1
      ElseIf Option1(1).Value = True Then
            intChoose = 2
      End If
      'end 2022/04/29
    
      'Modfiy by Amy 2022/05/02 原 if 判斷「請選擇系統類別」程式改至FromCheck,加判斷是否顯示代表圖
         If Check4.Value = vbChecked Then bolShowPic = True
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         '初始化頁數
         Page = 0
         'Add by Amy 2017/07/25 +欄位名(因與R050317_C 共用TempTB)
         strTBF = "(R013001,R013002,R013003,R013004,R013005,R013006,R013007,R013008,R013009,R013010,R013011,R013012,R013013,R013014,R013015,R013016,R013017,R013018,R013019" & _
                        ",ID,R013020,R013021,R013022,R013023,R013024,R013025,R013026,R013027)"
         'Modify by Amy 2022/05/02 +ReportName及Excel
         ReportName = "商標"
         If intChoose = 2 Then Page = 0
         ProcessT
         ReportName = "專利"
         If intChoose = 2 Then Page = 0
         ProcessP
         ReportName = "其他"
         If intChoose = 2 Then Page = 0
         ProcessO
         
         Screen.MousePointer = vbHourglass
         '將 word 存檔，寄信
         'Modify by Amy 2022/05/02 +Excel
         If IsOpenWord = True Or IsOpenExcel(1) = True Or IsOpenExcel(2) = True Or IsOpenExcel(3) = True Then
            If IsOpenWord = True Then
                strFileN = txt1(11).Text & "_" & txt1(0) & "_" & strSrvDate(1) & ".doc"
                fn_PutEnd strFileN
                strFileN = App.path & "\" & strFileN
            End If
            DoEvents
            
            'Modify By Sindy 2018/1/11
'            MAPISession1.LogonUI = False
'            MAPISession1.UserName = strUserNum
'            MAPISession1.SignOn
'            MAPIMessages1.SessionID = MAPISession1.SessionID
'            MAPIMessages1.MsgIndex = -1
'            MAPIMessages1.Compose
'            MAPIMessages1.MsgSubject = txt1(0) & "客戶案件預算預估表--電子檔"
'            MAPIMessages1.MsgNoteText = "Dear All：" & vbCrLf & "客戶案件預算預估表" & vbCrLf & "資料如附件！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
'            MAPIMessages1.AttachmentIndex = 0
'            MAPIMessages1.AttachmentPosition = 0
'            MAPIMessages1.AttachmentPathName = App.path & "\" & txt1(11).Text & "_" & txt1(0) & "_" & strSrvDate(1) & ".doc"
'            MAPIMessages1.RecipIndex = 0
'            MAPIMessages1.RecipDisplayName = IIf(Trim(txtMail.Text) = "", txt1(11), txtMail.Text)
'            MAPIMessages1.ResolveName
'            MAPIMessages1.Send
'            MAPISession1.SignOff
            PUB_SendMail strUserNum, IIf(Trim(txtMail.Text) = "", txt1(11), txtMail.Text), "", txt1(0) & "客戶案件預算預估表--電子檔", "Dear All：" & vbCrLf & "客戶案件預算預估表" & vbCrLf & "資料如附件！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", , strFileN, False
            '2018/1/11 END
            
            'Kill App.path & "\" & txt1(11).Text & "_" & txt1(0) & "_" & strSrvDate(1) & ".doc"
            If IsHaveData = True And Option1(0).Value = True Then
                Kill strFileN
            'Excel
            Else
                PUB_KillTempFile "預算表*.xls"  '刪預算表Excel
            End If
         End If
         'end 2022/05/02
         strMsg = "" 'Add by Amy 2022/07/19
         If IsHaveData = True Then
             If Option1(0).Value = True Then
                MsgBox "輸出完成!!" & vbCrLf & " 共 " & Page & " 頁！", , "輸出成功"
             Else
                'Modify by Amy 2022/07/19 原:檔案已寄出
                For Each oChk In Check1
                    If oChk.Value = 1 Then
                        If oChk.Index = 0 Then
                            strMsg = "專利 共" & intCntP & "筆" & vbCrLf & _
                                           "商標 共" & intCntT & "筆" & vbCrLf & _
                                           "其他 共" & intCntO & "筆"
                            Exit For
                        ElseIf oChk.Index = 1 Then
                            strMsg = strMsg & "專利 共" & intCntP & "筆"
                        ElseIf oChk.Index = 2 Then
                             strMsg = strMsg & "商標 共" & intCntT & "筆"
                        Else
                            strMsg = strMsg & "其他 共" & intCntO & "筆"
                        End If
                        If strMsg <> MsgText(601) Then strMsg = strMsg & vbCrLf
                    End If
                Next
                MsgBox "輸出完成!!" & vbCrLf & strMsg, , "輸出成功"
                'end 2022/07/19
             End If
         End If
         Me.Enabled = True
      'End If
   End If
   Screen.MousePointer = vbDefault
Case 1 '結束
   Unload Me
Case Else
End Select

Exit Sub

ErrorHandler:
   Select Case Err.Number
   Case 380
      MsgBox "印表機選擇錯誤!!!"
   Case Else
      MsgBox "(" & Err.Number & ")" & Err.Description
   End Select
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(2) = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
   
   Me.lblSalesName.Caption = ""
   '2014/11/27 add by sonia
   txtMail = strUserNum
   txt1(16) = strSrvDate(2)
   txt1(8) = "Y"
   '2014/11/27 end
   Call Option1_Click(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm210138 = Nothing
End Sub

'Add by Amy 2022/04/29
Private Sub Option1_Click(Index As Integer)
    Check4.Value = 0
    'Word
    If Index = 0 Then
        Check4.Enabled = False
    'Excel
    ElseIf Index = 1 Then
        If Check3(Index).Value = vbChecked Then
            Check4.Enabled = True
        Else
           Check4.Enabled = False
        End If
    End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    CloseIme
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '客戶編號起
    If Me.txt1(0).Text <> "" Then
        Me.txt1(0).Text = Left(Me.txt1(0).Text & String(9, "0"), 9)
        Me.txt1(1).Text = Left(Me.txt1(0).Text, 8) & "Z"
    End If
Case 1 '客戶編號迄
    If Me.txt1(1).Text <> "" Then
        Me.txt1(1).Text = Left(Me.txt1(1).Text & String(9, "0"), 9)
        Me.txt1(1).Text = Left(Me.txt1(1).Text, 8) & "Z"
    End If
Case 2 '系統類別
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(2)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(2).SetFocus
            txt1(2).SelStart = 0
            txt1(2).SelLength = Len(txt1(2))
            Exit Sub
        End If
     Next i
Case 7 '是否含核駁
     Select Case Trim(txt1(7))
     Case "N", "n", " ", ""
     Case Else
          s = MsgBox("是否含核駁輸入錯誤,只能 N 或空白!!", , "USER 輸入錯誤")
          txt1(7).SetFocus
          txt1(7).SelStart = 0
          txt1(7).SelLength = Len(txt1(7))
          Exit Sub
     End Select
Case 8 '國內外是否分開列印
     Select Case Trim(txt1(8))
     Case "Y", "y", " ", ""
     Case Else
          s = MsgBox("國內外是否分開列印只能 Y 或空白!!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          txt1(8).SelStart = 0
          txt1(8).SelLength = Len(txt1(8))
          Exit Sub
     End Select
Case 9 '輸出順序
     Select Case Trim(txt1(9))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("列印順序輸入錯誤,只能 1, 2, 3, 4 或 5 !!", , "USER 輸入錯誤")
          txt1(9).SetFocus
          txt1(9).SelStart = 0
          txt1(9).SelLength = Len(txt1(9))
          Exit Sub
     End Select
Case 1, 6 '客戶編號, 申請國家
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
'Case 3, 4 '收文日期
'    If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
'        Me.txt1(Index).SetFocus
'        txt1_GotFocus Index
'        Exit Sub
'    End If
'    If Index = 4 Then
'        If RunNick(txt1(Index - 1), txt1(Index)) Then
'            txt1(Index - 1).SetFocus
'            txt1_GotFocus (Index - 1)
'        Exit Sub
'        End If
'    End If
Case 13, 14 '申請日期
    If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
        Me.txt1(Index).SetFocus
        txt1_GotFocus Index
        Exit Sub
    End If
    If Index = 14 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
        Exit Sub
        End If
    End If
Case 16, 17 '本所期限
    If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
        Me.txt1(Index).SetFocus
        txt1_GotFocus Index
        Exit Sub
    End If
    If Index = 17 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
        Exit Sub
        End If
    End If
Case Else
End Select
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 9 '列印順序
        If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub RefreshColData(Index As Integer, strKind As String)
Select Case Index
Case 1 '中文報表
   If strKind = "申請人" Then
      m_ColCustName = "CU01||CU02||'   '||Nvl(CU04,Nvl(CU05,CU06)),Decode(CU04,Null,Decode(CU05,Null,CU88,''),''),Decode(CU04,Null,Decode(CU05,Null,CU89,''),''),Decode(CU04,Null,Decode(CU05,Null,CU90,''),'')"
      '聯絡地址-->POX-->申請地址
      m_ColCustAdd = "Nvl(CU31,Nvl(CU23,Nvl(CU65,Nvl(CU24,CU29))))," & _
         "Decode(CU31,Null,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),''),'')," & _
         "Decode(CU31,Null,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),''),'')," & _
         "Decode(CU31,Null,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),''),'')," & _
         "Decode(CU31,Null,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),''),'')"
   Else '代理人
      m_ColAgName = "FA01||FA02||'   '||Nvl(FA04,Nvl(FA05,FA06)),Decode(FA04,Null,Decode(FA05,Null,'',FA63),''),Decode(FA04,Null,Decode(FA05,Null,'',FA64),''),Decode(FA04,Null,Decode(FA05,Null,'',FA65),'')"
      m_ColAgAdd = "Nvl(FA17,Nvl(FA32,Nvl(FA18,FA23))),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA22),FA36),'')"
   End If
Case 2 '英文報表
   If strKind = "申請人" Then
      m_ColCustName = "CU01||CU02||'   '||Nvl(CU05,CU06),Decode(CU05,Null,'',CU88),Decode(CU05,Null,'',CU89),Decode(CU05,Null,'',CU90)"
      m_ColCustAdd = "Nvl(CU65,Nvl(CU24,CU29)),Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66)," & _
         "Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67)," & _
         "Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68)," & _
         "Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),Decode(CU65,Null,Decode(CU24,Null,'',CU102),'')"
   Else '代理人
      m_ColAgName = "FA01||FA02||'   '||Nvl(FA05,FA06),Decode(FA05,Null,'',FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)"
      m_ColAgAdd = "Nvl(FA32,Nvl(FA18,FA23)),Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),Decode(FA32,Null,Decode(FA18,Null,'',FA22),FA36)"
   End If
Case 3 '日文報表
   If strKind = "申請人" Then
      m_ColCustName = "CU01||CU02||'   '||Nvl(CU06,CU05),Decode(CU06,Null,CU88,''),Decode(CU06,Null,CU89,''),Decode(CU06,Null,CU90,'')"
      m_ColCustAdd = "Nvl(CU29,Nvl(CU65,CU24))," & _
         "Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),'')," & _
         "Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),'')," & _
         "Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),'')," & _
         "Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),'')"
   Else '代理人
      m_ColAgName = "FA01||FA02||'   '||Nvl(FA06,FA05),Decode(FA06,Null,FA63,''),Decode(FA06,Null,FA64,''),Decode(FA06,Null,FA65,'')"
      m_ColAgAdd = "Nvl(FA23,Nvl(FA32,FA18)),Decode(FA23,Null,Decode(FA32,Null,FA19,FA33),''),Decode(FA23,Null,Decode(FA32,Null,FA20,FA34),''),Decode(FA23,Null,Decode(FA32,Null,FA21,FA35),''),Decode(FA23,Null,Decode(FA32,Null,FA22,FA36),'')"
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case 11 '智權人員
        Me.lblSalesName.Caption = GetStaffName(Me.txt1(11).Text, True)
        If Me.txt1(11).Text <> "" And Me.lblSalesName.Caption = "" Then
            MsgBox "智權人員編號輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(11).SetFocus
            txt1_GotFocus 11
            Cancel = True
        End If
'2014/11/27 CANCEL BY SONIA 先寄給操作者看是否需專業部填報價
'        If Me.txt1(11).Text <> "" Then
'            If ChkStaffST04(txt1(11).Text, False) = False Then
'               txtMail = txt1(11).Text
'            Else
'               txtMail = ""
'            End If
'        End If
'2014/11/27 END
    Case Else
    End Select
End Sub

'只抓商標基本檔的資料
Sub ProcessT()
Dim strCharge As String, intCnt As Integer, strNP02 As String
Dim blnMatchKind As Boolean
Dim strCF03 As String
Dim intRow As Integer 'Add By Sindy 2012/8/28
Dim strNP07 As String 'Add By Sindy 2013/10/23
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'   '紙本
'   If Check3(0).Value = vbChecked Then
'      Printer.Orientation = 2
'      Printer.PaperSize = 9
'      o_PaperHeight = 10000
'   ElseIf Check3(1).Value = vbChecked Then
'      o_PaperHeight = 40
'   End If
   'Word
   If intChoose <> 2 Then
        o_PaperHeight = 40
   End If
   'end 2022/05/02
   blnMatchKind = False
   cnnConnection.Execute "delete from R050317_C WHERE ID='" & strUserNum & "' "
   strSQL2 = " AND NP08>=" & DBDATE(txt1(16)) & " AND NP08<=" & DBDATE(txt1(17)) & " "
   'Add By Sindy 2012/12/20
   '案件性質
   If Len(Trim(txt1(3))) <> 0 And Len(Trim(txt1(4))) <> 0 Then
      strSQL2 = strSQL2 & " AND NP07>=" & txt1(3) & " AND NP07<=" & txt1(4) & " " 'Modify By Sindy 2013/10/23
   End If
   '2012/12/20 End
   m_strSQL2_2 = strSQL2: m_strSQL2_3 = strSQL2: m_strSQL2_4 = strSQL2: m_strSQL2_5 = strSQL2
   strSQL2 = strSQL2 + " AND (TM23>='" & GetNewFagent(txt1(0)) & "' AND TM23<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL2_2 = m_strSQL2_2 + " AND (TM78>='" & GetNewFagent(txt1(0)) & "' AND TM78<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL2_3 = m_strSQL2_3 + " AND (TM79>='" & GetNewFagent(txt1(0)) & "' AND TM79<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL2_4 = m_strSQL2_4 + " AND (TM80>='" & GetNewFagent(txt1(0)) & "' AND TM80<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL2_5 = m_strSQL2_5 + " AND (TM81>='" & GetNewFagent(txt1(0)) & "' AND TM81<='" & GetNewFagent(txt1(1)) & "') "
   '若有勾選系統類別
   If Me.Check1(0).Value = vbChecked Or Me.Check1(2).Value = vbChecked Then
      blnMatchKind = True
      strNP02 = "' ',"
      If Me.Check1(0).Value = vbChecked Then
         strNP02 = strNP02 & SQLGrpStr(txt1(2), 2) & ","
      End If
      If Me.Check1(2).Value = vbChecked Then
         strNP02 = strNP02 & "'CFT','FCT','T','TF'" & ","
      End If
      strNP02 = Left(strNP02, Len(strNP02) - 1)
   End If
   If blnMatchKind = False Then Exit Sub
   
   '申請國家
   If Len(txt1(5)) <> 0 Then
       strSQL2 = strSQL2 + " AND TM10>='" & txt1(5) & "' "
       m_strSQL2_2 = m_strSQL2_2 + " AND TM10>='" & txt1(5) & "' "
       m_strSQL2_3 = m_strSQL2_3 + " AND TM10>='" & txt1(5) & "' "
       m_strSQL2_4 = m_strSQL2_4 + " AND TM10>='" & txt1(5) & "' "
       m_strSQL2_5 = m_strSQL2_5 + " AND TM10>='" & txt1(5) & "' "
   End If
   If Len(txt1(6)) <> 0 Then
       strSQL2 = strSQL2 + " AND TM10<='" & txt1(6) & "' "
       m_strSQL2_2 = m_strSQL2_2 + " AND TM10<='" & txt1(6) & "' "
       m_strSQL2_3 = m_strSQL2_3 + " AND TM10<='" & txt1(6) & "' "
       m_strSQL2_4 = m_strSQL2_4 + " AND TM10<='" & txt1(6) & "' "
       m_strSQL2_5 = m_strSQL2_5 + " AND TM10<='" & txt1(6) & "' "
   End If
   
   'Modify By Sindy 2013/10/23
'   '智權人員
'   If Me.txt1(11).Text <> "" Then
'       strSQL2 = strSQL2 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL2_2 = m_strSQL2_2 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL2_3 = m_strSQL2_3 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL2_4 = m_strSQL2_4 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL2_5 = m_strSQL2_5 + " AND NP10='" & Me.txt1(11).Text & "' "
'   '若未輸入智權人員
'   Else
'       strSQL2 = strSQL2 + " AND NP10 Is Null "
'       m_strSQL2_2 = m_strSQL2_2 + " AND NP10 Is Null "
'       m_strSQL2_3 = m_strSQL2_3 + " AND NP10 Is Null "
'       m_strSQL2_4 = m_strSQL2_4 + " AND NP10 Is Null "
'       m_strSQL2_5 = m_strSQL2_5 + " AND NP10 Is Null "
'   End If
   'Modify By Sindy 2015/9/23 改抓客戶業務員
   '智權人員
   If Me.txt1(11).Text <> "" Then
       strSQL2 = strSQL2 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL2_2 = m_strSQL2_2 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL2_3 = m_strSQL2_3 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL2_4 = m_strSQL2_4 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL2_5 = m_strSQL2_5 + " AND CU13='" & Me.txt1(11).Text & "' "
   End If
   '2015/9/23 END
   
   CheckOC
   'Modify By Sindy 2013/10/23
   '當無輸入案件性質或案件性質為領證
   If (Len(Trim(txt1(3))) = 0 Or (Trim(txt1(3)) >= "701" And Trim(txt1(4)) <= "701")) Then
      'modify by sonia 2017/11/20 增加大陸商標註冊證(核准後約8個月自動發證)np07='1701'
      strNP07 = "and ((" & Right(Trim(strNpSqlOfNoSalesDuty), Len(Trim(strNpSqlOfNoSalesDuty)) - 3) & ") or (CP10='101' AND NP07='305') or np07='1701')"
   Else
      strNP07 = "and (" & Right(Trim(strNpSqlOfNoSalesDuty), Len(Trim(strNpSqlOfNoSalesDuty)) - 3) & ")"
   End If
   '2013/10/23 END
                      strSql = " SELECT tm23,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(NA01,tm10),NP07,' '," & SQLDate("np08") & ",tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05," & SQLDate("CP57") & "||Decode(tm29,'Y','*',''),cp09, TM01, TM02, TM03, TM04," & SQLDate("TM11") & "," & SQLDate("CP27") & ", CP57,TM17 FROM trademark,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & " AND NP06 is null AND NP01=CP09(+) AND TM01=NP02(+) AND TM02=NP03(+) AND TM03=NP04(+) AND TM04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) AND (TM29<>'Y' or tm29 is null) " & strSQL2
   strSql = strSql & " union all SELECT tm78,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm78),tm12,NVL(NA01,tm10),NP07,' '," & SQLDate("np08") & ",tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05," & SQLDate("CP57") & "||Decode(tm29,'Y','*',''),cp09, TM01, TM02, TM03, TM04," & SQLDate("TM11") & "," & SQLDate("CP27") & ", CP57,TM17 FROM trademark,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & " AND NP06 is null AND NP01=CP09(+) AND TM01=NP02(+) AND TM02=NP03(+) AND TM03=NP04(+) AND TM04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(tm78,1,8)=CU01(+) AND DECODE(SUBSTR(tm78,9,1),NULL,'0',SUBSTR(tm78,9,1))=CU02(+) AND tm10=NA01(+) AND (TM29<>'Y' or tm29 is null) " & m_strSQL2_2
   strSql = strSql & " union all SELECT tm79,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm79),tm12,NVL(NA01,tm10),NP07,' '," & SQLDate("np08") & ",tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05," & SQLDate("CP57") & "||Decode(tm29,'Y','*',''),cp09, TM01, TM02, TM03, TM04," & SQLDate("TM11") & "," & SQLDate("CP27") & ", CP57,TM17 FROM trademark,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & " AND NP06 is null AND NP01=CP09(+) AND TM01=NP02(+) AND TM02=NP03(+) AND TM03=NP04(+) AND TM04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(tm79,1,8)=CU01(+) AND DECODE(SUBSTR(tm79,9,1),NULL,'0',SUBSTR(tm79,9,1))=CU02(+) AND tm10=NA01(+) AND (TM29<>'Y' or tm29 is null) " & m_strSQL2_3
   strSql = strSql & " union all SELECT tm80,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm80),tm12,NVL(NA01,tm10),NP07,' '," & SQLDate("np08") & ",tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05," & SQLDate("CP57") & "||Decode(tm29,'Y','*',''),cp09, TM01, TM02, TM03, TM04," & SQLDate("TM11") & "," & SQLDate("CP27") & ", CP57,TM17 FROM trademark,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & " AND NP06 is null AND NP01=CP09(+) AND TM01=NP02(+) AND TM02=NP03(+) AND TM03=NP04(+) AND TM04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(tm80,1,8)=CU01(+) AND DECODE(SUBSTR(tm80,9,1),NULL,'0',SUBSTR(tm80,9,1))=CU02(+) AND tm10=NA01(+) AND (TM29<>'Y' or tm29 is null) " & m_strSQL2_4
   strSql = strSql & " union all SELECT tm81,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm81),tm12,NVL(NA01,tm10),NP07,' '," & SQLDate("np08") & ",tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05," & SQLDate("CP57") & "||Decode(tm29,'Y','*',''),cp09, TM01, TM02, TM03, TM04," & SQLDate("TM11") & "," & SQLDate("CP27") & ", CP57,TM17 FROM trademark,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & " AND NP06 is null AND NP01=CP09(+) AND TM01=NP02(+) AND TM02=NP03(+) AND TM03=NP04(+) AND TM04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(tm81,1,8)=CU01(+) AND DECODE(SUBSTR(tm81,9,1),NULL,'0',SUBSTR(tm81,9,1))=CU02(+) AND tm10=NA01(+) AND (TM29<>'Y' or tm29 is null) " & m_strSQL2_5
   adoRecordset.CursorLocation = adUseClient
   k = 0
   intRow = 0 'Add By Sindy 2012/8/28
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount)
      With adoRecordset
         .MoveFirst
         DoEvents
         Do While .EOF = False
            'Modify By Sindy 2012/8/27 FCT,T,TF,CFT延展(102)和第二期(716)專用權須存在(TM17=Y)
            If (.Fields("TM01") = "FCT" Or .Fields("TM01") = "T" Or .Fields("TM01") = "TF" Or .Fields("TM01") = "CFT") And _
               (.Fields("NP07") = "102" Or .Fields("NP07") = "716") And _
               "" & .Fields("TM17") <> "Y" Then
               GoTo ReadNext_T
            End If
            intRow = intRow + 1
            '2012/8/27 End
            For i = LBound(strTemp) To UBound(strTemp)
                strTemp(i) = ""
            Next i
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Right(strTemp(16), 1) = "*" Then
                strTemp(9) = "*" + strTemp(9)
                strTemp(16) = Replace(strTemp(16), "*", "閉卷")
            End If
            '申請日
            strTemp(23) = "" & .Fields(22).Value
            '發文日
            strTemp(24) = "" & .Fields(23).Value
            CheckOC2
            strSql = "select min(pd05) from pridate where pd01='" & CheckStr(.Fields(18)) & "' and pd02='" & CheckStr(.Fields(19)) & "' and pd03='" & CheckStr(.Fields(20)) & "' and pd04='" & CheckStr(.Fields(21)) & "' "
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
                 strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
            Else
                 strTemp(6) = ""
            End If
            
            '預估費用:X28279000
            strCharge = "": intCnt = 0
            '102.延展
            If strTemp(5) = "102" Then
               'modify by sonia 2022/9/6 配合台灣案抓商品類別數，cf08+(cf13*1000) 改cf08,cf13，下面再計算
               strSql = "SELECT cf08,cf13 FROM casefee where cf01='" & .Fields(18).Value & "' and cf02='" & strTemp(4) & "' and cf03='" & strTemp(5) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  'modify by sonia 2022/9/6
                  'strCharge = "" & RsTemp.Fields(0)
                  strCharge = Val("" & RsTemp.Fields(0)) + (Val("" & RsTemp.Fields(1)) * 1000)
               End If
               'add by sonia 2022/9/5 台灣案抓商品類別數
               If strTemp(4) = "000" Then
                  intCnt = GetTMKindCnt(.Fields(18).Value, .Fields(19).Value, .Fields(20).Value, .Fields(21).Value)
                     'CF08規費要*商品類別數
                     If Val("" & RsTemp.Fields(0)) > 0 Then
                        strCharge = RsTemp.Fields(0) * intCnt
                     End If
                     'CF13標準價*1000(點數換算成金額)
                     If Val("" & RsTemp.Fields(1)) > 0 Then
                        strCharge = Val(strCharge) + (RsTemp.Fields(1) * 1000)
                     End If
               End If
               'end 2022/9/5
            
            '717.註冊費(新申請案催審),716第二期註冊費
            ElseIf strTemp(4) = "000" And (strTemp(5) = "717" Or strTemp(5) = "305" Or strTemp(5) = "716") Then
               If strTemp(5) = "305" Then
                  strCF03 = "717"
               Else
                  strCF03 = strTemp(5)
               End If
               '台灣商標註冊費:
               '抓商品類別數
               intCnt = GetTMKindCnt(.Fields(18).Value, .Fields(19).Value, .Fields(20).Value, .Fields(21).Value)
               strSql = "select cf08,cf13 from casefee where cf01='" & .Fields(18).Value & "' and cf02='" & strTemp(4) & "' and cf03='" & strCF03 & "' order by cf03 "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  'CF08規費要*商品類別數
                  If Val("" & RsTemp.Fields(0)) > 0 Then
                     strCharge = Val(strCharge) + (RsTemp.Fields(0) * intCnt)
                  End If
                  'CF13標準價*1000(點數換算成金額)
                  If Val("" & RsTemp.Fields(1)) > 0 Then
                     strCharge = Val(strCharge) + (RsTemp.Fields(1) * 1000)
                  End If
               End If
            'add by sonia 2017/11/20 大陸商標註冊證及申請案催審預估費用3000元
            ElseIf strTemp(4) = "020" And (strTemp(5) = "1701" Or strTemp(5) = "305") Then
               strCharge = 3000
               If Me.txt1(11).Text = "69010" Then
                  strCharge = "5000"
               ElseIf Me.txt1(11).Text = "76051" Then
                  strCharge = "6000"
               End If
            'end 2017/11/20
            End If
            
            strSql = "INSERT INTO R050317_C " & strTBF & " values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & ChgSQL(strTemp(18)) & "','" & strUserNum & "','" & strTemp(23) & "','" & strTemp(24) & "','" & .Fields(18).Value & "','" & .Fields(19).Value & "','" & .Fields(20).Value & "','" & .Fields(21).Value & "','" & "" & .Fields("CP57").Value & "'," & CNULL(strCharge) & ")"
            cnnConnection.Execute strSql
            IsHaveData = True
            k = k + 1
            DoEvents
ReadNext_T:
            .MoveNext
         Loop
      End With
      'Add By Sindy 2012/8/28
      If intRow = 0 Then
         InsertQueryLog (0)
         ShowNoData
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      '2012/8/28 End
   Else
      InsertQueryLog (0)
      ShowNoData
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   CheckOC
   
   PrintDataCt_A4
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'除專利及商標基本檔其他的相關資料
Private Sub ProcessO()
Dim strCharge As String, strNP02 As String
Dim blnMatchKind As Boolean
Dim stCon1 As String 'Added by Lydia 2018/10/12

On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
   '紙本
'   If Check3(0).Value = vbChecked Then
'      Printer.Orientation = 2
'      Printer.PaperSize = 9
'      o_PaperHeight = 10000
'   ElseIf Check3(1).Value = vbChecked Then
'      o_PaperHeight = 40
'   End If
   'Word
   If intChoose <> 2 Then
        o_PaperHeight = 40
   End If
   
   blnMatchKind = False
   cnnConnection.Execute "delete from R050317_C WHERE ID='" & strUserNum & "' "
   '初始化變數
   StrSQL3 = " AND NP08>=" & DBDATE(txt1(16)) & " AND NP08<=" & DBDATE(txt1(17)) & " "
   stCon1 = " and cp158=0 and cp159=0 AND cp54>=" & DBDATE(txt1(16)) & " AND cp54<=" & DBDATE(txt1(17)) & " " 'Added by Lydia 2018/10/12
   
   'Add By Sindy 2012/12/20
   '案件性質
   If Len(Trim(txt1(3))) <> 0 And Len(Trim(txt1(4))) <> 0 Then
      StrSQL3 = StrSQL3 & " AND NP07>=" & txt1(3) & " AND NP07<=" & txt1(4) & " " 'Modify By Sindy 2013/10/23
      stCon1 = stCon1 & " and cp10>=" & txt1(3) & " AND cp10<=" & txt1(4) & " " 'Added by Lydia 2018/10/12
   End If
   '2012/12/20 End
   m_strSQL3_2 = StrSQL3: m_strSQL3_3 = StrSQL3: m_strSQL3_4 = StrSQL3: m_strSQL3_5 = StrSQL3
   StrSQL4 = StrSQL3
   m_strSQL4_2 = StrSQL3: m_strSQL4_3 = StrSQL3: m_strSQL4_4 = StrSQL3: m_strSQL4_5 = StrSQL3
   strSQL5 = StrSQL3
   m_strSQL5_2 = StrSQL3: m_strSQL5_3 = StrSQL3: m_strSQL5_4 = StrSQL3: m_strSQL5_5 = StrSQL3
   StrSQL3 = StrSQL3 + " AND (LC11>='" & GetNewFagent(txt1(0)) & "' AND LC11<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL3_2 = m_strSQL3_2 + " AND (LC43>='" & GetNewFagent(txt1(0)) & "' AND LC43<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL3_3 = m_strSQL3_3 + " AND (LC44>='" & GetNewFagent(txt1(0)) & "' AND LC44<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL3_4 = m_strSQL3_4 + " AND (LC45>='" & GetNewFagent(txt1(0)) & "' AND LC45<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL3_5 = m_strSQL3_5 + " AND (LC46>='" & GetNewFagent(txt1(0)) & "' AND LC46<='" & GetNewFagent(txt1(1)) & "') "
   StrSQL4 = StrSQL4 + " AND (HC05>='" & GetNewFagent(txt1(0)) & "' AND HC05<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL4_2 = m_strSQL4_2 + " AND (HC24>='" & GetNewFagent(txt1(0)) & "' AND HC24<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL4_3 = m_strSQL4_3 + " AND (HC25>='" & GetNewFagent(txt1(0)) & "' AND HC25<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL4_4 = m_strSQL4_4 + " AND (HC26>='" & GetNewFagent(txt1(0)) & "' AND HC26<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL4_5 = m_strSQL4_5 + " AND (HC27>='" & GetNewFagent(txt1(0)) & "' AND HC27<='" & GetNewFagent(txt1(1)) & "') "
   strSQL5 = strSQL5 + " AND (SP08>='" & GetNewFagent(txt1(0)) & "' AND SP08<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL5_2 = m_strSQL5_2 + " AND (SP58>='" & GetNewFagent(txt1(0)) & "' AND SP58<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL5_3 = m_strSQL5_3 + " AND (SP59>='" & GetNewFagent(txt1(0)) & "' AND SP59<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL5_4 = m_strSQL5_4 + " AND (SP65>='" & GetNewFagent(txt1(0)) & "' AND SP65<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL5_5 = m_strSQL5_5 + " AND (SP66>='" & GetNewFagent(txt1(0)) & "' AND SP66<='" & GetNewFagent(txt1(1)) & "') "
   'Added by Lydia 2018/10/12
   stCon1 = stCon1 & " AND ((HC05>='" & GetNewFagent(txt1(0)) & "' AND HC05<='" & GetNewFagent(txt1(1)) & "') " & _
                "or (HC24>='" & GetNewFagent(txt1(0)) & "' AND HC24<='" & GetNewFagent(txt1(1)) & "') " & _
                "or (HC25>='" & GetNewFagent(txt1(0)) & "' AND HC25<='" & GetNewFagent(txt1(1)) & "') " & _
                "or (HC26>='" & GetNewFagent(txt1(0)) & "' AND HC26<='" & GetNewFagent(txt1(1)) & "') " & _
                "or (HC27>='" & GetNewFagent(txt1(0)) & "' AND HC27<='" & GetNewFagent(txt1(1)) & "')) "
   'end 2018/10/12
   '若有勾選系統類別
   If Me.Check1(0).Value = vbChecked Or Me.Check1(3).Value = vbChecked Or Me.Check1(4).Value = vbChecked Or Me.Check1(5).Value = vbChecked Or Me.Check1(6).Value = vbChecked Then
      blnMatchKind = True
      strNP02 = "' ',"
      If Me.Check1(0).Value = vbChecked Then
         strNP02 = strNP02 & SQLGrpStr(txt1(2), 3) & ","
         strNP02 = strNP02 & SQLGrpStr(txt1(2), 4) & ","
         strNP02 = strNP02 & SQLGrpStr(txt1(2), 5) & ","
      End If
      If Me.Check1(3).Value = vbChecked Then
         strNP02 = strNP02 & "'TC','CFC'" & ","
      End If
      If Me.Check1(4).Value = vbChecked Then
         strNP02 = strNP02 & "'L','CFL','FCL','LA','LIN','ACS'" & ","
      End If
      If Me.Check1(5).Value = vbChecked Then
         strNP02 = strNP02 & "'TB'" & ","
      End If
      If Me.Check1(6).Value = vbChecked Then
         strNP02 = strNP02 & "'TD'" & ","
      End If
      strNP02 = Left(strNP02, Len(strNP02) - 1)
   End If
   If blnMatchKind = False Then Exit Sub
   
   '申請國家
   If Len(txt1(5)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(5) & "' "
       m_strSQL3_2 = m_strSQL3_2 + " AND LC15>='" & txt1(5) & "' "
       m_strSQL3_3 = m_strSQL3_3 + " AND LC15>='" & txt1(5) & "' "
       m_strSQL3_4 = m_strSQL3_4 + " AND LC15>='" & txt1(5) & "' "
       m_strSQL3_5 = m_strSQL3_5 + " AND LC15>='" & txt1(5) & "' "
       strSQL5 = strSQL5 + " AND SP09>='" & txt1(5) & "' "
       m_strSQL5_2 = m_strSQL5_2 + " AND SP09>='" & txt1(5) & "' "
       m_strSQL5_3 = m_strSQL5_3 + " AND SP09>='" & txt1(5) & "' "
       m_strSQL5_4 = m_strSQL5_4 + " AND SP09>='" & txt1(5) & "' "
       m_strSQL5_5 = m_strSQL5_5 + " AND SP09>='" & txt1(5) & "' "
   End If
   If Len(txt1(6)) <> 0 Then
       StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(6) & "' "
       m_strSQL3_2 = m_strSQL3_2 + " AND LC15<='" & txt1(6) & "' "
       m_strSQL3_3 = m_strSQL3_3 + " AND LC15<='" & txt1(6) & "' "
       m_strSQL3_4 = m_strSQL3_4 + " AND LC15<='" & txt1(6) & "' "
       m_strSQL3_5 = m_strSQL3_5 + " AND LC15<='" & txt1(6) & "' "
       strSQL5 = strSQL5 + " AND SP09<='" & txt1(6) & "' "
       m_strSQL5_2 = m_strSQL5_2 + " AND SP09<='" & txt1(6) & "' "
       m_strSQL5_3 = m_strSQL5_3 + " AND SP09<='" & txt1(6) & "' "
       m_strSQL5_4 = m_strSQL5_4 + " AND SP09<='" & txt1(6) & "' "
       m_strSQL5_5 = m_strSQL5_5 + " AND SP09<='" & txt1(6) & "' "
   End If
   
   'Modify By Sindy 2013/10/23
'   '智權人員
'   If Me.txt1(11).Text <> "" Then
'       StrSQL3 = StrSQL3 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL3_2 = m_strSQL3_2 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL3_3 = m_strSQL3_3 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL3_4 = m_strSQL3_4 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL3_5 = m_strSQL3_5 + " AND NP10='" & Me.txt1(11).Text & "' "
'       StrSQL4 = StrSQL4 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL4_2 = m_strSQL4_2 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL4_3 = m_strSQL4_3 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL4_4 = m_strSQL4_4 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL4_5 = m_strSQL4_5 + " AND NP10='" & Me.txt1(11).Text & "' "
'       strSQL5 = strSQL5 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL5_2 = m_strSQL5_2 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL5_3 = m_strSQL5_3 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL5_4 = m_strSQL5_4 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL5_5 = m_strSQL5_5 + " AND NP10='" & Me.txt1(11).Text & "' "
'   '若未輸入智權人員
'   Else
'       StrSQL3 = StrSQL3 + " AND NP10 Is Null "
'       m_strSQL3_2 = m_strSQL3_2 + " AND NP10 Is Null "
'       m_strSQL3_3 = m_strSQL3_3 + " AND NP10 Is Null "
'       m_strSQL3_4 = m_strSQL3_4 + " AND NP10 Is Null "
'       m_strSQL3_5 = m_strSQL3_5 + " AND NP10 Is Null "
'       StrSQL4 = StrSQL4 + " AND NP10 Is Null "
'       m_strSQL4_2 = m_strSQL4_2 + " AND NP10 Is Null "
'       m_strSQL4_3 = m_strSQL4_3 + " AND NP10 Is Null "
'       m_strSQL4_4 = m_strSQL4_4 + " AND NP10 Is Null "
'       m_strSQL4_5 = m_strSQL4_5 + " AND NP10 Is Null "
'       strSQL5 = strSQL5 + " AND NP10 Is Null "
'       m_strSQL5_2 = m_strSQL5_2 + " AND NP10 Is Null "
'       m_strSQL5_3 = m_strSQL5_3 + " AND NP10 Is Null "
'       m_strSQL5_4 = m_strSQL5_4 + " AND NP10 Is Null "
'       m_strSQL5_5 = m_strSQL5_5 + " AND NP10 Is Null "
'   End If
   'Modify By Sindy 2015/9/23 改抓客戶業務員
   '智權人員
   If Me.txt1(11).Text <> "" Then
       StrSQL3 = StrSQL3 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL3_2 = m_strSQL3_2 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL3_3 = m_strSQL3_3 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL3_4 = m_strSQL3_4 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL3_5 = m_strSQL3_5 + " AND CU13='" & Me.txt1(11).Text & "' "
       StrSQL4 = StrSQL4 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL4_2 = m_strSQL4_2 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL4_3 = m_strSQL4_3 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL4_4 = m_strSQL4_4 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL4_5 = m_strSQL4_5 + " AND CU13='" & Me.txt1(11).Text & "' "
       strSQL5 = strSQL5 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL5_2 = m_strSQL5_2 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL5_3 = m_strSQL5_3 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL5_4 = m_strSQL5_4 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL5_5 = m_strSQL5_5 + " AND CU13='" & Me.txt1(11).Text & "' "
       stCon1 = stCon1 + " AND CU13='" & Me.txt1(11).Text & "' " 'Added by Lydia 2018/10/12
   End If
   '2015/9/23 END
   
   CheckOC
                       strSql = "Select LC11,lc23,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),''  ,NVL(NA01,LC15),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),''  ,0   ,''                                   ,''                            ,Nvl(lc05,Nvl(LC06,LC07)),CP05," & SQLDate("CP57") & "||Decode(LC08,'Y','*',''),cp09, LC01, LC02, LC03, LC04, ''                    , CP57 FROM LAWCASE        ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND LC01=NP02(+) AND LC02=NP03(+) AND LC03=NP04(+) AND LC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) AND (LC08<>'Y' or lc08 is null) " & StrSQL3
   strSql = strSql + " union all Select LC43,lc23,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC43),''  ,NVL(NA01,LC15),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),''  ,0   ,''                                   ,''                            ,Nvl(lc05,Nvl(LC06,LC07)),CP05," & SQLDate("CP57") & "||Decode(LC08,'Y','*',''),cp09, LC01, LC02, LC03, LC04, ''                    , CP57 FROM LAWCASE        ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND LC01=NP02(+) AND LC02=NP03(+) AND LC03=NP04(+) AND LC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(LC43,1,8)=CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1)) = CU02(+) AND LC15=NA01(+) AND (LC08<>'Y' or lc08 is null) " & m_strSQL3_2
   strSql = strSql + " union all Select LC44,lc23,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC44),''  ,NVL(NA01,LC15),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),''  ,0   ,''                                   ,''                            ,Nvl(lc05,Nvl(LC06,LC07)),CP05," & SQLDate("CP57") & "||Decode(LC08,'Y','*',''),cp09, LC01, LC02, LC03, LC04, ''                    , CP57 FROM LAWCASE        ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND LC01=NP02(+) AND LC02=NP03(+) AND LC03=NP04(+) AND LC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(LC44,1,8)=CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1)) = CU02(+) AND LC15=NA01(+) AND (LC08<>'Y' or lc08 is null) " & m_strSQL3_3
   strSql = strSql + " union all Select LC45,lc23,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC45),''  ,NVL(NA01,LC15),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),''  ,0   ,''                                   ,''                            ,Nvl(lc05,Nvl(LC06,LC07)),CP05," & SQLDate("CP57") & "||Decode(LC08,'Y','*',''),cp09, LC01, LC02, LC03, LC04, ''                    , CP57 FROM LAWCASE        ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND LC01=NP02(+) AND LC02=NP03(+) AND LC03=NP04(+) AND LC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(LC45,1,8)=CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1)) = CU02(+) AND LC15=NA01(+) AND (LC08<>'Y' or lc08 is null) " & m_strSQL3_4
   strSql = strSql + " union all Select LC46,lc23,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC46),''  ,NVL(NA01,LC15),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),''  ,0   ,''                                   ,''                            ,Nvl(lc05,Nvl(LC06,LC07)),CP05," & SQLDate("CP57") & "||Decode(LC08,'Y','*',''),cp09, LC01, LC02, LC03, LC04, ''                    , CP57 FROM LAWCASE        ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND LC01=NP02(+) AND LC02=NP03(+) AND LC03=NP04(+) AND LC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(LC46,1,8)=CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1)) = CU02(+) AND LC15=NA01(+) AND (LC08<>'Y' or lc08 is null) " & m_strSQL3_5
   strSql = strSql + " union all select HC05,''  ,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),HC05),''  ,NA01          ,NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",''  ,HC01||'-'||HC02||'-'||HC03||'-'||HC04                            ,''  ,0   ,''                                   ,''                            ,HC06                    ,CP05," & SQLDate("CP57") & "||Decode(HC09,'Y','*',''),cp09, HC01, HC02, HC03, HC04, ''                    , CP57 FROM HIRECASE       ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND HC01=NP02(+) AND HC02=NP03(+) AND HC03=NP04(+) AND HC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1)) = CU02(+) AND '000'=NA01(+) AND (HC09<>'Y' or hc09 is null) " & StrSQL4
   strSql = strSql + " union all select HC24,''  ,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),HC24),''  ,NA01          ,NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",''  ,HC01||'-'||HC02||'-'||HC03||'-'||HC04                            ,''  ,0   ,''                                   ,''                            ,HC06                    ,CP05," & SQLDate("CP57") & "||Decode(HC09,'Y','*',''),cp09, HC01, HC02, HC03, HC04, ''                    , CP57 FROM HIRECASE       ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND HC01=NP02(+) AND HC02=NP03(+) AND HC03=NP04(+) AND HC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(HC24,1,8)=CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1)) = CU02(+) AND '000'=NA01(+) AND (HC09<>'Y' or hc09 is null) " & m_strSQL4_2
   strSql = strSql + " union all select HC25,''  ,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),HC25),''  ,NA01          ,NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",''  ,HC01||'-'||HC02||'-'||HC03||'-'||HC04                            ,''  ,0   ,''                                   ,''                            ,HC06                    ,CP05," & SQLDate("CP57") & "||Decode(HC09,'Y','*',''),cp09, HC01, HC02, HC03, HC04, ''                    , CP57 FROM HIRECASE       ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND HC01=NP02(+) AND HC02=NP03(+) AND HC03=NP04(+) AND HC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(HC25,1,8)=CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1)) = CU02(+) AND '000'=NA01(+) AND (HC09<>'Y' or hc09 is null) " & m_strSQL4_3
   strSql = strSql + " union all select HC26,''  ,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),HC26),''  ,NA01          ,NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",''  ,HC01||'-'||HC02||'-'||HC03||'-'||HC04                            ,''  ,0   ,''                                   ,''                            ,HC06                    ,CP05," & SQLDate("CP57") & "||Decode(HC09,'Y','*',''),cp09, HC01, HC02, HC03, HC04, ''                    , CP57 FROM HIRECASE       ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND HC01=NP02(+) AND HC02=NP03(+) AND HC03=NP04(+) AND HC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(HC26,1,8)=CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1)) = CU02(+) AND '000'=NA01(+) AND (HC09<>'Y' or hc09 is null) " & m_strSQL4_4
   strSql = strSql + " union all select HC27,''  ,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),HC27),''  ,NA01          ,NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",''  ,HC01||'-'||HC02||'-'||HC03||'-'||HC04                            ,''  ,0   ,''                                   ,''                            ,HC06                    ,CP05," & SQLDate("CP57") & "||Decode(HC09,'Y','*',''),cp09, HC01, HC02, HC03, HC04, ''                    , CP57 FROM HIRECASE       ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND HC01=NP02(+) AND HC02=NP03(+) AND HC03=NP04(+) AND HC04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(HC27,1,8)=CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1)) = CU02(+) AND '000'=NA01(+) AND (HC09<>'Y' or hc09 is null) " & m_strSQL4_5
   strSql = strSql + " union all select sp08,sp27,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(NA01,SP09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,''                                   ,''                            ,Nvl(sp05,Nvl(SP06,SP07)),CP05," & SQLDate("CP57") & "||Decode(SP15,'Y','*',''),cp09, SP01, SP02, SP03, SP04," & SQLDate("SP10") & ", CP57 FROM SERVICEPRACTICE,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) AND (SP15<>'Y' or sp15 is null) " & strSQL5
   strSql = strSql + " union all select sp58,sp27,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP58),sp11,NVL(NA01,SP09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,''                                   ,''                            ,Nvl(sp05,Nvl(SP06,SP07)),CP05," & SQLDate("CP57") & "||Decode(SP15,'Y','*',''),cp09, SP01, SP02, SP03, SP04," & SQLDate("SP10") & ", CP57 FROM SERVICEPRACTICE,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(SP58,1,8)=CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1)) = CU02(+) AND SP09=NA01(+) AND (SP15<>'Y' or sp15 is null) " & m_strSQL5_2
   strSql = strSql + " union all select sp59,sp27,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP59),sp11,NVL(NA01,SP09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,''                                   ,''                            ,Nvl(sp05,Nvl(SP06,SP07)),CP05," & SQLDate("CP57") & "||Decode(SP15,'Y','*',''),cp09, SP01, SP02, SP03, SP04," & SQLDate("SP10") & ", CP57 FROM SERVICEPRACTICE,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(SP59,1,8)=CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1)) = CU02(+) AND SP09=NA01(+) AND (SP15<>'Y' or sp15 is null) " & m_strSQL5_3
   strSql = strSql + " union all select sp65,sp27,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP65),sp11,NVL(NA01,SP09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,''                                   ,''                            ,Nvl(sp05,Nvl(SP06,SP07)),CP05," & SQLDate("CP57") & "||Decode(SP15,'Y','*',''),cp09, SP01, SP02, SP03, SP04," & SQLDate("SP10") & ", CP57 FROM SERVICEPRACTICE,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(SP65,1,8)=CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1)) = CU02(+) AND SP09=NA01(+) AND (SP15<>'Y' or sp15 is null) " & m_strSQL5_4
   strSql = strSql + " union all select sp66,sp27,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP66),sp11,NVL(NA01,SP09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,''                                   ,''                            ,Nvl(sp05,Nvl(SP06,SP07)),CP05," & SQLDate("CP57") & "||Decode(SP15,'Y','*',''),cp09, SP01, SP02, SP03, SP04," & SQLDate("SP10") & ", CP57 FROM SERVICEPRACTICE,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNpSqlOfNoSalesDuty & " AND NP06 is null AND NP01=CP09(+) AND SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND SUBSTR(SP66,1,8)=CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1)) = CU02(+) AND SP09=NA01(+) AND (SP15<>'Y' or sp15 is null) " & m_strSQL5_5
   'Added by Lydia 2018/10/12 針對顧問LA案件，要增加顧問聘任未續簽的報價資料
   If Me.Check1(0).Value = vbChecked Or Me.Check1(4).Value = vbChecked Then
       strSql = strSql + "union all select HC05,''  ,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),HC05),''  ,NA01          ,CP10," & SQLDate("cp27") & "," & SQLDate("cp54") & ",''  ,HC01||'-'||HC02||'-'||HC03||'-'||HC04                            ,''  ,0   ,''                                   ,''                            ,HC06                    ,CP05," & SQLDate("CP57") & "||Decode(HC09,'Y','*',''),cp09, HC01, HC02, HC03, HC04, ''                    , CP57 FROM HIRECASE       ,CASEPROGRESS A,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01||cp10='LA0' AND HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1)) = CU02(+) AND '000'=NA01(+) AND (HC09<>'Y' or hc09 is null) " & stCon1
   End If
   'end 2018/10/12
   adoRecordset.CursorLocation = adUseClient
   k = 0
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount)
      With adoRecordset
         .MoveFirst
         DoEvents
         Do While .EOF = False
            For i = LBound(strTemp) To UBound(strTemp)
               strTemp(i) = ""
            Next i
            For i = 0 To 18
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Right(strTemp(16), 1) = "*" Then
               strTemp(9) = "*" + strTemp(9)
               strTemp(16) = Replace(strTemp(16), "*", "閉卷")
            End If
            '申請日
            strTemp(23) = "" & .Fields(22).Value
            CheckOC2
            strSql = "select min(pd05) from pridate where pd01='" & CheckStr(.Fields(18)) & "' and pd02='" & CheckStr(.Fields(19)) & "' and pd03='" & CheckStr(.Fields(20)) & "' and pd04='" & CheckStr(.Fields(21)) & "' "
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strTemp(24) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
            Else
               strTemp(24) = ""
            End If
            'X29787000 87013
            strSql = "INSERT INTO R050317_C  " & strTBF & " values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & ChgSQL(strTemp(18)) & "','" & strUserNum & "','" & strTemp(23) & "','" & strTemp(24) & "','" & .Fields(18).Value & "','" & .Fields(19).Value & "','" & .Fields(20).Value & "','" & .Fields(21).Value & "','" & "" & .Fields("CP57").Value & "',null)"
            cnnConnection.Execute strSql
            IsHaveData = True
            k = k + 1
            DoEvents
            .MoveNext
         Loop
      End With
   Else
      InsertQueryLog (0)
      ShowNoData
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   CheckOC
   
   PrintDataCo_A4
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'只抓專利基本檔的資料
Private Sub ProcessP()
Dim adoRst As New ADODB.Recordset
Dim strCharge As String, strNP02 As String
Dim m_NP02 As String, m_NP03 As String, m_NP04 As String, m_NP05 As String, m_NP07 As String, m_PA08 As String, m_PA09 As String
Dim m_strYear As String, m_PA91 As String, m_Nexttimes As String, strYF03 As String
Dim strDiscCase As String '年費是否可抵減
Dim m_bFirstYear As Boolean '是否繳第一次年費
Dim strPA72NextYear As String
Dim strPA72Year As String
Dim strMaxFeeYear As String '最大可繳費年度
Dim blnMatchKind As Boolean
Dim m_PA26 As String, m_CP44 As String, m_PA10 As String, m_PA20 As String
Dim str_P020Year As String, m_NP08 As String
Dim strNP07 As String 'Add By Sindy 2013/10/23

On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
   '紙本
'   If Check3(0).Value = vbChecked Then
'      Printer.Orientation = 2
'      Printer.PaperSize = 9
'      o_PaperHeight = 10000
'   ElseIf Check3(1).Value = vbChecked Then
'      o_PaperHeight = 40
'   End If
   'Word
   If intChoose <> 2 Then
        o_PaperHeight = 40
   End If
   'end 2022/05/02
   blnMatchKind = False
   cnnConnection.Execute "delete from R050317_C WHERE ID='" & strUserNum & "' "
   strSQL1 = " AND NP08>=" & DBDATE(txt1(16)) & " AND NP08<=" & DBDATE(txt1(17)) & " "
   'Add By Sindy 2012/12/20
   '案件性質
   If Len(Trim(txt1(3))) <> 0 And Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 & " AND NP07>=" & txt1(3) & " AND NP07<=" & txt1(4) & " " 'Modify By Sindy 2013/10/23
   End If
   '2012/12/20 End
   m_strSQL1_2 = strSQL1: m_strSQL1_3 = strSQL1: m_strSQL1_4 = strSQL1: m_strSQL1_5 = strSQL1
   strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(0)) & "' AND PA26<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL1_2 = m_strSQL1_2 + " AND (PA27>='" & GetNewFagent(txt1(0)) & "' AND PA27<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL1_3 = m_strSQL1_3 + " AND (PA28>='" & GetNewFagent(txt1(0)) & "' AND PA28<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL1_4 = m_strSQL1_4 + " AND (PA29>='" & GetNewFagent(txt1(0)) & "' AND PA29<='" & GetNewFagent(txt1(1)) & "') "
   m_strSQL1_5 = m_strSQL1_5 + " AND (PA30>='" & GetNewFagent(txt1(0)) & "' AND PA30<='" & GetNewFagent(txt1(1)) & "') "
   '若有勾選系統類別
   If Me.Check1(0).Value = vbChecked Or Me.Check1(1).Value = vbChecked Then
       blnMatchKind = True
       strNP02 = "' ',"
       If Me.Check1(0).Value = vbChecked Then
           strNP02 = strNP02 & SQLGrpStr(txt1(2), 1) & ","
       End If
       If Me.Check1(1).Value = vbChecked Then
           strNP02 = strNP02 & "'CFP','FCP','P'" & ","
       End If
       strNP02 = Left(strNP02, Len(strNP02) - 1)
   End If
   If blnMatchKind = False Then Exit Sub
   
   '申請國家
   If Len(txt1(5)) <> 0 Then
       strSQL1 = strSQL1 + " AND PA09>='" & txt1(5) & "' "
       m_strSQL1_2 = m_strSQL1_2 + " AND PA09>='" & txt1(5) & "' "
       m_strSQL1_3 = m_strSQL1_3 + " AND PA09>='" & txt1(5) & "' "
       m_strSQL1_4 = m_strSQL1_4 + " AND PA09>='" & txt1(5) & "' "
       m_strSQL1_5 = m_strSQL1_5 + " AND PA09>='" & txt1(5) & "' "
   End If
   If Len(txt1(6)) <> 0 Then
       strSQL1 = strSQL1 + " AND PA09<='" & txt1(6) & "' "
       m_strSQL1_2 = m_strSQL1_2 + " AND PA09<='" & txt1(6) & "' "
       m_strSQL1_3 = m_strSQL1_3 + " AND PA09<='" & txt1(6) & "' "
       m_strSQL1_4 = m_strSQL1_4 + " AND PA09<='" & txt1(6) & "' "
       m_strSQL1_5 = m_strSQL1_5 + " AND PA09<='" & txt1(6) & "' "
   End If
   
   'Modify By Sindy 2013/10/23
'   '智權人員
'   If Me.txt1(11).Text <> "" Then
'       strSQL1 = strSQL1 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL1_2 = m_strSQL1_2 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL1_3 = m_strSQL1_3 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL1_4 = m_strSQL1_4 + " AND NP10='" & Me.txt1(11).Text & "' "
'       m_strSQL1_5 = m_strSQL1_5 + " AND NP10='" & Me.txt1(11).Text & "' "
'   '若未輸入智權人員
'   Else
'       strSQL1 = strSQL1 + " AND NP10 Is Null "
'       m_strSQL1_2 = m_strSQL1_2 + " AND NP10 Is Null "
'       m_strSQL1_3 = m_strSQL1_3 + " AND NP10 Is Null "
'       m_strSQL1_4 = m_strSQL1_4 + " AND NP10 Is Null "
'       m_strSQL1_5 = m_strSQL1_5 + " AND NP10 Is Null "
'   End If
   'Modify By Sindy 2015/9/23 改抓客戶業務員
   '智權人員
   If Me.txt1(11).Text <> "" Then
       strSQL1 = strSQL1 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL1_2 = m_strSQL1_2 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL1_3 = m_strSQL1_3 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL1_4 = m_strSQL1_4 + " AND CU13='" & Me.txt1(11).Text & "' "
       m_strSQL1_5 = m_strSQL1_5 + " AND CU13='" & Me.txt1(11).Text & "' "
   End If
   '2015/9/23 END
   
   'Modify By Sindy 2013/10/23
   '當無輸入案件性質或案件性質為領證
   If (Len(Trim(txt1(3))) = 0 Or (Trim(txt1(3)) >= "601" And Trim(txt1(4)) <= "601")) Then
      strNP07 = "and ((" & Right(Trim(strNpSqlOfNoSalesDuty), Len(Trim(strNpSqlOfNoSalesDuty)) - 3) & ") or (CP10 in(" & UpdateCaseResultCP10List & ") AND NP07='411'))"
   Else
      strNP07 = "and (" & Right(Trim(strNpSqlOfNoSalesDuty), Len(Trim(strNpSqlOfNoSalesDuty)) - 3) & ")"
   End If
   '2013/10/23 END
                      'Modified by Morgan 2023/3/29 +pa179
                      strSql = " SELECT PA26,pa77,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(NA01,PA09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','存在','N','消滅',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05," & SQLDate("CP57") & "||Decode(pA57,'Y','*',''),cp09, PA01,PA02, PA03, PA04," & SQLDate("PA10") & ", CP57,PA08,PA91,PA26,CP44,PA10,PA20,np08,pa179 FROM PATENT,CASEPROGRESS A,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & _
                              " AND NP06 is null AND NP01=CP09(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) AND PA04='00' AND (PA57<>'Y' or pa57 is null) " & strSQL1
   strSql = strSql + " union all SELECT PA27,pa77,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA27),pa11,NVL(NA01,PA09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','存在','N','消滅',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05," & SQLDate("CP57") & "||Decode(pA57,'Y','*',''),cp09, PA01,PA02, PA03, PA04," & SQLDate("PA10") & ", CP57,PA08,PA91,PA26,CP44,PA10,PA20,np08,pa179 from PATENT,CASEPROGRESS A,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & _
                              " AND NP06 is null AND NP01=CP09(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND SUBSTR(PA27,1,8)=CU01(+) AND DECODE(SUBSTR(PA27,9,1),NULL,'0',SUBSTR(PA27,9,1))=CU02(+) AND PA09=NA01(+) AND PA04='00' AND (PA57<>'Y' or pa57 is null) " & m_strSQL1_2
   strSql = strSql + " union all SELECT PA28,pa77,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA28),pa11,NVL(NA01,PA09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','存在','N','消滅',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05," & SQLDate("CP57") & "||Decode(pA57,'Y','*',''),cp09, PA01,PA02, PA03, PA04," & SQLDate("PA10") & ", CP57,PA08,PA91,PA26,CP44,PA10,PA20,np08,pa179 from PATENT,CASEPROGRESS A,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & _
                              " AND NP06 is null AND NP01=CP09(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND SUBSTR(PA28,1,8)=CU01(+) AND DECODE(SUBSTR(PA28,9,1),NULL,'0',SUBSTR(PA28,9,1))=CU02(+) AND PA09=NA01(+) AND PA04='00' AND (PA57<>'Y' or pa57 is null) " & m_strSQL1_3
   strSql = strSql + " union all SELECT PA29,pa77,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA29),pa11,NVL(NA01,PA09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','存在','N','消滅',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05," & SQLDate("CP57") & "||Decode(pA57,'Y','*',''),cp09, PA01,PA02, PA03, PA04," & SQLDate("PA10") & ", CP57,PA08,PA91,PA26,CP44,PA10,PA20,np08,pa179 from PATENT,CASEPROGRESS A,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & _
                              " AND NP06 is null AND NP01=CP09(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND SUBSTR(PA29,1,8)=CU01(+) AND DECODE(SUBSTR(PA29,9,1),NULL,'0',SUBSTR(PA29,9,1))=CU02(+) AND PA09=NA01(+) AND PA04='00' AND (PA57<>'Y' or pa57 is null) " & m_strSQL1_4
   strSql = strSql + " union all SELECT PA30,pa77,NVL(Decode(CU04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA30),pa11,NVL(NA01,PA09),NP07," & SQLDate("cp27") & "," & SQLDate("np08") & ",pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','存在','N','消滅',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05," & SQLDate("CP57") & "||Decode(pA57,'Y','*',''),cp09, PA01,PA02, PA03, PA04," & SQLDate("PA10") & ", CP57,PA08,PA91,PA26,CP44,PA10,PA20,np08,pa179 from PATENT,CASEPROGRESS A,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION,NextProgress WHERE NP02 in(" & strNP02 & ") " & strNP07 & _
                              " AND NP06 is null AND NP01=CP09(+) AND NP02=CPM01(+) AND NP07=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND SUBSTR(PA30,1,8)=CU01(+) AND DECODE(SUBSTR(PA30,9,1),NULL,'0',SUBSTR(PA30,9,1))=CU02(+) AND PA09=NA01(+) AND PA04='00' AND (PA57<>'Y' or pa57 is null) " & m_strSQL1_5
   adoRst.CursorLocation = adUseClient
   k = 0
   adoRst.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRst.RecordCount <> 0 And adoRst.RecordCount > 0 Then
     InsertQueryLog (adoRst.RecordCount)
     With adoRst
         .MoveFirst
         DoEvents
         Do While .EOF = False
            For i = LBound(strTemp) To UBound(strTemp)
                strTemp(i) = ""
            Next i
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Right(strTemp(16), 1) = "*" Then
                strTemp(9) = "*" + strTemp(9)
                strTemp(16) = Replace(strTemp(16), "*", "閉卷")
            End If
            '申請日
            strTemp(23) = "" & .Fields(22).Value
            CheckOC2
            strSql = "select min(pd05) from pridate where pd01='" & CheckStr(.Fields(18)) & "' and pd02='" & CheckStr(.Fields(19)) & "' and pd03='" & CheckStr(.Fields(20)) & "' and pd04='" & CheckStr(.Fields(21)) & "' "
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
                 strTemp(24) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
            Else
                 strTemp(24) = ""
            End If
            
            '預估費用:X46513000 / CFP X66408000 (99031)
            strCharge = ""
            m_NP02 = CheckStr(.Fields(18))
            m_NP03 = CheckStr(.Fields(19))
            m_NP04 = Right("0" & CheckStr(.Fields(20)), 1)
            m_NP05 = Right("00" & CheckStr(.Fields(21)), 2)
            If m_NP02 = "P" Then
               strYF03 = "Y00000001"
            ElseIf m_NP02 = "CFP" And InStr("" & .Fields("PA91"), "大個體") > 0 Then
               strYF03 = "Y00000002"
            Else 'FCP,CFP
               strYF03 = "Y00000000"
            End If
                        
            'Added by Morgan 2023/3/29
            If m_NP02 = "CFP" And strSrvDate(1) >= PA179啟用日 Then
               If .Fields("PA179") = "1" Then
                  strYF03 = "Y00000002"
               ElseIf .Fields("PA179") = "3" Then
                  strYF03 = "Y00000003"
               End If
            End If
            'end 2023/3/29
            
            m_PA08 = CheckStr(.Fields("PA08"))
            m_PA09 = CheckStr(.Fields(4))
            m_NP07 = CheckStr(.Fields("NP07"))
            strDiscCase = PUB_GetCaseDiscStat(m_NP02 & m_NP03 & m_NP04 & m_NP05)
            m_PA26 = CheckStr(.Fields("PA26"))
            m_CP44 = CheckStr(.Fields("CP44"))
            m_PA10 = CheckStr(.Fields("PA10"))
            m_PA20 = CheckStr(.Fields("PA20"))
            m_NP08 = CheckStr(.Fields("np08"))
            
            '416.實體審查
            If m_NP07 = "416" Then
               'Modified by Morgan 2023/10/16
               'strSql = "SELECT YF06+YF07 FROM PATENTYEARFEE " & _
               '         "WHERE YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND F03='" & strYF03 & "' AND YF04='" & m_NP07 & "' AND YF05=1"
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               'If intI = 1 Then
               '   strCharge = "" & RsTemp.Fields(0)
               'End If
               'Added by Morgan 2024/10/4
               'CFP案要傳 strYF03 否則都會變小個體
               If m_NP02 = "CFP" Then
                  strCharge = PUB_GetYF0607(m_PA09, m_PA08, strYF03, m_NP07, "1", "1", m_NP02)
               Else
               'end 2024/10/4
                  strCharge = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, m_NP07, "1", "1", m_NP02)
               End If
               'end 2023/10/16
               
            
            ElseIf m_NP02 = "CFP" And (m_NP07 = "605" Or m_NP07 = "606" Or m_NP07 = "607") Then
               '取得下次繳費次數
               m_Nexttimes = PUB_Getnexttimes(m_NP02, m_NP03, m_NP04, m_NP05, m_strYear, m_PA91)
               'Modified by Morgan 2023/10/16
               'strSql = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE " & _
               '         "WHERE YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND YF03='" & strYF03 & "' AND YF04='" & m_NP07 & "' AND YF05=" & CNULL(m_strYear)
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               'If intI = 1 Then
               '   strCharge = RsTemp.Fields(0) + RsTemp.Fields(1)
               'End If
               'Modified by Morgan 2025/8/21
               'strCharge = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, m_NP07, m_strYear, m_strYear, m_NP02)
               strCharge = PUB_GetYF0607(m_PA09, m_PA08, strYF03, m_NP07, m_strYear, m_strYear, m_NP02)
               'end 2023/10/16
               
            '601.領證及繳年費
            ElseIf m_NP07 = "601" Or m_NP07 = "411" Then
               If m_PA09 = "020" Then '大陸
                  
                  'Modified by Morgan2014/12/24 改呼叫共用函數
                  Dim lTmp As Long
                  'Added by Morgan 2018/2/27
                  Dim pa(26) As String
                  Dim dblSFee As Double
                  'end 2018/2/27
                  
                  str_P020Year = PUB_GetChina605StartYear2(m_NP02, m_NP03, m_NP04, m_NP05)
                  If str_P020Year = "" Then
                     If m_PA20 = "" Then
                        str_P020Year = PUB_GetChina605StartYear(m_NP08, m_PA10)
                     Else
                        str_P020Year = PUB_GetChina605StartYear(m_PA20, m_PA10)
                     End If
                  End If
                  'Modified by Lydia 2015/03/30 +系統別
                  'Modified by Morgan 2018/2/27 取消 PUB_GetFee 改用 PUB_Get020601Fee(與接洽單一致)
                  'strCharge = PUB_GetFee(m_NP02, m_PA09, m_PA08, m_PA26, Val(str_P020Year), , m_CP44)
                  pa(1) = m_NP02: pa(8) = m_PA08: pa(9) = m_PA09: pa(26) = m_PA26
                  strCharge = PUB_Get020601Fee(pa, m_CP44, Val(str_P020Year), Val(str_P020Year), dblSFee)
                  strCharge = dblSFee + strCharge
                  'end 2018/2/27
                  
'                  Dim lTmp As Long, m_dbl601OfficialFee As Double, lBase As Long, lPlus As Long
'                  lTmp = PUB_GetYF06(m_PA09, m_PA08, ChangeCustomerL(m_PA26), "601", "1", "1")
'                  If lTmp = 0 Then
'                     If m_CP44 <> "" Then
'                        lTmp = PUB_GetYF0607(m_PA09, m_PA08, ChangeCustomerL(m_CP44), "601", "1", "1")
'                     End If
'                     If lTmp = 0 Then
'                        lTmp = PUB_GetYF0607(m_PA09, m_PA08, "Y00000001", "601", "1", "1")
'                        m_dbl601OfficialFee = PUB_GetYF07(m_PA09, m_PA08, "Y00000001", "601", "1", "1")
'                     Else
'                        m_dbl601OfficialFee = PUB_GetYF07(m_PA09, m_PA08, ChangeCustomerL(m_CP44), "601", "1", "1")
'                     End If
'                  Else
'                     If m_CP44 <> "" Then
'                        m_dbl601OfficialFee = PUB_GetYF07(m_PA09, m_PA08, ChangeCustomerL(m_CP44), "601", "1", "1")
'                     End If
'                     If m_dbl601OfficialFee = 0 Then
'                        m_dbl601OfficialFee = PUB_GetYF07(m_PA09, m_PA08, "Y00000001", "601", "1", "1")
'                     End If
'                     lTmp = lTmp + m_dbl601OfficialFee
'                  End If
'                  '大陸年度: 核准日+5個月為預定公告日--敏惠
'                  str_P020Year = ""
'                  If m_PA20 = "" Then
'                     str_P020Year = Int((TransDate(CompDate(1, 5, m_NP08), 2) - TransDate(m_PA10, 2)) / 10000) + 1
'                  Else
'                     str_P020Year = Int((TransDate(CompDate(1, 5, m_PA20), 2) - TransDate(m_PA10, 2)) / 10000) + 1
'                  End If
'                  If Val(i) > 0 And i <> "3" Then
'                     '年費(先抓1-3年的Base金額,再抓輸入大陸年度的金額, 計算出差額, 再加上上面之領證則為大陸領證費)
'                     lBase = PUB_GetYF07(m_PA09, m_PA08, ChangeCustomerL(m_PA26), "605", "3", "3")
'                     If lBase = 0 Then
'                        If m_CP44 <> "" Then
'                           lBase = PUB_GetYF07(m_PA09, m_PA08, ChangeCustomerL(m_CP44), "605", "3", "3")
'                        End If
'                        If lBase = 0 Then
'                           lBase = PUB_GetYF07(m_PA09, m_PA08, "Y00000001", "605", "3", "3")
'                        End If
'                     End If
'                     lPlus = PUB_GetYF07(m_PA09, m_PA08, ChangeCustomerL(m_PA26), "605", str_P020Year, str_P020Year)
'                     If lPlus = 0 Then
'                        If m_CP44 <> "" Then
'                           lPlus = PUB_GetYF07(m_PA09, m_PA08, ChangeCustomerL(m_CP44), "605", str_P020Year, str_P020Year)
'                        End If
'                        If lPlus = 0 Then
'                           lPlus = PUB_GetYF07(m_PA09, m_PA08, "Y00000001", "605", str_P020Year, str_P020Year)
'                        End If
'                     End If
'                  End If
'                  strCharge = lTmp + lPlus - lBase '大陸領證費; 若為下一年年費未屆時, 預估費用=大陸領證費
'                  m_dbl601OfficialFee = m_dbl601OfficialFee + lPlus - lBase

                  'end 2014/12/24
                  
                  '下一年年費將屆時, 預估費用=總費用
                  If PUB_CheckYear(m_PA08, m_PA10, m_PA20) = "Y" Then
                     strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0),NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND YF03='" & ChangeCustomerL(m_CP44) & "' AND YF04='605' AND YF05=2"
                     intI = 1
                     lTmp = 0
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        lTmp = Val(RsTemp.Fields(0))
                     Else
                        '內專抓代理人Y00000001
                        'Modified by Morgan 2023/10/16
                        'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0),NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND YF03='Y00000001' AND YF04='605' AND YF05=2"
                        'intI = 1
                        'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        'If intI = 1 Then
                        '   lTmp = Val(RsTemp.Fields(0))
                        'End If
                        lTmp = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, "605", "2", "2", m_NP02)
                        'end 2023/10/16
                     End If
                     strCharge = strCharge + lTmp '總費用
                  End If
                  
               Else
                  strExc(1) = "": strExc(2) = "": strExc(3) = ""
                  'Modified by Morgan 2023/10/16
                  'strSql = "Select YF06,YF07 From PatentYearFee " & _
                  '         "Where YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND YF03='" & strYF03 & "' AND YF04='601' AND YF05=1"
                  'intI = 1
                  'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  'If intI = 1 Then
                  '   strExc(1) = "" & RsTemp("YF06") '領證服務費
                  '   strExc(2) = "" & RsTemp("YF07") '領證規費
                  'End If
                  
                  'strSql = "Select YF07 From PatentYearFee " & _
                  '         "Where YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND YF03='" & strYF03 & "' AND YF04='605' AND YF05=1"
                  'intI = 1
                  'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  'If intI = 1 Then
                  '   strExc(3) = "" & RsTemp("YF07") '年費規費
                  '   If strDiscCase = "Y" Then
                  '      strExc(3) = Val(strExc(3)) - 800
                  '   End If
                  'End If
                  
                  'Added by Morgan 2024/10/4
                  'CFP案要傳 strYF03 否則都會變小個體
                  If m_NP02 = "CFP" Then
                     strExc(0) = PUB_GetYF0607(m_PA09, m_PA08, strYF03, "601", "1", "1", m_NP02, strExc(1), strExc(2))
                     strExc(0) = PUB_GetYF0607(m_PA09, m_PA08, strYF03, "605", "1", "1", m_NP02, , strExc(3))
                  Else
                  'end 2024/10/4
                     strExc(0) = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, "601", "1", "1", m_NP02, strExc(1), strExc(2))
                     strExc(0) = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, "605", "1", "1", m_NP02, , strExc(3))
                  End If
                  
                  If m_PA09 = "000" Then 'Added by Morgan 2023/12/4
                     If strDiscCase = "Y" Then
                        strExc(3) = Val(strExc(3)) - 800
                     End If
                  End If
                  'end 2023/10/16
                  strCharge = Val(strExc(1)) + Val(strExc(2)) + Val(strExc(3))
               End If
               
            '內專年費
            ElseIf m_NP07 = "605" Then
               '大陸,澳門(044)
               If m_PA09 = "020" Or m_PA09 = "044" Then
                  '取得下次繳費年度
                  strPA72NextYear = PUB_getPA72NextYear(m_NP02, m_NP03, m_NP04, m_NP05, , m_bFirstYear)
                  If strPA72NextYear <> "" Then
                     'Modified by Morgan 2023/10/16
                     'strCharge = PUB_GetYF0607(m_PA09, m_PA08, strYF03, m_NP07, strPA72NextYear, strPA72NextYear)
                     strCharge = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, m_NP07, strPA72NextYear, strPA72NextYear, m_NP02)
                     'end 2023/10/16
                  End If
               ElseIf m_PA09 = "013" Then '香港
                  '取得已繳費年度及專利種類
                  strPA72NextYear = PUB_getNextPayYear(m_NP02, m_NP03, m_NP04, m_NP05, strPA72Year)
                  If strPA72NextYear <> "" Then
                     'Modified by Morgan 2023/10/16
                     'strCharge = Val(PUB_GetYF0607(m_PA09, m_PA08, strYF03, m_NP07, strPA72NextYear, strPA72NextYear))
                     strCharge = Val(PUB_GetYF0607(m_PA09, m_PA08, m_PA26, m_NP07, strPA72NextYear, strPA72NextYear, m_NP02))
                     'end 2023/10/16
                  End If
               Else
                  '台灣
                  If m_PA09 = "000" Then
                     '取得下次繳費年度
                     strPA72NextYear = PUB_getPA72NextYear(m_NP02, m_NP03, m_NP04, m_NP05, strMaxFeeYear)
                     If Val(strPA72NextYear) > 0 Then
                        '服務費,規費
                        'Modified by Morgan 2023/10/16
                        'strSql = "Select YF06,YF07 From PatentYearFee Where YF01='" & m_PA09 & "' AND YF02='" & m_PA08 & "' AND YF03='" & strYF03 & "' AND YF04='" & m_NP07 & "' AND YF05=" & strPA72NextYear
                        'intI = 1
                        'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        'If intI = 1 Then
                        '   strExc(1) = "" & RsTemp("YF06") '服務費
                        '   strExc(2) = "" & RsTemp("YF07") '規費
                        'Else
                        '   strExc(1) = ""
                        '   strExc(2) = ""
                        'End If
                        strExc(0) = PUB_GetYF0607(m_PA09, m_PA08, m_PA26, m_NP07, strPA72NextYear, strPA72NextYear, m_NP02, strExc(1), strExc(2))
                        'end 2023/10/16
                        
                        If Val(strExc(2)) > 0 Then
                           If Val(strPA72NextYear) < 7 Then
                              '減免
                              If strDiscCase = "Y" Then
                                 If Val(strPA72NextYear) < 4 Then
                                    strExc(2) = Val(strExc(2)) - 800
                                 Else
                                    strExc(2) = Val(strExc(2)) - 1200
                                 End If
                              End If
                           End If
                        End If
                        strCharge = Val(strExc(1)) + Val(strExc(2))
                     End If
                  End If
               End If
            End If
            '專利年費資料檔因為加年費年度說明欄位有新增沒有費用的資料
            If strCharge = "0" Then strCharge = ""
            
            strSql = "INSERT INTO R050317_C  " & strTBF & " values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & ChgSQL(strTemp(18)) & "','" & strUserNum & "','" & strTemp(23) & "','" & strTemp(24) & "','" & .Fields(18).Value & "','" & .Fields(19).Value & "','" & .Fields(20).Value & "','" & .Fields(21).Value & "','" & "" & .Fields("CP57").Value & "'," & CNULL(strCharge) & ")"
            cnnConnection.Execute strSql
            IsHaveData = True
            k = k + 1
            DoEvents
            .MoveNext
         Loop
     End With
   Else
     InsertQueryLog (0)
     ShowNoData
     Screen.MousePointer = vbDefault
     Exit Sub
   End If
   
   If adoRst.State <> 0 Then
      adoRst.Close
   End If
   PrintDataCp_A4
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub PrintDataCp_A4()
Dim strCust As String '記錄申請人
Dim strSystemKind As String '記錄系統類別
Dim blnFirstPage As Boolean '判斷是否第一頁
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCaseNo As String '申請人+本所案號

'控制 word 換頁用
blnWordNewPage = True

On Error GoTo ErrHnd

'列印第一頁
blnFirstPage = True
'國內外分開列印(不同系統類別分開列印)
If txt1(8) = "Y" Then
   '修改排序方式申請人, 系統類別
   strSql = "SELECT DISTINCT R013001,R013019 FROM R050317_C WHERE ID='" & strUserNum & "' GROUP BY R013001,R013019 ORDER BY R013001,R013019 "
'國內外不分開列印(不管系統類別)
Else
   strSql = "SELECT DISTINCT R013001 FROM R050317_C WHERE ID='" & strUserNum & "' GROUP BY R013001 ORDER BY R013001 "
End If
CheckOC
'頁數加一
Page = Page + 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    RefreshColData 1, "申請人"
    With adoRecordset
        .MoveFirst
        '記錄申請人
        strCust = "" & .Fields("R013001").Value
        '記錄系統類別
        If txt1(8) = "Y" Then
            strSystemKind = "" & .Fields("R013019").Value
        Else
            strSystemKind = ""
        End If
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields("R013001").Value)
            If txt1(8) = "Y" Then
               strTemp(21) = CheckStr(.Fields("R013019").Value)
            Else
               strTemp(21) = ""
            End If
            strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU16||DECODE(CU17,NULL,'',','||CU17),CU20,CU18||DECODE(CU19,NULL,'',','||CU19),CU58,CU61 FROM CUSTOMER WHERE CU01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND CU02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
            Else
                For i = 0 To 13
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            '若申請人不同則跳頁
            If strCust <> strTemp(20) Then
                strCust = strTemp(20)
                strSystemKind = strTemp(21)
                If txt1(8) = "Y" Then
                   strSystemKind = strTemp(21)
                Else
                   strSystemKind = ""
                End If
                Page = Page + 1
                'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'                '紙本
'                If Check3(0).Value = vbChecked Then
'                    Printer.NewPage
'                ElseIf Check3(1).Value = vbChecked Then
'                    fn_DocNewPage
'                End If
                'Excel
                If intChoose = 2 Then
                    If Page <> 1 Then fn_SetExcel '換頁前設定
                    fn_CreateExcel Page, strSystemKind, blnFirstPage
                Else
                     fn_DocNewPage
                End If
                'end 2022/05/02
                PrintTitleCp_A4
            '若申請人相同
            Else
                '記錄系統類別
                If txt1(8) = "Y" Then
                    '若系統類別不同
                    If strSystemKind <> strTemp(21) Then
                        strSystemKind = strTemp(21)
                        Page = Page + 1
                        'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'                        '紙本
'                        If Check3(0).Value = vbChecked Then
'                            Printer.NewPage
'                        ElseIf Check3(1).Value = vbChecked Then
'                            fn_DocNewPage
'                        End If
                        'Excel
                        If intChoose = 2 Then
                            If Page <> 1 Then fn_SetExcel '換頁前設定
                            fn_CreateExcel Page, strSystemKind, blnFirstPage
                        Else
                            fn_DocNewPage
                        End If
                        'end 2022/05/02
                        PrintTitleCp_A4
                    End If
                End If
            End If
            '若為第一頁才印表頭
            If blnFirstPage = True Then
                'Modify by Amy 2022/05/02 +Excel
                If intChoose = 2 Then
                    fn_CreateExcel Page, strSystemKind, blnFirstPage
                Else
                    PrintTitleCp_A4
                End If
                blnFirstPage = False
            End If
            '國內外是否分開列印
            If txt1(8) = "Y" Then
               strSql = "SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, pa16 as CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027,NVL(Nvl(decode(pa09,'000',c1.cpm03,c1.cpm04),Nvl(c1.CPM10,c1.CPM13)),R013006) as CPMName, NVL(Nvl(decode(pa09,'000',c2.cpm03,c2.cpm04),Nvl(c2.CPM10,c2.CPM13)),cp10) as CPMName2,pa08,pa09 FROM R050317_C,caseprogress,patent,CASEPROPERTYMAP c1,CASEPROPERTYMAP c2 WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and R013022=c1.cpm01(+) and R013006=c1.cpm02(+) and cp01=c2.cpm01(+) and cp10=c2.cpm02(+) "
            Else
               strSql = "SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, pa16 as CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027,NVL(Nvl(decode(pa09,'000',c1.cpm03,c1.cpm04),Nvl(c1.CPM10,c1.CPM13)),R013006) as CPMName, NVL(Nvl(decode(pa09,'000',c2.cpm03,c2.cpm04),Nvl(c2.CPM10,c2.CPM13)),cp10) as CPMName2,pa08,pa09 FROM R050317_C,caseprogress,patent,CASEPROPERTYMAP c1,CASEPROPERTYMAP c2 WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and R013022=c1.cpm01(+) and R013006=c1.cpm02(+) and cp01=c2.cpm01(+) and cp10=c2.cpm02(+) "
            End If
            '依本所案號排序
            If Me.txt1(9).Text = "1" Then
               strSql = strSql & " ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010),r013015, Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '依案件名稱排序
            ElseIf Me.txt1(9).Text = "2" Then
               strSql = strSql & " ORDER BY r013015,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '依申請國家+本所案號排序
            ElseIf Me.txt1(9).Text = "3" Then
               strSql = strSql & " ORDER BY R013005, DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010),r013015, Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            'Add By Sindy 2013/12/4
            '依客戶案件案號排序
            ElseIf Me.txt1(9).Text = "5" Then
               strSql = strSql & " ORDER BY R013009,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '2013/12/4 END
            '依申請國家+案件名稱排序
            Else
               strSql = strSql & " ORDER BY R013005, r013015,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            End If
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    strTemp(0) = "" & adoRecordset1.Fields(0).Value
                    strTemp(1) = StrToStr("" & adoRecordset1.Fields(1).Value, 42)
                    strTemp(2) = "" & adoRecordset1.Fields(2).Value
                    strTemp(3) = StrToStr("" & adoRecordset1.Fields(3).Value, 10)
                    strTemp(4) = StrToStr("" & adoRecordset1.Fields(4).Value, 4)
                    strTemp(5) = ""
                    strTemp(6) = "" & adoRecordset1.Fields("R013008").Value '下一程序的本所期限
                    strTemp(7) = "" & Format(adoRecordset1.Fields("R013027").Value, "###,##0")
                    strTemp(8) = StrToStr("" & adoRecordset1.Fields(8).Value, 10)
                    strTemp(9) = StrToStr(GetNationName("" & adoRecordset1.Fields(9).Value, 0), 6)
                    strTemp(10) = "" & adoRecordset1.Fields(10).Value
                    strTemp(11) = ""
                    strTemp(12) = ""
                    strTemp(13) = ""
                    strTemp(14) = ""
                    '下一程序
                    If adoRecordset1.Fields("R013006").Value = "411" Then
                        strTemp(15) = "" & adoRecordset1.Fields("CPMName2").Value & "預估核准"
                    ElseIf adoRecordset1.Fields("R013006").Value = "605" Or adoRecordset1.Fields("R013006").Value = "606" Or adoRecordset1.Fields("R013006").Value = "607" Then
                        '下次繳費年度
                        Dim pa(4) As String, strText As String
                        pa(1) = adoRecordset1.Fields("R013022")
                        pa(2) = adoRecordset1.Fields("R013023")
                        pa(3) = adoRecordset1.Fields("R013024")
                        pa(4) = adoRecordset1.Fields("R013025")
                        Call PUB_GetNextYear(pa(), strText)
                        strTemp(15) = "" & adoRecordset1.Fields("CPMName").Value & "[" & strText & "]"
                    Else
                        strTemp(15) = "" & adoRecordset1.Fields("CPMName").Value
                    End If
                    GetPatentDuration "" & adoRecordset1("R013022").Value, "" & adoRecordset1("R013023").Value, "" & adoRecordset1("R013024").Value, "" & adoRecordset1("R013025").Value
                    'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
                    If intChoose = 2 Then
'*********  Excel *********
                        '依欄位順序顯示
                        '               本所案號                   客戶案件案號              商標名                     申日請日                       申請國                      申請號                審定號                         種類                            專用起日                   專用止日               本所期限                      下一程序                  預估費用
                        strXlsData = strTemp(0) & "$$" & strTemp(8) & "$$" & strTemp(1) & "$$" & strTemp(2) & "$$" & strTemp(9) & "$$" & strTemp(3) & "$$" & strTemp(10) & "$$" & strTemp(4) & "$$" & strXlsTp(1) & "$$" & strXlsTp(2) & "$$" & strTemp(6) & "$$" & strTemp(15) & "$$" & strTemp(7)
                        fn_PutExcel "" & adoRecordset1("R013022").Value, "" & adoRecordset1("R013023").Value, "" & adoRecordset1("R013024").Value, "" & adoRecordset1("R013025").Value
                        intXlsRow = intXlsRow + 1
'*********  End Excel *********
                    Else
'*********  Word *********
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.Line (0, iPrint)-(16000, iPrint)
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                            fn_DocNewPage
    '                        End If
                            'Word
                            fn_PutWords fn_StrLineToWord, , , 10
                            fn_DocNewPage
                            PrintTitleCp_A4
                        End If
                        '列印明細資料
                        PrintDatilCp_A4
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_DocNewPage
    '                        End If
                            fn_DocNewPage
                            PrintTitleCp_A4
                        End If
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.Line (0, iPrint)-(16000, iPrint)
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                            fn_DocNewPage
    '                        End If
                            fn_PutWords fn_StrLineToWord, , , 10
                            fn_DocNewPage
                            PrintTitleCp_A4
                        Else
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.CurrentX = PLeft(0)
    '                            Printer.CurrentY = iPrint
    '                            Printer.Print String(160, "-")
    '                            iPrint = iPrint + 300
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                        End If
                            fn_PutWords fn_StrLineToWord, , , 10
                            
                            If iPrint >= o_PaperHeight Then
                                Page = Page + 1
    '                            '紙本
    '                            If Check3(0).Value = vbChecked Then
    '                                Printer.NewPage
    '                            ElseIf Check3(1).Value = vbChecked Then
    '                                fn_DocNewPage
    '                            End If
                                fn_DocNewPage
                                PrintTitleCp_A4
                            End If
                        End If
'*********  End Word *********
                    End If
                    adoRecordset1.MoveNext
                Loop
                'Modify by Amy 2022/05/02 +if 不是Excel才run,紙本原本就沒使用
                If intChoose <> 2 Then
                    If iPrint >= o_PaperHeight Then
                        Page = Page + 1
    '                    '紙本
    '                    If Check3(0).Value = vbChecked Then
    '                        Printer.NewPage
    '                    ElseIf Check3(1).Value = vbChecked Then
    '                        fn_DocNewPage
    '                    End If
                        fn_DocNewPage
                        PrintTitleCp_A4
                    End If
                End If
                'end 2022/05/02
                If txt1(8) = "Y" Then
                  strSql = "SELECT DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050317_c WHERE R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' "
                Else
                  strSql = "SELECT DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050317_c WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' "
                End If
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                    IntTot = AdoRecordSet3.RecordCount
                End If
                'Modify by Amy 2022/05/02 +不是Excel才印總計,紙本原本就沒使用
'                '紙本
'                If Check3(0).Value = vbChecked Then
'                    Printer.CurrentX = PLeft(5)
'                    Printer.CurrentY = iPrint
'                    Printer.Print "總計：" & Format(IntTot, "##0")
'                    iPrint = iPrint + 300
'                ElseIf Check3(1).Value = vbChecked Then
'                    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 54.5) & "總計：" & Format(IntTot, "##0"), 64), , , 10
'                End If
                'Word
                If intChoose <> 2 Then
                    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 54.5) & "總計：" & Format(IntTot, "##0"), 64), , , 10
                End If
                'end 2022/05/02
                CheckOC3
            End If
            CheckOC2
            .MoveNext
        Loop
        'Add by Amy 3030/05/02 +Excel
        If intChoose = 2 Then
            fn_SetExcel
            fn_PutEndXls '存檔
        End If
    End With
End If
CheckOC
'Mark by Amy 2022/05/02 原本就沒使用
''紙本
'If Check3(0).Value = vbChecked Then
'   Printer.EndDoc
'End If
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub PrintTitleCp_A4()
'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
''紙本
'If Check3(0).Value = vbChecked Then
'    GetPleftCp_A4
'    iPrint = 500
'    Printer.Font.Size = 22
'    Printer.Font.Bold = True
'    Printer.Font.Underline = True
'    Printer.FontName = "細明體"
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "客戶案件預算預估表"
'    Printer.Font.Bold = False
'    Printer.Font.Underline = False
'    iPrint = iPrint + 500
'    Printer.Font.Size = 16
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "　　　　（專利）"
'    Printer.Font.Size = 10
'    'Add By Sindy 2012/10/31
'    iPrint = iPrint + 500
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所期限：" & ChangeTStringToTDateString(txt1(16)) & " - " & ChangeTStringToTDateString(txt1(17))
'    '2012/10/31 End
'    iPrint = iPrint + 500
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "收件人："
'    If Len(StrTemp5(12)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(12)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(13)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(13)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(0)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(0)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(1)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(1)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(2)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(2)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(3)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(3)
'       iPrint = iPrint + 300
'    End If
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint - 300
'
'    Printer.Print "電話：" & StrTemp5(5)
'
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "地址：" & StrTemp5(4)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "E-mail：" & StrTemp5(7)
'
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "Page：" & str(Page)
'    iPrint = iPrint + 300
'    Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'    Printer.CurrentY = iPrint
'    Printer.Print StrTemp5(6)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "傳真：" & StrTemp5(9)
'
'    iPrint = iPrint + 300
'    Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'    Printer.CurrentY = iPrint
'    Printer.Print StrTemp5(8)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "預設接洽人：" & GetCU08(strTemp(20))   '2019/8/29加預設二字
'    Printer.CurrentX = PLeft(4)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "智權人員：" & Me.txt1(11).Text & " " & GetStaffName(Me.txt1(11).Text, True)
'
'    iPrint = iPrint + 300
'    If Len(StrTemp5(10)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(10)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(11)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(11)
'       iPrint = iPrint + 300
'    End If
'    Printer.Line (0, iPrint)-(16000, iPrint)
'    iPrint = iPrint + 100
'
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所案號"
'    Printer.CurrentX = PLeft(1)
'    Printer.CurrentY = iPrint
'    Printer.Print "專利名稱"
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請日"
'    Printer.CurrentX = PLeft(3)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請案號"
'    Printer.CurrentX = PLeft(4)
'    Printer.CurrentY = iPrint
'    Printer.Print "種類"
'    Printer.CurrentX = PLeft(5)
'    Printer.CurrentY = iPrint
'    Printer.Print "專　用　期　限"
'    Printer.CurrentX = PLeft(6)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所期限"
'    Printer.CurrentX = PLeft(7)
'    Printer.CurrentY = iPrint
'    Printer.Print "預估費用"
'    iPrint = iPrint + 300
'
'    Printer.CurrentX = PLeft(8)
'    Printer.CurrentY = iPrint
'    Printer.Print "客戶案件案號"
'    Printer.CurrentX = PLeft(9)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請國家"
'    Printer.CurrentX = PLeft(10)
'    Printer.CurrentY = iPrint
'    Printer.Print "專利號數"
'    Printer.CurrentX = PLeft(15)
'    Printer.CurrentY = iPrint
'    Printer.Print "下一程序"
'    iPrint = iPrint + 300
'    Printer.Line (0, iPrint)-(16000, iPrint)
'    iPrint = iPrint + 100
'ElseIf Check3(1).Value = vbChecked Then
'Word
If intChoose <> 2 Then
'end 2022/05/02
    If blnWordNewPage = True And Page <> 1 Then
        fn_DocNewPage
    End If
    blnWordNewPage = False
    Dim isPrintDate As Boolean
    Dim o_tmp1 As String
    Dim o_tmp2 As String
    Dim o_tmp3 As String
    Dim o_tmp4 As String
    Dim o_tmp5 As String
    Dim o_tmp6 As String
    iPrint = 0
    o_tmp1 = StrTemp5(12)
    o_tmp2 = StrTemp5(13)
    o_tmp3 = StrTemp5(0)
    o_tmp4 = StrTemp5(1)
    o_tmp5 = StrTemp5(2)
    o_tmp6 = StrTemp5(3)
    If Len(CheckStr(o_tmp6)) <> 0 Then
        If Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            o_tmp6 = fn_StrToWordFmt("收件人：" & o_tmp6, 29) & "電話：" & StrTemp5(5)
        Else
            o_tmp6 = fn_StrToWordFmt("　　　：" & o_tmp6, 29) & "電話：" & StrTemp5(5)
        End If
    End If
    If Len(CheckStr(o_tmp5)) <> 0 Then
        If Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) = 0 Then
                o_tmp5 = fn_StrToWordFmt("收件人：" & o_tmp5, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp5 = fn_StrToWordFmt("收件人：" & o_tmp5, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) = 0 Then
                o_tmp5 = fn_StrToWordFmt("　　　：" & o_tmp5, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp5 = fn_StrToWordFmt("　　　：" & o_tmp5, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp4)) <> 0 Then
        If Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) = 0 Then
                o_tmp4 = fn_StrToWordFmt("收件人：" & o_tmp4, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp4 = fn_StrToWordFmt("收件人：" & o_tmp4, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) = 0 Then
                o_tmp4 = fn_StrToWordFmt("　　　：" & o_tmp4, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp4 = fn_StrToWordFmt("　　　：" & o_tmp4, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp3)) <> 0 Then
        If Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) = 0 Then
                o_tmp3 = fn_StrToWordFmt("收件人：" & o_tmp3, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp3 = fn_StrToWordFmt("收件人：" & o_tmp3, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) = 0 Then
                o_tmp3 = fn_StrToWordFmt("　　　：" & o_tmp3, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp3 = fn_StrToWordFmt("　　　：" & o_tmp3, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp2)) <> 0 Then
        If Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) = 0 Then
                o_tmp2 = fn_StrToWordFmt("收件人：" & o_tmp2, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp2 = fn_StrToWordFmt("收件人：" & o_tmp2, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) = 0 Then
                o_tmp2 = fn_StrToWordFmt("　　　：" & o_tmp2, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp2 = fn_StrToWordFmt("　　　：" & o_tmp2, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp1)) <> 0 Then
        If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) = 0 Then
            o_tmp1 = fn_StrToWordFmt("收件人：" & o_tmp1, 29) & "電話：" & StrTemp5(5)
        Else
            o_tmp1 = fn_StrToWordFmt("收件人：" & o_tmp1, 29)
        End If
    End If
    isPrintDate = False
    fn_PutWords "客戶案件預算預估表", wdAlignParagraphCenter, "細明體", 22, False, True, True
    fn_PutWords "（專利）", wdAlignParagraphCenter, , 16
    'Add By Sindy 2012/10/31 69=>64
    fn_PutWords fn_StrToWordFmt(" ", 64) & "本所期限：" & ChangeTStringToTDateString(txt1(16)) & " - " & ChangeTStringToTDateString(txt1(17)), , , 10
    '2012/10/31 End
    If Len(o_tmp1) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp1, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp1, 80), , , 10
        End If
    End If
    If Len(o_tmp2) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp2, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp2, 80), , , 10
        End If
    End If
    If Len(o_tmp3) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp3, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp3, 80), , , 10
        End If
    End If
    If Len(o_tmp4) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp4, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp4, 80), , , 10
        End If
    End If
    If Len(o_tmp5) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp5, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp5, 80), , , 10
        End If
    End If
    If Len(o_tmp6) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp6, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp6, 80), , , 10
        End If
    End If
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("地址：" & StrTemp5(4), 29) & "E-mail：" & StrTemp5(7), 64) & "Page：" & str(Page), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("　　　" & StrTemp5(6), 29) & "傳真：" & StrTemp5(9), 64), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(fn_StrToWordFmt("　　　" & StrTemp5(8), 29) & "預設接洽人：" & GetCU08(strTemp(20)), 47.5) & "智權人員：" & Me.txt1(11).Text & " " & GetStaffName(Me.txt1(11).Text, True), 64), , , 10
    If Len(StrTemp5(10)) <> 0 Then
        fn_PutWords fn_StrToWordFmt("　　　" & StrTemp5(10), 29), , , 10
    End If
    If Len(StrTemp5(11)) <> 0 Then
        fn_PutWords fn_StrToWordFmt("　　　" & StrTemp5(11), 29), , , 10
    End If
    fn_PutWords fn_StrLineToWord("="), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("本所案號", 11) & fn_StrToWordFmt("專利名稱", 18) & fn_StrToWordFmt("申請日", 7) & fn_StrToWordFmt("申請案號", 11) & fn_StrToWordFmt("種類", 5) & fn_StrToWordFmt("專　用　期　限", 12) & fn_StrToWordFmt("本所期限", 7) & "預估費用", 80), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("客戶案件案號", 29) & fn_StrToWordFmt("申請國家", 7) & fn_StrToWordFmt("專利號數", 28) & "下一程序", 80), , , 10
    fn_PutWords fn_StrLineToWord("="), , , 10
End If
End Sub

Private Sub GetPleftCp_A4()
Erase PLeft
'第一行
PLeft(0) = 0 '本所案號
PLeft(1) = 2100 '專利名稱
PLeft(2) = 6300 '申請日
PLeft(3) = 7600 '申請案號
PLeft(4) = 9800 '種類
PLeft(5) = 10700 '專用期限
PLeft(6) = 13100 '最近期限
PLeft(7) = 14500 '預估費用
'第二行
PLeft(8) = 0 '客戶案件案號
PLeft(9) = 6300 '申請國家
PLeft(10) = 7600 '專利號數
PLeft(11) = 9800 '准駁
PLeft(12) = 12100 '下次繳費日
PLeft(13) = 13100 '(年度)
PLeft(14) = 13100 '閉卷
PLeft(15) = 13100 '下一程序
End Sub

'取得專利專用期限
Private Sub GetPatentDuration(strCaseNo1 As String, strCaseNo2 As String, StrCaseNo3 As String, strCaseNo4 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim arrYearPay
Dim arrYearPaySet
Dim ii As Integer

StrSQLa = "Select PA24, PA25, PA57, PA72, PA08, NA21, NA23, NA25,pa09,pa21 From Patent, Nation Where PA09=NA01 And " & ChgPatent(strCaseNo1 & strCaseNo2 & StrCaseNo3 & strCaseNo4)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify by Amy 2022/05/02 +if 使用不同變數存資料 for Excel
    If intChoose = 2 Then
        For ii = LBound(strXlsTp) To UBound(strXlsTp)
            strXlsTp(ii) = ""
        Next ii
        
        strXlsTp(1) = ChangeWStringToWDateString("" & rsA.Fields("PA24").Value)
        strXlsTp(2) = ChangeWStringToWDateString("" & rsA.Fields("PA25").Value)
    Else
        strTemp(5) = ChangeWStringToWDateString("" & rsA("PA24").Value) & "-" & ChangeWStringToWDateString("" & rsA("PA25").Value)
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Sub PrintDatilCp_A4()
'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'紙本
'If Check3(0).Value = vbChecked Then
'    For i = 0 To 7
'       If i = 0 Then
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint
'          Printer.Print Replace(strTemp(i), "*", "＊")
'       ElseIf i = 1 Then
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint
'          Printer.Print Mid(strTemp(i), 1, 21)
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint + 300
'          Printer.Print Mid(strTemp(i), 22, 21)
'       Else
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint
'          Printer.Print strTemp(i)
'       End If
'    Next i
'    iPrint = iPrint + 300
'    For i = 8 To 14
'       Printer.CurrentX = PLeft(i)
'       Printer.CurrentY = iPrint
'       Printer.Print strTemp(i)
'    Next i
'    If strTemp(15) <> "" Then
'       Printer.CurrentX = PLeft(15)
'       Printer.CurrentY = iPrint
'       Printer.Print strTemp(15)
'    End If
'    iPrint = iPrint + 300
'ElseIf Check3(1).Value = vbChecked Then
'Excel
If intChoose = 2 Then
   
Else
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(Replace(strTemp(0), "*", "＊"), 11) & fn_StrToWordFmt(Mid(strTemp(1), 1, 21), 18) & fn_StrToWordFmt(strTemp(2), 7) & fn_StrToWordFmt(strTemp(3), 11) & fn_StrToWordFmt(strTemp(4), 5) & fn_StrToWordFmt(strTemp(5), 12) & fn_StrToWordFmt(strTemp(6), 7) & strTemp(7), 80), , , 10
    If iPrint >= o_PaperHeight Then
        Page = Page + 1
        'Mark by Amy 2022/05/02 紙本原本就沒使用
        '紙本
'        If Check3(0).Value = vbChecked Then
'            Printer.NewPage
'        ElseIf Check3(1).Value = vbChecked Then
            fn_DocNewPage
'        End If
        PrintTitleCp_A4
    End If
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(strTemp(8), 11) & fn_StrToWordFmt(Replace(strTemp(1), Trim(Replace(fn_StrToWordFmt(Mid(strTemp(1), 1, 21), 18), "　", "")), ""), 18) & fn_StrToWordFmt(strTemp(9), 7) & fn_StrToWordFmt(strTemp(10), 28) & strTemp(15), 80), , , 10
End If
End Sub

Private Function GetCU08(strCU0102 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

strCU0102 = Left(strCU0102 & "00000000", 9)
'接洽人不再抓舊欄位否則以刪除仍會印出
StrSQLa = "Select pcc05 From Customer,potcustcont Where CU01='" & Mid(strCU0102, 1, 8) & "' And CU02='" & Mid(strCU0102, 9, 1) & "' and pcc01(+)=cu01 and pcc02(+)=cu127"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCU08 = "" & rsA.Fields(0).Value
Else
    GetCU08 = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Mak by Amy 2022/05/02 未使用
'取得優先權資料
Private Sub PrintPriDate(strPD01 As String, strPD02 As String, strPD03 As String, strPD04 As String)
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim strPrintText As String
'
'strPrintText = ""
''要判斷已有優先權號(CFP會先輸國內案號等有申請號後才會更新)
'StrSQLa = "Select * From PriDate, Nation Where PD07=NA01 And PD01='" & strPD01 & "' And PD02='" & strPD02 & "' And PD03 ='" & strPD03 & "' And PD04='" & strPD04 & "' AND PD05>0 Order By PD05, PD07 "
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    While Not rsA.EOF
'        strPrintText = strPrintText & ChangeTStringToTDateString(ChangeWStringToTString(rsA("PD05").Value)) & " " & rsA("NA03").Value & ","
'        rsA.MoveNext
'    Wend
'    '若有優先權資料
'    If strPrintText <> "" Then
'        strPrintText = Left(strPrintText, Len(strPrintText) - 1)
'        '紙本
'        If Check3(0).Value = vbChecked Then
'            Printer.CurrentX = PLeft(0)
'            Printer.CurrentY = iPrint
'            Printer.Print "[優先權日]：" & strPrintText
'            iPrint = iPrint + 300
'        ElseIf Check3(1).Value = vbChecked Then
'            fn_PutWords fn_StrToWordFmt("[優先權日]：" & strPrintText, 80), , , 10
'        End If
'    End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
End Sub

Private Sub PrintTitleCo_A4()
'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'紙本
'If Check3(0).Value = vbChecked Then
'    GetPleftCo_A4
'    iPrint = 500
'    Printer.Font.Size = 22
'    Printer.Font.Bold = True
'    Printer.Font.Underline = True
'    Printer.FontName = "細明體"
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "客戶案件預算預估表"
'    Printer.Font.Bold = False
'    Printer.Font.Underline = False
'    iPrint = iPrint + 500
'    Printer.Font.Size = 16
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "　　　　（其他）"
'    Printer.Font.Size = 10
'    'Add By Sindy 2012/10/31
'    iPrint = iPrint + 500
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所期限：" & ChangeTStringToTDateString(txt1(16)) & " - " & ChangeTStringToTDateString(txt1(17))
'    '2012/10/31 End
'    iPrint = iPrint + 500
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "收件人："
'    If Len(StrTemp5(12)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(12)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(13)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(13)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(0)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(0)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(1)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(1)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(2)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(2)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(3)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(3)
'       iPrint = iPrint + 300
'    End If
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint - 300
'
'    Printer.Print "電話：" & StrTemp5(5)
'
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "地址：" & StrTemp5(4)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "E-mail：" & StrTemp5(7)
'
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "Page：" & str(Page)
'    iPrint = iPrint + 300
'    Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'    Printer.CurrentY = iPrint
'    Printer.Print StrTemp5(6)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "傳真：" & StrTemp5(9)
'
'    iPrint = iPrint + 300
'    Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'    Printer.CurrentY = iPrint
'    Printer.Print StrTemp5(8)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "預設接洽人：" & GetCU08(strTemp(20))     '2019/8/29加預設二字
'    Printer.CurrentX = PLeft(4)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "智權人員：" & Me.txt1(11).Text & " " & GetStaffName(Me.txt1(11).Text, True)
'
'    iPrint = iPrint + 300
'    If Len(StrTemp5(10)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(10)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(11)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(11)
'       iPrint = iPrint + 300
'    End If
'    Printer.Line (0, iPrint)-(16000, iPrint)
'    iPrint = iPrint + 100
'
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所案號"
'    Printer.CurrentX = PLeft(1)
'    Printer.CurrentY = iPrint
'    Printer.Print "案件名稱"
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請日"
'    Printer.CurrentX = PLeft(3)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請案號"
'    Printer.CurrentX = PLeft(4)
'    Printer.CurrentY = iPrint
'    Printer.Print "種類"
'    Printer.CurrentX = PLeft(5)
'    Printer.CurrentY = iPrint
'    Printer.Print "專　用　期　限"
'    Printer.CurrentX = PLeft(6)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所期限"
'    Printer.CurrentX = PLeft(7)
'    Printer.CurrentY = iPrint
'    Printer.Print "預估費用"
'    iPrint = iPrint + 300
'
'    Printer.CurrentX = PLeft(8)
'    Printer.CurrentY = iPrint
'    Printer.Print "客戶案件案號"
'    Printer.CurrentX = PLeft(9)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請國家"
'    Printer.CurrentX = PLeft(10)
'    Printer.CurrentY = iPrint
'    Printer.Print "審定號數"
'    Printer.CurrentX = PLeft(12)
'    Printer.CurrentY = iPrint
'    Printer.Print "條碼廠商號碼"
'    Printer.CurrentX = PLeft(15)
'    Printer.CurrentY = iPrint
'    Printer.Print "下一程序"
'    iPrint = iPrint + 300
'    Printer.Line (0, iPrint)-(16000, iPrint)
'    iPrint = iPrint + 100
'ElseIf Check3(1).Value = vbChecked Then
'Excel
If intChoose = 2 Then
Else
'end 2022/05/02
    If blnWordNewPage = True And Page <> 1 Then
        fn_DocNewPage
    End If
    blnWordNewPage = False
    Dim isPrintDate As Boolean
    Dim o_tmp1 As String
    Dim o_tmp2 As String
    Dim o_tmp3 As String
    Dim o_tmp4 As String
    Dim o_tmp5 As String
    Dim o_tmp6 As String
    iPrint = 0
    o_tmp1 = StrTemp5(12)
    o_tmp2 = StrTemp5(13)
    o_tmp3 = StrTemp5(0)
    o_tmp4 = StrTemp5(1)
    o_tmp5 = StrTemp5(2)
    o_tmp6 = StrTemp5(3)
    If Len(CheckStr(o_tmp6)) <> 0 Then
        If Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            o_tmp6 = fn_StrToWordFmt("收件人：" & o_tmp6, 29) & "電話：" & StrTemp5(5)
        Else
            o_tmp6 = fn_StrToWordFmt("　　　：" & o_tmp6, 29) & "電話：" & StrTemp5(5)
        End If
    End If
    If Len(CheckStr(o_tmp5)) <> 0 Then
        If Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) = 0 Then
                o_tmp5 = fn_StrToWordFmt("收件人：" & o_tmp5, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp5 = fn_StrToWordFmt("收件人：" & o_tmp5, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) = 0 Then
                o_tmp5 = fn_StrToWordFmt("　　　：" & o_tmp5, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp5 = fn_StrToWordFmt("　　　：" & o_tmp5, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp4)) <> 0 Then
        If Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) = 0 Then
                o_tmp4 = fn_StrToWordFmt("收件人：" & o_tmp4, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp4 = fn_StrToWordFmt("收件人：" & o_tmp4, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) = 0 Then
                o_tmp4 = fn_StrToWordFmt("　　　：" & o_tmp4, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp4 = fn_StrToWordFmt("　　　：" & o_tmp4, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp3)) <> 0 Then
        If Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) = 0 Then
                o_tmp3 = fn_StrToWordFmt("收件人：" & o_tmp3, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp3 = fn_StrToWordFmt("收件人：" & o_tmp3, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) = 0 Then
                o_tmp3 = fn_StrToWordFmt("　　　：" & o_tmp3, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp3 = fn_StrToWordFmt("　　　：" & o_tmp3, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp2)) <> 0 Then
        If Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) = 0 Then
                o_tmp2 = fn_StrToWordFmt("收件人：" & o_tmp2, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp2 = fn_StrToWordFmt("收件人：" & o_tmp2, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) = 0 Then
                o_tmp2 = fn_StrToWordFmt("　　　：" & o_tmp2, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp2 = fn_StrToWordFmt("　　　：" & o_tmp2, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp1)) <> 0 Then
        If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) = 0 Then
            o_tmp1 = fn_StrToWordFmt("收件人：" & o_tmp1, 29) & "電話：" & StrTemp5(5)
        Else
            o_tmp1 = fn_StrToWordFmt("收件人：" & o_tmp1, 29)
        End If
    End If
    isPrintDate = False
    fn_PutWords "客戶案件預算預估表", wdAlignParagraphCenter, "細明體", 22, False, True, True
    fn_PutWords "（其他）", wdAlignParagraphCenter, , 16
    'Add By Sindy 2012/10/31 69=>64
    fn_PutWords fn_StrToWordFmt(" ", 64) & "本所期限：" & ChangeTStringToTDateString(txt1(16)) & " - " & ChangeTStringToTDateString(txt1(17)), , , 10
    '2012/10/31 End
    If Len(o_tmp1) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp1, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp1, 80), , , 10
        End If
    End If
    If Len(o_tmp2) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp2, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp2, 80), , , 10
        End If
    End If
    If Len(o_tmp3) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp3, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp3, 80), , , 10
        End If
    End If
    If Len(o_tmp4) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp4, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp4, 80), , , 10
        End If
    End If
    If Len(o_tmp5) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp5, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp5, 80), , , 10
        End If
    End If
    If Len(o_tmp6) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp6, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp6, 80), , , 10
        End If
    End If
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("地址：" & StrTemp5(4), 29) & "E-mail：" & StrTemp5(7), 64) & "Page：" & str(Page), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("　　　" & StrTemp5(6), 29) & "傳真：" & StrTemp5(9), 64), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(fn_StrToWordFmt("　　　" & StrTemp5(8), 29) & "預設接洽人：" & GetCU08(strTemp(20)), 47.5) & "智權人員：" & Me.txt1(11).Text & " " & GetStaffName(Me.txt1(11).Text, True), 64), , , 10
    If Len(StrTemp5(10)) <> 0 Then
        fn_PutWords fn_StrToWordFmt("　　　" & StrTemp5(10), 29), , , 10
    End If
    If Len(StrTemp5(11)) <> 0 Then
        fn_PutWords fn_StrToWordFmt("　　　" & StrTemp5(11), 29), , , 10
    End If
    fn_PutWords fn_StrLineToWord("="), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("本所案號", 11) & fn_StrToWordFmt("案件名稱", 18) & fn_StrToWordFmt("申請日", 7) & fn_StrToWordFmt("申請案號", 11) & fn_StrToWordFmt("種類", 5) & fn_StrToWordFmt("專　用　期　限", 12) & fn_StrToWordFmt("本所期限", 7) & "預估費用", 80), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("客戶案件案號", 29) & fn_StrToWordFmt("申請國家", 7) & fn_StrToWordFmt("審定號數", 16) & fn_StrToWordFmt("條碼廠商號碼", 12) & "下一程序", 80), , , 10
    fn_PutWords fn_StrLineToWord("="), , , 10
End If
End Sub

Private Sub PrintDataCo_A4()
Dim strCust As String '記錄申請人
Dim strSystemKind As String '記錄系統類別
Dim blnFirstPage As Boolean '判斷是否第一頁
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCaseNo As String '申請人+本所案號

'控制 word 換頁用
blnWordNewPage = True

On Error GoTo ErrHnd

'列印第一頁
blnFirstPage = True
'國內外分開列印(不同系統類別分開列印)
If txt1(8) = "Y" Then
    '修改排序方式申請人, 系統類別
    strSql = "SELECT DISTINCT R013001,R013019 FROM R050317_C WHERE ID='" & strUserNum & "' GROUP BY R013001,R013019 ORDER BY R013001,R013019 "
'國內外不分開列印(不管系統類別)
Else
    strSql = "SELECT DISTINCT R013001 FROM R050317_C WHERE ID='" & strUserNum & "' GROUP BY R013001 ORDER BY R013001 "
End If
CheckOC
'頁數加一
Page = Page + 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    RefreshColData 1, "申請人"
    With adoRecordset
        .MoveFirst
        '記錄申請人
        strCust = "" & .Fields("R013001").Value
        '記錄系統類別
        If txt1(8) = "Y" Then
            strSystemKind = "" & .Fields("R013019").Value
        Else
            strSystemKind = ""
        End If
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields("R013001").Value)
            If txt1(8) = "Y" Then
               strTemp(21) = CheckStr(.Fields("R013019").Value)
            Else
               strTemp(21) = ""
            End If
            strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU16||DECODE(CU17,NULL,'',','||CU17),CU20,CU18||DECODE(CU19,NULL,'',','||CU19),CU58,CU61 FROM CUSTOMER WHERE CU01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND CU02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
            Else
                For i = 0 To 13
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            '若申請人不同則跳頁
            If strCust <> strTemp(20) Then
                strCust = strTemp(20)
                strSystemKind = strTemp(21)
                If txt1(8) = "Y" Then
                   strSystemKind = strTemp(21)
                Else
                   strSystemKind = ""
                End If
                Page = Page + 1
                'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'                '紙本
'                If Check3(0).Value = vbChecked Then
'                    Printer.NewPage
'                ElseIf Check3(1).Value = vbChecked Then
'                    fn_DocNewPage
'                End If
                'Excel
                If intChoose = 2 Then
                    If Page <> 1 Then fn_SetExcel '換頁前設定
                    fn_CreateExcel Page, strSystemKind, blnFirstPage
                Else
                    fn_DocNewPage
                End If
                'end 2022/05/02
                PrintTitleCo_A4
            '若申請人相同
            Else
                '記錄系統類別
                If txt1(8) = "Y" Then
                    '若系統類別不同
                    If strSystemKind <> strTemp(21) Then
                        strSystemKind = strTemp(21)
                        Page = Page + 1
                        'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'                        '紙本
'                        If Check3(0).Value = vbChecked Then
'                            Printer.NewPage
'                        ElseIf Check3(1).Value = vbChecked Then
'                            fn_DocNewPage
'                        End If
                        'Excel
                        If intChoose = 2 Then
                            If Page <> 1 Then fn_SetExcel '換頁前設定
                            fn_CreateExcel Page, strSystemKind, blnFirstPage
                        Else
                            fn_DocNewPage
                        End If
                        'end 2022/05/02
                        PrintTitleCo_A4
                    End If
                End If
            End If
            '若為第一頁才印表頭
            If blnFirstPage = True Then
                'Modify by Amy 2022/05/02 +Excel
                If intChoose = 2 Then
                    fn_CreateExcel Page, strSystemKind, blnFirstPage
                Else
                    PrintTitleCo_A4
                End If
                blnFirstPage = False
            End If
            '國內外是否分開列印
            If txt1(8) = "Y" Then
               strSql = "select R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, CPMName from ("
                         strSql = strSql & "SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),R013006) as CPMName FROM R050317_C,caseprogress,CASEPROPERTYMAP,SERVICEPRACTICE WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' and R013022=cpm01(+) and R013006=cpm02(+) and R013022=sp01 and R013023=sp02 and R013024=sp03 and R013025=sp04 "
               strSql = strSql & "union all SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),R013006) as CPMName FROM R050317_C,caseprogress,CASEPROPERTYMAP,LAWCASE         WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' and R013022=cpm01(+) and R013006=cpm02(+) and R013022=lc01 and R013023=lc02 and R013024=lc03 and R013025=lc04 "
               strSql = strSql & "union all SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, NVL(Nvl(cpm03,Nvl(CPM10,CPM13)),R013006) as CPMName                          FROM R050317_C,caseprogress,CASEPROPERTYMAP,HIRECASE        WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' and R013022=cpm01(+) and R013006=cpm02(+) and R013022=hc01 and R013023=hc02 and R013024=hc03 and R013025=hc04 "
               strSql = strSql & ")"
            Else
               strSql = "select R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, CPMName from ("
                         strSql = strSql & "SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),R013006) as CPMName FROM R050317_C,caseprogress,CASEPROPERTYMAP,SERVICEPRACTICE WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' and R013022=cpm01(+) and R013006=cpm02(+) and R013022=sp01 and R013023=sp02 and R013024=sp03 and R013025=sp04 "
               strSql = strSql & "union all SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),R013006) as CPMName FROM R050317_C,caseprogress,CASEPROPERTYMAP,LAWCASE         WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' and R013022=cpm01(+) and R013006=cpm02(+) and R013022=lc01 and R013023=lc02 and R013024=lc03 and R013025=lc04 "
               strSql = strSql & "union all SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027, NVL(Nvl(cpm03,Nvl(CPM10,CPM13)),R013006) as CPMName                          FROM R050317_C,caseprogress,CASEPROPERTYMAP,HIRECASE        WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' and R013022=cpm01(+) and R013006=cpm02(+) and R013022=hc01 and R013023=hc02 and R013024=hc03 and R013025=hc04 "
               strSql = strSql & ")"
            End If
            '依本所案號排序
            If Me.txt1(9).Text = "1" Then
               strSql = strSql & " ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010),r013015, Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '依案件名稱排序
            ElseIf Me.txt1(9).Text = "2" Then
               strSql = strSql & " ORDER BY r013015,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '依申請國家+本所案號排序
            ElseIf Me.txt1(9).Text = "3" Then
               strSql = strSql & " ORDER BY R013005, DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010),r013015, Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            'Add By Sindy 2013/12/4
            '依客戶案件案號排序
            ElseIf Me.txt1(9).Text = "5" Then
               strSql = strSql & " ORDER BY R013009,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '2013/12/4 END
            '依申請國家+案件名稱排序
            Else
               strSql = strSql & " ORDER BY R013005, r013015,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            End If
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    strTemp(0) = "" & adoRecordset1.Fields(0).Value '本所案號
                    strTemp(1) = StrToStr("" & adoRecordset1.Fields(1).Value, 42) '案件名稱
                    strTemp(2) = "" & adoRecordset1.Fields(2).Value '申請日
                    strTemp(3) = StrToStr("" & adoRecordset1.Fields(3).Value, 10) '申請案號
                    strTemp(4) = StrToStr("" & adoRecordset1.Fields(4).Value, 3) '種類
                    strTemp(5) = "" '專用期限
                    strTemp(6) = "" & adoRecordset1.Fields("R013008").Value '下一程序的本所期限
                    strTemp(7) = "" & Format(adoRecordset1.Fields("R013027").Value, "###,##0") '預估費用
                    strTemp(8) = StrToStr("" & adoRecordset1.Fields(8).Value, 10) '客戶案件案號
                    strTemp(9) = StrToStr(GetNationName("" & adoRecordset1.Fields(9).Value, 0), 6) '申請國家
                    strTemp(10) = StrToStr("" & adoRecordset1.Fields(10).Value, 10) '審定號數
                    strTemp(11) = ""
                    strTemp(12) = "" '條碼廠商號碼
                    strTemp(13) = ""
                    strTemp(14) = ""
                    strTemp(15) = "" & adoRecordset1.Fields("CPMName").Value '下一程序
                    GetOtherDuration "" & adoRecordset1("R013022").Value, "" & adoRecordset1("R013023").Value, "" & adoRecordset1("R013024").Value, "" & adoRecordset1("R013025").Value
                    'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
                    If intChoose = 2 Then
'*********  Excel *********
                        '依欄位順序顯示
                        '                本所案號                   客戶案件案號              商標名                     申日請日                       申請國                      申請號                審定號                       種類                          專用起日                   專用止日               條碼廠商號                 本所期限                      下一程序                  預估費用
                        strXlsData = strTemp(0) & "$$" & strTemp(8) & "$$" & strTemp(1) & "$$" & strTemp(2) & "$$" & strTemp(9) & "$$" & strTemp(3) & "$$" & strTemp(10) & "$$" & strXlsTp(4) & "$$" & strXlsTp(1) & "$$" & strXlsTp(2) & "$$" & strXlsTp(3) & "$$" & strTemp(6) & "$$" & strTemp(15) & "$$" & strTemp(7)
                        fn_PutExcel "" & adoRecordset1("R013022").Value, "" & adoRecordset1("R013023").Value, "" & adoRecordset1("R013024").Value, "" & adoRecordset1("R013025").Value
                        intXlsRow = intXlsRow + 1
'*********  End Excel *********
                    Else
'*********  Word *********
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.Line (0, iPrint)-(16000, iPrint)
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                            fn_DocNewPage
    '                        End If
                            fn_DocNewPage
                            PrintTitleCo_A4
                        End If
                        '列印明細資料
                        PrintDatilCo_A4
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_DocNewPage
    '                        End If
                            fn_DocNewPage
                            PrintTitleCo_A4
                        End If
                  
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.Line (0, iPrint)-(16000, iPrint)
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                            fn_DocNewPage
    '                        End If
                            fn_PutWords fn_StrLineToWord, , , 10
                            fn_DocNewPage
                            PrintTitleCo_A4
                        Else
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.CurrentX = PLeft(0)
    '                            Printer.CurrentY = iPrint
    '                            Printer.Print String(160, "-")
    '                            iPrint = iPrint + 300
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                        End If
                            fn_PutWords fn_StrLineToWord, , , 10
                         
                            If iPrint >= o_PaperHeight Then
                                Page = Page + 1
    '                            '紙本
    '                            If Check3(0).Value = vbChecked Then
    '                                Printer.NewPage
    '                            ElseIf Check3(1).Value = vbChecked Then
    '                                fn_DocNewPage
    '                            End If
                                fn_DocNewPage
                                PrintTitleCo_A4
                            End If
                        End If
'*********  End Word *********
                    End If
                    adoRecordset1.MoveNext
                Loop
                'Modify by Amy 2022/05/02 +if 不是Excel才run,紙本原本就沒使用
                If intChoose <> 2 Then
                    If iPrint >= o_PaperHeight Then
                        Page = Page + 1
    '                    '紙本
    '                    If Check3(0).Value = vbChecked Then
    '                        Printer.NewPage
    '                    ElseIf Check3(1).Value = vbChecked Then
    '                        fn_DocNewPage
    '                    End If
                        fn_DocNewPage
                        PrintTitleCo_A4
                    End If
                End If
                'end 2022/05/02
                
                If txt1(8) = "Y" Then
                  strSql = "SELECT DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050317_c WHERE R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' "
                Else
                  strSql = "SELECT DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050317_c WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' "
                End If
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                    IntTot = AdoRecordSet3.RecordCount
                End If
                'Modify by Amy 2022/05/02 +不是Excel才印總計,紙本原本就沒使用
'                '紙本
'                If Check3(0).Value = vbChecked Then
'                    Printer.CurrentX = PLeft(5)
'                    Printer.CurrentY = iPrint
'                    Printer.Print "總計：" & Format(IntTot, "##0")
'                    iPrint = iPrint + 300
'                ElseIf Check3(1).Value = vbChecked Then
'                    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 54.5) & "總計：" & Format(IntTot, "##0"), 64), , , 10
'                End If
                'Word
                If intChoose <> 2 Then
                    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 54.5) & "總計：" & Format(IntTot, "##0"), 64), , , 10
                End If
                'end 2022/05/02
                CheckOC3
            End If
            CheckOC2
            .MoveNext
        Loop
        'Add by Amy 3030/05/02 +Excel
        If intChoose = 2 Then
            fn_SetExcel
            fn_PutEndXls '存檔
        End If
    End With
End If
CheckOC
'Mark by Amy 2022/05/02 原本就沒使用
''紙本
'If Check3(0).Value = vbChecked Then
'   Printer.EndDoc
'End If
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'取得其他專用期限
Private Sub GetOtherDuration(strCaseNo1 As String, strCaseNo2 As String, StrCaseNo3 As String, strCaseNo4 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim arrYearPay
Dim arrYearPaySet
Dim ii As Integer

StrSQLa = "Select '', '', LC08, '' From Lawcase Where " & ChgLawcase(strCaseNo1 & strCaseNo2 & StrCaseNo3 & strCaseNo4)
StrSQLa = StrSQLa & " Union Select '', '', HC09, '' From Hirecase Where " & ChgHirecase(strCaseNo1 & strCaseNo2 & StrCaseNo3 & strCaseNo4)
StrSQLa = StrSQLa & " Union Select SP20||'', SP21||'', SP15, SP19 From Servicepractice Where " & ChgService(strCaseNo1 & strCaseNo2 & StrCaseNo3 & strCaseNo4)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify by Amy 2022/05/02 +if 使用不同變數存資料 for Excel
    If intChoose = 2 Then
        For ii = LBound(strXlsTp) To UBound(strXlsTp)
            strXlsTp(ii) = ""
        Next ii
        
        strXlsTp(1) = ChangeWStringToWDateString("" & rsA.Fields(0).Value)
        strXlsTp(2) = ChangeWStringToWDateString("" & rsA.Fields(1).Value)
        strXlsTp(3) = "" & rsA.Fields(3).Value '條碼廠商號碼
    Else
        strTemp(5) = ChangeWStringToWDateString("" & rsA.Fields(0).Value) & "-" & ChangeWStringToWDateString("" & rsA.Fields(1).Value)
        strTemp(12) = "" & rsA.Fields(3).Value '條碼廠商號碼
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Sub GetPleftCo_A4()
Erase PLeft

'第一行
PLeft(0) = 0 '本所案號
PLeft(1) = 2100 '專利名稱
PLeft(2) = 6300 '申請日
PLeft(3) = 7600 '申請案號
PLeft(4) = 9800 '種類
PLeft(5) = 10700 '專用期限
PLeft(6) = 13100 '最近期限
PLeft(7) = 14500 '預估費用
'第二行
PLeft(8) = 0 '客戶案件案號
PLeft(9) = 6300 '申請國家
PLeft(10) = 7600 '專利號數
PLeft(11) = 9800
PLeft(12) = 10700 '條碼廠商號碼
PLeft(13) = 13100
PLeft(14) = 13100 '閉卷
PLeft(15) = 13100 '下一程序
End Sub

Private Sub PrintDatilCo_A4()
'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'紙本
'If Check3(0).Value = vbChecked Then
'    For i = 0 To 7
'       If i = 0 Then
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint
'          Printer.Print Replace(strTemp(i), "*", "＊")
'       ElseIf i = 1 Then
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint
'          Printer.Print Mid(strTemp(i), 1, 21)
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint + 300
'          Printer.Print Mid(strTemp(i), 22, 21)
'       Else
'          Printer.CurrentX = PLeft(i)
'          Printer.CurrentY = iPrint
'          Printer.Print strTemp(i)
'       End If
'    Next i
'    iPrint = iPrint + 300
'    For i = 8 To 14
'       Printer.CurrentX = PLeft(i)
'       Printer.CurrentY = iPrint
'       Printer.Print strTemp(i)
'    Next i
'    If strTemp(15) <> "" Then
'       Printer.CurrentX = PLeft(15)
'       Printer.CurrentY = iPrint
'       Printer.Print strTemp(15)
'    End If
'    iPrint = iPrint + 300
'ElseIf Check3(1).Value = vbChecked Then
'Excel
If intChoose = 2 Then
    
Else
'end 2022/05/02
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(Replace(strTemp(0), "*", "＊"), 11) & fn_StrToWordFmt(Mid(strTemp(1), 1, 21), 18) & fn_StrToWordFmt(strTemp(2), 7) & fn_StrToWordFmt(strTemp(3), 11) & fn_StrToWordFmt(strTemp(4), 5) & fn_StrToWordFmt(strTemp(5), 12) & fn_StrToWordFmt(strTemp(6), 7) & strTemp(7), 80), , , 10
    If iPrint >= o_PaperHeight Then
        Page = Page + 1
        'Mark by Amy 2022/05/02 紙本原本就沒使用
'        '紙本
'        If Check3(0).Value = vbChecked Then
'            Printer.NewPage
'        ElseIf Check3(1).Value = vbChecked Then
            fn_DocNewPage
'        End If
        PrintTitleCo_A4
    End If
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(strTemp(8), 11) & fn_StrToWordFmt(Replace(strTemp(1), Trim(Replace(fn_StrToWordFmt(Mid(strTemp(1), 1, 21), 18), "　", "")), ""), 18) & fn_StrToWordFmt(strTemp(9), 7) & fn_StrToWordFmt(strTemp(10), 16) & fn_StrToWordFmt(strTemp(12), 12) & strTemp(15), 80), , , 10
End If
End Sub

'Mark by Amy 2022/05/02 不使用
Private Sub GetPleftCt_A4()
'Erase PLeft
''第一行
'PLeft(0) = 0 '本所案號
'PLeft(1) = 2100 '商標名稱
'PLeft(2) = 6300 '申請日
'PLeft(3) = 7600 '申請案號
'PLeft(4) = 9800 '種類
'PLeft(5) = 10700 '專用期限
'PLeft(6) = 13100 '最近期限
'PLeft(7) = 14500 '預估費用
''第二行
'PLeft(8) = 0 '客戶案件案號
'PLeft(9) = 6300 '申請國家
'PLeft(10) = 7600 '審定號數
'PLeft(11) = 9800 '准駁
'PLeft(12) = 10700 '商品類別
'PLeft(13) = 13100 '(年度)
'PLeft(14) = 13100 '閉卷
'PLeft(15) = 13100 '下一程序
End Sub

Private Sub PrintDataCt_A4()
Dim strCust As String '記錄申請人
Dim strSystemKind As String '記錄系統類別
Dim blnFirstPage As Boolean '判斷是否第一頁
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCaseNo As String '申請人+本所案號

'控制 word 換頁用
blnWordNewPage = True

On Error GoTo ErrHnd

'列印第一頁
blnFirstPage = True
'國內外分開列印(不同系統類別分開列印)
If txt1(8) = "Y" Then
    '修改排序方式申請人, 系統類別
    strSql = "SELECT DISTINCT R013001,R013019 FROM R050317_C WHERE ID='" & strUserNum & "' GROUP BY R013001,R013019 ORDER BY R013001,R013019 "
'國內外不分開列印(不管系統類別)
Else
    strSql = "SELECT DISTINCT R013001 FROM R050317_C WHERE ID='" & strUserNum & "' GROUP BY R013001 ORDER BY R013001 "
End If
CheckOC
'頁數加一
Page = Page + 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    RefreshColData 1, "申請人"
    With adoRecordset
        .MoveFirst
        '記錄申請人
        strCust = "" & .Fields("R013001").Value
        '記錄系統類別
        If txt1(8) = "Y" Then
            strSystemKind = "" & .Fields("R013019").Value
        Else
            strSystemKind = ""
        End If
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields("R013001").Value)
            If txt1(8) = "Y" Then
               strTemp(21) = CheckStr(.Fields("R013019").Value)
            Else
               strTemp(21) = ""
            End If
            strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU16||DECODE(CU17,NULL,'',','||CU17),CU20,CU18||DECODE(CU19,NULL,'',','||CU19),CU58,CU61 FROM CUSTOMER WHERE CU01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND CU02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
            Else
                For i = 0 To 13
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            '若申請人不同則跳頁
            If strCust <> strTemp(20) Then
                strCust = strTemp(20)
                strSystemKind = strTemp(21)
                If txt1(8) = "Y" Then
                   strSystemKind = strTemp(21)
                Else
                   strSystemKind = ""
                End If
                Page = Page + 1
                'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'                '紙本
'                If Check3(0).Value = vbChecked Then
'                    Printer.NewPage
'                ElseIf Check3(1).Value = vbChecked Then
'                    fn_DocNewPage
'                End If
                'Excel
                If intChoose = 2 Then
                    If Page <> 1 Then fn_SetExcel '換頁前設定
                    fn_CreateExcel Page, strSystemKind, blnFirstPage
                Else
                    fn_DocNewPage
                End If
                'end 2022/05/02
                PrintTitleCt_A4
            '若申請人相同
            Else
               '記錄系統類別
                If txt1(8) = "Y" Then
                    '若系統類別不同
                    If strSystemKind <> strTemp(21) Then
                        strSystemKind = strTemp(21)
                        Page = Page + 1
                        'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
'                        '紙本
'                        If Check3(0).Value = vbChecked Then
'                            Printer.NewPage
'                        ElseIf Check3(1).Value = vbChecked Then
'                            fn_DocNewPage
'                        End If
                        'Excel
                        If intChoose = 2 Then
                            If Page <> 1 Then fn_SetExcel '換頁前設定
                            fn_CreateExcel Page, strSystemKind, blnFirstPage
                        Else
                            fn_DocNewPage
                        End If
                        'end 2022/05/02
                        PrintTitleCt_A4
                    End If
                End If
            End If
            '若為第一頁才印表頭
            If blnFirstPage = True Then
                'Modify by Amy 2022/05/02 +Excel
                If intChoose = 2 Then
                    fn_CreateExcel Page, strSystemKind, blnFirstPage
                Else
                    PrintTitleCt_A4
                End If
                'end 2022/05/02
                blnFirstPage = False
            End If
            '國內外是否分開列印
            If txt1(8) = "Y" Then
               strSql = "SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, tm16 as CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027,NVL(Nvl(decode(tm10,'000',c1.cpm03,c1.cpm04),Nvl(c1.CPM10,c1.CPM13)),R013006) as CPMName, NVL(Nvl(decode(tm10,'000',c2.cpm03,c2.cpm04),Nvl(c2.CPM10,c2.CPM13)),cp10) as CPMName2 FROM R050317_C,caseprogress,trademark,CASEPROPERTYMAP c1,CASEPROPERTYMAP c2 WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and R013022=c1.cpm01(+) and R013006=c1.cpm02(+) and cp01=c2.cpm01(+) and cp10=c2.cpm02(+) "
            Else
               strSql = "SELECT R013010, R013015, R013020, R013004, R013014, R013006, '', R013008, R013009, R013005, R013011, tm16 as CP24, '', '', '', '', R013022, R013023, R013024, R013025, R013007, R013027,NVL(Nvl(decode(tm10,'000',c1.cpm03,c1.cpm04),Nvl(c1.CPM10,c1.CPM13)),R013006) as CPMName, NVL(Nvl(decode(tm10,'000',c2.cpm03,c2.cpm04),Nvl(c2.CPM10,c2.CPM13)),cp10) as CPMName2 FROM R050317_C,caseprogress,trademark,CASEPROPERTYMAP c1,CASEPROPERTYMAP c2 WHERE r013018=CP09(+) And R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and R013022=c1.cpm01(+) and R013006=c1.cpm02(+) and cp01=c2.cpm01(+) and cp10=c2.cpm02(+) "
            End If
            '依本所案號排序
            If Me.txt1(9).Text = "1" Then
               strSql = strSql & " ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010),r013015, Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '依案件名稱排序
            ElseIf Me.txt1(9).Text = "2" Then
               strSql = strSql & " ORDER BY r013015,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '依申請國家+本所案號排序
            ElseIf Me.txt1(9).Text = "3" Then
               strSql = strSql & " ORDER BY R013005, DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010),r013015, Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            'Add By Sindy 2013/12/4
            '依客戶案件案號排序
            ElseIf Me.txt1(9).Text = "5" Then
               strSql = strSql & " ORDER BY R013009,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            '2013/12/4 END
            '依申請國家+案件名稱排序
            Else
               strSql = strSql & " ORDER BY R013005, r013015,DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010), Decode(r013007,NULL,9999999,To_Number(Replace(r013007,'/',''))) "
            End If
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    strTemp(0) = "" & adoRecordset1.Fields(0).Value '本所案號
                    strTemp(1) = StrToStr("" & adoRecordset1.Fields(1).Value, 42) '案件名稱
                    strTemp(2) = "" & adoRecordset1.Fields(2).Value '申請日
                    strTemp(3) = StrToStr("" & adoRecordset1.Fields(3).Value, 10) '申請案號
                    strTemp(4) = ""
                    strTemp(5) = ""
                    strTemp(6) = "" & adoRecordset1.Fields("R013008").Value '下一程序的本所期限
                    strTemp(7) = "" & Format(adoRecordset1.Fields("R013027").Value, "###,##0") '預估費用
                    strTemp(8) = StrToStr("" & adoRecordset1.Fields(8).Value, 10) '客戶案件案號
                    strTemp(9) = StrToStr(GetNationName("" & adoRecordset1.Fields(9).Value, 0), 6) '申請國家
                    strTemp(10) = StrToStr("" & adoRecordset1.Fields(10).Value, 10) '審定號數
                    strTemp(11) = ""
                    strTemp(12) = ""
                    strTemp(13) = ""
                    strTemp(14) = ""
                    '下一程序
                    If "" & adoRecordset1.Fields("R013006").Value = "305" Then
                        strTemp(15) = "" & adoRecordset1.Fields("CPMName2").Value & "預估核准"
                    Else
                        strTemp(15) = "" & adoRecordset1.Fields("CPMName").Value
                    End If
                    GetTrademarkDuration "" & adoRecordset1("R013022").Value, "" & adoRecordset1("R013023").Value, "" & adoRecordset1("R013024").Value, "" & adoRecordset1("R013025").Value
                    'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
                    If intChoose = 2 Then
'*********  Excel *********
                        '依欄位順序顯示
                        '               本所案號                   客戶案件案號              商標名                     申日請日                       申請國                      申請號                審定號                       種類                          專用起日                   專用止日          商品類別                 本所期限                      下一程序                  預估費用
                        strXlsData = strTemp(0) & "$$" & strTemp(8) & "$$" & strTemp(1) & "$$" & strTemp(2) & "$$" & strTemp(9) & "$$" & strTemp(3) & "$$" & strTemp(10) & "$$" & strXlsTp(4) & "$$" & strXlsTp(1) & "$$" & strXlsTp(2) & "$$" & strXlsTp(3) & "$$" & strTemp(6) & "$$" & strTemp(15) & "$$" & strTemp(7)
                        fn_PutExcel "" & adoRecordset1("R013022").Value, "" & adoRecordset1("R013023").Value, "" & adoRecordset1("R013024").Value, "" & adoRecordset1("R013025").Value
                        intXlsRow = intXlsRow + 1
'*********  End Excel *********
                    Else
'*********  Word *********
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.Line (0, iPrint)-(16000, iPrint)
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                            fn_DocNewPage
    '                        End If
                            fn_PutWords fn_StrLineToWord, , , 10
                            fn_DocNewPage
                            PrintTitleCt_A4
                        End If
                        '列印明細資料
                        PrintDatilCt_A4
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_DocNewPage
    '                        End If
                            fn_DocNewPage
                            PrintTitleCt_A4
                        End If
                        If iPrint >= o_PaperHeight Then
                            Page = Page + 1
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.Line (0, iPrint)-(16000, iPrint)
    '                            Printer.NewPage
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                            fn_DocNewPage
    '                        End If
                             fn_PutWords fn_StrLineToWord, , , 10
                             fn_DocNewPage
                             PrintTitleCt_A4
                        Else
    '                        '紙本
    '                        If Check3(0).Value = vbChecked Then
    '                            Printer.CurrentX = PLeft(0)
    '                            Printer.CurrentY = iPrint
    '                            Printer.Print String(160, "-")
    '                            iPrint = iPrint + 300
    '                        ElseIf Check3(1).Value = vbChecked Then
    '                            fn_PutWords fn_StrLineToWord, , , 10
    '                        End If
                            fn_PutWords fn_StrLineToWord, , , 10
                           
                            If iPrint >= o_PaperHeight Then
                                Page = Page + 1
    '                            '紙本
    '                            If Check3(0).Value = vbChecked Then
    '                                Printer.NewPage
    '                            ElseIf Check3(1).Value = vbChecked Then
    '                                fn_DocNewPage
    '                            End If
                                fn_DocNewPage
                                PrintTitleCt_A4
                            End If
                        End If
'*********  End Word *********
                    End If
                    adoRecordset1.MoveNext
                Loop
                'Modify by Amy 2022/05/02 +if 不是Excel才run,紙本原本就沒使用
                If intChoose <> 2 Then
                    If iPrint >= o_PaperHeight Then
                        Page = Page + 1
    '                    '紙本
    '                    If Check3(0).Value = vbChecked Then
    '                        Printer.NewPage
    '                    ElseIf Check3(1).Value = vbChecked Then
    '                        fn_DocNewPage
    '                    End If
                        fn_DocNewPage
                        PrintTitleCt_A4
                    End If
                End If
                'end 2022/05/02
                If txt1(8) = "Y" Then
                  strSql = "SELECT DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050317_c WHERE R013001='" & strTemp(20) & "' AND R013019='" & strTemp(21) & "' AND ID='" & strUserNum & "' "
                Else
                  strSql = "SELECT DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050317_c WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' "
                End If
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                    IntTot = AdoRecordSet3.RecordCount
                End If
                'Modify by Amy 2022/05/02 +不是Excel才印總計,紙本原本就沒使用
'                '紙本
'                If Check3(0).Value = vbChecked Then
'                    Printer.CurrentX = PLeft(5)
'                    Printer.CurrentY = iPrint
'                    Printer.Print "總計：" & Format(IntTot, "##0")
'                    iPrint = iPrint + 300
'                ElseIf Check3(1).Value = vbChecked Then
'                    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 54.5) & "總計：" & Format(IntTot, "##0"), 64), , , 10
'                End If
                If intChoose <> 2 Then
                    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 54.5) & "總計：" & Format(IntTot, "##0"), 64), , , 10
                End If
                'end 2022/05/02
                CheckOC3
            End If
            CheckOC2
            .MoveNext
        Loop
        'Add by Amy 3030/05/02 +Excel
        If intChoose = 2 Then
            fn_SetExcel
            fn_PutEndXls '存檔
        End If
    End With
End If
CheckOC
'Mark by Amy 2022/05/02 原本就沒使用
''紙本
'If Check3(0).Value = vbChecked Then
'   Printer.EndDoc
'End If
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'取得商標專用期限
Private Sub GetTrademarkDuration(strCaseNo1 As String, strCaseNo2 As String, StrCaseNo3 As String, strCaseNo4 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim arrYearPay
Dim arrYearPaySet
Dim ii As Integer

StrSQLa = "Select TM21, TM22, TM29, TM09 ," & strIdfTag & " || PTM03 From Trademark, PatentTrademarkMap Where '2'=PTM01(+) And TM08=PTM02(+) And " & ChgTradeMark(strCaseNo1 & strCaseNo2 & StrCaseNo3 & strCaseNo4)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify by Amy 2022/05/02 +if 使用不同變數存資料 for Excel
    If intChoose = 2 Then
        For ii = LBound(strXlsTp) To UBound(strXlsTp)
            strXlsTp(ii) = ""
        Next ii
        
        strXlsTp(1) = ChangeWStringToWDateString("" & rsA.Fields("TM21").Value)
        strXlsTp(2) = ChangeWStringToWDateString("" & rsA.Fields("TM22").Value)
        strXlsTp(3) = "" & rsA.Fields("TM09").Value '商品類別
        strXlsTp(4) = StrToStr("" & rsA.Fields(4).Value, 4) '種類
    Else
        strTemp(5) = ChangeWStringToWDateString("" & rsA.Fields(0).Value) & "-" & ChangeWStringToWDateString("" & rsA.Fields(1).Value)
        strTemp(12) = "" & rsA.Fields(3).Value
        strTemp(4) = "" & rsA.Fields(4).Value
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Sub PrintDatilCt_A4()
'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset
'紙本
'If Check3(0).Value = vbChecked Then
'    For i = 0 To 7
'        If i = 0 Then
'            Printer.CurrentX = PLeft(i)
'            Printer.CurrentY = iPrint
'            Printer.Print Replace(strTemp(i), "*", "＊")
'        ElseIf i = 1 Then
'            StrSQLa = "Select * From Trademark Where " & ChgTradeMark(Replace(Replace(strTemp(0), "-", ""), "*", ""))
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'                '中文名稱
'                Printer.CurrentX = PLeft(i)
'                Printer.CurrentY = iPrint
'                Printer.Print Left("" & rsA("TM05").Value, 20)
'                '英文名稱
'                Printer.CurrentX = PLeft(i)
'                Printer.CurrentY = iPrint + 300
'                Printer.Print Mid("" & rsA("TM05").Value, 21, 20)
'            Else
'                Printer.CurrentX = PLeft(i)
'                Printer.CurrentY = iPrint
'                Printer.Print Mid(strTemp(i), 1, 21)
'                Printer.CurrentX = PLeft(i)
'                Printer.CurrentY = iPrint + 300
'                Printer.Print Mid(strTemp(i), 22, 21)
'            End If
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'        Else
'            Printer.CurrentX = PLeft(i)
'            Printer.CurrentY = iPrint
'            Printer.Print strTemp(i)
'        End If
'    Next i
'    iPrint = iPrint + 300
'    For i = 8 To 14
'        Printer.CurrentX = PLeft(i)
'        Printer.CurrentY = iPrint
'        Printer.Print strTemp(i)
'    Next i
'    If strTemp(15) <> "" Then
'        Printer.CurrentX = PLeft(15)
'        Printer.CurrentY = iPrint
'        Printer.Print strTemp(15)
'    End If
'    iPrint = iPrint + 300
'ElseIf Check3(1).Value = vbChecked Then
'Excel
If intChoose = 2 Then
Else
'end 2022/05/02
    If rsA.State = 1 Then rsA.Close
    StrSQLa = "Select * From Trademark Where " & ChgTradeMark(Replace(Replace(strTemp(0), "-", ""), "*", ""))
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        '中文名稱
        strTemp(1) = CheckStr(rsA("tm05").Value)
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(Replace(strTemp(0), "*", "＊"), 11) & fn_StrToWordFmt(Mid(strTemp(1), 1, 21), 18) & fn_StrToWordFmt(strTemp(2), 7) & fn_StrToWordFmt(strTemp(3), 11) & fn_StrToWordFmt(strTemp(4), 5) & fn_StrToWordFmt(strTemp(5), 12) & fn_StrToWordFmt(strTemp(6), 7) & strTemp(7), 80), , , 10
    If iPrint >= o_PaperHeight Then
        Page = Page + 1
        'Mark by Amy 2022/05/02 紙本原本就沒使用
'        '紙本
'        If Check3(0).Value = vbChecked Then
'            Printer.NewPage
'        ElseIf Check3(1).Value = vbChecked Then
            fn_DocNewPage
'        End If
        PrintTitleCt_A4
    End If
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(strTemp(8), 11) & fn_StrToWordFmt(Replace(strTemp(1), Trim(Replace(fn_StrToWordFmt(Mid(strTemp(1), 1, 21), 18), "　", "")), ""), 18) & fn_StrToWordFmt(strTemp(9), 7) & fn_StrToWordFmt(strTemp(10), 16) & fn_StrToWordFmt(strTemp(12), 12) & strTemp(15), 80), , , 10
End If
End Sub

Private Sub PrintTitleCt_A4()
'Modify by Amy 2022/05/02 +Excel,紙本原本就沒使用
''紙本
'If Check3(0).Value = vbChecked Then
'    GetPleftCt_A4
'    iPrint = 500
'    Printer.Font.Size = 22
'    Printer.Font.Bold = True
'    Printer.Font.Underline = True
'    Printer.FontName = "細明體"
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "客戶案件預算預估表"
'    Printer.Font.Bold = False
'    Printer.Font.Underline = False
'    iPrint = iPrint + 500
'    Printer.Font.Size = 16
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "　　　　（商標）"
'    Printer.Font.Size = 10
'    'Add By Sindy 2012/10/31
'    iPrint = iPrint + 500
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所期限：" & ChangeTStringToTDateString(txt1(16)) & " - " & ChangeTStringToTDateString(txt1(17))
'    '2012/10/31 End
'    iPrint = iPrint + 500
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "收件人："
'    If Len(StrTemp5(12)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(12)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(13)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(13)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(0)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(0)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(1)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(1)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(2)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(2)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(3)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("收件人：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(3)
'       iPrint = iPrint + 300
'    End If
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint - 300
'
'    Printer.Print "電話：" & StrTemp5(5)
'
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "地址：" & StrTemp5(4)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "E-mail：" & StrTemp5(7)
'
'    Printer.CurrentX = PLeft(13)
'    Printer.CurrentY = iPrint
'    Printer.Print "Page：" & str(Page)
'    iPrint = iPrint + 300
'    Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'    Printer.CurrentY = iPrint
'    Printer.Print StrTemp5(6)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "傳真：" & StrTemp5(9)
'
'    iPrint = iPrint + 300
'    Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'    Printer.CurrentY = iPrint
'    Printer.Print StrTemp5(8)
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "預設接洽人：" & GetCU08(strTemp(20))    '2019/8/29加預設二字
'    Printer.CurrentX = PLeft(4)
'    Printer.CurrentY = iPrint
'
'    Printer.Print "智權人員：" & Me.txt1(11).Text & " " & GetStaffName(Me.txt1(11).Text, True)
'
'    iPrint = iPrint + 300
'
'    Printer.CurrentX = PLeft(4)
'    Printer.CurrentY = iPrint
'    Printer.Print "A：原為聯合商標　B:原為服務標章　C:原為聯合服務標章"
'
'    iPrint = iPrint + 300
'    If Len(StrTemp5(10)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(10)
'       iPrint = iPrint + 300
'    End If
'    If Len(StrTemp5(11)) <> 0 Then
'       Printer.CurrentX = PLeft(0) + Printer.TextWidth("地址：")
'       Printer.CurrentY = iPrint
'       Printer.Print StrTemp5(11)
'       iPrint = iPrint + 300
'    End If
'    Printer.Line (0, iPrint)-(16000, iPrint)
'    iPrint = iPrint + 101
'
'    Printer.CurrentX = PLeft(0)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所案號"
'    Printer.CurrentX = PLeft(1)
'    Printer.CurrentY = iPrint
'    Printer.Print "商標名稱"
'    Printer.CurrentX = PLeft(2)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請日"
'    Printer.CurrentX = PLeft(3)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請案號"
'    Printer.CurrentX = PLeft(4) + 100
'    Printer.CurrentY = iPrint
'    Printer.Print "種類"
'    Printer.CurrentX = PLeft(5)
'    Printer.CurrentY = iPrint
'    Printer.Print "專　用　期　限"
'    Printer.CurrentX = PLeft(6)
'    Printer.CurrentY = iPrint
'    Printer.Print "本所期限"
'    Printer.CurrentX = PLeft(7)
'    Printer.CurrentY = iPrint
'    Printer.Print "預估費用"
'    iPrint = iPrint + 300
'
'    Printer.CurrentX = PLeft(8)
'    Printer.CurrentY = iPrint
'    Printer.Print "客戶案件案號"
'    Printer.CurrentX = PLeft(9)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請國家"
'    Printer.CurrentX = PLeft(10)
'    Printer.CurrentY = iPrint
'    Printer.Print "審定號數"
'    Printer.CurrentX = PLeft(12)
'    Printer.CurrentY = iPrint
'    Printer.Print "商品類別"
'    Printer.CurrentX = PLeft(15)
'    Printer.CurrentY = iPrint
'    Printer.Print "下一程序"
'    iPrint = iPrint + 300
'    Printer.Line (0, iPrint)-(16000, iPrint)
'    iPrint = iPrint + 100
'ElseIf Check3(1).Value = vbChecked Then
'Excel
If intChoose = 2 Then
Else
'end 2022/05/02
    If blnWordNewPage = True And Page <> 1 Then
        fn_DocNewPage
    End If
    blnWordNewPage = False
    Dim isPrintDate As Boolean
    Dim o_tmp1 As String
    Dim o_tmp2 As String
    Dim o_tmp3 As String
    Dim o_tmp4 As String
    Dim o_tmp5 As String
    Dim o_tmp6 As String
    iPrint = 0
    o_tmp1 = StrTemp5(12)
    o_tmp2 = StrTemp5(13)
    o_tmp3 = StrTemp5(0)
    o_tmp4 = StrTemp5(1)
    o_tmp5 = StrTemp5(2)
    o_tmp6 = StrTemp5(3)
    If Len(CheckStr(o_tmp6)) <> 0 Then
        If Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            o_tmp6 = fn_StrToWordFmt("收件人：" & o_tmp6, 29) & "電話：" & StrTemp5(5)
        Else
            o_tmp6 = fn_StrToWordFmt("　　　：" & o_tmp6, 29) & "電話：" & StrTemp5(5)
        End If
    End If
    If Len(CheckStr(o_tmp5)) <> 0 Then
        If Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) = 0 Then
                o_tmp5 = fn_StrToWordFmt("收件人：" & o_tmp5, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp5 = fn_StrToWordFmt("收件人：" & o_tmp5, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) = 0 Then
                o_tmp5 = fn_StrToWordFmt("　　　：" & o_tmp5, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp5 = fn_StrToWordFmt("　　　：" & o_tmp5, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp4)) <> 0 Then
        If Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) = 0 Then
                o_tmp4 = fn_StrToWordFmt("收件人：" & o_tmp4, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp4 = fn_StrToWordFmt("收件人：" & o_tmp4, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) = 0 Then
                o_tmp4 = fn_StrToWordFmt("　　　：" & o_tmp4, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp4 = fn_StrToWordFmt("　　　：" & o_tmp4, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp3)) <> 0 Then
        If Len(CheckStr(o_tmp2)) + Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) = 0 Then
                o_tmp3 = fn_StrToWordFmt("收件人：" & o_tmp3, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp3 = fn_StrToWordFmt("收件人：" & o_tmp3, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) = 0 Then
                o_tmp3 = fn_StrToWordFmt("　　　：" & o_tmp3, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp3 = fn_StrToWordFmt("　　　：" & o_tmp3, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp2)) <> 0 Then
        If Len(CheckStr(o_tmp1)) = 0 Then
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) = 0 Then
                o_tmp2 = fn_StrToWordFmt("收件人：" & o_tmp2, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp2 = fn_StrToWordFmt("收件人：" & o_tmp2, 29)
            End If
        Else
            If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) = 0 Then
                o_tmp2 = fn_StrToWordFmt("　　　：" & o_tmp2, 29) & "電話：" & StrTemp5(5)
            Else
                o_tmp2 = fn_StrToWordFmt("　　　：" & o_tmp2, 29)
            End If
        End If
    End If
    If Len(CheckStr(o_tmp1)) <> 0 Then
        If Len(CheckStr(o_tmp6)) + Len(CheckStr(o_tmp5)) + Len(CheckStr(o_tmp4)) + Len(CheckStr(o_tmp3)) + Len(CheckStr(o_tmp2)) = 0 Then
            o_tmp1 = fn_StrToWordFmt("收件人：" & o_tmp1, 29) & "電話：" & StrTemp5(5)
        Else
            o_tmp1 = fn_StrToWordFmt("收件人：" & o_tmp1, 29)
        End If
    End If
    isPrintDate = False
    fn_PutWords "客戶案件預算預估表", wdAlignParagraphCenter, "細明體", 22, False, True, True
    fn_PutWords "（商標）", wdAlignParagraphCenter, , 16
    'Add By Sindy 2012/10/31 69=>64
    fn_PutWords fn_StrToWordFmt(" ", 64) & "本所期限：" & ChangeTStringToTDateString(txt1(16)) & " - " & ChangeTStringToTDateString(txt1(17)), , , 10
    '2012/10/31 End
    If Len(o_tmp1) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp1, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp1, 80), , , 10
        End If
    End If
    If Len(o_tmp2) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp2, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp2, 80), , , 10
        End If
    End If
    If Len(o_tmp3) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp3, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp3, 80), , , 10
        End If
    End If
    If Len(o_tmp4) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp4, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp4, 80), , , 10
        End If
    End If
    If Len(o_tmp5) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp5, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp5, 80), , , 10
        End If
    End If
    If Len(o_tmp6) <> 0 Then
        If isPrintDate = False Then
            fn_PutWords fn_StrToWordFmt(o_tmp6, 64) & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)), , , 10
            isPrintDate = True
        Else
            fn_PutWords fn_StrToWordFmt(o_tmp6, 80), , , 10
        End If
    End If
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("地址：" & StrTemp5(4), 29) & "E-mail：" & StrTemp5(7), 64) & "Page：" & str(Page), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("　　　" & StrTemp5(6), 29) & "傳真：" & StrTemp5(9), 64), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt(fn_StrToWordFmt("　　　" & StrTemp5(8), 29) & "預設接洽人：" & GetCU08(strTemp(20)), 47.5) & "智權人員：" & Me.txt1(11).Text & " " & GetStaffName(Me.txt1(11).Text, True), 64), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("", 47) & "A：原為聯合商標　B:原為服務標章　C:原為聯合服務標章", 80), , , 10
    If Len(StrTemp5(10)) <> 0 Then
        fn_PutWords fn_StrToWordFmt("　　　" & StrTemp5(10), 29), , , 10
    End If
    If Len(StrTemp5(11)) <> 0 Then
        fn_PutWords fn_StrToWordFmt("　　　" & StrTemp5(11), 29), , , 10
    End If
    fn_PutWords fn_StrLineToWord("="), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("本所案號", 11) & fn_StrToWordFmt("商標名稱", 18) & fn_StrToWordFmt("申請日", 7) & fn_StrToWordFmt("申請案號", 11) & fn_StrToWordFmt("種類", 5) & fn_StrToWordFmt("專　用　期　限", 12) & fn_StrToWordFmt("本所期限", 7) & "預估費用", 80), , , 10
    fn_PutWords fn_StrToWordFmt(fn_StrToWordFmt("客戶案件案號", 29) & fn_StrToWordFmt("申請國家", 7) & fn_StrToWordFmt("審定號數", 16) & fn_StrToWordFmt("商品類別", 12) & "下一程序", 80), , , 10
    fn_PutWords fn_StrLineToWord("="), , , 10
End If
End Sub

'依照不同紙張建立 word 文件
Sub fn_CreateWord(Optional IsA4 As Boolean = True)
   Dim oldtmpOr
   Dim Prs As Integer
   Dim oldPrItem As Integer
On Error GoTo ErrHnd
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    
    g_WordAp.Documents.add
    g_WordAp.Visible = True
    g_WordAp.Application.WindowState = wdWindowStateMinimize
    g_WordAp.Application.Selection.Font.Name = "標楷體"
   'US 48 行
    If IsA4 = False Then
        tmpPrName = ""
        oldtmpPrName = Printer.DeviceName
        oldtmpOr = Printer.Orientation
        For Prs = 0 To Printers.Count - 1
            Set Printer = Printers(Prs)
             If InStr(1, Printer.DeviceName, "7800") <> 0 Then
                tmpPrName = Printer.DeviceName
             End If
            If Printer.DeviceName = oldtmpPrName Then
                oldPrItem = Prs
            End If
        Next Prs
        Set Printer = Printers(oldPrItem)
        g_WordAp.ActivePrinter = tmpPrName
        With g_WordAp.ActiveDocument.PageSetup
            .LineNumbering.Active = False
            .Orientation = wdOrientPortrait
            .TopMargin = g_WordAp.CentimetersToPoints(0)
            .BottomMargin = g_WordAp.CentimetersToPoints(0)
            .LeftMargin = g_WordAp.CentimetersToPoints(0)
            .RightMargin = g_WordAp.CentimetersToPoints(3.25)
            .Gutter = g_WordAp.CentimetersToPoints(0)
            .HeaderDistance = g_WordAp.CentimetersToPoints(0)
            .FooterDistance = g_WordAp.CentimetersToPoints(0)
            .PageWidth = g_WordAp.CentimetersToPoints(37.78)
            .PageHeight = g_WordAp.CentimetersToPoints(27.94)
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .VerticalAlignment = wdAlignVerticalTop
            .SuppressEndnotes = False
            .MirrorMargins = False
            .TwoPagesOnOne = False
            .GutterOnTop = False
            .CharsLine = 81
            .LinesPage = 48
        End With
    '其他
    Else    'A4 橫印 34 行
        With g_WordAp.ActiveDocument.PageSetup
                .LineNumbering.Active = False
                .Orientation = wdOrientLandscape
                .TopMargin = g_WordAp.CentimetersToPoints(0.56)
                .BottomMargin = g_WordAp.CentimetersToPoints(0.67)
                .LeftMargin = g_WordAp.CentimetersToPoints(0.64)
                .RightMargin = g_WordAp.CentimetersToPoints(0.64)
                .Gutter = g_WordAp.CentimetersToPoints(0)
                .HeaderDistance = g_WordAp.CentimetersToPoints(0)
                .FooterDistance = g_WordAp.CentimetersToPoints(0)
                .PageWidth = g_WordAp.CentimetersToPoints(29.7)
                .PageHeight = g_WordAp.CentimetersToPoints(21)
                .FirstPageTray = wdPrinterDefaultBin
                .OtherPagesTray = wdPrinterDefaultBin
                .SectionStart = wdSectionNewPage
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .VerticalAlignment = wdAlignVerticalTop
                .SuppressEndnotes = False
                .MirrorMargins = False
                .TwoPagesOnOne = False
                .GutterOnTop = False
                .CharsLine = 67
                .LinesPage = 34
            End With
    End If
    g_WordAp.Application.Selection.ParagraphFormat.DisableLineHeightGrid = True
    IsOpenWord = True
    
ErrHnd:

   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91:
            g_WordAp.Documents.add
            Resume Next
         Case 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            Resume Next
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
End Sub

'將字放入 word 中
Sub fn_PutWords(oStr As String, _
                           Optional oAlignment As WdParagraphAlignment = wdAlignParagraphJustify, _
                           Optional oFontName As String = "細明體", _
                           Optional oFontSize As Integer = 12, _
                           Optional IsNewLine As Boolean = True, _
                           Optional IsUnderline As Boolean = False, _
                           Optional IsBold As Boolean = False)
If IsOpenWord = False Then
    fn_CreateWord True
End If
                           
    With g_WordAp.Application.Selection
        If IsNewLine = True Then
            .TypeParagraph
        End If
        .ParagraphFormat.Alignment = oAlignment
        .Font.Size = oFontSize
        .Font.Name = oFontName
        If IsUnderline = True Then
            If .Font.Underline = wdUnderlineNone Then
                .Font.Underline = wdUnderlineSingle
            Else
                .Font.Underline = wdUnderlineNone
            End If
        End If
        If IsBold = True Then
            .Font.Bold = wdToggle
        End If
        .TypeText Text:=oStr
        If IsUnderline = True Then
            If .Font.Underline = wdUnderlineNone Then
                .Font.Underline = wdUnderlineSingle
            Else
                .Font.Underline = wdUnderlineNone
            End If
        End If
        If IsBold = True Then
            .Font.Bold = wdToggle
        End If
    End With
    iPrint = iPrint + 1
End Sub

'結束將字距縮小
Sub fn_PutEnd(oStrFileName As String)
If IsOpenWord = False Then
    fn_CreateWord True
End If
If Dir(App.path & "\" & oStrFileName) <> "" Then
    Kill App.path & "\" & oStrFileName
End If
With g_WordAp
    .ChangeFileOpenDirectory App.path & "\"
    .ActiveDocument.SaveAs FileName:=oStrFileName, FileFormat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
        True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
        False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
    .ActiveWindow.Close
    .ActivePrinter = oldtmpPrName 'Add By Sindy 2012/11/29 改回預設印表機
    IsOpenWord = False
End With
    g_WordAp.Visible = False
    g_WordAp.Quit
    Set g_WordAp = Nothing
End Sub

'將文字調成 word 格式
Function fn_StrToWordFmt(oStr As String, oCLen As Integer) As String
'中英混合時有問題
Dim stTmp As String, stTmpChr As String, strTail As String, iP As Integer, iSpace As Integer, iCnt As Integer
iCnt = 0
stTmp = ""
For iP = 1 To Len(oStr)
   stTmpChr = strConV(Mid(oStr, iP, 1), vbFromUnicode)
   If iCnt + LenB(stTmpChr) > 2 * oCLen Then
      Exit For
   Else
      iCnt = iCnt + LenB(stTmpChr)
   End If
   stTmp = stTmp & stTmpChr
Next
stTmp = strConV(MidB(stTmp, 1, iCnt), vbUnicode)
iSpace = 2 * oCLen - iCnt
If iSpace > 0 Then strTail = String(iSpace \ 2, "　") & String(iSpace Mod 2, " ")   '補全形空白,最後若有半字則補半型空白
fn_StrToWordFmt = stTmp & strTail
End Function

'放置相同符號到 word 內
Function fn_StrLineToWord(Optional oStr As String = "－", Optional IsA4 As Boolean = True) As String
If IsOpenWord = False Then
    fn_CreateWord True
End If

If IsA4 = True Then
    fn_StrLineToWord = String((160 / LenB(strConV(oStr, vbFromUnicode))), oStr)
Else
    fn_StrLineToWord = String((192 / LenB(strConV(oStr, vbFromUnicode))), oStr)
End If
End Function

'換頁
Function fn_DocNewPage()
If IsOpenWord = False Then
    fn_CreateWord True
End If
    fn_PutWords "", , , 10
End Function

Private Sub txtMail_GotFocus()
   TextInverse txtMail
End Sub

'Add by Amy 2022/05/02
'從cmdok_Click 搬過來修改
Private Function FormCheck() As Boolean
   FormCheck = False
   'Add by Amy 2022/04/29 輸出方式必擇一選
   If Check3(0).Value = 0 And Check3(1).Value = 0 Then
      s = MsgBox("輸出方式需擇一選擇!!", , "USER 輸入錯誤")
      Check3(1).SetFocus
      Exit Function
   End If
   '檢查客戶編號
   If Len(txt1(1)) = 0 Then
      s = MsgBox("客戶編號區間不可空白!!", , "USER 輸入錯誤")
      txt1(0).SetFocus
      txt1_GotFocus (0)
      Exit Function
   Else
      If Mid(Trim(txt1(0)), 1, 6) <> Mid(Trim(txt1(1)), 1, 6) Then
         s = MsgBox("客戶編號前六碼必須相同!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Function
      End If
    End If
    If Me.txt1(11).Text <> "" Then
        Me.lblSalesName.Caption = GetStaffName(Me.txt1(11).Text, True)
        If Me.lblSalesName.Caption = "" Then
            MsgBox "智權人員編號輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(11).SetFocus
            txt1_GotFocus 11
            Exit Function
        End If
    End If
    '檢查本所期限
    If Len(txt1(16)) = 0 Then
         s = MsgBox("本所期限區間不可空白!!", , "USER 輸入錯誤")
         txt1(16).SetFocus
         txt1_GotFocus (16)
         Exit Function
    ElseIf PUB_CheckKeyInDate(Me.txt1(16)) = -1 Then
          Me.txt1(16).SetFocus
          txt1_GotFocus 16
          Exit Function
    End If
    'Add by Amy 2022/05/02 未下本所期限止日會Error
    If Len(txt1(17)) = 0 Then
         s = MsgBox("本所期限區間不可空白!!", , "USER 輸入錯誤")
         txt1(17).SetFocus
         txt1_GotFocus (17)
         Exit Function
    ElseIf PUB_CheckKeyInDate(Me.txt1(17)) = -1 Then
          Me.txt1(17).SetFocus
          txt1_GotFocus 17
          Exit Function
    End If
    '檢查E-Mail
    If Len(Trim(txtMail)) = 0 Then
         s = MsgBox("E-Mail欄位不可空白!!", , "USER 輸入錯誤")
         txtMail.SetFocus
         Exit Function
    End If
    '系統類別
    If (Me.Check1(0).Value = vbUnchecked And Me.Check1(1).Value = vbUnchecked And Me.Check1(2).Value = vbUnchecked And _
         Me.Check1(3).Value = vbUnchecked And Me.Check1(4).Value = vbUnchecked And Me.Check1(5).Value = vbUnchecked And _
         Me.Check1(6).Value = vbUnchecked) Then
         MsgBox "請選擇系統類別!!!", vbExclamation + vbOKOnly
         Exit Function
    End If
    '輸出順序
    If Len(txt1(9)) = 0 Then
         s = MsgBox("輸出順序不可空白!!", , "USER 輸入錯誤")
         txt1(9).SetFocus
         Exit Function
    End If
    FormCheck = True
End Function

Sub fn_CreateExcel(ByRef intXlsSheet As Integer, ByRef strSystemKind As String, ByRef blnFirstPage As Boolean)
Dim ii As Integer, jj As Integer, kk As Integer
Dim strAllField As String, strAllWidth As String '欄位 第幾個字換行「-」分隔/欄寬
Dim strSetField_txt As String, strSetField As String, strTp3 As String, strTp4 As String
Dim strTp1() As String, strTp2() As String
Dim intReportN As Integer, bolOpen As Boolean '表名位置/已開Excel

On Error GoTo ErrHnd
    
intXlsRow = 1

If ReportName = "商標" Then
    bolOpen = IsOpenExcel(1)
ElseIf ReportName = "專利" Then
    bolOpen = IsOpenExcel(2)
Else
    bolOpen = IsOpenExcel(3)
End If
If bolOpen = False Or intXlsSheet > 1 Then
    If intXlsSheet = 1 Then
        xlsCustPoint.Workbooks.add
        xlsCustPoint.Visible = True
        xlsCustPoint.Application.WindowState = xlMinimized
    'Excel 2013 開啟只會有一個工作表,造成產生第二個工作表抓名稱會error
    ElseIf intXlsSheet > xlsCustPoint.Sheets.Count Then
        xlsCustPoint.Worksheets.add After:=wksrpt '插入sheet
    End If
    '工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(xlsCustPoint.Worksheets(1).Name, Len(xlsCustPoint.Worksheets(1).Name) - 1)
    Set wksrpt = xlsCustPoint.Worksheets(strWkName & intXlsSheet)
    wksrpt.Activate
    
    '報表抬頭格式設定
    wksrpt.Range(Chr(intField) & intXlsRow).Font.Size = 18
    wksrpt.Range(Chr(intField) & intXlsRow).Font.Bold = True
    wksrpt.Range(Chr(intField) & intXlsRow).Value = "客戶案件預算預估表(" & ReportName & ")"
    intReportN = intXlsRow
    
    intXlsRow = intXlsRow + 1
    wksrpt.Range(Chr(intField) & intXlsRow).Value = "收件人:" & StrTemp5(0)
    
    intXlsRow = intXlsRow + 1
    If blnFirstPage Then
        Erase setFieldNo_txt
        Erase setFieldNo
        Erase setFieldType
        
        jj = 0: kk = 0
        strAllField = "本所案號-2,客戶案件案號-3,案件名稱-2,申請日,申請國家-2,申請案號-2,審定號數-2,種類,專用期限(起)-4,專用期限(迄)-4"
        strAllWidth = "14.5,10.38,12.88,8.25,8.13,13,15,8.13,9.5,9.5"
        Select Case ReportName
            Case "專利"
            Case "商標"
                strAllField = strAllField & ",商品類別-2"
                strAllWidth = strAllWidth & ",9"
            Case "其他", "法務"
                strAllField = strAllField & ",商品類別-2"
                strAllWidth = strAllWidth & ",9"
        End Select
        strAllField = strAllField & ",本所期限-2,下一程序,預估費用"
        strAllWidth = strAllWidth & ",8.25,12.38,13"
        '勾「代表圖」
        If bolShowPic = True Then
            strAllField = "案件圖樣-2," & strAllField
            strAllWidth = "16.63," & strAllWidth
        End If
        '文字格式欄位
        strSetField_txt = "客戶案件案號,案件名稱,申請案號,審定號數"
        If ReportName <> "專利" Then
            strSetField_txt = strSetField_txt & ",商品類別"
        End If
        '其他格式設定
        strSetField = "申請日-e/mm/dd,專用期限(起)-yyyy/mm/dd,專用期限(迄)-yyyy/mm/dd"
        
        SetTitle = Split(strAllField, ",")
        setXlsWidth = Split(strAllWidth, ",")
        strTp1() = Split(strSetField_txt, ",")
        strTp2() = Split(strSetField, ",")
        
        ReDim setFieldNo_txt(UBound(strTp1)) '文字格式欄位
        ReDim setFieldNo(UBound(strTp2)) '其他格式欄位
        ReDim setFieldType(UBound(strTp2)) '其他格式欄位
        
        For ii = 0 To UBound(SetTitle)
            If jj <= UBound(strTp1) Then
                strTp3 = SetTitle(ii)
                If InStr(strTp3, "-") > 0 Then strTp3 = Mid(strTp3, 1, InStr(strTp3, "-") - 1)
                If strTp3 = strTp1(jj) Then
                    '設定文字欄位
                    setFieldNo_txt(jj) = Chr(ii + 65)
                    jj = jj + 1
                End If
            End If
            If kk <= UBound(strTp2) Then
                strTp3 = SetTitle(ii): strTp4 = strTp2(kk)
                If InStr(strTp3, "-") > 0 Then strTp3 = Mid(strTp3, 1, InStr(strTp3, "-") - 1)
                If InStr(strTp4, "-") > 0 Then strTp4 = Mid(strTp4, InStr(strTp4, "-") + 1)
                If strTp3 = Mid(strTp2(kk), 1, InStr(strTp2(kk), "-") - 1) Then
                    '設定需設格式的欄位
                    setFieldNo(kk) = Chr(ii + intField)
                    setFieldType(kk) = strTp4 '格式
                    kk = kk + 1
                End If
            End If
        Next ii
    End If
    '欄位
    For ii = LBound(SetTitle) To UBound(SetTitle)
        wksrpt.Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = setXlsWidth(ii)
        strTp3 = SetTitle(ii)
        If InStr(strTp3, "-") > 0 Then
            strTp4 = Mid(strTp3, InStr(SetTitle(ii), "-") + 1)
            strTp3 = Mid(strTp3, 1, InStr(SetTitle(ii), "-") - 1)
            strTp3 = Mid(strTp3, 1, Val(strTp4)) & Chr(10) & Mid(strTp3, Val(strTp4) + 1)
        End If
        wksrpt.Range(Chr(intField + ii) & intXlsRow).Value = strTp3
    Next ii
    intTitleR = intXlsRow
    
    '設定 表名 跨欄置中
    wksrpt.Range(Chr(intField) & intReportN & ":" & Chr(intField + UBound(SetTitle)) & intReportN).HorizontalAlignment = xlCenter
    wksrpt.Range(Chr(intField) & intReportN & ":" & Chr(intField + UBound(SetTitle)) & intReportN).MergeCells = True
    '設定 欄名 置中
    wksrpt.Range(Chr(intField) & intXlsRow & ":" & Chr(intField + UBound(SetTitle)) & intXlsRow).HorizontalAlignment = xlCenter
    
    intXlsRow = intXlsRow + 1
    
    '國內外分開印
    If txt1(8) = "Y" And strSystemKind <> MsgText(601) Then
        wksrpt.Name = strTemp(20) & "-" & strSystemKind
    Else
        wksrpt.Name = strTemp(20) & "-" & ReportName
    End If
    bolOpen = True
    If ReportName = "商標" Then
        IsOpenExcel(1) = bolOpen
    ElseIf ReportName = "專利" Then
        IsOpenExcel(2) = bolOpen
    Else
        IsOpenExcel(3) = bolOpen
    End If
End If
    
ErrHnd:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 91:
                xlsCustPoint.Workbooks.add
                Resume Next
            Case 462:
                Set xlsCustPoint = New Excel.Application
                xlsCustPoint.Workbooks.add
                Resume Next
            Case Else:
                MsgBox "錯誤 : " & Err.Description, vbCritical
        End Select
    End If
End Sub

Sub fn_PutExcel(strNo1 As String, strNo2 As String, strNo3 As String, strNo4 As String)
    Dim rsPic As New ADODB.Recordset
    Dim ii As Integer, intR As Integer
    Dim oShape, strData
    Dim stSQL As String, nowCaseNo As String, stFileName As String
    Dim oWidth As Single, oHeight As Single, wkWidth As Single

    '設定文字格式
    For ii = 0 To UBound(setFieldNo_txt)
        wksrpt.Range(setFieldNo_txt(ii) & intXlsRow).NumberFormatLocal = "@"
    Next ii
    
    If bolShowPic = True Then strXlsData = "$$" & strXlsData
    '資料寫入
    strData = Split(strXlsData, "$$")
    
    '修改顯示筆數(同案號算一筆,原一筆資料算一筆)
    nowCaseNo = strNo1 & "-" & strNo2 & strNo3 & strNo4
    For ii = 0 To UBound(strData)
        If bolShowPic = True And ii = GetValue("案件圖樣") Then
            If ReportName = "專利" Or ReportName = "商標" Then
                '同案號只印一個圖
                If oldCaseNo <> nowCaseNo Then
                    '讀取代表圖
                    'Modify by Amy 2023/07/26 改抓共用函數
'                    If GetImgByteFile_Case(strNo1, strNo2, strNo3, strNo4, stFileName, 0) = True Then
'                        wksrpt.Rows(intXlsRow).RowHeight = 110
'                        stFileName = Replace(stFileName, App.path & "\", "")
'                        Set oShape = wksrpt.Shapes.AddPicture(FileName:=App.path & "\" & stFileName, LinkToFile:=False, SaveWithDocument:=True, Left:=100, Top:=100, Width:=-1, Height:=-1) '.ConvertToShape '
'                        oShape.Select
'                        oWidth = wksrpt.Range(Chr(ii + 65) & intXlsRow).Width / oShape.Width
'                        oHeight = 110 / oShape.Height
'                        If oWidth > oHeight Then
'                            xlsCustPoint.Selection.ShapeRange.ScaleWidth Round(oHeight, 2) - 0.02, True, 0 '等比例縮放
'                        Else
'                            xlsCustPoint.Selection.ShapeRange.ScaleWidth Round(oWidth, 2) - 0.02, True, 0 '-0.02避免覆蓋格線上
'                        End If
'                        oShape.Left = wksrpt.Columns(Chr(ii + 65)).Left + 1
'                        oShape.Top = wksrpt.Range(Chr(ii + 65) & intXlsRow).Top + 1
'                    End If
                     Call PutXlsImg(Me.Name, wksrpt, Chr(ii + 65), intXlsRow, strNo1, strNo2, strNo3, strNo4)
                    'end 2023/07/26
                End If
            End If
        '其他非代表圖欄位
        'Modify by Amy 2022/07/19 未勾選代表圖,本所案號不會顯示,拿掉 If ii <> GetValue("案件圖樣")
        Else
            wksrpt.Range(Chr(ii + 65) & intXlsRow).Value = strData(ii)
        End If
    Next ii
    
    'Add by Amy 2022/07/19 每一系統類別的案件數-秀玲
    Select Case ReportName
        Case "專利"
            intCntP = intCntP + 1
        Case "商標"
            intCntT = intCntT + 1
        Case Else
            intCntO = intCntO + 1
    End Select
    'end 2022/07/19
    oldCaseNo = nowCaseNo
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(SetTitle)
       If UCase(SetTitle(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

'設定Excel
Sub fn_SetExcel()
    Dim ii As Integer
    
    '儲存格格式設定
    For ii = LBound(setFieldNo) To UBound(setFieldNo)
        wksrpt.Range(setFieldNo(ii) & intTitleR + 1 & ":" & setFieldNo(ii) & intXlsRow).NumberFormatLocal = setFieldType(ii)
    Next ii
    
    '欄位抬頭、內容格式設定
    With wksrpt.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(SetTitle)) & intXlsRow - 1)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Font.Name = "新細明體"
        .Font.Size = 11
    End With
    
    With wksrpt.PageSetup
        .PaperSize = 9       '設定紙張 A4
        .Orientation = xlLandscape                   '橫印
        .PrintTitleRows = "$1:$" & intTitleR       '表頭
        
        .LeftMargin = xlsCustPoint.InchesToPoints(0.5)
        .RightMargin = xlsCustPoint.InchesToPoints(0.5)
        .TopMargin = xlsCustPoint.InchesToPoints(0.3)
        .BottomMargin = xlsCustPoint.InchesToPoints(0.3)
        .HeaderMargin = xlsCustPoint.InchesToPoints(0.5)
        .FooterMargin = xlsCustPoint.InchesToPoints(0.5)
                    
        .Zoom = 100 '縮放比例
    End With
End Sub

'存檔
Sub fn_PutEndXls()
    Dim stFildName As String
    
    stFildName = "預算表" & ReportName & txt1(11).Text & "_" & txt1(0) & "_" & strSrvDate(1) & ".xls"
    If Dir(App.path & "\" & stFildName) <> "" Then
        Kill App.path & "\" & stFildName
    End If

        
    '判斷版本
    If Val(xlsCustPoint.Version) < 12 Then
        xlsCustPoint.Workbooks(1).SaveAs FileName:=App.path & "\" & stFildName, FileFormat:=-4143
    Else
        xlsCustPoint.Workbooks(1).SaveAs FileName:=App.path & "\" & stFildName, FileFormat:=56
    End If
    xlsCustPoint.Workbooks.Close
    xlsCustPoint.Quit
    Set wksrpt = Nothing
    Set xlsCustPoint = Nothing
    
    strFileN = strFileN & App.path & "\" & stFildName & "*"
    PUB_KillTempFile "$$*.*" '刪代表圖檔
End Sub

