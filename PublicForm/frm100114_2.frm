VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100114_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件查詢"
   ClientHeight    =   5730
   ClientLeft      =   1990
   ClientTop       =   1120
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   2520
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   15
      Width           =   720
   End
   Begin VB.CheckBox ChkPct 
      Caption         =   "Check1"
      Height          =   255
      Left            =   270
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件性質統計"
      Height          =   400
      Index           =   4
      Left            =   3255
      TabIndex        =   10
      Top             =   15
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8448
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   4776
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   15
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   6300
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7524
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   15
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   1590
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7267
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
      Height          =   1536
      Left            =   0
      TabIndex        =   2
      Top             =   4176
      Visible         =   0   'False
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   2716
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   3
      Left            =   960
      TabIndex        =   17
      Top             =   1290
      Width           =   8250
      VariousPropertyBits=   27
      Size            =   "14552;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Top             =   1005
      Width           =   8250
      VariousPropertyBits=   27
      Size            =   "14552;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   720
      Width           =   8250
      VariousPropertyBits=   27
      Size            =   "14552;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   1140
      TabIndex        =   14
      Top             =   480
      Width           =   3450
      VariousPropertyBits=   27
      Size            =   "6085;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：●代表銷卷＊代表閉卷"
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
      Left            =   6450
      TabIndex        =   11
      Top             =   480
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日文名稱："
      Height          =   180
      Index           =   3
      Left            =   30
      TabIndex        =   9
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文名稱："
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   8
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "中文名稱："
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "frm100114_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lbl1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim i As Integer, j As Integer, s As Integer, strSql As String, strTemp As Variant
Dim StrTest As String, intK As Integer, StrTest2 As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String, strSQL8 As String
Dim BolFrom100114 As Boolean
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'add by  nickc 2005/10/04  判斷是否法務
Public bolIsL As Boolean
'add by nickc 2005/10/04 若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
Private Const cntFaSql As String = " DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65))"

'add by nickc 2007/10/22 修正問題
Dim strSQL111 As String, StrSQL221 As String, StrSQL331 As String, StrSQL441 As String, strSQL551 As String, StrSQL661 As String, strSQL881 As String
Dim strSQL112 As String, StrSQL222 As String, StrSQL332 As String, StrSQL442 As String, strSQL552 As String, StrSQL662 As String, strSQL882 As String

'Add by Morgan 2008/11/20
'為了要能共用，前畫面條件改以參數方式傳遞
Public m_Sys As String '系統類別
Public m_Cty1 As String, m_Cty2 As String '申請國家
Public m_Pty1 As String, m_Pty2 As String '案件性質
Public m_Type As String '收發文別
Public m_Date1 As String, m_Date2 As String '日期
Public m_CKind As String '是否含C類來函 N:不含
'Added by Lydia 2019/11/01 利益衝突案件-管制
'Mark by Lydia 2019/12/26
'Dim m_CuFaArea As String '利益衝突案件：XY特殊權限管制-系統別
'Dim m_CuFaRight As String '利益衝突案件：XY特殊權限管制-可使用系統別
'Dim stCuFaSQL As String '利益衝突案件：查詢權限內的案件SQL
'Dim stConPA As String   '組合條件(Patent)
'Dim stConSP As String   '組合條件(ServicePractice)
'Dim m_adoRst As ADODB.Recordset
'Dim rsCnt As New ADODB.Recordset
'end 2019/12/26
'Memo by Lydia 2019/12/26 利益衝突案件：於後面增加欄位; 從外層SQL控制，改成逐案比對。
Dim intCufaCnt As Integer '限閱案件X件
Dim m_AllSys As String
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth, iDep As String
Dim iCol As Integer

iDep = PUB_GetST06(strUserNum)

If bolIsL = False Then
    'add by nickc 2007/03/23 加入 PCT 欄
    'edit by nickc 2007/12/21
    'If frm100114_1.ChkPct.Value = vbChecked Then
    'Modify By Sindy 2012/10/22 Grid 中 在申請人2 前加 申請人1
    If ChkPCT.Value = vbChecked Then
        'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
        arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                           , "申請案號", "PCT", "准駁" _
                           , "商品類別", "專用期間", "申請人1", "申請人2" _
                           , "申請人3", "申請人4", "申請人5", "FSORT" _
                           , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
         arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                           , 1200, 620, 450 _
                           , 800, 1800, 1100, 1100 _
                           , 1100, 1100, 1100, 0 _
                           , 0, 0, 0, 0, 0, 0)
    Else
        'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
        arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                           , "申請案號", "審定號專利號數", "准駁" _
                           , "商品類別", "專用期間", "申請人1", "申請人2" _
                           , "申請人3", "申請人4", "申請人5", "FSORT" _
                           , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
         arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                           , 1200, 1150, 450 _
                           , 800, 1800, 1100, 1100 _
                           , 1100, 1100, 1100, 0 _
                           , 0, 0, 0, 0, 0, 0)
    End If
Else
   'Modify By Sindy 2012/10/22 法務案在 收文日 前加 當事人1
   'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
   'arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                     , "申請日", "准駁", "當事人1", "收文日", "案件性質", "智權人員", "承辦人", _
                     "本所期限", "法定期限", "發文日" _
                     , "取消收文日", "代理人", "結果", "相關人", "進度備註" _
                     , "", "")
   '      arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                           , 850, 450, 850, 850, 800, 800, 800, _
                           850, 850, 850, _
                           850, 1200, 450, 1100, 1100, 0, 0)
   arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                     , "申請日", "准駁", "當事人1", "收文日", "案件性質", "智權人員", "承辦人", _
                     "本所期限", "法定期限", "發文日" _
                     , "取消收文日", "代理人", "結果", "相關人", "進度備註", "FSORT" _
                     , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
         arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                           , 850, 450, 850, 850, 800, 800, 800 _
                           , 850, 850, 850 _
                           , 850, 1200, 450, 1100, 1100, 0 _
                           , 0, 0, 0, 0, 0, 0)
End If

grdDataList1.Cols = UBound(arrGridHeadText) + 1
For iCol = 0 To grdDataList1.Cols - 1
   grdDataList1.row = 0
   grdDataList1.col = iCol
   grdDataList1.Text = arrGridHeadText(iCol)
   grdDataList1.ColWidth(iCol) = arrGridHeadWidth(iCol)
   grdDataList1.CellAlignment = flexAlignCenterCenter
Next iCol
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList1.ColWidth(2) = 0
Else
    grdDataList1.ColWidth(2) = 620
End If

End Sub

Private Sub cmdcp10_Click()
'92.04.16 nick 紀錄作用按鍵
cmdState = 4
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效

 Screen.MousePointer = vbHourglass
 frm100114_3.Show
 frm100114_3.StrMenu LBL1(0).Caption, BolFrom100114
 Screen.MousePointer = vbDefault
 Me.Hide
 Do
 DoEvents
 If bolToEndByNick = True Then Unload Me: Exit Sub
 Loop Until Not frm100114_3.Visible
 Unload frm100114_3
 Me.Show
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
'2cmd
Select Case cmdState
Case 0 '案件基本資料
      Me.Enabled = False
      For i = 1 To grdDataList1.Rows - 1
      grdDataList1.col = 0
      grdDataList1.row = i
      If Trim(grdDataList1.Text) = "V" Then
        grdDataList1.col = 0
        grdDataList1.Text = ""
        For j = 0 To grdDataList1.Cols - 1
           grdDataList1.col = j
           grdDataList1.CellBackColor = QBColor(15)
        Next j
        Dim Str01 As String
        grdDataList1.col = 1
        Str01 = SystemNumber(grdDataList1, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList1.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Select Case Pub_RplStr(Str01)
            Case "CFP", "FCP", "P"   '專利
                  Screen.MousePointer = vbHourglass
                  frm100101_3.Show
                  frm100101_3.Tag = Pub_RplStr(grdDataList1.Text)
                  frm100101_3.StrMenu
                  Screen.MousePointer = vbDefault
            Case "CFT", "FCT", "T", "TF"   '商標
                  Screen.MousePointer = vbHourglass
                  frm100101_4.Show
                  frm100101_4.Tag = Pub_RplStr(grdDataList1.Text)
                  frm100101_4.StrMenu
                  Screen.MousePointer = vbDefault
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"     '法務
                  Screen.MousePointer = vbHourglass
                  frm100101_5.Show
                  frm100101_5.Tag = Pub_RplStr(grdDataList1.Text)
                  frm100101_5.StrMenu
                  Screen.MousePointer = vbDefault
            Case "LA"            '顧問
                  Screen.MousePointer = vbHourglass
                  frm100101_6.Show
                  frm100101_6.Tag = Pub_RplStr(grdDataList1.Text)
                  frm100101_6.StrMenu
                  Screen.MousePointer = vbDefault
            Case Else                  '服務
                 Select Case Pub_RplStr(Str01)
                     Case "TB"    '條碼
                           Screen.MousePointer = vbHourglass
                         frm100101_7.Show
                         frm100101_7.Tag = Pub_RplStr(grdDataList1.Text)
                         frm100101_7.StrMenu
                         Screen.MousePointer = vbDefault
                     Case "TM"
                        Screen.MousePointer = vbHourglass
                         frm100101_8.Show
                         frm100101_8.Tag = Pub_RplStr(grdDataList1.Text)
                         frm100101_8.StrMenu
                         Screen.MousePointer = vbDefault
                     Case "TD"
                        Screen.MousePointer = vbHourglass
                         frm100101_9.Show
                         frm100101_9.Tag = Pub_RplStr(grdDataList1.Text)
                         frm100101_9.StrMenu
                         Screen.MousePointer = vbDefault
                     Case "TC", "CFC"
                         Screen.MousePointer = vbHourglass
                         frm100101_A.Show
                         frm100101_A.Tag = Pub_RplStr(grdDataList1.Text)
                         frm100101_A.StrMenu
                         Screen.MousePointer = vbDefault
                     Case Else
                         Screen.MousePointer = vbHourglass
                         frm100101_B.Show
                         frm100101_B.Tag = Pub_RplStr(grdDataList1.Text)
                         frm100101_B.StrMenu
                         Screen.MousePointer = vbDefault
                  End Select
            End Select
             Me.Enabled = True
             Exit Sub
        End If
     End If
     Next i
     Me.Enabled = True
Case 1 '案件進度
     Me.Enabled = False
     For i = 1 To grdDataList1.Rows - 1
     grdDataList1.col = 0
     grdDataList1.row = i
     If Trim(grdDataList1.Text) = "V" Then
        grdDataList1.col = 0
        grdDataList1.Text = ""
        For j = 0 To grdDataList1.Cols - 1
           grdDataList1.col = j
           grdDataList1.CellBackColor = QBColor(15)
        Next j
         grdDataList1.col = 1
         If Not IsNull(grdDataList1.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(grdDataList1.Text)
            'add by nickc 2005/10/06 加入分所號
            frm100101_2.Label15.Caption = grdDataList1.TextMatrix(i, 2)
            
            'Modify by Morgan 2008/11/21
            '申請人畫面的是否含來函條件改傳參數方式
            'If BolFrom100114 = False Then
            '    If Len(Trim(frm100102_1.Text8)) = 0 Then
            '        frm100101_2.StrMenu
            '    Else
            '        frm100101_2.StrMenu1
            '    End If
            'Else
            '    frm100101_2.StrMenu
            'End If
            'Modify By Sindy 2021/4/21 不要為了排除C類而寫2個函數 StrMenu,StrMenu1(Mark)
'            If Trim(Me.m_CKind) <> "" Then
'               frm100101_2.StrMenu1
'            Else
'               frm100101_2.StrMenu
'            End If
            frm100101_2.m_CKind = Me.m_CKind
            frm100101_2.StrMenu
            'end 2008/11/21
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   Me.Enabled = True
Case 2
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
   fnCloseAllFrm100
Case 4 '案件性質統計
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
       Me.Enabled = True
       Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100114_3.Show
   frm100114_3.StrMenu LBL1(0).Caption, BolFrom100114
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Exit Sub
'Add By Sindy 2019/1/15
Case 5 '卷宗區
   Me.Enabled = False
   For i = 1 To grdDataList1.Rows - 1
      grdDataList1.col = 0
      grdDataList1.row = i
      If Trim(grdDataList1.Text) = "V" Then
         grdDataList1.col = 0
         grdDataList1.Text = ""
         For j = 0 To grdDataList1.Cols - 1
            grdDataList1.col = j
            grdDataList1.CellBackColor = QBColor(15)
         Next j
         grdDataList1.col = 1
         If Not IsNull(grdDataList1.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_L.m_strKey = Pub_RplStr(grdDataList1.Text)
            'frm100101_L.Hide
            frm100101_L.SetParent Me
            If frm100101_L.QueryData = True Then
               frm100101_L.Show
               Me.Hide
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Activate()
If bolFNation = False Then
    s = MsgBox("國內人員不可查詢代理人案件", , "違規.....")
    Unload Me
    Exit Sub
End If
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth

'92.04.16 nick
cmdState = -1
End Sub

'Mark by Amy 2023/01/13 語法改共用函數
Sub StrMenu_Old()
'BolFrom100114 = True
'Me.Enabled = False
''顯示表單上頭資料
'lbl1(0).Caption = Me.Tag
'
''Add By Sindy 2011/01/03 檢查國內外權限
'If CheckSR12(Me.Tag) = False Then
'   Me.Enabled = True
'   Screen.MousePointer = vbDefault
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
'   Exit Sub
'End If
'
''edit by nickc 2005/12/06
''strSQL = "SELECT FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06 FROM FAGENT WHERE FA01='" & Left(GetNewFagent(Me.Tag), 8) & "' AND FA02='" & Right(GetNewFagent(Me.Tag), 1) & "' "
'strSql = "SELECT FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,fa77 FROM FAGENT WHERE FA01='" & Left(GetNewFagent(Me.Tag), 8) & "' AND FA02='" & Right(GetNewFagent(Me.Tag), 1) & "' "
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If Not IsNull(adoRecordset.Fields(0)) Then
'        lbl1(1).Caption = adoRecordset.Fields(0)
'    Else
'        lbl1(1).Caption = ""
'    End If
'    If Not IsNull(adoRecordset.Fields(1)) Then
'        lbl1(2).Caption = adoRecordset.Fields(1)
'    Else
'        lbl1(2) = ""
'    End If
'    If Not IsNull(adoRecordset.Fields(2)) Then
'        lbl1(3) = adoRecordset.Fields(2)
'    Else
'        lbl1(3) = ""
'    End If
'    If CheckStr(adoRecordset.Fields("fa77")) = "Y" Then
'        lbl1(0).ForeColor = &HFF&
'    Else
'        lbl1(0).ForeColor = &H80000012
'    End If
'Else
'    lbl1(1).Caption = ""
'    lbl1(2).Caption = ""
'    lbl1(3).Caption = ""
'End If
'CheckOC
''開始搜尋
'Dim strSQL11 As String
'Dim strSQL22 As String
'Dim strSQL33 As String
'Dim strSQL44 As String
'strSQL1 = ""
'strSQL2 = ""
'StrSQL3 = ""
'StrSQL4 = ""
'strSQL5 = ""
'StrSQL6 = ""
'
''系統類別
'If Len(Trim(m_Sys)) <> 0 Then
''edit by nickc 2007/10/22 修正
''   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
''   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
''   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
''   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
'   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
'   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
'   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
'   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
'   strSQL11 = strSQL11 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
'   strSQL22 = strSQL22 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
'   strSQL33 = strSQL33 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
'   strSQL44 = strSQL44 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
'End If
'
'm_AllSys = IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, "ALL")) 'Added by Lydia 2019/11/01
' 'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
' SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
' SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
' SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
' SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
' SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
' 'end 2019/12/26
'
''申請國家
'm_Cty1 = Trim(m_Cty1): m_Cty2 = Trim(m_Cty2)
'If m_Cty1 <> "" And m_Cty1 = m_Cty2 Then
'   strSQL1 = strSQL1 + " AND TM10='" & m_Cty1 & "' "
'   strSQL2 = strSQL2 + " AND PA09='" & m_Cty1 & "' "
'   StrSQL3 = StrSQL3 + " AND SP09='" & m_Cty1 & "' "
'   StrSQL4 = StrSQL4 + " AND LC15='" & m_Cty1 & "' "
'   strSQL11 = strSQL11 + " AND TM10='" & m_Cty1 & "' "
'   strSQL22 = strSQL22 + " AND PA09='" & m_Cty1 & "' "
'   strSQL33 = strSQL33 + " AND SP09='" & m_Cty1 & "' "
'   strSQL44 = strSQL44 + " AND LC15='" & m_Cty1 & "' "
'Else
'   If m_Cty1 <> "" Then
'      strSQL1 = strSQL1 + " AND TM10>='" & m_Cty1 & "' "
'      strSQL2 = strSQL2 + " AND PA09>='" & m_Cty1 & "' "
'      StrSQL3 = StrSQL3 + " AND SP09>='" & m_Cty1 & "' "
'      StrSQL4 = StrSQL4 + " AND LC15>='" & m_Cty1 & "' "
'      strSQL11 = strSQL11 + " AND TM10>='" & m_Cty1 & "' "
'      strSQL22 = strSQL22 + " AND PA09>='" & m_Cty1 & "' "
'      strSQL33 = strSQL33 + " AND SP09>='" & m_Cty1 & "' "
'      strSQL44 = strSQL44 + " AND LC15>='" & m_Cty1 & "' "
'   End If
'   If m_Cty2 <> "" Then
'      strSQL1 = strSQL1 + " AND TM10<='" & m_Cty2 & "' "
'      strSQL2 = strSQL2 + " AND PA09<='" & m_Cty2 & "' "
'      StrSQL3 = StrSQL3 + " AND SP09<='" & m_Cty2 & "' "
'      StrSQL4 = StrSQL4 + " AND LC15<='" & m_Cty2 & "' "
'      strSQL11 = strSQL11 + " AND TM10<='" & m_Cty2 & "' "
'      strSQL22 = strSQL22 + " AND PA09<='" & m_Cty2 & "' "
'      strSQL33 = strSQL33 + " AND SP09<='" & m_Cty2 & "' "
'      strSQL44 = strSQL44 + " AND LC15<='" & m_Cty2 & "' "
'   End If
'End If
'
''案件性質
'If Len(Trim(m_Pty1)) <> 0 Then
'    strSQL1 = strSQL1 + " AND CP10>='" & m_Pty1 & "' "
'    strSQL2 = strSQL2 + " AND CP10>='" & m_Pty1 & "' "
'    StrSQL3 = StrSQL3 + " AND CP10>='" & m_Pty1 & "' "
'    StrSQL4 = StrSQL4 + " AND CP10>='" & m_Pty1 & "' "
'    strSQL11 = strSQL11 + " AND CP10>='" & m_Pty1 & "' "
'    strSQL22 = strSQL22 + " AND CP10>='" & m_Pty1 & "' "
'    strSQL33 = strSQL33 + " AND CP10>='" & m_Pty1 & "' "
'    strSQL44 = strSQL44 + " AND CP10>='" & m_Pty1 & "' "
'End If
'If Len(Trim(m_Pty2)) <> 0 Then
'    strSQL1 = strSQL1 + " AND CP10<='" & m_Pty2 & "' "
'    strSQL2 = strSQL2 + " AND CP10<='" & m_Pty2 & "' "
'    StrSQL3 = StrSQL3 + " AND CP10<='" & m_Pty2 & "' "
'    StrSQL4 = StrSQL4 + " AND CP10<='" & m_Pty2 & "' "
'    strSQL11 = strSQL11 + " AND CP10<='" & m_Pty2 & "' "
'    strSQL22 = strSQL22 + " AND CP10<='" & m_Pty2 & "' "
'    strSQL33 = strSQL33 + " AND CP10<='" & m_Pty2 & "' "
'    strSQL44 = strSQL44 + " AND CP10<='" & m_Pty2 & "' "
'End If
'
''收文
'If m_Type = "1" Then
'   If Len(m_Date1) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(m_Date1)) & " "
'   End If
'   If Len(m_Date2) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(m_Date2)) & " "
'   'Add By Cheng 2002/03/18
'   Else
'      If Len(m_Date1) > 0 Then
'         strSQL5 = strSQL5 + " AND CP05<=" & strSrvDate(1) & " "
'      End If
'   End If
''發文
'Else
'   If Len(m_Date1) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(m_Date1)) & " "
'   End If
'   If Len(m_Date2) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(m_Date2)) & " "
'   'Add By Cheng 2002/03/18
'   Else
'      If Len(m_Date1) > 0 Then
'         strSQL5 = strSQL5 + " AND CP05<=" & strSrvDate(1) & " "
'      End If
'   End If
'End If
'
''Added by Lydia 2019/11/01 非法務案+屬於利益衝突案件之XY編號
''Mark by Lydia 2019/12/26
''stConPA = "": stConSP = ""
''If bolIsL = False And strSrvDate(1) >= XY特殊權限啟用日 And InStr(XY特殊權限範圍, Left(GetNewFagent(Me.Tag), 8)) > 0 Then
''    cnnConnection.Execute "delete from R100102_2 where R02201='" & strUserNum & "' and R02202='" & Me.Name & "' " '清空暫存檔
''    If PUB_ChkCuFa_Right(Me.Name, Me.Tag, m_AllSys, m_CuFaRight, m_CuFaArea) = True Then
''    End If
''    '有管制系統別=>組合SQL條件
''    If m_CuFaArea <> "" Then
''        stConPA = Pub_CufaConSQL(Me.Name, "PA", Me.Tag, m_CuFaRight, m_CuFaArea)
''        stConSP = Pub_CufaConSQL(Me.Name, "SP", Me.Tag, m_CuFaRight, m_CuFaArea)
''    End If
''End If
'''end 2019/11/01
''end 2019/12/26
'
''Modify By Sindy 2012/10/22 Grid 中 在申請人2 前加 申請人1
''                           法務案須抓申請人資料
''add by nickc 2005/10/05
'If bolIsL = False Then
''edit by nickc 2005/05/13
''    strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM27,'Y','＊','') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and tm44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
''    strSQL = strSQL + " union select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 其他申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 其他申請人5 FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and PA75='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5
'    'edit by nickc 2006/12/11
'    'strSQL = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and tm44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColTM
'    strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間," & _
'             "NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort" & SeColTM & _
'             " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and tm44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+)  and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'             " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColPA
'    strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort" & SeColPA & _
'             " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and PA75='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5
'
''edit by nickc 2005/05/13
''    strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
''             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
''             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (SP58='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & StrSQL5
''    strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
''             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
''             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (SP26='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5
''    strSQL = strSQL + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM LAWCASE,nation,CUSTOMER,caseprogress WHERE lc15=na01(+) and lc22='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  " & StrSQL4 & strSQL5
''    strSQL = strSQL & " union SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM27,'Y','＊','') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and cp44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1 & strSQL5
''    strSQL = strSQL + " union select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
''             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 其他申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 其他申請人5 FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and cp44='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
''             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
''             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  " & strSQL2 & strSQL5
''    strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
''             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
''             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (cp44='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  " & StrSQL3 & strSQL5
''    strSQL = strSQL + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM LAWCASE,nation,CUSTOMER,caseprogress WHERE lc15=na01(+) and cp44='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   " & StrSQL4 & strSQL5
''edit by nickc  2006/12/11
''    strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
'             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (SP26='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColSP
'    'Modify by Amy 2020/02/05 +SP73 商品類別
'    strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,NVL(SP73,'') AS商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
'             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort" & SeColSP & _
'             " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,customer c4,customer c5,caseprogress WHERE sp09=na01(+) and (SP26='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'             " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+)  and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'    strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
'             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort" & SeColLC & _
'             " FROM LAWCASE,nation,CUSTOMER C1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE lc15=na01(+) and lc22='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=C1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = C1.CU02(+)" & _
'             " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
'             " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  " & StrSQL4 & strSQL5
'
'    'edit by nickc 2006/12/11
'    'strSQL = strSQL & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and cp44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColTM
'    strSql = strSql & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3," & _
'                      "NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5 ,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort" & SeColTM & _
'                      " FROM TRADEMARK,nation,CUSTOMER c1,CUSTOMER c2,CUSTOMER c3,CUSTOMER c4,CUSTOMER c5,caseprogress WHERE tm10=na01(+) and cp44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColPA
'    strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort" & SeColPA & _
'             " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and cp44='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  " & strSQL2 & strSQL5
'    'edit by nickc 2006/12/11
'    'strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
'             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (cp44='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  " & StrSQL3 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColSP
'    'Modify by Amy 2020/02/05 +SP73 商品類別
'    strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,NVL(SP73,'') AS商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
'             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort" & SeColSP & _
'             " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER c4,CUSTOMER c5,caseprogress WHERE sp09=na01(+) and (cp44='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'             " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  " & StrSQL3 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'    strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
'             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort" & SeColLC & _
'             " FROM LAWCASE,nation,CUSTOMER C1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE lc15=na01(+) and cp44='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=C1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = C1.CU02(+)" & _
'             " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
'             " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   " & StrSQL4 & strSQL5
'Else
'    'Modify By Sindy 2012/10/22 法務案在 收文日 前加 當事人1
'    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'    strSql = "select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁, NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 當事人1, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,C0.CU01||C0.CU02,C0.CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & SeColLC & _
'                 " FROM LAWCASE,nation,CUSTOMER C0,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind,CUSTOMER C1 " & _
'                 " WHERE lc15=na01(+) and LC22='" & Me.Tag & "'  AND LC04='00' AND SUBSTR(cp56,1,8)=C0.CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = C0.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & _
'                 " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
'                   StrSQL4 & strSQL5
'    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'    strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁, NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 當事人1, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,C0.CU01||C0.CU02,C0.CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & SeColLC & _
'                " FROM LAWCASE,nation,CUSTOMER C0,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind,CUSTOMER C1 " & _
'                " WHERE lc15=na01(+) and cp44='" & Me.Tag & "'  AND LC04='00' AND SUBSTR(cp56,1,8)=C0.CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = C0.CU02(+) and Cp01=lC01(+) AND Cp02=lC02(+) AND Cp03=lC03(+) AND Cp04=lC04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & _
'                " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
'                StrSQL4 & strSQL5
'End If
'    'edit by nickc 2005/05/13
'    'strSQL = strSQL + " ORDER BY 本所案號"
'    strSql = strSql + " ORDER BY FSort,本所案號"
'
''Added by Lydia 2019/11/01 利益衝突案件：處理替換字串
''Mark by Lydia 2019/12/26
''If m_CuFaArea <> "" And stConPA & stConSP <> "" Then
''    stCuFaSQL = strSql
''    stCuFaSQL = Replace(stCuFaSQL, "CUFA_PA", stConPA)
''    stCuFaSQL = Replace(stCuFaSQL, "CUFA_SP", stConSP)
''    intI = 1
''    Set rsCnt = Nothing
''    Set rsCnt = ClsLawReadRstMsg(intI, stCuFaSQL)
''End If
''strSql = Replace(strSql, "CUFA_PA", "")
''strSql = Replace(strSql, "CUFA_SP", "")
'''end 2019/11/01
''end 2019/12/26
'
'CheckOC
's = 0
'StrTest2 = ""
''add by nickc 2007/03/23 更換 PCT 欄
''edit by nickc 2007/12/21
''If frm100114_1.ChkPct.Value = vbChecked Then
'If ChkPct.Value = vbChecked Then
'    strSql = Replace(Replace(Replace(Replace(UCase(strSql), "TM15 AS 審定號專利號數", "'' as PCT"), "SP14 AS 審定號專利號數", "'' as PCT"), "' ' AS 審定號專利號數", "'' as PCT"), "PA22 AS 審定號專利號數", "pa46 as PCT")
'End If
'adoRecordset.CursorLocation = adUseClient
''Modified by Lydia 2019/12/26 改變型態
''adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
'
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    'Added by Lydia 2019/12/26 利益衝突案件：逐案號判斷
'     If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
'        intCufaCnt = 0
'        adoRecordset.MoveFirst
'        Do While adoRecordset.EOF = False
'            '利益衝突案件：逐案號判斷
'            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
'                intCufaCnt = intCufaCnt + 1
'                adoRecordset.Delete
'            End If
'            adoRecordset.MoveNext
'        Loop
'        '利益衝突案件：限閱案件
'        If intCufaCnt > 0 Then
'            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
'        End If
'        If adoRecordset.RecordCount = 0 Then
'              GoTo JumpToNoData
'        End If
'     End If
'    'end 2019/12/26
'    cmdOK(0).Enabled = True
'    cmdOK(1).Enabled = True
'Else
'JumpToNoData: 'Added by Lydia 2019/11/01
'    cmdOK(0).Enabled = False
'    cmdOK(1).Enabled = False
'    Me.Enabled = True
'    ShowNoData
'    Screen.MousePointer = vbDefault
'    '92.04.18 nick
'    'Me.Hide
'    tmpBol = fnCancelNowFormAndShowParentForm(Me)
'    Exit Sub
'End If
'Set grdDataList1.Recordset = adoRecordset
'SetDataListWidth
'CheckOC
'Me.Enabled = True
End Sub 'End StrMenu_Old 語法改共用函數

'由申請人畫面來   910801
'Memo by Lydia 2019/11/01 (2008/11/21)已改用StrMenu
Sub StrMenu2()
BolFrom100114 = False
Me.Enabled = False
'顯示表單上頭資料
LBL1(0).Caption = Me.Tag

'Add By Sindy 2011/01/03 檢查國內外權限
If CheckSR12(Me.Tag) = False Then
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

strSql = "SELECT FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06 FROM FAGENT WHERE FA01='" & Left(GetNewFagent(Me.Tag), 8) & "' AND FA02='" & Right(GetNewFagent(Me.Tag), 1) & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If Not IsNull(adoRecordset.Fields(0)) Then
        LBL1(1).Caption = adoRecordset.Fields(0)
    Else
        LBL1(1).Caption = ""
    End If
    If Not IsNull(adoRecordset.Fields(1)) Then
        LBL1(2).Caption = adoRecordset.Fields(1)
    Else
        LBL1(2) = ""
    End If
    If Not IsNull(adoRecordset.Fields(2)) Then
        LBL1(3) = adoRecordset.Fields(2)
    Else
        LBL1(3) = ""
    End If
Else
    LBL1(1).Caption = ""
    LBL1(2).Caption = ""
    LBL1(3).Caption = ""
End If
CheckOC
'開始搜尋
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQL6 = ""
strSQL8 = ""
'add by nickc 2007/10/22 修正問題
strSQL111 = ""
strSQL112 = ""
StrSQL221 = ""
StrSQL222 = ""
StrSQL331 = ""
StrSQL332 = ""
StrSQL441 = ""
StrSQL442 = ""
strSQL881 = ""
strSQL882 = ""
'系統類別
If Len(Trim(m_Sys)) <> 0 Then
   'Modify  By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(m_Sys, 2) & ") and cp01 in (" & SQLGrpStr(m_Sys, 2) & ") "
'   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(m_Sys, 1) & ") and cp01 in (" & SQLGrpStr(m_Sys, 1) & ") "
'   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(m_Sys, 5) & ") and cp01 in (" & SQLGrpStr(m_Sys, 5) & ") "
'   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(m_Sys, 3) & ") and cp01 in (" & SQLGrpStr(m_Sys, 3) & ") "
'edit by nickc 2007/10/22 修正問題
'   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 2) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 2) & ") "
'   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 1) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 1) & ") "
'   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 5) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 5) & ") "
'   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 3) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 3) & ") "
'   strSQL8 = strSQL8 & " and hc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 4) & ") and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(,m_Sys)), 4) & ") "
   strSQL111 = strSQL111 & " and tm01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
   StrSQL221 = StrSQL221 & " and pa01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
   StrSQL331 = StrSQL331 & " and sp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
   StrSQL441 = StrSQL441 & " and lc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
   strSQL881 = strSQL881 & " and hc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 4) & ") "
   strSQL112 = strSQL112 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
   StrSQL222 = StrSQL222 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
   StrSQL332 = StrSQL332 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
   StrSQL442 = StrSQL442 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
   strSQL882 = strSQL882 & " and cp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 4) & ") "
End If

m_AllSys = IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, "ALL")) 'Added by Lydia 2019/11/01
 'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
 SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
 SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
 SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
 SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
 SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
 'end 2019/12/26
 
If Len(Trim(m_Pty1)) <> 0 Then            '檢查案件性質
    strSQL1 = strSQL1 + " AND CP10>='" & m_Pty1 & "' "
    strSQL2 = strSQL2 + " AND CP10>='" & m_Pty1 & "' "
    StrSQL3 = StrSQL3 + " AND CP10>='" & m_Pty1 & "' "
    StrSQL4 = StrSQL4 + " AND CP10>='" & m_Pty1 & "' "
    strSQL8 = strSQL8 + " AND CP10>='" & m_Pty1 & "' "
End If
If Len(Trim(m_Pty2)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & m_Pty2 & "' "
    strSQL2 = strSQL2 + " AND CP10<='" & m_Pty2 & "' "
    StrSQL3 = StrSQL3 + " AND CP10<='" & m_Pty2 & "' "
    StrSQL4 = StrSQL4 + " AND CP10<='" & m_Pty2 & "' "
    strSQL8 = strSQL8 + " AND CP10<='" & m_Pty2 & "' "
End If

   If Len(m_Date1) <> 0 Then
      strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(m_Date1)) & " "
   End If
   If Len(m_Date2) <> 0 Then
      strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(m_Date2)) & " "
   'Add By Cheng 2002/03/18
   Else
      If Len(m_Date1) > 0 Then
         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
      End If
   End If
'Added by Lydia 2019/11/01 非法務案+屬於利益衝突案件之XY編號
'Mark by Lydia 2019/12/26
'stConPA = "": stConSP = ""
'If bolIsL = False And strSrvDate(1) >= XY特殊權限啟用日 And InStr(XY特殊權限範圍, Left(GetNewFagent(Me.Tag), 8)) > 0 Then
'    cnnConnection.Execute "delete from R100102_2 where R02201='" & strUserNum & "' and R02202='" & Me.Name & "' " '清空暫存檔
'    If PUB_ChkCuFa_Right(Me.Name, Me.Tag, m_AllSys, m_CuFaRight, m_CuFaArea) = True Then
'    End If
'    '有管制系統別=>組合SQL條件
'    If m_CuFaArea <> "" Then
'        stConPA = Pub_CufaConSQL(Me.Name, "PA", Me.Tag, m_CuFaRight, m_CuFaArea)
'        stConSP = Pub_CufaConSQL(Me.Name, "SP", Me.Tag, m_CuFaRight, m_CuFaArea)
'    End If
'End If
''end 2019/11/01
'end 2019/12/26

'edit by  nickc 2005/05/13
'    strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM27,'Y','＊','') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and tm44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 其他申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 其他申請人5 FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and PA75='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
'             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (SP58='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM LAWCASE,nation,CUSTOMER,caseprogress WHERE lc15=na01(+) and lc22='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  " & StrSQL4 & strSQL5
'    strSQL = strSQL & " union SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM27,'Y','＊','') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and cp44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1 & strSQL5
'    strSQL = strSQL + " union select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 其他申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 其他申請人5 FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and cp44='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  " & strSQL2 & strSQL5
'    strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
'             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (cp44='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  " & StrSQL3 & strSQL5
'    strSQL = strSQL + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5 FROM LAWCASE,nation,CUSTOMER,caseprogress WHERE lc15=na01(+) and cp44='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   " & StrSQL4 & strSQL5
'
'    strSQL = strSQL + " ORDER BY 本所案號"
'add by nickc 2005/10/05
'Modify By Sindy 2012/10/22 Grid 中 在申請人2 前加 申請人1
'                           法務案須抓申請人資料
If bolIsL = False Then
    'edit by nickc 2006/12/11
    'strSQL = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and tm44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
    'Modified by Lydia 2019/12/26 +增加欄位SeColTM
    strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4," & _
             "NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort" & SeColTM & _
             " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and tm44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+)  and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
             " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5 & strSQL111
    
    'Modified by Lydia 2019/12/26 +增加欄位SeColPA
    strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort" & SeColPA & _
             " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and PA75='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5 & StrSQL221
    'edit by nickc 2006/12/11
    'strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (SP58='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5
    'Modified by Lydia 2019/12/26 +增加欄位SeColSP
    'Modify by Amy 2020/02/05 +SP73 商品類別
    strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,NVL(SP73,'') AS商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort" & SeColSP & _
             " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (SP58='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
             " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+)  and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5 & StrSQL331
    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
    strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort" & SeColLC & _
             " FROM LAWCASE,nation,CUSTOMER C1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE lc15=na01(+) and lc22='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=C1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = C1.CU02(+)" & _
             " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
             " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  " & StrSQL4 & strSQL5 & StrSQL441
    'edit by nickc 2006/12/11
    'strSQL = strSQL & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,' ' AS 其他申請人2,' ' AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort FROM TRADEMARK,nation,CUSTOMER,caseprogress WHERE tm10=na01(+) and cp44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = CU02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1 & strSQL5
    'Modified by Lydia 2019/12/26 +增加欄位SeColTM
    strSql = strSql & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號,tm15 AS 審定號專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 目前准駁,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3," & _
             "NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4," & _
             "NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5," & _
             "TM01||'-'||TM02||'-'||TM03||'-'||TM04 as FSort" & SeColTM & _
             " FROM TRADEMARK,nation,CUSTOMER c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and cp44='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
             " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)  " & strSQL1 & strSQL5 & strSQL112
    
    'Modified by Lydia 2019/12/26 +增加欄位SeColPA
    strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號,PA22 AS 審定號專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 目前准駁,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
             "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1,NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort" & SeColPA & _
             " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and cp44='" & Me.Tag & "'  and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
             " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
             " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  " & strSQL2 & strSQL5 & StrSQL222
    'edit by nickc 2006/12/11
    'strSQL = strSQL + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間" & _
             ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 其他申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 其他申請人3,' ' AS 其他申請人4,' ' AS 其他申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,caseprogress WHERE sp09=na01(+) and (cp44='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  " & StrSQL3 & strSQL5
    'Modified by Lydia 2019/12/26 +增加欄位SeColSP
    'Modify by Amy 2020/02/05 +SP73 商品類別
    strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號,SP14 AS 審定號專利號數,' ' AS 目前准駁,NVL(SP73,'') AS商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
             "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as FSort" & SeColSP & _
             " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (cp44='" & Me.Tag & "')  AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
             " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+)  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)  " & StrSQL3 & strSQL5 & StrSQL332
    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
    strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號,' ' AS 審定號專利號數,' ' AS 目前准駁,' ' AS 商品類別,'-' AS 專用期間,NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 申請人1," & _
             "NVL(C2.CU04,DECODE(C2.cu05,null,C2.CU06,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90)) AS 申請人2,NVL(C3.CU04,DECODE(C3.cu05,null,C3.CU06,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90)) AS 申請人3,NVL(C4.CU04,DECODE(C4.cu05,null,C4.CU06,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90)) AS 申請人4,NVL(C5.CU04,DECODE(C5.cu05,null,C5.CU06,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as FSort" & SeColLC & _
             " FROM LAWCASE,nation,CUSTOMER C1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE lc15=na01(+) and cp44='" & Me.Tag & "'  AND SUBSTR(LC11,1,8)=C1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = C1.CU02(+)" & _
             " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
             " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)   " & StrSQL4 & strSQL5 & StrSQL442
Else
    'Modify By Sindy 2012/10/22 法務案在 收文日 前加 當事人1
    'Modified by Lydia 2019/12/26 +增加欄位SeColLC
    strSql = "select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁, NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 當事人1, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,C0.CU01||C0.CU02,C0.CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & SeColLC & _
                " FROM LAWCASE,nation,CUSTOMER C0,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind,CUSTOMER C1 " & _
                   " WHERE lc15=na01(+) and LC22='" & Me.Tag & "'  AND LC04='00' AND SUBSTR(cp56,1,8)=C0.CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = C0.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & _
                   " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                   StrSQL4 & strSQL5 & StrSQL441
    strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁, NVL(C1.CU04,DECODE(C1.cu05,null,C1.CU06,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90)) AS 當事人1, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,C0.CU01||C0.CU02,C0.CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & SeColLC & _
                " FROM LAWCASE,nation,CUSTOMER C0,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind,CUSTOMER C1 " & _
                   " WHERE lc15=na01(+) and cp44='" & Me.Tag & "'  AND LC04='00' AND SUBSTR(cp56,1,8)=C0.CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = C0.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & _
                   " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                   StrSQL4 & strSQL5 & StrSQL441
End If
    strSql = strSql + " ORDER BY FSort,本所案號"

'Added by Lydia 2019/11/01 利益衝突案件：處理替換字串
'Mark by Lydia 2019/12/26
'If m_CuFaArea <> "" And stConPA & stConSP <> "" Then
'    stCuFaSQL = strSql
'    stCuFaSQL = Replace(stCuFaSQL, "CUFA_PA", stConPA)
'    stCuFaSQL = Replace(stCuFaSQL, "CUFA_SP", stConSP)
'    intI = 1
'    Set rsCnt = Nothing
'    Set rsCnt = ClsLawReadRstMsg(intI, stCuFaSQL)
'End If
'strSql = Replace(strSql, "CUFA_PA", "")
'strSql = Replace(strSql, "CUFA_SP", "")
''end 2019/11/01
'end 2019/12/26

CheckOC
s = 0
StrTest2 = ""
'add by nickc 2007/03/23 更換 PCT 欄
'edit by nickc 2007/12/21
'If frm100114_1.ChkPct.Value = vbChecked Then
If ChkPCT.Value = vbChecked Then
    strSql = Replace(Replace(Replace(Replace(UCase(strSql), "TM15 AS 審定號專利號數", "'' as PCT"), "SP14 AS 審定號專利號數", "'' as PCT"), "' ' AS 審定號專利號數", "'' as PCT"), "PA22 AS 審定號專利號數", "pa46 as PCT")
End If
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/12/26 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Added by Lydia 2019/12/26 利益衝突案件：逐案號判斷
     If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        intCufaCnt = 0
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
     End If
    'end 2019/12/26
    cmdOK(0).Enabled = True
    cmdOK(1).Enabled = True
Else
JumpToNoData: 'Added by Lydia 2019/11/01
    cmdOK(0).Enabled = False
    cmdOK(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList1.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100114_2 = Nothing
End Sub

Private Sub grdDataList1_Click()
grdDataList1.Visible = False
grdDataList1.row = grdDataList1.MouseRow
grdDataList1.col = 0
If grdDataList1.row <> 0 Then
If grdDataList1.Text = "V" Then
     grdDataList1.Text = ""
     For i = 0 To grdDataList1.Cols - 1
          grdDataList1.col = i
          grdDataList1.CellBackColor = QBColor(15)
    Next i
Else
     grdDataList1.Text = "V"
     For i = 0 To grdDataList1.Cols - 1
         grdDataList1.col = i
         grdDataList1.CellBackColor = &HFFC0C0
     Next i

End If
End If
grdDataList1.Visible = True
End Sub

'Add by Amy 2023/01/13 改共用函數,並整理程式
Sub StrMenu()
Dim dblRow As Double 'Add By Sindy 2025/9/3

    BolFrom100114 = True
    Me.Enabled = False
    '顯示表單中代理人編號
    LBL1(0).Caption = Me.Tag
    'Add by Amy 2023/10/13 目前代理人聯絡人編號未抓案件資料,故Show 代理人編號即可-秀玲
    If InStr(Me.Tag, "-") > 0 Then
        LBL1(0).Caption = Mid(Me.Tag, 1, Val(InStr(Me.Tag, "-")) - 1)
    End If
    
    '檢查國內外權限
    If CheckSR12(Me.Tag) = False Then
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
    pub_QL05 = pub_QL05 & ";代理人編號：" & Me.Tag & "(案件)" 'Add By Sindy 2025/8/13
    
    'Modify by Amy 2023/10/13 目前代理人聯絡人編號未抓案件資料,故Show 代理人編號,原Me.Tag
    strSql = "SELECT FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,fa77 FROM FAGENT WHERE FA01='" & Left(GetNewFagent(LBL1(0).Caption), 8) & "' AND FA02='" & Right(GetNewFagent(LBL1(0).Caption), 1) & "' "
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        If Not IsNull(adoRecordset.Fields(0)) Then
            LBL1(1).Caption = adoRecordset.Fields(0)
        Else
            LBL1(1).Caption = ""
        End If
        If Not IsNull(adoRecordset.Fields(1)) Then
            LBL1(2).Caption = adoRecordset.Fields(1)
        Else
            LBL1(2) = ""
        End If
        If Not IsNull(adoRecordset.Fields(2)) Then
            LBL1(3) = adoRecordset.Fields(2)
        Else
            LBL1(3) = ""
        End If
        If CheckStr(adoRecordset.Fields("fa77")) = "Y" Then
            LBL1(0).ForeColor = &HFF&
        Else
            LBL1(0).ForeColor = &H80000012
        End If
    Else
        LBL1(1).Caption = ""
        LBL1(2).Caption = ""
        LBL1(3).Caption = ""
    End If
    CheckOC
        
    strSQL1 = ""
    '案件性質
    If Len(Trim(m_Pty1)) <> 0 Then
        strSQL1 = strSQL1 + " AND CP10>='" & m_Pty1 & "' "
    End If
    If Len(Trim(m_Pty2)) <> 0 Then
        strSQL1 = strSQL1 + " AND CP10<='" & m_Pty2 & "' "
    End If
    'Add By Sindy 2025/8/13
    If Len(m_Pty1) <> 0 Or Len(m_Pty2) <> 0 Then
        pub_QL05 = pub_QL05 & ";案件性質：" & m_Pty1 & "-" & m_Pty2
    End If
    '2025/8/13 END

    '收文
    If m_Type = "1" Then
        If Len(m_Date1) <> 0 Then
            strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(m_Date1)) & " "
        End If
        If Len(m_Date2) <> 0 Then
            strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(m_Date2)) & " "
        Else
            If Len(m_Date1) > 0 Then
                strSQL1 = strSQL1 + " AND CP05<=" & strSrvDate(1) & " "
            End If
        End If
        'Add By Sindy 2025/8/13
        If Len(m_Date1) <> 0 Or Len(m_Date2) <> 0 Then
            pub_QL05 = pub_QL05 & ";收文日期：" & m_Date1 & "-" & m_Date2
        End If
        '2025/8/13 END
    '發文
    Else
        If Len(m_Date1) <> 0 Then
            strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(m_Date1)) & " "
        End If
        If Len(m_Date2) <> 0 Then
            strSQL1 = strSQL1 + " AND CP27<=" & Val(ChangeTStringToWString(m_Date2)) & " "
        Else
            If Len(m_Date1) > 0 Then
                strSQL1 = strSQL1 + " AND CP05<=" & strSrvDate(1) & " "
            End If
        End If
        'Add By Sindy 2025/8/13
        If Len(m_Date1) <> 0 Or Len(m_Date2) <> 0 Then
            pub_QL05 = pub_QL05 & ";發文日期：" & m_Date1 & "-" & m_Date2
        End If
        '2025/8/13 END
    End If
    'Modify by Amy 2023/10/13 於共用函數判斷取聯絡人編號,原:Pub_GetCusCaseSql(Me.Name, Me.Tag, Me.Tag, …)
    'Modify by Amy 2023/01/19 +if
    '不是 法務專用
    If bolIsL = False Then
        strSql = Pub_GetCusCaseSql(Me.Name, Me.Tag, m_Sys, bolIsL, ChkPCT.Value, , strSQL1, m_Cty1, m_Cty2)
    '法務專用 (前畫面按「法務進度」鈕)
    Else
        strSql = Pub_GetCusCaseSql(Me.Name, Me.Tag, m_Sys, bolIsL, ChkPCT.Value, cntFaSql, strSQL1, m_Cty1, m_Cty2)
    End If
    'end 2023/10/13
    strSql = strSql + " ORDER BY FSort,本所案號"
    CheckOC
    s = 0
    StrTest2 = ""
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

        '利益衝突案件：逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            intCufaCnt = 0
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/13
        End If
        cmdOK(0).Enabled = True
        cmdOK(1).Enabled = True
    Else
        If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
JumpToNoData:
        cmdOK(0).Enabled = False
        cmdOK(1).Enabled = False
        Me.Enabled = True
        ShowNoData
        Screen.MousePointer = vbDefault
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
    Set grdDataList1.Recordset = adoRecordset
    SetDataListWidth
    CheckOC
    Me.Enabled = True
End Sub

