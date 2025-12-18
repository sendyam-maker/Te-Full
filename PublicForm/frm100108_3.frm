VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100108_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "關聯案件資料及正聯商標查詢"
   ClientHeight    =   5710
   ClientLeft      =   120
   ClientTop       =   4920
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5710
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   1488
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   3012
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   4
      Left            =   7584
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "本案案件進度"
      Height          =   400
      Index           =   3
      Left            =   6060
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "本案案件基本資料"
      Height          =   400
      Index           =   2
      Left            =   4236
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   8508
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3660
      Left            =   45
      TabIndex        =   6
      Top             =   2040
      Width           =   9210
      _ExtentX        =   16228
      _ExtentY        =   6456
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
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
      Height          =   255
      Index           =   10
      Left            =   7170
      TabIndex        =   21
      Top             =   720
      Width           =   2025
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱（日）："
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   1506
      Width           =   1440
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1575
      TabIndex        =   19
      Top             =   1506
      Width           =   7590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13388;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱（英）："
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   1242
      Width           =   1440
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1575
      TabIndex        =   17
      Top             =   1242
      Width           =   7590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13388;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1575
      TabIndex        =   16
      Top             =   978
      Width           =   7590
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13388;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   15
      Top             =   714
      Width           =   2775
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4895;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   6030
      TabIndex        =   14
      Top             =   450
      Width           =   2775
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4895;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1110
      TabIndex        =   13
      Top             =   450
      Width           =   2775
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4895;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1005
      TabIndex        =   12
      Top             =   1770
      Width           =   8205
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "14473;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱（中）："
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   978
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   255
      Left            =   30
      TabIndex        =   10
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/證書號數："
      Height          =   255
      Left            =   4230
      TabIndex        =   9
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   714
      Width           =   900
   End
End
Attribute VB_Name = "frm100108_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lbl1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String, StrSQL7 As String, strSQL8 As String, strSQL22 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer, StrTempSystemKind As String
Dim strTemp As String, intK As Integer
Dim strNumber As String, strTemp3 As Variant, StrTest4 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
'利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String


Private Sub SetDataListWidth()
'Added by Lydia 2019/11/01
Dim intField As Integer
intField = 15
grdDataList.Cols = intField
'end 2019/11/01

   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 1: grdDataList.Text = "本所案號"
   grdDataList.ColWidth(1) = 1550
   grdDataList.CellAlignment = flexAlignCenterCenter
   Dim iDep As String
   iDep = PUB_GetST06(strUserNum)
   grdDataList.col = 2: grdDataList.Text = "分所號"
   '電腦中心，跟分所才秀
   If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
       grdDataList.ColWidth(2) = 0
   Else
       grdDataList.ColWidth(2) = 620
   End If
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "案件名稱"
   grdDataList.ColWidth(3) = 2000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "申請人"
   grdDataList.ColWidth(4) = 2000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(5) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "專利商標種類"
   grdDataList.ColWidth(6) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 7: grdDataList.Text = "目前准駁"
   grdDataList.ColWidth(7) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 8: grdDataList.Text = "閉卷日期"
   grdDataList.ColWidth(8) = 810
   grdDataList.CellAlignment = flexAlignCenterCenter
   'Add By Cheng 2002/02/25
   grdDataList.col = 9: grdDataList.Text = "審定號/專利號數"
   grdDataList.ColWidth(9) = 1400
   grdDataList.CellAlignment = flexAlignLeftCenter
   
'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 10 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
   
   Select Case cmdState
   Case 0
        Me.Enabled = False
        For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           For j = 0 To grdDataList.Cols - 1
                 grdDataList.col = j
                 grdDataList.CellBackColor = QBColor(15)
           Next j
           Dim Str01 As String
           grdDataList.col = 1
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
                     frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_3.StrMenu
                     Screen.MousePointer = vbDefault
               Case "CFT", "FCT", "T", "TF"   '商標
                     Screen.MousePointer = vbHourglass
                     frm100101_4.Show
                     frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_4.StrMenu
                     Screen.MousePointer = vbDefault
               'Modify By Sindy 2009/07/24 增加LIN系統類別
               'modify by sonia 2019/7/29 +ACS系統類別
               Case "CFL", "FCL", "L", "LIN", "ACS"  '法務
                     Screen.MousePointer = vbHourglass
                     frm100101_5.Show
                     frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_5.StrMenu
                     Screen.MousePointer = vbDefault
               Case "LA"            '顧問
                     Screen.MousePointer = vbHourglass
                     frm100101_6.Show
                     frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_6.StrMenu
                     Screen.MousePointer = vbDefault
               Case Else                  '服務
                    Select Case Pub_RplStr(Str01)
                        Case "TB"    '條碼
                            Screen.MousePointer = vbHourglass
                           frm100101_7.Show
                            frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_7.StrMenu
                            Screen.MousePointer = vbDefault
                        Case "TM"
                            Screen.MousePointer = vbHourglass
                              frm100101_8.Show
                            frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_8.StrMenu
                            Screen.MousePointer = vbDefault
                        Case "TD"
                            Screen.MousePointer = vbHourglass
                           frm100101_9.Show
                            frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_9.StrMenu
                            Screen.MousePointer = vbDefault
                        Case "TC", "CFC"
                              Screen.MousePointer = vbHourglass
                              frm100101_A.Show
                            frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_A.StrMenu
                            Screen.MousePointer = vbDefault
                        Case Else
                            Screen.MousePointer = vbHourglass
                           frm100101_B.Show
                            frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
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
   Case 1
        Me.Enabled = False
        
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                 grdDataList.col = j
                 grdDataList.CellBackColor = QBColor(15)
           Next j
            grdDataList.col = 1
            If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100101_2.Show
               frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
               frm100101_2.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        End If
        Next i
        Me.Enabled = True
   Case 2
        cmdState = -1
        Me.Enabled = False
        Str01 = SystemNumber(lbl1(0), 1)
        If Not IsNull(lbl1(0)) Then
           If fnSaveParentForm(Me) = False Then
               Me.Enabled = True
               Exit Sub
           End If
           Select Case Pub_RplStr(Str01)
           Case "CFP", "FCP", "P"   '專利
               Screen.MousePointer = vbHourglass
               frm100101_3.Show
               frm100101_3.Tag = lbl1(0)
               frm100101_3.StrMenu
               Screen.MousePointer = vbDefault
           Case "CFT", "FCT", "T", "TF"   '商標
               Screen.MousePointer = vbHourglass
               frm100101_4.Show
               frm100101_4.Tag = lbl1(0)
               frm100101_4.StrMenu
               Screen.MousePointer = vbDefault
           'Modify By Sindy 2009/07/24 增加LIN系統類別
           'modify by sonia 2019/7/29 +ACS系統類別
           Case "CFL", "FCL", "L", "LIN", "ACS"  '法務
               Screen.MousePointer = vbHourglass
               frm100101_5.Show
               frm100101_5.Tag = lbl1(0)
               frm100101_5.StrMenu
               Screen.MousePointer = vbDefault
           Case "LA"            '顧問
               Screen.MousePointer = vbHourglass
               frm100101_6.Show
               frm100101_6.Tag = lbl1(0)
               frm100101_6.StrMenu
               Screen.MousePointer = vbDefault
           Case Else                  '服務
               Select Case Pub_RplStr(Str01)
                   Case "TB"    '條碼
                       Screen.MousePointer = vbHourglass
                       frm100101_7.Show
                       frm100101_7.Tag = lbl1(0)
                       frm100101_7.StrMenu
                       Screen.MousePointer = vbDefault
                   Case "TM"
                       Screen.MousePointer = vbHourglass
                       frm100101_8.Show
                       frm100101_8.Tag = lbl1(0)
                       frm100101_8.StrMenu
                       Screen.MousePointer = vbDefault
                   Case "TD"
                       Screen.MousePointer = vbHourglass
                       frm100101_9.Show
                       frm100101_9.Tag = lbl1(0)
                       frm100101_9.StrMenu
                       Screen.MousePointer = vbDefault
                   Case "TC", "CFC"
                       Screen.MousePointer = vbHourglass
                       frm100101_A.Show
                       frm100101_A.Tag = lbl1(0)
                       frm100101_A.StrMenu
                       Screen.MousePointer = vbDefault
                   Case Else
                       Screen.MousePointer = vbHourglass
                       frm100101_B.Show
                       frm100101_B.Tag = lbl1(0)
                       frm100101_B.StrMenu
                       Screen.MousePointer = vbDefault
                   End Select
               End Select
               Me.Enabled = True
               Exit Sub
        End If
        Me.Enabled = True
   Case 3
        cmdState = -1
        Me.Enabled = False
        If Not IsNull(lbl1(0)) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100101_2.Show
               frm100101_2.Tag = lbl1(0)
               frm100101_2.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
        End If
        Me.Enabled = True
   Case 4
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Case 5
        fnCloseAllFrm100
   Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub

End Sub

Private Sub Form_Load()
'Added by Lydia 2021/12/21
Dim Lbl As Object

For Each Lbl In Me.lbl1
    Lbl.BackColor = &H8000000F
    Lbl.Caption = ""
Next
'end 2021/12/21

   bolToEndByNick = False
      MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()                 '相關卷號與正聯商標 'Memo by Lydia 2020/11/10 從frm100108_1來,所以抓表單的欄位
Dim strTM27 As String '正商標號數
'Add By Cheng 2002/07/02
Dim strTM12 As String '申請案號
Dim dblRow As Double 'Add By Sindy 2025/9/3

   Me.Enabled = False
   
   lbl1(0).Caption = Me.Tag
    'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
    SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
    SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
    SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
    SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
    SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
    m_AllSys = SystemNumber(Me.Tag, 1)
    intCufaCnt = 0
    'end 2019/11/01
   
   'Modify By Cheng 2002/04/25
   '多顯示正商標號數
   'strSQL = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),tm23) FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
   'strSQL = strSQL + " union all select PA11,PA05,PA06,PA07,PA22,nvl(nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),pa26) FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
   'strSQL = strSQL + " union all select SP11,SP05,SP06,SP07,SP14,nvl(nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),sp08) FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
   'strSQL = strSQL + " union all select '' ,LC05,LC06,LC07,'',nvl(nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),lc11) FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
   'strSQL = strSQL + " union all select '',HC06,'','','',nvl(nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)),hc05) FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
'   strSql = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23),TM27 FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
'   strSql = strSql + " union all select PA11,PA05,PA06,PA07,PA22,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26),'' FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
'   strSql = strSql + " union all select SP11,SP05,SP06,SP07,SP14,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08),'' FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
'   strSql = strSql + " union all select '' ,LC05,LC06,LC07,'',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11),'' FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
'   strSql = strSql + " union all select '',HC06,'','','',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05),'' FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   strSql = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23),TM27 as 正商標號數" & SeColTM & _
                                " FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
   strSql = strSql + " union all select PA11,PA05,PA06,PA07,PA22,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26),'' as 正商標號數" & SeColPA & _
                                " FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
   strSql = strSql + " union all select SP11,SP05,SP06,SP07,SP14,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08),''  as 正商標號數" & SeColSP & _
                                " FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
   strSql = strSql + " union all select '' ,LC05,LC06,LC07,'',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11),''  as 正商標號數" & SeColLC & _
                                " FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
   strSql = strSql + " union all select '',HC06,'','','',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05),''  as 正商標號數" & SeColHC & _
                                " FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   'end 2019/11/01
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   'Modified by Lydia 2019/11/01 改變型態
   'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

   '若有資料
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
         dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, lbl1(0).Caption, "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109), vbInformation, MsgText(1110)
            End If
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/9/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpNoDataMain
            End If
            adoRecordset.MoveFirst
         Else
            If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/9/3
         End If
        'end 2019/11/01
        lbl1(1).Caption = CheckStr(adoRecordset.Fields(0))
        lbl1(2).Caption = CheckStr(adoRecordset.Fields(1))
        lbl1(3).Caption = CheckStr(adoRecordset.Fields(2))
        lbl1(4).Caption = CheckStr(adoRecordset.Fields(3))
        lbl1(6).Caption = CheckStr(adoRecordset.Fields(4))
        lbl1(5).Caption = CheckStr(adoRecordset.Fields(5))
        'Add By Cheng 2002/04/25
        strTM27 = "" & adoRecordset.Fields(6).Value
        'Add By Cheng 2002/07/02
        If IsNull(adoRecordset.Fields(4).Value) Then
           strTM12 = "" & adoRecordset.Fields(0).Value
        Else
           strTM12 = ""
        End If
   '若無資料
   Else
       If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/3
JumpNoDataMain: 'Added by Lydia 2019/11/01
       lbl1(1).Caption = ""
       lbl1(2).Caption = ""
       lbl1(3).Caption = ""
       lbl1(4).Caption = ""
       lbl1(6).Caption = ""
       Me.Enabled = True
       Screen.MousePointer = vbDefault
   End If
   
   CheckOC
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   strSQL22 = ""
   'If Len(Trim(frm100108_1.txt1(6))) <> 0 Then
      'Modify By Cheng 2002/03/14
   '   strSQL1 = strSQL1 & " and PA01 in (" & SQLGrpStr(frm100108_1.txt1(6), 1) & ") "
   '   strSQL2 = strSQL2 & " and TM01 in (" & SQLGrpStr(frm100108_1.txt1(6), 2) & ") "
   '   StrSQL3 = StrSQL3 & " and LC01 in (" & SQLGrpStr(frm100108_1.txt1(6), 3) & ") "
   '   StrSQL4 = StrSQL4 & " and HC01 in (" & SQLGrpStr(frm100108_1.txt1(6), 4) & ") "
   '   StrSQL5 = StrSQL5 & " and SP01 in (" & SQLGrpStr(frm100108_1.txt1(6), 5) & ") "
   '   StrSQL22 = StrSQL22 & " and TM01 in (" & SQLGrpStr(frm100108_1.txt1(6), 2) & ") "
      strSQL1 = strSQL1 & " and PA01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
      strSQL2 = strSQL2 & " and TM01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
      StrSQL3 = StrSQL3 & " and LC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 3) & ") "
      StrSQL4 = StrSQL4 & " and HC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 4) & ") "
      strSQL5 = strSQL5 & " and SP01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
      strSQL22 = strSQL22 & " and TM01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
   'End If
   
    'Added by Lydia 2019/11/01 利益衝突案件
    m_AllSys = IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(, frm100108_1.txt1(6).Text))
    intCufaCnt = 0
    'end 2019/11/01
        
   If frm100108_1.txt1(7) = "1" Then     '查詢相關卷號
   'Modify By Cheng 2002/02/25
   '                    strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,DeCODE(pa09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,''   AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03           AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,''   AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cR07=hc03(+) and cR08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & StrSQL5
   'Modify By Cheng 2002/04/25
   '若已閉卷, 則在本所案號後加"＊"號
   '                        strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,DeCODE(pa09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,''   AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03           AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,''   AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cR07=hc03(+) and cR08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & StrSQL5
   'edit by nickc 2006/06/22 改 CASERELATION1
   '                        strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,''   AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03           AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,''   AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cR07=hc03(+) and cR08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   '    strSQL = strSQL & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
       'Modified by Lydia 2019/11/01 +增加欄位SeColTM, SeColPA, SeColSP, SeColLC, SeColHC
                           strSql = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 " & SeColPA & _
                                                " FROM CASERELATION1,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
       strSql = strSql & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                                " FROM CASERELATION1,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
       strSql = strSql & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,nvl(lc05,nvl(lc06,lc07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,''   AS 審定專利號數 " & SeColLC & _
                                                " FROM CASERELATION1,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
       strSql = strSql & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc06                     AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) AS 申請人,na03           AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,''   AS 審定專利號數 " & SeColHC & _
                                                " FROM CASERELATION1,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cR07=hc03(+) and cR08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
       strSql = strSql & " union all select '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 " & SeColSP & _
                                                " FROM CASERELATION1,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
       'end 2019/11/01
       
       'Modified by Lydia 2019/11/01 改成模組ProcDataByCase
'       CheckOC
'       adoRecordset.CursorLocation = adUseClient
'       adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'       If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'           cmdOK(0).Enabled = True
'           cmdOK(1).Enabled = True
'           cmdOK(2).Enabled = True
'           cmdOK(3).Enabled = True
'       Else
'           cmdOK(0).Enabled = False
'           cmdOK(1).Enabled = False
'           cmdOK(2).Enabled = False
'           cmdOK(3).Enabled = False
'           Me.Enabled = True
'           ShowNoData
'           Screen.MousePointer = vbDefault
'           '92.04.18 nick
'           'Me.Hide
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'           Exit Sub
'       End If
'       Set grdDataList.Recordset = adoRecordset
'       CheckOC
       Call ProcDataByCase("1", strSql)
       'end 2019/11/01
       
   Else     '查詢正聯商標
       strTemp = SystemNumber(Me.Tag, 1)
       'Modify By Cheng 2002/02/25
       '以正商標查聯合商標
       If frm100108_1.txt1(7).Text = 2 Then
          If strTemp = "CFT" Or strTemp = "FCT" Or strTemp = "T" Or strTemp = "TF" Then
               'Modify By Cheng 2002/02/25
      '        strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期 FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP WHERE TM27='" & lbl1(6) & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & StrSQL22
               'Modify By Cheng 2002/04/25
   '            strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP WHERE TM27='" & lbl1(6) & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & StrSQL22
               'Modify By Cheng 2003/06/03
   '            strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP WHERE TM27='" & lbl1(6) & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & StrSQL22
               'Modified by Lydia 2019/11/01 +增加欄位SeColTM
               strSql = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                " FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP, (Select TM08 As T08, TM15 As T15 From Trademark Where " & ChgTradeMark(Replace(Me.Tag, "-", "")) & " ) T1 WHERE TM27=T1.T15 And (Decode(T1.T08,'1','2','')=TM08 Or Decode(T1.T08,'1','3','')=TM08 Or Decode(T1.T08,'4','5','')=TM08 Or Decode(T1.T08,'4','6','')=TM08) And TM27='" & lbl1(6) & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL22
               'Add By Cheng 2002/07/02
               '抓同日申請的資料(當同日申請時, 正聯商標皆無審定號數, 且聯合商標的正商標號數為正商標的申請號數)
               If strTM12 <> "" And (frm100108_1.Option1(0).Value Or frm100108_1.Option1(1).Value) Then
                   'Modify By Cheng 2003/06/03
   '               strSQL = strSQL + " Union SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP WHERE TM27='" & strTM12 & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & StrSQL22
                  'Modified by Lydia 2019/11/01 +增加欄位SeColTM
                  strSql = strSql + " Union SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                            " FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP, (Select TM08 As T08, TM12 As T12 From Trademark Where " & ChgTradeMark(Replace(Me.Tag, "-", "")) & " ) T1 WHERE TM27=T1.T12 And (Decode(T1.T08,'1','2','')=TM08 Or Decode(T1.T08,'1','3','')=TM08 Or Decode(T1.T08,'4','5','')=TM08 Or Decode(T1.T08,'4','6','')=TM08) And TM27='" & strTM12 & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL22
               End If
              
              'Modified by Lydia 2019/11/01 改成模組ProcDataByCase
'              CheckOC
'              If Len(Trim(frm100108_1.txt1(6))) <> 0 Then
'                  'Modify By Cheng 2002/03/14
'   '               strTemp3 = Split(frm100108_1.txt1(6), ",")
'                  strTemp3 = Split(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), ",")
'              End If
'              adoRecordset.CursorLocation = adUseClient
'              adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'              If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'                  cmdOK(0).Enabled = True
'                  cmdOK(1).Enabled = True
'                  cmdOK(2).Enabled = True
'                  cmdOK(3).Enabled = True
'              Else
'                  cmdOK(0).Enabled = False
'                  cmdOK(1).Enabled = False
'                  cmdOK(2).Enabled = False
'                  cmdOK(3).Enabled = False
'                  Me.Enabled = True
'                  ShowNoData
'                  Screen.MousePointer = vbDefault
'                  '92.04.18 nick
'                  'Me.Hide
'                 tmpBol = fnCancelNowFormAndShowParentForm(Me)
'                   Exit Sub
'              End If
'              Set grdDataList.Recordset = adoRecordset
'              CheckOC
              Call ProcDataByCase("2", strSql)
              'end 2019/11/01
          Else
              s = MsgBox("此本所案號沒有正聯商標, 無法查詢!!" & Me.Tag & "  ", , "錯誤")
              cmdOK(0).Enabled = False
              cmdOK(1).Enabled = False
              cmdOK(2).Enabled = False
              cmdOK(3).Enabled = False
              Screen.MousePointer = vbDefault
              '92.04.18 nick
              'Me.Hide
             tmpBol = fnCancelNowFormAndShowParentForm(Me)
          End If
      
      '以聯合商標查正商標
      ElseIf frm100108_1.txt1(7).Text = "3" Then
          If strTemp = "CFT" Or strTemp = "FCT" Or strTemp = "T" Or strTemp = "TF" Then
               'Modify By Cheng 2002/02/25
               strSQL1 = ""
               strSQL2 = ""
               strSQL5 = ""
               StrSQL6 = ""
               StrSQL7 = ""
               strSQL8 = ""
               If Len(Trim(frm100108_1.txt1(6))) <> 0 Then
                  'Modify By Cheng 2002/03/14
   '               strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 1) & ") "
   '               strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 2) & ") "
   '               StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 5) & ") "
   '               StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 1) & ") "
   '               StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 2) & ") "
   '               StrSQL8 = StrSQL8 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 5) & ") "
                  strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
                  strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
                  StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
                  StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
                  StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
                  strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
               End If
      
      '        strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期 FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP WHERE TM27='" & lbl1(6) & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & StrSQL22
   '           strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP WHERE TM15='" & lbl1(6) & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & StrSQL22
              
               'Modify By Cheng 2002/04/23
   '                                strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & frm100108_1.txt1(5).Text & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
   '            strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & frm100108_1.txt1(5).Text & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
   '            strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & frm100108_1.txt1(5).Text & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL5
   '
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & frm100108_1.txt1(5).Text & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & frm100108_1.txt1(5).Text & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & frm100108_1.txt1(5).Text & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL8
                               
               'Modify By Cheng 2002/04/25
   '                            strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Me.lbl1(6).Caption & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
   '            strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & Me.lbl1(6).Caption & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
   '            strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & Me.lbl1(6).Caption & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL5
   '
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Me.lbl1(6).Caption & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Me.lbl1(6).Caption & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Me.lbl1(6).Caption & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL8
                                   'Modify By Cheng 2003/06/03
   '                                strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & strTM27 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
               'Modified by Lydia 2019/11/01 +增加欄位SeColTM, SeColPA, SeColSP
                                   strSql = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                                    " FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP, (Select TM08 As T08, TM27 As T27 From Trademark Where " & ChgTradeMark(Replace(Me.Tag, "-", "")) & " ) T1 WHERE TM15=T1.T27 And (Decode(T1.T08,'2','1','')=TM08 Or Decode(T1.T08,'3','1','')=TM08 Or Decode(T1.T08,'5','4','')=TM08 Or Decode(T1.T08,'6','4','')=TM08) And TM15='" & strTM27 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
               strSql = strSql + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數 " & SeColPA & _
                                                    " FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & strTM27 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
               strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數 " & SeColSP & _
                                                    " FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & strTM27 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
               'end 2019/11/01
               
               'Add By Cheng 2002/07/02
               '抓同日申請的資料
               If frm100108_1.Option1(0).Value Or frm100108_1.Option1(1).Value Then
                   'Modify By Cheng 2003/06/03
   '               strSQL = strSQL + " union all select ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM12='" & strTM27 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2 & " AND TM15 IS NULL "
                  'Modified by Lydia 2019/11/01 +增加欄位SeColTM, SeColPA, SeColSP
                  strSql = strSql + " union all select ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數" & SeColTM & _
                                          " FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP, (Select TM08 As T08, TM27 As T27 From Trademark Where " & ChgTradeMark(Replace(Me.Tag, "-", "")) & " ) T1 WHERE TM12=T1.T27 And (Decode(T1.T08,'2','1','')=TM08 Or Decode(T1.T08,'3','1','')=TM08 Or Decode(T1.T08,'5','4','')=TM08 Or Decode(T1.T08,'6','4','')=TM08) And  TM12='" & strTM27 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2 & " AND TM15 IS NULL "
                  strSql = strSql + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數" & SeColPA & _
                                          " FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA11='" & strTM27 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1 & " AND PA22 IS NULL "
                  strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數" & SeColSP & _
                                          " FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP11='" & strTM27 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5 & " AND SP14 IS NULL "
                  'end 2019/11/01
               End If
               'Modify By Cheng 2003/06/03
   '            strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & strTM27 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL7
               'Modified by Lydia 2019/11/01 +增加欄位SeColTM, SeColPA, SeColSP
               strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("TM30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                 " FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP, (Select TM08 As T08, TM27 As T27 From Trademark Where " & ChgTradeMark(Replace(Me.Tag, "-", "")) & " ) T1 WHERE CP36=T1.T27 And (Decode(T1.T08,'2','1','')=TM08 Or Decode(T1.T08,'3','1','')=TM08 Or Decode(T1.T08,'5','4','')=TM08 Or Decode(T1.T08,'6','4','')=TM08) And CP36='" & strTM27 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
               strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("PA58") & " AS 閉卷日期,PA22 AS 審定專利號數 " & SeColPA & _
                                " FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & strTM27 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
               strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,nvl(NA03,sp09) AS 申請國家,''                             AS 專利商標種類,''                                AS 目前准駁," & SQLDate("SP16") & " AS 閉卷日期,SP14 AS 審定專利號數 " & SeColSP & _
                               " FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & strTM27 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
              
              'Modified by Lydia 2019/11/01 改成模組ProcDataByCase
'              CheckOC
'              If Len(Trim(frm100108_1.txt1(6))) <> 0 Then
'                  'Modify By Cheng 2002/03/14
'   '               strTemp3 = Split(frm100108_1.txt1(6), ",")
'                  strTemp3 = Split(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), ",")
'              End If
'              adoRecordset.CursorLocation = adUseClient
'              adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'              If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'                  cmdOK(0).Enabled = True
'                  cmdOK(1).Enabled = True
'                  cmdOK(2).Enabled = True
'                  cmdOK(3).Enabled = True
'              Else
'                  cmdOK(0).Enabled = False
'                  cmdOK(1).Enabled = False
'                  cmdOK(2).Enabled = False
'                  cmdOK(3).Enabled = False
'                  Me.Enabled = True
'                  ShowNoData
'                  Screen.MousePointer = vbDefault
'                  '92.04.18 nick
'                  'Me.Hide
'                 tmpBol = fnCancelNowFormAndShowParentForm(Me)
'                   Exit Sub
'              End If
'              Set grdDataList.Recordset = adoRecordset
'              CheckOC
              Call ProcDataByCase("3", strSql)
              'end 2019/11/01
          Else
              s = MsgBox("此本所案號沒有正聯商標, 無法查詢!!" & Me.Tag & "  ", , "錯誤")
              cmdOK(0).Enabled = False
              cmdOK(1).Enabled = False
              cmdOK(2).Enabled = False
              cmdOK(3).Enabled = False
              Screen.MousePointer = vbDefault
              '92.04.18 nick
              'Me.Hide
             tmpBol = fnCancelNowFormAndShowParentForm(Me)
          End If
      End If
   End If
   'Me.Enabled = True 'Remove by Lydia 2019/11/01
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100108_3 = Nothing
End Sub

Private Sub grdDataList_SelChange()
   
   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
   If grdDataList.Text = "V" Then
        grdDataList.Text = ""
        For i = 0 To grdDataList.Cols - 1
             grdDataList.col = i
             grdDataList.CellBackColor = QBColor(15)
       Next i
   Else
        grdDataList.Text = "V"
        For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellBackColor = &HFFC0C0
        Next i
   End If
   End If
   grdDataList.Visible = True

End Sub

Sub StrMenu1()               '相關卷號 frm100110_1 用
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
   Me.Enabled = False
   
   lbl1(0).Caption = Me.Tag
    'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
    SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
    SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
    SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
    SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
    SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
    m_AllSys = SystemNumber(Me.Tag, 1)
    intCufaCnt = 0
    'end 2019/11/01
    
   'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
'   strSql = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23) FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
'   strSql = strSql + " union all select PA11,PA05,PA06,PA07,PA22,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
'   strSql = strSql + " union all select SP11,SP05,SP06,SP07,SP14,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
'   strSql = strSql + " union all select '' ,LC05,LC06,LC07,'',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
'   strSql = strSql + " union all select '',HC06,'','','',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   strSql = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23) as 申請人1 " & SeColTM & _
                                " FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
   strSql = strSql + " union all select PA11,PA05,PA06,PA07,PA22,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) as 申請人1 " & SeColPA & _
                                " FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
   strSql = strSql + " union all select SP11,SP05,SP06,SP07,SP14,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) as 申請人1 " & SeColSP & _
                                " FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
   strSql = strSql + " union all select '' ,LC05,LC06,LC07,'',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) as 申請人1 " & SeColLC & _
                                " FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
   strSql = strSql + " union all select '',HC06,'','','',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) as 申請人1 " & SeColHC & _
                                " FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   'end 2019/11/01
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   'Modified by Lydia 2019/11/01 改變型態
   'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
         dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, lbl1(0).Caption, "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109), vbInformation, MsgText(1110)
            End If
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/9/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpNoDataMain
            End If
            adoRecordset.MoveFirst
         Else
            If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/9/3
         End If
        'end 2019/11/01
      lbl1(1).Caption = CheckStr(adoRecordset.Fields(0))
      lbl1(2).Caption = CheckStr(adoRecordset.Fields(1))
      lbl1(3).Caption = CheckStr(adoRecordset.Fields(2))
      lbl1(4).Caption = CheckStr(adoRecordset.Fields(3))
      lbl1(6).Caption = CheckStr(adoRecordset.Fields(4))
      lbl1(5).Caption = CheckStr(adoRecordset.Fields(5))
   Else
       If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/3
JumpNoDataMain: 'Added by Lydia 2019/11/01
       lbl1(1).Caption = ""
       lbl1(2).Caption = ""
       lbl1(3).Caption = ""
       lbl1(4).Caption = ""
       lbl1(6).Caption = ""
       Me.Enabled = True
   
   End If
   
   CheckOC
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   'Remove by Lydia 2019/11/01 不使用
   'If Len(Trim(frm100110_1.txt1(6))) <> 0 Then
   '   strSQL1 = strSQL1 & " and cr01 in (" & SQLGrpStr(frm100110_1.txt1(6), 1) & ") "
   '   strSQL2 = strSQL2 & " and cr01 in (" & SQLGrpStr(frm100110_1.txt1(6), 2) & ") "
   '   StrSQL3 = StrSQL3 & " and cr01 in (" & SQLGrpStr(frm100110_1.txt1(6), 3) & ") "
   '   StrSQL4 = StrSQL4 & " and cr01 in (" & SQLGrpStr(frm100110_1.txt1(6), 4) & ") "
   '  strSQL5 = strSQL5 & " and cr01 in (" & SQLGrpStr(frm100110_1.txt1(6), 5) & ") "
   'End If
   'end 2019/11/01
   
   'add by nick 2004/11/12
   'modify by sonia 2015/3/19 不在此管制可看之系統類別, 否則FCP串P(FMP),外專程序會看不到 FCP-051717串P-111066
   'strSQL1 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 1) & ") "
   'strSQL2 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 2) & ") "
   'StrSQL3 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 3) & ") "
   'StrSQL4 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 4) & ") "
   'strSQL5 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 5) & ") "
   strSQL1 = "AND PA01 IS NOT NULL "
   strSQL2 = "AND TM01 IS NOT NULL "
   StrSQL3 = "AND LC01 IS NOT NULL "
   StrSQL4 = "AND HC01 IS NOT NULL "
   strSQL5 = "AND SP01 IS NOT NULL "
   'end 2015/3/19
   
   'Added by Lydia 2019/11/01 利益衝突案件
   m_AllSys = GetAllSysKind(, "ALL")
   intCufaCnt = 0
   'end 2019/11/01
   
   'Modify By Cheng 2002/02/25
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,DeCODE(pa09,'000',PTM03,PTM04) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and hc07=hc03(+) and hc08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & StrSQL5
   
   'Modify By Cheng 2002/04/25
   '若已閉卷, 則在本所案號後加"＊"號
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,DeCODE(pa09,'000',PTM03,PTM04) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and hc07=hc03(+) and hc08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & StrSQL5
   'edit by nick 2004/11/12 沒連結到
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and hc07=hc03(+) and hc08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   'edit by nickc 2006/06/22  改 CASERELATION1
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cr07=hc03(+) and cr08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   'Modified by Lydia 2019/11/01 +增加欄位SeColTM, SeColPA, SeColSP, SeColLC, SeColHC
   strSql = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 " & SeColPA & _
                                " FROM CASERELATION1,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                " FROM CASERELATION1,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,nvl(lc05,nvl(lc06,lc07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 " & SeColLC & _
                                " FROM CASERELATION1,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc06 AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 " & SeColHC & _
                                " FROM CASERELATION1,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cr07=hc03(+) and cr08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 " & SeColSP & _
                                " FROM CASERELATION1,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   'end 2019/11/01
   
   'Modified by Lydia 2019/11/01 改成模組ProcDataByCase
'   CheckOC
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'       cmdOK(0).Enabled = True
'       cmdOK(1).Enabled = True
'       cmdOK(2).Enabled = True
'       cmdOK(3).Enabled = True
'   Else
'       cmdOK(0).Enabled = False
'       cmdOK(1).Enabled = False
'       cmdOK(2).Enabled = False
'       cmdOK(3).Enabled = False
'       ShowNoData
'       Me.Enabled = True
'       '92.04.18 nick
'       'Me.Hide
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'       Exit Sub
'   End If
'   Set grdDataList.Recordset = adoRecordset
'   CheckOC
   ' Me.Enabled = True
    Call ProcDataByCase("4", strSql)
    'end 2019/11/01

End Sub

Sub StrMenu2()               '相關卷號  frm 100101_3 用
Dim dblRow As Double 'Add By Sindy 2025/9/3

   Me.Enabled = False
   
   lbl1(0).Caption = Me.Tag
    'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
    SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
    SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
    SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
    SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
    SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
    m_AllSys = SystemNumber(Me.Tag, 1)
    intCufaCnt = 0
    'end 2019/11/01
    
   'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
'   strSql = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23) FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
'   strSql = strSql + " union all select PA11,PA05,PA06,PA07,PA22,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
'   strSql = strSql + " union all select SP11,SP05,SP06,SP07,SP14,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
'   strSql = strSql + " union all select '' ,LC05,LC06,LC07,'',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
'   strSql = strSql + " union all select '',HC06,'','','',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   strSql = "SELECT TM12,TM05,TM06,TM07,TM15,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23) as 申請人1 " & SeColTM & _
                                " FROM TRADEMARK,customer WHERE TM01='" & SystemNumber(Me.Tag, 1) & "' AND TM02='" & SystemNumber(Me.Tag, 2) & "' AND TM03='" & SystemNumber(Me.Tag, 3) & "' and TM04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("tm23", "cu")
   strSql = strSql + " union all select PA11,PA05,PA06,PA07,PA22,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) as 申請人1 " & SeColPA & _
                                " FROM PATENT,customer WHERE PA01='" & SystemNumber(Me.Tag, 1) & "' AND PA02='" & SystemNumber(Me.Tag, 2) & "' AND PA03='" & SystemNumber(Me.Tag, 3) & "' and PA04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("pa26", "cu")
   strSql = strSql + " union all select SP11,SP05,SP06,SP07,SP14,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) as 申請人1 " & SeColSP & _
                                " FROM SERVICEPRACTICE,customer WHERE SP01='" & SystemNumber(Me.Tag, 1) & "' AND SP02='" & SystemNumber(Me.Tag, 2) & "' AND SP03='" & SystemNumber(Me.Tag, 3) & "' and SP04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("sp08", "cu")
   strSql = strSql + " union all select '' ,LC05,LC06,LC07,'',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) as 申請人1 " & SeColLC & _
                                " FROM LAWCASE,customer WHERE LC01='" & SystemNumber(Me.Tag, 1) & "' AND LC02='" & SystemNumber(Me.Tag, 2) & "' AND LC03='" & SystemNumber(Me.Tag, 3) & "' and LC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("lc11", "cu")
   strSql = strSql + " union all select '',HC06,'','','',nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) as 申請人1 " & SeColHC & _
                                " FROM HIRECASE,customer WHERE HC01='" & SystemNumber(Me.Tag, 1) & "' AND HC02='" & SystemNumber(Me.Tag, 2) & "' AND HC03='" & SystemNumber(Me.Tag, 3) & "' and HC04='" & SystemNumber(Me.Tag, 4) & "' and " & SQLNewFag("hc05", "cu")
   CheckOC
       
   adoRecordset.CursorLocation = adUseClient
   'Modified by Lydia 2019/11/01 改變型態
   'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
         dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, lbl1(0).Caption, "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109), vbInformation, MsgText(1110)
            End If
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/9/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpNoDataMain
            End If
            adoRecordset.MoveFirst
         Else
            If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/9/3
         End If
        'end 2019/11/01
        
      lbl1(1).Caption = CheckStr(adoRecordset.Fields(0))
      lbl1(2).Caption = CheckStr(adoRecordset.Fields(1))
      lbl1(3).Caption = CheckStr(adoRecordset.Fields(2))
      lbl1(4).Caption = CheckStr(adoRecordset.Fields(3))
      lbl1(6).Caption = CheckStr(adoRecordset.Fields(4))
      lbl1(5).Caption = CheckStr(adoRecordset.Fields(5))
   Else
       If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/3
JumpNoDataMain: 'Added by Lydia 2019/11/01
       lbl1(1).Caption = ""
       lbl1(2).Caption = ""
       lbl1(3).Caption = ""
       lbl1(4).Caption = ""
       lbl1(6).Caption = ""
       Me.Enabled = True
   
   End If
   
   CheckOC
   'edit by nick 2004/11/12
   'strSQL1 = ""
   'strSQL2 = ""
   'StrSQL3 = ""
   'StrSQL4 = ""
   'strSQL5 = ""
   'modify by sonia 2015/3/19 不在此管制可看之系統類別, 否則FCP串P(FMP),外專程序會看不到 FCP-051717串P-111066
   'strSQL1 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 1) & ") "
   'strSQL2 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 2) & ") "
   'StrSQL3 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 3) & ") "
   'StrSQL4 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 4) & ") "
   'strSQL5 = "AND CR05 IN (" & SQLGrpStr(GetSystemKindByNick, 5) & ") "
   strSQL1 = "AND PA01 IS NOT NULL "
   strSQL2 = "AND TM01 IS NOT NULL "
   StrSQL3 = "AND LC01 IS NOT NULL "
   StrSQL4 = "AND HC01 IS NOT NULL "
   strSQL5 = "AND SP01 IS NOT NULL "
   'end 2015/3/19
   
   'Added by Lydia 2019/11/01 利益衝突案件
   m_AllSys = GetAllSysKind(, "ALL")
   intCufaCnt = 0
   'end 2019/11/01
   
   'edit by nick  根本沒連結 2004/11/12
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+) and cr06=lc02(+) and cr07=lc03(+) and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and hc07=hc03(+) and hc08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   'edit by nickc 2006/06/22 改  CASERELATION 1
   'strSQL = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 FROM CASERELATION,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 FROM CASERELATION,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,lawcase,nation,customer                                                                                                                                                                                           WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+)  and cr06=lc02(+) and cr07=lc03(+)   and cr08=lc04(+) and lc15=na01(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 FROM CASERELATION,hirecase,nation,customer                                                                                                                                                                                                                                 WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cr07=hc03(+) and cr08=hc04(+) and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   'strSQL = strSQL & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 FROM CASERELATION,servicepractice,nation,customer                                                                                                                                                            WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   'Modified by Lydia 2019/11/01 +增加欄位SeColTM, SeColPA, SeColSP, SeColLC, SeColHC
   strSql = "SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) AS 申請人,nvl(na03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(pa09,'000',PTM03,PTM04)) 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁," & SQLDate("pa58") & " AS 閉卷日期,PA22 AS 審定專利號數 " & SeColPA & _
                                    " FROM CASERELATION1,patent,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09=na01(+) and '1'=ptm01(+) and pa08=ptm02(+) and pa01 is not null and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁," & SQLDate("tm30") & " AS 閉卷日期,TM15 AS 審定專利號數 " & SeColTM & _
                                    " FROM CASERELATION1,trademark,nation,customer,PATENTTRADEMARKMAP WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=tm01(+) and cr06=tm02(+) and cr07=tm03(+) and cr08=tm04(+) and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and tm01 is not null and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,nvl(lc05,nvl(lc06,lc07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) AS 申請人,nvl(na03,lc15) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("lc09") & " AS 閉卷日期,'' AS 審定專利號數 " & SeColLC & _
                                     " FROM CASERELATION1,lawcase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=lc01(+)  and cr06=lc02(+) and cr07=lc03(+)   and cr08=lc04(+) and lc15=na01(+) and lc01 is not null and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc06 AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) AS 申請人,na03 AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("hc10") & " AS 閉卷日期,'' AS 審定專利號數 " & SeColHC & _
                                        " FROM CASERELATION1,hirecase,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=hc01(+) and cr06=hc02(+) and cr07=hc03(+) and cr08=hc04(+) and hc01 is not null and '000'=na01(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   strSql = strSql & " union SELECT '' AS V,CR05||'-'||CR06||'-'||CR07||'-'||CR08||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) AS 申請人,nvl(na03,sp09) AS 申請國家,'' 專利商標種類,'' AS 目前准駁," & SQLDate("sp16") & " AS 閉卷日期,SP14 AS 審定專利號數 " & SeColSP & _
                                        " FROM CASERELATION1,servicepractice,nation,customer WHERE CR01='" & SystemNumber(Me.Tag, 1) & "' AND CR02='" & SystemNumber(Me.Tag, 2) & "' AND CR03='" & SystemNumber(Me.Tag, 3) & "' and CR04='" & SystemNumber(Me.Tag, 4) & "' and cr05=sp01(+) and cr06=sp02(+) and cr07=sp03(+) and cr08=sp04(+) and sp09=na01(+) and sp01 is not null and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   'end 2019/11/01
   
   strSql = strSql & " order by 1,2" 'Added by Morgan 2023/12/14
   'Modified by Lydia 2019/11/01 改成模組ProcDataByCase
'   CheckOC
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'       cmdOK(0).Enabled = True
'       cmdOK(1).Enabled = True
'       cmdOK(2).Enabled = True
'       cmdOK(3).Enabled = True
'   Else
'       cmdOK(0).Enabled = False
'       cmdOK(1).Enabled = False
'       cmdOK(2).Enabled = False
'       cmdOK(3).Enabled = False
'       ShowNoData
'       Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'       Exit Sub
'   End If
'   Set grdDataList.Recordset = adoRecordset
'   CheckOC
    'Me.Enabled = True
    Call ProcDataByCase("5", strSql)
    'end 2019/11/01
   
End Sub

'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
Private Sub ProcDataByCase(ByVal pType As String, ByRef pSQL As String)
Dim strMid As String, strGrp As String
Dim strJumpList As String '已排除的本所案號
Dim strA1 As String
Dim dblRow As Double 'Add By Sindy 2025/9/3
    
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open pSQL, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If adoRecordset.RecordCount > 0 Then
        dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                strMid = "" & adoRecordset.Fields("本所案號")
                '利益衝突案件：逐案號判斷
                If Len(strMid) > 9 Then
                    If strJumpList <> "" And InStr(strJumpList, strMid) > 0 Then
                        '剔除重複的本所案號
                        adoRecordset.Delete
                    Else
                        If strGrp <> strMid Then
                            If PUB_ChkCufaByCase(Me.Name, m_AllSys, strMid, "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                                strJumpList = strJumpList & strMid & ","
                                intCufaCnt = intCufaCnt + 1
                                adoRecordset.Delete
                            End If
                        End If
                    End If
                End If
                strGrp = strMid
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/9/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/9/3
        End If
        If Val(pType) <= 3 Then
            'StrMenu
        End If
        cmdOK(0).Enabled = True
        cmdOK(1).Enabled = True
        cmdOK(2).Enabled = True
        cmdOK(3).Enabled = True
        Me.Enabled = True
    Else
        If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/3
JumpToNoData:
        cmdOK(0).Enabled = False
        cmdOK(1).Enabled = False
        cmdOK(2).Enabled = False
        cmdOK(3).Enabled = False
        Me.Enabled = True
        If Val(pType) <= 3 Then
            'StrMenu
        End If
        ShowNoData
        If Val(pType) > 3 Then
            'StrMenu1, StrMenu2
            Me.Hide
        End If
        Screen.MousePointer = vbDefault
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
    Set grdDataList.Recordset = adoRecordset
    CheckOC
    
End Sub

