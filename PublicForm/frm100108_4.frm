VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100108_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "關聯案件資料及正聯商標查詢"
   ClientHeight    =   5720
   ClientLeft      =   4380
   ClientTop       =   3150
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9320
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   8508
      TabIndex        =   17
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "母案案件基本資料"
      Height          =   400
      Index           =   2
      Left            =   4170
      TabIndex        =   16
      Top             =   10
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "母案案件進度"
      Height          =   400
      Index           =   3
      Left            =   6000
      TabIndex        =   15
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   4
      Left            =   7530
      TabIndex        =   14
      Top             =   10
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "分割案案件進度"
      Height          =   400
      Index           =   1
      Left            =   2475
      TabIndex        =   13
      Top             =   10
      Width           =   1680
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "分割案案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   450
      TabIndex        =   12
      Top             =   30
      Width           =   2010
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3840
      Left            =   45
      TabIndex        =   0
      Top             =   1860
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   6773
      _Version        =   393216
      Cols            =   9
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
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   7230
      TabIndex        =   18
      Top             =   780
      Width           =   2025
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   2595
      VariousPropertyBits=   27
      Caption         =   " (1.相關卷號2.正聯商標 3.分割案)"
      Size            =   "4577;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "查詢內容："
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   450
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "794;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   8
      Top             =   1281
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
      TabIndex        =   7
      Top             =   727
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
      Left            =   1740
      TabIndex        =   6
      Top             =   1004
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
      Left            =   1080
      TabIndex        =   5
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1281
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/證書號數："
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1004
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   727
      Width           =   900
   End
End
Attribute VB_Name = "frm100108_4"
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
'add by nick 930915  判斷分割還是聯合商標
Public frm100108_txt_7 As String
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
'利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim StrCaseList As String
Dim strDELlist As String


Public Sub SetDataListWidth()
Dim iDep As String
'Added by Lydia 2019/11/01
Dim intField As Integer
intField = 23
grdDataList.Cols = intField
'end 2019/11/01

'edit by nick 2004/09/14
'If frm100108_1.Txt1(7).Text = "4" Then
'edit by nick 2004/09/15
'If frm100108_1.Txt1(7).Text = "3" Then
If frm100108_txt_7 = "3" Then
        grdDataList.row = 0
        grdDataList.col = 0: grdDataList.Text = "V"
        grdDataList.ColWidth(0) = 200
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 1: grdDataList.Text = "母案本所案號"
        grdDataList.ColWidth(1) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
        iDep = PUB_GetST06(strUserNum)
        grdDataList.col = 2: grdDataList.Text = "母案分所號"
        '電腦中心，跟分所才秀
        If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
            grdDataList.ColWidth(2) = 0
        Else
            grdDataList.ColWidth(2) = 1240
        End If
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 3: grdDataList.Text = "母案案件名稱"
        grdDataList.ColWidth(3) = 2500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 4: grdDataList.Text = "商品類別"
        grdDataList.ColWidth(4) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 5: grdDataList.Text = "分割案本所案號"
        grdDataList.ColWidth(5) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 6: grdDataList.Text = "分割案分所號"
        '電腦中心，跟分所才秀
        If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
            grdDataList.ColWidth(6) = 0
        Else
            grdDataList.ColWidth(6) = 1240
        End If
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 7: grdDataList.Text = "分割案案件名稱"
        grdDataList.ColWidth(7) = 2500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 8: grdDataList.Text = "商品類別"
        grdDataList.ColWidth(8) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
Else
        grdDataList.row = 0
        grdDataList.col = 0: grdDataList.Text = "V"
        grdDataList.ColWidth(0) = 200
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 1: grdDataList.Text = "正商標本所案號"
        grdDataList.ColWidth(1) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
        iDep = PUB_GetST06(strUserNum)
        grdDataList.col = 2: grdDataList.Text = "正商標分所號"
        '電腦中心，跟分所才秀
        If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
            grdDataList.ColWidth(2) = 0
        Else
            grdDataList.ColWidth(2) = 1240
        End If
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 3: grdDataList.Text = "正商標案件名稱"
        grdDataList.ColWidth(3) = 2500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 4: grdDataList.Text = "商品類別"
        grdDataList.ColWidth(4) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 5: grdDataList.Text = "聯合商標本所案號"
        grdDataList.ColWidth(5) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 6: grdDataList.Text = "聯合商標分所號"
        '電腦中心，跟分所才秀
        If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
            grdDataList.ColWidth(6) = 0
        Else
            grdDataList.ColWidth(6) = 1240
        End If
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 7: grdDataList.Text = "聯合商標案件名稱"
        grdDataList.ColWidth(7) = 2500
        grdDataList.CellAlignment = flexAlignCenterCenter
        grdDataList.col = 8: grdDataList.Text = "商品類別"
        grdDataList.ColWidth(8) = 1500
        grdDataList.CellAlignment = flexAlignCenterCenter
End If

'Added by Lydia 2019/11/01  隱藏欄位：母案案號,申請人1~5, FC代理人/ 子案案號,申請人1~5, FC代理人
For intI = 9 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI

End Sub

'92.04.16 nick
Public Sub PubShowNextData()
        Dim Str01 As String
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
        grdDataList.col = 5
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
                  'add by nickc 2006/10/18
                  frm100101_3.cmdOK(5).Visible = False
                  Screen.MousePointer = vbDefault
            Case "CFT", "FCT", "T", "TF"   '商標
                  Screen.MousePointer = vbHourglass
                  frm100101_4.Show
                  frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_4.StrMenu
                  'add by nickc 2006/10/18
                  frm100101_4.cmdOK(5).Visible = False
                  Screen.MousePointer = vbDefault
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
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
         grdDataList.col = 5
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            'add by nickc 2006/10/18
            frm100101_2.cmdOK(8).Visible = False
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 2
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
                  'add by nickc 2006/10/18
                  frm100101_3.cmdOK(5).Visible = False
                  Screen.MousePointer = vbDefault
            Case "CFT", "FCT", "T", "TF"   '商標
                  Screen.MousePointer = vbHourglass
                  frm100101_4.Show
                  frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_4.StrMenu
                  'add by nickc 2006/10/18
                  frm100101_4.cmdOK(5).Visible = False
                  Screen.MousePointer = vbDefault
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
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
Case 3
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
            'add by nickc 2006/10/18
            frm100101_2.cmdOK(8).Visible = False
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
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
'edit by nick 2004/09/15 移除
'SetDataListWidth
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()                 '分割案 'Memo by Lydia 2019/11/01 從frm100108_1,frm100108_2 過來
Dim strTM27 As String '正商標號數
Dim strTM12 As String '申請案號
Dim Strsql20040730 As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

   Me.Enabled = False
   '畫面上方改成輸入畫面的條件
   If frm100108_1.Option1(0).Value = True Then
       'lbl1(0).Caption = frm100108_1.txt1(0).Text & "-" & frm100108_1.txt1(1).Text & "-" & frm100108_1.txt1(2).Text & "-" & frm100108_1.txt1(3).Text
       lbl1(0).Caption = frm100108_1.txt1(0).Text & "-" & frm100108_1.txt1(1).Text & "-" & IIf(Len(Trim(frm100108_1.txt1(2).Text)) = 0, "0", frm100108_1.txt1(2).Text) & "-" & IIf(Len(Trim(frm100108_1.txt1(3).Text)) = 0, "00", frm100108_1.txt1(3).Text)
   Else
       lbl1(0).Caption = ""
   End If
   If frm100108_1.Option1(1).Value = True Then
       lbl1(1).Caption = frm100108_1.txt1(4).Text
   Else
       lbl1(1).Caption = ""
   End If
   If frm100108_1.Option1(2).Value = True Then
       lbl1(6).Caption = frm100108_1.txt1(5).Text
   Else
       lbl1(6).Caption = ""
   End If
   lbl1(2).Caption = frm100108_1.txt1(6).Text
   lbl1(3).Caption = frm100108_1.txt1(7).Text
   
   Call SetCUFA(1) 'Added by Lydia 2019/11/01 利益衝突案件：預設
    
'add by nick 2004/07/30 加入檢查是母案或分割案
Dim IsMother As Boolean
Dim s As Integer
   'edit by nick 2004/09/14
   'If frm100108_1.Txt1(7) = "4" Then
   If frm100108_1.txt1(7) = "3" Then  '前一畫面: 分割案
        strSql = "select AA.A,AA.B,BB.A,BB.B from ("
        strSql = strSql & "select '1' A,count(*) B from divisioncase where dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' ) AA,"
        strSql = strSql & " (select '2' A,count(*) B from divisioncase where dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "') BB "
        CheckOC
        
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If Val(CheckStr(adoRecordset.Fields(1).Value)) = 0 And Val(CheckStr(adoRecordset.Fields(3).Value)) = 0 Then
                InsertQueryLog (0) 'Add By Sindy 2010/11/3
                s = MsgBox("資料庫中搜尋不到此案有分割!!", , "沒有資料")
                cmdOK(0).Enabled = False
                cmdOK(1).Enabled = False
                cmdOK(2).Enabled = False
                cmdOK(3).Enabled = False
                Me.Enabled = True
                Screen.MousePointer = vbDefault
                 tmpBol = fnCancelNowFormAndShowParentForm(Me)
                Exit Sub
            End If
            If Val(CheckStr(adoRecordset.Fields(1).Value)) > 0 Then
                IsMother = False
            Else
                If Val(CheckStr(adoRecordset.Fields(3).Value)) > 0 Then
                    IsMother = True
                Else
                    InsertQueryLog (0) 'Add By Sindy 2010/11/3
                    s = MsgBox("資料庫中搜尋不到此案有分割!!", , "沒有資料")
                    cmdOK(0).Enabled = False
                    cmdOK(1).Enabled = False
                    cmdOK(2).Enabled = False
                    cmdOK(3).Enabled = False
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                     tmpBol = fnCancelNowFormAndShowParentForm(Me)
                    Exit Sub
                End If
            End If
        End If
        CheckOC
        If IsMother = True Then  '抓母案
            strSQL1 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL2 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL3 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL4 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL5 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            'Modified by Lydia 2019/11/01
            'strSQL1 = strSQL1 & " and T1.PA01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
            'strSQL2 = strSQL2 & " and T1.TM01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
            'StrSQL3 = StrSQL3 & " and T1.LC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 3) & ") "
            'StrSQL4 = StrSQL4 & " and T1.HC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 4) & ") "
            'strSQL5 = strSQL5 & " and T1.SP01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
            strSQL1 = strSQL1 & " and T1.PA01 in (" & SQLGrpStr(m_AllSys, 1) & ") "
            strSQL2 = strSQL2 & " and T1.TM01 in (" & SQLGrpStr(m_AllSys, 2) & ") "
            StrSQL3 = StrSQL3 & " and T1.LC01 in (" & SQLGrpStr(m_AllSys, 3) & ") "
            StrSQL4 = StrSQL4 & " and T1.HC01 in (" & SQLGrpStr(m_AllSys, 4) & ") "
            strSQL5 = strSQL5 & " and T1.SP01 in (" & SQLGrpStr(m_AllSys, 5) & ") "
            'end 2019/11/01
        Else     '抓子案
            strSQL1 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL2 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL3 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL4 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL5 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            'Modified by Lydia 2019/11/01
            'strSQL1 = strSQL1 & " and T2.PA01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
            'strSQL2 = strSQL2 & " and T2.TM01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
            'StrSQL3 = StrSQL3 & " and T2.LC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 3) & ") "
            'StrSQL4 = StrSQL4 & " and T2.HC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 4) & ") "
            'strSQL5 = strSQL5 & " and T2.SP01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
            strSQL1 = strSQL1 & " and T2.PA01 in (" & SQLGrpStr(m_AllSys, 1) & ") "
            strSQL2 = strSQL2 & " and T2.TM01 in (" & SQLGrpStr(m_AllSys, 2) & ") "
            StrSQL3 = StrSQL3 & " and T2.LC01 in (" & SQLGrpStr(m_AllSys, 3) & ") "
            StrSQL4 = StrSQL4 & " and T2.HC01 in (" & SQLGrpStr(m_AllSys, 4) & ") "
            strSQL5 = strSQL5 & " and T2.SP01 in (" & SQLGrpStr(m_AllSys, 5) & ") "
            '2019/11/01
        End If
   Else   '前一畫面: 正聯商標 =>  txt1(7) = "2") And (txt1(0) = "CFT" Or txt1(0) = "FCT" Or txt1(0) = "T" Or txt1(0) = "TF")
        strSql = "select AAA.A,AAA.B,BBB.A,BBB.B from ("
        strSql = strSql & "select '1' A,count(*) B from (SELECT AA.tm01 tt1,aa.tm02 tt2,aa.tm03 tt3,aa.tm04 tt4,aa.tm08 tt5,aa.tm15 tt6,"
        strSql = strSql & " bb.tm01 tt7,bb.tm02 tt8,bb.tm03 tt9,bb.tm04 tt10,bb.tm08 tt11,bb.tm27 tt12 FROM"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM15 FROM TRADEMARK WHERE TM08='1' AND TM15 IS NOT NULL) aa,"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM27 FROM TRADEMARK WHERE TM08='2' AND TM27 IS NOT NULL) bb"
        strSql = strSql & " Where aa.TM15 = bb.tm27"
        strSql = strSql & " union SELECT aa.tm01 tt1,aa.tm02 tt2,aa.tm03 tt3,aa.tm04 tt4,aa.tm08 tt5,aa.tm15 tt6,"
        strSql = strSql & " bb.tm01 tt7,bb.tm02 tt8,bb.tm03 tt9,bb.tm04 tt10,bb.tm08 tt11,bb.tm27 tt12 FROM"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM15 FROM TRADEMARK WHERE TM08='4' AND TM15 IS NOT NULL) aa,"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM27 FROM TRADEMARK WHERE TM08='5' AND TM27 IS NOT NULL) bb"
        strSql = strSql & " where aa.tm15=bb.tm27) divisioncase"
        strSql = strSql & " where tt1='" & SystemNumber(Me.Tag, 1) & "' and tt2='" & SystemNumber(Me.Tag, 2) & "' and tt3='" & SystemNumber(Me.Tag, 3) & "' and tt4='" & SystemNumber(Me.Tag, 4) & "' ) AAA,"
        strSql = strSql & " (select '2' A,count(*) B from (SELECT AA.tm01 tt1,aa.tm02 tt2,aa.tm03 tt3,aa.tm04 tt4,aa.tm08 tt5,aa.tm15 tt6,"
        strSql = strSql & " bb.tm01 tt7,bb.tm02 tt8,bb.tm03 tt9,bb.tm04 tt10,bb.tm08 tt11,bb.tm27 tt12 FROM"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM15 FROM TRADEMARK WHERE TM08='1' AND TM15 IS NOT NULL) aa,"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM27 FROM TRADEMARK WHERE TM08='2' AND TM27 IS NOT NULL) bb"
        strSql = strSql & " Where aa.TM15 = bb.tm27"
        strSql = strSql & " union SELECT aa.tm01 tt1,aa.tm02 tt2,aa.tm03 tt3,aa.tm04 tt4,aa.tm08 tt5,aa.tm15 tt6,"
        strSql = strSql & " bb.tm01 tt7,bb.tm02 tt8,bb.tm03 tt9,bb.tm04 tt10,bb.tm08 tt11,bb.tm27 tt12 FROM"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM15 FROM TRADEMARK WHERE TM08='4' AND TM15 IS NOT NULL) aa,"
        strSql = strSql & " (SELECT TM01,TM02,TM03,TM04,TM08,TM27 FROM TRADEMARK WHERE TM08='5' AND TM27 IS NOT NULL) bb"
        strSql = strSql & " where aa.tm15=bb.tm27) divisioncase"
        strSql = strSql & " where tt7='" & SystemNumber(Me.Tag, 1) & "' and tt8='" & SystemNumber(Me.Tag, 2) & "' and tt9='" & SystemNumber(Me.Tag, 3) & "' and tt10='" & SystemNumber(Me.Tag, 4) & "') BBB "
        CheckOC
        
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If Val(CheckStr(adoRecordset.Fields(1).Value)) = 0 And Val(CheckStr(adoRecordset.Fields(3).Value)) = 0 Then
                InsertQueryLog (0) 'Add By Sindy 2010/11/3
                s = MsgBox("資料庫中搜尋不到此案有聯合商標或正商標!!", , "沒有資料")
                cmdOK(0).Enabled = False
                cmdOK(1).Enabled = False
                cmdOK(2).Enabled = False
                cmdOK(3).Enabled = False
                Me.Enabled = True
                Screen.MousePointer = vbDefault
                 tmpBol = fnCancelNowFormAndShowParentForm(Me)
                Exit Sub
            End If
            If Val(CheckStr(adoRecordset.Fields(1).Value)) > 0 Then
                IsMother = False
            Else
                If Val(CheckStr(adoRecordset.Fields(3).Value)) > 0 Then
                    IsMother = True
                Else
                    InsertQueryLog (0) 'Add By Sindy 2010/11/3
                    s = MsgBox("資料庫中搜尋不到此案有聯合商標或正商標!!", , "沒有資料")
                    cmdOK(0).Enabled = False
                    cmdOK(1).Enabled = False
                    cmdOK(2).Enabled = False
                    cmdOK(3).Enabled = False
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                     tmpBol = fnCancelNowFormAndShowParentForm(Me)
                    Exit Sub
                End If
            End If
        End If
        CheckOC
        If IsMother = True Then '抓母案
            strSQL1 = " and tt7='" & SystemNumber(Me.Tag, 1) & "' and tt8='" & SystemNumber(Me.Tag, 2) & "' and tt9='" & SystemNumber(Me.Tag, 3) & "' and tt10='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL2 = " and tt7='" & SystemNumber(Me.Tag, 1) & "' and tt8='" & SystemNumber(Me.Tag, 2) & "' and tt9='" & SystemNumber(Me.Tag, 3) & "' and tt10='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL3 = " and tt7='" & SystemNumber(Me.Tag, 1) & "' and tt8='" & SystemNumber(Me.Tag, 2) & "' and tt9='" & SystemNumber(Me.Tag, 3) & "' and tt10='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL4 = " and tt7='" & SystemNumber(Me.Tag, 1) & "' and tt8='" & SystemNumber(Me.Tag, 2) & "' and tt9='" & SystemNumber(Me.Tag, 3) & "' and tt10='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL5 = " and tt7='" & SystemNumber(Me.Tag, 1) & "' and tt8='" & SystemNumber(Me.Tag, 2) & "' and tt9='" & SystemNumber(Me.Tag, 3) & "' and tt10='" & SystemNumber(Me.Tag, 4) & "' "
            'Modified by Lydia 2019/11/01
            'strSQL1 = strSQL1 & " and T1.PA01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
            'strSQL2 = strSQL2 & " and T1.TM01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
            'StrSQL3 = StrSQL3 & " and T1.LC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 3) & ") "
            'StrSQL4 = StrSQL4 & " and T1.HC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 4) & ") "
            'strSQL5 = strSQL5 & " and T1.SP01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
            strSQL1 = strSQL1 & " and T1.PA01 in (" & SQLGrpStr(m_AllSys, 1) & ") "
            strSQL2 = strSQL2 & " and T1.TM01 in (" & SQLGrpStr(m_AllSys, 2) & ") "
            StrSQL3 = StrSQL3 & " and T1.LC01 in (" & SQLGrpStr(m_AllSys, 3) & ") "
            StrSQL4 = StrSQL4 & " and T1.HC01 in (" & SQLGrpStr(m_AllSys, 4) & ") "
            strSQL5 = strSQL5 & " and T1.SP01 in (" & SQLGrpStr(m_AllSys, 5) & ") "
            'end 2019/11/01
        Else                '抓子案
            strSQL1 = " and tt1='" & SystemNumber(Me.Tag, 1) & "' and tt2='" & SystemNumber(Me.Tag, 2) & "' and tt3='" & SystemNumber(Me.Tag, 3) & "' and tt4='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL2 = " and tt1='" & SystemNumber(Me.Tag, 1) & "' and tt2='" & SystemNumber(Me.Tag, 2) & "' and tt3='" & SystemNumber(Me.Tag, 3) & "' and tt4='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL3 = " and tt1='" & SystemNumber(Me.Tag, 1) & "' and tt2='" & SystemNumber(Me.Tag, 2) & "' and tt3='" & SystemNumber(Me.Tag, 3) & "' and tt4='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL4 = " and tt1='" & SystemNumber(Me.Tag, 1) & "' and tt2='" & SystemNumber(Me.Tag, 2) & "' and tt3='" & SystemNumber(Me.Tag, 3) & "' and tt4='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL5 = " and tt1='" & SystemNumber(Me.Tag, 1) & "' and tt2='" & SystemNumber(Me.Tag, 2) & "' and tt3='" & SystemNumber(Me.Tag, 3) & "' and tt4='" & SystemNumber(Me.Tag, 4) & "' "
            'Modified by Lydia 2019/11/01
            'strSQL1 = strSQL1 & " and T2.PA01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 1) & ") "
            'strSQL2 = strSQL2 & " and T2.TM01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 2) & ") "
            'StrSQL3 = StrSQL3 & " and T2.LC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 3) & ") "
            'StrSQL4 = StrSQL4 & " and T2.HC01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 4) & ") "
            'strSQL5 = strSQL5 & " and T2.SP01 in (" & SQLGrpStr(IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6))), 5) & ") "
            strSQL1 = strSQL1 & " and T2.PA01 in (" & SQLGrpStr(m_AllSys, 1) & ") "
            strSQL2 = strSQL2 & " and T2.TM01 in (" & SQLGrpStr(m_AllSys, 2) & ") "
            StrSQL3 = StrSQL3 & " and T2.LC01 in (" & SQLGrpStr(m_AllSys, 3) & ") "
            StrSQL4 = StrSQL4 & " and T2.HC01 in (" & SQLGrpStr(m_AllSys, 4) & ") "
            strSQL5 = strSQL5 & " and T2.SP01 in (" & SQLGrpStr(m_AllSys, 5) & ") "
            'end 2019/11/01
        End If
   End If
   strTemp = SystemNumber(Me.Tag, 1)
   'edit by nick 2004/09/14
   'If frm100108_1.Txt1(7).Text = "4" Then
   If frm100108_1.txt1(7).Text = "3" Then   '前一畫面: 分割案
        'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
        strSql = "SELECT '' AS V,T1.PA01||'-'||T1.PA02||'-'||T1.PA03||'-'||T1.PA04||DECODE(length(nvl(t1.pa108,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.pa136,'')),null,'','●')||t1.pa47 as 母案分所號,NVL(T1.PA05,NVL(T1.PA06,T1.PA07)) AS 母案案件名稱,''  as 商品類別," & _
                    "T2.PA01||'-'||T2.PA02||'-'||T2.PA03||'-'||T2.PA04||DECODE(length(nvl(t2.pa108,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.pa136,'')),null,'','●')||t2.pa47 as 分割案分所號,NVL(T2.PA05,NVL(T2.PA06,T2.PA07)) AS 分割案案件名稱,''  as 商品類別" & SeColPA & _
                    "FROM PATENT T1,PATENT T2,DIVISIONcase " & _
                    "WHERE DC01=T2.PA01(+) AND DC02=T2.PA02(+) AND DC03=T2.PA03(+) AND DC04=T2.PA04(+) AND DC05=T1.PA01(+) AND DC06=T1.PA02(+) AND DC07=T1.PA03(+) AND DC08=T1.PA04(+) " & strSQL1
        strSql = strSql & " union SELECT '' AS V,T1.TM01||'-'||T1.TM02||'-'||T1.TM03||'-'||T1.TM04||DECODE(length(nvl(t1.tm57,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.tm73,'')),null,'','●')||t1.tm34 as 母案分所號,NVL(T1.TM05,NVL(T1.TM06,T1.TM07)) AS 母案案件名稱,T1.tm09 as 商品類別," & _
                    "T2.TM01||'-'||T2.TM02||'-'||T2.TM03||'-'||T2.TM04||DECODE(length(nvl(t2.tm57,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.tm73,'')),null,'','●')||t2.tm34 as 分割案分所號,NVL(T2.TM05,NVL(T2.TM06,T2.TM07)) AS 分割案案件名稱,T2.tm09 as 商品類別" & SeColTM & _
                    "FROM TRADEMARK T1,TRADEMARK T2,DIVISIONcase " & _
                    "WHERE DC01=T2.TM01(+) AND DC02=T2.TM02(+) AND DC03=T2.TM03(+) AND DC04=T2.TM04(+) AND DC05=T1.TM01(+) AND DC06=T1.TM02(+) AND DC07=T1.TM03(+) AND DC08=T1.TM04(+) " & strSQL2
        strSql = strSql & " union SELECT '' AS V,T1.LC01||'-'||T1.LC02||'-'||T1.LC03||'-'||T1.LC04||DECODE(length(nvl(t1.lc34,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.lc36,'')),null,'','●')||t1.lc16 as 母案分所號,NVL(T1.LC05,NVL(T1.LC06,T1.LC07)) AS 母案案件名稱,'' as 商品類別," & _
                    "T2.LC01||'-'||T2.LC02||'-'||T2.LC03||'-'||T2.LC04||DECODE(length(nvl(t2.lc34,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.lc36,'')),null,'','●')||t2.lc16 as 分割案分所號,NVL(T2.LC05,NVL(T2.LC06,T2.LC07)) AS 分割案案件名稱,''  as 商品類別" & SeColLC & _
                    "FROM LAWCASE T1,LAWCASE T2,DIVISIONcase " & _
                    "WHERE DC01=T2.LC01(+) AND DC02=T2.LC02(+) AND DC03=T2.LC03(+) AND DC04=T2.LC04(+) AND DC05=T1.LC01(+) AND DC06=T1.LC02(+) AND DC07=T1.LC03(+) AND DC08=T1.LC04(+) " & StrSQL3
        strSql = strSql & " union SELECT '' AS V,T1.HC01||'-'||T1.HC02||'-'||T1.HC03||'-'||T1.HC04||DECODE(length(nvl(t1.hc19,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.hc20,'')),null,'','●')||t1.hc07 as 母案分所號,T1.HC06 AS 母案案件名稱,'' as 商品類別," & _
                    "T2.HC01||'-'||T2.HC02||'-'||T2.HC03||'-'||T2.HC04||DECODE(length(nvl(t2.hc19,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.hc20,'')),null,'','●')||t2.hc07 as 分割案分所號,T2.HC06 AS 分割案案件名稱,''  as 商品類別" & SeColHC & _
                    "FROM HIRECASE T1,HIRECASE T2,DIVISIONcase " & _
                     "WHERE DC01=T2.HC01(+) AND DC02=T2.HC02(+) AND DC03=T2.HC03(+) AND DC04=T2.HC04(+) AND DC05=T1.HC01(+) AND DC06=T1.HC02(+) AND DC07=T1.HC03(+) AND DC08=T1.HC04(+) " & StrSQL4
        strSql = strSql & " union SELECT '' AS V,T1.SP01||'-'||T1.SP02||'-'||T1.SP03||'-'||T1.SP04||DECODE(length(nvl(t1.sp61,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.sp68,'')),null,'','●')||t1.sp28 as 母案分所號,NVL(T1.SP05,NVL(T1.SP06,T1.SP07)) AS 母案案件名稱,'' as 商品類別," & _
                    "T2.SP01||'-'||T2.SP02||'-'||T2.SP03||'-'||T2.SP04||DECODE(length(nvl(t2.sp61,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.sp68,'')),null,'','●')||t2.sp28 as 分割案分所號,NVL(T2.SP05,NVL(T2.SP06,T2.SP07)) AS 分割案案件名稱,''  as 商品類別" & SeColSP & _
                    "FROM SERVICEPRACTICE T1,SERVICEPRACTICE T2,DIVISIONcase " & _
                    "WHERE DC01=T2.SP01(+) AND DC02=T2.SP02(+) AND DC03=T2.SP03(+) AND DC04=T2.SP04(+) AND DC05=T1.SP01(+) AND DC06=T1.SP02(+) AND DC07=T1.SP03(+) AND DC08=T1.SP04(+) " & strSQL5
   Else      '前一畫面: 正聯商標
        Strsql20040730 = "(SELECT AA.tm01 tt1,aa.tm02 tt2,aa.tm03 tt3,aa.tm04 tt4,aa.tm08 tt5,aa.tm15 tt6,"
        Strsql20040730 = Strsql20040730 & " bb.tm01 tt7,bb.tm02 tt8,bb.tm03 tt9,bb.tm04 tt10,bb.tm08 tt11,bb.tm27 tt12 FROM"
        Strsql20040730 = Strsql20040730 & " (SELECT TM01,TM02,TM03,TM04,TM08,TM15 FROM TRADEMARK WHERE TM08='1' AND TM15 IS NOT NULL) aa,"
        Strsql20040730 = Strsql20040730 & " (SELECT TM01,TM02,TM03,TM04,TM08,TM27 FROM TRADEMARK WHERE TM08='2' AND TM27 IS NOT NULL) bb"
        Strsql20040730 = Strsql20040730 & " Where aa.TM15 = bb.tm27"
        Strsql20040730 = Strsql20040730 & " union SELECT aa.tm01 tt1,aa.tm02 tt2,aa.tm03 tt3,aa.tm04 tt4,aa.tm08 tt5,aa.tm15 tt6,"
        Strsql20040730 = Strsql20040730 & " bb.tm01 tt7,bb.tm02 tt8,bb.tm03 tt9,bb.tm04 tt10,bb.tm08 tt11,bb.tm27 tt12 FROM"
        Strsql20040730 = Strsql20040730 & " (SELECT TM01,TM02,TM03,TM04,TM08,TM15 FROM TRADEMARK WHERE TM08='4' AND TM15 IS NOT NULL) aa,"
        Strsql20040730 = Strsql20040730 & " (SELECT TM01,TM02,TM03,TM04,TM08,TM27 FROM TRADEMARK WHERE TM08='5' AND TM27 IS NOT NULL) bb"
        Strsql20040730 = Strsql20040730 & " where aa.tm15=bb.tm27) divisioncase"

        'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
        strSql = "SELECT '' AS V,T1.PA01||'-'||T1.PA02||'-'||T1.PA03||'-'||T1.PA04||DECODE(length(nvl(t1.pa108,'')),null,'','●') AS 正商標本所案號,DECODE(length(nvl(t1.pa136,'')),null,'','●')||t1.pa47 as 正商標分所號,NVL(T1.PA05,NVL(T1.PA06,T1.PA07)) AS 正商標案件名稱,''  as 商品類別," & _
                    "T2.PA01||'-'||T2.PA02||'-'||T2.PA03||'-'||T2.PA04||DECODE(length(nvl(t2.pa108,'')),null,'','●') AS 聯合商標本所案號,DECODE(length(nvl(t2.pa136,'')),null,'','●')||t2.pa47 as 聯合商標分所號,NVL(T2.PA05,NVL(T2.PA06,T2.PA07)) AS 聯合商標案件名稱,''  as 商品類別" & SeColPA & _
                    "FROM PATENT T1,PATENT T2," & Strsql20040730 & " " & _
                    "WHERE tt1=T2.PA01(+) AND tt2=T2.PA02(+) AND tt3=T2.PA03(+) AND tt4=T2.PA04(+) AND tt7=T1.PA01(+) AND tt8=T1.PA02(+) AND tt9=T1.PA03(+) AND tt10=T1.PA04(+) " & strSQL1
        strSql = strSql & " union SELECT '' AS V,T1.TM01||'-'||T1.TM02||'-'||T1.TM03||'-'||T1.TM04||DECODE(length(nvl(t1.tm57,'')),null,'','●') AS 正商標本所案號,DECODE(length(nvl(t1.tm73,'')),null,'','●')||t1.tm34 as 正商標分所號,NVL(T1.TM05,NVL(T1.TM06,T1.TM07)) AS 正商標案件名稱,T1.tm09 as 商品類別," & _
                    "T2.TM01||'-'||T2.TM02||'-'||T2.TM03||'-'||T2.TM04||DECODE(length(nvl(t2.tm57,'')),null,'','●') AS 聯合商標本所案號,DECODE(length(nvl(t2.tm73,'')),null,'','●')||t2.tm34 as 聯合商標分所號,NVL(T2.TM05,NVL(T2.TM06,T2.TM07)) AS 聯合商標案件名稱,T2.tm09 as 商品類別 " & SeColTM & _
                    "FROM TRADEMARK T1,TRADEMARK T2," & Strsql20040730 & " " & _
                    "WHERE tt1=T2.TM01(+) AND tt2=T2.TM02(+) AND tt3=T2.TM03(+) AND tt4=T2.TM04(+) AND tt7=T1.TM01(+) AND tt8=T1.TM02(+) AND tt9=T1.TM03(+) AND tt10=T1.TM04(+) " & strSQL2
        strSql = strSql & " union SELECT '' AS V,T1.LC01||'-'||T1.LC02||'-'||T1.LC03||'-'||T1.LC04||DECODE(length(nvl(t1.lc34,'')),null,'','●') AS 正商標本所案號,DECODE(length(nvl(t1.lc36,'')),null,'','●')||t1.lc16 as 正商標分所號,NVL(T1.LC05,NVL(T1.LC06,T1.LC07)) AS 正商標案件名稱,'' as 商品類別," & _
                    "T2.LC01||'-'||T2.LC02||'-'||T2.LC03||'-'||T2.LC04||DECODE(length(nvl(t2.lc34,'')),null,'','●') AS 聯合商標本所案號,DECODE(length(nvl(t2.lc36,'')),null,'','●')||t2.lc16 as 聯合商標分所號,NVL(T2.LC05,NVL(T2.LC06,T2.LC07)) AS 聯合商標案件名稱,''  as 商品類別" & SeColLC & _
                    "FROM LAWCASE T1,LAWCASE T2," & Strsql20040730 & " " & _
                    "WHERE tt1=T2.LC01(+) AND tt2=T2.LC02(+) AND tt3=T2.LC03(+) AND tt4=T2.LC04(+) AND tt7=T1.LC01(+) AND tt8=T1.LC02(+) AND tt9=T1.LC03(+) AND tt10=T1.LC04(+) " & StrSQL3
        strSql = strSql & " union SELECT '' AS V,T1.HC01||'-'||T1.HC02||'-'||T1.HC03||'-'||T1.HC04||DECODE(length(nvl(t1.hc19,'')),null,'','●') AS 正商標本所案號,DECODE(length(nvl(t1.hc20,'')),null,'','●')||t1.hc07 as 正商標分所號,T1.HC06 AS 正商標案件名稱,'' as 商品類別," & _
                    "T2.HC01||'-'||T2.HC02||'-'||T2.HC03||'-'||T2.HC04||DECODE(length(nvl(t2.hc19,'')),null,'','●') AS 聯合商標本所案號,DECODE(length(nvl(t2.hc20,'')),null,'','●')||t2.hc07 as 聯合商標分所號,T2.HC06 AS 聯合商標案件名稱,''  as 商品類別" & SeColHC & _
                    "FROM HIRECASE T1,HIRECASE T2," & Strsql20040730 & " " & _
                    "WHERE tt1=T2.HC01(+) AND tt2=T2.HC02(+) AND tt3=T2.HC03(+) AND tt4=T2.HC04(+) AND tt7=T1.HC01(+) AND tt8=T1.HC02(+) AND tt9=T1.HC03(+) AND tt10=T1.HC04(+) " & StrSQL4
        strSql = strSql & " union SELECT '' AS V,T1.SP01||'-'||T1.SP02||'-'||T1.SP03||'-'||T1.SP04||DECODE(length(nvl(t1.sp61,'')),null,'','●') AS 正商標本所案號,DECODE(length(nvl(t1.sp68,'')),null,'','●')||t1.sp28 as 正商標分所號,NVL(T1.SP05,NVL(T1.SP06,T1.SP07)) AS 正商標案件名稱,'' as 商品類別," & _
                    "T2.SP01||'-'||T2.SP02||'-'||T2.SP03||'-'||T2.SP04||DECODE(length(nvl(t2.sp61,'')),null,'','●') AS 聯合商標本所案號,DECODE(length(nvl(t2.sp68,'')),null,'','●')||t2.sp28 as 聯合商標分所號,NVL(T2.SP05,NVL(T2.SP06,T2.SP07)) AS 聯合商標案件名稱,''  as 商品類別" & SeColSP & _
                    "FROM SERVICEPRACTICE T1,SERVICEPRACTICE T2," & Strsql20040730 & " " & _
                    "WHERE tt1=T2.SP01(+) AND tt2=T2.SP02(+) AND tt3=T2.SP03(+) AND tt4=T2.SP04(+) AND tt7=T1.SP01(+) AND tt8=T1.SP02(+) AND tt9=T1.SP03(+) AND tt10=T1.SP04(+) " & strSQL5
   End If
   adoRecordset.CursorLocation = adUseClient
   'Modified by Lydia 2019/11/01 改變型態
   'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      
    dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

    'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            If "" & adoRecordset.Fields("casenoA") <> "" Then
                If strDELlist <> "" And InStr(strDELlist, "" & adoRecordset.Fields("casenoA")) > 0 Then '已檢查過並且為限閱案件
                    adoRecordset.Delete
                    GoTo JumpToNext
                Else
                    If InStr(StrCaseList, "" & adoRecordset.Fields("casenoA")) = 0 Then
                        StrCaseList = StrCaseList & adoRecordset.Fields("casenoA") & ","   '記錄已檢查過的案號
                        If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("casenoA"), "" & adoRecordset.Fields("custA1") & "," & adoRecordset.Fields("custA2") & "," & adoRecordset.Fields("custA3") & "," & adoRecordset.Fields("custA4") & "," & adoRecordset.Fields("custA5"), "" & adoRecordset.Fields("fcnoA")) = False Then
                            strDELlist = strDELlist & adoRecordset.Fields("casenoA") '記錄已檢查過並且為限閱案件的案號
                            intCufaCnt = intCufaCnt + 1
                            adoRecordset.Delete
                            GoTo JumpToNext
                        End If
                    End If
                End If
            End If
            
            If "" & adoRecordset.Fields("casenoB") <> "" Then
                If strDELlist <> "" And InStr(strDELlist, "" & adoRecordset.Fields("casenoB")) > 0 Then '已檢查過並且為限閱案件
                    adoRecordset.Delete
                    GoTo JumpToNext
                Else
                    If InStr(StrCaseList, "" & adoRecordset.Fields("casenoB")) = 0 Then
                        StrCaseList = StrCaseList & adoRecordset.Fields("casenoB") & ","   '記錄已檢查過的案號
                        If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("casenoB"), "" & adoRecordset.Fields("custB1") & "," & adoRecordset.Fields("custB2") & "," & adoRecordset.Fields("custB3") & "," & adoRecordset.Fields("custB4") & "," & adoRecordset.Fields("custB5"), "" & adoRecordset.Fields("fcnoB")) = False Then
                            strDELlist = strDELlist & adoRecordset.Fields("casenoB") '記錄已檢查過並且為限閱案件的案號
                            intCufaCnt = intCufaCnt + 1
                            adoRecordset.Delete
                            GoTo JumpToNext
                        End If
                    End If
                End If
            End If
JumpToNext:
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
    Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
    End If
    'end 2019/11/01
   
      cmdOK(0).Enabled = True
      cmdOK(1).Enabled = True
      cmdOK(2).Enabled = True
      cmdOK(3).Enabled = True
      Set grdDataList.Recordset = adoRecordset
      CheckOC
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
      'edit by nick 2004/09/14
      'If frm100108_1.Txt1(7).Text = "4" Then
      If frm100108_1.txt1(7).Text = "3" Then
          s = MsgBox("此本所案號沒有分割案, 無法查詢!!" & Me.Tag & "  ", , "錯誤")
      Else
          s = MsgBox("此本所案號沒有聯合商標或正商標, 無法查詢!!" & Me.Tag & "  ", , "錯誤")
      End If
      cmdOK(0).Enabled = False
      cmdOK(1).Enabled = False
      cmdOK(2).Enabled = False
      cmdOK(3).Enabled = False
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
    End If

   Me.Enabled = True
End Sub

Sub StrMenu1()                 '分割案 'Memo by Lydia 2019/11/01 從frm100101_2, frm100101_3, frm100101_4 過來
Dim strTM27 As String '正商標號數
Dim strTM12 As String '申請案號
Dim Strsql20040730 As String
Me.Enabled = False
'畫面上方改成輸入畫面的條件

lbl1(0).Caption = SystemNumber(Me.Tag, 1) & "-" & SystemNumber(Me.Tag, 2) & "-" & SystemNumber(Me.Tag, 3) & "-" & SystemNumber(Me.Tag, 4)
lbl1(2).Caption = SystemNumber(Me.Tag, 1)
lbl1(3).Caption = "3"

    Call SetCUFA(2) 'Added by Lydia 2019/11/01 利益衝突案件：預設
    
'add by nick 2004/07/30 加入檢查是母案或分割案
Dim IsMother As Boolean
Dim IsChild As Boolean 'Add By Sindy 2010/6/29
Dim s As Integer
        
        strSql = "select AA.A,AA.B,BB.A,BB.B from ("
        strSql = strSql & "select '1' A,count(*) B from divisioncase where dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' ) AA,"
        strSql = strSql & " (select '2' A,count(*) B from divisioncase where dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "') BB "
        CheckOC
        
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If Val(CheckStr(adoRecordset.Fields(1).Value)) = 0 And Val(CheckStr(adoRecordset.Fields(3).Value)) = 0 Then
                s = MsgBox("資料庫中搜尋不到此案有分割!!", , "沒有資料")
                cmdOK(0).Enabled = False
                cmdOK(1).Enabled = False
                cmdOK(2).Enabled = False
                cmdOK(3).Enabled = False
                Me.Enabled = True
                Screen.MousePointer = vbDefault
                 tmpBol = fnCancelNowFormAndShowParentForm(Me)
                Exit Sub
            End If
            'Modify By Sindy 2010/6/29
            IsMother = False
            IsChild = False
            If Val(CheckStr(adoRecordset.Fields(1).Value)) > 0 Then
                IsChild = True
            End If
            If Val(CheckStr(adoRecordset.Fields(3).Value)) > 0 Then
                IsMother = True
            End If
            If IsMother = False And IsChild = False Then
                s = MsgBox("資料庫中搜尋不到此案有分割!!", , "沒有資料")
                cmdOK(0).Enabled = False
                cmdOK(1).Enabled = False
                cmdOK(2).Enabled = False
                cmdOK(3).Enabled = False
                Me.Enabled = True
                Screen.MousePointer = vbDefault
                 tmpBol = fnCancelNowFormAndShowParentForm(Me)
                Exit Sub
            End If
        End If
        CheckOC
        strTemp = SystemNumber(Me.Tag, 1)
        'Modify By Sindy 2010/6/29
        strSql = ""
        If IsMother = True Then
            strSQL1 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL2 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL3 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL4 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL5 = " and dc05='" & SystemNumber(Me.Tag, 1) & "' and dc06='" & SystemNumber(Me.Tag, 2) & "' and dc07='" & SystemNumber(Me.Tag, 3) & "' and dc08='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL1 = strSQL1 & " and T1.PA01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 1) & ") "
            strSQL2 = strSQL2 & " and T1.TM01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 2) & ") "
            StrSQL3 = StrSQL3 & " and T1.LC01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 3) & ") "
            StrSQL4 = StrSQL4 & " and T1.HC01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 4) & ") "
            strSQL5 = strSQL5 & " and T1.SP01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 5) & ") "
            
            'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
            'strSql = "SELECT '' AS V,T1.PA01||'-'||T1.PA02||'-'||T1.PA03||'-'||T1.PA04||DECODE(length(nvl(t1.pa108,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.pa136,'')),null,'','●')||t1.pa47 as 母案分所號,NVL(T1.PA05,NVL(T1.PA06,T1.PA07)) AS 母案案件名稱,''  as 商品類別,T2.PA01||'-'||T2.PA02||'-'||T2.PA03||'-'||T2.PA04||DECODE(length(nvl(t2.pa108,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.pa136,'')),null,'','●')||t2.pa47 as 分割案分所號,NVL(T2.PA05,NVL(T2.PA06,T2.PA07)) AS 分割案案件名稱,''  as 商品類別 FROM PATENT T1,PATENT T2,DIVISIONcase " & _
                           "WHERE DC01=T2.PA01(+) AND DC02=T2.PA02(+) AND DC03=T2.PA03(+) AND DC04=T2.PA04(+) AND DC05=T1.PA01(+) AND DC06=T1.PA02(+) AND DC07=T1.PA03(+) AND DC08=T1.PA04(+) " & strSQL1
            'strSql = strSql & " union SELECT '' AS V,T1.TM01||'-'||T1.TM02||'-'||T1.TM03||'-'||T1.TM04||DECODE(length(nvl(t1.tm57,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.tm73,'')),null,'','●')||t1.tm34 as 母案分所號,NVL(T1.TM05,NVL(T1.TM06,T1.TM07)) AS 母案案件名稱,T1.tm09 as 商品類別,T2.TM01||'-'||T2.TM02||'-'||T2.TM03||'-'||T2.TM04||DECODE(length(nvl(t2.tm57,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.tm73,'')),null,'','●')||t2.tm34 as 分割案分所號,NVL(T2.TM05,NVL(T2.TM06,T2.TM07)) AS 分割案案件名稱,T2.tm09 as 商品類別 FROM TRADEMARK T1,TRADEMARK T2,DIVISIONcase " & _
                           "WHERE DC01=T2.TM01(+) AND DC02=T2.TM02(+) AND DC03=T2.TM03(+) AND DC04=T2.TM04(+) AND DC05=T1.TM01(+) AND DC06=T1.TM02(+) AND DC07=T1.TM03(+) AND DC08=T1.TM04(+) " & strSQL2
            'strSql = strSql & " union SELECT '' AS V,T1.LC01||'-'||T1.LC02||'-'||T1.LC03||'-'||T1.LC04||DECODE(length(nvl(t1.lc34,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.lc36,'')),null,'','●')||t1.lc16 as 母案分所號,NVL(T1.LC05,NVL(T1.LC06,T1.LC07)) AS 母案案件名稱,'' as 商品類別,T2.LC01||'-'||T2.LC02||'-'||T2.LC03||'-'||T2.LC04||DECODE(length(nvl(t2.lc34,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.lc36,'')),null,'','●')||t2.lc16 as 分割案分所號,NVL(T2.LC05,NVL(T2.LC06,T2.LC07)) AS 分割案案件名稱,''  as 商品類別 FROM LAWCASE T1,LAWCASE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.LC01(+) AND DC02=T2.LC02(+) AND DC03=T2.LC03(+) AND DC04=T2.LC04(+) AND DC05=T1.LC01(+) AND DC06=T1.LC02(+) AND DC07=T1.LC03(+) AND DC08=T1.LC04(+) " & StrSQL3
            'strSql = strSql & " union SELECT '' AS V,T1.HC01||'-'||T1.HC02||'-'||T1.HC03||'-'||T1.HC04||DECODE(length(nvl(t1.hc19,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.hc20,'')),null,'','●')||t1.hc07 as 母案分所號,T1.HC06 AS 母案案件名稱,'' as 商品類別,T2.HC01||'-'||T2.HC02||'-'||T2.HC03||'-'||T2.HC04||DECODE(length(nvl(t2.hc19,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.hc20,'')),null,'','●')||t2.hc07 as 分割案分所號,T2.HC06 AS 分割案案件名稱,''  as 商品類別 FROM HIRECASE T1,HIRECASE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.HC01(+) AND DC02=T2.HC02(+) AND DC03=T2.HC03(+) AND DC04=T2.HC04(+) AND DC05=T1.HC01(+) AND DC06=T1.HC02(+) AND DC07=T1.HC03(+) AND DC08=T1.HC04(+) " & StrSQL4
            'strSql = strSql & " union SELECT '' AS V,T1.SP01||'-'||T1.SP02||'-'||T1.SP03||'-'||T1.SP04||DECODE(length(nvl(t1.sp61,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.sp68,'')),null,'','●')||t1.sp28 as 母案分所號,NVL(T1.SP05,NVL(T1.SP06,T1.SP07)) AS 母案案件名稱,'' as 商品類別,T2.SP01||'-'||T2.SP02||'-'||T2.SP03||'-'||T2.SP04||DECODE(length(nvl(t2.sp61,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.sp68,'')),null,'','●')||t2.sp28 as 分割案分所號,NVL(T2.SP05,NVL(T2.SP06,T2.SP07)) AS 分割案案件名稱,''  as 商品類別 FROM SERVICEPRACTICE T1,SERVICEPRACTICE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.SP01(+) AND DC02=T2.SP02(+) AND DC03=T2.SP03(+) AND DC04=T2.SP04(+) AND DC05=T1.SP01(+) AND DC06=T1.SP02(+) AND DC07=T1.SP03(+) AND DC08=T1.SP04(+) " & strSQL5
            strSql = "SELECT '' AS V,T1.PA01||'-'||T1.PA02||'-'||T1.PA03||'-'||T1.PA04||DECODE(length(nvl(t1.pa108,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.pa136,'')),null,'','●')||t1.pa47 as 母案分所號,NVL(T1.PA05,NVL(T1.PA06,T1.PA07)) AS 母案案件名稱,''  as 商品類別," & _
                        "T2.PA01||'-'||T2.PA02||'-'||T2.PA03||'-'||T2.PA04||DECODE(length(nvl(t2.pa108,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.pa136,'')),null,'','●')||t2.pa47 as 分割案分所號,NVL(T2.PA05,NVL(T2.PA06,T2.PA07)) AS 分割案案件名稱,''  as 商品類別" & SeColPA & _
                        "FROM PATENT T1,PATENT T2,DIVISIONcase " & _
                        "WHERE DC01=T2.PA01(+) AND DC02=T2.PA02(+) AND DC03=T2.PA03(+) AND DC04=T2.PA04(+) AND DC05=T1.PA01(+) AND DC06=T1.PA02(+) AND DC07=T1.PA03(+) AND DC08=T1.PA04(+) " & strSQL1
            strSql = strSql & " union SELECT '' AS V,T1.TM01||'-'||T1.TM02||'-'||T1.TM03||'-'||T1.TM04||DECODE(length(nvl(t1.tm57,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.tm73,'')),null,'','●')||t1.tm34 as 母案分所號,NVL(T1.TM05,NVL(T1.TM06,T1.TM07)) AS 母案案件名稱,T1.tm09 as 商品類別," & _
                        "T2.TM01||'-'||T2.TM02||'-'||T2.TM03||'-'||T2.TM04||DECODE(length(nvl(t2.tm57,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.tm73,'')),null,'','●')||t2.tm34 as 分割案分所號,NVL(T2.TM05,NVL(T2.TM06,T2.TM07)) AS 分割案案件名稱,T2.tm09 as 商品類別" & SeColTM & _
                        "FROM TRADEMARK T1,TRADEMARK T2,DIVISIONcase " & _
                        "WHERE DC01=T2.TM01(+) AND DC02=T2.TM02(+) AND DC03=T2.TM03(+) AND DC04=T2.TM04(+) AND DC05=T1.TM01(+) AND DC06=T1.TM02(+) AND DC07=T1.TM03(+) AND DC08=T1.TM04(+) " & strSQL2
            strSql = strSql & " union SELECT '' AS V,T1.LC01||'-'||T1.LC02||'-'||T1.LC03||'-'||T1.LC04||DECODE(length(nvl(t1.lc34,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.lc36,'')),null,'','●')||t1.lc16 as 母案分所號,NVL(T1.LC05,NVL(T1.LC06,T1.LC07)) AS 母案案件名稱,'' as 商品類別," & _
                        "T2.LC01||'-'||T2.LC02||'-'||T2.LC03||'-'||T2.LC04||DECODE(length(nvl(t2.lc34,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.lc36,'')),null,'','●')||t2.lc16 as 分割案分所號,NVL(T2.LC05,NVL(T2.LC06,T2.LC07)) AS 分割案案件名稱,''  as 商品類別" & SeColLC & _
                        "FROM LAWCASE T1,LAWCASE T2,DIVISIONcase " & _
                         "WHERE DC01=T2.LC01(+) AND DC02=T2.LC02(+) AND DC03=T2.LC03(+) AND DC04=T2.LC04(+) AND DC05=T1.LC01(+) AND DC06=T1.LC02(+) AND DC07=T1.LC03(+) AND DC08=T1.LC04(+) " & StrSQL3
            strSql = strSql & " union SELECT '' AS V,T1.HC01||'-'||T1.HC02||'-'||T1.HC03||'-'||T1.HC04||DECODE(length(nvl(t1.hc19,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.hc20,'')),null,'','●')||t1.hc07 as 母案分所號,T1.HC06 AS 母案案件名稱,'' as 商品類別," & _
                        "T2.HC01||'-'||T2.HC02||'-'||T2.HC03||'-'||T2.HC04||DECODE(length(nvl(t2.hc19,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.hc20,'')),null,'','●')||t2.hc07 as 分割案分所號,T2.HC06 AS 分割案案件名稱,''  as 商品類別" & SeColHC & _
                        "FROM HIRECASE T1,HIRECASE T2,DIVISIONcase " & _
                        "WHERE DC01=T2.HC01(+) AND DC02=T2.HC02(+) AND DC03=T2.HC03(+) AND DC04=T2.HC04(+) AND DC05=T1.HC01(+) AND DC06=T1.HC02(+) AND DC07=T1.HC03(+) AND DC08=T1.HC04(+) " & StrSQL4
            strSql = strSql & " union SELECT '' AS V,T1.SP01||'-'||T1.SP02||'-'||T1.SP03||'-'||T1.SP04||DECODE(length(nvl(t1.sp61,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.sp68,'')),null,'','●')||t1.sp28 as 母案分所號,NVL(T1.SP05,NVL(T1.SP06,T1.SP07)) AS 母案案件名稱,'' as 商品類別," & _
                        "T2.SP01||'-'||T2.SP02||'-'||T2.SP03||'-'||T2.SP04||DECODE(length(nvl(t2.sp61,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.sp68,'')),null,'','●')||t2.sp28 as 分割案分所號,NVL(T2.SP05,NVL(T2.SP06,T2.SP07)) AS 分割案案件名稱,''  as 商品類別" & SeColSP & _
                        "FROM SERVICEPRACTICE T1,SERVICEPRACTICE T2,DIVISIONcase " & _
                         "WHERE DC01=T2.SP01(+) AND DC02=T2.SP02(+) AND DC03=T2.SP03(+) AND DC04=T2.SP04(+) AND DC05=T1.SP01(+) AND DC06=T1.SP02(+) AND DC07=T1.SP03(+) AND DC08=T1.SP04(+) " & strSQL5
            'end 2019/11/01
        End If
        If IsChild = True Then
            strSQL1 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL2 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL3 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            StrSQL4 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL5 = " and dc01='" & SystemNumber(Me.Tag, 1) & "' and dc02='" & SystemNumber(Me.Tag, 2) & "' and dc03='" & SystemNumber(Me.Tag, 3) & "' and dc04='" & SystemNumber(Me.Tag, 4) & "' "
            strSQL1 = strSQL1 & " and T2.PA01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 1) & ") "
            strSQL2 = strSQL2 & " and T2.TM01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 2) & ") "
            StrSQL3 = StrSQL3 & " and T2.LC01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 3) & ") "
            StrSQL4 = StrSQL4 & " and T2.HC01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 4) & ") "
            strSQL5 = strSQL5 & " and T2.SP01 in (" & SQLGrpStr(SystemNumber(Me.Tag, 1), 5) & ") "
            If strSql <> "" Then strSql = strSql & " union "
            'Modified by Lydia 2019/11/01 +增加欄位SeColPA, SeColTM, SeColSP, SeColLC, SeColHC
            'strSql = strSql & "SELECT '' AS V,T1.PA01||'-'||T1.PA02||'-'||T1.PA03||'-'||T1.PA04||DECODE(length(nvl(t1.pa108,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.pa136,'')),null,'','●')||t1.pa47 as 母案分所號,NVL(T1.PA05,NVL(T1.PA06,T1.PA07)) AS 母案案件名稱,''  as 商品類別,T2.PA01||'-'||T2.PA02||'-'||T2.PA03||'-'||T2.PA04||DECODE(length(nvl(t2.pa108,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.pa136,'')),null,'','●')||t2.pa47 as 分割案分所號,NVL(T2.PA05,NVL(T2.PA06,T2.PA07)) AS 分割案案件名稱,''  as 商品類別 FROM PATENT T1,PATENT T2,DIVISIONcase " & _
                           "WHERE DC01=T2.PA01(+) AND DC02=T2.PA02(+) AND DC03=T2.PA03(+) AND DC04=T2.PA04(+) AND DC05=T1.PA01(+) AND DC06=T1.PA02(+) AND DC07=T1.PA03(+) AND DC08=T1.PA04(+) " & strSQL1
            'strSql = strSql & " union SELECT '' AS V,T1.TM01||'-'||T1.TM02||'-'||T1.TM03||'-'||T1.TM04||DECODE(length(nvl(t1.tm57,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.tm73,'')),null,'','●')||t1.tm34 as 母案分所號,NVL(T1.TM05,NVL(T1.TM06,T1.TM07)) AS 母案案件名稱,T1.tm09 as 商品類別,T2.TM01||'-'||T2.TM02||'-'||T2.TM03||'-'||T2.TM04||DECODE(length(nvl(t2.tm57,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.tm73,'')),null,'','●')||t2.tm34 as 分割案分所號,NVL(T2.TM05,NVL(T2.TM06,T2.TM07)) AS 分割案案件名稱,T2.tm09 as 商品類別 FROM TRADEMARK T1,TRADEMARK T2,DIVISIONcase " & _
                           "WHERE DC01=T2.TM01(+) AND DC02=T2.TM02(+) AND DC03=T2.TM03(+) AND DC04=T2.TM04(+) AND DC05=T1.TM01(+) AND DC06=T1.TM02(+) AND DC07=T1.TM03(+) AND DC08=T1.TM04(+) " & strSQL2
            'strSql = strSql & " union SELECT '' AS V,T1.LC01||'-'||T1.LC02||'-'||T1.LC03||'-'||T1.LC04||DECODE(length(nvl(t1.lc34,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.lc36,'')),null,'','●')||t1.lc16 as 母案分所號,NVL(T1.LC05,NVL(T1.LC06,T1.LC07)) AS 母案案件名稱,'' as 商品類別,T2.LC01||'-'||T2.LC02||'-'||T2.LC03||'-'||T2.LC04||DECODE(length(nvl(t2.lc34,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.lc36,'')),null,'','●')||t2.lc16 as 分割案分所號,NVL(T2.LC05,NVL(T2.LC06,T2.LC07)) AS 分割案案件名稱,''  as 商品類別 FROM LAWCASE T1,LAWCASE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.LC01(+) AND DC02=T2.LC02(+) AND DC03=T2.LC03(+) AND DC04=T2.LC04(+) AND DC05=T1.LC01(+) AND DC06=T1.LC02(+) AND DC07=T1.LC03(+) AND DC08=T1.LC04(+) " & StrSQL3
            'strSql = strSql & " union SELECT '' AS V,T1.HC01||'-'||T1.HC02||'-'||T1.HC03||'-'||T1.HC04||DECODE(length(nvl(t1.hc19,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.hc20,'')),null,'','●')||t1.hc07 as 母案分所號,T1.HC06 AS 母案案件名稱,'' as 商品類別,T2.HC01||'-'||T2.HC02||'-'||T2.HC03||'-'||T2.HC04||DECODE(length(nvl(t2.hc19,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.hc20,'')),null,'','●')||t2.hc07 as 分割案分所號,T2.HC06 AS 分割案案件名稱,''  as 商品類別 FROM HIRECASE T1,HIRECASE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.HC01(+) AND DC02=T2.HC02(+) AND DC03=T2.HC03(+) AND DC04=T2.HC04(+) AND DC05=T1.HC01(+) AND DC06=T1.HC02(+) AND DC07=T1.HC03(+) AND DC08=T1.HC04(+) " & StrSQL4
            'strSql = strSql & " union SELECT '' AS V,T1.SP01||'-'||T1.SP02||'-'||T1.SP03||'-'||T1.SP04||DECODE(length(nvl(t1.sp61,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.sp68,'')),null,'','●')||t1.sp28 as 母案分所號,NVL(T1.SP05,NVL(T1.SP06,T1.SP07)) AS 母案案件名稱,'' as 商品類別,T2.SP01||'-'||T2.SP02||'-'||T2.SP03||'-'||T2.SP04||DECODE(length(nvl(t2.sp61,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.sp68,'')),null,'','●')||t2.sp28 as 分割案分所號,NVL(T2.SP05,NVL(T2.SP06,T2.SP07)) AS 分割案案件名稱,''  as 商品類別 FROM SERVICEPRACTICE T1,SERVICEPRACTICE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.SP01(+) AND DC02=T2.SP02(+) AND DC03=T2.SP03(+) AND DC04=T2.SP04(+) AND DC05=T1.SP01(+) AND DC06=T1.SP02(+) AND DC07=T1.SP03(+) AND DC08=T1.SP04(+) " & strSQL5
            strSql = strSql & "SELECT '' AS V,T1.PA01||'-'||T1.PA02||'-'||T1.PA03||'-'||T1.PA04||DECODE(length(nvl(t1.pa108,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.pa136,'')),null,'','●')||t1.pa47 as 母案分所號,NVL(T1.PA05,NVL(T1.PA06,T1.PA07)) AS 母案案件名稱,''  as 商品類別," & _
                            "T2.PA01||'-'||T2.PA02||'-'||T2.PA03||'-'||T2.PA04||DECODE(length(nvl(t2.pa108,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.pa136,'')),null,'','●')||t2.pa47 as 分割案分所號,NVL(T2.PA05,NVL(T2.PA06,T2.PA07)) AS 分割案案件名稱,''  as 商品類別" & SeColPA & _
                            "FROM PATENT T1,PATENT T2,DIVISIONcase " & _
                            "WHERE DC01=T2.PA01(+) AND DC02=T2.PA02(+) AND DC03=T2.PA03(+) AND DC04=T2.PA04(+) AND DC05=T1.PA01(+) AND DC06=T1.PA02(+) AND DC07=T1.PA03(+) AND DC08=T1.PA04(+) " & strSQL1
            strSql = strSql & " union SELECT '' AS V,T1.TM01||'-'||T1.TM02||'-'||T1.TM03||'-'||T1.TM04||DECODE(length(nvl(t1.tm57,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.tm73,'')),null,'','●')||t1.tm34 as 母案分所號,NVL(T1.TM05,NVL(T1.TM06,T1.TM07)) AS 母案案件名稱,T1.tm09 as 商品類別," & _
                           "T2.TM01||'-'||T2.TM02||'-'||T2.TM03||'-'||T2.TM04||DECODE(length(nvl(t2.tm57,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.tm73,'')),null,'','●')||t2.tm34 as 分割案分所號,NVL(T2.TM05,NVL(T2.TM06,T2.TM07)) AS 分割案案件名稱,T2.tm09 as 商品類別" & SeColTM & _
                           "FROM TRADEMARK T1,TRADEMARK T2,DIVISIONcase " & _
                           "WHERE DC01=T2.TM01(+) AND DC02=T2.TM02(+) AND DC03=T2.TM03(+) AND DC04=T2.TM04(+) AND DC05=T1.TM01(+) AND DC06=T1.TM02(+) AND DC07=T1.TM03(+) AND DC08=T1.TM04(+) " & strSQL2
            strSql = strSql & " union SELECT '' AS V,T1.LC01||'-'||T1.LC02||'-'||T1.LC03||'-'||T1.LC04||DECODE(length(nvl(t1.lc34,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.lc36,'')),null,'','●')||t1.lc16 as 母案分所號,NVL(T1.LC05,NVL(T1.LC06,T1.LC07)) AS 母案案件名稱,'' as 商品類別," & _
                           "T2.LC01||'-'||T2.LC02||'-'||T2.LC03||'-'||T2.LC04||DECODE(length(nvl(t2.lc34,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.lc36,'')),null,'','●')||t2.lc16 as 分割案分所號,NVL(T2.LC05,NVL(T2.LC06,T2.LC07)) AS 分割案案件名稱,''  as 商品類別" & SeColLC & _
                           "FROM LAWCASE T1,LAWCASE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.LC01(+) AND DC02=T2.LC02(+) AND DC03=T2.LC03(+) AND DC04=T2.LC04(+) AND DC05=T1.LC01(+) AND DC06=T1.LC02(+) AND DC07=T1.LC03(+) AND DC08=T1.LC04(+) " & StrSQL3
            strSql = strSql & " union SELECT '' AS V,T1.HC01||'-'||T1.HC02||'-'||T1.HC03||'-'||T1.HC04||DECODE(length(nvl(t1.hc19,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.hc20,'')),null,'','●')||t1.hc07 as 母案分所號,T1.HC06 AS 母案案件名稱,'' as 商品類別," & _
                           "T2.HC01||'-'||T2.HC02||'-'||T2.HC03||'-'||T2.HC04||DECODE(length(nvl(t2.hc19,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.hc20,'')),null,'','●')||t2.hc07 as 分割案分所號,T2.HC06 AS 分割案案件名稱,''  as 商品類別" & SeColHC & _
                           "FROM HIRECASE T1,HIRECASE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.HC01(+) AND DC02=T2.HC02(+) AND DC03=T2.HC03(+) AND DC04=T2.HC04(+) AND DC05=T1.HC01(+) AND DC06=T1.HC02(+) AND DC07=T1.HC03(+) AND DC08=T1.HC04(+) " & StrSQL4
            strSql = strSql & " union SELECT '' AS V,T1.SP01||'-'||T1.SP02||'-'||T1.SP03||'-'||T1.SP04||DECODE(length(nvl(t1.sp61,'')),null,'','●') AS 母案本所案號,DECODE(length(nvl(t1.sp68,'')),null,'','●')||t1.sp28 as 母案分所號,NVL(T1.SP05,NVL(T1.SP06,T1.SP07)) AS 母案案件名稱,'' as 商品類別," & _
                           "T2.SP01||'-'||T2.SP02||'-'||T2.SP03||'-'||T2.SP04||DECODE(length(nvl(t2.sp61,'')),null,'','●') AS 分割案本所案號,DECODE(length(nvl(t2.sp68,'')),null,'','●')||t2.sp28 as 分割案分所號,NVL(T2.SP05,NVL(T2.SP06,T2.SP07)) AS 分割案案件名稱,''  as 商品類別" & SeColSP & _
                           "FROM SERVICEPRACTICE T1,SERVICEPRACTICE T2,DIVISIONcase " & _
                           "WHERE DC01=T2.SP01(+) AND DC02=T2.SP02(+) AND DC03=T2.SP03(+) AND DC04=T2.SP04(+) AND DC05=T1.SP01(+) AND DC06=T1.SP02(+) AND DC07=T1.SP03(+) AND DC08=T1.SP04(+) " & strSQL5
            'end 2019/11/01
        End If
      adoRecordset.CursorLocation = adUseClient
      'Modified by Lydia 2019/11/01 改變型態
      'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

      If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
            If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                adoRecordset.MoveFirst
                Do While adoRecordset.EOF = False
                    If "" & adoRecordset.Fields("casenoA") <> "" Then
                        If strDELlist <> "" And InStr(strDELlist, "" & adoRecordset.Fields("casenoA")) > 0 Then '已檢查過並且為限閱案件
                            adoRecordset.Delete
                            GoTo JumpToNext
                        Else
                            If InStr(StrCaseList, "" & adoRecordset.Fields("casenoA")) = 0 Then
                                StrCaseList = StrCaseList & adoRecordset.Fields("casenoA") & ","   '記錄已檢查過的案號
                                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("casenoA"), "" & adoRecordset.Fields("custA1") & "," & adoRecordset.Fields("custA2") & "," & adoRecordset.Fields("custA3") & "," & adoRecordset.Fields("custA4") & "," & adoRecordset.Fields("custA5"), "" & adoRecordset.Fields("fcnoA")) = False Then
                                    strDELlist = strDELlist & adoRecordset.Fields("casenoA") '記錄已檢查過並且為限閱案件的案號
                                    intCufaCnt = intCufaCnt + 1
                                    adoRecordset.Delete
                                    GoTo JumpToNext
                                End If
                            End If
                        End If
                    End If
                    
                    If "" & adoRecordset.Fields("casenoB") <> "" Then
                        If strDELlist <> "" And InStr(strDELlist, "" & adoRecordset.Fields("casenoB")) > 0 Then '已檢查過並且為限閱案件
                            adoRecordset.Delete
                            GoTo JumpToNext
                        Else
                            If InStr(StrCaseList, "" & adoRecordset.Fields("casenoB")) = 0 Then
                                StrCaseList = StrCaseList & adoRecordset.Fields("casenoB") & ","   '記錄已檢查過的案號
                                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("casenoB"), "" & adoRecordset.Fields("custB1") & "," & adoRecordset.Fields("custB2") & "," & adoRecordset.Fields("custB3") & "," & adoRecordset.Fields("custB4") & "," & adoRecordset.Fields("custB5"), "" & adoRecordset.Fields("fcnoB")) = False Then
                                    strDELlist = strDELlist & adoRecordset.Fields("casenoB") '記錄已檢查過並且為限閱案件的案號
                                    intCufaCnt = intCufaCnt + 1
                                    adoRecordset.Delete
                                    GoTo JumpToNext
                                End If
                            End If
                        End If
                    End If
JumpToNext:
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
            'end 2019/11/01
         'Modified by Lydia 2023/02/13
         'cmdOK(0).Enabled = False
         'cmdOK(1).Enabled = False
         'cmdOK(2).Enabled = False
         'cmdOK(3).Enabled = False
         cmdOK(0).Enabled = True
         cmdOK(1).Enabled = True
         cmdOK(2).Enabled = True
         cmdOK(3).Enabled = True
         'end 2023/02/13
         Set grdDataList.Recordset = adoRecordset
         CheckOC
      Else
JumpToNoData:   'Added by Lydia 2019/11/01
         s = MsgBox("此本所案號沒有分割案, 無法查詢!!" & Me.Tag & "  ", , "錯誤")
         cmdOK(0).Enabled = False
         cmdOK(1).Enabled = False
         cmdOK(2).Enabled = False
         cmdOK(3).Enabled = False
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      End If
      Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100108_4 = Nothing
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

'Added by Lydia 2019/11/01  利益衝突案件：預設
Private Sub SetCUFA(ByVal iKind As String)

    '於後面增加欄位
    SeColTM = " ,t1.tm01||'-'||t1.tm02||'-'||t1.tm03||'-'||t1.tm04 as casenoA,t1.tm23 as custA1,t1.tm78 as custA2,t1.tm79 as custA3,t1.tm80 as custA4,t1.tm81 as custA5,t1.tm44 as fcnoA " & _
                        " ,t2.tm01||'-'||t2.tm02||'-'||t2.tm03||'-'||t2.tm04 as casenoB,t2.tm23 as custB1,t2.tm78 as custB2,t2.tm79 as custB3,t2.tm80 as custB4,t2.tm81 as custB5,t2.tm44 as fcnoB "
    SeColPA = " ,t1.pa01||'-'||t1.pa02||'-'||t1.pa03||'-'||t1.pa04 as casenoA,t1.pa26 as custA1,t1.pa27 as custA2,t1.pa28 as custA3,t1.pa29 as custA4,t1.pa30 as custA5,t1.pa75 as fcnoA " & _
                       " ,t2.pa01||'-'||t2.pa02||'-'||t2.pa03||'-'||t2.pa04 as casenoB,t2.pa26 as custB1,t2.pa27 as custB2,t2.pa28 as custB3,t2.pa29 as custB4,t2.pa30 as custB5,t2.pa75 as fcnoB "
    SeColSP = " ,t1.sp01||'-'||t1.sp02||'-'||t1.sp03||'-'||t1.sp04 as casenoA,t1.sp08 as custA1,t1.sp58 as custA2,t1.sp59 as custA3,t1.sp65 as custA4,t1.sp66 as custA5,t1.sp26 as fcnoA " & _
                    " ,t2.sp01||'-'||t2.sp02||'-'||t2.sp03||'-'||t2.sp04 as casenoB,t2.sp08 as custB1,t2.sp58 as custB2,t2.sp59 as custB3,t2.sp65 as custB4,t2.sp66 as custB5,t2.sp26 as fcnoB "
    SeColLC = " ,t1.lc01||'-'||t1.lc02||'-'||t1.lc03||'-'||t1.lc04 as casenoA,t1.lc11 as custA1,t1.lc43 as custA2,t1.lc44 as custA3,t1.lc45 as custA4,t1.lc46 as custA5,t1.lc22 as fcnoA " & _
                     " ,t2.lc01||'-'||t2.lc02||'-'||t2.lc03||'-'||t2.lc04 as casenoB,t2.lc11 as custB1,t2.lc43 as custB2,t2.lc44 as custB3,t2.lc45 as custB4,t2.lc46 as custB5,t2.lc22 as fcnoB "
    SeColHC = " ,t1.hc01||'-'||t1.hc02||'-'||t1.hc03||'-'||t1.hc04 as casenoA,t1.hc05 as custA1,t1.hc24 as custA2,t1.hc25 as custA3,t1.hc26 as custA4,t1.hc27 as custA5,'' as fcnoA " & _
                    " ,t2.hc01||'-'||t2.hc02||'-'||t2.hc03||'-'||t2.hc04 as casenoB,t2.hc05 as custB1,t2.hc24 as custB2,t2.hc25 as custB3,t2.hc26 as custB4,t2.hc27 as custB5,'' as fcnoB "
    
    If iKind = "1" Then
        m_AllSys = IIf(frm100108_1.txt1(6).Text <> "ALL", frm100108_1.txt1(6).Text, GetAllSysKind(frm100108_1.txt1(6)))
    Else
        m_AllSys = lbl1(2).Caption
    End If
    intCufaCnt = 0
    StrCaseList = ""
    strDELlist = ""
End Sub

