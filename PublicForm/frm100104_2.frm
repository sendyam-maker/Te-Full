VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100104_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以收／發文日查詢"
   ClientHeight    =   5712
   ClientLeft      =   156
   ClientTop       =   996
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8508
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   0
      Left            =   4536
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   6060
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7284
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5016
      Left            =   0
      TabIndex        =   4
      Top             =   684
      Width           =   9288
      _ExtentX        =   16383
      _ExtentY        =   8827
      _Version        =   393216
      Cols            =   19
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
      _Band(0).Cols   =   19
   End
   Begin VB.Label lblTot 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "共 0 筆"
      Height          =   180
      Left            =   8670
      TabIndex        =   10
      Top             =   450
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   2490
      TabIndex        =   9
      Top             =   180
      Width           =   1980
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   1
      Left            =   4992
      TabIndex        =   8
      Top             =   444
      Width           =   2640
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   0
      Left            =   1056
      TabIndex        =   7
      Top             =   444
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   3930
      TabIndex        =   6
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文期間："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm100104_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/18 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, i As Integer, j As Integer, strTemp As Variant, strTemp1 As Variant, s As Integer
Dim StrTag As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add By Amy 2016/06/22
Dim strFieldN()
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
Dim iCol As Integer

'Modify by Morgan 2007/7/23
'grdDataList.Cols = 20
'Modify by Amy 2016/06/22 原Cols=21:
'Modified by Lydia 2019/05/16 +2
'grdDataList.Cols = 25
'Modified by Lydia 2019/11/01
'grdDataList.Cols = 27
grdDataList.Cols = 32
ReDim strFieldN(1 To grdDataList.Cols - 1)

iCol = 0
grdDataList.row = 0
grdDataList.col = iCol: grdDataList.Text = "V"
grdDataList.ColWidth(iCol) = 200
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "收文日": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "本所案號": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = iCol: grdDataList.Text = "分所號": strFieldN(iCol) = grdDataList.Text
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(iCol) = 0
Else
    grdDataList.ColWidth(iCol) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "案件名稱": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 1300
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "案件性質": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 800
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
'Add By nick 2004/08/23
grdDataList.col = iCol: grdDataList.Text = "發文規費": strFieldN(iCol) = grdDataList.Text
If frm100104_1.txt1(26).Text = "Y" Then
    grdDataList.ColWidth(iCol) = 1000
Else
    grdDataList.ColWidth(iCol) = 0
End If
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/05/16 顯示發文規費欄, 請再其後增加發文時間,扣款日
iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "發文時間": strFieldN(iCol) = grdDataList.Text
If frm100104_1.txt1(26).Text = "Y" Then
    grdDataList.ColWidth(iCol) = 860
Else
    grdDataList.ColWidth(iCol) = 0
End If
grdDataList.CellAlignment = flexAlignCenterCenter
iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "扣款日": strFieldN(iCol) = grdDataList.Text
If frm100104_1.txt1(26).Text = "Y" Then
    grdDataList.ColWidth(iCol) = 860
Else
    grdDataList.ColWidth(iCol) = 0
End If
grdDataList.CellAlignment = flexAlignCenterCenter
'end 2019/05/16

'Add By Morgan 2007/7/20
iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "工作時數": strFieldN(iCol) = grdDataList.Text
If frm100104_1.txt1(28).Text = "Y" Then
    grdDataList.ColWidth(iCol) = 1000
Else
    grdDataList.ColWidth(iCol) = 0
End If
grdDataList.CellAlignment = flexAlignCenterCenter

'end 2007/7/20
iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "承辦人": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 700
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "智權人員": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 700
grdDataList.CellAlignment = flexAlignCenterCenter

'Add By Sindy 2023/4/14
iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "簽": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 300
grdDataList.CellAlignment = flexAlignCenterCenter
'2023/4/14 END

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "本所期限": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
'edit by nickc 2007/03/23 更換 PCT
'Modify By Sindy 2012/3/7
'If frm100104_1.ChkPCT.Value = vbChecked Then
If frm100104_1.txt1(35) = "Y" Then
    grdDataList.col = iCol: grdDataList.Text = "PCT": strFieldN(iCol) = grdDataList.Text
    grdDataList.ColWidth(iCol) = 620
Else
    grdDataList.col = iCol: grdDataList.Text = "法定期限": strFieldN(iCol) = grdDataList.Text
    grdDataList.ColWidth(iCol) = 810
End If
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "發文日": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "申請人": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 600
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "點數": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 500
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "是否出名": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 800
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "取消收文日": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "申請國家": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 800
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "總收文號": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "收款情形": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
'End

iCol = iCol + 1
'add by nickc 2005/05/10
grdDataList.col = iCol: grdDataList.Text = "": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Add by Amy 2016/06/22
iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "CP13": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "CP10": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "Nation": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "Apply": strFieldN(iCol) = grdDataList.Text
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人2~5, FC代理人
If iCol <= UBound(strFieldN) Then
   For i = iCol To UBound(strFieldN)
       grdDataList.col = i
       grdDataList.ColWidth(i) = 0
   Next i
End If

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
        grdDataList.col = 2
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
                Case "CFL", "FCL", "L", "LIN", "ACS" '法務
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
        End If
         Me.Enabled = True
         Exit Sub
     End If
     Next i
     Me.Enabled = True
Case 1
     Me.Enabled = False
     StrTag = ""
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
         grdDataList.col = 2
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
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
Case Else
End Select
End Sub


Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Select Case Index
Case 0
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        Dim Str01 As String
        grdDataList.col = 2
        Str01 = SystemNumber(grdDataList, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList.Text) Then
        Select Case Pub_RplStr(Str01)
            Case "CFP", "FCP", "P"   '專利
                  Screen.MousePointer = vbHourglass
                  frm100101_3.Show
                  'frm100101_3.Hide
                  'Modify By Cheng 2002/06/25
'                  frm100101_3.Tag = grdDataList.Text
                  frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_3.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_3.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_3.Visible
                  Unload frm100101_3
            Case "CFT", "FCT", "T", "TF"   '商標
                  Screen.MousePointer = vbHourglass
                  frm100101_4.Show
                  'frm100101_4.Hide
                  'Modify By Cheng 2002/06/25
'                  frm100101_4.Tag = grdDataList.Text
                  frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_4.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_4.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_4.Visible
                  Unload frm100101_4
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"  '法務
                  Screen.MousePointer = vbHourglass
                  frm100101_5.Show
                  'frm100101_5.Hide
                  'Modify By Cheng 2002/06/25
'                  frm100101_5.Tag = grdDataList.Text
                  frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_5.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_5.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_5.Visible
                  Unload frm100101_5
            Case "LA"            '顧問
                  Screen.MousePointer = vbHourglass
                  frm100101_6.Show
                  'frm100101_6.Hide
                  'Modify By Cheng 2002/06/25
'                  frm100101_6.Tag = grdDataList.Text
                  frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_6.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Hide
                  'frm100101_6.Show
                  Do
                  DoEvents
                  If bolToEndByNick = True Then Unload Me: Exit Sub
                  Loop Until Not frm100101_6.Visible
                  Unload frm100101_6
            Case Else                  '服務
                 Select Case Pub_RplStr(Str01)
                     Case "TB"    '條碼
                         Screen.MousePointer = vbHourglass
                        frm100101_7.Show
                        'frm100101_7.Hide
                        'Modify By Cheng 2002/06/25
'                        frm100101_7.Tag = grdDataList.Text
                        frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_7.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_7.Show
                        Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_7.Visible
                         Unload frm100101_7
                     Case "TM"
                         Screen.MousePointer = vbHourglass
                        frm100101_8.Show
                        'frm100101_8.Hide
                        'Modify By Cheng 2002/06/25
'                        frm100101_8.Tag = grdDataList.Text
                        frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_8.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_8.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_8.Visible
                         Unload frm100101_8
                     Case "TD"
                         Screen.MousePointer = vbHourglass
                        frm100101_9.Show
                        'frm100101_9.Hide
                        'Modify By Cheng 2002/06/25
'                        frm100101_9.Tag = grdDataList.Text
                        frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_9.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_9.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_9.Visible
                         Unload frm100101_9
                     Case "TC", "CFC"
                         Screen.MousePointer = vbHourglass
                        frm100101_A.Show
                        'frm100101_A.Hide
                        'Modify By Cheng 2002/06/25
'                        frm100101_A.Tag = grdDataList.Text
                        frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_A.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_A.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_A.Visible
                         Unload frm100101_A
                     Case Else
                         Screen.MousePointer = vbHourglass
                        frm100101_B.Show
                        'frm100101_B.Hide
                        'Modify By Cheng 2002/06/25
'                        frm100101_B.Tag = grdDataList.Text
                        frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                        frm100101_B.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Hide
                        'frm100101_B.Show
                         Do
                         DoEvents
                         If bolToEndByNick = True Then Unload Me: Exit Sub
                         Loop Until Not frm100101_B.Visible
                         Unload frm100101_B
                  End Select
        End Select
        End If
        grdDataList.col = 0
        grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
             grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
     End If
     Next i
     Me.Enabled = True
     Me.Show
Case 1
     Me.Enabled = False
     StrTag = ""
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 2
         If Not IsNull(grdDataList.Text) Then
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            'frm100101_2.Hide
            'Modify By Cheng 2002/06/25
'            frm100101_2.Tag = grdDataList.Text ' StrTag
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Hide
            'frm100101_2.Show
            Do
            DoEvents
            If bolToEndByNick = True Then Unload Me: Exit Sub
            Loop Until Not frm100101_2.Visible
            Unload frm100101_2
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
               grdDataList.col = j
               grdDataList.CellBackColor = QBColor(15)
            Next j

         End If
     End If
     Next i
     Me.Enabled = True
     Me.Show
Case 2
     Me.Hide
Case 3
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth
'92.04.16 nick
cmdState = -1
m_blnColOrderAsc = True 'Add by Amy 2016/06/22
End Sub

Sub StrMenu()
'Add By Cheng 2002/01/22
Dim strSQL11 As String
Dim strSQL21 As String
Dim strSQL31 As String
Dim strSQL41 As String
Dim strSQL51 As String
'Add By Cheng 2002/10/22
Dim StrSQLa As String
'Add By Cheng 2002/12/04
Dim strSQL1_1 As String
Dim strSQL1_2 As String
Dim strSQL1_3 As String
Dim strSQL1_4 As String
Dim strSQL1_5 As String
'Add By Cheng 2003/04/23
Dim strTKind As String
Dim arrTKind
Dim ii As Double
'Add by Amy 2017/02/24
Dim strMCTF As String, strCU13 As String
'Added by Lydia 2019/11/01 利益衝突案件：於Apply後面，增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim strTmp As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
'Added by Lydia 2019/11/01 利益衝突案件：於Apply後面，增加欄位
SeColTM = " ,tm78 as apply02,tm79 as apply03,tm80 as apply04,tm81 as apply05,tm44 as fcno "
SeColPA = " ,pa27 as apply02,pa28 as apply03,pa29 as apply04,pa30 as apply05,pa75 as fcno "
SeColSP = " ,sp58 as apply02,sp59 as apply03,sp65 as apply04,sp66 as apply05,sp26 as fcno "
SeColLC = " ,lc43 as apply02,lc44 as apply03,lc45 as apply04,lc46 as apply05,lc22 as fcno "
SeColHC = " ,hc24 as apply02,hc25 as apply03,hc26 as apply04,hc27 as apply05,'' as fcno "
intCufaCnt = 0
m_AllSys = IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, GetAllSysKind(, frm100104_1.txt1(3).Text))
'end 2019/11/01

'Add By Cheng 2002/01/22
strSQL11 = "": strSQL21 = "": strSQL31 = "": strSQL41 = "": strSQL51 = ""
'Add By Cheng 2002/12/04
strSQL1_1 = "": strSQL1_2 = "": strSQL1_3 = "": strSQL1_4 = "": strSQL1_5 = ""
If frm100104_1.txt1(9) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label8, 4) & "查詢" 'Add By Sindy 2010/01/22
ElseIf frm100104_1.txt1(9) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label8, 4) & "印表" 'Add By Sindy 2010/01/22
End If
If frm100104_1.txt1(0) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label1(0), 4) & "收文 " 'Add By Sindy 2010/01/22
ElseIf frm100104_1.txt1(0) = "2" Then
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label1(0), 4) & "發文" 'Add By Sindy 2010/01/22
End If
If frm100104_1.txt1(0) = "1" Then
   If Len(Trim(frm100104_1.txt1(1))) <> 0 Then
      strSQL1 = strSQL1 + " and CP05>=" & ChangeTStringToWString(frm100104_1.txt1(1)) & " "
        'Add By Cheng 2002/12/04
      strSQL1_1 = strSQL1_1 + " and CP05>=" & ChangeTStringToWString(frm100104_1.txt1(1)) & " "
   End If
   If Len(Trim(frm100104_1.txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & ChangeTStringToWString(frm100104_1.txt1(2)) & " "
        'Add By Cheng 2002/12/04
      strSQL1_1 = strSQL1_1 & " AND CP05<=" & ChangeTStringToWString(frm100104_1.txt1(2)) & " "
   End If
   
   If frm100104_1.ChkCP159.Value = 0 Then 'Added by Lydia 2022/05/17 原本預設「收文量不含已取消收文案件」，現在改為勾選項「□是否含已取消收文案件」
      'Added by Lydia 2016/09/06  +判斷未取消收文 CP159=0
      strSQL1 = strSQL1 & " AND CP159=0 "
      strSQL1_1 = strSQL1_1 & " AND CP159=0 "
   End If 'Added by Lydia 2022/05/17
   
   If Len(Trim(frm100104_1.txt1(1))) <> 0 Or Len(Trim(frm100104_1.txt1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & "收文" & frm100104_1.Label2 & frm100104_1.txt1(1) & "-" & frm100104_1.txt1(2) 'Add By Sindy 2010/01/22
   End If
Else
   If Len(Trim(frm100104_1.txt1(1))) <> 0 Then
      strSQL1 = strSQL1 + " and CP27>=" & ChangeTStringToWString(frm100104_1.txt1(1)) & " "
        'Add By Cheng 2002/12/04
      strSQL1_1 = strSQL1_1 + " and CP27>=" & ChangeTStringToWString(frm100104_1.txt1(1)) & " "
   End If
   If Len(Trim(frm100104_1.txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP27<=" & ChangeTStringToWString(frm100104_1.txt1(2)) & " "
        'Add By Cheng 2002/12/04
      strSQL1_1 = strSQL1_1 & " AND CP27<=" & ChangeTStringToWString(frm100104_1.txt1(2)) & " "
   End If
   If Len(Trim(frm100104_1.txt1(1))) <> 0 Or Len(Trim(frm100104_1.txt1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & "發文" & frm100104_1.Label2 & frm100104_1.txt1(1) & "-" & frm100104_1.txt1(2) 'Add By Sindy 2010/01/22
   End If
End If
lbl1(1) = frm100104_1.txt1(4) + "－" + frm100104_1.txt1(5)
If Len(Trim(frm100104_1.txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & frm100104_1.txt1(4) & "' "
     'Add By Cheng 2002/12/04
    strSQL1_1 = strSQL1_1 + " AND CP10>='" & frm100104_1.txt1(4) & "' "
End If
If Len(Trim(frm100104_1.txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & frm100104_1.txt1(5) & "' "
    'Add By Cheng 2002/12/04
    strSQL1_1 = strSQL1_1 + " AND CP10<='" & frm100104_1.txt1(5) & "' "
End If
If Len(Trim(frm100104_1.txt1(4))) <> 0 Or Len(Trim(frm100104_1.txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label4 & frm100104_1.txt1(4) & "-" & frm100104_1.txt1(5) 'Add By Sindy 2010/01/22
End If
If Len(Trim(frm100104_1.txt1(7))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & frm100104_1.txt1(7) & "' "
    'Add By Cheng 2002/12/04
    strSQL1_1 = strSQL1_1 + " AND CP14='" & frm100104_1.txt1(7) & "' "
    pub_QL05 = pub_QL05 & ";" & frm100104_1.Label6 & frm100104_1.txt1(7) & frm100104_1.lbl1(0)  'Add By Sindy 2010/01/22
End If

'Add by Morgan 2003/12/18
If Len(Trim(frm100104_1.txt1(24))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP12>='" & frm100104_1.txt1(24) & "' "
    strSQL1_1 = strSQL1_1 + " AND CP12>='" & frm100104_1.txt1(24) & "' "
End If
If Len(Trim(frm100104_1.txt1(25))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP12<='" & frm100104_1.txt1(25) & "' "
    strSQL1_1 = strSQL1_1 + " AND CP12<='" & frm100104_1.txt1(25) & "' "
End If
If Len(Trim(frm100104_1.txt1(24))) <> 0 Or Len(Trim(frm100104_1.txt1(25))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label18 & frm100104_1.txt1(24) & "-" & frm100104_1.txt1(25) 'Add By Sindy 2010/01/22
End If
'Add end 2003/12/18

'Add By Sindy 2023/4/14
If Len(Trim(frm100104_1.txt1(38))) <> 0 Then
   If frm100104_1.txt1(38) = "1" Then
      strExc(9) = "Y" '一般簽核
      pub_QL05 = pub_QL05 & ";主管簽核：一般簽核"
   Else
      strExc(9) = "特" '特例簽核
      pub_QL05 = pub_QL05 & ";主管簽核：特例簽核"
   End If
   strSQL1 = strSQL1 + " AND instr(GetSignF0202TypeNm(cp140),'" & strExc(9) & "')>0 AND CP140 is not null"
   strSQL1_1 = strSQL1_1 + " AND instr(GetSignF0202TypeNm(cp140),'" & strExc(9) & "')>0 AND CP140 is not null"
End If
'2023/4/14 END

If Len(Trim(frm100104_1.txt1(8))) <> 0 Then
    'Modify by Amy 2017/02/24 智權人員加MCTF
    If Left(frm100104_1.txt1(8), 4) = "MCTM" Or Left(frm100104_1.txt1(8), 4) = "MCTF" Then
        '下MCTM則三組分組都抓
        'Modify by Amy 2019/07/19 多增加MCTF04/05 且可能有離職人員無法記錄到,故增加 MCTMember
        If Left(frm100104_1.txt1(8), 5) = "MCTM" Then
'            strMCTF = ",'" & Replace(Pub_GetSpecMan("MCTF0", True), ";", "','") & "','MCTF01','MCTF02','MCTF03' "
'            strCU13 = " And SubStr(F1.fa120,1,5)='MCTF0' "
            strMCTF = ",'" & Replace(Pub_GetSpecMan("MCTMember", False), ";", "','") & "' "
            strCU13 = " And SubStr(cp161,1,4)='MCTF' "
        Else
            strMCTF = ",'" & Replace(Pub_GetSpecMan(frm100104_1.txt1(8)), ";", "','") & "','" & frm100104_1.txt1(8) & "' "
            'strCU13 = " And F1.fa120='" & frm100104_1.txt1(8) & "' "
            strCU13 = " And cp161='" & frm100104_1.txt1(8) & "' "
        End If
'        strSQL1 = strSQL1 + " AND CP13 in (" & Mid(strMCTF, 2) & ") "
'        strSQL1_1 = strSQL1_1 + " AND CP13 in (" & Mid(strMCTF, 2) & ") "
        strSQL1 = strSQL1 + " AND CP161 in (" & Mid(strMCTF, 2) & ") "
        strSQL1_1 = strSQL1_1 + " AND CP161 in (" & Mid(strMCTF, 2) & ") "
        'end 2019/07/19
        strSQL21 = strSQL21 + " And TM44 is not null" & strCU13
        strSQL31 = strSQL31 + " And LC22 is not null" & strCU13 'Add by Amy 2017/06/06 法務也加 ex:1060508 86048收文之L-005714不出現
        'modify by sonia 2024/4/30 查MCTM才會顯示法律所案源收文
        'strSQL51 = strSQL51 + " And SP26 is not null" & strCU13
        strSQL51 = strSQL51 + " And (SP26 is not null or (sp26 is null and sp01||sp02='TT999999'))" & strCU13
    Else
        strSQL1 = strSQL1 + " AND CP13='" & frm100104_1.txt1(8) & "' "
        'Add By Cheng 2002/12/04
        strSQL1_1 = strSQL1_1 + " AND CP13='" & frm100104_1.txt1(8) & "' "
    End If
    'end 2017/02/24
    pub_QL05 = pub_QL05 & ";" & frm100104_1.Label7 & frm100104_1.txt1(8) & frm100104_1.lbl1(1) 'Add By Sindy 2010/01/22
End If
'Modify By Cheng 2002/03/05
'If Len(Trim(frm100104_1.txt1(10))) = 0 Then
'    strSQL1 = strSQL1 + " AND CP09 < 'C' "
'End If

'若含內部收文資料, 但不含來函資料時
If Len(Trim(frm100104_1.txt1(15).Text)) > 0 And Len(Trim(frm100104_1.txt1(10).Text)) <= 0 Then
    strSQL1 = strSQL1 + " AND CP09 < 'C' "
    'Add By Cheng 2003/03/03
    strSQL1_1 = strSQL1_1 + " AND CP09 < 'C' "
'若含內部收文資料, 且含來函資料時
ElseIf Len(Trim(frm100104_1.txt1(15).Text)) > 0 And Len(Trim(frm100104_1.txt1(10).Text)) > 0 Then
   '不做處理
'若不含內部收文資料, 但含來函資料時
ElseIf Len(Trim(frm100104_1.txt1(15).Text)) <= 0 And Len(Trim(frm100104_1.txt1(10).Text)) > 0 Then
   'Modify By Cheng 2002/05/09
'    strSQL1 = strSQL1 + " AND (CP09 < 'B' Or CP09 > 'B') "
    strSQL1 = strSQL1 + " AND (CP09 < 'B' Or CP09 > 'C') "
    'Add By Cheng 2002/12/04
    strSQL1_1 = strSQL1_1 + " AND (CP09 < 'B' Or CP09 > 'C') "
'若不含內部收文資料, 且不含來函資料時
Else
    strSQL1 = strSQL1 + " AND CP09 < 'B' "
    'Add By Cheng 2002/12/04
    strSQL1_1 = strSQL1_1 + " AND CP09 < 'B' "
End If
If Len(Trim(frm100104_1.txt1(15))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label9(1), 10) & frm100104_1.txt1(15) 'Add By Sindy 2010/01/22
End If
If Len(Trim(frm100104_1.txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label9(0), 8) & frm100104_1.txt1(10) 'Add By Sindy 2010/01/22
End If

If Len(Trim(frm100104_1.txt1(6))) <> 0 Then
   strSQL1 = strSQL1 & " and cp10 not in (" & GetAddStr(frm100104_1.txt1(6)) & ") "
    'Add By Cheng 2002/12/04
   strSQL1_1 = strSQL1_1 & " and cp10 not in (" & GetAddStr(frm100104_1.txt1(6)) & ") "
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label5 & frm100104_1.txt1(6)  'Add By Sindy 2010/01/22
End If

'Add by Lydia 2015/02/12 + 是否只統計新申請案 (服務業務之新申請案案件性質為801、802、805、806)
If Len(Trim(frm100104_1.txt1(37))) <> 0 Then
    strExc(1) = "": strExc(2) = "": strExc(3) = ""
    'Modified by Lydia 2016/02/24 判斷跨部門權限
'    strExc(1) = SQLGrpStr(GetSystemKindByNick, 1) '專利
'    strExc(2) = SQLGrpStr(GetSystemKindByNick, 2) '商標
'    strExc(3) = SQLGrpStr(GetSystemKindByNick, 5) '服務
    If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, "ALL", strExc(4), False) Then
    End If
    strExc(1) = SQLGrpStr(strExc(4), 1) '專利
    strExc(2) = SQLGrpStr(strExc(4), 2) '商標
    strExc(3) = SQLGrpStr(strExc(4), 5) '服務
    'end 2016/02/24
    strExc(1) = Replace(strExc(1), ",' '", "")
    strExc(2) = Replace(strExc(2), ",' '", "")
    strExc(3) = Replace(strExc(3), ",' '", "")
    strExc(0) = ""
   If Len(strExc(1)) > 0 Then
      'Modified by Lydia 2016/08/01 + 含改請
      'strExc(0) = "(cp01 in (" & strExc(1) & ") and instr('" & NewCasePtyList & "',CP10)>0 ) "
      'Modified by Lydia 2025/09/19 改模組
      'strExc(0) = "(cp01 in (" & strExc(1) & ") and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3')) "
      strExc(0) = "(cp01 in (" & strExc(1) & ") and " & PUB_GetForNewCaseSql("1") & ") "
   End If
   If Len(strExc(2)) > 0 Then
      If Len(strExc(0)) > 0 Then strExc(0) = strExc(0) & " or "
      'Modified by Lydia 2025/09/19 改模組
      'strExc(0) = strExc(0) & "(cp01 in (" & strExc(2) & ") and CP10='101') "
      strExc(0) = strExc(0) & "(cp01 in (" & strExc(2) & ") " & PUB_GetForNewCaseSql("2") & ") "
   End If
   If Len(strExc(3)) > 0 Then
      If Len(strExc(0)) > 0 Then strExc(0) = strExc(0) & " or "
      'Modified by Lydia 2025/09/19 改模組
      'strExc(0) = strExc(0) & "(cp01 in (" & strExc(3) & ") and instr('801,802,805,806',CP10)>0) "
      strExc(0) = strExc(0) & "(cp01 in (" & strExc(3) & ") " & PUB_GetForNewCaseSql("5") & ") "
   End If
   strSQL1 = strSQL1 & " and (" & strExc(0) & ") "
   strSQL1_1 = strSQL1_1 & " and (" & strExc(0) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label9(9), 10) & frm100104_1.txt1(37)
End If
'end 2015/02/12

'Modify By Cheng 2003/06/18
''Add By Cheng 2002/04/24
''代理人
'If Len(Trim(frm100104_1.txt1(16).Text)) > 0 Then
'   strSQL1 = strSQL1 & " And CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' "
'End If
'If Len(Trim(frm100104_1.txt1(17).Text)) > 0 Then
'   strSQL1 = strSQL1 & " And CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' "
'End If
'Add By Cheng 2002/01/22
'申請人國籍
If Len(Trim(frm100104_1.txt1(13).Text)) > 0 Then
   strSQL1 = strSQL1 & " And C1.CU10 >= '" & frm100104_1.txt1(13).Text & "' "
    'Add By Cheng 2002/12/04
   strSQL1_1 = strSQL1_1 & " And C1.CU10 >= '" & frm100104_1.txt1(13).Text & "' "
End If
If Len(Trim(frm100104_1.txt1(14).Text)) > 0 Then
   '911203 nick 要將前 3 碼都抓出來
   'strSQL1 = strSQL1 & " And CU10 <= '" & frm100104_1.txt1(14).Text & "' "
   strSQL1 = strSQL1 & " And C1.CU10 <= '" & frm100104_1.txt1(14).Text & "z' "
    'Add By Cheng 2002/12/04
   strSQL1_1 = strSQL1_1 & " And C1.CU10 <= '" & frm100104_1.txt1(14).Text & "z' "
End If
If Len(Trim(frm100104_1.txt1(13).Text)) > 0 Or Len(Trim(frm100104_1.txt1(14).Text)) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label13 & frm100104_1.txt1(13) & "-" & frm100104_1.txt1(14) 'Add By Sindy 2010/01/22
End If
'Add By Cheng 2003/08/20
'是否只考慮有本所期限的資料
If Len(Trim(frm100104_1.txt1(23).Text)) > 0 Then
   strSQL1 = strSQL1 & " And CP06 Is Not Null "
   strSQL1_1 = strSQL1_1 & " And CP06 Is Not Null "
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label9(2), 14) & frm100104_1.txt1(23)  'Add By Sindy 2010/01/22
End If

'Add by Morgan 2011/1/17
'只列電子送件資料
'Modified by Morgan 2019/11/6 改為（空白：全部 1：電子送件 2：紙本送件）
'If frm100104_1.txt1(31).Text = "Y" Then
If frm100104_1.txt1(31) <> "" Then
   'Modified by Morgan 2013/7/16 +自動扣款的
   'strSQL1 = strSQL1 & " And CP118='Y'"
   'strSQL1_1 = strSQL1_1 & " And CP118='Y' "
   'Modified by Morgan 2019/11/6 有發文才算,紙本還要有經發文室--敏莉
   'strSQL1 = strSQL1 & " And (CP118='Y' or CP118='A') "
   'strSQL1_1 = strSQL1_1 & " And (CP118='Y' or CP118='A') "
   strSQL1 = strSQL1 & " and cp27>19221111"
   strSQL1_1 = strSQL1_1 & " and cp27>19221111"
   If frm100104_1.txt1(31) = "1" Then
      strSQL1 = strSQL1 & " And CP118 is not null"
      strSQL1_1 = strSQL1_1 & " And CP118 is not null"
   Else
      strSQL1 = strSQL1 & " And CP118 is null and cp123 is not null"
      strSQL1_1 = strSQL1_1 & " And CP118 is null and cp123 is not null"
   End If
'end 2019/11/6
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label9(6), 9) & frm100104_1.txt1(31)
End If

'Added by Lydia 2019/05/16 智慧局扣款日
If Trim(frm100104_1.txt1(40).Text) <> "" Then
   strSQL1 = strSQL1 & " And CP152=" & DBDATE(frm100104_1.txt1(40).Text)
   strSQL1_1 = strSQL1_1 & " And CP152=" & DBDATE(frm100104_1.txt1(40).Text)
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label23.Caption & frm100104_1.txt1(40).Text
End If

strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
'Add By Cheng 2002/12/04
strSQL1_2 = strSQL1_1
strSQL1_3 = strSQL1_1
strSQL1_4 = strSQL1_1
strSQL1_5 = strSQL1_1

If Len(Trim(frm100104_1.txt1(3))) <> 0 Then
   'Modify By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(frm100104_1.txt1(3), 1) & ") "
'   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(frm100104_1.txt1(3), 2) & ") "
'   StrSQL3 = StrSQL3 & " and cp01 in (" & SQLGrpStr(frm100104_1.txt1(3), 3) & ") "
'   StrSQL4 = StrSQL4 & " and cp01 in (" & SQLGrpStr(frm100104_1.txt1(3), 4) & ") "
'   StrSQL5 = StrSQL5 & " and cp01 in (" & SQLGrpStr(frm100104_1.txt1(3), 5) & ") "
   'Modified by Lydia 2016/02/25 "ALL"=使用者所有部門查詢權限
'   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 1) & ") "
'   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 2) & ") "
'   StrSQL3 = StrSQL3 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 3) & ") "
'   StrSQL4 = StrSQL4 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 4) & ") "
'   strSQL5 = strSQL5 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 5) & ") "
'    'Add By Cheng 2002/12/04
'   strSQL1_1 = strSQL1_1 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 1) & ") "
'   strSQL1_2 = strSQL1_2 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 2) & ") "
'   strSQL1_3 = strSQL1_3 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 3) & ") "
'   strSQL1_4 = strSQL1_4 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 4) & ") "
'   strSQL1_5 = strSQL1_5 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.Txt1(3).Text <> "ALL", frm100104_1.Txt1(3).Text, GetAllSysKind(frm100104_1.Txt1(3))), 5) & ") "
   If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, "ALL", strExc(4), False) Then
   End If
   
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 2) & ") "
   StrSQL3 = StrSQL3 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 3) & ") "
   StrSQL4 = StrSQL4 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 4) & ") "
   strSQL5 = strSQL5 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 5) & ") "
   strSQL1_1 = strSQL1_1 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 1) & ") "
   strSQL1_2 = strSQL1_2 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 2) & ") "
   strSQL1_3 = strSQL1_3 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 3) & ") "
   strSQL1_4 = strSQL1_4 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 4) & ") "
   strSQL1_5 = strSQL1_5 & " and cp01 in (" & SQLGrpStr(IIf(frm100104_1.txt1(3).Text <> "ALL", frm100104_1.txt1(3).Text, strExc(4)), 5) & ") "
   'end 2016/02/15
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label3, 5) & frm100104_1.txt1(3)  'Add By Sindy 2010/01/22
End If

'申請國家
If Len(Trim(frm100104_1.txt1(11).Text)) > 0 Then
   strSQL11 = " And PA09 >= '" & frm100104_1.txt1(11).Text & "' "
   strSQL21 = " And TM10 >= '" & frm100104_1.txt1(11).Text & "' "
   strSQL31 = " And LC15 >= '" & frm100104_1.txt1(11).Text & "' "
   strSQL41 = " And '000' >= '" & frm100104_1.txt1(11).Text & "' "
   strSQL51 = " And SP09 >= '" & frm100104_1.txt1(11).Text & "' "
End If
If Len(Trim(frm100104_1.txt1(12).Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA09 <= '" & frm100104_1.txt1(12).Text & "' "
   strSQL21 = strSQL21 & " And TM10 <= '" & frm100104_1.txt1(12).Text & "' "
   strSQL31 = strSQL31 & " And LC15 <= '" & frm100104_1.txt1(12).Text & "' "
   strSQL41 = strSQL41 & " And '000' <= '" & frm100104_1.txt1(12).Text & "' "
   strSQL51 = strSQL51 & " And SP09 <= '" & frm100104_1.txt1(12).Text & "' "
End If
If Len(Trim(frm100104_1.txt1(11).Text)) > 0 Or Len(Trim(frm100104_1.txt1(12).Text)) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label12 & frm100104_1.txt1(11) & "-" & frm100104_1.txt1(12) 'Add By Sindy 2010/01/22
End If
'Add By Cheng 2002/04/24
If Len(Trim(frm100104_1.txt1(18).Text)) > 0 Then
   'Modify By Sindy 2009/07/21
'   strSQL11 = " And PA26 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
'   strSQL21 = " And TM23 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
'   strSQL31 = " And LC11 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
'   strSQL41 = " And HC05 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
'   strSQL51 = " And SP08 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
   strSQL11 = strSQL11 & " And PA26 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
   strSQL21 = strSQL21 & " And TM23 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
   strSQL31 = strSQL31 & " And LC11 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
   strSQL41 = strSQL41 & " And HC05 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
   strSQL51 = strSQL51 & " And SP08 >= '" & Left(frm100104_1.txt1(18).Text & "000000000", 9) & "' "
End If
If Len(Trim(frm100104_1.txt1(19).Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA26 <= '" & Left(frm100104_1.txt1(19).Text & "000000000", 9) & "' "
   strSQL21 = strSQL21 & " And TM23 <= '" & Left(frm100104_1.txt1(19).Text & "000000000", 9) & "' "
   strSQL31 = strSQL31 & " And LC11 <= '" & Left(frm100104_1.txt1(19).Text & "000000000", 9) & "' "
   strSQL41 = strSQL41 & " And HC05 <= '" & Left(frm100104_1.txt1(19).Text & "000000000", 9) & "' "
   strSQL51 = strSQL51 & " And SP08 <= '" & Left(frm100104_1.txt1(19).Text & "000000000", 9) & "' "
End If
If Len(Trim(frm100104_1.txt1(18).Text)) > 0 Or Len(Trim(frm100104_1.txt1(19).Text)) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label15 & frm100104_1.txt1(18) & "-" & frm100104_1.txt1(19) 'Add By Sindy 2010/01/22
End If
'Add By Cheng 2002/12/04
'代理人
'add by nickc 2006/10/12 當前後都下時，原先方法會有 bug ex.cfp-15480
If Len(Trim(frm100104_1.txt1(16).Text)) > 0 And Len(Trim(frm100104_1.txt1(17).Text)) > 0 Then
       strSQL11 = strSQL11 & " And ((PA75 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and PA75 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "') Or (CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "')) "
       strSQL21 = strSQL21 & " And ((TM44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and TM44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "') Or (CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "')) "
       strSQL31 = strSQL31 & " And ((LC22 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and LC22 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "') Or (CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "')) "
       strSQL41 = strSQL41 & " And ((CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "')) "
       strSQL51 = strSQL51 & " And ((SP26 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and SP26 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "') Or (CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' and CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "')) "
       pub_QL05 = pub_QL05 & ";" & frm100104_1.Label14 & frm100104_1.txt1(16) & "-" & frm100104_1.txt1(17) 'Add By Sindy 2010/01/22
Else
    If Len(Trim(frm100104_1.txt1(16).Text)) > 0 Then
        'Modify By Cheng 2003/06/18
    '   strSQL11 = strSQL11 & " And PA75 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' "
    '   strSQL21 = strSQL21 & " And TM44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' "
    '   strSQL31 = strSQL31 & " And LC22 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' "
    '   strSQL51 = strSQL51 & " And SP26 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' "
       strSQL11 = strSQL11 & " And (PA75 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' Or CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' ) "
       strSQL21 = strSQL21 & " And (TM44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' Or CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' ) "
       strSQL31 = strSQL31 & " And (LC22 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' Or CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' ) "
       strSQL41 = strSQL41 & " And (CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' ) "
       strSQL51 = strSQL51 & " And (SP26 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "'  Or CP44 >= '" & Left(frm100104_1.txt1(16).Text & "000000000", 9) & "' ) "
    End If
    If Len(Trim(frm100104_1.txt1(17).Text)) > 0 Then
        'Modify By Cheng 2003/06/18
    '   strSQL11 = strSQL11 & " And PA75 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' "
    '   strSQL21 = strSQL21 & " And TM44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' "
    '   strSQL31 = strSQL31 & " And LC22 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' "
    '   strSQL51 = strSQL51 & " And SP26 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' "
       strSQL11 = strSQL11 & " And (PA75 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' Or CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' ) "
       strSQL21 = strSQL21 & " And (TM44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' Or CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' ) "
       strSQL31 = strSQL31 & " And (LC22 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' Or CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' ) "
       strSQL41 = strSQL41 & " And (CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' ) "
       strSQL51 = strSQL51 & " And (SP26 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' Or CP44 <= '" & Left(frm100104_1.txt1(17).Text & "000000000", 9) & "' ) "
    End If
    If Len(Trim(frm100104_1.txt1(16).Text)) > 0 Or Len(Trim(frm100104_1.txt1(17).Text)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100104_1.Label14 & frm100104_1.txt1(16) & "-" & frm100104_1.txt1(17) 'Add By Sindy 2010/01/22
    End If
End If

'Add By Sindy 2012/2/24
'Modify By Sindy 2013/11/18 +z,因國籍有4碼的,如輸入011就查不出0113,0114等,故加z
If Len(Trim(frm100104_1.txt1(32).Text)) > 0 And Len(Trim(frm100104_1.txt1(33).Text)) > 0 Then
       strSQL11 = strSQL11 & " And ((N4.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N4.NA01 <= '" & frm100104_1.txt1(33).Text & "z') Or (N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z')) "
       strSQL21 = strSQL21 & " And ((N2.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N2.NA01 <= '" & frm100104_1.txt1(33).Text & "z') Or (N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z')) "
       strSQL31 = strSQL31 & " And ((N2.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N2.NA01 <= '" & frm100104_1.txt1(33).Text & "z') Or (N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z')) "
       strSQL41 = strSQL41 & " And ((N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z')) "
       strSQL51 = strSQL51 & " And ((N2.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N2.NA01 <= '" & frm100104_1.txt1(33).Text & "z') Or (N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' and N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z')) "
       pub_QL05 = pub_QL05 & ";" & frm100104_1.Label19 & frm100104_1.txt1(32) & "-" & frm100104_1.txt1(33)
Else
    If Len(Trim(frm100104_1.txt1(32).Text)) > 0 Then
       strSQL11 = strSQL11 & " And (N4.NA01 >= '" & frm100104_1.txt1(32).Text & "' Or N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' ) "
       strSQL21 = strSQL21 & " And (N2.NA01 >= '" & frm100104_1.txt1(32).Text & "' Or N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' ) "
       strSQL31 = strSQL31 & " And (N2.NA01 >= '" & frm100104_1.txt1(32).Text & "' Or N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' ) "
       strSQL41 = strSQL41 & " And (N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' ) "
       strSQL51 = strSQL51 & " And (N2.NA01 >= '" & frm100104_1.txt1(32).Text & "'  Or N7.NA01 >= '" & frm100104_1.txt1(32).Text & "' ) "
    End If
    If Len(Trim(frm100104_1.txt1(33).Text)) > 0 Then
       strSQL11 = strSQL11 & " And (N4.NA01 <= '" & frm100104_1.txt1(33).Text & "z' Or N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z' ) "
       strSQL21 = strSQL21 & " And (N2.NA01 <= '" & frm100104_1.txt1(33).Text & "z' Or N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z' ) "
       strSQL31 = strSQL31 & " And (N2.NA01 <= '" & frm100104_1.txt1(33).Text & "z' Or N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z' ) "
       strSQL41 = strSQL41 & " And (N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z' ) "
       strSQL51 = strSQL51 & " And (N2.NA01 <= '" & frm100104_1.txt1(33).Text & "z' Or N7.NA01 <= '" & frm100104_1.txt1(33).Text & "z' ) "
    End If
    If Len(Trim(frm100104_1.txt1(32).Text)) > 0 Or Len(Trim(frm100104_1.txt1(33).Text)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100104_1.Label19 & frm100104_1.txt1(32) & "-" & frm100104_1.txt1(33)
    End If
End If
'2012/2/24 End

'Add By Cheng 2003/04/23
'商品類別
If Len(Trim(frm100104_1.txt1(20).Text)) > 0 Then
    arrTKind = Split(frm100104_1.txt1(20).Text, ",")
    For ii = LBound(arrTKind) To UBound(arrTKind)
        strTKind = strTKind & "'" & arrTKind(ii) & "',"
    Next ii
    If strTKind <> "" Then
        strTKind = Left(strTKind, Len(strTKind) - 1)
        strSQL21 = strSQL21 & " And TM09 IN (" & strTKind & " ) "
    End If
    pub_QL05 = pub_QL05 & ";" & frm100104_1.Label16 & frm100104_1.txt1(20) 'Add By Sindy 2010/01/22
End If
'Add By Cheng 2003/05/30
'專利/商標種類
If Len(Trim(frm100104_1.txt1(21).Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA08 >= '" & frm100104_1.txt1(21).Text & "' "
   strSQL21 = strSQL21 & " And TM08 >= '" & frm100104_1.txt1(21).Text & "' "
End If
If Len(Trim(frm100104_1.txt1(22).Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA08 <= '" & frm100104_1.txt1(22).Text & "' "
   strSQL21 = strSQL21 & " And TM08 <= '" & frm100104_1.txt1(22).Text & "' "
End If
If Len(Trim(frm100104_1.txt1(21).Text)) > 0 Or Len(Trim(frm100104_1.txt1(22).Text)) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label17 & frm100104_1.txt1(21) & "-" & frm100104_1.txt1(22) 'Add By Sindy 2010/01/22
End If
'add by nick  2005/02/04
If Trim(frm100104_1.txt1(27).Text) = "Y" Then
   strSQL11 = strSQL11 & " and pa46='Y' and pa09<>'056' "
   pub_QL05 = pub_QL05 & ";" & Left(frm100104_1.Label9(4), 10) & frm100104_1.txt1(27) 'Add By Sindy 2010/01/22
End If
'Add By Sindy 2009/05/22
'FCP管制人
If Len(Trim(frm100104_1.txt1(29).Text)) > 0 Then
   'Modify By Sindy 2016/7/15 增加案件性質區分抓取年費管制人及非年費管制人 ex.FCP-44616
   'strSQL11 = strSQL11 & " and NVL(NVL(N5.NA16,N6.NA16),NVL(N4.NA16,NVL(N3.NA16,N2.NA16)))>='" & Trim(frm100104_1.Txt1(29).Text) & "' "
   'Modified by Lydia 2017/02/13 2017/02/13+ FMP管制人
   If strSrvDate(1) < FMP管制人啟用日 Then
        strSQL11 = strSQL11 & " and ((cp10='605' and NVL(NVL(N5.NA16,N6.NA16),NVL(N4.NA16,NVL(N3.NA16,N2.NA16)))>='" & Trim(frm100104_1.txt1(29).Text) & "') or " & _
                                    "(cp10<>'605' and NVL(N4.NA16,N2.NA16)>='" & Trim(frm100104_1.txt1(29).Text) & "'))"
        '2016/7/15 END
        strSQL51 = strSQL51 & " and NVL(N2.NA16,N3.NA16)>='" & Trim(frm100104_1.txt1(29).Text) & "' "
   Else
        strSQL11 = strSQL11 & "and ((cp10='605' and DECODE(CP01,'P',NVL(NVL(NVL(NVL(NVL(N5.NA79,N5.NA16),NVL(N6.NA79,N6.NA16)),NVL(N4.NA79,N4.NA16)),NVL(N3.NA79,N3.NA16)),NVL(N2.NA79,N2.NA16)),NVL(NVL(N5.NA16,N6.NA16),NVL(N4.NA16,NVL(N3.NA16,N2.NA16))))>='" & Trim(frm100104_1.txt1(29).Text) & "') " & _
                              " or (cp10<>'605' and DECODE(CP01,'P',NVL(NVL(N4.NA79,N4.NA16),NVL(N2.NA79,N2.NA16)),NVL(N4.NA16,N2.NA16))>='" & Trim(frm100104_1.txt1(29).Text) & "')) "
        strSQL51 = strSQL51 & " AND DECODE(SP01,'PS',NVL(NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),NVL(N2.NA16,N3.NA16))>='" & Trim(frm100104_1.txt1(29).Text) & "' "
   End If
   'end 2017/02/13
End If
If Len(Trim(frm100104_1.txt1(30).Text)) > 0 Then
   'Modify By Sindy 2016/7/15 增加案件性質區分抓取年費管制人及非年費管制人 ex.FCP-44616
   'strSQL11 = strSQL11 & " and NVL(NVL(N5.NA16,N6.NA16),NVL(N4.NA16,NVL(N3.NA16,N2.NA16)))<='" & Trim(frm100104_1.Txt1(30).Text) & "' "
   'Modified by Lydia 2017/02/13 2017/02/13+ FMP管制人
   If strSrvDate(1) < FMP管制人啟用日 Then
        strSQL11 = strSQL11 & " and ((cp10='605' and NVL(NVL(N5.NA16,N6.NA16),NVL(N4.NA16,NVL(N3.NA16,N2.NA16)))<='" & Trim(frm100104_1.txt1(29).Text) & "') or " & _
                                    "(cp10<>'605' and NVL(N4.NA16,N2.NA16)<='" & Trim(frm100104_1.txt1(29).Text) & "'))"
        '2016/7/15 END
        strSQL51 = strSQL51 & " and NVL(N2.NA16,N3.NA16)<='" & Trim(frm100104_1.txt1(30).Text) & "' "
   Else
        strSQL11 = strSQL11 & "and ((cp10='605' and DECODE(CP01,'P',NVL(NVL(NVL(NVL(NVL(N5.NA79,N5.NA16),NVL(N6.NA79,N6.NA16)),NVL(N4.NA79,N4.NA16)),NVL(N3.NA79,N3.NA16)),NVL(N2.NA79,N2.NA16)),NVL(NVL(N5.NA16,N6.NA16),NVL(N4.NA16,NVL(N3.NA16,N2.NA16))))<='" & Trim(frm100104_1.txt1(30).Text) & "') " & _
                              " or (cp10<>'605' and DECODE(CP01,'P',NVL(NVL(N4.NA79,N4.NA16),NVL(N2.NA79,N2.NA16)),NVL(N4.NA16,N2.NA16))<='" & Trim(frm100104_1.txt1(30).Text) & "')) "
        strSQL51 = strSQL51 & " AND DECODE(SP01,'PS',NVL(NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),NVL(N2.NA16,N3.NA16))<='" & Trim(frm100104_1.txt1(30).Text) & "' "
   End If
   'end 2017/02/13
End If
If Len(Trim(frm100104_1.txt1(29).Text)) > 0 Or Len(Trim(frm100104_1.txt1(30).Text)) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label1(8) & frm100104_1.txt1(29) & "-" & frm100104_1.txt1(30) 'Add By Sindy 2010/01/22
End If
'2009/05/22 End

'Modify By Sindy 2012/3/7
'If frm100104_1.ChkPCT.Value = vbChecked Then
If frm100104_1.txt1(35) = "Y" Then
   pub_QL05 = pub_QL05 & ";" & "顯示PCT 案(僅查詢)" 'Add By Sindy 2010/01/22
End If

'Add By Sindy 2012/3/8 +國際分類
If Len(Trim(frm100104_1.txt1(34).Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA160 = '" & frm100104_1.txt1(34).Text & "' "
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label20 & frm100104_1.txt1(34)
End If
'Add By Sindy 2014/7/9
'專利發明/新型案件屬性
If Len(Trim(frm100104_1.Combo1.Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA158 = '" & Left(frm100104_1.Combo1.Text, 1) & "' And PA08<>'3' "
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label1(168) & frm100104_1.Combo1
End If
'台灣設計案件屬性
If Len(Trim(frm100104_1.Combo2.Text)) > 0 Then
   strSQL11 = strSQL11 & " And PA158 = '" & Left(frm100104_1.Combo2.Text, 1) & "' And PA08='3' "
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label1(1) & frm100104_1.Combo2
End If
'2014/7/9 END
'Add by Lydia 2014/11/18 增加FCP工程師組別
If Len(Trim(frm100104_1.txt1(36))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label9(8) & frm100104_1.txt1(36) & frm100104_1.lblName
   strSQL1 = strSQL1 + " And pa150 ='" & frm100104_1.txt1(36) & "' "
   strSQL2 = strSQL2 + "And TM01='" & frm100104_1.txt1(36) & "'" 'FCP工程師組別條件會令T,LC,HC案無資料
   StrSQL3 = StrSQL3 + "And LC01='" & frm100104_1.txt1(36) & "'"
   StrSQL4 = StrSQL4 + "And HC01='" & frm100104_1.txt1(36) & "'"
   strSQL5 = strSQL5 + "And SP79='" & frm100104_1.txt1(36) & "'"
End If
'end 'Add by Lydia 2014/11/18 增加FCP工程師組別

'Add By Sindy 2019/1/21
'特殊商標
If Len(Trim(frm100104_1.Combo3(0).Text)) > 0 Then
   strSQL21 = strSQL21 & " And TM72 >= '" & Left(frm100104_1.Combo3(0).Text, 1) & "' "
End If
If Len(Trim(frm100104_1.Combo3(1).Text)) > 0 Then
   strSQL21 = strSQL21 & " And TM72 <= '" & Left(frm100104_1.Combo3(1).Text, 1) & "' "
End If
If Len(Trim(frm100104_1.Combo3(0).Text)) > 0 Or Len(Trim(frm100104_1.Combo3(1).Text)) > 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100104_1.Label22 & frm100104_1.Combo3(0) & "-" & frm100104_1.Combo3(1)
End If
'2019/1/21 END

'edit by nickc 2005/05/10
'                    strSQL = "SELECT ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),pa26) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,patent,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=NA01(+) " & strSQL11 & strSQL1
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(DECODE(TM10,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),tm23) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,trademark,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And TM10=NA01(+) " & strSQL21 & strSQL2
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(DECODE(LC15,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),lc11) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,lawcase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and  cp14=s1.st01(+) and cp13=s2.st01(+) And LC15=NA01(+) " & strSQL31 & StrSQL3
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,nvl(cpm03,cp10)                          AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),hc05) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,hirecase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And NA01 = '000' " & strSQL41 & StrSQL4
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),sp08) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,servicepractice,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) And SP09=NA01(+) " & strSQL51 & strSQL5
'
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),pa26) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,patent,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=NA01(+) " & strSQL11 & strSQL1_1
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(DECODE(TM10,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),tm23) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,trademark,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And TM10=NA01(+) " & strSQL21 & strSQL1_2
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(DECODE(LC15,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),lc11) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,lawcase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and  cp14=s1.st01(+) and cp13=s2.st01(+) And LC15=NA01(+) " & strSQL31 & strSQL1_3
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,nvl(cpm03,cp10)                          AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),hc05) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,hirecase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And NA01 = '000' " & strSQL41 & strSQL1_4
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),sp08) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形
'FROM CASEPROGRESS,servicepractice,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) And SP09=NA01(+) " & strSQL51 & strSQL1_5
'add by nickc 2007/03/23 更換 PCT 欄
'                strSQL = "SELECT ' ' AS V," & SQLDate("CP05") & " AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),pa26) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,patent,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=NA01(+) " & strSQL11 & strSQL1
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(DECODE(TM10,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),tm23) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,trademark,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And TM10=NA01(+) " & strSQL21 & strSQL2
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(DECODE(LC15,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),lc11) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,lawcase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and  cp14=s1.st01(+) and cp13=s2.st01(+) And LC15=NA01(+) " & strSQL31 & StrSQL3
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,nvl(cpm03,cp10)                          AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),hc05) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,hirecase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And NA01 = '000' " & strSQL41 & StrSQL4
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),sp08) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,servicepractice,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) And SP09=NA01(+) " & strSQL51 & strSQL5
'Modify by Morgan 2007/7/20 加工作時數
'2010/9/10 MODIFY BY SONIA 改百年日期排序問題
'                strSQL = "SELECT ' ' AS V," & SQLDate("CP05") & " AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & IIf(frm100104_1.ChkPCT.Value = vbChecked, "pa46 as PCT", SQLDate("CP07") & " AS 法定期限") & "," & SQLDate("CP27") & " AS 發文日,nvl(NVL(C1.CU04,NVL(C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90,C1.CU06)),pa26) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,patent,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F1,FAGENT F2,FAGENT F3,customer C2,Nation N2,Nation N3,Nation N4,Nation N5,Nation N6 " & _
'                               " WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=C1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=C1.cu02(+) " & _
'                               " AND SUBSTR(C1.CU96,1,8)=F1.FA01(+) AND SUBSTR(C1.CU96||'0',9,1)=F1.FA02(+) AND SUBSTR(PA75,1,8)=F2.FA01(+) AND SUBSTR(PA75||'0',9,1)=F2.FA02(+) AND SUBSTR(PA76,1,8)=F3.FA01(+) AND SUBSTR(PA76||'0',9,1)=F3.FA02(+) AND SUBSTR(PA76,1,8)=C2.CU01(+) AND SUBSTR(PA76||'0',9,1)=C2.CU02(+) " & _
'                               " AND C1.CU10=N2.NA01(+) AND F1.FA10=N3.NA01(+) AND F2.FA10=N4.NA01(+) AND F3.FA10=N5.NA01(+) AND C2.CU10=N6.NA01(+) " & _
'                               " and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=N1.NA01(+) " & strSQL11 & strSQL1
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(DECODE(TM10,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & IIf(frm100104_1.ChkPCT.Value = vbChecked, "'' as PCT", SQLDate("CP07") & " AS 法定期限") & "," & SQLDate("CP27") & " AS 發文日,nvl(NVL(C1.CU04,NVL(C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90,C1.CU06)),tm23) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,trademark,customer C1,staff s1,staff s2,casepropertymap,Nation N1 WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=C1.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=C1.cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And TM10=N1.NA01(+) " & strSQL21 & strSQL2
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(DECODE(LC15,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & IIf(frm100104_1.ChkPCT.Value = vbChecked, "'' as PCT", SQLDate("CP07") & " AS 法定期限") & "," & SQLDate("CP27") & " AS 發文日,nvl(NVL(C1.CU04,NVL(C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90,C1.CU06)),lc11) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,lawcase,customer C1,staff s1,staff s2,casepropertymap,Nation N1 WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=C1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=C1.cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and  cp14=s1.st01(+) and cp13=s2.st01(+) And LC15=N1.NA01(+) " & strSQL31 & StrSQL3
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,nvl(cpm03,cp10)                          AS 案件性質,cp84 as 發文規費,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & IIf(frm100104_1.ChkPCT.Value = vbChecked, "'' as PCT", SQLDate("CP07") & " AS 法定期限") & "," & SQLDate("CP27") & " AS 發文日,nvl(NVL(C1.CU04,NVL(C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90,C1.CU06)),hc05) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,hirecase,customer C1,staff s1,staff s2,casepropertymap,Nation N1 WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=C1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=C1.cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And N1.NA01 = '000' " & strSQL41 & StrSQL4
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & IIf(frm100104_1.ChkPCT.Value = vbChecked, "'' as PCT", SQLDate("CP07") & " AS 法定期限") & "," & SQLDate("CP27") & " AS 發文日,nvl(NVL(C1.CU04,NVL(C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90,C1.CU06)),sp08) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,servicepractice,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F1,Nation N2,Nation N3 " & _
'                               " WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=C1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=C1.cu02(+) " & _
'                               " AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=F1.FA02(+) AND C1.CU10=N3.NA01(+) AND F1.FA10=N2.NA01(+) " & _
'                               " and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) And SP09=N1.NA01(+) " & strSQL51 & strSQL5
                
'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
'Modify By Sindy 2012/2/24 增加代理人國籍條件
'Modify By Sindy 2012/3/7 frm100104_1.ChkPCT.Value = vbChecked==>frm100104_1.txt1(35) = "Y"
'Modify by Amy 2016/06/22 +ST15||CP13/CP10/NA01,申請人no for 排序
'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'Modified by Lydia 2019/05/16 +發文時間,扣款日 (,SQLTIME6(CP82) AS 發文時間,SQLDATET(CP152) AS 扣款日)
'Modified by Lydia 2019/11/01 +增加欄位SeColTM,SeColPA,SeColSP,SeColLC,SeColHC
                strSql = "SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,SQLTIME6(CP82) AS 發文時間,SQLDATET(CP152) AS 扣款日,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,substr(GetSignF0202TypeNm(cp140),1,1) AS 簽,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限," & IIf(frm100104_1.txt1(35) = "Y", "pa46 as PCT", "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限") & ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(NVL(c1.CU04,DECODE(c1.cu05,null,c1.CU06,c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90)),pa26) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 出名,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09" & _
                               ",decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort,S2.ST15||CP13 as CP13,CP10,N1.NA01 as Nation,PA26 as Apply " & SeColPA & _
                               " FROM CASEPROGRESS,patent,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F1,FAGENT F2,FAGENT F3,customer C2,Nation N2,Nation N3,Nation N4,Nation N5,Nation N6,FAGENT F4,Nation N7 " & _
                               " WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=C1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=C1.cu02(+) " & _
                               " AND SUBSTR(C1.CU96,1,8)=F1.FA01(+) AND SUBSTR(C1.CU96||'0',9,1)=F1.FA02(+) AND SUBSTR(PA75,1,8)=F2.FA01(+) AND SUBSTR(PA75||'0',9,1)=F2.FA02(+) AND SUBSTR(PA76,1,8)=F3.FA01(+) AND SUBSTR(PA76||'0',9,1)=F3.FA02(+) AND SUBSTR(CP44,1,8)=F4.FA01(+) AND SUBSTR(CP44||'0',9,1)=F4.FA02(+) AND SUBSTR(PA76,1,8)=C2.CU01(+) AND SUBSTR(PA76||'0',9,1)=C2.CU02(+) " & _
                               " AND C1.CU10=N2.NA01(+) AND F1.FA10=N3.NA01(+) AND F2.FA10=N4.NA01(+) AND F3.FA10=N5.NA01(+) AND C2.CU10=N6.NA01(+) AND F4.FA10=N7.NA01(+) " & _
                               " and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=N1.NA01(+) " & strSQL11 & strSQL1
strSql = strSql & " union select ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(DECODE(TM10,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,SQLTIME6(CP82) AS 發文時間,SQLDATET(CP152) AS 扣款日,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,substr(GetSignF0202TypeNm(cp140),1,1) AS 簽,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限," & IIf(frm100104_1.txt1(35) = "Y", "'' as PCT", "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限") & ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(NVL(c1.CU04,DECODE(c1.cu05,null,c1.CU06,c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90)),tm23) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 出名,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09" & _
                               ",decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort,S2.ST15||CP13 as CP13,CP10,N1.NA01 as Nation,TM23 as Apply " & SeColTM & _
                               " FROM CASEPROGRESS,trademark,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F1,Nation N2,FAGENT F4,Nation N7 WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=C1.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=C1.cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And TM10=N1.NA01(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND SUBSTR(TM44||'0',9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F4.FA01(+) AND SUBSTR(CP44||'0',9,1)=F4.FA02(+) AND F1.FA10=N2.NA01(+) AND F4.FA10=N7.NA01(+) " & strSQL21 & strSQL2
strSql = strSql & " union select ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(DECODE(LC15,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,SQLTIME6(CP82) AS 發文時間,SQLDATET(CP152) AS 扣款日,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,substr(GetSignF0202TypeNm(cp140),1,1) AS 簽,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限," & IIf(frm100104_1.txt1(35) = "Y", "'' as PCT", "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限") & ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(NVL(c1.CU04,DECODE(c1.cu05,null,c1.CU06,c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90)),lc11) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 出名,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09" & _
                               ",decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort,S2.ST15||CP13 as CP13,CP10,N1.NA01 as Nation,LC11 as Apply " & SeColLC & _
                               " FROM CASEPROGRESS,lawcase,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F1,Nation N2,FAGENT F4,Nation N7 WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=C1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=C1.cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and  cp14=s1.st01(+) and cp13=s2.st01(+) And LC15=N1.NA01(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND SUBSTR(LC22||'0',9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F4.FA01(+) AND SUBSTR(CP44||'0',9,1)=F4.FA02(+) AND F1.FA10=N2.NA01(+) AND F4.FA10=N7.NA01(+) " & strSQL31 & StrSQL3
strSql = strSql & " union select ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,nvl(cpm03,cp10)                          AS 案件性質,cp84 as 發文規費,SQLTIME6(CP82) AS 發文時間,SQLDATET(CP152) AS 扣款日,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,substr(GetSignF0202TypeNm(cp140),1,1) AS 簽,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限," & IIf(frm100104_1.txt1(35) = "Y", "'' as PCT", "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限") & ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(NVL(c1.CU04,DECODE(c1.cu05,null,c1.CU06,c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90)),hc05) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 出名,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09" & _
                               ",decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort,S2.ST15||CP13 as CP13,CP10,N1.NA01 as Nation,HC05 as Apply " & SeColHC & _
                               " FROM CASEPROGRESS,hirecase,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F4,Nation N7 WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=C1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=C1.cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And N1.NA01 = '000' AND SUBSTR(CP44,1,8)=F4.FA01(+) AND SUBSTR(CP44||'0',9,1)=F4.FA02(+) AND F4.FA10=N7.NA01(+) " & strSQL41 & StrSQL4
strSql = strSql & " union select ' ' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,SQLTIME6(CP82) AS 發文時間,SQLDATET(CP152) AS 扣款日,nvl(cp113,cp114) as 工作時數,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員,substr(GetSignF0202TypeNm(cp140),1,1) AS 簽,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限," & IIf(frm100104_1.txt1(35) = "Y", "'' as PCT", "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限") & ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,nvl(NVL(c1.CU04,DECODE(c1.cu05,null,c1.CU06,c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90)),sp08) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 出名,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,nvl(N1.NA03,N1.NA04) As 申請國家, CP09" & _
                               ",decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort,S2.ST15||CP13 as CP13,CP10,N1.NA01 as Nation,SP08 as Apply " & SeColSP & _
                               " FROM CASEPROGRESS,servicepractice,customer C1,staff s1,staff s2,casepropertymap,Nation N1,FAGENT F1,Nation N2,Nation N3,FAGENT F4,Nation N7 " & _
                               " WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=C1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=C1.cu02(+) " & _
                               " AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F4.FA01(+) AND SUBSTR(CP44||'0',9,1)=F4.FA02(+) AND C1.CU10=N3.NA01(+) AND F1.FA10=N2.NA01(+) AND F4.FA10=N7.NA01(+) " & _
                               " and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) And SP09=N1.NA01(+) " & strSQL51 & strSQL5
'2010/9/10 END
'end 2007/7/20
'edit by nickc 2006/07/10 重複的語法
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(DECODE(PA09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),pa26) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,patent,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And PA09=NA01(+) " & strSQL11 & strSQL1_1
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(DECODE(TM10,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),tm23) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,trademark,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And TM10=NA01(+) " & strSQL21 & strSQL1_2
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(DECODE(LC15,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),lc11) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,lawcase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and  cp14=s1.st01(+) and cp13=s2.st01(+) And LC15=NA01(+) " & strSQL31 & strSQL1_3
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,nvl(cpm03,cp10)                          AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),hc05) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,hirecase,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) And NA01 = '000' " & strSQL41 & strSQL1_4
'strSQL = strSQL & " union select ' ' AS V," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質,cp84 as 發文規費,nvl(s1.st02,CP14) AS 承辦人,nvl(s2.st02,CP13) AS 智權人員," & SQLDate("CP06") & " AS 本所期限," & SQLDate("CP07") & " AS 法定期限," & SQLDate("CP27") & " AS 發文日,nvl(NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),sp08) AS 申請人,CP18 AS 點數,decode(CP22,'Y','是','N','否','') AS 是否出名," & SQLDate("CP57") & " AS 取消收文日,nvl(NA03,NA04) As 申請國家, CP09, CP60 As 收款情形,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort " & _
'                               " FROM CASEPROGRESS,servicepractice,customer,staff s1,staff s2,casepropertymap,Nation WHERE cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) And SP09=NA01(+) " & strSQL51 & strSQL1_5

'End

'Added by Lydia 2015/11/04 傳SQL到列印
frm100104_1.mQueryStr = strSql

'Add By Cheng 2003/08/20
'若只考慮有本所期限的資料
If frm100104_1.txt1(23).Text <> "" Then
    If Trim(frm100104_1.txt1(0).Text) = "1" Then
        'edit by nickc 2005/05/10
        'strSQL = strSQL + " ORDER BY 本所期限, 本所案號"
        strSql = strSql + " ORDER BY 本所期限,FSort, 本所案號"
    Else
        'edit by nickc 2005/05/10
        'strSQL = strSQL + " ORDER BY 本所期限, 本所案號"
        strSql = strSql + " ORDER BY 本所期限, FSort,本所案號"
    End If
'Added by Lydia 2019/05/16 智慧局扣款日: 依發文日期+發文時間+本所案號排序
ElseIf Trim(frm100104_1.txt1(40).Text) <> "" Then
        strSql = strSql + " ORDER BY 發文日,發文時間, FSort,本所案號"
        
'若不考慮是否有本所期限的資料
Else
    'Modify By Cheng 2002/03/05
    'If Len(Trim(frm100104_1.txt1(0))) = "1" Then
    If Trim(frm100104_1.txt1(0).Text) = "1" Then
        'edit by nickc 2005/05/10
        'strSQL = strSQL + " ORDER BY 收文日,本所案號"
        strSql = strSql + " ORDER BY 收文日,FSort,本所案號"
    Else
        'edit by nickc 2005/05/10
        'strSQL = strSQL + " ORDER BY 發文日,本所案號"
        strSql = strSql + " ORDER BY 發文日,FSort,本所案號"
    End If
End If
lbl1(0) = frm100104_1.txt1(1) + "－" + frm100104_1.txt1(2)
'Move by Lydia 2015/11/04

'若為列印
If Trim(frm100104_1.txt1(9)) = "2" Then
    'Added by Lydia 2019/11/01 '需要經過利益衝突案件的比對
    frm100104_1.mQueryStrEsc = ""
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        GoTo JumpToCheck
    End If
    'end 2019/11/01
    
    'Remove by Lydia 2015/11/04 傳SQL到列印
'   With adoRecordset
'      .MoveFirst
'      cnnConnection.Execute "delete from r100104 where id='" & strUserNum & "' "
'      Do While .EOF = False
'         If frm100104_1.txt1(0) = "1" Then
'            'Modify By Cheng 2002/10/22
''            cnnConnection.Execute "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(3))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & "','" & ChgSQL(CheckStr(.Fields(6))) & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(10))) & "'," & Val(CheckStr(.Fields(11))) & ",'" & ChgSQL(CheckStr(.Fields(12))) & "','" & ChgSQL(CheckStr(.Fields(13))) & "','" & strUserNum & "') "
'            'Modify By Cheng 2003/08/15
''            strSQLA = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(3))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & "','" & ChgSQL(CheckStr(.Fields(6))) & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(12))) & "'," & Val(CheckStr(.Fields(11))) & ",'" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(13))) & "','" & strUserNum & "') "
'            'edit by nick 2004/08/26
'            'strSQLA = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(3))) & "','" & ChgSQL(CheckStr(.Fields(4))) & PUB_GetRelateCasePropertyName("" & .Fields("CP09").Value, "1") & "','" & ChgSQL(CheckStr(.Fields(5))) & "','" & ChgSQL(CheckStr(.Fields(6))) & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(12))) & "'," & Val(CheckStr(.Fields(11))) & ",'" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(13))) & "','" & strUserNum & "') "
'            'Modify by Morgan 2007/10/1 有加欄位
'            'StrSQLa = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & PUB_GetRelateCasePropertyName("" & .Fields("CP09").Value, "1") & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(11))) & "','" & ChgSQL(CheckStr(.Fields(14))) & "'," & Val(CheckStr(.Fields(13))) & ",'" & ChgSQL(CheckStr(.Fields(12))) & "','" & ChgSQL(CheckStr(.Fields(15))) & "','" & strUserNum & "') "
'            StrSQLa = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & PUB_GetRelateCasePropertyName("" & .Fields("CP09").Value, "1") & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(11))) & "','" & ChgSQL(CheckStr(.Fields(12))) & "','" & ChgSQL(CheckStr(.Fields(15))) & "'," & Val(CheckStr(.Fields(14))) & ",'" & ChgSQL(CheckStr(.Fields(13))) & "','" & ChgSQL(CheckStr(.Fields(16))) & "','" & strUserNum & "') "
'            'end 2007/10/1
'            cnnConnection.Execute StrSQLa
'         Else
'            'Modify By Cheng 2002/10/22
''            cnnConnection.Execute "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(3))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & "','" & ChgSQL(CheckStr(.Fields(6))) & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(10))) & "'," & Val(CheckStr(.Fields(11))) & ",'" & ChgSQL(CheckStr(.Fields(12))) & "','" & ChgSQL(CheckStr(.Fields(13))) & "','" & strUserNum & "') "
'            'Modify By Cheng 2003/08/15
''            strSQLA = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(3))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & "','" & ChgSQL(CheckStr(.Fields(6))) & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(12))) & "'," & Val(CheckStr(.Fields(11))) & ",'" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(13))) & "','" & strUserNum & "') "
'            'edit by nick 2004/08/26
'            'strSQLA = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(3))) & "','" & ChgSQL(CheckStr(.Fields(4))) & PUB_GetRelateCasePropertyName("" & .Fields("CP09").Value, "1") & "','" & ChgSQL(CheckStr(.Fields(5))) & "','" & ChgSQL(CheckStr(.Fields(6))) & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(12))) & "'," & Val(CheckStr(.Fields(11))) & ",'" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(13))) & "','" & strUserNum & "') "
'            'Modify by Morgan 2007/10/1 有加欄位
'            'StrSQLa = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(11))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & PUB_GetRelateCasePropertyName("" & .Fields("CP09").Value, "1") & "','" & ChgSQL(CheckStr(.Fields(7))) & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(14))) & "'," & Val(CheckStr(.Fields(13))) & ",'" & ChgSQL(CheckStr(.Fields(12))) & "','" & ChgSQL(CheckStr(.Fields(15))) & "','" & strUserNum & "') "
'            StrSQLa = "insert into r100104 values ('" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(12))) & "','" & ChgSQL(CheckStr(.Fields(2))) & "','" & ChgSQL(CheckStr(.Fields(4))) & "','" & ChgSQL(CheckStr(.Fields(5))) & PUB_GetRelateCasePropertyName("" & .Fields("CP09").Value, "1") & "','" & ChgSQL(CheckStr(.Fields(8))) & "','" & ChgSQL(CheckStr(.Fields(9))) & "','" & ChgSQL(CheckStr(.Fields(10))) & "','" & ChgSQL(CheckStr(.Fields(11))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "','" & ChgSQL(CheckStr(.Fields(15))) & "'," & Val(CheckStr(.Fields(14))) & ",'" & ChgSQL(CheckStr(.Fields(13))) & "','" & ChgSQL(CheckStr(.Fields(16))) & "','" & strUserNum & "') "
'            'end 2007/10/1
'            cnnConnection.Execute StrSQLa
'         End If
'         .MoveNext
'      Loop
'   End With
    'end 2015/11/04
    
'若為查詢
Else
    'Move by Lydia 2015/11/04 從上面移到這裡
JumpToCheck:   'Added by Lydia 2019/11/01
    CheckOC
    Dim StrTest1 As String, StrTest2 As String
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
                StrTest2 = adoRecordset.Fields(2)
                If Left(StrTest2, 1) = "N" Then
                   StrTest2 = Right(StrTest2, Len(StrTest2) - 1)
                End If
                If StrTest1 <> StrTest2 Then
                    '利益衝突案件：逐案號判斷
                    If PUB_ChkCufaByCase(Me.Name, m_AllSys, StrTest2, "" & adoRecordset.Fields("apply") & "," & adoRecordset.Fields("apply02") & "," & adoRecordset.Fields("apply03") & "," & adoRecordset.Fields("apply04") & "," & adoRecordset.Fields("apply05"), "" & adoRecordset.Fields("fcno")) = False Then
                        intCufaCnt = intCufaCnt + 1
                        strTmp = strTmp & StrTest2 & ","
                        adoRecordset.Delete
                    Else
                        StrTest1 = StrTest2
                    End If
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
            If adoRecordset.RecordCount = 0 Then
                If Trim(frm100104_1.txt1(9)) = "2" Then
                     InsertQueryLog (0)
                     frm100104_1.mQueryStrEsc = strTmp
                     GoTo JumpToPrint
                Else
                     GoTo JumpToNoData
                End If
            End If
       Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
       End If
       'end 2019/11/01
       lblTot = "共 " & adoRecordset.RecordCount & " 筆" 'Added by Morgan 2019/11/6
       frm100104_1.m_blnNoData = False
       Me.cmdOK(0).Enabled = True
       Me.cmdOK(1).Enabled = True
    Else
       InsertQueryLog (0) 'Add By Sindy 2010/01/22
JumpToNoData:           'Added by Lydia 2019/11/01
       frm100104_1.m_blnNoData = True
       ShowNoData
       Me.cmdOK(0).Enabled = False
       Me.cmdOK(1).Enabled = False
       Me.Enabled = True
       Screen.MousePointer = vbDefault
       '92.04.18 nick
       'me.hide
       tmpBol = fnCancelNowFormAndShowParentForm(Me)
       Exit Sub
    End If
    'Added by Lydia 2019/11/01 後面回到frm100104_1列印
    If Trim(frm100104_1.txt1(9)) = "2" Then
        frm100104_1.mQueryStrEsc = strTmp
        GoTo JumpToPrint
    End If
    'end 2019/11/01
    
    Me.grdDataList.Visible = False
    Set grdDataList.Recordset = adoRecordset: DoEvents
    SetDataListWidth
    'Modify By Sindy 2019/12/12 會錯,移至上面
'    lblTot = "共 " & adoRecordset.RecordCount & " 筆" 'Added by Morgan 2019/11/6
    '2019/12/12 END
    'Add By Cheng 2003/08/15
    For ii = 1 To Me.grdDataList.Rows - 1
        'edit by nick 2004/08/23
        'Me.grdDataList.TextMatrix(ii, 4) = Me.grdDataList.TextMatrix(ii, 4) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 15), "1")
        'Modify by Morgan 2007/7/20
        'Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 16), "1")
        Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 18), "1")
        'end 2007/7/20
        'Add By Cheng 2004/02/09
        '收款情形
        Dim IntTemp1 As Long
        Dim IntTemp2 As Long
        IntTemp1 = 0
        IntTemp2 = 0
        Me.grdDataList.row = ii
        'edit by nick 2004/08/23
        'Me.grdDataList.col = 16
        Me.grdDataList.col = 19
        If Not IsNull(grdDataList.Text) And grdDataList.Text <> "" Then
            '2009/12/8 modify by sonia 加請款單
            'strSQL = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
            'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
            'If Mid(grdDataList.Text, 1, 1) = "E" Then
            '   strSql = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
            'Else
            If Mid(grdDataList.Text, 1, 1) = "X" Then
            'end 2011/8/15
            
               strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0 FROM ACC1K0 WHERE A1K01='" & grdDataList.Text & "'"
            'End If 'Remove by Morgan 2011/8/15
            '2009/12/8 end
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   If Not IsNull(adoRecordset1.Fields(0)) Then
                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
                   End If
                   If Not IsNull(adoRecordset1.Fields(1)) Then
                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
                   End If
                   If Not IsNull(adoRecordset1.Fields(4)) Then
                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
                   End If
                   If Not IsNull(adoRecordset1.Fields(5)) Then
                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
                   End If
                   If IntTemp1 = IntTemp2 Then
                        grdDataList.Text = "收回"
                   Else
                        If IntTemp2 = 0 Then
                            grdDataList.Text = "未收"
                        Else
                            If IntTemp1 > IntTemp2 Then
                                grdDataList.Text = "部分收回"
                            End If
                        End If
                    End If
               Else
                    'grdDataList.Text = "查無此收據編號"   '2010/3/18 CANCEL BY SONIA
               End If
            End If 'Add by Morgan 2011/8/15
        End If
        CheckOC2
        DoEvents
        'End
    Next ii
    Me.grdDataList.Visible = True
End If
JumpToPrint: 'Added by Lydia 2019/11/01
CheckOC
Me.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100104_2 = Nothing
End Sub

'Add by Amy 2016/06/22
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If grdDataList.MouseCol < 0 Or grdDataList.MouseRow < 0 Then Exit Sub
    grdDataList.col = grdDataList.MouseCol
    grdDataList.row = grdDataList.MouseRow
    If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
        '數字
        If grdDataList.col = GetValue("智權人員") Then grdDataList.col = GetValue("智權人員", True)
        If grdDataList.col = GetValue("案件性質") Then grdDataList.col = GetValue("案件性質", True)
        If grdDataList.col = GetValue("申請國家") Then grdDataList.col = GetValue("申請國家", True)
        If grdDataList.col = GetValue("申請人") Then grdDataList.col = GetValue("申請人", True)
        
        If grdDataList.col = GetValue("CP10") Or grdDataList.col = GetValue("Nation") Then
            If m_blnColOrderAsc = True Then
                Me.grdDataList.Sort = 3 '數值昇冪
                m_blnColOrderAsc = False
            Else
                Me.grdDataList.Sort = 4 '數值降冪
                m_blnColOrderAsc = True
            End If
        '文字
        Else
            If m_blnColOrderAsc = True Then
                Me.grdDataList.Sort = 5 '字串昇冪
                m_blnColOrderAsc = False
            Else
                Me.grdDataList.Sort = 6 '字串降冪
                m_blnColOrderAsc = True
            End If
        End If
    End If
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

Sub StrMenu1()
StrMenu
End Sub

'Add by Amy 2016/06/22
Private Function GetValue(pFieldN As String, Optional ByVal bolChange As Boolean = False) As Integer
   Dim jj As Integer, ii As Integer
   Dim strFind As String
 
    For jj = 1 To UBound(strFieldN)
        If UCase(strFieldN(jj)) = UCase(pFieldN) Then
            If bolChange = True Then
                Select Case UCase(pFieldN)
                    Case "智權人員"
                        strFind = "CP13"
                    Case "案件性質"
                        strFind = "CP10"
                    Case "申請國家"
                        strFind = "Nation"
                    Case "申請人"
                        strFind = "Apply"
                End Select
                For ii = 1 To UBound(strFieldN)
                    If UCase(strFieldN(ii)) = UCase(strFind) Then
                        GetValue = ii
                        Exit For
                    End If
                Next ii
            Else
                GetValue = jj
            End If
            Exit For
        End If
    Next jj
End Function
