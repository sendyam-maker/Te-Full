VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100110_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案件查詢"
   ClientHeight    =   5720
   ClientLeft      =   210
   ClientTop       =   990
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   3
      Left            =   3390
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   75
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   4
      Left            =   4935
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   75
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8496
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   0
      Left            =   6048
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   75
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7272
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   75
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4920
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8678
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
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
      Height          =   195
      Index           =   10
      Left            =   7020
      TabIndex        =   8
      Top             =   525
      Width           =   2025
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1230
      TabIndex        =   4
      Top             =   525
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "機關文號 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   525
      Width           =   870
   End
End
Attribute VB_Name = "frm100110_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
'grdDataList.Cols = 14
Dim intField As Integer
intField = 20
grdDataList.Cols = intField
'end 2019/11/01

   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   'Add By Cheng 2002/07/01
   grdDataList.CellAlignment = flexAlignCenterCenter
   '2013/8/19 add by sonia 加入法院案號條件 P-099556
   If frm100110_1.Option1(4).Value = True Then
      grdDataList.col = 1: grdDataList.Text = "法院案號"
   Else
   '2013/8/19 end
      grdDataList.col = 1: grdDataList.Text = "機關文號"
   End If
   grdDataList.ColWidth(1) = 1850
   '2002/07/01 End
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "收文日"
   grdDataList.ColWidth(2) = 810
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "本所案號"
   grdDataList.ColWidth(3) = 1450
   grdDataList.CellAlignment = flexAlignCenterCenter
   Dim iDep As String
   iDep = PUB_GetST06(strUserNum)
   grdDataList.col = 4: grdDataList.Text = "分所號"
   '電腦中心，跟分所才秀
   If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
       grdDataList.ColWidth(4) = 0
   Else
       grdDataList.ColWidth(4) = 620
   End If
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "案件名稱"
   grdDataList.ColWidth(5) = 1400
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "案件性質"
   grdDataList.ColWidth(6) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 7: grdDataList.Text = "證書審定號"
   grdDataList.ColWidth(7) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 8: grdDataList.Text = "承辦人"
   grdDataList.ColWidth(8) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 9: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(9) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 10: grdDataList.Text = "目前准駁"
   grdDataList.ColWidth(10) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 11: grdDataList.Text = "專用權是否存在"
   grdDataList.ColWidth(11) = 1400
   grdDataList.CellAlignment = flexAlignCenterCenter
   'Add By Cheng 2003/08/15
   grdDataList.col = 12: grdDataList.Text = "CP09"
   grdDataList.ColWidth(12) = 0
   grdDataList.CellAlignment = flexAlignCenterCenter
   'add by nickc 2005/05/10
   grdDataList.col = 13: grdDataList.Text = ""
   grdDataList.ColWidth(13) = 0
   grdDataList.CellAlignment = flexAlignCenterCenter
   
'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 14 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Dim Str01 As String
Dim StrTag As String

   Select Case cmdState
      Case 0 '相關卷號
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
               grdDataList.col = 3
               If Not IsNull(grdDataList.Text) Then
                  Screen.MousePointer = vbHourglass
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  frm100108_3.Show
                  frm100108_3.Tag = Pub_RplStr(grdDataList.Text)
                  frm100108_3.StrMenu1
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      Case 1 '回前畫面
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 2
         fnCloseAllFrm100
      Case 3 '案件基本資料
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
               grdDataList.col = 3
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
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      Case 4 '案件進度
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
               grdDataList.col = 3
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
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()        '機關文號
Dim ii As Integer
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

   ClearQueryLog ("frm100110_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
   Me.Enabled = False
   '2013/8/19 MODIFY BY SONIA 加入法院案號條件
   'Label2.Caption = frm100110_1.txt1(0)
   If frm100110_1.Option1(0).Value = True Then
      Label1(0).Caption = "機關文號 :"
      Label2.Caption = frm100110_1.txt1(0)
   Else
      Label1(0).Caption = "法院案號 :"
      Label2.Caption = frm100110_1.txt1(13)
   End If
   '2013/8/19 END
   
   'Add By Sindy 2010/02/05 增加查詢Lawcase,Hirecase,Serverpractice
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
   strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
   
    'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
    SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
    SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
    SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
    SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
    SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
    m_AllSys = GetAllSysKind(, "ALL")
    intCufaCnt = 0
    'end 2019/11/01
      
   '收文起日
   If Len(Trim(frm100110_1.txt1(4))) <> 0 Then
       strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
       strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
       StrSQL3 = StrSQL3 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
       StrSQL4 = StrSQL4 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
       strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
   End If
   '收文迄日
   If Len(Trim(frm100110_1.txt1(5))) <> 0 Then
       strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
       strSQL2 = strSQL2 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
       StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
       StrSQL4 = StrSQL4 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
       strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
   End If
   If Len(Trim(frm100110_1.txt1(4))) <> 0 Or Len(Trim(frm100110_1.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100110_1.Label5 & frm100110_1.txt1(4) & "-" & frm100110_1.txt1(5) 'Add By Sindy 2010/11/3
   End If
   '案件性質
   If Len(Trim(frm100110_1.txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
      strSQL2 = strSQL2 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
      StrSQL3 = StrSQL3 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
      StrSQL4 = StrSQL4 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
      strSQL5 = strSQL5 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
      pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(0) & frm100110_1.txt1(6) 'Add By Sindy 2010/11/3
   End If

   '申請人起
   'Modified by Lydia 2019/11/01 改成申請人1~5
'   If Len(Trim(frm100110_1.txt1(7))) <> 0 Then
'       strSQL1 = strSQL1 + " AND TM23>='" & frm100110_1.txt1(7) & "' "
'       strSQL2 = strSQL2 + " AND PA26>='" & frm100110_1.txt1(7) & "' "
'       StrSQL3 = StrSQL3 + " AND LC11>='" & frm100110_1.txt1(7) & "' "
'       StrSQL4 = StrSQL4 + " AND HC05>='" & frm100110_1.txt1(7) & "' "
'       strSQL5 = strSQL5 + " AND SP08>='" & frm100110_1.txt1(7) & "' "
'   End If
'   '申請人迄
'   If Len(Trim(frm100110_1.txt1(8))) <> 0 Then
'      strSQL1 = strSQL1 & " AND TM23<='" & frm100110_1.txt1(8) & "' "
'      strSQL2 = strSQL2 & " AND PA26<='" & frm100110_1.txt1(8) & "' "
'      StrSQL3 = StrSQL3 & " AND LC11<='" & frm100110_1.txt1(8) & "' "
'      StrSQL4 = StrSQL4 & " AND HC05<='" & frm100110_1.txt1(8) & "' "
'      strSQL5 = strSQL5 & " AND SP08<='" & frm100110_1.txt1(8) & "' "
'   End If
   'Memo by Lydia 2019/11/01 改成變數
   strExc(1) = ""
   If Len(Trim(frm100110_1.txt1(7))) <> 0 And Len(Trim(frm100110_1.txt1(8))) <> 0 Then
       strExc(1) = " AND ((PA26>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA26<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA27>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA27<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA28>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA28<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA29>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA29<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                 " OR (PA30>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  AND PA30<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )) "
   Else
       If Len(Trim(frm100110_1.txt1(7))) <> 0 Then
            strExc(1) = " AND ((PA26>='" & GetNewFagent(frm100110_1.txt1(7)) & "' )" & _
                                      " OR (PA27>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )" & _
                                      " OR (PA28>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )" & _
                                      " OR (PA29>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )" & _
                                      " OR (PA30>='" & GetNewFagent(frm100110_1.txt1(7)) & "'  )) "
       End If
       If Len(Trim(frm100110_1.txt1(8))) <> 0 Then
            strExc(1) = " AND ((PA26<='" & GetNewFagent(frm100110_1.txt1(8)) & "' )" & _
                                      " OR (PA27<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )" & _
                                      " OR (PA28<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )" & _
                                      " OR (PA29<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )" & _
                                      " OR (PA30<='" & GetNewFagent(frm100110_1.txt1(8)) & "'  )) "
       End If
   End If
   'end 2019/11/01
   
   If Len(Trim(frm100110_1.txt1(7))) <> 0 Or Len(Trim(frm100110_1.txt1(8))) <> 0 Then
      'Added by Lydia 2019/11/01 合併SQL
      strSQL1 = strSQL1 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "TM23"), "PA27", "TM78"), "PA28", "TM79"), "PA29", "TM80"), "PA30", "TM81")
      strSQL2 = strSQL2 & strExc(1)
      StrSQL3 = StrSQL3 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "LC11"), "PA27", "LC43"), "PA28", "LC44"), "PA29", "LC45"), "PA30", "LC46")
      StrSQL4 = StrSQL4 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "HC05"), "PA27", "HC24"), "PA28", "HC25"), "PA29", "HC26"), "PA30", "HC27")
      strSQL5 = strSQL5 & Replace(Replace(Replace(Replace(Replace(strExc(1), "PA26", "SP08"), "PA27", "SP58"), "PA28", "SP59"), "PA29", "SP65"), "PA30", "SP66")
      'end 2019/11/01
      pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(1) & frm100110_1.txt1(7) & "-" & frm100110_1.txt1(8) 'Add By Sindy 2010/11/3
   End If
   '代理人起
   If Len(Trim(frm100110_1.txt1(9))) <> 0 Then
       strSQL1 = strSQL1 + " AND TM44>='" & frm100110_1.txt1(9) & "' "
       strSQL2 = strSQL2 + " AND PA75>='" & frm100110_1.txt1(9) & "' "
       StrSQL3 = StrSQL3 + " AND LC22>='" & frm100110_1.txt1(9) & "' "
   '    strSQL4 = 無代理人
       strSQL5 = strSQL5 + " AND SP26>='" & frm100110_1.txt1(9) & "' "
   End If
   '代理人迄
   If Len(Trim(frm100110_1.txt1(10))) <> 0 Then
      strSQL1 = strSQL1 & " AND TM44<='" & frm100110_1.txt1(10) & "' "
      strSQL2 = strSQL2 & " AND PA75<='" & frm100110_1.txt1(10) & "' "
      StrSQL3 = StrSQL3 & " AND LC22<='" & frm100110_1.txt1(10) & "' "
   '   strSQL4 = 無代理人
      strSQL5 = strSQL5 & " AND SP26<='" & frm100110_1.txt1(10) & "' "
   End If
   If Len(Trim(frm100110_1.txt1(9))) <> 0 Or Len(Trim(frm100110_1.txt1(10))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(2) & frm100110_1.txt1(9) & "-" & frm100110_1.txt1(10) 'Add By Sindy 2010/11/3
   End If
   
   '2013/8/19 add by sonia 加入法院案號條件 P-099556
   If frm100110_1.Option1(4).Value = True Then
      '法院案號
      If Len(Trim(frm100110_1.txt1(13))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP35 = '" & frm100110_1.txt1(13) & "' "
         strSQL2 = strSQL2 & " AND CP35 = '" & frm100110_1.txt1(13) & "' "
         StrSQL3 = StrSQL3 & " AND CP35 = '" & frm100110_1.txt1(13) & "' "
         StrSQL4 = StrSQL4 & " AND CP35 = '" & frm100110_1.txt1(13) & "' "
         strSQL5 = strSQL5 & " AND CP35 = '" & frm100110_1.txt1(13) & "' "
         pub_QL05 = pub_QL05 & ";" & frm100110_1.Option1(4).Caption & frm100110_1.txt1(13)
      End If
   Else
   '2013/8/19 end
      '機關文號
      If Len(Trim(frm100110_1.txt1(0))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP08 Like '%" & frm100110_1.txt1(0) & "%' "
         strSQL2 = strSQL2 & " AND CP08 Like '%" & frm100110_1.txt1(0) & "%' "
         StrSQL3 = StrSQL3 & " AND CP08 Like '%" & frm100110_1.txt1(0) & "%' "
         StrSQL4 = StrSQL4 & " AND CP08 Like '%" & frm100110_1.txt1(0) & "%' "
         strSQL5 = strSQL5 & " AND CP08 Like '%" & frm100110_1.txt1(0) & "%' "
         pub_QL05 = pub_QL05 & ";" & frm100110_1.Option1(0).Caption & frm100110_1.txt1(0) 'Add By Sindy 2010/11/3
      End If
   End If
   'Add by Amy 2014/09/25 +申請國家
   If Len(Trim(frm100110_1.txt1(14))) <> 0 Then
       strSQL1 = strSQL1 + " AND TM10>='" & frm100110_1.txt1(14) & "' "
       strSQL2 = strSQL2 + " AND PA09>='" & frm100110_1.txt1(14) & "' "
       StrSQL3 = StrSQL3 + " AND LC15>='" & frm100110_1.txt1(14) & "' "
       strSQL5 = strSQL5 + " AND SP09>='" & frm100110_1.txt1(14) & "' "
   End If
   If Len(Trim(frm100110_1.txt1(15))) <> 0 Then
       strSQL1 = strSQL1 + " AND TM10<='" & frm100110_1.txt1(15) & "' "
       strSQL2 = strSQL2 + " AND PA09<='" & frm100110_1.txt1(15) & "' "
       StrSQL3 = StrSQL3 + " AND LC15<='" & frm100110_1.txt1(15) & "' "
       strSQL5 = strSQL5 + " AND SP09<='" & frm100110_1.txt1(15) & "' "
   End If
   If Len(Trim(frm100110_1.txt1(14))) <> 0 Or Len(Trim(frm100110_1.txt1(15))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(3) & frm100110_1.txt1(14) & "-" & frm100110_1.txt1(15)
   End If
   'end 2014/09/25
   
   'Add/Modify By Cheng 2002/07/01 加抓專利系統, 且機關文號可模糊比對
   ''Modify By Cheng 2002/04/25若已閉卷, 則在本所案號後加"*"號
   '94.3.4 MODIFY BY SONIA 加前後可模糊比對
   '2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
   '2013/8/19 add by sonia 加入法院案號條件 P-099556
   If frm100110_1.Option1(4).Value = True Then
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM,SeColPA
      'strSql = "SELECT '' AS V,CP35 AS 法院案號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP35>' '),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
      'strSql = strSql + " UNION SELECT '' AS V,CP35 AS 法院案號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP35>' '),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
      strSql = "SELECT '' AS V,CP35 AS 法院案號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColTM & _
                  " FROM (select * from CASEPROGRESS where CP35>' '),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
      strSql = strSql + " UNION SELECT '' AS V,CP35 AS 法院案號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColPA & _
                  " FROM (select * from CASEPROGRESS where CP35>' '),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
      'end 2019/11/01
   Else
   '2013/8/19 end
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM,SeColPA
      'strSql = "SELECT '' AS V,CP08 AS 機關文號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP08>' '),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
      'strSql = strSql + " UNION SELECT '' AS V,CP08 AS 機關文號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP08>' '),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
      strSql = "SELECT '' AS V,CP08 AS 機關文號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColTM & _
                   " FROM (select * from CASEPROGRESS where CP08>' '),TRADEMARK,STAFF S1,STAFF S2,casepropertymap WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
      strSql = strSql + " UNION SELECT '' AS V,CP08 AS 機關文號,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 證書審定號,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','存在','N','不存在','') AS 專用權是否存在, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColPA & _
                   " FROM (select * from CASEPROGRESS where CP08>' '),PATENT,STAFF S1,STAFF S2,casepropertymap WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
      'end 2019/11/01
   End If
   '取消lawcase,hirecase,servicepractice
   strSql = strSql + " ORDER BY 收文日,FSort,本所案號 "
   
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
            InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
        End If
        'end 2019/11/01
       
       cmdOK(0).Enabled = True
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
       Me.Enabled = True
       cmdOK(0).Enabled = False
       ShowNoData
       Screen.MousePointer = vbDefault
       '92.04.18 nick
       'Me.Hide
       tmpBol = fnCancelNowFormAndShowParentForm(Me)
       Exit Sub
   End If
   intK = adoRecordset.RecordCount
   Me.grdDataList.Visible = False
   Set grdDataList.Recordset = adoRecordset
   SetDataListWidth
   For ii = 1 To Me.grdDataList.Rows - 1
       Me.grdDataList.TextMatrix(ii, 6) = Me.grdDataList.TextMatrix(ii, 6) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 12), "1")
   Next ii
   Me.grdDataList.Visible = True
   CheckOC
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100110_2 = Nothing
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
