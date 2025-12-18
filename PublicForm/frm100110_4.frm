VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100110_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案件查詢"
   ClientHeight    =   5730
   ClientLeft      =   140
   ClientTop       =   990
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   3
      Left            =   3390
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   4
      Left            =   4935
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   0
      Left            =   6072
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7296
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5076
      Left            =   24
      TabIndex        =   5
      Top             =   624
      Width           =   9252
      _ExtentX        =   16334
      _ExtentY        =   8943
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
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
      Left            =   1170
      TabIndex        =   8
      Top             =   150
      Width           =   2025
   End
   Begin VB.Label Label2 
      Height          =   180
      Left            =   1110
      TabIndex        =   4
      Top             =   432
      Width           =   8196
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "條款:"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Top             =   435
      Width           =   405
   End
End
Attribute VB_Name = "frm100110_4"
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

Dim strSQL1 As String
'Add By Cheng 2002/07/01
Dim strSQL2 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer, intK As Integer
Dim strTemp As Variant, StrTest4 As String, STRTEMP12 As Variant, StrTemp10 As Variant
Dim StrTest2 As String, strTemp1 As String
Dim StrTest5 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String

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
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "收文日"
grdDataList.ColWidth(1) = 810
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "本所案號"
grdDataList.ColWidth(2) = 1450
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 3: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(3) = 0
Else
    grdDataList.ColWidth(3) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(4) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "案件性質"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "審定號"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "承辦人"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "條款"
grdDataList.ColWidth(8) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Modify By Cheng 2002/07/16
'grdDataList.Col = 8: grdDataList.Text = "是否出名"
grdDataList.col = 9: grdDataList.Text = "實際結果"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "對造名稱"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = "對造號數"
grdDataList.ColWidth(11) = 800
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

Private Sub SetDataListWidth2()
'Modified by Lydia 2019/11/01
'grdDataList.Cols = 16
Dim intField As Integer
intField = 22
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "收文日"
grdDataList.ColWidth(1) = 810
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "本所案號"
grdDataList.ColWidth(2) = 1450
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 3: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(3) = 0
Else
    grdDataList.ColWidth(3) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(4) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "案件性質"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "審定號"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "承辦人"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "商品類別"
grdDataList.ColWidth(8) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Modify By Cheng 2002/07/16
'grdDataList.Col = 8: grdDataList.Text = "是否出名"
grdDataList.col = 9: grdDataList.Text = "實際結果"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "對造名稱"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = "對造號數"
grdDataList.ColWidth(11) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = ""
grdDataList.ColWidth(12) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 13: grdDataList.Text = ""
grdDataList.ColWidth(13) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/15
grdDataList.col = 14: grdDataList.Text = "CP09"
grdDataList.ColWidth(14) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/10
grdDataList.col = 15: grdDataList.Text = ""
grdDataList.ColWidth(15) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 16 To intField - 1
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
          grdDataList.col = 2
          If Not IsNull(grdDataList.Text) Then
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
               Screen.MousePointer = vbHourglass
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
Case 2 '結束
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

Sub StrMenu()        '條款
Dim ii As Integer
Dim strTmp As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

ClearQueryLog ("frm100110_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
Me.Enabled = False
Label1(0).Caption = "條款："
Label2.Caption = frm100110_1.txt1(3)

strSQL1 = ""
strSQL2 = ""
'Modify By Sindy 2010/02/10 改寫SQL語法
strSQL1 = strSQL1 + " AND ("
StrTemp10 = Split(frm100110_1.txt1(3), ",")
For i = 0 To UBound(StrTemp10)
'   strSQL1 = strSQL1 + " instr(CP49,'" & StrTemp10(i) & "')>0  "
'   If i <> UBound(StrTemp10) Then
'      strSQL1 = strSQL1 + " OR "
'   End If
   'Modify By Sindy 2012/7/12 可同時查詢新舊條款,若為舊條款時須一併抓取對應的新條款
   If Len(Trim(StrTemp10(i))) = 5 Then '新條款+主張內容
      strSQL2 = strSQL2 + " or instr(CP49,'" & StrTemp10(i) & "')>0"
   ElseIf Len(Trim(StrTemp10(i))) = 4 Then '舊條款+主張內容 或 新條款
      '檢查是否為新條款
      strTmp = "select * from law where lw01='" & StrTemp10(i) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
      If intI = 1 Then
         '新條款
         strSQL2 = strSQL2 + " or instr(CP49,'" & StrTemp10(i) & "')>0"
      Else
         '舊條款+主張內容
         strSQL2 = strSQL2 + " or instr(CP49,'" & StrTemp10(i) & "')>0"
         '檢查是否有對應的新條款
         strTmp = "select * from law where instr(lw04,'" & Left(StrTemp10(i), 3) & "')>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               '為舊條款則須一併抓取對應的新條款+主張內容
               strSQL2 = strSQL2 + " or instr(CP49,'" & RsTemp.Fields("lw01") & Right(StrTemp10(i), 1) & "')>0"
               RsTemp.MoveNext
            Loop
         End If
      End If
   ElseIf Len(Trim(StrTemp10(i))) = 3 Then '舊條款
      strSQL2 = strSQL2 + " or substr(CP49,1,3)='" & StrTemp10(i) & "'"
      strSQL2 = strSQL2 + " or instr(CP49,'," & StrTemp10(i) & "')>0"
      '檢查是否有對應的新條款
      strTmp = "select * from law where instr(lw04,'" & StrTemp10(i) & "')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            '為舊條款則須一併抓取對應的新條款
            strSQL2 = strSQL2 + " or instr(CP49,'" & RsTemp.Fields("lw01") & "')>0"
            RsTemp.MoveNext
         Loop
      End If
   Else
      strSQL2 = strSQL2 + " or instr(CP49,'" & StrTemp10(i) & "')>0"
   End If
   '2012/7/12 End
Next i
'strSQL1 = strSQL1 + " ) "
'Modify By Sindy 2012/7/12
If strSQL2 <> "" Then strSQL2 = Mid(Trim(strSQL2), 4, Len(strSQL2))
strSQL1 = strSQL1 + strSQL2 + ") "
'2012/7/12 End
strSQL2 = strSQL1
pub_QL05 = pub_QL05 & ";" & frm100110_1.Option1(2).Caption & frm100110_1.txt1(3) 'Add By Sindy 2010/11/3
strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "

'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
m_AllSys = GetAllSysKind(, "ALL")
intCufaCnt = 0
'end 2019/11/01

'收文日起
If Len(Trim(frm100110_1.txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
    strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
End If
'收文日迄
If Len(Trim(frm100110_1.txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
    strSQL2 = strSQL2 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
End If
If Len(Trim(frm100110_1.txt1(4))) <> 0 Or Len(Trim(frm100110_1.txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label5 & frm100110_1.txt1(4) & "-" & frm100110_1.txt1(5) 'Add By Sindy 2010/11/3
End If
'案件性質
If Len(Trim(frm100110_1.txt1(6))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
   strSQL2 = strSQL2 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(0) & frm100110_1.txt1(6) 'Add By Sindy 2010/11/3
End If
'申請人起
'Modified by Lydia 2019/11/01 改成申請人1~5
'If Len(Trim(frm100110_1.txt1(7))) <> 0 Then
'    strSQL1 = strSQL1 + " AND TM23>='" & frm100110_1.txt1(7) & "' "
'    strSQL2 = strSQL2 + " AND PA26>='" & frm100110_1.txt1(7) & "' "
'End If
''申請人迄
'If Len(Trim(frm100110_1.txt1(8))) <> 0 Then
'   strSQL1 = strSQL1 & " AND TM23<='" & frm100110_1.txt1(8) & "' "
'   strSQL2 = strSQL2 & " AND PA26<='" & frm100110_1.txt1(8) & "' "
'End If
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
    'end 2019/11/01
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(1) & frm100110_1.txt1(7) & "-" & frm100110_1.txt1(8) 'Add By Sindy 2010/11/3
End If
'代理人起
If Len(Trim(frm100110_1.txt1(9))) <> 0 Then
    strSQL1 = strSQL1 + " AND TM44>='" & frm100110_1.txt1(9) & "' "
    strSQL2 = strSQL2 + " AND PA75>='" & frm100110_1.txt1(9) & "' "
End If
'代理人迄
If Len(Trim(frm100110_1.txt1(10))) <> 0 Then
   strSQL1 = strSQL1 & " AND TM44<='" & frm100110_1.txt1(10) & "' "
   strSQL2 = strSQL2 & " AND PA75<='" & frm100110_1.txt1(10) & "' "
End If
If Len(Trim(frm100110_1.txt1(9))) <> 0 Or Len(Trim(frm100110_1.txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(2) & frm100110_1.txt1(9) & "-" & frm100110_1.txt1(10) 'Add By Sindy 2010/11/3
End If
'Add by Amy 2014/09/25 +申請國家
If Len(Trim(frm100110_1.txt1(14))) <> 0 Then
   strSQL1 = strSQL1 & " AND TM10>='" & frm100110_1.txt1(14) & "' "
   strSQL2 = strSQL2 & " AND PA09>='" & frm100110_1.txt1(14) & "' "
End If
If Len(Trim(frm100110_1.txt1(15))) <> 0 Then
   strSQL1 = strSQL1 & " AND TM10<='" & frm100110_1.txt1(15) & "' "
   strSQL2 = strSQL2 & " AND PA09<='" & frm100110_1.txt1(15) & "' "
End If
If Len(Trim(frm100110_1.txt1(14))) <> 0 Or Len(Trim(frm100110_1.txt1(15))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(3) & frm100110_1.txt1(14) & "-" & frm100110_1.txt1(15)
End If
'end 2014/09/25
'edit by nickc 2005/05/10
'strSQL = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09 FROM CASEPROGRESS,TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'strSQL = strSQL + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09 FROM CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
'strSQL = strSQL + " ORDER BY 收文日,本所案號 "
'Modify By Sindy 2010/02/09
'strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM CASEPROGRESS,TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
'2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
'Modified by Lydia 2019/11/01 +增加欄位 SeColTM,SeColPA
'strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP49>' '),TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP49>' '),PATENT,STAFF,CASEPROPERTYMAP WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColTM & _
            " FROM (select * from CASEPROGRESS where CP49>' '),TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
strSql = strSql + " Union SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,PA22 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,CP49 AS 條款,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColPA & _
            " FROM (select * from CASEPROGRESS where CP49>' '),PATENT,STAFF,CASEPROPERTYMAP WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2

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
    Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 12), "1")
Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True
End Sub

Sub StrMenu2()        '商品類別
Dim ii As Integer
Dim dblRow As Double 'Add By Sindy 2025/9/3

ClearQueryLog ("frm100110_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
Me.Enabled = False
Label1(0).Caption = "商品類別："
Label2.Caption = frm100110_1.txt1(12)

strSQL1 = ""
strSQL2 = ""
strSQL1 = strSQL1 + " AND ( "
StrTemp10 = Split(frm100110_1.txt1(12), ",")
For i = 0 To UBound(StrTemp10)
    strSQL1 = strSQL1 + " (instr(CP80,'" & StrTemp10(i) & "')>0 or instr(tm09,'" & StrTemp10(i) & "')>0)"
    If i <> UBound(StrTemp10) Then
        strSQL1 = strSQL1 + " OR "
    End If
Next i
strSQL1 = strSQL1 + " ) "
strSQL2 = strSQL1
pub_QL05 = pub_QL05 & ";" & frm100110_1.Option1(3).Caption & frm100110_1.txt1(12) 'Add By Sindy 2010/11/3
'Modify By Sindy 2010/02/10 改寫SQL語法
strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "

'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
m_AllSys = GetAllSysKind(, "ALL")
intCufaCnt = 0
'end 2019/11/01

'收文日起
If Len(Trim(frm100110_1.txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
    strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(frm100110_1.txt1(4))) & " "
End If
'收文日迄
If Len(Trim(frm100110_1.txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
    strSQL2 = strSQL2 + " AND CP05<=" & Val(ChangeTStringToWString(frm100110_1.txt1(5))) & " "
End If
If Len(Trim(frm100110_1.txt1(4))) <> 0 Or Len(Trim(frm100110_1.txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label5 & frm100110_1.txt1(4) & "-" & frm100110_1.txt1(5) 'Add By Sindy 2010/11/3
End If
'案件性質
If Len(Trim(frm100110_1.txt1(6))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
   strSQL2 = strSQL2 & " AND CP10 IN (" & GetAddStr(frm100110_1.txt1(6)) & ") "
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(0) & frm100110_1.txt1(6) 'Add By Sindy 2010/11/3
End If
'申請人起
'Modified by Lydia 2019/11/01 改成申請人1~5
'If Len(Trim(frm100110_1.txt1(7))) <> 0 Then
'    strSQL1 = strSQL1 + " AND TM23>='" & frm100110_1.txt1(7) & "' "
'    strSQL2 = strSQL2 + " AND PA26>='" & frm100110_1.txt1(7) & "' "
'End If
''申請人迄
'If Len(Trim(frm100110_1.txt1(8))) <> 0 Then
'   strSQL1 = strSQL1 & " AND TM23<='" & frm100110_1.txt1(8) & "' "
'   strSQL2 = strSQL2 & " AND PA26<='" & frm100110_1.txt1(8) & "' "
'End If
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
    'end 2019/11/01
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(1) & frm100110_1.txt1(7) & "-" & frm100110_1.txt1(8) 'Add By Sindy 2010/11/3
End If
'代理人起
If Len(Trim(frm100110_1.txt1(9))) <> 0 Then
    strSQL1 = strSQL1 + " AND TM44>='" & frm100110_1.txt1(9) & "' "
    strSQL2 = strSQL2 + " AND PA75>='" & frm100110_1.txt1(9) & "' "
End If
'代理人迄
If Len(Trim(frm100110_1.txt1(10))) <> 0 Then
   strSQL1 = strSQL1 & " AND TM44<='" & frm100110_1.txt1(10) & "' "
   strSQL2 = strSQL2 & " AND PA75<='" & frm100110_1.txt1(10) & "' "
End If
If Len(Trim(frm100110_1.txt1(9))) <> 0 Or Len(Trim(frm100110_1.txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(2) & frm100110_1.txt1(9) & "-" & frm100110_1.txt1(10) 'Add By Sindy 2010/11/3
End If
'Add by Amy 2014/09/25 +申請國家
If Len(Trim(frm100110_1.txt1(14))) <> 0 Then
   strSQL1 = strSQL1 & " AND TM10>='" & frm100110_1.txt1(14) & "' "
   strSQL2 = strSQL2 & " AND PA09>='" & frm100110_1.txt1(14) & "' "
End If
If Len(Trim(frm100110_1.txt1(15))) <> 0 Then
   strSQL1 = strSQL1 & " AND TM10<='" & frm100110_1.txt1(15) & "' "
   strSQL2 = strSQL2 & " AND PA09<='" & frm100110_1.txt1(15) & "' "
End If
If Len(Trim(frm100110_1.txt1(14))) <> 0 Or Len(Trim(frm100110_1.txt1(15))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100110_1.Label6(3) & frm100110_1.txt1(14) & "-" & frm100110_1.txt1(15)
End If
'end 2014/09/25

'edit by nickc 2005/05/10
'strSQL = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,'' AS 商品類別,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數,Tm09,cp80, CP09 FROM CASEPROGRESS,TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'strSQL = strSQL + " ORDER BY 收文日,本所案號 "
'2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
'Modified by Lydia 2019/11/01 +增加欄位 SeColTM
'strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,TM09 AS 商品類別,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數,Tm09,cp80, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort FROM (select * from CASEPROGRESS where CP80>' '),TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,TM15 AS 專利審定號,NVL(ST02,CP14) AS 承辦人,TM09 AS 商品類別,DECODE(CP24,'1','勝','2','敗','') AS 實際結果,NVL(CP40,NVL(CP41,CP42)) AS 對造名稱,CP36 AS 對造號數,Tm09,cp80, CP09,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as FSort" & SeColTM & _
            " FROM (select * from CASEPROGRESS where CP80>' '),TRADEMARK,STAFF,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
            
strSql = strSql + " ORDER BY 收文日,FSort,本所案號 "
CheckOC
SetDataListWidth2
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
SetDataListWidth2
For ii = 1 To Me.grdDataList.Rows - 1
    Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 12), "1")
Next ii
Me.grdDataList.Visible = True
Dim intItem As Integer
Dim arrTmp As Variant
Dim ArrTmpIndex As Integer
Dim BolCheckData As Boolean
Dim StrTmpString As String
Dim StrTmpData As String
StrTmpString = frm100110_1.txt1(12)
If Mid(StrTmpString, 1, 1) <> "," Then
    StrTmpString = "," & StrTmpString
End If
If Right(StrTmpString, 1) <> "," Then
    StrTmpString = StrTmpString & ","
End If
With grdDataList
    .Visible = False
    For intItem = .Rows - 1 To 1 Step -1
        BolCheckData = False
        .row = intItem
        .col = 11
        If Trim(.Text) <> "" Then
            arrTmp = Split(.Text, ",")
            For ArrTmpIndex = 0 To UBound(arrTmp)
                If InStr(1, StrTmpString, "," & Trim(arrTmp(ArrTmpIndex)) & ",") > 0 Then
                    StrTmpData = .Text
                    .col = 7
                    .Text = StrTmpData
                    BolCheckData = True
                    Exit For
                End If
            Next ArrTmpIndex
        End If
        .col = 12
        If Trim(.Text) <> "" And BolCheckData = False Then
            arrTmp = Split(.Text, ",")
            For ArrTmpIndex = 0 To UBound(arrTmp)
                If InStr(1, StrTmpString, "," & Trim(arrTmp(ArrTmpIndex)) & ",") > 0 Then
                    StrTmpData = .Text
                    .col = 7
                    .Text = StrTmpData
                    BolCheckData = True
                    Exit For
                End If
            Next ArrTmpIndex
        End If
        'Modified by Lydia 2019/11/01 無法移除最後一個非固定列
        'If BolCheckData = False Then
        If BolCheckData = False And intItem > 1 Then
            .RemoveItem .row
        End If
    Next intItem
    .Visible = True
End With
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100110_4 = Nothing
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
