VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100107_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文未發文查詢"
   ClientHeight    =   5720
   ClientLeft      =   100
   ClientTop       =   980
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdOK 
      Caption         =   "核駁分析等未發文案件(&Word)"
      CausesValidation=   0   'False
      Height          =   330
      Index           =   5
      Left            =   6168
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   5352
      Visible         =   0   'False
      Width           =   3048
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "未發文原因"
      Height          =   400
      Index           =   4
      Left            =   3450
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   10
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8532
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   0
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   6084
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7308
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4848
      Left            =   0
      TabIndex        =   0
      Top             =   444
      Width           =   9228
      _ExtentX        =   16281
      _ExtentY        =   8555
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
   Begin VB.CheckBox Check1 
      Caption         =   "排除國外部收文案件"
      Height          =   252
      Left            =   96
      TabIndex        =   8
      Top             =   5376
      Visible         =   0   'False
      Width           =   2124
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷△介紹案源"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   2880
   End
End
Attribute VB_Name = "frm100107_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, strTemp As String, strTemp1 As String, i As Integer, j As Integer
Dim StrTag As String, intK As Integer, strTemp3 As Variant, StrTemp7 As Variant
Dim ArrTmpNoData As String, arrTmp As Variant
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件


Private Sub SetDataListWidth()
'Added by Lydia 2019/11/01
Dim intField As Integer
intField = 20 'Modified by Morgan 2024/3/18 +CP10 19->20
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "智權人員"
grdDataList.ColWidth(1) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "收文日"
grdDataList.ColWidth(2) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "本所案號"
grdDataList.ColWidth(3) = 1500
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
grdDataList.ColWidth(5) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "案件性質"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "申請人"
grdDataList.ColWidth(7) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "承辦人"
grdDataList.ColWidth(8) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "取消收文日"
'grdDataList.ColWidth(8) = 1500
grdDataList.ColWidth(9) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "收款情形"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2002/01/24
grdDataList.col = 11: grdDataList.Text = "申請國家"
grdDataList.ColWidth(11) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = "CP09"
grdDataList.ColWidth(12) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 13 To intField - 1
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
Case 2
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
'Add by Amy 2015/09/21
Case 4 '發文延誤原因
    Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
            Next j
            If Not IsNull(grdDataList.TextMatrix(i, 12)) Then
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Call frm090638_1.SetParent(Me)
                frm090638_1.BFormPeople = 3
                frm090638_1.m_NC01 = grdDataList.TextMatrix(i, 12)
                frm090638_1.Show
                If frm090638_1.QueryData = False Then
                    ShowNoData
                    frm090638_1.cmdState = 2
                    frm090638_1.PubShowNextData
                End If
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Sub
            End If
        End If
     Next i
     Me.Enabled = True
'Added by Morgan 2024/3/18
Case 5 'CFP核駁分析等未發文案件(&Word)
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   ExportWord
   Me.Enabled = True
   Screen.MousePointer = vbDefault
Case Else
End Select
End Sub


Private Sub Check1_Click()
   If Check1.Visible Then StrMenu
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

'Added by Morgan 2024/3/18
If Left(Pub_StrUserSt03, 2) = "P1" Then
   cmdOK(5).Visible = True
   Check1.Visible = True
   Check1.Value = vbChecked
Else
   grdDataList.Height = grdDataList.Height + cmdOK(5).Height + 50
   cmdOK(5).Visible = False
   Check1.Value = vbUnchecked
   Check1.Visible = False
End If
'end 2024/3/18

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100107_2 = Nothing
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

Sub StrMenu()
Dim dblRow As Double 'Add By Sindy 2025/9/3

' nickc 91.07.31
ArrTmpNoData = ""

'Add By Cheng 2002/01/23
Dim strSQL11 As String
Dim strSQL21 As String
Dim strSQL31 As String
Dim strSQL41 As String
Dim strSQL51 As String
Dim strLosQ As String 'Added by Lydia 2021/03/23 法律所介紹案源的介紹人員(智權人員)

Me.Enabled = False
grdDataList.Visible = False
   
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""

'Added by Morgan 2025/6/26 內專預設排除國外部案件--柏翰
If Check1.Value = vbChecked Then
   strSQL1 = " and CP12 not like 'F%'"
End If
'end 2025/6/26

If Len(Trim(frm100107_1.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05>=" & Val(ChangeTStringToWString(frm100107_1.txt1(0))) & " "
End If
If Len(Trim(frm100107_1.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(frm100107_1.txt1(1))) & " "
End If
If Len(Trim(frm100107_1.txt1(0))) <> 0 Or Len(Trim(frm100107_1.txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(0) & frm100107_1.txt1(0) & "-" & frm100107_1.txt1(1) 'Add By Sindy 2010/01/22
End If
If Len(Trim(frm100107_1.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND SUBSTR(CP09,1,1) IN (" & GetAddStr(frm100107_1.txt1(2)) & ") "
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(6) & frm100107_1.txt1(2) & frm100107_1.Label2 'Add By Sindy 2010/01/22
End If
If Len(Trim(frm100107_1.txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP14='" & frm100107_1.txt1(3) & "' "
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(5) & frm100107_1.txt1(3) & frm100107_1.lbl1(0)  'Add By Sindy 2010/01/22
End If
If Len(Trim(frm100107_1.txt1(4))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP13='" & frm100107_1.txt1(4) & "' "
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(4) & frm100107_1.txt1(4) & frm100107_1.lbl1(1)   'Add By Sindy 2010/01/22
End If
If Len(Trim(frm100107_1.txt1(6))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP10>='" & frm100107_1.txt1(6) & "' "
End If
If Len(Trim(frm100107_1.txt1(7))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP10<='" & frm100107_1.txt1(7) & "' "
End If
If Len(Trim(frm100107_1.txt1(6))) <> 0 Or Len(Trim(frm100107_1.txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(2) & frm100107_1.txt1(6) & "-" & frm100107_1.txt1(7)  'Add By Sindy 2010/01/22
End If
'Add By Cheng 2002/01/23
'申請人國籍
If Len(Trim(frm100107_1.txt1(10))) <> 0 Then
   strSQL1 = strSQL1 & " AND CU10>='" & frm100107_1.txt1(10) & "' "
End If
If Len(Trim(frm100107_1.txt1(11))) <> 0 Then
   strSQL1 = strSQL1 & " AND CU10<='" & frm100107_1.txt1(11) & "z' "
End If
If Len(Trim(frm100107_1.txt1(10))) <> 0 Or Len(Trim(frm100107_1.txt1(11))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(7) & frm100107_1.txt1(10) & "-" & frm100107_1.txt1(11)   'Add By Sindy 2010/01/22
End If
'Add By Cheng 2003/06/02
'是否含已取消收文資料
If Len(Trim(frm100107_1.txt1(12))) = 0 Then
   strSQL1 = strSQL1 & " AND CP57 Is Null "
   pub_QL05 = pub_QL05 & ";含已取消收文資料" 'Add By Sindy 2010/01/22
End If
strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
'Added by Lydia 2021/03/23 判斷條件有法務案+智權人員
strLosQ = ""
If frm100107_1.txt1(4) <> "" Then
   strExc(0) = GetST15(frm100107_1.txt1(4))
   If Left(strExc(0), 1) <> "L" And strExc(0) <> "P31" Then  '參考frm010007客戶所屬智權人員的部門判斷 =>判斷法務案是否有案源 (非法律所的客戶)
        strExc(0) = Replace(SQLGrpStr(IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(frm100107_1.txt1(5))), 3), "'ACS',", "")
        If InStr(strExc(0), "L") > 0 Then
           strLosQ = Replace(StrSQL3, " AND CP13='" & frm100107_1.txt1(4) & "' ", " AND INSTR(LOS04,'" & frm100107_1.txt1(4) & "') > 0 ")
           strLosQ = strLosQ & " AND CP01 IN (" & strExc(0) & ")"
        End If
   End If
End If
'end 2021/03/23

If Len(Trim(frm100107_1.txt1(5))) <> 0 Then
   'Modify By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(frm100107_1.txt1(5), 1) & ") "
'   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(frm100107_1.txt1(5), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(frm100107_1.txt1(5), 3) & ") "
'   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(frm100107_1.txt1(5), 4) & ") "
'   StrSQL5 = StrSQL5 & " AND CP01 IN (" & SQLGrpStr(frm100107_1.txt1(5), 5) & ") "
   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(frm100107_1.txt1(5))), 1) & ") "
   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(frm100107_1.txt1(5))), 2) & ") "
   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(frm100107_1.txt1(5))), 3) & ") "
   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(frm100107_1.txt1(5))), 4) & ") "
   strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(frm100107_1.txt1(5))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm100107_1.Label1(3), 5) & frm100107_1.txt1(5) 'Add By Sindy 2010/01/22
End If
'Add By Cheng 2002/01/23
'申請國家
strSQL11 = "": strSQL21 = "": strSQL31 = "": strSQL41 = "": strSQL51 = ""
If Len(Trim(frm100107_1.txt1(8))) <> 0 Then
   strSQL11 = " And PA09 >= '" & frm100107_1.txt1(8).Text & "' "
   strSQL21 = " And TM10 >= '" & frm100107_1.txt1(8).Text & "' "
   strSQL31 = " And LC15 >= '" & frm100107_1.txt1(8).Text & "' "
   strSQL41 = " And '000' >= '" & frm100107_1.txt1(8).Text & "' "
   strSQL51 = " And SP09 >= '" & frm100107_1.txt1(8).Text & "' "
End If
If Len(Trim(frm100107_1.txt1(9))) <> 0 Then
   strSQL11 = strSQL11 & " And PA09 <= '" & frm100107_1.txt1(9).Text & "' "
   strSQL21 = strSQL21 & " And TM10 <= '" & frm100107_1.txt1(9).Text & "' "
   strSQL31 = strSQL31 & " And LC15 <= '" & frm100107_1.txt1(9).Text & "' "
   strSQL41 = strSQL41 & " And '000' <= '" & frm100107_1.txt1(9).Text & "' "
   strSQL51 = strSQL51 & " And SP09 <= '" & frm100107_1.txt1(9).Text & "' "
End If
If Len(Trim(frm100107_1.txt1(8))) <> 0 Or Len(Trim(frm100107_1.txt1(9))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100107_1.Label1(1) & frm100107_1.txt1(8) & "-" & frm100107_1.txt1(9) 'Add By Sindy 2010/01/22
End If
If frm100107_1.Check1.Value = 1 Then
   strSQL1 = strSQL1 & " AND CP60 is not null  "
   strSQL2 = strSQL2 & " AND CP60 is not null  "
   StrSQL3 = StrSQL3 & " AND CP60 is not null  "
   StrSQL4 = StrSQL4 & " AND CP60 is not null  "
   strSQL5 = strSQL5 & " AND CP60 is not null  "
   pub_QL05 = pub_QL05 & ";已收款未發文" 'Add By Sindy 2010/01/22
End If

'Added by Lydia 2019/11/01 利益衝突案件
m_AllSys = IIf(frm100107_1.txt1(5).Text <> "ALL", frm100107_1.txt1(5).Text, GetAllSysKind(, frm100107_1.txt1(5).Text))
intCufaCnt = 0
'end 2019/11/01

'Modify By Cheng 2002/01/23
'加搜尋條件--申請國家
'strSQL = "SELECT '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,CP09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (TM29<>'Y' OR TM29 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("TM23", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL " & strSQL2
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("PA26", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL " & strSQL1
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03 AND CP04=SP04(+) AND (SP15<>'Y' OR SP15 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("SP08", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL  " & StrSQL5
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),LC11) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,CP09 FROM CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (LC08<>'Y' OR LC08 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL  " & StrSQL3
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(CPM03,CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),HC05) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,CP09 FROM CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND (HC09<>'Y' OR HC09 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL " & StrSQL4
'2010/9/10 MODIFY BY SONIA 改百年日期排序問題
'strSQL = "SELECT '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (TM29<>'Y' OR TM29 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("TM23", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And TM10=NA01(+) " & strSQL21 & strSQL2
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("PA26", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And PA09=NA01(+) " & strSQL11 & strSQL1
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03 AND CP04=SP04(+) AND (SP15<>'Y' OR SP15 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("SP08", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And SP09=NA01(+) " & strSQL51 & strSQL5
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),LC11) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (LC08<>'Y' OR LC08 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And LC15=NA01(+) " & strSQL31 & StrSQL3
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(CPM03,CP10)                          AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),HC05) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND (HC09<>'Y' OR HC09 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL AND CP10<>'0' And NA01='000' " & strSQL41 & StrSQL4

'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'Modified by Lydia 2019/11/01 +增加欄位: 申請人1~5(cust01~cust05), FC代理人(fcno)
'strSql = "SELECT '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),TM23) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
'         "WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (TM29<>'Y' OR TM29 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("TM23", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And TM10=NA01(+) " & strSQL21 & strSQL2
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
'         "WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("PA26", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And PA09=NA01(+) " & strSQL11 & strSQL1
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),SP08) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
'         "WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03 AND CP04=SP04(+) AND (SP15<>'Y' OR SP15 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("SP08", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And SP09=NA01(+) " & strSQL51 & strSQL5
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
'         "WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (LC08<>'Y' OR LC08 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And LC15=NA01(+) " & strSQL31 & StrSQL3
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(CPM03,CP10)                          AS 案件性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),HC05) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
'         "WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND (HC09<>'Y' OR HC09 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL AND CP10<>'0' And NA01='000' " & strSQL41 & StrSQL4
'2010/9/10 END
'Modified by Morgan 2024/3/18 +CP10
strSql = "SELECT '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質" & _
          ",NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),TM23) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09" & _
          ",tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno,cp10" & _
          " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
         "WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (TM29<>'Y' OR TM29 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("TM23", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And TM10=NA01(+) " & strSQL21 & strSQL2
strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質" & _
         ",NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09" & _
         ",pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno,cp10" & _
         " FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
         "WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (PA57<>'Y' OR PA57 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("PA26", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And PA09=NA01(+) " & strSQL11 & strSQL1
strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) AS 案件性質" & _
         ",NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),SP08) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09" & _
         ",sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno,cp10" & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
         "WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03 AND CP04=SP04(+) AND (SP15<>'Y' OR SP15 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("SP08", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And SP09=NA01(+) " & strSQL51 & strSQL5
strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質" & _
          ",NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09" & _
         ",lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno,cp10" & _
         " FROM CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
         "WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (LC08<>'Y' OR LC08 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And LC15=NA01(+) " & strSQL31 & StrSQL3
'Added by Lydia 2021/03/23 判斷條件有法務案+智權人員(非法律所)：額外抓法律所介紹案源的資料(介紹人員=智權人員)
If strLosQ <> "" Then
    strSql = strSql & " union all select '' AS V,GETSTAFFNAMELIST(LOS04)||'△' AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質" & _
              ",NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11) AS 申請人, Nvl(S1.St02,Cp13) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09" & _
             ",lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno,cp10" & _
             " FROM CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,STAFF S1,CUSTOMER,Nation,LawOfficeSource " & _
             "WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (LC08<>'Y' OR LC08 IS NULL) AND CP13=S1.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And LC15=NA01(+) AND CP162=LOS15(+) " & strSQL31 & strLosQ
End If
'end 2021/03/23
strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,NVL(CPM03,CP10) AS 案件性質" & _
         ",NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),HC05) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) As 收款情形,nvl(NA03,NA04) As 申請國家,CP09" & _
         ",hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno,cp10" & _
         " FROM CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation " & _
         "WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND (HC09<>'Y' OR HC09 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL AND CP10<>'0' And NA01='000' " & strSQL41 & StrSQL4
'end 2019/11/01

'2006/8/23 MODIFY BY SONIA 剔除顧問聘任
'strSql = strSql & " union all select '' AS V,NVL(S1.ST02,CP13) AS 智權人員," & SQLDate("CP05") & " AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(CPM03,CP10)                          AS 案件性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),HC05) AS 申請人,NVL(S2.ST02,CP14) AS 承辦人," & SQLDate("CP57") & " AS 取消收文日,CP60 AS 收款情形,nvl(NA03,NA04) As 申請國家,CP09 FROM CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER,Nation WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND (HC09<>'Y' OR HC09 IS NULL) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP27 IS NULL And NA01='000' " & strSQL41 & StrSQL4

'
If frm100107_1.Option1(0).Value = True Then
   strSql = strSql & " ORDER BY 收文日,本所案號 "
   pub_QL05 = pub_QL05 & ";資料順序:" & frm100107_1.Option1(0).Caption 'Add By Sindy 2010/01/22
Else
   If frm100107_1.Option1(1).Value = True Then
        strSql = strSql & " ORDER BY 本所案號,收文日"
        pub_QL05 = pub_QL05 & ";資料順序:" & frm100107_1.Option1(1).Caption 'Add By Sindy 2010/01/22
   End If
End If
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
        InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
     Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
     End If
    'end 2019/11/01
    
Else
    InsertQueryLog (0) 'Add By Sindy 2010/01/22
JumpToNoData:   'Added by Lydia 2019/11/01
    Me.Enabled = True
    cmdOK(0).Enabled = False
    cmdOK(1).Enabled = False
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'me.hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If

Set grdDataList.Recordset = adoRecordset
intK = adoRecordset.RecordCount
CheckOC
For i = 1 To grdDataList.Rows - 1
    'Add By Cheng 2003/08/15
    Me.grdDataList.TextMatrix(i, 6) = Me.grdDataList.TextMatrix(i, 6) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 12), "1")
    grdDataList.row = i

    Dim IntTemp1 As Long
    Dim IntTemp2 As Long
    IntTemp1 = 0
    IntTemp2 = 0
    grdDataList.col = 10
    If Not IsNull(grdDataList.Text) Then
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
                      '91.07.31 nick
                      If frm100107_1.Check1.Value = 1 Then
                          ArrTmpNoData = ArrTmpNoData & Trim(grdDataList.row) & ","
                      End If
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
Next i
'91.07.31 nickc
Me.Enabled = True
grdDataList.Visible = False
Me.Enabled = False
Dim ArrTmpIndex As Integer
If Trim(ArrTmpNoData) <> "" Then
    arrTmp = Split(ArrTmpNoData, ",")
    For ArrTmpIndex = UBound(arrTmp) To 0 Step -1
        If Trim(arrTmp(ArrTmpIndex)) <> "" Then
            If arrTmp(ArrTmpIndex) = 1 And grdDataList.Rows = 2 Then
                grdDataList.Clear
                SetDataListWidth
                Me.Enabled = True
                cmdOK(0).Enabled = False
                cmdOK(1).Enabled = False
                grdDataList.Visible = True
                ShowNoData
                Screen.MousePointer = vbDefault
                '92.04.18 nick
                'Me.Hide
                tmpBol = fnCancelNowFormAndShowParentForm(Me)
                Exit Sub
            Else
                grdDataList.Visible = False
                grdDataList.RemoveItem arrTmp(ArrTmpIndex)
            End If
        End If
    Next ArrTmpIndex
End If
grdDataList.Visible = True
Me.Enabled = True
End Sub


'Added by Morgan 2024/3/18
Private Sub ExportWord()
   Dim iResumeCnt As Integer
   Dim stTmp As String
   Dim oTable As Word.Table
   Dim iCol As Integer, iRow As Integer, ii As Integer
   
On Error GoTo ErrHnd
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   
   If g_WordAp.Visible And g_WordAp.Documents.Count > 0 Then
      If MsgBox("輸出資料是否附加在目前的文件後面？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         g_WordAp.Selection.EndKey Unit:=wdStory
         g_WordAp.Selection.TypeParagraph
      Else
         g_WordAp.Documents.add
      End If
   Else
      g_WordAp.Documents.add
   End If
   
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Visible = True
      
      '邊框設單線
      With .Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
      End With
      '橫印
      .Selection.PageSetup.Orientation = wdOrientLandscape
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
      
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 12
            
      stTmp = "1002核駁分析、1006最終核駁、1201通知修正、1206選取、1209檢索報告，7天以上還未發文："
      .Selection.TypeText Text:=stTmp
      .Selection.TypeParagraph
      
      '新增表格
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=6)
      
      'oTable.AllowAutoFit = True
      .Selection.SelectRow
      With .Selection.Borders(wdBorderTop)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderLeft)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderBottom)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderRight)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderHorizontal)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      With .Selection.Borders(wdBorderVertical)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      
      ii = 1
      oTable.Columns(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(1).Select
      .Selection.TypeText "工程師"
      
      oTable.Columns(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(2).Select
      .Selection.TypeText "案號"
      
      oTable.Columns(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(3).Select
      .Selection.TypeText "收文日"
      
      oTable.Columns(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(4).Select
      .Selection.TypeText "總承辦天數 "
      
      oTable.Columns(5).SetWidth ColumnWidth:=.CentimetersToPoints(5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(5).Select
      .Selection.TypeText "備註"
      
      oTable.Rows(ii).Cells(6).Select
      .Selection.TypeText "承諾會稿日 "
      
      
      For iRow = 1 To grdDataList.Rows - 1
         'Modified by Morgan 2025/6/26 抓畫面的所有資料，不必限制系統及性質--柏翰
         'If Left(grdDataList.TextMatrix(iRow, 3), 4) = "CFP-" And Left(grdDataList.TextMatrix(iRow, 12), 1) = "C" And InStr("1002,1006,1201,1206,1209", grdDataList.TextMatrix(iRow, 19)) > 0 Then
            ii = ii + 1
            oTable.Rows.add
            oTable.Rows(ii).Cells(1).Select
            .Selection.TypeText grdDataList.TextMatrix(iRow, 8) '"工程師"
            oTable.Rows(ii).Cells(2).Select
            .Selection.TypeText Replace(grdDataList.TextMatrix(iRow, 3), "-0-00", "") '"案號"
            oTable.Rows(ii).Cells(3).Select
            .Selection.TypeText grdDataList.TextMatrix(iRow, 2) '"收文日"
            oTable.Rows(ii).Cells(4).Select
            .Selection.TypeText GetWorkDay(strSrvDate(1), DBDATE(grdDataList.TextMatrix(iRow, 2))) '"總承辦天數"
         'End If
      Next
      
      .Activate
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤" & Err.Number & " : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Sub
