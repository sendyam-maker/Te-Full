VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100114_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件性質統計"
   ClientHeight    =   4440
   ClientLeft      =   2100
   ClientTop       =   2505
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9540
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   315
      Index           =   2
      Left            =   8400
      TabIndex        =   3
      Top             =   45
      Width           =   1080
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   315
      Index           =   1
      Left            =   7200
      TabIndex        =   2
      Top             =   45
      Width           =   1080
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "明細"
      Height          =   315
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   45
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3960
      Left            =   30
      TabIndex        =   0
      Top             =   435
      Visible         =   0   'False
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   6985
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS:指定國家之內部收文不統計, 但明細資料仍顯示"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3915
   End
End
Attribute VB_Name = "frm100114_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/06 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'91.08.01   nick 第二階段 編號 805 新撰寫
Option Explicit
Dim BolFrom100114 As Boolean
Dim tmpFa As String
Dim bolHaveData As Boolean
Dim s As Integer
Dim tmpCp10  As String
Dim tmpCP01 As String
'add by nickc 2007/12/21
Dim tmpCFFC As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件-管制
'Memo by Lydia 2019/12/26 利益衝突案件：從外層SQL控制，改成逐案比對。
Dim intCufaCnt As Integer '限閱案件X件
Dim m_AllSys As String

'92.04.16 nick
Public Sub PubShowNextData()
Dim j As Integer
Dim i As Integer
Select Case cmdState
Case 0
    cmdState = -1
    tmpCp10 = ""
    bolHaveData = False
    Screen.MousePointer = vbHourglass
    grdDataList.Visible = False
    For j = 1 To grdDataList.Rows - 1
        grdDataList.row = j
        grdDataList.col = 0
        'add by nickc 2007/12/21
        tmpCFFC = ""
        If grdDataList.CellBackColor = &HFFC0C0 Then
            grdDataList.col = 0
            tmpCP01 = grdDataList.Text
            'edit by nickc 2007/12/21 加欄位，修正
            'grdDataList.col = 3
            grdDataList.col = IIf(Mid(UCase(tmpFa), 1, 1) = "Y", 4, 3)
            tmpCp10 = grdDataList.Text
            'add by nickc 2007/12/21
            If Mid(UCase(tmpFa), 1, 1) = "Y" Then
                grdDataList.col = 1
                tmpCFFC = grdDataList.Text
            End If
            bolHaveData = True
            Exit For
        End If
    Next j
    grdDataList.Visible = True
    If bolHaveData = False Then
        Screen.MousePointer = vbDefault
        s = MsgBox("請選擇一筆才能顯示明細！", , "警告！")
        Exit Sub
    End If
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    frm100114_4.Show
    'edit by nickc 2007/12/21
    'frm100114_4.StrMenu tmpFa, BolFrom100114, tmpCP10, tmpCP01
    frm100114_4.StrMenu tmpFa, BolFrom100114, tmpCp10, tmpCP01, tmpCFFC
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
Case 1
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2
     fnCloseAllFrm100
Case Else
End Select
End Sub


Private Sub cmdOK_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Dim j As Integer
Dim i As Integer
Select Case Index
Case 0
    tmpCp10 = ""
    bolHaveData = False
    Screen.MousePointer = vbHourglass
    grdDataList.Visible = False
    For j = 1 To grdDataList.Rows - 1
        grdDataList.row = j
        grdDataList.col = 0
        If grdDataList.CellBackColor = &HFFC0C0 Then
            grdDataList.col = 0
            tmpCP01 = grdDataList.Text
            grdDataList.col = 3
            tmpCp10 = grdDataList.Text
            bolHaveData = True
            Exit For
        End If
    Next j
    grdDataList.Visible = True
    If bolHaveData = False Then
        Screen.MousePointer = vbDefault
        s = MsgBox("請選擇一筆才能顯示明細！", , "警告！")
        Exit Sub
    End If
    frm100114_4.Show
    frm100114_4.StrMenu tmpFa, BolFrom100114, tmpCp10, tmpCP01
    Screen.MousePointer = vbDefault
    Me.Hide
    Do
    DoEvents
    If bolToEndByNick = True Then Unload Me: Exit Sub
    Loop Until Not frm100114_4.Visible
    Unload frm100114_4
    Me.Show
Case 1
     Me.Hide
Case 2
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu(oStrFA As String, From100114 As Boolean)
tmpFa = oStrFA
BolFrom100114 = From100114
'edit by nickc 2007/12/21
'If BolFrom100114 = True Then
If Mid(UCase(tmpFa), 1, 1) = "Y" Then
'從代理人進來
    BolFrom100114 = True
    StrMenu1 oStrFA
Else
'從申請人進來
    BolFrom100114 = False
    StrMenu2 oStrFA
End If
End Sub

Private Sub SetDataListWidth()
With grdDataList
.Visible = False
'edit by nickc 2007/12/21 區分代理人及申請人
If BolFrom100114 = True Then
    .Cols = 5
Else
    .Cols = 4
End If
.row = 0
.col = 0: .Text = "系統別"
.ColWidth(0) = 800
.CellAlignment = flexAlignCenterCenter
'edit by nickc 2007/12/21 區分代理人及申請人
If BolFrom100114 = True Then
    .col = 1: .Text = "區分"
    .ColWidth(1) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "案件性質"
    .ColWidth(2) = 3000
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "小計"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = ""
    .ColWidth(4) = 0
    .CellAlignment = flexAlignCenterCenter
Else
    .col = 1: .Text = "案件性質"
    .ColWidth(1) = 3000
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "小計"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = ""
    .ColWidth(3) = 0
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End If
.Visible = True
End With
End Sub

'從代理人來
Sub StrMenu1(oStrFA As String)
Me.Enabled = False
'顯示表單上頭資料
'開始搜尋
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
Dim StrSQL6 As String
'Add By Cheng 2002/12/13
Dim StrSQL7 As String

Dim frm As Form
Dim IsFrom100114 As Boolean
IsFrom100114 = False
For Each frm In Forms
    Select Case UCase(frm.Name)
    Case "FRM100114_1"
            IsFrom100114 = True
    End Select
Next

strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
'2008/10/23 modify by sonia指定國家之B類不統計但明細資料仍顯示
'strSQL5 = ""
strSQL5 = " and not (CP04<>'00' and substr(CP09,1,1)='B') "
'2008/10/23 end
StrSQL6 = ""
StrSQL7 = ""
If IsFrom100114 = True Then
            If Len(Trim(frm100114_1.txt1(2))) <> 0 Then
               strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") "
               strSQL2 = strSQL2 & " and PA01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") "
               StrSQL3 = StrSQL3 & " and SP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") "
               StrSQL4 = StrSQL4 & " and LC01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") "
               StrSQL7 = StrSQL7 & " and HC01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 4) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 4) & ") "
            End If
            If Len(Trim(frm100114_1.txt1(8))) <> 0 Then           '檢查申請國家
               strSQL1 = strSQL1 + " AND TM10='" & frm100114_1.txt1(8) & "' "
               strSQL2 = strSQL2 + " AND PA09='" & frm100114_1.txt1(8) & "' "
               StrSQL3 = StrSQL3 + " AND SP09='" & frm100114_1.txt1(8) & "' "
               StrSQL4 = StrSQL4 + " AND LC15='" & frm100114_1.txt1(8) & "' "
            End If
            If Len(Trim(frm100114_1.txt1(6))) <> 0 Then            '檢查案件性質
                strSQL1 = strSQL1 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                strSQL2 = strSQL2 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                StrSQL3 = StrSQL3 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                StrSQL4 = StrSQL4 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                StrSQL7 = StrSQL7 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
            End If
            If Len(Trim(frm100114_1.txt1(7))) <> 0 Then
                strSQL1 = strSQL1 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                strSQL2 = strSQL2 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                StrSQL3 = StrSQL3 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                StrSQL4 = StrSQL4 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                StrSQL7 = StrSQL7 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
            End If
            If frm100114_1.txt1(3) = "1" Then        '收文
               If Len(frm100114_1.txt1(4)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
               End If
               If Len(frm100114_1.txt1(5)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
               Else
                  If Len(frm100114_1.txt1(4).Text) > 0 Then
                     strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                  End If
               End If
            Else
               If Len(frm100114_1.txt1(4)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
               End If
               If Len(frm100114_1.txt1(5)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
               Else
                  If Len(frm100114_1.txt1(4).Text) > 0 Then
                     strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                  End If
               End If
            End If
Else
        '系統類別
        If Len(Trim(frm100102_1.Text3.Text)) <> 0 Then
           strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 2) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 2) & ") "
           strSQL2 = strSQL2 & " and PA01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 1) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 1) & ") "
           StrSQL3 = StrSQL3 & " and SP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 5) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 5) & ") "
           StrSQL4 = StrSQL4 & " and LC01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 3) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 3) & ") "
           'Add By Cheng 2002/12/13
           StrSQL7 = StrSQL7 & " and hc01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 4) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 4) & ") "
        End If
        '2010/3/11 add by sonia
        If Len(Trim(frm100102_1.txtCountry(0))) <> 0 Then           '檢查申請國家
           strSQL1 = strSQL1 + " AND TM10>='" & frm100102_1.txtCountry(0) & "' "
           strSQL2 = strSQL2 + " AND PA09>='" & frm100102_1.txtCountry(0) & "' "
           StrSQL3 = StrSQL3 + " AND SP09>='" & frm100102_1.txtCountry(0) & "' "
           StrSQL4 = StrSQL4 + " AND LC15>='" & frm100102_1.txtCountry(0) & "' "
        End If
        If Len(Trim(frm100102_1.txtCountry(1))) <> 0 Then           '檢查申請國家
           strSQL1 = strSQL1 + " AND TM10<='" & frm100102_1.txtCountry(1) & "' "
           strSQL2 = strSQL2 + " AND PA09<='" & frm100102_1.txtCountry(1) & "' "
           StrSQL3 = StrSQL3 + " AND SP09<='" & frm100102_1.txtCountry(1) & "' "
           StrSQL4 = StrSQL4 + " AND LC15<='" & frm100102_1.txtCountry(1) & "' "
        End If
        '2010/3/11 end
        If Len(Trim(frm100102_1.Text6)) <> 0 Then            '檢查案件性質
            strSQL1 = strSQL1 + " AND CP10>='" & frm100102_1.Text6 & "' "
            strSQL2 = strSQL2 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL3 = StrSQL3 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL4 = StrSQL4 + " AND CP10>='" & frm100102_1.Text6 & "' "
            'Add By Cheng 2002/12/13
            StrSQL7 = StrSQL7 + " AND CP10>='" & frm100102_1.Text6 & "' "
        End If
        If Len(Trim(frm100102_1.Text7)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP10<='" & frm100102_1.Text7 & "' "
            strSQL2 = strSQL2 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL3 = StrSQL3 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL4 = StrSQL4 + " AND CP10<='" & frm100102_1.Text7 & "' "
            'Add By Cheng 2002/12/13
            StrSQL7 = StrSQL7 + " AND CP10<='" & frm100102_1.Text7 & "' "
        End If
        
           If Len(frm100102_1.Text4) <> 0 Then
              strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100102_1.Text4)) & " "
           End If
           If Len(frm100102_1.Text5) <> 0 Then
              strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100102_1.Text5)) & " "
           Else
              If Len(frm100102_1.Text4) > 0 Then
                 strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
              End If
           End If
End If


'edit by nickc  2007/12/21 加入區分欄位
'    strSQL = "SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA75='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP26='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'
'    strSQL = strSQL & " union SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and CP01=tm01(+) and CP02=tm02(+) and CP03=tm03(+) and CP04=tm04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and cp44='" & oStrFA & "' and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (cp44='" & oStrFA & "')  and CP01=SP01(+) and CP02=SP02(+) and CP03=SP03(+) and CP04=SP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and CP01=LC01(+) and CP02=LC02(+) and CP03=LC03(+) and CP04=LC04(+)   and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'
'    strSQL = strSQL & " Union SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA26='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA27='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA28='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA29='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA30='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP08='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP58='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP59='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select HC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM Hirecase ,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "'   and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL7 & strSQL5 & " group by HC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL & " order by 1,4 "
'Mark by Lydia 2019/12/26 利益衝突案件：改成先丟暫存檔，逐案件比對
'    strSql = "SELECT TM01,'FC代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSql = strSql + " union select PA01,'FC代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA75='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSql = strSql + " union select SP01,'FC代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP26='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'
'    strSql = strSql + " union select LC01,'FC代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSql = strSql & " union SELECT TM01,'CF代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and CP01=tm01(+) and CP02=tm02(+) and CP03=tm03(+) and CP04=tm04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSql = strSql + " union select PA01,'CF代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and cp44='" & oStrFA & "' and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSql = strSql + " union select SP01,'CF代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (cp44='" & oStrFA & "')  and CP01=SP01(+) and CP02=SP02(+) and CP03=SP03(+) and CP04=SP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSql = strSql + " union select LC01,'CF代理',decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and CP01=LC01(+) and CP02=LC02(+) and CP03=LC03(+) and CP04=LC04(+)   and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'end 2019/12/26
'edit by nickc 2007/12/21 秀玲說取消，因為這個 function 應該只會有代理人
'    strSQL = strSQL & " Union SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA26='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA27='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA28='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA29='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA30='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP08='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP58='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP59='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select HC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM Hirecase ,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "'   and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL7 & strSQL5 & " group by HC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "

'Modified by Lydia 2019/12/26 利益衝突案件：改成先丟暫存檔逐案件比對
'    strSql = strSql & " order by 1,5 "
    cnnConnection.Execute " delete from R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name)
    strExc(1) = "INSERT INTO R100102_1 (ID,R021001,R021006,R021003,R021007,R021014,R021015,R021016,R021017,R021018,R021020,R021021,R021022,R021023,R021024,R021025) "
    
    '案件別：FC代理
    strExc(1) = strExc(1) & " SELECT '" & strUserNum & "@" & Me.Name & "' as idname,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as caseno, CP09 , 'FC代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno" & _
                    " FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as caseno, CP09 , 'FC代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno" & _
                    " FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA75='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as caseno, CP09 , 'FC代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,SP08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno" & _
                    " FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP26='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as caseno, CP09 , 'FC代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno" & _
                    " FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'  and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5
    '案件別：CF代理
    strExc(1) = strExc(1) & "Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as caseno, CP09 , 'CF代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno" & _
                    " FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and CP01=tm01(+) and CP02=tm02(+) and CP03=tm03(+) and CP04=tm04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as caseno, CP09 , 'CF代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno" & _
                    " FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and cp44='" & oStrFA & "' and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as caseno, CP09 , 'CF代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,SP08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno" & _
                    " FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (cp44='" & oStrFA & "')  and CP01=SP01(+) and CP02=SP02(+) and CP03=SP03(+) and CP04=SP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as caseno, CP09 , 'CF代理' as 案件別,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno" & _
                    " FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and CP01=LC01(+) and CP02=LC02(+) and CP03=LC03(+) and CP04=LC04(+)   and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5
    cnnConnection.Execute strExc(1), intI
    '逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
         intCufaCnt = 0
         'Added by Lydia 2020/11/09 判斷來源
         If IsFrom100114 = False Then
             m_AllSys = IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3))
         Else
         'end 2020/11/09
             m_AllSys = IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(, frm100114_1.txt1(2).Text))
         End If 'Added by Lydia 2020/11/09
         strSql = "SELECT R021001, R021020, R021021, R021022, R021023, R021024, R021025 FROM R100102_1 WHERE id = " & CNULL(strUserNum & "@" & Me.Name) & _
                     " GROUP BY R021001, R021020, R021021, R021022, R021023, R021024, R021025 ORDER BY R021001 ASC "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
             RsTemp.MoveFirst
             Do While Not RsTemp.EOF
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & RsTemp.Fields("R021001"), "" & RsTemp.Fields("R021020") & "," & RsTemp.Fields("R021021") & "," & RsTemp.Fields("R021022") & "," & RsTemp.Fields("R021023") & "," & RsTemp.Fields("R021024"), "" & RsTemp.Fields("R021025")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    cnnConnection.Execute " delete from R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name) & " and R021001='" & RsTemp.Fields("R021001") & "' ", intI
                End If
                RsTemp.MoveNext
             Loop
         End If
         '利益衝突案件：限閱案件
         If intCufaCnt > 0 Then
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
         End If
    End If
    '加總
    strSql = "SELECT R021014 AS 系統別, R021003 AS 案件別, R021007 AS 案件性質, COUNT(*) AS CNT, R021018 AS CP10 " & _
                " FROM R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name) & " GROUP BY R021014, R021003, R021007, R021018 " & _
                " ORDER BY 1, 5 "
    'end 2019/12/26
    
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cmdok(0).Enabled = True
    cmdok(1).Enabled = True
Else
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True

End Sub

'由申請人畫面來   910801
Sub StrMenu2(oStrFA As String)
Me.Enabled = False
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
Dim StrSQL6 As String
'Add By Cheng 2002/12/13
Dim StrSQL7 As String

Dim frm As Form
Dim IsFrom100114 As Boolean
IsFrom100114 = False
For Each frm In Forms
    Select Case UCase(frm.Name)
    Case "FRM100114_1"
            IsFrom100114 = True
    End Select
Next

'開始搜尋
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
'2008/10/23 modify by sonia指定國家之B類不統計但明細資料仍顯示
'strSQL5 = ""
strSQL5 = " and not (CP04<>'00' and substr(CP09,1,1)='B') "
'2008/10/23 end
StrSQL6 = ""
'Add By Cheng 2002/12/13
StrSQL7 = ""
If IsFrom100114 = True Then
            If Len(Trim(frm100114_1.txt1(2))) <> 0 Then
               strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 2) & ") "
               strSQL2 = strSQL2 & " and PA01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 1) & ") "
               StrSQL3 = StrSQL3 & " and SP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 5) & ") "
               StrSQL4 = StrSQL4 & " and LC01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 3) & ") "
               StrSQL7 = StrSQL7 & " and HC01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 4) & ") and CP01 in (" & SQLGrpStr(IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(frm100114_1.txt1(2))), 4) & ") "
            End If
            If Len(Trim(frm100114_1.txt1(8))) <> 0 Then           '檢查申請國家
               strSQL1 = strSQL1 + " AND TM10='" & frm100114_1.txt1(8) & "' "
               strSQL2 = strSQL2 + " AND PA09='" & frm100114_1.txt1(8) & "' "
               StrSQL3 = StrSQL3 + " AND SP09='" & frm100114_1.txt1(8) & "' "
               StrSQL4 = StrSQL4 + " AND LC15='" & frm100114_1.txt1(8) & "' "
            End If
            If Len(Trim(frm100114_1.txt1(6))) <> 0 Then            '檢查案件性質
                strSQL1 = strSQL1 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                strSQL2 = strSQL2 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                StrSQL3 = StrSQL3 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                StrSQL4 = StrSQL4 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
                StrSQL7 = StrSQL7 + " AND CP10>='" & frm100114_1.txt1(6) & "' "
            End If
            If Len(Trim(frm100114_1.txt1(7))) <> 0 Then
                strSQL1 = strSQL1 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                strSQL2 = strSQL2 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                StrSQL3 = StrSQL3 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                StrSQL4 = StrSQL4 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
                StrSQL7 = StrSQL7 + " AND CP10<='" & frm100114_1.txt1(7) & "' "
            End If
            If frm100114_1.txt1(3) = "1" Then        '收文
               If Len(frm100114_1.txt1(4)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
               End If
               If Len(frm100114_1.txt1(5)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
               Else
                  If Len(frm100114_1.txt1(4).Text) > 0 Then
                     strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                  End If
               End If
            Else
               If Len(frm100114_1.txt1(4)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(frm100114_1.txt1(4))) & " "
               End If
               If Len(frm100114_1.txt1(5)) <> 0 Then
                  strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(frm100114_1.txt1(5))) & " "
               Else
                  If Len(frm100114_1.txt1(4).Text) > 0 Then
                     strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
                  End If
               End If
            End If
Else
        '系統類別
        If Len(Trim(frm100102_1.Text3.Text)) <> 0 Then
           strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 2) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 2) & ") "
           strSQL2 = strSQL2 & " and PA01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 1) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 1) & ") "
           StrSQL3 = StrSQL3 & " and SP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 5) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 5) & ") "
           StrSQL4 = StrSQL4 & " and LC01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 3) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 3) & ") "
           'Add By Cheng 2002/12/13
           StrSQL7 = StrSQL7 & " and hc01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 4) & ") and CP01 in (" & SQLGrpStr(IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3)), 4) & ") "
        End If
        '2010/3/11 add by sonia
        If Len(Trim(frm100102_1.txtCountry(0))) <> 0 Then           '檢查申請國家
           strSQL1 = strSQL1 + " AND TM10>='" & frm100102_1.txtCountry(0) & "' "
           strSQL2 = strSQL2 + " AND PA09>='" & frm100102_1.txtCountry(0) & "' "
           StrSQL3 = StrSQL3 + " AND SP09>='" & frm100102_1.txtCountry(0) & "' "
           StrSQL4 = StrSQL4 + " AND LC15>='" & frm100102_1.txtCountry(0) & "' "
        End If
        If Len(Trim(frm100102_1.txtCountry(1))) <> 0 Then           '檢查申請國家
           strSQL1 = strSQL1 + " AND TM10<='" & frm100102_1.txtCountry(1) & "' "
           strSQL2 = strSQL2 + " AND PA09<='" & frm100102_1.txtCountry(1) & "' "
           StrSQL3 = StrSQL3 + " AND SP09<='" & frm100102_1.txtCountry(1) & "' "
           StrSQL4 = StrSQL4 + " AND LC15<='" & frm100102_1.txtCountry(1) & "' "
        End If
        '2010/3/11 end
        If Len(Trim(frm100102_1.Text6)) <> 0 Then            '檢查案件性質
            strSQL1 = strSQL1 + " AND CP10>='" & frm100102_1.Text6 & "' "
            strSQL2 = strSQL2 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL3 = StrSQL3 + " AND CP10>='" & frm100102_1.Text6 & "' "
            StrSQL4 = StrSQL4 + " AND CP10>='" & frm100102_1.Text6 & "' "
            'Add By Cheng 2002/12/13
            StrSQL7 = StrSQL7 + " AND CP10>='" & frm100102_1.Text6 & "' "
        End If
        If Len(Trim(frm100102_1.Text7)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP10<='" & frm100102_1.Text7 & "' "
            strSQL2 = strSQL2 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL3 = StrSQL3 + " AND CP10<='" & frm100102_1.Text7 & "' "
            StrSQL4 = StrSQL4 + " AND CP10<='" & frm100102_1.Text7 & "' "
            'Add By Cheng 2002/12/13
            StrSQL7 = StrSQL7 + " AND CP10<='" & frm100102_1.Text7 & "' "
        End If
        
           If Len(frm100102_1.Text4) <> 0 Then
              strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(frm100102_1.Text4)) & " "
           End If
           If Len(frm100102_1.Text5) <> 0 Then
              strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(frm100102_1.Text5)) & " "
           Else
              If Len(frm100102_1.Text4) > 0 Then
                 strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
              End If
           End If
End If

    'Modify By Cheng 2002/12/13
'    strSQL = "SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA75='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (SP58='" & oStrFA & "')  and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL & " union SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and CP01=tm01(+) and CP02=tm02(+) and CP03=tm03(+) and CP04=tm04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and cp44='" & oStrFA & "' and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (cp44='" & oStrFA & "')  and CP01=SP01(+) and CP02=SP02(+) and CP03=SP03(+) and CP04=SP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and CP01=LC01(+) and CP02=LC02(+) and CP03=LC03(+) and CP04=LC04(+)   and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "

'edit by nickc 2007/12/21 修正統計方式
'    strSQL = "SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) ,count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa26='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa27='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa28='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa29='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa30='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP08='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP58='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP59='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select HC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM HIRECASE,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "'   and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL7 & strSQL5 & " group by Hc01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'Mark by Lydia 2019/12/26 利益衝突案件：改成先丟暫存檔，逐案件比對
'    strSql = "SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as CPM,CP09 as CPMCount,CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm23='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5
'    strSql = strSql + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa26='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa27='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa28='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa29='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and pa30='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
'    strSql = strSql + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP08='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP58='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP59='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP65='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP65='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
'    strSql = strSql + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc11='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5
'    strSql = strSql + " union select HC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP09,CP10 FROM HIRECASE,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and HC05='" & oStrFA & "'   and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL7 & strSQL5
'end 2019/12/26
'edit by nickc 2007/12/21 秀玲說取消，因為這個 function 應該只會有申請人
'    strSQL = strSQL & " Union SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and tm44='" & oStrFA & "' AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and PA75='" & oStrFA & "'and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and SP26='" & oStrFA & "' and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and lc22='" & oStrFA & "'   and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'
'    strSQL = strSQL & " union SELECT TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and cp44='" & oStrFA & "' and CP01=tm01(+) and CP02=tm02(+) and CP03=tm03(+) and CP04=tm04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5 & " group by TM01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and cp44='" & oStrFA & "' and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5 & " group by PA01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (cp44='" & oStrFA & "')  and CP01=SP01(+) and CP02=SP02(+) and CP03=SP03(+) and CP04=SP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5 & " group by SP01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
'    strSQL = strSQL + " union select LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),count(*),CP10 FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and cp44='" & oStrFA & "'  and CP01=LC01(+) and CP02=LC02(+) and CP03=LC03(+) and CP04=LC04(+)   and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5 & " group by LC01,decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03),CP10 "
    
'Modified by Lydia 2019/12/26 利益衝突案件：改成先丟暫存檔逐案件比對
    'add by nickc 2007/12/21  相同的資料要合併
    'strSql = "select TM01,CPM,count(CPMCount),CP10 from (" & strSql & ") group by tm01,cpm,CP10 "
    'strSql = strSql & " order by 1,4 "
    cnnConnection.Execute " delete from R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name)
    strExc(1) = "INSERT INTO R100102_1 (ID,R021001,R021006,R021007,R021014,R021015,R021016,R021017,R021018,R021020,R021021,R021022,R021023,R021024,R021025) "
    
    strExc(1) = strExc(1) & " SELECT '" & strUserNum & "@" & Me.Name & "' as idname,TM01||'-'||TM02||'-'||TM03||'-'||TM04 as caseno, CP09 , decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno" & _
                    " FROM TRADEMARK,nation,caseprogress,casepropertymap WHERE tm10=na01(+) and (nvl(tm23,'N')='" & oStrFA & "' or nvl(tm78,'N')='" & oStrFA & "' or nvl(tm79,'N')='" & oStrFA & "' or nvl(tm80,'N')='" & oStrFA & "' or nvl(tm81,'N')='" & oStrFA & "')" & _
                    " AND tm01=CP01(+) and tm02=CP02(+) and tm03=CP03(+) and tm04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL1 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as caseno, CP09 , decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno" & _
                    " FROM PATENT,nation,caseprogress,casepropertymap WHERE PA09=na01(+) and (nvl(pa26,'N')='" & oStrFA & "' or nvl(pa27,'N')='" & oStrFA & "' or nvl(pa28,'N')='" & oStrFA & "' or nvl(pa29,'N')='" & oStrFA & "' or nvl(pa30,'N')='" & oStrFA & "') " & _
                    " and PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & strSQL2 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,SP01||'-'||SP02||'-'||SP03||'-'||SP04 as caseno, CP09 , decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,SP08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno" & _
                    " FROM SERVICEPRACTICE,nation,caseprogress,casepropertymap WHERE SP09=na01(+) and (nvl(sp08,'N')='" & oStrFA & "' or nvl(sp52,'N')='" & oStrFA & "' or nvl(sp59,'N')='" & oStrFA & "' or nvl(sp65,'N')='" & oStrFA & "' or nvl(sp66,'N')='" & oStrFA & "') " & _
                    " and SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL3 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,LC01||'-'||LC02||'-'||LC03||'-'||LC04 as caseno, CP09 , decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno" & _
                    " FROM LAWCASE,nation,caseprogress,casepropertymap WHERE lc15=na01(+) and (nvl(lc11,'N')='" & oStrFA & "' or nvl(lc43,'N')='" & oStrFA & "' or nvl(lc44,'N')='" & oStrFA & "' or nvl(lc45,'N')='" & oStrFA & "' or nvl(lc46,'N')='" & oStrFA & "') " & _
                    " and LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL4 & strSQL5
    strExc(1) = strExc(1) & " Union All SELECT '" & strUserNum & "@" & Me.Name & "' as idname,HC01||'-'||HC02||'-'||HC03||'-'||HC04 as caseno, CP09 , decode(ltrim(rtrim(cpm03)),'（無）',cpm04,cpm03) as cpm03,CP01,CP02,CP03,CP04,CP10,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno" & _
                    " FROM HIRECASE,nation,caseprogress,casepropertymap WHERE '000'=na01(+) and (nvl(hc05,'N')='" & oStrFA & "' or nvl(hc24,'N')='" & oStrFA & "' or nvl(hc25,'N')='" & oStrFA & "' or nvl(hc26,'N')='" & oStrFA & "' or nvl(hc27,'N')='" & oStrFA & "') " & _
                    " and HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+)  and CP01=cpm01(+) and CP10=cpm02(+) " & StrSQL7 & strSQL5
    cnnConnection.Execute strExc(1), intI
    '逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
         intCufaCnt = 0
         'Added by Lydia 2020/11/09 判斷來源
         If IsFrom100114 = False Then
             m_AllSys = IIf(frm100102_1.Text3.Text <> "ALL", frm100102_1.Text3.Text, GetAllSysKind(frm100102_1.Text3))
         Else
         'end 2020/11/09
             m_AllSys = IIf(frm100114_1.txt1(2).Text <> "ALL", frm100114_1.txt1(2).Text, GetAllSysKind(, frm100114_1.txt1(2).Text))
         End If 'Added by Lydia 2020/11/09
         strSql = "SELECT R021001, R021020, R021021, R021022, R021023, R021024, R021025 FROM R100102_1 WHERE id = " & CNULL(strUserNum & "@" & Me.Name) & _
                     " GROUP BY R021001, R021020, R021021, R021022, R021023, R021024, R021025 ORDER BY R021001 ASC "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
             RsTemp.MoveFirst
             Do While Not RsTemp.EOF
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & RsTemp.Fields("R021001"), "" & RsTemp.Fields("R021020") & "," & RsTemp.Fields("R021021") & "," & RsTemp.Fields("R021022") & "," & RsTemp.Fields("R021023") & "," & RsTemp.Fields("R021024"), "" & RsTemp.Fields("R021025")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    cnnConnection.Execute " delete from R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name) & " and R021001='" & RsTemp.Fields("R021001") & "' ", intI
                End If
                RsTemp.MoveNext
             Loop
         End If
         '利益衝突案件：限閱案件
         If intCufaCnt > 0 Then
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
         End If
    End If
    '加總
    strSql = "SELECT R021014 AS 系統別, R021007 AS 案件性質, COUNT(*) AS CNT, R021018 AS CP10 " & _
                " FROM R100102_1 where id = " & CNULL(strUserNum & "@" & Me.Name) & " GROUP BY R021014, R021007, R021018 " & _
                " ORDER BY 1, 4 "
    'end 2019/12/26
    
CheckOC
s = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    cmdok(0).Enabled = True
    cmdok(1).Enabled = True
Else
    cmdok(0).Enabled = False
    cmdok(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100114_3 = Nothing
End Sub

Private Sub GrdDataList_Click()
Dim j As Integer
Dim i As Integer
Dim tmpcur As Integer
tmpcur = grdDataList.MouseRow
grdDataList.Visible = False
For j = 1 To grdDataList.Rows - 1
    grdDataList.row = j
    For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = QBColor(15)
    Next i
Next j

grdDataList.row = tmpcur
grdDataList.col = 0
If grdDataList.row <> 0 Then
    For i = 0 To grdDataList.Cols - 1
        grdDataList.col = i
        grdDataList.CellBackColor = &HFFC0C0
    Next i
End If
grdDataList.Visible = True
End Sub

